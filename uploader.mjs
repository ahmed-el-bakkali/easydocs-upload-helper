import { chromium } from "playwright-core";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { pathToFileURL, fileURLToPath } from "node:url";
import readline from "node:readline/promises";

const APP_DIR = path.dirname(fileURLToPath(import.meta.url));
// Everything the uploader keeps lives in its own folder, never in the user's folders.
const PROFILE_DIR = path.join(APP_DIR, ".browser-profile");
const PROGRESS_PATH = path.join(APP_DIR, "progress.json");
const OFFICE_SCRIPT = path.join(APP_DIR, "office-to-pdf.ps1");
const LEGACY_PROGRESS_FILENAME = ".easydocs-upload-progress.json";

const UPLOAD_URL = "https://easydocs.ma/upload.html";
const SUCCESS_URL = /resource(?:\.html)?(?:[?#]|$)/i;
const FILE_INPUT = "input[type='file']";
const TITLE_INPUT = "#title, #resourceTitle, input[name='title']";
const OFFICE_EXTENSIONS = [".ppt", ".pptx", ".doc", ".docx", ".rtf"];
// Earlier versions renamed skipped files with this prefix; keep ignoring them.
const LEGACY_SKIPPED_PREFIX = "skipped ";

// The EasyDocs form cascades: school -> year -> semester -> major. Each select
// stays disabled until its parent is chosen, and is then filled from the site.
const ACADEMIC_FIELDS = [
  { selector: "#school", label: "School (École)" },
  { selector: "#year", label: "Year (Année)" },
  { selector: "#semester", label: "Semester (Semestre)" },
  { selector: "#major", label: "Major (Filière)" }
];

const execFileAsync = promisify(execFile);
const sleep = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));

async function readJson(filePath, fallback) {
  try {
    return JSON.parse(await fs.readFile(filePath, "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") return fallback;
    throw error;
  }
}

async function writeJsonAtomic(filePath, value) {
  const temporaryPath = `${filePath}.tmp`;
  await fs.writeFile(temporaryPath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  await fs.rename(temporaryPath, filePath);
}

// Lists every file under the folder, at any depth, whose extension is accepted.
async function findFiles(folder, extensions) {
  const accepted = new Set(extensions);
  const results = [];

  async function visit(currentFolder) {
    const entries = await fs.readdir(currentFolder, { withFileTypes: true });
    entries.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));

    for (const entry of entries) {
      const fullPath = path.join(currentFolder, entry.name);
      if (entry.isDirectory()) {
        await visit(fullPath);
      } else if (
        entry.isFile() &&
        !entry.name.toLowerCase().startsWith(LEGACY_SKIPPED_PREFIX) &&
        accepted.has(path.extname(entry.name).toLowerCase())
      ) {
        results.push(fullPath);
      }
    }
  }

  await visit(folder);
  return results;
}

// Splits paths into batches that keep the PowerShell command line well under
// Windows' ~32K character limit.
function batchFilePaths(paths, maxChars = 8000, maxFiles = 100) {
  const batches = [];
  let current = [];
  let currentChars = 0;

  for (const filePath of paths) {
    if (current.length > 0 && (current.length >= maxFiles || currentChars + filePath.length + 1 > maxChars)) {
      batches.push(current);
      current = [];
      currentChars = 0;
    }
    current.push(filePath);
    currentChars += filePath.length + 1;
  }
  if (current.length > 0) batches.push(current);
  return batches;
}

// Writes a PDF next to every Word/PowerPoint file that does not have one yet,
// using installed Microsoft Office. Originals are left untouched.
async function convertOfficeDocuments(folder) {
  if (process.platform !== "win32") {
    console.log("Word/PowerPoint conversion only works on Windows.");
    return;
  }

  const sources = await findFiles(folder, OFFICE_EXTENSIONS);
  const pending = [];
  for (const source of sources) {
    const target = `${source.slice(0, -path.extname(source).length)}.pdf`;
    try {
      await fs.access(target);
    } catch {
      pending.push(source);
    }
  }

  console.log(`Found ${sources.length} Word/PowerPoint file(s); ${pending.length} need a PDF.`);
  if (pending.length === 0) return;

  console.log("Converting with Microsoft Office. This can take a few seconds per file...");
  let converted = 0;
  for (const batch of batchFilePaths(pending)) {
    try {
      const { stdout } = await execFileAsync(
        "powershell",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", OFFICE_SCRIPT, "-Files", ...batch],
        { windowsHide: true, maxBuffer: 10 * 1024 * 1024 }
      );
      for (const line of stdout.split(/\r?\n/).map((item) => item.trim())) {
        if (line.startsWith("OK ")) {
          converted += 1;
          console.log(`  converted: ${path.basename(line.slice(3))}`);
        } else if (line.startsWith("ERR ")) {
          console.log(`  failed: ${line.slice(4)}`);
        }
      }
    } catch (error) {
      console.log(`Office conversion could not run: ${error.message}`);
      break;
    }
  }
  console.log(`Converted ${converted} file(s) to PDF.`);
}

async function findInstalledBrowser() {
  const candidates = [process.env.CHROME_PATH].filter(Boolean);

  if (process.platform === "win32") {
    const roots = [process.env.LOCALAPPDATA, process.env.PROGRAMFILES, process.env["PROGRAMFILES(X86)"]].filter(Boolean);
    for (const root of roots) candidates.push(path.join(root, "Google", "Chrome", "Application", "chrome.exe"));
    for (const root of roots) candidates.push(path.join(root, "Microsoft", "Edge", "Application", "msedge.exe"));
  } else if (process.platform === "darwin") {
    candidates.push(
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
      "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
    );
  } else {
    candidates.push("/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/microsoft-edge");
  }

  for (const candidate of candidates) {
    try {
      await fs.access(candidate);
      return candidate;
    } catch {
      // Try the next browser path.
    }
  }
  throw new Error("Google Chrome or Microsoft Edge was not found. Install one, or set CHROME_PATH to its executable.");
}

function titleFromFilename(filePath) {
  return path.basename(filePath, path.extname(filePath)).replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
}

function isUploadUrl(currentUrl) {
  const current = new URL(currentUrl);
  const upload = new URL(UPLOAD_URL);
  return current.origin === upload.origin && current.pathname.replace(/\/$/, "") === upload.pathname.replace(/\/$/, "");
}

async function waitForUploadForm(page) {
  while (true) {
    try {
      const input = page.locator(FILE_INPUT).first();
      await input.waitFor({ state: "attached", timeout: 1500 });
      return input;
    } catch {
      if (page.isClosed()) throw new Error("The EasyDocs tab was closed.");
      console.log("Waiting for the EasyDocs upload form. Log in in the browser if asked...");
      await sleep(1500);
    }
  }
}

// Calls onSkip whenever Ctrl+S is pressed on the page, including after reloads.
async function listenForSkip(page, onSkip) {
  await page.exposeFunction("__easydocsSkipCurrent", onSkip);
  const handler = () => {
    if (window.__easydocsSkipListenerInstalled) return;
    window.__easydocsSkipListenerInstalled = true;
    window.addEventListener("keydown", (event) => {
      if (event.ctrlKey && event.key.toLowerCase() === "s") {
        event.preventDefault();
        event.stopPropagation();
        void window.__easydocsSkipCurrent();
      }
    }, true);
  };
  await page.addInitScript(handler);
  await page.evaluate(handler).catch(() => {});
}

async function addQueueOverlay(page, text) {
  await page.evaluate((text) => {
    document.querySelector("#easydocs-queue-helper")?.remove();
    const helper = document.createElement("aside");
    helper.id = "easydocs-queue-helper";
    helper.style.cssText = [
      "position:fixed", "right:18px", "bottom:18px", "z-index:2147483647",
      "max-width:340px", "padding:14px 16px", "border-radius:12px",
      "background:#0d1b2a", "color:white", "font:14px/1.4 system-ui,sans-serif",
      "box-shadow:0 10px 35px rgba(0,0,0,.28)"
    ].join(";");
    helper.textContent = text;
    document.body.appendChild(helper);
  }, text);
}

async function setWindowBounds(context, page, bounds) {
  const session = await context.newCDPSession(page);
  try {
    const { windowId } = await session.send("Browser.getWindowForTarget");
    await session.send("Browser.setWindowBounds", { windowId, bounds: { ...bounds, windowState: "normal" } });
  } finally {
    await session.detach();
  }
}

// Upload form on the left half of the screen, PDF preview on the right half.
async function splitScreen(page) {
  const screen = await page
    .evaluate(() => ({
      left: window.screen.availLeft ?? 0,
      top: window.screen.availTop ?? 0,
      width: window.screen.availWidth,
      height: window.screen.availHeight
    }))
    .catch(() => ({ left: 0, top: 0, width: 1920, height: 1040 }));
  const half = Math.floor(screen.width / 2);
  return {
    upload: { left: screen.left, top: screen.top, width: half, height: screen.height },
    preview: { left: screen.left + half, top: screen.top, width: screen.width - half, height: screen.height }
  };
}

// Waits until the select is enabled and has loaded its real options.
async function readSelectOptions(page, selector) {
  const handle = await page.waitForFunction(
    (sel) => {
      const select = document.querySelector(sel);
      if (!select || select.disabled) return null;
      const options = [...select.options]
        .filter((option) => option.value)
        .map((option) => ({ value: option.value, label: option.textContent.trim() }));
      return options.length > 0 ? options : null;
    },
    selector,
    { timeout: 30000 }
  );
  return handle.jsonValue();
}

// Lets the user pick school/year/semester/major once, from the options the
// live form offers. Pressing Enter leaves that field and the ones after it manual.
async function promptAcademicPreset(rl, page) {
  const preset = [];
  console.log("\nPick the values to fill on every document. Press Enter to leave a field (and the ones after it) manual.");

  for (const field of ACADEMIC_FIELDS) {
    let options;
    try {
      options = await readSelectOptions(page, field.selector);
    } catch {
      console.log(`No ${field.label} options loaded; the remaining fields stay manual.`);
      break;
    }

    let chosen;
    if (preset.length === 0 && options.length === 1) {
      chosen = options[0];
      console.log(`\n${field.label}: ${chosen.label} (only option)`);
    } else {
      console.log(`\n${field.label}:`);
      options.forEach((option, index) => console.log(`  ${index + 1} - ${option.label}`));
      while (true) {
        const answer = (await rl.question(`Choose [1-${options.length}, Enter = manual]: `)).trim();
        if (!answer) break;
        const number = Number(answer);
        if (Number.isInteger(number) && number >= 1 && number <= options.length) {
          chosen = options[number - 1];
          break;
        }
        console.log("Not a valid choice.");
      }
    }
    if (!chosen) break;

    await page.selectOption(field.selector, chosen.value);
    preset.push({ ...field, value: chosen.value, option: chosen.label });
  }

  if (preset.length > 0) {
    console.log("\nEvery document will be prefilled with:");
    for (const item of preset) console.log(`  ${item.label}: ${item.option}`);
  }
  return preset;
}

async function applyAcademicPreset(page, preset) {
  for (const item of preset) {
    try {
      await page.waitForFunction(
        ({ selector, value }) => {
          const select = document.querySelector(selector);
          return select && !select.disabled && [...select.options].some((option) => option.value === value);
        },
        { selector: item.selector, value: item.value },
        { timeout: 30000 }
      );
      await page.selectOption(item.selector, item.value);
    } catch {
      console.log(`Could not prefill ${item.label}; please fill it and the fields after it by hand.`);
      return;
    }
  }
}

// Progress for every folder is kept in progress.json next to the uploader,
// keyed by the folder's full path. Paths inside are relative to that folder.
async function loadProgress(folder) {
  const all = await readJson(PROGRESS_PATH, {});
  if (!all[folder]) {
    const legacy = await readJson(path.join(folder, LEGACY_PROGRESS_FILENAME), {});
    all[folder] = { completed: legacy.completed ?? [], skipped: legacy.skipped ?? [] };
  }
  return {
    progress: all[folder],
    save: () => writeJsonAtomic(PROGRESS_PATH, all)
  };
}

async function upload(rl, folder) {
  const allDocuments = await findFiles(folder, [".pdf"]);
  const { progress, save } = await loadProgress(folder);
  const done = new Set([...progress.completed, ...progress.skipped].map((item) => path.resolve(folder, item)));
  const pending = allDocuments.filter((filePath) => !done.has(filePath));

  console.log(`Found ${allDocuments.length} PDF(s); ${pending.length} still to upload.`);
  if (pending.length === 0) return;

  const prefill = (await rl.question("\nFill year, semester and major automatically? [y/N]: ")).trim().toLowerCase() === "y";

  const browserExecutable = await findInstalledBrowser();
  console.log(`Using browser: ${browserExecutable}`);

  let skipRequested = false;
  const requestSkip = async () => {
    skipRequested = true;
  };

  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    executablePath: browserExecutable,
    viewport: null
  });
  const uploadPage = context.pages()[0] ?? (await context.newPage());
  await listenForSkip(uploadPage, requestSkip);
  await uploadPage.goto(UPLOAD_URL, { waitUntil: "domcontentloaded" });

  const bounds = await splitScreen(uploadPage);
  await setWindowBounds(context, uploadPage, bounds.upload);

  // The preview uses a separate, throwaway browser so it opens as its own window.
  const previewBrowser = await chromium.launch({ headless: false, executablePath: browserExecutable });
  const previewContext = await previewBrowser.newContext({ viewport: null });
  const previewPage = await previewContext.newPage();
  await listenForSkip(previewPage, requestSkip);
  await setWindowBounds(previewContext, previewPage, bounds.preview);

  try {
    let preset = [];
    if (prefill) {
      await waitForUploadForm(uploadPage);
      preset = await promptAcademicPreset(rl, uploadPage);
    }

    for (const documentPath of pending) {
      const position = allDocuments.indexOf(documentPath) + 1;
      const filename = path.basename(documentPath);
      const relativePath = path.relative(folder, documentPath);
      skipRequested = false;
      console.log(`\n[${position}/${allDocuments.length}] ${filename}`);

      await previewPage.goto(pathToFileURL(documentPath).href, { waitUntil: "domcontentloaded" });
      if (!isUploadUrl(uploadPage.url())) await uploadPage.goto(UPLOAD_URL, { waitUntil: "domcontentloaded" });

      const fileInput = await waitForUploadForm(uploadPage);
      await fileInput.setInputFiles(documentPath);

      const titleInput = uploadPage.locator(TITLE_INPUT).first();
      if ((await titleInput.count()) > 0 && !(await titleInput.inputValue()).trim()) {
        await titleInput.fill(titleFromFilename(documentPath));
      }
      if (preset.length > 0) await applyAcademicPreset(uploadPage, preset);

      await addQueueOverlay(uploadPage, `EasyDocs queue ${position}/${allDocuments.length} — ${filename}`);
      await uploadPage.bringToFront();
      console.log("Check the form and press the EasyDocs upload button. Press Ctrl+S in the browser to skip this PDF.");

      while (!skipRequested) {
        if (uploadPage.isClosed()) throw new Error("The EasyDocs window was closed before the upload finished.");
        const currentUrl = uploadPage.url();
        if (!isUploadUrl(currentUrl) && SUCCESS_URL.test(currentUrl)) break;
        await sleep(250);
      }

      if (skipRequested) {
        progress.skipped.push(relativePath);
        console.log(`Skipped: ${filename}`);
        await uploadPage.goto(UPLOAD_URL, { waitUntil: "domcontentloaded" });
      } else {
        progress.completed.push(relativePath);
        console.log(`Uploaded: ${filename}`);
      }
      await save();
    }

    console.log("\nAll PDFs in this folder are done.");
  } finally {
    await previewBrowser.close().catch(() => {});
    await context.close().catch(() => {});
  }
}

async function main() {
  console.log(`
========================================
     EasyDocs Assisted Uploader
========================================`);

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  try {
    const argFolder = process.argv.slice(2).join(" ").trim();
    const answer = argFolder || (await rl.question("Paste the folder path: "));
    const requested = answer.trim().replace(/^"|"$/g, "");
    if (!requested) throw new Error("No folder was provided.");

    const folder = path.resolve(requested);
    if (!(await fs.stat(folder)).isDirectory()) throw new Error(`Not a folder: ${folder}`);

    console.log(`
1 - Upload the PDFs in this folder (and all subfolders)
2 - Convert Word/PowerPoint files in this folder to PDF
0 - Exit
`);
    const mode = (await rl.question("Choose [1/2/0]: ")).trim();
    if (mode === "1") await upload(rl, folder);
    else if (mode === "2") await convertOfficeDocuments(folder);
    else console.log("Exiting.");
  } finally {
    rl.close();
  }
}

main().catch((error) => {
  console.error(`\nError: ${error.message}`);
  process.exitCode = 1;
});
