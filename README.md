# EasyDocs Upload Helper

Helps you upload a whole folder of PDFs to [EasyDocs](https://easydocs.ma), one after another.

For each PDF it opens the EasyDocs upload form, attaches the file, fills in the title and
(optionally) your school, year, semester and major, and shows the PDF next to the form.
You check the details and press **Upload** yourself, and then it moves on to the next PDF.

## What you need

- Windows, macOS or Linux
- [Node.js](https://nodejs.org/) (the LTS version)
- Google Chrome or Microsoft Edge
- An EasyDocs account
- Word/PowerPoint conversion only: Windows with Microsoft Office

## How to use it

1. Download this project (**Code → Download ZIP**) and unzip it.
2. **Windows:** double-click `start-easydocs-uploader.bat`.
   **macOS/Linux:** open a terminal in the folder and run `npm install`, then `npm start`.
3. Paste the path of the folder that has your PDFs. Subfolders are included.
4. Choose **1** to upload. Log in to EasyDocs in the browser window that opens (only the first time).
5. For each PDF, check the form and press **Upload**. Press **Ctrl+S** to skip a PDF.

If you have Word or PowerPoint files, choose **2** first. It saves a PDF next to each one,
and then you can upload them with **1**.

You can close the uploader whenever you like. Next time, it continues where you stopped.

## What it stores on your computer

Everything is kept in this project's folder. Your PDF folders are not changed. The only exception is
option **2**, which adds the converted PDFs next to the originals.

- `.browser-profile/` holds your EasyDocs login. **Don't share it.** Delete it to log out.
- `progress.json` lists which PDFs are already uploaded or skipped. Delete it to start over.

To remove the uploader completely, delete the folder.

## License

MIT. See [LICENSE](LICENSE).
