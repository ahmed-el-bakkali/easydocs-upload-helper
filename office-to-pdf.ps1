# Converts Word and PowerPoint files to PDF using installed Microsoft Office
# applications (COM automation). One PDF is written next to each source file.
# Existing PDFs are left alone. Output lines are machine-readable: "OK <path>",
# "SKIP <path>", or "ERR <path> :: <message>".

param(
  [Parameter(Mandatory = $true)]
  [string[]] $Files
)

$ErrorActionPreference = "Stop"
$ppSaveAsPDF = 32
$wdFormatPDF = 17
$wordExtensions = @(".doc", ".docx", ".rtf")

$powerpoint = $null
$word = $null

function Get-PowerPoint {
  param()
  if ($script:powerpoint) { return $script:powerpoint }
  $script:powerpoint = New-Object -ComObject PowerPoint.Application
  return $script:powerpoint
}

function Get-Word {
  param()
  if ($script:word) { return $script:word }
  $script:word = New-Object -ComObject Word.Application
  $script:word.Visible = $false
  return $script:word
}

try {
  foreach ($file in $Files) {
    try {
      $source = (Resolve-Path -LiteralPath $file).Path
      $target = [System.IO.Path]::ChangeExtension($source, ".pdf")

      if (Test-Path -LiteralPath $target) {
        Write-Output "SKIP $target"
        continue
      }

      $extension = [System.IO.Path]::GetExtension($source).ToLowerInvariant()

      if ($wordExtensions -contains $extension) {
        $app = Get-Word
        # Open read-only, without confirming conversions, without a visible window.
        $document = $app.Documents.Open($source, $false, $true, $false)
        try {
          $document.SaveAs([ref] $target, [ref] $wdFormatPDF)
        } finally {
          $document.Close($false)
        }
      } else {
        $app = Get-PowerPoint
        # Open read-only, untitled, without a document window.
        $presentation = $app.Presentations.Open($source, $true, $false, $false)
        try {
          $presentation.SaveAs($target, $ppSaveAsPDF)
        } finally {
          $presentation.Close()
        }
      }

      Write-Output "OK $target"
    } catch {
      Write-Output "ERR $file :: $($_.Exception.Message)"
    }
  }
} finally {
  if ($script:powerpoint) {
    try { $script:powerpoint.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:powerpoint) | Out-Null } catch {}
  }
  if ($script:word) {
    try { $script:word.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:word) | Out-Null } catch {}
  }
}
