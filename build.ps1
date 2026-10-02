# Builds the Neudata Word templates (.dotx) and PDF previews.
# Requires Python 3 with python-docx, and Microsoft Word on Windows.
# Run:  powershell -ExecutionPolicy Bypass -File build.ps1

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
python (Join-Path $here "build_templates.py")
if ($LASTEXITCODE -ne 0) { throw "build_templates.py failed" }

$build = Join-Path $here "build"
$previews = Join-Path $here "previews"
New-Item -ItemType Directory -Force $previews | Out-Null

foreach ($f in Get-ChildItem $build -Filter *.docx) {
  $base = [IO.Path]::GetFileNameWithoutExtension($f.Name)
  # A fresh Word instance per file: Word sometimes crashes while closing
  # after saving a template, which would otherwise stop the whole build.
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($f.FullName, $false, $false)
  foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
  $doc.Fields.Update() | Out-Null
  foreach ($s in $doc.Sections) {
    foreach ($h in $s.Headers) { $h.Range.Fields.Update() | Out-Null }
    foreach ($ft in $s.Footers) { $ft.Range.Fields.Update() | Out-Null }
  }
  $doc.ExportAsFixedFormat((Join-Path $previews "$base.pdf"), 17)   # wdExportFormatPDF
  $doc.SaveAs2((Join-Path $here "$base.dotx"), 14)                   # wdFormatXMLTemplate
  try { $doc.Close($false); $word.Quit() } catch { }
  Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } |
    Stop-Process -Force -ErrorAction SilentlyContinue
  "Built $base.dotx"
}
