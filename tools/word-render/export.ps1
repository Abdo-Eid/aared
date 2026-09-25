param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputDir
)

$ErrorActionPreference = "Stop"

$input = [System.IO.Path]::GetFullPath($InputPath)
$output = [System.IO.Path]::GetFullPath($OutputDir)
if (-not [System.IO.File]::Exists($input)) {
  throw "aared Word export: input document does not exist: $input"
}

[System.IO.Directory]::CreateDirectory($output) | Out-Null
# Word opens a private copy so the fixture is never touched: no `~$` lock file
# appears beside it and a stray save cannot alter the frozen bytes.
$copy = Join-Path $output "word-source.docx"
Copy-Item -LiteralPath $input -Destination $copy -Force
$pdf = Join-Path $output "word.pdf"

$application = $null
$document = $null

function Release-ComObject([object]$Target) {
  if ($Target -ne $null) {
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Target)
  }
}

try {
  $application = New-Object -ComObject Word.Application
  # Read-only automation: no dialogs, macros force-disabled, document never saved.
  $application.Visible = $false
  $application.DisplayAlerts = 0 # wdAlertsNone
  $previousSecurity = $application.AutomationSecurity
  $application.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
  try {
    # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    $document = $application.Documents.Open($copy, $false, $true, $false)
    # PAGE/NUMPAGES and TOC page numbers print as Word computes them, not as last cached.
    [void]$document.Fields.Update()
    $document.Repaginate()
    $pageCount = [int]$document.ComputeStatistics(2) # wdStatisticPages
    $document.ExportAsFixedFormat($pdf, 17) # wdExportFormatPDF
    if (-not [System.IO.File]::Exists($pdf)) {
      throw "aared Word export: Word did not create the PDF"
    }
    [ordered]@{
      wordVersion = [string]$application.Version
      pageCount = $pageCount
      pdf = "word.pdf"
    } | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $output "word-export.json") -Encoding utf8
  } finally {
    $application.AutomationSecurity = $previousSecurity
    if ($document -ne $null) {
      $document.Close($false)
      Release-ComObject $document
    }
  }
} finally {
  if ($application -ne $null) {
    $application.Quit($false)
    Release-ComObject $application
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
