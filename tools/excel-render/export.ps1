param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputDir,
  [string]$PartsJson = "[]"
)

$ErrorActionPreference = "Stop"

$input = [System.IO.Path]::GetFullPath($InputPath)
$output = [System.IO.Path]::GetFullPath($OutputDir)
if (-not [System.IO.File]::Exists($input)) {
  throw "aared Excel export: input workbook does not exist: $input"
}

# Parts arrive as JSON: [{sheet (1-based tab position), name, range}].
# Names become file slugs.
$parsedParts = @()
foreach ($entry in ($PartsJson | ConvertFrom-Json)) {
  $slug = ($entry.name -replace '[^A-Za-z0-9._-]+', '-').Trim('-')
  if ($slug -eq "") { throw "aared Excel export: part name '$($entry.name)' has no usable file characters" }
  $parsedParts += [pscustomobject]@{ Sheet = [int]$entry.sheet; Name = [string]$entry.name; Slug = $slug; Range = [string]$entry.range }
}

[System.IO.Directory]::CreateDirectory($output) | Out-Null
$application = $null
$workbook = $null

function Release-ComObject([object]$Target) {
  if ($Target -ne $null) {
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Target)
  }
}

try {
  $application = New-Object -ComObject Excel.Application
  # Read-only automation that must never alter the fixture: no link updates,
  # macros force-disabled, no dialogs, workbook never saved. Unlike
  # CopyPicture, ExportAsFixedFormat renders through the print pipeline and
  # needs no visible window.
  $application.Visible = $false
  $application.DisplayAlerts = $false
  $application.ScreenUpdating = $false
  $application.EnableEvents = $false
  $application.AskToUpdateLinks = $false
  # "Scale content for A4 or 8.5 x 11" maps A4 to the locale's paper on
  # export (US Letter here); off, the hardcoded A4 policy prints on A4.
  $previousMapPaper = $application.MapPaperSize
  $application.MapPaperSize = $false
  $previousSecurity = $application.AutomationSecurity
  $application.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
  try {
    $workbook = $application.Workbooks.Open($input, 0, $true)
    $exported = @()
    $position = 0
    foreach ($sheet in $workbook.Worksheets) {
      $used = $null
      try {
        if ($sheet.Visible -ne -1) { continue } # xlSheetVisible
        $position++
        # Hardcoded print policy (deliberate: one consistent A4 truth owned by
        # tooling, not read from the file): A4 portrait, fit sheet on one
        # page, gridlines plus row/column headings, normal margins (Excel reads
        # margins in points, so inches go through InchesToPoints).
        $setup = $sheet.PageSetup
        # Page setup is sent to the printer driver as one batch: set one at a
        # time, the driver keeps its own paper (US Letter on a default
        # "Microsoft Print to PDF") although PaperSize reads back as A4.
        $application.PrintCommunication = $false
        try {
          $setup.PaperSize = 9 # xlPaperA4
          $setup.Orientation = 1 # xlPortrait
          $setup.Zoom = $false
          $setup.FitToPagesWide = 1
          $setup.FitToPagesTall = 1
          $setup.PrintGridlines = $true
          $setup.PrintHeadings = $true
          $setup.TopMargin = $application.InchesToPoints(0.75)
          $setup.BottomMargin = $application.InchesToPoints(0.75)
          $setup.LeftMargin = $application.InchesToPoints(0.7)
          $setup.RightMargin = $application.InchesToPoints(0.7)
          $setup.HeaderMargin = $application.InchesToPoints(0.3)
          $setup.FooterMargin = $application.InchesToPoints(0.3)
        } finally {
          $application.PrintCommunication = $true
          Release-ComObject $setup
        }
        $used = $sheet.UsedRange
        $address = [string]$used.Address($false, $false)
        $isEmpty = $used.Cells.Count -eq 1 -and [string]::IsNullOrEmpty([string]$used.Text)
        $fileName = $null
        $partFiles = @()
        if (-not $isEmpty) {
          $fileName = "excel-sheet-$position.pdf"
          # 0 = xlTypePDF, 0 = xlQualityStandard.
          $sheet.ExportAsFixedFormat(0, (Join-Path $output $fileName), 0, $true, $false)
          if (-not [System.IO.File]::Exists((Join-Path $output $fileName))) {
            throw "aared Excel export: Excel did not create $fileName"
          }
          foreach ($part in @($parsedParts | Where-Object { $_.Sheet -eq $position })) {
            $partFile = "excel-sheet-$position-part-$($part.Slug).pdf"
            $sheet.PageSetup.PrintArea = $part.Range
            try {
              $sheet.ExportAsFixedFormat(0, (Join-Path $output $partFile), 0, $true, $false)
            } finally {
              $sheet.PageSetup.PrintArea = ""
            }
            if (-not [System.IO.File]::Exists((Join-Path $output $partFile))) {
              throw "aared Excel export: Excel did not create $partFile"
            }
            $partFiles += [ordered]@{ name = $part.Name; range = $part.Range; file = $partFile }
          }
        } elseif (@($parsedParts | Where-Object { $_.Sheet -eq $position }).Count -gt 0) {
          throw "aared Excel export: sheet $position is empty but has curated parts"
        }
        $exported += [ordered]@{
          index = $position
          name = [string]$sheet.Name
          direction = $(if ($sheet.DisplayRightToLeft) { "rtl" } else { "ltr" })
          usedRange = $address
          empty = [bool]$isEmpty
          file = $fileName
          parts = $partFiles
        }
      } finally {
        Release-ComObject $used
        Release-ComObject $sheet
      }
    }

    [ordered]@{
      excelVersion = [string]$application.Version
      sheetCount = [int]$exported.Count
      sheets = $exported
    } | ConvertTo-Json -Compress -Depth 6 | Set-Content -LiteralPath (Join-Path $output "excel-export.json") -Encoding utf8
  } finally {
    $application.AutomationSecurity = $previousSecurity
    $application.MapPaperSize = $previousMapPaper
    if ($workbook -ne $null) {
      $workbook.Close($false)
      Release-ComObject $workbook
    }
  }
} finally {
  if ($application -ne $null) {
    $application.Quit()
    Release-ComObject $application
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
