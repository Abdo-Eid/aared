# Export the chart oracle corpus through Office itself (Windows, PowerPoint + Excel installed).
#
#   powershell -ExecutionPolicy Bypass -File tools/chart-oracle/export-office.ps1
#
# Writes into tools/chart-oracle/out/ (git-ignored):
#   chart-oracle.pdf          PowerPoint's vector PDF of chart-oracle.pptx (one chart per slide);
#                             score.py measures every word and gridline in it
#   xlsx-cNN.png              Excel's own picture of the chart on sheet cNN, for eyeballing
#   office-version.txt        which Office build drew them
param([string]$Dir = $PSScriptRoot, [string[]]$Sheets = @())
$ErrorActionPreference = "Stop"
$out = Join-Path $Dir "out"
New-Item -ItemType Directory -Force $out | Out-Null

$pp = New-Object -ComObject PowerPoint.Application
try {
  # Open(FileName, ReadOnly, Untitled, WithWindow)
  $pres = $pp.Presentations.Open((Join-Path $Dir "chart-oracle.pptx"), -1, 0, 0)
  $pres.SaveAs((Join-Path $out "chart-oracle.pdf"), 32)   # ppSaveAsPDF
  $pres.Close()
  "PowerPoint $($pp.Version) build $($pp.Build)" | Out-File (Join-Path $out "office-version.txt")
} finally { $pp.Quit() }

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Open((Join-Path $Dir "chart-oracle.xlsx"), 0, $true)
  # Excel recalculates chart links on open, so each sheet holds its chart's
  # data in the cells the chart names (gen_corpus.py, write_case_data).
  $names = if ($Sheets.Count -gt 0) { $Sheets } else { @($wb.Worksheets | ForEach-Object { $_.Name }) }
  foreach ($name in $names) {
    $chart = $wb.Worksheets.Item($name).ChartObjects(1).Chart
    [void]$chart.Export((Join-Path $out "xlsx-$name.png"), "PNG")
  }
  $wb.Close($false)
  "Excel $($xl.Version) build $($xl.Build)" | Out-File -Append (Join-Path $out "office-version.txt")
} finally { $xl.Quit() }

Write-Host "Wrote $out"
