param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputDir,
  [Parameter(Mandatory = $true)][int]$Width
)

$ErrorActionPreference = "Stop"

if ($Width -le 0) {
  throw "aared PowerPoint export: Width must be greater than zero."
}

$source = [System.IO.Path]::GetFullPath($InputPath)
$output = [System.IO.Path]::GetFullPath($OutputDir)
if (-not [System.IO.File]::Exists($source)) {
  throw "aared PowerPoint export: input presentation does not exist: $source"
}

[System.IO.Directory]::CreateDirectory($output) | Out-Null
$application = $null
$presentation = $null

try {
  $application = New-Object -ComObject PowerPoint.Application
  # Open read-only and without a window. This must never alter the fixture.
  $presentation = $application.Presentations.Open($source, $false, $true, $false)
  $slideWidth = [double]$presentation.PageSetup.SlideWidth
  $slideHeight = [double]$presentation.PageSetup.SlideHeight
  if ($slideWidth -le 0 -or $slideHeight -le 0) {
    throw "aared PowerPoint export: presentation has invalid slide dimensions."
  }
  $height = [int][Math]::Round($Width * $slideHeight / $slideWidth)

  for ($index = 1; $index -le $presentation.Slides.Count; $index++) {
    $path = Join-Path $output "powerpoint-slide-$index.png"
    $presentation.Slides.Item($index).Export($path, "PNG", $Width, $height)
    if (-not [System.IO.File]::Exists($path)) {
      throw "aared PowerPoint export: PowerPoint did not create $path"
    }
  }

  [ordered]@{
    powerpointVersion = [string]$application.Version
    slideCount = [int]$presentation.Slides.Count
    width = $Width
    height = $height
  } | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $output "powerpoint-export.json") -Encoding utf8
} finally {
  if ($presentation -ne $null) {
    $presentation.Close()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
  }
  if ($application -ne $null) {
    $application.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
