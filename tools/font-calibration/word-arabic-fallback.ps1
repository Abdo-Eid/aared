# What does Word actually draw for Arabic text whose complex-script font has NO Arabic
# glyphs (Cambria, Calibri)? Lay the same Arabic samples out in real Word, one paragraph
# per (font, sample), export to PDF, and read each line's real width with PyMuPDF
# (word-arabic-fallback.py) — geometry from the export, not caret positions (Word's
# Information() is unreliable for bidi caret ends). Compare each font's total width to
# the fonts we can also measure: a ratio near 1.00 identifies Word's fallback.
#
#   powershell -File tools/font-calibration/word-arabic-fallback.ps1
#   python tools/font-calibration/word-arabic-fallback.py
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$samples = (Get-Content -Raw -Encoding UTF8 "$here/arabic-survey.last.json" | ConvertFrom-Json).samples |
    ForEach-Object { if ($_.Length -gt 36) { $_.Substring(0, 36).Trim() } else { $_ } } |
    Where-Object { $_.Length -ge 14 } | Select-Object -First 14
$fonts = @("Arial", "Times New Roman", "Tahoma", "Cambria", "Calibri", "Segoe UI",
           "Simplified Arabic", "Traditional Arabic", "Sakkal Majalla", "Arabic Typesetting")

$wasRunning = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
try {
    $doc = $word.Documents.Add()
    $doc.PageSetup.PageWidth = 1584      # widest allowed: no line wrapping
    $doc.PageSetup.LeftMargin = 36
    $doc.PageSetup.RightMargin = 36
    $order = @()
    $first = $true
    foreach ($font in $fonts) {
        foreach ($sample in $samples) {
            if (-not $first) { $doc.Content.InsertParagraphAfter() }
            $first = $false
            $paragraph = $doc.Paragraphs.Last
            $range = $paragraph.Range
            $range.InsertBefore($sample)
            $range = $doc.Paragraphs.Last.Range
            $range.ParagraphFormat.ReadingOrder = 0      # RTL
            $range.ParagraphFormat.Alignment = 0
            $range.ParagraphFormat.PageBreakBefore = -1  # one sample per page: an exact sample-to-page mapping
            $range.ParagraphFormat.SpaceAfter = 0
            $range.ParagraphFormat.SpaceBefore = 0
            $range.Font.NameBi = $font                   # the complex-script slot Arabic uses
            $range.Font.SizeBi = 40
            $range.Font.BoldBi = 0
            $order += ("{0}`t{1}" -f $font, $sample)
        }
    }
    $pdf = "$here/word-arabic-fallback.pdf"
    $doc.ExportAsFixedFormat($pdf, 17)
    $order | Set-Content -Encoding UTF8 "$here/word-arabic-fallback.order.tsv"
    Write-Output ("exported {0} lines to {1}" -f $order.Count, $pdf)
}
finally {
    if ($doc -ne $null) { $doc.Close(0) }
    if (-not $wasRunning) { $word.Quit(0) }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
