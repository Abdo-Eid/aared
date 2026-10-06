# Aptos has no Arabic glyphs. What do Excel and Word draw when Arabic text is set in
# Aptos, Aptos Narrow or Aptos Display? Excel: one cell font per cell, no complex-script
# slot, so this is the default case for Arabic typed into a new workbook (Aptos Narrow
# is Excel's default font). Word: a run without w:rtl whose Latin font is Aptos, and a
# w:rtl run whose complex-script font is Aptos. Each case is labelled in Arial beside
# the Arabic, exported to PDF by the app itself, and read by aptos-arabic-fallback.py
# (which font the PDF says drew the Arabic, and its width against Arial's).
#
#   powershell -File tools/font-calibration/aptos-arabic-fallback.ps1
#   python tools/font-calibration/aptos-arabic-fallback.py
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$samples = (Get-Content -Raw -Encoding UTF8 "$here/arabic-survey.last.json" | ConvertFrom-Json).samples |
    ForEach-Object { if ($_.Length -gt 36) { $_.Substring(0, 36).Trim() } else { $_ } } |
    Where-Object { $_.Length -ge 14 } | Select-Object -First 8
$fonts = @("Aptos", "Aptos Narrow", "Aptos Display", "Arial", "Calibri")

# --- Excel ---------------------------------------------------------------------------
$excelWasRunning = [bool](Get-Process EXCEL -ErrorAction SilentlyContinue)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$book = $null
try {
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Columns.Item(1).ColumnWidth = 30
    $sheet.Columns.Item(2).ColumnWidth = 120
    $row = 1
    foreach ($font in $fonts) {
        foreach ($sample in $samples) {
            $label = $sheet.Cells.Item($row, 1)
            $label.Value2 = "EXCEL|$font"
            $label.Font.Name = "Arial"
            $label.Font.Size = 8
            $cell = $sheet.Cells.Item($row, 2)
            $cell.Value2 = $sample
            $cell.Font.Name = $font
            $cell.Font.Size = 20
            $row++
        }
    }
    $sheet.Rows.RowHeight = 30
    $setup = $sheet.PageSetup
    $excel.PrintCommunication = $false
    $setup.Zoom = $false
    $setup.FitToPagesWide = 1
    $setup.FitToPagesTall = $false
    $excel.PrintCommunication = $true
    $book.ExportAsFixedFormat(0, "$here/aptos-arabic-fallback.excel.pdf")
    Write-Output "Excel: exported $($row - 1) cells"
}
finally {
    if ($book -ne $null) { $book.Close($false) }
    if (-not $excelWasRunning) { $excel.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# --- Word ----------------------------------------------------------------------------
$wordWasRunning = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
try {
    $doc = $word.Documents.Add()
    $doc.PageSetup.PageWidth = 1584
    $doc.PageSetup.LeftMargin = 36
    $doc.PageSetup.RightMargin = 36
    $first = $true
    foreach ($mode in @("latin-slot", "cs-slot")) {
        foreach ($font in $fonts) {
            foreach ($sample in $samples) {
                if (-not $first) { $doc.Content.InsertParagraphAfter() }
                $first = $false
                $range = $doc.Paragraphs.Last.Range
                $range.InsertBefore("WORD|$mode|$font ")
                $range = $doc.Paragraphs.Last.Range
                $range.Font.Name = "Arial"
                $range.Font.NameBi = "Arial"
                $range.Font.Size = 8
                $range.ParagraphFormat.SpaceAfter = 0
                $end = $range.End - 1
                $range.InsertAfter($sample)
                $arabic = $doc.Range($end, $doc.Paragraphs.Last.Range.End - 1)
                $arabic.Font.Size = 20
                $arabic.Font.SizeBi = 20
                if ($mode -eq "latin-slot") {
                    # No w:rtl: Word uses the Latin font and glyph fallback (DECISION_LOG §23).
                    $arabic.Font.Name = $font
                    $arabic.Select()
                    $word.Selection.LtrRun()
                }
                else {
                    $arabic.Select()
                    $word.Selection.RtlRun()
                    $arabic.Font.NameBi = $font
                }
            }
        }
    }
    $doc.ExportAsFixedFormat("$here/aptos-arabic-fallback.word.pdf", 17)
    Write-Output "Word: exported"
}
catch {
    Write-Output "Word failed: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
}
finally {
    if ($doc -ne $null) { $doc.Close([ref]0) }
    if (-not $wordWasRunning) { $word.Quit([ref]0) }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
