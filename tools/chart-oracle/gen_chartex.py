"""Generate tools/chart-oracle/chartex-oracle.xlsx: one Office 2016 chart per
sheet (waterfall, funnel, histogram, pareto, box and whisker, treemap,
sunburst, and a few variants), written by a private Excel instance so the
chartex XML is exactly what Office writes.

    python gen_chartex.py <output.xlsx>

Uses DispatchEx (a separate Excel process, never the user's open one).
Each chart's PNG (Excel's own render) is exported next to the workbook: it
is the visual oracle the engine's layout is read against.

Finding: Excel writes an XLSX chartex part with no data cache. Each `cx:f`
names a hidden `_xlchart.v1.N` defined name, and the values live in the
cells; only PPTX and DOCX (which embed a workbook) carry `cx:lvl` caches.
"""
import os, sys
import win32com.client

OUT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "chartex-oracle.xlsx")

# (sheet name, xl chart type): xlWaterfall 119, xlFunnel 123, xlHistogram 118,
# xlPareto 122, xlBoxwhisker 121, xlTreemap 117, xlSunburst 120.
SHEETS = [
    ("waterfall", 119), ("funnel", 123), ("histogram", 118), ("pareto", 122),
    ("boxwhisker", 121), ("treemap", 117), ("sunburst", 120),
    ("histogram2", 118), ("treemap2", 117), ("sunburst2", 120),
    ("waterfall2", 119), ("pareto2", 122),
    # Settings the engine reads but no sheet above sets (DECISION_LOG §74): a bin count with overflow
    # and underflow bins, inclusive quartiles with the mean line, banner parent labels.
    ("histogram-bins", 118), ("box-inclusive", 121), ("treemap-banner", 117),
]

# Series settings per sheet, set over COM after the chart is made (Excel's own property names).
# BinsType 4 is xlBinsTypeBinCount; ParentDataLabelOption 1 is xlParentDataLabelOptionsBanner.
SETTINGS = {
    "histogram-bins": [("BinsType", 4), ("BinsCountValue", 5), ("BinsOverflowEnabled", True), ("BinsOverflowValue", 20),
                       ("BinsUnderflowEnabled", True), ("BinsUnderflowValue", 3)],
    "box-inclusive": [("QuartileCalculationInclusiveMedian", True), ("ShowMeanLine", True)],
    "treemap-banner": [("ParentDataLabelOption", 1)],
}

# 1-based points set as totals (the chart's own "Set as Total").
TOTALS = {"waterfall": [5], "waterfall2": [5, 7]}

TREE = [["Branch", "Item", "Sales"], ["فاكهة", "تفاح", 40], ["فاكهة", "موز", 25], ["فاكهة", "برتقال", 15], ["خضار", "جزر", 30], ["خضار", "خيار", 10], ["ألبان", "حليب", 20]]
# Unsorted: shows whether a chart sorts its branches and leaves.
TREE_UNSORTED = [["Branch", "Item", "Sales"], ["B", "b1", 10], ["B", "b2", 30], ["A", "a1", 5], ["A", "a2", 15], ["C", "c1", 60]]

DATA = {
    "waterfall": [["Item", "Amount"], ["البداية", 100], ["المبيعات", 60], ["التكاليف", -40], ["الضرائب", -15], ["الصافي", 105]],
    "funnel": [["Stage", "Count"], ["الزيارات", 1000], ["التسجيل", 640], ["التجربة", 380], ["الشراء", 120]],
    "histogram": [["Value"]] + [[v] for v in [1, 2, 2, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 7, 8, 8, 9, 12, 13, 13, 14, 15, 18, 21, 22, 25, 29, 35]],
    "histogram2": [["Value"]] + [[v] for v in [1204, 1350, 1480, 1500, 1620, 1710, 1800, 1850, 1900, 2050, 2100, 2300, 2450, 2600, 3000, 3400, 4100, 5200, 6100, 9800]],
    "pareto": [["Defect", "Count"], ["Scratch", 42], ["Dent", 25], ["Crack", 14], ["Stain", 9], ["Other", 4]],
    "boxwhisker": [["Group", "Score"]] + [[g, v] for g, vs in (("Alpha", [3, 5, 6, 7, 8, 9, 10, 11, 25]), ("بيتا", [10, 12, 13, 14, 15, 16, 18, 19, 20])) for v in vs],
    "treemap": TREE,
    "sunburst": TREE,
    # A total that is not the running sum shows whether a total bar draws its own value or the sum.
    "waterfall2": [["Item", "Amount"], ["Start", 100], ["Up", 60], ["Down", -40], ["Down2", -15], ["Total A", 90], ["Up2", 10], ["Total B", 120]],
    "pareto2": [["Defect", "Count"], ["Dent", 25], ["Scratch", 42], ["Stain", 9], ["Dent", 5], ["Crack", 14]],
    "treemap2": TREE_UNSORTED,
    "sunburst2": TREE_UNSORTED,
}
DATA["histogram-bins"] = DATA["histogram"]
DATA["box-inclusive"] = DATA["boxwhisker"]
DATA["treemap-banner"] = TREE

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = None
try:
    wb = excel.Workbooks.Add()
    while wb.Worksheets.Count < len(SHEETS):
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    for i, (name, code) in enumerate(SHEETS):
        ws = wb.Worksheets(i + 1)
        ws.Name = name
        rows = DATA[name]
        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                ws.Cells(r + 1, c + 1).Value = v
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), len(rows[0])))
        # Office 2016 chart types ignore SetSourceData: they take the selection.
        ws.Activate()
        rng.Select()
        shape = ws.Shapes.AddChart2(-1, code, 200, 10, 480, 300)
        shape.Name = name + " chart"
        for point in TOTALS.get(name, []):
            try:
                shape.Chart.FullSeriesCollection(1).Points(point).IsTotal = True
            except Exception as e:
                print("IsTotal failed:", e)
        for prop, value in SETTINGS.get(name, []):
            try:
                setattr(shape.Chart.FullSeriesCollection(1), prop, value)
            except Exception as e:
                print(f"{name}: {prop} failed:", e)
    for i, (name, _) in enumerate(SHEETS):
        ws = wb.Worksheets(i + 1)
        ws.Activate()
        ws.ChartObjects(1).Chart.Export(os.path.join(os.path.dirname(OUT), f"chartex-{name}.png"))
    wb.SaveAs(OUT, 51)
    print("saved", OUT)
finally:
    if wb is not None:
        wb.Close(0)
    excel.Quit()
