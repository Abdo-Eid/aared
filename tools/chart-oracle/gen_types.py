"""
Writes types-oracle.xlsx with Excel itself: one stock chart of each kind
(xlStockHLC 88, OHLC 89, VHLC 90, VOHLC 91) and the two of-pie charts
(xlPieOfPie 68, xlBarOfPie 71), each on its own sheet over a small data range,
and exports every chart to PNG next to the workbook.

Excel is driven through its own instance (DispatchEx), hidden, and quit in
`finally`; nothing else on the machine is touched.

    python tools/chart-oracle/gen_types.py [out_dir]

`out_dir` defaults to tools/chart-oracle/out/types (gitignored with the other
oracle outputs); the workbook is saved there as types-oracle.xlsx, the PNGs as
types-<sheet>.png. Copy the workbook over tools/chart-oracle/types-oracle.xlsx
to refresh the corpus the tests read.
"""

import os
import sys

import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "out", "types"))
os.makedirs(OUT, exist_ok=True)

# Eight trading days; Arabic weekday names label the categories of the first sheet.
DAYS_AR = [
    "الأحد", "الاثنين", "الثلاثاء",
    "الأربعاء", "الخميس", "الأحد 2",
    "الاثنين 2", "الثلاثاء 2",
]
DAYS_EN = [f"D{i}" for i in range(1, 9)]
#         open   high   low    close  volume
OHLCV = [
    (20.0, 24.0, 18.5, 23.0, 1200),   # up
    (23.0, 25.5, 21.0, 21.5, 1500),   # down
    (21.5, 22.5, 19.0, 22.0, 900),    # up
    (22.0, 26.0, 21.5, 25.0, 2100),   # up
    (25.0, 27.0, 22.5, 23.0, 1800),   # down
    (23.0, 24.0, 20.0, 20.5, 1300),   # down
    (20.5, 23.5, 20.0, 23.0, 1000),   # up
    (23.0, 28.0, 22.0, 27.0, 2500),   # up
]
PIE = [("A", 38), ("B", 25), ("C", 14), ("D", 9), ("E", 7), ("F", 4), ("G", 2), ("H", 1)]
PIE_AR = ["تمر", "زيتون", "جبن", "لبن",
          "عسل", "خبز", "أرز", "شاي"]

# (sheet, chart type, columns of OHLCV used, category labels)
STOCKS = [
    ("HLC", 88, [1, 2, 3], DAYS_AR),
    ("OHLC", 89, [0, 1, 2, 3], DAYS_EN),
    ("VHLC", 90, [4, 1, 2, 3], DAYS_EN),
    ("VOHLC", 91, [4, 0, 1, 2, 3], DAYS_EN),
]
NAMES = ["Open", "High", "Low", "Close", "Volume"]


def main() -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        book = excel.Workbooks.Add()
        while book.Worksheets.Count > 1:
            book.Worksheets(book.Worksheets.Count).Delete()
        first = True
        charts: list[tuple[str, object]] = []

        def sheet(name: str):
            nonlocal first
            if first:
                ws = book.Worksheets(1)
                first = False
            else:
                ws = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
            ws.Name = name
            # The XLSX reader sizes a chart from its anchor only inside the sheet's used range
            # (it measures columns and rows up to the last cell holding data), so mark the corner.
            ws.Cells(22, 16).Value = "end"
            return ws

        for name, kind, cols, cats in STOCKS:
            ws = sheet(name)
            ws.Cells(1, 1).Value = "Day"
            for j, c in enumerate(cols):
                ws.Cells(1, 2 + j).Value = NAMES[c]
            for i, row in enumerate(OHLCV):
                ws.Cells(2 + i, 1).Value = cats[i]
                for j, c in enumerate(cols):
                    ws.Cells(2 + i, 2 + j).Value = row[c]
            last = ws.Cells(1 + len(OHLCV), 1 + len(cols)).Address
            obj = ws.Shapes.AddChart2(-1, kind, 220, 10, 480, 300)
            obj.Chart.SetSourceData(ws.Range(f"A1:{last}"))
            charts.append((name, obj))

        for name, kind, labels in (("PieOfPie", 68, PIE_AR), ("BarOfPie", 71, [p[0] for p in PIE])):
            ws = sheet(name)
            ws.Cells(1, 1).Value = "Item"
            ws.Cells(1, 2).Value = "Share"
            for i, (_, v) in enumerate(PIE):
                ws.Cells(2 + i, 1).Value = labels[i]
                ws.Cells(2 + i, 2).Value = v
            obj = ws.Shapes.AddChart2(-1, kind, 220, 10, 480, 300)
            obj.Chart.SetSourceData(ws.Range(f"A1:B{1 + len(PIE)}"))
            charts.append((name, obj))

        for name, obj in charts:
            obj.Chart.Export(os.path.join(OUT, f"types-{name}.png"))
        book.SaveAs(os.path.join(OUT, "types-oracle.xlsx"), 51)
        print("saved", os.path.join(OUT, "types-oracle.xlsx"))
    finally:
        if book is not None:
            book.Close(0)
        excel.Quit()


if __name__ == "__main__":
    main()
