"""
Build cecl_allowance.xlsx, then let Excel calculate it and add the sensitivity tables.

    python build_workbook.py        # needs Excel on Windows for the second step

The workbook estimates the loss-rate model itself with LINEST on the History
sheet, drives it with the Fed's 2026 scenarios, and runs a mortgage book
through the forecast, reversion and long-run periods. Every number outside the
Inputs and data sheets is a formula, so changing an assumption reprices it.
"""
import time
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import cecl

ROOT = Path(__file__).parent
OUT = ROOT / "cecl_allowance.xlsx"
BLUE, BOLD, TITLE, GREY = Font(color="0000FF"), Font(bold=True), Font(bold=True, size=14), Font(color="808080")
HEAD, INPUT = PatternFill("solid", fgColor="DCE6F1"), PatternFill("solid", fgColor="FFF2CC")
PCT, PCT3, USD, NUM4 = "0.00%", "0.000%", "#,##0.00", "0.0000"
FED_9Q_LOSS = 0.015      # 2026 stress test, first-lien mortgages, 32 banks (public_results_DFAST_2026.csv)
MAX_Q = 120
SCEN = [("Baseline", "BL"), ("Severely adverse", "SA")]


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def header(ws, row, labels, col=1, height=30):
    for j, h in enumerate(labels, col):
        c = ws.cell(row, j, h)
        c.font, c.fill = BOLD, HEAD
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height


INPUTS = [
    ("Mortgage book", None, None, None, None),
    ("Balance", "Balance at 31 Dec 2025, $m", 1000, USD, None),
    ("NoteRate", "Note rate", 0.065, PCT, None),
    ("RemQ", "Remaining term, quarters", 112, "0", "up to 120"),
    ("CPR", "Prepayment speed (CPR)", 0.08, PCT, None),
    ("CECL assumptions", None, None, None, None),
    ("RSInput", "Reasonable and supportable period, quarters", 8, "0", "1 to 13, the length of the Fed scenarios"),
    ("RevInput", "Reversion period, quarters", 4, "0", "straight line to the long-run rate"),
    ("LongRunOverride", "Long-run loss rate, leave blank for the 1991-2025 average", None, PCT3, None),
    ("WeightBL", "Weight on the Fed baseline", 0.6, "0%", None),
    ("WeightSA", "Weight on the Fed severely adverse", 0.4, "0%", "weights should add to 100%"),
    ("Loss-rate model", None, None, None, None),
    ("Eps", "Offset inside the log", cecl.EPS, "0.00", "ln(max(charge-off rate, 0) + offset), rates in %"),
]


def inputs_sheet(wb):
    ws = wb.active
    ws.title = "Inputs"
    ws["A1"], ws["A1"].font = "CECL allowance for a mortgage book: inputs", TITLE
    ws["A2"] = "Yellow cells with blue text are inputs. Every other sheet is formulas."
    row = 4
    for nm, label, value, fmt, note in INPUTS:
        if label is None:
            row += 1
            ws.cell(row, 1, nm).font = BOLD
            row += 1
            continue
        ws.cell(row, 1, label)
        c = ws.cell(row, 2, value)
        c.font, c.fill, c.number_format = BLUE, INPUT, fmt
        ws.cell(row, 3, nm).font = Font(color="808080", italic=True)
        if note:
            ws.cell(row, 4, note).font = Font(color="595959")
        name(wb, nm, f"Inputs!$B${row}")
        if nm in ("RSInput", "RevInput"):
            dv = DataValidation(type="whole", operator="between", formula1="1",
                                formula2="13" if nm == "RSInput" else "40")
            ws.add_data_validation(dv)
            dv.add(c)
        row += 1
    row += 1
    ws.cell(row, 1, "Used").font = BOLD
    # Allowance!H7 is the long-run rate table's driver cell; empty except while that table runs
    derived = [("LongRun", "Long-run loss rate used",
                "=IF(ISBLANK(Allowance!$H$7),IF(ISBLANK(LongRunOverride),HistMean,LongRunOverride),Allowance!$H$7)",
                PCT3),
               ("SMM", "Quarterly prepayment rate", "=1-(1-CPR)^(1/4)", PCT3)]
    for nm, label, f, fmt in derived:
        row += 1
        ws.cell(row, 1, label)
        ws.cell(row, 2, f).number_format = fmt
        ws.cell(row, 3, nm).font = Font(color="808080", italic=True)
        name(wb, nm, f"Inputs!$B${row}")
    for col, w in zip("ABCD", (52, 12, 16, 44)):
        ws.column_dimensions[col].width = w


def history_sheet(wb, hist):
    ws = wb.create_sheet("History")
    ws["A1"], ws["A1"].font = "Loss-rate model, estimated with LINEST", TITLE
    ws["A2"] = ("Bank net charge-offs on single-family mortgages (FRED CORSFRMACBS, annualised) against unemployment and "
                "the 12-quarter change in house prices, both from the Fed's 2026 historical file.")
    header(ws, 5, ["Quarter", "Unemployment, %", "House prices", "12-quarter change, %", "Charge-off rate, %",
                   "ln(max(rate, 0) + offset)", "Fitted rate, %"])
    h = hist.loc["1988Q1":"2025Q4"]
    first = 6
    for i, (q, r) in enumerate(h.iterrows(), first):
        ws.cell(i, 1, str(q))
        ws.cell(i, 2, float(r.ur)).number_format = "0.0"
        ws.cell(i, 3, float(r.hpi)).number_format = "0.0"
        if i >= first + cecl.HPA_QUARTERS:
            ws.cell(i, 4, f"=(C{i}/C{i - cecl.HPA_QUARTERS}-1)*100").number_format = "0.00"
        if pd.notna(r.co) and str(q) >= cecl.FIRST:
            ws.cell(i, 5, float(r.co)).number_format = "0.00"
            ws.cell(i, 6, f"=LN(MAX(E{i},0)+Eps)").number_format = NUM4
            ws.cell(i, 7, f"=MAX(EXP(Beta0+BetaUR*B{i}+BetaHPA*D{i})-Eps,0)").number_format = "0.00"
    last = first + len(h) - 1
    r0 = first + cecl.HPA_QUARTERS          # 1991Q1
    assert ws.cell(r0, 1).value == cecl.FIRST
    y = f"$F${r0}:$F${last}"
    # LINEST wants the regressors as one contiguous block, so B and D are joined with CHOOSE
    xs = f"CHOOSE({{1,2}},$B${r0}:$B${last},$D${r0}:$D${last})"
    lin = f"LINEST({y},{xs},TRUE,TRUE)"
    ws["I4"], ws["I4"].font = "Estimated on 1991Q1 to 2025Q4", BOLD
    header(ws, 5, ["", "Coefficient", "Std error", "t-stat"], col=9, height=30)
    for k, (label, col, nm) in enumerate((("Constant", 3, "Beta0"), ("Unemployment", 2, "BetaUR"),
                                           ("House price change", 1, "BetaHPA")), 6):
        ws.cell(k, 9, label)
        ws.cell(k, 10, f"=INDEX({lin},1,{col})").number_format = NUM4
        ws.cell(k, 11, f"=INDEX({lin},2,{col})").number_format = NUM4
        ws.cell(k, 12, f"=J{k}/K{k}").number_format = "0.0"
        name(wb, nm, f"History!$J${k}")
    ws["I10"], ws["J10"] = "R squared", f"=INDEX({lin},3,1)"
    ws["I11"], ws["J11"] = "Quarters", f"=COUNT({y})"
    ws["I12"], ws["J12"] = "Average charge-off rate, %", f"=AVERAGE($E${r0}:$E${last})"
    labels = [ws.cell(i, 1).value for i in range(first, last + 1)]
    gap0, gap1 = first + labels.index("2008Q1"), first + labels.index("2013Q4")
    before, after = f"$E${r0}:$E${gap0 - 1}", f"$E${gap1 + 1}:$E${last}"
    ws["I13"], ws["J13"] = "Average excluding 2008-2013, %", (
        f"=(SUM({before})+SUM({after}))/(COUNT({before})+COUNT({after}))")
    for c in ("J10", "J12", "J13"):
        ws[c].number_format = NUM4 if c == "J10" else "0.000"
    name(wb, "HistMeanPct", "History!$J$12")
    wb.defined_names["HistMean"] = DefinedName("HistMean", attr_text="History!$J$12/100")
    for col, w in zip("ABCDEFGHIJKL", (9, 12, 10, 12, 11, 13, 10, 2, 30, 11, 10, 8)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B6"

    ch = LineChart()
    ch.title, ch.height, ch.width = "Charge-off rate, actual and fitted, %", 7.5, 16
    for col in (5, 7):
        ch.add_data(Reference(ws, min_col=col, min_row=5, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    ch.x_axis.tickLblSkip = ch.x_axis.tickMarkSkip = 20
    show_axes(ch)
    ws.add_chart(ch, "I16")
    return r0, last


def scenario_sheet(wb, hist, scen):
    ws = wb.create_sheet("Scenarios")
    ws["A1"], ws["A1"].font = "Federal Reserve 2026 supervisory scenarios", TITLE
    ws["A2"] = ("Unemployment and house prices from the Fed's final 2026 scenarios. The first twelve rows are history, "
                "so the 12-quarter house price change can start in 2026Q1.")
    header(ws, 5, ["Quarter", "Baseline unemployment", "Baseline house prices", "Baseline 12q change, %",
                   "Baseline loss rate, %", "Severely adverse unemployment", "Severely adverse house prices",
                   "Severely adverse 12q change, %", "Severely adverse loss rate, %"], height=45)
    pre = hist.loc["2023Q1":"2025Q4"]
    rows = [(str(q), r.ur, r.hpi, r.ur, r.hpi, True) for q, r in pre.iterrows()]
    bl, sa = scen["Baseline"], scen["Severely adverse"]
    rows += [(str(q), bl.loc[q, "ur"], bl.loc[q, "hpi"], sa.loc[q, "ur"], sa.loc[q, "hpi"], False) for q in bl.index]
    first = 6
    for i, (q, ub, hb, us, hs, is_hist) in enumerate(rows, first):
        ws.cell(i, 1, q).font = GREY if is_hist else Font()
        for col, v in ((2, ub), (3, hb), (6, us), (7, hs)):
            c = ws.cell(i, col, float(v))
            c.number_format = "0.0"
            if is_hist:
                c.font = GREY
        if not is_hist:
            for hpa_col, hpi_col, ur_col, rate_col in (("D", "C", "B", "E"), ("H", "G", "F", "I")):
                ws[f"{hpa_col}{i}"] = f"=({hpi_col}{i}/{hpi_col}{i - 12}-1)*100"
                ws[f"{rate_col}{i}"] = f"=MAX(EXP(Beta0+BetaUR*{ur_col}{i}+BetaHPA*{hpa_col}{i})-Eps,0)"
                ws[f"{hpa_col}{i}"].number_format = "0.00"
                ws[f"{rate_col}{i}"].number_format = "0.000"
    s0, s1 = first + 12, first + len(rows) - 1
    name(wb, "RateBL", f"Scenarios!$E${s0}:$E${s1}")
    name(wb, "RateSA", f"Scenarios!$I${s0}:$I${s1}")
    ws["K5"], ws["K5"].font = "Nine quarters, 2026Q1 to 2028Q1", BOLD
    ws["K6"], ws["L6"] = "Baseline cumulative loss", f"=SUM(E{s0}:E{s0 + 8})/400"
    ws["K7"], ws["L7"] = "Severely adverse cumulative loss", f"=SUM(I{s0}:I{s0 + 8})/400"
    ws["K8"], ws["L8"] = "Fed's own projection, severely adverse", FED_9Q_LOSS
    ws["K9"] = "(2026 stress test, first-lien mortgages, 32 banks)"
    ws["K9"].font = GREY
    for c in ("L6", "L7", "L8"):
        ws[c].number_format = PCT
    name(wb, "NineQBL", "Scenarios!$L$6")
    name(wb, "NineQSA", "Scenarios!$L$7")
    for col, w in zip("ABCDEFGHIJKL", (9, 12, 11, 11, 10, 12, 11, 11, 10, 2, 36, 9)):
        ws.column_dimensions[col].width = w


def projection_sheet(wb):
    ws = wb.create_sheet("Projection")
    ws["A1"], ws["A1"].font = "The book quarter by quarter under each scenario", TITLE
    ws["A2"] = ("Loss rates are annualised, in %. Each quarter loses its beginning balance times the rate / 4, then pays "
                "scheduled principal and prepays on what is left. Amounts in $m.")
    first = 6
    ws.cell(5, 1, "Quarter").font = BOLD
    ws.cell(5, 1).fill = HEAD
    for i in range(first, first + MAX_Q):
        ws.cell(i, 1, i - first + 1)
    blocks = {}
    for b, (label, tag) in enumerate(SCEN):
        c0 = 2 + b * 8
        cols = [chr(64 + c0 + k) for k in range(7)]
        rate, beg, loss, perf, sched, prep, end = cols
        blocks[tag] = cols
        ws.cell(4, c0, label).font = BOLD
        header(ws, 5, ["Loss rate, %", "Begin", "Loss", "Performing", "Scheduled", "Prepaid", "End"], col=c0)
        for i in range(first, first + MAX_Q):
            t = f"$A{i}"
            path = (f"IF({t}<=RSUsed,INDEX(Rate{tag},{t}),IF({t}<=RSUsed+RevUsed,"
                    f"INDEX(Rate{tag},RSUsed)+(LongRun*100-INDEX(Rate{tag},RSUsed))*({t}-RSUsed)/RevUsed,LongRun*100))")
            ws[f"{rate}{i}"] = f"=IF({t}>RemQ,0,{path})"
            ws[f"{beg}{i}"] = "=Balance" if i == first else f"={end}{i - 1}"
            ws[f"{loss}{i}"] = f"={beg}{i}*{rate}{i}/400"
            ws[f"{perf}{i}"] = f"={beg}{i}-{loss}{i}"
            ws[f"{sched}{i}"] = f"=IF({t}>RemQ,0,-PMT(NoteRate/4,RemQ-{t}+1,{perf}{i})-{perf}{i}*NoteRate/4)"
            ws[f"{prep}{i}"] = f"=({perf}{i}-{sched}{i})*SMM"
            ws[f"{end}{i}"] = f"={perf}{i}-{sched}{i}-{prep}{i}"
            ws[f"{rate}{i}"].number_format = "0.000"
            for c in (beg, loss, perf, sched, prep, end):
                ws[f"{c}{i}"].number_format = USD
        for c, w in zip(cols, (9, 10, 8, 10, 9, 9, 10)):
            ws.column_dimensions[c].width = w
        ws.column_dimensions[chr(64 + c0 + 7)].width = 2
        last = first + MAX_Q - 1
        name(wb, f"ACL{tag}", f"SUM(Projection!${loss}${first}:${loss}${last})")
        name(wb, f"LossEarly{tag}", f"SUM(Projection!${loss}${first}:${loss}${first + 11})")
    ws.freeze_panes = "B6"
    return blocks, first


def allowance_sheet(wb, blocks, pfirst):
    ws = wb.create_sheet("Allowance")
    ws["A1"], ws["A1"].font = "Allowance for credit losses", TITLE
    ws["A2"] = "Lifetime expected losses on the book, undiscounted, under each scenario and weighted."
    header(ws, 4, ["", "Baseline", "Severely adverse", "Weighted"])
    rows = [("Allowance, $m", "=ACLBL", "=ACLSA", "=WeightBL*ACLBL+WeightSA*ACLSA", USD),
            ("Allowance, % of the book", "=B5/Balance", "=C5/Balance", "=D5/Balance", PCT),
            ("Losses in the first 12 quarters, $m", "=LossEarlyBL", "=LossEarlySA",
             "=WeightBL*LossEarlyBL+WeightSA*LossEarlySA", USD),
            ("Share from after quarter 12", "=1-B7/B5", "=1-C7/C5", "=1-D7/D5", "0%"),
            ("Nine-quarter loss rate, 2026Q1 to 2028Q1", "=NineQBL", "=NineQSA", None, PCT)]
    for i, (label, *cells, fmt) in enumerate(rows, 5):
        ws.cell(i, 1, label)
        for j, f in enumerate(cells, 2):
            if f:
                ws.cell(i, j, f).number_format = fmt
    name(wb, "ACLWeighted", "Allowance!$D$5")
    ws["A11"], ws["A11"].font = "Fed's own severely adverse projection, first-lien mortgages", GREY
    ws["C11"] = "=Scenarios!L8"
    ws["C11"].number_format = PCT

    ws["A13"], ws["A13"].font = "Weighted allowance, $m: forecast period down, reversion period across", BOLD
    ws["A14"] = "=ACLWeighted"
    ws["A14"].number_format = USD
    ws["A14"].font = GREY
    for j, rev in enumerate((1, 4, 8, 12), 2):
        ws.cell(14, j, rev).font = BOLD
    for i, rs in enumerate((4, 6, 8, 10, 12), 15):
        ws.cell(i, 1, rs).font = BOLD
        for j in range(2, 6):
            ws.cell(i, j).number_format = USD
    ws["A21"], ws["A21"].font = "Allowance, $m, against the long-run loss rate", BOLD
    header(ws, 22, ["Long-run rate", "Baseline", "Severely adverse", "Weighted"])
    ws["B23"], ws["C23"], ws["D23"] = "=ACLBL", "=ACLSA", "=ACLWeighted"
    for c in ("B23", "C23", "D23"):
        ws[c].font, ws[c].number_format = GREY, USD
    for i, lr in enumerate((0.0005, 0.001, 0.002, 0.003, 0.004, 0.005), 24):
        ws.cell(i, 1, lr).number_format = PCT3
        for j in range(2, 5):
            ws.cell(i, j).number_format = USD

    # Data tables can only feed cells on their own sheet. These stay empty and
    # the model uses them in place of the Inputs values while a table runs.
    ws["G4"], ws["G4"].font = "Table drivers, leave empty", BOLD
    drivers = (("Forecast period driver", "RSUsed", "RSInput", "H5"),
               ("Reversion period driver", "RevUsed", "RevInput", "H6"),
               ("Long-run rate driver", None, None, "H7"))
    for k, (label, used, base, cell) in enumerate(drivers, 5):
        ws.cell(k, 7, label)
        ws[cell].fill = INPUT
        if used:
            name(wb, used, f"IF(ISBLANK(Allowance!${cell[0]}${cell[1:]}),{base},Allowance!${cell[0]}${cell[1:]})")
    for col, w in zip("ABCDEFGH", (44, 13, 15, 13, 13, 2, 26, 10)):
        ws.column_dimensions[col].width = w

    ch = LineChart()
    ch.title, ch.height, ch.width = "Loss rate path, first 40 quarters, %", 7.5, 15
    ch.y_axis.number_format = "0.0"
    pws = wb["Projection"]
    for tag, label in (("BL", "Baseline"), ("SA", "Severely adverse")):
        col = ord(blocks[tag][0]) - 64
        ch.series.append(Series(Reference(pws, min_col=col, min_row=pfirst, max_row=pfirst + 39), title=label))
    ch.set_categories(Reference(pws, min_col=1, min_row=pfirst, max_row=pfirst + 39))
    ch.x_axis.tickLblSkip = ch.x_axis.tickMarkSkip = 4
    show_axes(ch)
    ch.legend.position = "r"
    ws.add_chart(ch, "G13")
    return (14, 19), (23, 29)


def notes_sheet(wb):
    ws = wb.create_sheet("Notes")
    lines = [
        ("How the allowance is built", True),
        ("CECL (ASC 326) books lifetime expected credit losses on day one. This workbook follows the common "
         "top-down route: a loss-rate model driven by economic forecasts for a reasonable and supportable period, a "
         "straight-line reversion, and the historical average loss rate for the rest of the book's life.", False),
        ("The loss-rate model is estimated on the History sheet with LINEST, so it re-fits if the data change. The "
         "log form, ln(max(rate, 0) + 0.02), lets losses stay near zero when owners have equity and rise steeply when "
         "prices fall.", False),
        ("Scenarios are the Federal Reserve's final 2026 supervisory baseline and severely adverse paths. Weights are "
         "an input; a bank would usually add its own upside and downside cases.", False),
        ("Limits", True),
        ("The model is fitted to the whole banking system's charge-offs, which in 2008-2011 included far weaker "
         "loans than banks hold today, so it overstates losses for a prime book. Losses are undiscounted and taken as "
         "charge-offs, with no separate delinquency or recovery timing.", False),
        ("Checks", True),
        ("cecl.py rebuilds the regression, the loss paths and the book in Python, and tests/test_workbook.py "
         "compares them with the values Excel calculated.", False),
    ]
    for i, (text, bold) in enumerate(lines, 1):
        c = ws.cell(i, 1, text)
        c.font = BOLD if bold else Font()
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 110


def show_axes(chart):
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "b"


def build():
    hist, scen = cecl.load()
    wb = Workbook()
    inputs_sheet(wb)
    history_sheet(wb, hist)
    scenario_sheet(wb, hist, scen)
    blocks, pfirst = projection_sheet(wb)
    tables = allowance_sheet(wb, blocks, pfirst)
    notes_sheet(wb)
    wb.move_sheet("Allowance", offset=-4)
    wb.save(OUT)
    return tables


def finish_in_excel(tables):
    import pythoncom
    import win32com.client as win32
    (g0, g1), (l0, l1) = tables
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible, xl.DisplayAlerts = False, False
    try:
        wb = xl.Workbooks.Open(str(OUT.resolve()))
        xl.Calculation = -4135
        ws = wb.Worksheets("Allowance")
        ws.Range(f"A{g0}:E{g1}").Table(ws.Range("H6"), ws.Range("H5"))
        ws.Range(f"A{l0}:D{l1}").Table(pythoncom.Empty, ws.Range("H7"))
        xl.CalculateFull()
        wait_until_calculated(xl)
        xl.Calculation = -4105
        wait_until_calculated(xl)
        figs = ROOT / "figures"
        figs.mkdir(exist_ok=True)
        # Excel's own charts, exported directly. Unlike a picture of a range,
        # this needs neither a printer nor a visible screen.
        for sheet, out in (("History", "chart_fit.png"), ("Allowance", "chart_loss_path.png")):
            chart = wb.Worksheets(sheet).ChartObjects(1).Chart
            chart.Export(str((figs / out).resolve()))
        wb.Worksheets("Allowance").Activate()
        wb.Save()
        wb.Close()
    finally:
        try:
            xl.Quit()
        except Exception:
            pass


def wait_until_calculated(xl, timeout=600):
    import pywintypes
    start = time.time()
    while time.time() - start < timeout:
        try:
            if xl.CalculationState == 0:
                return
        except pywintypes.com_error:
            pass
        time.sleep(0.5)
    raise TimeoutError("Excel did not finish calculating")


if __name__ == "__main__":
    tables = build()
    finish_in_excel(tables)
    print("wrote", OUT)
