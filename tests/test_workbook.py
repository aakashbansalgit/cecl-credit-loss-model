"""Check the Excel workbook against cecl.py.

openpyxl reads the values Excel saved with the file, so this runs without Excel.

    python -m pytest tests
"""
import sys
from pathlib import Path

import numpy as np
import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import cecl  # noqa: E402

WB = load_workbook(ROOT / "cecl_allowance.xlsx", data_only=True)
HIST, SCEN = cecl.load()
BETA, USED = cecl.fit(HIST)


def named(name):
    sheet, cell = next(iter(WB.defined_names[name].destinations))
    return WB[sheet][cell.replace("$", "")].value


def book(**override):
    b = cecl.Book(balance=named("Balance"), note_rate=named("NoteRate"), remaining_quarters=int(named("RemQ")),
                  cpr=named("CPR"), rs_quarters=int(named("RSInput")), reversion_quarters=int(named("RevInput")),
                  weights={"Baseline": named("WeightBL"), "Severely adverse": named("WeightSA")})
    for k, v in override.items():
        setattr(b, k, v)
    return b


def test_linest_matches_least_squares():
    ws = WB["History"]
    for row, key in ((6, "const"), (7, "ur"), (8, "hpa")):
        assert ws.cell(row, 10).value == pytest.approx(BETA[key], rel=1e-10)
    assert ws["J11"].value == len(USED) == 140
    assert ws["J12"].value == pytest.approx(USED.co.mean(), rel=1e-12)
    calm = USED.co.drop(USED.loc["2008Q1":"2013Q4"].index)
    assert ws["J13"].value == pytest.approx(calm.mean(), rel=1e-12)


def test_scenario_loss_rates():
    ws = WB["Scenarios"]
    for col, name in ((5, "Baseline"), (9, "Severely adverse")):
        s = SCEN[name]
        want = cecl.predict(BETA, s.ur.to_numpy(), s.hpa.to_numpy())
        got = [ws.cell(r, col).value for r in range(18, 18 + len(s))]
        assert np.allclose(got, want, rtol=1e-10, atol=1e-12)
    res = cecl.allowance(book())
    assert ws["L6"].value * 100 == pytest.approx(res["scenarios"]["Baseline"]["nine_quarter_loss_pct"], rel=1e-10)
    assert ws["L7"].value * 100 == pytest.approx(res["scenarios"]["Severely adverse"]["nine_quarter_loss_pct"],
                                                 rel=1e-10)


def test_projection_quarter_by_quarter():
    ws = WB["Projection"]
    b = book()
    res = cecl.allowance(b)
    n = b.remaining_quarters
    for c0, name in ((2, "Baseline"), (10, "Severely adverse")):
        r = res["scenarios"][name]
        cols = {"rate": c0, "begin": c0 + 1, "loss": c0 + 2, "scheduled": c0 + 4, "prepaid": c0 + 5}
        got = {k: np.array([ws.cell(i, c).value for i in range(6, 6 + n)]) for k, c in cols.items()}
        assert np.allclose(got["rate"], r["rates"], rtol=1e-10, atol=1e-12)
        for k in ("begin", "loss", "scheduled", "prepaid"):
            assert np.allclose(got[k], r["cashflows"][k], rtol=1e-10, atol=1e-9), (name, k)
        tail = [ws.cell(i, cols["begin"]).value for i in range(6 + n, 6 + 120)]
        assert max(abs(v) for v in tail) < 1e-6           # the book is gone after its last quarter


def test_allowance_summary():
    ws = WB["Allowance"]
    res = cecl.allowance(book())
    assert ws["B5"].value == pytest.approx(res["scenarios"]["Baseline"]["acl"], rel=1e-10)
    assert ws["C5"].value == pytest.approx(res["scenarios"]["Severely adverse"]["acl"], rel=1e-10)
    assert ws["D5"].value == pytest.approx(res["weighted_acl"], rel=1e-10)


def test_forecast_and_reversion_table():
    ws = WB["Allowance"]
    revs = [ws.cell(14, j).value for j in range(2, 6)]
    for i in range(15, 20):
        rs = ws.cell(i, 1).value
        for j, rev in enumerate(revs, 2):
            want = cecl.allowance(book(rs_quarters=rs, reversion_quarters=rev))["weighted_acl"]
            assert ws.cell(i, j).value == pytest.approx(want, rel=1e-10), (rs, rev)


def test_long_run_rate_table():
    ws = WB["Allowance"]
    for i in range(24, 30):
        lr = ws.cell(i, 1).value
        res = cecl.allowance(book(long_run=lr * 100))
        assert ws.cell(i, 2).value == pytest.approx(res["scenarios"]["Baseline"]["acl"], rel=1e-10)
        assert ws.cell(i, 3).value == pytest.approx(res["scenarios"]["Severely adverse"]["acl"], rel=1e-10)
        assert ws.cell(i, 4).value == pytest.approx(res["weighted_acl"], rel=1e-10)


def test_losses_rise_with_the_long_run_rate_and_the_stress():
    ws = WB["Allowance"]
    baseline = [ws.cell(i, 2).value for i in range(24, 30)]
    assert all(b > a for a, b in zip(baseline, baseline[1:]))
    assert ws["C5"].value > ws["D5"].value > ws["B5"].value
