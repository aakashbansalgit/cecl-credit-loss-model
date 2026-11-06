"""
The CECL allowance in Python, written independently of the workbook.

    python cecl.py

1. Loss-rate model. The quarterly net charge-off rate on single-family
   mortgages at US banks (FRED CORSFRMACBS, annualised, 1991 to 2025) is
   regressed on the unemployment rate and the 12-quarter change in house
   prices, both from the Federal Reserve's historical stress-test file:

       ln(max(charge-off rate, 0) + 0.02) = a + b * unemployment + c * house price change

   The log form lets losses stay near zero when owners have equity and rise
   steeply when prices fall, which a straight line cannot do.

2. Loss-rate path. For the reasonable and supportable period, eight quarters
   by default, the model is driven by one of the Fed's 2026 scenarios. Over
   the next four quarters the rate moves in a straight line to its long-run
   average, and stays there for the rest of the portfolio's life.

3. Allowance. A mortgage book amortises and prepays quarter by quarter; each
   quarter loses its beginning balance times the loss rate / 4. The allowance
   is the sum of those losses (undiscounted), weighted across scenarios.
"""
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).parent
DATA = ROOT / "data"
EPS = 0.02
HPA_QUARTERS = 12
FIRST, LAST = "1991Q1", "2025Q4"


def _quarters(df):
    q = pd.PeriodIndex(df["Date"].str.replace(" ", ""), freq="Q")
    return pd.DataFrame({"ur": df["Unemployment rate"].values,
                         "hpi": df["House Price Index (Level)"].values}, index=q)


def load():
    hist = _quarters(pd.read_csv(DATA / "fed_2026_historic.csv"))
    co = pd.read_csv(DATA / "fred_chargeoffs.csv", index_col=0, parse_dates=True).iloc[:, 0]
    co.index = co.index.to_period("Q")
    hist["co"] = co.reindex(hist.index)
    hist["hpa"] = (hist.hpi / hist.hpi.shift(HPA_QUARTERS) - 1) * 100
    scen = {}
    for name, f in (("Baseline", "fed_2026_baseline.csv"), ("Severely adverse", "fed_2026_severely_adverse.csv")):
        s = pd.concat([hist[["ur", "hpi"]], _quarters(pd.read_csv(DATA / f))])
        s["hpa"] = (s.hpi / s.hpi.shift(HPA_QUARTERS) - 1) * 100
        scen[name] = s.loc["2026Q1":]
    return hist, scen


def fit(hist):
    d = hist.loc[FIRST:LAST].dropna(subset=["co", "hpa"])
    y = np.log(d.co.clip(lower=0) + EPS).to_numpy()
    X = np.column_stack([np.ones(len(d)), d.ur, d.hpa])
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    return dict(zip(("const", "ur", "hpa"), beta)), d


def predict(beta, ur, hpa):
    return np.maximum(np.exp(beta["const"] + beta["ur"] * ur + beta["hpa"] * hpa) - EPS, 0.0)


@dataclass
class Book:
    balance: float = 1000.0         # $m
    note_rate: float = 0.065
    remaining_quarters: int = 112
    cpr: float = 0.08
    rs_quarters: int = 8            # reasonable and supportable period
    reversion_quarters: int = 4
    long_run: float = None          # annualised %, None for the 1991-2025 average
    weights: dict = field(default_factory=lambda: {"Baseline": 0.6, "Severely adverse": 0.4})


def loss_path(beta, scen, long_run, book):
    """Annualised loss rate for each quarter of the book's remaining life."""
    n = book.remaining_quarters
    s = scen.iloc[:book.rs_quarters]
    rates = list(predict(beta, s.ur.to_numpy(), s.hpa.to_numpy()))
    last = rates[-1]
    for k in range(1, book.reversion_quarters + 1):
        rates.append(last + (long_run - last) * k / book.reversion_quarters)
    rates += [long_run] * max(0, n - len(rates))
    return np.array(rates[:n])


def run_book(book, rates):
    q = book.note_rate / 4
    smm = 1 - (1 - book.cpr) ** 0.25
    bal = book.balance
    out = {"begin": [], "loss": [], "scheduled": [], "prepaid": []}
    for t, rate in enumerate(rates, 1):
        loss = bal * rate / 100 / 4
        perf = bal - loss
        remaining = book.remaining_quarters - t + 1
        payment = perf * q / (1 - (1 + q) ** -remaining)
        sched = payment - perf * q
        prepaid = (perf - sched) * smm
        for k, v in (("begin", bal), ("loss", loss), ("scheduled", sched), ("prepaid", prepaid)):
            out[k].append(v)
        bal = perf - sched - prepaid
    return {k: np.array(v) for k, v in out.items()}


def allowance(book=None):
    book = book or Book()
    hist, scen = load()
    beta, used = fit(hist)
    long_run = float(used.co.mean()) if book.long_run is None else book.long_run
    res = {"beta": beta, "long_run": long_run, "scenarios": {}}
    for name, s in scen.items():
        rates = loss_path(beta, s, long_run, book)
        cf = run_book(book, rates)
        res["scenarios"][name] = {
            "rates": rates, "cashflows": cf, "acl": float(cf["loss"].sum()),
            "nine_quarter_loss_pct": float(predict(beta, s.ur.to_numpy()[:9], s.hpa.to_numpy()[:9]).sum() / 4),
        }
    res["weighted_acl"] = sum(book.weights[k] * v["acl"] for k, v in res["scenarios"].items())
    return res


if __name__ == "__main__":
    r = allowance()
    print("coefficients", {k: round(v, 4) for k, v in r["beta"].items()}, "long-run rate", round(r["long_run"], 4))
    for k, v in r["scenarios"].items():
        print(f"{k}: allowance ${v['acl']:.2f}m, nine-quarter loss {v['nine_quarter_loss_pct']:.2f}%,"
              f" peak rate {v['rates'].max():.2f}%")
    print(f"weighted allowance ${r['weighted_acl']:.2f}m on a $1,000m book")
