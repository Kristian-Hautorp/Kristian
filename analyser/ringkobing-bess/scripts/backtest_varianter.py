"""Varianter af backtesten: batteriets varighed (1h/4h) og mætning af systemydelser (reservepriser x1, x0,5, x0,25).

Brug: python3 backtest_varianter.py <sti til backtest_dk1.py> <datamappe>
"""
import importlib.util, sys

spec = importlib.util.spec_from_file_location("bt", sys.argv[1]); bt = importlib.util.module_from_spec(spec)
sys.argv = [sys.argv[1], sys.argv[2], "2.0", "18.6"]; spec.loader.exec_module(bt)
LTM = ("2025-10-13", "2026-10-13")

for E in (1.0, 4.0):
    bt.E = E
    df = bt.run(*LTM, False); f = 365 / len(df)
    r = df[["arb", "grid", "mwh_out"]].sum() * f
    print("varighed", E, {k: round(v) for k, v in r.items()}, "netto", round(r.arb + r.grid), flush=True)

bt.E = 2.0
fcp0, afu0, afd0, mfu0 = bt.fcp.copy(), bt.afu.copy(), bt.afd.copy(), bt.mfu.copy()
bt.fcp = fcp0 * 0  # FCR udelades: DK1-volumen fra danske enheder er kun ca. 2-10 MW
for k in (1.0, 0.5, 0.25):
    bt.afu, bt.afd, bt.mfu = afu0 * k, afd0 * k, mfu0 * k
    df = bt.run(*LTM, True); f = 365 / len(df)
    r = df[["arb", "grid", "afrr", "fcr", "mfrr", "mwh_out"]].sum() * f
    print("reservepris x", k, {kk: round(v) for kk, v in r.items()}, "netto", round(r.arb + r.grid + r.afrr + r.mfrr), flush=True)
