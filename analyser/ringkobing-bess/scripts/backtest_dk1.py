"""Perfect-foresight daily LP backtest for a 1 MW BESS in DK1 using Energinet data."""
import json, sys
import numpy as np, pandas as pd
from scipy.optimize import linprog
from scipy.sparse import lil_matrix

D = sys.argv[1]
def load(f): return json.load(open(f"{D}/{f}"))["records"]

# --- prices (EUR) ---
h = pd.DataFrame(load("spot_h.json")); h["t"] = pd.to_datetime(h.HourDK); h = h.set_index("t")["SpotPriceEUR"]
q = pd.DataFrame(load("da_15.json")); q["t"] = pd.to_datetime(q.TimeDK); q = q.set_index("t")["DayAheadPriceEUR"]
af = pd.DataFrame(load("afrr.json")); af["t"] = pd.to_datetime(af.TimeDK); af = af.set_index("t")
fc = pd.DataFrame(load("fcr.json")); fc["t"] = pd.to_datetime(fc.HourDK); fc = fc.set_index("t").sort_index()
mf = pd.DataFrame(load("mfrr.json")); mf["t"] = pd.to_datetime(mf.TimeDK); mf = mf.set_index("t")
afu = af.UpPriceEUR.groupby(level=0).mean(); afd = af.DownPriceEUR.groupby(level=0).mean()
fcp = fc.FCRdk_EUR.groupby(level=0).mean(); mfu = mf.UpPriceEUR.groupby(level=0).mean()

P, E = 1.0, float(sys.argv[2]) if len(sys.argv) > 2 else 2.0
RTE = 0.86; ec = ed = RTE ** 0.5
NCYC = 1.5
TC, TD = float(sys.argv[3]) if len(sys.argv) > 3 else 18.6, 1.9   # EUR/MWh grid cost on charge / discharge
TAU_A, TAU_F = 1.0, 0.5

def solve_day(price, dt, pu, pd_, pf, pm, use_res):
    n = len(price); nv = 7 * n  # c,d,ru,rd,f,m,s
    ic, id_, iu, idn, iff, im, isx = [np.arange(k * n, (k + 1) * n) for k in range(7)]
    cost = np.zeros(nv)
    cost[ic] = (price + TC) * dt; cost[id_] = -(price - TD) * dt
    if use_res:
        cost[iu] = -pu * dt; cost[idn] = -pd_ * dt; cost[iff] = -pf * dt; cost[im] = -pm * dt
    bounds = [(0, P)] * (6 * n) + [(0, E)] * n
    if not use_res:
        for k in list(iu) + list(idn) + list(iff) + list(im): bounds[k] = (0, 0)
    rows = []; A = lil_matrix((6 * n + 1, nv)); b = []
    r = 0
    for t in range(n):
        A[r, id_[t]] = 1; A[r, iu[t]] = 1; A[r, iff[t]] = 1; A[r, im[t]] = 1; b.append(P); r += 1
        A[r, ic[t]] = 1; A[r, idn[t]] = 1; A[r, iff[t]] = 1; b.append(P); r += 1
        # SoC must cover up-reserves before and after step
        A[r, iu[t]] = TAU_A / ed; A[r, im[t]] = TAU_A / ed; A[r, iff[t]] = TAU_F / ed; A[r, isx[t]] = -1; b.append(0); r += 1
        A[r, idn[t]] = TAU_A * ec; A[r, iff[t]] = TAU_F * ec; A[r, isx[t]] = 1; b.append(E); r += 1
        if t > 0:
            A[r, iu[t]] = TAU_A / ed; A[r, im[t]] = TAU_A / ed; A[r, iff[t]] = TAU_F / ed; A[r, isx[t-1]] = -1; b.append(0); r += 1
            A[r, idn[t]] = TAU_A * ec; A[r, iff[t]] = TAU_F * ec; A[r, isx[t-1]] = 1; b.append(E); r += 1
        else:
            A[r, iu[t]] = TAU_A / ed; A[r, im[t]] = TAU_A / ed; A[r, iff[t]] = TAU_F / ed; b.append(E / 2); r += 1
            A[r, idn[t]] = TAU_A * ec; A[r, iff[t]] = TAU_F * ec; b.append(E / 2); r += 1
    for k in id_: A[r, k] = dt
    b.append(NCYC * E); r += 1
    A = A[:r].tocsr(); b = np.array(b[:r])
    # SoC dynamics
    Aeq = lil_matrix((n + 4 * n, nv)); beq = []; e = 0
    for t in range(n):
        Aeq[e, isx[t]] = 1; Aeq[e, ic[t]] = -ec * dt; Aeq[e, id_[t]] = dt / ed
        if t > 0: Aeq[e, isx[t-1]] = -1; beq.append(0)
        else: beq.append(E / 2)
        e += 1
    Aeq[e, isx[n-1]] = 1; beq.append(E / 2); e += 1
    # reserves constant within the hour (15-min days)
    if dt < 1:
        spq = int(round(1 / dt))
        for t in range(n):
            if t % spq:
                for arr in (iu, idn, iff, im):
                    Aeq[e, arr[t]] = 1; Aeq[e, arr[t-1]] = -1; beq.append(0); e += 1
    Aeq = Aeq[:e].tocsr(); beq = np.array(beq)
    res = linprog(cost, A_ub=A, b_ub=b, A_eq=Aeq, b_eq=beq, bounds=bounds, method="highs")
    x = res.x
    out = dict(
        arb=float(((price - 0) * (x[id_] - x[ic]) * dt).sum()),
        grid=float((-(TC * x[ic] + TD * x[id_]) * dt).sum()),
        afrr=float(((pu * x[iu] + pd_ * x[idn]) * dt).sum()),
        fcr=float((pf * x[iff] * dt).sum()), mfrr=float((pm * x[im] * dt).sum()),
        mwh_out=float((x[id_] * dt).sum()))
    return out

def run(start, end, use_res):
    days = pd.date_range(start, end, freq="D", inclusive="left")
    tot = []
    for d in days:
        d1 = d + pd.Timedelta(days=1)
        s = q[(q.index >= d) & (q.index < d1)]
        if len(s) >= 92: dt = 0.25
        else:
            s = h[(h.index >= d) & (h.index < d1)]; dt = 1.0
            if len(s) < 23: continue
        hrs = s.index.floor("h")
        pu = afu.reindex(hrs).fillna(0).values; pdn = afd.reindex(hrs).fillna(0).values
        pf = fcp.reindex(hrs).fillna(0).values; pm = mfu.reindex(hrs).fillna(0).values
        o = solve_day(s.values.astype(float), dt, pu, pdn, pf, pm, use_res); o["day"] = d; tot.append(o)
    return pd.DataFrame(tot)

if __name__ == "__main__":
    periods = {"2024": ("2024-01-01", "2025-01-01"), "2025": ("2025-01-01", "2026-01-01"),
               "LTM": ("2025-10-13", "2026-10-13"), "2026YTD": ("2026-01-01", "2026-10-13")}
    rows = []
    for name, (a, z) in periods.items():
        for use_res in (False, True):
            df = run(a, z, use_res); ndays = len(df); f = 365.0 / ndays
            r = df[["arb", "grid", "afrr", "fcr", "mfrr", "mwh_out"]].sum() * f
            r["net"] = r.arb + r.grid + r.afrr + r.fcr + r.mfrr
            r["period"] = name; r["reserves"] = use_res; r["days"] = ndays; rows.append(r)
            print(name, "res" if use_res else "DA", ndays, {k: round(v) for k, v in r.items() if isinstance(v, float)}, flush=True)
    pd.DataFrame(rows).to_csv(f"{D}/backtest_E{E}_TC{TC}.csv", index=False)
