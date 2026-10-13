"""Bygger Excel-modellen for en tidlig investeringsberegning af et BESS ved Ringkøbing Fjernvarme.

Brug: python3 byg_model.py <datamappe> <output.xlsx>
Datamappen skal indeholde Energi Data Service-filerne (se hent_data.sh) og backtest-CSV'en fra backtest_dk1.py.
"""
import json, sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

DATA, OUT = sys.argv[1], sys.argv[2]
DATO = "13.10.2026"

# ---------- styles ----------
F_TITLE = Font(bold=True, size=14)
F_H = Font(bold=True, color="FFFFFF")
F_B = Font(bold=True)
F_NOTE = Font(italic=True, color="555555", size=9)
F_INPUT = Font(color="1F4E9E")
FILL_H = PatternFill("solid", fgColor="1F3A5F")
FILL_SEC = PatternFill("solid", fgColor="D9E2EF")
FILL_IN = PatternFill("solid", fgColor="FFF6D5")
FILL_KPI = PatternFill("solid", fgColor="E8F3E8")
THIN = Side(style="thin", color="BBBBBB")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
EUR = '#,##0;[Red]-#,##0'
EURK = '#,##0," k";[Red]-#,##0," k"'
PCT = '0.0%;[Red]-0.0%'
NUM2 = '0.00'

def header(ws, row, labels, widths=None):
    for i, t in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAP
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w

wb = Workbook()

# =====================================================================
# Forudsætninger
# =====================================================================
wi = wb.active; wi.title = "Forudsætninger"
wi["A1"] = "Forudsætninger: BESS ved Ringkøbing Fjernvarme (tidlig beregning, til godkendelse)"; wi["A1"].font = F_TITLE
wi["A2"] = (f"Udarbejdet {DATO}. Gule celler er input og kan ændres; scenarie-arkene regner automatisk. "
            "Status: Data = fra kilde, Estimat = min vurdering, Skal verificeres / Ved ikke = skal afklares før beslutning.")
wi["A2"].font = F_NOTE
header(wi, 4, ["Nøgle", "Parameter", "Enhed", "Pessimistisk", "Base", "Optimistisk", "Status", "Kilde / begrundelse"],
       [16, 46, 14, 13, 13, 13, 17, 90])

P = {}  # key -> row
rows = [
    ("SEC", "Projekt"),
    ("mw", "Batteriets effekt", "MW", 10, 10, 10, "Ved ikke",
     "Referencestørrelse. Den reelle størrelse afhænger af ledig import- og eksportkapacitet i RFV's tilslutning, som ikke er kendt."),
    ("dur", "Varighed (1, 2 eller 4 timer)", "timer", 2, 2, 2, "Estimat",
     "Backtesten findes for 1, 2 og 4 timer (se arket Backtest). 4 timer giver mest indtjening pr. MW nettilslutning, men højere capex."),
    ("cod", "Første driftsår (COD)", "år", 2028, 2028, 2028, "Estimat",
     "Antaget: udvikling, netaftale, indkøb og byggeri ca. 15-24 mdr. fra beslutning."),
    ("life", "Økonomisk levetid", "år", 15, 15, 15, "Estimat", "Ingen restværdi og ingen udvidelse (augmentation) antaget. Maks. 20 år i modellen."),
    ("avail", "Tilgængelighed inkl. konflikt med RFV's elkedel og gasmotorer", "%", 0.85, 0.92, 0.97, "Estimat",
     "Batteriet deler tilslutning med elkedel (import) og gasmotorer (eksport), der bruger kapaciteten i de samme timer. Skal afklares med RFV-data."),
    ("fade", "Kapacitetstab pr. år (påvirker arbitrage)", "%/år", 0.025, 0.02, 0.015, "Estimat", "Typisk for LFP ved ca. 1 cyklus/døgn."),
    ("rte", "Round-trip-virkningsgrad (AC-AC)", "%", 0.86, 0.86, 0.86, "Estimat", "Samme værdi som i backtesten."),
    ("SEC", "Anlægsomkostninger"),
    ("ckw", "Effektrelateret capex (PCS, trafo, HV/MV, BoP, EPC)", "EUR/kW", 250, 200, 160, "Estimat",
     "Kalibreret så 2h all-in ligger på 375/300/250 EUR/kWh. Ref.: Modo Energy (16.03.2026) 700k EUR/MW for 2h i Tyskland (scope ikke oplyst); "
     "BNEF Cost Survey 2025 Europa turnkey 177 USD/kWh ekskl. EPC og nettilslutning (via Energy-Storage.news, dec. 2025)."),
    ("ckwh", "Energirelateret capex (batteriblokke inkl. montage)", "EUR/kWh", 250, 200, 170, "Estimat", "Se ovenfor. Skal erstattes af tilbud."),
    ("dev", "Udvikling, rådgivning, netstudier, finansiering", "EUR", 400000, 250000, 150000, "Estimat", "Juridisk, skat (bindende svar), netselskab, teknisk DD."),
    ("conn", "Tilslutningsbidrag / ombygning af tilslutning", "EUR", 300000, 0, 0, "Skal verificeres",
     "Elnet.dk/Green Power Denmark: intet tilslutningsbidrag, hvis batteriet bruger eksisterende kapacitet uden at der søges ny. "
     "Pessimistisk: netselskabet kræver ombygning eller udvidelse."),
    ("SEC", "Indtægter (2026-priser)"),
    ("lvl", "Arbitrage-niveau ift. backtest seneste 12 mdr.", "faktor", 0.836, 1.0, 1.126, "Data",
     "Pess. = kalenderår 2025, Base = okt. 2025 til okt. 2026, Opt. = 2026 år-til-dato (annualiseret). 2h-backtest, se arket Backtest."),
    ("cap", "Realiseringsgrad ift. perfekt forudsigelse (day-ahead)", "%", 0.80, 0.85, 0.90, "Estimat", "Backtesten antager perfekt kendskab til priserne."),
    ("idu", "Tillæg fra intraday/ubalance", "%", 0.0, 0.10, 0.20, "Estimat", "Ikke backtestet. Kan være væsentligt i DK1, men usikkert."),
    ("arbtr", "Real udvikling i arbitrage-indtjening", "%/år", -0.03, -0.01, 0.01, "Estimat",
     "Flere batterier presser spreads; mere sol og vind øger dem. Spreads i 2026 er høje sammen med et højt prisniveau."),
    ("anc1", "Systemydelser (aFRR + mFRR), netto merindtægt år 1", "EUR/MW/år", 15000, 30000, 70000, "Data + estimat",
     "Backtest (2h, seneste 12 mdr.): merindtægt ift. ren arbitrage = 156k ved fulde priser, 70k ved 50 % og 30k ved 25 %. "
     "Priserne i Q3 2026 ligger på ca. 50 % af 12-mdr.-gennemsnittet. Base antager yderligere halvering frem til 2028."),
    ("anctr", "Udvikling i systemydelser", "%/år", -0.15, -0.05, -0.05, "Estimat", "aFRR op i DK1 er faldet fra ca. 70 til ca. 6-10 EUR/MW/h (Q4 2024 til 2026)."),
    ("ancfl", "Gulv for systemydelser", "EUR/MW/år", 0, 10000, 20000, "Estimat", ""),
    ("SEC", "Nettariffer og afgifter (betales af batteriet)"),
    ("t_en", "Energinet forbrugstarif (net + system)", "øre/kWh", 12.5, 10.5, 10.5, "Data",
     "Energinet-pressemeddelelse 04.09.2026: 10,5 øre/kWh i 2027 og loft på 12,5 øre/kWh frem til 2030."),
    ("t_dso", "Netselskabets forbrugstarif (gns. i ladetimer)", "øre/kWh", 5.0, 3.0, 2.0, "Skal verificeres",
     "Netselskab og tarif for RFV's tilslutning er ikke kendt. Elnet.dk: batterier betaler både forbrugs- og indfødningstarif."),
    ("t_afg", "Elafgift på opladning efter godtgørelse", "øre/kWh", 0.4, 0.4, 0.4, "Skal verificeres",
     "KRITISK. Lov nr. 1775 af 29.12.2025: 0,8 øre/kWh i 2026-27, derefter 60,9 øre (2015-niveau). Antaget: momsreg. virksomhed får "
     "godtgørelse ned til 0,4 øre. Gælder det ikke for et batteri fra 2028, falder business casen. Kræver bindende svar."),
    ("t_enp", "Energinet indfødnings- og balancetarif", "øre/kWh", 0.92, 0.92, 0.92, "Data", "Energinet-pressemeddelelse 04.09.2026: 0,5 + 0,42 øre/kWh (2027)."),
    ("t_dsop", "Netselskabets indfødningstarif", "øre/kWh", 0.5, 0.5, 0.5, "Skal verificeres", "Estimat."),
    ("fx", "Valutakurs", "DKK/EUR", 7.46, 7.46, 7.46, "Data", "Fast kurs (ERM2)."),
    ("SEC", "Driftsomkostninger (2026-priser)"),
    ("om", "Drift og vedligehold (LTSA)", "% af anlægscapex", 0.02, 0.015, 0.012, "Estimat", ""),
    ("ins", "Forsikring", "% af anlægscapex", 0.005, 0.004, 0.003, "Estimat", ""),
    ("fee", "Optimerings-/aggregatorfee", "% af markedsindtægt", 0.15, 0.10, 0.07, "Estimat", "Alternativ: fast fee pr. MW eller gevinstdeling."),
    ("rfv", "Betaling til RFV for netadgang og areal", "EUR/MW/år", 10000, 6000, 3000, "Ved ikke",
     "Til forhandling. Kan også erstattes af ejerandel til RFV."),
    ("fix", "Faste omkostninger (adm., revision, BRP, SCADA, abonnementer)", "EUR/år", 60000, 40000, 30000, "Estimat", ""),
    ("SEC", "Økonomi og skat"),
    ("infl", "Inflation (indeksering fra 2026-niveau)", "%/år", 0.02, 0.02, 0.02, "Estimat", ""),
    ("tax", "Selskabsskat", "%", 0.22, 0.22, 0.22, "Data", "Dansk selskabsskat."),
    ("saldo", "Skattemæssig afskrivning (saldo)", "%/år", 0.25, 0.25, 0.25, "Skal verificeres",
     "Antaget almindelige driftsmidler (afskrivningsloven). Underskud fremføres; 60 %-begrænsningen er ikke modelleret."),
    ("disc", "Diskonteringsrente til NPV (nominel)", "%", 0.08, 0.08, 0.08, "Estimat", "Indikativt afkastkrav for merchant-BESS."),
]
r = 5
for row in rows:
    if row[0] == "SEC":
        c = wi.cell(row=r, column=1, value=row[1]); c.font = F_B
        for col in range(1, 9): wi.cell(row=r, column=col).fill = FILL_SEC
    else:
        key, name, unit, lo, base, hi, status, src = row
        P[key] = r
        vals = [key, name, unit, lo, base, hi, status, src]
        for col, v in enumerate(vals, 1):
            c = wi.cell(row=r, column=col, value=v); c.border = BOX; c.alignment = WRAP
        fmt = PCT if unit.startswith("%") else (NUM2 if unit in ("faktor", "øre/kWh", "DKK/EUR") else EUR)
        if unit in ("år", "timer", "MW"): fmt = "0"
        for col in (4, 5, 6):
            c = wi.cell(row=r, column=col); c.number_format = fmt; c.fill = FILL_IN; c.font = F_INPUT
        if status in ("Skal verificeres", "Ved ikke"): wi.cell(row=r, column=7).font = Font(bold=True, color="B00020")
    r += 1
wi.freeze_panes = "D5"

# =====================================================================
# Backtest (statisk, beregnet fra Energinet-data)
# =====================================================================
wb_ = wb.create_sheet("Backtest")
bt = pd.read_csv(f"{DATA}/backtest_E2.0_TC18.6.csv")
wb_["A1"] = "Backtest af batteri i DK1 på Energinet-data (perfekt forudsigelse)"; wb_["A1"].font = F_TITLE
notes = [
    f"Beregnet {DATO} med scripts/backtest_dk1.py på data fra Energinets Energi Data Service (Elspotprices, DayAheadPrices, AfrrReservesNordic, FcrDK1, mFRRCapacityMarket).",
    "Metode: lineær optimering pr. døgn for 1 MW. Round-trip 86 %, maks. 1,5 cyklus/døgn, SoC 50 % ved døgnstart og -slut. Reserver kræver energi til 60 min (aFRR/mFRR) og 30 min (FCR).",
    "Netomkostning i optimeringen: 18,6 EUR/MWh på opladning og 1,9 EUR/MWh på afladning (svarer til base-tarifferne). Kvartersopløsning fra 1.10.2025, timeopløsning før.",
    "Perfekt forudsigelse overvurderer indtjeningen; det korrigeres med realiseringsgraden i Forudsætninger. Ubegrænset reservevolumen er urealistisk i DK1 (lille marked), se mætningsvarianter.",
]
for i, n in enumerate(notes): wb_.cell(row=2 + i, column=1, value=n).font = F_NOTE
r = 7
wb_.cell(row=r, column=1, value="A. Indtjening pr. MW pr. år, 2h-batteri (EUR, annualiseret)").font = F_B; r += 1
header(wb_, r, ["Periode", "Strategi", "Arbitrage brutto", "Netomkostning", "aFRR", "FCR", "mFRR", "Netto", "MWh afladet", "Dage"],
       [26, 34, 16, 16, 14, 14, 14, 14, 14, 10]); r += 1
names = {"2024": "Kalenderår 2024", "2025": "Kalenderår 2025", "LTM": "13.10.2025 til 12.10.2026", "2026YTD": "2026 år-til-dato"}
for _, x in bt.iterrows():
    vals = [names[x.period], "Day-ahead + alle reserver (ubegrænset volumen)" if x.reserves else "Kun day-ahead-arbitrage",
            x.arb, x.grid, x.afrr, x.fcr, x.mfrr, x.net, x.mwh_out, int(x.days)]
    for col, v in enumerate(vals, 1):
        c = wb_.cell(row=r, column=col, value=v); c.border = BOX
        if 3 <= col <= 8: c.number_format = EUR
        if col == 9: c.number_format = "#,##0"
    r += 1
ltm_da = bt[(bt.period == "LTM") & (~bt.reserves)].iloc[0]
r += 1
wb_.cell(row=r, column=1, value="B. Varighed, kun day-ahead, seneste 12 mdr. (bruges af modellen)").font = F_B; r += 1
header(wb_, r, ["Varighed (timer)", "Kilde", "Arbitrage brutto EUR/MW/år", "MWh afladet pr. MW/år", "Netto EUR/MW/år"]); r += 1
DUR_ROW0 = r
dur_rows = [(1, "backtest_varianter.py", 46403, 425, 36404),
            (2, "backtest_dk1.py", float(ltm_da.arb), float(ltm_da.mwh_out), float(ltm_da.arb + ltm_da.grid)),
            (4, "backtest_varianter.py", 135158, 1302, 104536)]
for d in dur_rows:
    for col, v in enumerate(d, 1):
        c = wb_.cell(row=r, column=col, value=v); c.border = BOX
        if col in (3, 5): c.number_format = EUR
        if col == 4: c.number_format = "#,##0"
    r += 1
DUR_ROW1 = r - 1
r += 1
wb_.cell(row=r, column=1, value="C. Mætning af systemydelser: 2h, seneste 12 mdr., aFRR + mFRR (uden FCR), reservepriser skaleret").font = F_B; r += 1
header(wb_, r, ["Reservepris-faktor", "Arbitrage brutto", "Netomkostning", "aFRR", "mFRR", "Netto", "Merindtægt ift. ren arbitrage"]); r += 1
base_net = float(ltm_da.arb + ltm_da.grid)
for k, a, g, af, mf in [(1.0, 54252, -11364, 95069, 82536), (0.5, 64544, -13162, 44689, 37960), (0.25, 72429, -14877, 20003, 16979)]:
    net = a + g + af + mf
    for col, v in enumerate([k, a, g, af, mf, net, net - base_net], 1):
        c = wb_.cell(row=r, column=col, value=v); c.border = BOX
        c.number_format = "0.00" if col == 1 else EUR
    r += 1
r += 1
# kvartalstabel
def load(f): return pd.DataFrame(json.load(open(f"{DATA}/{f}"))["records"])
af = load("afrr.json"); af["t"] = pd.to_datetime(af.TimeDK)
fc = load("fcr.json"); fc["t"] = pd.to_datetime(fc.HourDK)
mf = load("mfrr.json"); mf["t"] = pd.to_datetime(mf.TimeDK)
h = load("spot_h.json"); h["t"] = pd.to_datetime(h.HourDK); h["p"] = h.SpotPriceEUR
q = load("da_15.json"); q["t"] = pd.to_datetime(q.TimeDK); q["p"] = q.DayAheadPriceEUR
sp = pd.concat([h[["t", "p"]], q[["t", "p"]]]); sp["hr"] = sp.t.dt.floor("h")
hp = sp.groupby("hr").p.mean().reset_index(); hp["d"] = hp.hr.dt.floor("D")
ds = hp.groupby("d").p.apply(lambda v: (sum(sorted(v)[-2:]) - sum(sorted(v)[:2])) / 2).rename("s").reset_index()
ds["m"] = ds.d.dt.to_period("Q"); hp["m"] = hp.hr.dt.to_period("Q")
tab = pd.DataFrame({"spread": ds.groupby("m").s.mean(), "mean": hp.groupby("m").p.mean()})
for nm, df, col in [("afu", af, "UpPriceEUR"), ("afd", af, "DownPriceEUR"), ("fcr", fc, "FCRdk_EUR"), ("mfu", mf, "UpPriceEUR")]:
    df["m"] = df.t.dt.to_period("Q"); tab[nm] = df.groupby("m")[col].mean()
af["m"] = af.t.dt.to_period("Q"); tab["afud"] = af.groupby("m").UpDemandMW.mean(); tab["afdd"] = af.groupby("m").DownDemandMW.mean()
fc["m"] = fc.t.dt.to_period("Q"); tab["fcrd"] = fc.groupby("m").FCRdomestic_MW.mean()
mf["m"] = mf.t.dt.to_period("Q"); tab["mfd"] = mf.groupby("m").UpDemandMW.mean()
wb_.cell(row=r, column=1, value="D. DK1 pr. kvartal: prisspænd og reservepriser (gennemsnit). Q4 2026 er kun 1.-13. oktober.").font = F_B; r += 1
header(wb_, r, ["Kvartal", "Dagligt spænd top-2h minus bund-2h (EUR/MWh)", "Gns. spotpris (EUR/MWh)", "aFRR op (EUR/MW/h)",
                "aFRR ned (EUR/MW/h)", "FCR DK1 (EUR/MW/h)", "mFRR op (EUR/MW/h)", "aFRR op behov (MW)", "aFRR ned behov (MW)",
                "FCR fra danske enheder (MW)", "mFRR op behov (MW)"]); r += 1
for m, x in tab.iterrows():
    vals = [str(m)] + [None if pd.isna(v) else round(float(v), 1) for v in
                       (x.spread, x["mean"], x.afu, x.afd, x.fcr, x.mfu, x.afud, x.afdd, x.fcrd, x.mfd)]
    for col, v in enumerate(vals, 1):
        c = wb_.cell(row=r, column=col, value=v); c.border = BOX
        if col > 1: c.number_format = "0.0"
    r += 1
wb_.row_dimensions[8].height = 30

# =====================================================================
# Scenarie-ark
# =====================================================================
NY = 20  # max år
FI = "'Forudsætninger'"

def scenario_sheet(title, idx):
    ws = wb.create_sheet(title)
    ws["A1"] = f"Cash flow: {title} scenarie (EUR, nominelt)"; ws["A1"].font = F_TITLE
    ws["A2"] = "Scenarie-indeks (1 = pess., 2 = base, 3 = opt.)"; ws["B2"] = idx
    ws.column_dimensions["A"].width = 52; ws.column_dimensions["B"].width = 16
    for c in range(3, 3 + NY + 1): ws.column_dimensions[L(c)].width = 12
    # parameter pull
    ws["A4"] = "Forudsætninger for scenariet (hentes fra arket Forudsætninger)"; ws["A4"].font = F_B
    A = {}
    rr = 5
    for row in rows:
        if row[0] == "SEC": continue
        key, name = row[0], row[1]
        ws.cell(row=rr, column=1, value=f"{name} [{row[2]}]")
        c = ws.cell(row=rr, column=2, value=f"=INDEX({FI}!$D${P[key]}:$F${P[key]},1,$B$2)")
        c.number_format = PCT if row[2].startswith("%") else (NUM2 if row[2] in ("faktor", "øre/kWh", "DKK/EUR") else EUR)
        A[key] = f"$B${rr}"; rr += 1
    rr += 1
    ws.cell(row=rr, column=1, value="Afledte størrelser").font = F_B; rr += 1
    BT = "'Backtest'"
    derived = [
        ("mwh", "Energi (MWh)", f"={A['mw']}*{A['dur']}", EUR),
        ("capex_pl", "Anlægscapex (effekt + energi)", f"={A['mw']}*1000*{A['ckw']}+{A['mw']}*{A['dur']}*1000*{A['ckwh']}", EUR),
        ("capex", "Samlet investering", f"=B{{capex_pl}}+{A['dev']}+{A['conn']}", EUR),
        ("ceur_kwh", "Samlet investering pr. kWh", f"=B{{capex}}/(B{{mwh}}*1000)", EUR),
        ("ceur_mw", "Samlet investering pr. MW", f"=B{{capex}}/{A['mw']}", EUR),
        ("arb_ltm", "Arbitrage brutto, backtest (EUR/MW/år)",
         f"=INDEX({BT}!$C${DUR_ROW0}:$C${DUR_ROW1},MATCH({A['dur']},{BT}!$A${DUR_ROW0}:$A${DUR_ROW1},0))", EUR),
        ("mwh_ltm", "MWh afladet, backtest (pr. MW/år)",
         f"=INDEX({BT}!$D${DUR_ROW0}:$D${DUR_ROW1},MATCH({A['dur']},{BT}!$A${DUR_ROW0}:$A${DUR_ROW1},0))", "#,##0"),
        ("tc", "Omkostning pr. MWh opladet (EUR, 2026-niveau)", f"=({A['t_en']}+{A['t_dso']}+{A['t_afg']})*10/{A['fx']}", NUM2),
        ("td", "Omkostning pr. MWh afladet (EUR, 2026-niveau)", f"=({A['t_enp']}+{A['t_dsop']})*10/{A['fx']}", NUM2),
    ]
    for key, name, f, fmt in derived:
        A[key] = f"$B${rr}"; ws.cell(row=rr, column=1, value=name)
        ws.cell(row=rr, column=2, value=f); ws.cell(row=rr, column=2).number_format = fmt; rr += 1
    # resolve {key} placeholders
    for row_ in ws.iter_rows(min_row=1, max_row=rr, min_col=2, max_col=2):
        c = row_[0]
        if isinstance(c.value, str) and "{" in c.value:
            v = c.value
            for k, addr in A.items(): v = v.replace("B{" + k + "}", addr.replace("$", ""))
            c.value = v
    rr += 1
    # cash flow
    ws.cell(row=rr, column=1, value="Cash flow").font = F_B; rr += 1
    R = {}
    def line(key, label, fn, fmt=EUR, bold=False):
        nonlocal rr
        R[key] = rr
        ws.cell(row=rr, column=1, value=label).font = F_B if bold else Font()
        for j in range(NY + 1):
            col = L(3 + j); prev = L(2 + j)
            v = fn(j, col, prev)
            c = ws.cell(row=rr, column=3 + j, value=v); c.number_format = fmt
            if bold: c.font = F_B
        rr += 1
    line("t", "År", lambda j, c, p: j, "0", True)
    line("cal", "Kalenderår", lambda j, c, p: "" if j == 0 else f"={A['cod']}+{c}{R['t']}-1", "0")
    line("on", "Drift (1/0)", lambda j, c, p: 0 if j == 0 else f"=IF({c}{R['t']}<={A['life']},1,0)", "0")
    line("idx", "Prisindeks (2026 = 1)", lambda j, c, p: 1 if j == 0 else f"=(1+{A['infl']})^({c}{R['cal']}-2026)", NUM2)
    line("cf", "Kapacitetsfaktor", lambda j, c, p: 1 if j == 0 else f"=(1-{A['fade']})^({c}{R['t']}-1)", NUM2)
    line("arb", "Arbitrage (day-ahead + intraday)", lambda j, c, p: 0 if j == 0 else
         f"={A['mw']}*{A['arb_ltm']}*{A['lvl']}*{A['cap']}*(1+{A['idu']})*(1+{A['arbtr']})^({c}{R['t']}-1)*{c}{R['cf']}*{A['avail']}*{c}{R['idx']}*{c}{R['on']}")
    line("anc", "Systemydelser, netto merindtægt", lambda j, c, p: 0 if j == 0 else
         f"={A['mw']}*MAX({A['anc1']}*(1+{A['anctr']})^({c}{R['t']}-1),{A['ancfl']})*{A['avail']}*{c}{R['idx']}*{c}{R['on']}")
    line("rev", "Markedsindtægt i alt", lambda j, c, p: f"={c}{R['arb']}+{c}{R['anc']}", EUR, True)
    line("grid", "Nettariffer og elafgift", lambda j, c, p: 0 if j == 0 else
         f"=-{A['mw']}*{A['mwh_ltm']}*{c}{R['cf']}*{A['avail']}*({A['tc']}/{A['rte']}+{A['td']})*{c}{R['idx']}*{c}{R['on']}")
    line("om", "Drift, vedligehold og forsikring", lambda j, c, p: 0 if j == 0 else
         f"=-({A['om']}+{A['ins']})*{A['capex_pl']}*{c}{R['idx']}*{c}{R['on']}")
    line("fee", "Optimeringsfee", lambda j, c, p: f"=-{A['fee']}*{c}{R['rev']}")
    line("rfv", "Betaling til RFV (netadgang og areal)", lambda j, c, p: 0 if j == 0 else f"=-{A['rfv']}*{A['mw']}*{c}{R['idx']}*{c}{R['on']}")
    line("fix", "Faste omkostninger", lambda j, c, p: 0 if j == 0 else f"=-{A['fix']}*{c}{R['idx']}*{c}{R['on']}")
    line("ebitda", "EBITDA", lambda j, c, p: f"={c}{R['rev']}+{c}{R['grid']}+{c}{R['om']}+{c}{R['fee']}+{c}{R['rfv']}+{c}{R['fix']}", EUR, True)
    line("inv", "Investering", lambda j, c, p: f"=-{A['capex']}" if j == 0 else 0)
    line("cfpre", "Cash flow før skat", lambda j, c, p: f"={c}{R['ebitda']}+{c}{R['inv']}", EUR, True)
    rr += 1
    ws.cell(row=rr, column=1, value="Skat").font = F_B; rr += 1
    line("sp", "Saldo primo", lambda j, c, p: 0 if j == 0 else (f"={A['capex']}" if j == 1 else f"={p}{rr + 2}"))
    line("dep", "Skattemæssig afskrivning", lambda j, c, p: 0 if j == 0 else
         f"=IF({c}{R['t']}={A['life']},{c}{R['sp']},{c}{R['sp']}*{A['saldo']})*{c}{R['on']}")
    line("su", "Saldo ultimo", lambda j, c, p: 0 if j == 0 else f"={c}{R['sp']}-{c}{R['dep']}")
    line("ti", "Skattepligtig indkomst før underskud", lambda j, c, p: 0 if j == 0 else f"={c}{R['ebitda']}-{c}{R['dep']}")
    line("lp", "Fremført underskud primo", lambda j, c, p: 0 if j == 0 else f"={p}{rr + 2}")
    line("lu", "Anvendt underskud", lambda j, c, p: 0 if j == 0 else f"=MIN({c}{R['lp']},MAX({c}{R['ti']},0))")
    line("lul", "Fremført underskud ultimo", lambda j, c, p: 0 if j == 0 else f"={c}{R['lp']}-{c}{R['lu']}+MAX(-{c}{R['ti']},0)")
    line("tax", "Selskabsskat", lambda j, c, p: 0 if j == 0 else f"=-MAX({c}{R['ti']}-{c}{R['lu']},0)*{A['tax']}")
    line("cfpost", "Cash flow efter skat", lambda j, c, p: f"={c}{R['cfpre']}+{c}{R['tax']}", EUR, True)
    line("cum", "Akkumuleret cash flow efter skat", lambda j, c, p: f"={c}{R['cfpost']}" if j == 0 else f"={p}{rr}+{c}{R['cfpost']}")
    line("pb", "Tilbagebetalt i år (hjælpelinje)", lambda j, c, p: "" if j == 0 else
         f"=IF(AND({c}{R['cum']}>=0,{p}{R['cum']}<0),{c}{R['t']},\"\")", "0")
    rr += 1
    first, last = L(3), L(3 + NY)
    K = {}
    ws.cell(row=rr, column=1, value="Nøgletal").font = F_B; rr += 1
    kpis = [
        ("irr_pre", "IRR før skat (projekt, ugearet)", f"=IFERROR(IRR({first}{R['cfpre']}:{last}{R['cfpre']}),\"negativ\")", PCT),
        ("irr_post", "IRR efter skat (projekt, ugearet)", f"=IFERROR(IRR({first}{R['cfpost']}:{last}{R['cfpost']}),\"negativ\")", PCT),
        ("npv", "NPV efter skat", f"=NPV({A['disc']},{L(4)}{R['cfpost']}:{last}{R['cfpost']})+{first}{R['cfpost']}", EUR),
        ("pb", "Simpel tilbagebetalingstid (år)", f"=IF(COUNT({L(4)}{R['pb']}:{last}{R['pb']})=0,\"> levetid\",MIN({L(4)}{R['pb']}:{last}{R['pb']}))", "0"),
        ("capex", "Samlet investering", f"={A['capex']}", EUR),
        ("ceur_kwh", "Investering pr. kWh", f"={A['ceur_kwh']}", EUR),
        ("ceur_mw", "Investering pr. MW", f"={A['ceur_mw']}", EUR),
        ("rev1", "Markedsindtægt år 1 pr. MW", f"={L(4)}{R['rev']}/{A['mw']}", EUR),
        ("ebitda1", "EBITDA år 1", f"={L(4)}{R['ebitda']}", EUR),
        ("ebm1", "EBITDA-margin år 1", f"=IFERROR({L(4)}{R['ebitda']}/{L(4)}{R['rev']},0)", PCT),
        ("yield1", "EBITDA år 1 / investering", f"={L(4)}{R['ebitda']}/{A['capex']}", PCT),
    ]
    for key, name, f, fmt in kpis:
        ws.cell(row=rr, column=1, value=name)
        c = ws.cell(row=rr, column=2, value=f); c.number_format = fmt; c.fill = FILL_KPI; c.font = F_B
        K[key] = f"'{title}'!$B${rr}"; rr += 1
    ws.freeze_panes = "C5"
    return ws, R, K, A

sheets = {}
for i, t in enumerate(["Pessimistisk", "Base", "Optimistisk"], 1):
    sheets[t] = scenario_sheet(t, i)

# =====================================================================
# Følsomhed (base, IRR før skat)
# =====================================================================
wf = wb.create_sheet("Følsomhed")
wf["A1"] = "Følsomhed: IRR før skat i base-scenariet ved ændret investering og markedsindtægt"; wf["A1"].font = F_TITLE
wf["A2"] = ("Markedsindtægt skaleres (arbitrage + systemydelser); nettariffer følger volumen og skaleres ikke. "
            "Drift og forsikring følger investeringen. Ændr faktorerne i de gule celler.").strip()
wf["A2"].font = F_NOTE
_, RB, KB, AB = sheets["Base"]
capf = [0.6, 0.8, 0.9, 1.0, 1.1, 1.2]
revf = [0.7, 0.85, 1.0, 1.15, 1.3, 1.5]
wf["A4"] = "Markedsindtægt (faktor) ned / investering (faktor) hen"; wf["A4"].font = F_B
wf.column_dimensions["A"].width = 44
for j, cf in enumerate(capf):
    c = wf.cell(row=4, column=2 + j, value=cf); c.fill = FILL_IN; c.font = F_INPUT; c.number_format = "0%"
    wf.column_dimensions[L(2 + j)].width = 12
for i, rf in enumerate(revf):
    c = wf.cell(row=5 + i, column=1, value=rf); c.fill = FILL_IN; c.font = F_INPUT; c.number_format = "0%"
# helper block
H0 = 14
wf.cell(row=H0 - 1, column=1, value="Hjælpeberegning (cash flow før skat pr. kombination)").font = F_NOTE
hr_ = H0
for i in range(len(revf)):
    for j in range(len(capf)):
        wf.cell(row=hr_, column=1, value=f"=\"Indtægt \"&TEXT($A${5+i},\"0%\")&\" / investering \"&TEXT({L(2+j)}$4,\"0%\")")
        for y in range(NY + 1):
            col = L(3 + y)  # same column letters as Base sheet
            fr, fc_ = f"$A${5+i}", f"{L(2+j)}$4"
            if y == 0:
                f = f"='Base'!{col}{RB['inv']}*{fc_}"
            else:
                f = (f"={fr}*'Base'!{col}{RB['rev']}*(1-{AB['fee'].replace('$B', chr(39) + 'Base' + chr(39) + '!$B')})"
                     f"+'Base'!{col}{RB['grid']}+{fc_}*'Base'!{col}{RB['om']}+'Base'!{col}{RB['rfv']}+'Base'!{col}{RB['fix']}")
            wf.cell(row=hr_, column=2 + y, value=f).number_format = EUR
        wf.cell(row=5 + i, column=2 + j, value=f"=IFERROR(IRR(B{hr_}:{L(2+NY)}{hr_}),\"negativ\")").number_format = PCT
        hr_ += 1
for i in range(len(revf)):
    for j in range(len(capf)):
        c = wf.cell(row=5 + i, column=2 + j); c.border = BOX; c.font = F_B

# =====================================================================
# Resumé
# =====================================================================
wr = wb.create_sheet("Resumé", 0)
wr["A1"] = "BESS ved Ringkøbing Fjernvarme: tidlig investeringsberegning"; wr["A1"].font = F_TITLE
wr["A2"] = (f"Udarbejdet {DATO} til Kristian Hautorp. Status: TIL GODKENDELSE. Tidlig skrivebordsberegning uden data fra Ringkøbing Fjernvarme (RFV), "
            "uden tilbud og uden afklaring med netselskab og Skattestyrelsen.")
wr["A2"].font = F_NOTE
wr.column_dimensions["A"].width = 64
for c in "BCD": wr.column_dimensions[c].width = 18
wr.column_dimensions["E"].width = 70
header(wr, 4, ["Nøgletal", "Pessimistisk", "Base", "Optimistisk"])
krows = [("Effekt (MW)", None, "mw", "0"), ("Varighed (timer)", None, "dur", "0"),
         ("Samlet investering (EUR)", "capex", None, EUR), ("Investering pr. kWh (EUR)", "ceur_kwh", None, EUR),
         ("Investering pr. MW (EUR)", "ceur_mw", None, EUR), ("Markedsindtægt år 1 pr. MW (EUR)", "rev1", None, EUR),
         ("EBITDA år 1 (EUR)", "ebitda1", None, EUR), ("EBITDA år 1 / investering", "yield1", None, PCT),
         ("IRR før skat (ugearet)", "irr_pre", None, PCT), ("IRR efter skat (ugearet)", "irr_post", None, PCT),
         ("NPV efter skat ved 8 % (EUR)", "npv", None, EUR), ("Simpel tilbagebetalingstid (år)", "pb", None, "0")]
for i, (name, k, a, fmt) in enumerate(krows):
    rr = 5 + i
    wr.cell(row=rr, column=1, value=name).border = BOX
    for j, t in enumerate(["Pessimistisk", "Base", "Optimistisk"]):
        _, R_, K_, A_ = sheets[t]
        f = f"={K_[k]}" if k else f"='{t}'!{A_[a]}"
        c = wr.cell(row=rr, column=2 + j, value=f); c.number_format = fmt; c.border = BOX
        if k in ("irr_pre", "irr_post"): c.font = F_B; c.fill = FILL_KPI
rr = 5 + len(krows) + 1
wr.cell(row=rr, column=1, value="Hvad flytter base-scenariet? (én ændring ad gangen, IRR efter skat)").font = F_B; rr += 1
wr.cell(row=rr, column=1, value="Statiske tal beregnet 13.10.2026 med scripts/kontroller_model.py (samme logik som arkene). Opdateres ikke automatisk.").font = F_NOTE; rr += 1
for t, v in [("Base som beregnet", "ca. -1 %"),
             ("Systemydelser holder dagens niveau (70k EUR/MW i stedet for 30k)", "ca. 6 %"),
             ("Investering 250 EUR/kWh i stedet for ca. 310", "ca. 2 %"),
             ("Ingen betaling til RFV for netadgang", "ca. 1 %"),
             ("Intraday/ubalance-tillæg 20 % i stedet for 10 %", "ca. 1 %"),
             ("Levetid 20 år i stedet for 15", "ca. 1 %"),
             ("4 timer i stedet for 2 timer", "ca. -2 %"),
             ("Fuld elafgift 60,9 øre/kWh på opladning", "negativ (casen falder)"),
             ("Krav for 8 % efter skat: investering", "ca. 190 EUR/kWh"),
             ("Krav for 8 % efter skat: markedsindtægt", "ca. +43 %")]:
    wr.cell(row=rr, column=1, value=t).border = BOX
    c = wr.cell(row=rr, column=2, value=v); c.border = BOX; c.font = F_B
    rr += 1
rr += 1
text = [
    ("Sådan læses beregningen", True),
    ("Indtjeningen bygger på en backtest af faktiske DK1-priser fra Energinet (okt. 2025 til okt. 2026), korrigeret for at et rigtigt batteri ikke kender priserne på forhånd. "
     "Systemydelser er sat lavere end i dag, fordi priserne allerede er faldet kraftigt, efterhånden som flere batterier er kommet på markedet. Se arket Backtest.", False),
    ("Base-scenariet er bevidst konservativt. Resultatet er meget følsomt over for investering og markedsindtægt; se arket Følsomhed.", False),
    ("", False),
    ("Stop-risici (skal afklares før der bruges penge på udvikling)", True),
    ("1. Elafgift: modellen antager godtgørelse ned til 0,4 øre/kWh på opladning. Hvis den fulde afgift (60,9 øre/kWh fra 2028) rammer batteriet, falder business casen. Kræver bindende svar.", False),
    ("2. Netadgang: om netselskabet tillader et batteri (evt. ejet af et andet selskab end RFV) bag RFV's tilslutning, og om eksisterende kapacitet kan bruges uden nyt tilslutningsbidrag.", False),
    ("3. Samtidighed: elkedel og gasmotorer bruger tilslutningen i de samme timer, som batteriet vil bruge. Kræver RFV's 15-min. måledata og tilslutningsaftale.", False),
    ("4. Investering: base svarer til ca. 300 EUR/kWh all-in for 2h. Skal erstattes af 2-3 indikative tilbud.", False),
    ("5. Selskabsform: RFV er et forbrugerejet a.m.b.a. under varmeforsyningsloven. Deltagelse skal formentlig ske via et separat selskab. Skal verificeres juridisk.", False),
    ("", False),
    ("Bemærk: modellen regner ugearet projektafkast. Gæld hæver kun egenkapitalafkastet, hvis projektafkastet ligger over lånerenten.", False),
]
for t, bold in text:
    c = wr.cell(row=rr, column=1, value=t); c.font = F_B if bold else Font()
    c.alignment = Alignment(wrap_text=True, vertical="top")
    wr.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=5)
    if not bold and t: wr.row_dimensions[rr].height = 32
    rr += 1

# =====================================================================
# Kilder
# =====================================================================
wk = wb.create_sheet("Kilder")
header(wk, 1, ["Kilde", "Dato", "Bruges til", "Link"], [60, 14, 60, 90])
src = [
    ("Energinet, Energi Data Service: Elspotprices, DayAheadPrices, AfrrReservesNordic, FcrDK1, mFRRCapacityMarket (DK1)", "hentet 13.10.2026", "Backtest af indtjening og prisudvikling", "https://www.energidataservice.dk"),
    ("Energinet: Energinet sænker elforbrugernes tariffer yderligere (Via Ritzau)", "04.09.2026", "Energinets forbrugs-, indfødnings- og balancetarif 2027 og tarifloft", "https://via.ritzau.dk/pressemeddelelse/15138077/energinet-saenker-elforbrugernes-tariffer-yderligere"),
    ("Elnet.dk (Green Power Denmark): Tarifering af batterier og ellager", "ikke dateret", "Batterier betaler forbrugs- og indfødningstarif; intet tilslutningsbidrag ved brug af eksisterende kapacitet", "https://elnet.dk/nettariffer-priser-gebyrer/tarifering-batterier-ellager"),
    ("Lov nr. 1775 af 29.12.2025 om nedsættelse af elafgiften i 2026 og 2027", "29.12.2025", "Elafgift 0,8 øre/kWh 2026-27, 60,9 øre fra 2028", "https://www.retsinformation.dk/eli/lta/2025/1775/pdf"),
    ("Skattestyrelsen: Midlertidig nedsættelse af elafgiften i 2026 og 2027", "2025", "Godtgørelse ned til 0,4 øre/kWh for momsreg. virksomheder", "https://skat.dk/erhverv/afgifter-paa-varer-og-ydelser-punktafgifter/nyhedsbrev-afgifter/midlertidig-nedsaettelse-af-elafgiften-i-2026-og-2027"),
    ("Modo Energy: German BESS investment outlook Q2 2026, executive summary", "16.03.2026", "Capex-reference 2h/4h og IRR-reference (Tyskland, ikke DK1)", "https://modoenergy.com/research/en/germany-battery-investment-outlook-q2-2026-executive-summary"),
    ("Energy-Storage.news om BNEF Energy Storage Systems Cost Survey 2025", "dec. 2025", "Turnkey-pris Europa 177 USD/kWh ekskl. EPC og nettilslutning", "https://www.energy-storage.news/?p=92133"),
    ("Dansk Fjernvarme: Ringkøbing vinker farvel til gas", "30.07.2025", "RFV's anlæg: elkedel og gasmotorer på elmarkedet, 500 MWh akkutanke, fase 1 og 2", "https://danskfjernvarme.dk/aktuelt/nyheder/2025/ringkoebing-vinker-farvel-til-gas"),
    ("Radius Elnet: Tilslutning af batterier", "ikke dateret", "Proces og krav ved batteritilslutning (Radius' område, kun som indikation)", "https://radiuselnet.dk/ny-tilslutning/tilslutning/tilslutning-af-batterier/"),
    ("Energistyrelsen: Energifællesskab Rødkærsbro Fjernvarmeværk og Nordic Energy Storage", "12.02.2025", "Præcedens: batteri på fjernvarmeværk", "https://ens.dk/media/7751/download"),
]
for i, s in enumerate(src, 2):
    for j, v in enumerate(s, 1):
        c = wk.cell(row=i, column=j, value=v); c.alignment = WRAP; c.border = BOX

for ws_ in (wr, wi, wk):
    ws_.page_setup.orientation = "landscape"
    ws_.sheet_properties.pageSetUpPr.fitToPage = True
    ws_.page_setup.fitToWidth = 1
    ws_.page_setup.fitToHeight = 0

wb.save(OUT)
print("gemt", OUT)
