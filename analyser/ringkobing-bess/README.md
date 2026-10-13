# BESS ved Ringkøbing Fjernvarme: tidlig investeringsberegning

Status: **til godkendelse**. Skrivebordsberegning fra 13.10.2026. Den bygger kun på offentlige kilder og indeholder ingen data fra Ringkøbing Fjernvarme (RFV), ingen tilbud og ingen afklaring med netselskab eller Skattestyrelsen. Der indgår ingen Stern- eller Encavis-data.

## Filer
- `Ringkobing_BESS_tidlig_beregning.xlsx`: modellen. Start i arket *Resumé*. Gule celler i *Forudsætninger* kan ændres, og scenarierne regner automatisk.
- `scripts/hent_data.sh`: henter DK1-data fra Energinets Energi Data Service (spot, aFRR, FCR og mFRR).
- `scripts/backtest_dk1.py`: døgnvis optimering (perfekt forudsigelse) af et 1 MW-batteri på de historiske priser.
- `scripts/backtest_varianter.py`: varianter med 1h/4h og med reservepriser skaleret ned (mætning).
- `scripts/byg_model.py`: bygger Excel-filen.
- `scripts/kontroller_model.py`: uafhængig Python-replika, der kontrollerer Excel-formlerne og beregner enkeltfølsomhederne.

## Kør igen
```bash
D=/sti/til/data
scripts/hent_data.sh $D
python3 scripts/backtest_dk1.py $D 2.0 18.6
python3 scripts/backtest_varianter.py scripts/backtest_dk1.py $D
python3 scripts/byg_model.py $D model.xlsx
soffice --headless --calc --convert-to xlsx --outdir ud model.xlsx   # genberegn formler
python3 scripts/kontroller_model.py ud/model.xlsx
```
Tallene fra `backtest_varianter.py` står i dag fast i `byg_model.py` (arket Backtest, afsnit B og C) og skal opdateres manuelt, hvis backtesten køres igen.
