#!/usr/bin/env bash
# Henter DK1-markedsdata fra Energinets Energi Data Service (offentligt API) til mappen angivet som $1.
set -euo pipefail
D="${1:?angiv datamappe}"; mkdir -p "$D"
API="https://api.energidataservice.dk/dataset"
END="${2:-2026-10-14}"
curl -gsS -m 300 "$API/Elspotprices?start=2023-01-01&end=2025-10-01&filter={\"PriceArea\":[\"DK1\"]}&limit=0&sort=HourUTC%20ASC" -o "$D/spot_h.json"
curl -gsS -m 300 "$API/DayAheadPrices?start=2025-10-01&end=$END&filter={\"PriceArea\":[\"DK1\"]}&limit=0&sort=TimeUTC%20ASC" -o "$D/da_15.json"
curl -gsS -m 300 "$API/AfrrReservesNordic?start=2023-01-01&end=$END&filter={\"PriceArea\":[\"DK1\"]}&limit=0&sort=TimeUTC%20ASC" -o "$D/afrr.json"
sleep 100  # API'et har rate limit
curl -gsS -m 300 "$API/FcrDK1?start=2023-01-01&end=$END&limit=0" -o "$D/fcr.json"
curl -gsS -m 300 "$API/mFRRCapacityMarket?start=2023-01-01&end=$END&filter={\"PriceArea\":[\"DK1\"]}&limit=0&sort=TimeUTC%20ASC" -o "$D/mfrr.json"
