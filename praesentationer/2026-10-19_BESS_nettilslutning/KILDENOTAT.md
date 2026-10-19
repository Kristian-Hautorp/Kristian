# Kildenotat: "Right-sized batteries for a faster grid connection"

Internt notat til Kristian Hautorp. Udarbejdet 19.10.2026. **Status: udkast, til godkendelse.**

Præsentation (privat link, skal deles før Antoni kan åbne den): https://claude.ai/artifact/Dge56kbKAw88BbzayYJ7kc
Kildefilerne til præsentationen ligger i `deck/` i denne mappe.

## Opdatering 2, 19.10.2026: nyt design er 4 timer

Kristian 19.10.2026: alle 2-timers størrelser er det tidligere design, alle 4-timers størrelser er det nye. 3-timers tallene fra FM sep. 2025 er taget helt ud.

Præsentationen er rettet (slide 9-13, forsiden og opsummeringen): mindre effekt, 4 timers lager og begrænset netadgang.

| Projekt | Tidligere (2 t) | Nyt (4 t) | Ændring MW / MWh | Kilde for nyt design |
|---|---|---|---|---|
| M01B | 90 MW / 180 MWh | **[mangler]** | | Kristian: ny ansøgning netop indsendt |
| M23 | 120 MW / 240 MWh | **[mangler]** | | Ved ikke om der er indsendt ny |
| M30 | 90 MW / 180 MWh | 60 MW / 240 MWh | −33 % / +33 % | Indsendt til Nord Energi Net 17.09.2026 |
| M34 | 56,25 MW / 112,5 MWh | 45 MW / 180 MWh | −20 % / +60 % | Indsendt til Nord Energi Net 17.09.2026 |
| M30 + M34 | 146,25 MW / 292,5 MWh | 105 MW / 420 MWh | −28 % / +44 % | Beregnet |

Kilde for 2- og 4-timers tallene: Kristians egen præsentation "Mermaid Portfolio: Four Danish Solar + BESS Projects" (claude.ai artifact, opdateret 21.09.2026), slide "Battery storage". Den angiver at de reviderede ansøgninger dimensionerer lageret til 50 % af tilslutningskapaciteten på M30 og 60 % på M34 mod accept af permanent curtailment, og at svar fra Nord Energi Net udestår. Tjek: 60/120 MWac = 50 %, 45/75 MWac = 60 %.

**Skal verificeres:**
1. "Lower grid connection costs expected for a curtailable connection" (slide 12) bygger på GreenGos SteerCo-slides (`260610 SteerCo_Mermaid_Recap (corrected).docx`: DKK 60.000 vs. 20.000/MW/år) og Marco Bertazzis mail 04.09.2026 ("The expectation is that this would ... reduce connection costs"). Ikke bekræftet af netselskaberne.
2. M01B- og M23-tallene står som `[__ MW / __ MWh]` på slide 10 og 13 indtil de er leveret.

## Opdatering 1, 19.10.2026 efter Kristians svar

Kristians afgørelser: (1) M01B 90/180 er korrekt, men er den gamle; en ny ansøgning er netop indsendt. (2) M34-netansøgningen blev opdateret midt i september. (3) Tallene for det tidligere design indeholder fejl. (4) Ikrafttrædelsesdato: "October 2026" er fint.

Søgt i Drive-kopien af skrivebordsarkivet ("My Laptop / Mermaid Portefølje"). Kopien er fra 14.09.2026, så alt indsendt efter den dato mangler. Det rigtige skrivebord ligger i Sterns OneDrive/SharePoint, som denne session ikke har adgang til. Arkivets egen faktafil (`CLAUDE.md` i roden) og `Mermaid_Portfolio_overview_v1.5.xlsx` er ikke med i kopien.

Fundet (seneste før 14.09):

| Projekt | Fund | Kilde |
|---|---|---|
| M30 | M30b BESS 45 MW ac + M30c BESS 45 MW ac, import 45 MW hver, FNA-F | `M30B BESS Erklæring af modenhed.pdf` og `M30C BESS Erklæring af modenhed.pdf`, dateret 31.08.2026, til Nord Energi Net |
| M30 | "min. 90 MW / 180 MWh (2h)" | `_STATUS_M30.md`, autogenereret 14.09.2026 fra arkivets CLAUDE.md |
| M34 | 56,25 MW ac, import 56,25 MW, FNA-F | `M34 BESS Erklæring af modenhed.pdf`, dateret 31.08.2026, til Nord Energi Net |
| M30/M34 | Opdateret ansøgning skal afklare "the change from FNA to LNA"; BESS-konfiguration "the lower the effect the better" | Mail fra Mathias Schou Gammelgaard (BattMan Energy) til Encavis, 14.09.2026 |
| Portefølje | Mål: flytte samplaceret BESS fra kategori 3 til kategori 2 via netvenlige tiltag | Mail fra Alexander Wobedo (Encavis) til Aurora, 09.09.2026 |
| M01B/M23 | Energinets model: behind-the-meter (ingen netimport) rangerer højere end grid import; abonnement fuld kapacitet DKK 60.000/MW/år, begrænset DKK 20.000/MW/år | `260610 SteerCo_Mermaid_Recap (corrected).docx` (GreenGos slides, skal verificeres) |

Ikke fundet: M01B's nye ansøgning, M34's ansøgning fra midt september, og korrekte tal for det tidligere design. **Slide 9, 10 og 13 afventer disse tal og er ikke rettet endnu.**

## Antagelser

1. **Sprog:** engelsk, fordi Antoni Beltran Dauden (Chief Development Officer, Encavis) skal bruge den over for investorer. Kilde til titel: `Kopie von Mermaid_Email_List_2026-06-16.xlsx`.
2. **"Skifte spørgelse og kapacitet"** er tolket som ændret størrelse (MWh, lagerkapacitet) og effekt (MW). Hvis du mente spændingsniveau (kV), skal slide 9 og 10 laves om.
3. **Investorer behandles som eksterne.** Derfor nævner hverken slides eller speaker notes interne dokumenter (Q2 Strategy Paper, FM, SteerCo), IRR, capex, sunk cost, tvister eller DSA-frister. Tallene står som "draft planning basis". Speaker notes kan læses af alle med linket.
4. **Billeder:** Jeg havde ingen projektfotos til rådighed, så alle illustrationer er tegnet (SVG). Fotos fra pladserne kan sættes ind.

## Kilder pr. slide

| Slide | Påstand | Kilde |
|---|---|---|
| 2 | Firedobling af sol + landvind til 2030; 12 TWh (2021), 25 TWh (2030 uden nye tiltag), ca. 50 TWh (2030 med mål) | KEFM, "Faktaark - Firedobling af VE på land", Klimaaftale om grøn strøm og varme 2022 |
| 3 | Citat: "energilagring i fjernvarmen og batterier kan have en vigtig rolle for at udnytte den fulde kapacitet i elnettet"; lager "kan også belaste nettet uhensigtsmæssigt, hvis de ikke agerer netvenligt" | Aftale om akutplan for elnettet, 1.10.2026, s. 5 og s. 3 |
| 4 | Ca. 60 GW nyt forbrug i kø, maks. dansk elforbrug ca. 7 GW; drevet af datacentre, batterier, PtX | Energinet pressemeddelelse via Ritzau, 2.3.2026 |
| 4 | "Et batteri tæller som forbrug når det lader" | Energinet regner batterier med i "nyt forbrug" (2.3.2026); jeres egen screening af M01b/M23 BESS er anmeldt som "Forbrug" (`ISR - M01b-M23 BESS.pdf`, udskrevet 04.06.2026) |
| 5 | Pause 2.3.2026, pulje-model 3.6.2026, vedtaget 1.10.2026, næste tildeling efterår 2026, status senest ultimo 2027 | Aftale om akutplan for elnettet, s. 2 og s. 6; KEFM nyhed 1.10.2026 |
| 6 | Fire kategorier; kategori 3 = lager, inkl. lager i samplacerede anlæg "ud over brugerens almindelige produktions- eller forbrugsbehov"; kan prioriteres højere i kategorien; afvisning mulig i kat. 3 og 4 | Aftale om akutplan for elnettet, s. 3 og s. 5 |
| 7 | Fire netvenlighedskriterier; citat "Projekter, der belaster nettet mindre, kan få en bedre placering i køen..." | Aftale om akutplan for elnettet, s. 4 og s. 5 |
| 8 | Sol + batteri bag samme tilslutning | Illustrativt koncept, ingen projektdata |
| 9-10, 13 | Erstattet: se "Opdatering 2" øverst. 3-timers tallene fra `Mermaid_Working_Dashboard_v1.1.xlsx` bruges ikke længere | |
| 13 | PV MWp 168 / 224 / 169 / 105 | `Mermaid_Portfolio_Forside_v4.xlsx` (2026-05-21), "Confirmed ... all PV MWp figures (SteerCo April 2026)" |
| 13 | M01B + M23 deler 150 kV Lem Kær (280 MWac IA); M30 + M34 via Nord Energi 60 kV | `Energinet_Grid_Prioritisation_v3.docx` (2026-05-08); `Mermaid_Working_Dashboard_v1.1.xlsx` |

Beregningen af −25 % / −50 % på 3-timers grundlag er udgået (se Opdatering 2).

## Kildekonflikter: du skal afgøre

1. **M01B BESS:** 88 MW/176 MWh (`Mermaid_Working_Dashboard_v1.1.xlsx`, 2026-05-09; `Energinet_Grid_Prioritisation_v3.docx`, 2026-05-08) mod 90 MW/180 MWh markeret "?" (`Mermaid_Portfolio_Forside_v4.xlsx`, 2026-05-21). Præsentationen bruger 88/176.
2. **M23 BESS:** 120 MW/240 MWh (Forside v4; Working Dashboard "DSA 6 May 26") mod 119 MW (`Energinet_Grid_Prioritisation_v3.docx`). Præsentationen bruger 120/240.
3. **M30 BESS DSA-status:** Forside v4 kalder 90/180 "Confirmed ... (BESS DSAs)", mens Working Dashboard v1.1 skriver "Not signed". `00 JURIST-GUIDE - start her (06.09.2026).pdf` nævner "M30 BESS-tillæg + præcisering" som signeret. Tallene er ens (90/180), men grundlaget er uklart.
4. **M34 netansøgning vs. design (vigtig):** PDR 56,25 MW/112,5 MWh, men BESS-netansøgningen er på 75 MW (`Mermaid_Grid_Flexibility_Decision_GPT.pdf`, 11.09.2026: "Applications and drawings use different BESS capacities"). Hvis den indsendte ansøgning stadig siger 75 MW, afspejler ansøgningen ikke den mindre størrelse, som præsentationen bygger på.
5. **Porteføljetotal tidligere design:** Working Dashboard v1.1 angiver "~430 MW / ~1,300 MWh" (FM), men summen af projekttallene i samme fil er 398 MW/1.195 MWh for tre projekter (M34 har ingen FM-værdi). Jeg har ikke selv set FM-filen. Præsentationen bruger projektsummen og skriver "n/a" for M34.
6. **PV-størrelse:** Forside v4 (168/224/169/105 MWp, SteerCo april 2026) mod FM sep. 2025 i Working Dashboard (165,5/224/168,1/105). Præsentationen bruger Forside v4.
7. **Ikrafttræden af akutplanen:** KEFM (1.10.2026) skriver "Efter planen træder loven i kraft 12. oktober". TV2 Øst skriver "Mandag træder ... i kraft" uden dato, og en sekundær kilde nævner 19.10. Præsentationen skriver kun "October 2026".
8. **Høringsfrist for prioriteringsbekendtgørelsen:** `Mermaid_Grid_Flexibility_Decision_GPT.pdf` siger 22.09.2026; Høringsportalen (høring 71398) viser 07.10.2026 for den supplerende høring af § 15. Ikke brugt i præsentationen.

## Skal verificeres før Antoni bruger den

1. **Kategori 3 eller ej:** Argumentet om at et batteri på 75 % af PV-effekten og 2 timer ikke er "lager ud over almindelige produktionsbehov" er vores fortolkning. Der findes ingen bekræftet forholdstal-test (`Mermaid_Grid_Flexibility_Decision_GPT.pdf`, 11.09.2026). Bør bekræftes skriftligt med Energinet og Nord Energi.
2. **Gruppe 2 hos Energinet (M01B/M23):** Kun gyldig hvis BESS holdes inden for den eksisterende 280 MWac IA. Åben handling i `Energinet_Grid_Prioritisation_v3.docx`: bekræft med Kim Boe Jensen (GreenGo). Status efter loven af 1.10.2026: ved ikke.
3. **"Accepterer begrænset netadgang":** Det interne flex-notat anbefaler en fast import-kerne plus en afgrænset fleksibel del, ikke ubegrænset afbrydelighed, og vinter-PV kan ikke alene lade batteriet. Slide 11 siger derfor kun "better placed to accept". Antoni bør ikke love fuld afbrydelighed.
4. **M34 eksportkapacitet:** PV-aftalen giver 25 MW indtil netforstærkning, 75 MW endeligt (samme notat). Påvirker udsagnet "fits inside the solar park’s connection" for M34.
5. **Samlet anlæg eller flere:** Loven giver Energinet og netselskaberne ret til at afgøre, om et projekt behandles som ét samlet anlæg eller flere (akutplan s. 5). Relevant for hvordan PV + BESS vurderes.

## Forslag

1. Få tallene på slide 9, 10 og 13 bekræftet af Marco Bertazzi (Team Lead BESS Competence Center, Encavis) før deling.
2. Lad Viktoria Schneider (Legal) læse slide 5-7 og 11, da de gengiver ny lovgivning i forenklet form.
3. Sæt Encavis-skabelon og projektfotos på, hvis Antoni skal præsentere i Encavis-design. Præsentationen kan downloades som PowerPoint.
