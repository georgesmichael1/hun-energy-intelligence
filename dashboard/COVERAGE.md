# Metric coverage: Bosnia framework vs Hungary

Status legend: LIVE = fetched by the workflow from a public source; STATIC = fixed reference text;
RESEARCHED = a public source exists, no feed built yet; NONE = no public Hungarian equivalent found.
Actual / forecast labels are in the CSV column names (Plan, Forecast, Actual). Nothing is estimated.

| Bosnia dashboard metric | Hungary equivalent | Status | Source |
|---|---|---|---|
| NOS BiH 24h dispatch (plan vs actual gen/load) | MAVIR 15-min load actual/plan, generation by fuel, interconnector flows | LIVE (actual + plan) | rtdwweb.mavir.hu chart export |
| Regional / day-ahead prices | HUPX day-ahead 15-min (energy-charts, CC BY 4.0) + ENTSO-E A44 latest day | LIVE (wholesale only) | api.energy-charts.info, ENTSO-E |
| ENTSO-E load forecast, physical flows | Same datasets for HU and 7 neighbours | LIVE (14 directions x 31 days; Ukraine uses area 10Y1001C--00003F) | ENTSO-E |
| Balancing prices | MAVIR aFRR/mFRR/FRR offered capacity (MW). Prices (EUR/MW) not in this feed | LIVE (capacity); price NONE yet | MAVIR |
| Aggregators table (23 firms, GWh bought/sold) | MEKH register of 100 aggregators (names, decision no., city). Volumes are not published | LIVE register; volumes NONE | MEKH licensed-entity XLSX |
| Balance responsible parties | MAVIR BRP register, 198 parties (not aggregators) | LIVE | mavir.hu EIC register |
| BESS / RTB brownfield pipeline (33 projects) | MEKH storage licences: 213 rows, nominal MW and MWh. Licensed, not necessarily operating | LIVE | MEKH |
| Generation fleet assets (22 plants, GWh) | MEKH XLSX also has sheets Nagyeromuvek (large plants) and Kiseromuvek; MAVIR generation by fuel is live. Per-plant GWh not published | RESEARCHED (sheets present, feed not built) | MEKH |
| 154 substations, MVA, queue MW, free MW | Hosting-capacity tables published as PDF/XLSX by MAVIR (2nd publication procedure) and DSOs (E.ON, MVM). Not machine-readable feeds | RESEARCHED, not built | mavir.hu/web/mavir/masodik-kozzeteteli-eljaras ; eon.hu szabad_kapacitas_ELMU.pdf ; mvmhalozat.hu attachments/41914 |
| PV yield / GHI per substation | No Hungarian public per-substation series found | NONE | |
| End-user tariffs, household / non-household | Eurostat nrg_pc_204/205 semi-annual, all taxes | LIVE (separate from wholesale) | Eurostat |
| Tariff surge series 2020-2026 | Same Eurostat series, 2020-S2 to 2026-S1 | LIVE | Eurostat |
| Macro: GDP, FX, CPI, population | IMF WEO, MNB, Eurostat HICP, Eurostat/World Bank | LIVE | as named |
| Monthly generation (Generation_Monthly) | MAVIR generation is 15-min and covers 31 days; monthly rollup not built | NONE yet | |

Unit and aggregation notes
- Price KPIs: "latest" is the last non-empty 15-minute interval; the 31-day average is the plain mean of saved
  15-minute prices on complete days (equal weights, no volume weighting). A day the source omitted is left out whole.
- Load and generation are MW at the interval, not MWh.
- Wholesale spot (HUPX) and customer tariffs (Eurostat) are never mixed in a KPI.

Dashboard files: build_dashboard.py (builds Hungary_Energy_Dashboard.xlsx with 6 pages), Hungary_Module.bas (VBA refresh).
Install: open the workbook in Excel, Alt+F11, File > Import File > Hungary_Module.bas, save as .xlsm, run RefreshHungary.
Limitation: the VBA was written and reviewed without desktop Excel and has not been executed. The workbook was rendered
and inspected through LibreOffice only.
