# Metric coverage: Bosnia framework vs Hungary

Status legend: LIVE = fetched by the workflow from a public source; STATIC = fixed reference text;
RESEARCHED = a public source exists, no feed built yet; NONE = no public Hungarian equivalent found.
Actual / forecast labels are in the CSV column names (Plan, Forecast, Actual). Nothing is estimated.

| Bosnia dashboard metric | Hungary equivalent | Status | Source |
|---|---|---|---|
| NOS BiH 24h dispatch (plan vs actual gen/load) | MAVIR 15-min load actual/plan, generation by fuel, interconnector flows | LIVE (actual + plan) | rtdwweb.mavir.hu chart export |
| Regional / day-ahead prices | HUPX day-ahead 15-min (HUPX Labs primary, energy-charts CC BY 4.0 fallback) + ENTSO-E A44 latest day | LIVE (wholesale only) | labs.hupx.hu, api.energy-charts.info, ENTSO-E |
| ENTSO-E load forecast, physical flows | Same datasets for HU and 7 neighbours | LIVE (14 directions x 31 days; Ukraine uses area 10Y1001C--00003F) | ENTSO-E |
| Balancing prices | MAVIR aFRR/mFRR/FRR offered capacity (MW). Prices (EUR/MW) not in this feed | LIVE (capacity); price NONE yet | MAVIR |
| Aggregators table (23 firms, GWh bought/sold) | MEKH register of 100 aggregators (names, decision no., city). Volumes are not published | LIVE register; volumes NONE | MEKH licensed-entity XLSX |
| Balance responsible parties | MAVIR BRP register, 198 parties (not aggregators) | LIVE | mavir.hu EIC register |
| BESS / RTB brownfield pipeline (33 projects) | MEKH storage licences: 213 rows, nominal MW and MWh. Licensed, not necessarily operating | LIVE | MEKH |
| Generation fleet assets (22 plants, GWh) | MEKH XLSX also has sheets Nagyeromuvek (large plants) and Kiseromuvek; MAVIR generation by fuel is live. Per-plant GWh not published | RESEARCHED (sheets present, feed not built) | MEKH |
| Substations: free offtake MVA (5-year probable) | E.ON Hungary hosting-capacity PDF, 51 rows (EEL/EDE labels as published), parsed by substation_capacity_feed.py to hun_substation_capacity.csv | PARTIAL: E.ON areas only, MVA offtake side, quarterly. MVM Halozat, MVM Emasz (PDF has no extractable pages) and per-substation generation capacity NOT parsed | eon.hu szabad_kapacitas_ELMU.pdf |
| Connection queue (MW by technology, per substation) | MAVIR 2nd publication procedure closing publication, 653 requests, MVA per technology, rejected flag, parsed by mavir_connection_requests_feed.py to hun_connection_requests.csv | STATIC_REFERENCE: closed procedure, September 2024 file, not a live queue; MVA not MW | mavir.hu/web/mavir/masodik-kozzeteteli-eljaras |
| Installed capacity MW by technology (system level) | MAVIR Energia MIX gross installed capacity PDF, discovered from the MAVIR page at run time; 12 fuel groups x 12 periods, sum-checked against the published total | BUILT (OK): mavir_installed_capacity.csv | mavir.hu energia-mix-eromuvi-beepitett-teljesitokepesseg-adatok |
| Solar capture ratio (daily) | Derived: HUPX DAM 15-min prices weighted by MAVIR 15-min measured solar; complete Budapest days only; blank when day average <= 0 | BUILT (OK): solar_capture_daily.csv, page 3 | derived from LIVE sources |
| BESS best chronological spread (daily, 1/2/4h) | Derived: HUPX DAM 15-min prices; cheapest then most expensive non-overlapping window, charge ends before discharge starts; raw spread before efficiency, losses, fees | BUILT (OK): bess_spread_daily.csv, page 3; raw spread, NOT a revenue or hybrid uplift metric | derived from LIVE sources |
| Per-plant MW | MEKH large/small plant sheets list licensees only (no MW or technology) | NOT BUILT, source gap: per-plant MW not found in a machine-readable public feed | mekh.hu Villamosenergia_engedelyesek |
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

Dashboard files: build_dashboard.py (builds Hungary_Energy_Dashboard.xlsx with 7 pages), Hungary_Module.bas (VBA refresh).
Install: open the workbook in Excel, Alt+F11, File > Import File > Hungary_Module.bas, save as .xlsm, run RefreshHungary.
Limitation: the VBA was written and reviewed without desktop Excel and has not been executed. The workbook was rendered
and inspected through LibreOffice only.


## Solar_BESS page (2023 backtest)

| Item | Value |
| --- | --- |
| Weather | PVGIS v5_3 seriescalc, PVGIS-SARAH3 (5 km satellite) + ERA5 meteo; slope 35 deg south, 14% system loss; 2022-2023 fetched, sliced to the local-2023 UTC window |
| Timestamp convention | SARAH3 stamps are instantaneous satellite sample times (~HH:10 for Hungary); each record is assigned to the UTC hour containing the sample (JRC user manual 9.3). Instantaneous-vs-hourly-average approximation is a documented limitation; alignment sensitivity covered in tests |
| Prices | Energy-Charts API HU day-ahead, hourly UTC, EUR/MWh, licence CC BY 4.0 (source Bundesnetzagentur / SMARD.de); full local year 2023, 8760 h, zero gaps/nulls, DST-safe |
| Model | Daily MILP (scipy/HiGHS): binary charge/discharge exclusivity, net grid position, free PV curtailment, daily SOC neutrality, grid limit = PV MWp; degradation 10 EUR/MWh discharged in-objective; O&M/capex/WACC/lifetime are editable REPORTING-only assumptions, clearly flagged as unverified placeholders |
| Scope | Ex-post price-taker upper bound; not implementable as-is (DAM results known after gate closure); no intraday/imbalance/fees/taxes; interday flexibility not modelled; representative coordinates, not verified plots; 5 km substation filter separate |
| Verification | tests/test_solar_bess.py (13 tests incl. brute-force oracle); solar_bess_verification.json: 54 scenarios x 8760 h, max abs constraint violation 5e-14; workbook selectors switch among precomputed scenarios only |
