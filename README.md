# hun-energy-intelligence

Automated daily data feed for the Hungary Solar PV, BESS & Grid investment-screening workbook.

Hungarian counterpart of [bih-energy-intelligence](https://github.com/georgesmichael1/bih-energy-intelligence): same architecture, same validation philosophy, Hungarian sources and values. No Bosnia data is reused; every figure below comes from a Hungarian (or Hungary-filtered European) source.

## What runs when

| Workflow | Schedule (UTC) | Token? | What it does |
|---|---|---|---|
| `hungary_update.yml` | every 30 min + manual "Run workflow" | Only for the ENTSO-E batch | One orchestrated run: token-free feeds always (MAVIR, HUPX-zone prices, MNB, IMF, World Bank, Eurostat, BRP register); the ENTSO-E daily batch (load forecast, official A44 prices, 31-day physical flows) runs once a day after 03:30 Europe/Budapest when the token is set |

Until the `ENTSOE_API_TOKEN` repository secret is set (Settings > Secrets and variables > Actions), the ENTSO-E datasets record SKIPPED (a warning; the run stays green). Once set, their failures turn the run red. Use the same personal ENTSO-E Transparency Platform token as the Bosnia repo. The token is never written to code, CSVs, status files or logs; it travels only as the `securityToken` query parameter.

## Data sources and what each number is

| Output | Source | What it is | What it is NOT |
|---|---|---|---|
| `mavir_load.csv` | MAVIR rtdwweb chart 7678 (Actual and Plan System Loads) | 15-min published system load: gross/net, plan vs actual/estimate | Real-time SCADA; today holds published completed intervals only |
| `mavir_generation_15min.csv` | MAVIR rtdwweb chart 12058 (net settlement metering) | 15-min actual generation per production type | Forecast; excludes prosumer/HSPP PV per MAVIR's methodology |
| `mavir_balancing_offers.csv` | MAVIR rtdwweb chart 1244886 | 15-min FRR/aFRR/mFRR offer volumes up/down | Activated energy prices; imbalance settlement prices |
| `mavir_flows_interconnectors.csv` | MAVIR rtdwweb chart 5229 | 15-min flows per interconnection + MAVIR schedules, signed as published | Directionally explicit measurements (use `entsoe_physical_flows.csv` for those) |
| `hupx_day_ahead_prices.csv` | energy-charts bzn=HU (CC BY 4.0, Bundesnetzagentur/SMARD.de) | Complete current-day HUPX-zone auction curve, 15-min intervals | A customer tariff; intraday price |
| `entsoe_day_ahead_prices.csv` | ENTSO-E Transparency Platform A44 | Official previous-day auction curve (primary owner HUPX) | CC-BY data (ENTSO-E Terms of Use apply) |
| `entsoe_load_forecast.csv` | ENTSO-E A65 | Day-ahead total load forecast, previous day | Actual load |
| `entsoe_physical_flows.csv` | ENTSO-E A11 | Raw directional MW per border per direction; net export = HU outbound minus inbound | Signed net flow; missing pairs are gaps, never zero |
| `mnb_fx.csv` | MNB official daily rates | EUR/HUF, USD/HUF with publication date | A peg (HUF floats) |
| `current_gdp.csv` | IMF WEO datamapper (HUN) | Current-year nominal GDP forecast/estimate | Finalized national accounts |
| `eurostat_hicp.csv` | Eurostat prc_hicp_manr | All-items HICP annual rate, monthly | Core inflation |
| `eurostat_tariffs.csv` | Eurostat nrg_pc_204 / nrg_pc_205 | Final household (band DC) and non-household (band IC) customer prices per kWh, semi-annual | Wholesale market prices |
| `balance_responsible_parties.csv` | MAVIR BRP register (EIC_YYYYMM.xlsx) | Every registered balance responsible party (BRP) with EIC and validity date | An aggregator register: BRPs settle imbalances; only some are aggregators or traders. MAVIR accredits aggregators separately |
| `hun_data.csv` | this repo | Master sheet, same schema as the Bosnia master; live rows auto-synced | Fully verified master: see `status.json` per-dataset state |

## Validation philosophy (same as Bosnia)

- Strict source contract checks: headers, units, dates, interval grids, licence strings. A changed source shape fails loudly instead of mislabeling data.
- A failed fetch never erases previously saved intervals or days.
- No typed-in fallback values anywhere. If a source is down, the dataset records FAILED and the run goes red.
- Wholesale market prices and customer tariffs are separate datasets and are never mixed.
- Forecast/plan columns are labeled as such; actuals are published measurements, not real-time telemetry.
- `status.json` carries per-dataset state, source, coverage and last success. The 30-minute workflow fails if any dataset is not OK (STATIC_REFERENCE rows excepted).

## Offline tests

`python -m unittest discover -s tests -v` runs 22 tests with no network: MAVIR export parsing (header/grid/duplicate/future-interval rejection), MNB page parsing, BRP register parsing, HUPX day validation (including 23/25-hour DST days), area/direction wiring and a no-token-in-code check.

## Repo layout

- `update_hun.py` - 30-minute orchestrator (token-free sources) + master sync
- `mavir_adapter.py` - rtdwweb chart export client (load, generation, balancing, flows)
- `hupx_prices.py` - energy-charts HU day-ahead curve
- `entsoe_adapter.py` - token feeds: load forecast, physical flows, A44 prices
- `physical_flow_history.py` - rolling 31-day flow history for 14 directions
- `mnb_fx.py`, `eurostat_feeds.py`, `imf_gdp.py`, `brp_feed.py` - macro/FX/tariff/register feeds
- `pipeline_status.py`, `price_xml_parser.py` - shared plumbing (identical to Bosnia)
- `tests/` - offline unit tests


## Dashboard
See dashboard/COVERAGE.md for the metric coverage matrix against the Bosnia framework, build_dashboard.py for the
Excel workbook and Hungary_Module.bas for the VBA refresh. SOURCE_PARTIAL means the upstream source published nothing for
a day or direction; it is shown as a warning, never counted OK, and is refetched by later runs (price_hu on every run,
ENTSO-E physical flows in the daily batch) until it fills or the day leaves the 31-day window.
