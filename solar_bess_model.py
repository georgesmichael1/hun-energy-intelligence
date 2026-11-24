"""Project-level solar + BESS dispatch model (ex-post price-taker backtest).

WHAT THIS IS
------------
A perfectly-foresight (hindsight) MILP optimiser that co-optimises a PV plant
and a battery energy storage system against one matched historical year of
Hungarian day-ahead market prices. It answers: "for this site, plant size and
battery configuration, what is the theoretical upper-bound DAM revenue a
price-taker could have captured in the backtest year?" It is NOT an earnings
forecast: realised prices are not known at DAM bid submission (HUPX gate
closure 12:00 D-1, results ~13:00 D-1), so an implementable strategy would
earn less. Results are labelled accordingly everywhere they appear.

DATA SOURCES AND TIMESTAMP CONVENTION (verified, see README)
------------------------------------------------------------
- PV profile: JRC PVGIS v5_3 seriescalc (PVGIS-SARAH3 satellite irradiance,
  ERA5 meteo), hourly records whose timestamp is the INSTANTANEOUS satellite
  sample time (PVGIS user manual section 9.3), not an hourly average. Each
  record is assigned to the UTC clock hour CONTAINING its sample (stamp
  truncated to the hour); no interval shift is applied. The instantaneous-
  versus-hourly-average approximation is a documented JRC data limitation;
  its impact on revenue is bounded by an alignment sensitivity test
  (tests/test_solar_bess.py).
- Prices: Energy-Charts API day-ahead prices, HU bidding zone (source
  Bundesnetzagentur | SMARD.de, licence CC BY 4.0, attribution required),
  hourly UTC interval starts, EUR/MWh.
- Matched year: the Budapest local delivery year 2023 spans UTC
  2022-12-31 23:00 through 2023-12-31 22:00 (8760 delivery hours, DST-safe in
  UTC). PVGIS is fetched for 2022-2023 and sliced to exactly that window;
  completeness (no missing hour) is asserted, never assumed.

MODEL
-----
Daily MILP (the DAM is a daily auction; interday flexibility is a known
limitation of this daily decomposition):
  max  sum_t price_t * net_t  -  c_deg * sum_t dis_t
  s.t. ch_pv_t + ch_grid_t <= B * z_t        (charge limit, exclusivity link)
       dis_t <= B * (1 - z_t)                (discharge limit, exclusivity)
       soc_t = soc_{t-1} + eta_c*(ch_pv_t+ch_grid_t) - dis_t/eta_d
       0 <= soc_t <= E ;  soc_0 = soc_N = 0  (daily SOC neutrality)
       net_t = pv_t - curt_t - ch_pv_t + dis_t - ch_grid_t
       -G <= net_t <= G                       (single net grid position)
       curt_t free                            (free PV curtailment)
       ch_grid_t = 0 under the "pv_only" strategy
  z_t binary => physically impossible same-interval charge+discharge cycling
  at negative prices is excluded by construction (not merely made unprofitable
  by round-trip losses). Intertemporal negative-price arbitrage (charge at
  p<0, discharge later) remains allowed and is covered by tests.

ACCOUNTING (no double counting)
-------------------------------
The objective already subtracts degradation. Reported lines are separate:
gross energy revenue (price x net position), degradation cost (throughput),
fixed O&M (workbook, editable), net = gross - degradation - O&M. NPV/IRR in
the workbook use only clearly-labelled editable assumptions.
"""
import csv
import datetime as dt
import json
import math
import urllib.request
from pathlib import Path

import numpy as np
from scipy.optimize import milp, LinearConstraint, Bounds

UTC = dt.timezone.utc
BACKTEST_YEAR = 2023

SITES = {
    "Szeged": (46.25, 20.15),
    "Budapest": (47.50, 19.05),
    "Nyiregyhaza": (47.95, 21.72),
}
PV_MWP_OPTIONS = (1, 10, 50)
BESS_HOURS_OPTIONS = (0, 2, 4)
STRATEGIES = ("pv_only", "grid_pv")
BESS_POWER_RATIO = 0.5          # BESS power (MW) per PV MWp
ETA_CHARGE = 0.95
ETA_DISCHARGE = 0.95
DEGRADATION_EUR_PER_MWH = 10.0  # per MWh discharged; labelled assumption
PVGIS_SLOPE_DEG = 35
PVGIS_AZIMUTH_DEG = 0           # 0 = south in PVGIS convention
PVGIS_SYSTEM_LOSS_PCT = 14

PVGIS_URL = "https://re.jrc.ec.europa.eu/api/v5_3/seriescalc"
PRICE_URL = "https://api.energy-charts.info/price"
PRICE_LICENCE = "CC BY 4.0 (creativecommons.org/licenses/by/4.0) from Bundesnetzagentur | SMARD.de"


# ---------------------------------------------------------------- data fetch
def _get_json(url, timeout=120):
    req = urllib.request.Request(url, headers={"User-Agent": "hun-energy-intelligence/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def fetch_pvgis(lat, lon, startyear, endyear):
    """Hourly PVGIS records for 1 kWp (W) with explicit, identical settings.

    Uses pvcalculation=1: the pvcalc alias is silently ignored by the API
    (verified: response then omits the P column and nulls pv_module fields).
    Raises unless every record carries P and the echoed settings match.
    """
    url = (f"{PVGIS_URL}?lat={lat}&lon={lon}&pvcalculation=1&peakpower=1"
           f"&loss={PVGIS_SYSTEM_LOSS_PCT}&angle={PVGIS_SLOPE_DEG}&aspect={PVGIS_AZIMUTH_DEG}"
           f"&startyear={startyear}&endyear={endyear}&outputformat=json")
    d = _get_json(url)
    pv = d["inputs"].get("pv_module", {})
    ms = d["inputs"]["mounting_system"]["fixed"]
    if pv.get("peak_power") != 1.0 or pv.get("system_loss") != float(PVGIS_SYSTEM_LOSS_PCT):
        raise ValueError("PVGIS pv_module settings not echoed as requested: " + json.dumps(pv))
    if ms["slope"]["value"] != PVGIS_SLOPE_DEG or ms["azimuth"]["value"] != PVGIS_AZIMUTH_DEG:
        raise ValueError("PVGIS mounting settings not echoed as requested: " + json.dumps(ms))
    hourly = d["outputs"]["hourly"]
    if not hourly or any("P" not in h for h in hourly):
        raise ValueError("PVGIS response missing P column")
    return hourly


def stamp_to_utc_hour(stamp):
    """PVGIS 'yyyymmdd:HHMM' sample stamp -> UTC clock hour containing the sample."""
    t = dt.datetime.strptime(stamp, "%Y%m%d:%H%M").replace(tzinfo=UTC)
    return t.replace(minute=0, second=0, microsecond=0)


def pv_by_utc_hour(hourly):
    out = {}
    for h in hourly:
        out[stamp_to_utc_hour(h["time"])] = float(h["P"])  # W for 1 kWp
    return out


def fetch_prices_year(year):
    """Full local delivery year's hourly DAM prices (EUR/MWh), keyed by UTC start.

    The API returns UTC unix seconds; requesting start=<year>-01-01 to the next
    year's 01-01 (inclusive end day) yields the local year plus an extra day,
    which is trimmed. Gaps, non-hourly steps or nulls raise.
    """
    url = f"{PRICE_URL}?bzn=HU&start={year}-01-01&end={year + 1}-01-01"
    d = _get_json(url)
    if d.get("unit") != "EUR / MWh":
        raise ValueError("unexpected price unit: " + str(d.get("unit")))
    ts, price = d["unix_seconds"], d["price"]
    if len(ts) != len(price) or not ts:
        raise ValueError("malformed price series")
    for i in range(len(ts) - 1):
        if ts[i + 1] - ts[i] != 3600:
            raise ValueError(f"non-hourly step at unix {ts[i]}")
    if any(p is None for p in price):
        raise ValueError("null price in series")
    first = dt.datetime.fromtimestamp(ts[0], UTC)
    hours = {}
    for t, p in zip(ts, price):
        hours[dt.datetime.fromtimestamp(t, UTC)] = float(p)
    # local year = 8760 hours starting at local midnight of Jan 1 (CET, UTC+1)
    start = dt.datetime(year, 1, 1, 0, 0) - dt.timedelta(hours=1)
    start = start.replace(tzinfo=UTC)
    out = {}
    for i in range(8760):
        h = start + dt.timedelta(hours=i)
        if h not in hours:
            raise ValueError("price hour missing for local year: " + h.isoformat())
        out[h] = hours[h]
    if first > start:
        raise ValueError("price series does not cover the local year start")
    return out


def aligned_year(site, cache_dir, year=BACKTEST_YEAR, fetch=True):
    """(hours_utc[8760], pv_w_per_kwp[8760], price_eur_mwh[8760]) fully matched.

    Reads cached CSVs when present; fetches live otherwise. Completeness is
    asserted: 8760 distinct UTC hours, each with both a PV and a price value.
    """
    cache_dir = Path(cache_dir)
    pv_csv = cache_dir / f"pv_{site.lower()}_{year}_hourly.csv"
    pr_csv = cache_dir / f"prices_hu_{year}_hourly.csv"
    if pv_csv.exists() and pr_csv.exists():
        with pv_csv.open() as f:
            pv = {dt.datetime.fromisoformat(r["utc_hour_start"]): float(r["pv_w_per_kwp"])
                  for r in csv.DictReader(f)}
        with pr_csv.open() as f:
            pr = {dt.datetime.fromisoformat(r["utc_hour_start"]): float(r["price_eur_mwh"])
                  for r in csv.DictReader(f)}
    elif fetch:
        lat, lon = SITES[site]
        pv = pv_by_utc_hour(fetch_pvgis(lat, lon, year - 1, year))
        pr = fetch_prices_year(year)
        write_cache(site, year, pv, pr, cache_dir)
    else:
        raise FileNotFoundError(f"cache missing for {site} {year}")
    start = dt.datetime(year, 1, 1, tzinfo=UTC) - dt.timedelta(hours=1)
    hours = [start + dt.timedelta(hours=i) for i in range(8760)]
    missing_pv = [h for h in hours if h not in pv]
    missing_pr = [h for h in hours if h not in pr]
    if missing_pv or missing_pr:
        raise ValueError(f"alignment incomplete: {len(missing_pv)} pv, {len(missing_pr)} price hours missing")
    return hours, np.array([pv[h] for h in hours]), np.array([pr[h] for h in hours])


def write_cache(site, year, pv, pr, cache_dir):
    cache_dir = Path(cache_dir)
    with (cache_dir / f"pv_{site.lower()}_{year}_hourly.csv").open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["utc_hour_start", "pv_w_per_kwp", "source", "convention"])
        for h in sorted(pv):
            w.writerow([h.isoformat(), f"{pv[h]:.2f}",
                        f"PVGIS v5_3 SARAH3+ERA5 slope{PVGIS_SLOPE_DEG} az{PVGIS_AZIMUTH_DEG} loss{PVGIS_SYSTEM_LOSS_PCT}",
                        "instantaneous sample assigned to containing UTC hour"])
    with (cache_dir / f"prices_hu_{year}_hourly.csv").open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["utc_hour_start", "price_eur_mwh", "licence"])
        for h in sorted(pr):
            w.writerow([h.isoformat(), f"{pr[h]:.2f}", PRICE_LICENCE])


# ------------------------------------------------------------------- MILP
def dispatch_day(pv_mw, price_eur_mwh, energy_mwh, power_mw, grid_limit_mw,
                 strategy, eta_c=ETA_CHARGE, eta_d=ETA_DISCHARGE,
                 c_deg=DEGRADATION_EUR_PER_MWH):
    """One day's optimal dispatch. Returns dict of per-hour arrays + revenue split.

    pv_mw, price_eur_mwh: length-n arrays (n hours). Raises on infeasible input.
    """
    pv = np.asarray(pv_mw, dtype=float)
    price = np.asarray(price_eur_mwh, dtype=float)
    n = len(pv)
    if len(price) != n:
        raise ValueError("pv/price length mismatch")
    grid_charge = 1.0 if strategy == "grid_pv" else 0.0
    B, E, G = float(power_mw), float(energy_mwh), float(grid_limit_mw)

    # variables: ch_pv[0:n] ch_grid[n:2n] dis[2n:3n] net[3n:4n] z[4n:5n] soc[5n:6n+1]
    N = 6 * n + 1
    CHPV, CHGR, DIS, NET, Z, SOC = 0, n, 2 * n, 3 * n, 4 * n, 5 * n
    c = np.zeros(N)
    c[NET:NET + n] = -price          # maximise sum price*net
    c[DIS:DIS + n] = c_deg           # degradation cost on discharge throughput

    lb = np.zeros(N)
    ub = np.full(N, np.inf)
    ub[CHPV:CHPV + n] = B
    ub[CHGR:CHGR + n] = B * grid_charge
    ub[DIS:DIS + n] = B
    lb[NET:NET + n] = -G
    ub[NET:NET + n] = G
    ub[Z:Z + n] = 1.0
    ub[SOC:SOC + n + 1] = E
    integrality = np.zeros(N)
    integrality[Z:Z + n] = 1
    bounds = Bounds(lb, ub)

    rows, lo, hi = [], [], []

    def add(coefs, l, u):
        row = np.zeros(N)
        for k, v in coefs:
            row[k] = v
        rows.append(row)
        lo.append(l)
        hi.append(u)

    for t in range(n):
        # exclusivity: ch_pv + ch_grid <= B*z  ;  dis <= B*(1-z)
        add([(CHPV + t, 1), (CHGR + t, 1), (Z + t, -B)], -np.inf, 0)
        add([(DIS + t, 1), (Z + t, B)], -np.inf, B)
        # soc dynamics: soc_t - soc_{t-1} - eta_c*ch_t + dis_t/eta_d = 0
        add([(SOC + t + 1, 1), (SOC + t, -1),
             (CHPV + t, -eta_c), (CHGR + t, -eta_c), (DIS + t, 1.0 / eta_d)], 0, 0)
        # net position: net = pv_used - ch_pv + dis - ch_grid with pv_used in [0, pv]
        # (curtailment is the slack); both bounds are needed: the upper bound caps
        # PV use at availability, the lower bound forbids phantom import/dump load -
        # every imported MWh must physically enter the battery.
        add([(NET + t, 1), (CHPV + t, 1), (CHGR + t, 1), (DIS + t, -1)], -np.inf, pv[t])
        add([(NET + t, 1), (CHPV + t, 1), (CHGR + t, 1), (DIS + t, -1)], 0, np.inf)
        # pv_used >= ch_pv: charging the battery from PV requires that much PV
        # actually used; without this, phantom import (net < 0 with no storage
        # inflow) earns money at negative prices.
        add([(NET + t, 1), (CHGR + t, 1), (DIS + t, -1)], 0, np.inf)
        # PV used cannot exceed available: ch_pv + (export part) ... implied by
        # curt >= 0 combined with net >= -G; add explicit pv cap on ch_pv:
        add([(CHPV + t, 1)], -np.inf, pv[t])
    # daily SOC neutrality: soc_0 = 0 and soc_n = 0
    add([(SOC, 1)], 0, 0)
    add([(SOC + n, 1)], 0, 0)

    A = np.array(rows)
    res = milp(c=c, integrality=integrality, bounds=bounds,
               constraints=LinearConstraint(A, np.array(lo), np.array(hi)),
               options={"time_limit": 60})
    if res.status not in (0, 1):
        raise ValueError(f"MILP failed: status={res.status} {res.message}")
    x = res.x
    ch_pv, ch_grid, dis = x[CHPV:CHPV + n], x[CHGR:CHGR + n], x[DIS:DIS + n]
    net, z, soc = x[NET:NET + n], x[Z:Z + n], x[SOC:SOC + n + 1]
    pv_used = net + ch_pv + ch_grid - dis
    curt = pv - pv_used
    gross = float(np.dot(price, net))
    degradation = float(c_deg * dis.sum())
    return {
        "ch_pv": ch_pv, "ch_grid": ch_grid, "dis": dis, "net": net, "z": z,
        "soc": soc[1:], "curt": curt, "pv": pv, "price": price,
        "gross_eur": gross, "degradation_eur": degradation,
        "net_eur": gross - degradation,
        "discharge_mwh": float(dis.sum()), "charge_mwh": float((ch_pv + ch_grid).sum()),
        "curtail_mwh": float(curt.sum()),
    }


# ------------------------------------------------------------------ backtest
def local_day_slices(hours, year=BACKTEST_YEAR):
    """Group UTC hours into Budapest local delivery days (23/24/25 h, DST-safe)."""
    from zoneinfo import ZoneInfo
    bud = ZoneInfo("Europe/Budapest")
    slices, start = [], 0
    for i in range(1, len(hours) + 1):
        if i == len(hours) or hours[i].astimezone(bud).date() != hours[start].astimezone(bud).date():
            slices.append((start, i))
            start = i
    assert len(slices) == 365 if year == 2023 else True
    return slices


def run_scenario(site, pv_mwp, bess_hours, strategy, hours, pv_w_per_kwp, price,
                 day_slices=None):
    """Yearly backtest for one configuration; returns aggregates + per-day rows."""
    pv_mw = pv_w_per_kwp * pv_mwp / 1000.0          # W per kWp -> MW for the plant
    B = BESS_POWER_RATIO * pv_mwp if bess_hours > 0 else 0.0
    E = bess_hours * B
    G = float(pv_mwp)                                # grid limit = PV MWp (labelled assumption)
    if day_slices is None:
        day_slices = local_day_slices(hours)
    days = []
    for i, (a, b) in enumerate(day_slices):
        r = dispatch_day(pv_mw[a:b], price[a:b], E, B, G, strategy)
        days.append(r)
    gross = sum(d["gross_eur"] for d in days)
    deg = sum(d["degradation_eur"] for d in days)
    pv_energy = float(pv_mw[:(day_slices[-1][1])].sum())
    return {
        "site": site, "pv_mwp": pv_mwp, "bess_hours": bess_hours,
        "bess_power_mw": B, "bess_energy_mwh": E, "strategy": strategy,
        "pv_energy_mwh": pv_energy,
        "gross_eur": gross, "degradation_eur": deg, "net_eur": gross - deg,
        "discharge_mwh": sum(d["discharge_mwh"] for d in days),
        "curtail_mwh": sum(d["curtail_mwh"] for d in days),
        "capture_eur_per_mwh": (gross / pv_energy) if pv_energy > 0 else None,
        "days": days,
    }


def representative_days(days, n_types=("top_revenue", "most_negative", "mid_summer")):
    """Pick representative day indices: highest net revenue, most negative-price
    hours, and the midsummer day containing June 15 (index 165)."""
    top = max(range(len(days)), key=lambda i: days[i]["net_eur"])
    neg = max(range(len(days)), key=lambda i: int((days[i]["price"] < 0).sum()))
    picks = {"top_revenue": top, "most_negative": neg, "mid_summer": 165}
    return picks


# ------------------------------------------------------------- CSV outputs
SCENARIOS_CSV = "solar_bess_scenarios.csv"
DISPATCH_CSV = "solar_bess_dispatch_hourly.csv"

SCEN_HEADER = ["Scenario_Key", "Site", "PV_MWp", "BESS_Hours", "BESS_Power_MW",
               "BESS_Energy_MWh", "Strategy", "PV_Energy_MWh", "Capture_EUR_per_MWh",
               "Gross_EUR", "Degradation_EUR", "Net_EUR", "Discharge_MWh",
               "Curtail_MWh", "Year", "Basis"]
DISP_HEADER = ["Scenario_Key", "Day_Type", "Delivery_Date", "Hour_Local",
               "PV_MW", "Charge_MW", "Discharge_MW", "SOC_MWh", "Net_MW",
               "Curtailed_MW", "Price_EUR_MWh", "Row_Key"]


def scenario_key(site, pv_mwp, bess_hours, strategy):
    return f"{site}|{pv_mwp}|{bess_hours}|{strategy}"


def update_solar_bess(cache_dir="."):
    """Run the full scenario grid on cached/fetched data; write the two CSVs.
    Returns (scenario_rows, dispatch_rows). Fetches live only when the cache
    is absent, so the daily CI run costs seconds once CSVs are committed."""
    from zoneinfo import ZoneInfo
    bud = ZoneInfo("Europe/Budapest")
    scen_rows, disp_rows = [], []
    aligned = {}
    for site in SITES:
        aligned[site] = aligned_year(site, cache_dir)
    hours0, _, price = aligned["Szeged"]
    for site in SITES:
        hours, pv, _ = aligned[site]
        slices = local_day_slices(hours)
        for pv_mwp in PV_MWP_OPTIONS:
            for bess_hours in BESS_HOURS_OPTIONS:
                for strategy in STRATEGIES:
                    r = run_scenario(site, pv_mwp, bess_hours, strategy, hours, pv, price,
                                     day_slices=slices)
                    key = scenario_key(site, pv_mwp, bess_hours, strategy)
                    scen_rows.append({
                        "Scenario_Key": key, "Site": site, "PV_MWp": pv_mwp,
                        "BESS_Hours": bess_hours, "BESS_Power_MW": round(r["bess_power_mw"], 3),
                        "BESS_Energy_MWh": round(r["bess_energy_mwh"], 3), "Strategy": strategy,
                        "PV_Energy_MWh": round(r["pv_energy_mwh"], 1),
                        "Capture_EUR_per_MWh": round(r["capture_eur_per_mwh"], 2) if r["capture_eur_per_mwh"] else "",
                        "Gross_EUR": round(r["gross_eur"], 0),
                        "Degradation_EUR": round(r["degradation_eur"], 0),
                        "Net_EUR": round(r["net_eur"], 0),
                        "Discharge_MWh": round(r["discharge_mwh"], 1),
                        "Curtail_MWh": round(r["curtail_mwh"], 1),
                        "Year": BACKTEST_YEAR,
                        "Basis": "ex-post price-taker upper bound, not a forecast",
                    })
                    picks = representative_days(r["days"])
                    for day_type, di in picks.items():
                        d = r["days"][di]
                        a, b = slices[di]
                        ddate = hours[a].astimezone(bud).date().isoformat()
                        for j in range(b - a):
                            local_h = hours[a + j].astimezone(bud).strftime("%Y-%m-%d %H:%M")
                            disp_rows.append({
                                "Scenario_Key": key, "Day_Type": day_type,
                                "Delivery_Date": ddate, "Hour_Local": local_h,
                                "PV_MW": round(float(d["pv"][j]), 4),
                                "Charge_MW": round(float(d["ch_pv"][j] + d["ch_grid"][j]), 4),
                                "Discharge_MW": round(float(d["dis"][j]), 4),
                                "SOC_MWh": round(float(d["soc"][j]), 4),
                                "Net_MW": round(float(d["net"][j]), 4),
                                "Curtailed_MW": round(float(d["curt"][j]), 4),
                                "Price_EUR_MWh": round(float(d["price"][j]), 2),
                                "Row_Key": f"{key}|{day_type}|{j}",
                            })
    cache = Path(cache_dir)
    with (cache / SCENARIOS_CSV).open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=SCEN_HEADER)
        w.writeheader()
        w.writerows(scen_rows)
    with (cache / DISPATCH_CSV).open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=DISP_HEADER)
        w.writeheader()
        w.writerows(disp_rows)
    return scen_rows, disp_rows
