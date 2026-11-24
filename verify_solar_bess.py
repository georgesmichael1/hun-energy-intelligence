"""Residual/constraint verification across every scenario and hour of the
solar+BESS backtest. Re-runs the full grid and checks the raw dispatch:
SOC dynamics, bounds, exclusivity, grid limit, energy balance (no phantom
import), and accounting reconciliation. Writes solar_bess_verification.json.
"""
import json
import numpy as np
from solar_bess_model import (SITES, PV_MWP_OPTIONS, BESS_HOURS_OPTIONS, STRATEGIES,
                              aligned_year, local_day_slices, run_scenario,
                              ETA_CHARGE, ETA_DISCHARGE, DEGRADATION_EUR_PER_MWH)

worst = {"soc_low": 0.0, "soc_high": 0.0, "exclusivity": 0.0, "grid": 0.0,
         "phantom_import": 0.0, "soc_dynamics": 0.0, "soc_neutrality": 0.0,
         "accounting": 0.0, "pv_balance": 0.0}
hours_checked = 0
for site in SITES:
    hours, pv_w, price = aligned_year(site, ".")
    slices = local_day_slices(hours)
    for pv_mwp in PV_MWP_OPTIONS:
        for bh in BESS_HOURS_OPTIONS:
            for strat in STRATEGIES:
                r = run_scenario(site, pv_mwp, bh, strat, hours, pv_w, price, day_slices=slices)
                B, E, G = r["bess_power_mw"], r["bess_energy_mwh"], float(pv_mwp)
                acc_g = acc_d = 0.0
                for d in r["days"]:
                    ch = d["ch_pv"] + d["ch_grid"]; dis = d["dis"]; soc = d["soc"]; net = d["net"]
                    n = len(ch)
                    worst["exclusivity"] = max(worst["exclusivity"], float(np.min(np.stack([ch, dis]), axis=0).max()))
                    worst["soc_low"] = max(worst["soc_low"], float(-soc.min()))
                    worst["soc_high"] = max(worst["soc_high"], float(soc.max() - E))
                    worst["grid"] = max(worst["grid"], float(np.abs(net).max() - G))
                    worst["phantom_import"] = max(worst["phantom_import"], float((-net - d["ch_grid"] - d["ch_pv"]).max()))
                    dyn = soc - np.concatenate([[0.0], soc[:-1]]) - ETA_CHARGE * ch + dis / ETA_DISCHARGE
                    worst["soc_dynamics"] = max(worst["soc_dynamics"], float(np.abs(dyn).max()))
                    worst["soc_neutrality"] = max(worst["soc_neutrality"], float(abs(soc[-1])))
                    pvb = d["pv"] - d["curt"] - (net + d["ch_pv"] + d["ch_grid"] - dis)
                    worst["pv_balance"] = max(worst["pv_balance"], float(np.abs(pvb).max()))
                    acc_g += d["gross_eur"]; acc_d += d["degradation_eur"]
                    hours_checked += n
                worst["accounting"] = max(worst["accounting"],
                    abs(acc_g - r["gross_eur"]) + abs(acc_d - r["degradation_eur"]),
                    abs((r["gross_eur"] - r["degradation_eur"]) - r["net_eur"]))
out = {"scenarios": len(SITES) * len(PV_MWP_OPTIONS) * len(BESS_HOURS_OPTIONS) * len(STRATEGIES),
       "hours_checked": hours_checked, "max_abs_violations": worst,
       "note": "all values must be ~0 within solver tolerance; MW/MWh units per row"}
with open("solar_bess_verification.json", "w") as f:
    json.dump(out, f, indent=2)
print(json.dumps(out, indent=2))
