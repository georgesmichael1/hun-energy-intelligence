"""Build Hungary_Energy_Dashboard.xlsx from this repo's CSV outputs.

Usage: python dashboard/build_dashboard.py [csv_dir] [out.xlsx]
csv_dir may be a local folder (default: repo root) or a raw GitHub base URL ending in '/'.
Six dashboard pages sit in front of the raw data sheets (D_*). Every number on a page is a
formula over a data sheet, so the VBA module (Hungary_Module.bas) only has to replace the
data sheets and recalculate. Nothing is invented: a missing value stays empty, never zero.
"""
import csv, io, json, sys, datetime as dt, urllib.request
from pathlib import Path
from zoneinfo import ZoneInfo
import xlsxwriter

SRC = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent.parent)
OUT = sys.argv[2] if len(sys.argv) > 2 else "Hungary_Energy_Dashboard.xlsx"
BUD = ZoneInfo("Europe/Budapest")
NAVY, TEAL, GOLD, RED, GREY, LIGHT = "#12263F", "#1B998B", "#F2A541", "#C8553D", "#6B7A8F", "#F3F6FA"


def read(name):
    if SRC.startswith("http"):
        text = urllib.request.urlopen(SRC + name, timeout=60).read().decode("utf-8")
    else:
        text = (Path(SRC) / name).read_text(encoding="utf-8")
    return list(csv.DictReader(io.StringIO(text)))


def local(utc_text):
    return dt.datetime.fromisoformat(utc_text.replace("Z", "+00:00")).astimezone(BUD).replace(tzinfo=None)


def num(v):
    try:
        return float(v) if v not in ("", None) else None
    except ValueError:
        return None


wb = xlsxwriter.Workbook(OUT)
F = lambda **k: wb.add_format(k)
f_title = F(bold=True, font_size=20, font_color="white", bg_color=NAVY, valign="vcenter")
f_sub = F(font_size=10, font_color="white", bg_color=NAVY, valign="vcenter")
f_h = F(bold=True, font_color="white", bg_color=TEAL, border=1, text_wrap=True, valign="vcenter")
f_cell = F(border=1)
f_num = F(border=1, num_format="#,##0.0")
f_num2 = F(border=1, num_format="#,##0.0000")
f_dt = F(num_format="yyyy-mm-dd hh:mm")
f_d = F(num_format="yyyy-mm-dd", border=1)
f_kl = F(font_size=9, font_color=GREY, bg_color=LIGHT, top=2, top_color=TEAL, align="center", text_wrap=True)
f_kv = F(font_size=22, bold=True, font_color=NAVY, bg_color=LIGHT, align="center", num_format="#,##0.0")
f_kv0 = F(font_size=22, bold=True, font_color=NAVY, bg_color=LIGHT, align="center", num_format="#,##0")
f_ks = F(font_size=9, font_color=GREY, bg_color=LIGHT, align="center", text_wrap=True)
f_note = F(font_size=9, italic=True, font_color=GREY, text_wrap=True, valign="top")
f_sec = F(bold=True, font_size=12, font_color=NAVY, bottom=2, bottom_color=TEAL)

# ---------------- data sheets ----------------
price = [r for r in read("hupx_day_ahead_prices.csv")]
eprice = read("entsoe_day_ahead_prices.csv")
load = read("mavir_load.csv"); gen = read("mavir_generation_15min.csv")
flow = read("mavir_flows_interconnectors.csv"); bal = read("mavir_balancing_offers.csv")
elf = read("entsoe_load_forecast.csv")
tar = read("eurostat_tariffs.csv"); hicp = read("eurostat_hicp.csv"); fx = read("mnb_fx.csv")
brp = read("balance_responsible_parties.csv"); master = read("hun_data.csv")
try:
    agg = read("mekh_aggregators.csv"); sto = read("mekh_storage_licences.csv")
except Exception:
    agg, sto = [], []   # feed not yet committed: pages show 'no data', never zero
status = json.loads(urllib.request.urlopen(SRC + "status.json").read() if SRC.startswith("http")
                    else (Path(SRC) / "status.json").read_bytes())

pages = {}
names = ["Overview", "DayAhead", "System", "Flows_Balancing", "Tariffs_Macro", "Registers_Health"]
for n in names:
    pages[n] = wb.add_worksheet(n)


def data_sheet(name, header, rows, fmts=None):
    ws = wb.add_worksheet(name)
    for c, h in enumerate(header):
        ws.write(0, c, h, f_h)
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row):
            if isinstance(v, dt.datetime):
                ws.write_datetime(r, c, v, f_dt)
            elif v is None or v == "":
                pass                                  # missing stays blank, never zero
            else:
                ws.write(r, c, v)
    ws.set_column(0, len(header) - 1, 16)
    ws.freeze_panes(1, 0)
    return ws, len(rows)


D = {}
D["price"] = data_sheet("D_Price", ["Delivery_Date", "Local_Start", "Price_EUR_MWh", "Source"],
                        [(r["Delivery_Date"], local(r["Interval_Start_UTC"]), num(r["Price_EUR_MWh"]), "HUPX day-ahead via energy-charts (CC BY 4.0)") for r in price])
D["eprice"] = data_sheet("D_EntsoePrice", ["Delivery_Date", "Local_Start", "Price_EUR_MWh", "Source"],
                         [(r["Delivery_Date"], local(r["Interval_Start_UTC"]), num(r["Price_EUR_MWh"]), "ENTSO-E A44 (primary owner HUPX)") for r in eprice])
LC = ["Gross_Actual_System_Load_MW", "Gross_Plan_System_Load_MW", "Net_Actual_Load_Settlement_MW"]
D["load"] = data_sheet("D_Load", ["Local_Start"] + LC, [(local(r["Interval_Start_UTC"]),) + tuple(num(r[c]) for c in LC) for r in load])
GC = ["Nuclear_MW", "Lignite_MW", "Gas_MW", "Wind_Onshore_MW", "Biomass_MW", "Solar_PV_Measured_MW", "Hydro_RunOfRiver_MW", "Waste_Incineration_MW"]
D["gen"] = data_sheet("D_Gen", ["Local_Start", "Net_Domestic_Generation_Actual_MW"] + GC,
                      [(local(r["Interval_Start_UTC"]), num(r["Net_Domestic_Generation_Actual_MW"])) + tuple(num(r[c]) for c in GC) for r in gen])
FC = ["HU_AT_Actual_MW", "HU_SK_Actual_MW", "HU_RO_Actual_MW", "HU_RS_Actual_MW", "HU_HR_Actual_MW", "HU_SI_Actual_MW", "HU_UA_Actual_MW"]
D["flow"] = data_sheet("D_Flows", ["Local_Start"] + FC, [(local(r["Interval_Start_UTC"]),) + tuple(num(r[c]) for c in FC) for r in flow])
BC = ["AFRR_Up_Offers_MW", "AFRR_Down_Offers_MW", "MFRR_Up_Offers_MW", "MFRR_Down_Offers_MW"]
D["bal"] = data_sheet("D_Balancing", ["Local_Start"] + BC, [(local(r["Interval_Start_UTC"]),) + tuple(num(r[c]) for c in BC) for r in bal])
D["elf"] = data_sheet("D_LoadForecast", ["Local_Start", "Forecast_MW"], [(local(r["Interval_Start_UTC"]), num(r["Value_MW"])) for r in elf])
D["tar"] = data_sheet("D_Tariffs", ["Period", "Group", "Band", "EUR_per_kWh_all_taxes"],
                      [(r["Period"], r["Consumer_Group"], r["Consumption_Band"], num(r["Price_EUR_Per_kWh"])) for r in tar])
D["hicp"] = data_sheet("D_HICP", ["Period", "HICP_annual_pct"], [(r["Period"], num(r["HICP_Annual_Rate_Pct"])) for r in hicp])
D["fx"] = data_sheet("D_FX", ["Rate_Date", "Currency", "HUF_per_unit"], [(r["Rate_Date"], r["Currency"], num(r["HUF_Per_Unit"])) for r in fx])
D["agg"] = data_sheet("D_Aggregators", ["Aggregator (MEKH register)", "Decision_No", "City"], [(r["Aggregator"], r["Registration_Decision_No"], r["City"]) for r in agg])
D["sto"] = data_sheet("D_Storage", ["Licensee", "Site", "Power_MW", "Capacity_MWh", "Revoked"],
                      [(r["Licensee"], r["Site"], num(r["Nominal_Power_MW"]), num(r["Storage_Capacity_MWh"]), r["Revoked_Date"]) for r in sto])
D["brp"] = data_sheet("D_BRP", ["Balance responsible party (MAVIR)", "EIC", "Valid_From"], [(r["Balance_Responsible_Party"], r["EIC_Code"], r["Valid_From"]) for r in brp])
det = status["dataset_details"]
D["stat"] = data_sheet("D_Status", ["Dataset", "State", "Detail", "Last_success"],
                       [(k, v.get("state"), (v.get("detail") or "")[:240], v.get("last_success_at") or "") for k, v in det.items()])


def rng(key, col, last=True):
    ws, n = D[key]
    L = chr(ord("A") + col)
    return f"{ws.get_name()}!${L}$2:${L}${max(n, 1) + 1}"


def sheet_of(key):
    return D[key][0].get_name()


def header(ws, title, sub, page_no):
    ws.hide_gridlines(2)
    ws.set_landscape(); ws.set_paper(8); ws.fit_to_pages(1, 1)
    ws.set_column(0, 0, 2); ws.set_column(1, 14, 13)
    ws.set_row(0, 34); ws.set_row(1, 18)
    for c in range(0, 15):
        ws.write(0, c, "", f_title); ws.write(1, c, "", f_sub)
    ws.write(0, 1, "HUNGARY ENERGY  |  " + title, f_title)
    ws.write(1, 1, sub, f_sub)
    for i, n in enumerate(names):
        ws.write_url(2, 1 + i * 2, f"internal:'{n}'!A1", F(font_size=9, bold=(n == ws.get_name()), font_color=TEAL if n != ws.get_name() else NAVY, underline=0 if n == ws.get_name() else 1), string=f"{i+1}. {n.replace('_', ' & ')}")


def kpi(ws, row, col, label, formula, fmt, sub, val=0):
    ws.merge_range(row, col, row, col + 1, label, f_kl)
    ws.merge_range(row + 1, col, row + 1, col + 1, "", fmt); ws.write_formula(row + 1, col, formula, fmt, val)
    ws.merge_range(row + 2, col, row + 2, col + 1, sub, f_ks)
    ws.set_row(row + 1, 34); ws.set_row(row + 2, 26)


def chart(ws, cell, title, kind="line", series=(), ytitle="", size=(560, 280), sub=None, stacked=False, xint=480, xfmt="dd mmm"):
    ch = wb.add_chart({"type": kind, **({"subtype": "stacked"} if stacked else {})})
    for nm, cats, vals, color in series:
        s = {"name": nm, "categories": cats, "values": vals}
        if kind == "line":
            s["line"] = {"color": color, "width": 1.25}
        else:
            s["fill"] = {"color": color}
        ch.add_series(s)
    ch.set_title({"name": title, "name_font": {"size": 11, "color": NAVY}})
    ch.set_y_axis({"name": ytitle, "name_font": {"size": 9}, "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#DDDDDD"}}})
    ch.set_x_axis({"text_axis": True, "label_position": "low", "num_font": {"size": 8, "rotation": -30}, "interval_unit": xint, "interval_tick": xint, "num_format": xfmt} if kind != "column" else {"num_font": {"size": 8}})
    ch.show_blanks_as("gap")
    ch.set_legend({"position": "bottom", "font": {"size": 8}})
    ch.set_size({"width": size[0], "height": size[1]})
    ws.insert_chart(cell, ch)


def last(rows, key):
    v = [num(r[key]) for r in rows if num(r[key]) is not None]
    return v[-1] if v else "no data"


def mean(rows, key):
    v = [num(r[key]) for r in rows if num(r[key]) is not None]
    return sum(v) / len(v) if v else "no data"


def eur_huf():
    v = [num(r["HUF_Per_Unit"]) for r in fx if r["Currency"] == "EUR"]
    return v[0] if v else "no data"


pal = [NAVY, TEAL, GOLD, RED, "#7B6D8D", "#4A90A4", "#8CB369", GREY]
PS, PN = sheet_of("price"), D["price"][1]
# ---------- 1 Overview ----------
ws = pages["Overview"]
header(ws, "EXECUTIVE OVERVIEW", "Wholesale market, system and register snapshot. Every figure is a live formula over the data sheets.", 1)
kpi(ws, 4, 1, "LATEST HUPX DAY-AHEAD (EUR/MWh)", f"=IFERROR(LOOKUP(2,1/({PS}!$C$2:$C${PN+1}<>\"\"),{PS}!$C$2:$C${PN+1}),\"no data\")", f_kv, "wholesale spot, latest published 15-min interval", last(price, "Price_EUR_MWh"))
kpi(ws, 4, 3, "31-DAY AVERAGE (EUR/MWh)", f"=IFERROR(AVERAGE({PS}!$C$2:$C${PN+1}),\"no data\")", f_kv, "complete saved days only; omitted days are not zero", mean(price, "Price_EUR_MWh"))
kpi(ws, 4, 5, "LATEST GROSS SYSTEM LOAD (MW)", f"=IFERROR(LOOKUP(2,1/(D_Load!$B$2:$B${D['load'][1]+1}<>\"\"),D_Load!$B$2:$B${D['load'][1]+1}),\"no data\")", f_kv0, "MAVIR actual, 15-min", last(load, "Gross_Actual_System_Load_MW"))
kpi(ws, 4, 7, "LATEST NET GENERATION (MW)", f"=IFERROR(LOOKUP(2,1/(D_Gen!$B$2:$B${D['gen'][1]+1}<>\"\"),D_Gen!$B$2:$B${D['gen'][1]+1}),\"no data\")", f_kv0, "MAVIR actual domestic", last(gen, "Net_Domestic_Generation_Actual_MW"))
kpi(ws, 4, 9, "NUCLEAR SHARE, LATEST", f"=IFERROR(LOOKUP(2,1/(D_Gen!$C$2:$C${D['gen'][1]+1}<>\"\"),D_Gen!$C$2:$C${D['gen'][1]+1})/LOOKUP(2,1/(D_Gen!$B$2:$B${D['gen'][1]+1}<>\"\"),D_Gen!$B$2:$B${D['gen'][1]+1}),\"no data\")", F(font_size=22, bold=True, font_color=NAVY, bg_color=LIGHT, align="center", num_format="0%"), "Paks, share of net generation", (last(gen, "Nuclear_MW") / last(gen, "Net_Domestic_Generation_Actual_MW")))
kpi(ws, 4, 11, "EUR / HUF (MNB)", "=IFERROR(INDEX(D_FX!$C$2:$C$10,MATCH(\"EUR\",D_FX!$B$2:$B$10,0)),\"no data\")", f_kv, "official daily rate", eur_huf())
chart(ws, "B9", "HUPX day-ahead price, 15-minute (EUR/MWh)", series=[("HUPX day-ahead", f"={PS}!$B$2:$B${PN+1}", f"={PS}!$C$2:$C${PN+1}", TEAL)], ytitle="EUR/MWh", size=(600, 290))
chart(ws, "I9", "MAVIR gross system load, actual vs plan (MW)", series=[("Actual", f"=D_Load!$A$2:$A${D['load'][1]+1}", f"=D_Load!$B$2:$B${D['load'][1]+1}", NAVY), ("Plan", f"=D_Load!$A$2:$A${D['load'][1]+1}", f"=D_Load!$C$2:$C${D['load'][1]+1}", GOLD)], ytitle="MW", size=(600, 290))
ws.merge_range("B25:O26", "Honesty rules: wholesale spot prices (HUPX) are never mixed with customer tariffs (Eurostat, page 5). A day the source did not publish is left out whole, never zero or estimated. Dataset states are on page 6.", f_note)

# ---------- 2 Day-ahead ----------
ws = pages["DayAhead"]
header(ws, "DAY-AHEAD PRICES", "Published 15-minute wholesale intervals. HUPX series via energy-charts (CC BY 4.0); ENTSO-E A44 latest day as a cross-check.", 2)
chart(ws, "B5", "HUPX day-ahead, all saved days (EUR/MWh)", series=[("HUPX", f"={PS}!$B$2:$B${PN+1}", f"={PS}!$C$2:$C${PN+1}", TEAL)], ytitle="EUR/MWh", size=(900, 300))
ES, EN = sheet_of("eprice"), D["eprice"][1]
chart(ws, "B21", "ENTSO-E A44 day-ahead, latest delivery day (EUR/MWh)", series=[("ENTSO-E A44", f"={ES}!$B$2:$B${EN+1}", f"={ES}!$C$2:$C${EN+1}", NAVY)], ytitle="EUR/MWh", size=(900, 260), xint=8, xfmt="hh:mm")
days = sorted({r["Delivery_Date"] for r in price})
ws.write(5, 12, "Delivery day", f_h); ws.write(5, 13, "Avg", f_h); ws.write(5, 14, "Min", f_h); ws.write(5, 15, "Max", f_h)
ws.set_column(12, 15, 12)
for i, d in enumerate(days[-31:]):
    r = 6 + i
    ws.write(r, 12, d, f_cell)
    for c, fn in ((13, "AVERAGEIFS"), (14, "MINIFS"), (15, "MAXIFS")):
        vals = [num(x["Price_EUR_MWh"]) for x in price if x["Delivery_Date"] == d and num(x["Price_EUR_MWh"]) is not None]
        cv = (sum(vals) / len(vals) if fn == "AVERAGEIFS" else min(vals) if fn == "MINIFS" else max(vals)) if vals else ""
        ws.write_formula(r, c, f'={fn}({PS}!$C$2:$C${PN+1},{PS}!$A$2:$A${PN+1},M{r+1})', f_num, cv)
ws.write(4, 12, "Daily summary (wholesale, EUR/MWh)", f_sec)

# ---------- 3 System ----------
ws = pages["System"]
header(ws, "SYSTEM: LOAD & GENERATION", "MAVIR actual load, plan, generation mix by fuel, and ENTSO-E load forecast.", 3)
LN, GN = D["load"][1], D["gen"][1]
chart(ws, "B5", "Gross system load: actual, plan, settlement (MW)", series=[(h, f"=D_Load!$A$2:$A${LN+1}", f"=D_Load!${chr(66+i)}$2:${chr(66+i)}${LN+1}", pal[i]) for i, h in enumerate(["Actual", "Plan", "Net settlement"])], ytitle="MW", size=(900, 290))
chart(ws, "B20", "Generation by fuel (MW, stacked)", kind="area", stacked=True, series=[(h.replace("_MW", ""), f"=D_Gen!$A$2:$A${GN+1}", f"=D_Gen!${chr(67+i)}$2:${chr(67+i)}${GN+1}", pal[i]) for i, h in enumerate(GC)], ytitle="MW", size=(900, 320))
chart(ws, "B38", "ENTSO-E day-ahead load forecast, latest day (MW)", series=[("Forecast", f"=D_LoadForecast!$A$2:$A${D['elf'][1]+1}", f"=D_LoadForecast!$B$2:$B${D['elf'][1]+1}", TEAL)], ytitle="MW", size=(900, 250), xint=8, xfmt="hh:mm")

# ---------- 4 Flows & balancing ----------
ws = pages["Flows_Balancing"]
header(ws, "CROSS-BORDER FLOWS & BALANCING", "MAVIR actual interconnector flows (sign as published by MAVIR) and balancing-capacity offers.", 4)
FN, BN = D["flow"][1], D["bal"][1]
chart(ws, "B5", "Actual interconnector flows by border (MW, MAVIR sign convention)", series=[(h.replace("_Actual_MW", ""), f"=D_Flows!$A$2:$A${FN+1}", f"=D_Flows!${chr(66+i)}$2:${chr(66+i)}${FN+1}", pal[i]) for i, h in enumerate(FC)], ytitle="MW", size=(900, 320))
chart(ws, "B22", "Balancing offers: aFRR / mFRR up and down (MW)", series=[(h.replace("_Offers_MW", ""), f"=D_Balancing!$A$2:$A${BN+1}", f"=D_Balancing!${chr(66+i)}$2:${chr(66+i)}${BN+1}", pal[i]) for i, h in enumerate(BC)], ytitle="MW", size=(900, 280))
ws.merge_range("B37:O38", "ENTSO-E physical flows (14 directions, 31 days) are in entsoe_physical_flows.csv with coverage in physical_flows_status.json. Directions the source does not publish are reported as SOURCE_NO_DATA gaps, never zero.", f_note)

# ---------- 5 Tariffs & macro ----------
ws = pages["Tariffs_Macro"]
header(ws, "CUSTOMER TARIFFS & MACRO", "Eurostat retail prices (all taxes included) kept separate from wholesale. HICP, FX and master-sheet macro rows.", 5)
TN = D["tar"][1]
hh = [(i, r) for i, r in enumerate(tar) if r["Consumer_Group"] == "Household"]
nh = [(i, r) for i, r in enumerate(tar) if r["Consumer_Group"] != "Household"]
for ch_i, (grp, rows_, col) in enumerate((("Household (2500-4999 kWh)", hh, TEAL), ("Non-household (500-1999 MWh)", nh, NAVY))):
    a, b = rows_[0][0] + 2, rows_[-1][0] + 2
    chart(ws, "B5" if ch_i == 0 else "I5", grp + " price, EUR/kWh all taxes", kind="column", series=[(grp, f"=D_Tariffs!$A${a}:$A${b}", f"=D_Tariffs!$D${a}:$D${b}", col)], ytitle="EUR/kWh", size=(430, 280))
chart(ws, "B21", "HICP annual inflation (%)", kind="column", series=[("HICP", f"=D_HICP!$A$2:$A${D['hicp'][1]+1}", f"=D_HICP!$B$2:$B${D['hicp'][1]+1}", GOLD)], ytitle="%", size=(430, 260))
ws.write(20, 8, "Master-sheet macro rows (hun_data.csv)", f_sec)
macro = [r for r in master if r["Domain_Category"].startswith("8.")]
for i, r in enumerate(macro):
    ws.merge_range(21 + i, 8, 21 + i, 10, r["Entity_Project_Indicator_Name"][:48], f_cell); ws.merge_range(21 + i, 11, 21 + i, 14, (r["Primary_Value"] or "")[:70] + " " + (r["Primary_Unit"] or ""), f_cell)

# ---------- 6 Registers & health ----------
ws = pages["Registers_Health"]
header(ws, "REGISTERS & DATA HEALTH", "MEKH aggregator register, MEKH storage licences, MAVIR balance responsible parties (a different thing), and every dataset state.", 6)
kpi(ws, 4, 1, "AGGREGATORS REGISTERED (MEKH)", f"=IF(COUNTA(D_Aggregators!$A$2:$A$500)=0,\"no data\",COUNTA(D_Aggregators!$A$2:$A$500))", f_kv0, "independent aggregator register", len(agg) or "no data")
kpi(ws, 4, 3, "STORAGE LICENCE POWER (MW)", f"=IF(COUNT(D_Storage!$C$2:$C$500)=0,\"no data\",SUMIFS(D_Storage!$C$2:$C$500,D_Storage!$E$2:$E$500,\"\"))", f_kv0, "licensed, not necessarily operating", round(sum(num(r["Nominal_Power_MW"]) or 0 for r in sto if not r["Revoked_Date"]), 1) if sto else "no data")
kpi(ws, 4, 5, "BALANCE RESPONSIBLE PARTIES", "=COUNTA(D_BRP!$A$2:$A$500)", f_kv0, "MAVIR imbalance settlement, not aggregators", len(brp))
kpi(ws, 4, 7, "DATASETS OK", f"=COUNTIF(D_Status!$B$2:$B$60,\"OK\")", f_kv0, "of all datasets in status.json", sum(1 for v in det.values() if v.get("state") == "OK"))
kpi(ws, 4, 9, "SOURCE_PARTIAL", f"=COUNTIF(D_Status!$B$2:$B$60,\"SOURCE_PARTIAL\")", f_kv0, "upstream published nothing for a day/direction", sum(1 for v in det.values() if v.get("state") == "SOURCE_PARTIAL"))
kpi(ws, 4, 11, "PARTIAL / FAILED / SKIPPED", f"=COUNTIF(D_Status!$B$2:$B$60,\"PARTIAL\")+COUNTIF(D_Status!$B$2:$B$60,\"FAILED\")+COUNTIF(D_Status!$B$2:$B$60,\"SKIPPED\")", f_kv0, "transport, validation, auth or skipped", sum(1 for v in det.values() if v.get("state") in ("PARTIAL", "FAILED", "SKIPPED")))
ws.write(9, 1, "Dataset states (status.json)", f_sec)
ws.set_column(1, 1, 26)
ws.write(10, 1, "Dataset", f_h); ws.write(10, 2, "State", f_h); ws.merge_range(10, 3, 10, 14, "Detail", f_h)
for i, (k, v) in enumerate(det.items()):
    r = 11 + i
    ws.write(r, 1, k, f_cell)
    ws.write(r, 2, v.get("state"), f_cell)
    ws.merge_range(r, 3, r, 14, (v.get("detail") or "")[:200], f_cell)
ws.write(11 + len(det) + 1, 1, "Largest unrevoked storage licences (MW)", f_sec)
top = sorted([r for r in sto if not r["Revoked_Date"] and num(r["Nominal_Power_MW"])], key=lambda r: -num(r["Nominal_Power_MW"]))[:10]
for i, r in enumerate(top):
    rr = 11 + len(det) + 2 + i
    ws.merge_range(rr, 1, rr, 4, r["Licensee"][:60], f_cell); ws.write(rr, 5, num(r["Nominal_Power_MW"]), f_num); ws.write(rr, 6, num(r["Storage_Capacity_MWh"]), f_num)
    ws.merge_range(rr, 7, rr, 10, r["Site"][:50], f_cell)
wb.close()
print("wrote", OUT)
