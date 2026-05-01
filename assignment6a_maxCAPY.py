"""
APM Project 2026 - Assignment 6a
Idea: Sensitivity analysis on the maximum expansion cap of Workstation Y.
"""

import gurobipy as gp
from gurobipy import GRB
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST,
    DEMAND_FORECAST, DEMAND_REALIZED, BACKORDER_COST
)

# ── helpers ───────────────────────────────────────────────────────────────────
def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

def clean_num(val, tol=1e-6):
    v = float(val)
    return 0.0 if abs(v) < tol else round(v, 6)

# ── constants (identical to 5a/5b) ───────────────────────────────────────────
CAP_X_BASE     = 800
CAP_X_MAX_EXP  = 200
COST_EXP_X     = 10
CAP_X_OT_MAX   = 300
COST_OT_X      = 2

CAP_Y_BASE     = 60 * 24 * 7 - 80          # 10 000 min/week
COST_EXP_Y_PCT = 1_500
CAP_Y_OT_MAX   = 38
COST_OT_Y      = 120

PROC_X = {END_PRODUCT: 1}
PROC_Y = {"B1401": 3, "B2302": 2}

periods = range(1, T + 1)
BIG_M   = {i: sum(DEMAND_FORECAST) * 25 for i in PARTS}

# ── sensitivity range ─────────────────────────────────────────────────────────
REF_DY_PCT = 14.1
CAP_VALUES = [round(REF_DY_PCT - 2 * 5, 1), round(REF_DY_PCT - 1 * 5, 1)] + [round(REF_DY_PCT + i * 5, 1) for i in range(0, 14)]
# → [14.1, 19.1, 24.1, 29.1, 34.1, 39.1, 44.1, 49.1, 54.1, 59.1, 64.1, 69.1, 74.1, 79.1]

# ── solve one scenario ────────────────────────────────────────────────────────
def solve_scenario(cap_y_max_pct: float) -> dict | None:
    """
    Build and solve the 5a model with a given max-expansion cap for Y.
    Returns a dict with forecast costs, the schedule, and dy_pct chosen.
    Returns None if no optimal solution found.
    """
    m = gp.Model(f"APM_6a_cap{int(cap_y_max_pct)}")
    m.setParam("OutputFlag", 0)

    x      = m.addVars(PARTS, periods, name="x",  vtype=GRB.INTEGER, lb=0)
    y      = m.addVars(PARTS, periods, name="y",  vtype=GRB.BINARY)
    I      = m.addVars(PARTS, periods, name="I",  vtype=GRB.INTEGER, lb=0)
    dx     = m.addVar(name="dx",     vtype=GRB.INTEGER, lb=0, ub=CAP_X_MAX_EXP)
    dy_pct = m.addVar(name="dy_pct", lb=cap_y_max_pct, ub=cap_y_max_pct)
    ot_x   = m.addVars(periods, name="ot_x", vtype=GRB.INTEGER, lb=0, ub=CAP_X_OT_MAX)
    ot_y   = m.addVars(periods, name="ot_y", lb=0, ub=CAP_Y_OT_MAX)

    m.setObjective(
        gp.quicksum(SETUP_COST[i]*y[i,t] + HOLDING_COST[i]*I[i,t]
                    for i in PARTS for t in periods)
        + COST_EXP_X * dx + COST_EXP_Y_PCT * dy_pct
        + gp.quicksum(COST_OT_X*ot_x[t] + COST_OT_Y*ot_y[t] for t in periods),
        GRB.MINIMIZE
    )

    for i in PARTS:
        for t in periods:
            inv_prev = INIT_INV[i] if t == 1 else I[i, t-1]
            op  = t - LEAD_TIME[i]
            rec = x[i, op] if op >= 1 else 0
            int_d = gp.quicksum(BOM[p][i]*x[p,t] for p in get_parents(i))
            ext_d = DEMAND_FORECAST[t-1] if i == END_PRODUCT else 0
            m.addConstr(I[i,t] == inv_prev + rec - int_d - ext_d)
            m.addConstr(x[i,t] >= MIN_LOT[i] * y[i,t])
            m.addConstr(x[i,t] <= BIG_M[i]   * y[i,t])

    for t in periods:
        m.addConstr(
            gp.quicksum(PROC_X.get(i,0)*x[i,t] for i in PARTS)
            <= CAP_X_BASE + dx + ot_x[t]
        )
        m.addConstr(
            gp.quicksum(PROC_Y.get(i,0)*x[i,t] for i in PARTS)
            <= CAP_Y_BASE + (CAP_Y_BASE / 100.0) * dy_pct + 60.0 * ot_y[t]
        )

    m.optimize()

    if m.status != GRB.OPTIMAL:
        return None

    dx_val     = clean_num(dx.X)
    dy_val     = clean_num(dy_pct.X)
    invest_X   = COST_EXP_X     * dx_val
    invest_Y   = COST_EXP_Y_PCT * dy_val
    total_setup   = sum(SETUP_COST[i]   * y[i,t].X for i in PARTS for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i,t].X for i in PARTS for t in periods)
    total_ot_x    = sum(COST_OT_X * ot_x[t].X for t in periods)
    total_ot_y    = sum(COST_OT_Y * ot_y[t].X for t in periods)

    schedule = {i: {t: x[i,t].X for t in periods} for i in PARTS}
    ot_x_vals = {t: clean_num(ot_x[t].X) for t in periods}
    ot_y_vals = {t: clean_num(ot_y[t].X) for t in periods}

    return dict(
        cap_y_max_pct = cap_y_max_pct,
        dx_val        = dx_val,
        dy_val        = dy_val,
        invest_X      = invest_X,
        invest_Y      = invest_Y,
        forecast_setup   = total_setup,
        forecast_holding = total_holding,
        forecast_ot_x    = total_ot_x,
        forecast_ot_y    = total_ot_y,
        forecast_total   = m.ObjVal,
        schedule      = schedule,
        ot_x_vals     = ot_x_vals,
        ot_y_vals     = ot_y_vals,
    )


# ── simulate realized demand (identical logic to 5b) ─────────────────────────
def simulate_realized(res: dict) -> dict:
    schedule  = res["schedule"]
    ot_x_vals = res["ot_x_vals"]
    ot_y_vals = res["ot_y_vals"]

    # fixed costs from the plan
    setup_cost_fixed = sum(
        SETUP_COST[i]
        for i in PARTS for t in periods
        if schedule[i].get(t, 0.0) > 0.5
    )

    inventory  = {i: {} for i in PARTS}
    held       = {i: {} for i in PARTS}
    backorders = {t: 0.0 for t in periods}

    # components
    for i in PARTS:
        if i == END_PRODUCT:
            continue
        inv_prev = float(INIT_INV[i])
        for t in periods:
            op  = t - LEAD_TIME[i]
            rec = schedule[i].get(op, 0.0) if op >= 1 else 0.0
            int_d = sum(BOM[p][i] * schedule[p].get(t, 0.0) for p in get_parents(i))
            net = inv_prev + rec - int_d
            inventory[i][t] = clean_num(net)
            held[i][t]      = clean_num(max(0.0, net))
            inv_prev = net

    # end product with backorders
    inv_prev = float(INIT_INV[END_PRODUCT])
    bo_prev  = 0.0
    for t in periods:
        op  = t - LEAD_TIME[END_PRODUCT]
        rec = schedule[END_PRODUCT].get(op, 0.0) if op >= 1 else 0.0
        net = inv_prev - bo_prev + rec - DEMAND_REALIZED[t-1]
        if net >= 0:
            inventory[END_PRODUCT][t] = clean_num(net)
            held[END_PRODUCT][t]      = clean_num(net)
            backorders[t] = 0.0
        else:
            inventory[END_PRODUCT][t] = 0.0
            held[END_PRODUCT][t]      = 0.0
            backorders[t] = clean_num(-net)
        inv_prev = inventory[END_PRODUCT][t]
        bo_prev  = backorders[t]

    total_holding   = sum(HOLDING_COST[i] * held[i][t] for i in PARTS for t in periods)
    total_backorder = sum(BACKORDER_COST  * backorders[t] for t in periods)
    total_cost      = (setup_cost_fixed + total_holding + total_backorder
                       + res["invest_X"] + res["invest_Y"]
                       + sum(COST_OT_X * ot_x_vals[t] for t in periods)
                       + sum(COST_OT_Y * ot_y_vals[t] for t in periods))

    periods_no_bo = sum(1 for t in periods if backorders[t] == 0)
    service_level = periods_no_bo / T
    total_demand  = sum(DEMAND_REALIZED)
    total_new_bo  = 0.0
    prev_bo = 0.0
    for t in periods:
        new_bo = max(0.0, backorders[t] - prev_bo)
        total_new_bo += new_bo
        prev_bo = backorders[t]
    fill_rate = 1.0 - (total_new_bo / total_demand) if total_demand > 0 else 1.0

    return dict(
        setup_cost     = setup_cost_fixed,
        holding_cost   = total_holding,
        backorder_cost = total_backorder,
        total_cost     = total_cost,
        service_level  = service_level,
        fill_rate      = fill_rate,
        total_new_bo   = total_new_bo,
        backorders     = backorders,
        inventory      = inventory,
    )


# ── run all scenarios ─────────────────────────────────────────────────────────
print("Running sensitivity analysis on max expansion cap of Workstation Y …")
print(f"{'Cap (%)':>8}  {'dy chosen (%)':>14}  {'Cost forecast':>14}  "
      f"{'Cost realized':>14}  {'SL':>6}  {'FR':>7}")
print("-" * 72)

results = []
for cap in CAP_VALUES:
    res = solve_scenario(cap)
    if res is None:
        print(f"{cap:>8}%  {'INFEASIBLE':>14}")
        continue
    sim = simulate_realized(res)
    res["sim"] = sim
    results.append(res)
    print(f"{cap:>7}%  {res['dy_val']:>13.1f}%  "
          f"€{res['forecast_total']:>13,.2f}  "
          f"€{sim['total_cost']:>13,.2f}  "
          f"{sim['service_level']*100:>5.1f}%  "
          f"{sim['fill_rate']*100:>6.2f}%")

# ── write to OUTPUT.xlsx ──────────────────────────────────────────────────────
OUTPUT_FILE = "OUTPUT.xlsx"
SHEET_NAME  = "Output_6a_maxY"

if os.path.exists(OUTPUT_FILE):
    wb = load_workbook(OUTPUT_FILE)
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = wb.create_sheet(SHEET_NAME)
else:
    wb = Workbook()
    ws  = wb.active
    ws.title = SHEET_NAME

# ── style helpers ─────────────────────────────────────────────────────────────
NO_FILL    = PatternFill(fill_type=None)
BLACK_FILL = PatternFill("solid", start_color="000000", end_color="000000")
GREY_FILL  = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
BLUE_FILL  = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
LGREY_FILL = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")
GREEN_FILL = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")

none_border = Border()
bot_medium  = Border(bottom=Side(style="medium", color="000000"))
bot_thin    = Border(bottom=Side(style="thin",   color="CCCCCC"))
full_thin   = Border(
    left  =Side(style="thin", color="CCCCCC"),
    right =Side(style="thin", color="CCCCCC"),
    top   =Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

def plain(cell, val, bold=False, align="left", size=9,
          color="000000", fmt=None, border=None, fill=None):
    cell.value = val
    cell.font  = Font(name="Calibri", bold=bold, size=size, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center",
                                wrap_text=False)
    cell.fill  = fill if fill else NO_FILL
    if border: cell.border = border
    if fmt:    cell.number_format = fmt

def hdr(cell, val, fill=None):
    cell.value = val
    cell.font  = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    cell.fill  = fill if fill else BLACK_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = none_border

def section_title(ws, r, last_col, text):
    ws.row_dimensions[r].height = 14
    plain(ws.cell(r, 2), text, size=8, color="888888", border=bot_medium)
    for col in range(3, last_col + 1):
        ws.cell(r, col).border = bot_medium

# column widths
ws.column_dimensions["A"].width = 1
col_labels = [
    ("B", 28),
    ("C", 10), ("D", 10), ("E", 14), ("F", 14),
    ("G", 14), ("H", 14), ("I", 10), ("J", 10),
]
for col_letter, w in col_labels:
    ws.column_dimensions[col_letter].width = w

LAST_COL = 10   # column J

# ── Title ─────────────────────────────────────────────────────────────────────
ws.row_dimensions[1].height = 28
ws.merge_cells("B1:J1")
plain(ws.cell(1,2),
      "Assignment 6a – Sensitivity analysis: max expansion cap Workstation Y",
      bold=True, size=13)

ws.row_dimensions[2].height = 14
ws.merge_cells("B2:J2")
plain(ws.cell(2,2),
      "Idea: Force Workstation Y to produce B1401/B2302 at higher capacity → "
      "build component buffer → reduce backorders under realized demand.",
      size=8, color="555555")

# ── Summary table (forecast + realized) ──────────────────────────────────────
r = 4
section_title(ws, r, LAST_COL,
              "Sensitivity table: forecast cost vs realized cost per max-cap scenario")
r += 1
ws.row_dimensions[r].height = 18

headers = [
    "Max cap Y (%)", "dy chosen (%)",
    "Forecast total (€)", "Realized total (€)",
    "Holding (real.) (€)", "Backorder (real.) (€)",
    "Service level", "Fill rate", "Backorders (units)"
]
for c, h in enumerate(headers, start=2):
    hdr(ws.cell(r, c), h)

for res in results:
    r += 1
    ws.row_dimensions[r].height = 16
    sim = res["sim"]

    # highlight the reference row (cap = 40, identical to 5a/5b)
    is_ref = (res["cap_y_max_pct"] == 40)
    row_fill = LGREY_FILL if is_ref else NO_FILL

    # find best realized cost to highlight
    best_realized = min(r2["sim"]["total_cost"] for r2 in results)
    is_best = abs(sim["total_cost"] - best_realized) < 0.01

    row_fill = LGREY_FILL if is_ref else (GREEN_FILL if is_best else NO_FILL)
    ref_label = " ← ref (5a/5b)" if is_ref else (" ← best" if is_best else "")

    data = [
        (f"{int(res['cap_y_max_pct'])}%{ref_label}", False, "left",  None,          "000000"),
        (f"{res['dy_val']:.1f}%",                     False, "center",None,          "000000"),
        (res["forecast_total"],                        False, "right", '"€"#,##0.00',"000000"),
        (sim["total_cost"],                            is_best,"right",'"€"#,##0.00',"000000"),
        (sim["holding_cost"],                          False, "right", '"€"#,##0.00',"000000"),
        (sim["backorder_cost"],                        False, "right", '"€"#,##0.00',
         "CC0000" if sim["backorder_cost"] > 0 else "000000"),
        (f"{sim['service_level']*100:.1f}%",           False, "center",None,          "000000"),
        (f"{sim['fill_rate']*100:.2f}%",               False, "center",None,          "000000"),
        (sim["total_new_bo"],                           False, "right", "#,##0.00",
         "CC0000" if sim["total_new_bo"] > 0 else "000000"),
    ]
    for c, (val, bold, align, fmt, color) in enumerate(data, start=2):
        plain(ws.cell(r, c), val, bold=bold, align=align,
              fmt=fmt, color=color, border=full_thin, fill=row_fill)

# ── Cost breakdown per scenario ───────────────────────────────────────────────
r += 2
section_title(ws, r, LAST_COL, "Full cost breakdown (realized demand)")
r += 1
ws.row_dimensions[r].height = 18

hdr2_labels = [
    "Max cap Y (%)", "Setup (€)", "Holding (€)", "Backorder (€)",
    "OT cost X (€)", "OT cost Y (€)", "Invest X (€)", "Invest Y (€)", "Total (€)"
]
for c, h in enumerate(hdr2_labels, start=2):
    hdr(ws.cell(r, c), h)

for res in results:
    r += 1
    ws.row_dimensions[r].height = 16
    sim   = res["sim"]
    ot_x_total = sum(COST_OT_X * res["ot_x_vals"][t] for t in periods)
    ot_y_total = sum(COST_OT_Y * res["ot_y_vals"][t] for t in periods)
    is_ref = (res["cap_y_max_pct"] == 40)
    best_realized = min(r2["sim"]["total_cost"] for r2 in results)
    is_best = abs(sim["total_cost"] - best_realized) < 0.01
    row_fill = LGREY_FILL if is_ref else (GREEN_FILL if is_best else NO_FILL)

    row_data = [
        f"{int(res['cap_y_max_pct'])}%",
        sim["setup_cost"],
        sim["holding_cost"],
        sim["backorder_cost"],
        ot_x_total,
        ot_y_total,
        res["invest_X"],
        res["invest_Y"],
        sim["total_cost"],
    ]
    for c, val in enumerate(row_data, start=2):
        is_money = isinstance(val, float)
        plain(ws.cell(r, c), val,
              bold=(c == 10 and is_best),
              align="right" if is_money else "center",
              fmt='"€"#,##0.00' if is_money else None,
              border=full_thin, fill=row_fill,
              color="CC0000" if (c == 5 and val > 0) else "000000")


# ── Backorders per period ─────────────────────────────────────────────────────
r += 2
BO_LAST_COL = 2 + T + 1
section_title(ws, r, BO_LAST_COL, "Backorders per period (units, realized demand)")
r += 1
ws.row_dimensions[r].height = 18
hdr(ws.cell(r, 2), "Max cap Y (%)")
for t in periods:
    col = 2 + t
    ws.column_dimensions[get_column_letter(col)].width = 8
    hdr(ws.cell(r, col), f"P{t}")
hdr(ws.cell(r, BO_LAST_COL), "Total units")
ws.column_dimensions[get_column_letter(BO_LAST_COL)].width = 12

for res in results:
    r += 1
    ws.row_dimensions[r].height = 16
    sim = res["sim"]
    is_ref  = abs(res["cap_y_max_pct"] - REF_DY_PCT) < 0.05
    best_realized = min(r2["sim"]["total_cost"] for r2 in results)
    is_best = abs(sim["total_cost"] - best_realized) < 0.01
    row_fill = LGREY_FILL if is_ref else (GREEN_FILL if is_best else NO_FILL)

    plain(ws.cell(r, 2), f"{res['cap_y_max_pct']:.1f}%",
          align="left", border=full_thin, fill=row_fill)
    total_bo = 0.0
    for t in periods:
        col = 2 + t
        bo  = sim["backorders"][t]
        total_bo += bo
        plain(ws.cell(r, col),
              round(bo, 2) if bo > 0 else "",
              align="right",
              fmt="#,##0.00" if bo > 0 else None,
              border=full_thin, fill=row_fill,
              color="CC0000" if bo > 0 else "000000")
    plain(ws.cell(r, BO_LAST_COL),
          round(total_bo, 2) if total_bo > 0 else "",
          bold=True, align="right",
          fmt="#,##0.00" if total_bo > 0 else None,
          border=full_thin, fill=row_fill,
          color="CC0000" if total_bo > 0 else "000000")

# ══════════════════════════════════════════════════════════════════════════
# Chart: Forecast cost vs Realized cost per cap scenario
# ══════════════════════════════════════════════════════════════════════════
from openpyxl.chart import BarChart, Reference, Series

r += 2
section_title(ws, r, LAST_COL, "Cost comparison chart (forecast vs realized total cost)")
chart_data_start_row = r + 1

# Write a small data table for the chart
ws.cell(chart_data_start_row, 2).value = "Cap Y (%)"
ws.cell(chart_data_start_row, 3).value = "Forecast total"
ws.cell(chart_data_start_row, 4).value = "Realized total"
ws.cell(chart_data_start_row, 5).value = "Backorder cost"
ws.cell(chart_data_start_row, 6).value = "Holding cost"
for col in [2, 3, 4, 5, 6]:
    c = ws.cell(chart_data_start_row, col)
    c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c.fill = BLACK_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")

for idx, res in enumerate(results):
    row = chart_data_start_row + 1 + idx
    sim = res["sim"]
    is_ref = abs(res["cap_y_max_pct"] - REF_DY_PCT) < 0.05
    rfill = LGREY_FILL if is_ref else NO_FILL
    ws.cell(row, 2).value = f"{res['cap_y_max_pct']:.1f}%"
    ws.cell(row, 3).value = round(res["forecast_total"], 2)
    ws.cell(row, 4).value = round(sim["total_cost"], 2)
    ws.cell(row, 5).value = round(sim["backorder_cost"], 2)
    ws.cell(row, 6).value = round(sim["holding_cost"], 2)
    for col in [2, 3, 4, 5, 6]:
        cell = ws.cell(row, col)
        cell.font      = Font(name="Calibri", size=9)
        cell.alignment = Alignment(horizontal="center")
        cell.border    = full_thin
        cell.fill      = rfill

n_scenarios = len(results)
chart_data_end_row = chart_data_start_row + n_scenarios

# ── Bar chart: forecast vs realized total ────────────────────────────────
chart = BarChart()
chart.type      = "col"
chart.grouping  = "clustered"
chart.title     = "Total Cost per Y-Expansion Scenario"
chart.y_axis.title = "Total Cost (\u20ac)"
chart.x_axis.title = "Max expansion cap Workstation Y"
chart.style     = 10
chart.width     = 24
chart.height    = 14

forecast_vals = Reference(ws, min_col=3, max_col=3,
    min_row=chart_data_start_row, max_row=chart_data_end_row)
realized_vals = Reference(ws, min_col=4, max_col=4,
    min_row=chart_data_start_row, max_row=chart_data_end_row)

s1 = Series(forecast_vals, title="Forecast total (\u20ac)")
s1.graphicalProperties.solidFill = "4472C4"
s2 = Series(realized_vals, title="Realized total (\u20ac)")
s2.graphicalProperties.solidFill = "ED7D31"

chart.append(s1)
chart.append(s2)

labels = Reference(ws, min_col=2, max_col=2,
    min_row=chart_data_start_row + 1, max_row=chart_data_end_row)
chart.set_categories(labels)

chart_anchor_row = chart_data_end_row + 2
ws.add_chart(chart, f"B{chart_anchor_row}")

# ── Stacked bar chart: cost breakdown under realized demand ───────────────
chart2 = BarChart()
chart2.type      = "col"
chart2.grouping  = "stacked"
chart2.title     = "Cost Breakdown under Realized Demand"
chart2.y_axis.title = "Cost (\u20ac)"
chart2.x_axis.title = "Max expansion cap Workstation Y"
chart2.style     = 10
chart2.width     = 24
chart2.height    = 14

bo_vals   = Reference(ws, min_col=5, max_col=5,
    min_row=chart_data_start_row, max_row=chart_data_end_row)
hold_vals = Reference(ws, min_col=6, max_col=6,
    min_row=chart_data_start_row, max_row=chart_data_end_row)

s3 = Series(hold_vals, title="Holding cost (\u20ac)")
s3.graphicalProperties.solidFill = "4472C4"
s4 = Series(bo_vals, title="Backorder cost (\u20ac)")
s4.graphicalProperties.solidFill = "FF0000"

chart2.append(s3)
chart2.append(s4)
chart2.set_categories(labels)

chart2_anchor_row = chart_anchor_row + 22
ws.add_chart(chart2, f"B{chart2_anchor_row}")

wb.save(OUTPUT_FILE)
print(f"\nResults written to {OUTPUT_FILE} → sheet '{SHEET_NAME}'")