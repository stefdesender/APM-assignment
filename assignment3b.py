"""
Assignment 3b

"""

import gurobipy as gp
from gurobipy import GRB
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_REALIZED, BACKORDER_COST
)

def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

def clean_num(val, tol=1e-6):
    v = float(val)
    return 0.0 if abs(v) < tol else round(v, 6)

CAPACITY_X        = 800
CAPACITY_X_OT_MAX = 300
COST_OT_X         = 2

CAPACITY_Y        = 7 * 24 * 60 - 80
CAPACITY_Y_OT_MAX = 38
COST_OT_Y         = 120

PROC_TIME_Y = {
    'B1401': 3,
    'B2302': 2
}

model = gp.Model("APM_Assignment3b")
periods = range(1, T + 1)
parts   = PARTS

x = model.addVars(parts, periods, name="x", vtype=GRB.INTEGER, lb=0)
y = model.addVars(parts, periods, name="y", vtype=GRB.BINARY)
I = model.addVars(parts, periods, name="I", vtype=GRB.INTEGER, lb=0)

B = model.addVars(periods, name="B", vtype=GRB.INTEGER, lb=0)

ot_x = model.addVars(periods, name="ot_x", vtype=GRB.INTEGER, lb=0, ub=CAPACITY_X_OT_MAX)
ot_y = model.addVars(periods, name="ot_y", lb=0, ub=CAPACITY_Y_OT_MAX)

BIG_M = {i: sum(DEMAND_REALIZED) * 25 for i in parts}

model.setObjective(
    gp.quicksum(
        SETUP_COST[i] * y[i, t] + HOLDING_COST[i] * I[i, t]
        for i in parts for t in periods
    )
    + gp.quicksum(
        COST_OT_X * ot_x[t] + COST_OT_Y * ot_y[t]
        for t in periods
    )
    + gp.quicksum(
        BACKORDER_COST * B[t]
        for t in periods
    ),
    GRB.MINIMIZE
)

for i in parts:
    for t in periods:
        inv_prev = INIT_INV[i] if t == 1 else I[i, t - 1]

        order_period = t - LEAD_TIME[i]
        receipts = x[i, order_period] if order_period >= 1 else 0

        internal_demand = gp.quicksum(
            BOM[p][i] * x[p, t]
            for p in get_parents(i)
        )

        if i != END_PRODUCT:
            model.addConstr(
                I[i, t] == inv_prev + receipts - internal_demand,
                name=f"inv_balance_{i}_{t}"
            )
        else:
            backlog_prev = 0 if t == 1 else B[t - 1]
            demand_t = DEMAND_REALIZED[t - 1]

            model.addConstr(
                I[i, t] - B[t] == inv_prev - backlog_prev + receipts - demand_t,
                name=f"inv_backlog_balance_{i}_{t}"
            )

        model.addConstr(
            x[i, t] >= MIN_LOT[i] * y[i, t],
            name=f"min_lot_{i}_{t}"
        )

        model.addConstr(
            x[i, t] <= BIG_M[i] * y[i, t],
            name=f"bigM_{i}_{t}"
        )

for t in periods:
    model.addConstr(
        x['E2801', t] <= CAPACITY_X + ot_x[t],
        name=f"capacity_X_{t}"
    )

for t in periods:
    model.addConstr(
        PROC_TIME_Y['B1401'] * x['B1401', t] +
        PROC_TIME_Y['B2302'] * x['B2302', t]
        <= CAPACITY_Y + 60 * ot_y[t],
        name=f"capacity_Y_{t}"
    )

model.optimize()

if model.status == GRB.OPTIMAL:
    total_setup     = sum(SETUP_COST[i] * y[i, t].X for i in parts for t in periods)
    total_holding   = sum(HOLDING_COST[i] * I[i, t].X for i in parts for t in periods)
    total_ot_x      = sum(COST_OT_X * ot_x[t].X for t in periods)
    total_ot_y      = sum(COST_OT_Y * ot_y[t].X for t in periods)
    total_backorder = sum(BACKORDER_COST * B[t].X for t in periods)

    total_demand = sum(DEMAND_REALIZED)

    served = {}
    periods_without_backorder = 0

    for t in periods:
        backlog_prev = 0 if t == 1 else B[t - 1].X
        backlog_now = B[t].X
        demand_t = DEMAND_REALIZED[t - 1]

        served[t] = demand_t + backlog_prev - backlog_now

        if backlog_now <= 1e-6:
            periods_without_backorder += 1

    total_served_immediately = sum(served[t] for t in periods)

    service_level = periods_without_backorder / T
    fill_rate = total_served_immediately / total_demand if total_demand > 0 else 1.0

    print(f"\n{'='*70}")
    print("ASSIGNMENT 3b - OPTIMAL SOLUTION (realized demand + backorders)")
    print(f"{'='*70}")
    print(f"Total cost:           €{model.ObjVal:,.2f}")
    print(f"Setup cost:           €{total_setup:,.2f}")
    print(f"Holding cost:         €{total_holding:,.2f}")
    print(f"Overtime cost X:      €{total_ot_x:,.2f}")
    print(f"Overtime cost Y:      €{total_ot_y:,.2f}")
    print(f"Backorder cost:       €{total_backorder:,.2f}")
    print(f"Total overtime cost:  €{total_ot_x + total_ot_y:,.2f}")

    print(f"\nService level:        {service_level:.4f} ({100*service_level:.2f}%)")
    print(f"Fill rate:            {fill_rate:.4f} ({100*fill_rate:.2f}%)")

    output = {
        "status": "OPTIMAL",
        "total_cost": clean_num(model.ObjVal),
        "cost_breakdown": {
            "setup_cost": clean_num(total_setup),
            "holding_cost": clean_num(total_holding),
            "overtime_cost_X": clean_num(total_ot_x),
            "overtime_cost_Y": clean_num(total_ot_y),
            "total_overtime_cost": clean_num(total_ot_x + total_ot_y),
            "backorder_cost": clean_num(total_backorder)
        },
        "service_metrics": {
            "service_level": clean_num(service_level),
            "fill_rate": clean_num(fill_rate)
        },
        "capacity_parameters": {
            "X_regular_units": CAPACITY_X,
            "X_overtime_max_units": CAPACITY_X_OT_MAX,
            "X_overtime_cost_per_unit": COST_OT_X,
            "Y_regular_minutes": CAPACITY_Y,
            "Y_overtime_max_hours": CAPACITY_Y_OT_MAX,
            "Y_overtime_cost_per_hour": COST_OT_Y
        },
        "production_schedule": {
            i: {str(t): clean_num(x[i, t].X) for t in periods if x[i, t].X > 0.5}
            for i in parts
        },
        "inventory": {
            i: {str(t): clean_num(I[i, t].X) for t in periods}
            for i in parts
        },
        "backorders": {
            str(t): clean_num(B[t].X) for t in periods
        },
        "overtime_usage": {
            str(t): {
                "X_overtime_units": clean_num(ot_x[t].X),
                "Y_overtime_hours": clean_num(ot_y[t].X),
                "Y_overtime_minutes": clean_num(60 * ot_y[t].X)
            }
            for t in periods
        }
    }


    OUTPUT_FILE = "OUTPUT.xlsx"
    SHEET_NAME = "Output_3b"

    if os.path.exists(OUTPUT_FILE):
        wb = load_workbook(OUTPUT_FILE)
        if SHEET_NAME in wb.sheetnames:
            del wb[SHEET_NAME]
        ws = wb.create_sheet(SHEET_NAME)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME

    NO_FILL    = PatternFill(fill_type=None)
    BLACK_FILL = PatternFill("solid", start_color="000000", end_color="000000")
    none_border = Border()
    bot_medium  = Border(bottom=Side(style="medium", color="000000"))
    bot_thin    = Border(bottom=Side(style="thin", color="CCCCCC"))

    def plain(cell, val, bold=False, align="left", size=9, color="000000", fmt=None, border=None):
        cell.value = val
        cell.font = Font(name="Calibri", bold=bold, size=size, color=color)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.fill = NO_FILL
        if border:
            cell.border = border
        if fmt:
            cell.number_format = fmt

    def hdr(cell, val):
        cell.value = val
        cell.font = Font(name="Calibri", bold=False, size=9, color="FFFFFF")
        cell.fill = BLACK_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = none_border

    def section_title(ws, r, last_col, text):
        ws.row_dimensions[r].height = 14
        plain(ws.cell(r, 2), text, size=8, color="888888", border=bot_medium)
        for col in range(3, last_col + 1):
            ws.cell(r, col).border = bot_medium

    ws.column_dimensions["A"].width = 1
    ws.column_dimensions["B"].width = 18
    for col in range(3, T + 4):
        ws.column_dimensions[get_column_letter(col)].width = 5.5

    last_col = T + 2

    ws.row_dimensions[1].height = 26
    ws.merge_cells(f"B1:{get_column_letter(last_col)}1")
    plain(ws.cell(1, 2), "Assignment 3b - Optimal solution with realized demand (overtime)", bold=True, size=13)

    ws.row_dimensions[2].height = 4

    cost_rows = [
        ("Setup cost", total_setup, False),
        ("Holding cost", total_holding, False),
        ("Overtime cost X", total_ot_x, False),
        ("Overtime cost Y", total_ot_y, False),
        ("Total overtime cost", total_ot_x + total_ot_y, False),
        ("Backorder cost", total_backorder, False),
        ("Total cost", model.ObjVal, True),
    ]

    for r, (label, val, bold) in enumerate(cost_rows, start=3):
        ws.row_dimensions[r].height = 16
        plain(ws.cell(r, 2), label, size=9, color="555555")
        plain(ws.cell(r, 3), round(clean_num(val), 2), bold=bold, size=9, fmt='€ #,##0.00')
        ws.merge_cells(f"C{r}:{get_column_letter(last_col)}{r}")

    ws.row_dimensions[11].height = 6
    section_title(ws, 12, last_col, "Service metrics")
    service_rows = [
        ("Service level", f"{service_level*100:.2f}%  ({periods_without_backorder}/{T} periods without backorder)"),
        ("Fill rate", f"{fill_rate*100:.2f}%  ({total_served_immediately:,.0f} / {total_demand:,.0f} units served)"),
        ("Total demand", f"{total_demand:,.0f} units"),
        ("Total backorder", f"{sum(B[t].X for t in periods):,.0f} units"),
    ]

    for r, (label, text) in enumerate(service_rows, start=13):
        ws.row_dimensions[r].height = 16
        plain(ws.cell(r, 2), label, size=9, color="555555")
        plain(ws.cell(r, 3), text, size=9)
        ws.merge_cells(f"C{r}:{get_column_letter(last_col)}{r}")

    r = 19
    section_title(ws, r, last_col, "Production / order schedule")
    r += 1
    hdr(ws.cell(r, 2), "Part")
    for t in periods:
        hdr(ws.cell(r, t + 2), str(t))

    for i in parts:
        r += 1
        plain(ws.cell(r, 2), i, size=9, border=bot_thin)
        for t in periods:
            val = round(clean_num(x[i, t].X))
            plain(
                ws.cell(r, t + 2),
                val if val > 0 else "",
                bold=val > 0,
                size=9,
                align="center",
                fmt="#,##0" if val > 0 else None,
                border=bot_thin
            )

    r += 2
    section_title(ws, r, last_col, "Inventory levels")
    r += 1
    hdr(ws.cell(r, 2), "Part")
    for t in periods:
        hdr(ws.cell(r, t + 2), str(t))

    for i in parts:
        r += 1
        plain(ws.cell(r, 2), i, size=9, border=bot_thin)
        for t in periods:
            val = round(clean_num(I[i, t].X))
            color = "BBBBBB" if val == 0 else "000000"
            plain(
                ws.cell(r, t + 2),
                val,
                size=9,
                align="center",
                color=color,
                fmt="#,##0",
                border=bot_thin
            )

    r += 2
    section_title(ws, r, last_col, f"Backorders - {END_PRODUCT}")
    r += 1
    hdr(ws.cell(r, 2), "Part")
    for t in periods:
        hdr(ws.cell(r, t + 2), str(t))

    r += 1
    plain(ws.cell(r, 2), END_PRODUCT, size=9, border=bot_thin)
    for t in periods:
        val = round(clean_num(B[t].X))
        plain(
            ws.cell(r, t + 2),
            val if val > 0 else "",
            bold=val > 0,
            size=9,
            align="center",
            color="CC0000" if val > 0 else "BBBBBB",
            fmt="#,##0" if val > 0 else None,
            border=bot_thin
        )

    r += 2
    section_title(ws, r, last_col, "Overtime usage")
    r += 1
    hdr(ws.cell(r, 2), "")
    for t in periods:
        hdr(ws.cell(r, t + 2), str(t))

    r += 1
    plain(ws.cell(r, 2), "X overtime units", size=9, border=bot_thin)
    for t in periods:
        val = round(clean_num(ot_x[t].X))
        plain(
            ws.cell(r, t + 2),
            val if val > 0 else "",
            bold=val > 0,
            size=9,
            align="center",
            color="CC0000" if val > 0 else "BBBBBB",
            fmt="#,##0" if val > 0 else None,
            border=bot_thin
        )

    r += 1
    plain(ws.cell(r, 2), "Y overtime hours", size=9, border=bot_thin)
    for t in periods:
        val = clean_num(ot_y[t].X)
        plain(
            ws.cell(r, t + 2),
            val if val > 0 else "",
            bold=val > 0,
            size=9,
            align="center",
            color="CC0000" if val > 0 else "BBBBBB",
            fmt="0.00" if val > 0 else None,
            border=bot_thin
        )

    r += 2
    section_title(ws, r, last_col, "Setup decisions")
    r += 1
    hdr(ws.cell(r, 2), "Part")
    for t in periods:
        hdr(ws.cell(r, t + 2), str(t))

    for i in parts:
        r += 1
        plain(ws.cell(r, 2), i, size=9, border=bot_thin)
        for t in periods:
            setup = 1 if x[i, t].X > 0.5 else 0
            plain(
                ws.cell(r, t + 2),
                "x" if setup else "",
                bold=True,
                size=9,
                align="center",
                border=bot_thin
            )

    wb.save(OUTPUT_FILE)

    print(f"Results written to {OUTPUT_FILE} -> sheet '{SHEET_NAME}'")
