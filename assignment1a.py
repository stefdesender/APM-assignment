"""
APM Project 2026 - Assignment 1a
MIP model for production planning (infinite capacity, forecasted demand)
"""

import gurobipy as gp
from gurobipy import GRB
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_FORECAST
)

def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

model = gp.Model("APM_Assignment1a")
periods = range(1, T + 1)
parts   = PARTS

x = model.addVars(parts, periods, name="x", lb=0)
y = model.addVars(parts, periods, name="y", vtype=GRB.BINARY)
I = model.addVars(parts, periods, name="I", lb=0)

BIG_M = {i: sum(DEMAND_FORECAST) * 25 for i in parts}

model.setObjective(
    gp.quicksum(
        SETUP_COST[i] * y[i, t] + HOLDING_COST[i] * I[i, t]
        for i in parts for t in periods
    ),
    GRB.MINIMIZE
)

for i in parts:
    for t in periods:
        inv_prev = INIT_INV[i] if t == 1 else I[i, t - 1]
        order_placement_period = t - LEAD_TIME[i]
        receipts = x[i, order_placement_period] if order_placement_period >= 1 else 0
        internal_demand = gp.quicksum(BOM[p][i] * x[p, t] for p in get_parents(i))
        external_demand = DEMAND_FORECAST[t - 1] if i == END_PRODUCT else 0
        model.addConstr(I[i, t] == inv_prev + receipts - internal_demand - external_demand,
                        name=f"inv_balance_{i}_{t}")
        model.addConstr(x[i, t] >= MIN_LOT[i] * y[i, t], name=f"min_lot_{i}_{t}")
        model.addConstr(x[i, t] <= BIG_M[i] * y[i, t],   name=f"bigM_{i}_{t}")

model.optimize()

if model.status == GRB.OPTIMAL:
    total_setup   = sum(SETUP_COST[i]   * y[i, t].X for i in parts for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i, t].X for i in parts for t in periods)

    OUTPUT_FILE = "OUTPUT.xlsx"
    SHEET_NAME  = "Output_1a"

    if os.path.exists(OUTPUT_FILE):
        wb = load_workbook(OUTPUT_FILE)
        if SHEET_NAME in wb.sheetnames:
            del wb[SHEET_NAME]
        ws = wb.create_sheet(SHEET_NAME)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME

    # ── Styles ────────────────────────────────────────────────────────────
    NO_FILL   = PatternFill(fill_type=None)
    BLACK_FILL = PatternFill("solid", start_color="000000", end_color="000000")

    none_border = Border()
    blk  = Side(style="thin", color="000000")
    gray = Side(style="thin", color="AAAAAA")
    bot_black = Border(bottom=Side(style="medium", color="000000"))
    full_gray = Border(left=gray, right=gray, top=gray, bottom=gray)
    full_black = Border(left=blk, right=blk, top=blk, bottom=blk)
    bottom_only = Border(bottom=Side(style="thin", color="CCCCCC"))

    def plain(cell, val, bold=False, align="left", size=10, color="000000", fmt=None, border=None):
        cell.value = val
        cell.font  = Font(name="Calibri", bold=bold, size=size, color=color)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.fill  = NO_FILL
        if border:
            cell.border = border
        if fmt:
            cell.number_format = fmt

    def header_cell(cell, val):
        """Black background, white text, centered."""
        cell.value = val
        cell.font  = Font(name="Calibri", bold=False, size=9, color="FFFFFF")
        cell.fill  = BLACK_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = none_border

    # ── Column widths ─────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 1
    ws.column_dimensions["B"].width = 9
    for col in range(3, T + 4):
        ws.column_dimensions[get_column_letter(col)].width = 5.5

    last_col = T + 2

    # ══════════════════════════════════════════════════════════════════════
    # Row 1: Title
    # ══════════════════════════════════════════════════════════════════════
    ws.row_dimensions[1].height = 26
    ws.merge_cells(f"B1:{get_column_letter(last_col)}1")
    plain(ws.cell(1, 2), "Assignment 1a - Optimal solution", bold=True, size=13)

    # ══════════════════════════════════════════════════════════════════════
    # Rows 3-5: Cost summary
    # ══════════════════════════════════════════════════════════════════════
    ws.row_dimensions[2].height = 4
    for r, (label, val) in enumerate([
        ("Total cost",   model.ObjVal),
        ("Setup cost",   total_setup),
        ("Holding cost", total_holding),
    ], start=3):
        ws.row_dimensions[r].height = 16
        plain(ws.cell(r, 2), label, size=9, color="555555")
        vc = ws.cell(r, 3)
        plain(vc, round(val, 2), bold=(r == 3), size=9, align="left", fmt='"€"#,##0.00')
        ws.merge_cells(f"C{r}:{get_column_letter(last_col)}{r}")

    # ══════════════════════════════════════════════════════════════════════
    # Section: Production schedule
    # ══════════════════════════════════════════════════════════════════════
    r = 7
    ws.row_dimensions[r].height = 6
    r = 8
    ws.row_dimensions[r].height = 14
    plain(ws.cell(r, 2), "Production / order schedule (units)", size=8,
          color="888888", border=bot_black)
    for col in range(3, last_col + 1):
        ws.cell(r, col).border = bot_black

    r = 9
    ws.row_dimensions[r].height = 16
    header_cell(ws.cell(r, 2), "Part")
    for t in periods:
        header_cell(ws.cell(r, t + 2), str(t))

    for idx, i in enumerate(parts, start=1):
        r += 1
        ws.row_dimensions[r].height = 15
        plain(ws.cell(r, 2), i, size=9, border=bottom_only)
        for t in periods:
            val = round(x[i, t].X) if x[i, t].X > 0.5 else ""
            vc = ws.cell(r, t + 2)
            plain(vc, val, bold=bool(val), size=9, align="center",
                  fmt='#,##0' if val != "" else None, border=bottom_only)

    # ══════════════════════════════════════════════════════════════════════
    # Section: Inventory levels
    # ══════════════════════════════════════════════════════════════════════
    r += 2
    ws.row_dimensions[r].height = 14
    plain(ws.cell(r, 2), "Inventory levels (end of period)", size=8,
          color="888888", border=bot_black)
    for col in range(3, last_col + 1):
        ws.cell(r, col).border = bot_black

    r += 1
    ws.row_dimensions[r].height = 16
    header_cell(ws.cell(r, 2), "Part")
    for t in periods:
        header_cell(ws.cell(r, t + 2), str(t))

    for idx, i in enumerate(parts, start=1):
        r += 1
        ws.row_dimensions[r].height = 15
        plain(ws.cell(r, 2), i, size=9, border=bottom_only)
        for t in periods:
            v = round(I[i, t].X)
            col_val = "000000" if v > 0 else "BBBBBB"
            plain(ws.cell(r, t + 2), v, size=9, align="center",
                  color=col_val, fmt='#,##0', border=bottom_only)

    # ══════════════════════════════════════════════════════════════════════
    # Section: Setup decisions
    # ══════════════════════════════════════════════════════════════════════
    r += 2
    ws.row_dimensions[r].height = 14
    plain(ws.cell(r, 2), "Setup decisions", size=8,
          color="888888", border=bot_black)
    for col in range(3, last_col + 1):
        ws.cell(r, col).border = bot_black

    r += 1
    ws.row_dimensions[r].height = 16
    header_cell(ws.cell(r, 2), "Part")
    for t in periods:
        header_cell(ws.cell(r, t + 2), str(t))

    for idx, i in enumerate(parts, start=1):
        r += 1
        ws.row_dimensions[r].height = 15
        plain(ws.cell(r, 2), i, size=9, border=bottom_only)
        for t in periods:
            setup = int(round(y[i, t].X))
            plain(ws.cell(r, t + 2), "x" if setup else "",
                  bold=True, size=9, align="center",
                  color="000000", border=bottom_only)

    wb.save(OUTPUT_FILE)
    print(f"\nResults written to {OUTPUT_FILE} -> sheet '{SHEET_NAME}'")
    print(f"Total cost:    EUR {model.ObjVal:,.2f}")
    print(f"  Setup cost:  EUR {total_setup:,.2f}")
    print(f"  Holding cost: EUR {total_holding:,.2f}")

else:
    print(f"Model status: {model.status} -- no optimal solution found.")