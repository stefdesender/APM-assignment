"""
APM Project 2026 - Assignment 4a
MIP model with PERMANENT CAPACITY EXPANSION (modernized machines).

Builds on Assignment 2a (finite capacity) by adding:
  - dx     : integer variable – extra units permanently added to workstation X
  - dy_pct : integer variable – number of 1%-increments permanently added to workstation Y

New parameters vs. Assignment 2a:
  CAP_X_BASE      = 800 units/week
  CAP_X_MAX_EXP   = 200 units       (max permanent expansion on X)
  COST_EXP_X      = €10 per extra unit of capacity on X

  CAP_Y_BASE      = 60*24*7 - 80 minutes/week
  CAP_Y_MAX_PCT   = 40 %             (max permanent expansion on Y)
  COST_EXP_Y_PCT  = €1 500 per 1 % increment on Y

Workstation assignments (same as Assignment 2a):
  Workstation X : E2801 – 1 unit needs 1 slot  (capacity in units/week)
  Workstation Y : B1401 → 3 min/unit,  B2302 → 2 min/unit
"""

import gurobipy as gp
from gurobipy import GRB
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_FORECAST
)

# ── Workstation parameters ─────────────────────────────────────────────────
CAP_X_BASE       = 800
PROC_X           = {END_PRODUCT: 1}
CAP_Y_BASE       = 60 * 24 * 7 - 80          # 10 000 min/week
PROC_Y           = {"B1401": 3, "B2302": 2}

# Permanent expansion limits & costs
CAP_X_MAX_EXP    = 200
COST_EXP_X       = 10

CAP_Y_MAX_PCT    = 40
COST_EXP_Y_PCT   = 1_500

# ── Helper ─────────────────────────────────────────────────────────────────
def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

# ── Build model ────────────────────────────────────────────────────────────
model = gp.Model("APM_Assignment4a")

periods = range(1, T + 1)
parts   = PARTS

# ── Decision variables ─────────────────────────────────────────────────────
x = model.addVars(parts, periods, name="x", lb=0)
y = model.addVars(parts, periods, name="y", vtype=GRB.BINARY)
I = model.addVars(parts, periods, name="I", lb=0)

# Permanent expansion variables
dx     = model.addVar(name="dx",     vtype=GRB.INTEGER, lb=0, ub=CAP_X_MAX_EXP)
dy_pct = model.addVar(name="dy_pct", vtype=GRB.INTEGER, lb=0, ub=CAP_Y_MAX_PCT)

# ── Big-M ──────────────────────────────────────────────────────────────────
BIG_M = {i: sum(DEMAND_FORECAST) * 25 for i in parts}

# ── Objective ──────────────────────────────────────────────────────────────
model.setObjective(
    gp.quicksum(
        SETUP_COST[i] * y[i, t] + HOLDING_COST[i] * I[i, t]
        for i in parts for t in periods
    )
    + COST_EXP_X     * dx
    + COST_EXP_Y_PCT * dy_pct,
    GRB.MINIMIZE
)

# ── Constraints ────────────────────────────────────────────────────────────
for i in parts:
    for t in periods:

        # 1. Inventory balance
        inv_prev = INIT_INV[i] if t == 1 else I[i, t - 1]
        order_placement_period = t - LEAD_TIME[i]
        receipts = x[i, order_placement_period] if order_placement_period >= 1 else 0

        internal_demand = gp.quicksum(
            BOM[p][i] * x[p, t] for p in get_parents(i)
        )
        external_demand = DEMAND_FORECAST[t - 1] if i == END_PRODUCT else 0

        model.addConstr(
            I[i, t] == inv_prev + receipts - internal_demand - external_demand,
            name=f"inv_balance_{i}_{t}"
        )

        # 2. Minimum lot size
        model.addConstr(
            x[i, t] >= MIN_LOT[i] * y[i, t],
            name=f"min_lot_{i}_{t}"
        )

        # 3. Big-M linking
        model.addConstr(
            x[i, t] <= BIG_M[i] * y[i, t],
            name=f"bigM_{i}_{t}"
        )

# 4. Workstation X capacity (per period)
for t in periods:
    model.addConstr(
        gp.quicksum(PROC_X[i] * x[i, t] for i in PROC_X if i in parts)
        <= CAP_X_BASE + dx,
        name=f"cap_X_{t}"
    )

# 5. Workstation Y capacity (per period)
#    CAP_Y_BASE * (1 + dy_pct/100)  →  linearised as  CAP_Y_BASE + (CAP_Y_BASE/100)*dy_pct
for t in periods:
    model.addConstr(
        gp.quicksum(PROC_Y[i] * x[i, t] for i in PROC_Y if i in parts)
        <= CAP_Y_BASE + (CAP_Y_BASE / 100.0) * dy_pct,
        name=f"cap_Y_{t}"
    )

# ── Solve ──────────────────────────────────────────────────────────────────
model.optimize()

# ── Output ─────────────────────────────────────────────────────────────────
if model.status == GRB.OPTIMAL:
    dx_val     = dx.X
    dy_pct_val = dy_pct.X
    invest_X   = COST_EXP_X     * dx_val
    invest_Y   = COST_EXP_Y_PCT * dy_pct_val

    total_setup   = sum(SETUP_COST[i]   * y[i,t].X for i in parts for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i,t].X for i in parts for t in periods)

    print(f"\n{'='*60}")
    print(f"ASSIGNMENT 4a - OPTIMAL SOLUTION (Permanent Capacity Expansion)")
    print(f"{'='*60}")
    print(f"Total cost:         €{model.ObjVal:,.2f}")
    print(f"  Setup cost:       €{total_setup:,.2f}")
    print(f"  Holding cost:     €{total_holding:,.2f}")
    print(f"  Investment X:     €{invest_X:,.2f}  (+{dx_val:.0f} units → new cap {CAP_X_BASE + dx_val:.0f} u/week)")
    print(f"  Investment Y:     €{invest_Y:,.2f}  (+{dy_pct_val:.0f}% → new cap {CAP_Y_BASE*(1+dy_pct_val/100):.0f} min/week)")

    print(f"\n{'─'*60}")
    print("PRODUCTION/ORDER SCHEDULE (non-zero quantities):")
    print(f"{'─'*60}")
    for i in parts:
        orders = [(t, x[i,t].X) for t in periods if x[i,t].X > 0.5]
        if orders:
            print(f"\n{i}:")
            for t, qty in orders:
                print(f"  Week {t:2d}: {qty:,.0f} units")

    print(f"\n{'─'*60}")
    print("END INVENTORY PER PART (end of horizon):")
    print(f"{'─'*60}")
    for i in parts:
        print(f"  {i}: {I[i, T].X:,.0f} units")

    print(f"\n{'─'*60}")
    print("WORKSTATION UTILISATION:")
    print(f"{'─'*60}")
    cap_x_new = CAP_X_BASE + dx_val
    cap_y_new = CAP_Y_BASE * (1 + dy_pct_val / 100)
    for t in periods:
        load_x = sum(PROC_X.get(i, 0) * x[i,t].X for i in parts)
        load_y = sum(PROC_Y.get(i, 0) * x[i,t].X for i in parts)
        if load_x > 0 or load_y > 0:
            print(f"  Week {t:2d}: X {load_x:6.0f}/{cap_x_new:.0f} ({100*load_x/cap_x_new:5.1f}%)   "
                  f"Y {load_y:7.0f}/{cap_y_new:.0f} ({100*load_y/cap_y_new:5.1f}%)")

    # Write output
    output = {
        "status": "OPTIMAL",
        "total_cost": model.ObjVal,
        "setup_cost": total_setup,
        "holding_cost": total_holding,
        "investment_cost_X": invest_X,
        "investment_cost_Y": invest_Y,
        "expansion_X_units": dx_val,
        "expansion_Y_pct": dy_pct_val,
        "new_cap_X": CAP_X_BASE + dx_val,
        "new_cap_Y_minutes": cap_y_new,
        "production_schedule": {
            i: {str(t): x[i,t].X for t in periods if x[i,t].X > 0.5}
            for i in parts
        },
        "inventory": {
            i: {str(t): I[i,t].X for t in periods}
            for i in parts
        }
    }
    with open("output_4a.json", "w") as f:
        json.dump(output, f, indent=2)
    print("\nResults written to output_4a.json")

else:
    print(f"Model status: {model.status} — no optimal solution found.")
