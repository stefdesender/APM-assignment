"""
APM Project 2026 - Assignment 2a
MIP model for production planning with finite capacity
- Demand = forecasted demand
- Workstation X capacity for E2801
- Workstation Y capacity for B1401 and B2302
"""

import gurobipy as gp
from gurobipy import GRB
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_FORECAST
)

# ── Helper functions ────────────────────────────────────────────────────────
def get_children(part):
    """Return direct children of a part in the BOM."""
    return BOM.get(part, {})

def get_parents(part):
    """Return all parents of a part and the qty needed per parent unit."""
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

# ── Capacity data for assignment 2a ────────────────────────────────────────
CAPACITY_X = 800  # units/week for E2801
CAPACITY_Y = 7 * 24 * 60 - 80  # minutes/week = 10000

PROC_TIME_Y = {
    'B1401': 3,  # min/unit
    'B2302': 2   # min/unit
}

# ── Build model ────────────────────────────────────────────────────────────
model = gp.Model("APM_Assignment2a")

periods = range(1, T + 1)
parts = PARTS

# ── Decision variables ─────────────────────────────────────────────────────
# x[i,t] : order/production quantity of part i in period t
x = model.addVars(parts, periods, name="x", vtype=GRB.INTEGER, lb=0)

# y[i,t] : 1 if part i is ordered/produced in period t
y = model.addVars(parts, periods, name="y", vtype=GRB.BINARY)

# I[i,t] : inventory of part i at end of period t
I = model.addVars(parts, periods, name="I", lb=0)

# ── Big-M values ───────────────────────────────────────────────────────────
# Safe upper bound, similar to your 1a approach
BIG_M = {i: sum(DEMAND_FORECAST) * 25 for i in parts}

# ── Objective: minimize setup + holding costs ──────────────────────────────
model.setObjective(
    gp.quicksum(
        SETUP_COST[i] * y[i, t] + HOLDING_COST[i] * I[i, t]
        for i in parts for t in periods
    ),
    GRB.MINIMIZE
)

# ── Constraints ────────────────────────────────────────────────────────────
for i in parts:
    for t in periods:

        # 1. Inventory balance
        if t == 1:
            inv_prev = INIT_INV[i]
        else:
            inv_prev = I[i, t - 1]

        order_placement_period = t - LEAD_TIME[i]
        if order_placement_period >= 1:
            receipts = x[i, order_placement_period]
        else:
            receipts = 0

        # Internal demand from parent items
        internal_demand = gp.quicksum(
            BOM[p][i] * x[p, t]
            for p in get_parents(i)
        )

        # External demand only for end product
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

# 4. Capacity constraint on workstation X
# Only E2801 uses workstation X, max 800 units/week
for t in periods:
    model.addConstr(
        x['E2801', t] <= CAPACITY_X,
        name=f"capacity_X_{t}"
    )

# 5. Capacity constraint on workstation Y
# B1401 needs 3 min/unit, B2302 needs 2 min/unit
for t in periods:
    model.addConstr(
        3 * x['B1401', t] + 2 * x['B2302', t] <= CAPACITY_Y,
        name=f"capacity_Y_{t}"
    )

# ── Solve ──────────────────────────────────────────────────────────────────
model.optimize()

# ── Output ─────────────────────────────────────────────────────────────────
if model.status == GRB.OPTIMAL:
    print(f"\n{'='*60}")
    print("ASSIGNMENT 2a - OPTIMAL SOLUTION")
    print(f"{'='*60}")
    print(f"Total cost:     €{model.ObjVal:,.2f}")

    total_setup = sum(SETUP_COST[i] * y[i, t].X for i in parts for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i, t].X for i in parts for t in periods)

    print(f"Setup cost:     €{total_setup:,.2f}")
    print(f"Holding cost:   €{total_holding:,.2f}")

    print(f"\n{'─'*60}")
    print("PRODUCTION / ORDER SCHEDULE (non-zero quantities)")
    print(f"{'─'*60}")
    for i in parts:
        orders = [(t, x[i, t].X) for t in periods if x[i, t].X > 0.5]
        if orders:
            print(f"\n{i}:")
            for t, qty in orders:
                print(f"  Week {t:2d}: {qty:,.0f} units")

    print(f"\n{'─'*60}")
    print("END INVENTORY (week 30)")
    print(f"{'─'*60}")
    for i in parts:
        print(f"  {i}: {I[i, 30].X:,.0f} units")

    # Optional: show capacity usage
    print(f"\n{'─'*60}")
    print("CAPACITY USAGE")
    print(f"{'─'*60}")
    for t in periods:
        usage_x = x['E2801', t].X
        usage_y = 3 * x['B1401', t].X + 2 * x['B2302', t].X
        print(
            f"Week {t:2d} | X: {usage_x:7.1f} / {CAPACITY_X:4d} units"
            f" | Y: {usage_y:7.1f} / {CAPACITY_Y:5d} min"
        )

    output = {
        "status": "OPTIMAL",
        "total_cost": model.ObjVal,
        "setup_cost": total_setup,
        "holding_cost": total_holding,
        "capacity": {
            "workstation_X_units_per_week": CAPACITY_X,
            "workstation_Y_minutes_per_week": CAPACITY_Y
        },
        "production_schedule": {
            i: {str(t): x[i, t].X for t in periods if x[i, t].X > 0.5}
            for i in parts
        },
        "inventory": {
            i: {str(t): I[i, t].X for t in periods}
            for i in parts
        },
        "capacity_usage": {
            str(t): {
                "X_used_units": x['E2801', t].X,
                "Y_used_minutes": 3 * x['B1401', t].X + 2 * x['B2302', t].X
            }
            for t in periods
        }
    }

    with open("output_2aleo.json", "w") as f:
        json.dump(output, f, indent=2)

    print("\nResults written to output_2aleo.json")

else:
    print(f"Model status: {model.status} — no optimal solution found.")