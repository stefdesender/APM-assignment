"""
APM Project 2026 - Assignment 2a
MIP model for production planning (FINITE capacity, forecasted demand)

New constraints vs Assignment 1a:
  - Workstation X: max 800 units/week for E2801
  - Workstation Y: B1401 needs 3 min/unit, B2302 needs 2 min/unit
                   runs 24/7 = 7*24*60 = 10080 min/week minus 80 min maintenance
                   => 10000 min/week available
"""

import gurobipy as gp
from gurobipy import GRB
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_FORECAST
)

# ── Workstation capacities ─────────────────────────────────────────────────
# Workstation X: processes E2801, max 800 units/week
CAP_X = 800  # units/week

# Workstation Y: processes B1401 (3 min/unit) and B2302 (2 min/unit)
# Runs 24/7 = 7 days * 24 hours * 60 min = 10080 min/week
# Minus 80 min maintenance => 10000 min/week available
MINUTES_PER_WEEK = 7 * 24 * 60   # 10080
MAINTENANCE      = 80
CAP_Y            = MINUTES_PER_WEEK - MAINTENANCE  # 10000 min/week

PROC_TIME = {
    'B1401': 3,  # minutes per unit
    'B2302': 2,  # minutes per unit
}

# ── Helper: compute gross requirements via BOM ─────────────────────────────
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

# ── Build model ────────────────────────────────────────────────────────────
model = gp.Model("APM_Assignment2a")

periods = range(1, T + 1)   # weeks 1..30
parts   = PARTS

# ── Decision variables ─────────────────────────────────────────────────────
# x[i,t] : production/order quantity of part i in period t (continuous >= 0)
x = model.addVars(parts, periods, name="x", lb=0)

# y[i,t] : binary setup variable (1 if part i is produced/ordered in period t)
y = model.addVars(parts, periods, name="y", vtype=GRB.BINARY)

# I[i,t] : inventory of part i at end of period t
I = model.addVars(parts, periods, name="I", lb=0)

# ── Big-M for lot-size linking constraint ──────────────────────────────────
BIG_M = {i: sum(DEMAND_FORECAST) * 25 for i in parts}  # safe upper bound

# ── Objective: minimize total setup + holding costs ────────────────────────
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

        # ── 1. Inventory balance ───────────────────────────────────────────
        if t == 1:
            inv_prev = INIT_INV[i]
        else:
            inv_prev = I[i, t - 1]

        order_placement_period = t - LEAD_TIME[i]
        if order_placement_period >= 1:
            receipts = x[i, order_placement_period]
        else:
            receipts = 0

        internal_demand = gp.quicksum(
            BOM[p][i] * x[p, t]
            for p in get_parents(i)
        )

        external_demand = DEMAND_FORECAST[t - 1] if i == END_PRODUCT else 0

        model.addConstr(
            I[i, t] == inv_prev + receipts - internal_demand - external_demand,
            name=f"inv_balance_{i}_{t}"
        )

        # ── 2. Minimum lot size ────────────────────────────────────────────
        model.addConstr(
            x[i, t] >= MIN_LOT[i] * y[i, t],
            name=f"min_lot_{i}_{t}"
        )

        # ── 3. Big-M: can only produce if y=1 ─────────────────────────────
        model.addConstr(
            x[i, t] <= BIG_M[i] * y[i, t],
            name=f"bigM_{i}_{t}"
        )

# ── 4. NEW: Workstation X capacity constraint ──────────────────────────────
# E2801 can be processed at most 800 units per week on workstation X
for t in periods:
    model.addConstr(
        x[END_PRODUCT, t] <= CAP_X,
        name=f"cap_X_{t}"
    )

# ── 5. NEW: Workstation Y capacity constraint ──────────────────────────────
# B1401 (3 min/unit) + B2302 (2 min/unit) <= 10000 min/week
for t in periods:
    model.addConstr(
        gp.quicksum(PROC_TIME[i] * x[i, t] for i in PROC_TIME) <= CAP_Y,
        name=f"cap_Y_{t}"
    )

# ── Solve ──────────────────────────────────────────────────────────────────
model.optimize()

# ── Output ─────────────────────────────────────────────────────────────────
if model.status == GRB.OPTIMAL:
    print(f"\n{'='*60}")
    print(f"ASSIGNMENT 2a - OPTIMAL SOLUTION (Finite Capacity)")
    print(f"{'='*60}")
    print(f"Total cost:    €{model.ObjVal:,.2f}")

    # Cost breakdown
    total_setup   = sum(SETUP_COST[i]   * y[i,t].X for i in parts for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i,t].X for i in parts for t in periods)
    print(f"  Setup cost:  €{total_setup:,.2f}")
    print(f"  Holding cost:€{total_holding:,.2f}")

    # Workstation utilization summary
    print(f"\n{'─'*60}")
    print("WORKSTATION UTILIZATION:")
    print(f"{'─'*60}")
    for t in periods:
        usage_x = x[END_PRODUCT, t].X
        usage_y = sum(PROC_TIME[i] * x[i, t].X for i in PROC_TIME)
        if usage_x > 0 or usage_y > 0:
            print(f"  Week {t:2d}: WS-X = {usage_x:6.0f}/{CAP_X} units "
                  f"| WS-Y = {usage_y:7.1f}/{CAP_Y} min")

    # Production schedule per part
    print(f"\n{'─'*60}")
    print("PRODUCTION/ORDER SCHEDULE (non-zero quantities):")
    print(f"{'─'*60}")
    for i in parts:
        orders = [(t, x[i,t].X) for t in periods if x[i,t].X > 0.5]
        if orders:
            print(f"\n{i}:")
            for t, qty in orders:
                print(f"  Week {t:2d}: order {qty:,.0f} units")

    # Inventory levels at end of horizon
    print(f"\n{'─'*60}")
    print("END INVENTORY PER PART (end of horizon):")
    print(f"{'─'*60}")
    for i in parts:
        print(f"  {i}: {I[i,30].X:,.0f} units")

    # Write to output file
    output = {
        "status": "OPTIMAL",
        "total_cost": model.ObjVal,
        "setup_cost": total_setup,
        "holding_cost": total_holding,
        "workstation_capacities": {
            "X": {"capacity_units_per_week": CAP_X, "part": END_PRODUCT},
            "Y": {"capacity_minutes_per_week": CAP_Y,
                  "processing_times": PROC_TIME}
        },
        "production_schedule": {
            i: {str(t): x[i,t].X for t in periods if x[i,t].X > 0.5}
            for i in parts
        },
        "inventory": {
            i: {str(t): I[i,t].X for t in periods}
            for i in parts
        },
        "workstation_usage": {
            "X": {str(t): x[END_PRODUCT, t].X for t in periods},
            "Y": {str(t): sum(PROC_TIME[i] * x[i, t].X for i in PROC_TIME)
                  for t in periods}
        }
    }
    with open("output_2a.json", "w") as f:
        json.dump(output, f, indent=2)
    print("\nResults written to output_2a.json")

else:
    print(f"Model status: {model.status} — no optimal solution found.")