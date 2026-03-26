"""
APM Project 2026 - Assignment 1a
MIP model for production planning (infinite capacity, forecasted demand)
"""

import gurobipy as gp
from gurobipy import GRB
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_FORECAST
)

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
model = gp.Model("APM_Assignment1a")

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
# Upper bound on order quantity: total forecasted demand * max BOM multiplier
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
        # Inventory(t) = Inventory(t-1) + receipts(t) - demand(t)
        # Receipts in period t come from orders placed in period t - lead_time
        if t == 1:
            inv_prev = INIT_INV[i]
        else:
            inv_prev = I[i, t - 1]
            
        # Orders placed in period (t - lead_time[i]) arrive in period t
        order_placement_period = t - LEAD_TIME[i]
        if order_placement_period >= 1:
            receipts = x[i, order_placement_period]
        elif order_placement_period == 0:
            # Order placed before horizon — assume 0 (already in init inv)
            receipts = 0
        else:
            receipts = 0

        # Internal demand: how much of part i is consumed by its parents
        # A parent p ordered/produced in period t needs qty[i] units of i
        # available at period t (i.e., the consumption happens when parent is produced)
        internal_demand = gp.quicksum(
            BOM[p][i] * x[p, t]
            for p in get_parents(i)
        )

        # External demand only applies to end product
        external_demand = DEMAND_FORECAST[t - 1] if i == END_PRODUCT else 0

        model.addConstr(
            I[i, t] == inv_prev + receipts - internal_demand - external_demand,
            name=f"inv_balance_{i}_{t}"
        )

        # ── 2. Minimum lot size (only order if y=1, and at least MIN_LOT) ──
        model.addConstr(
            x[i, t] >= MIN_LOT[i] * y[i, t],
            name=f"min_lot_{i}_{t}"
        )

        # ── 3. Big-M: can only produce if y=1 ─────────────────────────────
        model.addConstr(
            x[i, t] <= BIG_M[i] * y[i, t],
            name=f"bigM_{i}_{t}"
        )

# ── Solve ──────────────────────────────────────────────────────────────────
model.optimize()

# ── Output ─────────────────────────────────────────────────────────────────
if model.status == GRB.OPTIMAL:
    print(f"\n{'='*60}")
    print(f"ASSIGNMENT 1a - OPTIMAL SOLUTION")
    print(f"{'='*60}")
    print(f"Total cost:    €{model.ObjVal:,.2f}")

    # Cost breakdown
    total_setup   = sum(SETUP_COST[i]   * y[i,t].X for i in parts for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i,t].X for i in parts for t in periods)
    print(f"  Setup cost:  €{total_setup:,.2f}")
    print(f"  Holding cost:€{total_holding:,.2f}")

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

    # Inventory levels
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
        "production_schedule": {
            i: {str(t): x[i,t].X for t in periods if x[i,t].X > 0.5}
            for i in parts
        },
        "inventory": {
            i: {str(t): I[i,t].X for t in periods}
            for i in parts
        }
    }
    with open("output_1a.json", "w") as f:
        json.dump(output, f, indent=2)
    print("\nResults written to output_1a.json")

else:
    print(f"Model status: {model.status} — no optimal solution found.")
