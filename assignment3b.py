"""
APM Project 2026 - Assignment 3b
MIP model for production planning with finite capacity + overtime option
using REALIZED demand and allowing BACKORDERS

- Workstation X: capacity 800 units/week, overtime up to 300 extra units at €2/unit
- Workstation Y: capacity 7*24*60 - 80 = 10000 min/week,
                 overtime up to 38 hours at €120/hour
- Backorder cost: €250 per unit per period
- Demand = realized demand
"""

import gurobipy as gp
from gurobipy import GRB
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_REALIZED, BACKORDER_COST
)

# ── Helper functions ────────────────────────────────────────────────────────
def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

# ── Capacity data ───────────────────────────────────────────────────────────
CAPACITY_X        = 800
CAPACITY_X_OT_MAX = 300
COST_OT_X         = 2

CAPACITY_Y        = 7 * 24 * 60 - 80   # = 10000 min/week
CAPACITY_Y_OT_MAX = 38                 # hours/week
COST_OT_Y         = 120                # €/hour

PROC_TIME_Y = {
    'B1401': 3,   # min/unit
    'B2302': 2    # min/unit
}

# ── Build model ────────────────────────────────────────────────────────────
model = gp.Model("APM_Assignment3b")
periods = range(1, T + 1)
parts   = PARTS

# ── Decision variables ─────────────────────────────────────────────────────
x = model.addVars(parts, periods, name="x", vtype=GRB.INTEGER, lb=0)
y = model.addVars(parts, periods, name="y", vtype=GRB.BINARY)
I = model.addVars(parts, periods, name="I", vtype=GRB.INTEGER, lb=0)

# Backorders only for end product
B = model.addVars(periods, name="B", vtype=GRB.INTEGER, lb=0)

# Overtime
ot_x = model.addVars(periods, name="ot_x", vtype=GRB.INTEGER, lb=0, ub=CAPACITY_X_OT_MAX)
ot_y = model.addVars(periods, name="ot_y", lb=0, ub=CAPACITY_Y_OT_MAX)

# ── Big-M values ───────────────────────────────────────────────────────────
BIG_M = {i: sum(DEMAND_REALIZED) * 25 for i in parts}

# ── Objective ──────────────────────────────────────────────────────────────
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

# ── Constraints ────────────────────────────────────────────────────────────
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
            # Standard inventory balance for components/subassemblies
            model.addConstr(
                I[i, t] == inv_prev + receipts - internal_demand,
                name=f"inv_balance_{i}_{t}"
            )
        else:
            # End product with backorders
            # Net inventory = inventory - backlog
            # I_t - B_t = I_{t-1} - B_{t-1} + receipts - demand
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

# ── Workstation X capacity ─────────────────────────────────────────────────
for t in periods:
    model.addConstr(
        x['E2801', t] <= CAPACITY_X + ot_x[t],
        name=f"capacity_X_{t}"
    )

# ── Workstation Y capacity ─────────────────────────────────────────────────
for t in periods:
    model.addConstr(
        PROC_TIME_Y['B1401'] * x['B1401', t] +
        PROC_TIME_Y['B2302'] * x['B2302', t]
        <= CAPACITY_Y + 60 * ot_y[t],
        name=f"capacity_Y_{t}"
    )

# ── Solve ──────────────────────────────────────────────────────────────────
model.optimize()

# ── Output ─────────────────────────────────────────────────────────────────
if model.status == GRB.OPTIMAL:
    print(f"\n{'='*70}")
    print("ASSIGNMENT 3b - OPTIMAL SOLUTION (realized demand + backorders)")
    print(f"{'='*70}")
    print(f"Total cost:           €{model.ObjVal:,.2f}")

    total_setup     = sum(SETUP_COST[i] * y[i, t].X for i in parts for t in periods)
    total_holding   = sum(HOLDING_COST[i] * I[i, t].X for i in parts for t in periods)
    total_ot_x      = sum(COST_OT_X * ot_x[t].X for t in periods)
    total_ot_y      = sum(COST_OT_Y * ot_y[t].X for t in periods)
    total_backorder = sum(BACKORDER_COST * B[t].X for t in periods)

    print(f"Setup cost:           €{total_setup:,.2f}")
    print(f"Holding cost:         €{total_holding:,.2f}")
    print(f"Overtime cost X:      €{total_ot_x:,.2f}")
    print(f"Overtime cost Y:      €{total_ot_y:,.2f}")
    print(f"Backorder cost:       €{total_backorder:,.2f}")
    print(f"Total overtime cost:  €{total_ot_x + total_ot_y:,.2f}")

    # ── Service level and fill rate ────────────────────────────────────────
    total_demand = sum(DEMAND_REALIZED)

    # demand served in period t = demand_t + backlog_prev - backlog_t
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

    print(f"\n{'─'*70}")
    print("SERVICE PERFORMANCE")
    print(f"{'─'*70}")
    print(f"Service level:        {service_level:.4f} ({100*service_level:.2f}%)")
    print(f"Fill rate:            {fill_rate:.4f} ({100*fill_rate:.2f}%)")

    print(f"\n{'─'*70}")
    print("PRODUCTION / ORDER SCHEDULE (non-zero quantities)")
    print(f"{'─'*70}")
    for i in parts:
        orders = [(t, x[i, t].X) for t in periods if x[i, t].X > 0.5]
        if orders:
            print(f"\n{i}:")
            for t, qty in orders:
                print(f"  Week {t:2d}: {qty:,.0f} units")

    print(f"\n{'─'*70}")
    print("END INVENTORY / BACKORDERS (week 30)")
    print(f"{'─'*70}")
    for i in parts:
        print(f"  {i}: {I[i, 30].X:,.0f} units")
    print(f"  Backorders E2801: {B[30].X:,.0f} units")

    print(f"\n{'─'*70}")
    print("OVERTIME USAGE (only periods with overtime)")
    print(f"{'─'*70}")
    ot_used = False
    for t in periods:
        ox = ot_x[t].X
        oy = ot_y[t].X
        if ox > 1e-6 or oy > 1e-6:
            ot_used = True
            print(f"  Week {t:2d}: X = {ox:6.0f} extra units | Y = {oy:6.2f} extra hours ({60*oy:7.1f} min)")
    if not ot_used:
        print("  No overtime used.")

    print(f"\n{'─'*70}")
    print("BACKORDERS PER WEEK")
    print(f"{'─'*70}")
    bo_used = False
    for t in periods:
        if B[t].X > 1e-6:
            bo_used = True
            print(f"  Week {t:2d}: {B[t].X:,.0f} units backordered")
    if not bo_used:
        print("  No backorders.")

    output = {
        "status": "OPTIMAL",
        "total_cost": float(model.ObjVal),
        "cost_breakdown": {
            "setup_cost": float(total_setup),
            "holding_cost": float(total_holding),
            "overtime_cost_X": float(total_ot_x),
            "overtime_cost_Y": float(total_ot_y),
            "total_overtime_cost": float(total_ot_x + total_ot_y),
            "backorder_cost": float(total_backorder)
        },
        "service_metrics": {
            "service_level": float(service_level),
            "fill_rate": float(fill_rate)
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
            i: {str(t): float(x[i, t].X) for t in periods if x[i, t].X > 0.5}
            for i in parts
        },
        "inventory": {
            i: {str(t): float(I[i, t].X) for t in periods}
            for i in parts
        },
        "backorders": {
            str(t): float(B[t].X) for t in periods
        },
        "overtime_usage": {
            str(t): {
                "X_overtime_units": float(ot_x[t].X),
                "Y_overtime_hours": float(ot_y[t].X),
                "Y_overtime_minutes": float(60 * ot_y[t].X)
            }
            for t in periods
        }
    }

    with open("output_3b.json", "w") as f:
        json.dump(output, f, indent=2)

    print("\nResults written to output_3b.json")
