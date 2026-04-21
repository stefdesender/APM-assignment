"""
Assignment 3a
"""

import gurobipy as gp
from gurobipy import GRB
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_FORECAST
)

# Helper functions
def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

def clean_num(val, tol=1e-6, digits=6):
    """
    Cleans tiny numerical noise from solver output.
    Example: -0.0, 1e-9, -2e-8 -> 0.0
    """
    v = float(val)
    if abs(v) < tol:
        return 0.0
    return round(v, digits)

# Capacity data
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

# Build model 
model = gp.Model("APM_Assignment3a")
periods = range(1, T + 1)
parts   = PARTS

#Decision variables 
x    = model.addVars(parts, periods, name="x", vtype=GRB.INTEGER, lb=0)
y    = model.addVars(parts, periods, name="y", vtype=GRB.BINARY)
I    = model.addVars(parts, periods, name="I", vtype=GRB.INTEGER, lb=0)

# Overtime on X in units
ot_x = model.addVars(periods, name="ot_x", vtype=GRB.INTEGER, lb=0, ub=CAPACITY_X_OT_MAX)

# Overtime on Y in hours
ot_y = model.addVars(periods, name="ot_y", lb=0, ub=CAPACITY_Y_OT_MAX)

#Big-M values 
BIG_M = {i: sum(DEMAND_FORECAST) * 25 for i in parts}

# Objective 
model.setObjective(
    gp.quicksum(
        SETUP_COST[i] * y[i, t] + HOLDING_COST[i] * I[i, t]
        for i in parts for t in periods
    )
    + gp.quicksum(
        COST_OT_X * ot_x[t] +
        COST_OT_Y * ot_y[t]
        for t in periods
    ),
    GRB.MINIMIZE
)

#  Constraints 
for i in parts:
    for t in periods:
        inv_prev = INIT_INV[i] if t == 1 else I[i, t - 1]

        order_period = t - LEAD_TIME[i]
        receipts = x[i, order_period] if order_period >= 1 else 0

        internal_demand = gp.quicksum(
            BOM[p][i] * x[p, t]
            for p in get_parents(i)
        )

        external_demand = DEMAND_FORECAST[t - 1] if i == END_PRODUCT else 0

        model.addConstr(
            I[i, t] == inv_prev + receipts - internal_demand - external_demand,
            name=f"inv_balance_{i}_{t}"
        )

        model.addConstr(
            x[i, t] >= MIN_LOT[i] * y[i, t],
            name=f"min_lot_{i}_{t}"
        )

        model.addConstr(
            x[i, t] <= BIG_M[i] * y[i, t],
            name=f"bigM_{i}_{t}"
        )

# Workstation X capacity
for t in periods:
    model.addConstr(
        x['E2801', t] <= CAPACITY_X + ot_x[t],
        name=f"capacity_X_{t}"
    )

# Workstation Y capacity 
for t in periods:
    model.addConstr(
        PROC_TIME_Y['B1401'] * x['B1401', t] +
        PROC_TIME_Y['B2302'] * x['B2302', t]
        <= CAPACITY_Y + 60 * ot_y[t],
        name=f"capacity_Y_{t}"
    )

# Solve 
model.optimize()

#  Output 
if model.status == GRB.OPTIMAL:
    total_setup   = sum(SETUP_COST[i]   * y[i, t].X for i in parts for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i, t].X for i in parts for t in periods)
    total_ot_x    = sum(COST_OT_X * ot_x[t].X for t in periods)
    total_ot_y    = sum(COST_OT_Y * ot_y[t].X for t in periods)

    print(f"\n{'='*60}")
    print("ASSIGNMENT 3a - OPTIMAL SOLUTION (with overtime)")
    print(f"{'='*60}")
    print(f"Total cost:          €{clean_num(model.ObjVal):,.2f}")
    print(f"Setup cost:          €{clean_num(total_setup):,.2f}")
    print(f"Holding cost:        €{clean_num(total_holding):,.2f}")
    print(f"Overtime cost X:     €{clean_num(total_ot_x):,.2f}")
    print(f"Overtime cost Y:     €{clean_num(total_ot_y):,.2f}")
    print(f"Total overtime cost: €{clean_num(total_ot_x + total_ot_y):,.2f}")

    print(f"\n{'─'*60}")
    print("PRODUCTION / ORDER SCHEDULE (non-zero quantities)")
    print(f"{'─'*60}")
    for i in parts:
        orders = [(t, clean_num(x[i, t].X)) for t in periods if x[i, t].X > 0.5]
        if orders:
            print(f"\n{i}:")
            for t, qty in orders:
                print(f"  Week {t:2d}: {qty:,.0f} units")

    print(f"\n{'─'*60}")
    print("END INVENTORY (week 30)")
    print(f"{'─'*60}")
    for i in parts:
        print(f"  {i}: {clean_num(I[i, 30].X):,.0f} units")

    print(f"\n{'─'*60}")
    print("OVERTIME USAGE (only periods with overtime)")
    print(f"{'─'*60}")
    ot_used = False
    for t in periods:
        ox = clean_num(ot_x[t].X)
        oy = clean_num(ot_y[t].X)
        if ox > 1e-6 or oy > 1e-6:
            ot_used = True
            print(
                f"  Week {t:2d}: "
                f"X = {ox:6.0f} extra units | "
                f"Y = {oy:6.2f} extra hours ({clean_num(60 * oy):7.1f} min)"
            )
    if not ot_used:
        print("  No overtime used.")

    print(f"\n{'─'*60}")
    print("CAPACITY USAGE PER WEEK")
    print(f"{'─'*60}")
    print(f"  {'Wk':>2}  {'X used':>8} / {'X cap':>8}  |  {'Y used':>8} / {'Y cap':>8}")
    for t in periods:
        usage_x   = clean_num(x['E2801', t].X)
        usage_y   = clean_num(
            PROC_TIME_Y['B1401'] * x['B1401', t].X +
            PROC_TIME_Y['B2302'] * x['B2302', t].X
        )
        cap_x_tot = clean_num(CAPACITY_X + ot_x[t].X)
        cap_y_tot = clean_num(CAPACITY_Y + 60 * ot_y[t].X)
        print(f"  {t:2d}  {usage_x:8.0f} / {cap_x_tot:8.1f}  |  {usage_y:8.0f} / {cap_y_tot:8.1f}")

    output = {
        "status": "OPTIMAL",
        "total_cost": clean_num(model.ObjVal),
        "cost_breakdown": {
            "setup_cost": clean_num(total_setup),
            "holding_cost": clean_num(total_holding),
            "overtime_cost_X": clean_num(total_ot_x),
            "overtime_cost_Y": clean_num(total_ot_y),
            "total_overtime_cost": clean_num(total_ot_x + total_ot_y)
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
            i: {
                str(t): clean_num(x[i, t].X)
                for t in periods if x[i, t].X > 0.5
            }
            for i in parts
        },
        "inventory": {
            i: {
                str(t): clean_num(I[i, t].X)
                for t in periods
            }
            for i in parts
        },
        "overtime_usage": {
            str(t): {
                "X_overtime_units": clean_num(ot_x[t].X),
                "Y_overtime_hours": clean_num(ot_y[t].X),
                "Y_overtime_minutes": clean_num(60 * ot_y[t].X),
                "overtime_cost_X": clean_num(COST_OT_X * ot_x[t].X),
                "overtime_cost_Y": clean_num(COST_OT_Y * ot_y[t].X),
                "total_overtime_cost": clean_num(COST_OT_X * ot_x[t].X + COST_OT_Y * ot_y[t].X)
            }
            for t in periods
        }
    }

    with open("output_3a.json", "w") as f:
        json.dump(output, f, indent=2)

    print("\nResults written to output_3a.json")