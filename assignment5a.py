"""
APM Project 2026 - Assignment 5a
MIP model with OVERTIME + PERMANENT CAPACITY EXPANSION (combined).
 
Combines Assignment 3a (overtime) and Assignment 4a (permanent expansion):
 
Workstation X:
  - Base capacity:        800 units/week
  - Permanent expansion:  dx integer, max 200 units, cost €10/unit
  - Overtime:             ot_x integer, max 300 units/week, cost €2/unit/week
 
Workstation Y:
  - Base capacity:        60*24*7 - 80 = 10 000 min/week
  - Permanent expansion:  dy_pct integer, max 40%, cost €1500 per 1%-increment
  - Overtime:             ot_y continuous, max 38 hours/week, cost €120/hour/week
 
Effective capacity per period:
  X: (CAP_X_BASE + dx) + ot_x[t]
  Y: CAP_Y_BASE * (1 + dy_pct/100) + 60 * ot_y[t]
     = (CAP_Y_BASE + (CAP_Y_BASE/100)*dy_pct) + 60*ot_y[t]   [linearised]
"""
 
import gurobipy as gp
from gurobipy import GRB
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_FORECAST
)
 
# ── Helper functions ────────────────────────────────────────────────────────
def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents
 
def clean_num(val, tol=1e-6, digits=6):
    v = float(val)
    if abs(v) < tol:
        return 0.0
    return round(v, digits)
 
# ── Capacity parameters ─────────────────────────────────────────────────────
CAP_X_BASE       = 800
CAP_X_MAX_EXP    = 200
COST_EXP_X       = 10
CAP_X_OT_MAX     = 300
COST_OT_X        = 2
 
CAP_Y_BASE       = 60 * 24 * 7 - 80          # 10 000 min/week
CAP_Y_MAX_PCT    = 40
COST_EXP_Y_PCT   = 1_500
CAP_Y_OT_MAX     = 38                         # hours/week
COST_OT_Y        = 120
 
PROC_X = {END_PRODUCT: 1}
PROC_Y = {"B1401": 3, "B2302": 2}            # min/unit
 
# ── Build model ─────────────────────────────────────────────────────────────
model = gp.Model("APM_Assignment5a")
 
periods = range(1, T + 1)
parts   = PARTS
 
# ── Decision variables ──────────────────────────────────────────────────────
x      = model.addVars(parts, periods, name="x",  vtype=GRB.INTEGER, lb=0)
y      = model.addVars(parts, periods, name="y",  vtype=GRB.BINARY)
I      = model.addVars(parts, periods, name="I",  vtype=GRB.INTEGER, lb=0)
 
# Permanent expansion (once, for the whole horizon)
dx     = model.addVar(name="dx",     vtype=GRB.INTEGER, lb=0, ub=CAP_X_MAX_EXP)
dy_pct = model.addVar(name="dy_pct", lb=0, ub=CAP_Y_MAX_PCT)
 
# Overtime per period
ot_x   = model.addVars(periods, name="ot_x", vtype=GRB.INTEGER, lb=0, ub=CAP_X_OT_MAX)
ot_y   = model.addVars(periods, name="ot_y", lb=0, ub=CAP_Y_OT_MAX)
 
# ── Big-M ───────────────────────────────────────────────────────────────────
BIG_M = {i: sum(DEMAND_FORECAST) * 25 for i in parts}
 
# ── Objective ───────────────────────────────────────────────────────────────
model.setObjective(
    gp.quicksum(
        SETUP_COST[i] * y[i, t] + HOLDING_COST[i] * I[i, t]
        for i in parts for t in periods
    )
    # Permanent investment (one-off)
    + COST_EXP_X     * dx
    + COST_EXP_Y_PCT * dy_pct
    # Recurring overtime costs
    + gp.quicksum(
        COST_OT_X * ot_x[t] + COST_OT_Y * ot_y[t]
        for t in periods
    ),
    GRB.MINIMIZE
)
 
# ── Constraints ─────────────────────────────────────────────────────────────
for i in parts:
    for t in periods:
 
        # 1. Inventory balance
        inv_prev = INIT_INV[i] if t == 1 else I[i, t - 1]
        order_period = t - LEAD_TIME[i]
        receipts = x[i, order_period] if order_period >= 1 else 0
 
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
 
# 4. Workstation X capacity per period
#    regular base + permanent expansion + overtime
for t in periods:
    model.addConstr(
        gp.quicksum(PROC_X[i] * x[i, t] for i in PROC_X if i in parts)
        <= CAP_X_BASE + dx + ot_x[t],
        name=f"cap_X_{t}"
    )
 
# 5. Workstation Y capacity per period (linearised permanent expansion)
#    CAP_Y_BASE*(1 + dy_pct/100) + 60*ot_y[t]
for t in periods:
    model.addConstr(
        gp.quicksum(PROC_Y[i] * x[i, t] for i in PROC_Y if i in parts)
        <= CAP_Y_BASE + (CAP_Y_BASE / 100.0) * dy_pct + 60.0 * ot_y[t],
        name=f"cap_Y_{t}"
    )
 
# ── Solve ────────────────────────────────────────────────────────────────────
model.optimize()
 
# ── Output ───────────────────────────────────────────────────────────────────
if model.status == GRB.OPTIMAL:
    dx_val     = clean_num(dx.X)
    dy_pct_val = clean_num(dy_pct.X)
 
    invest_X   = COST_EXP_X     * dx_val
    invest_Y   = COST_EXP_Y_PCT * dy_pct_val
 
    total_setup   = sum(SETUP_COST[i]   * y[i, t].X for i in parts for t in periods)
    total_holding = sum(HOLDING_COST[i] * I[i, t].X for i in parts for t in periods)
    total_ot_x    = sum(COST_OT_X * ot_x[t].X for t in periods)
    total_ot_y    = sum(COST_OT_Y * ot_y[t].X for t in periods)
 
    cap_x_new = CAP_X_BASE + dx_val
    cap_y_new = CAP_Y_BASE * (1 + dy_pct_val / 100)
 
    print(f"\n{'='*65}")
    print("ASSIGNMENT 5a – OPTIMAL SOLUTION (Overtime + Permanent Expansion)")
    print(f"{'='*65}")
    print(f"Total cost:              €{clean_num(model.ObjVal):>12,.2f}")
    print(f"  Setup cost:            €{clean_num(total_setup):>12,.2f}")
    print(f"  Holding cost:          €{clean_num(total_holding):>12,.2f}")
    print(f"  Investment X:          €{invest_X:>12,.2f}  (+{dx_val:.0f} units → new base {cap_x_new:.0f} u/wk)")
    print(f"  Investment Y:          €{invest_Y:>12,.2f}  (+{dy_pct_val:.0f}% → new base {cap_y_new:.0f} min/wk)")
    print(f"  Overtime cost X:       €{clean_num(total_ot_x):>12,.2f}")
    print(f"  Overtime cost Y:       €{clean_num(total_ot_y):>12,.2f}")
    print(f"  Total overtime cost:   €{clean_num(total_ot_x + total_ot_y):>12,.2f}")
 
    print(f"\n{'─'*65}")
    print("PRODUCTION / ORDER SCHEDULE (non-zero quantities)")
    print(f"{'─'*65}")
    for i in parts:
        orders = [(t, clean_num(x[i, t].X)) for t in periods if x[i, t].X > 0.5]
        if orders:
            print(f"\n{i}:")
            for t, qty in orders:
                print(f"  Week {t:2d}: {qty:,.0f} units")
 
    print(f"\n{'─'*65}")
    print("END INVENTORY (week 30)")
    print(f"{'─'*65}")
    for i in parts:
        print(f"  {i}: {clean_num(I[i, T].X):,.0f} units")
 
    print(f"\n{'─'*65}")
    print("OVERTIME USAGE (only periods with overtime)")
    print(f"{'─'*65}")
    ot_used = False
    for t in periods:
        ox = clean_num(ot_x[t].X)
        oy = clean_num(ot_y[t].X)
        if ox > 1e-6 or oy > 1e-6:
            ot_used = True
            print(
                f"  Week {t:2d}: "
                f"X = {ox:6.0f} extra units | "
                f"Y = {oy:6.2f} extra hours ({clean_num(60*oy):7.1f} min)"
            )
    if not ot_used:
        print("  No overtime used.")
 
    print(f"\n{'─'*65}")
    print("WORKSTATION UTILISATION PER WEEK")
    print(f"{'─'*65}")
    print(f"  {'Wk':>2}  {'X used':>8} / {'X cap':>8}  {'X%':>6}  |"
          f"  {'Y used':>8} / {'Y cap':>8}  {'Y%':>6}")
    for t in periods:
        load_x   = clean_num(sum(PROC_X.get(i, 0) * x[i, t].X for i in parts))
        load_y   = clean_num(sum(PROC_Y.get(i, 0) * x[i, t].X for i in parts))
        cap_x_t  = cap_x_new + clean_num(ot_x[t].X)
        cap_y_t  = cap_y_new + 60 * clean_num(ot_y[t].X)
        pct_x    = 100 * load_x / cap_x_t if cap_x_t > 0 else 0
        pct_y    = 100 * load_y / cap_y_t if cap_y_t > 0 else 0
        print(f"  {t:2d}  {load_x:8.0f} / {cap_x_t:8.0f}  {pct_x:5.1f}%  |"
              f"  {load_y:8.0f} / {cap_y_t:8.0f}  {pct_y:5.1f}%")
 
    # ── Write JSON output ────────────────────────────────────────────────────
    output = {
        "status": "OPTIMAL",
        "total_cost": clean_num(model.ObjVal),
        "cost_breakdown": {
            "setup_cost":            clean_num(total_setup),
            "holding_cost":          clean_num(total_holding),
            "investment_cost_X":     invest_X,
            "investment_cost_Y":     invest_Y,
            "overtime_cost_X":       clean_num(total_ot_x),
            "overtime_cost_Y":       clean_num(total_ot_y),
            "total_overtime_cost":   clean_num(total_ot_x + total_ot_y),
        },
        "capacity_parameters": {
            "X_base_units":          CAP_X_BASE,
            "X_permanent_expansion": dx_val,
            "X_new_base":            cap_x_new,
            "X_overtime_max_units":  CAP_X_OT_MAX,
            "X_cost_expansion":      COST_EXP_X,
            "X_cost_overtime":       COST_OT_X,
            "Y_base_minutes":        CAP_Y_BASE,
            "Y_permanent_pct":       dy_pct_val,
            "Y_new_base_minutes":    cap_y_new,
            "Y_overtime_max_hours":  CAP_Y_OT_MAX,
            "Y_cost_expansion_pct":  COST_EXP_Y_PCT,
            "Y_cost_overtime_hour":  COST_OT_Y,
        },
        "production_schedule": {
            i: {str(t): clean_num(x[i, t].X) for t in periods if x[i, t].X > 0.5}
            for i in parts
        },
        "inventory": {
            i: {str(t): clean_num(I[i, t].X) for t in periods}
            for i in parts
        },
        "overtime_usage": {
            str(t): {
                "X_overtime_units":  clean_num(ot_x[t].X),
                "Y_overtime_hours":  clean_num(ot_y[t].X),
                "Y_overtime_minutes": clean_num(60 * ot_y[t].X),
                "cost_X":            clean_num(COST_OT_X * ot_x[t].X),
                "cost_Y":            clean_num(COST_OT_Y * ot_y[t].X),
            }
            for t in periods
        }
    }
    with open("output_5a.json", "w") as f:
        json.dump(output, f, indent=2)
    print("\nResults written to output_5a.json")
 
else:
    print(f"Model status: {model.status} — no optimal solution found.")
 