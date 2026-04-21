"""
APM Project 2026 - Assignment 4b
Evaluate the Assignment 4a production plan against REALIZED demand.

What changes vs 4a:
  - The production schedule (x[i,t]) is FIXED to the 4a optimal plan
    (read from output_4a.json).
  - Demand is replaced by DEMAND_REALIZED from input_data.py.
  - Unmet demand is backordered at BACKORDER_COST per unit per period.
  - We compute:
      * Total cost (setup + holding + backorder + investment)
      * Service level  = fraction of periods with zero backorder for end product
      * Fill rate      = fraction of demand satisfied immediately
  - Results are compared with 4a.
"""

import json
import gurobipy as gp
from gurobipy import GRB
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_REALIZED, BACKORDER_COST
)

# ── Workstation parameters (same as 4a) ───────────────────────────────────
CAP_X_BASE  = 800
PROC_X      = {END_PRODUCT: 1}
CAP_Y_BASE  = 60 * 24 * 7 - 80          # 10 000 min/week
PROC_Y      = {"B1401": 3, "B2302": 2}

COST_EXP_X     = 10
COST_EXP_Y_PCT = 1_500

# ── Load 4a results ───────────────────────────────────────────────────────
with open("output_4a.json") as f:
    plan_4a = json.load(f)

# Fixed production quantities from 4a
fixed_x = {
    i: {int(t): qty for t, qty in plan_4a["production_schedule"].get(i, {}).items()}
    for i in PARTS
}

def get_x(i, t):
    return fixed_x[i].get(t, 0.0)

# Permanent expansion decisions carried over from 4a
dx_val     = plan_4a["expansion_X_units"]
dy_pct_val = plan_4a["expansion_Y_pct"]
invest_X   = plan_4a["investment_cost_X"]
invest_Y   = plan_4a["investment_cost_Y"]

# ── Helper ────────────────────────────────────────────────────────────────
def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

# ── Build evaluation model ────────────────────────────────────────────────
model = gp.Model("APM_Assignment4b")
model.Params.OutputFlag = 0

periods = range(1, T + 1)
parts   = PARTS

# Inventory (end of period) and backorder variables
I  = model.addVars(parts, periods, name="I",vtype=GRB.INTEGER,  lb=0)
BO = model.addVars(periods,        name="BO", lb=0)   # cumulative backorder, end product only

# Setup cost is determined by the fixed schedule (no setup variables needed)
setup_cost_fixed = sum(
    SETUP_COST[i] for i in parts for t in periods if get_x(i, t) > 0.5
)

# ── Objective: holding + backorder costs (setup & investment fixed from 4a) ─
model.setObjective(
    gp.quicksum(HOLDING_COST[i] * I[i, t] for i in parts for t in periods)
    + BACKORDER_COST * gp.quicksum(BO[t] for t in periods),
    GRB.MINIMIZE
)

# ── Constraints ────────────────────────────────────────────────────────────
for i in parts:
    for t in periods:
        inv_prev = INIT_INV[i] if t == 1 else I[i, t - 1]

        order_placement_period = t - LEAD_TIME[i]
        receipts = get_x(i, order_placement_period) if order_placement_period >= 1 else 0

        internal_demand = sum(
            BOM[p][i] * get_x(p, t) for p in get_parents(i)
        )

        if i == END_PRODUCT:
            real_demand = DEMAND_REALIZED[t - 1]
            bo_prev = BO[t - 1] if t > 1 else 0
            # I[t] - BO[t] = I[t-1] - BO[t-1] + receipts - demand
            model.addConstr(
                I[i, t] - BO[t] == inv_prev - bo_prev + receipts - internal_demand - real_demand,
                name=f"inv_balance_{i}_{t}"
            )
        else:
            model.addConstr(
                I[i, t] == inv_prev + receipts - internal_demand,
                name=f"inv_balance_{i}_{t}"
            )

# ── Solve ──────────────────────────────────────────────────────────────────
model.optimize()

# ── Results ───────────────────────────────────────────────────────────────
if model.status in (GRB.OPTIMAL, GRB.SUBOPTIMAL):

    total_holding   = sum(HOLDING_COST[i] * I[i,t].X for i in parts for t in periods)
    total_backorder = BACKORDER_COST * sum(BO[t].X for t in periods)
    total_cost_4b   = setup_cost_fixed + total_holding + total_backorder + invest_X + invest_Y

    # ── Service level ──────────────────────────────────────────────────────
    # Fraction of periods with zero backorder
    periods_no_bo = sum(1 for t in periods if BO[t].X < 0.5)
    service_level = periods_no_bo / T

    # ── Fill rate ──────────────────────────────────────────────────────────
    # Fraction of total realized demand satisfied without backorder
    total_realized = sum(DEMAND_REALIZED)
    new_bo = []
    for t in periods:
        bo_t   = BO[t].X
        bo_tm1 = BO[t - 1].X if t > 1 else 0.0
        new_bo.append(max(0.0, bo_t - bo_tm1))
    total_new_bo = sum(new_bo)
    fill_rate = 1.0 - total_new_bo / total_realized

    # ── Print results ──────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"ASSIGNMENT 4b — REALIZED DEMAND EVALUATION")
    print(f"{'='*60}")

    print(f"\n── Cost breakdown ─────────────────────────────────────────")
    print(f"  Setup cost (from 4a plan):   €{setup_cost_fixed:>12,.2f}")
    print(f"  Holding cost:                €{total_holding:>12,.2f}")
    print(f"  Backorder cost:              €{total_backorder:>12,.2f}")
    print(f"  Investment cost X:           €{invest_X:>12,.2f}")
    print(f"  Investment cost Y:           €{invest_Y:>12,.2f}")
    print(f"  {'─'*42}")
    print(f"  TOTAL cost (4b):             €{total_cost_4b:>12,.2f}")

    print(f"\n── Comparison with 4a ─────────────────────────────────────")
    total_cost_4a = plan_4a["total_cost"]
    print(f"  Total cost 4a (forecast):    €{total_cost_4a:>12,.2f}")
    print(f"  Total cost 4b (realized):    €{total_cost_4b:>12,.2f}")
    print(f"  Difference (4b − 4a):        €{total_cost_4b - total_cost_4a:>+12,.2f}")

    print(f"\n── Service metrics ────────────────────────────────────────")
    print(f"  Service level:   {service_level*100:6.2f}%  ({periods_no_bo}/{T} periods without backorder)")
    print(f"  Fill rate:       {fill_rate*100:6.2f}%  "
          f"({total_realized - total_new_bo:,.0f} / {total_realized:,.0f} units met immediately)")

    print(f"\n── Backorder detail per period (end product) ──────────────")
    print(f"  {'Week':>4}  {'Forecast':>9}  {'Realized':>9}  {'Received':>9}  {'New BO':>8}  {'Cum BO':>8}")
    for t in periods:
        from input_data import DEMAND_FORECAST
        rec = get_x(END_PRODUCT, t - LEAD_TIME[END_PRODUCT]) if (t - LEAD_TIME[END_PRODUCT]) >= 1 else 0
        print(f"  {t:>4}  {DEMAND_FORECAST[t-1]:>9.0f}  {DEMAND_REALIZED[t-1]:>9.0f}  "
              f"{rec:>9.0f}  {new_bo[t-1]:>8.1f}  {BO[t].X:>8.1f}")

    print(f"\n── Workstation utilisation (4a schedule, expanded caps) ───")
    cap_x_new = CAP_X_BASE + dx_val
    cap_y_new = CAP_Y_BASE * (1 + dy_pct_val / 100)
    print(f"  Cap X = {cap_x_new:.0f} u/week   Cap Y = {cap_y_new:.0f} min/week")
    for t in periods:
        load_x = sum(PROC_X.get(i, 0) * get_x(i, t) for i in parts)
        load_y = sum(PROC_Y.get(i, 0) * get_x(i, t) for i in parts)
        if load_x > 0 or load_y > 0:
            print(f"  Week {t:2d}: X {load_x:6.0f}/{cap_x_new:.0f} ({100*load_x/cap_x_new:5.1f}%)   "
                  f"Y {load_y:7.0f}/{cap_y_new:.0f} ({100*load_y/cap_y_new:5.1f}%)")

    # ── Write output ──────────────────────────────────────────────────────
    output = {
        "status": "EVALUATED",
        "total_cost_4b": total_cost_4b,
        "setup_cost": setup_cost_fixed,
        "holding_cost": total_holding,
        "backorder_cost": total_backorder,
        "investment_cost_X": invest_X,
        "investment_cost_Y": invest_Y,
        "service_level": service_level,
        "fill_rate": fill_rate,
        "total_new_backorder_units": total_new_bo,
        "comparison": {
            "total_cost_4a": total_cost_4a,
            "total_cost_4b": total_cost_4b,
            "difference": total_cost_4b - total_cost_4a
        },
        "backorder_per_period": {str(t): BO[t].X for t in periods},
        "inventory": {
            i: {str(t): I[i,t].X for t in periods} for i in parts
        }
    }
    with open("output_4b.json", "w") as f:
        json.dump(output, f, indent=2)
    print("\nResults written to output_4b.json")

else:
    print(f"Model status: {model.status} — evaluation failed.")