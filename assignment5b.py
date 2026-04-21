"""
APM Project 2026 - Assignment 5b
Evaluate the production plan from Assignment 5a using REALIZED demand.

Approach:
  - Fix the production schedule from output_5a.json
  - Simulate inventory with realized demand
  - Backorder if inventory goes negative, at €250/unit/period
  - Report cost types, service level and fill rate
"""

import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME,
    INIT_INV, SETUP_COST, HOLDING_COST, DEMAND_REALIZED
)

BACKORDER_COST = 250

# ── Load production schedule from 5a ────────────────────────────────────────
with open("output_5a.json", "r") as f:
    plan_5a = json.load(f)

schedule = plan_5a["production_schedule"]

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

periods = range(1, T + 1)
parts   = PARTS

# ── State tracking ───────────────────────────────────────────────────────────
inventory = {i: {} for i in parts}     # positive inventory only
held      = {i: {} for i in parts}     # inventory used for holding cost

# End product backlog carried through time
backorders = {t: 0.0 for t in periods}
net_end    = {}

# ── Simulate all non-end-product parts ──────────────────────────────────────
for i in parts:
    if i == END_PRODUCT:
        continue

    inv_prev = INIT_INV[i]

    for t in periods:
        order_period = t - LEAD_TIME[i]
        receipts = float(schedule.get(i, {}).get(str(order_period), 0)) if order_period >= 1 else 0.0

        internal_demand = sum(
            BOM[p][i] * float(schedule.get(p, {}).get(str(t), 0))
            for p in get_parents(i)
        )

        net = inv_prev + receipts - internal_demand

        inventory[i][t] = clean_num(net)
        held[i][t] = clean_num(max(0.0, net))

        inv_prev = net

# ── Simulate end product with carried backorders ────────────────────────────
inv_prev = INIT_INV[END_PRODUCT]
bo_prev  = 0.0

for t in periods:
    order_period = t - LEAD_TIME[END_PRODUCT]
    receipts = float(schedule.get(END_PRODUCT, {}).get(str(order_period), 0)) if order_period >= 1 else 0.0

    demand_t = DEMAND_REALIZED[t - 1]

    # backlog is carried into next period
    net = inv_prev - bo_prev + receipts - demand_t

    if net >= 0:
        inventory[END_PRODUCT][t] = clean_num(net)
        held[END_PRODUCT][t] = clean_num(net)
        backorders[t] = 0.0
    else:
        inventory[END_PRODUCT][t] = 0.0
        held[END_PRODUCT][t] = 0.0
        backorders[t] = clean_num(-net)

    net_end[t] = clean_num(net)

    inv_prev = inventory[END_PRODUCT][t]
    bo_prev  = backorders[t]

# ── Cost calculation ─────────────────────────────────────────────────────────
# Setup cost stays fixed because the production plan from 5a is fixed
total_setup = sum(
    SETUP_COST[i] * (1 if float(schedule.get(i, {}).get(str(t), 0)) > 0.5 else 0)
    for i in parts for t in periods
)

total_holding = sum(
    HOLDING_COST[i] * held[i][t]
    for i in parts for t in periods
)

total_backorder = sum(
    BACKORDER_COST * backorders[t]
    for t in periods
)

# Investment and overtime costs stay exactly the same as in 5a
invest_X   = float(plan_5a["cost_breakdown"]["investment_cost_X"])
invest_Y   = float(plan_5a["cost_breakdown"]["investment_cost_Y"])
total_ot_x = float(plan_5a["cost_breakdown"]["overtime_cost_X"])
total_ot_y = float(plan_5a["cost_breakdown"]["overtime_cost_Y"])

total_cost = (
    total_setup
    + total_holding
    + total_backorder
    + invest_X
    + invest_Y
    + total_ot_x
    + total_ot_y
)

# ── Service level & fill rate ────────────────────────────────────────────────
periods_no_bo = sum(1 for t in periods if backorders[t] == 0)
service_level = periods_no_bo / T

total_demand = sum(DEMAND_REALIZED)

# Same logic as in 1b: count only NEW backlog creation
total_new_bo = 0.0
prev_bo = 0.0
for t in periods:
    new_bo = max(0.0, backorders[t] - prev_bo)
    total_new_bo += new_bo
    prev_bo = backorders[t]

fill_rate = 1.0 - (total_new_bo / total_demand) if total_demand > 0 else 1.0
units_on_time = total_demand - total_new_bo

# ── Print results ────────────────────────────────────────────────────────────
print(f"\n{'='*65}")
print("ASSIGNMENT 5b – REALIZED DEMAND EVALUATION")
print(f"{'='*65}")
print(f"Total cost:              €{clean_num(total_cost):>12,.2f}")
print(f"  Setup cost:            €{clean_num(total_setup):>12,.2f}")
print(f"  Holding cost:          €{clean_num(total_holding):>12,.2f}")
print(f"  Backorder cost:        €{clean_num(total_backorder):>12,.2f}")
print(f"  Investment X:          €{clean_num(invest_X):>12,.2f}")
print(f"  Investment Y:          €{clean_num(invest_Y):>12,.2f}")
print(f"  Overtime cost X:       €{clean_num(total_ot_x):>12,.2f}")
print(f"  Overtime cost Y:       €{clean_num(total_ot_y):>12,.2f}")

print(f"\n{'─'*65}")
print(f"Service level:           {100*service_level:>8.2f}%  ({periods_no_bo}/{T} periods without backorder)")
print(f"Fill rate:               {100*fill_rate:>8.2f}%  ({clean_num(units_on_time):.0f}/{total_demand:.0f} units delivered on time)")

print(f"\n{'─'*65}")
print("BACKORDERS PER PERIOD (only periods with backorders)")
print(f"{'─'*65}")
bo_found = False
for t in periods:
    if backorders[t] > 0:
        bo_found = True
        print(f"  Week {t:2d}: {backorders[t]:,.0f} units backordered")
if not bo_found:
    print("  No backorders.")

print(f"\n{'─'*65}")
print("END INVENTORY PER PART (week 30)")
print(f"{'─'*65}")
for i in parts:
    print(f"  {i}: {inventory[i][T]:,.0f} units")

# ── Write JSON output ────────────────────────────────────────────────────────
output = {
    "status": "EVALUATED",
    "based_on_plan": "output_5a.json",
    "demand_type": "realized",
    "total_cost": clean_num(total_cost),
    "cost_breakdown": {
        "setup_cost":         clean_num(total_setup),
        "holding_cost":       clean_num(total_holding),
        "backorder_cost":     clean_num(total_backorder),
        "investment_cost_X":  clean_num(invest_X),
        "investment_cost_Y":  clean_num(invest_Y),
        "overtime_cost_X":    clean_num(total_ot_x),
        "overtime_cost_Y":    clean_num(total_ot_y),
    },
    "performance": {
        "service_level":            clean_num(service_level),
        "service_level_pct":        clean_num(100 * service_level),
        "fill_rate":                clean_num(fill_rate),
        "fill_rate_pct":            clean_num(100 * fill_rate),
        "total_demand_realized":    total_demand,
        "units_on_time":            clean_num(units_on_time),
        "periods_no_backorder":     periods_no_bo,
        "total_units_backordered":  clean_num(sum(backorders[t] for t in periods)),
    },
    "backorders": {
        str(t): clean_num(backorders[t]) for t in periods
    },
    "end_product_net_position": {
        str(t): clean_num(net_end[t]) for t in periods
    },
    "inventory": {
        i: {str(t): clean_num(inventory[i][t]) for t in periods}
        for i in parts
    }
}

with open("output_5b.json", "w") as f:
    json.dump(output, f, indent=2)

print("\nResults written to output_5b.json")