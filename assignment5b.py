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

# Production schedule: x[i][t] (t as string in JSON)
schedule = plan_5a["production_schedule"]

def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents

periods = range(1, T + 1)
parts   = PARTS

# ── Simulate inventory with realized demand ──────────────────────────────────
inventory  = {i: {0: INIT_INV[i]} for i in parts}   # inventory[i][t]
backorders = {t: 0 for t in periods}                  # backorders per period (E2801 only)

for t in periods:
    for i in parts:
        inv_prev = inventory[i][t - 1]

        # Receipts: order placed t - LEAD_TIME[i] periods ago
        order_period = t - LEAD_TIME[i]
        receipts = float(schedule.get(i, {}).get(str(order_period), 0)) if order_period >= 1 else 0

        # Internal demand from BOM
        internal_demand = sum(
            BOM[p][i] * float(schedule.get(p, {}).get(str(t), 0))
            for p in get_parents(i)
        )

        # External demand (only E2801, realized)
        external_demand = DEMAND_REALIZED[t - 1] if i == END_PRODUCT else 0

        net = inv_prev + receipts - internal_demand - external_demand

        if net < 0:
            backorders[t] += abs(net)   # only possible for E2801
            inventory[i][t] = 0
        else:
            inventory[i][t] = net

# ── Cost calculation ─────────────────────────────────────────────────────────
total_setup   = sum(
    SETUP_COST[i] * (1 if float(schedule.get(i, {}).get(str(t), 0)) > 0.5 else 0)
    for i in parts for t in periods
)
total_holding = sum(
    HOLDING_COST[i] * inventory[i][t]
    for i in parts for t in periods
)
total_backorder = sum(BACKORDER_COST * backorders[t] for t in periods)

# Investment and overtime costs stay the same as 5a (fixed plan)
invest_X   = plan_5a["cost_breakdown"]["investment_cost_X"]
invest_Y   = plan_5a["cost_breakdown"]["investment_cost_Y"]
total_ot_x = plan_5a["cost_breakdown"]["overtime_cost_X"]
total_ot_y = plan_5a["cost_breakdown"]["overtime_cost_Y"]

total_cost = total_setup + total_holding + total_backorder + invest_X + invest_Y + total_ot_x + total_ot_y

# ── Service level & fill rate ────────────────────────────────────────────────
# Service level: fraction of periods with NO backorder
periods_with_bo = sum(1 for t in periods if backorders[t] > 0)
service_level   = (T - periods_with_bo) / T

# Fill rate: fraction of total demand that was delivered on time
total_demand    = sum(DEMAND_REALIZED)
total_backorder_units = sum(backorders[t] for t in periods)
fill_rate       = 1 - total_backorder_units / total_demand

# ── Print results ─────────────────────────────────────────────────────────────
print(f"\n{'='*65}")
print("ASSIGNMENT 5b – REALIZED DEMAND EVALUATION")
print(f"{'='*65}")
print(f"Total cost:              €{total_cost:>12,.2f}")
print(f"  Setup cost:            €{total_setup:>12,.2f}")
print(f"  Holding cost:          €{total_holding:>12,.2f}")
print(f"  Backorder cost:        €{total_backorder:>12,.2f}")
print(f"  Investment X:          €{invest_X:>12,.2f}")
print(f"  Investment Y:          €{invest_Y:>12,.2f}")
print(f"  Overtime cost X:       €{total_ot_x:>12,.2f}")
print(f"  Overtime cost Y:       €{total_ot_y:>12,.2f}")

print(f"\n{'─'*65}")
print(f"Service level:           {100*service_level:>8.2f}%  ({T - periods_with_bo}/{T} periods without backorder)")
print(f"Fill rate:               {100*fill_rate:>8.2f}%  ({total_demand - total_backorder_units:.0f}/{total_demand:.0f} units delivered on time)")

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
    "total_cost": total_cost,
    "cost_breakdown": {
        "setup_cost":          total_setup,
        "holding_cost":        total_holding,
        "backorder_cost":      total_backorder,
        "investment_cost_X":   invest_X,
        "investment_cost_Y":   invest_Y,
        "overtime_cost_X":     total_ot_x,
        "overtime_cost_Y":     total_ot_y,
    },
    "performance": {
        "service_level":          service_level,
        "fill_rate":              fill_rate,
        "total_demand_realized":  total_demand,
        "total_units_backordered": total_backorder_units,
        "periods_with_backorder": periods_with_bo,
    },
    "backorders": {str(t): backorders[t] for t in periods},
    "inventory":  {i: {str(t): inventory[i][t] for t in periods} for i in parts}
}

with open("output_5b.json", "w") as f:
    json.dump(output, f, indent=2)
print("\nResults written to output_5b.json")