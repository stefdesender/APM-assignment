"""
APM Project 2026 - Assignment 1b
Evaluate the plan from 1a against REALIZED demand.
Calculates backorders, service level, fill rate, and updated costs.
"""

import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME, MIN_LOT,
    INIT_INV, SETUP_COST, HOLDING_COST,
    DEMAND_FORECAST, DEMAND_REALIZED, BACKORDER_COST
)

# ── Load production schedule from 1a ──────────────────────────────────────
with open("output_1a.json") as f:
    plan = json.load(f)

schedule = plan["production_schedule"]  # {part: {period: qty}}

periods = range(1, T + 1)

# ── Simulate with realized demand ─────────────────────────────────────────
inventory  = {i: INIT_INV[i] for i in PARTS}
backorders = {t: 0 for t in periods}   # backorders of end product per period
delivered  = {t: 0 for t in periods}   # units delivered per period

holding_cost  = 0.0
setup_cost    = 0.0
backorder_cost_total = 0.0

# Track demand met per period for service level / fill rate
demand_per_period   = DEMAND_REALIZED
periods_fully_met   = 0

for t in periods:
    # 1. Receive arrivals (orders placed t - lead_time periods ago)
    for i in PARTS:
        lt = LEAD_TIME[i]
        source_period = str(t - lt)
        qty = schedule.get(i, {}).get(source_period, 0)
        inventory[i] += qty

        # Setup cost: incurred when order was placed (already counted in 1a)
        # For 1b we recalculate setup cost from the same schedule
        if qty > 0:
            setup_cost += SETUP_COST[i]

    # 2. Consume components for internal production
    # (assuming production follows the same schedule)
    for parent, children in BOM.items():
        prod_qty = schedule.get(parent, {}).get(str(t), 0)
        if prod_qty > 0:
            for child, qty_per_unit in children.items():
                inventory[child] -= prod_qty * qty_per_unit

    # 3. Fulfill realized end-product demand
    d = demand_per_period[t - 1]
    available = inventory[END_PRODUCT]

    if available >= d:
        delivered[t] = d
        inventory[END_PRODUCT] -= d
        periods_fully_met += 1
    else:
        delivered[t] = available
        backorders[t] = d - available
        inventory[END_PRODUCT] = 0

    # 4. Accumulate holding costs (end of period)
    for i in PARTS:
        if inventory[i] > 0:
            holding_cost += HOLDING_COST[i] * inventory[i]

    # 5. Backorder cost
    backorder_cost_total += BACKORDER_COST * backorders[t]

# ── Results ────────────────────────────────────────────────────────────────
total_demand    = sum(DEMAND_REALIZED)
total_delivered = sum(delivered.values())
total_backorder = sum(backorders.values())

service_level = periods_fully_met / T * 100
fill_rate     = total_delivered / total_demand * 100
total_cost    = setup_cost + holding_cost + backorder_cost_total

print(f"\n{'='*60}")
print(f"ASSIGNMENT 1b - REALIZED DEMAND EVALUATION")
print(f"{'='*60}")
print(f"Total cost:         €{total_cost:,.2f}")
print(f"  Setup cost:       €{setup_cost:,.2f}")
print(f"  Holding cost:     €{holding_cost:,.2f}")
print(f"  Backorder cost:   €{backorder_cost_total:,.2f}")
print(f"\nTotal demand:       {total_demand:,} units")
print(f"Total delivered:    {total_delivered:,} units")
print(f"Total backordered:  {total_backorder:,} units")
print(f"\nService level:      {service_level:.1f}%  (periods fully met)")
print(f"Fill rate:          {fill_rate:.1f}%  (units delivered / demanded)")

print(f"\n{'─'*60}")
print("BACKORDERS PER PERIOD:")
for t in periods:
    if backorders[t] > 0:
        print(f"  Week {t:2d}: {backorders[t]:,} units backordered")

# Write output
output_1b = {
    "total_cost": total_cost,
    "setup_cost": setup_cost,
    "holding_cost": holding_cost,
    "backorder_cost": backorder_cost_total,
    "service_level_pct": service_level,
    "fill_rate_pct": fill_rate,
    "backorders": backorders,
    "delivered": delivered
}
with open("output_1b.json", "w") as f:
    json.dump(output_1b, f, indent=2)
print("\nResults written to output_1b.json")