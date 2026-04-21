"""
APM Project 2026 - Assignment 2b
Evaluate the plan from 2a (finite capacity) against REALIZED demand.
Calculates backorders, service level, fill rate, and updated costs.
"""

import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME,
    INIT_INV, SETUP_COST, HOLDING_COST,
    DEMAND_REALIZED, BACKORDER_COST
)

# ── Load production schedule from 2a ──────────────────────────────────────
with open("output_2a.json") as f:
    plan = json.load(f)

schedule = plan["production_schedule"]  # {part: {period: qty}}

periods = range(1, T + 1)

# ── Simulate with realized demand ─────────────────────────────────────────
inventory  = {i: INIT_INV[i] for i in PARTS}
backorders = {t: 0 for t in periods}   # backorders of end product per period
delivered  = {t: 0 for t in periods}   # units delivered per period

holding_cost         = 0.0
setup_cost           = 0.0
backorder_cost_total = 0.0

periods_fully_met = 0
demand_per_period = DEMAND_REALIZED

for t in periods:
    # 1. Receive arrivals (orders placed t - lead_time periods ago)
    for i in PARTS:
        lt = LEAD_TIME[i]
        source_period = str(t - lt)
        qty = schedule.get(i, {}).get(source_period, 0)
        inventory[i] += qty

        # Setup cost: incurred whenever an order arrives
        if qty > 0:
            setup_cost += SETUP_COST[i]

    # 2. Consume components for internal production
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

# ── Performance metrics ────────────────────────────────────────────────────
total_demand    = sum(DEMAND_REALIZED)
total_delivered = sum(delivered.values())
total_backorder = sum(backorders.values())

service_level = periods_fully_met / T * 100
fill_rate     = total_delivered / total_demand * 100
total_cost    = setup_cost + holding_cost + backorder_cost_total

# ── Print results ──────────────────────────────────────────────────────────
print(f"\n{'='*60}")
print(f"ASSIGNMENT 2b - REALIZED DEMAND EVALUATION (Finite Capacity Plan)")
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
any_bo = False
for t in periods:
    if backorders[t] > 0:
        any_bo = True
        print(f"  Week {t:2d}: {backorders[t]:,} units backordered")
if not any_bo:
    print("  No backorders — all demand met on time!")

# ── Comparison with 1b ─────────────────────────────────────────────────────
try:
    with open("output_1b.json") as f:
        res_1b = json.load(f)
    print(f"\n{'─'*60}")
    print("COMPARISON: 2b (finite capacity) vs 1b (infinite capacity)")
    print(f"{'─'*60}")
    print(f"{'Metric':<25} {'1b':>12} {'2b':>12} {'Difference':>12}")
    print(f"{'─'*60}")

    def diff(v2, v1, fmt=".2f"):
        d = v2 - v1
        sign = "+" if d >= 0 else ""
        return f"{sign}{d:{fmt}}"

    print(f"{'Total cost (€)':<25} {res_1b['costs']['total_cost']:>12,.2f} {total_cost:>12,.2f} {diff(total_cost, res_1b['costs']['total_cost']):>12}")
    print(f"{'Setup cost (€)':<25} {res_1b['costs']['setup_cost']:>12,.2f} {setup_cost:>12,.2f} {diff(setup_cost, res_1b['costs']['setup_cost']):>12}")
    print(f"{'Holding cost (€)':<25} {res_1b['costs']['holding_cost']:>12,.2f} {holding_cost:>12,.2f} {diff(holding_cost, res_1b['costs']['holding_cost']):>12}")
    print(f"{'Backorder cost (€)':<25} {res_1b['costs']['backorder_cost']:>12,.2f} {backorder_cost_total:>12,.2f} {diff(backorder_cost_total, res_1b['costs']['backorder_cost']):>12}")
    print(f"{'Service level (%)':<25} {res_1b['service_metrics']['service_level_pct']:>12.1f} {service_level:>12.1f} {diff(service_level, res_1b['service_metrics']['service_level_pct'], '.1f'):>12}")
    print(f"{'Fill rate (%)':<25} {res_1b['service_metrics']['fill_rate_pct']:>12.1f} {fill_rate:>12.1f} {diff(fill_rate, res_1b['service_metrics']['fill_rate_pct'], '.1f'):>12}")

except FileNotFoundError:
    print("\n(output_1b.json not found — skipping comparison table)")

# ── Write output ───────────────────────────────────────────────────────────
output_2b = {
    "total_cost": total_cost,
    "setup_cost": setup_cost,
    "holding_cost": holding_cost,
    "backorder_cost": backorder_cost_total,
    "service_level_pct": service_level,
    "fill_rate_pct": fill_rate,
    "backorders": {str(t): backorders[t] for t in periods},
    "delivered":  {str(t): delivered[t]  for t in periods}
}
with open("output_2b.json", "w") as f:
    json.dump(output_2b, f, indent=2)
print("\nResults written to output_2b.json")