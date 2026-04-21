"""
APM Project 2026 - Assignment 3b

Evaluate the fixed production plan from Assignment 3a
under REALIZED demand, allowing BACKORDERS.

Important:
- No re-optimization is done in 3b
- The production schedule from output_3a.json is kept fixed
- Overtime is NOT newly chosen in 3b
- Any unmet realized demand becomes backorder
"""

import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME,
    INIT_INV, HOLDING_COST, DEMAND_REALIZED, BACKORDER_COST
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

# ── Read fixed plan from 3a ────────────────────────────────────────────────
with open("output_3a.json", "r") as f:
    plan_3a = json.load(f)

schedule = plan_3a["production_schedule"]
cost_3a = plan_3a["total_cost"]
setup_cost_3a = plan_3a["cost_breakdown"]["setup_cost"]
overtime_cost_x_3a = plan_3a["cost_breakdown"]["overtime_cost_X"]
overtime_cost_y_3a = plan_3a["cost_breakdown"]["overtime_cost_Y"]
total_overtime_cost_3a = plan_3a["cost_breakdown"]["total_overtime_cost"]

periods = range(1, T + 1)

# ── State tracking ──────────────────────────────────────────────────────────
inv = {i: {} for i in PARTS}          # positive inventory only for components
h_pos = {i: {} for i in PARTS}        # held inventory for holding cost

# For end product we track net inventory and backlog explicitly
net_end = {}
bo = {}                               # backorders of end product
held_end = {}

# ── Simulate all components except end product ─────────────────────────────
for i in PARTS:
    if i == END_PRODUCT:
        continue

    inv_prev = INIT_INV[i]

    for t in periods:
        order_period = t - LEAD_TIME[i]
        if order_period >= 1:
            receipts = float(schedule.get(i, {}).get(str(order_period), 0))
        else:
            receipts = 0.0

        internal_demand = 0.0
        for p, qty_per_parent in get_parents(i).items():
            parent_prod = float(schedule.get(p, {}).get(str(t), 0))
            internal_demand += qty_per_parent * parent_prod

        net = inv_prev + receipts - internal_demand

        inv[i][t] = clean_num(net)
        h_pos[i][t] = clean_num(max(0.0, net))

        inv_prev = net

# ── Simulate end product with realized demand + backorders ─────────────────
inv_prev = INIT_INV[END_PRODUCT]
bo_prev = 0.0

for t in periods:
    order_period = t - LEAD_TIME[END_PRODUCT]
    if order_period >= 1:
        receipts = float(schedule.get(END_PRODUCT, {}).get(str(order_period), 0))
    else:
        receipts = 0.0

    demand_t = DEMAND_REALIZED[t - 1]

    # Net stock position after satisfying current demand and previous backlog
    net = inv_prev - bo_prev + receipts - demand_t

    end_inventory = max(0.0, net)
    end_backorder = max(0.0, -net)

    net_end[t] = clean_num(net)
    bo[t] = clean_num(end_backorder)
    held_end[t] = clean_num(end_inventory)

    inv[END_PRODUCT][t] = clean_num(end_inventory)
    h_pos[END_PRODUCT][t] = clean_num(end_inventory)

    inv_prev = end_inventory
    bo_prev = end_backorder

# ── Cost calculation ────────────────────────────────────────────────────────
# Setup and overtime stay exactly the same as in 3a because the plan is fixed
total_setup = setup_cost_3a
total_ot_x = overtime_cost_x_3a
total_ot_y = overtime_cost_y_3a
total_overtime = total_overtime_cost_3a

total_holding = sum(
    HOLDING_COST[i] * h_pos[i][t]
    for i in PARTS
    for t in periods
)

total_backorder = sum(
    BACKORDER_COST * bo[t]
    for t in periods
)

total_cost_3b = total_setup + total_holding + total_overtime + total_backorder

# ── Service metrics ─────────────────────────────────────────────────────────
periods_no_bo = sum(1 for t in periods if bo[t] == 0)
service_level = periods_no_bo / T

total_demand = sum(DEMAND_REALIZED)

total_new_bo = 0.0
prev_bo = 0.0
for t in periods:
    new_bo = max(0.0, bo[t] - prev_bo)
    total_new_bo += new_bo
    prev_bo = bo[t]

fill_rate = 1.0 - (total_new_bo / total_demand) if total_demand > 0 else 1.0

units_on_time = total_demand - total_new_bo

# ── Print results ───────────────────────────────────────────────────────────
print(f"\n{'='*70}")
print("ASSIGNMENT 3b — EVALUATION OF 3a PLAN UNDER REALIZED DEMAND")
print(f"{'='*70}")
print(f"Setup cost:           €{clean_num(total_setup):>12,.2f}")
print(f"Holding cost:         €{clean_num(total_holding):>12,.2f}")
print(f"Overtime cost X:      €{clean_num(total_ot_x):>12,.2f}")
print(f"Overtime cost Y:      €{clean_num(total_ot_y):>12,.2f}")
print(f"Backorder cost:       €{clean_num(total_backorder):>12,.2f}")
print(f"{'─'*47}")
print(f"Total cost (3b):      €{clean_num(total_cost_3b):>12,.2f}")
print(f"Total cost (3a):      €{clean_num(cost_3a):>12,.2f}")
print(f"Difference:           €{clean_num(total_cost_3b - cost_3a):>+12,.2f}")

print(f"\n{'─'*70}")
print("SERVICE PERFORMANCE")
print(f"{'─'*70}")
print(f"Service level:        {service_level*100:.2f}% ({periods_no_bo}/{T} periods without backorder)")
print(f"Fill rate:            {fill_rate*100:.2f}% ({clean_num(units_on_time):,.0f} / {total_demand:,.0f} units on time)")

print(f"\n{'─'*70}")
print("BACKORDERS PER WEEK — END PRODUCT")
print(f"{'─'*70}")
print(f"  {'Week':>4}  {'Realized demand':>16}  {'Net position':>14}  {'Backorder':>10}")
print(f"  {'─'*56}")
bo_used = False
for t in periods:
    if bo[t] > 0:
        bo_used = True
        print(f"  {t:>4}  {DEMAND_REALIZED[t-1]:>16}  {net_end[t]:>14,.0f}  {bo[t]:>10,.0f}")
if not bo_used:
    print("  No backorders.")

# ── JSON output ─────────────────────────────────────────────────────────────
output = {
    "status": "SIMULATED",
    "based_on_plan": "output_3a.json",
    "demand_type": "realized",
    "costs": {
        "setup_cost": clean_num(total_setup),
        "holding_cost": clean_num(total_holding),
        "overtime_cost_X": clean_num(total_ot_x),
        "overtime_cost_Y": clean_num(total_ot_y),
        "total_overtime_cost": clean_num(total_overtime),
        "backorder_cost": clean_num(total_backorder),
        "total_cost": clean_num(total_cost_3b),
    },
    "comparison_with_3a": {
        "total_cost_3a": clean_num(cost_3a),
        "total_cost_3b": clean_num(total_cost_3b),
        "difference": clean_num(total_cost_3b - cost_3a),
    },
    "service_metrics": {
        "service_level_pct": clean_num(service_level * 100),
        "fill_rate_pct": clean_num(fill_rate * 100),
        "periods_no_backorder": periods_no_bo,
        "total_periods": T,
        "total_demand_units": total_demand,
        "units_on_time": clean_num(units_on_time),
    },
    "inventory": {
        i: {
            str(t): clean_num(inv[i][t])
            for t in periods
        }
        for i in PARTS
    },
    "backorders": {
        str(t): clean_num(bo[t])
        for t in periods
    },
    "end_product_net_position": {
        str(t): clean_num(net_end[t])
        for t in periods
    },
    "overtime_usage_from_3a_plan": plan_3a.get("overtime_usage", {})
}

with open("output_3bbackorder.json", "w") as f:
    json.dump(output, f, indent=2)

print("\nResults written to output_3b.json")