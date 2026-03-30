"""
APM Project 2026 - Assignment 1b
 
Evaluate the 1a production plan (optimised on forecasted demand) against
the *realized* demand.  Shortfalls are backordered at €250 / unit / period.
 
Steps
-----
1. Load the 1a production schedule from output_1a.json.
2. Re-simulate the inventory balance week-by-week using realized demand.
3. Allow inventory to go negative (= backorder position).
4. Compute total costs: setup (unchanged) + holding + backorder.
5. Compute service level and fill rate.
6. Write results to output_1b.json and print a comparison with 1a.
"""
 
import json
from input_data import (
    PARTS, END_PRODUCT, T, BOM, LEAD_TIME,
    INIT_INV, SETUP_COST, HOLDING_COST,
    DEMAND_REALIZED, BACKORDER_COST,
)
 
# ── Load the 1a production schedule ──────────────────────────────────────────
with open("output_1a.json") as f:
    plan_1a = json.load(f)
 
# production_schedule[part][str(t)] = quantity
schedule = plan_1a["production_schedule"]   # {part: {str(t): qty}}
cost_1a  = plan_1a["total_cost"]
 
periods = range(1, T + 1)
 
# ── Helper: BOM parents ───────────────────────────────────────────────────────
def get_parents(part):
    parents = {}
    for parent, children in BOM.items():
        if part in children:
            parents[parent] = children[part]
    return parents
 
# ── Simulate inventory with backorders ───────────────────────────────────────
# inv[part][t]  : net inventory at end of period t  (negative = backorder)
inv   = {i: {} for i in PARTS}
bo    = {i: {} for i in PARTS}   # backorder units per period (>= 0)
h_pos = {i: {} for i in PARTS}   # positive inventory (for holding cost)
 
for i in PARTS:
    inv_prev = INIT_INV[i]
 
    for t in periods:
        # ── Receipts: orders placed in period (t - lead_time) arrive now ──
        order_period = t - LEAD_TIME[i]
        if order_period >= 1:
            receipts = float(schedule.get(i, {}).get(str(order_period), 0))
        else:
            receipts = 0.0
 
        # ── Internal demand (same logic as 1a) ───────────────────────────
        internal_demand = 0.0
        for p, qty_per_parent in get_parents(i).items():
            # units of parent p produced in period t (from the fixed plan)
            parent_prod = float(schedule.get(p, {}).get(str(t), 0))
            internal_demand += qty_per_parent * parent_prod
 
        # ── External demand (realized, end product only) ──────────────────
        external_demand = DEMAND_REALIZED[t - 1] if i == END_PRODUCT else 0.0
 
        # ── Net inventory (can be negative = backorder) ───────────────────
        net = inv_prev + receipts - internal_demand - external_demand
 
        inv[i][t] = net
        bo[i][t]  = max(0.0, -net)          # backorder units
        h_pos[i][t] = max(0.0, net)         # units held in stock
 
        inv_prev = net  # carry forward (including negative balance)
 
# ── Cost calculation ──────────────────────────────────────────────────────────
# Setup costs: same setups as 1a (plan is unchanged)
total_setup = plan_1a["setup_cost"]
 
# Holding costs: only on positive inventory
total_holding = sum(
    HOLDING_COST[i] * h_pos[i][t]
    for i in PARTS
    for t in periods
)
 
# Backorder costs: only on end product backorders
# A backorder in period t costs €250 per unit *per period* it remains outstanding.
# The standard approach used in MRP/lot-sizing literature is to charge the
# backorder cost for each period the unit is short (i.e. bo[END_PRODUCT][t] each week).
total_backorder = sum(
    BACKORDER_COST * bo[END_PRODUCT][t]
    for t in periods
)
 
total_cost_1b = total_setup + total_holding + total_backorder
 
# ── Service level & fill rate (end product only) ──────────────────────────────
# Service level = fraction of periods with NO backorder at end of period
periods_no_bo   = sum(1 for t in periods if bo[END_PRODUCT][t] == 0)
service_level   = periods_no_bo / T
 
# Fill rate = fraction of demanded units delivered *on time* (not backordered)
# = 1 - (total units backordered that are NEW in each period) / total demand
# We track *newly created* backorders rather than the cumulative position.
# New backorders in period t = increase in backorder position vs previous period
# (after any receipts cleared old backorders).
# Equivalently: new_bo[t] = max(0, demand_t - available_before_demand_t)
#
# For simplicity (and consistency with the simulation above which allows
# carried-over backorders), we use:
#   new_bo[t] = max(0, bo[t] - bo[t-1])  where bo[0] = 0
# This correctly counts units that *first* became short in period t.
 
total_demand     = sum(DEMAND_REALIZED)
total_new_bo     = 0.0
prev_bo          = 0.0
for t in periods:
    new_bo        = max(0.0, bo[END_PRODUCT][t] - prev_bo)
    total_new_bo += new_bo
    prev_bo       = bo[END_PRODUCT][t]
 
fill_rate = 1.0 - (total_new_bo / total_demand) if total_demand > 0 else 1.0
 
# ── Print results ─────────────────────────────────────────────────────────────
print(f"\n{'='*60}")
print("ASSIGNMENT 1b — REALIZED DEMAND EVALUATION")
print(f"{'='*60}")
print(f"  Setup cost:      €{total_setup:>12,.2f}  (unchanged from 1a)")
print(f"  Holding cost:    €{total_holding:>12,.2f}")
print(f"  Backorder cost:  €{total_backorder:>12,.2f}")
print(f"  {'─'*38}")
print(f"  Total cost (1b): €{total_cost_1b:>12,.2f}")
print(f"  Total cost (1a): €{cost_1a:>12,.2f}")
print(f"  Difference:      €{total_cost_1b - cost_1a:>+12,.2f}")
 
print(f"\n{'─'*60}")
print("SERVICE METRICS (end product):")
print(f"  Service level : {service_level*100:.1f}%  "
      f"({periods_no_bo}/{T} periods without backorder)")
print(f"  Fill rate     : {fill_rate*100:.2f}%  "
      f"({total_demand - total_new_bo:,.0f} / {total_demand:,.0f} units on time)")
 
print(f"\n{'─'*60}")
print("BACKORDER DETAIL — END PRODUCT (E2801):")
print(f"  {'Week':>4}  {'Realized demand':>16}  {'Inventory':>12}  {'Backorder':>10}")
print(f"  {'─'*48}")
for t in periods:
    if bo[END_PRODUCT][t] > 0 or inv[END_PRODUCT][t] < 0:
        marker = " ◄"
    else:
        marker = ""
    print(f"  {t:>4}  {DEMAND_REALIZED[t-1]:>16}  "
          f"{inv[END_PRODUCT][t]:>12,.0f}  "
          f"{bo[END_PRODUCT][t]:>10,.0f}{marker}")
 
# ── Write output ──────────────────────────────────────────────────────────────
output = {
    "status": "SIMULATED",
    "based_on_plan": "output_1a.json",
    "demand_type": "realized",
    "costs": {
        "setup_cost":     total_setup,
        "holding_cost":   total_holding,
        "backorder_cost": total_backorder,
        "total_cost":     total_cost_1b,
    },
    "comparison_with_1a": {
        "total_cost_1a":  cost_1a,
        "total_cost_1b":  total_cost_1b,
        "difference":     total_cost_1b - cost_1a,
    },
    "service_metrics": {
        "service_level_pct": round(service_level * 100, 2),
        "fill_rate_pct":     round(fill_rate * 100, 2),
        "periods_no_backorder": periods_no_bo,
        "total_periods":     T,
        "total_demand_units": total_demand,
        "units_on_time":     total_demand - total_new_bo,
    },
    "inventory_and_backorders": {
        i: {
            str(t): {
                "net_inventory": inv[i][t],
                "backorder":     bo[i][t],
                "held":          h_pos[i][t],
            }
            for t in periods
        }
        for i in PARTS
    },
}
 
with open("output_1b_test.json", "w") as f:
    json.dump(output, f, indent=2)
 
print("\nResults written to output_1b_test.json")