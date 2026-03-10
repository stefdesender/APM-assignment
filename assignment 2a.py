import gurobipy as gp
from gurobipy import GRB
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════════
# 1.  READ INPUT DATA FROM EXCEL
# ═══════════════════════════════════════════════════════════════════════════

INPUT_FILE  = "input_data.xlsx"
OUTPUT_FILE = "output_solution.xlsx"

def load_inputs(path):
    wb = load_workbook(path, data_only=True)

    # ── Part parameters (sheet "Parameters", rows 3-9, cols A-H) ──────────
    ws = wb["Parameters"]
    parts_order = []
    params = {}
    for row in ws.iter_rows(min_row=3, max_row=9, min_col=1, max_col=8, values_only=True):
        part, lt, min_lot, inv0, setup, hold, ws_id, proc = row
        parts_order.append(part)
        params[part] = {
            "lt":      int(lt),
            "min_lot": int(min_lot),
            "inv0":    int(inv0),
            "setup":   float(setup),
            "hold":    float(hold),
            "ws":      str(ws_id).strip() if ws_id else None,
            "proc":    float(proc) if proc not in ("", None) else 0.0,
        }

    # ── BOM (rows 13-18, cols A-C) ────────────────────────────────────────
    bom = {}
    for row in ws.iter_rows(min_row=13, max_row=18, min_col=1, max_col=3, values_only=True):
        parent, component, qty = row
        bom.setdefault(parent, {})[component] = int(qty)

    # ── Capacity (rows 22-23, cols A-B) ───────────────────────────────────
    cap = {}
    for row in ws.iter_rows(min_row=22, max_row=23, min_col=1, max_col=2, values_only=True):
        ws_name, value = row
        cap[str(ws_name).strip()] = float(value)

    # ── Demand (sheet "Demand", rows 3-32, cols A-C) ──────────────────────
    ws2 = wb["Demand"]
    demand_forecast = {}
    demand_realized = {}
    for row in ws2.iter_rows(min_row=3, max_row=32, min_col=1, max_col=3, values_only=True):
        week, fc, re = row
        demand_forecast[int(week)] = int(fc)
        demand_realized[int(week)] = int(re)

    # ── Backorder cost (sheet "Backorder", row 3, col B) ──────────────────
    ws3 = wb["Backorder"]
    backorder_cost = float(ws3.cell(3, 2).value)

    wb.close()
    return parts_order, params, bom, cap, demand_forecast, demand_realized, backorder_cost


# ═══════════════════════════════════════════════════════════════════════════
# 2.  MIP MODEL
# ═══════════════════════════════════════════════════════════════════════════

def build_and_solve(parts, params, bom, cap, demand_dict, T,
                    label="forecast", backorder_cost=None):
    periods = list(range(1, T + 1))
    BIG_M   = 1_000_000

    model = gp.Model(f"APM_Assign2_{label}")
    model.Params.LogToConsole = 1

    # ── Variables ──────────────────────────────────────────────────────────
    x, y, I, B = {}, {}, {}, {}
    for p in parts:
        for t in periods:
            x[p, t] = model.addVar(lb=0, vtype=GRB.INTEGER,    name=f"x_{p}_{t}")
            y[p, t] = model.addVar(vtype=GRB.BINARY,           name=f"y_{p}_{t}")
            I[p, t] = model.addVar(lb=0, vtype=GRB.CONTINUOUS, name=f"I_{p}_{t}")
            if backorder_cost is not None:
                B[p, t] = model.addVar(lb=0, vtype=GRB.CONTINUOUS, name=f"B_{p}_{t}")

    # ── Objective ──────────────────────────────────────────────────────────
    obj = (gp.quicksum(params[p]["setup"] * y[p, t] for p in parts for t in periods)
         + gp.quicksum(params[p]["hold"]  * I[p, t] for p in parts for t in periods))
    if backorder_cost is not None:
        obj += gp.quicksum(backorder_cost * B["E2801", t] for t in periods)
    model.setObjective(obj, GRB.MINIMIZE)

    # ── Inventory balance + lot-size constraints ───────────────────────────
    for p in parts:
        lt = params[p]["lt"]
        for t in periods:
            inv_prev  = params[p]["inv0"] if t == 1 else I[p, t - 1]
            prod_recv = x[p, t - lt] if (t - lt) >= 1 else 0

            if p == "E2801":
                use = demand_dict[t]
                if backorder_cost is not None:
                    back_prev = B[p, t - 1] if t > 1 else 0
                    model.addConstr(
                        inv_prev + prod_recv + back_prev - use == I[p, t] - B[p, t],
                        name=f"inv_bal_{p}_{t}")
                else:
                    model.addConstr(inv_prev + prod_recv - use == I[p, t],
                                    name=f"inv_bal_{p}_{t}")
            else:
                use_expr = gp.quicksum(
                    children[p] * x[parent, t]
                    for parent, children in bom.items() if p in children)
                model.addConstr(inv_prev + prod_recv - use_expr == I[p, t],
                                name=f"inv_bal_{p}_{t}")

            model.addConstr(x[p, t] >= params[p]["min_lot"] * y[p, t],
                            name=f"min_lot_{p}_{t}")
            model.addConstr(x[p, t] <= BIG_M * y[p, t],
                            name=f"setup_link_{p}_{t}")

    # ── Workstation capacity constraints ───────────────────────────────────
    cap_X = cap.get("X")
    cap_Y = cap.get("Y_weekly_minutes")
    ws_X  = [p for p in parts if params[p]["ws"] == "X"]
    ws_Y  = [p for p in parts if params[p]["ws"] == "Y"]

    for t in periods:
        if cap_X and ws_X:
            model.addConstr(gp.quicksum(x[p, t] for p in ws_X) <= cap_X,
                            name=f"cap_X_{t}")
        if cap_Y and ws_Y:
            model.addConstr(
                gp.quicksum(params[p]["proc"] * x[p, t] for p in ws_Y) <= cap_Y,
                name=f"cap_Y_{t}")

    model.optimize()

    if model.Status not in (GRB.OPTIMAL, GRB.SUBOPTIMAL):
        print(f"[{label}] No optimal solution. Status = {model.Status}")
        return model, None

    # ── Extract solution ───────────────────────────────────────────────────
    sol = {
        "status": model.Status, "obj": model.ObjVal,
        "periods": periods, "parts": parts,
        "x": {}, "y": {}, "I": {}, "B": {},
        "cap_X": cap_X, "cap_Y": cap_Y, "demand": demand_dict,
        "setup_cost": 0.0, "holding_cost": 0.0, "backorder_cost_total": 0.0,
    }
    for p in parts:
        sol["x"][p] = {}; sol["y"][p] = {}; sol["I"][p] = {}; sol["B"][p] = {}
        for t in periods:
            sol["x"][p][t] = round(x[p, t].X)
            sol["y"][p][t] = int(round(y[p, t].X))
            sol["I"][p][t] = round(I[p, t].X, 2)
            bval = round(B[p, t].X, 2) if (backorder_cost and p == "E2801") else 0.0
            sol["B"][p][t] = bval
            sol["setup_cost"]   += params[p]["setup"] * sol["y"][p][t]
            sol["holding_cost"] += params[p]["hold"]  * sol["I"][p][t]

    if backorder_cost:
        sol["backorder_cost_total"] = sum(
            backorder_cost * sol["B"]["E2801"][t] for t in periods)
        no_back = sum(1 for t in periods if sol["B"]["E2801"][t] < 0.5)
        sol["service_level"] = no_back / len(periods) * 100
        tot_dem = sum(demand_dict[t] for t in periods)
        tot_bk  = sum(sol["B"]["E2801"][t] for t in periods)
        sol["fill_rate"] = (1 - tot_bk / tot_dem) * 100

    sol["util_X"] = {t: sum(sol["x"][p][t] for p in ws_X) for t in periods}
    sol["util_Y"] = {t: sum(params[p]["proc"] * sol["x"][p][t]
                            for p in ws_Y) for t in periods}
    return model, sol


# ═══════════════════════════════════════════════════════════════════════════
# 3.  WRITE OUTPUT TO EXCEL
# ═══════════════════════════════════════════════════════════════════════════

HBLU = "2F5496"; HBLU2 = "4472C4"
GREY = "F2F2F2"; WHITE = "FFFFFF"
BLUE_F = "0070C0"; GREEN_F = "375623"; RED_F = "C00000"

def _hdr(cell, val, bg=HBLU, fc="FFFFFF", bold=True):
    cell.value = val
    cell.font  = Font(bold=bold, color=fc, name="Arial", size=10)
    cell.fill  = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def _val(cell, val, color="000000", bold=False, fmt=None):
    cell.value = val
    cell.font  = Font(color=color, name="Arial", size=10, bold=bold)
    cell.alignment = Alignment(horizontal="center")
    if fmt:
        cell.number_format = fmt

def _thin(ws, r1, r2, c1, c2):
    th = Side(style="thin")
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for c in row:
            c.border = Border(left=th, right=th, top=th, bottom=th)

def write_output(sol_a, sol_b, params, path):
    wb = Workbook()

    # ── Summary ────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Summary"
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 30
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22

    _hdr(ws["A1"], "Assignment 2 – Cost & Performance Summary", bg=HBLU)
    ws.merge_cells("A1:C1")
    for col, h in enumerate(["Cost Type", "2a – Forecast", "2b – Realized (backorder)"], 1):
        _hdr(ws.cell(2, col), h, bg=HBLU2)

    summary_rows = [
        ("Setup cost (€)",      sol_a["setup_cost"],          sol_b["setup_cost"]),
        ("Holding cost (€)",    sol_a["holding_cost"],         sol_b["holding_cost"]),
        ("Backorder cost (€)",  "–",                           sol_b["backorder_cost_total"]),
        ("Total objective (€)", sol_a["obj"],                  sol_b["obj"]),
    ]
    for r, (lbl, va, vb) in enumerate(summary_rows, 3):
        fill = PatternFill("solid", start_color=GREY if r % 2 == 0 else WHITE)
        is_tot = lbl.startswith("Total")
        ws.cell(r, 1).value = lbl
        ws.cell(r, 1).font  = Font(name="Arial", size=10, bold=is_tot)
        ws.cell(r, 1).fill  = fill
        for col, v in enumerate([va, vb], 2):
            c = ws.cell(r, col); c.fill = fill
            if isinstance(v, float):
                _val(c, v, fmt='#,##0.00 "€"', bold=is_tot,
                     color="000000" if is_tot else BLUE_F)
            else:
                _val(c, v)

    if "service_level" in sol_b:
        for r, (lbl, v) in enumerate([("Service Level (%)", sol_b["service_level"]),
                                       ("Fill Rate (%)",     sol_b["fill_rate"])], 7):
            ws.cell(r, 1).value = lbl
            ws.cell(r, 1).font  = Font(name="Arial", size=10)
            _val(ws.cell(r, 2), "–")
            _val(ws.cell(r, 3), round(v, 1), color=GREEN_F, fmt='0.0"%"')
    _thin(ws, 2, 8, 1, 3)

    # ── Production plan sheet helper ───────────────────────────────────────
    def plan_sheet(wb, sol, title):
        ws = wb.create_sheet(title)
        parts   = sol["parts"]
        periods = sol["periods"]
        has_bk  = any(sol["B"].get("E2801", {}).get(t, 0) > 0 for t in periods)

        col_hdrs = ["Period", "Demand (E2801)"]
        for p in parts:
            col_hdrs += [f"{p} – Prod/Order", f"{p} – Setup (0/1)", f"{p} – Inventory"]
            if p == "E2801" and has_bk:
                col_hdrs.append(f"{p} – Backlog")

        n_cols = len(col_hdrs)
        ws.column_dimensions["A"].width = 9
        ws.column_dimensions["B"].width = 18
        for i in range(2, n_cols):
            ws.column_dimensions[get_column_letter(i + 1)].width = 17

        _hdr(ws["A1"], title, bg=HBLU)
        ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
        for col, h in enumerate(col_hdrs, 1):
            _hdr(ws.cell(2, col), h, bg=HBLU2)

        for r, t in enumerate(periods, 3):
            fill = PatternFill("solid", start_color=GREY if r % 2 == 0 else WHITE)
            col = 1
            _val(ws.cell(r, col), t);              ws.cell(r, col).fill = fill; col += 1
            _val(ws.cell(r, col), sol["demand"][t]); ws.cell(r, col).fill = fill; col += 1
            for p in parts:
                _val(ws.cell(r, col), sol["x"][p][t], color=BLUE_F)
                ws.cell(r, col).fill = fill; col += 1
                _val(ws.cell(r, col), sol["y"][p][t])
                ws.cell(r, col).fill = fill; col += 1
                _val(ws.cell(r, col), sol["I"][p][t], color=BLUE_F)
                ws.cell(r, col).fill = fill; col += 1
                if p == "E2801" and has_bk:
                    bv = sol["B"]["E2801"].get(t, 0)
                    _val(ws.cell(r, col), bv, color=RED_F if bv > 0 else "000000")
                    ws.cell(r, col).fill = fill; col += 1

        _thin(ws, 2, 2 + len(periods), 1, n_cols)

    plan_sheet(wb, sol_a, "2a – Forecast Plan")
    plan_sheet(wb, sol_b, "2b – Realized Plan")

    # ── Capacity utilisation sheet ─────────────────────────────────────────
    wsc = wb.create_sheet("Capacity Utilisation")
    for col, w in zip("ABCDE", [9, 20, 22, 20, 22]):
        wsc.column_dimensions[get_column_letter(
            "ABCDE".index(col) + 1)].width = w

    _hdr(wsc["A1"], "Workstation Capacity Utilisation per Period", bg=HBLU)
    wsc.merge_cells("A1:E1")
    cap_X = sol_a["cap_X"]; cap_Y = int(sol_a["cap_Y"])
    for col, h in enumerate(["Period",
                               f"2a – WS X (cap {int(cap_X)})",
                               f"2a – WS Y (cap {cap_Y} min)",
                               f"2b – WS X (cap {int(cap_X)})",
                               f"2b – WS Y (cap {cap_Y} min)"], 1):
        _hdr(wsc.cell(2, col), h, bg=HBLU2)

    for r, t in enumerate(sol_a["periods"], 3):
        fill = PatternFill("solid", start_color=GREY if r % 2 == 0 else WHITE)
        vals = [t, sol_a["util_X"][t], sol_a["util_Y"][t],
                   sol_b["util_X"][t], sol_b["util_Y"][t]]
        for col, v in enumerate(vals, 1):
            c = wsc.cell(r, col); c.fill = fill
            over = ((col in (2, 4) and v > cap_X) or
                    (col in (3, 5) and v > cap_Y))
            _val(c, v, color=RED_F if over else "000000")

    ref = 3 + len(sol_a["periods"])
    wsc.cell(ref, 1).value = "Capacity →"
    wsc.cell(ref, 1).font  = Font(bold=True, name="Arial", size=10)
    for col in (2, 4): _val(wsc.cell(ref, col), int(cap_X), bold=True)
    for col in (3, 5): _val(wsc.cell(ref, col), cap_Y,      bold=True)
    _thin(wsc, 2, ref, 1, 5)

    wb.save(path)
    print(f"\nOutput written to: {path}")


# ═══════════════════════════════════════════════════════════════════════════
# 4.  MAIN
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("Reading input data from:", INPUT_FILE)
    parts, params, bom, cap, demand_fc, demand_re, back_cost = load_inputs(INPUT_FILE)
    T = 30

    print(f"Capacity X = {cap.get('X')} units/week")
    print(f"Capacity Y = {cap.get('Y_weekly_minutes')} min/week")

    print("\n" + "="*60)
    print("ASSIGNMENT 2a – Finite Capacity, Forecasted Demand")
    print("="*60)
    _, sol_a = build_and_solve(parts, params, bom, cap, demand_fc, T,
                               label="forecast")

    print("\n" + "="*60)
    print("ASSIGNMENT 2b – Finite Capacity, Realized Demand (backordering)")
    print("="*60)
    _, sol_b = build_and_solve(parts, params, bom, cap, demand_re, T,
                               label="realized", backorder_cost=back_cost)

    if sol_a and sol_b:
        print("\n── Cost Summary ──────────────────────────────────")
        print(f"  2a  Setup cost   : €{sol_a['setup_cost']:>12,.2f}")
        print(f"  2a  Holding cost : €{sol_a['holding_cost']:>12,.2f}")
        print(f"  2a  Total        : €{sol_a['obj']:>12,.2f}")
        print(f"  2b  Setup cost   : €{sol_b['setup_cost']:>12,.2f}")
        print(f"  2b  Holding cost : €{sol_b['holding_cost']:>12,.2f}")
        print(f"  2b  Backord cost : €{sol_b['backorder_cost_total']:>12,.2f}")
        print(f"  2b  Total        : €{sol_b['obj']:>12,.2f}")
        if "service_level" in sol_b:
            print(f"  Service level   :  {sol_b['service_level']:.1f}%")
            print(f"  Fill rate       :  {sol_b['fill_rate']:.1f}%")

        write_output(sol_a, sol_b, params, OUTPUT_FILE)
