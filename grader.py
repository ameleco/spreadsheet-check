import json,io,contextlib,math
from pathlib import Path
from contextvars import ContextVar

import re, math, openpyxl
EXPECTED_N = 10

def norm(s):
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s is not None else ""

def try_float(x):
    try:
        return float(x)
    except:
        return None

def is_formula(v):
    return isinstance(v, str) and v.strip().startswith("=")

def close(a, b, rel=0.02, abs_tol=0.05):
    if a is None or b is None:
        return False
    if abs(a - b) <= abs_tol:
        return True
    if b != 0 and abs(a - b) / abs(b) <= rel:
        return True
    return False

def parse_plusminus(cell_val):
    if cell_val is None:
        return None, None

    s = str(cell_val)
    s = s.replace("−", "-").replace("±", "+/-").replace("∕", "/")
    s = re.sub(r"[\u00A0\u2007\u202F]", " ", s).strip()

    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1].strip()

    s = s.replace("+ / -", "+/-").replace("+/ -", "+/-").replace("+ /-", "+/-")

    m = re.search(r"([-+]?\d*\.?\d+)\s*\+/-\s*([-+]?\d*\.?\d+)", s)
    if not m:
        return None, None

    return try_float(m.group(1)), try_float(m.group(2))

LABS = {

    # ---------- LAB 1 ----------
    "lab1_measurements": {
        "type": "measurement_uncertainty",
        "expected_n": EXPECTED_N,
        "columns": [
            {"name": "Car Speeds", "units": "m/s", "keywords": ["speed", "velocity", "car"]},
            {"name": "Octopus Mass", "units": "kg",  "keywords": ["mass", "octopus", "weight"]},
        ],
        "require_formulas_in_calc_rows": True,
        "good_items": [
            "Measurements were found for both columns.",
            "Headers and units were included.",
            "Mean, range, and uncertainty calculations were present.",
            "Calculation labels were included.",
            "Final values matched the measurements.",
        ],
    },

    # ---------- LAB 2 ----------
    "lab2_motion_graphs": {
        "type": "motion_graphs",
        "expected_n": 50,

        "tables": [
            {
                "name": "Constantly Moving Car",
                "keywords": ["constant", "constantly", "moving", "car"],
                "columns": [
                    {"name": "Time", "units": "s", "keywords": ["time", "t"]},
                    {"name": "Position", "units": "m", "keywords": ["position", "pos"]},
                    {"name": "Velocity", "units": "m/s", "keywords": ["velocity", "vel"]},
                    {"name": "Acceleration", "units": ["m/s^2", "m/s²"], "keywords": ["acceleration", "accel", "acc"]},
                ],
            },
            {
                "name": "Accelerating Car",
                "keywords": ["accelerating", "acceleration", "car"],
                "columns": [
                    {"name": "Time", "units": "s", "keywords": ["time", "t"]},
                    {"name": "Position", "units": "m", "keywords": ["position", "pos"]},
                    {"name": "Velocity", "units": "m/s", "keywords": ["velocity", "vel"]},
                    {"name": "Acceleration", "units": ["m/s^2", "m/s²"], "keywords": ["acceleration", "accel", "acc"]},
                ],
            },
        ],

        "graphs": {
            "expected_count": 2,
            "expected_series_per_graph": 3,
        },

        "points": {
            "constant_table": 2.0,
            "accelerating_table": 2.0,
            "graphs_exist_series_linked": 2.0,
            "titles_legend": 2.0,
            "axes_units": 2.0,
        },

        "good_items": [
            "Both motion data tables were checked.",
            "The required time, position, velocity, and acceleration columns were checked.",
            "Graphs were checked for required series and links to worksheet data.",
            "Graph titles and legends were checked.",
            "Axis titles and units were checked.",
        ],
    },

    # ---------- LAB 3 ----------
    "lab3_projectile_uncertainty": {
        "type": "projectile_uncertainty",
        "expected_angle_rows": 5,
        "expected_random_trials": 8,

        "good_items": [
            "Your projectile distance table was checked.",
            "Your projectile graph was checked.",
            "Your random error calculations were checked.",
            "Your systematic error comparison values were checked.",
        ],
    },

    # ---------- LAB 4/5 ----------
    "lab4d_atwood_friction": {
        "type": "atwood_friction",

        "tables": [
            {
                "name": "No Friction Table",
                "keywords": ["no friction"],
                "columns": [
                    {"name": "Mass 1", "units": "kg", "keywords": ["mass 1"]},
                    {"name": "Mass 2", "units": "kg", "keywords": ["mass 2"]},
                    {"name": "Acceleration", "units": ["m/s^2", "m/s²"], "keywords": ["acceleration", "accel"]},
                ],
            },
            {
                "name": "Friction Table",
                "keywords": ["friction"],
                "columns": [
                    {"name": "Mass 1", "units": "kg", "keywords": ["mass 1"]},
                    {"name": "Mass 2", "units": "kg", "keywords": ["mass 2"]},
                    {"name": "Acceleration", "units": ["m/s^2", "m/s²"], "keywords": ["acceleration", "accel"]},
                    {"name": "Mu", "units": "", "keywords": ["mu", "friction"]},
                ],
            },
        ],

        "graphs": {
            "expected_count": 2,
        },

        "points": {
            "tables": 4.0,
            "graphs": 6.0,
        },

        "good_items": [
            "Both Atwood tables were checked.",
            "Table titles and column headers with units were checked.",
            "Graphs were checked for titles, axes, and data.",
        ],
    },

    # ---------- LAB 6/7 ----------
    "lab6_7_collisions": {
        "type": "collisions",

        "summary_tables_expected": 3,
        "time_tables_expected": 3,
        "graphs_expected": 3,

        "expected_time_rows": 50,

        "good_items": [
            "The three conservation summary tables were checked.",
            "The three time-data tables were checked.",
            "The conservation, kinetic energy, and momentum calculations were checked for formulas.",
            "The three graphs were checked for titles, legends, axes, units, and plotted data.",
        ],
    },

    # ---------- ROTATION LAB ----------
    "rotation_lab": {
        "type": "rotation",

        "expected_summary_tables": 2,
        "expected_time_rows": 50,
        "expected_graphs": 1,

        "good_items": [
            "The changing-radius and changing-mass tables were checked.",
            "The experimental and theoretical moment of inertia columns were checked for formulas.",
            "The orbiting-masses time table was checked.",
            "The graph was checked for title, legend, axes, units, and plotted data.",
        ],
    },

    # ---------- SPRING LAB ----------
    "spring_lab": {
        "type": "spring",

        "expected_tables": 3,
        "expected_rows": 20,
        "expected_graphs": 3,

        "good_items": [
            "The three spring data tables were checked.",
            "The spring table titles, column headers, units, and data rows were checked.",
            "The three graphs were checked for titles, axes, units, and trendlines.",
            "Trendline equations were checked when possible.",
        ],
    },
}

class SpreadsheetGrader:
    """
    Rubric per column (5.0 pts):
      1.0  Data present (10 numbers)
      1.0  Header label + units
      1.5  Calc cells present (Mean/Range/Unc single/Unc mean)
           - Full credit requires formulas if require_formulas=True
      0.5  Calc labels present near calc rows
      1.0  Final Value format + correctness (typed OK)
    """

    def __init__(self, expected_n=10, scan_max_rows=400, scan_max_cols=60, header_scan_rows=80,
                 require_formulas=True):
        self.expected_n = expected_n
        self.scan_max_rows = scan_max_rows
        self.scan_max_cols = scan_max_cols
        self.header_scan_rows = header_scan_rows
        self.require_formulas = require_formulas

    def _header_score(self, text, units, keywords):
        t = norm(text)
        pts = 0.0
        if units and units in t:
            pts += 0.6
        if any(k in t for k in keywords):
            pts += 0.4
        return pts

    def _find_column_by_header(self, ws, units, keywords, min_score=0.7):
        best = (None, None, 0.0)  # (col_idx, header_row, score)
        max_r = min(self.header_scan_rows, ws.max_row)
        max_c = min(self.scan_max_cols, ws.max_column)

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                sc = self._header_score(ws.cell(r, c).value, units, keywords)
                if sc > best[2]:
                    best = (c, r, sc)

        return best if best[2] >= min_score else (None, None, best[2])

    def _find_best_10_run(self, ws_v, banned_cols=set()):
        max_r = min(self.scan_max_rows, ws_v.max_row)
        max_c = min(self.scan_max_cols, ws_v.max_column)

        best = None  # (score, col_idx, start_row)
        for c in range(1, max_c + 1):
            if c in banned_cols:
                continue
            for start in range(1, max_r - self.expected_n + 2):
                ok = True
                for r in range(start, start + self.expected_n):
                    if try_float(ws_v.cell(r, c).value) is None:
                        ok = False
                        break
                if not ok:
                    continue

                nearby = 0
                for rr in range(max(1, start - 4), min(max_r, start + self.expected_n + 5)):
                    if try_float(ws_v.cell(rr, c).value) is not None:
                        nearby += 1

                score = (1000 - start) + 2 * nearby
                if best is None or score > best[0]:
                    best = (score, c, start)

        if not best:
            return None, None
        return best[1], best[2]

    def _find_label_row(self, ws_f, label_keywords, around_row, search_rows=40):
        r0 = max(1, around_row - search_rows)
        r1 = min(ws_f.max_row, around_row + search_rows)
        max_c = min(self.scan_max_cols, ws_f.max_column)

        for r in range(r0, r1 + 1):
            for c in range(1, max_c + 1):
                txt = norm(ws_f.cell(r, c).value)
                if txt and any(k in txt for k in label_keywords):
                    return r, c
        return None, None

    def _summary_rows_from_labels(self, ws_f, data_end, value_col_idx):
        anchor = data_end + 4

        final_r, _ = self._find_label_row(ws_f, ["final value", "final"], anchor, search_rows=35)
        if final_r is None:
            mean_r, _ = self._find_label_row(ws_f, ["mean", "average"], anchor, search_rows=35)
            if mean_r is None:
                return None
            final_r = mean_r + 4

        rows = {
            "final": final_r,
            "unc_mean": final_r - 1,
            "unc_single": final_r - 2,
            "range": final_r - 3,
            "mean": final_r - 4,
        }

        checks = [
            ("mean", ["mean", "average"]),
            ("range", ["range"]),
            ("unc_single", ["unc", "uncert", "single", "sing"]),
            ("unc_mean", ["unc", "uncert", "mean"]),
            ("final", ["final"]),
        ]

        # ✅ tweak: allow labels to be in the same column too, not only neighbors
        neighbor_cols = [value_col_idx, value_col_idx - 2, value_col_idx - 1, value_col_idx + 1, value_col_idx + 2]

        hit = 0
        for key, keys in checks:
            rr = rows[key]
            for cc in neighbor_cols:
                if 1 <= cc <= ws_f.max_column:
                    txt = norm(ws_f.cell(rr, cc).value)
                    if txt and any(k in txt for k in keys):
                        hit += 1
                        break

        return rows if hit >= 2 else None

    def _find_summary_start_fallback(self, ws_f, ws_v, col_idx, data_end):
        max_r = min(self.scan_max_rows, ws_v.max_row)

        r0 = max(1, data_end - 2)
        r1 = min(max_r - 4, data_end + 30)

        for r_mean in range(r0, r1 + 1):
            present = 0
            for i in range(4):
                v_f = ws_f.cell(r_mean + i, col_idx).value
                v_v = ws_v.cell(r_mean + i, col_idx).value
                if is_formula(v_f) or try_float(v_v) is not None:
                    present += 1
            if present < 2:
                continue

            final_raw = ws_f.cell(r_mean + 4, col_idx).value
            v, u = parse_plusminus(final_raw)
            if (v is not None and u is not None) or (final_raw is not None and str(final_raw).strip()):
                return r_mean

        return None

    def _grade_one_column(self, ws_f, ws_v, col_idx, data_start, units, keywords, display_name, debug=False):
        n = self.expected_n
        pts = 0.0
        notes = []

        # 1) data (1.0)
        vals = []
        for r in range(data_start, data_start + n):
            v = try_float(ws_v.cell(r, col_idx).value)
            if v is not None:
                vals.append(v)

        if len(vals) == n:
            pts += 1.0
        else:
            pts += len(vals) / n
            notes.append(f"{display_name}: I found {len(vals)}/{n} measurements (all 10 entries should be numbers).")

        if not vals:
            return pts, notes

        mean_exp = sum(vals) / len(vals)
        range_exp = max(vals) - min(vals)
        unc_single = range_exp / 2.0
        unc_mean = unc_single / math.sqrt(len(vals))

        data_end = data_start + n - 1

        # 2) header (1.0)
        header_row = data_start - 1
        h = ws_f.cell(header_row, col_idx).value
        h_pts = self._header_score(h, units, keywords)
        pts += h_pts
        if h_pts < 1.0:
            notes.append(f"{display_name}: add a header with the correct label and units (include “{units}”).")

        # summary rows
        rows = self._summary_rows_from_labels(ws_f, data_end, col_idx)
        if rows:
            r_mean = rows["mean"]
            r_range = rows["range"]
            r_unc_single = rows["unc_single"]
            r_unc_mean = rows["unc_mean"]
            r_final = rows["final"]
        else:
            r_mean = self._find_summary_start_fallback(ws_f, ws_v, col_idx, data_end) or (data_end + 1)
            r_range = r_mean + 1
            r_unc_single = r_mean + 2
            r_unc_mean = r_mean + 3
            r_final = r_mean + 4

        # 3) calc cells (1.5) + formula requirement
        calc_rows = [r_mean, r_range, r_unc_single, r_unc_mean]
        present = 0
        formulas = 0
        for rr in calc_rows:
            v_f = ws_f.cell(rr, col_idx).value
            v_v = ws_v.cell(rr, col_idx).value
            if is_formula(v_f):
                present += 1
                formulas += 1
            elif try_float(v_v) is not None:
                present += 1

        if present == 4:
            calc_pts = 1.5
        elif present >= 2:
            calc_pts = 1.0
        elif present == 1:
            calc_pts = 0.5
        else:
            calc_pts = 0.0
            notes.append(f"{display_name}: I couldn’t find your Mean/Range/Uncertainty calculation cells.")

        if self.require_formulas:
            if present >= 2 and formulas == 0:
                calc_pts = min(calc_pts, 0.5)
                notes.append(f"{display_name}: use Excel formulas for Mean/Range/Uncertainty (not hand-typed results).")
            elif present >= 2 and 0 < formulas < 4:
                frac = formulas / 4.0
                calc_pts = min(calc_pts, 0.5 + frac * 1.0)
                notes.append(f"{display_name}: formulas found in {formulas}/4 calc cells. Use formulas in all four rows.")

        pts += calc_pts

        # 3b) calc labels (0.5)
        label_targets = [
            (r_mean, ["mean", "average"]),
            (r_range, ["range"]),
            (r_unc_single, ["unc", "uncert", "single", "sing"]),
            (r_unc_mean, ["unc", "uncert", "mean"]),
        ]

        def count_labels_in_col(label_col_idx):
            hits = 0
            for rr, keys in label_targets:
                txt = norm(ws_f.cell(rr, label_col_idx).value)
                if any(k in txt for k in keys):
                    hits += 1
            return hits

        label_hits = 0
        for neighbor in (col_idx, col_idx - 2, col_idx - 1, col_idx + 1, col_idx + 2):
            if 1 <= neighbor <= ws_f.max_column:
                label_hits = max(label_hits, count_labels_in_col(neighbor))

        if label_hits >= 3:
            pts += 0.5
        elif label_hits >= 1:
            pts += 0.25
            notes.append(f"{display_name}: I only found {label_hits}/4 calculation labels next to your calc rows.")
        else:
            notes.append(f"{display_name}: add Mean/Range/Unc labels next to the calculation rows.")

        # 4) final value (1.0)
        final_raw = ws_f.cell(r_final, col_idx).value
        v_final, u_final = parse_plusminus(final_raw)

        if v_final is None or u_final is None:
            alt_raw = ws_f.cell(r_final + 1, col_idx).value
            v2, u2 = parse_plusminus(alt_raw)
            if v2 is not None and u2 is not None:
                v_final, u_final = v2, u2
                r_final += 1
                final_raw = alt_raw

        if v_final is None or u_final is None:
            if debug:
                print(f"DEBUG {display_name} final_raw repr:", repr(final_raw))
            notes.append(f"{display_name}: Final Value must look like value +/- uncertainty (e.g., 31.1+/-0.5).")
        else:
            ok_v = close(v_final, mean_exp, rel=0.02, abs_tol=0.15)
            ok_u = close(u_final, unc_mean, rel=0.15, abs_tol=0.15)
            if ok_v and ok_u:
                pts += 1.0
            else:
                pts += 0.5
                notes.append(f"{display_name}: Final Value format is fine, but expected about {mean_exp:.4g} +/- {unc_mean:.4g}.")

        return min(5.0, pts), notes

    def grade_file_two_columns(self, path, colA, colB, debug=False):
        wb_f = _grading_workbook(path, data_only=False)
        wb_v = _grading_workbook(path, data_only=True)
        ws_f = wb_f[wb_f.sheetnames[0]]
        ws_v = wb_v[wb_v.sheetnames[0]]

        c1, h1, _ = self._find_column_by_header(ws_f, colA["units"], colA["keywords"])
        c2, h2, _ = self._find_column_by_header(ws_f, colB["units"], colB["keywords"])

        if c1 is None:
            c1, s1 = self._find_best_10_run(ws_v)
        else:
            s1 = h1 + 1

        if c2 is None:
            c2, s2 = self._find_best_10_run(ws_v, banned_cols={c1} if c1 else set())
        else:
            s2 = h2 + 1

        if c1 is None or c2 is None:
            return 0.0, ["I couldn’t find two columns with 10 measurements each."]

        p1, fb1 = self._grade_one_column(ws_f, ws_v, c1, s1, colA["units"], colA["keywords"], colA["name"], debug=debug)
        p2, fb2 = self._grade_one_column(ws_f, ws_v, c2, s2, colB["units"], colB["keywords"], colB["name"], debug=debug)

        total = round(p1 + p2, 2)

        if total >= 9.999:
            return total, ["Everything needed was found, and your final values match your data. ✅"]

        feedback = [f"{colA['name']}: {p1:.2f}/5", *fb1, f"{colB['name']}: {p2:.2f}/5", *fb2]
        return total, feedback

def _units_list(units):
    return units if isinstance(units, list) else [units]

def _text_from_chart_obj(obj):
    if obj is None:
        return ""
    try:
        return str(obj)
    except Exception:
        return ""

def _cell_text_matches(text, keywords, units):
    t = norm(text)
    unit_list = _units_list(units)

    keyword_ok = any(k in t for k in keywords)
    unit_ok = any(u.lower().replace("²", "^2") in t.replace("²", "^2") for u in unit_list)

    return keyword_ok

def _find_lab2_table_candidates(ws, table_def, expected_n=50, scan_rows=120, scan_cols=80):
    """
    Finds rows where the four required headers appear close together.
    Uses keyword-only detection so merged title rows and unit formatting do not break it.
    """
    max_r = min(scan_rows, ws.max_row)
    max_c = min(scan_cols, ws.max_column)

    cols = table_def["columns"]
    candidates = []

    for r in range(1, max_r + 1):
        row_texts = [norm(ws.cell(r, c).value) for c in range(1, max_c + 1)]

        # skip rows that are just table titles
        headerish_count = 0
        for col_def in cols:
            if any(any(k in txt for k in col_def["keywords"]) for txt in row_texts):
                headerish_count += 1

        if headerish_count < 4:
            continue

        hits = []
        used_cols = set()

        for col_def in cols:
            found_col = None
            for c in range(1, max_c + 1):
                if c in used_cols:
                    continue

                txt = row_texts[c - 1]

                if any(k in txt for k in col_def["keywords"]):
                    found_col = c
                    used_cols.add(c)
                    break

            hits.append(found_col)

        if all(h is not None for h in hits):
            spread = max(hits) - min(hits)

            # Lab 2 tables are four nearby columns
            if spread <= 10:
                candidates.append({
                    "row": r,
                    "cols": hits,
                    "spread": spread,
                })

    return candidates

def _count_numeric_rows(ws_v, header_row, col_indexes, expected_n=50):
    start = header_row + 1
    good_rows = 0

    for r in range(start, start + expected_n):
        numeric_count = 0
        for c in col_indexes:
            if try_float(ws_v.cell(r, c).value) is not None:
                numeric_count += 1

        # require at least 3 of 4 numeric entries on a row
        if numeric_count >= 3:
            good_rows += 1

    return good_rows

def _grade_lab2_table(ws_f, ws_v, table_def, expected_n=50):
    pts = 0.0
    notes = []

    candidates = _find_lab2_table_candidates(ws_f, table_def, expected_n=expected_n)

    if not candidates:
        notes.append(f"{table_def['name']}: I could not find the required time, position, velocity, and acceleration headers with units.")
        return 0.0, notes

    # choose first reasonable candidate
    cand = candidates[0]
    header_row = cand["row"]
    col_indexes = cand["cols"]

    # 0.75 headers/units
    pts += 0.75

    # 0.75 rows of numeric data
    numeric_rows = _count_numeric_rows(ws_v, header_row, col_indexes, expected_n=expected_n)
    # 0.75 rows of numeric data, but do not over-penalize short tables
    row_fraction = min(numeric_rows / expected_n, 1.0)

    if numeric_rows >= expected_n:
        row_pts = 0.75
    elif numeric_rows >= 10:
        row_pts = 0.50
    elif numeric_rows >= 5:
        row_pts = 0.25
    else:
        row_pts = 0.0

    pts += row_pts

    if numeric_rows < expected_n:
        notes.append(f"{table_def['name']}: I found about {numeric_rows}/{expected_n} usable numeric data rows.")

# 0.50 usable structure
# Give structure credit if the table columns are correct and there is a reasonable amount of data.
    if len(set(col_indexes)) == 4 and numeric_rows >= 10:
        pts += 0.50
    else:
        notes.append(f"{table_def['name']}: the table structure was not fully usable.")

    return round(min(2.0, pts), 2), notes

def _chart_title_exists(chart):
    return bool(_text_from_chart_obj(getattr(chart, "title", None)).strip())

def _axis_title_text(axis):
    return _text_from_chart_obj(getattr(axis, "title", None))

def _has_time_unit(axis_text):
    s = norm(axis_text)
    return "s" in s

def _has_y_units(axis_text):
    """
    Returns partial credit for:
    position: m
    velocity: m/s
    acceleration: m/s^2 or m/s²
    Avoids giving position credit just because m appears inside m/s.
    """
    raw = str(axis_text).lower().replace("²", "^2")

    has_acc = ("m/s^2" in raw)
    has_vel = ("m/s" in raw)

    cleaned = raw.replace("m/s^2", "").replace("m/s", "")
    has_pos = "m" in cleaned

    pts = 0.0
    missing = []

    if has_pos:
        pts += 0.33
    else:
        missing.append("position unit m")

    if has_vel:
        pts += 0.33
    else:
        missing.append("velocity unit m/s")

    if has_acc:
        pts += 0.34
    else:
        missing.append("acceleration unit m/s²")

    return pts, missing

def _series_is_linked(series):
    """
    Checks whether a chart series appears linked to worksheet cell ranges.
    """
    blob = str(series)
    return "!" in blob and "$" in blob

def _series_title_text(series):
    return _text_from_chart_obj(getattr(series, "tx", None))

def _series_name_hits(series_list):
    blob = " ".join(_series_title_text(s) for s in series_list).lower()

    pos = ("pos" in blob) or ("position" in blob)
    vel = ("vel" in blob) or ("velocity" in blob)
    acc = ("acc" in blob) or ("accel" in blob) or ("acceleration" in blob)

    return pos, vel, acc

def _collect_charts(workbook):
    charts = []
    for ws in workbook.worksheets:
        for ch in getattr(ws, "_charts", []):
            charts.append(ch)
    return charts

def grade_lab2_motion_graphs(path, rubric):
    wb_f = _grading_workbook(path, data_only=False)
    wb_v = _grading_workbook(path, data_only=True)

    ws_f = wb_f[wb_f.sheetnames[0]]
    ws_v = wb_v[wb_v.sheetnames[0]]

    total = 0.0
    notes = []

    expected_n = rubric.get("expected_n", 50)

    # 1–2. Tables: 4 pts total
    table_scores = []
    for table_def in rubric["tables"]:
        best_p = 0.0
        best_fb = [f"{table_def['name']}: I could not find the required time, position, velocity, and acceleration headers with units."]

        for sheet_name in wb_f.sheetnames:
            ws_f_try = wb_f[sheet_name]
            ws_v_try = wb_v[sheet_name]

            p, fb = _grade_lab2_table(ws_f_try, ws_v_try, table_def, expected_n=expected_n)

            if p > best_p:
                best_p = p
                best_fb = fb

        table_scores.append(best_p)
        notes.extend(best_fb)

    total += sum(table_scores)

    # 3–5. Graph checks: 6 pts total
    charts = _collect_charts(wb_f)
    expected_charts = rubric["graphs"]["expected_count"]
    expected_series = rubric["graphs"]["expected_series_per_graph"]

    # Graph existence + series + linkage: 2 pts
    graph_pts = 0.0

    if len(charts) >= expected_charts:
        graph_pts += 0.50
    else:
        graph_pts += 0.50 * min(len(charts) / expected_charts, 1.0)
        notes.append(f"Graphs: I found {len(charts)}/{expected_charts} required graphs.")

    charts_to_check = charts[:expected_charts]

    meaningful_legend_count = 0
    for ch in charts_to_check:
        series_list = list(getattr(ch, "series", []))
        pos, vel, acc = _series_name_hits(series_list)
        if pos and vel and acc:
            meaningful_legend_count += 1

    series_ok_count = 0
    linked_ok_count = 0

    for ch in charts_to_check:
        series_list = list(getattr(ch, "series", []))

        if len(series_list) >= expected_series:
            series_ok_count += 1

        if series_list and all(_series_is_linked(s) for s in series_list[:expected_series]):
            linked_ok_count += 1

    graph_pts += 0.25 * (series_ok_count / expected_charts if expected_charts else 0)
    graph_pts += 1.25 * (linked_ok_count / expected_charts if expected_charts else 0)

    if series_ok_count < expected_charts:
        notes.append("Graphs: each graph should include three data series: position, velocity, and acceleration.")

    if linked_ok_count < expected_charts:
        notes.append("Graphs: chart series should be linked to worksheet data ranges.")

    total += min(2.0, graph_pts)

    # Titles + legend: 2 pts
    title_legend_pts = 0.0

    title_count = sum(1 for ch in charts_to_check if _chart_title_exists(ch))
    title_legend_pts += 0.50 * (title_count / expected_charts if expected_charts else 0)

    legend_count = sum(1 for ch in charts_to_check if getattr(ch, "legend", None) is not None)
    title_legend_pts += 1.00 * (legend_count / expected_charts if expected_charts else 0)

    title_legend_pts += 0.50 * (meaningful_legend_count / expected_charts if expected_charts else 0)

    if title_count < expected_charts:
        notes.append("Graphs: each graph should have a title.")

    if legend_count < expected_charts:
        notes.append("Graphs: each graph should include a legend.")

    if meaningful_legend_count < expected_charts:
        notes.append("Graphs: legend/series names should clearly identify position, velocity, and acceleration.")

    total += min(2.0, title_legend_pts)

    # Axes + units: 2 pts
    axes_pts = 0.0

    x_good = 0
    y_unit_pts_total = 0.0
    y_missing_messages = []

    for ch in charts_to_check:
        x_text = _axis_title_text(getattr(ch, "x_axis", None))
        y_text = _axis_title_text(getattr(ch, "y_axis", None))

        if x_text and _has_time_unit(x_text):
            x_good += 1

        y_pts, missing = _has_y_units(y_text)
        y_unit_pts_total += y_pts

        if missing:
            y_missing_messages.append(", ".join(missing))

    # x-axis: 1 point total
    axes_pts += 1.0 * (x_good / expected_charts if expected_charts else 0)

    # y-axis: 1 point total, partial by unit
    axes_pts += min(1.0, y_unit_pts_total / expected_charts if expected_charts else 0)

    if x_good < expected_charts:
        notes.append("Graphs: x-axis title should include time units, s.")

    if y_missing_messages:
        notes.append("Graphs: y-axis title should include units for position (m), velocity (m/s), and acceleration (m/s²).")

    total += min(2.0, axes_pts)

    total = round(min(10.0, total), 2)

    if total >= 9.995:
        return total, ["Everything needed was found: tables, graphs, series, titles, legends, axis units, and linked data. ✅"]

    return total, notes

def _sheet_text_blob(ws, max_rows=120, max_cols=80):
    parts = []
    for r in range(1, min(ws.max_row, max_rows) + 1):
        for c in range(1, min(ws.max_column, max_cols) + 1):
            v = ws.cell(r, c).value
            if v is not None:
                parts.append(str(v).lower())
    return " ".join(parts)

def _count_numeric_near_headers(ws_v, ws_f, header_keywords_a, header_keywords_b, max_rows=80, max_cols=40):
    best = 0

    for r in range(1, min(ws_f.max_row, max_rows) + 1):
        cols_a = []
        cols_b = []

        for c in range(1, min(ws_f.max_column, max_cols) + 1):
            txt = norm(ws_f.cell(r, c).value)
            if any(k in txt for k in header_keywords_a):
                cols_a.append(c)
            if any(k in txt for k in header_keywords_b):
                cols_b.append(c)

        for ca in cols_a:
            for cb in cols_b:
                if abs(ca - cb) <= 3 and ca != cb:
                    count = 0
                    for rr in range(r + 1, min(ws_v.max_row, r + 30) + 1):
                        if try_float(ws_v.cell(rr, ca).value) is not None and try_float(ws_v.cell(rr, cb).value) is not None:
                            count += 1
                    best = max(best, count)

    return best

def _has_label_and_value(ws_v, ws_f, label_keywords, max_rows=120, max_cols=80):
    for r in range(1, min(ws_f.max_row, max_rows) + 1):
        for c in range(1, min(ws_f.max_column, max_cols) + 1):
            txt = norm(ws_f.cell(r, c).value)
            if any(k in txt for k in label_keywords):
                # Look nearby for a number/formula result
                for rr in range(max(1, r - 1), min(ws_f.max_row, r + 2) + 1):
                    for cc in range(max(1, c - 3), min(ws_f.max_column, c + 4) + 1):
                        if try_float(ws_v.cell(rr, cc).value) is not None:
                            return True
    return False

def _count_random_trial_columns(ws_v, ws_f, max_rows=120, max_cols=80):
    best = 0

    for c in range(1, min(ws_f.max_column, max_cols) + 1):
        count = 0
        for r in range(1, min(ws_f.max_row, max_rows) + 1):
            if try_float(ws_v.cell(r, c).value) is not None:
                count += 1
        best = max(best, count)

    return best

def grade_lab3_projectile_uncertainty(path, rubric):
    wb_f = _grading_workbook(path, data_only=False)
    wb_v = _grading_workbook(path, data_only=True)

    total = 0.0
    notes = []

    # Collect all sheets
    sheets_f = wb_f.worksheets
    sheets_v = wb_v.worksheets

    # ---------- 1. Projectile table + graph: 2 pts ----------
    best_projectile_rows = 0
    for ws_f, ws_v in zip(sheets_f, sheets_v):
        rows = _count_numeric_near_headers(
            ws_v, ws_f,
            header_keywords_a=["angle", "deg", "degree"],
            header_keywords_b=["distance", "dist"],
        )
        best_projectile_rows = max(best_projectile_rows, rows)

    if best_projectile_rows >= rubric.get("expected_angle_rows", 5):
        total += 2.0
    elif best_projectile_rows >= 3:
        total += 1.5
        notes.append(f"Projectile table: I found about {best_projectile_rows}/5 angle-distance rows.")
    elif best_projectile_rows > 0:
        total += 1.0
        notes.append(f"Projectile table: I found only about {best_projectile_rows}/5 angle-distance rows.")
    else:
        notes.append("Projectile table: I could not find the angle and distance data table.")

    charts = _collect_charts(wb_f)
    if len(charts) >= 1:
        total += 2.0
    else:
        notes.append("Projectile graph: I could not find a graph for angle vs. distance.")

    # ---------- 2. Random error analysis: 4 pts ----------
    best_trials = 0
    for ws_f, ws_v in zip(sheets_f, sheets_v):
        best_trials = max(best_trials, _count_random_trial_columns(ws_v, ws_f))

    if best_trials >= rubric.get("expected_random_trials", 8):
        total += 1.0
    elif best_trials >= 5:
        total += 0.5
        notes.append(f"Random error: I found about {best_trials} trial values.")
    else:
        notes.append("Random error: I could not find enough repeated trial values.")

    blob_all = " ".join(_sheet_text_blob(ws) for ws in sheets_f)

    avg_found = any(_has_label_and_value(wv, wf, ["average", "mean"]) for wf, wv in zip(sheets_f, sheets_v))
    range_found = any(_has_label_and_value(wv, wf, ["range"]) for wf, wv in zip(sheets_f, sheets_v))
    range_unc_found = any(_has_label_and_value(
        wv, wf,
        ["range method", "range method unc", "range uncertainty", "range unc", "range unc error", "range. unc", "range. unc. error", "range error"]
    ) for wf, wv in zip(sheets_f, sheets_v))
    std_found = any(_has_label_and_value(wv, wf, ["std", "standard dev", "st.dev", "st dev"]) for wf, wv in zip(sheets_f, sheets_v))

    if avg_found:
        total += 1.0
    else:
        notes.append("Random error: I could not find the average/mean calculation.")

    if range_found and range_unc_found:
        total += 1.0
    elif range_found or range_unc_found:
        total += 0.5
        notes.append("Random error: I found part of the range/range-method uncertainty work, but not all of it.")
    else:
        notes.append("Random error: I could not find the range and range-method uncertainty.")

    if std_found:
        total += 1.0
    else:
        notes.append("Random error: I could not find the standard deviation method uncertainty.")

    # ---------- 3. Systematic error comparison: 2 pts ----------
    systematic_keywords = ["systematic", "error", "deg", "%", "no error"]
    systematic_hits = sum(1 for k in systematic_keywords if k in blob_all)

    numeric_systematic_values = 0
    for ws_v in sheets_v:
        for r in range(1, min(ws_v.max_row, 120) + 1):
            for c in range(1, min(ws_v.max_column, 80) + 1):
                v = try_float(ws_v.cell(r, c).value)
                if v is not None and 0 < v < 100:
                    numeric_systematic_values += 1

    if systematic_hits >= 3:
        total += 1.0
    else:
        notes.append("Systematic error: labels should clearly identify the systematic error comparisons.")

    if numeric_systematic_values >= 3:
        total += 1.0
    else:
        notes.append("Systematic error: I could not find the three comparison values.")

    total = round(min(10.0, total), 2)

    if total >= 9.995:
        return total, ["Everything needed was found: projectile table, graph, random error calculations, and systematic error values. ✅"]

    return total, notes

def _find_simple_table(ws_f, ws_v, table_def, scan_rows=80, scan_cols=40):
    max_r = min(scan_rows, ws_f.max_row)
    max_c = min(scan_cols, ws_f.max_column)

    for r in range(1, max_r + 1):
        found_cols = []

        for col_def in table_def["columns"]:
            col_match = None

            for c in range(1, max_c + 1):
                txt = norm(ws_f.cell(r, c).value)

                if any(k in txt for k in col_def["keywords"]):
                    col_match = c
                    break

            if col_match:
                found_cols.append(col_match)

        if len(found_cols) == len(table_def["columns"]):
            return r, found_cols

    return None, None

def _count_data_rows(ws_v, header_row, cols, max_rows=50):
    count = 0

    for r in range(header_row + 1, header_row + 1 + max_rows):
        numeric = sum(1 for c in cols if try_float(ws_v.cell(r, c).value) is not None)
        if numeric >= max(2, len(cols) - 1):
            count += 1

    return count

def _grade_lab4d_table(ws_f, ws_v, table_def):
    pts = 0.0
    notes = []

    header_row, cols = _find_simple_table(ws_f, ws_v, table_def)

    if header_row is None:
        notes.append(f"{table_def['name']}: I could not find the required columns.")
        return 0.0, notes

    is_friction_table = any(
        any(k in ["mu", "friction"] for k in col_def["keywords"])
        for col_def in table_def["columns"]
    )

    # ---------- No-friction table: 2 pts ----------
    if not is_friction_table:
        pts += 1.0  # headers/title/columns found

        data_rows = _count_data_rows(ws_v, header_row, cols)

        if data_rows >= 10:
            pts += 1.0
        elif data_rows >= 5:
            pts += 0.5
            notes.append(f"{table_def['name']}: I found about {data_rows} data rows.")
        else:
            notes.append(f"{table_def['name']}: not enough usable data rows.")

        return min(2.0, pts), notes

    # ---------- Friction table: 2 pts ----------
    # 1.0 table structure/data
    data_rows = _count_data_rows(ws_v, header_row, cols)

    if data_rows >= 10:
        pts += 1.0
    elif data_rows >= 5:
        pts += 0.5
        notes.append(f"{table_def['name']}: I found about {data_rows} data rows.")
    else:
        notes.append(f"{table_def['name']}: not enough usable data rows.")

    # Find the mu/friction coefficient column
    mu_col = None
    for i, col_def in enumerate(table_def["columns"]):
        joined = " ".join(col_def["keywords"]).lower()
        if "mu" in joined or "friction" in joined:
            mu_col = cols[i]
            break

    if mu_col is None:
        notes.append("Friction table: I could not find the calculated hidden friction coefficient column.")
        return min(2.0, pts), notes

    formula_count = 0
    correct_count = 0
    value_count = 0

    for r in range(header_row + 1, header_row + 21):
        cell_formula = ws_f.cell(r, mu_col).value
        cell_value = try_float(ws_v.cell(r, mu_col).value)

        if cell_value is None:
            continue

        value_count += 1

        if is_formula(cell_formula):
            formula_count += 1

        # We are not grading sig figs.
        # Rounded values like 0.011 should pass.
        if close(cell_value, 0.011, rel=0.25, abs_tol=0.005):
            correct_count += 1

    # 0.5 formula/process credit
    if formula_count >= 10:
        pts += 0.5
    elif formula_count >= 5:
        pts += 0.25
        notes.append("Friction table: only some μ*k values appear to be calculated with formulas.")
    else:
        notes.append("Friction table: μ*k values should be calculated using formulas, not typed.")

    # 0.5 approximate correctness credit
    if correct_count >= 10:
        pts += 0.5
    elif correct_count >= 5:
        pts += 0.25
        notes.append("Friction table: only some μ*k values are close to the expected value.")
    else:
        notes.append("Friction table: μ*k values are not close to the expected value around 0.011.")

    return min(2.0, pts), notes

def grade_lab4d_atwood_friction(path, rubric):
    wb_f = _grading_workbook(path, data_only=False)
    wb_v = _grading_workbook(path, data_only=True)

    total = 0.0
    notes = []

    # ---------- Tables (4 pts) ----------
    for table_def in rubric["tables"]:
        best = 0.0
        best_fb = []

        for ws_name in wb_f.sheetnames:
            p, fb = _grade_lab4d_table(
                wb_f[ws_name],
                wb_v[ws_name],
                table_def
            )
            if p > best:
                best = p
                best_fb = fb

        total += best
        notes.extend(best_fb)

    # ---------- Graphs (6 pts) ----------
    charts = _collect_charts(wb_f)

    if len(charts) >= 2:
        total += 2.0
    else:
        total += len(charts)
        notes.append(f"Graphs: I found {len(charts)}/2 graphs.")

    # titles (2 pts)
    title_count = sum(1 for ch in charts if _chart_title_exists(ch))
    total += min(2.0, title_count)

    if title_count < 2:
        notes.append("Graphs: each graph should have a title.")

    # axes (2 pts)
    # OpenPyXL sometimes fails to read Excel axis titles cleanly,
    # so we check more broadly for axis-title text inside the chart object.
    axes_good = 0

    for ch in charts[:2]:
        chart_blob = str(ch).lower().replace("²", "^2")

        x = _axis_title_text(getattr(ch, "x_axis", None))
        y = _axis_title_text(getattr(ch, "y_axis", None))

        axis_blob = " ".join([str(x), str(y), chart_blob]).lower().replace("²", "^2")

        has_x_unit = (
            "mass 2" in axis_blob
            or "kg" in axis_blob
        )

        has_y_unit = (
            "acceleration" in axis_blob
            or "m/s^2" in axis_blob
            or "mu" in axis_blob
            or "friction" in axis_blob
            or "unitless" in axis_blob
        )

        if has_x_unit and has_y_unit:
            axes_good += 1

    total += min(2.0, axes_good)

    if axes_good < 2:
        notes.append("Graphs: include axis labels with units.")

    total = round(min(10.0, total), 2)

    if total >= 9.995:
        return total, ["Everything needed was found: both tables and both graphs. ✅"]

    return total, notes

def _cell_has_any(text, keywords):
    t = norm(text)
    return any(k in t for k in keywords)

def _find_rows_with_keywords(ws, keyword_groups, scan_rows=120, scan_cols=80):
    matches = []

    for r in range(1, min(ws.max_row, scan_rows) + 1):
        row_blob = " ".join(
            norm(ws.cell(r, c).value)
            for c in range(1, min(ws.max_column, scan_cols) + 1)
            if ws.cell(r, c).value is not None
        )

        if all(any(k in row_blob for k in group) for group in keyword_groups):
            matches.append(r)

    return matches

def _count_formulas_in_row_range(ws_f, header_row, cols, n_rows=5):
    formula_count = 0
    value_count = 0

    for r in range(header_row + 1, header_row + 1 + n_rows):
        for c in cols:
            v = ws_f.cell(r, c).value
            if v is not None:
                value_count += 1
            if is_formula(v):
                formula_count += 1

    return formula_count, value_count

def _find_summary_table_headers(ws_f, scan_rows=80, scan_cols=80):
    """
    Looks for compact conservation summary tables with KE, p, and conservation columns.
    """
    candidates = []

    for r in range(1, min(ws_f.max_row, scan_rows) + 1):
        hits = {
            "ke": [],
            "p": [],
            "conservation": [],
        }

        for c in range(1, min(ws_f.max_column, scan_cols) + 1):
            txt = norm(ws_f.cell(r, c).value)

            if "ke" in txt or "joule" in txt:
                hits["ke"].append(c)

            if "p " in f" {txt} " or "kg*m/s" in txt or "momentum" in txt:
                hits["p"].append(c)

            if "conserved" in txt or "conservation" in txt:
                hits["conservation"].append(c)

        if len(hits["ke"]) >= 2 and len(hits["p"]) >= 2 and len(hits["conservation"]) >= 1:
            candidates.append(r)

    return candidates

def _grade_one_summary_table(ws_f, ws_v, header_row):
    pts = 0.0
    notes = []

    max_c = min(ws_f.max_column, 80)

    conservation_cols = []
    for c in range(1, max_c + 1):
        txt = norm(ws_f.cell(header_row, c).value)
        txt_above = norm(ws_f.cell(header_row - 1, c).value) if header_row > 1 else ""

        if "conserved" in txt or "conservation" in txt or "conserved" in txt_above:
            conservation_cols.append(c)

    # fallback: often the final two columns are the conservation columns
    if len(conservation_cols) < 2:
        conservation_cols = list(range(max(1, max_c - 1), max_c + 1))

    conservation_cols = conservation_cols[-2:]

    formula_count, value_count = _count_formulas_in_row_range(
        ws_f, header_row, conservation_cols, n_rows=5
    )

    values = []
    for r in range(header_row + 1, header_row + 6):
        for c in conservation_cols:
            v = try_float(ws_v.cell(r, c).value)
            if v is not None:
                values.append(v)

    if value_count >= 5:
        pts += 0.35
    else:
        notes.append("Summary table: conservation values are missing or incomplete.")

    if formula_count >= 5:
        pts += 0.35
    elif formula_count > 0:
        pts += 0.20
        notes.append("Summary table: only some conservation values use formulas.")
    else:
        notes.append("Summary table: conservation values should be calculated with formulas.")

    reasonable = [v for v in values if 0 <= v <= 110]

    if len(reasonable) >= max(3, len(values) * 0.7):
        pts += 0.30
    else:
        notes.append("Summary table: conservation percentages should usually be between 0 and 100.")

    return min(1.0, pts), notes

def _find_time_table_headers(ws_f, scan_rows=250, scan_cols=120):
    """
    Looks for data-over-time collision tables.

    Expected idea:
    time, v1, v2, KE total, p total / momentum total

    This version is intentionally forgiving because students may write:
    - t (s)
    - Time
    - v1 (m/s)
    - v2 (m/s)
    - KE Total (J)
    - p Total (kg*m/s)
    - Momentum
    """
    candidates = []

    for r in range(1, min(ws_f.max_row, scan_rows) + 1):
        row_blob = " ".join(
            norm(ws_f.cell(r, c).value)
            for c in range(1, min(ws_f.max_column, scan_cols) + 1)
            if ws_f.cell(r, c).value is not None
        )

        row_blob = row_blob.lower().replace("²", "^2")

        has_time = (
            "time" in row_blob
            or "t (s)" in row_blob
            or ("t" in row_blob and "s" in row_blob)
        )

        has_v1 = "v1" in row_blob or "velocity 1" in row_blob
        has_v2 = "v2" in row_blob or "velocity 2" in row_blob

        has_ke = (
            "ke" in row_blob
            or "kinetic" in row_blob
            or "energy" in row_blob
        )

        has_p = (
            "p total" in row_blob
            or "momentum" in row_blob
            or "kg*m/s" in row_blob
            or "kg m/s" in row_blob
            or "kg·m/s" in row_blob
        )

        if has_time and has_v1 and has_v2 and has_ke and has_p:
            candidates.append(r)

    return candidates

def _grade_one_time_table(ws_f, ws_v, header_row, expected_rows=50):
    pts = 0.0
    notes = []

    max_c = min(ws_f.max_column, 120)

    cols = {
        "time": None,
        "v1": None,
        "v2": None,
        "ke": None,
        "p": None,
    }

    # Look at header row plus the row above/below in case headers are split/merged
    for c in range(1, max_c + 1):
        txt_parts = []
        for rr in [header_row - 1, header_row, header_row + 1]:
            if 1 <= rr <= ws_f.max_row:
                txt_parts.append(norm(ws_f.cell(rr, c).value))

        txt = " ".join(txt_parts).lower().replace("²", "^2")

        if cols["time"] is None and (
            "time" in txt or "t (s)" in txt or txt.strip() in ["t", "t s"]
        ):
            cols["time"] = c

        elif cols["v1"] is None and (
            "v1" in txt or "velocity 1" in txt or "velocity one" in txt
        ):
            cols["v1"] = c

        elif cols["v2"] is None and (
            "v2" in txt or "velocity 2" in txt or "velocity two" in txt
        ):
            cols["v2"] = c

        elif cols["ke"] is None and (
            "ke" in txt or "kinetic" in txt or "energy" in txt
        ):
            cols["ke"] = c

        elif cols["p"] is None and (
            "p total" in txt or "momentum" in txt or "kg*m/s" in txt or "kg m/s" in txt or "kg·m/s" in txt
        ):
            cols["p"] = c

    if not all(cols.values()):
        notes.append(
            "Time table: required columns should include time, v1, v2, KE total, and total momentum."
        )
        return 0.0, notes

    numeric_rows = 0
    for r in range(header_row + 1, header_row + 1 + expected_rows):
        numeric_count = sum(
            1 for c in cols.values()
            if try_float(ws_v.cell(r, c).value) is not None
        )
        if numeric_count >= 4:
            numeric_rows += 1

    if numeric_rows >= 45:
        pts += 0.40
    elif numeric_rows >= 25:
        pts += 0.25
        notes.append(f"Time table: I found about {numeric_rows}/{expected_rows} usable rows.")
    else:
        notes.append(f"Time table: I found only about {numeric_rows}/{expected_rows} usable rows.")

    formula_cols = [cols["ke"], cols["p"]]
    formula_count, value_count = _count_formulas_in_row_range(
        ws_f,
        header_row,
        formula_cols,
        n_rows=max(1, min(numeric_rows, expected_rows))
    )

    if formula_count >= 0.8 * max(1, value_count):
        pts += 0.40
    elif formula_count > 0:
        pts += 0.20
        notes.append("Time table: only some KE total and momentum values use formulas.")
    else:
        notes.append("Time table: KE total and total momentum should be calculated with formulas.")

    pts += 0.20  # columns/structure present

    return min(1.0, pts), notes

def _grade_collision_graphs(wb_f, expected_graphs=3):
    total = 0.0
    notes = []

    charts = _collect_charts(wb_f)
    charts_to_check = charts[:expected_graphs]

    # 1 pt — graph count
    if len(charts) >= expected_graphs:
        total += 1.0
    else:
        total += len(charts) / expected_graphs
        notes.append(f"Graphs: I found {len(charts)}/{expected_graphs} graphs.")

    # 1 pt — titles + legends
    title_count = sum(1 for ch in charts_to_check if _chart_title_exists(ch))
    legend_count = sum(1 for ch in charts_to_check if getattr(ch, "legend", None) is not None)

    default_title_count = 0
    generic_legend_count = 0

    for ch in charts_to_check:
        title_text = _text_from_chart_obj(getattr(ch, "title", None)).lower()
        if "chart title" in title_text:
            default_title_count += 1

        series_blob = " ".join(str(s) for s in getattr(ch, "series", [])).lower()
        if "series 1" in series_blob or "series1" in series_blob:
            generic_legend_count += 1

    title_legend_pts = 0.0
    title_legend_pts += 0.5 * (title_count / expected_graphs)
    title_legend_pts += 0.5 * (legend_count / expected_graphs)

    # small penalty if titles/legends exist but are generic
    if default_title_count > 0:
        title_legend_pts = max(0.0, title_legend_pts - 0.25)

    if generic_legend_count > 0:
        title_legend_pts = max(0.0, title_legend_pts - 0.25)

    total += min(1.0, title_legend_pts)

    if title_count < expected_graphs:
        notes.append("Graphs: each graph should have a clear, descriptive title.")
    elif default_title_count > 0:
        notes.append(
            'Graphs: replace the default "Chart Title" with a descriptive title that explains what the graph shows.'
        )

    if legend_count < expected_graphs:
        notes.append("Graphs: each graph should include a legend.")
    elif generic_legend_count > 0:
        notes.append(
            "Graphs: legend labels should identify v1, v2, KE total, and total momentum instead of generic Series names."
        )

    # 1 pt — axes
    axes_pts = 0.0
    swapped_axis_count = 0
    missing_axis_count = 0

    for ch in charts_to_check:
        chart_blob = str(ch).lower().replace("²", "^2")

        x = _axis_title_text(getattr(ch, "x_axis", None))
        y = _axis_title_text(getattr(ch, "y_axis", None))

        x_blob = str(x).lower().replace("²", "^2")
        y_blob = str(y).lower().replace("²", "^2")
        axis_blob = " ".join([x_blob, y_blob, chart_blob]).lower().replace("²", "^2")

        x_has_time = (
            "time" in x_blob
            or "t (s)" in x_blob
            or ("t" in x_blob and "s" in x_blob)
        )

        y_has_data_units = (
            "m/s" in y_blob
            or "joule" in y_blob
            or "j)" in y_blob
            or "kg*m/s" in y_blob
            or "kg m/s" in y_blob
            or "kg·m/s" in y_blob
            or "ke" in y_blob
            or "momentum" in y_blob
            or "velocity" in y_blob
        )

        y_has_time = (
            "time" in y_blob
            or "t (s)" in y_blob
            or ("t" in y_blob and "s" in y_blob)
        )

        x_has_data_units = (
            "m/s" in x_blob
            or "joule" in x_blob
            or "j)" in x_blob
            or "kg*m/s" in x_blob
            or "kg m/s" in x_blob
            or "kg·m/s" in x_blob
            or "ke" in x_blob
            or "momentum" in x_blob
            or "velocity" in x_blob
        )

        if x_has_time and y_has_data_units:
            axes_pts += 1 / expected_graphs
        elif y_has_time or x_has_data_units:
            # labels are present, but likely assigned to the wrong axes
            axes_pts += 0.5 / expected_graphs
            swapped_axis_count += 1
        elif x_blob or y_blob or axis_blob:
            axes_pts += 0.25 / expected_graphs
            missing_axis_count += 1
        else:
            missing_axis_count += 1

    total += min(1.0, axes_pts)

    if swapped_axis_count > 0:
        notes.append(
            "Graphs: axis labels appear to be swapped or placed on the wrong axes; time should be on the x-axis."
        )
    elif missing_axis_count > 0:
        notes.append("Graphs: axes should be labeled with units, including time (s).")

    # 1 pt — correct data series
    series_pts = 0.0

    for ch in charts_to_check:
        series_blob = " ".join(str(s) for s in getattr(ch, "series", [])).lower()

        has_v1 = "v1" in series_blob or "velocity 1" in series_blob
        has_v2 = "v2" in series_blob or "velocity 2" in series_blob
        has_ke = "ke" in series_blob or "kinetic" in series_blob or "energy" in series_blob
        has_p = "p total" in series_blob or "momentum" in series_blob

        hits = sum([has_v1, has_v2, has_ke, has_p])
        series_pts += (hits / 4) / expected_graphs

    total += min(1.0, series_pts)

    if series_pts < 0.99:
        notes.append(
            "Graphs: the legend/series labels should clearly identify v1, v2, KE total, and total momentum."
        )

    return min(4.0, total), notes

def grade_lab6_7_collisions(path, rubric):
    wb_f = _grading_workbook(path, data_only=False)
    wb_v = _grading_workbook(path, data_only=True)

    total = 0.0
    notes = []

    expected_time_rows = rubric.get("expected_time_rows", 50)

    # ---------- Summary tables: 3 pts ----------
    summary_scores = []

    for ws_name in wb_f.sheetnames:
        ws_f = wb_f[ws_name]
        ws_v = wb_v[ws_name]

        header_rows = _find_summary_table_headers(ws_f)

        for hr in header_rows:
            p, fb = _grade_one_summary_table(ws_f, ws_v, hr)
            summary_scores.append((p, fb))

    summary_scores = sorted(summary_scores, key=lambda x: x[0], reverse=True)[:3]

    total += sum(p for p, fb in summary_scores)

    if len(summary_scores) < 3:
        notes.append(f"Summary tables: I found {len(summary_scores)}/3 conservation summary tables.")

    for p, fb in summary_scores:
        if p < 0.99:
            notes.extend(fb)

    # ---------- Time tables: 3 pts ----------
    time_scores = []

    for ws_name in wb_f.sheetnames:
        ws_f = wb_f[ws_name]
        ws_v = wb_v[ws_name]

        header_rows = _find_time_table_headers(ws_f)

        for hr in header_rows:
            p, fb = _grade_one_time_table(ws_f, ws_v, hr, expected_rows=expected_time_rows)
            time_scores.append((p, fb))

    time_scores = sorted(time_scores, key=lambda x: x[0], reverse=True)[:3]

    total += sum(p for p, fb in time_scores)

    if len(time_scores) < 3:
        likely_side_by_side_time_tables = any(
            ws.max_row >= 45 and ws.max_column >= 17
            for ws in wb_f.worksheets
        )

        if len(time_scores) >= 1 and likely_side_by_side_time_tables:
            print("[DEBUG] Time tables likely side-by-side; detection incomplete.")

            missing = 3 - len(time_scores)
            total += 0.70 * missing

        else:
            notes.append(f"Time tables: I found {len(time_scores)}/3 collision-over-time tables.")

    for p, fb in time_scores:
        if p < 0.99:
            notes.extend(fb)

    # ---------- Graphs: 4 pts ----------
    graph_pts, graph_fb = _grade_collision_graphs(
        wb_f,
        expected_graphs=rubric.get("graphs_expected", 3)
    )

    total += graph_pts
    notes.extend(graph_fb)

    total = round(min(10.0, total), 2)

    if total >= 9.995:
        return total, ["Everything needed was found: summary tables, time tables, formulas, and graphs. ✅"]

    # remove duplicate feedback lines
    deduped = []
    seen = set()
    for n in notes:
        key = n.lower().strip()
        if key not in seen:
            seen.add(key)
            deduped.append(n)

    notes = deduped

    return total, notes

def _is_iexp_label(txt):
    txt = norm(txt).lower().replace("²", "^2")
    txt = txt.replace(" ", "").replace(".", "")

    return (
        "i_exp" in txt
        or "i-exp" in txt
        or "i(exp" in txt
        or "iexp" in txt
        or "lexp" in txt
        or "experimental" in txt
        or "exp" in txt
    )

def _is_itheo_label(txt):
    txt = norm(txt).lower().replace("²", "^2")
    txt = txt.replace(" ", "").replace(".", "")

    return (
        "i_theo" in txt
        or "i-theo" in txt
        or "i(theo" in txt
        or "itheo" in txt
        or "ltheo" in txt
        or "theoretical" in txt
        or "theo" in txt
    )

def _find_rotation_summary_headers(ws_f, scan_rows=80, scan_cols=80):
    candidates = []

    for r in range(1, min(ws_f.max_row, scan_rows) + 1):
        row_blob = " ".join(
            norm(ws_f.cell(r, c).value)
            for c in range(1, min(ws_f.max_column, scan_cols) + 1)
            if ws_f.cell(r, c).value is not None
        ).lower().replace("²", "^2")

        has_mass = "mass" in row_blob and "kg" in row_blob
        has_radius = "radius" in row_blob or "r(" in row_blob or "r (" in row_blob
        has_alpha = (
            "α" in row_blob
            or "alpha" in row_blob
            or "a (rad/s^2" in row_blob
            or "a(rad/s^2" in row_blob
            or "rad/s^2" in row_blob
            or "rad/s²" in row_blob
        )
        has_iexp = _is_iexp_label(row_blob)
        has_itheo = _is_itheo_label(row_blob)

        # trial is optional for detection, but graded later
        has_inertia_col = has_iexp or has_itheo

        # Detect the table even if one required non-calculation column is missing.
        # Missing columns are penalized later in _grade_one_rotation_summary_table.
        core_hits = sum([has_mass, has_radius, has_alpha, has_inertia_col])

        if core_hits >= 3 and has_inertia_col:
            candidates.append(r)

    return candidates

def _grade_one_rotation_summary_table(ws_f, ws_v, header_row):
    pts = 0.0
    notes = []

    max_c = min(ws_f.max_column, 80)

    # Find I_exp and I_theo columns separately so missing one does not erase the whole table.
    iexp_cols = []
    itheo_cols = []

    for c in range(1, max_c + 1):
        txt = ws_f.cell(header_row, c).value

        if _is_iexp_label(txt):
            iexp_cols.append(c)

        if _is_itheo_label(txt):
            itheo_cols.append(c)

    # 1.0 pts — usable 10-row data table
    numeric_rows = 0
    for r in range(header_row + 1, header_row + 11):
        numeric_count = 0
        for c in range(1, max_c + 1):
            if try_float(ws_v.cell(r, c).value) is not None:
                numeric_count += 1
        if numeric_count >= 5:
            numeric_rows += 1

    if numeric_rows >= 10:
        pts += 1.0
    elif numeric_rows >= 5:
        pts += 0.5
        notes.append("Rotation summary table: I found only part of the expected 10 rows.")
    else:
        notes.append("Rotation summary table: not enough usable data rows.")

    # 1.25 pts — I_exp/I_theo formulas, split evenly
    formula_pts = 0.0
    n_rows_to_check = max(1, numeric_rows)

    if iexp_cols:
        fc, vc = _count_formulas_in_row_range(
            ws_f, header_row, [iexp_cols[-1]], n_rows=n_rows_to_check
        )

        if fc >= 0.8 * max(1, vc):
            formula_pts += 0.625
        elif fc > 0:
            formula_pts += 0.325
            notes.append("Rotation summary table: only some I_exp values use formulas.")
        else:
            notes.append("Rotation summary table: I_exp should be calculated with formulas.")
    else:
        notes.append("Rotation summary table: missing the I_exp calculation column.")

    if itheo_cols:
        fc, vc = _count_formulas_in_row_range(
            ws_f, header_row, [itheo_cols[-1]], n_rows=n_rows_to_check
        )

        if fc >= 0.8 * max(1, vc):
            formula_pts += 0.625
        elif fc > 0:
            formula_pts += 0.325
            notes.append("Rotation summary table: only some I_theo values use formulas.")
        else:
            notes.append("Rotation summary table: I_theo should be calculated with formulas.")
    else:
        notes.append("Rotation summary table: missing the I_theo calculation column.")

    pts += formula_pts

    # 0.75 pts — structure/header credit, partial by required columns
    row_blob = " ".join(
        norm(ws_f.cell(header_row, c).value)
        for c in range(1, max_c + 1)
        if ws_f.cell(header_row, c).value is not None
    ).lower().replace("²", "^2")

    structure_pts = 0.0

    has_trial = "trial" in row_blob
    has_mass = "mass" in row_blob and "kg" in row_blob
    has_radius = "radius" in row_blob or "r(" in row_blob or "r (" in row_blob
    has_alpha = (
        "α" in row_blob
        or "alpha" in row_blob
        or "a (rad/s^2" in row_blob
        or "a(rad/s^2" in row_blob
        or "rad/s^2" in row_blob
        or "rad/s²" in row_blob
    )

    if has_trial:
        structure_pts += 0.15
    else:
        notes.append("Rotation summary table: include the Trial # column.")

    if has_mass:
        structure_pts += 0.15
    else:
        notes.append("Rotation summary table: include mass column headers with kg.")

    if has_radius:
        structure_pts += 0.15
    else:
        notes.append("Rotation summary table: include radius column headers with m.")

    if has_alpha:
        structure_pts += 0.15
    else:
        notes.append("Rotation summary table: include angular acceleration with rad/s².")

    if iexp_cols:
        structure_pts += 0.075
    else:
        notes.append("Rotation summary table: include the I_exp header with units.")

    if itheo_cols:
        structure_pts += 0.075
    else:
        notes.append("Rotation summary table: include the I_theo header with units.")

    pts += structure_pts

    return min(3.0, pts), notes

def _find_rotation_time_table_headers(ws_f, scan_rows=200, scan_cols=120):
    candidates = []

    for r in range(1, min(ws_f.max_row, scan_rows) + 1):
        row_blob = " ".join(
            norm(ws_f.cell(r, c).value)
            for c in range(1, min(ws_f.max_column, scan_cols) + 1)
            if ws_f.cell(r, c).value is not None
        ).lower().replace("²", "^2")

        has_time = (
            "time" in row_blob
            or "t (s)" in row_blob
            or ("t" in row_blob and "s" in row_blob)
        )

        has_theta = (
            "θ" in row_blob
            or "theta" in row_blob
            or "o(rad" in row_blob
            or "o (" in row_blob
        )

        has_omega = (
            "ω" in row_blob
            or "omega" in row_blob
            or "w(rad/s" in row_blob
            or "w (" in row_blob
            or "rad/s" in row_blob
        )

        has_alpha = (
            "α" in row_blob
            or "alpha" in row_blob
            or "a(rad/s^2" in row_blob
            or "a (" in row_blob
            or "rad/s^2" in row_blob
        )

        if has_time and has_theta and has_omega and has_alpha:
            candidates.append(r)

    return candidates

def _grade_rotation_time_table(ws_f, ws_v, header_row, expected_rows=50):
    pts = 0.0
    notes = []

    max_c = min(ws_f.max_column, 120)

    cols = {
        "time": None,
        "theta": None,
        "omega": None,
        "alpha": None,
    }

    for c in range(1, max_c + 1):
        txt_parts = []
        for rr in [header_row - 1, header_row, header_row + 1]:
            if 1 <= rr <= ws_f.max_row:
                txt_parts.append(norm(ws_f.cell(rr, c).value))

        txt = " ".join(txt_parts).lower().replace("²", "^2")

        if cols["time"] is None and ("time" in txt or "t (s)" in txt or txt.strip() in ["t", "t s"]):
            cols["time"] = c
        elif cols["theta"] is None and ("θ" in txt or "theta" in txt):
            cols["theta"] = c
        elif cols["omega"] is None and ("ω" in txt or "omega" in txt or "rad/s" in txt):
            cols["omega"] = c
        elif cols["alpha"] is None and ("α" in txt or "alpha" in txt or "rad/s^2" in txt):
            cols["alpha"] = c

    if not all(cols.values()):
        notes.append("Rotation time table: required columns should include time, θ, ω, and α.")
        return 0.0, notes

    numeric_rows = 0
    for r in range(header_row + 1, header_row + 1 + expected_rows):
        numeric_count = sum(
            1 for c in cols.values()
            if try_float(ws_v.cell(r, c).value) is not None
        )
        if numeric_count >= 3:
            numeric_rows += 1

    # 1.5 pts — enough numeric rows
    if numeric_rows >= 45:
        pts += 1.5
    elif numeric_rows >= 25:
        pts += 1.0
        notes.append(f"Rotation time table: I found about {numeric_rows}/{expected_rows} usable rows.")
    elif numeric_rows > 0:
        pts += 0.5
        notes.append(f"Rotation time table: I found only about {numeric_rows}/{expected_rows} usable rows.")
    else:
        notes.append("Rotation time table: I could not find usable numeric rows.")

    # 0.5 pts — structure/columns present
    pts += 0.5

    return min(2.0, pts), notes

def _grade_rotation_graphs(wb_f, expected_graphs=1):
    total = 0.0
    notes = []

    charts = _collect_charts(wb_f)

    # 0.75 pts — graph present
    if len(charts) >= expected_graphs:
        total += 0.75
    else:
        total += 0.75 * (len(charts) / expected_graphs)
        notes.append(f"Graph: I found {len(charts)}/{expected_graphs} required graph.")
        return total, notes

    ch = charts[0]

    # 0.50 pts — title + legend
    title_exists = _chart_title_exists(ch)
    legend_exists = getattr(ch, "legend", None) is not None

    title_text = _text_from_chart_obj(getattr(ch, "title", None)).lower()
    default_title = "chart title" in title_text

    title_legend_pts = 0.0

    if title_exists:
        title_legend_pts += 0.25
    if legend_exists:
        title_legend_pts += 0.25

    if default_title:
        title_legend_pts = max(0.0, title_legend_pts - 0.15)
        notes.append(
            'Graph: replace the default "Chart Title" with a descriptive title that explains what the graph shows.'
        )

    if not title_exists:
        notes.append("Graph: include a descriptive chart title.")
    if not legend_exists:
        notes.append("Graph: include a legend identifying θ, ω, and α.")

    total += min(0.50, title_legend_pts)

    # 0.50 pts — axes
    x = _axis_title_text(getattr(ch, "x_axis", None))
    y = _axis_title_text(getattr(ch, "y_axis", None))

    x_blob = str(x).lower().replace("²", "^2")
    y_blob = str(y).lower().replace("²", "^2")

    x_has_time = "time" in x_blob or "t (s)" in x_blob or ("t" in x_blob and "s" in x_blob)
    y_has_rotation_units = (
        "rad" in y_blob
        or "rad/s" in y_blob
        or "rad/s^2" in y_blob
        or "theta" in y_blob
        or "omega" in y_blob
        or "alpha" in y_blob
        or "θ" in y_blob
        or "ω" in y_blob
        or "α" in y_blob
    )

    y_has_time = "time" in y_blob or "t (s)" in y_blob
    x_has_rotation_units = (
        "rad" in x_blob
        or "theta" in x_blob
        or "omega" in x_blob
        or "alpha" in x_blob
        or "θ" in x_blob
        or "ω" in x_blob
        or "α" in x_blob
    )

    if x_has_time and y_has_rotation_units:
        total += 0.50
    elif y_has_time or x_has_rotation_units:
        total += 0.25
        notes.append("Graph: axis labels appear to be swapped or placed on the wrong axes; time should be on the x-axis.")
    else:
        total += 0.10
        notes.append("Graph: axes should be labeled with units, including time (s) and radian-based quantities.")

    # 0.25 pts — series labels
    series_blob = " ".join(str(s) for s in getattr(ch, "series", [])).lower()

    has_theta = "θ" in series_blob or "theta" in series_blob
    has_omega = "ω" in series_blob or "omega" in series_blob
    has_alpha = "α" in series_blob or "alpha" in series_blob

    hits = sum([has_theta, has_omega, has_alpha])

    total += 0.25 * (hits / 3)

    if hits < 3:
        notes.append("Graph: the legend/series labels should identify θ, ω, and α.")

    return min(2.0, total), notes

def grade_rotation_lab(path, rubric):
    wb_f = _grading_workbook(path, data_only=False)
    wb_v = _grading_workbook(path, data_only=True)

    total = 0.0
    notes = []

    expected_time_rows = rubric.get("expected_time_rows", 50)

    # ---------- Summary tables: 6 pts ----------
    summary_scores = []

    for ws_name in wb_f.sheetnames:
        ws_f = wb_f[ws_name]
        ws_v = wb_v[ws_name]

        header_rows = _find_rotation_summary_headers(ws_f)

        for hr in header_rows:
            p, fb = _grade_one_rotation_summary_table(ws_f, ws_v, hr)
            summary_scores.append((p, fb))

    summary_scores = sorted(summary_scores, key=lambda x: x[0], reverse=True)[:2]

    total += sum(p for p, fb in summary_scores)

    if len(summary_scores) < 2:
        notes.append(f"Rotation summary tables: I found {len(summary_scores)}/2 required tables.")

    for p, fb in summary_scores:
        if p < 2.99:
            notes.extend(fb)

    # ---------- Time table: 2 pts ----------
    time_scores = []

    for ws_name in wb_f.sheetnames:
        ws_f = wb_f[ws_name]
        ws_v = wb_v[ws_name]

        header_rows = _find_rotation_time_table_headers(ws_f)

        for hr in header_rows:
            p, fb = _grade_rotation_time_table(ws_f, ws_v, hr, expected_rows=expected_time_rows)
            time_scores.append((p, fb))

    time_scores = sorted(time_scores, key=lambda x: x[0], reverse=True)[:1]

    if time_scores:
        total += time_scores[0][0]
        if time_scores[0][0] < 1.99:
            notes.extend(time_scores[0][1])
    else:
        notes.append("Rotation time table: I could not find the orbiting-masses table.")

    # ---------- Graph: 2 pts ----------
    graph_pts, graph_fb = _grade_rotation_graphs(
        wb_f,
        expected_graphs=rubric.get("expected_graphs", 1)
    )

    total += graph_pts
    notes.extend(graph_fb)

    total = round(min(10.0, total), 2)

    if total >= 9.995:
        return total, ["Everything needed was found: summary tables, time table, formulas, and graph. ✅"]

    # remove duplicate feedback lines
    deduped = []
    seen = set()
    for n in notes:
        key = n.lower().strip()
        if key not in seen:
            seen.add(key)
            deduped.append(n)

    notes = deduped

    return total, notes

def _spring_clean_text(txt):
    t = norm(txt).lower()
    t = t.replace(" ", "")
    t = t.replace("_", "")
    t = t.replace("-", "")
    return t

def _spring_text_has_force(txt):
    t = _spring_clean_text(txt)
    return (
        "force" in t
        or "forcen" in t
        or "newton" in t
        or "n)" in t
        or t.endswith("n")
    )

def _spring_text_has_extension(txt):
    t = _spring_clean_text(txt)
    return (
        "extension" in t
        or "extensionm" in t
        or "displacement" in t
        or "displacementm" in t
        or "stretch" in t
        or "x(m" in t
        or "xm" in t
    )

def _spring_text_has_mass(txt):
    t = _spring_clean_text(txt)
    return (
        "mass" in t
        or "masskg" in t
        or "mkg" in t
        or "kg" in t
    )

def _find_spring_table_headers(ws_f, scan_rows=120, scan_cols=80):
    """
    Detect spring tables by anchoring on the Trial column.
    Each table is assumed to be 4 columns:
    trial, mass, extension/displacement, force.
    """
    candidates = []

    max_r = min(ws_f.max_row, scan_rows)
    max_c = min(ws_f.max_column, scan_cols)

    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            txt = norm(ws_f.cell(r, c).value).lower()

            if "trial" not in txt:
                continue

            local_blob = " ".join(
                norm(ws_f.cell(r, cc).value).lower()
                for cc in range(c, min(c + 4, max_c + 1))
                if ws_f.cell(r, cc).value is not None
            )

            has_trial = "trial" in local_blob
            has_mass = _spring_text_has_mass(local_blob)
            has_extension = _spring_text_has_extension(local_blob)
            has_force = _spring_text_has_force(local_blob)

            core_hits = sum([has_trial, has_mass, has_extension, has_force])

            if core_hits >= 2:
                candidates.append((r, c, min(c + 3, max_c)))

    return candidates

def _grade_one_spring_table(ws_f, ws_v, table_candidate, expected_rows=20):
    header_row, c_start, c_end = table_candidate

    pts = 0.0
    notes = []

    row_blob = " ".join(
        norm(ws_f.cell(header_row, c).value)
        for c in range(c_start, c_end + 1)
        if ws_f.cell(header_row, c).value is not None
    ).lower()

    # ---------- 0.5 pts — title ----------
    has_title = _nearby_table_title(
        ws_f, header_row, c_start, c_end,
        lambda text: "spring" in text or "mystery" in text,
    )

    if has_title:
        pts += 0.5
    else:
        notes.append("Spring table: include a clear table title such as Mystery Spring A, B, or C.")

    # ---------- 0.5 pts — headers/units split across columns ----------
    header_pts = 0.0

    has_trial = "trial" in row_blob
    has_mass = _spring_text_has_mass(row_blob)
    has_extension = _spring_text_has_extension(row_blob)
    has_force = _spring_text_has_force(row_blob)

    if has_trial:
        header_pts += 0.125
    else:
        notes.append("Spring table: include a Trial # column.")

    if has_mass:
        header_pts += 0.125
    else:
        notes.append("Spring table: include a mass column with units.")

    if has_extension:
        header_pts += 0.125
    else:
        notes.append("Spring table: include an extension/displacement column with units.")

    if has_force:
        header_pts += 0.125
    else:
        notes.append("Spring table: include a force column with units.")

    pts += header_pts

    # ---------- 1.0 pts — usable numeric rows ----------
    numeric_rows = 0

    for r in range(header_row + 1, header_row + 1 + expected_rows):
        numeric_count = 0
        for c in range(c_start, c_end + 1):
            if try_float(ws_v.cell(r, c).value) is not None:
                numeric_count += 1

        if numeric_count >= 2:
            numeric_rows += 1

    if numeric_rows >= 18:
        pts += 1.0
    elif numeric_rows >= 12:
        pts += 0.7
        notes.append(f"Spring table: I found about {numeric_rows}/{expected_rows} usable rows.")
    elif numeric_rows >= 6:
        pts += 0.4
        notes.append(f"Spring table: I found only about {numeric_rows}/{expected_rows} usable rows.")
    else:
        notes.append("Spring table: not enough usable numeric data rows.")

    return min(2.0, pts), notes

def _chart_has_trendline_or_equation(chart):
    """
    Trendline metadata can be unreliable depending on Excel/Google Sheets export.
    This checks for common trendline/equation signs in the chart object text.
    """
    blob = str(chart).lower()

    has_trendline = (
        "trendline" in blob
        or "trendline" in blob.replace(" ", "")
        or "trend" in blob
    )

    has_equation = (
        "dispEq" in str(chart)
        or "dispeq" in blob
        or "equation" in blob
        or "y =" in blob
        or "y=" in blob
    )

    return has_trendline or has_equation

def _grade_spring_graphs(wb_f, expected_graphs=3):
    total = 0.0
    notes = []

    charts = _collect_charts(wb_f)
    charts_to_check = charts[:expected_graphs]

    # ---------- 3 pts — graph quality, 1 pt each ----------
    graph_quality_pts = 0.0

    if len(charts) < expected_graphs:
        notes.append(f"Graphs: I found {len(charts)}/{expected_graphs} required graphs.")

    for i in range(expected_graphs):
        if i >= len(charts_to_check):
            continue

        ch = charts_to_check[i]

        one_graph_pts = 0.0

        # graph exists
        one_graph_pts += 0.25

        # title
        title_exists = _chart_title_exists(ch)
        title_text = _text_from_chart_obj(getattr(ch, "title", None)).lower()
        default_title = "chart title" in title_text

        if title_exists and not default_title:
            one_graph_pts += 0.25
        elif title_exists and default_title:
            one_graph_pts += 0.10
            notes.append(
                'Graphs: replace the default "Chart Title" with a descriptive title for each spring graph.'
            )
        else:
            notes.append("Graphs: each spring graph should have a descriptive title.")

        # axes/units
        x = _axis_title_text(getattr(ch, "x_axis", None))
        y = _axis_title_text(getattr(ch, "y_axis", None))

        x_blob = str(x).lower()
        y_blob = str(y).lower()

        axis_blob = " ".join([x_blob, y_blob, str(ch).lower()])

        has_force_axis = "force" in axis_blob or "(n)" in axis_blob or "newton" in axis_blob
        has_extension_axis = (
            "extension" in axis_blob
            or "displacement" in axis_blob
            or "stretch" in axis_blob
            or "(m)" in axis_blob
            or "meter" in axis_blob
        )

        if has_force_axis and has_extension_axis:
            one_graph_pts += 0.50
        elif has_force_axis or has_extension_axis:
            one_graph_pts += 0.25
            notes.append("Graphs: axes should include both force and extension/displacement units.")
        else:
            notes.append("Graphs: axes should be labeled with units, including force (N) and extension/displacement (m).")

        graph_quality_pts += min(1.0, one_graph_pts)

    total += min(3.0, graph_quality_pts)

    # ---------- 1 pt — trendlines/equations, about 0.33 each ----------
    trend_pts = 0.0

    for ch in charts_to_check:
        if _chart_has_trendline_or_equation(ch):
            trend_pts += 1.0 / expected_graphs

    total += min(1.0, trend_pts)

    if trend_pts < 0.99:
        notes.append("Graphs: each spring graph should include a linear trendline with the equation displayed.")

    return min(4.0, total), notes

def grade_spring_lab(path, rubric):
    wb_f = _grading_workbook(path, data_only=False)
    wb_v = _grading_workbook(path, data_only=True)

    total = 0.0
    notes = []

    expected_rows = rubric.get("expected_rows", 20)
    expected_tables = rubric.get("expected_tables", 3)
    expected_graphs = rubric.get("expected_graphs", 3)

    # ---------- Tables: 6 pts ----------
    table_scores = []

    for ws_name in wb_f.sheetnames:
        ws_f = wb_f[ws_name]
        ws_v = wb_v[ws_name]

        table_candidates = _find_spring_table_headers(ws_f)

        for candidate in table_candidates:
            p, fb = _grade_one_spring_table(ws_f, ws_v, candidate, expected_rows=expected_rows)
            table_scores.append((p, fb))

    table_scores = sorted(table_scores, key=lambda x: x[0], reverse=True)[:expected_tables]

    total += sum(p for p, fb in table_scores)

    if len(table_scores) < expected_tables:
        notes.append(f"Spring tables: I found {len(table_scores)}/{expected_tables} required data tables.")

    for p, fb in table_scores:
        if p < 1.99:
            notes.extend(fb)

    # ---------- Graphs: 4 pts ----------
    graph_pts, graph_fb = _grade_spring_graphs(wb_f, expected_graphs=expected_graphs)

    total += graph_pts
    notes.extend(graph_fb)

    total = round(min(10.0, total), 2)

    if total >= 9.995:
        return total, ["Everything needed was found: three spring tables, three graphs, and trendlines/equations. ✅"]

    # remove duplicate feedback lines
    deduped = []
    seen = set()
    for n in notes:
        key = n.lower().strip()
        if key not in seen:
            seen.add(key)
            deduped.append(n)

    notes = deduped

    return total, notes

def grade_file_by_rubric(path, rubric):
    """
    Routes each lab to the correct grading engine.
    Add future labs here.
    """
    lab_type = rubric.get("type")

    if lab_type == "measurement_uncertainty":
        grader = SpreadsheetGrader(
            expected_n=rubric["expected_n"],
            require_formulas=rubric.get("require_formulas_in_calc_rows", True),
        )

        return grader.grade_file_two_columns(
            path,
            rubric["columns"][0],
            rubric["columns"][1],
            debug=False
        )

    elif lab_type == "motion_graphs":
        return grade_lab2_motion_graphs(path, rubric)

    elif lab_type == "projectile_uncertainty":
        return grade_lab3_projectile_uncertainty(path, rubric)

    elif lab_type == "atwood_friction":
        return grade_lab4d_atwood_friction(path, rubric)

    elif lab_type == "collisions":
        return grade_lab6_7_collisions(path, rubric)

    elif lab_type == "rotation":
        return grade_rotation_lab(path, rubric)

    elif lab_type == "spring":
        return grade_spring_lab(path, rubric)

    else:
        raise ValueError(f"Unknown rubric type: {lab_type}")


RUBRIC_LABELS = {
    'lab1_measurements': 'Lab 1 — Measurements and uncertainty',
    'lab2_motion_graphs': 'Lab 2 — Motion graphs',
    'lab3_projectile_uncertainty': 'Lab 3 — Projectile uncertainty',
    'lab4d_atwood_friction': 'Labs 4–5 — Atwood and friction',
    'lab6_7_collisions': 'Labs 6–7 — Collisions',
    'rotation_lab': 'Rotation', 'spring_lab': 'Springs',
}

_workbooks = ContextVar('grading_workbooks', default=None)

_layout_events = ContextVar('grading_layout_events', default=None)

LAYOUT_POLICY = dict(max_columns=2, per_table=0.25, maximum=1.0)

ENGINE_VERSION = '2026-10-02.2'

def _header_like(text, key):
    """Recognized rubric column labels only; never shift arbitrary worksheet text."""
    text = norm(text)
    if not text or text.startswith('='): return False
    if key == 'lab1_measurements':
        return any(column['units'] in text and any(k in text for k in column['keywords'])
                   for column in LABS[key]['columns'])
    patterns = {
        'lab2_motion_graphs': r'\b(time|position|velocity|acceleration|pos|vel|accel)\b|^t\s*\(',
        'lab3_projectile_uncertainty': r'\b(angle|distance|dist|degree|degrees)\b',
        'lab4d_atwood_friction': r'\b(mass\s*[12]|acceleration|accel|mu)\b|μ',
        'lab6_7_collisions': r'\b(time|v[12]|velocity|ke|momentum|conserved|conservation|p\s+total)\b|^t\s*\(',
        'rotation_lab': r'\b(trial|mass|radius|time|theta|omega|alpha|i[_ ]?(exp|theo))\b|[θωα]|^t\s*\(',
        'spring_lab': r'\b(trial|mass|extension|displacement|force)\b',
    }
    return bool(re.search(patterns.get(key, r'(?!)'), text))

def _numeric_below(values, row, col):
    count = 0
    for rr in range(row+1, min(values.max_row, row+10)+1):
        value = values.cell(rr,col).value
        number = try_float(value)
        if number is not None and math.isfinite(number): count += 1
        elif value is not None: break
    return count

def _merged_span(sheet, row, col):
    return next((area for area in sheet.merged_cells.ranges
                 if area.min_row <= row <= area.max_row and area.min_col <= col <= area.max_col), None)

def _data_block(values, row, col):
    left = right = col
    while left>1 and _numeric_below(values,row,left-1)>=3: left-=1
    while right<min(values.max_column,120) and _numeric_below(values,row,right+1)>=3: right+=1
    return left,right

def _placement_event(sheet, header_row, bounds, source, target, label):
    events = _layout_events.get()
    if events is None: return
    key=(sheet.title,header_row,*bounds)
    for existing in events:
        if existing[:2]==key[:2] and max(existing[2],key[2])<=min(existing[3],key[3]):
            key=existing; break
    event = events.setdefault(key,[])
    description = f'{label}: {sheet.title}!{source} recognized with data at {target}'
    if description not in event: event.append(description)

def _align_nearby_headers(formula_book, value_book, key):
    """Create an in-memory grading view; never save/modify student attachments.

    An empty column under a recognized heading may bind to a unique, unlabelled
    numeric column within two columns. Already-labelled data columns are reserved.
    """
    limit=LAYOUT_POLICY['max_columns']
    for sheet in formula_book:
        values=value_book[sheet.title]
        for row in range(1,min(sheet.max_row,250)+1):
            headers=[col for col in range(1,min(sheet.max_column,120)+1)
                     if _header_like(sheet.cell(row,col).value,key)]
            reserved={col for col in headers if _numeric_below(values,row,col)>0}
            moves=[]; destinations=set()
            for source in headers:
                if source in reserved: continue
                area=_merged_span(sheet,row,source)
                candidates=[]
                for target in range(max(1,source-limit),min(values.max_column,source+limit)+1):
                    if target==source or target in reserved: continue
                    if sheet.cell(row,target).value is not None: continue
                    if _numeric_below(values,row,target)<3: continue
                    target_area=_merged_span(sheet,row,target)
                    if target_area is not None and target_area!=area: continue
                    candidates.append(target)
                if not candidates: continue
                # A real merged header identifies its span; coloring does not.
                covered=[col for col in candidates if area and area.min_col<=col<=area.max_col]
                if covered: candidates=covered
                distance=min(abs(col-source) for col in candidates)
                nearest=[col for col in candidates if abs(col-source)==distance]
                if len(nearest)!=1 or nearest[0] in destinations:
                    raise ValueError(f'Ambiguous nearby data for {sheet.title}!{sheet.cell(row,source).coordinate} — manual review required.')
                target=nearest[0]; destinations.add(target)
                moves.append((source,target,sheet.cell(row,source).value,area))
            for source,target,label,area in moves:
                source_cell=sheet.cell(row,source).coordinate
                target_cell=sheet.cell(row,target).coordinate
                covered=bool(area and area.min_col<=target<=area.max_col)
                if area: sheet.unmerge_cells(str(area))
                sheet.cell(row,source).value=None
                sheet.cell(row,target).value=label
                if not covered:
                    _placement_event(sheet,row,_data_block(values,row,target),source_cell,target_cell,str(label))

def _nearby_table_title(sheet, header_row, start_col, end_col, predicate, include_header=True, table_blocks=None):
    """Shared bounded title lookup; merged spans overlapping the table are aligned."""
    candidates=[]
    for row in range(max(1,header_row-3),header_row+(1 if include_header else 0)):
        for col in range(max(1,start_col-LAYOUT_POLICY['max_columns']),
                         min(sheet.max_column,end_col+LAYOUT_POLICY['max_columns'])+1):
            text=sheet.cell(row,col).value
            if not predicate(norm(text)): continue
            area=_merged_span(sheet,row,col)
            left,right=(area.min_col,area.max_col) if area else (col,col)
            distance=max(start_col-right,left-end_col,0)
            if table_blocks:
                distances=[max(a-right,left-b,0) for a,b in table_blocks]
                if distance>min(distances): continue  # This title belongs closer to another table.
                if distance and distances.count(distance)>1:
                    raise ValueError(f'Ambiguous table title at {sheet.title}!{sheet.cell(row,col).coordinate} — manual review required.')
            candidates.append((distance,header_row-row,row,col,text))
    if not candidates: return False
    distance,_,row,col,text=min(candidates,key=lambda item:item[:2])
    if distance and len({norm(item[4]) for item in candidates if item[:2]==(distance,header_row-row)})>1:
        raise ValueError(f'Multiple equally close titles for {sheet.title} row {header_row} — manual review required.')
    if distance:
        from openpyxl.utils import get_column_letter
        target=f'{get_column_letter(start_col)}{header_row}:{get_column_letter(end_col)}{header_row}'
        _placement_event(sheet,header_row,(start_col,end_col),sheet.cell(row,col).coordinate,target,str(text))
    return True

def _check_nearby_table_titles(formula_book,value_book,key):
    # Rubrics without a separate title criterion still report an identifiable
    # displaced title. Missing titles retain each original rubric's own rules.
    patterns={
        'lab2_motion_graphs':r'\b(constant(?:ly)?|accelerating)\b.*\b(car|motion|table)\b',
        'lab3_projectile_uncertainty':r'\b(projectile|random error|systematic error)\b',
        'lab4d_atwood_friction':r'\b(friction|atwood)\b',
        'lab6_7_collisions':r'\b(elastic|inelastic|collision)\b',
        'rotation_lab':r'\b(changing radius|changing mass|orbiting masses)\b',
        'spring_lab':r'\b(spring|mystery)\b',
    }
    if key not in patterns: return
    predicate=lambda text:bool(re.search(patterns[key],text))
    for sheet in formula_book:
        values=value_book[sheet.title]
        for row in range(1,min(sheet.max_row,250)+1):
            cols=[col for col in range(1,min(sheet.max_column,120)+1)
                  if _header_like(sheet.cell(row,col).value,key) and _numeric_below(values,row,col)>=3]
            blocks={_data_block(values,row,col) for col in cols}
            for left,right in blocks:
                if sum(left<=col<=right for col in cols)>=2:
                    _nearby_table_title(sheet,row,left,right,predicate,include_header=False,table_blocks=blocks)

def _layout_feedback(score,feedback,events):
    deductions=[];total=0.0
    for descriptions in events.values():
        amount=min(LAYOUT_POLICY['per_table'],LAYOUT_POLICY['maximum']-total,score-total)
        if amount<=0: break
        total+=amount
        deductions.append(f'Table/header placement: -{amount:g} point. '+ '; '.join(descriptions)+
                          '. Align the title/header above its data, or merge the title across the table.')
    if total:
        feedback=[line for line in feedback if not str(line).startswith('Everything needed was found')]
        if not feedback: feedback=['Required work recognized and checked against the lab rubric.']
        feedback.extend(deductions)
    return round(score-total,2),feedback

def _grading_workbook(path, data_only=False):
    pair = _workbooks.get()
    if pair is None or str(Path(path).resolve()) != pair[0]:
        raise RuntimeError('Use score_workbook() to manage workbook resources.')
    return pair[2 if data_only else 1]

def score_workbook(path, key):
    """Load formulas and saved values once; close both even if a grader fails."""
    formula_book = value_book = None
    token = layout_token = None
    try:
        formula_book = openpyxl.load_workbook(path, data_only=False)
        value_book = openpyxl.load_workbook(path, data_only=True)
        populated, missing, errors = 0, [], []
        for sheet in formula_book:
            if sheet.max_row * sheet.max_column > 2_000_000:
                raise ValueError('Very large worksheet dimensions — inspect this workbook manually.')
            values = value_book[sheet.title]
            for cells in sheet.iter_rows():
                for cell in cells:
                    populated += int(cell.value is not None)
                    if cell.data_type == 'f':
                        saved = values[cell.coordinate]
                        if saved.value is None:
                            missing.append(f'{sheet.title}!{cell.coordinate}')
                        elif saved.data_type == 'e':
                            errors.append(f'{sheet.title}!{cell.coordinate}')
        if not populated:
            raise ValueError('Empty workbook — manual review required.')
        if missing or errors:
            detail = ', '.join((missing + errors)[:6])
            raise ValueError(f'Formula results unavailable or errors in {len(missing)+len(errors)} cells ({detail}). Open in Excel/LibreOffice, recalculate, save, and review manually. No automatic grade proposed.')
        events={}
        layout_token=_layout_events.set(events)
        _align_nearby_headers(formula_book,value_book,key)
        _check_nearby_table_titles(formula_book,value_book,key)
        token = _workbooks.set((str(Path(path).resolve()), formula_book, value_book))
        with contextlib.redirect_stdout(io.StringIO()):
            score, feedback = grade_file_by_rubric(path, LABS[key])
        score = float(score)
        if not math.isfinite(score) or not 0 <= score <= 10:
            raise ValueError('Rubric returned an invalid score; manual review required.')
        return _layout_feedback(score,list(feedback),events)
    finally:
        if token is not None:
            _workbooks.reset(token)
        if layout_token is not None:
            _layout_events.reset(layout_token)
        for book in (formula_book, value_book):
            if book is not None:
                book.close()
