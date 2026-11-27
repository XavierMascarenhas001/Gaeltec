"""Tool 3 - Target Price -> Control File (writes Auxiliar / Block1 / PA CONTROL).

Options in the web page:
  * sheets          - "Pole Materials" always, "OHL" when ticked (as in the desktop script)
  * threshold       - % similarity for two Line names to be flagged (default 90)
  * seg_threshold   - % word/phrase similarity used to drop the Scheme Title / Section
                      from a segment when they repeat the Line (default 90)
  * pm              - Project Manager for the segment names (instead of Target!D21)
  * master + sources- ENIDs of this Target Price already in those Control Files of the
                      Master go to a new sheet "Enids already in CF" (only if any), item by
                      item: added (green), removed (red), quantity changes, and the action to
                      take in that Control File to match the Target Price

Original logic kept verbatim. The tkinter pieces were replaced:
  * the two askopenfilename dialogs  -> `control_path`, `tp_path` arguments
  * the "Choose Line" pop-up         -> a two-step flow in the web page:
        1) analyze()  returns every similar-line pair found
        2) run()      receives the user's choice for each pair
           ("1", "2" or "keep_both" - same three answers as before,
            anything else falls back to "1" exactly like the old dialog)
"""
import pandas as pd
import re
from rapidfuzz import fuzz
from itertools import combinations
import difflib
import os
import shutil
from openpyxl import load_workbook

CONTROL_FILETYPES = [("Excel files", "*.xlsx *.xlsm")]
TP_FILETYPES = [("Excel Macro-Enabled Workbook", "*.xlsm"), ("All files", "*.*")]


# ---------------------------
# --- String cleaning function ---
# ---------------------------
def clean_string(s):
    if pd.isna(s):
        return ""
    s = re.sub(r'\s+', ' ', str(s)).strip()
    s = s.replace(' ', ' ')
    return s.lower()


# ---------------------------
# --- Highlight differences ---
# ---------------------------
def highlight_diff(s1, s2, mode="change"):
    if mode == "keep_both":
        return {"original_1": s1, "original_2": s2}

    seq = difflib.SequenceMatcher(None, s1, s2)
    highlighted = []

    for tag, i1, i2, j1, j2 in seq.get_opcodes():
        if tag == "equal":
            highlighted.append(s1[i1:i2])
        elif tag in ("replace", "delete"):
            highlighted.append(f"[{s1[i1:i2]}]")
        elif tag == "insert":
            highlighted.append(f"(+{s2[j1:j2]})")
    return "".join(highlighted)


# ---------------------------
# --- Detect similar lines ---
# ---------------------------
def detect_similar_lines_unique(strings, threshold=80):
    unique_strings = list(dict.fromkeys(strings))  # Only unique strings
    similar_pairs = []

    for s1, s2 in combinations(unique_strings, 2):
        similarity = fuzz.ratio(s1, s2)
        if similarity >= threshold and similarity < 100:
            diff = highlight_diff(s1, s2)
            similar_pairs.append((s1, s2, similarity, diff))

    return similar_pairs


def choice_to_lines(s1, s2, chosen, log=print):
    """Same rules as the old choose_string_gui dialog."""
    result = [s1]  # default
    if chosen is not None:
        choice = str(chosen).strip().lower()
        if choice == "1":
            result = [s1]
        elif choice == "2":
            result = [s2]
        elif choice == "keep_both":
            result = [s1, s2]
        else:
            log(f"Invalid choice - using first line by default: {s1}")
    return result


# ---------------------------
# --- Build segment string ---
# ---------------------------
SIM_THRESHOLD = 90


def build_segment(row, prefix, pm_value, schemetitle_value, seg_threshold=None):
    thr = SIM_THRESHOLD if seg_threshold is None else seg_threshold
    similarity = fuzz.token_set_ratio(str(schemetitle_value), str(row["Line"]))
    scheme_part = "" if similarity >= thr else f"{schemetitle_value} - "

    line_section_similarity = fuzz.token_set_ratio(str(row["Line"]), str(row["Section"]))
    section_part = "" if line_section_similarity >= thr else f" - {row['Section']}"

    return f"{prefix}{pm_value} - {scheme_part}{row['Circuit']} - {row['Line']}{section_part}"


def _read_tp_sheet(tp_path, sheet):
    usecols = [1, 2, 3, 4, 6, 8, 9, 11, 13]  # B,D,E,G,I,J,L,N
    df = pd.read_excel(tp_path, sheet_name=sheet, header=3, usecols=usecols, engine="openpyxl")
    df.columns = [
        "Circuit", "Section", "Line", "ENID", "Item",
        "Work Description", "Span", "Quantity", "Comment"
    ]
    # Clean "Line"
    df["Line"] = df["Line"].astype(str).apply(clean_string)
    return df


def process_target_sheet(tp_path, sheet, pm_value, schemetitle_value, choices=None,
                         threshold=90, log=print, seg_threshold=None):
    """
    Reads a sheet from Target Price, detects similar lines, applies the
    user's choices, keeps all original rows intact, and updates the Segment
    column correctly. `choices` is a list aligned with the similar pairs
    (index i -> "1" / "2" / "keep_both").
    """
    df = _read_tp_sheet(tp_path, sheet)

    # Detect similar lines
    similar = detect_similar_lines_unique(df["Line"].tolist(), threshold=threshold)
    choices = choices or []

    if similar:
        log(f"⚠️ Similar lines detected in '{sheet}':")
        for i, (s1, s2, score, diff) in enumerate(similar):
            log(f"[{score}% similarity]\n  1: {s1}\n  2: {s2}\n  Diff: {diff}\n")
            chosen = choices[i] if i < len(choices) else "1"
            chosen_lines = choice_to_lines(s1, s2, chosen, log=log)
            log(f"   → choice: {chosen}")

            if len(chosen_lines) == 1:
                # Replace all occurrences of s1 and s2 with the chosen line
                df.loc[df["Line"].isin([s1, s2]), "Line"] = chosen_lines[0]

            # If "keep_both", do nothing, keep all rows as they are
    else:
        log(f"✅ No similar lines detected in '{sheet}'.")

    # Skip rows with "Missing" in 2 or more columns
    missing_counts = df.apply(lambda row: row.astype(str).str.contains("missing", case=False, na=False).sum(), axis=1)
    df = df[missing_counts < 2]

    # Add Segment based on current Line (after possible renaming)
    prefix = "M - " if sheet == "Pole Materials" else "C - "
    df["Segment"] = df.apply(build_segment, axis=1, prefix=prefix, pm_value=pm_value, schemetitle_value=schemetitle_value,
                             seg_threshold=seg_threshold)

    log(f"✅ Processed {len(df)} rows from '{sheet}'")
    return df


def read_target_values(tp_path, log=print):
    # --- Read PM value from Target Price sheet 'Target', cell D21 ---
    try:
        wb_tp = load_workbook(tp_path, data_only=True)
        ws_target = wb_tp["Target"]

        pm_value = ws_target["D21"].value or "Unknown"
        schemetitle_value = ws_target["D4"].value or "Unknown"
        PID = ws_target["D22"].value or ""

        wb_tp.close()
        log(f"🔹 PM value read from Target Price (Target!D21): {pm_value}")
        log(f"🔹 Scheme Title read from Target Price (Target!D4): {schemetitle_value}")

    except Exception as e:
        log(f"⚠️ Could not read values from Target sheet: {e}")
        pm_value = "Unknown"
        schemetitle_value = "Unknown"
        PID = ""
    return pm_value, schemetitle_value, PID


TP_SHEETS = ["Pole Materials"]  # default: only the Pole Materials sheet ("OHL" can be added)


def _sheets(sheets):
    out = [s for s in (sheets or TP_SHEETS) if s in ("Pole Materials", "OHL")]
    return out or list(TP_SHEETS)


def _choices_for(choices, sheet, n):
    """choices may be a flat list (old pages, Pole Materials only) or {sheet: [..]}."""
    if isinstance(choices, dict):
        return list(choices.get(sheet) or [])
    return list(choices or []) if sheet == "Pole Materials" else []


# ---------------------------
# --- ENIDs already in the Master ---
# ---------------------------
_ENID_SKIP = {"", "nan", "none", "0", "missing", "-", "n/a", "na", "tbc", "tba"}


def norm_enid(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    s = re.sub(r"\s+", "", s).upper()
    if s.lower() in _ENID_SKIP:
        return ""
    # only real ENIDs (an 8-digit number) are checked - "New pole", "NEW", "TBC"... are ignored
    return s if re.fullmatch(r"\d{8}", s) else ""


MASTER_COLS = ["sourcefile", "job", "enid", "description", "qty", "project", "project manager",
               "circuit", "district", "done", "plan1"]


def master_info(master_path, log=print):
    """Project Managers and Control Files (source files) found in the Master."""
    cols = ["project manager", "sourcefile"]
    df = pd.read_parquet(master_path, columns=cols)
    pm = df["project manager"].fillna("").astype(str).str.strip()
    pms = pm[pm != ""].value_counts()
    src = df["sourcefile"].fillna("").astype(str).str.strip()
    srcs = src[src != ""].value_counts()

    def _num(name):
        m = re.match(r"\s*(\d+)", name)
        return int(m.group(1)) if m else 10 ** 6
    sources = sorted(({"name": k, "rows": int(v)} for k, v in srcs.items()), key=lambda d: (-_num(d["name"]), d["name"]))
    log(f"Master: {len(pms)} project managers, {len(sources)} Control Files.")
    return {"pms": [{"name": k, "rows": int(v)} for k, v in pms.items()], "sources": sources}


def enids_in_master(tp_df, master_path, sources=None, log=print):
    """Rows of the Master (only the chosen Control Files) whose ENID is also in this Target Price."""
    want = {}
    for _, r in tp_df.iterrows():
        k = norm_enid(r.get("ENID"))
        if k:
            want.setdefault(k, []).append(r)
    if not want:
        return pd.DataFrame(), {}
    try:
        import pyarrow.parquet as pq
        have = set(pq.read_schema(master_path).names)
    except Exception:
        have = set(MASTER_COLS)
    cols = [c for c in MASTER_COLS if c in have]
    m = pd.read_parquet(master_path, columns=cols)
    if sources is not None:
        keep = {str(x).strip() for x in sources}
        m = m[m["sourcefile"].fillna("").astype(str).str.strip().isin(keep)]
    m = m.assign(_k=m["enid"].map(norm_enid))
    m = m[m["_k"].isin(want.keys())]
    per = m.groupby("_k")["sourcefile"].nunique().to_dict() if len(m) else {}
    log(f"🔎 {len(per)} of the {len(want)} ENIDs in this Target Price are already in "
        f"{m['sourcefile'].nunique() if len(m) else 0} Control File(s) of the Master.")
    return m, want


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(",", "").strip() or 0)
    except (TypeError, ValueError):
        return 0.0


def _desc_key(v):
    s = "" if v is None else str(v)
    if s.lower() in ("nan", "none"):
        s = ""
    s = re.sub(r"\s+", " ", s.replace(" ", " ")).strip().lower()
    return re.sub(r"[\s.;,:]+$", "", s)


def _kind_of_job(job):
    j = str(job or "").strip().lower()
    return "M" if j.startswith("m -") or j.startswith("m-") else ("C" if j.startswith("c -") or j.startswith("c-") else "")


def compare_items(hits, tp_df, log=print):
    """For every ENID already in a Control File: what the Target Price adds, removes or changes
    (item = description, M/C kept apart, quantities summed) and the action for the Control File."""
    tp = tp_df.copy()
    tp["_k"] = tp["ENID"].map(norm_enid)
    tp = tp[tp["_k"].isin(set(hits["_k"]))]
    tp["_kind"] = tp["Segment"].astype(str).str[:1].str.upper()
    tp["_d"] = tp["Work Description"].map(_desc_key)
    tp["_q"] = tp["Quantity"].map(_num)
    hits = hits.copy()
    hits["_kind"] = hits["job"].map(_kind_of_job)
    hits["_d"] = hits["description"].map(_desc_key)
    hits["_q"] = hits["qty"].map(_num)
    kinds_read = set(tp_df["Segment"].astype(str).str[:1].str.upper())
    rows = []
    for (k, src), cf in hits.groupby(["_k", "sourcefile"], sort=True):
        cf = cf[cf["_kind"].isin(kinds_read) | (cf["_kind"] == "")]
        t = tp[tp["_k"] == k]
        info = cf.iloc[0] if len(cf) else hits[(hits["_k"] == k) & (hits["sourcefile"] == src)].iloc[0]
        jobs = {kd: g["job"].iloc[0] for kd, g in cf.groupby("_kind")}
        tp_items, cf_items = {}, {}
        for _, r in t.iterrows():
            e = tp_items.setdefault((r["_kind"], r["_d"]), {"q": 0.0, "desc": r["Work Description"], "seg": r["Segment"]})
            e["q"] += r["_q"]
        for _, r in cf.iterrows():
            e = cf_items.setdefault((r["_kind"], r["_d"]), {"q": 0.0, "desc": r["description"], "job": r["job"]})
            e["q"] += r["_q"]
        # descriptions written slightly differently are still the same item
        left_tp = [x for x in tp_items if x not in cf_items]
        left_cf = [x for x in cf_items if x not in tp_items]
        for x in list(left_tp):
            best, score = None, 0
            for y in left_cf:
                if y[0] != x[0]:
                    continue
                sc = fuzz.token_sort_ratio(x[1], y[1])
                if sc > score:
                    best, score = y, sc
            if best is not None and score >= 92:
                cf_items[x] = cf_items.pop(best)
                cf_items[x]["reworded"] = True
                left_cf.remove(best)
        for key in sorted(set(tp_items) | set(cf_items), key=lambda z: (z[0], z[1])):
            a, c = tp_items.get(key), cf_items.get(key)
            kind = key[0]
            job = (c or {}).get("job") or jobs.get(kind) or (a or {}).get("seg") or ""
            desc = (a or c)["desc"]
            q_tp, q_cf = (a or {}).get("q", 0.0), (c or {}).get("q", 0.0)

            def fq(v):
                return int(v) if float(v).is_integer() else round(v, 3)
            if a and not c:
                status, change = "Added", q_tp
                action = f"Add '{desc}' × {fq(q_tp)} to job '{job}' in {src}"
            elif c and not a:
                status, change = "Removed", -q_cf
                action = f"Remove '{desc}' (qty {fq(q_cf)}) from job '{job}' in {src}"
            elif abs(q_tp - q_cf) > 1e-9:
                status = "Quantity up" if q_tp > q_cf else "Quantity down"
                change = q_tp - q_cf
                action = f"Change the quantity of '{desc}' in job '{job}' ({src}) from {fq(q_cf)} to {fq(q_tp)}"
            else:
                status, change, action = "Same", 0.0, "No change"
            if c and c.get("reworded") and status != "Removed":
                action += f"  (Control File says '{c['desc']}')"
            rows.append({"ENID": k, "Control File": src, "Job in Control File": job, "Type": kind,
                         "Description": desc, "Qty in Control File": fq(q_cf) if c else None,
                         "Qty in Target Price": fq(q_tp) if a else None, "Change": fq(change),
                         "Status": status, "Action needed": action,
                         "Project": info.get("project"), "Project Manager": info.get("project manager"),
                         "Circuit": info.get("circuit"), "Plan date": info.get("plan1"), "Done": info.get("done")})
    out = pd.DataFrame(rows)
    if len(out):
        n = out["Status"].value_counts().to_dict()
        log(f"🔁 Changes on poles already in Control Files: {n.get('Added', 0)} to add, {n.get('Removed', 0)} to remove, "
            f"{n.get('Quantity up', 0) + n.get('Quantity down', 0)} quantity change(s), {n.get('Same', 0)} the same.")
    return out


def _write_enid_sheet(wb, cmp, log=print):
    """New sheet 'Enids already in CF' - only when there are repeated ENIDs.
    Green font = added / more, red font = removed / less, with the action for each line."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    name = "Enids already in CF"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = "C0392B"
    head = ["ENID", "Control File", "Job in Control File", "Type", "Description", "Qty in Control File",
            "Qty in Target Price", "Change", "Status", "Action needed (in the Control File)",
            "Project", "Project Manager", "Circuit", "Plan date", "Done"]
    keys = ["ENID", "Control File", "Job in Control File", "Type", "Description", "Qty in Control File",
            "Qty in Target Price", "Change", "Status", "Action needed", "Project", "Project Manager", "Circuit", "Plan date", "Done"]
    n = cmp["Status"].value_counts().to_dict()
    ws.append([f"ENIDs of this Target Price already in other Control Files – {cmp['ENID'].nunique()} ENIDs: "
               f"{n.get('Added', 0)} to add, {n.get('Removed', 0)} to remove, "
               f"{n.get('Quantity up', 0) + n.get('Quantity down', 0)} quantity changes, {n.get('Same', 0)} the same"])
    ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F3B57")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(head))
    ws.append(["Green = added / more in the Target Price    Red = removed / less in the Target Price    Black = the same"])
    ws["A2"].font = Font(italic=True, color="5D6D7E")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(head))
    ws.append(head)
    thin = Side(style="thin", color="C9D3DD")
    for c in range(1, len(head) + 1):
        cell = ws.cell(row=3, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E86C1" if c <= 9 else "C0392B")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = Border(bottom=thin)
    GREEN, RED = "1E8449", "C0392B"
    band = PatternFill("solid", fgColor="F4F7FA")
    order = {"Removed": 0, "Quantity down": 1, "Quantity up": 2, "Added": 3, "Same": 4}
    cmp = cmp.assign(_o=cmp["Status"].map(order)).sort_values(["ENID", "Control File", "_o", "Description"], kind="stable")
    grp_i = {k: i for i, k in enumerate(dict.fromkeys(zip(cmp["ENID"], cmp["Control File"])))}
    row_i = 4
    for _, r in cmp.iterrows():
        st = r["Status"]
        colour = GREEN if st in ("Added", "Quantity up") else RED if st in ("Removed", "Quantity down") else "000000"
        bold = st != "Same"
        for c, key in enumerate(keys, start=1):
            v = r[key]
            try:
                if v is None or pd.isna(v):
                    v = None
            except (TypeError, ValueError):
                pass
            if hasattr(v, "to_pydatetime"):
                v = v.to_pydatetime()
            if key == "ENID" and isinstance(v, str) and v.isdigit():
                v = int(v)
            if key == "Change" and isinstance(v, (int, float)) and v > 0:
                v = f"+{v}"
            cell = ws.cell(row=row_i, column=c, value=v)
            cell.font = Font(color=colour, bold=bold and key in ("Description", "Change", "Status", "Action needed"))
            if hasattr(v, "year"):
                cell.number_format = "dd/mm/yyyy"
            if key == "Action needed":
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if grp_i[(r["ENID"], r["Control File"])] % 2:
                cell.fill = band
        row_i += 1
    widths = [12, 40, 44, 7, 46, 11, 11, 9, 14, 70, 18, 18, 14, 12, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(head))}{max(3, row_i - 1)}"
    log(f"📑 Added sheet '{name}' with {row_i - 4} line(s).")


def analyze(tp_path, threshold=90, log=print, sheets=None, master_path=None, sources=None):
    """Step 1 - returns the header values and every similar-line pair the
    user needs to decide on (replaces the pop-up questions)."""
    pm_value, schemetitle_value, PID = read_target_values(tp_path, log=log)
    pairs = []
    frames = []
    for sheet in _sheets(sheets):
        try:
            df = _read_tp_sheet(tp_path, sheet)
        except Exception as e:
            log(f"⚠️ Error reading '{sheet}': {e}")
            continue
        df["Segment"] = ("M - " if sheet == "Pole Materials" else "C - ")
        frames.append(df)
        for s1, s2, score, diff in detect_similar_lines_unique(df["Line"].tolist(), threshold=threshold):
            pairs.append({"sheet": sheet, "s1": s1, "s2": s2, "similarity": round(float(score), 1), "diff": diff})
    log(f"Found {len(pairs)} similar line pair(s) to review.")
    out = {"pm": str(pm_value), "scheme_title": str(schemetitle_value), "pid": str(PID), "pairs": pairs}
    if master_path and frames:
        allf = pd.concat(frames, ignore_index=True)
        hits, want = enids_in_master(allf, master_path, sources, log=log)
        if len(hits):
            by = (hits.groupby("sourcefile")["_k"].nunique().sort_values(ascending=False))
            out["enid_hits"] = {"enids": int(hits["_k"].nunique()), "lines": int(len(hits)), "tp_enids": len(want),
                                "files": [{"name": str(k), "enids": int(v)} for k, v in by.items()]}
            try:
                cmp = compare_items(hits, allf, log=log)
                st = cmp["Status"].value_counts().to_dict() if len(cmp) else {}
                out["enid_hits"]["changes"] = {"add": int(st.get("Added", 0)), "remove": int(st.get("Removed", 0)),
                                               "qty": int(st.get("Quantity up", 0) + st.get("Quantity down", 0)),
                                               "same": int(st.get("Same", 0))}
            except Exception as e:
                log(f"⚠️ Couldn't compare the items: {e}")
        else:
            out["enid_hits"] = {"enids": 0, "lines": 0, "tp_enids": len(want), "files": []}
    return out


def run(control_path, tp_path, choices=None, output_dir=None, log=print, threshold=90, seg_threshold=None,
        sheets=None, pm=None, master_path=None, sources=None):
    """Step 2 - writes the data into a copy of the Control File."""
    log(f"\n📘 Selected Control file: {control_path}")
    log(f"📗 Selected Target Price file: {tp_path}")

    pm_value, schemetitle_value, PID = read_target_values(tp_path, log=log)
    if pm and str(pm).strip():
        pm_value = str(pm).strip()
        log(f"🔹 Project Manager chosen in the page: {pm_value}")

    tp_data = []
    for sheet in _sheets(sheets):
        try:
            df = process_target_sheet(tp_path, sheet, pm_value, schemetitle_value,
                                      choices=_choices_for(choices, sheet, 0), threshold=threshold,
                                      log=log, seg_threshold=seg_threshold)
            tp_data.append(df)
        except Exception as e:
            log(f"⚠️ Error reading '{sheet}': {e}")

    # Combine sheets
    tp_combined = pd.concat(tp_data, ignore_index=True)

    # --- custom suffix for output filename ---
    suffix = schemetitle_value

    # --- Copy Control file to preserve formulas ---
    base, ext = os.path.splitext(control_path)
    if output_dir:
        base = os.path.join(output_dir, os.path.basename(base))
    updated_path = f"{base}_{suffix}{ext}"
    shutil.copy2(control_path, updated_path)
    log(f"📂 Copied Control file → {updated_path}")

    # --- Open copied Control file ---
    wb = load_workbook(updated_path, keep_vba=True)

    # --- Write to 'Auxiliar' sheet ---
    ws_aux = wb["Auxiliar"]
    first_empty_aux = next((row for row in range(4, ws_aux.max_row + 2) if ws_aux.cell(row=row, column=1).value in (None, "")), ws_aux.max_row + 1)

    for i, (_, row) in enumerate(tp_combined.iterrows()):
        ws_aux.cell(row=first_empty_aux + i, column=1, value=row["Segment"])
        ws_aux.cell(row=first_empty_aux + i, column=2, value=row["ENID"])
        ws_aux.cell(row=first_empty_aux + i, column=3, value=row["Work Description"])
        ws_aux.cell(row=first_empty_aux + i, column=4, value=row["Quantity"])
        ws_aux.cell(row=first_empty_aux + i, column=5, value=row["Span"])
        ws_aux.cell(row=first_empty_aux + i, column=6, value=row["Item"])
        ws_aux.cell(row=first_empty_aux + i, column=7, value=row["Comment"])

    # --- Write to 'Block1' sheet ---
    block_segments = []
    if "Block1" in wb.sheetnames:
        ws_block = wb["Block1"]
        first_empty_block = next((row for row in range(30, ws_block.max_row + 2) if ws_block.cell(row=row, column=1).value in (None, "")), ws_block.max_row + 1)

        for i, (_, row) in enumerate(tp_combined.iterrows()):
            target_row = first_empty_block + i
            ws_block.cell(row=target_row, column=1, value=row["Segment"])
            ws_block.cell(row=target_row, column=2, value=row["ENID"])
            ws_block.cell(row=target_row, column=3, value=row["Work Description"])
            ws_block.cell(row=target_row, column=4, value=row["Quantity"])
            ws_block.cell(row=target_row, column=6, value=row["Quantity"])
            ws_block.cell(row=target_row, column=20, value=row["Span"])
            ws_block.cell(row=target_row, column=21, value=row["Comment"])
            ws_block.cell(row=target_row, column=91, value=PID)
            block_segments.append(row["Segment"])
    else:
        log("⚠️ Sheet 'Block1' not found in Control file!")

    # --- Write to 'PA CONTROL' sheet ---
    if "PA CONTROL" in wb.sheetnames:
        ws_pa = wb["PA CONTROL"]
        first_empty_pa = next((row for row in range(3, ws_pa.max_row + 2) if ws_pa.cell(row=row, column=1).value in (None, "")), ws_pa.max_row + 1)

        unique_segments = sorted(set(s for s in block_segments if s and str(s).strip()))
        for i, seg in enumerate(unique_segments):
            target_row = first_empty_pa + i
            ws_pa.cell(row=target_row, column=1, value=seg)
            ws_pa.cell(row=target_row, column=2, value="MAT" if str(seg).startswith("M") else "CONST")
    else:
        log("⚠️ Sheet 'PA CONTROL' not found in Control file!")

    # --- ENIDs already in other Control Files (Master) ---
    if master_path:
        try:
            hits, want = enids_in_master(tp_combined, master_path, sources, log=log)
            if len(hits):
                _write_enid_sheet(wb, compare_items(hits, tp_combined, log=log), log=log)
            else:
                log("✅ None of the ENIDs are in the chosen Control Files - no extra sheet added.")
        except Exception as e:
            log(f"⚠️ Couldn't check the Master for ENIDs: {e}")

    # --- Save workbook ---
    wb.save(updated_path)
    wb.close()
    log(f"💾 Data successfully transferred to '{updated_path}' with PM='{pm_value}' and all formulas intact!")
    return [updated_path]
