"""
engine.py
---------
Core data logic - loading, filtering, materials matching (with the
qty_per_unit fix), and PDF/pole-vocabulary matching for maps. Extracted
unchanged (same functions, same behavior) from the existing desktop tool
so the Streamlit dashboard sits on top of exactly the same proven logic,
not a re-implementation.
"""
from __future__ import annotations

import os
import re
import threading
import concurrent.futures
from pathlib import Path
from datetime import date, datetime

import pandas as pd

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

DATE_FIELDS = {
    "Plan1": "plan1",
    "Done": "done",
    "DateToUse": "datetouse",
}

# ---------------------------------------------------------------------------
# Default network locations
# ---------------------------------------------------------------------------
MASTER_CONTROL_FOLDER = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker"
DEFAULT_MATERIALS_MASTER = (
    r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker"
    r"\Programs\Program files\materials_all.parquet"
)

# "Master_DD-MM-YYYY (vN)..." - matched at the start of the filename so it
# doesn't matter whether there's a ".parquet" extension or something else
# after the version number. Version suffix is OPTIONAL - a plain
# "Master_01-09-2026.parquet" with no "(vN)" is still a valid file
# (treated as v0), not something to silently skip.
_MASTER_FILE_REGEX = re.compile(r"^Master_(\d{2})-(\d{2})-(\d{4})(?:\s*\(v(\d+)\))?", re.IGNORECASE)


def find_all_master_control_files(folder: str) -> list[tuple[datetime, int, Path]]:
    """Scans `folder` (not recursive) for files named like
    "Master_DD-MM-YYYY[ (vN)]..." and returns a list of (date, version,
    Path) tuples sorted newest-first - date takes priority over version,
    so a same-day higher version only breaks ties, never overrides a
    genuinely later date. Returns [] if the folder isn't accessible or
    nothing matches the naming convention."""
    folder_path = Path(folder)
    if not folder_path.is_dir():
        return []

    candidates = []
    try:
        entries = list(folder_path.iterdir())
    except OSError:
        return []

    for entry in entries:
        if not entry.is_file():
            continue
        m = _MASTER_FILE_REGEX.match(entry.name)
        if not m:
            continue
        day, month, year, version = m.groups()
        try:
            file_date = datetime(int(year), int(month), int(day))
        except ValueError:
            continue  # not a real calendar date - skip rather than crash
        candidates.append((file_date, int(version or 0), entry))

    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)
    return candidates


def find_latest_master_control_file(folder: str) -> str | None:
    """Convenience wrapper: just the single newest path (as a str), or None."""
    candidates = find_all_master_control_files(folder)
    return str(candidates[0][2]) if candidates else None


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def normalize(text) -> str | None:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return None
    return " ".join(str(text).strip().lower().split())


def load_materials_master(path: str | Path):
    df = pd.read_parquet(path)

    def split(table_name):
        sub = df[df["table"] == table_name].copy()
        sub = sub.dropna(axis=1, how="all")
        return sub.reset_index(drop=True)

    poles = split("poles")
    free_issue = split("free_issue")
    guk_items = split("guk_items")
    guk_subitems = split("guk_subitems")

    poles["_norm"] = poles["description"].map(normalize)
    free_issue["_norm"] = free_issue["description"].map(normalize)
    guk_items["_norm"] = guk_items["description"].map(normalize)

    aliases = df[df["table"] == "aliases"][["item_code", "sub_code", "sub_description"]].copy()
    aliases = aliases.rename(columns={"sub_code": "target_table", "sub_description": "alias_description"})
    aliases["_norm"] = aliases["alias_description"].map(normalize)
    aliases = aliases.reset_index(drop=True)

    return poles, free_issue, guk_items, guk_subitems, aliases


def load_control_file(path: str | Path) -> pd.DataFrame:
    df = pd.read_parquet(path)
    for col in ("district", "project", "circuit", "cat"):
        if col in df.columns:
            df[col] = df[col].astype("string")
    return df


# ---------------------------------------------------------------------------
# PID summary (for the export - which PIDs are in scope, and their
# District/Project/Project Manager/Circuit context)
# ---------------------------------------------------------------------------
def build_pid_summary(filtered: pd.DataFrame) -> pd.DataFrame:
    cols = ["pid", "district", "project", "project manager", "circuit"]
    df = filtered[cols].copy()
    df["pid"] = df["pid"].astype(str).str.strip()
    df = df[df["pid"].notna() & ~df["pid"].isin(["", "nan", "None"])]
    for c in ("district", "project", "project manager", "circuit"):
        df[c] = df[c].astype(str).str.strip()
        df.loc[df[c].isin(["", "nan", "None"]), c] = None
    df = df.drop_duplicates().sort_values(["pid", "district", "project", "circuit"])
    return df.rename(columns={
        "pid": "PID", "district": "District", "project": "Project",
        "project manager": "Project Manager", "circuit": "Circuit",
    }).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Filter option discovery
# ---------------------------------------------------------------------------
def get_filter_options(df: pd.DataFrame) -> dict:
    def opts(col):
        vals = df[col].dropna().astype(str).str.strip()
        vals = sorted({v for v in vals if v and v.lower() != "nan"})
        return vals

    return {
        "district": opts("district"),
        "project": opts("project"),
        "circuit": opts("circuit"),
        "pid": opts("pid"),
        "enid": opts("enid"),
        "sourcefile": opts("sourcefile"),
    }


# ---------------------------------------------------------------------------
# Filtering
# ---------------------------------------------------------------------------
def filter_control(
    df: pd.DataFrame,
    districts: list[str] | None = None,
    projects: list[str] | None = None,
    circuits: list[str] | None = None,
    pids: list[str] | None = None,
    enids: list[str] | None = None,
    sourcefiles: list[str] | None = None,
    date_field: str = "datetouse",
    date_from: date | datetime | str | None = None,
    date_to: date | datetime | str | None = None,
) -> pd.DataFrame:
    if date_field not in df.columns:
        raise ValueError(f"Unknown date field: {date_field}")

    mask = pd.Series(True, index=df.index)
    if districts:
        mask &= df["district"].str.strip().isin(districts)
    if projects:
        mask &= df["project"].str.strip().isin(projects)
    if circuits:
        mask &= df["circuit"].str.strip().isin(circuits)
    if pids:
        mask &= df["pid"].astype(str).str.strip().isin(pids)
    if enids:
        mask &= df["enid"].astype(str).str.strip().isin(enids)
    if sourcefiles:
        mask &= df["sourcefile"].astype(str).str.strip().isin(sourcefiles)

    dates = df[date_field]
    if date_from:
        mask &= dates >= pd.Timestamp(date_from)
    if date_to:
        mask &= dates < (pd.Timestamp(date_to) + pd.Timedelta(days=1))

    return df.loc[mask].copy()


# ---------------------------------------------------------------------------
# Matching + aggregation (same logic/fix as the desktop tool)
# ---------------------------------------------------------------------------
GROUP = ["category", "material_item", "description", "code"]
GROUP_FREE_ISSUE = ["category", "material_item", "description", "unit", "code", "pid"]
GROUP_GUK = ["category", "material_item", "description", "unit", "code"]


def _first_non_null(s: pd.Series):
    """Reducer for groupby().agg() - returns the first non-null,
    non-blank value in the group, or None if every value is missing.
    Used for the control file's own "type" (unit) column, which is
    expected to be consistent per material rather than something to
    split rows on the way PID is."""
    for v in s:
        if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip() not in ("", "nan", "None"):
            return v
    return None


def _prepare_mat_rows(filtered: pd.DataFrame) -> pd.DataFrame:
    mat = filtered[filtered["cat"].str.upper() == "MAT"].copy()
    mat["_norm"] = mat["description"].map(normalize)
    mat["_qty"] = pd.to_numeric(mat["qsub"], errors="coerce").fillna(0.0)
    mat["pid"] = mat["pid"].astype(str).str.strip()
    mat.loc[mat["pid"].isin(["", "nan", "None"]), "pid"] = None
    # The control file's own "type" column (when present) records the
    # material's unit (e.g. "Metres", "Km", "Each") directly, independent
    # of whatever the materials master itself says. Not every control
    # file has this column, so this degrades to None rather than raising
    # if it's missing.
    if "type" in filtered.columns:
        mat["ctrl_type"] = filtered["type"].astype(str).str.strip()
        mat.loc[mat["ctrl_type"].isin(["", "nan", "None"]), "ctrl_type"] = None
    else:
        mat["ctrl_type"] = None
    return mat


def _apply_aliases(mat: pd.DataFrame, aliases: pd.DataFrame, target_table: str, extra_cols=None) -> pd.DataFrame:
    extra_cols = extra_cols or []
    cols = ["_qty"] + extra_cols
    sub = aliases[aliases["target_table"] == target_table]
    if sub.empty:
        empty = mat.iloc[0:0][cols].copy()
        empty["item_code"] = pd.Series(dtype="object")
        return empty
    hit = mat.merge(sub[["_norm", "item_code"]], on="_norm", how="inner")
    return hit[["item_code"] + cols]


def build_breakdown(
    filtered_control: pd.DataFrame,
    poles: pd.DataFrame,
    free_issue: pd.DataFrame,
    guk_items: pd.DataFrame,
    guk_subitems: pd.DataFrame,
    aliases: pd.DataFrame,
):
    """Returns poles_out, free_issue_out, guk_out, guk_sub_out (each fully
    summed - no date/PID/PO breakdown), and unmatched_out for QA."""
    mat = _prepare_mat_rows(filtered_control)
    mat_ctx = mat.drop(columns=["description"])

    # ---- Poles ------------------------------------------------------------
    p_match = mat_ctx.merge(
        poles[["_norm", "description", "short_code", "pole_spec", "treatment"]],
        on="_norm", how="inner",
    )
    poles_out = (
        p_match.assign(
            category=p_match["treatment"].fillna(p_match["pole_spec"]),
            material_item=p_match["short_code"],
            description=p_match["description"],
            code=None,
        )
        .groupby(GROUP, dropna=False, as_index=False)
        .agg(qty=("_qty", "sum"), ctrl_type=("ctrl_type", _first_non_null))
    )

    # ---- Free Issue -----------------------------------------------------------
    f_direct = mat.merge(free_issue[["_norm", "item_code"]], on="_norm", how="inner")[["item_code", "_qty", "pid", "ctrl_type"]]
    f_alias = _apply_aliases(mat, aliases, "free_issue", extra_cols=["pid", "ctrl_type"])
    f_hits = pd.concat([f_direct, f_alias], ignore_index=True)
    f_match = f_hits.merge(
        free_issue[["item_code", "category", "description", "commodity_code", "unit"]],
        on="item_code", how="left",
    )
    free_issue_out = (
        f_match.assign(
            material_item=f_match["item_code"],
            description=f_match["description"],
            code=f_match["commodity_code"],
            unit=f_match["unit"],
        )
        .groupby(GROUP_FREE_ISSUE, dropna=False, as_index=False)
        .agg(qty=("_qty", "sum"), ctrl_type=("ctrl_type", _first_non_null))
    )

    # ---- GUK assemblies -------------------------------------------------------
    g_direct = mat.merge(guk_items[["_norm", "item_code"]], on="_norm", how="inner")[["item_code", "_qty", "ctrl_type"]]
    g_alias = _apply_aliases(mat, aliases, "guk_items", extra_cols=["ctrl_type"])
    g_hits = pd.concat([g_direct, g_alias], ignore_index=True)
    g_match = g_hits.merge(
        guk_items[["item_code", "category", "description", "unit"]],
        on="item_code", how="left",
    )
    guk_out = (
        g_match.assign(
            material_item=g_match["item_code"],
            description=g_match["description"],
            code=None,
            unit=g_match["unit"],
        )
        .groupby(GROUP_GUK, dropna=False, as_index=False)
        .agg(qty=("_qty", "sum"), ctrl_type=("ctrl_type", _first_non_null))
    )

    # ---- GUK subdivisions (qty_per_unit multiplier applied) -------------------
    g_detail = g_match[["item_code", "category", "description", "_qty", "ctrl_type"]].rename(
        columns={"item_code": "parent_item_code"}
    )
    sub_exploded = g_detail.merge(guk_subitems, left_on="parent_item_code", right_on="item_code", how="inner")
    sub_exploded["qty_per_unit"] = pd.to_numeric(sub_exploded["qty_per_unit"], errors="coerce").fillna(1.0)
    sub_exploded["_component_qty"] = sub_exploded["_qty"] * sub_exploded["qty_per_unit"]
    guk_sub_out = (
        sub_exploded.assign(
            material_item=sub_exploded["parent_item_code"],
            code=sub_exploded["sub_code"],
            description=sub_exploded["sub_description"],
            unit=sub_exploded["unit"],
        )
        .groupby(GROUP_GUK + ["qty_per_unit"], dropna=False, as_index=False)
        .agg(qty=("_component_qty", "sum"), ctrl_type=("ctrl_type", _first_non_null))
    )

    # ---- Unmatched (QA) -----------------------------------------------------
    matched_norms = (
        set(poles["_norm"]) | set(free_issue["_norm"]) | set(guk_items["_norm"]) | set(aliases["_norm"])
    )
    unmatched = mat.loc[~mat["_norm"].isin(matched_norms)]
    unmatched_out = (
        unmatched.groupby(["description"], dropna=False)["_qty"]
        .agg(total_qsub="sum", n_rows="count")
        .reset_index()
        .sort_values("total_qsub", ascending=False)
    )

    for out in (poles_out, free_issue_out, guk_out, guk_sub_out):
        out.sort_values(["category", "material_item"], inplace=True, na_position="last")

    return poles_out, free_issue_out, guk_out, guk_sub_out, unmatched_out


def _clean_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    v = round(float(v), 3)
    return int(v) if v == int(v) else v


# ---------------------------------------------------------------------------
# Pole-identifier vocabulary + PDF map scanning (unchanged from the desktop
# tool, including the network-scan date-filter fix)
# ---------------------------------------------------------------------------
EIGHT_DIGIT_RE = re.compile(r"\b\d{8}\b")
EIGHT_DIGIT_SUFFIX_RE = re.compile(r"\b\d{8}\s*\(?[A-Za-z0-9]{0,4}\)?")


def _normalize_pole_text(s) -> str:
    s = str(s or "")
    s = re.sub(r"[_\-]+", " ", s)
    s = re.sub(r"[()]", " ", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def build_pole_vocabulary(control_df: pd.DataFrame) -> dict:
    enids = control_df["enid"].dropna().astype(str).str.strip()
    enids = enids[enids != ""]
    unique_enids = enids.unique().tolist()

    digit8 = set()
    named_norm = {}
    for e in unique_enids:
        if re.match(r"^\d{8}$", e):
            digit8.add(e)
        else:
            norm = _normalize_pole_text(e)
            if norm and len(norm) >= 4:
                if norm not in named_norm or len(e) < len(named_norm[norm]):
                    named_norm[norm] = e

    if named_norm:
        alternatives = sorted(named_norm.keys(), key=len, reverse=True)
        pattern = r"(?<!\w)(?:" + "|".join(re.escape(a) for a in alternatives) + r")(?!\w)"
        named_regex = re.compile(pattern)
    else:
        named_regex = None

    return {"digit8": digit8, "named_norm": named_norm, "named_regex": named_regex}


def find_poles_in_text(text: str, vocab: dict) -> set:
    text = str(text or "")
    found = set()

    digit8_seen = EIGHT_DIGIT_RE.findall(text)
    if digit8_seen:
        for base in digit8_seen:
            if base in vocab["digit8"]:
                found.add(base)
        for m in EIGHT_DIGIT_SUFFIX_RE.finditer(text):
            candidate = m.group(0)
            norm = _normalize_pole_text(candidate)
            if norm in vocab["named_norm"]:
                found.add(vocab["named_norm"][norm])
        return found

    if vocab["named_regex"] is None:
        return found
    norm_text = _normalize_pole_text(text)
    if not norm_text:
        return found

    for m in vocab["named_regex"].finditer(norm_text):
        found.add(vocab["named_norm"][m.group(0)])

    return found


# ---------------------------------------------------------------------------
# Automatic network-path discovery of "Workpack zones" map folders
# ---------------------------------------------------------------------------
NETWORK_ROOT = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\1 - Outages Programme"

YEAR_DIR_REGEX = re.compile(r"^\d{4}$")
MONTH_DIR_REGEX = re.compile(r"^\d{2}\s*-\s*.+$")
WORKPACK_ZONES_NAME = "workpack zones"

_OUTAGE_RANGE_TO_REGEX = re.compile(
    r"^(\d{2})-(\d{2})-(\d{4})\s+to\s+(\d{2})-(\d{2})-(\d{4})\s*-\s*(.+)$", re.IGNORECASE
)
_OUTAGE_SINGLE_OR_DAYRANGE_REGEX = re.compile(r"^(\d{2})(?:_(\d{2}))?-(\d{2})-(\d{4})\s*-\s*(.+)$")


def _parse_outage_folder_dates(name):
    m = _OUTAGE_RANGE_TO_REGEX.match(name)
    if m:
        d1, mo1, y1, d2, mo2, y2, desc = m.groups()
        try:
            start = datetime(int(y1), int(mo1), int(d1))
            end = datetime(int(y2), int(mo2), int(d2))
        except ValueError:
            return None
        return start, end, desc

    m = _OUTAGE_SINGLE_OR_DAYRANGE_REGEX.match(name)
    if m:
        d1, d2, mo, y, desc = m.groups()
        try:
            start = datetime(int(y), int(mo), int(d1))
            end = datetime(int(y), int(mo), int(d2)) if d2 else start
        except ValueError:
            return None
        return start, end, desc

    return None


PDF_READ_WORKERS = min(32, (os.cpu_count() or 4) * 5)
NETWORK_SCAN_WORKERS = 8


def _resolve_outage_units(folder_path, folder_name, start_date, end_date):
    try:
        children = list(os.scandir(folder_path))
    except OSError:
        return [(folder_path, folder_name, start_date, end_date)]

    dated_children = []
    for entry in children:
        if not entry.is_dir():
            continue
        parsed = _parse_outage_folder_dates(entry.name)
        if parsed:
            child_start, child_end, _desc = parsed
            dated_children.append((entry, child_start, child_end))

    if not dated_children:
        return [(folder_path, folder_name, start_date, end_date)]

    units = []
    for entry, child_start, child_end in dated_children:
        units.extend(_resolve_outage_units(entry.path, entry.name, child_start, child_end))
    return units


def _scan_outage_for_zone_folders(outage_path, outage_name):
    found = []
    leaf_pdf_dirs = []
    for dirpath, dirnames, filenames in os.walk(outage_path):
        matches = [d for d in dirnames if d.strip().lower() == WORKPACK_ZONES_NAME]
        for d in matches:
            wz_path = os.path.join(dirpath, d)
            found.append({"path": wz_path, "day_outage": outage_name, "folder": os.path.basename(dirpath)})
            dirnames.remove(d)
        if not dirnames and any(fn.lower().endswith(".pdf") for fn in filenames):
            leaf_pdf_dirs.append(dirpath)

    if found:
        return found

    return [{"path": d, "day_outage": outage_name, "folder": os.path.basename(d)} for d in leaf_pdf_dirs]


def _to_datetime_or_none(d):
    """Normalizes a date-like value into a datetime.datetime (or None).
    `st.date_input()` returns a plain `datetime.date`, while the folder-name
    parsing here (_parse_outage_folder_dates) builds `datetime.datetime`
    objects - Python refuses to compare those two types directly
    ("'<' not supported between instances of 'datetime.datetime' and
    'datetime.date'"), which is exactly what was crashing the network
    scan. Converting whatever comes in to datetime.datetime once, right
    at the entry point, means every comparison downstream is guaranteed
    to be apples-to-apples regardless of which type the caller passed."""
    if d is None:
        return None
    if isinstance(d, datetime):
        return d
    return datetime.combine(d, datetime.min.time())


def find_workpack_zone_folders(root, date_from=None, date_to=None, progress_cb=None,
                                max_workers=NETWORK_SCAN_WORKERS):
    if not os.path.isdir(root):
        raise FileNotFoundError(f"Network path not found or not accessible: {root}")

    date_from = _to_datetime_or_none(date_from)
    date_to = _to_datetime_or_none(date_to)

    candidate_top_outages = []
    for year_entry in sorted(os.scandir(root), key=lambda e: e.name):
        if not year_entry.is_dir() or not YEAR_DIR_REGEX.match(year_entry.name):
            continue
        for month_entry in sorted(os.scandir(year_entry.path), key=lambda e: e.name):
            if not month_entry.is_dir() or not MONTH_DIR_REGEX.match(month_entry.name):
                continue
            for outage_entry in sorted(os.scandir(month_entry.path), key=lambda e: e.name):
                if not outage_entry.is_dir():
                    continue
                parsed = _parse_outage_folder_dates(outage_entry.name)
                if not parsed:
                    continue
                top_start, top_end, _desc = parsed
                if date_from is not None and top_end < date_from:
                    continue
                if date_to is not None and top_start > date_to:
                    continue
                candidate_top_outages.append((outage_entry.path, outage_entry.name, top_start, top_end))

    if not candidate_top_outages:
        return [], []

    def _resolve_and_scan(path, name, top_start, top_end):
        units = _resolve_outage_units(path, name, top_start, top_end)
        found = []
        in_range_names = []
        for unit_path, unit_label, unit_start, unit_end in units:
            if date_from is not None and unit_end < date_from:
                continue
            if date_to is not None and unit_start > date_to:
                continue
            in_range_names.append(unit_label)
            found.extend(_scan_outage_for_zone_folders(unit_path, unit_label))
        return found, in_range_names

    results = []
    scanned_names = set()
    total = len(candidate_top_outages)
    done = 0
    lock = threading.Lock()

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {
            ex.submit(_resolve_and_scan, path, name, top_start, top_end): name
            for path, name, top_start, top_end in candidate_top_outages
        }
        for fut in concurrent.futures.as_completed(futures):
            outage_name = futures[fut]
            try:
                found, in_range_names = fut.result()
                results.extend(found)
                scanned_names.update(in_range_names)
            except Exception:
                pass
            if progress_cb:
                with lock:
                    done += 1
                    progress_cb(f"Scanned {done}/{total} outage folder(s) ({outage_name})...")

    return results, sorted(scanned_names)


def find_workpack_zone_pdfs(root, date_from=None, date_to=None, progress_cb=None):
    zone_folders, scanned_outages = find_workpack_zone_folders(
        root, date_from=date_from, date_to=date_to, progress_cb=progress_cb
    )
    entries = []
    for zf in zone_folders:
        for dirpath, _dirnames, filenames in os.walk(zf["path"]):
            for fn in filenames:
                if fn.lower().endswith(".pdf"):
                    entries.append({"path": os.path.join(dirpath, fn), "folder": zf["folder"], "day_outage": zf["day_outage"]})
    outages_with_zones = {zf["day_outage"] for zf in zone_folders}
    outages_without_zones = sorted(o for o in scanned_outages if o not in outages_with_zones)
    return entries, len(zone_folders), len(entries), outages_without_zones


def collect_pdf_entries(selected_folders: list[str]) -> list[dict]:
    entries = []
    for folder in selected_folders:
        norm_folder = os.path.normpath(folder)
        folder_label = os.path.basename(norm_folder)
        day_outage_label = os.path.basename(os.path.dirname(norm_folder))
        for root_dir, _dirs, files in os.walk(norm_folder):
            for fn in files:
                if fn.lower().endswith(".pdf"):
                    entries.append({"path": os.path.join(root_dir, fn), "folder": folder_label, "day_outage": day_outage_label})
    return entries


def _read_pdf_poles(entry, vocab):
    path = entry["path"]
    fname = os.path.basename(path)
    map_key = (entry["day_outage"], entry["folder"], fname)
    found = set()
    doc = fitz.open(path)
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            page_had_annots = False
            for annot in page.annots() or []:
                info = annot.info or {}
                content = (info.get("content") or "").strip()
                if content:
                    page_had_annots = True
                    found |= find_poles_in_text(content, vocab)
            if not page_had_annots:
                text = page.get_text() or ""
                for line in text.splitlines():
                    line = line.strip()
                    if line:
                        found |= find_poles_in_text(line, vocab)
    finally:
        doc.close()
    return map_key, found


def build_map_pole_index_from_pdfs(pdf_entries: list[dict], vocab: dict, progress_cb=None,
                                    max_workers=PDF_READ_WORKERS) -> dict:
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed. Run: pip install pymupdf")
    index: dict = {}
    total = len(pdf_entries)
    done = 0
    lock = threading.Lock()

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {ex.submit(_read_pdf_poles, e, vocab): e for e in pdf_entries}
        for fut in concurrent.futures.as_completed(futures):
            entry = futures[fut]
            fname = os.path.basename(entry["path"])
            try:
                map_key, found = fut.result()
                if found:
                    index.setdefault(map_key, set()).update(found)
            except Exception as e:
                if progress_cb:
                    progress_cb(f"Failed to read {fname}: {e}")
            if progress_cb:
                with lock:
                    done += 1
                    progress_cb(f"Read {done}/{total} PDF(s) ({fname})")
    return index
