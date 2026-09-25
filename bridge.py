"""
bridge.py - glue between the web page and the tool modules.

Runs inside the browser (Pyodide, in a web worker). Every function takes
plain JSON-style arguments plus a `log` callback and returns plain data.
Files the user dropped are available as normal paths (under /in/...), and
anything a tool writes goes to a fresh folder under /out/<job>/, which the
page then offers as downloads. The tool modules themselves are the same
files you run in Jupyter - nothing in their logic is changed here.
"""
import os
import sys
import uuid
from datetime import datetime

OUT_ROOT = "/out"


class UserError(Exception):
    """A problem the user can fix - shown as-is in the page."""


def _outdir():
    d = os.path.join(OUT_ROOT, uuid.uuid4().hex[:8])
    os.makedirs(d, exist_ok=True)
    return d


def _name(name, ext, default):
    name = os.path.basename((name or "").strip()) or default
    if ext and not name.lower().endswith(ext):
        name += ext
    return name


def _need(path, label):
    if not path:
        raise UserError(f"Drop the {label} first.")
    if not os.path.exists(path):
        raise UserError(f"{label} is no longer loaded - drop it again.")
    return path


# ---------------------------------------------------------------------------
# Tool 1 - Target Price -> Control File
# ---------------------------------------------------------------------------
def tp_analyze(tp, log=print):
    import tool_tp_to_cf
    return tool_tp_to_cf.analyze(_need(tp, "Target Price file"), log=log)


def tp_run(control, tp, choices=None, log=print):
    import tool_tp_to_cf
    return tool_tp_to_cf.run(_need(control, "Template_CF"), _need(tp, "Target Price file"),
                             choices=choices or [], output_dir=_outdir(), log=log)


# ---------------------------------------------------------------------------
# Tool 2 - Aggregate Control Files
# ---------------------------------------------------------------------------
def aggregate(files, output_name="CF_aggregated.xlsx", log=print):
    import tool_aggregate_cf
    if not files:
        raise UserError("Drop at least one Control File.")
    for f in files:
        _need(f, "Control File")
    out = os.path.join(_outdir(), _name(output_name, ".xlsx", "CF_aggregated.xlsx"))
    return tool_aggregate_cf.run(file_paths=files, output_file=out, log=log)


# ---------------------------------------------------------------------------
# Tool 3 - Build Master Parquet
# ---------------------------------------------------------------------------
def master(aggregated, tracker, misc, output_name="", include_man_day=True, log=print):
    import tool_build_master
    _need(aggregated, "CF_aggregated.parquet")
    _need(tracker, "Project Tracker.parquet")
    _need(misc, "miscelaneous.parquet")
    default = "Master_" + datetime.now().strftime("%d-%m-%Y") + ".parquet"
    out = os.path.join(_outdir(), _name(output_name, ".parquet", default))
    return tool_build_master.run(aggregated, tracker, misc, out, include_man_day=bool(include_man_day), log=log)


# ---------------------------------------------------------------------------
# Tool 4 - Outputs (CV Excel report)
# ---------------------------------------------------------------------------
_DF_CACHE = {}


def _cached(kind, path, loader):
    key = (kind, path, os.path.getmtime(path))
    if key not in _DF_CACHE:
        if len(_DF_CACHE) > 4:
            _DF_CACHE.pop(next(iter(_DF_CACHE)))
        _DF_CACHE[key] = loader(path)
    return _DF_CACHE[key]


def _outputs_df(path):
    import tool_materials_report as M
    return _cached("outputs", _need(path, "master file"), M.load_file)


def outputs_options(path, log=print):
    import tool_materials_report as M
    return M.filter_options(_outputs_df(path))


def _outputs_filtered(path, filters, dates, start, end):
    import tool_materials_report as M
    try:
        return M.apply_filters(_outputs_df(path), filters or {}, dates or [], start or "", end or "")
    except (ValueError, TypeError) as e:
        raise UserError(f"Could not apply the date filter: {e}")


def outputs_apply(path, filters=None, dates=None, start="", end="", log=print):
    df, _ = _outputs_filtered(path, filters, dates, start, end)
    return {"rows": int(len(df))}


def outputs_export(path, filters=None, dates=None, start="", end="", cv_groups=None,
                   output_name="Outputs_Report.xlsx", logo_left="", logo_right="", log=print):
    import tool_materials_report as M
    df, date_filters = _outputs_filtered(path, filters, dates, start, end)
    # The report puts the Gaeltec / SPEN logos in the top corner. In the
    # browser the network path isn't reachable, so logos dropped on the page
    # are used instead (and simply left out if none were given).
    M.IMG_LEFT = logo_left if logo_left and os.path.exists(logo_left) else M.IMG_LEFT
    M.IMG_RIGHT = logo_right if logo_right and os.path.exists(logo_right) else M.IMG_RIGHT
    out = os.path.join(_outdir(), _name(output_name, ".xlsx", "Outputs_Report.xlsx"))
    return M.export(df, date_filters, cv_groups or [], out, log=log)


# ---------------------------------------------------------------------------
# Tool 5 - Work Instructions & map check
# ---------------------------------------------------------------------------
def _wi_df(path):
    import tool_work_instructions as WI
    try:
        return _cached("wi", _need(path, "master parquet"), WI.load_master_parquet)
    except ValueError as e:
        raise UserError(str(e))


def _wi_filtered(path, filters):
    import tool_work_instructions as WI
    df = _wi_df(path)
    try:
        return df, WI.filtered(df, filters or {})
    except ValueError as e:
        raise UserError(str(e))


def wi_options(path, log=print):
    import tool_work_instructions as WI
    return WI.filter_options(_wi_df(path))


def wi_preview(path, filters=None, log=print):
    import tool_work_instructions as WI
    _full, df = _wi_filtered(path, filters)
    return {"count": int(len(df)), **WI.preview_rows(df)}


def _entries_out(entries):
    out = []
    for e in entries:
        e = dict(e)
        for k in ("outage_start", "outage_end"):
            if isinstance(e.get(k), datetime):
                e[k] = e[k].strftime("%Y-%m-%d")
        out.append(e)
    return out


def _entries_in(entries):
    out = []
    for e in entries or []:
        e = dict(e)
        if not os.path.exists(e.get("path", "")):
            continue  # map was cleared from the page
        for k in ("outage_start", "outage_end"):
            if e.get(k):
                e[k] = datetime.strptime(e[k][:10], "%Y-%m-%d")
            else:
                e.pop(k, None)
        out.append(e)
    return out


def wi_add_folder(folder, log=print):
    """Same as 'Add folder manually' - every PDF under the dropped folder."""
    import tool_work_instructions as WI
    try:
        return {"entries": WI.entries_from_folder(folder)}
    except FileNotFoundError as e:
        raise UserError(str(e))


def wi_add_files(files, log=print):
    import tool_work_instructions as WI
    return {"entries": WI.entries_from_files(files)}


def wi_scan(root, date_from="", date_to="", log=print):
    """Same as 'Scan for maps in date range', run on the dropped Outages
    Programme folder(s) - laid out as <root>/<YYYY>/<MM - Month>/<outage>."""
    import tool_work_instructions as WI
    try:
        df, dt = WI.parse_date(date_from), WI.parse_date(date_to)
    except ValueError as e:
        raise UserError(str(e))
    res = WI.scan_network(root, date_from=df, date_to=dt, log=log)
    res["entries"] = _entries_out(res["entries"])
    return res


def wi_word(path, filters=None, pdf_entries=None, use_map_order=True,
            output_name="Work_Instructions.docx", log=print):
    import tool_work_instructions as WI
    _full, df = _wi_filtered(path, filters)
    if len(df) == 0:
        raise UserError("There are no rows to write - check the filters on step 1.")
    out = os.path.join(_outdir(), _name(output_name, ".docx", "Work_Instructions.docx"))
    return WI.generate_word(df, out, pdf_entries=_entries_in(pdf_entries),
                            use_map_order=bool(use_map_order), log=log)


def _fix_links(xlsx_path, scan_roots, target):
    """Point the map hyperlinks at the real files. Maps found by a scan sit
    under a dropped folder in the browser (one of `scan_roots`), laid out
    exactly like the real Outages Programme folder, so they are linked to
    the same place under `target` (the real network path, or a mapped drive
    letter). Maps added by hand only exist inside the browser, so their name
    is left as plain text instead of a broken link."""
    import openpyxl
    from openpyxl.styles import Font
    roots = sorted({r.rstrip("/") for r in scan_roots or [] if r}, key=len, reverse=True)
    target = (target or "").rstrip("\\/")
    wb = openpyxl.load_workbook(xlsx_path)
    changed = False
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                link = cell.hyperlink
                if link is None or not link.target or not link.target.startswith("/in/"):
                    continue
                root = next((r for r in roots if link.target.startswith(r + "/")), None)
                if root and target:
                    link.target = target + link.target[len(root):].replace("/", "\\")
                else:
                    cell.hyperlink = None
                    cell.font = Font()
                changed = True
    if changed:
        wb.save(xlsx_path)


def wi_report(path, filters=None, pdf_entries=None, threshold=85, scan_roots=None,
              real_root="", drive_letter="", output_name="Pole_PDF_CrossReference_Report.xlsx", log=print):
    import tool_work_instructions as WI
    full, df = _wi_filtered(path, filters)
    entries = _entries_in(pdf_entries)
    if len(df) == 0:
        raise UserError("Load a parquet file and apply filters first (no rows match).")
    if not entries:
        raise UserError("Add at least one map first - drop an Outages folder and scan it, or drop PDFs.")
    try:
        threshold = float(threshold)
    except (TypeError, ValueError):
        raise UserError("Threshold must be a number, e.g. 85")
    out = os.path.join(_outdir(), _name(output_name, ".xlsx", "Pole_PDF_CrossReference_Report.xlsx"))
    res = WI.generate_report(full, df, entries, out, threshold=threshold, log=log)
    # a mapped drive letter (e.g. Z:) stands for the real Outages folder, same as before
    _fix_links(out, scan_roots, (drive_letter or "").strip() or (real_root or "").strip())
    return res
