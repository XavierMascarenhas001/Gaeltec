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


# Columns each Build Master input must have (names compared lower-case, trimmed)
_TRACKER_COLS = ["shire", "project", "segmentcode", "job name", "pid", "po"]
_MISC_COLS = ["column_1", "column_2"]
_MASTER_HINT = ["job", "enid", "description", "qsub", "datetouse", "sourcefile"]


def _parquet_columns(path):
    try:
        import pyarrow.parquet as pq
        names = pq.read_schema(path).names
    except Exception:
        import pandas as pd
        names = list(pd.read_parquet(path).columns)
    return [str(c).strip().lower() for c in names]


def _check_inputs(tracker, misc, aggregated=None):
    """Fail fast, before any long work, if a file went into the wrong box."""
    def what_is(cols):
        if all(c in cols for c in _MASTER_HINT):
            return "a Master file"
        if all(c in cols for c in _TRACKER_COLS):
            return "the Project Tracker"
        if all(c in cols for c in _MISC_COLS):
            return "the miscelaneous file"
        return None
    problems = []
    for path, label, need in ((tracker, "Project Tracker.parquet", _TRACKER_COLS), (misc, "miscelaneous.parquet", _MISC_COLS)):
        cols = _parquet_columns(path)
        missing = [c for c in need if c not in cols]
        if missing:
            seems = what_is(cols)
            problems.append(
                f"• The file in the '{label}' box ({os.path.basename(path)}) is missing the column(s): {', '.join(missing)}."
                + (f" It looks like {seems} – drop {label} there instead." if seems else f" Make sure it's {label}."))
    if aggregated:
        cols = _parquet_columns(aggregated)
        if not all(c in cols for c in ("district", "project", "circuit", "job", "description")):
            problems.append(f"• The file in the 'CF_aggregated.parquet' box ({os.path.basename(aggregated)}) doesn't look like CF_aggregated_aggregated.parquet.")
    if problems:
        raise UserError("Nothing was run – a file looks like it's in the wrong box:\n" + "\n".join(problems))


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
    _check_inputs(tracker, misc, aggregated)
    default = "Master_" + datetime.now().strftime("%d-%m-%Y") + ".parquet"
    out = os.path.join(_outdir(), _name(output_name, ".parquet", default))
    return tool_build_master.run(aggregated, tracker, misc, out, include_man_day=bool(include_man_day), log=log)


# ---------------------------------------------------------------------------
# Tools 2 + 3 in one go - Control Files -> Master parquet + xlsx
# (exactly the same two functions as the separate tools, chained)
# ---------------------------------------------------------------------------
def build_master_direct(files, tracker, misc, output_name="", include_man_day=True, keep_aggregated=False, log=print):
    import tool_aggregate_cf
    import tool_build_master
    if not files:
        raise UserError("Drop at least one Control File.")
    for f in files:
        _need(f, "Control File")
    _need(tracker, "Project Tracker.parquet")
    _need(misc, "miscelaneous.parquet")
    _check_inputs(tracker, misc)
    out_dir = _outdir()
    log("━━━━━━━━ Step 1 of 2 · Aggregating the Control Files ━━━━━━━━")
    agg_outputs = tool_aggregate_cf.run(file_paths=files, output_file=os.path.join(out_dir, "CF_aggregated.xlsx"), log=log,
                                        write_excel=bool(keep_aggregated))
    agg_parquet = next((p for p in agg_outputs if p.endswith("_aggregated.parquet")), None)
    if not agg_parquet:
        raise UserError("No data was aggregated from those Control Files - see the log above.")
    log("━━━━━━━━ Step 2 of 2 · Building the Master ━━━━━━━━")
    default = "Master_" + datetime.now().strftime("%d-%m-%Y") + ".parquet"
    out = os.path.join(out_dir, _name(output_name, ".parquet", default))
    master_outputs = tool_build_master.run(agg_parquet, tracker, misc, out, include_man_day=bool(include_man_day), log=log)
    if keep_aggregated:
        return master_outputs + agg_outputs
    for p in agg_outputs:
        try:
            os.remove(p)
        except OSError:
            pass
    return master_outputs


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


def _find_loaded_misc():
    """A miscelaneous.parquet (Column_1 + Column_3) dropped anywhere on the
    page - files dropped on the page live under /in/<n>/. Files named
    like 'misc...' are tried first. Returns '' if there is none."""
    found = []
    for root, _dirs, files in os.walk("/in"):
        for fn in files:
            if fn.lower().endswith(".parquet"):
                found.append(os.path.join(root, fn))
    found.sort(key=lambda p: (0 if "misc" in os.path.basename(p).lower() else 1, p))
    for p in found:
        try:
            cols = _parquet_columns(p)
        except Exception:
            continue
        if "column_1" in cols and "column_3" in cols:
            return p
    return ""


def outputs_export(path, filters=None, dates=None, start="", end="", cv_groups=None,
                   output_name="Outputs_Report.xlsx", logo_left="", logo_right="",
                   misc="", service_code_sheet=False, log=print):
    import tool_materials_report as M
    df, date_filters = _outputs_filtered(path, filters, dates, start, end)
    # The report puts the Gaeltec / SPEN logos in the top corner. In the
    # browser the network path isn't reachable, so logos dropped on the page
    # are used instead (and simply left out if none were given).
    M.IMG_LEFT = logo_left if logo_left and os.path.exists(logo_left) else M.IMG_LEFT
    M.IMG_RIGHT = logo_right if logo_right and os.path.exists(logo_right) else M.IMG_RIGHT
    # Service codes come from the miscelaneous file dropped on the page -
    # if it wasn't dropped in this tool, use one already loaded anywhere on
    # the page (e.g. in the Build Master boxes)
    if not misc:
        misc = _find_loaded_misc()
        if misc:
            log(f"Using the miscelaneous file already loaded on the page: {os.path.basename(misc)}")
    if misc:
        _need(misc, "miscelaneous.parquet")
        cols = _parquet_columns(misc)
        if not all(c in cols for c in ("column_1", "column_3")):
            raise UserError(f"The file in the 'miscelaneous.parquet' box ({os.path.basename(misc)}) has no "
                            f"Column_1 / Column_3 (description / service code) - make sure it's miscelaneous.parquet.")
    elif service_code_sheet:
        raise UserError("Drop miscelaneous.parquet to add the 'Variations_service codes' sheet "
                        "(the service codes come from that file).")
    out = os.path.join(_outdir(), _name(output_name, ".xlsx", "Outputs_Report.xlsx"))
    return M.export(df, date_filters, cv_groups or [], out, service_codes_path=misc or "",
                    include_service_code_sheet=bool(service_code_sheet), log=log)


def outputs_weekly(path, docs=None, cv_groups=None, logo_left="", logo_right="", misc="",
                   service_code_sheet=False, log=print):
    """The weekly district reports in one go – one Excel per doc:
    docs = [{name, district, projects: [...], date_col: "done"/"plan1", start, end}]."""
    import tool_materials_report as M
    if not docs:
        raise UserError("Tick at least one weekly report.")
    M.IMG_LEFT = logo_left if logo_left and os.path.exists(logo_left) else M.IMG_LEFT
    M.IMG_RIGHT = logo_right if logo_right and os.path.exists(logo_right) else M.IMG_RIGHT
    if not misc:
        misc = _find_loaded_misc()
    if service_code_sheet and not misc:
        raise UserError("Drop miscelaneous.parquet to add variations (the service codes come from that file).")
    base = _outputs_df(path)
    out_dir, files = _outdir(), []
    for d in docs:
        try:
            df, date_filters = M.apply_filters(base, {"shire": [d["district"]], "project": d.get("projects") or []},
                                               [d.get("date_col") or "done"], d.get("start") or "", d.get("end") or "")
        except (ValueError, TypeError) as e:
            raise UserError(f"{d.get('name')}: could not apply the date filter: {e}")
        name = _name(d.get("name"), ".xlsx", "Outputs.xlsx")
        log(f"{name}: {len(df):,} row(s)")
        files += M.export(df, date_filters, cv_groups or [], os.path.join(out_dir, name),
                          service_codes_path=misc or "", include_service_code_sheet=bool(service_code_sheet),
                          log=lambda *_a, **_k: None)
    return files


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


# ---------------------------------------------------------------------------
# Dashboards (Master Control, Materials Breakdown, Network Job Tracker)
# ---------------------------------------------------------------------------
def _dash(fn, *a, **k):
    try:
        return fn(*a, **k)
    except (ValueError, KeyError, FileNotFoundError) as e:
        raise UserError(str(e).strip("'\""))


def mc_load(path, log=print):
    import dash_master as D
    return _dash(D.load, _need(path, "Master parquet"))


def mc_view(path, log=print, **filters):
    import dash_master as D
    return _dash(D.view, _need(path, "Master parquet"), **filters)


def mb_load(master_path, control_path, log=print):
    import dash_materials as D
    return _dash(D.load, _need(master_path, "materials master (materials_all.parquet)"), _need(control_path, "Master parquet"))


def mb_view(master_path, control_path, log=print, **kw):
    import dash_materials as D
    return _dash(D.view, master_path, control_path, **kw)


def mb_maps(master_path, control_path, filters=None, pdf_entries=None, log=print):
    import dash_materials as D
    return _dash(D.maps, master_path, control_path, filters=filters, pdf_entries=pdf_entries, log=log)


def mb_folders(folders, log=print):
    import dash_materials as D
    return _dash(D.collect_folders, folders, log=log)


def mb_scan(root, date_from="", date_to="", log=print):
    import dash_materials as D
    return _dash(D.scan, root, date_from or None, date_to or None, log=log)


def mb_export(master_path, control_path, filters=None, scope="All poles matching filters", pdf_entries=None, log=print):
    import dash_materials as D
    out = os.path.join(_outdir(), f"Materials_Breakdown_{datetime.now():%Y%m%d_%H%M}.xlsx")
    return _dash(D.export, master_path, control_path, out, filters=filters, scope=scope, pdf_entries=pdf_entries, log=log)


def nt_load(path, cols=None, log=print):
    import dash_tracker as D
    return _dash(D.load, _need(path, "Master parquet"), cols)


def nt_view(path, cols=None, filters=None, granularity="Month", detail=None, misc="", pid_mode="All", tab=None, log=print):
    import dash_tracker as D
    # service codes for the Totals tab: the file in the dashboard's slot, or one loaded anywhere on the page
    misc = misc if misc and os.path.exists(misc) else (_find_loaded_misc() if tab in (None, "totals") else "")
    return _dash(D.view, _need(path, "Master parquet"), cols, filters, granularity, detail, misc=misc, pid_mode=pid_mode, tab=tab)


def nt_variations_csv(path, cols=None, filters=None, misc="", log=print):
    """Totals tab: every row of 'Variations - every item' as a CSV download."""
    import dash_tracker as D
    misc = misc if misc and os.path.exists(misc) else _find_loaded_misc()
    out = os.path.join(_outdir(), f"Variations_every_item_{datetime.now():%Y-%m-%d_%H%M}.csv")
    return _dash(D.variations_csv, _need(path, "Master parquet"), cols, filters, misc, out)


def warmup(log=print):
    """Called by the page once Python is up: loads pandas' parquet reader and the
    dashboard modules in the background, so the first file you open is quicker."""
    import pyarrow.parquet  # noqa: F401
    import dash_tracker, dash_master  # noqa: F401
    return True


def nt_outages(outage_path, master_path=None, cols=None, filters=None, log=print, **kw):
    import dash_tracker as D
    mp = master_path if master_path and os.path.exists(master_path) else None
    return _dash(D.outages, _need(outage_path, "Outages Programme workbook"), mp, cols, filters, **kw)


def nt_outage_folders(outage_path, log=print, **kw):
    import dash_tracker as D
    return _dash(D.outage_folders_for, _need(outage_path, "High level planning workbook"), **kw)


def nt_ics(ics_path, log=print):
    import dash_tracker as D
    return _dash(D.ics, _need(ics_path, "calendar (.ics) file"))


def nt_forecast(path, log=print, **kw):
    import dash_tracker as D
    return _dash(D.forecast, _need(path, "Service Partner Workbank"), **kw)


# ---------------------------------------------------------------------------
# Outage Calendar (tool_outage_calendar.py): High level planning -> Excel calendar + Outlook .ics
#   oc_load -> (oc_folders -> oc_enrich) -> oc_build
# The parsed planning rows, folder look-ups, roster and absence data are kept for the session
# and shared with the Planning dashboard's Jobs calendar.
# ---------------------------------------------------------------------------
def _oc_team(roster=None, absence=None):
    import tool_outage_calendar as C
    if roster is not None:
        C.TEAM["roster"] = roster or ""
    if absence is not None:
        C.TEAM["absence"] = [p for p in absence if p]
    try:
        return C.team_context()
    except RuntimeError as e:
        raise UserError(f"Absence CSV: {e}")


def oc_load(planning="", password="", master="", absence=None, roster=None, date_field="Plan1", log=print):
    import tool_outage_calendar as C
    _need(planning, "High level planning workbook")
    try:
        rows = C.programme(planning, password)
    except (C.NeedPassword, RuntimeError) as e:
        raise UserError(str(e))
    dists = {}
    for r in rows:
        d = C._s(r.get("District"))
        if d:
            dists.setdefault(C._canonical_district(d), d.title() if d.isupper() else d)
    res = {"rows": len(rows), "years": sorted({r["Outage Date"].year for r in rows}),
           "districts": sorted(dists.values(), key=str.lower),
           "links": len(C.outage_folders(rows)),
           "enriched": sum(1 for f in C.outage_folders(rows) if f in C.ENRICHED),
           "projects": sorted({C._s(r.get("Project")) for r in rows if C._s(r.get("Project"))}, key=str.lower)}
    if master and os.path.exists(master):
        try:
            m = C.master_frame(master, date_field)
        except Exception as e:
            raise UserError(f"Master file: {e}")
        res["projects"] = sorted(set(res["projects"]) | {p for p in m["project"].unique() if p}, key=str.lower)
        res["master_days"] = int(m["date"].nunique())
    tc = _oc_team(roster, absence)
    roster_groups, leaders = C.parse_roster(C.TEAM.get("roster") or "")
    res["roster"] = {"groups": len(roster_groups), "teams": sum(1 for g in roster_groups if g.lower().startswith("team ")),
                     "people": sum(len(v) for v in roster_groups.values()), "leaders": len(leaders)}
    res["absence"] = {"files": len(C.TEAM.get("absence") or []), "people": len(tc[2]) if tc else 0}
    return res


def oc_folders(planning="", password="", districts=None, years=None, projects=None, log=print):
    """Outage folders (Windows paths) of the filtered rows that haven't been looked up yet."""
    import tool_outage_calendar as C
    rows = C.filter_rows(C.programme(_need(planning, "High level planning workbook"), password), districts, years,
                         projects=projects)
    folders = C.outage_folders(rows)
    return {"folders": [f for f in folders if f not in C.ENRICHED], "all": len(folders), "base": C.BASE_UNC_PATH}


def oc_enrich(dir="", found=None, log=print):
    import tool_outage_calendar as C
    return C.enrich_folders(found or {}, dir)


def oc_build(planning="", password="", master="", date_field="Plan1", districts=None, projects=None, years=None,
             include_comments=True, include_total=True, output_name="", log=print):
    import tool_outage_calendar as C
    rows = C.programme(_need(planning, "High level planning workbook"), password)
    rows = C.with_enrichment(C.filter_rows(rows, districts, years, projects=projects))
    mcf = None
    if master and os.path.exists(master):
        mcf = C.mcf_day_stats(C.master_frame(master, date_field), districts, projects, years)
    tc = _oc_team()
    if not rows and not mcf and not tc:
        raise UserError("No outages, Master figures or team data match these filters.")
    parts = []
    if districts:
        parts.append("District: " + ", ".join(sorted(districts)))
    if projects:
        parts.append("Project: " + ", ".join(sorted(projects)))
    if years:
        parts.append("Year: " + ", ".join(str(y) for y in sorted(years)))
    summary = "; ".join(parts) or None
    log(f"{len(rows)} outage(s) · {len(mcf or {})} Master day(s) · team breakdown {'on' if tc else 'off'}")
    stem = os.path.splitext(_name(output_name, "", f"outage_calendar_{datetime.now():%Y%m%d}"))[0]
    out = _outdir()
    x = os.path.join(out, stem + ".xlsx")
    C.build_calendar_workbook(rows, mcf_stats=mcf, filter_summary=summary, mcf_enabled=mcf is not None,
                              include_comments=include_comments, team_ctx=tc,
                              include_financial_total=include_total).save(x)
    i = os.path.join(out, stem + ".ics")
    with open(i, "w", encoding="utf-8", newline="") as fh:
        fh.write(C.build_ics(rows, mcf_stats=mcf, mcf_enabled=mcf is not None, team_ctx=tc,
                             include_financial_total=include_total))
    return [x, i]


# ---------------------------------------------------------------------------
# Generators <-> High-level planning (tool_generator_sync.py - the same file
# that runs in Jupyter). The page calls these one after the other:
#   gs_probe -> gs_load -> gs_days -> gs_preview -> gs_apply
# The two workbooks stay decrypted in memory between calls; after an Apply
# the updated copy is kept, so a second run continues from it (dropping a
# new file replaces it).
# ---------------------------------------------------------------------------
_GS = {}


def _gs_is_encrypted(path):
    with open(path, "rb") as f:
        return f.read(4) != b"PK\x03\x04"


def gs_probe(tracker="", planning="", log=print):
    """Which of the dropped workbooks are password protected (the page then
    loads the decryption package only when it's actually needed)."""
    out = {}
    for key, path in (("tracker", tracker), ("planning", planning)):
        out[key] = {"encrypted": bool(path and os.path.exists(path) and _gs_is_encrypted(path))}
    return out


def _gs_status():
    import tool_generator_sync as G
    s = {}
    t = _GS.get("trk")
    if t:
        s["tracker"] = (f"{_GS['trk_name']}{' (protected)' if _GS['trk_enc'] else ''} – {len(t['rows'])} rows, "
                        f"next free row {t['last_used_row'] + 1}")
    p = _GS.get("plan")
    if p:
        s["planning"] = (f"{_GS['plan_name']}{' (protected)' if _GS['plan_enc'] else ''} – {len(p['rows'])} outages "
                         f"on sheets {', '.join(p['sheets'])}")
    return s


def gs_load(tracker="", planning="", password="", log=print):
    import tool_generator_sync as G
    _need(tracker, "Generator Tracker")
    # the planning workbook is only needed for tasks 1-2 (the email uses the tracker alone)
    todo = [("trk", tracker, G.read_tracker)] + ([("plan", planning, G.read_planning)] if planning else [])
    if not planning:
        for k in ("plan", "plan_plain", "plan_enc", "plan_name", "plan_stamp"):
            _GS.pop(k, None)
    for key, path, reader in todo:
        stamp = (path, os.path.getmtime(path), password if _gs_is_encrypted(path) else "")
        if _GS.get(key + "_stamp") == stamp:
            continue  # same file (and password) as last time - keep it, incl. any applied changes
        label = "Generator Tracker" if key == "trk" else "High-level planning"
        raw = open(path, "rb").read()
        if raw[:4] != b"PK\x03\x04" and not password:
            raise UserError(f"The {label} ({os.path.basename(path)}) is password protected – type the password and click Read files.")
        try:
            plain, enc = G.decrypt_if_needed(raw, password)
            data = reader(plain)
        except RuntimeError as e:
            raise UserError(f"{label}: {e}")
        log(f"Read {label}: {os.path.basename(path)}")
        _GS.update({key: data, key + "_plain": plain, key + "_enc": enc, key + "_name": os.path.basename(path),
                    key + "_stamp": stamp})
    _GS["password"] = password
    _GS.pop("plan_res", None)
    return _gs_status()


def _gs_ready():
    if not (_GS.get("trk") and _GS.get("plan")):
        raise UserError("Drop both workbooks and click Read files first.")


def gs_email_list(date_from="", date_to="", log=print):
    """Generator Tracker rows (with a Type of supply) between the two dates, for the email."""
    import tool_generator_sync as G
    if not _GS.get("trk"):
        raise UserError("Drop the Generator Tracker and click Read files first (step 1).")
    try:
        a = datetime.strptime(date_from[:10], "%Y-%m-%d").date()
        b = datetime.strptime(date_to[:10], "%Y-%m-%d").date()
    except ValueError:
        raise UserError("Choose both dates (From and To).")
    a, b = min(a, b), max(a, b)
    cands = G.email_candidates(_GS["trk"], a, b)
    _GS["email_cands"] = {c["_id"]: c for c in cands}
    days = {}
    for c in cands:
        d = days.setdefault(c["_date"], {"date": c["_date"].isoformat(), "label": c["_date"].strftime("%a %d %b %Y"), "items": []})
        d["items"].append({"id": c["_id"], "label": c["_label"]})
    has_days = bool(_GS["trk"]["cols"].get("days"))
    return {"days": list(days.values()), "days_col": has_days}


def gs_email(ids=None, log=print):
    import tool_generator_sync as G
    rows = [_GS.get("email_cands", {}).get(i) for i in ids or []]
    rows = [r for r in rows if r]
    if not rows:
        raise UserError("Tick at least one generator.")
    return G.build_email(rows)


def gs_days(task=1, date_from="", date_to="", log=print):
    import tool_generator_sync as G
    _gs_ready()
    try:
        a = datetime.strptime(date_from[:10], "%Y-%m-%d").date()
        b = datetime.strptime(date_to[:10], "%Y-%m-%d").date()
    except ValueError:
        raise UserError("Choose both dates (From and To).")
    a, b = min(a, b), max(a, b)
    task = int(task)
    cands = G.candidates_task1(_GS["plan"], a, b) if task == 1 else G.candidates_task2(_GS["trk"], a, b)
    _GS["cands"] = {c["_id"]: c for c in cands}
    _GS["task"] = task
    _GS.pop("plan_res", None)
    days = {}
    for c in cands:
        d = days.setdefault(c["_date"], {"date": c["_date"].isoformat(), "label": c["_date"].strftime("%a %d %b %Y"), "items": []})
        d["items"].append({"id": c["_id"], "label": c["_label"]})
    return {"days": list(days.values())}


def gs_preview(task=1, ids=None, log=print):
    import tool_generator_sync as G
    _gs_ready()
    if int(task) != _GS.get("task"):
        raise UserError("The task was changed – click Show days & outages again.")
    items = [_GS["cands"][i] for i in ids or [] if i in _GS.get("cands", {})]
    if not items:
        raise UserError("Tick at least one day (and outage).")
    res = G.plan_task1(items, _GS["trk"]) if _GS["task"] == 1 else G.plan_task2(items, _GS["plan"])
    _GS["plan_res"] = res
    df = G.preview_frame(res)
    df = df.astype(object).where(df.notna(), "")
    conflicts = []
    for i, it in enumerate(res["items"]):
        if not it["conflict"]:
            continue
        if res["task"] == 1:
            conflicts.append({"i": i, "where": it["where"], "current": it["current"], "new": it["new"]})
        else:
            conflicts.append({"i": i, "where": f"{it['sheet']}!{it['col']}{it['row']} ({it['field']}) · {it['date']} "
                                                f"{it['outage']} {it['circuit']}".strip(),
                              "current": it["current"], "new": it["new"]})
    n_conf = len(conflicts)
    if res["task"] == 1:
        n_new = len(res["items"]) - n_conf
        n_mid = sum(1 for it in res["items"] if not it["conflict"] and "inserted" in it["where"])
        summary = (f"{n_new} outage(s) will be added to the Generators table in date order"
                   + (f" ({n_mid} inserted between existing dates)" if n_mid else "")
                   + f"; {n_conf} already in the tracker.")
        choices = G.T1_CHOICES
    else:
        n_w = sum(1 for it in res["items"] if not it["conflict"] and not it["same"])
        summary = (f"{n_w} empty cell(s) will be filled; {n_conf} cell(s) already have information; "
                   f"{len(res['unmatched'])} outage(s) could not be matched.")
        choices = G.T2_CHOICES
    return {
        "summary": summary,
        "columns": [str(c) for c in df.columns],
        "rows": [[str(v).replace("\n", "  ⏎  ") for v in r] for r in df.values.tolist()],
        "conflicts": conflicts,
        "choices": [{"label": l, "value": v} for l, v in choices],
        "can_apply": any(not it.get("same") for it in res["items"]),
    }


def gs_apply(choices=None, log=print):
    import tool_generator_sync as G
    res = _GS.get("plan_res")
    if not res:
        raise UserError("Click Preview first.")
    choices = {int(k): v for k, v in (choices or {}).items() if v}
    undecided = [it for i, it in enumerate(res["items"]) if it["conflict"] and i not in choices]
    if undecided:
        raise UserError(f"Choose what to do for the {len(undecided)} item(s) that already have information "
                        f"(or use 'Set all to').")
    pwd = _GS.get("password", "")
    if res["task"] == 1:
        edits, summary = G.build_edits_task1(res, choices, _GS["trk"])
        if not edits:
            raise UserError("Nothing to write – everything was kept as it is.")
        plain, wlog = G.apply_task1(_GS["trk_plain"], edits)
        final = G.encrypt(plain, pwd) if _GS["trk_enc"] else plain
        name = _GS["trk_name"]
        _GS.update(trk_plain=plain, trk=G.read_tracker(plain))
    else:
        edits, summary = G.build_edits_task2(res, choices)
        if not edits:
            raise UserError("Nothing to write – everything was kept as it is.")
        plain, wlog = G.apply_task2(_GS["plan_plain"], edits)
        final = G.encrypt(plain, pwd) if _GS["plan_enc"] else plain
        name = _GS["plan_name"]
        _GS.update(plan_plain=plain, plan=G.read_planning(plain))
    _GS.pop("plan_res", None)
    for line in wlog:
        log("• " + line)
    for s in summary:
        log(" · ".join(f"{v}" for v in s.values()).replace("\n", "  ⏎  "))
    stem, ext = os.path.splitext(name)
    out = os.path.join(_outdir(), f"{stem}_updated_{datetime.now():%Y-%m-%d_%H%M}{ext or '.xlsx'}")
    with open(out, "wb") as f:
        f.write(final)
    return [out]


def gs_status(log=print):
    return _gs_status()
