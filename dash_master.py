"""
dash_master.py - Master Control (job costing) dashboard, data side.

The calculations below are taken from master_control_dashboard.py unchanged
(loading/cleaning, date presets, cascading Circuit/Pole options, Panel 01-04
aggregations). Only the Streamlit screen code was replaced: instead of
drawing widgets and charts, each function returns plain data that the web
page draws.
"""
import datetime as dt
import os

import numpy as np
import pandas as pd

COLUMNS_NEEDED = [
    "job", "plan1", "plan2", "done", "datetouse", "invoice date", "total", "orig", "comment",
    "district", "project", "project manager", "team", "team lider", "pid",
    "circuit", "enid", "sourcefile",
]
# "Combined date" replaces DateToUse: Done when it's today/past, otherwise Plan1 when
# it's today/future (engine.combined_date) - added to the Master when it's read.
COMBINED = "Combined date"
DATE_FIELDS = {COMBINED: "combined", "Plan1": "plan1", "Plan2": "plan2", "Done": "done"}
PRESETS = ["All time", "Current week", "Last week", "Current month", "Last month", "This year", "Last year", "Custom range"]

# ---- Cascading filters: each one only offers what the filters before it allow ----
_BLANK_OPT = {"", "nan", "none", "<na>", "nat"}


def cascade_options(df, base_mask, order, selections, strs=None):
    """order = [(filter key, column)], top to bottom. Returns ({key: options}, mask after every level).
    strs(col), when given, returns df[col].astype(str).str.strip() from a cache."""
    mask = base_mask.copy()
    opts = {}
    for key, col in order:
        if not col or col not in df.columns:
            opts[key] = []
            continue
        col_s = strs(col) if strs else df[col].astype(str).str.strip()
        opts[key] = sorted({v for v in col_s[mask].unique() if isinstance(v, str) and v.lower() not in _BLANK_OPT})
        sel = selections.get(key) or []
        if sel:
            mask &= col_s.isin([str(x).strip() for x in sel])
    return opts, mask


# ---- Quick ranges on the calendar (today's date): whole weeks Mon-Sun, whole months ----
CAL_PRESETS = ("Current week", "Last week", "Current month", "Last month")


def calendar_range(preset, today=None):
    t = pd.Timestamp(today if today is not None else pd.Timestamp.today()).normalize()
    monday = t - pd.Timedelta(days=t.weekday())
    if preset == "Current week":
        return monday.date(), (monday + pd.Timedelta(days=6)).date()
    if preset == "Last week":
        return (monday - pd.Timedelta(days=7)).date(), (monday - pd.Timedelta(days=1)).date()
    first = t.replace(day=1)
    if preset == "Current month":
        return first.date(), (first + pd.offsets.MonthEnd(0)).date()
    last_end = first - pd.Timedelta(days=1)          # Last month
    return last_end.replace(day=1).date(), last_end.date()


# ---- Plan2: the most repeated comments for the rows in view ----
_BLANK_COMMENTS = {"", "nan", "none", "nat", "<na>", "null", "0", "-"}


def plan2_comments(frame, date_col, comment_col="comment", job_col="job", limit=500):
    if comment_col not in frame.columns or date_col not in frame.columns:
        return {"available": False}
    sub = frame[frame[date_col].notna()]
    txt = sub[comment_col].astype(str).str.strip()
    keep = ~txt.str.lower().isin(_BLANK_COMMENTS)
    sub, txt = sub[keep], txt[keep]
    if sub.empty:
        return {"available": True, "rows": int(len(frame[frame[date_col].notna()])), "with_comment": 0, "distinct": 0, "top": []}
    txt = txt.str.replace("\u00a0", " ").str.replace(r"\s+", " ", regex=True)
    g = pd.DataFrame({"c": txt.values, "k": txt.str.lower().str.strip(" .").values, "d": pd.to_datetime(sub[date_col]).values,
                      "j": sub[job_col].astype(str).values if job_col in sub.columns else ""})
    # same comment typed differently ("NO TIME" / "No Time") counts once, shown in its commonest spelling
    shown = g.groupby("k")["c"].agg(lambda x: x.value_counts().index[0])
    agg = (g.groupby("k").agg(n=("k", "size"), jobs=("j", "nunique"), first=("d", "min"), last=("d", "max"))
           .sort_values(["n", "jobs"], ascending=False))
    agg.index = shown.reindex(agg.index).values
    top = [[c, int(r["n"]), int(r["jobs"]), r["first"].strftime("%Y-%m-%d"), r["last"].strftime("%Y-%m-%d")] for c, r in agg.head(limit).iterrows()]
    return {"available": True, "rows": int(len(frame[frame[date_col].notna()])), "with_comment": int(len(sub)),
            "distinct": int(len(agg)), "top": top}



# --------------------------------------------------------------------------
# Data loading (unchanged)
# --------------------------------------------------------------------------
def load_data(path_str: str):
    """Returns (df, error_message). Never raises."""
    try:
        try:   # plan2 / comment are optional - older Master files may not have them
            import pyarrow.parquet as _pq
            _have = set(_pq.read_schema(path_str).names)
        except Exception:
            _have = set(COLUMNS_NEEDED)
        df = pd.read_parquet(path_str, columns=[c for c in COLUMNS_NEEDED if c in _have or c not in ("plan2", "comment", "datetouse")])
    except Exception as e:
        return None, str(e)

    # total / orig sometimes land as text — coerce to numeric
    df["total"] = pd.to_numeric(df["total"], errors="coerce")
    df["orig"] = pd.to_numeric(df["orig"], errors="coerce")

    for c in ["plan1", "plan2", "done", "datetouse", "invoice date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    for c in ["district", "project", "project manager", "team", "team lider", "pid", "job", "circuit", "enid", "sourcefile"]:
        df[c] = df[c].astype("string").str.strip()
        df[c] = df[c].replace("", pd.NA)

    job_lower = df["job"].str.lower()
    df["flag"] = "other"
    df.loc[job_lower.str.startswith(("m -", "m-"), na=False), "flag"] = "material"
    df.loc[job_lower.str.startswith(("c -", "c-"), na=False), "flag"] = "construction"

    for c in ["plan1", "plan2", "done", "datetouse"]:
        if c in df.columns:
            df[c] = df[c].where(df[c].dt.year > 1901)

    from engine import combined_date
    df["combined"] = combined_date(df["done"], df["plan1"])
    return df, None


# --------------------------------------------------------------------------
# Formatting helpers (unchanged)
# --------------------------------------------------------------------------
def auto_granularity(date_from: dt.date, date_to: dt.date) -> str:
    days = (date_to - date_from).days
    if days <= 31:
        return "Day"
    if days <= 180:
        return "Week"
    if days <= 900:
        return "Month"
    return "Year"


def bucket_series(dates: pd.Series, granularity: str) -> pd.Series:
    if granularity == "Day":
        return dates.dt.to_period("D").dt.start_time
    if granularity == "Week":
        return dates.dt.to_period("W-MON").dt.start_time
    if granularity == "Month":
        return dates.dt.to_period("M").dt.start_time
    return dates.dt.to_period("Y").dt.start_time


def bucket_label(ts: pd.Timestamp, granularity: str) -> str:
    if granularity == "Day":
        return ts.strftime("%-d %b") if hasattr(ts, "strftime") else str(ts)
    if granularity == "Week":
        return "w/c " + ts.strftime("%-d %b")
    if granularity == "Month":
        return ts.strftime("%b %Y")
    return ts.strftime("%Y")


def resolve_date_range(frame, active_col, preset, custom_from=None, custom_to=None):
    """pick_date_field_and_range's range logic - presets anchored to the
    LATEST date in the data, clipped to the data's own range."""
    if preset in CAL_PRESETS:
        return calendar_range(preset)
    col_min = frame[active_col].min().normalize()
    col_max = frame[active_col].max().normalize()
    if preset == "All time":
        date_from, date_to = col_min.date(), col_max.date()
    elif preset == "Last 7 days":
        date_from, date_to = (col_max - pd.Timedelta(days=6)).date(), col_max.date()
    elif preset == "Last 30 days":
        date_from, date_to = (col_max - pd.Timedelta(days=29)).date(), col_max.date()
    elif preset == "This month":
        date_from, date_to = col_max.replace(day=1).date(), col_max.date()
    elif preset == "This year":
        date_from, date_to = col_max.replace(month=1, day=1).date(), col_max.date()
    elif preset == "Last year":
        ly = col_max.year - 1
        date_from, date_to = pd.Timestamp(ly, 1, 1).date(), pd.Timestamp(ly, 12, 31).date()
    else:  # Custom range
        try:
            date_from = pd.Timestamp(custom_from).date() if custom_from else col_min.date()
            date_to = pd.Timestamp(custom_to).date() if custom_to else col_max.date()
        except ValueError:
            date_from, date_to = col_min.date(), col_max.date()
    date_from = max(date_from, col_min.date())
    date_to = min(date_to, col_max.date())
    if date_from > date_to:
        date_from, date_to = col_min.date(), col_max.date()
    return date_from, date_to


# --------------------------------------------------------------------------
# Session cache (the file is read once, then every filter change is fast)
# --------------------------------------------------------------------------
_CACHE = {}
_STRS = {}


def _strs(df):
    """df[col].astype(str).str.strip() for the cascading filters - worked out once per file."""
    def get(col):
        k = (id(df), col)
        if k not in _STRS:
            _STRS[k] = df[col].astype(str).str.strip()
        return _STRS[k]
    return get


def _df(path):
    key = (path, os.path.getmtime(path))
    if key not in _CACHE:
        _CACHE.clear(); _STRS.clear()
        df, err = load_data(path)
        if err:
            raise ValueError(
                f"Couldn't read {os.path.basename(path)} - it may be corrupted, or missing one of the "
                f"expected columns.\n\nDetails: {err}")
        _CACHE[key] = df
    return _CACHE[key]


def _num(v):
    v = float(v) if v is not None and not pd.isna(v) else 0.0
    return round(v, 2)


def _s(v):
    return "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v)


NO_TEAM = "No team"


def _team_list(df):
    """Teams found in the control file (A, B, C …) with the team leader that appears most for each."""
    sub = df.loc[df["flag"] == "construction", ["team", "team lider"]]
    if sub.empty:
        sub = df[["team", "team lider"]]
    team = sub["team"].fillna(NO_TEAM)
    out = []
    counts = team.value_counts()
    lead = (sub.assign(team=team).dropna(subset=["team lider"])
            .groupby(["team", "team lider"]).size().reset_index(name="n")
            .sort_values(["team", "n"], ascending=[True, False]).drop_duplicates("team")
            .set_index("team")["team lider"])
    for t in sorted(counts.index, key=lambda x: (x == NO_TEAM, str(x))):
        out.append({"team": str(t), "leader": _s(lead.get(t)), "rows": int(counts[t])})
    return out


def load(path):
    df = _df(path)
    available = {label: col for label, col in DATE_FIELDS.items() if col in df.columns and df[col].notna().any()}
    other = df.loc[df["flag"] == "other", "job"].dropna().drop_duplicates().head(15).tolist()
    return {
        "rows": int(len(df)),
        "name": os.path.basename(path),
        "date_fields": list(available.keys()),
        "default_date_field": COMBINED if COMBINED in available else (list(available)[0] if available else None),
        "presets": PRESETS,
        "other_count": int((df["flag"] == "other").sum()),
        "teams": _team_list(df),
        "other_sample": [_s(x) for x in other],
        "options": {
            "district": sorted(df["district"].dropna().unique().tolist()),
            "project": sorted(df["project"].dropna().unique().tolist()),
            "pid": sorted(df["pid"].dropna().unique().tolist()),
            "pm": sorted(df["project manager"].dropna().unique().tolist()),
            "sourcefile": sorted(df["sourcefile"].dropna().unique().tolist()),
        },
    }


# --------------------------------------------------------------------------
# One call per filter change -> everything the page shows
# --------------------------------------------------------------------------
def view(path, date_field=COMBINED, preset="All time", custom_from=None, custom_to=None,
         districts=None, projects=None, pids=None, pms=None, sourcefiles=None,
         circuits=None, poles=None, granularity="Auto", pid_mode="All", tab=None,
         teams=None, team_leaders=None):
    """tab = the tab on screen ("trends", "pid", "finance"): only that one is worked out.
    No tab = all three, as before."""
    df = _df(path)
    active_date_col = DATE_FIELDS.get(date_field, "combined")
    if active_date_col not in df.columns or not df[active_date_col].notna().any():
        raise ValueError(f"The master file has no usable '{date_field}' dates.")
    date_from, date_to = resolve_date_range(df, active_date_col, preset, custom_from, custom_to)
    date_min = df[active_date_col].min()
    date_max = df[active_date_col].max()
    districts, projects, pids, pms = districts or [], projects or [], pids or [], pms or []
    sourcefiles, circuits, poles_selected = sourcefiles or [], circuits or [], poles or []

    def _distinct(series: pd.Series) -> list:
        return sorted(series.dropna().unique().tolist())

    # Cascading filters: District > Project > PID > Project Manager > Source file > Circuit > Pole
    _scope = df[active_date_col].isna() | (
        (df[active_date_col] >= pd.Timestamp(date_from)) & (df[active_date_col] < pd.Timestamp(date_to) + pd.Timedelta(days=1)))
    options, _m = cascade_options(df, _scope, [("districts", "district"), ("projects", "project"), ("pids", "pid"), ("pms", "project manager"),
                                               ("sourcefiles", "sourcefile"), ("circuits", "circuit"), ("poles", "enid")],
                                  {"districts": districts, "projects": projects, "pids": pids, "pms": pms, "sourcefiles": sourcefiles,
                                   "circuits": circuits, "poles": poles_selected}, strs=_strs(df))
    circuit_options, pole_options = options["circuits"], options["poles"]

    # Apply filters (unchanged)
    mask = pd.Series(True, index=df.index)
    mask &= df["flag"] != "other"
    if districts:
        mask &= df["district"].isin(districts)
    if projects:
        mask &= df["project"].isin(projects)
    if pids:
        mask &= df["pid"].isin(pids)
    if pms:
        mask &= df["project manager"].isin(pms)
    if sourcefiles:
        mask &= df["sourcefile"].isin(sourcefiles)
    if circuits:
        mask &= df["circuit"].isin(circuits)
    if poles_selected:
        mask &= df["enid"].isin(poles_selected)
    fdf = df[mask].copy()
    date_mask = fdf[active_date_col].notna() & (fdf[active_date_col] >= pd.Timestamp(date_from)) & (fdf[active_date_col] < pd.Timestamp(date_to) + pd.Timedelta(days=1))
    fdf_dated = fdf[date_mask].copy()
    gran = auto_granularity(date_from, date_to) if granularity == "Auto" else granularity

    out = {
        "options": options, "circuit_options": circuit_options, "pole_options": pole_options,
        "date_from": str(date_from), "date_to": str(date_to),
        "data_min": str(date_min.date()), "data_max": str(date_max.date()),
        "records": int(len(fdf_dated)), "granularity": gran, "granularity_auto": granularity == "Auto",
        "plan2_comments": plan2_comments(fdf_dated, active_date_col) if active_date_col == "plan2" else None,
    }
    out["tab"] = tab
    if tab in (None, "trends"):
        out["trends"] = _trends(fdf_dated, active_date_col, gran, teams, team_leaders)
    if tab in (None, "pid"):
        out["pid"] = _pid_breakdown(fdf, pid_mode)
    if tab in (None, "finance"):
        out["finance"] = _finance(fdf_dated)
    return out


def _trends(fdf_dated, active_date_col, granularity, teams=None, team_leaders=None):
    construction = fdf_dated[fdf_dated["flag"] == "construction"].copy()
    material = fdf_dated[fdf_dated["flag"] == "material"].copy()
    res = {"panel1": None, "panel2": None}
    if construction.empty and material.empty:
        return res
    construction["bucket"] = bucket_series(construction[active_date_col], granularity)
    material["bucket"] = bucket_series(material[active_date_col], granularity)

    c_agg = construction.groupby("bucket", as_index=False).agg(total=("total", "sum"), orig=("orig", "sum"))
    m_agg = material.groupby("bucket", as_index=False).agg(material=("total", "sum"))
    chart_df = pd.merge(c_agg, m_agg, on="bucket", how="outer").fillna(0)
    chart_df = chart_df.sort_values("bucket")
    chart_df["variance"] = chart_df["total"] - chart_df["orig"]
    chart_df["base"] = chart_df[["total", "orig"]].min(axis=1)
    chart_df["cap"] = chart_df["variance"].abs()
    chart_df["label"] = chart_df["bucket"].apply(lambda ts: bucket_label(ts, granularity))
    shared_buckets = chart_df["bucket"].tolist()
    shared_labels = chart_df["label"].tolist()
    grand_total = chart_df["total"].sum() + chart_df["material"].sum()
    var_sum = chart_df["variance"].sum()
    res["panel1"] = {
        "kpis": {"grand_total": _num(grand_total), "construction": _num(chart_df["total"].sum()),
                 "original": _num(chart_df["orig"].sum()), "variance": _num(var_sum),
                 "materials": _num(chart_df["material"].sum())},
        "labels": shared_labels,
        "base": [_num(v) for v in chart_df["base"]], "cap": [_num(v) for v in chart_df["cap"]],
        "variance": [_num(v) for v in chart_df["variance"]], "material": [_num(v) for v in chart_df["material"]],
        "total": [_num(v) for v in chart_df["total"]], "orig": [_num(v) for v in chart_df["orig"]],
    }

    if construction.empty:
        return res
    # Panel 2 – Teams: the teams picked in the Teams tab (all when none picked),
    # named "A · leader" when a leader was typed for that team
    leaders = {str(k): str(v).strip() for k, v in (team_leaders or {}).items() if v and str(v).strip()}
    construction["team_grp"] = construction["team"].fillna(NO_TEAM).astype(str)
    present = construction.groupby("team_grp")["total"].sum()
    order = sorted(present.index, key=lambda x: (x == NO_TEAM, x))
    if teams:
        wanted = set(map(str, teams))
        order = [t for t in order if t in wanted]
    team_c = construction[construction["team_grp"].isin(order)]
    name = {t: (f"{t} · {leaders[t]}" if t in leaders else t) for t in order}
    pivot = (
        team_c.groupby(["bucket", "team_grp"])["total"].sum()
        .unstack("team_grp", fill_value=0)
        .reindex(index=shared_buckets, fill_value=0)
        .reindex(columns=order, fill_value=0)
    )
    var_sum = team_c["total"].sum() - team_c["orig"].sum()
    res["panel2"] = {
        "kpis": {"total": _num(team_c["total"].sum()), "original": _num(team_c["orig"].sum()), "variance": _num(var_sum)},
        "labels": shared_labels, "leaders": [name[t] for t in order],
        "values": {name[t]: [_num(v) for v in pivot[t].values] for t in order},
    }
    return res


def _pid_breakdown(fdf, mc_mode):
    pdf = fdf.dropna(subset=["pid"]).copy()
    if mc_mode != "All":
        pdf = pdf[pdf["flag"] == mc_mode.lower()]
    if pdf.empty:
        return None
    pdf["has_plan"] = pdf["plan1"].notna()
    pdf["has_done"] = pdf["done"].notna()
    pdf["has_inv"] = pdf["invoice date"].notna()
    pdf["stage"] = np.select(
        [pdf["has_inv"], pdf["has_done"], pdf["has_plan"]],
        ["invoiced", "done", "planned"],
        default="remaining",
    )

    def _group_mode(col):
        """Per PID, the same value as s.mode().iloc[0] (most frequent; ties -> the
        smallest, as mode() sorts; blanks ignored; None if there are none) - worked
        out with one grouped count instead of a Python call per PID."""
        sub = pdf[["pid", col]].dropna(subset=[col])
        if sub.empty:
            return pd.Series(dtype=object)
        cnt = sub.groupby(["pid", col], observed=True).size().reset_index(name="_n")
        cnt = cnt.sort_values(["pid", "_n", col], ascending=[True, False, True], kind="mergesort")
        return cnt.drop_duplicates("pid").set_index("pid")[col]

    base = pdf.groupby("pid", dropna=False).agg(total=("total", "sum")).reset_index()
    for out_col, src_col in (("job", "job"), ("district", "district"), ("project", "project"), ("pm", "project manager")):
        m = _group_mode(src_col)
        base[out_col] = base["pid"].map(m).astype(object).where(base["pid"].isin(m.index), None)
    stage_sums = pdf.groupby(["pid", "stage"], dropna=False)["total"].sum().unstack("stage", fill_value=0)
    for s in ["remaining", "planned", "done", "invoiced"]:
        if s not in stage_sums.columns:
            stage_sums[s] = 0.0
    stage_sums = stage_sums.reset_index()
    agg = base.merge(stage_sums, on="pid", how="left")
    _numeric_fill_cols = ["total", "remaining", "planned", "done", "invoiced"]
    agg[_numeric_fill_cols] = agg[_numeric_fill_cols].fillna(0)

    var_src = fdf.dropna(subset=["pid"])
    var_src = var_src[var_src["flag"] == "construction"]
    var_agg = var_src.groupby("pid", dropna=False).agg(c_total=("total", "sum"), c_orig=("orig", "sum")).reset_index()
    var_agg["variance"] = var_agg["c_total"] - var_agg["c_orig"]
    agg = agg.merge(var_agg[["pid", "variance"]], on="pid", how="left").fillna({"variance": 0.0})
    assert agg["pid"].is_unique

    agg["district_total"] = agg.groupby("district")["total"].transform("sum")
    agg["project_total"] = agg.groupby(["district", "project"])["total"].transform("sum")
    agg["pm_total"] = agg.groupby(["district", "project", "pm"])["total"].transform("sum")
    agg = agg.sort_values(["district_total", "project_total", "pm_total", "total"], ascending=[False, False, False, False])
    agg["pm_label"] = agg["pm"].fillna("Unassigned")
    agg["district_label"] = agg["district"].fillna("Unassigned")
    agg["project_label"] = agg["project"].fillna("Unassigned")

    rows = [{
        "district": _s(r.district_label), "project": _s(r.project_label), "pm": _s(r.pm_label),
        "pid": _s(r.pid), "job": _s(r.job), "total": _num(r.total), "remaining": _num(r.remaining),
        "planned": _num(r.planned), "done": _num(r.done), "invoiced": _num(r.invoiced), "variance": _num(r.variance),
    } for r in agg.itertuples(index=False)]
    return {
        "kpis": {"total": _num(agg["total"].sum()), "remaining": _num(agg["remaining"].sum()),
                 "planned": _num(agg["planned"].sum()), "done": _num(agg["done"].sum()),
                 "invoiced": _num(agg["invoiced"].sum()), "variance": _num(agg["variance"].sum())},
        "rows": rows,
    }


def _finance(fdf_dated):
    fin = fdf_dated.copy()
    fin_c = fin[fin["flag"] == "construction"].copy()
    fin_m = fin[fin["flag"] == "material"].copy()
    if fin.empty:
        return None
    fin["has_plan"] = fin["plan1"].notna()
    fin["has_done"] = fin["done"].notna()
    fin["has_inv"] = fin["invoice date"].notna()

    total_value = fin["total"].sum()
    orig_value = fin_c["orig"].sum()
    variance = fin_c["total"].sum() - orig_value
    variance_pct = (variance / orig_value * 100) if orig_value else 0.0
    invoiced_sum = fin.loc[fin["has_inv"], "total"].sum()
    wip_sum = fin.loc[fin["has_done"] & ~fin["has_inv"], "total"].sum()
    backlog_sum = fin.loc[~fin["has_done"], "total"].sum()
    invoiced_pct = (invoiced_sum / total_value * 100) if total_value else 0.0
    material_pct = (fin_m["total"].sum() / total_value * 100) if total_value else 0.0

    pipeline = [
        {"stage": "Remaining (not started)", "value": _num(fin.loc[~fin["has_plan"], "total"].sum())},
        {"stage": "Planned (not done)", "value": _num(fin.loc[fin["has_plan"] & ~fin["has_done"], "total"].sum())},
        {"stage": "Done (not invoiced)", "value": _num(wip_sum)},
        {"stage": "Invoiced", "value": _num(invoiced_sum)},
    ]

    lag_df = fin[fin["has_done"] & fin["has_inv"]].copy()
    lag_df["lag_days"] = (lag_df["invoice date"] - lag_df["done"]).dt.days
    lag_df = lag_df[(lag_df["lag_days"] >= 0) & (lag_df["lag_days"] <= 365)]
    lag = None
    if not lag_df.empty:
        counts, edges = np.histogram(lag_df["lag_days"], bins=min(30, max(1, lag_df["lag_days"].nunique())))
        lag = {"median": float(lag_df["lag_days"].median()), "mean": float(lag_df["lag_days"].mean()), "n": int(len(lag_df)),
               "counts": [int(c) for c in counts], "edges": [float(e) for e in edges]}

    proj_var = fin_c.groupby("project", dropna=False).agg(total=("total", "sum"), orig=("orig", "sum"))
    proj_var = proj_var[proj_var["orig"] != 0]
    proj_var["variance"] = proj_var["total"] - proj_var["orig"]
    proj_var["variance_pct"] = proj_var["variance"] / proj_var["orig"] * 100
    proj_var = proj_var.sort_values("variance_pct")
    projects = [{"project": _s(i) or "Unassigned", "total": _num(r.total), "orig": _num(r.orig),
                 "variance": _num(r.variance), "pct": round(float(r.variance_pct), 2)}
                for i, r in zip(proj_var.index, proj_var.itertuples(index=False))]

    pid_var = fin_c.dropna(subset=["pid"]).groupby(["pid", "district", "project"], dropna=False).agg(
        total=("total", "sum"), orig=("orig", "sum")).reset_index()
    pid_var = pid_var[pid_var["orig"] != 0]
    pid_var["variance"] = pid_var["total"] - pid_var["orig"]
    pid_var["variance_pct"] = pid_var["variance"] / pid_var["orig"] * 100

    def _lb(frame):
        return [{"pid": _s(r.pid), "district": _s(r.district), "project": _s(r.project),
                 "variance": _num(r.variance), "pct": round(float(r.variance_pct), 1)} for r in frame.itertuples(index=False)]

    return {
        "kpis": {"total": _num(total_value), "original": _num(orig_value), "variance": _num(variance),
                 "variance_pct": round(float(variance_pct), 2), "invoiced": _num(invoiced_sum),
                 "invoiced_pct": round(float(invoiced_pct), 1), "wip": _num(wip_sum), "backlog": _num(backlog_sum),
                 "material_pct": round(float(material_pct), 1)},
        "pipeline": pipeline, "lag": lag, "projects": projects,
        "worst": _lb(pid_var.sort_values("variance").head(8)) if not pid_var.empty else [],
        "best": _lb(pid_var.sort_values("variance", ascending=False).head(8)) if not pid_var.empty else [],
    }