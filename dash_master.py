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
    "job", "plan1", "done", "datetouse", "invoice date", "total", "orig",
    "district", "project", "project manager", "team", "team lider", "pid",
    "circuit", "enid", "sourcefile",
]
DATE_FIELDS = {"Plan1": "plan1", "Done": "done", "DateToUse": "datetouse"}
PRESETS = ["All time", "Last 7 days", "Last 30 days", "This month", "This year", "Last year", "Custom range"]


# --------------------------------------------------------------------------
# Data loading (unchanged)
# --------------------------------------------------------------------------
def load_data(path_str: str):
    """Returns (df, error_message). Never raises."""
    try:
        df = pd.read_parquet(path_str, columns=COLUMNS_NEEDED)
    except Exception as e:
        return None, str(e)

    # total / orig sometimes land as text — coerce to numeric
    df["total"] = pd.to_numeric(df["total"], errors="coerce")
    df["orig"] = pd.to_numeric(df["orig"], errors="coerce")

    for c in ["plan1", "done", "datetouse", "invoice date"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    for c in ["district", "project", "project manager", "team", "team lider", "pid", "job", "circuit", "enid", "sourcefile"]:
        df[c] = df[c].astype("string").str.strip()
        df[c] = df[c].replace("", pd.NA)

    job_lower = df["job"].str.lower()
    df["flag"] = "other"
    df.loc[job_lower.str.startswith(("m -", "m-"), na=False), "flag"] = "material"
    df.loc[job_lower.str.startswith(("c -", "c-"), na=False), "flag"] = "construction"

    for c in ["plan1", "done", "datetouse"]:
        df[c] = df[c].where(df[c].dt.year > 1901)

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


def _df(path):
    key = (path, os.path.getmtime(path))
    if key not in _CACHE:
        _CACHE.clear()
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


def load(path):
    df = _df(path)
    available = {label: col for label, col in DATE_FIELDS.items() if col in df.columns and df[col].notna().any()}
    other = df.loc[df["flag"] == "other", "job"].dropna().drop_duplicates().head(15).tolist()
    return {
        "rows": int(len(df)),
        "name": os.path.basename(path),
        "date_fields": list(available.keys()),
        "default_date_field": "DateToUse" if "DateToUse" in available else (list(available)[0] if available else None),
        "presets": PRESETS,
        "other_count": int((df["flag"] == "other").sum()),
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
def view(path, date_field="DateToUse", preset="All time", custom_from=None, custom_to=None,
         districts=None, projects=None, pids=None, pms=None, sourcefiles=None,
         circuits=None, poles=None, granularity="Auto", pid_mode="All"):
    df = _df(path)
    active_date_col = DATE_FIELDS.get(date_field, "datetouse")
    if active_date_col not in df.columns or not df[active_date_col].notna().any():
        raise ValueError(f"The master file has no usable '{date_field}' dates.")
    date_from, date_to = resolve_date_range(df, active_date_col, preset, custom_from, custom_to)
    date_min = df[active_date_col].min()
    date_max = df[active_date_col].max()
    districts, projects, pids, pms = districts or [], projects or [], pids or [], pms or []
    sourcefiles, circuits, poles_selected = sourcefiles or [], circuits or [], poles or []

    def _distinct(series: pd.Series) -> list:
        return sorted(series.dropna().unique().tolist())

    # Cascading Circuit / Pole options (unchanged)
    _scope_mask = pd.Series(True, index=df.index)
    if districts:
        _scope_mask &= df["district"].isin(districts)
    if projects:
        _scope_mask &= df["project"].isin(projects)
    if pids:
        _scope_mask &= df["pid"].isin(pids)
    if pms:
        _scope_mask &= df["project manager"].isin(pms)
    if sourcefiles:
        _scope_mask &= df["sourcefile"].isin(sourcefiles)
    _scope_mask &= df[active_date_col].isna() | (
        (df[active_date_col] >= pd.Timestamp(date_from)) & (df[active_date_col] < pd.Timestamp(date_to) + pd.Timedelta(days=1))
    )
    circuit_options = _distinct(df.loc[_scope_mask, "circuit"])
    _scope_mask_for_pole = _scope_mask.copy()
    if circuits:
        _scope_mask_for_pole &= df["circuit"].isin(circuits)
    pole_options = _distinct(df.loc[_scope_mask_for_pole, "enid"])

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
        "circuit_options": circuit_options, "pole_options": pole_options,
        "date_from": str(date_from), "date_to": str(date_to),
        "data_min": str(date_min.date()), "data_max": str(date_max.date()),
        "records": int(len(fdf_dated)), "granularity": gran, "granularity_auto": granularity == "Auto",
    }
    out["trends"] = _trends(fdf_dated, active_date_col, gran)
    out["pid"] = _pid_breakdown(fdf, pid_mode)
    out["finance"] = _finance(fdf_dated)
    return out


def _trends(fdf_dated, active_date_col, granularity):
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
    totals_by_lider = (
        construction.assign(lider=construction["team lider"].fillna("Unassigned"))
        .groupby("lider")["total"].sum()
        .sort_values(ascending=False)
    )
    TOP_N = 6
    top_liders = list(totals_by_lider.index[:TOP_N])
    has_other = len(totals_by_lider) > TOP_N
    construction["lider_grp"] = construction["team lider"].fillna("Unassigned")
    if has_other:
        construction["lider_grp"] = construction["lider_grp"].where(construction["lider_grp"].isin(top_liders), "Other")
    lider_order = top_liders + (["Other"] if has_other else [])
    lider_pivot = (
        construction.groupby(["bucket", "lider_grp"])["total"].sum()
        .unstack("lider_grp", fill_value=0)
        .reindex(index=shared_buckets, fill_value=0)
        .reindex(columns=lider_order, fill_value=0)
    )
    var_sum = construction["total"].sum() - construction["orig"].sum()
    res["panel2"] = {
        "kpis": {"total": _num(construction["total"].sum()), "original": _num(construction["orig"].sum()), "variance": _num(var_sum)},
        "labels": shared_labels, "leaders": [str(l) for l in lider_order], "has_other": has_other, "top_n": TOP_N,
        "values": {str(l): [_num(v) for v in lider_pivot[l].values] for l in lider_order},
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

    def _mode(s):
        m = s.mode()
        return m.iloc[0] if not m.empty else None

    base = pdf.groupby("pid", dropna=False).agg(
        total=("total", "sum"),
        job=("job", _mode),
        district=("district", _mode),
        project=("project", _mode),
        pm=("project manager", _mode),
    ).reset_index()
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
