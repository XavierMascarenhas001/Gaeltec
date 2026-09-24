"""
Master Control — Job Costing Dashboard
Streamlit app. Auto-loads the latest Master_DD-MM-YYYY (vN).parquet file
from the network share - no upload needed.
 
Run locally (this needs access to the \\gaeltec-gl share, so it must run
on a machine on-domain/VPN - NOT Streamlit Community Cloud, which has no
access to internal UNC paths):
    pip install -r requirements.txt
    streamlit run app.py
"""
 
import datetime as dt
import re
from pathlib import Path
 
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
 
# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
 
st.set_page_config(
    page_title="Master Control — Job Costing Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)
 
# Same folder/naming convention as the Network Job Tracker dashboard -
# adjust this if Master Control actually reads from a different file.
MASTER_DIR = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker"
# Version suffix "(vN)" is OPTIONAL - a plain "Master_01-09-2026.parquet"
# with no version tag is a perfectly valid file (treated as v0) and must
# still be picked up, not silently skipped.
MASTER_PATTERN = re.compile(r"Master_(\d{2})-(\d{2})-(\d{4})(?:\s*\(v(\d+)\))?\.parquet$", re.IGNORECASE)
 
COLUMNS_NEEDED = [
    "job", "plan1", "done", "datetouse", "invoice date", "total", "orig",
    "district", "project", "project manager", "team", "team lider", "pid",
    "circuit", "enid", "sourcefile",
]
 
COLOR_GREEN = "#1E9E5A"      # positive variation
COLOR_RED = "#D64545"        # negative variation / PID tab — remaining
COLOR_YELLOW = "#F0B429"     # materials
COLOR_NAVY = "#1F3A5F"       # "the rest" / base (construction, Panel 1)
COLOR_DARKBLUE = "#123A6B"   # PID tab — done
COLOR_TEAL = "#2C7DA0"       # PID tab — total (KPI accent)
COLOR_AMBER = "#E0A02A"      # PID tab — planned
COLOR_PURPLE = "#7C5CBF"     # PID tab — invoiced
COLOR_MUTED = "#B7BFC9"
 
CATEGORY_PALETTE = [
    "#2C7DA0", "#E0A02A", "#7C5CBF", "#1E9E5A", "#D64545",
    "#2F9E9E", "#A85CB0", "#7C8C2A", "#C25C82", "#4C6FC2",
]
 
TEXT_DARK = "#1B2430"
TEXT_MUTED = "#5B6B7C"
GRID_LIGHT = "#E2E7ED"
PANEL_BG = "#FFFFFF"
 
PLOTLY_LIGHT = dict(
    paper_bgcolor=PANEL_BG,
    plot_bgcolor=PANEL_BG,
    font=dict(color=TEXT_DARK, family="IBM Plex Mono, monospace", size=13),
)
 
 
# --------------------------------------------------------------------------
# Data loading
# --------------------------------------------------------------------------
 
def find_all_masters(directory: str):
    """Scans `directory` for Master_DD-MM-YYYY[ (vN)].parquet files and
    returns a list of (date, version, Path) tuples sorted newest-first -
    date takes priority over version, so a same-day higher version only
    matters as a tiebreaker, never overrides a genuinely later date.
    A missing version suffix counts as v0. Returns [] if the folder is
    unreachable or nothing matches - never raises."""
    found = []
    try:
        candidates = list(Path(directory).glob("Master_*.parquet"))
    except OSError:
        return found
    for path in candidates:
        m = MASTER_PATTERN.search(path.name)
        if not m:
            continue
        dd, mm, yyyy, ver = m.groups()
        try:
            date_key = dt.datetime(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        found.append((date_key, int(ver or 0), path))
    found.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return found


def find_latest_master(directory: str):
    """Convenience wrapper: just the single newest Path, or None."""
    all_masters = find_all_masters(directory)
    return all_masters[0][2] if all_masters else None
 
 
@st.cache_data(show_spinner="Loading Master Control data…", max_entries=2, ttl=300)
def load_data(path_str: str):
    """Returns (df, error_message). Never raises - a locked/corrupted file
    shows a clean st.error rather than crashing the app.
    ttl=300 (5 min): the file list can change between sessions - use the
    sidebar "Refresh from network" button for an immediate re-check."""
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
    # NOTE: rows that don't match either prefix ("other") are now KEPT here, not dropped.
    # Dropping them at this point used to silently shrink the date range shown in the
    # sidebar date picker whenever newer jobs didn't follow the "M -" / "C -" convention —
    # the picker's min/max were computed AFTER those rows were already gone. They're still
    # excluded from every chart downstream (via the flag filter applied to fdf), but they
    # now count toward the min/max date bounds, so the calendar always reflects the true
    # range of dates in the uploaded file.
 
    # plan1/done/datetouse sometimes carry a placeholder ~1900 date for "no
    # date" - treat as missing on all three now that any of them can be
    # picked as the active date field (previously this was only applied
    # to datetouse, since it was the only one ever used for filtering).
    for c in ["plan1", "done", "datetouse"]:
        df[c] = df[c].where(df[c].dt.year > 1901)
 
    return df, None
 
 
# --------------------------------------------------------------------------
# Formatting helpers
# --------------------------------------------------------------------------
 
def fmt_money(v: float) -> str:
    if pd.isna(v):
        v = 0.0
    sign = "-" if v < 0 else ""
    return f"{sign}£{abs(v):,.2f}"
 
 
def fmt_money_short(v: float) -> str:
    if pd.isna(v):
        v = 0.0
    sign = "-" if v < 0 else ""
    a = abs(v)
    if a >= 1_000_000:
        return f"{sign}£{a/1_000_000:.1f}m"
    if a >= 1_000:
        return f"{sign}£{a/1_000:.1f}k"
    return f"{sign}£{a:.0f}"
 
 
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
 
 
def metric_row(items):
    """items: list of (label, value_str, accent_color_or_None, highlight_bool)"""
    cols = st.columns(len(items))
    for col, (label, value, color, highlight) in zip(cols, items):
        with col:
            if highlight:
                box_bg = "#EAF3EC"
                box_border = color or COLOR_NAVY
                value_size = "26px"
                label_color = TEXT_DARK
                border_style = f"border:1px solid {box_border};border-left:5px solid {box_border};"
            else:
                box_bg = "#F6F8FA"
                value_size = "18px"
                label_color = TEXT_MUTED
                border_style = "border:1px solid #DDE3E9;"
            val_style = f"color:{color};" if color else f"color:{TEXT_DARK};"
            st.markdown(
                f"""
                <div style="{border_style}border-radius:6px;padding:10px 14px;background:{box_bg};">
                  <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.08em;
                              text-transform:uppercase;color:{label_color};margin-bottom:4px;">{label}</div>
                  <div style="font-family:'IBM Plex Mono',monospace;font-size:{value_size};font-weight:700;{val_style}">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
 
 
# --------------------------------------------------------------------------
# Sidebar — data source + filters
# --------------------------------------------------------------------------
 
st.sidebar.title("Master Control")
st.sidebar.caption("Job costing & progress — construction / materials")
 
col_a, col_b = st.sidebar.columns(2)
with col_a:
    if st.button("🔄 Refresh", help="Re-scan for the latest Master file."):
        st.cache_data.clear()
        (st.rerun if hasattr(st, "rerun") else st.experimental_rerun)()
with col_b:
    if st.button("🧹 Clear cache", help="Frees memory by dropping every cached dataframe."):
        st.cache_data.clear()
        st.success("Cleared.")
        (st.rerun if hasattr(st, "rerun") else st.experimental_rerun)()
 
all_masters = find_all_masters(MASTER_DIR)

if not all_masters:
    st.title("Master Control — Job Costing Dashboard")
    st.error(
        f"Couldn't find any `Master_*.parquet` file in:\n\n`{MASTER_DIR}`\n\n"
        "Check the folder path is correct and that this machine has access to the network share "
        "(this won't work on Streamlit Community Cloud - only when run on a machine with access "
        "to the \\\\gaeltec-gl share, e.g. locally or on an internal server)."
    )
    st.stop()

_master_labels = [
    f"{p.name}" + ("  ← newest" if i == 0 else "")
    for i, (_, _, p) in enumerate(all_masters)
]
_master_choice = st.sidebar.selectbox(
    "📄 Master file",
    options=range(len(all_masters)),
    format_func=lambda i: _master_labels[i],
    index=0,
    help="Newest by date first, then by version. Pick a different one if the auto-detected file isn't the one you want.",
)
latest_master_path = all_masters[_master_choice][2]
 
df, load_err = load_data(str(latest_master_path))
if load_err:
    st.title("Master Control — Job Costing Dashboard")
    st.error(
        f"Couldn't read `{latest_master_path.name}` - it may be corrupted, locked while being "
        f"written, or missing one of the expected columns.\n\n**Details:** {load_err}"
    )
    st.stop()
 
other_count = int((df["flag"] == "other").sum())
if other_count:
    with st.sidebar.expander(f"⚠️ {other_count:,} row(s) not classified"):
        st.caption(
            "These jobs don't start with 'M -' / 'M-' (material) or 'C -' / 'C-' "
            "(construction), so they're excluded from every chart below. Their dates "
            "still count toward the date-range picker's min/max. Sample job codes:"
        )
        sample = df.loc[df["flag"] == "other", "job"].dropna().drop_duplicates().head(15).tolist()
        st.write(sample)
 
def pick_date_field_and_range(frame: pd.DataFrame, field_options: dict, default_label: str = None, key_prefix: str = ""):
    """Lets the user choose WHICH date column to filter/bucket by (e.g.
    Plan1 / Done / DateToUse), then pick a range against it - either via
    a quick preset (fewer clicks for the common cases) or a custom
    calendar range. `field_options` is {display label: column name},
    filtered here to only the ones that actually exist with real data.
    Presets are anchored to the LATEST DATE ACTUALLY IN THE DATA, not
    today's real-world date - this is planning/completion data that may
    not extend to today, so "Last 30 days" means the last 30 days of
    data present, which is what's actually useful here.
    Returns (active_column_name, date_from, date_to)."""
    available = {label: col for label, col in field_options.items() if col in frame.columns and frame[col].notna().any()}
    if not available:
        return None, None, None

    labels = list(available.keys())
    default_index = labels.index(default_label) if default_label in labels else 0
    chosen_label = st.sidebar.radio("Date field", labels, index=default_index, horizontal=True, key=f"{key_prefix}date_field")
    active_col = available[chosen_label]

    col_min = frame[active_col].min().normalize()
    col_max = frame[active_col].max().normalize()

    preset = st.sidebar.selectbox(
        "Quick range",
        ["All time", "Last 7 days", "Last 30 days", "This month", "This year", "Last year", "Custom range"],
        key=f"{key_prefix}date_preset",
    )

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
        date_range = st.sidebar.date_input(
            "Custom date range", value=(col_min.date(), col_max.date()),
            min_value=col_min.date(), max_value=col_max.date(), key=f"{key_prefix}date_custom",
        )
        if isinstance(date_range, tuple) and len(date_range) == 2:
            date_from, date_to = date_range
        else:
            date_from, date_to = col_min.date(), col_max.date()

    # Clip to the data's own range so a preset never produces an empty
    # or invalid range (e.g. "Last year" when the data doesn't go back
    # that far).
    date_from = max(date_from, col_min.date())
    date_to = min(date_to, col_max.date())
    if date_from > date_to:
        date_from, date_to = col_min.date(), col_max.date()

    st.sidebar.caption(f"Showing: {date_from} -> {date_to}  (anchored to the latest date in the data, not today)")
    return active_col, date_from, date_to


st.sidebar.markdown("### Filters")

active_date_col, date_from, date_to = pick_date_field_and_range(
    df, {"Plan1": "plan1", "Done": "done", "DateToUse": "datetouse"}, default_label="DateToUse",
)
date_min = df[active_date_col].min()
date_max = df[active_date_col].max()

districts = st.sidebar.multiselect("District", sorted(df["district"].dropna().unique()))
projects = st.sidebar.multiselect("Project", sorted(df["project"].dropna().unique()))
pids = st.sidebar.multiselect("PID", sorted(df["pid"].dropna().unique()))
pms = st.sidebar.multiselect("Project Manager", sorted(df["project manager"].dropna().unique()))
sourcefiles = st.sidebar.multiselect("Source file", sorted(df["sourcefile"].dropna().unique()))


def _distinct(series: pd.Series) -> list:
    return sorted(series.dropna().unique().tolist())


# Circuit and Pole (enid) are CASCADING, matching the Materials dashboard:
# Circuit's options narrow to what's actually reachable given
# District/Project/Date/PID/Project Manager/Source file so far, and
# Pole's options narrow further given Circuit too - so neither ever
# offers a value that couldn't possibly match what's already selected.
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
circuits = st.sidebar.multiselect(
    "Circuit", circuit_options,
    help="Narrowed to circuits that actually exist given the other filters selected above.",
)

_scope_mask_for_pole = _scope_mask.copy()
if circuits:
    _scope_mask_for_pole &= df["circuit"].isin(circuits)
pole_options = _distinct(df.loc[_scope_mask_for_pole, "enid"])
poles_selected = st.sidebar.multiselect(
    "Pole (enid)", pole_options,
    help="Narrowed to poles that actually exist given the other filters (including Circuit) selected above.",
)

gran_choice = st.sidebar.selectbox("Date grouping", ["Auto", "Day", "Week", "Month", "Year"])

if st.sidebar.button("Reset filters"):
    st.rerun()

# --------------------------------------------------------------------------
# Apply filters
# --------------------------------------------------------------------------

mask = pd.Series(True, index=df.index)
mask &= df["flag"] != "other"  # charts only ever show material/construction, same as before
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
 
granularity = auto_granularity(date_from, date_to) if gran_choice == "Auto" else gran_choice
 
st.sidebar.markdown("---")
st.sidebar.markdown(
    f"<span style='font-family:monospace;font-size:12px;color:#7C8AA0;'>"
    f"RECORDS IN VIEW: <b style='color:#E3A64A;'>{len(fdf_dated):,}</b></span>",
    unsafe_allow_html=True,
)
 
# --------------------------------------------------------------------------
# Header
# --------------------------------------------------------------------------
 
st.title("Master Control — Job Costing Dashboard")
st.caption(
    f"Source: {latest_master_path.name} · {date_min.date()} → {date_max.date()} · "
    f"grouping: {granularity}{' (auto)' if gran_choice == 'Auto' else ''}"
)
 
tab_trend, tab_pid, tab_finance = st.tabs(["📈 Trends", "📊 PID Breakdown", "💰 Finance"])
 
# --------------------------------------------------------------------------
# TAB 1 — Trends
# --------------------------------------------------------------------------
 
with tab_trend:
    st.subheader("Panel 01 — Total by Date (Construction + Materials)")
    st.caption(
        "Each bar = base (navy) + variance vs. original (red/green) + materials (yellow), stacked"
    )
 
    construction = fdf_dated[fdf_dated["flag"] == "construction"].copy()
    material = fdf_dated[fdf_dated["flag"] == "material"].copy()
 
    if construction.empty and material.empty:
        st.info("No records under these filters. Widen the date range or clear a filter.")
    else:
        construction["bucket"] = bucket_series(construction[active_date_col], granularity)
        material["bucket"] = bucket_series(material[active_date_col], granularity)
 
        c_agg = construction.groupby("bucket", as_index=False).agg(total=("total", "sum"), orig=("orig", "sum"))
        m_agg = material.groupby("bucket", as_index=False).agg(material=("total", "sum"))
 
        chart_df = pd.merge(c_agg, m_agg, on="bucket", how="outer").fillna(0)
        chart_df = chart_df.sort_values("bucket")
        chart_df["variance"] = chart_df["total"] - chart_df["orig"]
        # base = the smaller of total/orig ("the rest"); cap = the gap between them, coloured by sign
        chart_df["base"] = chart_df[["total", "orig"]].min(axis=1)
        chart_df["cap"] = chart_df["variance"].abs()
        chart_df["cap_color"] = chart_df["variance"].apply(lambda v: COLOR_GREEN if v >= 0 else COLOR_RED)
        chart_df["label"] = chart_df["bucket"].apply(lambda ts: bucket_label(ts, granularity))
 
        shared_buckets = chart_df["bucket"].tolist()
        shared_labels = chart_df["label"].tolist()
 
        grand_total = chart_df["total"].sum() + chart_df["material"].sum()
        var_sum = chart_df["variance"].sum()
 
        metric_row([
            ("Total (construction + materials)", fmt_money(grand_total), COLOR_NAVY, True),
            ("Construction total", fmt_money(chart_df["total"].sum()), None, False),
            ("Original", fmt_money(chart_df["orig"].sum()), None, False),
            ("Variance", fmt_money(var_sum), COLOR_GREEN if var_sum >= 0 else COLOR_RED, False),
            ("Materials", fmt_money(chart_df["material"].sum()), COLOR_YELLOW, False),
        ])
 
        fig = go.Figure()
 
        fig.add_trace(go.Bar(
            x=chart_df["label"], y=chart_df["base"],
            marker_color=COLOR_NAVY,
            name="Base (original / the rest)",
            customdata=chart_df[["total", "orig", "variance", "material"]],
            hovertemplate="<b>%{x}</b><br>Base: £%{y:,.2f}<extra></extra>",
        ))
        fig.add_trace(go.Bar(
            x=chart_df["label"], y=chart_df["cap"],
            marker_color=chart_df["cap_color"],
            name="Variance (total − original)",
            hovertemplate="<b>%{x}</b><br>Variance: £%{y:,.2f}<extra></extra>",
            showlegend=False,
        ))
        fig.add_trace(go.Bar(
            x=chart_df["label"], y=chart_df["material"],
            marker_color=COLOR_YELLOW,
            marker_line=dict(color="#B98600", width=0.6),
            name="Materials",
            hovertemplate="<b>%{x}</b><br>Materials: £%{y:,.2f}<extra></extra>",
        ))
        # dummy traces purely to add red/green to the legend (variance trace itself is colour-mixed)
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=COLOR_GREEN, name="Positive variation"))
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=COLOR_RED, name="Negative variation"))
 
        fig.update_layout(
            **PLOTLY_LIGHT,
            height=400,
            barmode="stack",
            yaxis=dict(tickprefix="£", gridcolor=GRID_LIGHT, zeroline=False, tickfont=dict(size=12)),
            xaxis=dict(
                gridcolor=GRID_LIGHT, tickfont=dict(size=12),
                categoryorder="array", categoryarray=shared_labels,
            ),
            legend=dict(orientation="h", y=-0.18, bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
            bargap=0.25,
            margin=dict(l=10, r=10, t=30, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)
 
    st.markdown("---")
    st.subheader("Panel 02 — Total by Team Leader (Construction)")
    st.caption("Stacked by team leader · same filters, date grouping & x-axis as Panel 01")
 
    if construction.empty:
        st.info("No records under these filters.")
    else:
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
            construction["lider_grp"] = construction["lider_grp"].where(
                construction["lider_grp"].isin(top_liders), "Other"
            )
 
        lider_order = top_liders + (["Other"] if has_other else [])
        color_map = {l: (COLOR_MUTED if l == "Other" else CATEGORY_PALETTE[i % len(CATEGORY_PALETTE)])
                     for i, l in enumerate(lider_order)}
 
        def _short_name(name, n=16):
            return name if len(name) <= n else name[:n - 1] + "…"
 
        short_label = {l: (l if l == "Other" else _short_name(l)) for l in lider_order}
 
        # pivot onto the SAME bucket list & order used by Panel 01, so both x-axes match exactly
        lider_pivot = (
            construction.groupby(["bucket", "lider_grp"])["total"].sum()
            .unstack("lider_grp", fill_value=0)
            .reindex(index=shared_buckets, fill_value=0)
            .reindex(columns=lider_order, fill_value=0)
        )
 
        var_sum = construction["total"].sum() - construction["orig"].sum()
        metric_row([
            ("Total (construction)", fmt_money(construction["total"].sum()), COLOR_NAVY, True),
            ("Original", fmt_money(construction["orig"].sum()), None, False),
            ("Variance", fmt_money(var_sum), COLOR_GREEN if var_sum >= 0 else COLOR_RED, False),
        ])
 
        fig2 = go.Figure()
        for l in lider_order:
            fig2.add_trace(go.Bar(
                x=shared_labels,
                y=lider_pivot[l].values,
                name=short_label[l],
                marker_color=color_map[l],
                hovertemplate=f"<b>{l}</b><br>%{{x}}<br>£%{{y:,.2f}}<extra></extra>",
            ))
        fig2.update_layout(
            **PLOTLY_LIGHT,
            height=380,
            barmode="stack",
            yaxis=dict(tickprefix="£", gridcolor=GRID_LIGHT, zeroline=False, tickfont=dict(size=12)),
            xaxis=dict(
                gridcolor=GRID_LIGHT, tickfont=dict(size=12),
                categoryorder="array", categoryarray=shared_labels,
            ),
            legend=dict(
                orientation="h", y=-0.22, x=0.5, xanchor="center",
                bgcolor="rgba(0,0,0,0)", font=dict(size=11), tracegroupgap=4,
            ),
            bargap=0.25,
            margin=dict(l=10, r=10, t=30, b=10),
        )
        st.plotly_chart(fig2, use_container_width=True)
        if has_other:
            st.caption(
                "Top " + str(TOP_N) + " team leaders shown individually — the rest are grouped under **Other**."
            )
 
# --------------------------------------------------------------------------
# TAB 2 — PID Breakdown
# --------------------------------------------------------------------------
 
with tab_pid:
    st.subheader("Panel 03 — PID Breakdown")
    st.caption("District → Project → Project Manager → PID, one bar per PID · ordered by total value, highest first")
 
    mc_mode = st.radio("Show", ["All", "Construction", "Material"], horizontal=True)
 
    pdf = fdf.dropna(subset=["pid"]).copy()
    if mc_mode != "All":
        pdf = pdf[pdf["flag"] == mc_mode.lower()]
 
    if pdf.empty:
        st.info("No PIDs under these filters.")
    else:
        pdf["has_plan"] = pdf["plan1"].notna()
        pdf["has_done"] = pdf["done"].notna()
        pdf["has_inv"] = pdf["invoice date"].notna()
 
        # mutually exclusive progress stage per row, so the four segments sum exactly to "total"
        pdf["stage"] = np.select(
            [pdf["has_inv"], pdf["has_done"], pdf["has_plan"]],
            ["invoiced", "done", "planned"],
            default="remaining",
        )
 
        def _mode(s):
            m = s.mode()
            return m.iloc[0] if not m.empty else None
 
        # group by PID alone so every PID gets exactly one row/bar — a PID that touched more
        # than one district/project/PM in the raw data is attributed to its most common one
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
        # Only fill the numeric columns with 0 - filling the whole frame (including
        # the "string" dtype columns job/district/project/pm from base) raises
        # "Cannot set non-string value '0' into a StringArray" on newer pandas,
        # which enforces strict typing on the nullable string dtype.
        _numeric_fill_cols = ["total", "remaining", "planned", "done", "invoiced"]
        agg[_numeric_fill_cols] = agg[_numeric_fill_cols].fillna(0)
 
        # variance always reflects construction budget (total vs original), independent of the Show toggle
        var_src = fdf.dropna(subset=["pid"])
        var_src = var_src[var_src["flag"] == "construction"]
        var_agg = var_src.groupby("pid", dropna=False).agg(
            c_total=("total", "sum"), c_orig=("orig", "sum")
        ).reset_index()
        var_agg["variance"] = var_agg["c_total"] - var_agg["c_orig"]
        agg = agg.merge(var_agg[["pid", "variance"]], on="pid", how="left").fillna({"variance": 0.0})
 
        assert agg["pid"].is_unique  # one row = one bar, guaranteed
 
        # order: district money desc -> project money desc -> project manager money desc -> pid money desc
        agg["district_total"] = agg.groupby("district")["total"].transform("sum")
        agg["project_total"] = agg.groupby(["district", "project"])["total"].transform("sum")
        agg["pm_total"] = agg.groupby(["district", "project", "pm"])["total"].transform("sum")
        agg = agg.sort_values(
            ["district_total", "project_total", "pm_total", "total"],
            ascending=[False, False, False, False],
        )
 
        agg["pm_label"] = agg["pm"].fillna("Unassigned")
        agg["district_label"] = agg["district"].fillna("Unassigned")
        agg["project_label"] = agg["project"].fillna("Unassigned")
 
        var_sum_all = agg["variance"].sum()
        metric_row([
            ("Total", fmt_money(agg["total"].sum()), COLOR_NAVY, True),
            ("Remaining", fmt_money(agg["remaining"].sum()), COLOR_MUTED, False),
            ("Planned", fmt_money(agg["planned"].sum()), COLOR_AMBER, False),
            ("Done", fmt_money(agg["done"].sum()), COLOR_DARKBLUE, False),
            ("Invoiced", fmt_money(agg["invoiced"].sum()), COLOR_PURPLE, False),
            ("Variance", fmt_money(var_sum_all), COLOR_GREEN if var_sum_all >= 0 else COLOR_RED, False),
        ])
        st.caption(f"{len(agg)} PID(s) in view — Remaining/Planned/Done/Invoiced sum to Total; Variance is total vs. original budget")
 
        # ---- render: plain HTML/CSS grouped bars (District > Project > PM > PID) ----
        # Plotly's multicategory axis proved unreliable at this row count/depth — this is
        # deterministic: exactly one proportional bar row per PID, grouped under nested headers.
 
        def esc(s):
            return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
 
        STAGE_COLS = [
            ("invoiced", "Invoiced", COLOR_PURPLE),
            ("done", "Done", COLOR_DARKBLUE),
            ("planned", "Planned", COLOR_AMBER),
            ("remaining", "Remaining", COLOR_MUTED),
        ]
 
        # hover CSS: custom value tooltips on each bar segment + expand-on-hover for
        # truncated job names (pure CSS, no JS needed — works inside st.markdown)
        pid_hover_css = f"""
        <style>
        .pid-tooltip {{ position: relative; }}
        .pid-tooltip .tt-box {{
            visibility: hidden; opacity: 0; transition: opacity .12s ease;
            position: absolute; bottom: 130%; left: 50%; transform: translateX(-50%);
            background: {TEXT_DARK}; color: #fff; padding: 4px 8px; border-radius: 4px;
            font-family: "IBM Plex Mono", monospace; font-size: 11px; white-space: nowrap;
            z-index: 50; pointer-events: none; box-shadow: 0 2px 6px rgba(0,0,0,.25);
        }}
        .pid-tooltip .tt-box::after {{
            content: ""; position: absolute; top: 100%; left: 50%; transform: translateX(-50%);
            border: 5px solid transparent; border-top-color: {TEXT_DARK};
        }}
        .pid-tooltip:hover .tt-box {{ visibility: visible; opacity: 1; }}
 
        .job-name-wrap {{ position: relative; }}
        .job-name-full {{
            font-family: "IBM Plex Sans", sans-serif; font-size: 9.5px; color: {TEXT_MUTED};
            width: 230px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
            cursor: default;
        }}
        .job-name-wrap:hover .job-name-full {{
            position: absolute; left: 0; top: 16px; width: 230px;
            white-space: normal; word-break: break-word; overflow: visible;
            background: #FFFDE7; border: 1px solid #E6D98A; border-radius: 4px;
            padding: 4px 6px; box-shadow: 0 3px 10px rgba(0,0,0,.18); z-index: 40;
        }}
        </style>
        """
        st.markdown(pid_hover_css, unsafe_allow_html=True)
 
        legend_html = "".join(
            f"<span style='display:inline-flex;align-items:center;gap:5px;margin-right:16px;"
            f"font-family:\"IBM Plex Mono\",monospace;font-size:11px;color:{TEXT_MUTED};'>"
            f"<span style='width:11px;height:11px;border-radius:3px;background:{color};display:inline-block;'></span>"
            f"{label}</span>"
            for _, label, color in STAGE_COLS
        )
        legend_html += (
            f"<span style='display:inline-flex;align-items:center;gap:5px;margin-right:16px;"
            f"font-family:\"IBM Plex Mono\",monospace;font-size:11px;color:{TEXT_MUTED};'>"
            f"<span style='width:9px;height:9px;border-radius:50%;background:{COLOR_GREEN};display:inline-block;'></span>"
            f"Over value</span>"
            f"<span style='display:inline-flex;align-items:center;gap:5px;"
            f"font-family:\"IBM Plex Mono\",monospace;font-size:11px;color:{TEXT_MUTED};'>"
            f"<span style='width:9px;height:9px;border-radius:50%;background:{COLOR_RED};display:inline-block;'></span>"
            f"Under value</span>"
        )
        st.markdown(f"<div style='margin:6px 0 14px;'>{legend_html}</div>", unsafe_allow_html=True)
 
        rows_html = []
        last_district, last_project, last_pm = None, None, None
 
        for _, r in agg.iterrows():
            if r["district_label"] != last_district:
                d_total = agg.loc[agg["district_label"] == r["district_label"], "total"].sum()
                rows_html.append(
                    f"<div style='background:{COLOR_NAVY};color:#fff;padding:8px 14px;"
                    f"font-family:\"IBM Plex Mono\",monospace;font-size:13px;font-weight:700;"
                    f"letter-spacing:.04em;text-transform:uppercase;margin-top:14px;border-radius:4px;"
                    f"display:flex;justify-content:space-between;'>"
                    f"<span>{esc(r['district_label'])}</span><span>{fmt_money(d_total)}</span></div>"
                )
                last_district, last_project, last_pm = r["district_label"], None, None
 
            if r["project_label"] != last_project:
                p_total = agg.loc[
                    (agg["district_label"] == r["district_label"]) & (agg["project_label"] == r["project_label"]),
                    "total",
                ].sum()
                rows_html.append(
                    f"<div style='background:#EEF1F5;color:{TEXT_DARK};padding:6px 14px 6px 26px;"
                    f"font-family:\"IBM Plex Mono\",monospace;font-size:11.5px;font-weight:600;"
                    f"letter-spacing:.03em;text-transform:uppercase;margin-top:6px;border-radius:3px;"
                    f"display:flex;justify-content:space-between;'>"
                    f"<span>{esc(r['project_label'])}</span><span>{fmt_money(p_total)}</span></div>"
                )
                last_project, last_pm = r["project_label"], None
 
            if r["pm_label"] != last_pm:
                rows_html.append(
                    f"<div style='padding:4px 14px 2px 40px;font-family:\"IBM Plex Sans\",sans-serif;"
                    f"font-size:10.5px;font-style:italic;color:{TEXT_MUTED};margin-top:4px;'>"
                    f"{esc(r['pm_label'])}</div>"
                )
                last_pm = r["pm_label"]
 
            total = r["total"]
            pct = {k: (r[k] / total * 100 if total > 0 else 0) for k, _, _ in STAGE_COLS}
            invoiced_pct = (r["invoiced"] / total * 100) if total > 0 else 0.0
            var_color = COLOR_GREEN if r["variance"] >= 0 else COLOR_RED
            var_sign = "+" if r["variance"] >= 0 else ""
            var_label = "Over value" if r["variance"] >= 0 else "Under value"
 
            # bar order, left to right: Invoiced, Done, Planned, [variance marker], Remaining (grey)
            seg_parts = []
            for k, label, color in STAGE_COLS:
                if k == "remaining":
                    continue
                if pct[k] > 0:
                    seg_parts.append(
                        f"<div class='pid-tooltip' style='width:{pct[k]:.3f}%;background:{color};height:100%;'>"
                        f"<div class='tt-box'>{label}: {fmt_money(r[k])}</div>"
                        f"</div>"
                    )
            # variance marker always shown, between Planned and Remaining
            seg_parts.append(
                f"<div class='pid-tooltip' style='width:6px;flex:none;background:{var_color};height:100%;"
                f"border-left:1px solid #fff;border-right:1px solid #fff;'>"
                f"<div class='tt-box'>{var_label}: {var_sign}{fmt_money(r['variance'])}</div>"
                f"</div>"
            )
            if pct["remaining"] > 0:
                seg_parts.append(
                    f"<div class='pid-tooltip' style='width:{pct['remaining']:.3f}%;background:{COLOR_MUTED};height:100%;'>"
                    f"<div class='tt-box'>Remaining: {fmt_money(r['remaining'])}</div>"
                    f"</div>"
                )
            segs = "".join(seg_parts)
 
            job_full = esc(r["job"])
 
            rows_html.append(f"""
            <div style='display:flex;align-items:center;gap:10px;padding:3px 14px 3px 52px;
                        border-top:1px solid #F0F2F5;'>
              <div class='job-name-wrap' style='width:230px;flex:none;'>
                <div style='font-family:"IBM Plex Mono",monospace;font-size:11.5px;color:{TEXT_DARK};
                            font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'
                     title="{esc(r['job'])}">{esc(r['pid'])}</div>
                <div class='job-name-full'>{job_full}</div>
              </div>
              <div style='flex:1;height:15px;border-radius:3px;display:flex;background:#F0F2F5;'>
                {segs}
              </div>
              <div style='width:230px;flex:none;font-family:"IBM Plex Mono",monospace;font-size:11px;
                          color:{TEXT_DARK};text-align:right;'>
                <b>{fmt_money_short(total)}</b>&nbsp;·&nbsp;
                <span style='color:{COLOR_PURPLE};'>Inv {fmt_money_short(r["invoiced"])} ({invoiced_pct:.0f}%)</span>
              </div>
            </div>
            """)
 
        chart_html = "".join(rows_html)
        st.markdown(
            f"""
            <div style='max-height:900px;overflow-y:auto;border:1px solid {GRID_LIGHT};
                        border-radius:6px;padding:8px 0;background:#FFFFFF;'>
              {chart_html}
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption(
            "Bar order: Invoiced → Done → Planned → variance marker → Remaining (grey). "
            "Segments are proportional to that PID's own total, so proportions compare cleanly "
            "across very different PID sizes. Hover a segment to see its exact value, or hover "
            "a job name to read the full text. Scroll for the full list."
        )
 
# --------------------------------------------------------------------------
# TAB 3 — Finance
# --------------------------------------------------------------------------
 
with tab_finance:
    st.subheader("Panel 04 — Financial Health")
    st.caption("Budget performance, billing pipeline & invoicing speed · same filters as above")
 
    fin = fdf_dated.copy()
    fin_c = fin[fin["flag"] == "construction"].copy()
    fin_m = fin[fin["flag"] == "material"].copy()
 
    if fin.empty:
        st.info("No records under these filters. Widen the date range or clear a filter.")
    else:
        fin["has_plan"] = fin["plan1"].notna()
        fin["has_done"] = fin["done"].notna()
        fin["has_inv"] = fin["invoice date"].notna()
 
        total_value = fin["total"].sum()
        orig_value = fin_c["orig"].sum()
        variance = fin_c["total"].sum() - orig_value
        variance_pct = (variance / orig_value * 100) if orig_value else 0.0
 
        invoiced_sum = fin.loc[fin["has_inv"], "total"].sum()
        wip_sum = fin.loc[fin["has_done"] & ~fin["has_inv"], "total"].sum()          # earned, not yet billed
        backlog_sum = fin.loc[~fin["has_done"], "total"].sum()                        # not yet completed
        invoiced_pct = (invoiced_sum / total_value * 100) if total_value else 0.0
        material_pct = (fin_m["total"].sum() / total_value * 100) if total_value else 0.0
 
        # ---- KPI row 1 : budget performance ----
        metric_row([
            ("Total Value", fmt_money(total_value), COLOR_NAVY, True),
            ("Original Budget", fmt_money(orig_value), None, False),
            ("Variance", fmt_money(variance), COLOR_GREEN if variance >= 0 else COLOR_RED, False),
            ("Variance %", f"{variance_pct:+.2f}%", COLOR_GREEN if variance_pct >= 0 else COLOR_RED, False),
        ])
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        # ---- KPI row 2 : cash / pipeline ----
        metric_row([
            ("Invoiced", fmt_money(invoiced_sum) + f"  ({invoiced_pct:.1f}%)", COLOR_PURPLE, False),
            ("WIP — done, not invoiced", fmt_money(wip_sum), COLOR_GREEN, False),
            ("Backlog — not yet done", fmt_money(backlog_sum), COLOR_MUTED, False),
            ("Material share", f"{material_pct:.1f}%", COLOR_YELLOW, False),
        ])
 
        st.markdown("---")
 
        col_a, col_b = st.columns(2)
 
        # ---- Billing pipeline (single stacked bar) ----
        with col_a:
            st.markdown("**Billing pipeline**")
            st.caption("Where the total value currently sits, end to end")
 
            stage_vals = pd.Series({
                "Remaining\n(not started)": fin.loc[~fin["has_plan"], "total"].sum(),
                "Planned\n(not done)": fin.loc[fin["has_plan"] & ~fin["has_done"], "total"].sum(),
                "Done\n(not invoiced)": wip_sum,
                "Invoiced": invoiced_sum,
            })
            pipe_colors = [COLOR_MUTED, COLOR_AMBER, COLOR_GREEN, COLOR_PURPLE]
 
            fig4 = go.Figure()
            for stage, val, color in zip(stage_vals.index, stage_vals.values, pipe_colors):
                fig4.add_trace(go.Bar(
                    y=["Value"], x=[val], name=stage.replace("\n", " "), orientation="h",
                    marker_color=color,
                    text=[fmt_money_short(val)], textposition="inside", textfont=dict(color="#fff", size=12),
                    hovertemplate=f"<b>{stage.replace(chr(10),' ')}</b>: £%{{x:,.2f}}<extra></extra>",
                ))
            fig4.update_layout(
                **PLOTLY_LIGHT,
                height=180,
                barmode="stack",
                xaxis=dict(tickprefix="£", gridcolor=GRID_LIGHT, zeroline=False, tickfont=dict(size=12)),
                yaxis=dict(visible=False),
                legend=dict(orientation="h", y=-0.35, bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
                margin=dict(l=10, r=10, t=10, b=10),
            )
            st.plotly_chart(fig4, use_container_width=True)
 
        # ---- Invoicing lag ----
        with col_b:
            st.markdown("**Invoicing speed**")
            st.caption("Days between job marked *done* and *invoiced*")
 
            lag_df = fin[fin["has_done"] & fin["has_inv"]].copy()
            lag_df["lag_days"] = (lag_df["invoice date"] - lag_df["done"]).dt.days
            lag_df = lag_df[(lag_df["lag_days"] >= 0) & (lag_df["lag_days"] <= 365)]
 
            if lag_df.empty:
                st.info("Not enough done + invoiced pairs under these filters to measure lag.")
            else:
                med_lag = lag_df["lag_days"].median()
                mean_lag = lag_df["lag_days"].mean()
                st.markdown(
                    f"<span style='font-family:\"IBM Plex Mono\",monospace;font-size:13px;color:{TEXT_DARK};'>"
                    f"Median: <b>{med_lag:.0f} days</b> &nbsp;·&nbsp; Average: <b>{mean_lag:.0f} days</b> "
                    f"&nbsp;·&nbsp; n={len(lag_df):,}</span>",
                    unsafe_allow_html=True,
                )
                fig5 = go.Figure()
                fig5.add_trace(go.Histogram(
                    x=lag_df["lag_days"], nbinsx=30, marker_color=COLOR_TEAL,
                    hovertemplate="%{x} days<br>%{y} jobs<extra></extra>",
                ))
                fig5.add_vline(x=med_lag, line_dash="dash", line_color=COLOR_RED,
                                annotation_text="median", annotation_font_size=11)
                fig5.update_layout(
                    **PLOTLY_LIGHT,
                    height=180,
                    xaxis=dict(title="Days", gridcolor=GRID_LIGHT, tickfont=dict(size=11)),
                    yaxis=dict(title="Jobs", gridcolor=GRID_LIGHT, tickfont=dict(size=11)),
                    margin=dict(l=10, r=10, t=10, b=10),
                    bargap=0.1,
                )
                st.plotly_chart(fig5, use_container_width=True)
 
        st.markdown("---")
        st.markdown("**Variance by project**")
        st.caption("Construction total vs. original budget, worst to best")
 
        proj_var = fin_c.groupby("project", dropna=False).agg(total=("total", "sum"), orig=("orig", "sum"))
        proj_var = proj_var[proj_var["orig"] != 0]
        proj_var["variance"] = proj_var["total"] - proj_var["orig"]
        proj_var["variance_pct"] = proj_var["variance"] / proj_var["orig"] * 100
        proj_var = proj_var.sort_values("variance_pct")
 
        if proj_var.empty:
            st.info("No projects with a non-zero original budget under these filters.")
        else:
            fig6 = go.Figure()
            fig6.add_trace(go.Bar(
                y=proj_var.index.fillna("Unassigned"), x=proj_var["variance_pct"], orientation="h",
                marker_color=[COLOR_GREEN if v >= 0 else COLOR_RED for v in proj_var["variance_pct"]],
                text=[f"{v:+.1f}%" for v in proj_var["variance_pct"]],
                textposition="outside", textfont=dict(size=12),
                customdata=proj_var[["total", "orig", "variance"]],
                hovertemplate=(
                    "<b>%{y}</b><br>Total: £%{customdata[0]:,.2f}<br>Original: £%{customdata[1]:,.2f}"
                    "<br>Variance: £%{customdata[2]:,.2f} (%{x:+.1f}%)<extra></extra>"
                ),
            ))
            fig6.update_layout(
                **PLOTLY_LIGHT,
                height=max(220, 42 * len(proj_var) + 60),
                xaxis=dict(title="Variance %", gridcolor=GRID_LIGHT, zeroline=True, zerolinecolor="#B9C2CC", tickfont=dict(size=12)),
                yaxis=dict(gridcolor=GRID_LIGHT, tickfont=dict(size=13), automargin=True),
                margin=dict(l=10, r=40, t=10, b=10),
            )
            st.plotly_chart(fig6, use_container_width=True)
 
        st.markdown("---")
        st.markdown("**PID variance leaderboard**")
        st.caption("Largest overruns and largest underspends, by £ variance (construction only)")
 
        pid_var = fin_c.dropna(subset=["pid"]).groupby(
            ["pid", "district", "project"], dropna=False
        ).agg(total=("total", "sum"), orig=("orig", "sum")).reset_index()
        pid_var = pid_var[pid_var["orig"] != 0]
        pid_var["variance"] = pid_var["total"] - pid_var["orig"]
        pid_var["variance_pct"] = pid_var["variance"] / pid_var["orig"] * 100
 
        if pid_var.empty:
            st.info("No PIDs with a non-zero original budget under these filters.")
        else:
            lead_a, lead_b = st.columns(2)
            with lead_a:
                st.markdown(f"<span style='color:{COLOR_RED};font-weight:700;'>Top overruns</span>", unsafe_allow_html=True)
                worst = pid_var.sort_values("variance").head(8).copy()
                worst["Variance"] = worst["variance"].apply(fmt_money)
                worst["Variance %"] = worst["variance_pct"].apply(lambda v: f"{v:+.1f}%")
                st.dataframe(
                    worst[["pid", "district", "project", "Variance", "Variance %"]]
                    .rename(columns={"pid": "PID", "district": "District", "project": "Project"}),
                    hide_index=True, use_container_width=True,
                )
            with lead_b:
                st.markdown(f"<span style='color:{COLOR_GREEN};font-weight:700;'>Top underspends</span>", unsafe_allow_html=True)
                best = pid_var.sort_values("variance", ascending=False).head(8).copy()
                best["Variance"] = best["variance"].apply(fmt_money)
                best["Variance %"] = best["variance_pct"].apply(lambda v: f"{v:+.1f}%")
                st.dataframe(
                    best[["pid", "district", "project", "Variance", "Variance %"]]
                    .rename(columns={"pid": "PID", "district": "District", "project": "Project"}),
                    hide_index=True, use_container_width=True,
                )
 
st.markdown("---")
st.caption(f"Master Control Dashboard — data as of {date_max.date()}")
