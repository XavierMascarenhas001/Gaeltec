"""
app.py
------
Materials Breakdown Dashboard - a Streamlit app run locally (no usage
caps, no expiry - those only apply to Streamlit Community Cloud hosting;
`streamlit run app.py` on your own machine is just a local web server).

Run:
    streamlit run app.py

Requirements:
    pip install streamlit pandas pyarrow xlsxwriter plotly pymupdf
"""
from __future__ import annotations

import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px

import engine as eng

st.set_page_config(page_title="Materials Breakdown Dashboard", layout="wide")




# ---------------------------------------------------------------------------
# Cached loaders - only re-read from disk when the path actually changes
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Loading materials master...")
def cached_load_master(path: str, mtime: float):
    return eng.load_materials_master(path)


@st.cache_data(show_spinner="Loading control file...")
def cached_load_control(path: str, mtime: float):
    return eng.load_control_file(path)


@st.cache_data(show_spinner=False)
def cached_filter_options(control_df: pd.DataFrame):
    return eng.get_filter_options(control_df)


@st.cache_data(show_spinner="Matching materials...")
def cached_build_breakdown(filtered_control: pd.DataFrame, _poles, _free_issue, _guk_items, _guk_subitems, _aliases):
    # filtered_control is the cache key (small enough after filtering);
    # the master tables are prefixed with "_" so Streamlit doesn't try to
    # hash the (large, unchanging) master tables on every call.
    return eng.build_breakdown(filtered_control, _poles, _free_issue, _guk_items, _guk_subitems, _aliases)


_CATEGORY_PALETTE = [
    "#2E75B6", "#548235", "#BF8F00", "#7030A0", "#C00000",
    "#215868", "#E36C09", "#4472C4", "#70AD47", "#A9483D",
    "#31859C", "#948A54",
]


def _clean_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    v = round(float(v), 3)
    return int(v) if v == int(v) else v


def _write_table(book, writer, sheet_name, df, header_note=None, start_row=0, type_col="Type"):
    """Writes one dataframe as a plain formatted range with a merged,
    color-banded Type/Category column - the same polished layout as the
    desktop tool's export (merged cells, category color banding, clean
    number formatting, autofilter, frozen panes). Rows must already be
    sorted by the type column so same-type runs are contiguous."""
    if sheet_name not in writer.sheets:
        writer.sheets[sheet_name] = book.add_worksheet(sheet_name)
    ws = writer.sheets[sheet_name]

    row = start_row
    if header_note:
        note_fmt = book.add_format({"bold": True, "font_size": 12})
        ws.write(row, 0, header_note, note_fmt)
        row += 1

    if df.empty:
        ws.write(row, 0, "(no matching rows)")
        return row + 2

    header_fmt = book.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1, "align": "center"})
    for c, col in enumerate(df.columns):
        ws.write(row, c, col, header_fmt)

    has_type = type_col in df.columns
    type_col_idx = list(df.columns).index(type_col) if has_type else None

    cat_colors = {}
    if has_type:
        color_i = 0
        for cat in df[type_col]:
            key = cat if pd.notna(cat) else "(Uncategorized)"
            if key not in cat_colors:
                cat_colors[key] = _CATEGORY_PALETTE[color_i % len(_CATEGORY_PALETTE)]
                color_i += 1

    band_colors = ["#FFFFFF", "#F2F2F2"]
    n = len(df)
    data_start = row + 1

    r = data_start
    i = 0
    band_i = 0
    while i < n:
        j = i
        cat_i = df.iloc[i][type_col] if has_type else None
        while j + 1 < n and (df.iloc[j + 1][type_col] if has_type else None) == cat_i:
            j += 1

        band_color = band_colors[band_i % len(band_colors)]
        band_i += 1
        row_fmt = book.add_format({"bg_color": band_color, "border": 1})
        cat_fmt = book.add_format({
            "bg_color": cat_colors.get(cat_i if pd.notna(cat_i) else "(Uncategorized)", "#D9E1F2"),
            "font_color": "white", "bold": True, "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        })

        for row_idx in range(i, j + 1):
            rec = df.iloc[row_idx]
            for c, val in enumerate(rec):
                if c == type_col_idx:
                    continue
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    ws.write_blank(r + (row_idx - i), c, None, row_fmt)
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    ws.write(r + (row_idx - i), c, _clean_num(val), row_fmt)
                else:
                    ws.write(r + (row_idx - i), c, val, row_fmt)

        if has_type:
            top, bottom = r, r + (j - i)
            label = cat_i if pd.notna(cat_i) else "(Uncategorized)"
            if bottom > top:
                ws.merge_range(top, type_col_idx, bottom, type_col_idx, label, cat_fmt)
            else:
                ws.write(top, type_col_idx, label, cat_fmt)

        r += (j - i) + 1
        i = j + 1

    last_row = data_start + n - 1
    ws.autofilter(row, 0, last_row, len(df.columns) - 1)
    ws.freeze_panes(data_start, 0)

    for c, col in enumerate(df.columns):
        width = max(len(str(col)), df[col].astype(str).str.len().max() if len(df) else 0)
        ws.set_column(c, c, min(max(width + 2, 10), 50))

    return last_row + 2


def export_to_excel_formatted(
    poles_sheet: pd.DataFrame,
    fi_sheet: pd.DataFrame,
    guk_sheet: pd.DataFrame,
    guk_sub_sheet: pd.DataFrame,
    filter_summary: dict,
    pid_summary: pd.DataFrame | None = None,
    pole_summary: pd.DataFrame | None = None,
) -> bytes:
    """Polished layout matching the desktop tool's export: merged/color-
    banded Type column per sheet, GUK assemblies + subdivisions stacked
    in one sheet with header notes, autofilter, frozen panes, clean
    number formatting - instead of a bare pandas .to_excel() dump."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        book = writer.book

        ws = book.add_worksheet("Summary")
        writer.sheets["Summary"] = ws
        bold = book.add_format({"bold": True})
        header_fmt = book.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1})
        band_a = book.add_format({"border": 1, "bg_color": "#FFFFFF"})
        band_b = book.add_format({"border": 1, "bg_color": "#F2F2F2"})

        ws.write(0, 0, "Materials Breakdown Export", bold)
        r = 2
        for label, value in filter_summary.items():
            ws.write(r, 0, label, bold)
            ws.write(r, 1, value)
            r += 1

        if pid_summary is not None and not pid_summary.empty:
            r += 1
            ws.write(r, 0, "PIDs in this export", bold)
            r += 1
            for c, col in enumerate(pid_summary.columns):
                ws.write(r, c, col, header_fmt)
            r += 1
            for row_idx, rec in enumerate(pid_summary.itertuples(index=False)):
                fmt = band_a if row_idx % 2 == 0 else band_b
                for c, val in enumerate(rec):
                    if val is None or (isinstance(val, float) and pd.isna(val)):
                        ws.write_blank(r, c, None, fmt)
                    else:
                        ws.write(r, c, val, fmt)
                r += 1
            for c, col in enumerate(pid_summary.columns):
                width = max(len(str(col)), pid_summary[col].astype(str).str.len().max())
                ws.set_column(c, c, min(max(width + 2, 10), 40))

        _write_table(book, writer, "Poles", poles_sheet, type_col="Category")
        _write_table(book, writer, "Free Issue", fi_sheet, type_col="Type")
        next_row = _write_table(book, writer, "GUK", guk_sheet, header_note="GUK Assemblies", type_col="Type")
        _write_table(book, writer, "GUK", guk_sub_sheet, header_note="GUK Subdivisions (component breakdown)",
                     start_row=next_row, type_col="Type")

        if pole_summary is not None and not pole_summary.empty:
            ws_p = book.add_worksheet("Pole Summary")
            writer.sheets["Pole Summary"] = ws_p
            for c, col in enumerate(pole_summary.columns):
                ws_p.write(0, c, col, header_fmt)
            for row_idx, rec in enumerate(pole_summary.itertuples(index=False)):
                fmt = band_a if row_idx % 2 == 0 else band_b
                for c, val in enumerate(rec):
                    ws_p.write(row_idx + 1, c, val, fmt)
            ws_p.autofilter(0, 0, len(pole_summary), len(pole_summary.columns) - 1)
            ws_p.freeze_panes(1, 0)
            for c, col in enumerate(pole_summary.columns):
                width = max(len(str(col)), pole_summary[col].astype(str).str.len().max())
                ws_p.set_column(c, c, min(max(width + 2, 10), 40))

    return buf.getvalue()


# ---------------------------------------------------------------------------
# Sidebar: data files
# ---------------------------------------------------------------------------
st.sidebar.header("Data files")

# Auto-discover every "Master_DD-MM-YYYY[ (vN)]..." control file in the
# project tracker folder - most recent date wins, then highest version
# number for that date (a missing version suffix counts as v0, so plain
# "Master_01-09-2026.parquet" files are picked up, not skipped). The
# newest is pre-selected, but every candidate is offered in case the
# auto-pick guesses wrong.
_all_control_files = eng.find_all_master_control_files(eng.MASTER_CONTROL_FOLDER)

master_path = st.sidebar.text_input("Materials master parquet", value=eng.DEFAULT_MATERIALS_MASTER)

if _all_control_files:
    _control_labels = [
        f"{p.name}" + ("  ← newest" if i == 0 else "")
        for i, (_, _, p) in enumerate(_all_control_files)
    ]
    _control_choice = st.sidebar.selectbox(
        "Control file (Master parquet)",
        options=range(len(_all_control_files)),
        format_func=lambda i: _control_labels[i],
        index=0,
        help="Newest by date first, then by version. Pick a different one if the auto-detected file isn't the one you want.",
    )
    control_path = str(_all_control_files[_control_choice][2])
    with st.sidebar.expander("Or enter a path manually"):
        _manual_control_path = st.text_input("Control file parquet (manual override)", value="")
        if _manual_control_path:
            control_path = _manual_control_path
else:
    st.sidebar.warning(
        f"Couldn't auto-detect a 'Master_DD-MM-YYYY[ (vN)]...' file in:\n{eng.MASTER_CONTROL_FOLDER}\n"
        "Enter the control file path manually below."
    )
    control_path = st.sidebar.text_input("Control file parquet", value="")

if st.sidebar.button("Re-check for a newer Master file"):
    st.cache_data.clear()
    st.rerun()

if not master_path or not os.path.exists(master_path):
    st.error(f"Materials master not found: {master_path}")
    st.stop()
if not control_path or not os.path.exists(control_path):
    st.error(f"Control file not found: {control_path}")
    st.stop()

poles, free_issue, guk_items, guk_subitems, aliases = cached_load_master(
    master_path, os.path.getmtime(master_path)
)
control_df = cached_load_control(control_path, os.path.getmtime(control_path))
filter_options = cached_filter_options(control_df)

st.sidebar.caption(
    f"{len(poles)} poles, {len(free_issue)} free issue, {len(guk_items)} GUK items, "
    f"{len(guk_subitems)} GUK sub-lines, {len(aliases)} aliases | "
    f"{len(control_df):,} control-file rows"
)

# ---------------------------------------------------------------------------
# Filters (all multi-select). Circuit and Pole are CASCADING: their
# available options narrow to whatever actually exists given the
# District/Project/Date selections (and, for Pole, the Circuit selection
# too) - so you never see a circuit or pole that couldn't possibly match
# what's already chosen. District, Project, PID, and the date range itself
# are not narrowed by anything (there's no natural "parent" filter to
# narrow them by here).
# ---------------------------------------------------------------------------
st.sidebar.header("Filters")

date_field_label = st.sidebar.radio("Date field", list(eng.DATE_FIELDS.keys()), index=2, horizontal=True)
date_field = eng.DATE_FIELDS[date_field_label]
col_a, col_b = st.sidebar.columns(2)
date_from = col_a.date_input("From", value=None, format="YYYY-MM-DD")
date_to = col_b.date_input("To", value=None, format="YYYY-MM-DD")

districts = st.sidebar.multiselect("District", filter_options["district"])
projects = st.sidebar.multiselect("Project", filter_options["project"])
pids = st.sidebar.multiselect("PID", filter_options["pid"])
sourcefiles = st.sidebar.multiselect("Source file", filter_options["sourcefile"])


def _distinct(series: pd.Series) -> list[str]:
    vals = series.dropna().astype(str).str.strip()
    return sorted({v for v in vals if v and v.lower() != "nan"})


# Circuit options: narrowed by District/Project/Date only (not by Circuit
# or Pole themselves, which would be circular).
scope_for_circuit = eng.filter_control(
    control_df, districts=districts or None, projects=projects or None,
    date_field=date_field, date_from=date_from, date_to=date_to,
)
circuit_options = _distinct(scope_for_circuit["circuit"])
circuits = st.sidebar.multiselect(
    "Circuit", circuit_options,
    help="Narrowed to circuits that actually exist given the District/Project/Date selected above.",
)

# Pole (enid) options: narrowed by District/Project/Date/Circuit (every
# filter selected so far except Pole itself).
scope_for_pole = eng.filter_control(
    control_df, districts=districts or None, projects=projects or None,
    circuits=circuits or None, date_field=date_field, date_from=date_from, date_to=date_to,
)
enid_options = _distinct(scope_for_pole["enid"])
enids = st.sidebar.multiselect(
    "Pole (enid) - optional, search to narrow", enid_options,
    help="Narrowed to poles that actually exist given the District/Project/Circuit/Date selected above. Leave empty to include all of them.",
)

filtered = eng.filter_control(
    control_df,
    districts=districts or None,
    projects=projects or None,
    circuits=circuits or None,
    pids=pids or None,
    enids=enids or None,
    sourcefiles=sourcefiles or None,
    date_field=date_field,
    date_from=date_from,
    date_to=date_to,
)
mat_rows = (filtered["cat"].str.upper() == "MAT").sum()
st.sidebar.metric("Control-file rows matched", f"{len(filtered):,}", f"{mat_rows:,} material rows")

poles_out, fi_out, guk_out, guk_sub_out, unmatched_out = cached_build_breakdown(
    filtered, poles, free_issue, guk_items, guk_subitems, aliases
)

# ---------------------------------------------------------------------------
# Header / KPIs
# ---------------------------------------------------------------------------
st.title("Materials Breakdown Dashboard")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Poles", int(poles_out["qty"].sum()) if not poles_out.empty else 0)
k2.metric("Free Issue line items", len(fi_out))
k3.metric("GUK assemblies", len(guk_out))
k4.metric("GUK components", len(guk_sub_out))

tab_poles, tab_fi, tab_guk, tab_maps, tab_export = st.tabs(
    ["Poles", "Free Issue", "GUK", "Maps (optional)", "Export"]
)

# ---------------------------------------------------------------------------
# Poles tab
# ---------------------------------------------------------------------------
with tab_poles:
    st.subheader("Poles needed")
    if poles_out.empty:
        st.info("No poles matched for the current filters.")
    else:
        chart_df = poles_out.groupby("material_item", as_index=False)["qty"].sum().sort_values("qty", ascending=False)
        fig = px.bar(chart_df, x="material_item", y="qty", title="Pole quantities by size", labels={"material_item": "Pole size", "qty": "Quantity"})
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(
            poles_out.rename(columns={
                "category": "Category", "material_item": "Pole Size", "description": "Description", "qty": "Qty",
            })[["Category", "Pole Size", "Description", "Qty"]],
            use_container_width=True, hide_index=True,
        )

# ---------------------------------------------------------------------------
# Free Issue tab
# ---------------------------------------------------------------------------
with tab_fi:
    st.subheader("Free Issue materials")
    if fi_out.empty:
        st.info("No Free Issue materials matched for the current filters.")
    else:
        fi_categories = sorted(fi_out["category"].dropna().unique())
        chosen_cats = st.multiselect("Filter by Free Issue Type", fi_categories, key="fi_cat")
        fi_scope = fi_out[fi_out["category"].isin(chosen_cats)] if chosen_cats else fi_out

        materials_in_scope = sorted(fi_scope["material_item"].dropna().unique())
        chosen_materials = st.multiselect("Filter by Free Issue material", materials_in_scope, key="fi_mat")
        fi_display = fi_scope[fi_scope["material_item"].isin(chosen_materials)] if chosen_materials else fi_scope

        chart_data = fi_display.groupby(["category", "code", "description"], dropna=False, as_index=False)["qty"].sum()
        fig = px.bar(
            chart_data, x="code", y="qty", color="category", hover_data=["description"],
            title="Free Issue quantities by commodity code",
            labels={"code": "Commodity Code", "qty": "Quantity", "category": "Type"},
        )
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            fi_display.rename(columns={
                "category": "Type", "material_item": "Item", "description": "Description",
                "unit": "Unit", "code": "Commodity Code", "pid": "PID", "qty": "Qty",
            })[["Type", "Item", "Description", "Commodity Code", "PID", "Qty", "Unit"]],
            use_container_width=True, hide_index=True,
        )

# ---------------------------------------------------------------------------
# GUK tab
# ---------------------------------------------------------------------------
with tab_guk:
    st.subheader("GUK assemblies")
    if guk_out.empty:
        st.info("No GUK assemblies matched for the current filters.")
    else:
        guk_categories = sorted(guk_out["category"].dropna().unique())
        chosen_guk_cats = st.multiselect("Filter by GUK Type", guk_categories, key="guk_cat")
        guk_scope = guk_out[guk_out["category"].isin(chosen_guk_cats)] if chosen_guk_cats else guk_out

        guk_materials = sorted(guk_scope["material_item"].dropna().unique())
        chosen_guk_materials = st.multiselect("Filter by GUK material (assembly)", guk_materials, key="guk_mat")
        guk_display = guk_scope[guk_scope["material_item"].isin(chosen_guk_materials)] if chosen_guk_materials else guk_scope

        fig_guk = px.bar(
            guk_display.groupby(["category", "material_item"], as_index=False)["qty"].sum(),
            x="material_item", y="qty", color="category",
            title="GUK assembly quantities", labels={"material_item": "Assembly", "qty": "Quantity", "category": "Type"},
        )
        st.plotly_chart(fig_guk, use_container_width=True)
        st.dataframe(
            guk_display.rename(columns={
                "category": "Type", "material_item": "Item", "description": "Description",
                "unit": "Unit", "qty": "Qty",
            })[["Type", "Item", "Description", "Qty", "Unit"]],
            use_container_width=True, hide_index=True,
        )

        st.markdown("---")
        st.subheader("GUK sub-materials (component breakdown)")
        parent_scope = chosen_guk_materials if chosen_guk_materials else guk_materials
        sub_scope_df = guk_sub_out[guk_sub_out["material_item"].isin(parent_scope)]
        sub_materials = sorted(sub_scope_df["code"].dropna().unique())
        chosen_sub_materials = st.multiselect("Filter by sub GUK material", sub_materials, key="guk_submat")
        sub_display = sub_scope_df[sub_scope_df["code"].isin(chosen_sub_materials)] if chosen_sub_materials else sub_scope_df

        if sub_display.empty:
            st.info("No GUK sub-materials in scope.")
        else:
            fig_sub = px.bar(
                sub_display.groupby(["category", "code", "description"], dropna=False, as_index=False)["qty"].sum(),
                x="code", y="qty", color="category", hover_data=["description"],
                title="GUK sub-material quantities by sub code",
                labels={"code": "Sub Code", "qty": "Quantity", "category": "Type"},
            )
            st.plotly_chart(fig_sub, use_container_width=True)
            st.dataframe(
                sub_display.rename(columns={
                    "category": "Type", "material_item": "Parent Item", "description": "Description",
                    "code": "Sub Code", "unit": "Unit",
                    "qty_per_unit": "Qty Per Unit", "qty": "Qty",
                })[["Type", "Parent Item", "Sub Code", "Description", "Qty Per Unit", "Qty", "Unit"]],
                use_container_width=True, hide_index=True,
            )

# ---------------------------------------------------------------------------
# Maps tab (optional enrichment - never restricts the main totals above)
# ---------------------------------------------------------------------------
# These are always defined (even if no maps have been scanned this run) so
# the Export tab below can safely offer a "poles found on maps" export
# scope without erroring when that hasn't been used yet.
all_map_pole_enids: set = set()
matched_on_maps: set = set()
not_on_maps: set = set()
pole_summary_df = pd.DataFrame(columns=["Pole (enid)", "On Scanned Map"])

with tab_maps:
    st.subheader("Connect maps (optional)")
    st.caption(
        "Maps only ADD context (which poles were found on which map) - the Poles/Free "
        "Issue/GUK tabs above always include every pole matching your filters, whether "
        "or not it appears on any scanned map. Poles in your filtered range with no map "
        "match are listed separately below."
    )

    mode = st.radio("Source", ["Manual folder(s)", "Network path scan"], horizontal=True)
    pdf_entries = []

    if mode == "Manual folder(s)":
        folders_text = st.text_area(
            "Folder path(s), one per line",
            help='Each folder is scanned recursively. Its own name becomes "Folder"; the folder ONE LEVEL UP becomes "Day/Outage".',
        )
        folders = [f.strip() for f in folders_text.splitlines() if f.strip()]
        if st.button("Scan folders"):
            pdf_entries = eng.collect_pdf_entries(folders)
            st.session_state["map_pdf_entries"] = pdf_entries
            st.success(f"Found {len(pdf_entries)} PDF(s) in {len(folders)} folder(s).")
    else:
        network_root = st.text_input("Network root", value=eng.NETWORK_ROOT)
        if st.button("Scan network path for Workpack zones (uses the date filter above)"):
            with st.spinner("Scanning network path..."):
                try:
                    entries, zone_count, pdf_count, no_zone = eng.find_workpack_zone_pdfs(
                        network_root, date_from=date_from, date_to=date_to,
                    )
                    st.session_state["map_pdf_entries"] = entries
                    st.success(f"Found {zone_count} Workpack zones folder(s), {pdf_count} PDF(s).")
                    if no_zone:
                        st.warning(f"{len(no_zone)} outage folder(s) in range had no Workpack zones folder.")
                except Exception as e:
                    st.error(str(e))

    pdf_entries = st.session_state.get("map_pdf_entries", [])

    if pdf_entries:
        if eng.fitz is None:
            st.error("PyMuPDF is not installed. Run: pip install pymupdf")
        else:
            with st.spinner(f"Reading {len(pdf_entries)} PDF(s) for pole references..."):
                vocab = eng.build_pole_vocabulary(control_df)
                pole_index = eng.build_map_pole_index_from_pdfs(pdf_entries, vocab)

            for s in pole_index.values():
                all_map_pole_enids |= s

            filtered_enids = set(filtered["enid"].astype(str).str.strip())
            matched_on_maps = filtered_enids & all_map_pole_enids
            not_on_maps = filtered_enids - all_map_pole_enids

            # Per-pole summary (every filtered pole, flagged Yes/No) - used
            # both for on-screen display and as its own Excel sheet.
            pole_summary_df = pd.DataFrame({
                "Pole (enid)": sorted(filtered_enids),
            })
            pole_summary_df["On Scanned Map"] = pole_summary_df["Pole (enid)"].isin(matched_on_maps).map({True: "Yes", False: "No"})
            pole_summary_df = pole_summary_df.sort_values(["On Scanned Map", "Pole (enid)"], ascending=[False, True]).reset_index(drop=True)

            m1, m2, m3 = st.columns(3)
            m1.metric("Maps scanned", len(pdf_entries))
            m2.metric("Filtered poles found on a map", len(matched_on_maps))
            m3.metric("Filtered poles NOT on any map (materials below still included)", len(not_on_maps))

            map_rows = []
            for (day_outage, folder, fname), pole_set in pole_index.items():
                map_rows.append({
                    "Day/Outage": day_outage, "Folder": folder, "Map": fname,
                    "Poles found on map": len(pole_set),
                    "Poles also in your filter": len(pole_set & filtered_enids),
                })
            if map_rows:
                st.dataframe(pd.DataFrame(map_rows), use_container_width=True, hide_index=True)

            st.markdown("**Pole summary for your current filter** (every pole matching your sidebar filters, flagged by whether it was found on a scanned map):")
            st.dataframe(pole_summary_df, use_container_width=True, hide_index=True)
    else:
        st.info("No maps scanned yet.")

# ---------------------------------------------------------------------------
# Export tab
# ---------------------------------------------------------------------------
with tab_export:
    st.subheader("Export current filters to Excel")

    export_scope = st.radio(
        "Export scope",
        ["All poles matching filters", "Only poles found on scanned maps"],
        horizontal=True,
        help='"Only poles found on scanned maps" requires scanning maps on the Maps tab first.',
    )

    scope_ready = True
    if export_scope == "Only poles found on scanned maps":
        if not matched_on_maps:
            st.warning(
                "No maps have been scanned yet (or none of your filtered poles were found on any scanned "
                'map) - go to the Maps tab first, or switch back to "All poles matching filters".'
            )
            scope_ready = False
        else:
            export_scope_control_df = filtered[filtered["enid"].astype(str).str.strip().isin(matched_on_maps)]
            poles_out_x, fi_out_x, guk_out_x, guk_sub_out_x, unmatched_out_x = eng.build_breakdown(
                export_scope_control_df, poles, free_issue, guk_items, guk_subitems, aliases
            )
            st.caption(f"Exporting materials for {len(matched_on_maps)} pole(s) found on scanned maps (out of {len(filtered['enid'].unique())} matching your filters).")
    else:
        export_scope_control_df = filtered
        poles_out_x, fi_out_x, guk_out_x, guk_sub_out_x, unmatched_out_x = poles_out, fi_out, guk_out, guk_sub_out, unmatched_out

    if scope_ready:
        poles_sheet = poles_out_x.rename(columns={
            "category": "Category", "material_item": "Pole Size", "description": "Description", "qty": "Qty",
        })[["Category", "Pole Size", "Description", "Qty"]].sort_values(["Category", "Pole Size"])

        # Free Issue export drops PID as a row-level column (kept only in
        # the Summary sheet's PID table instead) - re-aggregate across PID
        # so rows that only differed by PID collapse into one combined
        # total per material, rather than leaving PID-driven "duplicate"
        # rows with no visible explanation for why they're separate.
        fi_for_export = (
            fi_out_x.groupby(["category", "material_item", "description", "unit", "code"], dropna=False, as_index=False)["qty"]
            .sum()
        )
        fi_sheet = fi_for_export.rename(columns={
            "category": "Type", "material_item": "Item", "description": "Description",
            "unit": "Unit", "code": "Commodity Code", "qty": "Qty",
        })[["Type", "Item", "Description", "Commodity Code", "Qty", "Unit"]].sort_values(["Type", "Item"])

        guk_sheet = guk_out_x.rename(columns={
            "category": "Type", "material_item": "Item", "description": "Description",
            "unit": "Unit", "qty": "Qty",
        })[["Type", "Item", "Description", "Qty", "Unit"]].sort_values(["Type", "Item"])

        guk_sub_sheet = guk_sub_out_x.rename(columns={
            "category": "Type", "material_item": "Parent Item", "description": "Description",
            "code": "Sub Code", "unit": "Unit",
            "qty_per_unit": "Qty Per Unit", "qty": "Qty",
        })[["Type", "Parent Item", "Sub Code", "Description", "Qty Per Unit", "Qty", "Unit"]].sort_values(["Type", "Parent Item"])

        # Summary reflects the ACTUAL data being exported, not the raw
        # sidebar filter selections - e.g. if District was left blank
        # ("all"), this shows the districts that actually turned up in the
        # exported rows, and Date Start/End are the real min/max dates
        # present rather than whatever was typed into the date pickers.
        def _actual_values(col):
            vals = export_scope_control_df[col].dropna().astype(str).str.strip()
            vals = sorted({v for v in vals if v and v.lower() != "nan"})
            return ", ".join(vals) if vals else "(none)"

        actual_dates = pd.to_datetime(export_scope_control_df[date_field], errors="coerce").dropna()
        date_start = actual_dates.min().strftime("%Y-%m-%d") if not actual_dates.empty else "(none)"
        date_end = actual_dates.max().strftime("%Y-%m-%d") if not actual_dates.empty else "(none)"

        filter_summary = {
            "Export scope": export_scope,
            "District": _actual_values("district"),
            "Project": _actual_values("project"),
            "Circuit": _actual_values("circuit"),
            "Date field": date_field_label,
            "Date start": date_start,
            "Date end": date_end,
            "Control-file rows exported": len(export_scope_control_df),
            "Exported": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }

        pid_summary_df = eng.build_pid_summary(export_scope_control_df)
        export_pole_summary = pole_summary_df if not pole_summary_df.empty else None

        excel_bytes = export_to_excel_formatted(
            poles_sheet, fi_sheet, guk_sheet, guk_sub_sheet, filter_summary,
            pid_summary=pid_summary_df, pole_summary=export_pole_summary,
        )
        st.download_button(
            "Download Materials_Breakdown.xlsx",
            data=excel_bytes,
            file_name=f"Materials_Breakdown_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if not unmatched_out_x.empty:
            with st.expander(f"Unmatched descriptions ({len(unmatched_out_x)}) - not exported above, shown here for QA"):
                st.dataframe(unmatched_out_x, use_container_width=True, hide_index=True)
