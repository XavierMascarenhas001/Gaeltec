"""
dash_materials.py - Materials Breakdown dashboard, data side.

engine.py (loading, filtering, materials matching, map scanning) is used
unchanged. The Excel export (_write_table / export_to_excel_formatted) is
copied from the Streamlit app.py unchanged. What was Streamlit screen code
is replaced by functions that return plain data for the web page.
"""
from __future__ import annotations

import io
import os
from datetime import datetime

import pandas as pd

import engine as eng

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
# Session cache - each file is read once
# ---------------------------------------------------------------------------
_MASTER = {}
_CONTROL = {}
_MAPS = {"key": None, "pole_index": None}


def _materials(path):
    key = (path, os.path.getmtime(path))
    if key not in _MASTER:
        _MASTER.clear()
        _MASTER[key] = eng.load_materials_master(path)
    return _MASTER[key]


def _control(path):
    key = (path, os.path.getmtime(path))
    if key not in _CONTROL:
        _CONTROL.clear()
        _CONTROL[key] = eng.load_control_file(path)
    return _CONTROL[key]


def _s(v):
    return "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v)


def _n(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return _clean_num(v)


def _records(df, cols):
    out = []
    for rec in df[cols].itertuples(index=False):
        out.append([_n(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else _s(v) for v in rec])
    return out


def load(master_path, control_path):
    poles, free_issue, guk_items, guk_subitems, aliases = _materials(master_path)
    control_df = _control(control_path)
    opts = eng.get_filter_options(control_df)
    return {
        "master_name": os.path.basename(master_path), "control_name": os.path.basename(control_path),
        "counts": {"poles": len(poles), "free_issue": len(free_issue), "guk_items": len(guk_items),
                   "guk_subitems": len(guk_subitems), "aliases": len(aliases), "control_rows": len(control_df)},
        "date_fields": list(eng.DATE_FIELDS.keys()), "default_date_field": list(eng.DATE_FIELDS.keys())[2],
        "options": {k: opts[k] for k in ("district", "project", "pid", "sourcefile")},
        "network_root": eng.NETWORK_ROOT,
    }


def _distinct(series: pd.Series) -> list[str]:
    vals = series.dropna().astype(str).str.strip()
    return sorted({v for v in vals if v and v.lower() != "nan"})


def _filtered(control_df, date_field_label="DateToUse", date_from=None, date_to=None, districts=None,
              projects=None, pids=None, sourcefiles=None, circuits=None, enids=None):
    date_field = eng.DATE_FIELDS[date_field_label]
    date_from = date_from or None
    date_to = date_to or None
    scope_for_circuit = eng.filter_control(
        control_df, districts=districts or None, projects=projects or None,
        date_field=date_field, date_from=date_from, date_to=date_to,
    )
    circuit_options = _distinct(scope_for_circuit["circuit"])
    scope_for_pole = eng.filter_control(
        control_df, districts=districts or None, projects=projects or None,
        circuits=circuits or None, date_field=date_field, date_from=date_from, date_to=date_to,
    )
    enid_options = _distinct(scope_for_pole["enid"])
    filtered = eng.filter_control(
        control_df, districts=districts or None, projects=projects or None, circuits=circuits or None,
        pids=pids or None, enids=enids or None, sourcefiles=sourcefiles or None,
        date_field=date_field, date_from=date_from, date_to=date_to,
    )
    return filtered, circuit_options, enid_options, date_field


def view(master_path, control_path, filters=None, fi_cats=None, fi_mats=None,
         guk_cats=None, guk_mats=None, guk_submats=None):
    poles, free_issue, guk_items, guk_subitems, aliases = _materials(master_path)
    control_df = _control(control_path)
    filtered, circuit_options, enid_options, _df_field = _filtered(control_df, **(filters or {}))
    mat_rows = int((filtered["cat"].str.upper() == "MAT").sum())
    poles_out, fi_out, guk_out, guk_sub_out, unmatched_out = eng.build_breakdown(
        filtered, poles, free_issue, guk_items, guk_subitems, aliases)

    res = {
        "circuit_options": circuit_options, "pole_options": enid_options,
        "rows": int(len(filtered)), "mat_rows": mat_rows,
        "kpis": {"poles": int(poles_out["qty"].sum()) if not poles_out.empty else 0,
                 "fi": int(len(fi_out)), "guk": int(len(guk_out)), "guk_sub": int(len(guk_sub_out))},
    }

    # ---- Poles ----
    if poles_out.empty:
        res["poles"] = None
    else:
        chart_df = poles_out.groupby("material_item", as_index=False)["qty"].sum().sort_values("qty", ascending=False)
        t = poles_out.rename(columns={"category": "Category", "material_item": "Pole Size", "description": "Description", "qty": "Qty"})
        res["poles"] = {"chart": {"x": [_s(v) for v in chart_df["material_item"]], "y": [_n(v) for v in chart_df["qty"]]},
                        "columns": ["Category", "Pole Size", "Description", "Qty"],
                        "rows": _records(t, ["Category", "Pole Size", "Description", "Qty"])}

    # ---- Free Issue ----
    if fi_out.empty:
        res["fi"] = None
    else:
        fi_categories = sorted(fi_out["category"].dropna().unique())
        chosen_cats = [c for c in (fi_cats or []) if c in fi_categories]
        fi_scope = fi_out[fi_out["category"].isin(chosen_cats)] if chosen_cats else fi_out
        materials_in_scope = sorted(fi_scope["material_item"].dropna().unique())
        chosen_materials = [m for m in (fi_mats or []) if m in materials_in_scope]
        fi_display = fi_scope[fi_scope["material_item"].isin(chosen_materials)] if chosen_materials else fi_scope
        chart_data = fi_display.groupby(["category", "code", "description"], dropna=False, as_index=False)["qty"].sum()
        t = fi_display.rename(columns={"category": "Type", "material_item": "Item", "description": "Description",
                                       "unit": "Unit", "code": "Commodity Code", "pid": "PID", "qty": "Qty"})
        cols = ["Type", "Item", "Description", "Commodity Code", "PID", "Qty", "Unit"]
        res["fi"] = {"categories": [_s(c) for c in fi_categories], "materials": [_s(m) for m in materials_in_scope],
                     "chart": [{"category": _s(r.category), "code": _s(r.code), "description": _s(r.description), "qty": _n(r.qty)}
                               for r in chart_data.itertuples(index=False)],
                     "columns": cols, "rows": _records(t, cols)}

    # ---- GUK ----
    if guk_out.empty:
        res["guk"] = None
    else:
        guk_categories = sorted(guk_out["category"].dropna().unique())
        chosen_guk_cats = [c for c in (guk_cats or []) if c in guk_categories]
        guk_scope = guk_out[guk_out["category"].isin(chosen_guk_cats)] if chosen_guk_cats else guk_out
        guk_materials = sorted(guk_scope["material_item"].dropna().unique())
        chosen_guk_materials = [m for m in (guk_mats or []) if m in guk_materials]
        guk_display = guk_scope[guk_scope["material_item"].isin(chosen_guk_materials)] if chosen_guk_materials else guk_scope
        gchart = guk_display.groupby(["category", "material_item"], as_index=False)["qty"].sum()
        t = guk_display.rename(columns={"category": "Type", "material_item": "Item", "description": "Description", "unit": "Unit", "qty": "Qty"})
        gcols = ["Type", "Item", "Description", "Qty", "Unit"]

        parent_scope = chosen_guk_materials if chosen_guk_materials else guk_materials
        sub_scope_df = guk_sub_out[guk_sub_out["material_item"].isin(parent_scope)]
        sub_materials = sorted(sub_scope_df["code"].dropna().unique())
        chosen_sub = [m for m in (guk_submats or []) if m in sub_materials]
        sub_display = sub_scope_df[sub_scope_df["code"].isin(chosen_sub)] if chosen_sub else sub_scope_df
        sub = None
        if not sub_display.empty:
            schart = sub_display.groupby(["category", "code", "description"], dropna=False, as_index=False)["qty"].sum()
            st_ = sub_display.rename(columns={"category": "Type", "material_item": "Parent Item", "description": "Description",
                                              "code": "Sub Code", "unit": "Unit", "qty_per_unit": "Qty Per Unit", "qty": "Qty"})
            scols = ["Type", "Parent Item", "Sub Code", "Description", "Qty Per Unit", "Qty", "Unit"]
            sub = {"chart": [{"category": _s(r.category), "code": _s(r.code), "description": _s(r.description), "qty": _n(r.qty)}
                             for r in schart.itertuples(index=False)],
                   "columns": scols, "rows": _records(st_, scols)}
        res["guk"] = {"categories": [_s(c) for c in guk_categories], "materials": [_s(m) for m in guk_materials],
                      "sub_materials": [_s(m) for m in sub_materials],
                      "chart": [{"category": _s(r.category), "item": _s(r.material_item), "qty": _n(r.qty)} for r in gchart.itertuples(index=False)],
                      "columns": gcols, "rows": _records(t, gcols), "sub": sub}

    res["unmatched"] = {"columns": ["Description", "Total qsub", "Rows"],
                        "rows": [[_s(r.description), _n(r.total_qsub), int(r.n_rows)] for r in unmatched_out.itertuples(index=False)]}
    return res


# ---------------------------------------------------------------------------
# Maps (optional enrichment - never restricts the main totals)
# ---------------------------------------------------------------------------
def collect_folders(folders, log=print):
    entries = eng.collect_pdf_entries(folders)
    return {"entries": entries, "status": f"Found {len(entries)} PDF(s) in {len(folders)} folder(s)."}


def scan(root, date_from=None, date_to=None, log=print):
    df_ = pd.Timestamp(date_from).to_pydatetime() if date_from else None
    dt_ = pd.Timestamp(date_to).to_pydatetime() if date_to else None
    entries, zone_count, pdf_count, no_zone = eng.find_workpack_zone_pdfs(root, date_from=df_, date_to=dt_, progress_cb=log)
    return {"entries": entries, "zone_count": zone_count, "pdf_count": pdf_count, "no_zone": no_zone,
            "status": f"Found {zone_count} Workpack zones folder(s), {pdf_count} PDF(s)."}


def _map_state(control_df, pdf_entries, filtered, log=print):
    entries = [e for e in (pdf_entries or []) if os.path.exists(e.get("path", ""))]
    if not entries:
        return None
    key = (id(control_df), tuple(sorted(e["path"] for e in entries)))
    if _MAPS["key"] != key:
        if eng.fitz is None:
            raise RuntimeError("PyMuPDF is not available, so maps can't be read.")
        vocab = eng.build_pole_vocabulary(control_df)
        _MAPS["pole_index"] = eng.build_map_pole_index_from_pdfs(entries, vocab, progress_cb=log)
        _MAPS["key"] = key
    pole_index = _MAPS["pole_index"]
    all_map_pole_enids = set()
    for s in pole_index.values():
        all_map_pole_enids |= s
    filtered_enids = set(filtered["enid"].astype(str).str.strip())
    matched_on_maps = filtered_enids & all_map_pole_enids
    not_on_maps = filtered_enids - all_map_pole_enids
    pole_summary_df = pd.DataFrame({"Pole (enid)": sorted(filtered_enids)})
    pole_summary_df["On Scanned Map"] = pole_summary_df["Pole (enid)"].isin(matched_on_maps).map({True: "Yes", False: "No"})
    pole_summary_df = pole_summary_df.sort_values(["On Scanned Map", "Pole (enid)"], ascending=[False, True]).reset_index(drop=True)
    return {"entries": entries, "pole_index": pole_index, "filtered_enids": filtered_enids,
            "matched": matched_on_maps, "not_on_maps": not_on_maps, "pole_summary": pole_summary_df}


def maps(master_path, control_path, filters=None, pdf_entries=None, log=print):
    control_df = _control(control_path)
    filtered, _c, _e, _f = _filtered(control_df, **(filters or {}))
    st_ = _map_state(control_df, pdf_entries, filtered, log=log)
    if st_ is None:
        return None
    map_rows = []
    for (day_outage, folder, fname), pole_set in st_["pole_index"].items():
        map_rows.append([_s(day_outage), _s(folder), _s(fname), len(pole_set), len(pole_set & st_["filtered_enids"])])
    return {
        "maps_scanned": len(st_["entries"]), "matched": len(st_["matched"]), "not_on_maps": len(st_["not_on_maps"]),
        "map_columns": ["Day/Outage", "Folder", "Map", "Poles found on map", "Poles also in your filter"], "map_rows": map_rows,
        "summary_columns": ["Pole (enid)", "On Scanned Map"],
        "summary_rows": st_["pole_summary"].values.tolist(),
    }


# ---------------------------------------------------------------------------
# Export tab (same workbook as before)
# ---------------------------------------------------------------------------
def export(master_path, control_path, out_path, filters=None, scope="All poles matching filters",
           pdf_entries=None, log=print):
    poles, free_issue, guk_items, guk_subitems, aliases = _materials(master_path)
    control_df = _control(control_path)
    filters = filters or {}
    filtered, _c, _e, date_field = _filtered(control_df, **filters)
    date_field_label = filters.get("date_field_label", "DateToUse")
    st_ = _map_state(control_df, pdf_entries, filtered, log=log) if pdf_entries else None
    matched_on_maps = st_["matched"] if st_ else set()
    pole_summary_df = st_["pole_summary"] if st_ else pd.DataFrame(columns=["Pole (enid)", "On Scanned Map"])

    if scope == "Only poles found on scanned maps":
        if not matched_on_maps:
            raise ValueError('No maps have been scanned yet (or none of your filtered poles were found on any scanned map) - '
                             'add maps on the Maps tab first, or switch back to "All poles matching filters".')
        export_scope_control_df = filtered[filtered["enid"].astype(str).str.strip().isin(matched_on_maps)]
        log(f"Exporting materials for {len(matched_on_maps)} pole(s) found on scanned maps (out of {len(filtered['enid'].unique())} matching your filters).")
    else:
        export_scope_control_df = filtered
    poles_out_x, fi_out_x, guk_out_x, guk_sub_out_x, unmatched_out_x = eng.build_breakdown(
        export_scope_control_df, poles, free_issue, guk_items, guk_subitems, aliases)

    poles_sheet = poles_out_x.rename(columns={
        "category": "Category", "material_item": "Pole Size", "description": "Description", "qty": "Qty",
    })[["Category", "Pole Size", "Description", "Qty"]].sort_values(["Category", "Pole Size"])
    fi_for_export = (
        fi_out_x.groupby(["category", "material_item", "description", "unit", "code"], dropna=False, as_index=False)["qty"].sum()
    )
    fi_sheet = fi_for_export.rename(columns={
        "category": "Type", "material_item": "Item", "description": "Description",
        "unit": "Unit", "code": "Commodity Code", "qty": "Qty",
    })[["Type", "Item", "Description", "Commodity Code", "Qty", "Unit"]].sort_values(["Type", "Item"])
    guk_sheet = guk_out_x.rename(columns={
        "category": "Type", "material_item": "Item", "description": "Description", "unit": "Unit", "qty": "Qty",
    })[["Type", "Item", "Description", "Qty", "Unit"]].sort_values(["Type", "Item"])
    guk_sub_sheet = guk_sub_out_x.rename(columns={
        "category": "Type", "material_item": "Parent Item", "description": "Description",
        "code": "Sub Code", "unit": "Unit", "qty_per_unit": "Qty Per Unit", "qty": "Qty",
    })[["Type", "Parent Item", "Sub Code", "Description", "Qty Per Unit", "Qty", "Unit"]].sort_values(["Type", "Parent Item"])

    def _actual_values(col):
        vals = export_scope_control_df[col].dropna().astype(str).str.strip()
        vals = sorted({v for v in vals if v and v.lower() != "nan"})
        return ", ".join(vals) if vals else "(none)"

    actual_dates = pd.to_datetime(export_scope_control_df[date_field], errors="coerce").dropna()
    date_start = actual_dates.min().strftime("%Y-%m-%d") if not actual_dates.empty else "(none)"
    date_end = actual_dates.max().strftime("%Y-%m-%d") if not actual_dates.empty else "(none)"
    filter_summary = {
        "Export scope": scope,
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
    with open(out_path, "wb") as fh:
        fh.write(excel_bytes)
    log(f"✅ Saved {os.path.basename(out_path)} ({len(export_scope_control_df):,} control-file rows)")
    return [out_path]
