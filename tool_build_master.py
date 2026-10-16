"""Tool 2 - Build Master Parquet (CF_aggregated + Project Tracker + miscelaneous).

Original logic kept verbatim; only the tkinter file dialogs were replaced by
function arguments. The two versions of this script you had are merged:
`include_man_day=True` is the newer version (adds Man_day from the
miscelaneous file's 'man_day' column), `include_man_day=False` behaves
exactly like the older version.
"""
import pandas as pd


def move_columns_before(df, cols_to_move, before):
    """Move cols_to_move (in the given order) to sit immediately
    before the `before` column, leaving all other columns untouched."""
    cols_to_move = [c for c in cols_to_move if c in df.columns]
    if before not in df.columns:
        return df
    remaining = [c for c in df.columns if c not in cols_to_move]
    idx = remaining.index(before)
    for offset, col in enumerate(cols_to_move):
        remaining.insert(idx + offset, col)
    return df[remaining]


def run(aggregated_file, tracker_file, misc_file, output_parquet_file,
        include_man_day=True, log=print):
    # -----------------------------
    # LOAD DATA
    # -----------------------------
    log(f"📘 Loading aggregated: {aggregated_file}")
    agg_df = pd.read_parquet(aggregated_file)
    log(f"📘 Loading tracker:    {tracker_file}")
    tracker_df = pd.read_parquet(tracker_file)
    log(f"📘 Loading misc:       {misc_file}")
    misc_df = pd.read_parquet(misc_file)

    # -----------------------------
    # CLEAN COLUMN NAMES
    # -----------------------------
    for df in [agg_df, tracker_df, misc_df]:
        df.columns = df.columns.str.strip().str.lower()

    # -----------------------------
    # NORMALIZE KEYS
    # -----------------------------
    AGG_SHIRE_COL = 'district'
    AGG_PROJECT_COL = 'project'
    AGG_CIRCUIT_COL = 'circuit'
    AGG_JOB_COL = 'job'
    AGG_ITEM_COL = 'description'

    TRACKER_SHIRE_COL = 'shire'
    TRACKER_PROJECT_COL = 'project'
    TRACKER_SEGMENTCODE_COL = 'segmentcode'
    TRACKER_JOBNAME_COL = 'job name'

    for col in [AGG_SHIRE_COL, AGG_PROJECT_COL, AGG_CIRCUIT_COL]:
        if col in agg_df.columns:
            agg_df[col] = agg_df[col].astype(str).str.strip().str.lower()
    for col in [TRACKER_SHIRE_COL, TRACKER_PROJECT_COL, TRACKER_SEGMENTCODE_COL]:
        if col in tracker_df.columns:
            tracker_df[col] = tracker_df[col].astype(str).str.strip().str.lower()

    # text normalization
    if AGG_JOB_COL in agg_df.columns:
        agg_df[AGG_JOB_COL] = agg_df[AGG_JOB_COL].astype(str).str.lower()
    if TRACKER_JOBNAME_COL in tracker_df.columns:
        tracker_df[TRACKER_JOBNAME_COL] = tracker_df[TRACKER_JOBNAME_COL].astype(str).str.lower()

    # ensure output columns exist
    agg_df['pid'] = None
    agg_df['po'] = None

    # -----------------------------
    # GROUP TRACKER FOR FAST LOOKUP
    # -----------------------------
    tracker_groups = tracker_df.groupby([TRACKER_SHIRE_COL, TRACKER_PROJECT_COL, TRACKER_SEGMENTCODE_COL])

    # -----------------------------
    # MATCH PID + PO LOGIC
    # -----------------------------
    # Same rule as before - for each row, the first tracker row (in tracker
    # order) of its (district, project, circuit) group whose job name appears
    # in the row's job text gives the PID / PO - but each distinct
    # (group, job text) is only worked out once instead of once per row.
    log("🔗 Matching PID / PO against the Project Tracker...")
    tvals = tracker_df.values                 # the values iterrows() would give
    tcols = list(tracker_df.columns)

    def tcol(name):
        return tcols.index(name) if name in tcols else None

    j_job, j_pid, j_po = tcol(TRACKER_JOBNAME_COL), tcol('pid'), tcol('po')
    candidates = {}
    for key, pos in tracker_groups.indices.items():
        lst = []
        for p in sorted(pos):
            trow = tvals[p]
            job_name = str(trow[j_job] if j_job is not None else '').strip()
            lst.append((job_name, trow[j_pid] if j_pid is not None else None,
                        trow[j_po] if j_po is not None else None))
        candidates[key] = lst

    def column(name):
        return agg_df[name].tolist() if name in agg_df.columns else [None] * len(agg_df)

    pids, pos_ = [None] * len(agg_df), [None] * len(agg_df)
    agg_job = column(AGG_JOB_COL) if AGG_JOB_COL in agg_df.columns else [''] * len(agg_df)
    found = {}
    for i, key in enumerate(zip(column(AGG_SHIRE_COL), column(AGG_PROJECT_COL), column(AGG_CIRCUIT_COL))):
        cand = candidates.get(key)
        if cand is None:
            continue
        segment_text = str(agg_job[i])
        fk = (key, segment_text)
        hit = found.get(fk, found)           # `found` itself = not worked out yet
        if hit is found:
            hit = None
            for job_name, pid, po in cand:
                if job_name and job_name in segment_text:
                    hit = (pid, po)
                    break
            found[fk] = hit
        if hit is not None:
            pids[i], pos_[i] = hit
    agg_df['pid'] = pd.Series(pids, index=agg_df.index, dtype=object)
    agg_df['po'] = pd.Series(pos_, index=agg_df.index, dtype=object)
    log(f"   ✅ PID matched on {agg_df['pid'].notna().sum()} of {len(agg_df)} rows")

    # -----------------------------
    # MD POLING (+ MAN_DAY) FROM MISC FILE
    # -----------------------------
    if AGG_ITEM_COL in agg_df.columns and 'column_1' in misc_df.columns and 'column_2' in misc_df.columns:
        agg_df[AGG_ITEM_COL] = agg_df[AGG_ITEM_COL].astype(str).str.strip().str.lower()
        misc_df['column_1'] = misc_df['column_1'].astype(str).str.strip().str.lower()
        item_to_column_2 = misc_df.set_index('column_1')['column_2'].to_dict()
        agg_df['MD Poling'] = agg_df[AGG_ITEM_COL].map(item_to_column_2)
        log("   ✅ MD Poling added")

    if include_man_day:
        if AGG_ITEM_COL in agg_df.columns and 'column_1' in misc_df.columns and 'man_day' in misc_df.columns:
            agg_df[AGG_ITEM_COL] = agg_df[AGG_ITEM_COL].astype(str).str.strip().str.lower()
            misc_df['column_1'] = misc_df['column_1'].astype(str).str.strip().str.lower()
            item_to_man_day = misc_df.set_index('column_1')['man_day'].to_dict()
            agg_df['Man_day'] = agg_df[AGG_ITEM_COL].map(item_to_man_day)
            log("   ✅ Man_day added")
        else:
            log("   ⚠️ 'man_day' column not found in miscelaneous file - Man_day skipped")

    # -----------------------------
    # FINAL COLUMN ORDER FIX
    # -----------------------------
    if include_man_day:
        agg_df = move_columns_before(agg_df, ['MD Poling', 'Man_day', 'pid', 'po'], 'sourcefile')
    else:
        agg_df = move_columns_before(agg_df, ['MD Poling', 'pid', 'po'], 'sourcefile')

    # -----------------------------
    # SAVE OUTPUTS
    # -----------------------------
    if not output_parquet_file.lower().endswith(".parquet"):
        output_parquet_file += ".parquet"
    agg_df.to_parquet(output_parquet_file, index=False)
    xlsx_file = output_parquet_file.replace(".parquet", ".xlsx")
    import fast_xlsx
    fast_xlsx.write(xlsx_file, [("Sheet1", agg_df)])
    log(f"💾 Saved: {output_parquet_file}")
    log(f"💾 Saved: {xlsx_file}")
    log("Done!")
    return [output_parquet_file, xlsx_file]
