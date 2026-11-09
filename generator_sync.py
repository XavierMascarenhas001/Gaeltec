"""
Generator Tracker  <->  High-level planning  sync tool
=====================================================

Run in JupyterLab:
    %run generator_sync.py          (or paste the whole file into a cell and run it)

Requirements (Anaconda already has most of them):
    pip install ipywidgets openpyxl pandas msoffcrypto-tool

Task 1  High-level planning  ->  Generator Tracker
    Copies District, Date, Weekday, Outage (Outage #), Circuit and PM (SPEN PM)
    of the selected days / outages into the first free rows at the bottom of the
    'Generators' table.

Task 2  Generator Tracker  ->  High-level planning
    For tracker rows whose 'Type of supply' is Generator, builds a line such as
        No Access, (x1) Generator from GWF, 3 phase 10 kVa, Setubal, 30 m of cable needed.
    and writes it into the 'Generators' column of the matching outage.
    For 'Battery pod' rows the Multiplier (number of pods) goes into 'Battery Pods'.
    Outages are matched on District, Date, Outage and Circuit (PM breaks ties).

Before anything is overwritten, the preview lists every cell / row that already
has information and asks: keep it, replace it, or add (below / as a new row).

How the files are written
    Both workbooks contain pivot tables, charts, timelines, images and external
    links. Saving through openpyxl would delete those, so this tool edits only the
    target cells directly inside the sheet XML and copies every other part of the
    file unchanged. Password-protected workbooks are decrypted in memory and the
    result is re-encrypted with the same password.
"""

from __future__ import annotations

import base64
import datetime as dt
import html
import io
import re
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook



def _msoffcrypto():
    """Imported only when a password-protected workbook is met
    (pip install msoffcrypto-tool)."""
    try:
        import msoffcrypto
        from msoffcrypto.format.ooxml import OOXMLFile
    except ImportError:
        raise RuntimeError("This workbook is password protected and the decryption package is missing. "
                           "Run:  pip install msoffcrypto-tool")
    return msoffcrypto, OOXMLFile


# ════════════════════════════════════════════════════════════════════════════
#  SETTINGS
# ════════════════════════════════════════════════════════════════════════════
TRACKER_SHEET = "Generators"          # sheet holding the 'Generators' table
TRACKER_HEADER_ROW = 2
PLANNING_HEADER_ROW = 7
PLANNING_SHEET_PATTERN = r"^\d{4}$"   # year sheets: 2024, 2025, 2026, 2027 ...

WEEKDAY_FORMULA = 'IF(B{r}="","",TEXT(B{r},"[$-en-US]dddd"))'
GENERATOR_TYPE = "generator"          # Type of supply -> text in planning 'Generators'
BATTERY_TYPE = "battery pod"          # Type of supply -> Multiplier in planning 'Battery Pods'
LINE_SEPARATOR = "\n"                 # between several generators on one outage
REFRESH_PIVOTS_ON_OPEN = True         # tracker Dashboard pivots refresh when opened

OUTPUT_DIR = Path.cwd()               # where updated files are saved


# ════════════════════════════════════════════════════════════════════════════
#  Small helpers
# ════════════════════════════════════════════════════════════════════════════
def _norm_header(h) -> str:
    return re.sub(r"\s+", " ", str(h or "")).strip().lower()


def find_col(headers: dict[str, str], *names: str, prefix: bool = False) -> str | None:
    """headers: normalised header -> column letter. Returns first match."""
    for n in names:
        n = _norm_header(n)
        if n in headers:
            return headers[n]
    if prefix:
        for n in names:
            n = _norm_header(n)
            for h, col in headers.items():
                if h.startswith(n):
                    return col
    return None


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def idx_to_col(i: int) -> str:
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return isinstance(v, str) and v.strip() == ""


def clean_text(v) -> str:
    """Value -> display text (40213.0 -> '40213')."""
    if is_blank(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def to_date(v) -> dt.date | None:
    if is_blank(v):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:  # Excel serial
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    try:
        d = pd.to_datetime(str(v), dayfirst=True, errors="coerce")
        return None if pd.isna(d) else d.date()
    except Exception:
        return None


def excel_serial(d: dt.date) -> int:
    return (d - dt.date(1899, 12, 30)).days


# key normalisation used for matching the two workbooks
def k_district(v) -> str:
    return clean_text(v).lower()


def k_outage(v) -> str:
    return re.sub(r"\s+", "", clean_text(v)).upper()


def k_circuit(v) -> str:
    c = re.sub(r"\s+", "", clean_text(v)).upper()
    return re.sub(r"^SP(?=\d)", "", c)  # 'SP40213' and 40213 are the same circuit


# ════════════════════════════════════════════════════════════════════════════
#  Encryption
# ════════════════════════════════════════════════════════════════════════════
def decrypt_if_needed(raw: bytes, password: str) -> tuple[bytes, bool]:
    """Returns (plain xlsx bytes, was_encrypted)."""
    if raw[:4] == b"PK\x03\x04":
        return raw, False
    msoffcrypto, _ = _msoffcrypto()
    office = msoffcrypto.OfficeFile(io.BytesIO(raw))
    try:
        office.load_key(password=password)
        out = io.BytesIO()
        office.decrypt(out)
    except Exception as exc:
        raise RuntimeError(f"Could not open the protected workbook with password '{password}': {exc}")
    return out.getvalue(), True


def encrypt(plain: bytes, password: str) -> bytes:
    _, OOXMLFile = _msoffcrypto()
    out = io.BytesIO()
    OOXMLFile(io.BytesIO(plain)).encrypt(password, out)
    return out.getvalue()


# ════════════════════════════════════════════════════════════════════════════
#  Reading the workbooks
# ════════════════════════════════════════════════════════════════════════════
def _headers(ws, header_row: int) -> dict[str, str]:
    out = {}
    for row in ws.iter_rows(min_row=header_row, max_row=header_row):
        for c in row:
            if c.value is not None and hasattr(c, "column_letter"):
                out.setdefault(_norm_header(c.value), c.column_letter)
    return out


def read_tracker(plain: bytes) -> dict:
    wb = load_workbook(io.BytesIO(plain), read_only=True, data_only=True)
    if TRACKER_SHEET not in wb.sheetnames:
        raise RuntimeError(f"Sheet '{TRACKER_SHEET}' not found in the Generator Tracker.")
    ws = wb[TRACKER_SHEET]
    hdr = _headers(ws, TRACKER_HEADER_ROW)
    cols = {
        "district": find_col(hdr, "District"),
        "date": find_col(hdr, "Date"),
        "weekday": find_col(hdr, "Weekday"),
        "outage": find_col(hdr, "Outage"),
        "circuit": find_col(hdr, "Circuit"),
        "pm": find_col(hdr, "PM"),
        "type": find_col(hdr, "Type of supply", prefix=True),
        "capacity": find_col(hdr, "Capacity (kVa)", "Capacity", prefix=True),
        "multiplier": find_col(hdr, "Multiplier", prefix=True),
        "isolation": find_col(hdr, "Isolation Point", prefix=True),
        "phase": find_col(hdr, "Phase"),
        "cable": find_col(hdr, "Cable length (m)", "Cable length", prefix=True),
        "provider": find_col(hdr, "Gen provider", prefix=True),
        "status": find_col(hdr, "Status"),
    }
    missing = [k for k in ("district", "date", "outage", "circuit", "pm") if not cols[k]]
    if missing:
        raise RuntimeError(f"Generator Tracker: columns not found: {missing}")
    max_col = max(col_to_idx(c) for c in hdr.values())
    idx = {k: col_to_idx(v) - 1 for k, v in cols.items() if v}

    rows, last_used = [], TRACKER_HEADER_ROW
    for r_i, vals in enumerate(
        ws.iter_rows(min_row=TRACKER_HEADER_ROW + 1, max_col=max_col, values_only=True),
        start=TRACKER_HEADER_ROW + 1,
    ):
        if any(not is_blank(v) for v in vals):
            last_used = r_i
        rec = {k: (vals[i] if i < len(vals) else None) for k, i in idx.items()}
        rec["_row"] = r_i
        rec["_date"] = to_date(rec.get("date"))
        if rec["_date"] or not is_blank(rec.get("outage")):
            rows.append(rec)
    wb.close()
    return {"cols": cols, "rows": rows, "last_used_row": last_used}


def read_planning(plain: bytes) -> dict:
    wb = load_workbook(io.BytesIO(plain), read_only=True, data_only=True)
    sheets, rows = {}, []
    for name in wb.sheetnames:
        if not re.match(PLANNING_SHEET_PATTERN, name):
            continue
        ws = wb[name]
        hdr = _headers(ws, PLANNING_HEADER_ROW)
        cols = {
            "district": find_col(hdr, "District"),
            "date": find_col(hdr, "Outage Date", "Date"),
            "weekday": find_col(hdr, "Weekday"),
            "scheme": find_col(hdr, "Scheme"),
            "outage": find_col(hdr, "Outage #", "Outage"),
            "circuit": find_col(hdr, "Circuit"),
            "pm": find_col(hdr, "SPEN PM"),
            "poi": find_col(hdr, "POI"),
            "generators": find_col(hdr, "Generators"),
            "battery": find_col(hdr, "Battery Pods", prefix=True),
        }
        if not (cols["date"] and cols["generators"] and cols["outage"]):
            continue  # not a planning sheet with the expected layout
        sheets[name] = cols
        idx = {k: col_to_idx(v) - 1 for k, v in cols.items() if v}
        max_col = max(idx.values()) + 1
        for r_i, vals in enumerate(
            ws.iter_rows(min_row=PLANNING_HEADER_ROW + 1, max_col=max_col, values_only=True),
            start=PLANNING_HEADER_ROW + 1,
        ):
            rec = {k: (vals[i] if i < len(vals) else None) for k, i in idx.items()}
            d = to_date(rec.get("date"))
            if not d:
                continue
            rec.update(_row=r_i, _sheet=name, _date=d)
            rows.append(rec)
    wb.close()
    if not sheets:
        raise RuntimeError("High-level planning: no year sheet with 'Outage Date' / 'Generators' headers on row 7.")
    return {"sheets": sheets, "rows": rows}


# ════════════════════════════════════════════════════════════════════════════
#  Matching
# ════════════════════════════════════════════════════════════════════════════
def k_pm(v) -> str:
    return re.sub(r"\s+", " ", clean_text(v)).lower()


def match_score(a: dict, b: dict) -> float | None:
    """None = incompatible. Otherwise how well two rows agree.
    * Date must be equal.
    * District / Outage / Circuit must be equal when filled on both sides
      (a blank on one side is ignored) – each agreement counts 1, and a blank
      on both sides counts a little (0.25) so 'no circuit' pairs with 'no circuit'.
    * PM is only a tie-breaker (+0.5 when equal): it separates two outages on
      the same day and circuit, but a spelling difference never blocks a match."""
    if a["_date"] != b["_date"]:
        return None
    score = 0.0
    for fn, key in ((k_district, "district"), (k_outage, "outage"), (k_circuit, "circuit")):
        x, y = fn(a.get(key)), fn(b.get(key))
        if x and y:
            if x != y:
                return None
            score += 1
        elif not x and not y:
            score += 0.25  # blank on both sides: weak agreement
    x, y = k_pm(a.get("pm")), k_pm(b.get("pm"))
    if x and x == y:
        score += 0.5
    elif not x and not y:
        score += 0.1
    return score


def best_match(item: dict, candidates: list[dict], require_same_outage: bool = False):
    """Returns (match or None, reason)."""
    scored = []
    for c in candidates:
        s = match_score(item, c)
        if s is None:
            continue
        if require_same_outage and k_outage(item.get("outage")) != k_outage(c.get("outage")):
            continue
        scored.append((s, c))
    if not scored:
        return None, "no matching outage in the other file"
    top = max(s for s, _ in scored)
    best = [c for s, c in scored if s == top]
    if len(best) > 1:
        return None, f"ambiguous – {len(best)} outages fit equally well"
    return best[0], "ok"


# ════════════════════════════════════════════════════════════════════════════
#  Generator description  ->  "No Access, (x1) Generator from GWF, 3 phase 10 kVa, Setubal, 30 m of cable needed."
# ════════════════════════════════════════════════════════════════════════════
def _num_txt(v) -> str:
    if is_blank(v):
        return ""
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else f"{f:g}"
    except (TypeError, ValueError):
        return str(v).strip()


def _is_num(s: str) -> bool:
    return bool(s) and s.replace(".", "", 1).isdigit()


def describe_generator(r: dict) -> str:
    """No Access, (x1) Generator from GWF, 3 phase 10 kVa, Setubal, 30 m of cable needed."""
    status, provider = clean_text(r.get("status")), clean_text(r.get("provider"))
    mult = _num_txt(r.get("multiplier"))
    mult = mult if _is_num(mult) and mult != "0" else "1"
    unit = f"(x{mult}) Generator" + (f" from {provider}" if provider else "")

    phase = clean_text(r.get("phase"))
    if phase and "phase" not in phase.lower():
        phase += " phase"
    cap = _num_txt(r.get("capacity"))
    cap_txt = f"{cap} kVa" if _is_num(cap) else cap
    spec = " ".join(p for p in (phase, cap_txt) if p)

    cable = _num_txt(r.get("cable"))
    if _is_num(cable):
        cable += " m of cable needed"
    elif cable:
        cable += " of cable needed"

    parts = [p for p in (status, unit, spec, clean_text(r.get("isolation")), cable) if p]
    return ", ".join(parts) + "."


def battery_count(r: dict) -> float:
    """Number of pods on a 'Battery pod' row = its Multiplier (blank -> 1)."""
    m = _num_txt(r.get("multiplier"))
    return float(m) if _is_num(m) else 1.0


# ════════════════════════════════════════════════════════════════════════════
#  Building the list of outages to choose from
# ════════════════════════════════════════════════════════════════════════════
def candidates_task1(plan: dict, start: dt.date, end: dt.date) -> list[dict]:
    out = [r for r in plan["rows"] if start <= r["_date"] <= end]
    out.sort(key=lambda r: (r["_date"], r["_sheet"], r["_row"]))
    for r in out:
        r["_id"] = f"P|{r['_sheet']}|{r['_row']}"
        r["_label"] = " · ".join([
            clean_text(r.get("outage")) or "—",
            clean_text(r.get("circuit")) or "no circuit",
            clean_text(r.get("district")),
            (clean_text(r.get("scheme"))[:45] or ""),
            "PM " + (clean_text(r.get("pm")) or "—"),
        ])
    return out


def candidates_task2(trk: dict, start: dt.date, end: dt.date) -> list[dict]:
    """One entry per outage (District/Date/Outage/Circuit/PM) holding its
    Generator and Battery pod rows."""
    groups: dict[tuple, dict] = {}
    for r in trk["rows"]:
        if not r["_date"] or not (start <= r["_date"] <= end):
            continue
        typ = clean_text(r.get("type")).lower()
        if typ not in (GENERATOR_TYPE, BATTERY_TYPE):
            continue
        key = (r["_date"], k_district(r.get("district")), k_outage(r.get("outage")),
               k_circuit(r.get("circuit")), k_pm(r.get("pm")))
        g = groups.setdefault(key, {
            "_date": r["_date"], "district": r.get("district"), "outage": r.get("outage"),
            "circuit": r.get("circuit"), "pm": r.get("pm"), "gens": [], "pods": [],
            "_id": "T|" + "|".join(map(str, key)),
        })
        (g["gens"] if typ == GENERATOR_TYPE else g["pods"]).append(r)
    out = sorted(groups.values(), key=lambda g: (g["_date"], (g["gens"] + g["pods"])[0]["_row"]))
    for g in out:
        what = []
        if g["gens"]:
            what.append(f"{len(g['gens'])} generator row(s)")
        if g["pods"]:
            what.append(f"{_num_txt(sum(battery_count(p) for p in g['pods']))} battery pod(s)")
        g["_label"] = " · ".join([
            clean_text(g["outage"]) or "—",
            clean_text(g["circuit"]) or "no circuit",
            clean_text(g["district"]) or "no district",
            "PM " + (clean_text(g["pm"]) or "—"),
            " + ".join(what),
            "tracker row " + ", ".join(str(x["_row"]) for x in g["gens"] + g["pods"]),
        ])
    return out


# ════════════════════════════════════════════════════════════════════════════
#  Plans – what would be written. Items flagged 'conflict' already have
#  information in the target and need a decision (keep / replace / ...).
# ════════════════════════════════════════════════════════════════════════════
T1_CHOICES = [("Keep existing row (don't add)", "keep"),
              ("Replace values in the existing row", "replace"),
              ("Add as a new row anyway", "add")]
T2_CHOICES = [("Keep existing text", "keep"),
              ("Replace with new text", "replace"),
              ("Add new text below existing", "append")]


def _trk_summary(t: dict) -> str:
    return " · ".join(x for x in (
        clean_text(t.get("district")), t["_date"].strftime("%d/%m/%Y") if t.get("_date") else "",
        clean_text(t.get("outage")), clean_text(t.get("circuit")), clean_text(t.get("pm"))) if x)


def plan_task1(selected: list[dict], trk: dict) -> dict:
    items = []
    for p in selected:
        existing, _ = best_match(p, trk["rows"], require_same_outage=True)
        if existing is None:  # several tracker rows fit -> use the first one as the conflict
            fits = [t for t in trk["rows"] if match_score(p, t) is not None
                    and k_outage(p.get("outage")) == k_outage(t.get("outage"))]
            existing = fits[0] if fits else None
        items.append({
            "p": p, "existing": existing, "conflict": existing is not None,
            "where": f"Tracker row {existing['_row']}" if existing else "new row",
            "current": _trk_summary(existing) if existing else "",
            "new": _trk_summary({**p}),
            "source": f"{p['_sheet']}!row {p['_row']}",
        })
    return {"task": 1, "items": items, "unmatched": []}


def _task1_cells(p: dict, row: int, cols: dict) -> dict:
    d = p["_date"]
    e = {
        cols["district"]: ("str", clean_text(p.get("district"))),
        cols["date"]: ("num", excel_serial(d)),
        cols["outage"]: ("str", clean_text(p.get("outage"))),
        cols["pm"]: ("str", clean_text(p.get("pm"))),
    }
    if cols.get("weekday"):
        e[cols["weekday"]] = ("formula", (WEEKDAY_FORMULA.format(r=row), d.strftime("%A")))
    circ = p.get("circuit")
    if isinstance(circ, (int, float)) and not is_blank(circ):
        e[cols["circuit"]] = ("num", _num_txt(circ))
    else:
        e[cols["circuit"]] = ("str", clean_text(circ))
    return e


def build_edits_task1(plan_res: dict, choices: dict[int, str], trk: dict) -> tuple[dict, list[dict]]:
    """Returns (edits {row: cells}, summary rows)."""
    cols = trk["cols"]
    edits, summary = {}, []
    next_row = trk["last_used_row"] + 1
    for i, it in enumerate(plan_res["items"]):
        action = choices.get(i, "keep") if it["conflict"] else "add"
        if action == "add":
            row, next_row = next_row, next_row + 1
        elif action == "replace":
            row = it["existing"]["_row"]
        else:
            summary.append({"Action": "kept existing", "Row": it["existing"]["_row"], "Outage": it["new"]})
            continue
        edits.setdefault(row, {}).update(_task1_cells(it["p"], row, cols))
        summary.append({"Action": "added" if action == "add" else "replaced", "Row": row, "Outage": it["new"]})
    return edits, summary


def plan_task2(selected_groups: list[dict], plan: dict) -> dict:
    targets: dict[tuple, dict] = {}
    unmatched = []
    for g in selected_groups:
        cands = [r for r in plan["rows"] if r["_date"] == g["_date"]]
        match, reason = best_match(g, cands)
        if not match:
            unmatched.append({"Date": g["_date"].strftime("%d/%m/%Y"), "Outage": clean_text(g["outage"]),
                              "Circuit": clean_text(g["circuit"]), "District": clean_text(g["district"]),
                              "PM": clean_text(g["pm"]), "Reason": reason})
            continue
        t = targets.setdefault((match["_sheet"], match["_row"]), {"match": match, "lines": [], "pods": 0.0, "n_pods": 0})
        t["lines"] += [x for x in (describe_generator(r) for r in g["gens"]) if x]
        t["pods"] += sum(battery_count(r) for r in g["pods"])
        t["n_pods"] += len(g["pods"])

    items = []
    for (sheet, row), t in sorted(targets.items()):
        m, scols = t["match"], plan["sheets"][sheet]
        base = {"sheet": sheet, "row": row, "date": m["_date"].strftime("%d/%m/%Y"),
                "outage": clean_text(m.get("outage")), "circuit": clean_text(m.get("circuit")),
                "district": clean_text(m.get("district")), "pm": clean_text(m.get("pm"))}
        if t["lines"]:
            new = LINE_SEPARATOR.join(t["lines"])
            cur = clean_text(m.get("generators"))
            items.append({**base, "field": "Generators", "col": scols["generators"],
                          "current": cur, "new": new, "value": ("str", new),
                          "conflict": bool(cur) and cur != new, "same": cur == new})
        if t["n_pods"]:
            if not scols.get("battery"):
                unmatched.append({**{k.title(): v for k, v in base.items() if k in ("date", "outage", "circuit", "district", "pm")},
                                  "Reason": f"sheet {sheet} has no 'Battery Pods' column"})
            else:
                new = _num_txt(t["pods"])
                cur = clean_text(m.get("battery"))
                items.append({**base, "field": "Battery Pods", "col": scols["battery"],
                              "current": cur, "new": new, "value": ("num", new),
                              "conflict": bool(cur) and _num_txt(cur) != new, "same": _num_txt(cur) == new})
    return {"task": 2, "items": items, "unmatched": unmatched}


def build_edits_task2(plan_res: dict, choices: dict[int, str]) -> tuple[dict, list[dict]]:
    edits, summary = {}, []
    for i, it in enumerate(plan_res["items"]):
        ref = f"{it['sheet']}!{it['col']}{it['row']}"
        if it["same"]:
            summary.append({"Action": "already up to date", "Cell": ref, "Text": it["new"]})
            continue
        action = choices.get(i, "keep") if it["conflict"] else "replace"
        if action == "keep":
            summary.append({"Action": "kept existing", "Cell": ref, "Text": it["current"]})
            continue
        if action == "append":
            value = ("str", it["current"] + LINE_SEPARATOR + it["new"])
        else:
            value = it["value"]
        edits.setdefault(it["sheet"], {}).setdefault(it["row"], {})[it["col"]] = value
        summary.append({"Action": "written" if action == "replace" else "added below", "Cell": ref, "Text": value[1]})
    return edits, summary


def preview_frame(plan_res: dict) -> pd.DataFrame:
    rows = []
    if plan_res["task"] == 1:
        for it in plan_res["items"]:
            p = it["p"]
            rows.append({
                "Status": "ALREADY IN TRACKER – choose below" if it["conflict"] else "will be added",
                "Where": it["where"], "District": clean_text(p.get("district")),
                "Date": p["_date"].strftime("%d/%m/%Y"), "Weekday": p["_date"].strftime("%A"),
                "Outage": clean_text(p.get("outage")), "Circuit": clean_text(p.get("circuit")),
                "PM": clean_text(p.get("pm")), "Existing tracker row": it["current"], "From": it["source"],
            })
    else:
        for it in plan_res["items"]:
            rows.append({
                "Status": ("already up to date" if it["same"] else
                           "HAS INFORMATION – choose below" if it["conflict"] else "will be written"),
                "Cell": f"{it['sheet']}!{it['col']}{it['row']}", "Field": it["field"],
                "Date": it["date"], "Outage": it["outage"], "Circuit": it["circuit"],
                "District": it["district"], "PM": it["pm"],
                "Current": it["current"], "New": it["new"],
            })
        for u in plan_res["unmatched"]:
            rows.append({"Status": "NOT WRITTEN – " + u.get("Reason", ""), **{k: v for k, v in u.items() if k != "Reason"}})
    return pd.DataFrame(rows)


# ════════════════════════════════════════════════════════════════════════════
#  Surgical XML writer (keeps pivots, charts, images, links … intact)
# ════════════════════════════════════════════════════════════════════════════
_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _esc(s: str) -> str:
    return _ILLEGAL_XML.sub("", html.escape(str(s), quote=False))


def _cell_xml(ref: str, style: str | None, kind: str, value) -> str:
    s = f' s="{style}"' if style else ""
    if kind == "str":
        if value in (None, ""):
            return f'<c r="{ref}"{s}/>'
        return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{_esc(value)}</t></is></c>'
    if kind == "num":
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    if kind == "formula":
        formula, cached = value
        return f'<c r="{ref}"{s} t="str"><f>{_esc(formula)}</f><v>{_esc(cached)}</v></c>'
    raise ValueError(kind)


_ROW_RE = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
_CELL_RE = re.compile(r"<c\b[^>]*?(?:/>|>.*?</c>)", re.S)


def _style_reference(xml: str, row_no: int) -> dict[str, str]:
    """col -> style id taken from a reference row (to style brand-new cells)."""
    m = re.search(rf'<row\b[^>]*\br="{row_no}"[^>]*?(?:/>|>.*?</row>)', xml, re.S)
    out = {}
    if m:
        for c in _CELL_RE.finditer(m.group(0)):
            tag = c.group(0)[: c.group(0).find(">") + 1]
            ref = re.search(r'\br="([A-Z]+)\d+"', tag)
            st = re.search(r'\bs="(\d+)"', tag)
            if ref and st:
                out[ref.group(1)] = st.group(1)
    return out


def patch_sheet_xml(xml: str, edits: dict[int, dict], style_row: int | None = None):
    """edits: {row: {col: (kind, value)}}. Returns (new_xml, log list, max_row)."""
    log = []
    ref_styles = _style_reference(xml, style_row) if style_row else {}
    sd_open = re.search(r"<sheetData\b[^>]*?(/>|>)", xml)
    if sd_open.group(1) == "/>":  # empty sheetData
        xml = xml[: sd_open.start()] + "<sheetData></sheetData>" + xml[sd_open.end():]
        sd_open = re.search(r"<sheetData\b[^>]*>", xml)
    sd_start = sd_open.end()
    sd_end = xml.index("</sheetData>", sd_start)
    body = xml[sd_start:sd_end]

    pending = dict(sorted(edits.items()))
    pieces, pos = [], 0
    for m in _ROW_RE.finditer(body):
        tag_end = m.group(0).find(">") + 1
        rtag = m.group(0)[:tag_end]
        r_no = int(re.search(r'\br="(\d+)"', rtag).group(1))
        # insert missing rows that come before this one
        for nr in [n for n in pending if n < r_no]:
            pieces.append(body[pos:m.start()]); pos = m.start()
            pieces.append(_build_row(nr, pending.pop(nr), ref_styles, log))
        if r_no in pending:
            pieces.append(body[pos:m.start()])
            pieces.append(_patch_row(m.group(0), r_no, pending.pop(r_no), ref_styles, log))
            pos = m.end()
    pieces.append(body[pos:])
    for nr in list(pending):
        pieces.append(_build_row(nr, pending.pop(nr), ref_styles, log))
    new_body = "".join(pieces)
    xml = xml[:sd_start] + new_body + xml[sd_end:]
    max_row = max([int(x) for x in re.findall(r'<row\b[^>]*?\br="(\d+)"', new_body)] or [0])
    return xml, log, max_row


def _build_row(r_no, cells, ref_styles, log):
    inner = "".join(
        _cell_xml(f"{c}{r_no}", ref_styles.get(c), k, v)
        for c, (k, v) in sorted(cells.items(), key=lambda kv: col_to_idx(kv[0]))
    )
    return f'<row r="{r_no}">{inner}</row>'


def _patch_row(row_xml, r_no, cells, ref_styles, log):
    tag_end = row_xml.find(">") + 1
    rtag = row_xml[:tag_end]
    self_closing = rtag.endswith("/>")
    rtag = re.sub(r'\sspans="[^"]*"', "", rtag)  # spans is only a hint; drop it
    if self_closing:
        rtag = rtag[:-2].rstrip() + ">"
        inner = ""
    else:
        inner = row_xml[tag_end:-len("</row>")]

    existing = []  # (col_idx, col, xml)
    for c in _CELL_RE.finditer(inner):
        ctag = c.group(0)[: c.group(0).find(">") + 1]
        col = re.search(r'\br="([A-Z]+)\d+"', ctag).group(1)
        existing.append((col_to_idx(col), col, c.group(0)))
    tail = _CELL_RE.sub("", inner)  # anything after the cells (e.g. extLst) – normally ''
    have = {col: x for _, col, x in existing}

    for col, (kind, val) in cells.items():
        old = have.get(col)
        if old is not None and "<f" in old:
            if kind == "formula":
                continue  # weekday formula already there – keep it
            log.append(f"{col}{r_no}: contains a formula – left unchanged")
            continue
        style = None
        if old is not None:
            st = re.search(r'\bs="(\d+)"', old[: old.find(">") + 1])
            style = st.group(1) if st else None
        if style is None:
            style = ref_styles.get(col)
        have[col] = _cell_xml(f"{col}{r_no}", style, kind, val)

    ordered = "".join(have[c] for c in sorted(have, key=col_to_idx))
    return f"{rtag}{ordered}{tail}</row>"


def _sheet_paths(z: zipfile.ZipFile) -> dict[str, str]:
    wb = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid_target = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        i, t = re.search(r'\bId="([^"]+)"', tag), re.search(r'\bTarget="([^"]+)"', tag)
        if i and t:
            target = t.group(1).lstrip("/")
            rid_target[i.group(1)] = target if target.startswith("xl/") else "xl/" + target
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        tag = m.group(0)
        name = html.unescape(re.search(r'\bname="([^"]*)"', tag).group(1))
        rid = re.search(r'\br:id="([^"]+)"', tag).group(1)
        out[name] = rid_target[rid]
    return out


def _sheet_table_paths(z: zipfile.ZipFile, sheet_path: str) -> list[str]:
    folder, fname = sheet_path.rsplit("/", 1)
    rel_path = f"{folder}/_rels/{fname}.rels"
    if rel_path not in z.namelist():
        return []
    rels = z.read(rel_path).decode("utf8")
    out = []
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        if "relationships/table" in tag:
            tgt = re.search(r'\bTarget="([^"]+)"', tag).group(1)
            out.append(str(Path(folder, tgt)).replace("\\", "/") if not tgt.startswith("/") else tgt.lstrip("/"))
    # normalise '../'
    norm = []
    for p in out:
        parts = []
        for seg in p.split("/"):
            if seg == "..":
                parts.pop()
            elif seg not in ("", "."):
                parts.append(seg)
        norm.append("/".join(parts))
    return norm


def _extend_ref(ref: str, max_row: int) -> str:
    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    if not m or int(m.group(4)) >= max_row:
        return ref
    return f"{m.group(1)}{m.group(2)}:{m.group(3)}{max_row}"


def write_workbook(plain: bytes, sheet_edits: dict[str, dict], style_rows: dict[str, int] | None = None,
                   grow_tables: bool = False, refresh_pivots: bool = False,
                   full_calc: bool = False) -> tuple[bytes, list[str]]:
    """sheet_edits: {sheet name: {row: {col: (kind, value)}}}"""
    style_rows = style_rows or {}
    zin = zipfile.ZipFile(io.BytesIO(plain))
    paths = _sheet_paths(zin)
    replaced: dict[str, bytes] = {}
    log: list[str] = []
    for sheet, edits in sheet_edits.items():
        if not edits:
            continue
        p = paths[sheet]
        xml = zin.read(p).decode("utf8")
        new_xml, lg, max_row = patch_sheet_xml(xml, edits, style_rows.get(sheet))
        log += [f"[{sheet}] {x}" for x in lg]
        # keep <dimension> covering the data
        new_xml = re.sub(r'(<dimension\b[^>]*\bref=")([^"]+)(")',
                         lambda m: m.group(1) + _extend_ref(m.group(2), max_row) + m.group(3), new_xml, count=1)
        replaced[p] = new_xml.encode("utf8")
        if grow_tables:
            for tp in _sheet_table_paths(zin, p):
                txml = zin.read(tp).decode("utf8")
                new_t = re.sub(r'(<(?:table|autoFilter)\b[^>]*?\bref=")([^"]+)(")',
                               lambda m: m.group(1) + _extend_ref(m.group(2), max_row) + m.group(3), txml)
                if new_t != txml:
                    replaced[tp] = new_t.encode("utf8")
                    log.append(f"[{sheet}] table extended to row {max_row}")
    if refresh_pivots:
        for n in zin.namelist():
            if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", n):
                x = zin.read(n).decode("utf8")
                if "refreshOnLoad=" not in x:
                    x = re.sub(r"<pivotCacheDefinition\b", '<pivotCacheDefinition refreshOnLoad="1"', x, count=1)
                    replaced[n] = x.encode("utf8")

    if full_calc:  # Excel recalculates every formula (e.g. Weekday) when the file is opened
        wbx = zin.read("xl/workbook.xml").decode("utf8")
        if "fullCalcOnLoad" not in wbx:
            wbx2 = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wbx, count=1)
            if wbx2 != wbx:
                replaced["xl/workbook.xml"] = wbx2.encode("utf8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replaced.get(item.filename, None)
            zout.writestr(item, data if data is not None else zin.read(item.filename))
    return out.getvalue(), log


# ════════════════════════════════════════════════════════════════════════════
#  Run a task end-to-end (also usable without the UI)
# ════════════════════════════════════════════════════════════════════════════
def apply_task1(tracker_plain: bytes, edits: dict) -> tuple[bytes, list[str]]:
    return write_workbook(
        tracker_plain, {TRACKER_SHEET: edits},
        style_rows={TRACKER_SHEET: TRACKER_HEADER_ROW + 1},
        grow_tables=True, refresh_pivots=REFRESH_PIVOTS_ON_OPEN, full_calc=True,
    )


def apply_task2(planning_plain: bytes, edits: dict) -> tuple[bytes, list[str]]:
    return write_workbook(planning_plain, edits)


def versioned_path(folder: Path, stem: str, suffix: str) -> Path:
    base = folder / f"{stem}_{dt.date.today():%Y-%m-%d}{suffix}"
    if not base.exists():
        return base
    v = 2
    while (p := folder / f"{stem}_{dt.date.today():%Y-%m-%d}_v{v}{suffix}").exists():
        v += 1
    return p


# ════════════════════════════════════════════════════════════════════════════
#  JupyterLab user interface
# ════════════════════════════════════════════════════════════════════════════
def _upload_payload(widget):
    """(name, bytes) from a FileUpload widget – works on ipywidgets 7 and 8."""
    v = widget.value
    if not v:
        return None, None
    if isinstance(v, dict):  # ipywidgets 7
        name, item = next(iter(v.items()))
        return name, bytes(item["content"])
    item = v[0]  # ipywidgets 8
    return item["name"], bytes(item["content"])


def _html_table(df: pd.DataFrame) -> str:
    if df.empty:
        return "<i>nothing</i>"
    shown = df.astype(object).where(df.notna(), "").map(lambda v: html.escape(str(v)).replace("\n", "<br>"))
    return shown.to_html(index=False, escape=False)


def launch():
    import ipywidgets as W
    from IPython.display import display, HTML

    state = {"trk": None, "plan": None, "trk_plain": None, "plan_plain": None,
             "trk_enc": False, "plan_enc": False, "trk_name": None, "plan_name": None,
             "days": {}, "task": None, "plan_res": None, "choice_w": {}}

    css = HTML("""<style>
      .gs-title{font-size:18px;font-weight:600;color:#1f3864;margin:4px 0 2px}
      .gs-sub{color:#555;margin-bottom:8px}
      .gs-conf{border-left:3px solid #e69500;padding:2px 8px;margin:4px 0;background:#fff8ec}
    </style>""")
    title = W.HTML('<div class="gs-title">Generator Tracker ⇄ High-level planning</div>'
                   '<div class="gs-sub">1 · upload both files (enter the password if protected)  '
                   '2 · choose task and dates  3 · tick days, then the outages inside them  '
                   '4 · preview and decide on existing information  5 · apply</div>')

    up_trk = W.FileUpload(accept=".xlsx,.xlsm", multiple=False, description="Generator Tracker",
                          layout=W.Layout(width="230px"))
    up_plan = W.FileUpload(accept=".xlsx,.xlsm", multiple=False, description="High-level planning",
                           layout=W.Layout(width="230px"))
    pwd = W.Password(value="", placeholder="file password", description="Password",
                     layout=W.Layout(width="240px"))
    btn_read = W.Button(description="Read files", button_style="info", icon="folder-open")
    files_msg = W.HTML()

    task = W.ToggleButtons(
        options=[("1 · Planning → Generator Tracker", 1), ("2 · Generator Tracker → Planning", 2)],
        value=1, style={"button_width": "260px"})
    today = dt.date.today()
    d_from = W.DatePicker(description="From", value=today)
    d_to = W.DatePicker(description="To", value=today + dt.timedelta(days=7))
    btn_list = W.Button(description="Show days & outages", button_style="primary", icon="calendar")

    btn_all = W.Button(description="Tick all days", layout=W.Layout(width="130px"))
    btn_none = W.Button(description="Clear all", layout=W.Layout(width="110px"))
    sel_box = W.VBox(layout=W.Layout(max_height="430px", overflow="auto",
                                     border="1px solid #ddd", padding="6px"))

    btn_prev = W.Button(description="Preview", button_style="warning", icon="eye")
    btn_apply = W.Button(description="Apply & save", button_style="success", icon="save", disabled=True)
    prev_out = W.Output()
    conflict_box = W.VBox()
    out = W.Output()

    # ---------------------------------------------------------------- reading
    def _status_line(which):
        if which == "trk":
            t = state["trk"]
            if not t:
                return "✖ Generator Tracker not loaded"
            return (f"✔ <b>{html.escape(state['trk_name'])}</b>{' (protected)' if state['trk_enc'] else ''} – "
                    f"{len(t['rows'])} rows, next free row <b>{t['last_used_row'] + 1}</b>")
        p = state["plan"]
        if not p:
            return "✖ High-level planning not loaded"
        return (f"✔ <b>{html.escape(state['plan_name'])}</b>{' (protected)' if state['plan_enc'] else ''} – "
                f"{len(p['rows'])} outages on sheets {', '.join(p['sheets'])}")

    def _read_one(which):
        widget = up_trk if which == "trk" else up_plan
        name, raw = _upload_payload(widget)
        if raw is None:
            return None
        if raw[:4] != b"PK\x03\x04" and not pwd.value:
            return f"🔒 <b>{html.escape(name)}</b> is password protected – type the password and click <b>Read files</b>."
        try:
            plain, enc = decrypt_if_needed(raw, pwd.value)
            data = read_tracker(plain) if which == "trk" else read_planning(plain)
        except Exception as exc:
            return f"<span style='color:#c00'>✖ {html.escape(name)}: {html.escape(str(exc))}</span>"
        state.update({which: data, f"{which}_plain": plain, f"{which}_enc": enc, f"{which}_name": name})
        return None

    def read_files(_=None, only=None):
        errors = []
        for which in (only,) if only else ("trk", "plan"):
            e = _read_one(which)
            if e:
                errors.append(e)
        files_msg.value = "<br>".join([_status_line("trk"), _status_line("plan")] + errors)
        reset_selection()

    def on_upload(which):
        def handler(change):
            # a new upload always replaces what was in memory (incl. results of an earlier Apply)
            state[which] = None
            read_files(only=which)
        return handler

    # ---------------------------------------------------------------- selection
    def reset_selection(*_):
        state.update(days={}, plan_res=None, choice_w={})
        sel_box.children = [W.HTML("<i>Click <b>Show days & outages</b>.</i>")]
        clear_preview()

    def clear_preview():
        state.update(plan_res=None, choice_w={})
        prev_out.clear_output()
        conflict_box.children = []
        btn_apply.disabled = True

    def build_selection(_=None):
        clear_preview()
        out.clear_output()
        if not (state["trk"] and state["plan"]):
            sel_box.children = [W.HTML("<span style='color:#c00'>Load both files first.</span>")]
            return
        a, b = d_from.value, d_to.value
        if not a or not b:
            sel_box.children = [W.HTML("<span style='color:#c00'>Choose both dates.</span>")]
            return
        a, b = min(a, b), max(a, b)
        cands = candidates_task1(state["plan"], a, b) if task.value == 1 else candidates_task2(state["trk"], a, b)
        state.update(task=task.value, days={})
        if not cands:
            sel_box.children = [W.HTML("<i>No outages in this date range"
                                       + (" with Type of supply = Generator / Battery pod" if task.value == 2 else "")
                                       + ".</i>")]
            return
        by_day: dict[dt.date, list] = {}
        for c in cands:
            by_day.setdefault(c["_date"], []).append(c)
        children = [W.HTML("<i>Tick a day to see its outages, then untick any outage you don't want.</i>")]
        for day, items in by_day.items():
            day_cb = W.Checkbox(value=False, indent=False, layout=W.Layout(width="auto"),
                                description=f"{day:%a %d %b %Y}  ({len(items)} outage{'s' if len(items) > 1 else ''})")
            item_cbs = [W.Checkbox(value=True, indent=False, description=it["_label"],
                                   layout=W.Layout(width="auto", margin="0 0 0 28px")) for it in items]
            inner = W.VBox(item_cbs, layout=W.Layout(display="none"))

            def show(change, inner=inner, item_cbs=item_cbs):
                inner.layout.display = "flex" if change["new"] else "none"
                if change["new"]:
                    for cb in item_cbs:
                        cb.value = True
                clear_preview()

            day_cb.observe(show, names="value")
            for cb in item_cbs:
                cb.observe(lambda _c: clear_preview(), names="value")
            state["days"][day] = (day_cb, list(zip(item_cbs, items)))
            children += [day_cb, inner]
        sel_box.children = children

    def set_all(val):
        for day_cb, _ in state["days"].values():
            day_cb.value = val

    def selected_items():
        return [it for day_cb, pairs in state["days"].values() if day_cb.value
                for cb, it in pairs if cb.value]

    # ---------------------------------------------------------------- preview
    def preview(_=None):
        clear_preview()
        out.clear_output()
        with prev_out:
            if not state["days"]:
                print("Click 'Show days & outages' first.")
                return
            if state["task"] != task.value:
                print("Task changed – click 'Show days & outages' again.")
                return
            items = selected_items()
            if not items:
                print("Tick at least one day (and outage).")
                return
            res = plan_task1(items, state["trk"]) if task.value == 1 else plan_task2(items, state["plan"])
            state["plan_res"] = res
            n_conf = sum(1 for it in res["items"] if it["conflict"])
            if task.value == 1:
                n_new = len(res["items"]) - n_conf
                display(HTML(f"<b>{n_new}</b> outage(s) will be added to the Generators table starting at row "
                             f"<b>{state['trk']['last_used_row'] + 1}</b>; <b>{n_conf}</b> already in the tracker."))
            else:
                n_w = sum(1 for it in res["items"] if not it["conflict"] and not it["same"])
                display(HTML(f"<b>{n_w}</b> empty cell(s) will be filled; <b>{n_conf}</b> cell(s) already have "
                             f"information; <b>{len(res['unmatched'])}</b> outage(s) could not be matched."))
            display(HTML(_html_table(preview_frame(res))))

        # one decision per item that already has information
        conflicts = [(i, it) for i, it in enumerate(res["items"]) if it["conflict"]]
        if conflicts:
            opts = [("— choose —", None)] + (T1_CHOICES if task.value == 1 else T2_CHOICES)
            set_all_dd = W.Dropdown(options=opts, value=None, description="Set all to",
                                    layout=W.Layout(width="420px"))
            rows = [W.HTML(f"<b style='color:#b26a00'>{len(conflicts)} item(s) already have information "
                           f"in the file – what do you want to do?</b>"), set_all_dd]
            for i, it in conflicts:
                if task.value == 1:
                    txt = (f"<b>{html.escape(it['where'])}</b> already has: {html.escape(it['current'])}"
                           f"<br>From planning: {html.escape(it['new'])}")
                else:
                    txt = (f"<b>{it['sheet']}!{it['col']}{it['row']}</b> ({it['field']}) · {it['date']} "
                           f"{html.escape(it['outage'])} {html.escape(it['circuit'])}"
                           f"<br>Current: {html.escape(it['current']).replace(chr(10), '<br>')}"
                           f"<br>New: {html.escape(it['new']).replace(chr(10), '<br>')}")
                dd = W.Dropdown(options=opts, value=None, layout=W.Layout(width="320px"))
                state["choice_w"][i] = dd
                rows.append(W.HBox([W.HTML(f"<div class='gs-conf'>{txt}</div>",
                                           layout=W.Layout(width="760px")), dd]))

            def _set_all(change):
                if change["new"] is not None:
                    for dd in state["choice_w"].values():
                        dd.value = change["new"]
            set_all_dd.observe(_set_all, names="value")
            conflict_box.children = rows
        btn_apply.disabled = not any(not it.get("same") for it in res["items"])

    # ---------------------------------------------------------------- apply
    def apply(_=None):
        res = state.get("plan_res")
        if not res:
            return
        out.clear_output()
        with out:
            undecided = [i for i, dd in state["choice_w"].items() if dd.value is None]
            if undecided:
                print(f"✖ Choose keep / replace for the {len(undecided)} item(s) that already have information "
                      f"(or use 'Set all to').")
                return
            choices = {i: dd.value for i, dd in state["choice_w"].items()}
            try:
                if res["task"] == 1:
                    edits, summary = build_edits_task1(res, choices, state["trk"])
                    if not edits:
                        print("Nothing to write – everything was kept as it is.")
                        return
                    plain, log = apply_task1(state["trk_plain"], edits)
                    final = encrypt(plain, pwd.value) if state["trk_enc"] else plain
                    name = state["trk_name"]
                    state.update(trk_plain=plain, trk=read_tracker(plain))
                else:
                    edits, summary = build_edits_task2(res, choices)
                    if not edits:
                        print("Nothing to write – everything was kept as it is.")
                        return
                    plain, log = apply_task2(state["plan_plain"], edits)
                    final = encrypt(plain, pwd.value) if state["plan_enc"] else plain
                    name = state["plan_name"]
                    state.update(plan_plain=plain, plan=read_planning(plain))
                path = versioned_path(OUTPUT_DIR, Path(name).stem + "_updated", Path(name).suffix)
                path.write_bytes(final)
                for line in log:
                    print("•", line)
                b64 = base64.b64encode(final).decode()
                display(HTML(
                    f"<p style='color:#1a7f37'><b>Saved:</b> {html.escape(str(path))}</p>"
                    f"<a download='{html.escape(path.name)}' "
                    f"href='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}'>"
                    f"⬇ Download {html.escape(path.name)}</a>" + _html_table(pd.DataFrame(summary))))
                files_msg.value = "<br>".join([_status_line("trk"), _status_line("plan")])
                clear_preview()
            except Exception as exc:
                print("✖ Failed:", exc)

    up_trk.observe(on_upload("trk"), names="value")
    up_plan.observe(on_upload("plan"), names="value")
    btn_read.on_click(lambda _: read_files())
    btn_list.on_click(build_selection)
    btn_all.on_click(lambda _: set_all(True))
    btn_none.on_click(lambda _: set_all(False))
    btn_prev.on_click(preview)
    btn_apply.on_click(apply)
    task.observe(reset_selection, names="value")
    d_from.observe(reset_selection, names="value")
    d_to.observe(reset_selection, names="value")

    reset_selection()
    ui = W.VBox([
        title,
        W.HBox([up_trk, up_plan, pwd, btn_read]), files_msg,
        W.HTML("<hr style='margin:6px 0'>"),
        task,
        W.HBox([d_from, d_to, btn_list]),
        W.HBox([btn_all, btn_none]),
        sel_box,
        W.HBox([btn_prev, btn_apply]),
        prev_out, conflict_box, out,
    ])
    display(css, ui)
    return ui


if __name__ == "__main__":
    try:
        get_ipython  # noqa: F821  – only launch the UI inside Jupyter/IPython
        launch()
    except NameError:
        print("Run this inside JupyterLab:  %run generator_sync.py")
