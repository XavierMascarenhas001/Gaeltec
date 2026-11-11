"""Outage Calendar – browser version of the desktop "Outage Calendar Exporter".

Reads the High level planning workbook (every year sheet, header row 7, columns found by
their header names), overlays the Master Control File figures (poles, switches,
transformers) and the team breakdown (roster + absence CSVs), optionally enriches each
outage from its linked folder (Outage Request, Work Instructions, Generators PDFs,
Workpack Zones / Access Maps Returned), and writes:
  * an Excel month-by-month calendar (.xlsx)
  * an Outlook calendar (.ics) – one event per outage + Day / Week summaries

The same functions feed the Planning dashboard's Jobs calendar (hover + click details).
Nothing secret (absence link / token, staff names) is stored in this file: the roster and
the absence data come from the user and stay in their browser.
"""
from __future__ import annotations

import calendar as py_calendar
import html
import io
import os
import re
import uuid
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, time as dtime, timedelta, timezone
from urllib.parse import quote, unquote

# ════════════════════════════════════════════════════════════════════════════
#  Configuration (same as the desktop script)
# ════════════════════════════════════════════════════════════════════════════
HEADER_ROW = 7
YEAR_SHEET = re.compile(r"^\s*(20\d\d)\s*$")

# field -> (header names to look for, fallback column letter)
SOURCE_FIELDS = {
    "District": (("District",), "A"),
    "Outage Date": (("Outage Date", "Date"), "B"),
    "Weekday": (("Weekday",), "C"),
    "Outage Type": (("Outage Type",), "D"),
    "Scheme": (("Scheme",), "E"),
    "Outage #": (("Outage #", "Outage"), "F"),
    "Circuit": (("Circuit",), "G"),
    "Project": (("Project",), "H"),
    "Outage Value": (("Outage Value £", "Outage Value"), "K"),
    "PID": (("PID",), "L"),
    "SPEN PM": (("SPEN PM",), "M"),
    "POI": (("POI",), "N"),
    "Status": (("Status",), "P"),
}
COMMENT_FIELDS = [
    ("Access", "Q"), ("Live line poling", "R"), ("Jointers", "S"), ("Wayleaves", "T"),
    ("Generators", "U"), ("Battery Pods", "V"), ("Materials", "W"), ("Chasing Materials", "X"),
    ("Warehouse", "Y"), ("Traffic Management", "Z"), ("Finance", "AA"), ("Chasing Funds", "AB"),
    ("Prevet", "AC"), ("Permit 1", "AD"), ("Permit 2", "AE"), ("Permit 3", "AF"),
    ("Permit 4", "AG"), ("Permit 5", "AH"), ("Permit 6", "AI"), ("Permit 7", "AJ"),
    ("Permit 8", "AK"), ("Permit 9", "AL"),
]

BASE_UNC_PATH = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\1 - Outages Programme"

DISTRICT_COLORS = {"LANARK": "FF0000", "AYRSHIRE": "0000FF", "GLASGOW": "008000"}
FALLBACK_PALETTE = ["7C3AED", "D97706", "0891B2", "BE185D", "4D7C0F", "334155"]
DEFAULT_COLOR = "334155"
DISTRICT_STATUS_COLORS = {
    "AYRSHIRE": {"green": "C6F6D5", "at_risk": "38A169"},
    "LANARK": {"green": "BEE3F8", "at_risk": "2B6CB0"},
    "GLASGOW": {"green": "E9D8FD", "at_risk": "805AD5"},
}
CANCELED_COLOR = "FEB2B2"
DISTRICT_ALIASES = {"LANARKSHIRE": "LANARK"}

WEEKDAY_LABELS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
EVENT_COLUMN_WIDTH = 46
WEEK_BREAKDOWN_COLUMN_WIDTH = 32

# ---- Master Control File ----
MCF_DATE_COL = "plan1"
RECOVER_AH_POLE_DESC = ("recover 'a' / 'h' pole, up to and including 15 metres in height, "
                        "and reinstate, all ground conditions")
RECOVER_SINGLE_POLE_DESC = ("recover single pole, up to and including 15 metres in height, "
                            "and reinstate, all ground conditions")
TRANSFORMER_LABELS = {
    "Transformer 1ph 50kVA": "TX 1ph (50kVA)", "Transformer 3ph 50kVA": "TX 3ph (50kVA)",
    "Transformer 1ph 100kVA": "TX 1ph (100kVA)", "Transformer 1ph 25kVA": "TX 1ph (25kVA)",
    "Transformer 3ph 200kVA": "TX 3ph (200kVA)", "Transformer 3ph 100kVA": "TX 3ph (100kVA)",
}
SWITCH_CATEGORIES = {
    "Noja": ["Noja"],
    "Soule": ["11kV PMSW (Soule)"],
    "ABSW 11kV": ["11kv ABSW Hookstick Standard", "11kv ABSW Hookstick Spring loaded mech"],
    "ABSW 33kV": ["33kv ABSW Hookstick Dependant"],
}

# ---- folder lookups ----
GENERATORS_FOLDER_NAMES = {"generators", "generator", "gens", "gen", "genny", "gennys", "genset", "gensets"}
WORKPACK_ZONES_FOLDER_NAMES = {"workpack zones", "workpack zone", "work pack zones", "work pack zone"}
ACCESS_MAPS_FOLDER_NAMES = {"access maps returned", "access map returned"}


# ════════════════════════════════════════════════════════════════════════════
#  Small helpers
# ════════════════════════════════════════════════════════════════════════════
def _s(v) -> str:
    if v is None:
        return ""
    try:
        if v != v:            # NaN
            return ""
    except Exception:
        pass
    return str(v).strip()


def _canonical_district(d) -> str:
    key = str(d or "").strip().upper()
    return DISTRICT_ALIASES.get(key, key)


def _norm_header(v) -> str:
    return re.sub(r"[\s_]+", " ", str(v or "")).strip().lower()


def _col_letter(i: int) -> str:
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_index(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + ord(ch) - 64
    return n


def _to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return date(1899, 12, 30) + timedelta(days=int(v))
    t = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(t, fmt).date()
        except ValueError:
            pass
    return None


def _parse_time_value(v):
    if v is None or v == "":
        return None
    if isinstance(v, dtime):
        return v
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, (int, float)) and 0 <= float(v) < 1:
        mins = round(float(v) * 24 * 60)
        return dtime(mins // 60 % 24, mins % 60)
    text = str(v).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            pass
    return None


def _format_time_value(v) -> str:
    t = _parse_time_value(v)
    return t.strftime("%H:%M:%S") if t else _s(v)


def _format_outage_value(v) -> str:
    if v is None or v == "":
        return ""
    try:
        return f"\u00a3{float(v):,.2f}"
    except (TypeError, ValueError):
        return _s(v)


def _status_tier(status_text) -> str:
    s = str(status_text or "").strip().lower()
    if "cancel" in s:
        return "canceled"
    if "risk" in s:
        return "at_risk"
    return "green"


def _status_label(status_text) -> str:
    return {"canceled": "Cancelled", "at_risk": "At Risk", "green": "Green to go"}[_status_tier(status_text)]


# ════════════════════════════════════════════════════════════════════════════
#  Links (UNC folder -> file:// for Outlook)
# ════════════════════════════════════════════════════════════════════════════
def normalize_link(raw_target: str) -> str:
    if not raw_target:
        return raw_target
    target = raw_target.strip()
    if re.match(r"^(https?|mailto|ftp):", target, flags=re.IGNORECASE):
        return target
    target = re.sub(r"^file:/+", "", target, flags=re.IGNORECASE)
    target = unquote(target).replace("/", "\\")
    if target.startswith("\\\\") or re.match(r"^[A-Za-z]:\\", target):
        return target
    return f"{BASE_UNC_PATH}\\{target.lstrip(chr(92))}"


def _to_file_uri(p: str) -> str:
    p = p.strip()
    if p.startswith("\\\\"):
        return "file://" + quote(p[2:].replace("\\", "/"), safe="/:")
    return "file:///" + quote(p.replace("\\", "/"), safe="/:")


def _link_uri(path_or_url: str) -> str:
    if re.match(r"^https?:", path_or_url or "", flags=re.IGNORECASE):
        return path_or_url
    return _to_file_uri(path_or_url)


def _google_maps_search_url(query: str) -> str:
    return "https://www.google.com/maps/search/?api=1&query=" + quote(query, safe="")


# ════════════════════════════════════════════════════════════════════════════
#  Reading the High level planning workbook
# ════════════════════════════════════════════════════════════════════════════
_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _sheet_hyperlinks(z: zipfile.ZipFile, sheet_name: str, column_letter: str) -> dict:
    """{row: external link target} for one column of one sheet, straight from the XML."""
    try:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        rid = None
        for sh in wb_xml.find(f"{_NS_MAIN}sheets"):
            if sh.get("name") == sheet_name:
                rid = sh.get(f"{_NS_REL}id")
                break
        if not rid:
            return {}
        target = None
        for rel in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if not target:
            return {}
        sheet_path = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        sdir, sfile = sheet_path.rsplit("/", 1)
        rels_path = f"{sdir}/_rels/{sfile}.rels"
        if rels_path not in z.namelist():
            return {}
        rid_url = {rel.get("Id"): rel.get("Target") for rel in ET.fromstring(z.read(rels_path))
                   if rel.get("TargetMode") == "External"}
        if not rid_url:
            return {}
        out = {}
        # the <hyperlinks> block sits at the end of the sheet XML – find it without parsing the whole sheet
        xml = z.read(sheet_path).decode("utf8", "replace")
        blk = re.search(r"<hyperlinks>(.*?)</hyperlinks>", xml, re.S)
        if not blk:
            return {}
        for m in re.finditer(r"<hyperlink\b([^>]*)/?>", blk.group(1)):
            a = m.group(1)
            ref = re.search(r'\bref="([A-Z]+)(\d+)(?::[A-Z]+\d+)?"', a)
            r_id = re.search(r'\br:id="([^"]+)"', a) or re.search(r'\bid="([^"]+)"', a)
            if not ref or not r_id or ref.group(1) != column_letter:
                continue
            if r_id.group(1) in rid_url:
                out[int(ref.group(2))] = rid_url[r_id.group(1)]
        return out
    except (KeyError, ET.ParseError, zipfile.BadZipFile, OSError, AttributeError, TypeError):
        return {}


def read_programme(plain: bytes) -> list[dict]:
    """Every outage row of every year sheet (2024, 2025, 2026 …): the source fields,
    the planning-check comments, and the Outage # folder link."""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(plain), read_only=True, data_only=True)
    z = zipfile.ZipFile(io.BytesIO(plain))
    rows: list[dict] = []
    for name in wb.sheetnames:
        if not YEAR_SHEET.match(name):
            continue
        ws = wb[name]
        hdr_vals = next(ws.iter_rows(min_row=HEADER_ROW, max_row=HEADER_ROW, values_only=True), ())
        by_name = {}
        for i, v in enumerate(hdr_vals, start=1):
            k = _norm_header(v)
            if k and k not in by_name:
                by_name[k] = _col_letter(i)

        def find(names, letter):
            for n in names:
                if _norm_header(n) in by_name:
                    return by_name[_norm_header(n)]
            for n in names:   # "Outage Value £" written as "Outage Value\n£" etc.
                for k, col in by_name.items():
                    if k.startswith(_norm_header(n)):
                        return col
            return letter
        src = {f: find(names, letter) for f, (names, letter) in SOURCE_FIELDS.items()}
        if _norm_header(hdr_vals[_col_index(src["Outage Date"]) - 1] if len(hdr_vals) >= _col_index(src["Outage Date"]) else "") \
                not in ("outage date", "date"):
            continue      # not a planning sheet
        com = [(label, find((label,), letter)) for label, letter in COMMENT_FIELDS]
        idx = {f: _col_index(c) - 1 for f, c in src.items()}
        cidx = [(label, _col_index(c) - 1) for label, c in com]
        max_col = max(list(idx.values()) + [i for _, i in cidx]) + 1
        links = _sheet_hyperlinks(z, name, src["Outage #"])
        for r, vals in enumerate(ws.iter_rows(min_row=HEADER_ROW + 1, max_col=max_col, values_only=True),
                                 start=HEADER_ROW + 1):
            get = lambda i: vals[i] if i < len(vals) else None
            rec = {f: get(i) for f, i in idx.items()}
            d = _to_date(rec.get("Outage Date"))
            if d is None:
                continue
            rec["Outage Date"] = d
            rec["_comments"] = [(label, _s(get(i))) for label, i in cidx if _s(get(i))]
            rec["_outage_link"] = normalize_link(links[r]) if r in links else None
            rec["_sheet"], rec["_row"] = name, r
            rec["_id"] = f"{name}:{r}"
            rows.append(rec)
    wb.close()
    if not rows:
        raise RuntimeError("No outage rows found – the year sheets (2026 …) need the headers on row 7 "
                           "(District, Outage Date, Outage #, Circuit …).")
    return rows


# ════════════════════════════════════════════════════════════════════════════
#  Master Control File – per day and per outage (date + circuit) figures
# ════════════════════════════════════════════════════════════════════════════
def _safe_numeric(series):
    import pandas as pd
    s = series.astype(str).str.strip().str.replace("\u00a0", "", regex=False)
    s = s.str.replace(",", "", regex=False).str.replace("£", "", regex=False)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def mcf_frame(df, date_series=None, desc_col="description", qsub_col="qsub", total_col="total",
              circuit_col="circuit", district_col="district", project_col="project", job_col="job"):
    """Master rows -> a small frame with the date, circuit and each row's pole / switch /
    transformer label (vectorised: one label per unique description)."""
    import pandas as pd
    d = pd.to_datetime(date_series if date_series is not None else df[MCF_DATE_COL], errors="coerce").dt.normalize()
    desc = df[desc_col].astype("string").str.strip().str.lower().fillna("")
    uniq = pd.unique(desc)
    sw_pairs = [(p.lower(), lab) for lab, ps in SWITCH_CATEGORIES.items() for p in ps]
    tx_pairs = [(p.lower(), lab) for p, lab in TRANSFORMER_LABELS.items()]
    sw_map, tx_map, pole_mult = {}, {}, {}
    for u in uniq:
        sw_map[u] = next((lab for p, lab in sw_pairs if u.startswith(p)), None)
        tx_map[u] = next((lab for p, lab in tx_pairs if u.startswith(p)), None)
        pole_mult[u] = 1.0 if u == RECOVER_SINGLE_POLE_DESC else (2.0 if u == RECOVER_AH_POLE_DESC else 0.0)
    q = _safe_numeric(df[qsub_col]) if qsub_col in df.columns else pd.Series(0.0, index=df.index)
    out = pd.DataFrame({
        "date": d.values,
        "circuit": (df[circuit_col].astype("string").str.strip().str.upper().fillna("").values
                    if circuit_col in df.columns else ""),
        "district": (df[district_col].astype("string").str.strip().fillna("").values if district_col in df.columns else ""),
        "project": (df[project_col].astype("string").str.strip().fillna("").values if project_col in df.columns else ""),
        "job": (df[job_col].astype("string").str.strip().fillna("").values if job_col in df.columns else ""),
        "total": (_safe_numeric(df[total_col]).values if total_col in df.columns else 0.0),
        "qsub": q.values,
        "sw": desc.map(sw_map).values,
        "tx": desc.map(tx_map).values,
        "poles": (q * desc.map(pole_mult).astype(float)).values,
    })
    return out[out["date"].notna()]


def _grouped_stats(m, keys: list[str], with_job: bool = False) -> dict:
    """One pass of group-bys (fast): {key: stats} where key is the date, or (date, circuit)."""
    if m is None or m.empty:
        return {}
    base = m.groupby(keys, sort=False)[["total", "poles"]].sum()
    sw = m[m["sw"].notna()].groupby(keys + ["sw"], sort=False)["qsub"].sum()
    tx = m[m["tx"].notna()].groupby(keys + ["tx"], sort=False)["qsub"].sum()
    swd, txd = {}, {}
    for k, v in sw.items():
        if v:
            swd.setdefault(k[:-1] if len(keys) > 1 else k[0], {})[k[-1]] = float(v)
    for k, v in tx.items():
        if v:
            txd.setdefault(k[:-1] if len(keys) > 1 else k[0], {})[k[-1]] = float(v)
    jobs = {}
    if with_job:
        jj = m[m["job"] != ""].groupby(keys + ["job"], sort=False).size().reset_index(name="n")
        jj = jj.sort_values("n", ascending=False).drop_duplicates(keys)
        jobs = {(tuple(r[:-2]) if len(keys) > 1 else r[0]): r[-2] for r in jj.itertuples(index=False)}
    out = {}
    for k, row in zip(base.index, base.itertuples(index=False)):
        s_ = swd.get(k, {})
        t_ = txd.get(k, {})
        st = {"total": float(row.total), "pole_recover": float(row.poles), "switch_breakdown": s_,
              "switch_count": float(sum(s_.values())), "transformer_breakdown": t_,
              "has_tx_or_switch": bool(s_ or any(v > 0 for v in t_.values()))}
        if with_job:
            st["job"] = jobs.get(k, "")
        out[k] = st
    return out


def mcf_day_stats(m, districts=None, projects=None, years=None) -> dict:
    """{date: stats} like the desktop script's compute_master_control_stats."""
    if m is None or m.empty:
        return {}
    f = m
    if districts:
        allowed = {_canonical_district(x) for x in districts}
        f = f[f["district"].map(_canonical_district).isin(allowed)]
    if projects:
        allowed_p = {p.strip().lower() for p in projects}
        f = f[f["project"].str.lower().isin(allowed_p)]
    if years:
        f = f[f["date"].dt.year.isin({int(y) for y in years})]
    return {k.date(): v for k, v in _grouped_stats(f, ["date"]).items()}


def mcf_outage_stats(m) -> dict:
    """{(date, CIRCUIT): stats + most repeated job} – the figures for one outage."""
    if m is None or m.empty:
        return {}
    sub = m[m["circuit"] != ""]
    return {(k[0].date(), k[1]): v for k, v in _grouped_stats(sub, ["date", "circuit"], with_job=True).items()}


def _switch_lines(sb) -> list[str]:
    sb = sb or {}
    return [f"{lab}: {sb[lab]:g}" for lab in SWITCH_CATEGORIES if sb.get(lab)]


def _day_breakdown_lines(st, include_financial_total=True) -> list[str]:
    lines = []
    if include_financial_total:
        lines.append(f"Total: \u00a3{st['total']:,.2f}")
    lines.append(f"Poles: {st['pole_recover']:g}")
    lines.extend(_switch_lines(st.get("switch_breakdown")))
    for lab, c in sorted(st.get("transformer_breakdown", {}).items()):
        if c:
            lines.append(f"{lab}: {c:g}")
    return lines


def _week_mcf_summary(week, month, mcf_stats):
    total = pole = 0.0
    sw, tx, flagged = {}, {}, []
    for d in week:
        if d.month != month:
            continue
        st = mcf_stats.get(d) if mcf_stats else None
        if not st:
            continue
        total += st["total"]
        pole += st["pole_recover"]
        for k, v in st.get("switch_breakdown", {}).items():
            sw[k] = sw.get(k, 0) + v
        for k, v in st.get("transformer_breakdown", {}).items():
            tx[k] = tx.get(k, 0) + v
        if st.get("has_tx_or_switch"):
            flagged.append(d.day)
    return total, pole, sw, tx, flagged


# ════════════════════════════════════════════════════════════════════════════
#  Team roster + absence data
# ════════════════════════════════════════════════════════════════════════════
ROSTER_TEMPLATE = """# One group per line:  Group: name, name, *leader name, name
# Field teams are the groups named "Team A", "Team B" …  – mark the team leader with *
# Other groups used by the breakdown: Supervisors, Digger Drivers, Jointers, Pole Lorry
Supervisors: Name One, Name Two
Team A: Name, Name, *Leader Name, Name
Digger Drivers: Name, Name
Jointers: Name, Name
Pole Lorry: Name, Name
"""


def parse_roster(text: str) -> tuple[dict, dict]:
    """'Team A: x, *y, z' lines -> ({group: [names]}, {team: leader})."""
    roster, leaders = {}, {}
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        group, names = line.split(":", 1)
        group = re.sub(r"\s+", " ", group).strip()
        members = []
        for n in names.split(","):
            n = n.strip()
            if not n:
                continue
            if n.startswith("*"):
                n = n.lstrip("* ").strip()
                leaders[group] = n
            members.append(n)
        if group and members:
            roster[group] = members
    return roster, leaders


def _group(roster: dict, *names) -> list[str]:
    want = {n.lower() for n in names}
    for k, v in roster.items():
        if k.lower() in want:
            return v
    return []


def read_absence_csvs(paths: list[str]) -> dict:
    """Absence/training CSVs (user, team, role, start, end, days, absence) ->
    {person: [(start, end, reason), …]}."""
    import pandas as pd
    frames = []
    for p in paths:
        df = pd.read_csv(p, sep=None, engine="python", encoding="utf-8-sig", dtype=str)
        df.columns = [str(c).strip().lstrip("\ufeff").lower() for c in df.columns]
        frames.append(df)
    if not frames:
        return {}
    df = pd.concat(frames, ignore_index=True)
    missing = {"user", "start", "end"} - set(df.columns)
    if missing:
        raise RuntimeError("Absence CSV is missing the columns: " + ", ".join(sorted(missing))
                           + " (found: " + ", ".join(df.columns) + ")")

    def parse(col):
        iso = pd.to_datetime(df[col], format="%Y-%m-%d", errors="coerce")
        rest = iso.isna()
        if rest.any():
            iso.loc[rest] = pd.to_datetime(df.loc[rest, col], dayfirst=True, errors="coerce")
        return iso.dt.date
    st, en = parse("start"), parse("end")
    out: dict = {}
    reason = df["absence"] if "absence" in df.columns else pd.Series("", index=df.index)
    for person, s, e, why in zip(df["user"], st, en, reason):
        person = _s(person)
        if not person or s is None or e is None or s != s or e != e:
            continue
        out.setdefault(person, []).append((s, e, _s(why)))
    return out


def _is_absent(person, day, intervals):
    for s, e, reason in intervals.get(person, []):
        if s <= day <= e:
            return reason or "off"
    return None


def team_breakdown(day: date, roster: dict, leaders: dict, intervals: dict) -> dict:
    complete, incomplete, leaders_off, members_off = 0, 0, [], {}
    for team, members in roster.items():
        if not team.lower().startswith("team "):
            continue
        off = [m for m in members if _is_absent(m, day, intervals)]
        if off:
            incomplete += 1
            members_off[team] = off
        else:
            complete += 1
        ld = leaders.get(team)
        if ld and _is_absent(ld, day, intervals):
            leaders_off.append(ld)
    avail = lambda *g: sum(1 for m in _group(roster, *g) if not _is_absent(m, day, intervals))
    return {"complete_teams": complete, "incomplete_teams": incomplete, "leaders_off": leaders_off,
            "team_members_off": members_off,
            "digger_drivers": avail("Digger Drivers", "Digger Drives", "Digger Driver"),
            "jointers": avail("Jointers", "Jointer"),
            "lorry_drivers": avail("Pole Lorry", "Lorry Drivers", "Lory drivers"),
            "supervisors_off": [m for m in _group(roster, "Supervisors", "Supervisor") if _is_absent(m, day, intervals)]}


def _team_breakdown_lines(b: dict) -> list[str]:
    lines = [f"Complete Teams - {b['complete_teams']}", f"Uncomplete teams - {b['incomplete_teams']}"]
    lines.append(f"Team liders off - {len(b['leaders_off'])} - {', '.join(b['leaders_off'])}"
                 if b["leaders_off"] else "Team liders off - 0")
    for team in sorted(b.get("team_members_off", {})):
        lines.append(f"{team} off: {', '.join(b['team_members_off'][team])}")
    lines += [f"Digger Drivers - {b['digger_drivers']}", f"Jointers - {b['jointers']}",
              f"Lory drivers - {b['lorry_drivers']}"]
    lines.append(f"Supervisors off - {len(b['supervisors_off'])} - {', '.join(b['supervisors_off'])}"
                 if b.get("supervisors_off") else "Supervisors off - 0")
    return lines


# ════════════════════════════════════════════════════════════════════════════
#  Folder lookups (the browser hands us the files it found in each outage folder)
# ════════════════════════════════════════════════════════════════════════════
_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
HOSPITAL_LABEL_RE = re.compile(r"(?:closest|nearest)\s+(?:a\s*&\s*e\s+)?hospital\s*[:\-]?\s*", re.IGNORECASE)
W3W_LABEL_RE = re.compile(r"what\s*3\s*words\s*[:\-]?\s*", re.IGNORECASE)
_ALL_LABEL_RES = [HOSPITAL_LABEL_RE, W3W_LABEL_RE]
W3W_WORDS_RE = re.compile(r"(?:/{3}\s*)?([^\W\d_]+\.[^\W\d_]+\.[^\W\d_]+)", re.UNICODE)
W3W_SLASHED_RE = re.compile(r"/{3}\s*([^\W\d_]+\.[^\W\d_]+\.[^\W\d_]+)", re.UNICODE)


def _paragraph_text(p) -> str:
    parts = []
    for el in p.iter():
        if el.tag == _W_NS + "t" and el.text:
            parts.append(el.text)
        elif el.tag in (_W_NS + "tab", _W_NS + "br", _W_NS + "cr"):
            parts.append(" ")
    return re.sub(r"\s+", " ", "".join(parts)).strip()


def _docx_text_chunks(data: bytes) -> list[str]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(_W_NS + "body")
    chunks: list[str] = []

    def walk(container):
        for child in container:
            if child.tag == _W_NS + "p":
                t = _paragraph_text(child)
                if t:
                    chunks.append(t)
            elif child.tag == _W_NS + "tbl":
                for tr in child.findall(_W_NS + "tr"):
                    for tc in tr.findall(_W_NS + "tc"):
                        t = ", ".join(x for x in (_paragraph_text(p) for p in tc.iter(_W_NS + "p")) if x)
                        if t:
                            chunks.append(t)
            elif child.tag == _W_NS + "sdt":
                c = child.find(_W_NS + "sdtContent")
                if c is not None:
                    walk(c)
    if body is not None:
        walk(body)
    return chunks


def _cut_at_other_labels(text, own_re):
    cut = len(text)
    for rx in _ALL_LABEL_RES:
        if rx is not own_re:
            m = rx.search(text)
            if m:
                cut = min(cut, m.start())
    return text[:cut]


def _value_after_label(chunks, label_re):
    for i, chunk in enumerate(chunks):
        m = label_re.search(chunk)
        if not m:
            continue
        value = _cut_at_other_labels(chunk[m.end():], label_re).strip(" ,;:-\t")
        if not value and i + 1 < len(chunks) and not chunks[i + 1].rstrip().endswith(":"):
            value = _cut_at_other_labels(chunks[i + 1], label_re).strip(" ,;:-\t")
        if value:
            return value
    return None


def work_instructions_site_info(data: bytes) -> dict:
    try:
        chunks = _docx_text_chunks(data)
    except Exception:
        return {}
    info = {}
    hospital = _value_after_label(chunks, HOSPITAL_LABEL_RE)
    if hospital:
        info["hospital"] = hospital
        info["hospital_url"] = _google_maps_search_url(hospital)
    words = None
    v = _value_after_label(chunks, W3W_LABEL_RE)
    if v:
        m = W3W_WORDS_RE.search(v)
        if m:
            words = m.group(1)
    if not words:
        for c in chunks:
            m = W3W_SLASHED_RE.search(c)
            if m:
                words = m.group(1)
                break
    if words:
        words = words.lower()
        info["w3w"] = words
        info["w3w_url"] = "https://what3words.com/" + quote(words, safe=".")
    return info


def outage_request_cells(data: bytes) -> dict | None:
    from openpyxl import load_workbook
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    except Exception:
        return None
    try:
        sn = next((s for s in wb.sheetnames if s.strip().lower() in ("substations", "substation")), None)
        if sn is None:
            return None
        ws = wb[sn]
        return {k: ws[k].value for k in ("F4", "J4", "L4", "J6")}
    except Exception:
        return None
    finally:
        try:
            wb.close()
        except Exception:
            pass


_GEN_PREFIX_RE = re.compile(r"^gen_+", re.IGNORECASE)
_DDMMYYYY_RE = re.compile(r"(?<!\d)\d{2}-\d{2}-\d{4}(?!\d)")


def clean_generator_pdf_name(name: str) -> str:
    stem, ext = os.path.splitext(name)
    cleaned = _DDMMYYYY_RE.sub("", _GEN_PREFIX_RE.sub("", stem))
    cleaned = re.sub(r"([\s_\-])[\s_\-]+", r"\1", cleaned).strip(" _-")
    return (cleaned or stem) + ext


def enrich_row(row: dict, found: dict, file_dir: str = "") -> dict:
    """found (from the page): {request: rel file, wi: rel file, wi_rel: path in folder,
    gens: [pdf names], maps: [[title, [[name, rel path], …]], …]}."""
    status = {"substation": False, "generators": False, "work_instructions": False,
              "site_info": False, "workpack_zones": False, "access_maps": False}
    folder = row.get("_outage_link") or ""
    join = lambda rel: folder.rstrip("\\") + "\\" + rel.replace("/", "\\")
    if found.get("request"):
        try:
            cells = outage_request_cells(open(os.path.join(file_dir, found["request"]), "rb").read())
        except OSError:
            cells = None
        if cells:
            job, cust = _s(cells.get("F4")), _s(cells.get("J6"))
            st, en = _format_time_value(cells.get("J4")), _format_time_value(cells.get("L4"))
            lines = []
            if job:
                lines.append(("Jay number", job))
            if cust:
                lines.append(("Number of Customers", cust))
                row["_customers"] = cust
            if st or en:
                lines.append(("Start / End", f"{st or '?'}  -  {en or '?'}"))
            if lines:
                row["_outage_request"] = lines
                status["substation"] = True
            t0, t1 = _parse_time_value(cells.get("J4")), _parse_time_value(cells.get("L4"))
            if t0:
                row["_start_time"] = t0
            if t1:
                row["_end_time"] = t1
    if found.get("wi"):
        row["_work_instructions"] = join(found.get("wi_rel") or found["wi"])
        status["work_instructions"] = True
        try:
            info = work_instructions_site_info(open(os.path.join(file_dir, found["wi"]), "rb").read())
        except OSError:
            info = {}
        if info:
            row["_site_info"] = info
            status["site_info"] = True
    groups = []
    for title, items in found.get("maps") or []:
        if items:
            groups.append((title, [(n, join(rel)) for n, rel in items]))
            status["workpack_zones" if title.startswith("Workpack") else "access_maps"] = True
    if groups:
        row["_map_groups"] = groups
    gens = sorted(clean_generator_pdf_name(n) for n in found.get("gens") or [])
    if gens:
        row["_generator_pdfs"] = gens
        status["generators"] = True
    return status


# ════════════════════════════════════════════════════════════════════════════
#  Event description (shared by Outlook, Excel comment and the dashboard panel)
# ════════════════════════════════════════════════════════════════════════════
DETAIL_FIELDS = (("Outage #", "Outage #"), ("Project", "Project"), ("Outage Type", "Outage Type"),
                 ("PID", "PID"), ("SPEN PM", "SPEN PM"), ("POI", "POI"), ("Status", "Status"))
LINK_CAPTION_PREFIX = "Open "


def event_sections(row: dict) -> list:
    sections = []
    details = [(label, _s(row.get(f)), None) for label, f in DETAIL_FIELDS if _s(row.get(f))]
    ov = _format_outage_value(row.get("Outage Value"))
    if ov:
        details.append(("Outage Value", ov, None))
    info = row.get("_site_info") or {}
    if info.get("hospital"):
        details.append(("Closest Hospital", info["hospital"], None))
        details.append(("Hospital Map", "Open in Google Maps", info["hospital_url"]))
    if info.get("w3w"):
        details.append(("What3words", f"///{info['w3w']}", None))
        details.append(("Site (what3words)", "Open in what3words", info["w3w_url"]))
    if details:
        sections.append(("Outage Details", details))
    req = [(l, t, None) for l, t in row.get("_outage_request", []) if t]
    if req:
        sections.append(("Outage Request", req))
    checks = [(l, t, None) for l, t in row.get("_comments", []) if t]
    if checks:
        sections.append(("Planning Checks", checks))
    gens = row.get("_generator_pdfs") or []
    if gens:
        sections.append((f"Generators ({len(gens)})", [("", n, None) for n in gens]))
    docs = []
    if row.get("_work_instructions"):
        wi = row["_work_instructions"]
        docs.append(("Work Instructions", wi.replace("/", "\\").rsplit("\\", 1)[-1], _link_uri(wi)))
    if row.get("_outage_link"):
        docs.append(("Outage Folder", "Open outage folder", _link_uri(row["_outage_link"])))
    if docs:
        sections.append(("Documents", docs))
    for title, pdfs in row.get("_map_groups", []):
        sections.append((f"{title} ({len(pdfs)})", [("", n, _link_uri(p)) for n, p in pdfs]))
    return sections


def _event_lines(row: dict) -> list[str]:
    lines = [x for x in (_s(row.get("District")), _s(row.get("Project"))) if x]
    if _s(row.get("Circuit")):
        lines.append(f"Circuit {_s(row.get('Circuit'))}")
    lines += [x for x in (_s(row.get("Scheme")), _s(row.get("Outage #")), _format_outage_value(row.get("Outage Value"))) if x]
    t0, t1 = row.get("_start_time"), row.get("_end_time")
    if t0 and t1:
        lines.append(f"{t0.strftime('%H:%M')} - {t1.strftime('%H:%M')}")
    return lines or ["Outage"]


def _event_comment_text(row: dict) -> str:
    out = []
    for title, items in event_sections(row):
        if title.startswith(("Workpack Zones", "Access Maps Returned")):
            out.append(f"{title.rsplit(' (', 1)[0]}: {len(items)} PDF(s)")
            continue
        lines = [f"{l}: {t}" if l else f"- {t}" for l, t, u in items if not (u and t.startswith(LINK_CAPTION_PREFIX))]
        if lines:
            if out:
                out.append("")
            out.append(f"[{title.upper()}]")
            out.extend(lines)
    return "\n".join(out).strip()


# ════════════════════════════════════════════════════════════════════════════
#  Colours
# ════════════════════════════════════════════════════════════════════════════
def district_color_map(rows) -> dict:
    colors = dict(DISTRICT_COLORS)
    extra = []
    for r in rows:
        d = _canonical_district(r.get("District"))
        if d and d not in colors and d not in extra:
            extra.append(d)
    for i, d in enumerate(extra):
        colors[d] = FALLBACK_PALETTE[i % len(FALLBACK_PALETTE)]
    return colors


def _lighten_hex(h, factor=0.55):
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"{int(r + (255 - r) * factor):02X}{int(g + (255 - g) * factor):02X}{int(b + (255 - b) * factor):02X}"


def _darken_hex(h, factor=0.3):
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"{int(r * (1 - factor)):02X}{int(g * (1 - factor)):02X}{int(b * (1 - factor)):02X}"


def readable_text_color(fill_hex):
    r, g, b = int(fill_hex[0:2], 16), int(fill_hex[2:4], 16), int(fill_hex[4:6], 16)
    return "1E293B" if (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.6 else "FFFFFF"


def event_fill_color(colors, district, status_text):
    tier = _status_tier(status_text)
    if tier == "canceled":
        return CANCELED_COLOR
    d = _canonical_district(district)
    preset = DISTRICT_STATUS_COLORS.get(d)
    if preset:
        return preset["at_risk"] if tier == "at_risk" else preset["green"]
    base = colors.get(d, DEFAULT_COLOR)
    return _darken_hex(base) if tier == "at_risk" else _lighten_hex(base)


# ════════════════════════════════════════════════════════════════════════════
#  Excel calendar
# ════════════════════════════════════════════════════════════════════════════
def _estimate_row_height(groups, col_width=EVENT_COLUMN_WIDTH):
    per = max(8, int(col_width * 1.1))
    mx = 1
    for lines in groups:
        mx = max(mx, sum(max(1, -(-len(l) // per)) for l in lines if l))
    return max(20, 17 * mx + 6)


def build_calendar_workbook(rows, mcf_stats=None, filter_summary=None, mcf_enabled=None,
                            include_comments=True, team_ctx=None, include_financial_total=True):
    """team_ctx = (roster, leaders, intervals) or None."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    if mcf_enabled is None:
        mcf_enabled = bool(mcf_stats)
    mcf_stats = mcf_stats or {}
    team_on = bool(team_ctx)
    day_info = mcf_enabled or team_on

    header_font = Font(bold=True, color="FFFFFF", size=13)
    header_fill = PatternFill("solid", fgColor="1E3A8A")
    weekday_fill = PatternFill("solid", fgColor="334155")
    weekday_font = Font(bold=True, color="FFFFFF")
    day_num_font = Font(bold=True, size=13, color="1E293B")
    day_fill = PatternFill("solid", fgColor="F8FAFC")
    breakdown_fill = PatternFill("solid", fgColor="D9D9D9")
    flagged_fill = PatternFill("solid", fgColor="FEE2E2")
    outside_fill = PatternFill("solid", fgColor="E2E8F0")
    empty_fill = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="94A3B8")
    wrap = Alignment(wrap_text=True, vertical="top", horizontal="left")
    center = Alignment(horizontal="center", vertical="center")

    colors = district_color_map(rows)
    wb = Workbook()
    wb.remove(wb.active)

    lg = wb.create_sheet("Legend")
    for c, t in (("A1", "District"), ("B1", "On Track"), ("C1", "At Risk")):
        lg[c] = t
        lg[c].font, lg[c].fill = weekday_font, weekday_fill
    for i, d in enumerate(sorted(colors), start=2):
        lg.cell(row=i, column=1, value=d.title())
        lg.cell(row=i, column=2, value="").fill = PatternFill("solid", fgColor=event_fill_color(colors, d, "Green"))
        lg.cell(row=i, column=3, value="").fill = PatternFill("solid", fgColor=event_fill_color(colors, d, "At Risk"))
    cr = len(colors) + 2
    lg.cell(row=cr, column=1, value="(any district)")
    lg.cell(row=cr, column=2, value="Cancelled")
    lg.cell(row=cr, column=3, value="").fill = PatternFill("solid", fgColor=CANCELED_COLOR)
    lg.column_dimensions["A"].width, lg.column_dimensions["B"].width, lg.column_dimensions["C"].width = 22, 14, 14
    if mcf_enabled:
        lg["A" + str(lg.max_row + 2)] = ("Each day also has its own Day Breakdown row (Total and Poles for that day; a red "
                                         "highlight plus one line per switch type (" + " / ".join(SWITCH_CATEGORIES)
                                         + ") and one line per Transformer capacity whenever any were logged that day). "
                                         "The 8th column has the same figures rolled up for the whole week, plus which days were flagged.")
    if team_on:
        lg["A" + str(lg.max_row + 2)] = ("Each day's Day Breakdown cell also shows that day's Team Breakdown: how many of the "
                                         "field teams are fully staffed, which team leaders are off, and Digger Driver / "
                                         "Jointer / Pole Lorry availability, from the absence / training data.")
    if filter_summary:
        lg["A" + str(lg.max_row + 2)] = f"Filtered to: {filter_summary}"

    by_month = {}
    for r in rows:
        by_month.setdefault((r["Outage Date"].year, r["Outage Date"].month), []).append(r)
    month_keys = set(by_month) | ({(d.year, d.month) for d in mcf_stats} if mcf_stats else set())

    for (year, month) in sorted(month_keys):
        events_by_day = {}
        for r in by_month.get((year, month), []):
            events_by_day.setdefault(r["Outage Date"].day, []).append(r)
        for evs in events_by_day.values():
            evs.sort(key=lambda r: (_s(r.get("Outage #")), _s(r.get("Scheme"))))
        ws = wb.create_sheet(f"{py_calendar.month_abbr[month]} {year}"[:31])
        last_col = 8 if mcf_enabled else 7
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
        t = ws.cell(row=1, column=1, value=f"{py_calendar.month_name[month]} {year}")
        t.font, t.fill, t.alignment = header_font, header_fill, center
        ws.row_dimensions[1].height = 24
        for i, wd in enumerate(WEEKDAY_LABELS, start=1):
            c = ws.cell(row=2, column=i, value=wd)
            c.font, c.fill, c.alignment = weekday_font, weekday_fill, center
            ws.column_dimensions[get_column_letter(i)].width = EVENT_COLUMN_WIDTH
        if mcf_enabled:
            c = ws.cell(row=2, column=8, value="Week Breakdown")
            c.font, c.fill, c.alignment = weekday_font, weekday_fill, center
            ws.column_dimensions["H"].width = WEEK_BREAKDOWN_COLUMN_WIDTH

        row_idx = 3
        for week in py_calendar.Calendar(firstweekday=0).monthdatescalendar(year, month):
            week_start = row_idx
            max_events = max(max((len(events_by_day.get(d.day, [])) for d in week if d.month == month), default=0), 1)
            for ci, d in enumerate(week, start=1):
                c = ws.cell(row=row_idx, column=ci, value=d.day)
                c.border = Border(left=thin if ci == 1 else None, right=thin, top=thin)
                if d.month == month:
                    c.font, c.fill = day_num_font, day_fill
                else:
                    c.font, c.fill = Font(size=13, color="94A3B8"), outside_fill
                c.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_idx].height = 20
            row_idx += 1

            if day_info:
                def day_lines(d):
                    lines, flagged = [], False
                    st = mcf_stats.get(d) if mcf_enabled else None
                    if st:
                        lines.extend(_day_breakdown_lines(st, include_financial_total))
                        flagged = flagged or bool(st.get("has_tx_or_switch"))
                    if team_on:
                        tb = team_breakdown(d, *team_ctx)
                        if lines:
                            lines.append("")
                        lines.extend(_team_breakdown_lines(tb))
                        flagged = flagged or bool(tb["leaders_off"])
                    return lines, flagged
                cache = {d: day_lines(d) for d in week if d.month == month}
                max_lines = max([len(v[0]) for v in cache.values()] + [0])
                for ci, d in enumerate(week, start=1):
                    c = ws.cell(row=row_idx, column=ci)
                    c.border = Border(left=thin if ci == 1 else None, right=thin)
                    c.alignment = wrap
                    if d.month != month:
                        c.fill = outside_fill
                        continue
                    lines, flagged = cache[d]
                    if not lines:
                        c.fill = day_fill
                        continue
                    c.value = "\n".join(lines)
                    c.fill = flagged_fill if flagged else breakdown_fill
                    c.font = Font(bold=True, size=9, color="B91C1C" if flagged else "0F172A")
                ws.row_dimensions[row_idx].height = 6 if max_lines == 0 else 13 * max_lines + 6
                row_idx += 1

            for slot in range(max_events):
                last = slot == max_events - 1
                groups = []
                for ci, d in enumerate(week, start=1):
                    c = ws.cell(row=row_idx, column=ci)
                    c.alignment = wrap
                    c.border = Border(left=thin if ci == 1 else None, right=thin, bottom=thin if last else None)
                    if d.month != month:
                        c.fill = outside_fill
                        continue
                    evs = events_by_day.get(d.day, [])
                    if slot < len(evs):
                        ev = evs[slot]
                        lines = _event_lines(ev)
                        c.value = "\n".join(lines)
                        groups.append(lines)
                        fill = event_fill_color(colors, _s(ev.get("District")), ev.get("Status"))
                        link = ev.get("_outage_link")
                        c.fill = PatternFill("solid", fgColor=fill)
                        c.font = Font(color=readable_text_color(fill), bold=True, underline="single" if link else None)
                        if link:
                            c.hyperlink = link
                        if include_comments:
                            txt = _event_comment_text(ev)
                            if txt:
                                c.comment = Comment(txt, "Outages Calendar")
                    else:
                        c.fill = empty_fill
                ws.row_dimensions[row_idx].height = _estimate_row_height(groups)
                row_idx += 1

            if mcf_enabled:
                tot, pole, sw, tx, flagged = _week_mcf_summary(week, month, mcf_stats)
                lines = []
                if include_financial_total:
                    lines.append(f"Total: \u00a3{tot:,.2f}")
                lines.append(f"Poles: {pole:g}")
                lines.extend(_switch_lines(sw))
                lines += [f"{k}: {v:g}" for k, v in sorted(tx.items()) if v]
                lines.append(f"Flagged days: {', '.join(str(d) for d in sorted(flagged)) if flagged else 'none'}")
                end = row_idx - 1
                ws.merge_cells(start_row=week_start, start_column=8, end_row=end, end_column=8)
                c = ws.cell(row=week_start, column=8, value="\n".join(lines))
                c.font = Font(bold=True, size=9, color="B91C1C" if flagged else "0F172A")
                c.fill = flagged_fill if flagged else breakdown_fill
                c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                for r_ in range(week_start, end + 1):
                    ws.cell(row=r_, column=8).border = Border(left=thin, right=thin,
                                                             top=thin if r_ == week_start else None,
                                                             bottom=thin if r_ == end else None)
        ws.freeze_panes = "A3"
    return wb


# ════════════════════════════════════════════════════════════════════════════
#  Outlook (.ics)
# ════════════════════════════════════════════════════════════════════════════
def _fold_ics_line(line: str) -> str:
    if len(line.encode("utf-8")) <= 75:
        return line
    out, start, limit = [], 0, 75
    while start < len(line):
        end = min(start + limit, len(line))
        while len(line[start:end].encode("utf-8")) > limit and end > start:
            end -= 1
        out.append(line[start:end])
        start, limit = end, 74
    return "\r\n ".join(out)


def _esc_ics(text: str) -> str:
    return (text.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r\n", "\\n").replace("\n", "\\n"))


def _event_title(row) -> str:
    parts = [p for p in (_s(row.get("District")), _s(row.get("Scheme")), _s(row.get("Circuit"))) if p]
    return " | ".join(parts) if parts else "Outage"


def _event_body_plain(row) -> str:
    out = []
    for title, items in event_sections(row):
        if out:
            out.append("")
        out.append(f"\u2501\u2501\u2501 {title.upper()} \u2501\u2501\u2501")
        for l, t, u in items:
            if u and t.startswith(LINK_CAPTION_PREFIX):
                out.append(f"{l}: {u}")
            elif u:
                out.append(f"{l}: {t}" if l else f"\u2022 {t}")
                out.append(f"    {u}")
            else:
                out.append(f"{l}: {t}" if l else f"\u2022 {t}")
    return "\n".join(out)


_SEC = "margin:14px 0 4px 0;padding:3px 8px;background:#1E3A8A;color:#FFFFFF;font:bold 10pt Segoe UI,Arial,sans-serif;"
_TBL = "border-collapse:collapse;font:10pt Segoe UI,Arial,sans-serif;"
_LAB = "padding:2px 12px 2px 8px;color:#475569;font-weight:bold;vertical-align:top;white-space:nowrap;"
_VAL = "padding:2px 8px;color:#0F172A;vertical-align:top;"


def _event_body_html(row) -> str:
    parts = []
    for title, items in event_sections(row):
        parts.append(f'<div style="{_SEC}">{html.escape(title.upper())}</div>')
        trs = []
        for l, t, u in items:
            v = html.escape(t)
            if u:
                v = f'<a href="{html.escape(u)}">{v}</a>'
            trs.append(f'<tr><td style="{_LAB}">{html.escape(l)}</td><td style="{_VAL}">{v}</td></tr>' if l
                       else f'<tr><td colspan="2" style="{_VAL}">&bull; {v}</td></tr>')
        parts.append(f'<table style="{_TBL}">{"".join(trs)}</table>')
    return "<html><body>" + "".join(parts) + "</body></html>"


def _ics_vevent(uid, dtstart, dtend, summary, category, body_plain, body_html, url=None):
    now = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    L = ["BEGIN:VEVENT", f"UID:{uid}", f"DTSTAMP:{now}"]
    if isinstance(dtstart, datetime):
        L += [f"DTSTART:{dtstart:%Y%m%dT%H%M%S}", f"DTEND:{dtend:%Y%m%dT%H%M%S}"]
    else:
        L += [f"DTSTART;VALUE=DATE:{dtstart:%Y%m%d}", f"DTEND;VALUE=DATE:{dtend:%Y%m%d}"]
    L.append(_fold_ics_line(f"SUMMARY:{_esc_ics(summary)}"))
    if category:
        L.append(_fold_ics_line(f"CATEGORIES:{_esc_ics(category)}"))
    L.append(_fold_ics_line(f"DESCRIPTION:{_esc_ics(body_plain)}"))
    L.append(_fold_ics_line(f"X-ALT-DESC;FMTTYPE=text/html:{_esc_ics(body_html)}"))
    if url:
        L.append(_fold_ics_line(f"URL;VALUE=URI:{_link_uri(url)}"))
    L += ["TRANSP:TRANSPARENT", "END:VEVENT"]
    return L


def build_ics(rows, mcf_stats=None, mcf_enabled=False, team_ctx=None, include_financial_total=True,
              calendar_name="Outages Programme") -> str:
    mcf_stats = mcf_stats or {}
    team_on = bool(team_ctx)
    L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Outages Programme Export//EN", "CALSCALE:GREGORIAN",
         f"X-WR-CALNAME:{_esc_ics(calendar_name)}"]
    for r in rows:
        d = r["Outage Date"]
        cat = f"{_canonical_district(r.get('District')).title() or 'Unknown'} - {_status_label(r.get('Status'))}"
        t0, t1 = r.get("_start_time"), r.get("_end_time")
        if t0 and t1:
            s, e = datetime.combine(d, t0), datetime.combine(d, t1)
            if e <= s:
                e += timedelta(days=1)
        else:
            s, e = d, d + timedelta(days=1)
        L.extend(_ics_vevent(f"{uuid.uuid4()}@outages-programme", s, e, _event_title(r), cat,
                             _event_body_plain(r), _event_body_html(r), url=r.get("_outage_link")))
    if mcf_enabled or team_on:
        if team_on:
            months = {(r["Outage Date"].year, r["Outage Date"].month) for r in rows} | {(d.year, d.month) for d in mcf_stats}
            days = {date(y, m, dd) for (y, m) in months for dd in range(1, py_calendar.monthrange(y, m)[1] + 1)}
        else:
            days = set(mcf_stats)
        for d in sorted(days):
            lines, flagged = [], False
            st = mcf_stats.get(d) if mcf_enabled else None
            if st:
                lines.extend(_day_breakdown_lines(st, include_financial_total))
                flagged = flagged or bool(st.get("has_tx_or_switch"))
            if team_on:
                tb = team_breakdown(d, *team_ctx)
                if lines:
                    lines.append("")
                lines.extend(_team_breakdown_lines(tb))
                flagged = flagged or bool(tb["leaders_off"])
            if not lines:
                continue
            L.extend(_ics_vevent(f"{uuid.uuid4()}@outages-programme-day-summary", d, d + timedelta(days=1),
                                 "📊 Day Summary" + (" (Flagged)" if flagged else ""), "Day Summary",
                                 "\n".join(lines), "<html><body>" + "<br>".join(html.escape(x) for x in lines) + "</body></html>"))
    if mcf_enabled and mcf_stats:
        groups = {}
        for d in mcf_stats:
            iy, iw, _ = d.isocalendar()
            groups.setdefault((iy, iw), []).append(d)
        for (iy, iw), ds in sorted(groups.items()):
            tot = pole = 0.0
            sw, tx, fl = {}, {}, []
            for d in sorted(ds):
                st = mcf_stats[d]
                tot += st["total"]
                pole += st["pole_recover"]
                for k, v in st.get("switch_breakdown", {}).items():
                    sw[k] = sw.get(k, 0) + v
                for k, v in st.get("transformer_breakdown", {}).items():
                    tx[k] = tx.get(k, 0) + v
                if st.get("has_tx_or_switch"):
                    fl.append(d)
            ws_ = date.fromisocalendar(iy, iw, 1)
            we = ws_ + timedelta(days=6)
            lines = [f"Week of {ws_:%d %b %Y} - {we:%d %b %Y}"]
            if include_financial_total:
                lines.append(f"Total: \u00a3{tot:,.2f}")
            lines.append(f"Poles: {pole:g}")
            lines.extend(_switch_lines(sw))
            lines += [f"{k}: {v:g}" for k, v in sorted(tx.items()) if v]
            lines.append(f"Flagged days: {', '.join(d.strftime('%d %b') for d in sorted(fl)) if fl else 'none'}")
            L.extend(_ics_vevent(f"{uuid.uuid4()}@outages-programme-week-summary", ws_, we + timedelta(days=1),
                                 "📊 Week Summary" + (" (Flagged)" if fl else ""), "Week Summary",
                                 "\n".join(lines), "<html><body>" + "<br>".join(html.escape(x) for x in lines) + "</body></html>"))
    L.append("END:VCALENDAR")
    return "\r\n".join(L) + "\r\n"


# ════════════════════════════════════════════════════════════════════════════
#  Filters
# ════════════════════════════════════════════════════════════════════════════
def filter_rows(rows, districts=None, years=None, date_from=None, date_to=None, pms=None, projects=None):
    out = rows
    if projects:
        pr = {p.strip().lower() for p in projects}
        out = [r for r in out if _s(r.get("Project")).lower() in pr]
    if districts:
        allowed = {_canonical_district(d) for d in districts}
        out = [r for r in out if _canonical_district(r.get("District")) in allowed]
    if years:
        ys = {int(y) for y in years}
        out = [r for r in out if r["Outage Date"].year in ys]
    if pms:
        p = {x.strip().lower() for x in pms}
        out = [r for r in out if _s(r.get("SPEN PM")).lower() in p]
    if date_from:
        out = [r for r in out if r["Outage Date"] >= date_from]
    if date_to:
        out = [r for r in out if r["Outage Date"] <= date_to]
    return out


def outage_folders(rows) -> list[str]:
    """Unique linked folders (Windows paths) of these rows, in date order."""
    seen, out = set(), []
    for r in sorted(rows, key=lambda r: r["Outage Date"]):
        p = r.get("_outage_link")
        if p and not re.match(r"^(https?|mailto|ftp):", p, re.I) and p not in seen:
            seen.add(p)
            out.append(p)
    return out


# ════════════════════════════════════════════════════════════════════════════
#  Session caches shared by the Calendar tool and the Planning dashboard
# ════════════════════════════════════════════════════════════════════════════
class NeedPassword(RuntimeError):
    pass


_PROG: dict = {}
ENRICHED: dict = {}          # outage folder (Windows path) -> fields found in it
_MCF: dict = {}
_ABS: dict = {}
TEAM: dict = {"roster": "", "absence": []}


def programme(path: str, password: str = "") -> list[dict]:
    raw_head = open(path, "rb").read(4)
    enc = raw_head != b"PK\x03\x04"
    key = (path, os.path.getmtime(path), password if enc else "")
    if key not in _PROG:
        raw = open(path, "rb").read()
        if enc:
            if not password:
                raise NeedPassword("The High level planning workbook is password protected – type the password.")
            from tool_generator_sync import decrypt_if_needed
            try:
                plain, _ = decrypt_if_needed(raw, password)
            except Exception as e:
                raise NeedPassword(f"Couldn't open the High level planning workbook: {e}")
        else:
            plain = raw
        _PROG.clear()
        _PROG[key] = read_programme(plain)
    return _PROG[key]


def with_enrichment(rows: list[dict]) -> list[dict]:
    return [dict(r, **ENRICHED[r["_outage_link"]]) if r.get("_outage_link") in ENRICHED else r for r in rows]


def enrich_folders(found_by_folder: dict, file_dir: str) -> dict:
    """found_by_folder: {folder: found dict (see enrich_row)} -> totals; kept in ENRICHED."""
    totals = {}
    for folder, found in found_by_folder.items():
        tmp = {"_outage_link": folder}
        st = enrich_row(tmp, found or {}, file_dir)
        tmp.pop("_outage_link")
        ENRICHED[folder] = tmp
        for k, v in st.items():
            totals[k] = totals.get(k, 0) + int(v)
    return totals


def master_frame(path: str, date_field: str = "Plan1"):
    """Master parquet -> mcf_frame on the chosen date (Plan1 / Combined date / Done)."""
    import pandas as pd
    key = (path, os.path.getmtime(path), date_field)
    if key not in _MCF:
        cols = ["description", "qsub", "total", "plan1", "done", "district", "project", "circuit", "job"]
        try:
            import pyarrow.parquet as pq
            have = set(pq.read_schema(path).names)
            df = pd.read_parquet(path, columns=[c for c in cols if c in have])
        except Exception:
            df = pd.read_parquet(path)
        need = {"description", "qsub", "plan1"} - set(df.columns)
        if need:
            raise RuntimeError("The Master file is missing: " + ", ".join(sorted(need)))
        if date_field == "Done" and "done" in df.columns:
            ds = df["done"]
        elif date_field.lower().startswith("combined") and "done" in df.columns:
            from engine import combined_date
            ds = combined_date(df["done"], df["plan1"])
        else:
            ds = df["plan1"]
        _MCF.clear()
        _MCF[key] = mcf_frame(df, ds)
    return _MCF[key]


def team_context():
    """(roster, leaders, intervals) from the roster text + absence CSVs, or None."""
    roster, leaders = parse_roster(TEAM.get("roster") or "")
    paths = tuple(p for p in TEAM.get("absence") or [] if os.path.exists(p))
    if not roster or not paths:
        return None
    if paths not in _ABS:
        _ABS.clear()
        _ABS[paths] = read_absence_csvs(list(paths))
    return roster, leaders, _ABS[paths]
