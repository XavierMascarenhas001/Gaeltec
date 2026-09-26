import os
import re
import io
from pathlib import Path
from datetime import datetime, date, time as dt_time

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# streamlit_calendar is an OPTIONAL third-party UI component (a calendar grid
# widget). If it isn't installed, the app still runs - the Outages Programme
# just falls back to a table view instead of crashing the whole dashboard.
try:
    from streamlit_calendar import calendar as st_calendar
    CALENDAR_LIB_AVAILABLE = True
except ImportError:
    st_calendar = None
    CALENDAR_LIB_AVAILABLE = False

st.set_page_config(page_title="Network Job Tracker", layout="wide")

# ============================================================
# NETWORK PATHS
# ============================================================
# These only work if the machine running this Streamlit app can actually
# see the \\gaeltec-gl share (i.e. running locally on-domain / on VPN, or
# on an internal server) - NOT on Streamlit Community Cloud, which has no
# access to internal UNC paths.

MASTER_DIR = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker"
OUTAGES_PATH = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\1 - Outages Programme\High-level_planning_2026.xlsx"
FORECAST_PATH = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\36. Workbank\2026\Service Partner Workbank_2026.xlsx"
ICS_DIR = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker\Calendar"

# Resolved relative to THIS script's own location (not the current working
# directory, which varies depending on how/where Streamlit is launched from).
# Deploy layout: put an "Images" folder directly next to this .py file.
SCRIPT_DIR = Path(__file__).resolve().parent
IMAGE_DIR = str(SCRIPT_DIR / "Images")

# Version suffix "(vN)" is OPTIONAL - a plain "Master_01-09-2026.parquet"
# with no version tag is a perfectly valid file (treated as v0) and must
# still be picked up, not silently skipped.
MASTER_PATTERN = re.compile(r"Master_(\d{2})-(\d{2})-(\d{4})(?:\s*\(v(\d+)\))?\.parquet$", re.IGNORECASE)
# Matches: outage_calendar_2026 (v6).ics - picks highest year, then highest
# version, so a stray "2027" file would win over "2026" regardless of vN,
# and within the same year the highest vN wins. Deploy new files without
# renaming and this always finds the newest one.
ICS_PATTERN = re.compile(r"outage_calendar_(\d{4})\s*\(v(\d+)\)\.ics$", re.IGNORECASE)


def find_latest_versioned_file(directory: str, glob_pattern: str, name_pattern: re.Pattern):
    """Generic 'find the newest by (group1, group2) parsed out of the
    filename' scanner, shared by find_latest_master() and find_latest_ics().
    name_pattern must have exactly 2 capture groups that parse to ints
    (e.g. year+version, or the master file's (date, version) after combining
    year/month/day into a sortable int). Returns None if the folder is
    unreachable or nothing matches - never raises."""
    best_path, best_key = None, None
    try:
        candidates = list(Path(directory).glob(glob_pattern))
    except OSError:
        return None
    for path in candidates:
        m = name_pattern.search(path.name)
        if not m:
            continue
        try:
            key = tuple(int(g) for g in m.groups())
        except ValueError:
            continue
        if best_key is None or key > best_key:
            best_key, best_path = key, path
    return best_path


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
            date_key = datetime(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        found.append((date_key, int(ver or 0), path))
    found.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return found


def find_latest_master(directory: str):
    """Convenience wrapper: just the single newest Path, or None."""
    all_masters = find_all_masters(directory)
    return all_masters[0][2] if all_masters else None


def find_latest_ics(directory: str):
    """Scans `directory` for outage_calendar_YYYY (vN).ics files and
    returns the Path of the one with the highest (year, version). Returns
    None if the folder is unreachable or nothing matches - never raises."""
    return find_latest_versioned_file(directory, "outage_calendar_*.ics", ICS_PATTERN)


# ============================================================
# .ICS PARSING (stdlib only - no external package required, so this never
# depends on a pip install succeeding on this machine/network)
# ============================================================
def _unfold_ics_lines(text: str) -> list:
    """RFC 5545 line folding: a line starting with a space or tab is a
    continuation of the previous line, not a new field. Unfold before
    parsing so multi-line SUMMARY/DESCRIPTION fields read correctly."""
    raw_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    unfolded = []
    for line in raw_lines:
        if line.startswith((" ", "\t")) and unfolded:
            unfolded[-1] += line[1:]
        else:
            unfolded.append(line)
    return unfolded


def _parse_ics_datetime(value: str):
    """Parses common ICS date/date-time forms: 'YYYYMMDD' (all-day) or
    'YYYYMMDDTHHMMSS[Z]'. Returns a datetime, a date, or the raw string
    if it doesn't match either (so the row still shows something instead
    of silently dropping the event)."""
    value = value.strip()
    for fmt in ("%Y%m%dT%H%M%SZ", "%Y%m%dT%H%M%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    try:
        return datetime.strptime(value, "%Y%m%d").date()
    except ValueError:
        return value


@st.cache_data(show_spinner="Reading outage calendar (.ics)...", max_entries=2, ttl=1800)
def parse_ics_events(file_bytes: bytes):
    """Extracts SUMMARY/DTSTART/DTEND/LOCATION/DESCRIPTION/UID from each
    VEVENT block using only the standard library. Returns (DataFrame,
    error_message); never raises - a malformed .ics shows a clean st.error
    rather than crashing the app."""
    try:
        text = file_bytes.decode("utf-8-sig", errors="replace")
        lines = _unfold_ics_lines(text)

        events = []
        current = None
        for line in lines:
            stripped = line.strip()
            if stripped == "BEGIN:VEVENT":
                current = {}
                continue
            if stripped == "END:VEVENT":
                if current is not None:
                    events.append(current)
                current = None
                continue
            if current is None or ":" not in line:
                continue

            key_part, _, value = line.partition(":")
            key = key_part.split(";")[0].strip().upper()

            if key == "SUMMARY":
                current["Summary"] = value.strip()
            elif key == "DTSTART":
                current["Start"] = _parse_ics_datetime(value)
            elif key == "DTEND":
                current["End"] = _parse_ics_datetime(value)
            elif key == "LOCATION":
                current["Location"] = value.strip()
            elif key == "DESCRIPTION":
                current["Description"] = value.strip().replace("\\n", " ").replace("\\,", ",")
            elif key == "UID":
                current["UID"] = value.strip()

        if not events:
            return pd.DataFrame(columns=["Summary", "Start", "End", "Location", "Description"]), None

        df = pd.DataFrame(events)
        for col in ["Summary", "Start", "End", "Location", "Description", "UID"]:
            if col not in df.columns:
                df[col] = ""
        df = df[["Summary", "Start", "End", "Location", "Description", "UID"]]
        df = df.sort_values("Start", key=lambda s: s.astype(str), na_position="last").reset_index(drop=True)
        return df, None
    except Exception as e:
        return None, str(e)


# ============================================================
# OUTAGES PROGRAMME
# ============================================================
@st.cache_data(show_spinner=False, max_entries=5, ttl=1800)
def build_calendar_events(cal_df: pd.DataFrame) -> list:
    """Vectorized FullCalendar event build (zip over plain Python lists,
    no .iterrows()) - cached so re-rendering the calendar on an unrelated
    rerun (a date click now costs a full script rerun) doesn't rebuild
    the same event list from scratch every time. Only used when
    CALENDAR_LIB_AVAILABLE is True."""
    palette = ["#2563eb", "#dc2626", "#059669", "#d97706", "#7c3aed", "#0891b2", "#be185d", "#4d7c0f"]
    districts_sorted = sorted(cal_df["District"].dropna().unique())
    district_colors = {d: palette[i % len(palette)] for i, d in enumerate(districts_sorted)}

    starts = cal_df["Outage Date"].dt.strftime("%Y-%m-%d").tolist()
    districts = cal_df["District"].tolist()
    schemes = cal_df["Scheme"].tolist()
    outage_nums = cal_df["Outage #"].tolist()
    circuits_evt = cal_df["Circuit"].tolist()
    pids_evt = cal_df["PID"].tolist()
    pms_evt = cal_df["SPEN PM"].tolist()
    pois_evt = cal_df["POI"].tolist()

    def _evt_str(v):
        return "" if pd.isna(v) or str(v).strip() == "" else str(v).strip()

    events = []
    for start, district, scheme, onum, circuit, pid, pm, poi in zip(
        starts, districts, schemes, outage_nums, circuits_evt, pids_evt, pms_evt, pois_evt
    ):
        d_str, s_str = _evt_str(district), _evt_str(scheme)
        parts = [d_str, s_str]
        for label, val in [("PID", pid), ("Circuit", circuit), ("Outage #", onum), ("PM", pm), ("POI", poi)]:
            v = _evt_str(val)
            if v:
                parts.append(f"{label} {v}")
        full_title = " — ".join(p for p in parts if p) or "Outage"
        color = district_colors.get(district, "#2563eb")
        events.append({
            "title": full_title,
            "start": start,
            "allDay": True,
            "backgroundColor": color,
            "borderColor": color,
        })
    return events


@st.cache_data(show_spinner="Reading outage programme...", max_entries=2, ttl=1800)
def load_outage_programme(file_bytes: bytes):
    """Cached on the file's bytes - re-parses only when the file content
    actually changes, not on every rerun/widget interaction.

    Returns (df, error_message). Never raises."""
    try:
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name="2026",
            header=6,                     # Excel row 7 is the header row (0-indexed = 6)
            usecols="A,B,C,E,F,G,L,M,N",  # District, Outage Date, Weekday, Scheme, Outage #, Circuit, PID, SPEN PM, POI
            engine="openpyxl",
        )
        df.columns = [
            "District", "Outage Date", "Weekday", "Scheme",
            "Outage #", "Circuit", "PID", "SPEN PM", "POI",
        ]
        df = df.dropna(how="all")
        df["Outage Date"] = pd.to_datetime(df["Outage Date"], errors="coerce")
        # PID mixes numbers with text like "Not delivered" - keep it as text
        # so Streamlit's table doesn't log an Arrow conversion error each time.
        df["PID"] = df["PID"].astype("string")
        return df, None
    except Exception as e:
        return None, str(e)


# ============================================================
# YOUR MAPPING DICTIONARIES
# ============================================================
CV7_erect = {
    "Erect Single HV/EHV Pole, up to and including 12 metre pole": "CV7 HV pole",
    "Erect Single HV/EHV Pole, up to and including 12 metre pole.": "CV7 HV pole",
}
CV7_erect_H = {
    "Erect Section Structure 'H' HV/EHV Pole, up to and including 12 metre pole.": "CV7 HV pole"
}
CV7_erect_lv = {
    "Erect LV Structure Single Pole, up to and including 12 metre pole": "CV7 LV pole",
}
CV7_recover = {
    "Recover single pole, up to and including 15 metres in height, and reinstate, all ground conditions": "CV7",
    "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions": "CV7 HV pole",
}
CV7_Tx = {
    "Erect pole mounted transformer up to 100kVA 1.ph.": "CV7 Tx",
    "Erect pole mounted transformer up to 200kVA 3.p.h.": "CV7 Tx",
    "Erect Voltage Regulator.": "CV7 Tx",
    "Erect Voltage Transformer (VT), RTU or Repeater": "CV7 Tx",
    "Erect 12kV/36kV Surge arrestors ( directly mounted ).": "CV7 Tx",
    "Remove pole mounted tranformer.": "CV7 Tx",
    "Remove platform mounted or 'H' pole mounted transformer.": "CV7 Tx",
}
transformer = {
    "Transformer 1ph 50kVA": "TX 1ph (50kVA)",
    "Transformer 3ph 50kVA": "TX 3ph (50kVA)",
    "Transformer 1ph 100kVA": "TX 1ph (100kVA)",
    "Transformer 1ph 25kVA": "TX 1ph (25kVA)",
    "Transformer 3ph 200kVA": "TX 3ph (200kVA)",
    "Transformer 3ph 100kVA": "TX 3ph (100kVA)",
}
CV7_OHL_CONDUCTOR_instal = {
    "Install bare conductor, run out, sag, terminate, bind in and connect jumpers; <100mm²": "CV7 OHL CONDUCTOR",
    "Install bare conductor, run out, sag, terminate, bind in and connect jumpers; >=100mm² <200mm²": "CV7 OHL CONDUCTOR",
    "Install conductor, run out, sag, terminate, clamp in and form jumper loops; >=200mm²": "CV7 OHL CONDUCTOR",
}
CV7_OHL_CONDUCTOR_recover = {
    "Recover overhead wire and fittings; HV/EHV overhead line or Hardex Pilot (1 conductor)": "CV7 OHL CONDUCTOR",
    "Recover overhead wire and fittings; HV/EHV overhead line or Hardex Pilot (2 conductor)": "CV7 OHL CONDUCTOR",
    "Recover overhead wire and fittings; HV/EHV overhead line or Hardex Pilot (3 conductor)": "CV7 OHL CONDUCTOR",
}
CV7_OHL_CONDUCTOR_LV_instal = {
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 2c": "CV7 OHL CONDUCTOR LV",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 4c": "CV7 OHL CONDUCTOR LV",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 2c + Earth": "CV7 OHL CONDUCTOR LV",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 4c + Earth": "CV7 OHL CONDUCTOR LV",
}
CV7_OHL_CONDUCTOR_LV_recover = {
    "Recover overhead wires and fittings; LV openwire overhead line (2 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV openwire overhead line (3 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV openwire overhead line (4 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV openwire overhead line (5 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 2 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 3 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 4 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 5 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover cleated service": "CV7 OHL CONDUCTOR LV",
}
Switch = {
    "Noja": "Noja",
    "11kV PMSW (Soule)": "11kV PMSW (Soule)",
    "11kv ABSW Hookstick Standard": "11kv ABSW Hookstick Standard",
    "11kv ABSW Hookstick Spring loaded mech": "11kv ABSW Hookstick Spring loaded mech",
    "33kv ABSW Hookstick Dependant": "33kv ABSW Hookstick Dependant",
}
Fuses = {
    "100A LV Fuse JPU 82.5mm": "100A LV Fuse JPU 82.5mm",
    "160A LV Fuse JPU 82.5mm": "160A LV Fuse JPU 82.5mm",
    "200A LV Fuse JPU 82.5mm": "200A LV Fuse JPU 82.5mm",
    "315A LV Fuse JPU 82.5mm": "315A LV Fuse JPU 82.5mm",
    "400A LV Fuse JPU 82.5mm": "400A LV Fuse JPU 82.5mm",
    "200A LV Fuse JSU 92mm": "200A LV Fuse JSU 92mm",
    "315A LV Fuse JSU 92mm": "315A LV Fuse JSU 92mm",
    "400A LV Fuse JSU 92mm": "400A LV Fuse JSU 92mm",
    "100A LV Fuse - Porcelain screw-in": "100A LV Fuse - Porcelain screw-in",
    "160A LV Fuse - Porcelain screw-in": "160A LV Fuse - Porcelain screw-in",
    "200A LV Fuse - Porcelain screw-in": "200A LV Fuse - Porcelain screw-in",
    "Single Phase cut out kit 100A Henley Series 7": "Single Phase cut out kit 100A Henley Series 7",
    "Three Phase cut out kit 100A Henley Series 7": "Three Phase cut out kit 100A Henley Series 7",
    "Three Phase 200A Cut out": "Three Phase 200A Cut out",
    "Cut out Fuse (MF) 60A": "Cut out Fuse (MF) 60A",
    "Cut out Fuse (MF) 80A": "Cut out Fuse (MF) 80A",
    "Cut out Fuse (MF) 100A": "Cut out Fuse (MF) 100A",
    "11KV FUSE UNIT - C-TYPE": "11KV FUSE UNIT - C-TYPE",
    "11KV SOLID LINK - C-TYPE": "11KV SOLID LINK - C-TYPE",
    "11KV OHL ASL C-TYPE RESET 20A 2 SHOT": "11KV OHL ASL C-TYPE RESET 20A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 25A 2 SHOT": "11KV OHL ASL C-TYPE RESET 25A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 40A 1 SHOT": "11KV OHL ASL C-TYPE RESET 40A 1 SHOT",
    "11KV OHL ASL C-TYPE RESET 40A 2 SHOT": "11KV OHL ASL C-TYPE RESET 40A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 63A 1 SHOT": "11KV OHL ASL C-TYPE RESET 63A 1 SHOT",
    "11KV OHL ASL C-TYPE RESET 63A 2 SHOT": "11KV OHL ASL C-TYPE RESET 63A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 63A 3 SHOT": "11KV OHL ASL C-TYPE RESET 63A 3 SHOT",
    "11KV OHL ASL C-TYPE RESET 100A 1 SHOT": "11KV OHL ASL C-TYPE RESET 100A 1 SHOT",
    "11KV OHL ASL C-TYPE RESET 100A 2 SHOT": "11KV OHL ASL C-TYPE RESET 100A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 100A 3 SHOT": "11KV OHL ASL C-TYPE RESET 100A 3 SHOT",
    "11KV OHL FUSE ELEMENT C-TYPE 15A": "11KV OHL FUSE ELEMENT C-TYPE 15A",
    "11KV OHL FUSE ELEMENT C-TYPE 25A": "11KV OHL FUSE ELEMENT C-TYPE 25A",
    "11KV OHL FUSE ELEMENT C-TYPE 30A": "11KV OHL FUSE ELEMENT C-TYPE 30A",
    "11KV OHL FUSE ELEMENT C-TYPE 40A": "11KV OHL FUSE ELEMENT C-TYPE 40A",
    "11KV OHL FUSE ELEMENT C-TYPE 50A": "11KV OHL FUSE ELEMENT C-TYPE 50A",
    "11KV OHL ASL DJP-TYPE 20A 2 SHOT": "11KV OHL ASL DJP-TYPE 20A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 25A 1 SHOT": "11KV OHL ASL DJP-TYPE 25A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 25A 2 SHOT": "11KV OHL ASL DJP-TYPE 25A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 40A 1 SHOT": "11KV OHL ASL DJP-TYPE 40A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 40A 2 SHOT": "11KV OHL ASL DJP-TYPE 40A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 63A 1 SHOT": "11KV OHL ASL DJP-TYPE 63A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 63A 2 SHOT": "11KV OHL ASL DJP-TYPE 63A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 63A 3 SHOT": "11KV OHL ASL DJP-TYPE 63A 3 SHOT",
    "11KV OHL ASL DJP-TYPE 100A 1 SHOT": "11KV OHL ASL DJP-TYPE 100A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 100A 2 SHOT": "11KV OHL ASL DJP-TYPE 100A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 100A 3 SHOT": "11KV OHL ASL DJP-TYPE 100A 3 SHOT",
    "11KV OHL FUSE ELEMENT DJP-TYPE 15A": "11KV OHL FUSE ELEMENT DJP-TYPE 15A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 25A": "11KV OHL FUSE ELEMENT DJP-TYPE 25A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 30A": "11KV OHL FUSE ELEMENT DJP-TYPE 30A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 40A": "11KV OHL FUSE ELEMENT DJP-TYPE 40A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 50A": "11KV OHL FUSE ELEMENT DJP-TYPE 50A",
}
CV31 = {
    "Replace / Fit safety or warning sign, number plates or name plate": "CV31",
    "Barbed Wire Wrap ACD (or Enhanced) single pole or stay - Replace/Repair": "CV31",
    "Steelwork bonding repair / fit.": "CV31",
    "Replace LV/HV/Earth guard missing / damaged.": "CV31",
}
CV8 = {
    "Tighten existing stay.": "CV8",
    "Erect/Replace stay above ground only.": "CV8",
    "Erect/Replace stay complete including block or driven type anchor": "CV8",
    "Erect/Replace stay complete including rock type anchor": "CV8",
    "Retrofit structure with Anchor Clamp fitting for Section / Angle / Terminal support": "CV8",
    "Erect Single Crossarm to single pole.": "CV8",
    "Erect Double Crossarm 'H' Pole formation": "CV8",
    "Remove Steelwork crossarm item only": "CV8",
    "Change 11kV Insulators to avoid contamination from old conductor": "CV8",
    "Change 33kV Insulators to avoid contamination from old conductor": "CV8",
    "Replace tension insulator, 11kV.": "CV8",
    "Replace tension insulator, 33kV.": "CV8",
    "Additional cost for fitting Stay Outrigger Bracket": "CV8",
    "Additional cost for fitting Angle / Terminal stay attachment plates on Heavy Construction as SP4009862": "CV8",
    "Recover and reinstate stay position,all ground conditions.": "CV8",
    "Fit foundation block to existing pole.": "CV8",
    "Fit bog shoe foundation to existing single pole.": "CV8",
    "Replace jumper / dropper mechanical connection with compression connection": "CV8",
    "Replace jumper / dropper with live line bail and flexible jumper conductor": "CV8",
    "Replace / Repair conductor with mid span joint using compression connection": "CV8",
    "Conductor repair; piece in conductor including compression joints": "CV8",
    "Bind In Conductors; 1.ph 11kV Intermediate / Pin Angle pole.": "CV8",
    "Bind In Conductors; 3.ph 11kV Intermediate / Pin Angle pole.": "CV8",
    "Conductor Terminations - 1.ph 11kV Section pole including jumpers.": "CV8",
    "Conductor Terminations - 3.ph 11kV Section pole including jumpers.": "CV8",
    "Conductor Terminations - 1.ph 11kV Terminal pole.": "CV8",
    "Conductor Terminations - 3.ph 11kV Terminal pole.": "CV8",
    "Unbind and reregulate existing conductors": "CV8",
    "Convert 1.ph 11kV Intermediate pole into Section Pole.": "CV8",
    "Convert 1.ph/3.p.h. 11kV line pole into Terminal Pole.": "CV8",
    "Convert 3.ph 11kV Intermediate pole into Section Pole.": "CV8",
    "Replace 11kV/33kV insulator pin and insulator, including unbinding and binding in": "CV8",
    "Replace 11kV/33kV insulator binder": "CV8",
    "Replace tension insulator, 11kV": "CV8",
    "Replace tension insulator, 33kV": "CV8",
    "Replace 11kV/33kV dead end termination": "CV8",
    "Additional cost for erection of pilot pin and insulator or pilot post insulator (11kV or 33kV)": "CV8",
    "Replace insulated conductor HV/LV earth above ground to first rod": "CV8",
    "Install Copper Covered Green / Yellow HV Earth or Black LV Earth to foot of pole": "CV8",
    "Install EHV/ HV Earth Electrode including excavate & reinstate (up to 8mtrs)": "CV8",
    "Install LV Earth Electrode including excavate & reinstate (up to 28mtrs)": "CV8",
    "Additional extra over for additional earthing excavated, laid & backfilled": "CV8",
    "Install Earth Electrode within cable trench": "CV8",
    "Erect 11kV Cable Termination ( incorporating surge arrestors )": "CV8",
    "Erect 33kV Cable Termination ( incorporating surge arrestors )": "CV8",
    "Steelwork bonding repair / fit": "CV8",
    "Erect 1.ph LV cable pole termination": "CV8",
    "Erect 3.ph LV cable pole termination": "CV8",
    "Remove 11kV/33kV Cable termination": "CV8",
    "Remove LV cable termination": "CV8",
    "Repair pole twist - including unbind / rebind.": "CV8",
}

POLE_CATEGORIES = {
    "CV7_erect": CV7_erect,
    "CV7_erect_H": CV7_erect_H,
    "CV7_erect_lv": CV7_erect_lv,
    "CV7_recover": CV7_recover,
}

POLE_DISPLAY_NAMES = {
    "CV7_erect": "CV7 (CV7_erect)",
    "CV7_erect_H": "CV7 H Pole (CV7_erect_H)",
    "CV7_erect_lv": "CV7 LV Pole (CV7_erect_lv)",
    "CV7_recover": "CV7 Recover (CV7_recover)",
}

ALL_CATEGORIES = {
    **POLE_CATEGORIES,
    "CV7_Tx": CV7_Tx,
    "transformer": transformer,
    "CV7_OHL_CONDUCTOR_instal": CV7_OHL_CONDUCTOR_instal,
    "CV7_OHL_CONDUCTOR_recover": CV7_OHL_CONDUCTOR_recover,
    "CV7_OHL_CONDUCTOR_LV_instal": CV7_OHL_CONDUCTOR_LV_instal,
    "CV7_OHL_CONDUCTOR_LV_recover": CV7_OHL_CONDUCTOR_LV_recover,
    "Switch": Switch,
    "Fuses": Fuses,
    "CV31": CV31,
    "CV8": CV8,
}

POLE_DEDUPE_CATEGORIES = {"CV8", "CV31"}

SWITCH_SUBTYPES = {
    "Noja": ["Noja"],
    "Soule": ["11kV PMSW (Soule)"],
    "ABSW": [
        "11kv ABSW Hookstick Standard",
        "11kv ABSW Hookstick Spring loaded mech",
        "33kv ABSW Hookstick Dependant",
    ],
}

# ============================================================
# IMAGE GROUPS for the Mapped Items tab
# Images live in the "Images" folder next to this script (see IMAGE_DIR above).
# ============================================================
CARD_GROUPS = [
    {
        "title": "Poles",
        "image": os.path.join(IMAGE_DIR, "Poles.png"),
        "categories": ["CV7_recover", "CV7_erect", "CV7_erect_H", "CV7_erect_lv"],
    },
    {
        "title": "Transformers",
        "image": os.path.join(IMAGE_DIR, "Transformer.png"),
        "categories": ["CV7_Tx", "transformer"],
    },
    {
        "title": "Conductor",
        "image": os.path.join(IMAGE_DIR, "Cable.png"),
        "categories": [
            "CV7_OHL_CONDUCTOR_instal",
            "CV7_OHL_CONDUCTOR_recover",
            "CV7_OHL_CONDUCTOR_LV_instal",
            "CV7_OHL_CONDUCTOR_LV_recover",
        ],
    },
    {
        "title": "Switch gear",
        "image": os.path.join(IMAGE_DIR, "Switchgear.png"),
        "subtypes": SWITCH_SUBTYPES,
    },
]

HV_POLE_KEY = "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions"
HV_POLE_MULTIPLIER = 2

# ============================================================
# HELPERS (aligned with the export script's normalization)
# ============================================================
def normalize_item(x):
    if pd.isna(x):
        return ""
    s = str(x).replace("\u200b", "").replace("\u200e", "").replace("\u200f", "").replace("\xa0", "").strip().upper()
    return re.sub(r"\s+", " ", s)


def normalize_item_series(s: pd.Series) -> pd.Series:
    """Vectorized equivalent of normalize_item(). .astype(object) forces
    plain Python strings (not pyarrow-backed) so \\uXXXX regex escapes
    work reliably across pandas/pyarrow versions."""
    out = s.where(s.notna(), "").astype(str).astype(object)
    out = out.str.replace(r"[\u200b\u200e\u200f\xa0]", "", regex=True)
    out = out.str.strip().str.upper()
    return out.str.replace(r"\s+", " ", regex=True)


def normalize_pole(p):
    if pd.isna(p):
        return ""
    s = str(p).replace("\u200b", "").replace("\u200e", "").replace("\u200f", "").replace("\xa0", "").strip().upper()
    return re.sub(r"\s+", "", s)


def normalize_pole_series(s: pd.Series) -> pd.Series:
    out = s.where(s.notna(), "").astype(str).astype(object)
    out = out.str.replace(r"[\u200b\u200e\u200f\xa0]", "", regex=True)
    out = out.str.strip().str.upper()
    return out.str.replace(r"\s+", "", regex=True)


def clean_job(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    text = re.sub(r'^[A-Za-z]\s*-\s*', '', text)
    m = re.search(r'map', text, flags=re.IGNORECASE)
    if m:
        text = text[: m.start()]
    text = re.sub(r'\bGSP\d+\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\bSP\d+\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\b\d+\b', '', text)
    text = re.sub(r'\s{2,}', ' ', text)
    text = re.sub(r'^[\s\-\u2013_,.:]+|[\s\-\u2013_,.:]+$', '', text)
    return text.strip()


def clean_job_series(s: pd.Series) -> pd.Series:
    out = s.where(s.notna(), "").astype(str).astype(object).str.strip()
    out = out.str.replace(r'^[A-Za-z]\s*-\s*', '', regex=True)
    out = out.str.replace(r'(?i)map.*$', '', regex=True)
    out = out.str.replace(r'(?i)\bGSP\d+\b', '', regex=True)
    out = out.str.replace(r'(?i)\bSP\d+\b', '', regex=True)
    out = out.str.replace(r'\b\d+\b', '', regex=True)
    out = out.str.replace(r'\s{2,}', ' ', regex=True)
    out = out.str.replace(r'^[\s\-\u2013_,.:]+|[\s\-\u2013_,.:]+$', '', regex=True)
    return out.str.strip()


UNIT_CONFIG = {
    "CV7_OHL_CONDUCTOR_recover": "m",
    "CV7_OHL_CONDUCTOR_LV_recover": "m",
    "CV7_OHL_CONDUCTOR_instal": "km",
    "CV7_OHL_CONDUCTOR_LV_instal": "km",
}


def format_length(value, native_unit):
    meters = value * 1000 if native_unit == "km" else value
    if meters >= 1000:
        return f"{meters / 1000:,.2f} km"
    return f"{meters:,.0f} m"


def show_total_banner(label, value_str):
    st.markdown(
        f"""
        <div style="text-align:center; padding: 4px 0 18px 0;">
            <div style="font-size:2.6rem; font-weight:800; color:#1e3a8a; line-height:1.15;">{value_str}</div>
            <div style="font-size:0.9rem; color:#475569; text-transform:uppercase; letter-spacing:0.05em;">{label}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False, max_entries=4, ttl=1800)
def build_pole_task_table(day_df: pd.DataFrame, item_col: str, qsub_col: str, pole_col) -> pd.DataFrame:
    if day_df.empty or item_col not in day_df.columns or qsub_col not in day_df.columns:
        return pd.DataFrame(columns=["Pole", "Task", "Qty", "Erect"])

    pole_vals = (
        day_df[pole_col].astype(str).str.strip()
        if pole_col and pole_col in day_df.columns
        else pd.Series("(no pole column mapped)", index=day_df.index)
    )
    out = pd.DataFrame({
        "Pole": pole_vals,
        "Task": day_df[item_col].astype(str).str.strip(),
        "Qty": day_df[qsub_col],
    })

    bad = {"", "nan", "none", "nat", "0"}
    out = out[~out["Pole"].str.lower().isin(bad) & ~out["Task"].str.lower().isin(bad)]
    if out.empty:
        return pd.DataFrame(columns=["Pole", "Task", "Qty", "Erect"])

    out["Erect"] = out["Task"].str.contains("erect", case=False, na=False).map({True: "Yes", False: ""})

    pole_order = out["Pole"].drop_duplicates().tolist()
    out["Pole"] = pd.Categorical(out["Pole"], categories=pole_order, ordered=True)
    out = out.sort_values("Pole", kind="stable").reset_index(drop=True)
    out["Pole"] = out["Pole"].astype(str)
    return out


@st.cache_data(show_spinner=False, max_entries=4, ttl=1800)
def render_pole_task_table_html(df: pd.DataFrame) -> str:
    if df.empty:
        return ""

    def esc(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    poles = df["Pole"].tolist()
    tasks = df["Task"].tolist()
    qtys = df["Qty"].tolist()
    erects = df["Erect"].tolist()
    n = len(poles)

    band_colors = ["#ffffff", "#eef2f7"]
    rows_html = []
    i, band = 0, 0
    while i < n:
        j = i
        while j + 1 < n and poles[j + 1] == poles[i]:
            j += 1
        span = j - i + 1
        bg = band_colors[band % 2]
        band += 1
        for k in range(i, j + 1):
            erect_style = "color:#dc2626;font-weight:600;" if erects[k] == "Yes" else ""
            cells = []
            if k == i:
                cells.append(
                    f"<td rowspan='{span}' style='background:{bg};font-weight:600;"
                    f"vertical-align:top;padding:6px 10px;border:1px solid #e3e7ee;"
                    f"white-space:nowrap;'>{esc(poles[k])}</td>"
                )
            cells.append(
                f"<td style='background:{bg};padding:6px 10px;border:1px solid #e3e7ee;{erect_style}'>"
                f"{esc(tasks[k])}</td>"
            )
            cells.append(
                f"<td style='background:{bg};padding:6px 10px;border:1px solid #e3e7ee;"
                f"text-align:right;white-space:nowrap;'>{esc(qtys[k])}</td>"
            )
            cells.append(
                f"<td style='background:{bg};padding:6px 10px;border:1px solid #e3e7ee;"
                f"text-align:center;'>{esc(erects[k])}</td>"
            )
            rows_html.append("<tr>" + "".join(cells) + "</tr>")
        i = j + 1

    header = (
        "<tr style='background:#1e3a8a;color:#ffffff;'>"
        "<th style='padding:8px 10px;text-align:left;'>Pole</th>"
        "<th style='padding:8px 10px;text-align:left;'>Task (MD Poling)</th>"
        "<th style='padding:8px 10px;text-align:right;'>Qty</th>"
        "<th style='padding:8px 10px;text-align:center;'>Erect</th>"
        "</tr>"
    )
    return (
        "<div style='max-height:480px;overflow-y:auto;border:1px solid #e3e7ee;border-radius:8px;'>"
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>"
        f"<thead>{header}</thead><tbody>{''.join(rows_html)}</tbody></table></div>"
    )


@st.cache_data(show_spinner="Loading latest master file from network...", max_entries=2, ttl=300)
def load_master_from_network(path_str: str):
    """Returns (df, error_message). Never raises.
    ttl=300 (5 min): the file list can change between sessions - use the
    sidebar "Refresh from network" button for an immediate re-check."""
    try:
        p = Path(path_str)
        if p.suffix.lower() == ".csv":
            df = pd.read_csv(p)
        else:
            df = pd.read_parquet(p)
        df.columns = df.columns.str.strip().str.lower()
        return df, None
    except Exception as e:
        return None, str(e)


@st.cache_data(show_spinner="Processing data...", max_entries=2, ttl=1800)
def process_data(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    df = df.copy()
    item_col = cols["item_col"]
    qsub_col = cols["qsub_col"]

    df["_item_norm"] = normalize_item_series(df[item_col])

    df["_qsub_raw"] = pd.to_numeric(df[qsub_col], errors="coerce").fillna(0).astype("float32")

    hv_key_norm = normalize_item(HV_POLE_KEY)
    hv_mask = df["_item_norm"] == hv_key_norm
    df["_qsub_adj"] = df["_qsub_raw"]
    df.loc[hv_mask, "_qsub_adj"] *= HV_POLE_MULTIPLIER

    item_to_cat = {}
    for cat_name, mapping in ALL_CATEGORIES.items():
        for desc, label in mapping.items():
            item_to_cat[normalize_item(desc)] = label
    df["_mapped_category"] = df["_item_norm"].map(item_to_cat)

    pole_col = cols.get("pole_col")
    if pole_col and pole_col in df.columns:
        df["_pole_norm"] = normalize_pole_series(df[pole_col])
    else:
        df["_pole_norm"] = ""

    job_col = cols.get("job_col")
    if job_col and job_col in df.columns:
        df["_job_clean"] = clean_job_series(df[job_col])
    else:
        df["_job_clean"] = ""

    date_col = cols.get("date_col")
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce") if date_col and date_col in df.columns else pd.NaT

    df["_item_norm"] = df["_item_norm"].astype("category")
    df["_mapped_category"] = df["_mapped_category"].astype("category")

    return df


@st.cache_data(show_spinner=False, max_entries=3, ttl=1800)
def cv7_dedupe_poles(frame: pd.DataFrame) -> set:
    keys = set()
    for mapping in POLE_CATEGORIES.values():
        keys |= {normalize_item(k) for k in mapping}
    poles = set(frame.loc[frame["_item_norm"].isin(keys), "_pole_norm"].dropna())
    poles.discard("")
    return poles


def cv_pole_resume(frame: pd.DataFrame, mapping: dict, cv7_poles: set) -> pd.DataFrame:
    keys = {normalize_item(k) for k in mapping}
    sub = frame[frame["_item_norm"].isin(keys)].copy()
    sub = sub[sub["_qsub_raw"] != 0]
    sub = sub.drop_duplicates(subset="_pole_norm")
    sub = sub[~sub["_pole_norm"].isin(cv7_poles)]
    return sub


def build_card(frame: pd.DataFrame, cat_name: str, mapping: dict, cv7_poles: set):
    if cat_name in POLE_DEDUPE_CATEGORIES:
        sub = cv_pole_resume(frame, mapping, cv7_poles)
        if sub.empty:
            return None
        return (cat_name, len(sub), sub)
    else:
        keys = {normalize_item(k) for k in mapping}
        sub = frame[frame["_item_norm"].isin(keys)]
        if sub.empty:
            return None
        return (cat_name, sub["_qsub_adj"].sum(), sub)


def build_subtype_card(frame: pd.DataFrame, subtype_name: str, descriptions: list):
    keys = {normalize_item(d) for d in descriptions}
    sub = frame[frame["_item_norm"].isin(keys)]
    if sub.empty:
        return None
    return (subtype_name, sub["_qsub_adj"].sum(), sub)


def render_metric(slot, cat_name: str, total_qty, display_name: str = None):
    label = display_name or cat_name
    if cat_name in UNIT_CONFIG:
        slot.metric(label, format_length(total_qty, UNIT_CONFIG[cat_name]))
    elif cat_name in POLE_DEDUPE_CATEGORIES:
        slot.metric(label, f"{total_qty:,.0f} poles")
    else:
        slot.metric(label, f"{total_qty:,.0f}")


# ============================================================
# APP
# ============================================================
st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        background-color: #f5f7fa;
        border: 1px solid #e3e7ee;
        border-radius: 10px;
        padding: 14px 16px;
    }
    div[data-testid="stMetricLabel"] { font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("⚡ Planning Dashboard")

with st.sidebar:
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("🔄 Refresh from network", help="Re-scan for the latest Master file and reload the outages/forecast/calendar files."):
            st.cache_data.clear()
            (st.rerun if hasattr(st, "rerun") else st.experimental_rerun)()
    with col_b:
        if st.button("🧹 Clear cache", help="Frees memory by dropping every cached file/dataframe."):
            st.cache_data.clear()
            st.success("Cache cleared.")
            (st.rerun if hasattr(st, "rerun") else st.experimental_rerun)()
    if not CALENDAR_LIB_AVAILABLE:
        st.caption("ℹ️ 'streamlit-calendar' package not installed - Outages Programme calendar view is unavailable (table view still works).")

# ---- Master parquet: auto-pick latest file from the network folder,
# with a manual override in case the auto-pick guesses wrong ----
all_masters = find_all_masters(MASTER_DIR)

if not all_masters:
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
raw_df, read_err = load_master_from_network(str(latest_master_path))
if read_err:
    st.error(
        f"Couldn't read `{latest_master_path.name}` - it may be corrupted or locked while being written.\n\n"
        f"**Details:** {read_err}"
    )
    st.stop()

with st.expander("Detected columns in your file (click to view)"):
    st.write(list(raw_df.columns))


def guess(*candidates):
    for c in candidates:
        if c in raw_df.columns:
            return c
    return None


with st.sidebar.expander("⚙️ Column mapping (advanced)", expanded=False):
    st.caption("Confirm each field maps to the right column in your file.")
    col_options = list(raw_df.columns)
    none_option = ["(none)"] + col_options

    def pick(label, default_col, key):
        opts = none_option
        idx = opts.index(default_col) if default_col in opts else 0
        if default_col is None:
            st.warning(f"Couldn't guess a column for **{label}** - pick one.")
        val = st.selectbox(label, opts, index=idx, key=key)
        return None if val == "(none)" else val

    cols = {
        "item_col": pick("Description / item", guess("item", "description"), "map_item"),
        "qsub_col": pick("Quantity (qsub)", guess("qsub", "quantity_used"), "map_qsub"),
        "district_col": pick("District", guess("shire", "district"), "map_district"),
        "project_col": pick("Project", guess("project"), "map_project"),
        "circuit_col": pick("Circuit", guess("segmentcode", "circuit"), "map_circuit"),
        "pole_col": pick("Pole / enid", guess("pole", "enid"), "map_pole"),
        "pid_col": pick("PID", guess("pid_ohl_nr", "pid"), "map_pid"),
        "total_col": pick("Total value", guess("total"), "map_total"),
        "orig_col": pick("Original value", guess("orig", "original"), "map_orig"),
        "job_col": pick("Job", guess("job", "sourcefile"), "map_job"),
        "date_col": pick("Date", guess("datetouse", "date", "plan1", "done"), "map_date"),
    }

missing_required = [k for k in ["item_col", "qsub_col", "district_col", "circuit_col"] if cols[k] is None]
if missing_required:
    st.error(f"These required fields still need a column picked in the sidebar 'Column mapping' section: {missing_required}")
    st.stop()

if cols.get("pole_col") is None:
    st.sidebar.warning("No Pole/enid column selected - CV8 and CV31 counts can't be deduplicated by pole and will be shown as raw row counts, which may not match the export tool.")

df = process_data(raw_df, cols)

district_col, project_col, circuit_col = cols["district_col"], cols["project_col"], cols["circuit_col"]
pole_col, pid_col, job_col = cols.get("pole_col"), cols.get("pid_col"), cols.get("job_col")

# ---- Sidebar filters ----
st.sidebar.header("🔍 Filters")


def pick_date_field_and_range(frame: pd.DataFrame, field_options: dict, fallback_series: pd.Series, default_label: str = None, key_prefix: str = ""):
    """Lets the user choose WHICH date column to filter by (Plan1 / Done /
    DateToUse, when present), then pick a range against it - either via a
    quick preset (fewer clicks for the common cases) or a custom calendar
    range. `field_options` is {display label: column name}, filtered here
    to only the ones that actually exist with real data; if none of them
    exist, falls back to `fallback_series` (the originally-mapped date
    column from the advanced Column mapping section) with no field choice
    shown. Presets are anchored to the LATEST DATE ACTUALLY IN THE DATA,
    not today's real-world date - this is planning/completion data that
    may not extend to today, so "Last 30 days" means the last 30 days of
    data present, which is what's actually useful here.
    Returns (dates_series, date_from, date_to)."""
    available = {}
    for label, col in field_options.items():
        if col in frame.columns:
            parsed = pd.to_datetime(frame[col], errors="coerce")
            # plan1/done/datetouse sometimes carry a placeholder ~1900 date
            # for "no date" - treat as missing, same cleanup as datetouse
            # already needed before any of these could be picked as the
            # active date field.
            parsed = parsed.where(parsed.dt.year > 1901)
            if parsed.notna().any():
                available[label] = parsed

    if not available:
        if fallback_series is None or fallback_series.isna().all():
            return fallback_series, None, None
        col_min, col_max = fallback_series.min(), fallback_series.max()
        date_range = st.sidebar.date_input("Date", value=(col_min.date(), col_max.date()), key=f"{key_prefix}date_fallback")
        if isinstance(date_range, tuple) and len(date_range) == 2:
            return fallback_series, date_range[0], date_range[1]
        return fallback_series, col_min.date(), col_max.date()

    labels = list(available.keys())
    default_index = labels.index(default_label) if default_label in labels else 0
    chosen_label = st.sidebar.radio("Date field", labels, index=default_index, horizontal=True, key=f"{key_prefix}date_field")
    dates = available[chosen_label]

    col_min = dates.min().normalize()
    col_max = dates.max().normalize()

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

    date_from = max(date_from, col_min.date())
    date_to = min(date_to, col_max.date())
    if date_from > date_to:
        date_from, date_to = col_min.date(), col_max.date()

    st.sidebar.caption(f"Showing: {date_from} -> {date_to}  (anchored to the latest date in the data, not today)")
    return dates, date_from, date_to


_dates_active, _date_from_picked, _date_to_picked = pick_date_field_and_range(
    df, {"Plan1": "plan1", "Done": "done", "DateToUse": "datetouse"},
    fallback_series=df["_date"], default_label="DateToUse",
)
df["_date"] = _dates_active

if _date_from_picked is not None and _date_to_picked is not None:
    date_range = (_date_from_picked, _date_to_picked)
else:
    date_range = None
    st.sidebar.caption("No usable dates found in the selected Date column.")


def multiselect_filter(label, col, key, options=None):
    if not col or col not in df.columns:
        return []
    opts = options if options is not None else sorted(df[col].dropna().astype(str).unique())
    return st.sidebar.multiselect(label, opts, key=key)


districts = multiselect_filter("District", district_col, "f_district")
projects = multiselect_filter("Project", project_col, "f_project")
pids = multiselect_filter("PID", pid_col, "f_pid")
sourcefiles = multiselect_filter("Source file", job_col, "f_sourcefile")

# Circuit and Pole (enid) are CASCADING, matching the Materials dashboard:
# Circuit's options narrow to what's actually reachable given
# District/Project/Date/PID/Source file so far, and Pole's options narrow
# further given Circuit too - so neither ever offers a value that
# couldn't possibly match what's already selected.
_scope_mask = pd.Series(True, index=df.index)
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    _start, _end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    _scope_mask &= df["_date"].between(_start, _end)
if districts:
    _scope_mask &= df[district_col].astype(str).isin(districts)
if projects:
    _scope_mask &= df[project_col].astype(str).isin(projects)
if pids and pid_col:
    _scope_mask &= df[pid_col].astype(str).isin(pids)
if sourcefiles and job_col:
    _scope_mask &= df[job_col].astype(str).isin(sourcefiles)

circuit_options = sorted(df.loc[_scope_mask, circuit_col].dropna().astype(str).unique()) if circuit_col else []
circuits = st.sidebar.multiselect(
    "Circuit", circuit_options, key="f_circuit",
    help="Narrowed to circuits that actually exist given the other filters selected above.",
)

_scope_mask_for_pole = _scope_mask.copy()
if circuits:
    _scope_mask_for_pole &= df[circuit_col].astype(str).isin(circuits)
pole_options = sorted(df.loc[_scope_mask_for_pole, pole_col].dropna().astype(str).unique()) if pole_col else []
poles_selected = multiselect_filter("Pole (enid)", pole_col, "f_pole", options=pole_options)

f = df.copy()
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    f = f[f["_date"].between(start, end)]
if districts:
    f = f[f[district_col].astype(str).isin(districts)]
if projects:
    f = f[f[project_col].astype(str).isin(projects)]
if circuits:
    f = f[f[circuit_col].astype(str).isin(circuits)]
if pids and pid_col:
    f = f[f[pid_col].astype(str).isin(pids)]
if sourcefiles and job_col:
    f = f[f[job_col].astype(str).isin(sourcefiles)]
if poles_selected and pole_col:
    f = f[f[pole_col].astype(str).isin(poles_selected)]

st.sidebar.divider()
st.sidebar.metric("Rows after filters", f"{len(f):,}", delta=f"of {len(df):,} total")

tab_overview, tab_jobs, tab_items, tab_forecast, tab_totals = st.tabs(
    ["📊 Overview", "🗂️ Jobs", "📦 Mapped Items", "📈 Pole Position", "💰 Totals"]
)

# ---- Overview: CV7_recover over time ----
with tab_overview:
    st.subheader("CV7_recover — count over time")
    recover_keys = {normalize_item(k) for k in CV7_recover}
    recover_df = f[f["_item_norm"].isin(recover_keys)]
    recover_total = recover_df["_qsub_adj"].sum()
    show_total_banner("Total CV7_recover count", f"{recover_total:,.0f}")

    if recover_df.empty or recover_df["_date"].isna().all():
        st.caption("No CV7_recover records (with a date) for the current filters.")
    else:
        granularity = st.radio("Group by", ["Day", "Week", "Month"], index=2, horizontal=True)
        freq = {"Day": "D", "Week": "W", "Month": "MS"}[granularity]
        trend = (
            recover_df.dropna(subset=["_date"])
            .set_index("_date")
            .resample(freq)["_qsub_adj"]
            .sum()
            .reset_index()
            .rename(columns={"_date": "Date", "_qsub_adj": "Count"})
        )
        fig = px.bar(trend, x="Date", y="Count", text="Count")
        fig.update_traces(marker_color="#2563eb")
        fig.update_layout(margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("All pole categories (erect only)")
    pole_rows = []
    for cat_name, mapping in POLE_CATEGORIES.items():
        if cat_name == "CV7_recover":
            continue
        keys = {normalize_item(k) for k in mapping}
        sub = f[f["_item_norm"].isin(keys)]
        if not sub.empty:
            pole_rows.append({
                "Pole type": POLE_DISPLAY_NAMES.get(cat_name, cat_name),
                "Count": sub["_qsub_adj"].sum(),
            })
    pole_summary = pd.DataFrame(pole_rows)

    pole_total = pole_summary["Count"].sum() if not pole_summary.empty else 0
    show_total_banner("Total poles (all categories)", f"{pole_total:,.0f}")

    if not pole_summary.empty:
        fig2 = px.bar(pole_summary, x="Pole type", y="Count", text="Count", color="Pole type")
        fig2.update_layout(showlegend=False, margin=dict(t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.caption("No pole records for the current filters.")

# ---- Jobs tab ----
with tab_jobs:
    st.subheader("Outages Programme 2026")

    outage_df = None
    if os.path.exists(OUTAGES_PATH):
        try:
            with open(OUTAGES_PATH, "rb") as fh:
                outage_bytes = fh.read()
            outage_df, outage_err = load_outage_programme(outage_bytes)
            if outage_err:
                st.error(
                    f"Couldn't read the outages workbook - check it has a sheet named '2026' with headers on row 7.\n\n"
                    f"**Details:** {outage_err}"
                )
                outage_df = None
        except OSError as e:
            st.error(f"Couldn't open the outages workbook.\n\n**Details:** {e}")
    else:
        st.error(f"Outages workbook not found at:\n\n`{OUTAGES_PATH}`")

    if outage_df is not None:
        st.caption(f"{len(outage_df):,} rows from High-level_planning_2026.xlsx (sheet '2026')")

        oc1, oc2, oc3 = st.columns(3)
        with oc1:
            outage_districts = st.multiselect(
                "District", sorted(outage_df["District"].dropna().unique()), key="outage_district"
            )
        with oc2:
            outage_pms = st.multiselect(
                "SPEN PM", sorted(outage_df["SPEN PM"].dropna().unique()), key="outage_pm"
            )
        with oc3:
            if outage_df["Outage Date"].notna().any():
                min_od, max_od = outage_df["Outage Date"].min(), outage_df["Outage Date"].max()
                outage_date_range = st.date_input(
                    "Outage date", value=(min_od.date(), max_od.date()), key="outage_date_range"
                )
            else:
                outage_date_range = None

        outage_f = outage_df.copy()
        if outage_districts:
            outage_f = outage_f[outage_f["District"].isin(outage_districts)]
        if outage_pms:
            outage_f = outage_f[outage_f["SPEN PM"].isin(outage_pms)]
        if outage_date_range and isinstance(outage_date_range, tuple) and len(outage_date_range) == 2:
            o_start = pd.Timestamp(outage_date_range[0])
            o_end = pd.Timestamp(outage_date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            outage_f = outage_f[outage_f["Outage Date"].between(o_start, o_end)]

        st.caption(f"{len(outage_f):,} rows after outage filters")

        view_options = ["Table", "Calendar"] if CALENDAR_LIB_AVAILABLE else ["Table"]
        outage_view = st.radio("View", view_options, index=0, horizontal=True, key="outage_view")

        if outage_view == "Table" or not CALENDAR_LIB_AVAILABLE:
            st.dataframe(outage_f, height=420, use_container_width=True, hide_index=True)
        else:
            cal_df = outage_f.dropna(subset=["Outage Date"])
            events = build_calendar_events(cal_df)

            calendar_options = {
                "initialView": "dayGridMonth",
                "headerToolbar": {
                    "left": "prev,next today",
                    "center": "title",
                    "right": "dayGridMonth,listMonth",
                },
                "height": "auto",
                "firstDay": 1,
                "dayMaxEvents": False,
            }

            calendar_custom_css = """
                .fc-event-title { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
                .fc-daygrid-event { white-space: nowrap; cursor: pointer; }
                .fc-daygrid-day-number { cursor: pointer; }
                .fc-daygrid-day:hover { background-color: #eef2ff; }
            """

            calendar_result = st_calendar(
                events=events,
                options=calendar_options,
                custom_css=calendar_custom_css,
                callbacks=["dateClick", "eventClick"],
                key="outage_calendar",
            )

            clicked_date = None
            if calendar_result.get("eventClick"):
                ev_start = calendar_result["eventClick"].get("event", {}).get("start")
                if ev_start:
                    clicked_date = pd.to_datetime(ev_start).date()
            elif calendar_result.get("dateClick"):
                dc = calendar_result["dateClick"]
                dc_date = dc.get("date") or dc.get("dateStr")
                if dc_date:
                    clicked_date = pd.to_datetime(dc_date).date()

            if clicked_date is not None:
                st.session_state["selected_outage_date"] = clicked_date

            selected_date = st.session_state.get("selected_outage_date")

            st.divider()
            if selected_date is None:
                st.caption("Click a date or an outage above to see its full breakdown and Work Instructions here.")
            else:
                day_rows = outage_f[outage_f["Outage Date"].dt.date == selected_date]
                st.markdown(f"**Outages on {selected_date:%d %b %Y}** ({len(day_rows):,})")
                if day_rows.empty:
                    st.caption("No outages on this date under the current filters.")
                else:
                    st.dataframe(day_rows, use_container_width=True, hide_index=True)

                st.markdown(f"**Work Instructions — poles & MD Poling tasks on {selected_date:%d %b %Y}**")
                if cols.get("date_col") is None or f["_date"].isna().all():
                    st.caption(
                        "No usable dates in the main dataset's mapped 'Date' column - pick one under "
                        "'Column mapping' in the sidebar to enable this breakdown."
                    )
                else:
                    day_job_rows = f[f["_date"].dt.date == selected_date]
                    pole_task_df = build_pole_task_table(
                        day_job_rows, cols["item_col"], cols["qsub_col"], cols.get("pole_col")
                    )
                    if pole_task_df.empty:
                        st.caption("No pole/task records for this date in the main dataset under the current filters.")
                    else:
                        st.caption(f"{pole_task_df['Pole'].nunique():,} pole(s) · {len(pole_task_df):,} task row(s)")
                        st.markdown(render_pole_task_table_html(pole_task_df), unsafe_allow_html=True)

    # ---- Outage Calendar (.ics) - auto-picks the newest version file ----
    st.divider()
    st.subheader("Outage Calendar (.ics)")

    latest_ics_path = find_latest_ics(ICS_DIR)
    if latest_ics_path is None:
        st.info(
            f"No `outage_calendar_YYYY (vN).ics` file found in:\n\n`{ICS_DIR}`\n\n"
            "Deploy one there (any version number, any year) and it'll be picked up automatically."
        )
    else:
        st.caption(f"📅 Using: **{latest_ics_path.name}**")
        try:
            ics_bytes = latest_ics_path.read_bytes()
            ics_df, ics_err = parse_ics_events(ics_bytes)
            if ics_err:
                st.error(f"Couldn't parse `{latest_ics_path.name}`.\n\n**Details:** {ics_err}")
            elif ics_df.empty:
                st.caption("No VEVENT entries found in this calendar file.")
            else:
                st.caption(f"{len(ics_df):,} event(s) found")
                st.dataframe(
                    ics_df[["Summary", "Start", "End", "Location", "Description"]],
                    height=360, use_container_width=True, hide_index=True,
                )
        except OSError as e:
            st.error(f"Couldn't read `{latest_ics_path.name}`.\n\n**Details:** {e}")

# ---- Mapped items tab: image-led groups, then the rest as a card grid ----
with tab_items:
    st.subheader("Mapped items")

    cv7_poles = cv7_dedupe_poles(f)
    all_card_data = []

    grouped_cat_names = {c for group in CARD_GROUPS for c in group.get("categories", [])}
    grouped_cat_names |= {c for group in CARD_GROUPS if "subtypes" in group for c in ["Switch"]}

    for group in CARD_GROUPS:
        title = group["title"]
        img_l, img_c, img_r = st.columns([1, 1, 1])
        with img_c:
            st.markdown(f"<h3 style='text-align:center; margin-bottom:0.3rem;'>{title}</h3>", unsafe_allow_html=True)
            image_path = group.get("image")
            if image_path and os.path.exists(image_path):
                st.image(image_path, width=300)
            elif image_path:
                st.caption(f"⚠️ Image not found: {image_path}")

        group_cards = []

        if "subtypes" in group:
            for subtype_name, descriptions in group["subtypes"].items():
                card = build_subtype_card(f, subtype_name, descriptions)
                if card:
                    group_cards.append(card)
                    all_card_data.append(card)
        else:
            for cat_name in group.get("categories", []):
                mapping = ALL_CATEGORIES.get(cat_name)
                if mapping is None:
                    continue
                card = build_card(f, cat_name, mapping, cv7_poles)
                if card:
                    group_cards.append(card)
                    all_card_data.append(card)

        if group_cards:
            row_cols = st.columns(len(group_cards))
            for slot, (cat_name, total_qty, _sub) in zip(row_cols, group_cards):
                render_metric(slot, cat_name, total_qty, display_name=POLE_DISPLAY_NAMES.get(cat_name))
        else:
            st.caption("No records for this group under the current filters.")

        st.divider()

    remaining_cat_names = [c for c in ALL_CATEGORIES if c not in grouped_cat_names]
    remaining_cards = []
    for cat_name in remaining_cat_names:
        card = build_card(f, cat_name, ALL_CATEGORIES[cat_name], cv7_poles)
        if card:
            remaining_cards.append(card)
            all_card_data.append(card)

    if remaining_cards:
        st.markdown("**Other items**")
        n_cols = 4
        rows = [remaining_cards[i:i + n_cols] for i in range(0, len(remaining_cards), n_cols)]
        for row in rows:
            row_cols = st.columns(n_cols)
            for slot, (cat_name, total_qty, _sub) in zip(row_cols, row):
                render_metric(slot, cat_name, total_qty)

    if not all_card_data:
        st.caption("No mapped items for the current filters.")
    else:
        st.divider()
        chosen = st.selectbox("View details for", [c[0] for c in all_card_data])
        _, _, sub = next(c for c in all_card_data if c[0] == chosen)
        detail = pd.DataFrame({
            "District": sub[district_col],
            "Job": sub["_job_clean"],
            "Circuit": sub[circuit_col],
            "enid": sub[cols["pole_col"]] if cols["pole_col"] in sub.columns else "",
        })
        st.dataframe(detail, height=320, use_container_width=True, hide_index=True)

# ---- Poles Forecast tab ----
@st.cache_data(show_spinner="Reading poles forecast workbook...", max_entries=2, ttl=1800)
def list_forecast_sheets(file_bytes: bytes):
    try:
        return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names, None
    except Exception as e:
        return [], str(e)


@st.cache_data(show_spinner="Reading poles forecast workbook...", max_entries=2, ttl=1800)
def load_forecast_workbook(file_bytes: bytes, sheet_name: str):
    try:
        fdf = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
        fdf.columns = fdf.columns.astype(str).str.strip()
        return fdf, None
    except Exception as e:
        return None, str(e)


# ---- Pole Position: status colours, column letters, hyperlinks, Gantt rows ----
# Status colours from the workbook's Status column (matched ignoring case).
FC_STATUS_COLOURS = [
    ("complete", "#00B050", "Complete"),
    ("planned", "#8ED973", "Planned"),
    ("awaiting outage plan", "#C00000", "Awaiting Outage Plan"),
    ("land access", "#C00000", "Land Access Issues"),
    ("still to be handed over", "#000000", "Still to be handed over"),
    ("to be priced", "#0070C0", "To be priced"),
]
FC_STATUS_OTHER = "#9AA3AF"

# (key, label, header guesses, fallback column letter)
FC_FIELDS = [
    ("district", "District", ("District",), None),
    ("pid", "Project ID", ("Project ID", "PID"), None),
    ("project", "Project (column C)", ("Project", "Project Name"), "C"),
    ("circuit", "Circuit", ("Circuit",), None),
    ("voltage", "Voltage", ("Voltage",), None),
    ("forecast", "Forecasted Total poles", ("Forecasted Total poles", "Forecasted Total Poles", "Forecast Total Poles"), None),
    ("disposed", "Poles Disposed", ("Poles Disposed", "Poles disposed"), None),
    ("status", "Status", ("Status", "Job Status", "Project Status", "Stage"), None),
    ("start_date", "Start Date (column I)", ("Start Date", "StartDate", "Start"), "I"),
    ("finish_date", "Finish Date (column J)", ("Finish Date", "End Date", "Finish", "Completion Date", "FinishDate"), "J"),
    ("comment", "Comment (column L)", ("Comment", "Comments", "Notes"), "L"),
    ("control_file", "Control File (column O)", ("Control File", "Control file", "Control Files", "CF"), "O"),
]
FC_REQUIRED = ["project", "circuit", "pid", "forecast", "disposed"]
_HYPERLINK_FORMULA = re.compile(r'^=\s*HYPERLINK\(\s*"([^"]*)"\s*(?:[,;]\s*"([^"]*)")?', re.IGNORECASE)


def forecast_status_colour(status):
    s = "" if status is None or (isinstance(status, float) and pd.isna(status)) else str(status).strip().lower()
    for key, colour, _label in FC_STATUS_COLOURS:
        if key in s:
            return colour
    return FC_STATUS_OTHER


def forecast_guess_columns(fdf):
    """Header-name guess first, then the column letter the workbook uses (C, I, J, L, O)."""
    lower_map = {c.lower(): c for c in fdf.columns}
    cols = list(fdf.columns)
    out = {}
    for key, _label, guesses, letter in FC_FIELDS:
        found = next((lower_map[g.lower()] for g in guesses if g.lower() in lower_map), None)
        if found is None and letter:
            i = ord(letter) - ord("A")
            found = cols[i] if i < len(cols) else None
        out[key] = found
    return out


def read_forecast_links(file_bytes, sheet_name, fdf, col_names):
    """Hyperlinks (cell links or =HYPERLINK formulas) for the given columns.
    Returns {column: {row_index: (url, friendly_text)}} - header is row 1,
    so dataframe row i is worksheet row i + 2."""
    import openpyxl
    out = {c: {} for c in col_names if c}
    if not out:
        return out
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb[sheet_name]
    except Exception:
        return out
    positions = {c: list(fdf.columns).index(c) + 1 for c in out}
    for c, col_i in positions.items():
        for i in range(len(fdf)):
            cell = ws.cell(row=i + 2, column=col_i)
            url, text = None, None
            if cell.hyperlink is not None and cell.hyperlink.target:
                url = cell.hyperlink.target
            elif isinstance(cell.value, str):
                m = _HYPERLINK_FORMULA.match(cell.value.strip())
                if m:
                    url, text = m.group(1), m.group(2)
            if url:
                out[c][i] = (url, text)
    return out


def build_forecast_rows(fdf, f_cols, links):
    """One row per job with everything the Gantt and the job panel need,
    ordered District -> Voltage -> Project -> Circuit -> PID."""
    def col(key):
        return fdf[f_cols[key]] if f_cols.get(key) else pd.Series([None] * len(fdf), index=fdf.index)

    def txt(v):
        return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()

    proj_links = links.get(f_cols.get("project"), {}) if f_cols.get("project") else {}
    cf_links = links.get(f_cols.get("control_file"), {}) if f_cols.get("control_file") else {}
    project = col("project").copy()
    for i, (_url, text) in proj_links.items():      # formula links: take the friendly name if the cell has no value
        if text and i < len(project) and txt(project.iloc[i]) == "":
            project.iloc[i] = text
    df = pd.DataFrame({
        "District": col("district").map(txt),
        "Voltage": col("voltage").map(txt),
        "Project": project.map(txt),
        "Circuit": col("circuit").map(txt),
        "PID": col("pid").map(txt),
        "Status": col("status").map(txt),
        "Forecast": pd.to_numeric(col("forecast"), errors="coerce").fillna(0),
        "Disposed": pd.to_numeric(col("disposed"), errors="coerce").fillna(0),
        "Start Date": pd.to_datetime(col("start_date"), errors="coerce"),
        "Finish Date": pd.to_datetime(col("finish_date"), errors="coerce"),
        "Comment": col("comment").map(txt),
        "Control File": col("control_file").map(txt),
    })
    df["_row"] = range(len(df))
    df["Project link"] = [proj_links.get(i, (None, None))[0] for i in df["_row"]]
    df["Control File link"] = [cf_links.get(i, (None, None))[0] for i in df["_row"]]
    df.loc[df["Control File"].eq("") & df["Control File link"].notna(), "Control File"] = df["Control File link"]
    df = df[df["Project"] != ""]
    df["Disposed"] = df[["Disposed", "Forecast"]].min(axis=1)
    df["Remaining"] = (df["Forecast"] - df["Disposed"]).clip(lower=0)
    df["Year"] = df["Start Date"].dt.year
    df["Colour"] = df["Status"].map(forecast_status_colour)
    for k in ("District", "Voltage", "Project", "Circuit", "PID"):
        df[k + "_sort"] = df[k].str.lower().replace("", "~")
    df = df.sort_values(["District_sort", "Voltage_sort", "Project_sort", "Circuit_sort", "PID_sort", "Start Date"], kind="stable")
    return df.drop(columns=[c for c in df.columns if c.endswith("_sort")])


@st.cache_data(show_spinner="Reading links from the workbook...", max_entries=4, ttl=1800)
def cached_forecast_links(file_bytes: bytes, sheet_name: str, project_col, control_col):
    fdf, _err = load_forecast_workbook(file_bytes, sheet_name)
    return read_forecast_links(file_bytes, sheet_name, fdf, [project_col, control_col])


with tab_forecast:
    st.subheader("Pole Position")

    forecast_bytes = None
    if os.path.exists(FORECAST_PATH):
        try:
            with open(FORECAST_PATH, "rb") as fh:
                forecast_bytes = fh.read()
        except OSError as e:
            st.error(f"Couldn't open the forecast workbook.\n\n**Details:** {e}")
    else:
        st.error(
            f"Forecast workbook not found at:\n\n`{FORECAST_PATH}`\n\n"
            "Update `FORECAST_PATH` near the top of this script with the real path."
        )

    if forecast_bytes is None:
        pass
    else:
        sheet_names, sheet_err = list_forecast_sheets(forecast_bytes)
        if sheet_err:
            st.error(f"Couldn't open that workbook.\n\n**Details:** {sheet_err}")
            sheet_names = []

        if not sheet_names:
            pass
        else:
            sheet_choice = (
                st.selectbox("Sheet", sheet_names, key="forecast_sheet")
                if len(sheet_names) > 1 else sheet_names[0]
            )
            fdf, fdf_err = load_forecast_workbook(forecast_bytes, sheet_choice)
            if fdf_err:
                st.error(f"Couldn't read sheet '{sheet_choice}'.\n\n**Details:** {fdf_err}")
                fdf = None

            if fdf is None:
                pass
            else:
                with st.expander("Detected columns (click to view)"):
                    st.write(list(fdf.columns))

                fcol_options = ["(none)"] + list(fdf.columns)
                _fc_guess = forecast_guess_columns(fdf)

                with st.expander("⚙️ Column mapping", expanded=False):
                    def fpick(label, default_col, key):
                        idx = fcol_options.index(default_col) if default_col in fcol_options else 0
                        if default_col is None:
                            st.warning(f"Couldn't guess a column for **{label}** - pick one.")
                        val = st.selectbox(label, fcol_options, index=idx, key=key)
                        return None if val == "(none)" else val

                    f_cols = {key: fpick(label, _fc_guess[key], f"fc_{key}") for key, label, _g, _l in FC_FIELDS}

                missing = [k for k in FC_REQUIRED if f_cols[k] is None]
                if missing:
                    st.error(f"Please map these columns in 'Column mapping' above: {missing}")
                else:
                    fc_links = cached_forecast_links(forecast_bytes, sheet_choice, f_cols["project"], f_cols["control_file"])
                    jobs = build_forecast_rows(fdf, f_cols, fc_links)

                    fc1, fc2, fc3, fc4 = st.columns(4)
                    with fc1:
                        forecast_districts = (
                            st.multiselect("District", sorted({v for v in jobs["District"] if v}), key="forecast_district")
                            if f_cols["district"] else []
                        )
                    with fc2:
                        forecast_voltages = (
                            st.multiselect("Voltage", sorted({v for v in jobs["Voltage"] if v}), key="forecast_voltage")
                            if f_cols["voltage"] else []
                        )
                    with fc3:
                        forecast_statuses = (
                            st.multiselect("Status", sorted({v for v in jobs["Status"] if v}), key="forecast_status")
                            if f_cols["status"] else []
                        )
                    with fc4:
                        forecast_years = st.multiselect(
                            "Start year", sorted(jobs["Year"].dropna().unique().astype(int)), key="forecast_year",
                        )

                    if forecast_districts:
                        jobs = jobs[jobs["District"].isin(forecast_districts)]
                    if forecast_voltages:
                        jobs = jobs[jobs["Voltage"].isin(forecast_voltages)]
                    if forecast_statuses:
                        jobs = jobs[jobs["Status"].isin(forecast_statuses)]
                    if forecast_years:
                        jobs = jobs[jobs["Year"].isin(forecast_years)]

                    total_forecast = jobs["Forecast"].sum()
                    total_disposed = jobs["Disposed"].sum()
                    show_total_banner(
                        "Poles disposed vs forecasted",
                        f"{total_disposed:,.0f} / {total_forecast:,.0f}"
                        + (f"  ({total_disposed / total_forecast:.0%})" if total_forecast else ""),
                    )
                    not_found = [label for key, label, _g, _l in FC_FIELDS
                                 if key in ("status", "start_date", "finish_date", "comment", "control_file") and not f_cols[key]]
                    if not_found:
                        st.caption("⚠️ Not found in this sheet: " + ", ".join(not_found) + " - pick them under Column mapping if they're there under another name.")

                    if jobs.empty:
                        st.caption("No rows to chart for the current filters.")
                    else:
                        # ---- legend: the workbook's status colours ----
                        present = jobs["Status"].map(lambda s: next((lab for k, _c, lab in FC_STATUS_COLOURS if k in s.lower()), "Other / blank")).value_counts()
                        chips = [(lab, c) for _k, c, lab in FC_STATUS_COLOURS] + [("Other / blank", FC_STATUS_OTHER)]
                        st.markdown(" ".join(
                            f"<span style='display:inline-flex;align-items:center;gap:6px;margin:0 14px 6px 0;font-size:13px;'>"
                            f"<span style='width:12px;height:12px;border-radius:3px;background:{c};display:inline-block;'></span>{lab} ({present[lab]})</span>"
                            for lab, c in chips if lab in present.index), unsafe_allow_html=True)

                        dated = jobs.dropna(subset=["Start Date"])
                        t0 = t1 = today = None
                        if dated.empty:
                            st.caption("No Start Dates in these rows - map the Start Date column under Column mapping.")
                        else:
                            ends = dated["Finish Date"].where(dated["Finish Date"] >= dated["Start Date"])
                            t0 = dated["Start Date"].min().replace(day=1)
                            t1 = (pd.concat([ends.dropna(), dated["Start Date"]]).max() + pd.offsets.MonthBegin(1)).normalize()
                            today = pd.Timestamp.today().normalize()

                        def _ink(hex_colour):
                            h_ = hex_colour.lstrip("#")
                            r_, g_, b_ = int(h_[0:2], 16), int(h_[2:4], 16), int(h_[4:6], 16)
                            return "#0b0b0b" if (0.2126 * r_ + 0.7152 * g_ + 0.0722 * b_) / 255 > 0.55 else "#ffffff"

                        # ---- Gantt: one chart per District / Voltage, rows = Project - Circuit - PID ----
                        picked_row = st.session_state.get("fc_pick")
                        for g_i, ((g_district, g_voltage), grp) in enumerate(jobs.groupby(["District", "Voltage"], sort=False)):
                            st.markdown(
                                f"<div style='background:#0E6E78;color:#fff;padding:7px 12px;border-radius:6px;font-weight:700;"
                                f"letter-spacing:.04em;text-transform:uppercase;display:flex;justify-content:space-between;margin-top:12px;'>"
                                f"<span>{g_district or '(blank)'} · {g_voltage or '(blank)'}</span>"
                                f"<span>{grp['Disposed'].sum():,.0f} / {grp['Forecast'].sum():,.0f} poles</span></div>",
                                unsafe_allow_html=True)
                            gd = grp.dropna(subset=["Start Date"])
                            no_date = grp[grp["Start Date"].isna()]
                            if not no_date.empty:
                                st.caption("No Start Date: " + ", ".join(f"{p_} — {c_} — PID {pid_}" for p_, c_, pid_ in zip(no_date["Project"], no_date["Circuit"], no_date["PID"])))
                            if gd.empty or t0 is None:
                                continue
                            finish = gd["Finish Date"].where(gd["Finish Date"] >= gd["Start Date"])
                            end = (finish + pd.Timedelta(days=1)).fillna(gd["Start Date"] + pd.Timedelta(days=21))
                            labels = (gd["Project"] + " — " + gd["Circuit"] + " — PID " + gd["PID"]).tolist()
                            fig = go.Figure(go.Bar(
                                y=labels, base=gd["Start Date"], x=(end - gd["Start Date"]).dt.total_seconds() * 1000,
                                orientation="h", marker_color=gd["Colour"], marker_line=dict(color="rgba(0,0,0,.25)", width=0.5),
                                text=[f"Disposed {d:,.0f} · Forecast {f_:,.0f}" for d, f_ in zip(gd["Disposed"], gd["Forecast"])],
                                textposition="auto", insidetextanchor="start",
                                textfont=dict(color=[_ink(c) for c in gd["Colour"]], size=12),
                                outsidetextfont=dict(color="#1e293b", size=12),
                                customdata=gd[["_row"]].values,
                                hovertemplate=[
                                    f"<b>{lab}</b><br>Status: {st_ or '—'}<br>{sd:%d %b %Y} → "
                                    + (f"{fd:%d %b %Y}" if pd.notna(fd) else "no finish date")
                                    + f"<br>Disposed {dp:,.0f} / Forecast {fo:,.0f}<extra>click for details</extra>"
                                    for lab, st_, sd, fd, dp, fo in zip(labels, gd["Status"], gd["Start Date"], gd["Finish Date"], gd["Disposed"], gd["Forecast"])],
                            ))
                            if t0 <= today <= t1:
                                fig.add_shape(type="line", x0=today, x1=today, y0=0, y1=1, yref="paper", line=dict(width=1.5, color="#475569"), opacity=0.6)
                            fig.update_layout(
                                height=max(140, 36 * len(gd) + 60), margin=dict(l=330, r=10, t=10, b=10), showlegend=False,
                                xaxis=dict(type="date", range=[t0, t1], gridcolor="#E2E7ED", side="top"),
                                yaxis=dict(categoryorder="array", categoryarray=labels[::-1], autorange=True, automargin=False,
                                           tickvals=labels, ticktext=[l if len(l) <= 52 else l[:51] + "…" for l in labels]),
                                bargap=0.3, plot_bgcolor="#FFFFFF",
                            )
                            try:
                                ev = st.plotly_chart(fig, use_container_width=True, on_select="rerun", selection_mode="points", key=f"fc_gantt_{g_i}")
                                try:
                                    pts = ev["selection"]["points"] if ev else []
                                except (KeyError, TypeError, AttributeError):
                                    pts = []
                                if pts:
                                    cd = pts[0].get("customdata")
                                    picked_row = int(cd[0] if isinstance(cd, (list, tuple)) else cd)
                                    st.session_state["fc_pick"] = picked_row
                            except TypeError:       # older Streamlit without chart selections
                                st.plotly_chart(fig, use_container_width=True)
                        st.caption(f"{len(jobs):,} project/circuit rows shown · bars run from Start Date to Finish Date (open-ended when there's no Finish Date)")

                        # ---- job details: comment (L), control file (O), project link (C) ----
                        st.markdown("#### Job details")
                        job_labels = {int(rw): f"{d_} · {v_} · {p_} — {c_} — PID {pid_}" for rw, d_, v_, p_, c_, pid_
                                      in zip(jobs["_row"], jobs["District"], jobs["Voltage"], jobs["Project"], jobs["Circuit"], jobs["PID"])}
                        keys = list(job_labels)
                        default_i = keys.index(picked_row) if picked_row in keys else 0
                        sel_row = st.selectbox("Click a bar above, or pick a job", keys, index=default_i, format_func=lambda k: job_labels[k], key=f"fc_job_{picked_row}")
                        j = jobs[jobs["_row"] == sel_row].iloc[0]
                        with st.container(border=True):
                            title = f"[{j['Project']} ↗]({j['Project link']})" if j["Project link"] else j["Project"]
                            st.markdown(f"### {title}")
                            st.markdown(
                                f"<span style='display:inline-flex;align-items:center;gap:7px;border:1px solid #DCE3EC;border-radius:99px;padding:3px 10px;font-weight:600;'>"
                                f"<span style='width:13px;height:13px;border-radius:4px;background:{j['Colour']};display:inline-block;'></span>{j['Status'] or 'No status'}</span>",
                                unsafe_allow_html=True)
                            d1, d2, d3, d4 = st.columns(4)
                            d1.metric("Forecasted Total poles", f"{j['Forecast']:,.0f}")
                            d2.metric("Poles Disposed", f"{j['Disposed']:,.0f}")
                            d3.metric("Start Date", j["Start Date"].strftime("%d %b %Y") if pd.notna(j["Start Date"]) else "—")
                            d4.metric("Finish Date", j["Finish Date"].strftime("%d %b %Y") if pd.notna(j["Finish Date"]) else "—")
                            st.markdown(f"**District:** {j['District'] or '—'} · **Voltage:** {j['Voltage'] or '—'} · **Circuit:** {j['Circuit'] or '—'} · **PID:** {j['PID'] or '—'} · workbook row {int(j['_row']) + 2}")
                            st.markdown("**Comment**")
                            st.write(j["Comment"] or "No comment in column L for this job.")
                            st.markdown("**Control File**")
                            cf_target = j["Control File link"] or j["Control File"]
                            if not cf_target:
                                st.write("No control file in column O.")
                            else:
                                if j["Control File link"]:
                                    st.markdown(f"[{j['Control File'] or cf_target}]({j['Control File link']})")
                                st.code(cf_target, language=None)   # copy button on the right
                            if j["Project link"]:
                                st.link_button("Open project link", j["Project link"])

                        with st.expander("All jobs as a table"):
                            st.dataframe(
                                jobs[["District", "Voltage", "Project", "Circuit", "PID", "Status", "Start Date", "Finish Date",
                                      "Forecast", "Disposed", "Remaining", "Comment", "Control File", "Project link"]],
                                use_container_width=True, hide_index=True,
                                column_config={"Project link": st.column_config.LinkColumn("Project link"),
                                               "Start Date": st.column_config.DateColumn(format="DD MMM YYYY"),
                                               "Finish Date": st.column_config.DateColumn(format="DD MMM YYYY")},
                            )


with tab_totals:
    GBP_COLUMN_CONFIG = lambda col_name: {col_name: st.column_config.NumberColumn(col_name, format="£%.2f")}

    total_val = pd.to_numeric(f[cols["total_col"]], errors="coerce").sum() if cols["total_col"] in f.columns else None
    orig_val = pd.to_numeric(f[cols["orig_col"]], errors="coerce").sum() if cols["orig_col"] in f.columns else None

    if total_val is not None:
        show_total_banner("Total value (£)", f"£{total_val:,.2f}")

    c1, c2 = st.columns(2)
    if total_val is not None:
        c1.metric("Total value", f"£{total_val:,.2f}")
    if total_val is not None and orig_val is not None:
        c2.metric("Difference vs original", f"£{total_val - orig_val:,.2f}")

    if total_val is not None and project_col in f.columns:
        st.divider()
        st.subheader("Total value by Project")
        f["_total_val_row"] = pd.to_numeric(f[cols["total_col"]], errors="coerce")
        by_project_total = (
            f.groupby(project_col)["_total_val_row"].sum()
            .reset_index()
            .rename(columns={project_col: "Project", "_total_val_row": "Total value (£)"})
            .sort_values("Total value (£)", ascending=False)
        )
        st.caption(f"{len(by_project_total):,} projects under the current filters")

        fig_proj = px.bar(
            by_project_total.sort_values("Total value (£)", ascending=True),
            x="Total value (£)", y="Project", orientation="h", text="Total value (£)",
        )
        fig_proj.update_traces(marker_color="#2563eb", texttemplate="£%{text:,.0f}")
        fig_proj.update_layout(height=max(360, 32 * len(by_project_total)), margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig_proj, use_container_width=True)

        st.dataframe(
            by_project_total, height=320, use_container_width=True, hide_index=True,
            column_config=GBP_COLUMN_CONFIG("Total value (£)"),
        )

    if total_val is not None and orig_val is not None:
        f["_row_variance"] = pd.to_numeric(f[cols["total_col"]], errors="coerce") - pd.to_numeric(f[cols["orig_col"]], errors="coerce")
        variance_rows = f[f["_row_variance"] != 0]

        st.subheader("Jobs where total ≠ original")
        variance_table = pd.DataFrame({
            "District": variance_rows[district_col],
            "Job": variance_rows["_job_clean"],
            "Circuit": variance_rows[circuit_col],
            "Difference (£)": variance_rows["_row_variance"],
        })
        st.caption(f"{len(variance_table):,} rows with a variance under the current filters")
        st.dataframe(
            variance_table.sort_values("Difference (£)", key=abs, ascending=False),
            height=320, use_container_width=True, hide_index=True,
            column_config=GBP_COLUMN_CONFIG("Difference (£)"),
        )

        st.subheader("Difference by Job")
        by_job = (
            variance_rows.assign(Job=variance_rows["_job_clean"])
            .groupby("Job")["_row_variance"].sum()
            .reset_index()
            .rename(columns={"_row_variance": "Difference (£)"})
            .sort_values("Difference (£)", key=abs, ascending=False)
        )
        st.dataframe(
            by_job, height=280, use_container_width=True, hide_index=True,
            column_config=GBP_COLUMN_CONFIG("Difference (£)"),
        )

        st.subheader("Difference by Project")
        if project_col in variance_rows.columns:
            by_project = (
                variance_rows.groupby(project_col)["_row_variance"].sum()
                .reset_index()
                .rename(columns={project_col: "Project", "_row_variance": "Difference (£)"})
                .sort_values("Difference (£)", key=abs, ascending=False)
            )
            st.dataframe(
                by_project, height=280, use_container_width=True, hide_index=True,
                column_config=GBP_COLUMN_CONFIG("Difference (£)"),
            )
