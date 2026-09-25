"""
dash_tracker.py - Network Job Tracker ("Planning Dashboard"), data side.

Every helper, mapping dictionary and calculation below is copied from
network_job_tracker.py unchanged (only the @st.cache_data decorators were
removed). The Streamlit screen code is replaced by the functions at the
bottom, which return plain data for the web page. Files come from what you
drop on the page instead of fixed network paths.
"""
import os
import re
import io
from pathlib import Path
from datetime import datetime, date, time as dt_time

import pandas as pd

IMAGE_DIR = ""   # group pictures are drawn by the page instead

MASTER_PATTERN = re.compile(r"Master_(\d{2})-(\d{2})-(\d{4})(?:\s*\(v(\d+)\))?\.parquet$", re.IGNORECASE)
# Matches: outage_calendar_2026 (v6).ics - picks highest year, then highest
# version, so a stray "2027" file would win over "2026" regardless of vN,
# and within the same year the highest vN wins. Deploy new files without
# renaming and this always finds the newest one.
ICS_PATTERN = re.compile(r"outage_calendar_(\d{4})\s*\(v(\d+)\)\.ics$", re.IGNORECASE)
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

def list_forecast_sheets(file_bytes: bytes):
    try:
        return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names, None
    except Exception as e:
        return [], str(e)


def load_forecast_workbook(file_bytes: bytes, sheet_name: str):
    try:
        fdf = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
        fdf.columns = fdf.columns.astype(str).str.strip()
        return fdf, None
    except Exception as e:
        return None, str(e)



# ============================================================
# WEB PAGE FUNCTIONS - replace the Streamlit screen code. The filter,
# overview, mapped-items, jobs, pole-position and totals calculations are
# the same statements as in the Streamlit app, returning data instead of
# drawing widgets.
# ============================================================
PRESETS = ["All time", "Last 7 days", "Last 30 days", "This month", "This year", "Last year", "Custom range"]
DATE_FIELD_OPTIONS = {"Plan1": "plan1", "Done": "done", "DateToUse": "datetouse"}
COL_FIELDS = [  # (key, label, guesses)
    ("item_col", "Description / item", ("item", "description")),
    ("qsub_col", "Quantity (qsub)", ("qsub", "quantity_used")),
    ("district_col", "District", ("shire", "district")),
    ("project_col", "Project", ("project",)),
    ("circuit_col", "Circuit", ("segmentcode", "circuit")),
    ("pole_col", "Pole / enid", ("pole", "enid")),
    ("pid_col", "PID", ("pid_ohl_nr", "pid")),
    ("total_col", "Total value", ("total",)),
    ("orig_col", "Original value", ("orig", "original")),
    ("job_col", "Job", ("job", "sourcefile")),
    ("date_col", "Date", ("datetouse", "date", "plan1", "done")),
]
REQUIRED = ["item_col", "qsub_col", "district_col", "circuit_col"]

_RAW = {}
_PROC = {}
_FILE_CACHE = {}


def _s(v):
    if v is None:
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return "" if pd.isna(v) else v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def _n(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(v) else round(v, 2)


def _cell(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return _n(v)
    return _s(v)


def _table(df):
    return {"columns": [str(c) for c in df.columns], "rows": [[_cell(v) for v in r] for r in df.itertuples(index=False)]}


def load_master_from_network(path_str: str):
    """Same as the Streamlit loader - the path is now the dropped file."""
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


def _raw(path):
    key = (path, os.path.getmtime(path))
    if key not in _RAW:
        _RAW.clear(); _PROC.clear()
        df, err = load_master_from_network(path)
        if err:
            raise ValueError(f"Couldn't read {os.path.basename(path)} - it may be corrupted.\n\nDetails: {err}")
        _RAW[key] = df
    return _RAW[key]


def _clean_cols(raw_df, cols):
    cols = dict(cols or {})
    out = {}
    for key, _label, guesses in COL_FIELDS:
        val = cols.get(key, "__guess__")
        if val == "__guess__":
            val = next((c for c in guesses if c in raw_df.columns), None)
        out[key] = val if val in raw_df.columns else None
    return out


def _proc(path, cols):
    raw_df = _raw(path)
    key = (path, tuple(sorted((k, v or "") for k, v in cols.items())))
    if key not in _PROC:
        _PROC.clear()
        _PROC[key] = process_data(raw_df, cols)
    return _PROC[key].copy()


def load(path, cols=None):
    raw_df = _raw(path)
    cols = _clean_cols(raw_df, cols)
    missing_required = [k for k in REQUIRED if cols[k] is None]
    res = {
        "name": os.path.basename(path), "rows": int(len(raw_df)),
        "columns": list(raw_df.columns), "cols": cols,
        "col_fields": [{"key": k, "label": l} for k, l, _g in COL_FIELDS],
        "missing_required": missing_required, "presets": PRESETS,
        "warn_no_pole": cols.get("pole_col") is None,
    }
    if missing_required:
        return res
    df = _proc(path, cols)
    available = []
    for label, col in DATE_FIELD_OPTIONS.items():
        if col in df.columns:
            parsed = pd.to_datetime(df[col], errors="coerce")
            parsed = parsed.where(parsed.dt.year > 1901)
            if parsed.notna().any():
                available.append(label)
    res["date_fields"] = available
    res["default_date_field"] = "DateToUse" if "DateToUse" in available else (available[0] if available else None)

    def opts(col):
        return sorted(df[col].dropna().astype(str).unique()) if col and col in df.columns else []
    res["options"] = {"district": opts(cols["district_col"]), "project": opts(cols["project_col"]),
                      "pid": opts(cols["pid_col"]), "sourcefile": opts(cols["job_col"])}
    return res


def _date_pick(frame, date_field, preset, custom_from, custom_to):
    """pick_date_field_and_range - same field/preset/clip logic."""
    field_options = DATE_FIELD_OPTIONS
    fallback_series = frame["_date"]
    available = {}
    for label, col in field_options.items():
        if col in frame.columns:
            parsed = pd.to_datetime(frame[col], errors="coerce")
            parsed = parsed.where(parsed.dt.year > 1901)
            if parsed.notna().any():
                available[label] = parsed
    if not available:
        if fallback_series is None or fallback_series.isna().all():
            return fallback_series, None, None
        col_min, col_max = fallback_series.min(), fallback_series.max()
        try:
            return fallback_series, (pd.Timestamp(custom_from).date() if custom_from else col_min.date()), \
                (pd.Timestamp(custom_to).date() if custom_to else col_max.date())
        except ValueError:
            return fallback_series, col_min.date(), col_max.date()
    labels = list(available.keys())
    chosen_label = date_field if date_field in labels else ("DateToUse" if "DateToUse" in labels else labels[0])
    dates = available[chosen_label]
    col_min = dates.min().normalize()
    col_max = dates.max().normalize()
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
    else:
        try:
            date_from = pd.Timestamp(custom_from).date() if custom_from else col_min.date()
            date_to = pd.Timestamp(custom_to).date() if custom_to else col_max.date()
        except ValueError:
            date_from, date_to = col_min.date(), col_max.date()
    date_from = max(date_from, col_min.date())
    date_to = min(date_to, col_max.date())
    if date_from > date_to:
        date_from, date_to = col_min.date(), col_max.date()
    return dates, date_from, date_to


def _filtered(path, cols, filters):
    """The sidebar filters: date field/range, District, Project, PID, Source
    file, then cascading Circuit and Pole - same statements as the app."""
    raw_df = _raw(path)
    cols = _clean_cols(raw_df, cols)
    df = _proc(path, cols)
    fl = filters or {}
    district_col, project_col, circuit_col = cols["district_col"], cols["project_col"], cols["circuit_col"]
    pole_col, pid_col, job_col = cols.get("pole_col"), cols.get("pid_col"), cols.get("job_col")
    _dates_active, _date_from_picked, _date_to_picked = _date_pick(
        df, fl.get("date_field"), fl.get("preset", "All time"), fl.get("custom_from"), fl.get("custom_to"))
    df["_date"] = _dates_active
    date_range = (_date_from_picked, _date_to_picked) if _date_from_picked is not None and _date_to_picked is not None else None
    districts, projects = fl.get("districts") or [], fl.get("projects") or []
    pids, sourcefiles = fl.get("pids") or [], fl.get("sourcefiles") or []
    circuits, poles_selected = fl.get("circuits") or [], fl.get("poles") or []

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
    _scope_mask_for_pole = _scope_mask.copy()
    if circuits:
        _scope_mask_for_pole &= df[circuit_col].astype(str).isin(circuits)
    pole_options = sorted(df.loc[_scope_mask_for_pole, pole_col].dropna().astype(str).unique()) if pole_col else []

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
    return df, f, cols, {"circuit_options": circuit_options, "pole_options": pole_options,
                         "date_from": _s(date_range[0]) if date_range else "", "date_to": _s(date_range[1]) if date_range else "",
                         "rows_total": int(len(df)), "rows": int(len(f))}


def metric_value(cat_name, total_qty):
    """render_metric's formatting, returned as text."""
    if cat_name in UNIT_CONFIG:
        return format_length(total_qty, UNIT_CONFIG[cat_name])
    elif cat_name in POLE_DEDUPE_CATEGORIES:
        return f"{total_qty:,.0f} poles"
    return f"{total_qty:,.0f}"


def view(path, cols=None, filters=None, granularity="Month", detail=None):
    df, f, cols, meta = _filtered(path, cols, filters)
    district_col, project_col, circuit_col = cols["district_col"], cols["project_col"], cols["circuit_col"]
    res = dict(meta)

    # ---- Overview ----
    recover_keys = {normalize_item(k) for k in CV7_recover}
    recover_df = f[f["_item_norm"].isin(recover_keys)]
    recover_total = recover_df["_qsub_adj"].sum()
    trend = None
    if not (recover_df.empty or recover_df["_date"].isna().all()):
        freq = {"Day": "D", "Week": "W", "Month": "MS"}.get(granularity, "MS")
        t = (recover_df.dropna(subset=["_date"]).set_index("_date").resample(freq)["_qsub_adj"].sum().reset_index()
             .rename(columns={"_date": "Date", "_qsub_adj": "Count"}))
        trend = {"x": [_s(d) for d in t["Date"]], "y": [_n(v) for v in t["Count"]]}
    pole_rows = []
    for cat_name, mapping in POLE_CATEGORIES.items():
        if cat_name == "CV7_recover":
            continue
        keys = {normalize_item(k) for k in mapping}
        sub = f[f["_item_norm"].isin(keys)]
        if not sub.empty:
            pole_rows.append({"Pole type": POLE_DISPLAY_NAMES.get(cat_name, cat_name), "Count": sub["_qsub_adj"].sum()})
    pole_summary = pd.DataFrame(pole_rows)
    pole_total = pole_summary["Count"].sum() if not pole_summary.empty else 0
    res["overview"] = {"recover_total": _n(recover_total), "trend": trend, "pole_total": _n(pole_total),
                       "poles": [{"type": r["Pole type"], "count": _n(r["Count"])} for r in pole_rows]}

    # ---- Mapped items ----
    cv7_poles = cv7_dedupe_poles(f)
    all_card_data = []
    grouped_cat_names = {c for group in CARD_GROUPS for c in group.get("categories", [])}
    grouped_cat_names |= {c for group in CARD_GROUPS if "subtypes" in group for c in ["Switch"]}
    groups = []
    for group in CARD_GROUPS:
        group_cards = []
        if "subtypes" in group:
            for subtype_name, descriptions in group["subtypes"].items():
                card = build_subtype_card(f, subtype_name, descriptions)
                if card:
                    group_cards.append(card); all_card_data.append(card)
        else:
            for cat_name in group.get("categories", []):
                mapping = ALL_CATEGORIES.get(cat_name)
                if mapping is None:
                    continue
                card = build_card(f, cat_name, mapping, cv7_poles)
                if card:
                    group_cards.append(card); all_card_data.append(card)
        groups.append({"title": group["title"], "cards": [
            {"name": c, "label": POLE_DISPLAY_NAMES.get(c) or c, "value": metric_value(c, q)} for c, q, _s2 in group_cards]})
    remaining_cards = []
    for cat_name in [c for c in ALL_CATEGORIES if c not in grouped_cat_names]:
        card = build_card(f, cat_name, ALL_CATEGORIES[cat_name], cv7_poles)
        if card:
            remaining_cards.append(card); all_card_data.append(card)
    detail_table = None
    names = [c[0] for c in all_card_data]
    if all_card_data:
        chosen = detail if detail in names else names[0]
        _, _, sub = next(c for c in all_card_data if c[0] == chosen)
        detail_df = pd.DataFrame({
            "District": sub[district_col],
            "Job": sub["_job_clean"],
            "Circuit": sub[circuit_col],
            "enid": sub[cols["pole_col"]] if cols["pole_col"] in sub.columns else "",
        })
        detail_table = {"chosen": chosen, **_table(detail_df)}
    res["items"] = {"groups": groups, "other": [{"name": c, "label": c, "value": metric_value(c, q)} for c, q, _s2 in remaining_cards],
                    "detail_names": names, "detail": detail_table}

    # ---- Totals ----
    tot = {"total": None, "orig": None}
    total_val = pd.to_numeric(f[cols["total_col"]], errors="coerce").sum() if cols["total_col"] in f.columns else None
    orig_val = pd.to_numeric(f[cols["orig_col"]], errors="coerce").sum() if cols["orig_col"] in f.columns else None
    tot["total"], tot["orig"] = _n(total_val), _n(orig_val)
    if total_val is not None and project_col in f.columns:
        f["_total_val_row"] = pd.to_numeric(f[cols["total_col"]], errors="coerce")
        by_project_total = (
            f.groupby(project_col)["_total_val_row"].sum().reset_index()
            .rename(columns={project_col: "Project", "_total_val_row": "Total value (£)"})
            .sort_values("Total value (£)", ascending=False)
        )
        tot["by_project"] = _table(by_project_total)
    if total_val is not None and orig_val is not None:
        f["_row_variance"] = pd.to_numeric(f[cols["total_col"]], errors="coerce") - pd.to_numeric(f[cols["orig_col"]], errors="coerce")
        variance_rows = f[f["_row_variance"] != 0]
        variance_table = pd.DataFrame({
            "District": variance_rows[district_col],
            "Job": variance_rows["_job_clean"],
            "Circuit": variance_rows[circuit_col],
            "Difference (£)": variance_rows["_row_variance"],
        })
        tot["variance"] = _table(variance_table.sort_values("Difference (£)", key=abs, ascending=False))
        by_job = (
            variance_rows.assign(Job=variance_rows["_job_clean"])
            .groupby("Job")["_row_variance"].sum().reset_index()
            .rename(columns={"_row_variance": "Difference (£)"})
            .sort_values("Difference (£)", key=abs, ascending=False)
        )
        tot["by_job"] = _table(by_job)
        if project_col in variance_rows.columns:
            by_project = (
                variance_rows.groupby(project_col)["_row_variance"].sum().reset_index()
                .rename(columns={project_col: "Project", "_row_variance": "Difference (£)"})
                .sort_values("Difference (£)", key=abs, ascending=False)
            )
            tot["variance_by_project"] = _table(by_project)
    res["totals"] = tot
    return res


# ---- Jobs tab: Outages Programme + work instructions for a clicked day ----
def _bytes(path):
    key = (path, os.path.getmtime(path))
    if key not in _FILE_CACHE:
        if len(_FILE_CACHE) > 6:
            _FILE_CACHE.pop(next(iter(_FILE_CACHE)))
        with open(path, "rb") as fh:
            _FILE_CACHE[key] = fh.read()
    return _FILE_CACHE[key]


def outages(outage_path, master_path=None, cols=None, filters=None, o_districts=None, o_pms=None,
            o_from=None, o_to=None, selected_date=None):
    outage_df, outage_err = load_outage_programme(_bytes(outage_path))
    if outage_err:
        raise ValueError("Couldn't read the outages workbook - check it has a sheet named '2026' with headers on row 7.\n\n"
                         f"Details: {outage_err}")
    res = {"rows_total": int(len(outage_df)),
           "districts": sorted(str(v) for v in outage_df["District"].dropna().unique()),
           "pms": sorted(str(v) for v in outage_df["SPEN PM"].dropna().unique())}
    if outage_df["Outage Date"].notna().any():
        res["min"], res["max"] = _s(outage_df["Outage Date"].min()), _s(outage_df["Outage Date"].max())
    outage_f = outage_df.copy()
    if o_districts:
        outage_f = outage_f[outage_f["District"].astype(str).isin(o_districts)]
    if o_pms:
        outage_f = outage_f[outage_f["SPEN PM"].astype(str).isin(o_pms)]
    if o_from and o_to:
        o_start = pd.Timestamp(o_from)
        o_end = pd.Timestamp(o_to) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        outage_f = outage_f[outage_f["Outage Date"].between(o_start, o_end)]
    res["rows"] = int(len(outage_f))
    res["table"] = _table(outage_f)
    cal_df = outage_f.dropna(subset=["Outage Date"])
    res["events"] = build_calendar_events(cal_df) if not cal_df.empty else []
    for _e, _d in zip(res["events"], cal_df["District"].tolist()):   # district for the page's colour key
        _e["district"] = "" if pd.isna(_d) else str(_d).strip()

    if selected_date:
        sel = pd.Timestamp(selected_date).date()
        day_rows = outage_f[outage_f["Outage Date"].dt.date == sel]
        day = {"date": sel.isoformat(), "outages": _table(day_rows), "tasks": None, "note": ""}
        if master_path:
            _df, f, cols_, _meta = _filtered(master_path, cols, filters)
            if cols_.get("date_col") is None or f["_date"].isna().all():
                day["note"] = ("No usable dates in the main dataset's mapped 'Date' column - pick one under "
                               "'Column mapping' to enable this breakdown.")
            else:
                day_job_rows = f[f["_date"].dt.date == sel]
                pole_task_df = build_pole_task_table(day_job_rows, cols_["item_col"], cols_["qsub_col"], cols_.get("pole_col"))
                day["tasks"] = _table(pole_task_df)
                day["pole_count"] = int(pole_task_df["Pole"].nunique()) if not pole_task_df.empty else 0
        else:
            day["note"] = "Load the Master file to see the poles & MD Poling tasks for this day."
        res["day"] = day
    return res


def ics(ics_path):
    ics_df, ics_err = parse_ics_events(_bytes(ics_path))
    if ics_err:
        raise ValueError(f"Couldn't parse {os.path.basename(ics_path)}.\n\nDetails: {ics_err}")
    return {"name": os.path.basename(ics_path), "count": int(len(ics_df)),
            **_table(ics_df[["Summary", "Start", "End", "Location", "Description"]])}


# ---- Pole Position (forecast workbook) ----
FC_FIELDS = [
    ("district", "District", ("District",)),
    ("pid", "Project ID", ("Project ID", "PID")),
    ("project", "Project", ("Project", "Project Name")),
    ("circuit", "Circuit", ("Circuit",)),
    ("voltage", "Voltage", ("Voltage",)),
    ("forecast", "Forecasted Total poles", ("Forecasted Total poles", "Forecasted Total Poles", "Forecast Total Poles")),
    ("disposed", "Poles Disposed", ("Poles Disposed", "Poles disposed")),
    ("start_date", "Start Date", ("Start Date", "StartDate", "Start")),
]


def forecast(path, sheet=None, f_cols=None, districts=None, voltages=None, years=None):
    fb = _bytes(path)
    sheet_names, sheet_err = list_forecast_sheets(fb)
    if sheet_err:
        raise ValueError(f"Couldn't open that workbook.\n\nDetails: {sheet_err}")
    if not sheet_names:
        raise ValueError("The workbook has no sheets.")
    sheet_choice = sheet if sheet in sheet_names else sheet_names[0]
    fdf, fdf_err = load_forecast_workbook(fb, sheet_choice)
    if fdf_err:
        raise ValueError(f"Couldn't read sheet '{sheet_choice}'.\n\nDetails: {fdf_err}")

    def fguess(*candidates):
        lower_map = {c.lower(): c for c in fdf.columns}
        for cand in candidates:
            if cand.lower() in lower_map:
                return lower_map[cand.lower()]
        return None
    chosen = dict(f_cols or {})
    fc = {}
    for key, _label, guesses in FC_FIELDS:
        v = chosen.get(key, "__guess__")
        fc[key] = fguess(*guesses) if v == "__guess__" else (v if v in fdf.columns else None)
    res = {"sheets": sheet_names, "sheet": sheet_choice, "columns": list(fdf.columns), "f_cols": fc,
           "fields": [{"key": k, "label": l} for k, l, _g in FC_FIELDS]}
    required = ["project", "circuit", "pid", "forecast", "disposed"]
    res["missing"] = [k for k in required if fc[k] is None]
    if res["missing"]:
        return res
    f_cols = fc
    plot_df = pd.DataFrame({
        "District": fdf[f_cols["district"]] if f_cols["district"] else "",
        "PID": fdf[f_cols["pid"]],
        "Project": fdf[f_cols["project"]],
        "Circuit": fdf[f_cols["circuit"]],
        "Voltage": fdf[f_cols["voltage"]] if f_cols["voltage"] else "",
        "Forecast": pd.to_numeric(fdf[f_cols["forecast"]], errors="coerce").fillna(0),
        "Disposed": pd.to_numeric(fdf[f_cols["disposed"]], errors="coerce").fillna(0),
    })
    if f_cols["start_date"]:
        plot_df["Start Date"] = pd.to_datetime(fdf[f_cols["start_date"]], errors="coerce")
        plot_df["Year"] = plot_df["Start Date"].dt.year
    plot_df = plot_df.dropna(subset=["Project"])
    plot_df["Disposed"] = plot_df[["Disposed", "Forecast"]].min(axis=1)
    plot_df["Remaining"] = (plot_df["Forecast"] - plot_df["Disposed"]).clip(lower=0)
    plot_df["Label"] = (plot_df["Project"].astype(str) + " — " + plot_df["Circuit"].astype(str) + " — PID " + plot_df["PID"].astype(str))
    res["options"] = {
        "district": sorted(str(v) for v in plot_df["District"].dropna().unique()) if f_cols["district"] else [],
        "voltage": sorted(str(v) for v in plot_df["Voltage"].dropna().unique()) if f_cols["voltage"] else [],
        "year": sorted(int(v) for v in plot_df["Year"].dropna().unique()) if "Year" in plot_df.columns else [],
    }
    if districts:
        plot_df = plot_df[plot_df["District"].astype(str).isin(districts)]
    if voltages:
        plot_df = plot_df[plot_df["Voltage"].astype(str).isin(voltages)]
    if years:
        plot_df = plot_df[plot_df["Year"].isin([int(y) for y in years])]
    total_forecast = plot_df["Forecast"].sum()
    total_disposed = plot_df["Disposed"].sum()
    res["banner"] = {"disposed": _n(total_disposed), "forecast": _n(total_forecast),
                     "pct": (float(total_disposed / total_forecast) if total_forecast else None)}
    plot_df = plot_df.sort_values("Forecast", ascending=True)
    res["bars"] = [{"label": r.Label, "disposed": _n(r.Disposed), "remaining": _n(r.Remaining), "forecast": _n(r.Forecast)}
                   for r in plot_df.itertuples(index=False)]
    return res
