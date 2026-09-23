"""Tool 5 - Work Instructions & WorkPack map check (was the 3-tab reporting tool).
Original logic kept verbatim; only the tkinter window was replaced."""
import os
import re
import threading
import traceback
import concurrent.futures
from functools import lru_cache
from difflib import SequenceMatcher
from datetime import datetime
 
 
import pandas as pd
 
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_LINE_SPACING
 
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
 
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
 
try:
    from rapidfuzz import fuzz as _rf_fuzz  # much faster than difflib if installed
except ImportError:
    _rf_fuzz = None
 
 
# ---------------------------------------------------------------------------
# CONFIG - adjust these if your real column names differ even slightly
# ---------------------------------------------------------------------------
# NOTE: these were updated to match the NEW master parquet's column schema
# (produced by the aggregation / build-master pipeline), which renamed:
#   shire -> district, segmentcode -> circuit, pole -> enid
# plan1, project, sourcefile, "MD Poling", qsub, and comment kept their
# original names, so they're unchanged below.
COL_DATE = "plan1"
COL_SHIRE = "district"
COL_PROJECT = "project"
COL_CIRCUIT = "circuit"
COL_SOURCEFILE = "sourcefile"
COL_POLE = "enid"
COL_MDPOLING = "MD Poling"
COL_QSUB = "qsub"
COL_COMMENT = "comment"
# 🔥 New "Type" column (e.g. "Metre", "Km", ...) added upstream in the
# aggregation pipeline, right after qsub. It's OPTIONAL here (not in
# REQUIRED_COLUMNS) so master parquet files built before this column
# existed still load fine - it's only used to append a length-unit
# abbreviation onto the quantity for a specific set of MD Poling entries
# (see UNIT_ITEM_NAMES / _normalize_length_unit below).
COL_TYPE = "type"
 
REQUIRED_COLUMNS = [
    COL_DATE, COL_SHIRE, COL_PROJECT, COL_CIRCUIT, COL_SOURCEFILE,
    COL_POLE, COL_MDPOLING, COL_QSUB, COL_COMMENT,
]
 
# A "pole" is normally an 8-digit number. Some map comments instead use a
# name like "New_Pole" or "Mainline_Pole" (no digit sequence at all) - those
# are only treated as poles when the comment has NO 8-digit number in it,
# i.e. named poles are a fallback, never a replacement for a numeric match.
POLE_REGEX = re.compile(r"\b\d{8}\b")
NAMED_POLE_REGEX = re.compile(r"\b[A-Za-z][A-Za-z0-9]*_Pole\b", re.IGNORECASE)
 
DEFAULT_SIMILARITY_THRESHOLD = 85  # percent
 
# How many PDFs / outage folders to read concurrently. Both the PDF reads
# and the network directory scans are I/O bound (network share latency), so
# doing several at once instead of one-at-a-time is the main speed win.
# Bumped up from the original 8: reading maps is I/O (network share) bound,
# not CPU bound, so a higher thread count keeps more requests in flight at
# once without saturating the CPU. Lower this back down if it ever
# overwhelms the network share.
PDF_READ_WORKERS = 20
NETWORK_SCAN_WORKERS = 20
 
# ---------------------------------------------------------------------------
# Word document formatting (applies to the whole Work Instructions doc)
# ---------------------------------------------------------------------------
WORD_FONT_NAME = "Times New Roman"
WORD_FONT_SIZE_PT = 12
WORD_LINE_SPACING = 2.0  # "double" spacing
 
# ---------------------------------------------------------------------------
# Network folder structure for automatic Workpack-zones discovery
# ---------------------------------------------------------------------------
# Expected layout under NETWORK_ROOT:
#   <root> / <YYYY> / <MM - Month name> / <DD-MM-YYYY - Outage description> / ... / "Workpack zones" / *.pdf
# "Workpack zones" is not always nested at the same depth under the outage
# folder (it can sit directly inside it, or a level or two deeper inside an
# outage-description subfolder), so it is located with a recursive search.
NETWORK_ROOT = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\1 - Outages Programme"
 
YEAR_DIR_REGEX = re.compile(r"^\d{4}$")
MONTH_DIR_REGEX = re.compile(r"^\d{2}\s*-\s*.+$")
WORKPACK_ZONES_NAME = "workpack zones"
 
# Outage folder names come in three shapes:
#   "DD-MM-YYYY - description"                (single day)
#   "DD_DD-MM-YYYY - description"              (day range, same month)
#   "DD-MM-YYYY to DD-MM-YYYY - description"   (full date range, can span months)
_OUTAGE_RANGE_TO_REGEX = re.compile(
    r"^(\d{2})-(\d{2})-(\d{4})\s+to\s+(\d{2})-(\d{2})-(\d{4})\s*-\s*(.+)$", re.IGNORECASE
)
_OUTAGE_SINGLE_OR_DAYRANGE_REGEX = re.compile(r"^(\d{2})(?:_(\d{2}))?-(\d{2})-(\d{4})\s*-\s*(.+)$")
 
 
def _parse_outage_folder_dates(name):
    """Returns (start_date, end_date, description) if `name` looks like a
    dated outage folder in any of the three supported shapes, else None.
    The "DD-MM-YYYY to DD-MM-YYYY" shape is tried first since it's the
    most specific pattern."""
    m = _OUTAGE_RANGE_TO_REGEX.match(name)
    if m:
        d1, mo1, y1, d2, mo2, y2, desc = m.groups()
        try:
            start = datetime(int(y1), int(mo1), int(d1))
            end = datetime(int(y2), int(mo2), int(d2))
        except ValueError:
            return None
        return start, end, desc
 
    m = _OUTAGE_SINGLE_OR_DAYRANGE_REGEX.match(name)
    if m:
        d1, d2, mo, y, desc = m.groups()
        try:
            start = datetime(int(y), int(mo), int(d1))
            end = datetime(int(y), int(mo), int(d2)) if d2 else start
        except ValueError:
            return None
        return start, end, desc
 
    return None
 
 
# ---------------------------------------------------------------------------
# Core data helpers (no GUI dependency - these are independently testable)
# ---------------------------------------------------------------------------
def load_master_parquet(path):
    """Load the master parquet file and parse the date column."""
    df = pd.read_parquet(path)
 
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "The parquet file is missing these expected columns: "
            + ", ".join(missing)
            + "\nColumns found: " + ", ".join(df.columns)
        )
 
    df = df.copy()
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
 
    # Make sure text columns are plain strings (avoids NaN/float surprises)
    for c in [COL_SHIRE, COL_PROJECT, COL_CIRCUIT, COL_SOURCEFILE,
              COL_POLE, COL_MDPOLING, COL_QSUB, COL_COMMENT]:
        df[c] = df[c].fillna("").astype(str)
 
    # COL_TYPE is optional (older master files won't have it) - only clean
    # it up when present.
    if COL_TYPE in df.columns:
        df[COL_TYPE] = df[COL_TYPE].fillna("").astype(str)
 
    return df
 
 
def unique_sorted(df, col):
    vals = sorted(v for v in df[col].dropna().unique() if str(v).strip() != "")
    return vals
 
 
def apply_filters(df, shires=None, projects=None, circuits=None,
                   sourcefiles=None, poles=None, date_from=None, date_to=None):
    """
    Each of shires/projects/circuits/sourcefiles/poles is a list of values.
    An empty/None list means "no filter" on that field.
    date_from / date_to are datetime objects or None.
    """
    out = df
 
    def apply_list_filter(frame, col, values):
        if values:
            return frame[frame[col].isin(values)]
        return frame
 
    out = apply_list_filter(out, COL_SHIRE, shires)
    out = apply_list_filter(out, COL_PROJECT, projects)
    out = apply_list_filter(out, COL_CIRCUIT, circuits)
    out = apply_list_filter(out, COL_SOURCEFILE, sourcefiles)
    out = apply_list_filter(out, COL_POLE, poles)
 
    if date_from is not None:
        out = out[out[COL_DATE] >= date_from]
    if date_to is not None:
        out = out[out[COL_DATE] <= date_to]
 
    return out
 
 
def is_erect(mdpoling_text):
    """True if the MD Poling text mentions 'Erect' anywhere (case-insensitive).
    This also naturally covers 'Erect New SOULE/Auto Reclosure'."""
    return "erect" in str(mdpoling_text).lower()
 
 
def is_plumb(mdpoling_text):
    """True if the MD Poling text mentions 'Plumb' anywhere (case-insensitive)."""
    return "plumb" in str(mdpoling_text).lower()
 
 
# A handful of standard "Erect ..." MD Poling entries should NOT get the
# red-highlight treatment that other "Erect" entries get - they're shown
# as plain text instead. Written as flexible patterns (rather than exact
# strings) since punctuation/wording in the source data can vary slightly
# (e.g. "1ph" vs "1.ph", "Termination" vs "Term..."); each pattern only
# anchors on the safely-common prefix of the real MD Poling text.
ERECT_NO_HIGHLIGHT_PATTERNS = [
    re.compile(r"erect\s*1\.?\s*ph\s+lv\s+cable\s+pole\s+term", re.IGNORECASE),   # Erect 1ph/1.ph LV cable pole term(ination)
    re.compile(r"erect\s*3\.?\s*ph\s+lv\s+cable\s+pole\s+term", re.IGNORECASE),   # Erect 3ph/3.ph LV cable pole term(ination)
    re.compile(r"erect\s*11\s*kv\s+cable\s+term", re.IGNORECASE),                 # Erect 11kV Cable Term(ination) ...
    re.compile(r"erect\s+support\s+steelwork", re.IGNORECASE),                    # Erect support steelwork
]
 
 
def is_erect_no_highlight(mdpoling_text):
    """True if this 'Erect' MD Poling entry is one of the standard ones that
    should be shown as plain text instead of being highlighted red."""
    text = str(mdpoling_text or "")
    return any(p.search(text) for p in ERECT_NO_HIGHLIGHT_PATTERNS)
 
 
def similarity(a, b):
    """Simple % similarity between two strings (0-100). Uses rapidfuzz when
    available (much faster - written in C), falls back to stdlib difflib."""
    a, b = str(a or ""), str(b or "")
    if not a.strip() or not b.strip():
        return 0.0
    a_n, b_n = a.lower().strip(), b.lower().strip()
    if _rf_fuzz is not None:
        return _rf_fuzz.ratio(a_n, b_n)
    return SequenceMatcher(None, a_n, b_n).ratio() * 100.0
 
 
def find_poles(text):
    """Return all pole identifiers found in a piece of text: 8-digit pole
    numbers if any are present, otherwise fall back to named poles like
    "New_Pole" / "Mainline_Pole" (never both at once for the same text)."""
    text = str(text or "")
    numeric = POLE_REGEX.findall(text)
    if numeric:
        return numeric
    return NAMED_POLE_REGEX.findall(text)
 
 
# ---------------------------------------------------------------------------
# MD Poling "work instruction" rules
# ---------------------------------------------------------------------------
# Each rule says, for a given MD Poling instruction type:
#   structure : how the instruction text/comment is formatted in the Word doc
#                 "plain"                 - no special formatting
#                 "highlight_comment_red" - the "(...)" text is highlighted red
#                                           (if there's no "(...)" text, the MD
#                                           Poling text itself is highlighted red)
#                 "bold_comment"          - the "(...)" text is shown in bold
#                 "red_text"              - the MD Poling instruction text itself
#                                           is shown in red, underlined font. This
#                                           structure is exclusively used by the
#                                           "joint" items (see the formatting note
#                                           below) - qty/comment order and display
#                                           for these is handled specially in
#                                           _emit_pole_block.
#   order     : workflow position (float so items can be inserted between two
#                whole numbers, e.g. 2.1 sits right after all the "2"s and
#                before any "3"s). Used to order the different MD Poling
#                entries that belong to the SAME pole.
#   comment   : True if this instruction's own "comment" column value should
#                be included in the displayed text
#   qsub      : True if this instruction's own "qsub" column value should be
#                included in the displayed text.
#
# 🔥 QUANTITY / COMMENT DISPLAY RULES (updated):
#   - A quantity of exactly 1 is never shown (e.g. "x1" is redundant).
#   - For every rule EXCEPT the "red_text" ("joint") ones: the COMMENT comes
#     first, shown in its own parentheses, then the quantity follows as
#     plain text (NOT wrapped in parentheses), e.g.
#     "Erect Pole (some comment) x2".
#   - For the "red_text" ("joint") rules specifically: quantity and comment
#     are combined into ONE parenthetical, with the QUANTITY FIRST and the
#     comment after, e.g. "LV Straight Joint (x2 some comment)". These
#     entries are also always shown in red, underlined text.
#   - For a specific subset of rules (see UNIT_ITEM_NAMES below), a length
#     unit abbreviation is appended right after the quantity, taken from
#     that row's "Type" column - e.g. "x120 m" or "x1.5 Km". "Metre"/"per
#     meter"/"per m"/"per metre" (and similar) normalize to "m";
#     "Km"/"per Km"/"km"/"per km" (and similar) normalize to "Km"; any
#     other Type value is shown as-is (never silently dropped).
#
# NOTE: matching against the real MD Poling text in the data is done with
# match_mdpoling_rule() below - it tries an exact (case/whitespace
# insensitive) match first, then falls back to the longest rule phrase that
# appears as a substring of the real text (or vice versa), so small wording
# differences in the source data still resolve to the right rule.
MD_POLING_RULES = [
    ("Erect Pole", "highlight_comment_red", 1, True, False),
    ("Recover Pole", "plain", 2, True, False),
    ("Recover H pole", "plain", 2, True, True),
    ("Recover Tower", "plain", 2, True, False),
    ("Fit Block", "plain", 14, False, True),
    ("Fit bog shoe", "plain", 14, False, True),
    ("Install Bare Conductor", "plain", 10, True, True),
    ("Install Conductor", "plain", 10, True, True),
    ("Install ABC 2c", "plain", 10, True, True),
    ("Install ABC 4c", "plain", 10, True, True),
    ("Install ABC 2c + E", "plain", 10, True, True),
    ("Install ABC 4c + E", "plain", 10, True, True),
    ("Bind Conductor", "plain", 13, False, True),
    ("1ph cond term including jumpers", "plain", 13, False, True),
    ("3ph cond term including jumpers", "plain", 13, False, True),
    ("Unbind & reregulate conductors", "plain", 13, False, True),
    ("Recover OHL & fittings", "plain", 2.1, True, True),
    ("Recover Cleated Service", "plain", 2.1, True, True),
    ("Erect New 1ph TX", "bold_comment", 3, True, False),
    ("Erect New 3ph TX", "bold_comment", 3, True, False),
    ("Erect New ABSW", "bold_comment", 3, True, False),
    ("Erect New SOULE/Auto Reclosure", "bold_comment", 3, True, False),
    ("Erect Voltage Regulator", "plain", 4.1, True, False),
    ("Erect Voltage Transformer/RTU/Repeater", "plain", 4.1, True, False),
    ("Surge arrestors", "plain", 4.1, True, False),
    ("Erect 1ph fuse units", "plain", 4.1, True, True),
    ("Erect 3ph fuse units", "plain", 4.1, True, True),
    ("Fitting fuse outrigger bracket", "plain", 4.1, True, False),
    ("Lower/Raise existing TX", "plain", 4, True, False),
    ("Lower/Raise existing equipment (HV Fuses/Pole Box/ABSW)", "plain", 4, True, False),
    ("Remove TX", "plain", 4, True, False),
    ("Remove platform/TX", "plain", 4, True, False),
    ("Remove HV fuses", "plain", 4, True, False),
    ("Remove ABSW", "plain", 4, True, False),
    ("Remove Auto Reclosure", "plain", 4, True, False),
    ("Remove Street Light", "plain", 14, False, True),
    ("Stay Above", "plain", 5, False, True),
    ("SC", "plain", 5, False, True),
    ("SC (rock)", "plain", 5, False, True),
    ("Tighten stay", "plain", 5, False, True),
    ("HI-Vi SG", "plain", 5, False, False),
    ("Stay Outrigger", "plain", 5, True, False),
    ("Recover Stay", "plain", 5, True, False),
    ("Single X-Arm T-Off", "plain", 5, True, False),
    ("Double X-Arm T-Off", "plain", 5, True, False),
    ("Double X-Arm H T-Off", "plain", 5, True, False),
    ("Erect support steelwork", "plain", 5, True, False),
    ("Remove X-Arm", "plain", 5, True, False),
    ("Convert 1ph Inter to Section Pole", "plain", 2.1, False, False),
    ("Convert pole to Termial", "plain", 2.1, False, False),
    ("Convert 3ph Inter to Section Pole", "plain", 2.1, False, False),
    ("Change Insulators", "plain", 13, False, True),
    ("Replace Pin Insulators", "plain", 13, False, True),
    ("Replace Insulator Binder", "plain", 13, False, True),
    ("Tension Insulator", "plain", 13, False, True),
    ("Dead end termination", "plain", 13, False, True),
    ("Pilot Pin", "plain", 13, False, True),
    ("ABC T-off Connection to Inter/Section Support", "plain", 13, True, True),
    ("Rewire 1ph TX", "plain", 4.1, False, False),
    ("Rewire 3ph TX", "plain", 4.1, False, False),
    ("cable wiring including cleating", "plain", 13, True, True),
    ("Shrouding", "plain", 13, False, False),
    ("transition ABC to Open Wire 3wires", "plain", 13, True, True),
    ("transition ABC to Open Wire 5wires", "plain", 13, True, True),
    ("Install service span", "plain", 13, True, True),
    ("Take-off bracket", "plain", 13, True, True),
    ("Concentric cable termination", "plain", 13, True, True),
    ("Install Extension bracket", "plain", 13, True, True),
    ("Concentric Cable Tee Joint", "plain", 13, True, True),
    ("Wall box", "plain", 14, False, True),
    ("Install 1ph/Split cut out", "plain", 13, True, True),
    ("Install 3ph cut out", "plain", 13, True, True),
    ("Retrofit Suspension Clamp", "plain", 13, True, True),
    ("Retrofit Anchor Clamp", "plain", 13, True, True),
    ("Remake Jumpers", "plain", 13, False, True),
    ("Raise OH service", "plain", 13, True, True),
    ("LV fuse units 1ph transformer", "plain", 4.1, False, False),
    ("LV fuse units Split ph transformer", "plain", 4.1, False, False),
    ("LV fuse units 3ph transformer", "plain", 4.1, False, False),
    ("Replace/Reposition LV pole mounted fuse unit", "plain", 13, True, True),
    ("Reconnect LV earth", "plain", 11, True, True),
    ("Reconnect LV mains cable", "plain", 11, True, True),
    ("Reconnect LV service cable 1ph", "plain", 11, True, True),
    ("Reconnect LV service cable 3ph", "plain", 11, True, True),
    ("Reconnect oh 1ph service cables", "plain", 11, True, True),
    ("Reconnect oh 3ph service cables", "plain", 11, True, True),
    ("Reconnect ABC connection", "plain", 11, True, True),
    ("Reconnect street lamp", "plain", 13, False, False),
    ("Refit Telecom line", "plain", 13, False, False),
    ("Reconnect LV fuses 3.ph including wiring", "plain", 13, True, True),
    ("Reconnect distribution box", "plain", 13, True, True),
    ("Convert Inter pole to Terminal", "plain", 13, True, True),
    ("Convert Inter pole to section 2 wires", "plain", 13, True, True),
    ("Convert Inter pole to section 4 wires", "plain", 13, True, True),
    ("Plumb pole", "plain", 5, False, False),
    ("Repair pole twist", "plain", 2, False, False),
    ("NP", "plain", 14, True, True),
    ("Wrap ACD", "plain", 14, True, True),
    ("HV/LV Earth", "plain", 5.1, True, True),
    ("Bonding", "plain", 12, True, True),
    ("CG", "plain", 12, True, True),
    ("Replace Jumper", "plain", 12, True, True),
    ("Repair conductor (mid span)", "plain", 12, True, True),
    ("Conductor repair", "plain", 12, True, True),
    ("Green/Yellow or Black Earth", "plain", 12, True, True),
    ("HV Earth Electrode", "plain", 12, True, True),
    ("LV Earth Electrode", "plain", 12, True, True),
    ("Additional earthing", "plain", 12, True, True),
    ("Earth Electrode cable trench", "plain", 12, True, True),
    ("Erect 11kV Cable Term & surge arrestors", "plain", 7.1, True, True),
    ("Erect 33kV Cable Term & surge arrestors", "plain", 7.1, True, True),
    ("Erect 1ph LV cable pole term", "plain", 5.2, True, True),
    ("Erect 3ph LV cable pole term", "plain", 5.2, True, True),
    ("Remove HV cable term", "plain", 5.3, True, True),
    ("Remove LV cable term", "plain", 5.3, True, True),
    ("Joint Bay", "plain", 6.1, True, True),
    ("Exposure Bay", "plain", 6, True, True),
    ("3ph LV pole term", "plain", 6, True, True),
    ("1ph LV pole term", "plain", 6, True, True),
    ("11kV pole term", "plain", 6, True, True),
    ("33kV pole term", "plain", 6, True, True),
    ("LV Straight Joint", "red_text", 7, True, True),
    ("LV Breach Joint", "red_text", 7, True, True),
    ("LV Pot End", "red_text", 7, True, True),
    # 🔥 These two were "plain" before - moved to "red_text" so they get the
    # same red/underlined treatment and comment-then-quantity ordering as
    # the rest of the "joint" items below (per the user's grouping of all
    # 15 joint-type entries together).
    ("1ph LV Service Straight Joint", "red_text", 7, True, True),
    ("3ph Service LV Straight Joint", "red_text", 7, True, True),
    ("Single LV Service Breech Joint", "red_text", 7, True, True),
    ("Service Breech Joint", "red_text", 7, True, True),
    ("Service Pot End", "red_text", 7, True, True),
    ("11kV Straight Joint", "red_text", 7, True, True),
    ("11kV Transition Joint", "red_text", 7, True, True),
    ("11kV Breach/Loop", "red_text", 7, True, True),
    ("11kV Pot End", "red_text", 7, True, True),
    ("33kV Straight Joint", "red_text", 7, True, True),
    ("33kV Transition Joint", "red_text", 7, True, True),
    ("33kV Por End", "red_text", 7, True, True),
    ("Pilot Cable Joint", "plain", 8, True, True),
    ("UG LV Cable", "plain", 8, True, True),
    ("UG LV in trench dug by others", "plain", 8, True, True),
    ("UG LV in existing duct", "plain", 8, True, True),
    ("UG HV Cable", "plain", 8, True, True),
    ("UG HV in trench dug by others", "plain", 8, True, True),
    ("UG HV in existing duct", "plain", 8, True, True),
    ("Excavate, lay, backfill & reinstate", "plain", 7.1, True, True),
    ("Installation of cable only in trench", "plain", 9, True, True),
    ("Install cable in existing duct", "plain", 9, True, True),
    ("Install single duct in trench", "plain", 9, True, True),
    ("Install trefoil duct group in trench", "plain", 9, True, True),
    ("Generator", "plain", 9, True, True),
    ("Jointing Platform", "plain", 9, True, True),
    ("Rock", "plain", 9, True, True),
    ("Bird divertors", "plain", 9, False, True),
    ("Fit 1ph jumpers", "plain", 9, True, True),
    ("Fit 3ph jumpers", "plain", 9, True, True),
    ("Install live line bails 1ph", "plain", 9, True, True),
    ("Install live line bails 3ph", "plain", 9, True, True),
]
 
# 🔥 MD Poling rules whose quantity should also get a length-unit
# abbreviation appended (from that row's "Type" column) - see
# _normalize_length_unit. Matched against the rule's canonical name
# (case/whitespace-insensitive, via _normalize_mdpoling), same as the rest
# of the rule matching.
UNIT_ITEM_NAMES = {
    "Install Bare Conductor",
    "Install Conductor",
    "Install ABC 2c",
    "Install ABC 4c",
    "Install ABC 2c + E",
    "Install ABC 4c + E",
    "Unbind & reregulate conductors",
    "HV Earth Electrode",
    "Recover OHL & fittings",
    "Recover Cleated Service",
    "UG LV Cable",
    "UG LV in trench dug by others",
    "UG LV in existing duct",
    "UG HV Cable",
    "UG HV in trench dug by others",
    "UG HV in existing duct",
    "Excavate, lay, backfill & reinstate",
    "Installation of cable only in trench",
    "Install cable in existing duct",
    "Install single duct in trench",
    "Install trefoil duct group in trench",
}
 
# Fallback used when an MD Poling value in the data doesn't match anything
# in MD_POLING_RULES above. This intentionally reproduces the ORIGINAL
# (pre-rules-table) behaviour: "Erect ..." entries (other than the small
# ERECT_NO_HIGHLIGHT_PATTERNS set) are highlighted red, everything else is
# plain, and only the row's own "comment" column is ever shown - never qsub.
# Its order (999) puts unmatched entries at the very end of a pole's list.
_FALLBACK_ORDER = 999.0
 
# A short substring-fallback match (see match_mdpoling_rule) is only
# trustworthy when both sides of the comparison are reasonably descriptive
# text - a stray value like "1" or "0.0" is technically a substring of lots
# of real rule phrases ("1ph cond term including jumpers" contains "1"), so
# without this floor those junk values would spuriously match an unrelated
# rule and surface as a bare, disconnected "(x1 ...)" bullet. Real MD Poling
# rule phrases are all well above this length, so it never blocks a
# legitimate (if abbreviated) match.
_MIN_SUBSTRING_MATCH_LEN = 4
 
# A cleaned MD Poling value that is PURELY numeric (with an optional single
# decimal point) is never a real instruction - actual instructions are
# always descriptive text. This catches numeric noise (e.g. a qsub value
# that ended up in the MD Poling column, "0.0", "1", "07") that _clean()'s
# blank check alone wouldn't catch, so it doesn't get treated as a
# standalone "instruction" and generate a disconnected qsub/comment bullet.
_NUMERIC_JUNK_REGEX = re.compile(r"^\d+(\.\d+)?$")
 
 
def _is_junk_mdpoling(cleaned_text):
    """True if a cleaned MD Poling value is noise (purely numeric) rather
    than a real instruction, and should be treated the same as blank."""
    return bool(_NUMERIC_JUNK_REGEX.match(cleaned_text))
 
 
def _normalize_mdpoling(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()
 
 
# 🔥 Length-unit abbreviation lookup for UNIT_ITEM_NAMES. Km is checked
# first since "km" would otherwise also satisfy the looser "m" pattern.
_KM_UNIT_REGEX = re.compile(r"\bkm\b|kilomet", re.IGNORECASE)
_M_UNIT_REGEX = re.compile(r"\bm\b|metre|meter", re.IGNORECASE)
 
 
def _normalize_length_unit(type_text):
    """Maps a raw "Type" column value to a short unit abbreviation:
    "Metre" / "per meter" / "per m" / "per metre" (and similar) -> "m";
    "Km" / "per Km" / "km" / "per km" (and similar) -> "Km". A blank Type
    value returns "" (no unit shown). Anything else that doesn't match
    either pattern is shown AS-IS (the raw Type text) rather than being
    dropped - so every one of the UNIT_ITEM_NAMES rules always gets
    whatever Type value the row actually has."""
    t = str(type_text or "").strip()
    if not t:
        return ""
    if _KM_UNIT_REGEX.search(t):
        return "Km"
    if _M_UNIT_REGEX.search(t):
        return "m"
    return t
 
 
_UNIT_ITEM_NAMES_NORM = {_normalize_mdpoling(n) for n in UNIT_ITEM_NAMES}
_RULES_EXACT = {_normalize_mdpoling(rule[0]): rule for rule in MD_POLING_RULES}
# Longest phrase first, so substring matching prefers the most specific rule.
_RULES_BY_LENGTH = sorted(MD_POLING_RULES, key=lambda r: len(r[0]), reverse=True)
 
 
@lru_cache(maxsize=4096)
def match_mdpoling_rule(mdpoling_text):
    """
    Returns (structure, order, comment_flag, qsub_flag, matched, rule_name)
    for a real MD Poling value from the data. Tries, in order:
      1. an exact match (case/whitespace insensitive) against the rules table
      2. the longest rule phrase that appears as a substring of the real
         text, or where the real text appears as a substring of the rule
         phrase (covers minor wording differences either direction) - only
         attempted when the real text is at least _MIN_SUBSTRING_MATCH_LEN
         characters, so short/junk values can't spuriously match
      3. a legacy fallback (see _FALLBACK_ORDER above) when nothing matches
    `matched` is False only for the legacy fallback case. `rule_name` is the
    matched rule's canonical name (used to check UNIT_ITEM_NAMES), or None
    for the legacy fallback.
 
    Cached (@lru_cache) since the same MD Poling text repeats across many
    rows in a typical Control File, and this function is pure - memoizing
    it avoids re-running the whole rule search for every single row.
    """
    norm = _normalize_mdpoling(mdpoling_text)
    if not norm:
        return "plain", _FALLBACK_ORDER, False, False, False, None
 
    exact = _RULES_EXACT.get(norm)
    if exact is not None:
        name, structure, order, comment_flag, qsub_flag = exact
        return structure, float(order), comment_flag, qsub_flag, True, name
 
    if len(norm) >= _MIN_SUBSTRING_MATCH_LEN:
        for rule in _RULES_BY_LENGTH:
            rule_norm = _normalize_mdpoling(rule[0])
            if rule_norm and (rule_norm in norm or norm in rule_norm):
                name, structure, order, comment_flag, qsub_flag = rule
                return structure, float(order), comment_flag, qsub_flag, True, name
 
    # Legacy fallback - keep the original Erect-based red highlight so any
    # MD Poling text not (yet) listed in MD_POLING_RULES still behaves the
    # way this report always used to.
    if is_erect(mdpoling_text) and not is_erect_no_highlight(mdpoling_text):
        return "highlight_comment_red", _FALLBACK_ORDER, True, False, False, None
    return "plain", _FALLBACK_ORDER, True, False, False, None
 
 
# ---------------------------------------------------------------------------
# Word document #1: Work Instructions (pole / MD Poling listing)
# ---------------------------------------------------------------------------
def _dedup_keep_order(values):
    """Remove blanks/duplicates from a list while keeping first-seen order."""
    seen = set()
    out = []
    for v in values:
        v = _clean(v)
        if not v or v in seen:
            continue
        seen.add(v)
        out.append(v)
    return out
 
 
def _clean(value):
    """Strip a value and treat NaN/empty look-alikes ('nan', 'none', 'nat',
    '', '0') as blank. Returns '' for anything that should be omitted."""
    s = str(value).strip()
    if s.lower() in ("", "nan", "none", "nat", "0"):
        return ""
    return s
 
 
def _set_document_base_style(doc):
    """Apply the house style to the whole Work Instructions document: Times
    New Roman 12pt, double line spacing, for the Normal style (which almost
    every paragraph in this document inherits from)."""
    normal = doc.styles["Normal"]
    normal.font.name = WORD_FONT_NAME
    normal.font.size = Pt(WORD_FONT_SIZE_PT)
    # East-asian font fallback - keeps Word from silently substituting a
    # different font for the same run in some locales.
    rpr = normal.element.get_or_add_rPr()
    rFonts = rpr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    if rFonts is None:
        from docx.oxml.ns import qn
        rFonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rFonts)
    from docx.oxml.ns import qn
    rFonts.set(qn("w:eastAsia"), WORD_FONT_NAME)
 
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.line_spacing = WORD_LINE_SPACING
 
 
def _style_paragraph(p):
    """Make sure a paragraph (and anything already added to it) uses the
    house font/spacing - belt-and-braces on top of the Normal style, since
    some paragraph styles (e.g. 'List Bullet') don't always inherit cleanly."""
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    p.paragraph_format.line_spacing = WORD_LINE_SPACING
    for run in p.runs:
        run.font.name = WORD_FONT_NAME
        run.font.size = Pt(WORD_FONT_SIZE_PT)
 
 
def _build_display_parts(qsub_raw, comment_raw, comment_flag, qsub_flag, unit_abbrev=""):
    """
    Builds the two display fragments for one MD Poling entry from raw
    qsub/comment values, per its rule's comment/qsub flags:
 
      qty_text     - "x<value>" (e.g. "x2"), or "x<value> <unit>" (e.g.
                     "x120 m") when unit_abbrev is given. "" when qsub_flag
                     is False, the value is blank, or the value is exactly
                     "1" (a quantity of 1 is redundant and is never shown).
      comment_text - the cleaned comment value, or "" when comment_flag is
                     False or the value is blank.
 
    How these two get combined/positioned is decided by the caller
    (_emit_pole_block), since it differs for "red_text" (joint) rules vs
    everything else.
    """
    qty_text = ""
    if qsub_flag:
        q = _clean(qsub_raw)
        if q and q != "1":
            qty_text = f"x{q}"
            if unit_abbrev:
                qty_text += f" {unit_abbrev}"
 
    comment_text = ""
    if comment_flag:
        comment_text = _clean(comment_raw)
 
    return qty_text, comment_text
 
 
def _map_display_name(file_key):
    """Map file name shown in a block's meta line - the PDF's own file name,
    without its extension (e.g. "Workmap 1.pdf" -> "Workmap 1")."""
    return os.path.splitext(str(file_key or ""))[0] or str(file_key or "")
 
 
def _group_block_entries_by_map(block_entries, pole_sequence_map):
    """
    Splits one header block's poles into map-sized groups, "per the map
    sequence, from first to last":
      - poles found on a map are grouped by which map (PDF) they were found
        on, each group's poles ordered by their position within that map
      - map groups themselves are ordered by the earliest map position of
        any pole in them (so the block reads in the same order the maps
        would be worked through)
      - poles with NO map data (not found on any scanned map, or no maps
        were scanned at all) are kept - per Control File data is never
        dropped for lack of a map match - in their own trailing group, in
        their original data order, with no map name attached
 
    Returns a list of (map_display_name_or_None, [(pole, entries), ...]).
    When `pole_sequence_map` is empty/None, everything is unmapped and this
    returns a single (None, [...]) group in the original order, so the
    output is identical to the pre-map-splitting behaviour.
    """
    mapped = {}   # file_key -> list of (pole, entries, seq, fallback_index)
    unmapped = []  # list of (pole, entries, fallback_index)
 
    for pole, entries, fallback_index in block_entries:
        loc = pole_sequence_map.get(pole) if pole_sequence_map else None
        if loc is not None:
            file_key, seq = loc
            mapped.setdefault(file_key, []).append((pole, entries, seq, fallback_index))
        else:
            unmapped.append((pole, entries, fallback_index))
 
    ordered_maps = sorted(mapped.items(), key=lambda kv: min(item[2] for item in kv[1]))
 
    groups = []
    for file_key, items in ordered_maps:
        items.sort(key=lambda item: item[2])
        groups.append((_map_display_name(file_key), [(p, e) for p, e, _seq, _fi in items]))
 
    if unmapped or not groups:
        unmapped.sort(key=lambda item: item[2])
        groups.append((None, [(p, e) for p, e, _fi in unmapped]))
 
    return groups
 
 
def _emit_pole_block(doc, date_val, sourcefile, group_df, pole_sequence_map=None):
    """
    Write one header block (date / sourcefile / metadata line) followed by
    a bullet per pole in group_df, e.g.:
 
        03 August 2026
        108. Connections _2025 New  - Control file v0.xlsb
        ayrshire   |   connections   |   sp74823   |   Jack Murray   |   03-08-2026
        - <pole> ...
 
    Poles are grouped and ordered per the map sequence when
    `pole_sequence_map` is supplied: poles are split into one sub-group per
    map (PDF), each sub-group's meta line repeats the usual
    shire | project | circuit | qsub | date info with the map's name
    appended, and its poles are listed in that map's first-to-last order.
    Poles with no map match keep their own trailing sub-group (no map name).
    Without `pole_sequence_map`, the block is a single group in the
    original data order, with no map name - unchanged from before.
 
    Within a single pole, its MD Poling entries are ordered by
    MD_POLING_RULES' workflow "order" (see match_mdpoling_rule); the
    original behaviour is preserved for any MD Poling text not in that table.
 
    A row with a blank/"0"/NaN MD Poling value - or one that's purely
    numeric noise (e.g. a qsub value that ended up in the wrong column,
    "0.0", "1") - is never turned into an entry: such a row's comment/qsub
    is not "connected" to any real work instruction, so it's left out
    entirely (this also means a pole whose rows are ALL like that
    contributes nothing and is dropped).
 
    Formatting: the pole number is the only bold text. For every rule
    except the "red_text" (joint) ones, the comment (if any) is shown first,
    in its own parentheses, then the quantity (when not 1) follows as plain
    text right after (e.g. "x2" - plus a unit abbreviation for
    UNIT_ITEM_NAMES rules, e.g. "x120 m"). "highlight_comment_red" still
    highlights its comment text red (or the instruction text if there's no
    comment), and "bold_comment" shows its comment in parentheses like every
    other entry. "red_text" rules show the instruction text itself in red,
    underlined font, with quantity and comment combined into ONE
    parenthetical - quantity first, comment after.
 
    Returns True if anything was written (a block with no usable MD Poling
    entries in any of its poles is skipped entirely, and nothing is written).
    """
    block_entries = []
    for fallback_index, (pole, pole_df) in enumerate(group_df.groupby("_pole", sort=False)):
        if not pole:
            continue
        # Pull the columns we need as plain lists once per pole, then zip -
        # much faster than pole_df.iterrows() (which builds a full pandas
        # Series per row) for what is usually a very small number of rows.
        mdpoling_vals = pole_df[COL_MDPOLING].tolist()
        qsub_vals = pole_df[COL_QSUB].tolist()
        comment_vals = pole_df[COL_COMMENT].tolist()
        if COL_TYPE in pole_df.columns:
            type_vals = pole_df[COL_TYPE].tolist()
        else:
            type_vals = [""] * len(pole_df)
        entries = []
        for mdpoling_raw, qsub_raw, comment_raw, type_raw in zip(
            mdpoling_vals, qsub_vals, comment_vals, type_vals
        ):
            mdpoling = _clean(mdpoling_raw)
            if not mdpoling or _is_junk_mdpoling(mdpoling):
                continue  # blank/"0"/purely-numeric MD Poling - skip this
                          # row entirely, its comment/qsub (if any) has no
                          # real instruction to attach to and is never
                          # shown on its own.
            structure, order, comment_flag, qsub_flag, _matched, rule_name = match_mdpoling_rule(mdpoling)
 
            unit_abbrev = ""
            if rule_name and _normalize_mdpoling(rule_name) in _UNIT_ITEM_NAMES_NORM:
                unit_abbrev = _normalize_length_unit(type_raw)
 
            qty_text, comment_text = _build_display_parts(
                qsub_raw, comment_raw, comment_flag, qsub_flag, unit_abbrev
            )
            entries.append({
                "mdpoling": mdpoling,
                "qty_text": qty_text,
                "comment_text": comment_text,
                "structure": structure,
                "order": order,
            })
        if entries:
            # Stable sort - entries with the same order keep their original
            # (data) relative order.
            entries.sort(key=lambda e: e["order"])
            block_entries.append((pole, entries, fallback_index))
 
    if not block_entries:
        return False
 
    map_groups = _group_block_entries_by_map(block_entries, pole_sequence_map)
 
    first_row = group_df.iloc[0]
 
    if date_val is not None:
        p_date = doc.add_paragraph()
        run = p_date.add_run(date_val.strftime("%d %B %Y"))
        run.bold = True
        _style_paragraph(p_date)
    if sourcefile:
        p_sf = doc.add_paragraph()
        run = p_sf.add_run(sourcefile)
        run.bold = True
        _style_paragraph(p_sf)
 
    short_date = date_val.strftime("%d-%m-%Y") if date_val is not None else ""
    base_meta_parts = [
        _clean(first_row[COL_SHIRE]),
        _clean(first_row[COL_PROJECT]),
        _clean(first_row[COL_CIRCUIT]),
        _clean(first_row[COL_QSUB]),
        short_date,
    ]
    base_meta_parts = [m for m in base_meta_parts if m]
 
    for map_name, poles_in_group in map_groups:
        meta_parts = list(base_meta_parts)
        if map_name:
            meta_parts.append(map_name)
        if meta_parts:
            p_meta = doc.add_paragraph("   |   ".join(meta_parts))
            _style_paragraph(p_meta)
 
        for pole, entries in poles_in_group:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(6)
 
            for i, entry in enumerate(entries):
                if i == 0:
                    # Only the pole number itself is bold - everything else
                    # (MD Poling text, comment, qsub) is plain weight.
                    run_pole = p.add_run(f"{pole} \u2013 ")
                    run_pole.bold = True
                    run_main = p.add_run(entry["mdpoling"])
                else:
                    p.add_run(", ")
                    run_main = p.add_run(entry["mdpoling"])
 
                structure = entry["structure"]
                qty_text = entry["qty_text"]
                comment_text = entry["comment_text"]
 
                if structure == "red_text":
                    # "Joint" items: instruction text itself shown in red,
                    # underlined font. Quantity and comment are combined
                    # into ONE parenthetical - quantity first, comment after.
                    run_main.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
                    run_main.font.underline = True
                    paren_parts = []
                    if qty_text:
                        paren_parts.append(qty_text)
                    if comment_text:
                        paren_parts.append(comment_text)
                    if paren_parts:
                        p.add_run(f" ({' '.join(paren_parts)})")
 
                elif structure == "highlight_comment_red":
                    if comment_text:
                        # "(...)" text shown highlighted red, connected to its
                        # own MD Poling entry (like a highlighter marker).
                        p.add_run(" (")
                        run_extra = p.add_run(comment_text)
                        run_extra.font.highlight_color = WD_COLOR_INDEX.RED
                        p.add_run(")")
                    else:
                        # No "(...)" text for this entry - highlight the MD
                        # Poling text itself in red instead, so it still stands out.
                        run_main.font.highlight_color = WD_COLOR_INDEX.RED
 
                else:  # "plain" and "bold_comment"
                    # Comment shown first, in its own parentheses; quantity
                    # (plus unit, if any) follows as plain text - NOT in
                    # parentheses.
                    if comment_text:
                        p.add_run(f" ({comment_text})")
                    if qty_text:
                        p.add_run(f" {qty_text}")
 
            _style_paragraph(p)
 
    p_spacer = doc.add_paragraph("")  # spacing before the next block
    _style_paragraph(p_spacer)
    return True
 
 
def generate_pole_docx(df, save_path, title="Work Instructions", pole_sequence_map=None):
    """
    pole_sequence_map: optional output of build_pole_sequence_map() - when
    given, poles within each (date, sourcefile) block are split into
    per-map sub-groups and ordered per the map sequence (see
    _group_block_entries_by_map); when omitted, poles keep the same order
    as before (first appearance in the filtered data), with no per-map
    sub-grouping. Distribution into blocks - by date, then sourcefile, i.e.
    Control File / shire / day - is unaffected either way.
    """
    doc = Document()
    _set_document_base_style(doc)
 
    doc.add_heading(title, level=1)
    p_gen = doc.add_paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M}")
    _style_paragraph(p_gen)
    p_total = doc.add_paragraph(f"Total entries: {len(df)}")
    _style_paragraph(p_total)
    p_blank = doc.add_paragraph("")
    _style_paragraph(p_blank)
 
    work = df.copy()
    work["_sourcefile"] = work[COL_SOURCEFILE].apply(_clean)
    work["_pole"] = work[COL_POLE].apply(_clean)
 
    # Header blocks: one per (date, sourcefile) combination - dated rows
    # first in date order (each date's sourcefiles kept in the order they
    # first appear in the data), then any undated rows grouped by sourcefile.
    # Using pandas groupby here (rather than re-filtering the whole frame
    # for every pole/date/sourcefile) keeps this fast even on large inputs.
    dated = work[work[COL_DATE].notna()]
    undated = work[work[COL_DATE].isna()]
 
    if not dated.empty:
        for date_val, date_group in dated.groupby(COL_DATE, sort=True):
            for sourcefile, sf_group in date_group.groupby("_sourcefile", sort=False):
                _emit_pole_block(doc, date_val, sourcefile, sf_group, pole_sequence_map)
 
    if not undated.empty:
        for sourcefile, sf_group in undated.groupby("_sourcefile", sort=False):
            _emit_pole_block(doc, None, sourcefile, sf_group, pole_sequence_map)
 
    doc.save(save_path)
    return save_path
 
 
def build_pole_sequence_map(pole_index):
    """
    Turns the {pole: [location, ...]} pole_index built from the scanned
    Workpack-zone maps (see build_pole_index_from_pdfs) into a
    {pole: (file_key, seq)} map, giving every pole a position "per the map
    sequence, from first to last": file_key clusters poles by which map
    (PDF) they were found in first (so the doc naturally distributes work
    instructions per map), and seq is that location's reading-order position
    within the map (page, then order of appearance on the page - see the
    "seq" field added by _read_pdf_pole_locations/_read_pdf_comments).
    When a pole was found in more than one map/location, its EARLIEST
    (file_key, seq) is used.
    """
    out = {}
    for pole, locations in pole_index.items():
        if not locations:
            continue
        best = min(
            locations,
            key=lambda loc: (loc.get("file", os.path.basename(loc.get("path", ""))), loc.get("seq", 0)),
        )
        file_key = best.get("file", os.path.basename(best.get("path", "")))
        out[pole] = (file_key, best.get("seq", 0))
    return out
 
 
# ---------------------------------------------------------------------------
# PDF comment extraction
# ---------------------------------------------------------------------------
def _read_pdf_comments(entry):
    """Read one PDF's annotation comments (or fall back to per-line plain
    text if it has no annotations). Runs in a worker thread."""
    path = entry["path"]
    folder = entry.get("folder", "")
    day_outage = entry.get("day_outage", "")
    fname = os.path.basename(path)
 
    comments = []
    doc = fitz.open(path)
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            page_had_annots = False
            local_seq = 0
            for annot in page.annots() or []:
                info = annot.info or {}
                content = (info.get("content") or "").strip()
                if content:
                    page_had_annots = True
                    comments.append({
                        "folder": folder, "day_outage": day_outage,
                        "file": fname, "path": path,
                        "page": page_index + 1, "text": content,
                        "seq": page_index * 100000 + local_seq,
                    })
                    local_seq += 1
            if not page_had_annots:
                # Fallback: treat non-empty lines of page text as
                # pseudo-comments so plain drawing text is still searched.
                text = page.get_text() or ""
                for line in text.splitlines():
                    line = line.strip()
                    if line:
                        comments.append({
                            "folder": folder, "day_outage": day_outage,
                            "file": fname, "path": path,
                            "page": page_index + 1, "text": line,
                            "seq": page_index * 100000 + local_seq,
                        })
                        local_seq += 1
    finally:
        doc.close()
    return comments
 
 
def extract_pdf_comments(pdf_entries, progress_cb=None, max_workers=PDF_READ_WORKERS):
    """
    pdf_entries: list of {"path": <full path>, "folder": <folder label>,
    "day_outage": <day/outage label>}. "folder" is whatever grouping label
    the caller wants shown in reports - e.g. the name of the outage-
    description folder the PDF's "Workpack zones" folder sits in, or ""
    when a PDF was picked individually (no folder grouping). "day_outage"
    is similarly the outage folder label, or "" when not applicable.
 
    Returns a list of dicts: {"folder": <label>, "day_outage": <label>,
    "file": <basename>, "path": <full path>, "page": <1-based page number>,
    "text": <comment text>, "seq": <reading-order position within the PDF>}
 
    PDFs are read concurrently (they're read from a network share, so this
    is I/O bound and reading several at once is much faster than one at a
    time). Primary source: PDF annotation "comments" (sticky notes /
    markups), read via PyMuPDF. If a PDF has no annotations at all, we
    fall back to splitting that page's plain text into line-based
    pseudo-comments, so poles mentioned in plain drawing text are still
    picked up.
    """
    if fitz is None:
        raise RuntimeError(
            "PyMuPDF (fitz) is not installed. Run: pip install pymupdf"
        )
 
    all_comments = []
    total = len(pdf_entries)
    done = 0
    lock = threading.Lock()
 
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {ex.submit(_read_pdf_comments, e): e for e in pdf_entries}
        for fut in concurrent.futures.as_completed(futures):
            entry = futures[fut]
            fname = os.path.basename(entry["path"])
            try:
                comments = fut.result()
                all_comments.extend(comments)
            except Exception as e:
                if progress_cb:
                    progress_cb(f"Failed to read {fname}: {e}")
            if progress_cb:
                with lock:
                    done += 1
                    progress_cb(f"Read {done}/{total} PDF(s) ({fname})")
    return all_comments
 
 
# ---------------------------------------------------------------------------
# Automatic network-path discovery of "Workpack zones" map folders
# ---------------------------------------------------------------------------
def _resolve_outage_units(folder_path, folder_name, start_date, end_date):
    """
    An outage folder's own name might be a broad range (e.g. "22-08-2026
    to 01-09-2026 - On Generations") while the real work dates live one
    level deeper, in subfolders that are THEMSELVES dated (e.g. a
    "24-08-2026 - ..." folder inside it). When that's the case, each such
    subfolder is treated as its own outage unit - checked against the date
    filter using ITS OWN, more specific date - instead of the broad parent
    range. This recurses, so it also handles a dated folder nested inside
    another dated folder. A folder with no dated subfolders is a leaf unit
    in its own right, using its own start/end date.
 
    Returns a list of (unit_path, day_outage_label, start_date, end_date).
    """
    try:
        children = list(os.scandir(folder_path))
    except OSError:
        return [(folder_path, folder_name, start_date, end_date)]
 
    dated_children = []
    for entry in children:
        if not entry.is_dir():
            continue
        parsed = _parse_outage_folder_dates(entry.name)
        if parsed:
            child_start, child_end, _desc = parsed
            dated_children.append((entry, child_start, child_end))
 
    if not dated_children:
        return [(folder_path, folder_name, start_date, end_date)]
 
    units = []
    for entry, child_start, child_end in dated_children:
        units.extend(_resolve_outage_units(entry.path, entry.name, child_start, child_end))
    return units
 
 
def _scan_outage_for_zone_folders(outage_path, outage_name, outage_start=None, outage_end=None):
    """Recursively finds every "Workpack zones" folder under one outage
    unit. If NONE exist anywhere in it, falls back to using every LEAF
    folder (one with no subdirectories of its own) that directly contains
    at least one PDF - some outages file their maps straight into the
    deepest folder instead of a "Workpack zones" folder. Runs in a worker
    thread - each outage unit is scanned independently so multiple units
    can be scanned in parallel.
 
    outage_start/outage_end (datetime or None) is this unit's own resolved
    date interval (see _resolve_outage_units) - carried along on every
    result so callers can later recognise a pole as genuinely dated even
    when its Control File date falls outside whatever narrower date range
    happens to be set on Tab 1 right now, as long as it's still within the
    outage folder's own interval (see build_cross_reference)."""
    found = []
    leaf_pdf_dirs = []
    for dirpath, dirnames, filenames in os.walk(outage_path):
        matches = [d for d in dirnames if d.strip().lower() == WORKPACK_ZONES_NAME]
        for d in matches:
            wz_path = os.path.join(dirpath, d)
            found.append({
                "path": wz_path,
                "day_outage": outage_name,
                "folder": os.path.basename(dirpath),
                "outage_start": outage_start,
                "outage_end": outage_end,
            })
            # Don't descend into a Workpack zones folder looking for
            # further nested ones inside it.
            dirnames.remove(d)
        if not dirnames and any(fn.lower().endswith(".pdf") for fn in filenames):
            leaf_pdf_dirs.append(dirpath)
 
    if found:
        return found
 
    # No "Workpack zones" folder anywhere in this outage unit - use maps
    # found directly in its deepest (leaf) folders instead.
    return [
        {
            "path": d, "day_outage": outage_name, "folder": os.path.basename(d),
            "outage_start": outage_start, "outage_end": outage_end,
        }
        for d in leaf_pdf_dirs
    ]
 
 
def find_workpack_zone_folders(root, date_from=None, date_to=None, progress_cb=None,
                                max_workers=NETWORK_SCAN_WORKERS):
    """
    Walk the standard network folder structure:
        root / <YYYY> / <MM - Month> / <dated outage folder> / ... / "Workpack zones"
 
    An outage folder's name can be a single day, a same-month day range, or
    a full "DD-MM-YYYY to DD-MM-YYYY" range (see _parse_outage_folder_dates).
    If that folder contains more specifically-dated subfolders, each of
    those becomes its own outage unit using its own date (see
    _resolve_outage_units) instead of the broader parent range.
 
    "Workpack zones" is searched for recursively under each matching outage
    unit (it isn't always at the same depth - sometimes it sits directly
    inside the unit, sometimes a level or two deeper inside a description
    subfolder), so every unit within the date range is walked to find every
    "Workpack zones" folder inside it. If a unit has NO "Workpack zones"
    folder anywhere, its maps are assumed to be filed directly in its
    deepest (leaf) folders instead, and those are used as the map source.
    Outage units are scanned concurrently, since this is a network share
    and the bottleneck is round-trip latency, not local CPU work.
 
    date_from / date_to are datetime objects or None (no bound on that side).
    Both are inclusive, and are compared against each unit's date (a range
    is considered in-range if it overlaps the filter at all, not just its
    first day).
 
    Returns a list of dicts: {"path": <Workpack zones folder path>,
    "day_outage": <outage folder name>, "folder": <name of the folder that
    directly contains this Workpack zones folder>, "outage_start":
    <datetime>, "outage_end": <datetime>} - outage_start/outage_end is this
    unit's own resolved date interval (see _resolve_outage_units), carried
    through so a pole found on this map can later be recognised as dated
    even outside today's narrower Tab 1 filter (see build_cross_reference) -
    and a sorted list of
    every outage folder NAME that was in the date range and scanned (used
    by callers to report which in-range outages had no "Workpack zones"
    folder at all - that's a legitimate "nothing to show" case, not a
    scanning bug, but worth surfacing so it's not mistaken for one).
    """
    if not os.path.isdir(root):
        raise FileNotFoundError(f"Network path not found or not accessible: {root}")
 
    # First pass: cheaply identify which outage folders are in range,
    # using os.scandir (fewer round trips than listdir + separate isdir
    # calls per entry).
    candidate_outages = []
    for year_entry in sorted(os.scandir(root), key=lambda e: e.name):
        if not year_entry.is_dir() or not YEAR_DIR_REGEX.match(year_entry.name):
            continue
        for month_entry in sorted(os.scandir(year_entry.path), key=lambda e: e.name):
            if not month_entry.is_dir() or not MONTH_DIR_REGEX.match(month_entry.name):
                continue
            for outage_entry in sorted(os.scandir(month_entry.path), key=lambda e: e.name):
                if not outage_entry.is_dir():
                    continue
                parsed = _parse_outage_folder_dates(outage_entry.name)
                if not parsed:
                    continue
                top_start, top_end, _desc = parsed
                # Resolve into the actual unit(s) to scan - if this folder
                # contains more specifically-dated subfolders, those (and
                # their own dates) are used instead of the broader range.
                for unit_path, unit_label, unit_start, unit_end in _resolve_outage_units(
                    outage_entry.path, outage_entry.name, top_start, top_end
                ):
                    if date_from is not None and unit_end < date_from:
                        continue
                    if date_to is not None and unit_start > date_to:
                        continue
                    candidate_outages.append((unit_path, unit_label, unit_start, unit_end))
 
    if not candidate_outages:
        return [], []
 
    # Second pass: scan each candidate outage folder for "Workpack zones"
    # folders in parallel.
    results = []
    total = len(candidate_outages)
    done = 0
    lock = threading.Lock()
 
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {
            ex.submit(_scan_outage_for_zone_folders, path, name, start, end): name
            for path, name, start, end in candidate_outages
        }
        for fut in concurrent.futures.as_completed(futures):
            outage_name = futures[fut]
            try:
                results.extend(fut.result())
            except Exception:
                pass
            if progress_cb:
                with lock:
                    done += 1
                    progress_cb(f"Scanned {done}/{total} outage folder(s) ({outage_name})...")
 
    return results, sorted({name for _path, name, _start, _end in candidate_outages})
 
 
def find_workpack_zone_pdfs(root, date_from=None, date_to=None, progress_cb=None):
    """
    Same as find_workpack_zone_folders, but expands each Workpack zones
    folder into the individual PDF files inside it (recursively), returning
    entries in the {"path", "folder", "day_outage"} shape expected by
    extract_pdf_comments.
 
    Returns (entries, zone_folder_count, pdf_count, outages_without_zones) -
    outages_without_zones lists the outage folder names that WERE in the
    date range and scanned, but had no "Workpack zones" folder anywhere
    inside them, so they legitimately contribute zero maps (e.g. an outage
    folder named "... - No DOP Req'd" is expected to have none). This is
    for distinguishing "nothing there" from "the scan missed it".
    """
    zone_folders, scanned_outages = find_workpack_zone_folders(
        root, date_from=date_from, date_to=date_to, progress_cb=progress_cb
    )
    entries = []
    for zf in zone_folders:
        for dirpath, _dirnames, filenames in os.walk(zf["path"]):
            for fn in filenames:
                if fn.lower().endswith(".pdf"):
                    entries.append({
                        "path": os.path.join(dirpath, fn),
                        "folder": zf["folder"],
                        "day_outage": zf["day_outage"],
                        "outage_start": zf.get("outage_start"),
                        "outage_end": zf.get("outage_end"),
                    })
    outages_with_zones = {zf["day_outage"] for zf in zone_folders}
    outages_without_zones = sorted(o for o in scanned_outages if o not in outages_with_zones)
    return entries, len(zone_folders), len(entries), outages_without_zones
 
 
def build_pole_index(comments):
    """Map each 8-digit pole number to the list of comment dicts it appears in."""
    index = {}
    for c in comments:
        for pole in find_poles(c["text"]):
            index.setdefault(pole, []).append(c)
    return index
 
 
def _read_pdf_pole_locations(entry):
    """Reads one PDF and returns a partial pole index (pole -> list of
    location dicts) built entirely within this worker thread - both the
    file read AND the per-annotation/per-line pole matching happen here,
    in parallel across PDFs, instead of reading every PDF first and then
    matching poles afterward in a separate single-threaded pass. Each
    location dict keeps its own text (needed later for the fuzzy
    similarity match against the MD Poling instructions) and its own "seq"
    (reading-order position within the PDF - page, then order of
    appearance on the page), used later to order poles "per the map
    sequence, from first to last" (see build_pole_sequence_map). This is
    exactly equivalent to extract_pdf_comments() + build_pole_index() -
    just fused into one pass."""
    path = entry["path"]
    folder = entry.get("folder", "")
    day_outage = entry.get("day_outage", "")
    outage_start = entry.get("outage_start")
    outage_end = entry.get("outage_end")
    fname = os.path.basename(path)
 
    local_index = {}
 
    def _record(text, page_index, local_seq):
        for pole in find_poles(text):
            local_index.setdefault(pole, []).append({
                "folder": folder, "day_outage": day_outage,
                "file": fname, "path": path,
                "page": page_index + 1, "text": text,
                "seq": page_index * 100000 + local_seq,
                "outage_start": outage_start, "outage_end": outage_end,
            })
 
    doc = fitz.open(path)
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            page_had_annots = False
            local_seq = 0
            for annot in page.annots() or []:
                info = annot.info or {}
                content = (info.get("content") or "").strip()
                if content:
                    page_had_annots = True
                    _record(content, page_index, local_seq)
                    local_seq += 1
            if not page_had_annots:
                text = page.get_text() or ""
                for line in text.splitlines():
                    line = line.strip()
                    if line:
                        _record(line, page_index, local_seq)
                        local_seq += 1
    finally:
        doc.close()
    return local_index
 
 
def build_pole_index_from_pdfs(pdf_entries, progress_cb=None, max_workers=PDF_READ_WORKERS):
    """Reads every PDF and builds the pole -> [locations] index in ONE
    concurrent pass, folding together what extract_pdf_comments() +
    build_pole_index() used to do as two separate steps (read everything
    into memory, then match poles afterward in a single-threaded pass).
    Returns the same {pole: [location_dict, ...]} shape build_pole_index()
    does (each location dict now also carries a "seq" field - see
    _read_pdf_pole_locations)."""
    if fitz is None:
        raise RuntimeError(
            "PyMuPDF (fitz) is not installed. Run: pip install pymupdf"
        )
 
    index = {}
    total = len(pdf_entries)
    done = 0
    lock = threading.Lock()
 
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {ex.submit(_read_pdf_pole_locations, e): e for e in pdf_entries}
        for fut in concurrent.futures.as_completed(futures):
            entry = futures[fut]
            fname = os.path.basename(entry["path"])
            try:
                local_index = fut.result()
                for pole, locs in local_index.items():
                    index.setdefault(pole, []).extend(locs)
            except Exception as e:
                if progress_cb:
                    progress_cb(f"Failed to read {fname}: {e}")
            if progress_cb:
                with lock:
                    done += 1
                    progress_cb(f"Read {done}/{total} PDF(s) ({fname})")
    return index
 
 
# ---------------------------------------------------------------------------
# Cross-reference report
# ---------------------------------------------------------------------------
def build_cross_reference(df, pole_index, threshold=DEFAULT_SIMILARITY_THRESHOLD,
                           master_poles=None, master_pole_dates=None):
    """
    df: the filtered dataframe (pole + MD Poling universe you chose)
    pole_index: output of build_pole_index()
    master_poles: optional set of every pole value present anywhere in the
                  full (unfiltered) master parquet file. When given, extra
                  poles found in the PDFs are also flagged as to whether they
                  exist anywhere in the master file at all.
    master_pole_dates: optional dict of {pole: date string} covering every
                  pole in the full (unfiltered) master parquet file. When
                  given, an extra pole that IS in the master file (see
                  master_poles above) also gets its master-file date carried
                  along, for display in the "Poles without dates on CF" sheet.
 
    Rows whose Pole or MD Poling is blank/NaN are ignored when building
    found_rows/poles_not_found (there's no MD Poling text to match against
    a map). A pole's DATE, however, is tracked from every row that has a
    valid pole + date, regardless of whether MD Poling is blank - that's
    the basis for "has a date in the current filter" everywhere below.
 
    Returns a dict with:
      found_rows        : list of {pole, mdpoling, comment, plan_date,
                           locations, best_match_pct, best_match_text,
                           matched} - "locations" is a list of {day_outage,
                           folder, file, path, page, text, pct, matched},
                           one per PDF comment that mentioned this pole.
                           "comment" is the MD Poling row's own comment
                           column (not the PDF text). "plan_date" is that
                           row's Control File date.
      mdpoling_not_found : rows where the pole WAS found in a PDF, but the
                           MD Poling text was NOT found (similarity < threshold)
                           inside the best-matching comment
      poles_not_found    : poles from df that do not appear in any PDF comment
                           at all - list of {pole, mdpoling, comment,
                           sourcefile, plan_date}
      extra_poles        : list of {pole, locations, in_master, date} for
                           poles found in PDF comments that do NOT have a
                           date in the current filter ("in_master" is
                           True/None when master_poles wasn't supplied, else
                           True/False; "date" is the master-file date when
                           in_master is True and master_pole_dates was
                           supplied, else ""). A pole is NOT considered
                           "missing" here (and so left out of this list)
                           when its master-file date, even though outside
                           the CURRENT filter, falls within the date
                           interval of the outage folder its map was found
                           in (e.g. a map filed under
                           "22-08-2026 to 01-09-2026 - On Generations" and a
                           Control File date of 25-08-2026 anywhere in the
                           master data - see outage_start/outage_end on
                           each pole_index location, set by
                           find_workpack_zone_pdfs when scanning the network).
      presence_summary   : list of {pole, has_filtered_date, in_pdf,
                           present_in_both, date} covering every pole seen
                           in either source - "has_filtered_date" is True
                           when the pole has a valid date in the current
                           filter (independent of MD Poling); "date" is
                           that date (blank for PDF-only poles)
    """
    found_rows = []
    mdpoling_not_found = []
    poles_not_found = []
 
    # --- Poles with a valid date in the CURRENTLY FILTERED data, regardless
    # of whether MD Poling is blank for that row - the basis for the Pole
    # Summary "Date on Control File" column, "Present In Both?", and what
    # "Poles without dates on CF" excludes.
    _tmp = df[[COL_POLE, COL_DATE]].copy()
    _tmp[COL_POLE] = _tmp[COL_POLE].apply(_clean)
    _tmp = _tmp[(_tmp[COL_POLE] != "") & _tmp[COL_DATE].notna()]
    _tmp["_date_str"] = _tmp[COL_DATE].apply(lambda d: d.strftime("%Y-%m-%d"))
    pole_to_date = (
        _tmp.drop_duplicates(subset=[COL_POLE], keep="first")
            .set_index(COL_POLE)["_date_str"]
            .to_dict()
    )
    poles_with_filtered_date = set(pole_to_date.keys())
 
    # --- Poles whose master-file date (regardless of the current Tab 1
    # filter) falls within the date interval of the outage folder a map of
    # theirs was found in - e.g. a map filed under
    # "22-08-2026 to 01-09-2026 - On Generations" legitimately covers any
    # Control File date in that window, so a pole dated 25-08-2026 is
    # genuinely dated even when Tab 1's own filter is narrower (say, just
    # 22-08-2026). Manually-added PDFs (not found via the network scan)
    # carry no outage_start/outage_end, so they never contribute here -
    # this only ever widens recognition, never narrows the strict
    # current-filter check above.
    poles_dated_via_outage_range = set()
    if master_pole_dates:
        for p, locations in pole_index.items():
            if p in poles_with_filtered_date:
                continue
            date_str = master_pole_dates.get(p)
            if not date_str:
                continue
            try:
                master_dt = datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                continue
            for loc in locations:
                o_start, o_end = loc.get("outage_start"), loc.get("outage_end")
                if o_start is not None and o_end is not None and o_start <= master_dt <= o_end:
                    poles_dated_via_outage_range.add(p)
                    break
 
    # Pull the columns needed by the loop below as plain Python lists once,
    # then zip - avoids the per-row pandas Series construction that
    # df.iterrows() does (a real cost on a large filtered Control File).
    pole_vals = df[COL_POLE].tolist()
    mdpoling_vals = df[COL_MDPOLING].tolist()
    comment_vals = df[COL_COMMENT].tolist()
    sourcefile_vals = df[COL_SOURCEFILE].tolist()
    date_strs = [d.strftime("%Y-%m-%d") if pd.notna(d) else "" for d in df[COL_DATE]]
 
    for pole_raw, mdpoling_raw, comment_raw, sourcefile_raw, plan_date in zip(
        pole_vals, mdpoling_vals, comment_vals, sourcefile_vals, date_strs
    ):
        pole = _clean(pole_raw)
        mdpoling = _clean(mdpoling_raw)
        if not pole or not mdpoling:
            continue  # ignore blank pole values and blank/NaN MD Poling entries
        comment = _clean(comment_raw)
 
        locations = pole_index.get(pole, [])
        if not locations:
            sourcefile = _clean(sourcefile_raw)
            poles_not_found.append({
                "pole": pole, "mdpoling": mdpoling, "comment": comment,
                "sourcefile": sourcefile, "plan_date": plan_date,
            })
            continue
 
        scored_locations = []
        best_pct = 0.0
        best_loc_text = ""
        for loc in locations:
            pct = round(similarity(mdpoling, loc["text"]), 1)
            scored_locations.append({
                "day_outage": loc.get("day_outage", ""),
                "folder": loc.get("folder", ""),
                "file": loc["file"],
                "path": loc.get("path", ""),
                "page": loc["page"],
                "text": loc["text"],
                "pct": pct,
                "matched": pct >= threshold,
            })
            if pct > best_pct:
                best_pct = pct
                best_loc_text = loc["text"]
 
        matched = best_pct >= threshold
        record = {
            "pole": pole,
            "mdpoling": mdpoling,
            "comment": comment,
            "plan_date": plan_date,
            "locations": scored_locations,
            "best_match_pct": round(best_pct, 1),
            "best_match_text": best_loc_text,
            "matched": matched,
        }
        found_rows.append(record)
        if not matched:
            mdpoling_not_found.append(record)
 
    extra_poles = []
    for p in sorted(pole_index.keys()):
        if p in poles_with_filtered_date:
            continue  # already has a date in the current filter - not "missing"
        if p in poles_dated_via_outage_range:
            continue  # has a master-file date within its map's outage-folder
                       # date interval, even though outside the current
                       # filter - not genuinely "missing" either
        in_master = True if master_poles is None else (p in master_poles)
        date_val = ""
        if in_master and master_pole_dates:
            date_val = master_pole_dates.get(p, "")
        extra_poles.append({
            "pole": p,
            "locations": [
                {
                    "day_outage": l.get("day_outage", ""),
                    "folder": l.get("folder", ""),
                    "file": l["file"], "path": l.get("path", ""), "page": l["page"], "text": l["text"],
                }
                for l in pole_index[p]
            ],
            "in_master": in_master,
            "date": date_val,
        })
 
    all_poles = poles_with_filtered_date | set(pole_index.keys())
    presence_summary = [
        {
            "pole": p,
            "has_filtered_date": p in poles_with_filtered_date,
            "in_pdf": p in pole_index,
            "present_in_both": p in poles_with_filtered_date and p in pole_index,
            "date": pole_to_date.get(p, ""),
        }
        for p in sorted(all_poles)
    ]
 
    return {
        "found_rows": found_rows,
        "mdpoling_not_found": mdpoling_not_found,
        "poles_not_found": poles_not_found,
        "extra_poles": extra_poles,
        "presence_summary": presence_summary,
        "threshold": threshold,
    }
 
 
def _autosize_columns(ws):
    """Set a reasonable column width based on the longest value in each column."""
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 60)
 
 
def _merge_column_runs(ws, col_idx, start_row, row_count, keys, wrap=False):
    """
    Merge vertically-adjacent cells in a column where `keys` (one entry per
    data row, same order as written) are equal. Only the first cell in each
    run keeps its value; the rest are cleared before merging. `keys` should
    already encode any outer grouping (e.g. include the file/pole so a
    matching value doesn't merge across an unrelated group).
    """
    align = Alignment(vertical="center", wrap_text=wrap)
    i = 0
    while i < row_count:
        j = i
        while j + 1 < row_count and keys[j + 1] == keys[i]:
            j += 1
        top = start_row + i
        bottom = start_row + j
        if bottom > top:
            for r in range(top + 1, bottom + 1):
                ws.cell(row=r, column=col_idx).value = None
            ws.merge_cells(start_row=top, start_column=col_idx, end_row=bottom, end_column=col_idx)
        for r in range(top, bottom + 1):
            ws.cell(row=r, column=col_idx).alignment = align
        i = j + 1
 
 
def _apply_pole_banding(ws, start_row, row_count, group_keys, col_start, col_end,
                         color_a="FFFFFF", color_b="D9D9D9"):
    """
    Fill columns [col_start, col_end] with alternating solid colors, one
    color per contiguous run in `group_keys` (e.g. each pole group gets its
    own color, alternating with the next pole group).
    """
    fill_a = PatternFill(start_color=color_a, end_color=color_a, fill_type="solid")
    fill_b = PatternFill(start_color=color_b, end_color=color_b, fill_type="solid")
    i = 0
    use_a = True
    while i < row_count:
        j = i
        while j + 1 < row_count and group_keys[j + 1] == group_keys[i]:
            j += 1
        fill = fill_a if use_a else fill_b
        for r in range(start_row + i, start_row + j + 1):
            for c in range(col_start, col_end + 1):
                ws.cell(row=r, column=c).fill = fill
        use_a = not use_a
        i = j + 1
 
 
def _map_hyperlink_path(path, network_root, drive_letter):
    """
    Rewrite a raw UNC path to use a mapped drive letter instead, when one
    is configured. Excel/Windows generally resolve a mapped drive letter
    (e.g. Z:\\...) much faster than a raw \\\\server\\share\\... path, so
    this is purely about making the resulting hyperlinks quicker to open -
    it doesn't change which file is linked.
    """
    if not drive_letter or not network_root:
        return path
    if not path.lower().startswith(network_root.lower()):
        return path
    rest = path[len(network_root):]
    letter = drive_letter.rstrip("\\/")
    return letter + rest
 
 
def _apply_hyperlinks(ws, col_idx, start_row, row_count, keys, paths):
    """
    Turn the top cell of each merged run in `col_idx` into a clickable
    hyperlink pointing at the given file path, so clicking the map name
    opens that PDF. Only the top cell of a merged run keeps a value after
    _merge_column_runs, so only it needs the hyperlink/style applied (this
    also keeps the number of hyperlink relationships written to the
    workbook to a minimum, which keeps the file itself fast to open).
    `paths` is aligned with `keys`/row order - one path per data row.
    """
    link_font = Font(color="0563C1", underline="single")
    i = 0
    while i < row_count:
        j = i
        while j + 1 < row_count and keys[j + 1] == keys[i]:
            j += 1
        top = start_row + i
        path = paths[i]
        if path:
            cell = ws.cell(row=top, column=col_idx)
            cell.hyperlink = path
            cell.font = link_font
        i = j + 1
 
 
def generate_crossref_xlsx(report, save_path, network_root=None, drive_letter=None):
    """
    Cross-reference report as an Excel workbook, aggregated by folder, then
    PDF file, then pole. Map name cells are hyperlinked to the actual PDF
    file so clicking one opens the map.
 
    network_root / drive_letter: if both are given, any hyperlink target
    starting with network_root has that prefix swapped for drive_letter
    (e.g. a mapped "Z:") so the links open faster from Excel.
    """
    header_font = Font(bold=True)
 
    wb = openpyxl.Workbook()
 
    # --- Sheet 1: poles found, grouped by Day/Outage, then Folder, then --
    # --- Map, then Pole ----------------------------------------------------
    ws1 = wb.active
    ws1.title = "Control File Vs maps"
    ws1.append([
        "Day/Outage", "Folder", "Map", "Pole", "Control File Instructions",
        "Control File Comment", "Date", "Match %", "Map Instructions",
    ])
    for cell in ws1[1]:
        cell.font = header_font
 
    by_day_folder_file = {}
    for r in report["found_rows"]:
        for loc in r["locations"]:
            key = (loc.get("day_outage", ""), loc.get("folder", ""), loc["file"])
            by_day_folder_file.setdefault(key, []).append({
                "page": loc["page"],
                "pole": r["pole"],
                "mdpoling": r["mdpoling"],
                "comment": r["comment"],
                "plan_date": r.get("plan_date", ""),
                "pct": loc["pct"],
                "text": loc["text"],
                "path": loc.get("path", ""),
            })
 
    rows_data = []
    for day_outage, folder, fname in sorted(by_day_folder_file.keys()):
        entries = sorted(by_day_folder_file[(day_outage, folder, fname)], key=lambda e: (e["pole"], e["page"]))
        for e in entries:
            rows_data.append({"day_outage": day_outage, "folder": folder, "file": fname, **e})
 
    start_row = ws1.max_row + 1
    for e in rows_data:
        ws1.append([
            e["day_outage"], e["folder"], e["file"], e["pole"], e["mdpoling"], e["comment"],
            e["plan_date"], f"{e['pct']}%", e["text"],
        ])
 
    if rows_data:
        day_keys = [e["day_outage"] for e in rows_data]
        folder_keys = [(e["day_outage"], e["folder"]) for e in rows_data]
        file_keys = [(e["day_outage"], e["folder"], e["file"]) for e in rows_data]
        pole_keys = [(e["day_outage"], e["folder"], e["file"], e["pole"]) for e in rows_data]
        comment_keys = [(e["day_outage"], e["folder"], e["file"], e["pole"], e["text"]) for e in rows_data]
        _merge_column_runs(ws1, 1, start_row, len(rows_data), day_keys)
        _merge_column_runs(ws1, 2, start_row, len(rows_data), folder_keys)
        _merge_column_runs(ws1, 3, start_row, len(rows_data), file_keys)
        _merge_column_runs(ws1, 4, start_row, len(rows_data), pole_keys)
        _merge_column_runs(ws1, 9, start_row, len(rows_data), comment_keys, wrap=True)
        # Alternate grey/white shading per pole group, from the Pole column
        # (4) onward only - Day/Outage, Folder and Map (1-3) are left unshaded.
        _apply_pole_banding(ws1, start_row, len(rows_data), pole_keys, col_start=4, col_end=9)
        # Hyperlink each Map cell (top cell of its merged run) to the PDF.
        paths = [_map_hyperlink_path(e.get("path", ""), network_root, drive_letter) for e in rows_data]
        _apply_hyperlinks(ws1, 3, start_row, len(rows_data), file_keys, paths)
        ws1.column_dimensions["I"].width = 50
 
    # --- Sheet 2: pole presence summary (right after sheet 1) --------------
    ws2 = wb.create_sheet("Pole Summary")
    ws2.append(["Pole", "Date on Control File", "In Map", "Present In Both?", "Date"])
    for cell in ws2[1]:
        cell.font = header_font
    for r in report["presence_summary"]:
        ws2.append([
            r["pole"],
            "Yes" if r["has_filtered_date"] else "No",
            "Yes" if r["in_pdf"] else "No",
            "Yes" if r["present_in_both"] else "No",
            r.get("date", ""),
        ])
 
    # --- Sheet 3: poles not found in any map --------------------------------
    ws3 = wb.create_sheet("Poles not found in maps")
    ws3.append(["Pole", "MD Poling", "MD Poling Comment", "Source File", "Plan Date"])
    for cell in ws3[1]:
        cell.font = header_font
    poles_not_found_sorted = sorted(report["poles_not_found"], key=lambda x: x["pole"])
    ws3_start_row = ws3.max_row + 1
    for r in poles_not_found_sorted:
        ws3.append([r["pole"], r["mdpoling"], r["comment"], r["sourcefile"], r["plan_date"]])
    if poles_not_found_sorted:
        pole_keys3 = [r["pole"] for r in poles_not_found_sorted]
        _merge_column_runs(ws3, 1, ws3_start_row, len(poles_not_found_sorted), pole_keys3)
 
    # --- Sheet 4: poles found in the maps that are not in the filtered CF's Pole column -
    ws4 = wb.create_sheet("Poles without dates on CF")
    ws4.append(["Day/Outage", "Folder", "Map", "Pole", "Map Instructions", "In Master Control File?", "Date"])
    for cell in ws4[1]:
        cell.font = header_font
 
    extra_rows = []
    for ep in report["extra_poles"]:
        for loc in ep["locations"]:
            extra_rows.append({
                "day_outage": loc.get("day_outage", ""),
                "folder": loc.get("folder", ""), "file": loc["file"],
                "path": loc.get("path", ""),
                "page": loc["page"], "pole": ep["pole"], "text": loc["text"],
                "in_master": ep["in_master"], "date": ep.get("date", ""),
            })
    extra_rows.sort(key=lambda e: (e["day_outage"], e["folder"], e["file"], e["pole"], e["page"]))
 
    ws4_start_row = ws4.max_row + 1
    for e in extra_rows:
        ws4.append([
            e["day_outage"], e["folder"], e["file"], e["pole"], e["text"],
            "Yes" if e["in_master"] else "No",
            e.get("date", ""),
        ])
 
    if extra_rows:
        day_keys4 = [e["day_outage"] for e in extra_rows]
        folder_keys4 = [(e["day_outage"], e["folder"]) for e in extra_rows]
        file_keys4 = [(e["day_outage"], e["folder"], e["file"]) for e in extra_rows]
        _merge_column_runs(ws4, 1, ws4_start_row, len(extra_rows), day_keys4)
        _merge_column_runs(ws4, 2, ws4_start_row, len(extra_rows), folder_keys4)
        _merge_column_runs(ws4, 3, ws4_start_row, len(extra_rows), file_keys4)
        paths4 = [_map_hyperlink_path(e.get("path", ""), network_root, drive_letter) for e in extra_rows]
        _apply_hyperlinks(ws4, 3, ws4_start_row, len(extra_rows), file_keys4, paths4)
 
    for ws in (ws1, ws2, ws3, ws4):
        _autosize_columns(ws)
    ws1.column_dimensions["I"].width = 50  # keep the wrapped comment column readable
 
    wb.save(save_path)
    return save_path
 

# =====================================================================
# WEB WRAPPER - replaces the tkinter App window (3 tabs). Each function
# below is the body of the matching button handler, unchanged; only the
# widgets/dialogs became function arguments.
# =====================================================================
PARQUET_FILETYPES = [("Parquet files", "*.parquet"), ("All files", "*.*")]
PDF_FILETYPES = [("PDF files", "*.pdf")]
DEFAULT_DOCX_NAME = "Work_Instructions.docx"
DEFAULT_REPORT_NAME = "Pole_PDF_CrossReference_Report.xlsx"

FILTER_FIELDS = [  # (key, label, column) - same 5 boxes as Tab 1
    ("shires", "District", COL_SHIRE),
    ("projects", "Project", COL_PROJECT),
    ("circuits", "Circuit", COL_CIRCUIT),
    ("sourcefiles", "Sourcefile", COL_SOURCEFILE),
    ("poles", "Pole", COL_POLE),
]
PREVIEW_COLUMNS = [COL_DATE, COL_SHIRE, COL_PROJECT, COL_CIRCUIT, COL_SOURCEFILE,
                   COL_POLE, COL_MDPOLING, COL_QSUB, COL_COMMENT]


def parse_date(s):
    """Same as App._parse_date (YYYY-MM-DD, blank = no bound)."""
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"Could not parse date '{s}'. Use YYYY-MM-DD.")


def filter_options(df):
    """Same as App.on_select_parquet - populates the 5 filter lists."""
    return {
        "rows": int(len(df)),
        "fields": [
            {"key": key, "label": label, "values": unique_sorted(df, col)}
            for key, label, col in FILTER_FIELDS
        ],
    }


def filtered(df, filters):
    """Same as App.on_apply_filters."""
    filters = filters or {}
    return apply_filters(
        df,
        shires=filters.get("shires") or [],
        projects=filters.get("projects") or [],
        circuits=filters.get("circuits") or [],
        sourcefiles=filters.get("sourcefiles") or [],
        poles=filters.get("poles") or [],
        date_from=parse_date(filters.get("date_from")),
        date_to=parse_date(filters.get("date_to")),
    )


def preview_rows(df, limit=500):
    """Same as App._refresh_preview."""
    rows = []
    for _, row in df.head(limit).iterrows():
        date_str = row[COL_DATE].strftime("%Y-%m-%d") if pd.notna(row[COL_DATE]) else ""
        rows.append([
            date_str, row[COL_SHIRE], row[COL_PROJECT], row[COL_CIRCUIT],
            row[COL_SOURCEFILE], row[COL_POLE], row[COL_MDPOLING],
            row[COL_QSUB], row[COL_COMMENT],
        ])
    return {"columns": PREVIEW_COLUMNS, "rows": rows}


def entries_from_folder(folder):
    """Same as App.on_add_pdf_folder."""
    found = []
    for root_dir, _dirs, files in os.walk(folder):
        for fn in files:
            if fn.lower().endswith(".pdf"):
                found.append(os.path.join(root_dir, fn))
    if not found:
        raise FileNotFoundError(f"No PDF files were found under:\n{folder}")
    norm_folder = os.path.normpath(folder)
    folder_label = os.path.basename(norm_folder)
    # Day/Outage = the name of the folder that CONTAINS the selected folder
    day_outage_label = os.path.basename(os.path.dirname(norm_folder))
    return [{"path": p, "folder": folder_label, "day_outage": day_outage_label} for p in found]


def entries_from_files(paths):
    """Same as App.on_select_pdfs."""
    out = []
    for p in paths:
        folder_dir = os.path.dirname(p)
        folder_label = os.path.basename(folder_dir)
        day_outage_label = os.path.basename(os.path.dirname(folder_dir))
        out.append({"path": p, "folder": folder_label, "day_outage": day_outage_label})
    return out


def scan_network(root, date_from=None, date_to=None, log=print):
    """Same as App.on_scan_network (after its confirmation prompt)."""
    log("Scanning network path for Workpack zones folders...")
    found_entries, zone_count, pdf_count, no_zone_outages = find_workpack_zone_pdfs(
        root, date_from=date_from, date_to=date_to, progress_cb=log,
    )
    status = f"Found {zone_count} Workpack zones folder(s), {pdf_count} PDF(s) added to the selection."
    if no_zone_outages:
        status += f" ({len(no_zone_outages)} outage folder(s) in range had no Workpack zones folder.)"
    log(status)
    return {"entries": found_entries, "zone_count": zone_count, "pdf_count": pdf_count,
            "no_zone_outages": no_zone_outages, "status": status}


# Shared pole-index cache (same idea as App._get_pole_index): Tab 2 and Tab 3
# re-use one read of the maps when the PDF selection hasn't changed.
_POLE_INDEX_CACHE = {}
_POLE_INDEX_LOCK = threading.Lock()


def get_pole_index(pdf_entries, log=print):
    key = frozenset(e["path"] for e in pdf_entries)
    with _POLE_INDEX_LOCK:
        if key in _POLE_INDEX_CACHE:
            log(f"Using {len(pdf_entries)} previously-read map(s) (cached)...")
            return _POLE_INDEX_CACHE[key]
    pole_index = build_pole_index_from_pdfs(pdf_entries, progress_cb=log)
    with _POLE_INDEX_LOCK:
        if len(_POLE_INDEX_CACHE) >= 8:
            _POLE_INDEX_CACHE.pop(next(iter(_POLE_INDEX_CACHE)))
        _POLE_INDEX_CACHE[key] = pole_index
    return pole_index


def generate_word(filtered_df, save_path, pdf_entries=None, use_map_order=True, log=print):
    """Same as App.on_generate_word."""
    if filtered_df is None or len(filtered_df) == 0:
        raise ValueError("There are no rows to write - check the filters.")
    if not save_path.lower().endswith(".docx"):
        save_path += ".docx"

    pole_sequence_map = None
    if use_map_order and pdf_entries:
        if fitz is None:
            log("⚠️ PyMuPDF is not installed, so maps can't be read to order the poles - "
                "using the original pole ordering. Run: pip install pymupdf")
        else:
            log("Reading maps to determine pole order, please wait...")
            try:
                pole_index = get_pole_index(pdf_entries, log=log)
                pole_sequence_map = build_pole_sequence_map(pole_index)
            except Exception as e:
                traceback.print_exc()
                log(f"⚠️ Maps could not be read, so the document will use the original pole ordering instead: {e}")
                pole_sequence_map = None

    generate_pole_docx(filtered_df, save_path, pole_sequence_map=pole_sequence_map)
    log(f"Saved: {save_path}")
    return [save_path]


def generate_report(full_df, filtered_df, pdf_entries, save_path, threshold=DEFAULT_SIMILARITY_THRESHOLD,
                    network_root="", drive_letter="", log=print):
    """Same as App.on_generate_report."""
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed. Run: pip install pymupdf")
    if filtered_df is None or len(filtered_df) == 0:
        raise ValueError("Load a parquet file and apply filters first (no rows match).")
    if not pdf_entries:
        raise ValueError("Add at least one folder or PDF file first, or use the network scan.")
    threshold = float(threshold)
    if not save_path.lower().endswith(".xlsx"):
        save_path += ".xlsx"

    log("Reading PDFs, please wait...")
    pole_index = get_pole_index(pdf_entries, log=log)
    master_poles = {_clean(v) for v in full_df[COL_POLE]}
    tmp = full_df[[COL_POLE, COL_DATE]].copy()
    tmp[COL_POLE] = tmp[COL_POLE].apply(_clean)
    tmp = tmp[tmp[COL_POLE] != ""]
    tmp["_date_str"] = tmp[COL_DATE].apply(
        lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else ""
    )
    master_pole_dates = (
        tmp.drop_duplicates(subset=[COL_POLE], keep="first")
           .set_index(COL_POLE)["_date_str"]
           .to_dict()
    )
    report = build_cross_reference(
        filtered_df, pole_index, threshold=threshold,
        master_poles=master_poles, master_pole_dates=master_pole_dates,
    )
    generate_crossref_xlsx(report, save_path, network_root=network_root, drive_letter=drive_letter)

    not_in_master = sum(1 for ep in report["extra_poles"] if not ep["in_master"])
    present_both = sum(1 for r in report["presence_summary"] if r["present_in_both"])
    log(f"Report saved: {save_path}")
    log(f"Poles found in maps: {len(report['found_rows'])}")
    log(f"Poles not found in maps: {len(report['poles_not_found'])}")
    log(f"Poles without dates on CF: {len(report['extra_poles'])}")
    log(f"  - of which not in master Control File at all: {not_in_master}")
    log(f"Poles present in both Control File and maps: {present_both}")
    return [save_path]
