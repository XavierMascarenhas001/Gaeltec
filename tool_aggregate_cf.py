"""Tool 1 - Aggregate Control Files (Block1 + PA CONTROL -> Excel + parquet).
Original logic kept verbatim; only the tkinter file dialogs were replaced."""
import pandas as pd
import os
import re
import numpy as np
from dateutil import parser
 
# ---------------- Utility Functions ----------------
def parse_excel_date(x):
    try:
        if pd.isna(x):
            return np.nan
        
        # Excel serial numbers (already correct)
        if isinstance(x, (int, float)):
            return pd.to_datetime(x, origin="1899-12-30", unit="D", errors="coerce")
        
        # US format (month first)
        return pd.to_datetime(str(x), errors="coerce")
    
    except Exception:
        return np.nan
 
def parse_segment_info(segment_str, project_mapping, default_pm=""):
    """
    Parse a segment string into type, description, code, and project manager.
    If no project manager is found, use the default_pm as backup.
    """
    if not isinstance(segment_str, str) or not segment_str.strip():
        return pd.Series([None, None, None, default_pm])
    
    segment_str = segment_str.strip()
    
    # Extract type (M or C)
    type_match = re.match(r"^\s*([MC])\s*-\s*", segment_str)
    segment_type = type_match.group(1) if type_match else None
 
    # Extract segment code (e.g., ABC12345)
    code_match = re.search(r"\b([A-Z]{0,3}\d{5})\b", segment_str)
    segment_code = code_match.group(1) if code_match else None
 
    # Remove type and code from description
    segment_desc = re.sub(r"^\s*[MC]\s*-\s*", "", segment_str)
    if segment_code:
        segment_desc = re.sub(re.escape(segment_code), "", segment_desc)
    segment_desc = re.sub(r"\s+", " ", segment_desc).strip(" -")
 
    # --- Try to detect Project Manager from project_mapping ---
    project_manager = ""
    text_lower = segment_str.lower()
    for pm, (pm_shire, pm_project) in project_mapping.items():
        if pm.lower() in text_lower:
            project_manager = pm
            segment_desc = re.sub(re.escape(pm), "", segment_desc, flags=re.IGNORECASE)
            break
 
    # --- Backup logic: use default_pm if no match found ---
    if not project_manager and default_pm:
        project_manager = default_pm
 
    segment_desc = re.sub(r"\s+", " ", segment_desc).strip(" -")
    return pd.Series([segment_type, segment_desc, segment_code, project_manager])
    
def extract_project_shire(filename):
    filename_lower = filename.lower()
 
    for key, value in file_project_mapping.items():
        if key.lower() in filename_lower:
            shire, project = value  # unpack list
            return project, shire
 
    return "", ""
 
def extract_location(segment_desc):
    if not isinstance(segment_desc, str):
        return ""
    segment_desc_lower = segment_desc.lower()
    for key, locations in mapping_region.items():
        if key.lower() in segment_desc_lower:
            return ", ".join(locations)
    return ""
 
# --- Normalize and map multiple codes ---
def map_teams(codes):
    if pd.isna(codes):
        return "UNKNOWN"
    
    codes = str(codes).upper().strip()
    
    # Handle empty strings and "nan"
    if codes == "" or codes == "NAN":
        return "UNKNOWN"
    
    names = []
    for char in codes:
        if char in teams:
            names.extend(teams[char])
    
    return ", ".join(names) if names else "UNKNOWN"
    
# ---------------- Mapping Dictionaries ----------------
# --- Project Manager Mapping ---
project_mapping = {
    "Jonathon Mcclung": ["Ayrshire", "PCB"],
    "Gary MacDonald": ["Ayrshire", "LV"],
    "Jim Gaffney": ["Lanark", "PCB"],
    "Calum Thomson": ["Ayrshire", "Connections"],
    "Calum Thompson": ["Ayrshire", "Connections"],
    "Calum Thomsom": ["Ayrshire", "Connections"],
    "David Jamieson": ["Lanark", "11kV"],
    "Andrew Galt": ["Ayrshire", "-"],
    "Henry Gordon": ["Ayrshire", "-"],
    "Jack Murray": ["Ayrshire", "Connections"],
    "Jonathan Douglas": ["Ayrshire", "11kV"],
    "Jonathon Douglas": ["Ayrshire", "11kV"],
    "Jordan Graham": ["Lanark", "-"],
    "Matt": ["Lanark", "-"],
    "Lee Fraser": ["Ayrshire", "Connections"],
    "Lee Frazer": ["Ayrshire", "Connections"],
    "Mark": ["Lanark", "Connections"],
    "Mark Nicholls": ["Ayrshire", "Connections"],
    "Cameron Fleming": ["Lanark", "Connections"],
    "Cameron Flemming": ["Lanark", "Connections"],
    "Ronnie Goodwin": ["Lanark", "Connections"],
    "Ian Young": ["Ayrshire", "Connections"],
    "Iain Cassidy": ["Ayrshire", "Connections"],
    "Matthew Watson": ["Lanark", "Connections"],
    "Martin Maxwell": ["Ayrshire", "Connections"],
    "Aileen Brese": ["Ayrshire", "Connections"],
    "Mark McGoldrick": ["Lanark", "Connections"],
    "Rachel Plant": ["Ayrshire", "Connections"]
}
 
# --- Pole / Equipment / Conductor Mappings ---
mapping = {
    "9x220 BIOCIDE LV POLE": "9m B",
    "9x275 BIOCIDE LV POLE": "9s B",
    "9x220 CREOSOTE LV POLE": "9m",
    "9x275 CREOSOTE LV POLE": "9s",
    "9x220 HV SINGLE POLE": "9m",
    "9x275 HV SINGLE POLE": "9s",
    "9x295 HV SINGLE POLE": "9es",
    "9x315 HV SINGLE POLE": "9esp",
    "10x230 BIOCIDE LV POLE": "10m B",
    "10x230 HV SINGLE POLE": "10m",
    "10x285 BIOCIDE LV POLE": "10s B",
    "10x285 H POLE HV Creosote": "10s",
    "10x285 HV SINGLE POLE": "10s",
    "10x305 HV SINGLE POLE": "10es",
    "11x295 HV SINGLE POLE": "11s",
    "11x295 H POLE HV Creosote": "11s",
    "11x295 BIOCIDE LV POLE": "11sB",
    "12x250 BIOCIDE LV POLE": "12m B",
    "12x305 BIOCIDE LV POLE": "12s B",
    "12x250 CREOSOTE LV POLE": "12m",
    "12x305 CREOSOTE LV POLE": "12s",
    "12x305 H POLE HV Creosote":"12s",
    "12x250 HV SINGLE POLE": "12m",
    "12x305 HV SINGLE POLE": "12s",
    "12x325 HV SINGLE POLE": "12es",
    "12x345 HV SINGLE POLE": "12esp",
    "13x260 BIOCIDE LV POLE": "13m B",
    "13x320 BIOCIDE LV POLE": "13s B",
    "13x260 CREOSOTE LV POLE": "13m",
    "13x320 CREOSOTE LV POLE": "13s",
    "13x260 HV SINGLE POLE": "13m",
    "13x320 HV SINGLE POLE": "13s",
    "13x340 HV SINGLE POLE": "13es",
    "13x365 HV SINGLE POLE": "13esp",
    "14x275 BIOCIDE LV POLE": "14m B",
    "14x335 BIOCIDE LV POLE": "14s B",
    "14x275 CREOSOTE LV POLE": "14m",
    "14x335 CREOSOTE LV POLE": "14s",
    "14x275 HV SINGLE POLE": "14m",
    "14x335 HV SINGLE POLE": "14s",
    "14x355 HV SINGLE POLE": "14es",
    "14x375 HV SINGLE POLE": "14esp",
    "16x305 BIOCIDE LV POLE": "16m B",
    "16x365 BIOCIDE LV POLE": "16s B",
    "16x305 CREOSOTE LV POLE": "16m",
    "16x365 CREOSOTE LV POLE": "16s",
    "16x305 HV SINGLE POLE": "16m",
    "16x365 HV SINGLE POLE": "16s",
    "16x385 HV SINGLE POLE": "16es",
    "16x405 HV SINGLE POLE": "16esp",
    "11x315 H POLE HV Creosote":"11es",
    "14x335 H POLE HV Creosote":"14s",
    "11x315 HV SINGLE POLE":"11es",
    "13x320 H POLE HV Creosote":"13s",
    "11x240 CREOSOTE LV POLE":"11",
    "11x240 HV SINGLE POLE":"11m",
    "10x230 CREOSOTE LV POLE":"10m",
    "11x335 H POLE HV Creosote":"11esp",
    "10x305 H POLE HV Creosote":"10es",
    "11x240 BIOCIDE LV POLE":"11m B",
    "16x365 H POLE HV Creosote":"16s",
    "16x405 EHV SINGLE POLE CREOSOTE":"16esp",
    "14x355 H Delta HVY SP4147830":"14es",
    "14x355 H POLE HV Creosote":"14es",
    "12x325 H POLE HV Creosote":"12es",
    "16x385 H POLE HV Creosote":"16es",
    "12x305 EHV SINGLE POLE CREOSOTE":"12s",
    "13x340 EHV SINGLE POLE CREOSOTE":"13es",
    "11x335 EHV SINGLE POLE CREOSOTE":"11es",
    "11x315 EHV SINGLE POLE CREOSOTE":"11es",
    "12x325 EHV SINGLE POLE CREOSOTE":"12es",
 
    # AAAC bare conductors
    "Hazel - 50mm² AAAC bare (1000m drums)": "Hazel 50mm² (1000m drums)",
    "Oak - 100mm² AAAC bare (1000m drums)": "Oak 100mm² (1000m drums)",
    "Ash - 150mm² AAAC bare (1000m drums)": "Ash 150mm² (1000m drums)",
    "Poplar - 200mm² AAAC bare (1000m drums)": "Poplar 200mm² (1000m drums)",
    "Upas - 300mm² AAAC bare (1000m drums)": "Upas 300mm² (1000m drums)",
    "Poplar OPPC - 200mm² AAAC equivalent bare": "Poplar OPPC 200mm²",
    "Upas OPPC - 300mm² AAAC equivalent bare": "Upas OPPC 300mm²",
 
    # ACSR conductors
    "Gopher - 25mm² ACSR bare (1000m drums)": "Gopher 25mm² (1000m drums)",
    "Caton - 25mm² Compacted ACSR bare (1000m drums)": "Caton 25mm² (1000m drums)",
    "Rabbit - 50mm² ACSR bare (1000m drums)": "Rabbit 50mm² (1000m drums)",
    "Wolf - 150mm² ACSR bare (1000m drums)": "Wolf 150mm² (1000m drums)",
    "Horse - 70mm² ACSR bare": "Horse 70mm²",
    "Dog - 100mm² ACSR bare (1000m drums)": "Dog 100mm² (1000m drums)",
    "Dingo - 150mm² ACSR bare (1000m drums)": "Dingo 150mm² (1000m drums)",
 
    # Copper conductors
    "Hard Drawn Copper 16mm² ( 3/2.65mm ) (500m drums)": "Copper 16mm² (500m drums)",
    "Hard Drawn Copper 32mm² ( 3/3.75mm ) (1000m drums)": "Copper 32mm² (500m drums)",
    "Hard Drawn Copper 70mm² (500m drums)": "Copper 70mm² (500m drums)",
    "Hard Drawn Copper 100mm² (500m drums)": "Copper 100mm² (500m drums)",
 
    # PVC covered copper
    "35mm² Copper (Green / Yellow PVC covered) (50m drums)": "Copper 35mm² GY PVC (50m drums)",
    "70mm² Copper (Green / Yellow PVC covered) (50m drums)": "Copper 70mm² GY PVC (50m drums)",
    "35mm² Copper (Blue PVC covered) (50m drums)": "Copper 35mm² Blue PVC (50m drums)",
    "70mm² Copper (Blue PVC covered) (50m drums)": "Copper 70mm² Blue PVC (50m drums)",
 
    # Double insulated cables
    "35mm² Double Insulated (Brown) (50m drums)": "Double Insulated 35mm² Brown (50m drums)",
    "35mm² Double Insulated (Blue) (50m drums)": "Double Insulated 35mm² Blue (50m drums)",
    "70mm² Double Insulated (Brown) (50m drums)": "Double Insulated 70mm² Brown (50m drums)",
    "70mm² Double Insulated (Blue) (50m drums)": "Double Insulated 70mm² Blue (50m drums)",
    "120mm² Double Insulated (Brown) (50m drums)": "Double Insulated 120mm² Brown (50m drums)",
    "120mm² Double Insulated (Blue) (50m drums)": "Double Insulated 120mm² Blue (50m drums)",
 
    # LV cables
    "LV Cable 1ph 4mm Concentric (250m drums)": "LV 1ph 4mm Concentric (250m drums)",
    "LV Cable 1ph 25mm CNE (250m drums)": "LV 1ph 25mm CNE (250m drums)",
    "LV Cable 1ph 25mm SNE (100m drums)": "LV 1ph 25mm SNE (100m drums)",
    "LV Cable 1ph 35mm CNE (250m drums)": "LV 1ph 35mm CNE (250m drums)",
    "LV Cable 1ph 35mm SNE (100m drums)": "LV 1ph 35mm SNE (100m drums)",
    "LV Cable 3ph 35mm Cu Split Con (250m drums)": "LV 3ph 35mm Cu Split Con (250m drums)",
    "LV Cable 3ph 35mm SNE (250m drums)": "LV 3ph 35mm SNE (250m drums)",
    "LV Cable 3ph 35mm CNE (250m drums)": "LV 3ph 35mm CNE (250m drums)",
    "LV Cable 3ph 35mm CNE Al (LSOH) (250m drums)": "LV 3ph 35mm CNE Al LSOH (250m drums)",
    "LV Cable 3c 95mm W/F (250m drums)": "LV 3c 95mm W/F (250m drums)",
    "LV Cable 3c 185mm W/F (250m drums)": "LV 3c 185mm W/F (250m drums)",
    "LV Cable 3c 300mm W/F (250m drums)": "LV 3c 300mm W/F (250m drums)",
    "LV Cable 4c 95mm W/F (250m drums)": "LV 4c 95mm W/F (250m drums)",
    "LV Cable 4c 185mm W/F (250m drums)": "LV 4c 185mm W/F (250m drums)",
    "LV Cable 4c 240mm W/F (250m drums)": "LV 4c 240mm W/F (250m drums)",
    "LV Marker Tape (365m roll)": "LV Marker Tape (365m roll)",
 
    # 11kV cables
    "11kv Cable 95mm 3c Poly (250m drums)": "11kV 3c 95mm Poly (250m drums)",
    "11kv Cable 185mm 3c Poly (250m drums)": "11kV 3c 185mm Poly (250m drums)",
    "11kv Cable 300mm 3c Poly (250m drums)": "11kV 3c 300mm Poly (250m drums)",
    "11kv Cable 95mm 1c Poly (250m drums)": "11kV 1c 95mm Poly (250m drums)",
    "11kv Cable 185mm 1c Poly (250m drums)": "11kV 1c 185mm Poly (250m drums)",
    "11kv Cable 300mm 1c Poly (250m drums)": "11kV 1c 300mm Poly (250m drums)",
    "11kV Marker Tape (40m roll)": "11kV Marker Tape (40m roll)",
 
    # --- Transformer & Steelwork mappings ---
    "Transformer 1ph 50kVA": "TX 1ph (50kVA)",
    "Transformer 3ph 50kVA": "TX 3ph (50kVA)",
    "Transformer 1ph 100kVA": "TX 1ph (100kVA)",
    "Transformer 1ph 25kVA": "TX 1ph (25kVA)",
    "Transformer 3ph 200kVA": "TX 3ph (200kVA)",
    "Transformer 3ph 100kVA": "TX 3ph (100kVA)",
 
    "Erect Single HV/EHV Pole, up to and including 12 metre pole":"Erect HV pole", 
    "Erect LV Structure Single Pole, up to and including 12 metre pole" :"Erect LV pole",
    "Erect Single HV/EHV Pole, up to and including 12 metre pole.":"Erect HV pole",
    "Erect Section Structure 'H' HV/EHV Pole, up to and including 12 metre pole.":"Erect H HV pole",
 
    "Plumb single pole":"Plumb pole",
    "Recover single pole, up to and including 15 metres in height, and reinstate, all ground conditions":"Recover single pole",
    "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions":"Recover H pole"
}
 
# Create mapping dict for 'Mapped' column
mapping_dict = mapping.copy()
 
# --- REGION MAPPING ---
file_project_mapping = {
 
    # ---------- AYRSHIRE ----------
    "pcb 2022": ["Ayrshire", "PCB"],
    "33kv Refurb 2021": ["Ayrshire", "33kV Refurb"],
    "Connections 2023": ["Ayrshire", "Connections"],
    "Aurs Road 40222": ["Ayrshire", "Aurs Road"],
    "Storms _2023": ["Ayrshire", "Storms"],
    "11kV Refurb 2023": ["Ayrshire", "11kV Refurb"],
    "SPEN Labour Provider": ["Ayrshire", "SPEN Labour"],
 
    # Duplicate 2023 refurb set
    "11kV Refurb 2023_2": ["Ayrshire", "11kV Refurb"],
 
    # 2024 sets
    "Connections 2024": ["Ayrshire", "Connections"],
    "PCB 2024": ["Ayrshire", "PCB"],
    "LVHi5_4 2024": ["Ayrshire", "LV"],
    "11kV Refurb 2024": ["Ayrshire", "11kV Refurb"],
    "Lanark 2024": ["Lanark", "Lanark"],   # ambiguous name but file is Ayrshire region
    "11kV Refurb Lethanhill 2024": ["Ayrshire", "11kV Refurb"],
 
    # 2025 sets
    "Connections 2025": ["Ayrshire", "Connections"],
    "LV Ayrshire 2025": ["Ayrshire", "LV"],
    "PCB 2025 Ayrshire": ["Ayrshire", "PCB"],
    "11kV Refurb Ayrshire": ["Ayrshire", "11kV Refurb"],
    "11kV Ref Ayr Pinwherry": ["Ayrshire", "11kV Refurb"],
    "Storms _2025": ["Scotland", "Storms"],
    "Storms _2025 New": ["Scotland", "Storms"],
    "Connections _2025 New": ["Ayrshire", "Connections"],
    "LV & ESQCR Lanark 2025New": ["Lanark", "LV"],   # belongs in Ayrshire dataset
    "PCB 2025 Ayrshire NEW": ["Ayrshire", "PCB"],
    "11kv Refurb Ayrshire NEW": ["Ayrshire", "11kV Refurb"],
    "11kV Refurb Ayrshire 2026": ["Ayrshire", "11kV Refurb"],
    "11kV Refurb Ayrshire Pinwherry": ["Ayrshire", "11kV Refurb"],
    "LV Ayrshire 2025 new": ["Ayrshire", "LV"],
    "33kV Ayrshire 2025": ["Ayrshire", "33kV Refurb"],
    "Hi5_4_Ayrshire_2026": ["Ayrshire", "11kV Refurb"],
    "Connections_2026": ["Ayrshire", "Connections"],
 
 
    # ---------- LANARK ----------
    "Lanark 2025_11kv Refurb": ["Lanark", "11kV Refurb"],
    "Lanark 2025_Connections": ["Lanark", "Connections"],
    "Lanark 2025_PCB": ["Lanark", "PCB"],
    "LV & ESQCR Lanark 2025": ["Lanark", "LV"],
    "Lanark 2025_Connections NEW": ["Lanark", "Connections"],
    "Lanark 2025_PCB NEW": ["Lanark", "PCB"],
    "Lanark 2025_11kV Refur NEW": ["Lanark", "11kV Refurb"],
    "Hi5_4_Lanark_2026": ["Lanark", "11kV Refurb"],
    "LV Lanark 2026New": ["Lanark", "LV"],
    "Lanark _11kv refur 2026 NEW": ["Lanark", "11kV Refurb"],
    "Lanark _33kv refur 2026 NEW": ["Lanark", "33kV Refurb"],
    "Reinforcement_Lanark":["Lanark", "Reinforcement"],   # (comma added - was missing)
 
 
    # ---------- Glasgow----------
 
    "Glasgow 2026_11kV": ["Glasgow", "11kV Refurb"],
}
 
# --- REGION MAPPING ---
mapping_region = {
    "Newmilns": ["Irvine Valley"],
    "New Cumnock": ["New Cumnock"],
    "Kilwinning": ["Kilwinning"],
    "Stewarton": ["Irvine Valley"],
    "Kilbirnie": ["Kilbirnie and Beith"],
    "Coylton": ["Ayr East"],
    "Irvine": ["Irvine Valley", "Irvine East", "Irvine West"],
    "TROON": ["Troon"],
    "Ayr": ["Ayr East", "Ayr North", "Ayr West"],
    "Maybole": ["Maybole, North Carrick and Coylton"],
    "Clerkland": ["Irvine Valley"],
    "Glengarnock": ["Kilbirnie and Beith"]
}
 
# --- TEAM MAPPING ---
teams = {
    "A": ["Paulo Marques"],
    "B": ["Rui Rocha"],
    "C": ["Craig Kerr"],
    "D": ["Robert Urie"],
    "E": ["Alistair Mcpherson"],
    "F": ["Kenny Campbell"],
    "G": ["Pele"],
    "S": ["Sub contracted"],
}
 
# ---------------- Final output column spec ----------------
# Maps the internal (lowercased) column name -> the final renamed column,
# and defines the exact column order for the Aggregated sheet / parquet output.
FINAL_COLUMN_RENAME = {
    "segment": "job",
    "pole": "Enid",
    "item": "description",
    "qty": "qty",
    "qcvi": "qcvi",
    "qsub": "qsub",
    "itemtype": "Type",
    "comment": "Comment",
    "plan1": "plan1",
    "plan2": "plan2",
    "done": "done",
    "datetouse": "datetouse",
    "date": "Invoice date",
    "total": "total",
    "orig": "orig",
    "shire": "district",
    "project": "project",
    "projectmanager": "project manager",
    "segmentcode": "circuit",
    "region": "region",
    "team": "team",
    "team_name": "team lider",
    "cat": "CAT",
    "defect code": "Defect Code",
    "sourcefile": "sourcefile",
}
 
FINAL_COLUMN_ORDER = [
    "job", "Enid", "description", "qty", "qcvi", "qsub", "Type", "Comment",
    "plan1", "plan2", "done", "datetouse", "Invoice date", "total", "orig",
    "district", "project", "project manager", "circuit", "region",
    "team", "team lider", "CAT", "Defect Code", "sourcefile",
]
 
 
def finalize_columns(df):
    """Rename to the requested output names, keep only the requested
    columns, and force them into the requested order. Any column not
    listed in FINAL_COLUMN_RENAME is dropped."""
    df = df.rename(columns=FINAL_COLUMN_RENAME)
    for col in FINAL_COLUMN_ORDER:
        if col not in df.columns:
            df[col] = np.nan
    return df[FINAL_COLUMN_ORDER]
 

# =====================================================================
# WEB WRAPPER - replaces the tkinter file dialogs. The processing below
# is the original main-script body, unchanged, with print() -> log().
#   file_paths  : list of Excel files (was askopenfilenames)
#   output_file : the .xlsx to save      (was asksaveasfilename)
# =====================================================================
FILETYPES = [("Excel files", "*.xlsx *.xlsm *.xls *.xlsb")]


def run(file_paths, output_file, log=print):
    outputs = []
    if not file_paths:
        log("❌ No files selected. Exiting.")
        return outputs

    aggregated_df = pd.DataFrame()
    resume_list_files = []  # for simple [file_name, project, shire] info
    resume_list_dfs = []    # for actual DataFrames like PA CONTROL

    # Added T (Type), W (Plan2) and AC (CAT) to the read range.
    BLOCK1_USECOLS_FULL = "A,B,C,D,E,F,T,U,V,W,AC,AL,AM,AO,CG,CH,CA"   # CA = Defect Code
    BLOCK1_USECOLS_FALLBACK = "A,B,C,D,E,F,T,U,V,W,AC,AL,AM,AO,CG,CH"  # used if CA is missing

    for file_path in file_paths:
        ext = os.path.splitext(file_path)[1].lower()
        file_name = os.path.basename(file_path)
        log(f"\n📘 Reading file: {file_name}")

        # --- Detect project + shire from filename ---
        project, shire = extract_project_shire(file_name)
        log(f" → Project: {project} | Shire: {shire}")

        # Store simple file info
        resume_list_files.append([file_name, project, shire])

        # --- BLOCK1 ---
        try:
            had_defect_code_col = True
            try:
                read_kwargs = dict(sheet_name="Block1", header=2, skiprows=range(3, 29),
                                   usecols=BLOCK1_USECOLS_FULL)
                if ext == ".xlsb":
                    read_kwargs["engine"] = "pyxlsb"
                df = pd.read_excel(file_path, **read_kwargs)
            except Exception:
                # Column CA (Defect Code) may not exist in this file — retry without it
                had_defect_code_col = False
                read_kwargs = dict(sheet_name="Block1", header=2, skiprows=range(3, 29),
                                   usecols=BLOCK1_USECOLS_FALLBACK)
                if ext == ".xlsb":
                    read_kwargs["engine"] = "pyxlsb"
                df = pd.read_excel(file_path, **read_kwargs)

            df.columns = df.columns.str.strip().str.lower()
            df.columns = df.columns.str.strip()

            if not had_defect_code_col:
                log("   ⚠️ 'Defect Code' (column CA) not found — left blank for this file")
                df["defect code"] = np.nan

            if "type" in df.columns:
                df = df.rename(columns={"type": "itemtype"})
            else:
                df["itemtype"] = np.nan

            # --- Drop rows where column A is empty or 'Stop' ---
            col_a_name = df.columns[0]  # Column A
            df = df[~df[col_a_name].str.lower().isin(['stop']) & df[col_a_name].notna()]
            # --- Parse dates ---
            df['plan1'] = df['plan1'].apply(parse_excel_date)
            if 'plan2' in df.columns:
                df['plan2'] = df['plan2'].apply(parse_excel_date)
            df['done'] = df['done'].apply(parse_excel_date)
            df['datetouse'] = df['done'].combine_first(df['plan1'])
            date_cols = [c for c in df.columns if c.startswith('date')]
            for col in date_cols:
                df[col] = df[col].apply(parse_excel_date)

            # Parse segment info
            if 'segment' in df.columns:
                df[['type', 'segmentdesc', 'segmentcode', 'projectmanager']] = df['segment'].apply(
                    lambda x: parse_segment_info(x, project_mapping)
                )

            # Add project info
            df['project'] = project
            df['shire'] = shire
            df['location'] = df['segmentdesc'].apply(extract_location)
            df['region'] = df['location'].where(df['location'].notna() & (df['location'] != ""), df['shire'])
            df['sourcefile'] = file_name
            #Add the teams to the output
            team_map = {k: v[0] for k, v in teams.items()}   # simplify mapping

            if 'team' in df.columns:
                df['team'] = df['team'].astype(str).str.strip().str.upper()
                df['team_name'] = df['team'].apply(map_teams)

            if 'item' in df.columns:
                for col in ['qty', 'qsub']:
                    if col in df.columns:
                        df.loc[df['item'].str.contains('H POLE', na=False), col] *= 2

            aggregated_df = pd.concat([aggregated_df, df], ignore_index=True)
            log(f"   ✅ 'Block1' loaded — {len(df)} rows")
        except Exception as e:
            log(f"   ❌ Error reading 'Block1': {e}")

        # --- PA CONTROL ---
        try:
            pa_kwargs = dict(sheet_name="PA CONTROL", header=0, skiprows=1, usecols=[0,2,4])
            if ext == ".xlsb":
                pa_kwargs["engine"] = "pyxlsb"

            pa_df = pd.read_excel(file_path, **pa_kwargs)
            pa_df.columns = ["section", "value_eur", "completion"]

            # Keep only MC sections
            pa_df = pa_df[pa_df['section'].astype(str).str.match(r'^[MC]', na=False)]

            # Convert numeric columns
            pa_df['value_eur'] = pd.to_numeric(pa_df['value_eur'], errors='coerce')
            pa_df['completion'] = pd.to_numeric(pa_df['completion'], errors='coerce')
            pa_df['%complete'] = (pa_df['completion'] / pa_df['value_eur'].replace(0, np.nan)) * 100

            # Parse segment info
            pa_df[['type', 'segmentdesc', 'segmentcode', 'projectmanager']] = pa_df['section'].apply(
                lambda x: parse_segment_info(x, project_mapping)
            )

            # Add project info
            pa_df['project'] = project
            pa_df['shire'] = shire
            pa_df['location'] = pa_df['segmentdesc'].apply(extract_location)
            pa_df['region'] = pa_df['location'].where(pa_df['location'].notna() & (pa_df['location'] != ""), pa_df['shire'])
            pa_df['sourcefile'] = file_name

            # Append to list
            resume_list_dfs.append(pa_df)
            log(f"   ✅ 'PA CONTROL' loaded — {len(pa_df)} rows")
        except Exception as e:
            log(f"   ⚠️ Could not read 'PA CONTROL': {e}")

    # --- SAVE OUTPUT ---
    if aggregated_df.empty:
        log("⚠️ No valid data found. Exiting.")
        return outputs

    aggregated_df = aggregated_df.sort_values(by='datetouse').reset_index(drop=True)

    # --- Reduce to the requested columns, renamed, in the requested order ---
    aggregated_df = finalize_columns(aggregated_df)

    if not output_file:
        log("❌ No output file selected. Exiting.")
        return outputs

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        aggregated_df.to_excel(writer, index=False, sheet_name="Aggregated")

        # Concatenate PA CONTROL DataFrames safely
        if resume_list_dfs:
            resume_df = pd.concat(resume_list_dfs, ignore_index=True)
            resume_df.to_excel(writer, index=False, sheet_name="Resume")
    log(f"✅ Excel saved: {output_file}")
    outputs.append(output_file)

    # --- Save Aggregated Parquet ---
    aggregated_parquet_file = os.path.splitext(output_file)[0] + "_aggregated.parquet"
    agg_df_copy = aggregated_df.copy()
    for col in agg_df_copy.select_dtypes(include=['object']).columns:
        agg_df_copy[col] = agg_df_copy[col].astype(str)
    agg_df_copy.to_parquet(aggregated_parquet_file, index=False)
    log(f"✅ Aggregated Parquet saved: {aggregated_parquet_file}")
    outputs.append(aggregated_parquet_file)

    # --- Save Resume Parquet ---
    if resume_list_dfs:
        resume_parquet_file = os.path.splitext(output_file)[0] + "_resume.parquet"
        resume_df_copy = resume_df.copy()
        for col in resume_df_copy.select_dtypes(include=['object']).columns:
            resume_df_copy[col] = resume_df_copy[col].astype(str)
        resume_df_copy.to_parquet(resume_parquet_file, index=False)
        log(f"✅ Resume Parquet saved: {resume_parquet_file}")
        outputs.append(resume_parquet_file)

    return outputs
