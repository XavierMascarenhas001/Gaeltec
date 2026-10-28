"""Tool 4 - Materials / CV Excel Export (was the "Excel Export Tool" window).
Original logic kept verbatim; only the tkinter window was replaced."""
import pandas as pd
import numpy as np
import io
import re
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import Outline
import threading
 
 
# -------------------------------
# 🔴 ADD YOUR MAPPINGS HERE
# -------------------------------
 
# --- Project Manager Mapping ---
project_mapping = {
    "Jonathon Mcclung": ["Ayrshire", "PCB"],
    "Gary MacDonald": ["Ayrshire", "LV"],
    "Jim Gaffney": ["Lanark", "PCB"],
    "Calum Thomson": ["Ayrshire", "Connections"],
    "Calum Thomsom": ["Ayrshire", "Connections"],
    "Calum Thompson": ["Ayrshire", "Connections"],
    "Andrew Galt": ["Ayrshire", "-"],
    "Henry Gordon": ["Ayrshire", "-"],
    "Jonathan Douglas": ["Ayrshire", "11 kV"],
    "Jonathon Douglas": ["Ayrshire", "11 kV"],
    "Matt": ["Lanark", ""],
    "Lee Fraser": ["Ayrshire", "Connections"],
    "Lee Frazer": ["Ayrshire", "Connections"],
    "Mark": ["Lanark", "Connections"],
    "Mark Nicholls": ["Ayrshire", "Connections"],
    "Cameron Fleming": ["Lanark", "Connections"],
    "Ronnie Goodwin": ["Lanark", "Connections"],
    "Ian Young": ["Ayrshire", "Connections"],
    "Matthew Watson": ["Lanark", "Connections"],
    "Aileen Brese": ["Ayrshire", "Connections"],
    "Mark McGoldrick": ["Lanark", "Connections"]
}
 
# --- Region Mapping ---
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
    "Glengarnock": ["Kilbirnie and Beith"],
    "Ayrshire": ["North Coast and Cumbraes","Prestwick", "Saltcoats and Stevenston", "Troon", "Ayr East", "Ayr North",
                 "Ayr West","Annick","Ardrossan and Arran","Dalry and West Kilbride","Girvan and South Carrick","Irvine East",
                 "Irvine Valley","Irvine West","Kilbirnie and Beith","Kilmarnock East and Hurlford","Kilmarnock North",
                 "Kilmarnock South","Kilmarnock West and Crosshouse","Kilwinning","Kyle","Maybole, North Carrick and Coylton",
                 "Ayr, Carrick and Cumnock","East_Ayrshire","North_Ayrshre","South_Ayrshre","Doon Valley"],
    "Lanark": ["Abronhill, Kildrum and the Village","Airdrie Central","Airdrie North","Airdrie South","Avondale and Stonehouse",
               "Ballochmyle","Bellshill","Blantyre","Bothwell and Uddingston","Cambuslang East","Cambuslang West",
               "Clydesdale East","Clydesdale North","Clydesdale South","Clydesdale West","Coatbridge North and Glenboig",
               "Coatbridge South","Coatbridge West","Cumbernauld North","Cumbernauld South",
               "East Kilbride Central North","East Kilbride Central South","East Kilbride East","East Kilbride South",
               "East Kilbride West","Fortissat","Hamilton North and East","Hamilton South","Hamilton West and Earnock",
               "Mossend and Holytown","Motherwell North","Motherwell South East and Ravenscraig","Motherwell West",
               "Rutherglen Central and North","Rutherglen South","Strathkelvin","Thorniewood","Wishaw","Larkhall",
               "Airdrie and Shotts","Cumbernauld, Kilsyth and Kirkintilloch East","East Kilbride, Strathaven and Lesmahagow",
               "Lanark and Hamilton East","Motherwell and Wishaw","North_Lanarkshire","South_Lanarkshire"]
}
 
# --- File Project Mapping ---
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
    "Storms _2025": ["Scotland", "Storms"],
    "Storms _2025 New": ["Scotland", "Storms"],
    "Connections _2025 New": ["Ayrshire", "Connections"],
    "LV & ESQCR Lanark 2025New": ["Lanark", "LV"],   # belongs in Ayrshire dataset
    "PCB 2025 Ayrshire NEW": ["Ayrshire", "PCB"],
    "11kv Refurb Ayrshire NEW": ["Ayrshire", "11kV Refurb"],
    "11kV Refurb Ayrshire 2026": ["Ayrshire", "11kV Refurb"],
    "11kV Ref Ayr Pinwherry": ["Ayrshire", "11kV Refurb"],
    "LV Ayrshire 2025 new": ["Ayrshire", "LV"],
    "33kV Ayrshire 2025": ["Ayrshire", "33kV Refurb"],
    "Hi5_4_Ayrshire_2026": ["Ayrshire", "11kV Refurb"],
 
 
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
 
    # ---------- Glasgow----------
 
    "Glasgow 2026_11kV": ["Glasgow", "11kV Refurb"],
}
 
CV7_erect = {
    "Erect Single HV/EHV Pole, up to and including 12 metre pole":"CV7 HV pole", 
    "Erect Single HV/EHV Pole, up to and including 12 metre pole.":"CV7  HV pole",
}
 
CV7_erect_H = {
    "Erect Section Structure 'H' HV/EHV Pole, up to and including 12 metre pole.":"CV7 HV pole"
}
 
CV7_erect_lv = {
    "Erect LV Structure Single Pole, up to and including 12 metre pole" :"CV7 LV pole",
}
 
CV7_recover = {
    "Recover single pole, up to and including 15 metres in height, and reinstate, all ground conditions":"CV7",
    "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions":"CV7  HV pole"
}
 
 
# --- Transformer Mappings ---
CV7_Tx = {
    "Erect pole mounted transformer up to 100kVA 1.ph.": "CV7 Tx",
    "Erect pole mounted transformer up to 200kVA 3.p.h.": "CV7 Tx",
    "Erect Voltage Regulator.": "CV7 Tx",
    "Erect Voltage Transformer (VT), RTU or Repeater": "CV7 Tx",
    "Erect 12kV/36kV Surge arrestors ( directly mounted ).": "CV7 Tx)",
    "Remove pole mounted tranformer.": "CV7 Tx)",
    "Remove platform mounted or 'H' pole mounted transformer.": "CV7 Tx)"
}
 
# --- Transformer Mappings ---
transformer = {
    "Transformer 1ph 50kVA": "TX 1ph (50kVA)",
    "Transformer 3ph 50kVA": "TX 3ph (50kVA)",
    "Transformer 1ph 100kVA": "TX 1ph (100kVA)",
    "Transformer 1ph 25kVA": "TX 1ph (25kVA)",
    "Transformer 3ph 200kVA": "TX 3ph (200kVA)",
    "Transformer 3ph 100kVA": "TX 3ph (100kVA)"
}
 
# --- Equipment / Conductor Mappings ---
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
 
    # LV cables per meter
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
 
 
CV7_SWITCHGEAR = {
    "Erect 11kV/33kV ABSW": "CV7 SWITCHGEAR",
    "Erect 11kV Remote Controlled Switch Disconnector ( Soule Auguste ) or Auto Reclosure unit c/w VT, Aerial, RTU & umbilical cable.": "CV7 SWITCHGEAR",
    "Erect 1.ph fuse units at single tee off pole or in line pole.": "CV7 SWITCHGEAR",
    "Erect 3.ph fuse units at single tee off pole or in line pole.": "CV7 SWITCHGEAR",
    "Additional cost for fitting fuse outrigger bracket.": "CV7 SWITCHGEAR",
    "Remove 11kV/33kV ABSW": "CV7 SWITCHGEAR",
}
 
CV7_UG = {
    "Installation of cable only in trench dug by others; 11kV Cable 3 x 1 core.": "CV7 UG 11 kV",
    "Install cable in existing duct; 11kV Cable 3 x 1 core.": "CV7 UG 11 kV",
    "Installation of cable only in trench dug by others; 33kV Cable 3 x 1 core.": "CV7 UG 33 kV",
    "Install cable in existing duct; 33kV Cable 3 x 1 core.": "CV7 UG 33 kV",
    "Installation of cable only in trench dug by others; LV Cable Large or 11kV Cable 1 x 3 Core": "CV7 UG",
    "Install cable in existing duct; LV Cable Large or 11kV Cable 1 x 3 Core": "CV7 UG",
    "Installation of cable only in trench dug by others; LV Service, Small LV or Pilot Cable.": "CV7 UG LV Service",
    "Install cable in existing duct; LV Service, Small LV or Pilot Cable.": "CV7 UG LV Service",
}
 
CV7_CB = {
    "Remove Auto Reclosure.": "CV7 CB",
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
    "Remove Steelwork crossarm item only": "CV8",
    "Convert 1.ph 11kV Intermediate pole into Section Pole.": "CV8",
    "Convert 1.ph/3.p.h. 11kV line pole into Terminal Pole.": "CV8",
    "Convert 3.ph 11kV Intermediate pole into Section Pole.": "CV8",
    "Change 11kV Insulators to avoid contamination from old conductor": "CV8",
    "Change 33kV Insulators to avoid contamination from old conductor": "CV8",
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
 
summary_items = [
    "Erect Single HV/EHV Pole, up to and including 12 metre pole.",
    "Erect Section Structure 'H' HV/EHV Pole, up to and including 12 metre pole",
    "Erect LV Structure Single Pole, up to and including 12 metre pole",
    "Recover single pole, up to and including 15 metres in height, and reinstate, all ground conditions",
    "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions",
    "Erect 11kV/33kV ABSW.",
    "Remove 11kV/33kV ABSW",
    "Noja"
    "0.5 kVa Tx for Noja"
    "11kV PMSW (Soule)"
    "Remove Auto Reclosure",
    "Erect pole mounted transformer up to 100kVA 1.ph",
    "Erect pole mounted transformer up to 200kVA 3.p.h",
    "Remove pole mounted transformer",
    "Remove platform mounted or 'H' pole mounted transformer",
    "Install bare conductor, run out, sag, terminate, bind in and connect jumpers; <100mm²",
    "Install bare conductor, run out, sag, terminate, bind in and connect jumpers; >=100mm² <200mm²",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 2c + Earth",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 4c + Earth",
    "Install service span including connection to mainline & building / structure",
    "Erect 3.ph fuse units at single tee off pole or in line pole"
    "Remove 1.ph or 3.ph HV fuses",    
]
 
categories = [
    ("CV7_erect", CV7_erect, "Quantity"),
    ("CV7_erect_H", CV7_erect_H, "Quantity"),
    ("CV7_erect_LV", CV7_erect_lv, "Quantity"),
    ("CV7_recover", CV7_recover, "Quantity"),
    ("CV7 Tx", CV7_Tx, "Quantity"),
    ("transformer", transformer, "Quantity"),
    ("CV7 OHL CONDUCTOR_instal", CV7_OHL_CONDUCTOR_instal, "Length (Km)"),
    ("CV7 OHL CONDUCTOR_recover", CV7_OHL_CONDUCTOR_recover, "Length (m)"),
    ("CV7 OHL CONDUCTOR LV_instal", CV7_OHL_CONDUCTOR_LV_instal, "Length (Km)"),
    ("CV7 OHL CONDUCTOR LV_recover", CV7_OHL_CONDUCTOR_LV_recover, "Length (m)"),
    ("CV7 SWITCHGEAR", CV7_SWITCHGEAR, "Quantity"),
    ("CV7_UG", CV7_UG, "Quantity"),
    ("CV7_CB", CV7_CB, "Quantity"),
    ("Switch", Switch, "Quantity"),
    ("Fuses", Fuses, "Quantity"),
    ("CV8", CV8, "Quantity"),
    ("CV31", CV31, "Quantity"),
]
 
column_rename_map = {
    "mapped": "Output",
    "segmentcode": "Circuit",
    "datetouse_display": "Date",
    "qty": "Quantity_original",
    "qcvi":"Variations",
    "qsub": "Quantity_used",
    "segmentdesc": "location",
    "shire": "District",
    "pid_ohl_nr": "PID",
    "projectmanager": "Project Manager"
}
 
export_columns = [
    'Output','comment', 'item', 'Quantity_original','Variations','Quantity_used', 'pole', 'Date',
    'District', 'project', 'Project Manager', 'Circuit', 'Segment',
    'team lider', 'PID','total', 'orig', 'sourcefile'
]
 
 
categories = [
    ("CV7_erect", CV7_erect, "Quantity"),
    ("CV7_erect_H", CV7_erect_H, "Quantity"),
    ("CV7_erect_LV", CV7_erect_lv, "Quantity"),
    ("CV7_recover", CV7_recover, "Quantity"),
    ("CV8", CV8, "Quantity"),
    ("CV31", CV31, "Quantity"),
]
 
extra_categories = [
    ("CV7 Tx", CV7_Tx),
    ("transformer", transformer),
    ("CV7 OHL CONDUCTOR_instal", CV7_OHL_CONDUCTOR_instal),
    ("CV7 OHL CONDUCTOR_recover", CV7_OHL_CONDUCTOR_recover),
    ("CV7 OHL CONDUCTOR LV_instal", CV7_OHL_CONDUCTOR_LV_instal),
    ("CV7 OHL CONDUCTOR LV_recover", CV7_OHL_CONDUCTOR_LV_recover),
    ("Switch", Switch),
    ("Fuses", Fuses),
]
 
 
# =========================================================
# 📌 SHEET SCHEMA (INTERNAL ONLY - DO NOT RENAME THESE)
# =========================================================
 
 
sheet_columns = [
    "segment", "pole", "item", "comment",
    "quantity_original", "variations", "quantity_used",
    "plan1", "done", "shire", "project",
    "segmentcode", "segmentdesc"
]
 
sheet_columns_lower = [c.lower() for c in sheet_columns]
 
# =========================================================
# 🔗 NEW MASTER-FILE -> INTERNAL SCHEMA BRIDGE
# =========================================================
# The aggregation / build-master pipeline renames several source
# columns before saving the master parquet. We translate them back
# to the old internal names immediately on load so everything below
# stays unchanged.
NEW_TO_OLD_COLUMNS = {
    "job": "segment",
    "enid": "pole",
    "description": "item",
    "qvci": "qcvi",
    "district": "shire",
    "circuit": "segmentcode",
    "project manager": "projectmanager",
    "pid": "pid_ohl_nr",
}
 
# 🔥 Columns already present in every "normal" (non-Summary) sheet's
# export by default. Used so that when the user picks extra date
# columns to filter on (e.g. "plan2"), we only add the ones that
# aren't already part of the base layout.
BASE_EXPORT_COLUMNS = {
    "segment", "enid", "item", "comment",
    "quantity_original", "variations", "quantity_used", "type",
    "plan1", "done", "shire", "project",
    "circuit", "job", "sourcefile"
}
 
# 🔥 Sheets that get the money columns "total" + "orig" right after
# "done" (and after any extra date-filter columns).
SHEETS_WITH_MONEY_COLUMNS = {"Outputs", "Variations"}
 
IMG_LEFT = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker\Images\GaeltecImage.png"
IMG_RIGHT = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker\Images\SPEN.png"
 
# 🔥 Service codes lookup file (description -> service code).
# Column_1 = description (same text as the master file's "description"),
# Column_3 = service code. ⚠️ Change this path to wherever the file lives.
SERVICE_CODES_PATH = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\28.Project Tracker\miscelaneous.parquet"
 
# Header used for the service code column in the exported sheets
SERVICE_CODE_COL = "service code"
 
# Name of the optional per-service-code variations sheet
SERVICE_CODE_SHEET = "Variations_service codes"
 
# 💷 Money format: negatives shown in red with a minus sign
MONEY_FORMAT = '£#,##0.00;[Red]-£#,##0.00'
 
 
# =========================================================
# 🔧 EXPORT RENAMING (ONLY USED AT FINAL STEP)
# =========================================================
def prepare_for_excel(df):
    df = df.copy()
 
    rename_map = {
        "pole": "enid",
        "segmentcode": "circuit",
        "segmentdesc": "project_name"
    }
 
    return df.rename(columns=rename_map)
 
 
# =========================================================
# 📅 DATE FORMATTING (SAFE)
# =========================================================
def format_excel_dates(df):
    df = df.copy()
 
    date_like_cols = {"plan1", "plan2", "done"}
 
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%d-%B-%y")
 
        elif col in date_like_cols or "date" in col.lower():
            df[col] = pd.to_datetime(df[col], errors="coerce")
            df[col] = df[col].dt.strftime("%d-%B-%y")
 
    return df
 
 
# =========================================================
# 🔧 NORMALISE FUNCTIONS
# =========================================================
def normalize_pole(p):
    if pd.isna(p):
        return ""
    s = str(p)
    s = (
        s.replace("\u200b", "")
         .replace("\u200e", "")
         .replace("\u200f", "")
         .replace("\xa0", "")
         .strip()
         .upper()
    )
    return re.sub(r"\s+", "", s)
 
 
def normalize_item(x):
    if pd.isna(x):
        return ""
    s = str(x)
    s = (
        s.replace("\u200b", "")
         .replace("\u200e", "")
         .replace("\u200f", "")
         .replace("\xa0", "")
         .strip()
         .upper()
    )
    return re.sub(r"\s+", " ", s)
 
 
def normalise_mapping_keys(mapping):
    return {normalize_item(k): v for k, v in mapping.items()}
 
 
def normalize_cv_name(s):
    """Normalize a category/cv_group name for comparison: strips spaces
    and underscores and lower-cases, so 'CV7 Tx', 'CV7_Tx', 'cv7tx' all match."""
    return re.sub(r"[\s_]+", "", str(s)).strip().lower()
 
 
def most_common_value(series):
    """Return the most frequently occurring non-blank value in a
    Series, or '' if there isn't one."""
    s = series.dropna()
    s = s[s.astype(str).str.strip() != ""]
    if s.empty:
        return ""
    return s.mode().iat[0]


# Text that should be treated as "no value" anywhere it shows up
BLANK_TEXT = {"", "nan", "none", "nat", "null", "<na>"}


def most_common_comment(series, min_words=2):
    """Most frequent comment, ignoring blanks / 'nan' / 'None' and any
    comment with fewer than `min_words` words. Returns '' if nothing
    qualifies."""
    s = series.dropna().astype(str).str.strip()
    s = s[~s.str.lower().isin(BLANK_TEXT)]
    s = s[s.str.split().str.len() >= min_words]
    if s.empty:
        return ""
    return s.mode().iat[0]


def blank_out_nans(df):
    """Replace NaN / 'nan' / 'None' style values with None so they come
    through as empty cells in Excel."""
    df = df.astype(object).where(pd.notna(df), None)
    for col in df.columns:
        df[col] = df[col].map(
            lambda v: None if isinstance(v, str) and v.strip().lower() in BLANK_TEXT else v
        )
    return df


# =========================================================
# 🏷️ SERVICE CODES
# =========================================================
# Status of the service-code lookup for each row (shown on the
# Variations_service codes sheet for rows that got no code)
SC_NO_CODE = "(no code in miscelaneous file)"
SC_NOT_FOUND = "(description not in miscelaneous file)"
 
# Filled in by generate_excel(): how many rows got a code, and which
# descriptions didn't (for the log / message after an export)
LAST_SERVICE_CODE_STATS = {}
 
_PUNCT_TRANS = str.maketrans({
    "\u2018": "'", "\u2019": "'", "\u00b4": "'", "`": "'",
    "\u201c": '"', "\u201d": '"',
    "\u2013": "-", "\u2014": "-", "\u2212": "-",
})
 
 
def loose_item_key(x):
    """A looser version of normalize_item, used only when the exact text
    isn't found: ignores case, spacing around punctuation/brackets, a
    trailing full stop, curly quotes, dashes and characters like '²'.
    So 'Replace tension insulator, 11kV' finds 'Replace tension insulator, 11kV.'
    and 'Cable Termination ( incorporating surge arrestors )' finds
    'Cable Termination (incorporating surge arrestors)'."""
    s = normalize_item(x)
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s).translate(_PUNCT_TRANS).upper()
    s = re.sub(r"\s*([,;:/()\[\]\-.'+&])\s*", r"\1", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip(" .,;:")
 
 
def load_service_codes(path=None):
    """Reads the service-codes file and returns a dict
    {normalised description: service code}. Descriptions come from
    Column_1 (or the first column) and codes from Column_3 (or a column
    with 'service' in its name). Descriptions whose code is blank or '0'
    are kept with code "" so the report can tell "no code in the file"
    apart from "description not in the file"."""
    path = path or SERVICE_CODES_PATH
    lower = str(path).lower()
    if lower.endswith(".parquet"):
        codes = pd.read_parquet(path)
    elif lower.endswith(".csv"):
        codes = pd.read_csv(path, dtype=str)
    else:
        codes = pd.read_excel(path, dtype=str)
 
    cols = list(codes.columns)
    desc_col = "Column_1" if "Column_1" in cols else cols[0]
    code_col = "Column_3" if "Column_3" in cols else next(
        (c for c in cols if "service" in str(c).lower()), cols[2] if len(cols) > 2 else cols[-1]
    )
 
    out = {}
    for d, c in zip(codes[desc_col], codes[code_col]):
        d = normalize_item(d)
        if not d:
            continue
        c = "" if c is None or (isinstance(c, float) and pd.isna(c)) else str(c).strip()
        c = re.sub(r"\.0$", "", c)                          # 30980981.0 -> 30980981
        if c.lower() in BLANK_TEXT or c == "0":
            c = ""
        if out.get(d):                                     # same description twice:
            continue                                       # keep the first real code
        out[d] = c
    return out
 
 
def service_code_summary():
    """One short text about the last export's service codes."""
    st = LAST_SERVICE_CODE_STATS
    if not st:
        return ("Service codes: the miscelaneous file wasn't loaded, so the 'service code' column is blank.\n"
                f"(Looked for: {SERVICE_CODES_PATH})")
    lines = [f"Service codes: {st['with_code']:,} of {st['rows']:,} rows have a code."]
    if st["no_code_in_file"]:
        lines.append(f"{st['no_code_in_file']:,} rows are items whose code is 0 / blank in the miscelaneous file.")
    if st["not_in_file"]:
        lines.append(f"{st['not_in_file']:,} rows have a description that isn't in the miscelaneous file, e.g.:")
        lines += [f"   - {d}" for d in st["not_in_file_items"][:8]]
    return "\n".join(lines)
 
 
def match_service_codes(item_norms, service_codes):
    """For each normalised description -> (code or None, status).
    1st try: the exact text (same matching as Build Master's Man_day).
    2nd try: loose_item_key - only used when every description in the
    file with that key has the same code, so a code is never guessed."""
    loose = {}
    for d, c in service_codes.items():
        loose.setdefault(loose_item_key(d), set()).add(c)
 
    cache = {}
    codes, status = [], []
    for it in item_norms:
        if it not in cache:
            if it in service_codes:
                c = service_codes[it]
            else:
                cands = loose.get(loose_item_key(it)) if it else None
                c = next(iter(cands)) if cands and len(cands) == 1 else None
            if c is None:
                cache[it] = (None, SC_NOT_FOUND)
            elif c == "":
                cache[it] = (None, SC_NO_CODE)
            else:
                cache[it] = (c, "")
        c, st = cache[it]
        codes.append(c)
        status.append(st)
    return codes, status
 
 
# =========================================================
# 🧹 CLEAN DATAFRAME
# =========================================================
def clean_dataframe(df):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    df = df.loc[:, ~df.columns.duplicated()]
    return df
 
 
# =========================================================
# 📏 AUTO COLUMN WIDTH (shared helper)
# =========================================================
def autofit_columns(ws, start_row=2, max_width=40):
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_length = 0
        for row in range(start_row, ws.max_row + 1):
            val = ws.cell(row=row, column=col_idx).value
            if val is not None:
                max_length = max(max_length, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max_length + 4, max_width)
 
 
# =========================================================
# 📤 WRITE SHEET (FINAL EXPORT LAYER)
# =========================================================
def write_sheet(wb, name, df, agg_mode="sum", extra_date_columns=None, header_row=2):
    """Writes one sheet and returns the worksheet.
    header_row lets a caller (the Variations sheet) push the table
    further down to leave room for something above it."""
 
    ws = wb.create_sheet(name[:31])
 
    if df is None or df.empty:
        ws.append(["No Data"])
        return ws
 
    df = prepare_for_excel(df)
    df = format_excel_dates(df)
    df = blank_out_nans(df)
 
    # =========================
    # IMAGES
    # =========================
    try:
        img1 = Image(IMG_LEFT)
        img2 = Image(IMG_RIGHT)
 
        img1.width = 110
        img1.height = 75
 
        img2.width = 140
        img2.height = 60
 
        ws.add_image(img1, "A1")
        ws.add_image(img2, "B1")
 
        ws.row_dimensions[1].height = 60
 
    except Exception as e:
        print(f"Image load failed: {e}")
 
    # =========================
    # STYLES
    # =========================
    HEADER_COLOR = "00CCFF"
    header_fill = PatternFill(start_color=HEADER_COLOR, end_color=HEADER_COLOR, fill_type="solid")
 
    header_font = Font(bold=True, size=15)
    normal_font = Font(size=11)
    medium_font = Font(size=13)
    project_font = Font(bold=True, size=14)
    big_font = Font(bold=True, size=14)
 
    center_align = Alignment(horizontal="center")
 
    # 💷 CURRENCY FORMAT
    currency_format = '£#,##0.00'
 
    thick_bottom = Border(
        bottom=Side(style="medium")
    )
 
    # =========================
    # SUMMARY SHEET (pivot style: District > project > circuit)
    # =========================
    # Expects a hidden "_level" column built in generate_excel:
    #   0 = district row, 1 = project row, 2 = circuit row, 3 = grand total
    # Project and circuit rows are Excel outline groups, so the [+]/[-]
    # buttons expand/collapse them like a pivot table.
    if name.lower() == "summary":
 
        levels = list(df["_level"]) if "_level" in df.columns else [2] * len(df)
        df = df.drop(columns=["_level"], errors="ignore")
 
        district_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        grand_fill = PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")
        money_cols = {"total", "orig", "Variations"}
        indent_cols = {"project_name": 2, "circuit": 4}
 
        ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=True)
 
        # HEADER
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
 
        # DATA
        for r_idx, (row, level) in enumerate(zip(df.itertuples(index=False), levels), header_row + 1):
            for c_idx, value in enumerate(row, 1):
                col_name = df.columns[c_idx - 1]
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
 
                if level == 0:
                    cell.font = Font(bold=True, size=16, color="FFFFFF")
                    cell.fill = district_fill
                elif level == 1:
                    cell.font = project_font
                    cell.fill = header_fill
                elif level == 3:
                    cell.font = big_font
                    cell.fill = grand_fill
                    cell.border = thick_bottom
                else:
                    cell.font = normal_font
                    if col_name in indent_cols:
                        cell.alignment = Alignment(indent=indent_cols[col_name])
 
                if col_name in money_cols:
                    cell.number_format = MONEY_FORMAT
 
            # outline: projects under districts, circuits under projects
            if level == 1:
                ws.row_dimensions[r_idx].outlineLevel = 1
            elif level == 2:
                ws.row_dimensions[r_idx].outlineLevel = 2
                ws.row_dimensions[r_idx].hidden = True   # circuits start collapsed
 
            # mark project rows collapsed when their circuits are hidden
            if level == 1:
                ws.row_dimensions[r_idx].collapsed = True
 
    # =========================
    # NORMAL SHEETS
    # =========================
    else:
 
        export_columns = [
            "segment", "enid", "item", "comment",
            "quantity_original", "variations", "quantity_used", "type",
            "plan1", "done", "shire", "project",
            "circuit", "job", 'sourcefile'
        ]
 
        # Extra date-filter columns (e.g. "plan2") go right after "done"
        if extra_date_columns:
            done_idx = export_columns.index("done")
            insert_at = done_idx + 1
            for col in extra_date_columns:
                if col not in export_columns:
                    export_columns.insert(insert_at, col)
                    insert_at += 1
 
        # 🔥 Outputs + Variations sheets: insert "total" and "orig"
        # right after "done" and any extra date-filter columns.
        if name in SHEETS_WITH_MONEY_COLUMNS:
            done_idx = export_columns.index("done")
            insert_pos = done_idx + 1
            while insert_pos < len(export_columns) and export_columns[insert_pos] in (extra_date_columns or []):
                insert_pos += 1
            export_columns = export_columns[:insert_pos] + ["total", "orig"] + export_columns[insert_pos:]
 
        # 🔥 "service code": right after "orig" on Outputs/Variations; on
        # the other sheets (no money columns) it sits in the same place,
        # right after "done" and any extra date-filter columns.
        if "orig" in export_columns:
            sc_pos = export_columns.index("orig") + 1
        else:
            sc_pos = export_columns.index("done") + 1
            while sc_pos < len(export_columns) and export_columns[sc_pos] in (extra_date_columns or []):
                sc_pos += 1
        export_columns.insert(sc_pos, SERVICE_CODE_COL)
 
        for c in export_columns:
            if c not in df.columns:
                df[c] = ""
 
        df = df[export_columns]
 
        for col_idx, col_name in enumerate(export_columns, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
 
        for r_idx, row in enumerate(df.itertuples(index=False), header_row + 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if export_columns[c_idx - 1] in ("total", "orig"):
                    cell.number_format = MONEY_FORMAT
 
    # =========================
    # AUTO COLUMN WIDTH
    # =========================
    autofit_columns(ws, start_row=header_row)
 
    return ws
 
 
# =========================================================
# 📊 VARIATIONS SHEET (pivot-style block + detail rows)
# =========================================================
def build_variations_pivot(df):
    """Aggregate the variations != 0 rows into:
       - job_stats: one row per (sourcefile, job) with Orig, Total,
         Variation (Total - Orig) and the most common comment (Reason)
       - src_stats: Orig, Total, Variation per sourcefile"""
    p = df.copy()
 
    def clean_text(col):
        if col not in p.columns:
            return pd.Series("(blank)", index=p.index)
        s = p[col].fillna("").astype(str).str.strip()
        return s.where(~s.str.lower().isin(BLANK_TEXT), "(blank)")
 
    p["_src"] = clean_text("sourcefile")
    p["_job"] = clean_text("job")
    if "comment" not in p.columns:
        p["comment"] = ""
    for c in ("orig", "total"):
        p[c] = pd.to_numeric(p[c], errors="coerce").fillna(0) if c in p.columns else 0.0
 
    job_stats = (
        p.groupby(["_src", "_job"], as_index=False)
         .agg(orig=("orig", "sum"), total=("total", "sum"),
              reason=("comment", most_common_comment))
         .sort_values(["_src", "_job"])
    )
    job_stats["variation"] = job_stats["total"] - job_stats["orig"]
 
    src_stats = p.groupby("_src")[["orig", "total"]].sum().sort_index()
    src_stats["variation"] = src_stats["total"] - src_stats["orig"]
 
    return job_stats, src_stats
 
 
def write_pivot_block(ws, job_stats, src_stats, start_row):
    """Writes a pivot-style table in columns A:E starting at start_row:
    Job | Orig | Total | Variation | Reason.
    Each source file is a bold row with its own Orig/Total/Variation;
    its jobs sit underneath as a collapsed Excel outline group, so the
    [+] next to a source file expands it into its jobs. Returns the
    last row used."""
 
    header_fill = PatternFill(start_color="00CCFF", end_color="00CCFF", fill_type="solid")
    src_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    grand_fill = PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")
    thick_bottom = Border(bottom=Side(style="medium"))
    money_cols = (2, 3, 4)            # Orig, Total, Variation
    n_cols = 5
 
    # Put the [+]/[-] button on the source-file row (above its jobs)
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=True)
 
    headers = ["Job", "Orig", "Total", "Variation", "Reason"]
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = Font(bold=True, size=15)
        cell.alignment = Alignment(horizontal="center")
 
    def write_money(row, values, font=None):
        for col, val in zip(money_cols, values):
            cell = ws.cell(row=row, column=col, value=float(val))
            cell.number_format = MONEY_FORMAT
            if font is not None:
                cell.font = font
 
    r = start_row + 1
    for src, st in src_stats.iterrows():
        # --- source file row (always visible) ---
        src_row = r
        ws.cell(row=src_row, column=1, value=src).font = Font(bold=True, size=12)
        write_money(src_row, (st["orig"], st["total"], st["variation"]), Font(bold=True, size=12))
        for c in range(1, n_cols + 1):
            ws.cell(row=src_row, column=c).fill = src_fill
        r += 1
 
        # --- job rows (grouped + collapsed) ---
        jobs = job_stats[job_stats["_src"] == src]
        first_job_row = r
        for _, jr in jobs.iterrows():
            ws.cell(row=r, column=1, value=jr["_job"]).alignment = Alignment(indent=2)
            write_money(r, (jr["orig"], jr["total"], jr["variation"]))
            ws.cell(row=r, column=5, value=jr["reason"] or None)
            r += 1
        last_job_row = r - 1
 
        if last_job_row >= first_job_row:
            ws.row_dimensions.group(first_job_row, last_job_row, outline_level=1, hidden=True)
            ws.row_dimensions[src_row].collapsed = True
 
    # --- grand total ---
    ws.cell(row=r, column=1, value="GRAND TOTAL")
    write_money(r, (src_stats["orig"].sum(), src_stats["total"].sum(), src_stats["variation"].sum()))
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=r, column=c)
        cell.fill = grand_fill
        cell.font = Font(bold=True, size=14)
        cell.border = thick_bottom
 
    return r
 
 
def write_variations_sheet(wb, df, extra_date_columns=None):
    """Variations sheet: pivot-style summary at the top (source file ->
    job, Sum of total, most common comment), then the detail rows
    underneath. Only rows where variations != 0 are passed in."""
 
    if df is None or df.empty:
        return write_sheet(wb, "Variations", df)
 
    job_stats, src_stats = build_variations_pivot(df)
 
    # Work out how tall the pivot block will be, so the detail table
    # can start below it (outline groups hide whole rows, so the two
    # can't sit side by side).
    pivot_start = 2
    n_pivot_rows = 1 + len(src_stats) + len(job_stats) + 1   # header + sources + jobs + grand total
    pivot_end = pivot_start + n_pivot_rows - 1
    table_header_row = pivot_end + 3          # two blank rows between pivot and table
 
    ws = write_sheet(
        wb, "Variations", df,
        extra_date_columns=extra_date_columns,
        header_row=table_header_row
    )
 
    write_pivot_block(ws, job_stats, src_stats, pivot_start)
 
    # widths now need to cover both the pivot and the table
    autofit_columns(ws, start_row=2)
 
    return ws
 

# =========================================================
# 🏷️ VARIATIONS BY SERVICE CODE (optional sheet)
# =========================================================
def write_service_code_sheet(wb, df):
    """One row per distinct service code (rows with variations != 0),
    with the description, summed quantities (qty / qsub), Orig, Total
    and Variation (Total - Orig), plus a GRAND TOTAL row. Items without
    a service code come last, one row per description, marked either
    '(no code in miscelaneous file)' or '(description not in miscelaneous file)'."""
 
    ws = wb.create_sheet(SERVICE_CODE_SHEET[:31])
 
    if df is None or df.empty:
        ws.append(["No Data"])
        return ws
 
    p = df.copy()
    code = p[SERVICE_CODE_COL].fillna("").astype(str).str.strip() if SERVICE_CODE_COL in p.columns \
        else pd.Series("", index=p.index)
    status = p["_sc_status"].fillna("") if "_sc_status" in p.columns else pd.Series("", index=p.index)
    status = status.where(status != "", SC_NOT_FOUND)
    has_code = code != ""
    p["_code"] = code.where(has_code, status)
    # rows without a code are listed per description, so it's clear which items they are
    p["_desc"] = p["item_norm"].where(~has_code, "") if "item_norm" in p.columns else ""
 
    agg = (
        p.groupby(["_code", "_desc"], as_index=False)
         .agg(item=("item", most_common_value),
              quantity_original=("quantity_original", "sum"),
              quantity_used=("quantity_used", "sum"),
              orig=("orig", "sum"),
              total=("total", "sum"))
    )
    agg["variation"] = agg["total"] - agg["orig"]
    # real codes first (sorted), then the rows with no code, then the ones not in the file
    order = {SC_NO_CODE: "1", SC_NOT_FOUND: "2"}
    agg["_sort"] = agg["_code"].map(order).fillna("0") + agg["_code"] + "\u0001" + agg["_desc"]
    agg = agg.sort_values("_sort")
 
    headers = [SERVICE_CODE_COL, "item", "quantity_original", "quantity_used",
               "Orig", "Total", "Variation"]
    data_cols = ["_code", "item", "quantity_original", "quantity_used",
                 "orig", "total", "variation"]
    money = {"Orig", "Total", "Variation"}
 
    header_fill = PatternFill(start_color="00CCFF", end_color="00CCFF", fill_type="solid")
    grand_fill = PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")
    header_row = 2
 
    try:
        img1 = Image(IMG_LEFT); img1.width, img1.height = 110, 75
        img2 = Image(IMG_RIGHT); img2.width, img2.height = 140, 60
        ws.add_image(img1, "A1"); ws.add_image(img2, "B1")
        ws.row_dimensions[1].height = 60
    except Exception as e:
        print(f"Image load failed: {e}")
 
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = Font(bold=True, size=15)
        cell.alignment = Alignment(horizontal="center")
 
    r = header_row + 1
    for _, row in agg.iterrows():
        for c_idx, (h, col) in enumerate(zip(headers, data_cols), 1):
            val = row[col]
            if col == "item" and (val == "" or pd.isna(val)):
                val = None
            cell = ws.cell(row=r, column=c_idx, value=val)
            if h in money:
                cell.number_format = MONEY_FORMAT
        r += 1
 
    # grand total
    ws.cell(row=r, column=1, value="GRAND TOTAL")
    for c_idx, (h, col) in enumerate(zip(headers, data_cols), 1):
        if col in ("quantity_original", "quantity_used", "orig", "total", "variation"):
            cell = ws.cell(row=r, column=c_idx, value=float(agg[col].sum()))
            if h in money:
                cell.number_format = MONEY_FORMAT
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=r, column=c)
        cell.fill = grand_fill
        cell.font = Font(bold=True, size=14)
        cell.border = Border(bottom=Side(style="medium"))
 
    autofit_columns(ws, start_row=header_row, max_width=60)
    return ws
 
 
# =========================================================
# 🔥 CV FUNCTIONS
# =========================================================
def cv7_set(df, mapping):
    tmp = df.copy()
 
    tmp["item_norm"] = tmp["item"].apply(normalize_item)
    tmp["pole"] = tmp["pole"].apply(normalize_pole)
 
    keys = set(normalize_item(k) for k in mapping.keys())
    tmp = tmp[tmp["item_norm"].isin(keys)]
 
    return set(tmp["pole"].dropna())
 
 
def process_cv(df, mapping, cv7_poles):
    tmp = df.copy()
 
    tmp["item_norm"] = tmp["item"].apply(normalize_item)
    tmp["pole"] = tmp["pole"].apply(normalize_pole)
 
    keys = set(normalize_item(k) for k in mapping.keys())
 
    tmp = tmp[tmp["item_norm"].isin(keys)]
    tmp["quantity_used"] = pd.to_numeric(tmp.get("qsub", 0), errors="coerce").fillna(0)
 
    tmp = tmp[tmp["quantity_used"] != 0]
    tmp = tmp.drop_duplicates(subset=["pole"])
    tmp = tmp[~tmp["pole"].isin(cv7_poles)]
 
    return tmp
 
 
# =========================================================
# 📊 EXCEL GENERATOR
# =========================================================
def generate_excel(df, categories, cv_groups, selected_cv_groups=None, extra_date_columns=None,
                   service_codes=None, include_service_code_sheet=False):
    """service_codes: dict {description: code} from load_service_codes().
    If None, it is loaded from SERVICE_CODES_PATH (if that fails, the
    service code column is just left blank).
    include_service_code_sheet: add the optional 'Variations_service codes' sheet."""
 
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
 
    if selected_cv_groups is None:
        selected_cv_groups = list(cv_groups.keys())
 
    cv_groups_norm = {normalize_cv_name(k) for k in cv_groups.keys()}
    selected_norm = {normalize_cv_name(s) for s in selected_cv_groups}
 
    # =========================
    # CLEAN INPUT
    # =========================
    df = clean_dataframe(df).copy()
 
    def safe_numeric(series):
        return pd.to_numeric(
            series.astype(str)
                  .str.replace(",", "", regex=False)
                  .str.replace("£", "", regex=False)
                  .str.replace(" ", "", regex=False),
            errors="coerce"
        ).fillna(0)
 
    def get_numeric_col(df, col):
        if col in df.columns:
            return safe_numeric(df[col])
        return pd.Series(0.0, index=df.index)
 
    df["quantity_used"] = get_numeric_col(df, "qsub")
    df["quantity_original"] = get_numeric_col(df, "qty")
    df["variations"] = get_numeric_col(df, "qcvi")
 
    df["total"] = get_numeric_col(df, "total")
    df["orig"] = get_numeric_col(df, "orig")
 
    df["item_norm"] = df["item"].apply(normalize_item)
 
    # =========================
    # 🏷️ SERVICE CODE (looked up from the description)
    # =========================
    if service_codes is None:
        try:
            service_codes = load_service_codes(SERVICE_CODES_PATH)
        except Exception as e:
            print(f"Service codes not loaded from {SERVICE_CODES_PATH} ({e}) - service code column will be blank")
            service_codes = {}
    codes, status = match_service_codes(df["item_norm"].tolist(), service_codes)
    df[SERVICE_CODE_COL] = pd.Series(codes, index=df.index, dtype=object)
    df["_sc_status"] = pd.Series(status, index=df.index, dtype=object)
 
    LAST_SERVICE_CODE_STATS.clear()
    if service_codes:
        missing = df.loc[df["_sc_status"] == SC_NOT_FOUND, "item"].dropna().astype(str).str.strip()
        LAST_SERVICE_CODE_STATS.update({
            "rows": int(len(df)),
            "with_code": int(df[SERVICE_CODE_COL].notna().sum()),
            "no_code_in_file": int((df["_sc_status"] == SC_NO_CODE).sum()),
            "not_in_file": int((df["_sc_status"] == SC_NOT_FOUND).sum()),
            "not_in_file_items": missing.value_counts().head(15).index.tolist(),
        })
 
    if "pole" in df.columns:
        df["pole"] = df["pole"].apply(normalize_pole)
 
    # =========================
    # 🔧 "job" COLUMN (most common job/segment text per circuit)
    # =========================
    if "circuit" in df.columns and "segment" in df.columns:
        circuit_job_mode = df.groupby("circuit")["segment"].agg(most_common_value)
        df["job"] = df["circuit"].map(circuit_job_mode)
    else:
        df["job"] = ""
 
    if "project_name" not in df.columns:
        df["project_name"] = ""
 
    # =========================
    # 🔧 HV POLE MULTIPLIER (CV7_recover special case)
    # =========================
    HV_POLE_MULTIPLIER = 2
    hv_pole_key = normalize_item(
        "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions"
    )
    hv_pole_mask = df["item_norm"] == hv_pole_key
 
    for col in ["quantity_original", "variations", "quantity_used"]:
        df.loc[hv_pole_mask, col] = df.loc[hv_pole_mask, col] * HV_POLE_MULTIPLIER
 
    # =========================
    # SHEETS
    # =========================
    # 🔥 "Combined" is now "Outputs" (with total + orig columns)
    write_sheet(wb, "Outputs", df, extra_date_columns=extra_date_columns)
 
    # 🔥 Variations: only rows where variations is NOT 0, plus pivot block
    variations_df = df[df["variations"] != 0].copy()
    write_variations_sheet(wb, variations_df, extra_date_columns=extra_date_columns)
 
    # 🔥 Optional: variations grouped by service code
    if include_service_code_sheet:
        write_service_code_sheet(wb, variations_df)
 
    for name, mapping, *_ in categories:
        if normalize_cv_name(name) in cv_groups_norm and normalize_cv_name(name) not in selected_norm:
            continue
        keys = set(normalise_mapping_keys(mapping).keys())
        write_sheet(wb, name, df[df["item_norm"].isin(keys)].copy(), extra_date_columns=extra_date_columns)
 
    try:
        for name, mapping in extra_categories:
            if normalize_cv_name(name) in cv_groups_norm and normalize_cv_name(name) not in selected_norm:
                continue
            try:
                keys = set(normalise_mapping_keys(mapping).keys())
                write_sheet(wb, name, df[df["item_norm"].isin(keys)].copy(), extra_date_columns=extra_date_columns)
            except:
                write_sheet(wb, name, pd.DataFrame())
    except:
        pass
 
    # =========================
    # CV LOGIC
    # =========================
    cv7_poles = set()
    cv7_poles |= cv7_set(df, CV7_erect)
    cv7_poles |= cv7_set(df, CV7_erect_H)
    cv7_poles |= cv7_set(df, CV7_erect_lv)
    cv7_poles |= cv7_set(df, CV7_recover)
 
    cv8_resume = process_cv(df, cv_groups["CV8"], cv7_poles) if "CV8" in selected_cv_groups else pd.DataFrame()
    cv31_resume = process_cv(df, cv_groups["CV31"], cv7_poles) if "CV31" in selected_cv_groups else pd.DataFrame()
 
    if "CV8" in selected_cv_groups:
        write_sheet(wb, "CV8_resume", cv8_resume, extra_date_columns=extra_date_columns)
    if "CV31" in selected_cv_groups:
        write_sheet(wb, "CV31_resume", cv31_resume, extra_date_columns=extra_date_columns)
 
    # =========================
    # SUMMARY BUILD
    # =========================
    if "project" in df.columns:
 
        group_cols = []
        if "shire" in df.columns:
            group_cols.append("shire")
        group_cols.append("project")
        if "project_name" in df.columns:
            group_cols.append("project_name")
        if "circuit" in df.columns:
            group_cols.append("circuit")
 
        numeric_cols = ["total", "orig"]
 
        summary_df = (
            df.groupby(group_cols, as_index=False)[numeric_cols]
              .sum()
        )
 
        # CATEGORY COLUMNS
        for name, mapping, *_ in categories:
            if normalize_cv_name(name) in cv_groups_norm and normalize_cv_name(name) not in selected_norm:
                continue
 
            keys = set(normalise_mapping_keys(mapping).keys())
 
            tmp = (
                df[df["item_norm"].isin(keys)]
                .groupby(group_cols, as_index=False)["quantity_used"]
                .sum()
                .rename(columns={"quantity_used": name})
            )
 
            summary_df = summary_df.merge(tmp, on=group_cols, how="outer")
 
        # extra categories
        try:
            for name, mapping in extra_categories:
                if normalize_cv_name(name) in cv_groups_norm and normalize_cv_name(name) not in selected_norm:
                    continue
                try:
                    keys = set(normalise_mapping_keys(mapping).keys())
 
                    tmp = (
                        df[df["item_norm"].isin(keys)]
                        .groupby(group_cols, as_index=False)["quantity_used"]
                        .sum()
                        .rename(columns={"quantity_used": name})
                    )
 
                    summary_df = summary_df.merge(tmp, on=group_cols, how="outer")
                except:
                    pass
        except:
            pass
 
        # CV COUNTS
        if "CV8" in selected_cv_groups and not cv8_resume.empty:
            cv8_tmp = (
                cv8_resume.groupby(group_cols)
                          .size()
                          .reset_index(name="CV8_resume")
            )
            summary_df = summary_df.merge(cv8_tmp, on=group_cols, how="outer")
 
        if "CV31" in selected_cv_groups and not cv31_resume.empty:
            cv31_tmp = (
                cv31_resume.groupby(group_cols)
                           .size()
                           .reset_index(name="CV31_resume")
            )
            summary_df = summary_df.merge(cv31_tmp, on=group_cols, how="outer")
 
        summary_df = summary_df.fillna(0)
 
        # DERIVED COLUMN
        summary_df["Variations"] = summary_df["total"] - summary_df["orig"]
 
        cols = list(summary_df.columns)
        cols.remove("Variations")
        cols.insert(cols.index("orig") + 1, "Variations")
        summary_df = summary_df[cols]
 
        # Value columns (everything that gets summed)
        numeric_cols_all = [c for c in summary_df.columns if c not in group_cols]
 
        # "project_name" - most common "job" text per circuit/project/district
        if "project_name" in summary_df.columns and "job" in df.columns:
            proj_group_keys = [c for c in group_cols if c != "project_name"]
            if proj_group_keys:
                job_mode_df = (
                    df.groupby(proj_group_keys, as_index=False)["job"]
                      .agg(most_common_value)
                      .rename(columns={"job": "_job_mode"})
                )
                summary_df = summary_df.merge(job_mode_df, on=proj_group_keys, how="left")
                summary_df["project_name"] = summary_df["_job_mode"].fillna("")
                summary_df = summary_df.drop(columns=["_job_mode"])
 
        # "Most Common Comment" (only when "plan2" is ticked) - ignores
        # blanks / nan and comments shorter than two words
        has_comment = (
            extra_date_columns and "plan2" in extra_date_columns
            and "circuit" in summary_df.columns and "comment" in df.columns
        )
        if has_comment:
            circuit_comment_mode = df.groupby("circuit")["comment"].agg(most_common_comment)
            summary_df["Most Common Comment"] = summary_df["circuit"].map(circuit_comment_mode).fillna("")
 
        # =========================
        # PIVOT LAYOUT: District row (district totals)
        #   > project row (project totals)
        #     > circuit rows (grouped, collapsed)
        # then GRAND TOTAL. "_level" tells write_sheet how to style and
        # group each row; it is not written to the sheet.
        # =========================
        has_shire = "shire" in group_cols
        out_cols = [c for c in ["shire", "project", "project_name", "circuit"] if c in summary_df.columns]
        if has_comment:
            out_cols.append("Most Common Comment")
        out_cols += numeric_cols_all
 
        def totals_row(sub, level, **labels):
            row = {c: None for c in out_cols}
            row.update(labels)
            for c in numeric_cols_all:
                row[c] = sub[c].sum()
            row["_level"] = level
            return row
 
        base = summary_df.sort_values(group_cols)
        pivot_rows = []
        shire_iter = base.groupby("shire", sort=True) if has_shire else [(None, base)]
 
        for shire_val, g_shire in shire_iter:
            if has_shire:
                pivot_rows.append(totals_row(g_shire, 0, shire=shire_val))
 
            for project, g_proj in g_shire.groupby("project", sort=True):
                pivot_rows.append(totals_row(g_proj, 1, project=project))
 
                for _, r in g_proj.iterrows():
                    circuit_row = {c: r[c] for c in out_cols if c not in ("shire", "project")}
                    circuit_row["_level"] = 2
                    pivot_rows.append(circuit_row)
 
        pivot_rows.append(totals_row(base, 3, project="GRAND TOTAL"))
 
        summary_df = pd.DataFrame(pivot_rows, columns=out_cols + ["_level"])
        summary_df = summary_df.rename(columns={"shire": "District"})
 
    else:
        summary_df = pd.DataFrame()
        
    # =========================
    # WRITE SUMMARY
    # =========================
    write_sheet(wb, "Summary", summary_df)
 
    if "Summary" in wb.sheetnames:
        ws = wb["Summary"]
        wb._sheets.remove(ws)
        wb._sheets.insert(0, ws)
 
    # =========================
    # EXPORT
    # =========================
    wb.save(output)
    output.seek(0)
 
    return output.getvalue()


# WEB WRAPPER - replaces the tkinter FilterApp window. Loading, filter
# population, Apply Filters and Export Excel are the original methods'
# bodies, unchanged; only the widgets became function arguments.
# =====================================================================
FILETYPES = [("Parquet", "*.parquet"), ("CSV", "*.csv")]

FILTER_COLUMNS = ["shire", "project", "circuit", "project_name", "cat"]
COLUMN_LABEL_MAP = {"cat": "Category"}

CV_GROUP_NAMES = [
    "CV8", "CV31", "CV7_erect", "CV7_erect_H", "CV7_erect_lv",
    "CV7_recover", "CV7_Tx", "transformer",
    "CV7_OHL_CONDUCTOR_instal", "CV7_OHL_CONDUCTOR_recover",
    "CV7_OHL_CONDUCTOR_LV_instal", "CV7_OHL_CONDUCTOR_LV_recover",
    "Switch", "Fuses"
]

CV_GROUPS = {
    "CV8": CV8,
    "CV31": CV31,
    "CV7_erect": CV7_erect,
    "CV7_erect_H": CV7_erect_H,
    "CV7_erect_lv": CV7_erect_lv,
    "CV7_recover": CV7_recover,
    "CV7_Tx": CV7_Tx,
    "transformer": transformer,
    "CV7_OHL_CONDUCTOR_instal": CV7_OHL_CONDUCTOR_instal,
    "CV7_OHL_CONDUCTOR_recover": CV7_OHL_CONDUCTOR_recover,
    "CV7_OHL_CONDUCTOR_LV_instal": CV7_OHL_CONDUCTOR_LV_instal,
    "CV7_OHL_CONDUCTOR_LV_recover": CV7_OHL_CONDUCTOR_LV_recover,
    "Switch": Switch,
    "Fuses": Fuses
}


def load_file(path):
    """Same as FilterApp.load_file worker."""
    df = pd.read_csv(path) if path.endswith(".csv") else pd.read_parquet(path)
    df = clean_dataframe(df)
    # Translate the new master file's column names to the
    # internal names the rest of this program expects.
    df = df.rename(columns=NEW_TO_OLD_COLUMNS)
    df = df.rename(columns={
        "segmentcode": "circuit",
        "segmentdesc": "project_name"
    })
    return df


def filter_options(df):
    """Same as FilterApp.on_load_success - returns what each box offers."""
    options = {}
    for col in FILTER_COLUMNS:
        if col == "cat":
            # Category filter is fixed to these two options
            options[col] = ["Const", "Mat"]
        elif col in df.columns:
            if col == "project":
                vals = ["All (No Connections)"] + sorted(df[col].dropna().astype(str).unique())
            else:
                vals = sorted(df[col].dropna().astype(str).unique())
            options[col] = vals
        else:
            options[col] = []

    date_cols = [c for c in ["plan1", "plan2", "done", "datetouse"] if c in df.columns]
    extra_options = ["all", "unplanned", "undone"]
    if date_cols:
        date_values, date_default = date_cols + extra_options, [date_cols[0]]
    else:
        date_values, date_default = extra_options, []

    return {
        "rows": int(len(df)),
        "columns": FILTER_COLUMNS,
        "labels": {c: COLUMN_LABEL_MAP.get(c, c) for c in FILTER_COLUMNS},
        "options": options,
        "date_options": date_values,
        "date_default": date_default,
        "cv_groups": CV_GROUP_NAMES,
    }


def apply_filters(df, text_filters, selected_dates, start="", end=""):
    """Same as FilterApp.apply_filters. Returns (filtered_df, selected_date_filters)."""
    df = df.copy()
    text_filters = text_filters or {}

    # =========================
    # TEXT FILTERS
    # =========================
    for col in FILTER_COLUMNS:
        if col not in df.columns:
            continue

        selected = text_filters.get(col) or []

        if not selected:
            continue

        df[col] = df[col].astype(str).str.strip()
        selected_lower = [v.lower() for v in selected]

        # SPECIAL CASE: PROJECT FILTER
        if col == "project" and "all (no connections)" in selected_lower:
            other_selected = [v for v in selected if v.lower() != "all (no connections)"]
            mask = df[col].str.lower() != "connections"
            if other_selected:
                mask = mask | df[col].str.lower().isin([v.lower() for v in other_selected])
            df = df[mask]
            continue

        df = df[df[col].str.lower().isin(selected_lower)]

    # =========================
    # DATE FILTER (additive / OR across the ticked columns)
    # =========================
    selected_dates = selected_dates or []
    start = (start or "").strip()
    end = (end or "").strip()

    specials = {"all", "unplanned", "undone"}
    special_selected = [d for d in selected_dates if d in specials]
    date_col_selected = [d for d in selected_dates if d not in specials]

    if not selected_dates or "all" in special_selected:
        pass
    else:
        mask = pd.Series(False, index=df.index)
        any_criterion = False

        if "unplanned" in special_selected and "plan1" in df.columns and "done" in df.columns:
            mask = mask | (df["plan1"].isna() & df["done"].isna())
            any_criterion = True

        if "undone" in special_selected and "done" in df.columns:
            mask = mask | df["done"].isna()
            any_criterion = True

        for date_col in date_col_selected:
            if date_col in df.columns:
                parsed = pd.to_datetime(df[date_col], errors="coerce")
                col_mask = parsed.notna()

                if start:
                    col_mask = col_mask & (parsed >= pd.to_datetime(start))

                if end:
                    col_mask = col_mask & (parsed <= pd.to_datetime(end))

                mask = mask | col_mask
                any_criterion = True

        if any_criterion:
            df = df[mask]

    selected_date_filters = [d for d in date_col_selected if d not in BASE_EXPORT_COLUMNS]
    return df, selected_date_filters


def export(df_filtered, selected_date_filters, selected_cv_groups, output_path,
           service_codes_path="", include_service_code_sheet=False, log=print):
    """Same as FilterApp.export_excel worker.
    service_codes_path: the miscelaneous file dropped on the page (Column_1 =
    description, Column_3 = service code). Without it the "service code"
    column is left blank - the network path can't be reached from the browser."""
    if not selected_cv_groups:
        selected_cv_groups = list(CV_GROUPS.keys())
    service_codes = {}
    if service_codes_path:
        service_codes = load_service_codes(service_codes_path)
        log(f"Service codes loaded: {len(service_codes)}")
    else:
        log("No miscelaneous file given - the service code column will be blank.")
    log(f"Generating Excel for {len(df_filtered)} rows, CV groups: {', '.join(selected_cv_groups)}")
    excel_bytes = generate_excel(
        df_filtered,
        categories,
        CV_GROUPS,
        selected_cv_groups=selected_cv_groups,
        extra_date_columns=selected_date_filters,
        service_codes=service_codes,
        include_service_code_sheet=bool(include_service_code_sheet)
    )
    if not output_path.lower().endswith(".xlsx"):
        output_path += ".xlsx"
    with open(output_path, "wb") as f:
        f.write(excel_bytes)
    if service_codes:
        for line in service_code_summary().split("\n"):
            log(line)
    log(f"✅ Excel exported successfully: {output_path}")
    return [output_path]
