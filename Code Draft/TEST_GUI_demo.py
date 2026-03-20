# ============================================================
# INDOTECH HARDWARE PLATFORM — UI v3  (FULLY INTEGRATED)
# ============================================================
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import time, re, os, json, threading, math, webbrowser
from collections import defaultdict
from datetime import datetime

# ── Optional heavy imports ───────────────────────────────────
try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.drawing.image import Image as XLImage
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    import win32com.client
    WIN32_OK = True
except ImportError:
    WIN32_OK = False

try:
    import openpyxl as _oxl
    _oxl_ok = True
except ImportError:
    _oxl_ok = False

# ============================================================
# BASE PATHS
# ============================================================
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
IMG_DIR     = os.path.join(BASE_DIR, "img")
SOUND_DIR   = os.path.join(BASE_DIR, "sound")
# Check root folder first (old behaviour), then img/ subfolder
_LOGO_ROOT = os.path.join(BASE_DIR, "logo.png")
_LOGO_IMG  = os.path.join(IMG_DIR,  "logo.png")
LOGO_PATH  = _LOGO_ROOT if os.path.exists(_LOGO_ROOT) else _LOGO_IMG
ICO_PATH    = os.path.join(IMG_DIR,  "app.ico")
WAV_PATH    = os.path.join(SOUND_DIR,"startup.wav")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# ============================================================
# CONFIG
# ============================================================
DEFAULT_CONFIG = {
    "weight_table_path" : "",
    "last_output_folder": "",
    "theme"             : "dark",
    "open_after_gen"    : True,
    "autosave_folder"   : False,
    "prepared_by"       : "RG",
    "checked_by"        : "JN",
    "approved_by"       : "JN",
    "format_number"     : "F/DSN/54",
}

def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH) as f:
                saved = json.load(f)
            cfg = DEFAULT_CONFIG.copy(); cfg.update(saved); return cfg
        except: pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    try:
        with open(CONFIG_PATH,"w") as f: json.dump(cfg, f, indent=2)
    except Exception as e: print(f"Config save failed: {e}")

CONFIG = load_config()

# ============================================================
# HARDCODED WEIGHT TABLE — fallback when Excel entry is missing
# ============================================================
WEIGHT_TABLE_BUILTIN = {
    "M4X10-HDG":800,"M4X12-HDG":750,"M4X16-HDG":680,"M4X20-HDG":620,
    "M5X10-HDG":800,"M5X16-HDG":680,"M5X20-HDG":620,"M5X30-HDG":500,
    "M5X40-HDG":420,"M5X60-HDG":300,
    "M6X10-HDG":420,"M6X16-HDG":350,"M6X20-HDG":310,"M6X30-HDG":245,
    "M6X40-HDG":200,"M6X50-HDG":170,"M6X60-HDG":148,
    "M8X16-HDG":200,"M8X20-HDG":175,"M8X30-HDG":136,"M8X40-HDG":110,
    "M8X50-HDG":92,"M8X60-HDG":80,"M8X80-HDG":62,
    "M10X20-HDG":110,"M10X30-HDG":95,"M10X40-HDG":85,"M10X50-HDG":75,
    "M10X60-HDG":67,"M10X70-HDG":60,"M10X80-HDG":54,"M10X100-HDG":45,
    "M12X20-HDG":78,"M12X30-HDG":65,"M12X40-HDG":58,"M12X50-HDG":52,
    "M12X60-HDG":46,"M12X70-HDG":42,"M12X80-HDG":38,"M12X90-HDG":34,
    "M12X100-HDG":31,
    "M16X20-HDG":44,"M16X30-HDG":37,"M16X40-HDG":32,"M16X50-HDG":28,
    "M16X60-HDG":25,"M16X70-HDG":22,"M16X80-HDG":20,"M16X90-HDG":18,
    "M16X100-HDG":17,"M16X120-HDG":15,"M16X130-HDG":14,
    "M20X30-HDG":22,"M20X40-HDG":18,"M20X50-HDG":16,"M20X60-HDG":14,
    "M20X70-HDG":12,"M20X80-HDG":11,"M20X100-HDG":9,"M20X120-HDG":8,
    "M24X50-HDG":9,"M24X60-HDG":9,"M24X70-HDG":8,"M24X80-HDG":7,
    "M24X90-HDG":6,"M24X100-HDG":6,"M24X120-HDG":5,
    "M30X60-HDG":5,"M30X80-HDG":4,"M30X100-HDG":3,"M30X120-HDG":3,"M30X150-HDG":2,
    "M36X80-HDG":2,"M36X100-HDG":2,"M36X120-HDG":2,"M36X150-HDG":1,
}


# ============================================================
# THEME
# ============================================================
THEMES = {
    "dark":  {"BG_DEEP":"#0A0C10","BG_PANEL":"#111318","BG_CARD":"#181C24",
               "BG_SIDEBAR":"#0D0F14","BORDER_COLOR":"#22283A",
               "TEXT_PRIMARY":"#EAEDF5","TEXT_SECONDARY":"#6B7799","TEXT_MUTED":"#3A4060"},
    "light": {"BG_DEEP":"#F0F2F7","BG_PANEL":"#FFFFFF","BG_CARD":"#F8F9FC",
               "BG_SIDEBAR":"#1A1F2E","BORDER_COLOR":"#DDE1EC",
               "TEXT_PRIMARY":"#0F1420","TEXT_SECONDARY":"#4A5270","TEXT_MUTED":"#9AA0B8"},
}
ACCENT=     "#FF6B00"
ACCENT_GLOW="#FF8C33"
ACCENT_DARK="#CC5500"
GREEN=      "#00C96A"
RED=        "#FF3B5C"
BLUE=       "#4FC3F7"
YELLOW=     "#FFD600"
PURPLE=     "#B388FF"
WHITE=      "#FFFFFF"
BADGE_PALETTE=[ACCENT,BLUE,GREEN,YELLOW,PURPLE,RED,"#FF8A65","#80CBC4","#F48FB1","#AED581"]

T = THEMES[CONFIG.get("theme","dark")].copy()

def apply_theme(name): T.update(THEMES[name]); CONFIG["theme"]=name; save_config(CONFIG)
def BG_DEEP():    return T["BG_DEEP"]
def BG_PANEL():   return T["BG_PANEL"]
def BG_CARD():    return T["BG_CARD"]
def BG_SIDEBAR(): return T["BG_SIDEBAR"]
def BORDER():     return T["BORDER_COLOR"]
def FG():         return T["TEXT_PRIMARY"]
def FG2():        return T["TEXT_SECONDARY"]
def FG_MUTED():   return T["TEXT_MUTED"]

# ============================================================
# GLOBAL APP STATE
# ============================================================
app_state = {
    "weight_table"      : {},
    "weight_table_path" : CONFIG.get("weight_table_path",""),
    "coatings_loaded"   : set(),
    "selected_pdf"      : "",
    "current_page"      : "dashboard",
    "history"           : [],
}

# ============================================================
# BUILT-IN WEIGHT TABLE (hardcoded HDG values — always available)
# Used when no Excel weight table is loaded.
# ============================================================

# ============================================================
# WEIGHT TABLE LOGIC
# ============================================================
def parse_size_from_description(desc):
    m = re.search(r'(M\d+)\s*[Xx]\s*(\d+)', desc.strip().upper())
    return (m.group(1), m.group(2)) if m else (None, None)

def extract_coating_from_sheet_name(sheet_name):
    name = sheet_name.strip()
    for s in [" hardware"," coating"," material"," bolts"," fasteners"]:
        name = re.sub(s,"",name,flags=re.IGNORECASE)
    name = name.strip().upper()
    norm = {"BLACKODIZED":"BLACK","BLACKODISED":"BLACK","BLACK OXIDE":"BLACK",
            "HIGH TENSILE":"HT","STAINLESS STEEL":"SS",
            "HOT DIP GALVANIZED":"HDG","HOT DIP GALVANISED":"HDG",
            "GALVANIZED":"GALV","GALVANISED":"GALV","MAGNIGOLD":"MAGNI","DACROMET":"DACRO"}
    return norm.get(name, name[:8])

def detect_sub_coating(desc, default):
    d = desc.upper()
    for kws, code in [
        (["HIGH TENSILE"," HT ","HT FT","-HT"],"HT"),
        (["STAINLESS"," SS ","SS 316","SS316"],"SS"),
        (["MAGNI","MAGNIGOLD"],"MAGNI"),
        (["DACROMET","DACRO"],"DACRO"),
        (["ZINC","GALVANIZED","GALVANISED"],"GALV"),
    ]:
        if any(k in d for k in kws): return code
    return default

def load_weight_table_from_excel(path):
    if not _oxl_ok:
        return None, set(), [("error","openpyxl not installed")]
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path)
    except Exception as e:
        return None, set(), [("error", f"Cannot open file: {e}")]

    weight_dict = {}; coatings = set()
    log_msgs = [("info", f"{len(wb.sheetnames)} sheet(s): {', '.join(wb.sheetnames)}")]

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        coating_code = extract_coating_from_sheet_name(sheet_name)
        hdr = [str(c.value).lower().strip() if c.value else ""
               for c in next(ws.iter_rows(max_row=1))]
        desc_col  = next((i for i,h in enumerate(hdr) if "desc" in h), 1)
        unit_col  = next((i for i,h in enumerate(hdr) if "unit" in h), 2)
        pcskg_col = next((i for i,h in enumerate(hdr)
                          if "per kg" in h or ("quantity" in h and i > 2)), None)
        count_kg = count_na = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[desc_col]: continue
            description = str(row[desc_col]).strip()
            unit = str(row[unit_col]).strip().lower() if row[unit_col] else "no"
            pcs_val = row[pcskg_col] if pcskg_col and len(row) > pcskg_col else None
            size, length = parse_size_from_description(description)
            if not size or not length: continue
            rounded = math.ceil(int(length)/10)*10
            sub_coating = detect_sub_coating(description, coating_code)
            key = f"{size}X{rounded}-{sub_coating}"
            if unit == "kg" and pcs_val and str(pcs_val).strip().replace('.','').isdigit():
                weight_dict[key] = int(float(str(pcs_val).strip()))
                count_kg += 1
            else:
                weight_dict[key] = "N/A"; count_na += 1
            coatings.add(sub_coating)
        log_msgs.append(("success",
            f"'{sheet_name}' [{coating_code}]: {count_kg} with weight, {count_na} qty-only"))

    log_msgs.append(("info", f"Total entries: {len(weight_dict)}"))
    log_msgs.append(("info", f"Coatings: {', '.join(sorted(coatings))}"))
    return weight_dict, coatings, log_msgs

# ============================================================
# BACKEND — PDF PARSING & EXCEL EXPORT
# ============================================================
def _extract_pdf_lines(path):
    lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text: lines.extend(text.split("\n"))
    return lines

def _extract_header(lines):
    h = {"WORK_ORDER":"","REVISION":"","CUSTOMER":"","RATING":""}
    for line in lines:
        u = line.upper()
        if "WORK ORDER" in u:
            m = re.search(r"WORK ORDER\s+(\d+)", u)
            if m: h["WORK_ORDER"] = m.group(1)
        if "REVISION" in u:
            m = re.search(r"REVISION\s+([A-Z0-9]+)", u)
            if m: h["REVISION"] = m.group(1)
        if "CUSTOMER" in u:
            m = re.search(r"CUSTOMER\s*(.*)", line, re.IGNORECASE)
            if m:
                name = m.group(1).strip()
                if name: h["CUSTOMER"] = name
        if "RATING" in u:
            m = re.search(r"RATING\s*(.*?MVA.*?kV)", line, re.IGNORECASE)
            if m: h["RATING"] = m.group(1).strip()
    return h

def _parse_hardware(lines):

    data = []
    size_pattern = re.compile(r"S?-?M\d+\s*[Xx]\s*\d+", re.IGNORECASE)

    for line in lines:

        u = line.upper()

        if not size_pattern.search(u):
            continue

        parts = line.split()

        for i, part in enumerate(parts):

            token = part.upper()

            if re.match(r"S?-?M\d+[Xx]\d+-?[A-Z]*", token):

                try:

                    prev = parts[i-1].upper() if i > 0 else ""

                    if "BOLT" in prev or "SCREW" in prev:
                        fastener = "BOLT"

                    elif "STUD" in prev or token.startswith("S-"):
                        fastener = "STUD"

                    else:
                        fastener = "BOLT"

                    qty = int(parts[i+1])
                    nut = int(parts[i+2])
                    lnut = int(parts[i+3])
                    pw  = int(parts[i+4])
                    sw  = int(parts[i+5])

                    if qty == 0:
                        continue

                    data.append({
                        "FASTENER": fastener,
                        "SIZE_MATERIAL": token,
                        "QTY": qty,
                        "NUT": nut,
                        "L_NUT": lnut,
                        "P_WASHER": pw,
                        "S_WASHER": sw,
                    })

                except:
                    continue

    return data

def _build_summary(data):
    summary = defaultdict(lambda: {"FULL_NUT":0,"LOCK_NUT":0,"PLAIN_WASHER":0,"SPRING_WASHER":0})
    for row in data:
        m = re.search(r"M\d+", row["SIZE_MATERIAL"].upper())
        if not m: continue
        size = m.group(0)
        mm = re.search(r"-([A-Z]+)$", row["SIZE_MATERIAL"].upper())
        if mm:
            extracted = mm.group(1)
            material = extracted if extracted in ["HDG","MAGNI","SS","HT","BLACK","GALV"] else "HDG"
        else:
            material = "HDG"
        key = (size, material)
        summary[key]["FULL_NUT"]     += row["NUT"]
        summary[key]["LOCK_NUT"]     += row["L_NUT"]
        summary[key]["PLAIN_WASHER"] += row["P_WASHER"]
        summary[key]["SPRING_WASHER"]+= row["S_WASHER"]
    return summary

def _is_below_m10(size_material):
    key = re.sub(r"^S-","", size_material.upper().strip())
    m = re.search(r"M(\d+)", key)
    return int(m.group(1)) < 10 if m else False

def _round_size_material(size_material):
    key = size_material.upper().strip()
    prefix = ""
    if key.startswith("S-"): prefix="S-"; key=key[2:]
    m = re.match(r"(M\d+)X(\d+)(-[A-Z]+)", key)
    if m:
        size_num = int(re.search(r"\d+", m.group(1)).group())
        coating  = "-SS" if size_num < 10 else m.group(3)
        rounded  = math.ceil(int(m.group(2))/10)*10
        return f"{prefix}{m.group(1)}X{rounded}{coating}"
    return size_material

def _get_weight(size_material, qty):
    wt = app_state.get("weight_table", {})
    key = re.sub(r"^S-","", size_material.upper().strip())
    m = re.match(r"(M\d+)X(\d+)(-[A-Z]+)", key)
    if m:
        rounded = math.ceil(int(m.group(2))/10)*10
        key = f"{m.group(1)}X{rounded}{m.group(3)}"
    # Try Excel-loaded table first, then fall back to builtin hardcoded table
    pcs = wt.get(key, None)
    if pcs is None or pcs == "N/A":
        pcs = WEIGHT_TABLE_BUILTIN.get(key, None)
    if pcs is None: return "N/A", "N/A"
    return pcs, round(qty / pcs, 2)

def run_export(pdf_path, out_folder, open_after, progress_cb, done_cb, error_cb):
    """Full pipeline. Runs in background thread."""
    try:
        if not PDF_OK:    raise RuntimeError("pdfplumber not installed. Run: pip install pdfplumber")
        if not PANDAS_OK: raise RuntimeError("pandas not installed. Run: pip install pandas")
        if not OPENPYXL_OK: raise RuntimeError("openpyxl not installed. Run: pip install openpyxl")

        progress_cb("Step 1/6 — Reading PDF...")
        lines = _extract_pdf_lines(pdf_path)
        if not lines: error_cb("No text found in PDF. Is it scanned?"); return

        progress_cb("Step 2/6 — Extracting header...")
        header = _extract_header(lines)

        progress_cb("Step 3/6 — Parsing hardware lines...")
        data = _parse_hardware(lines)
        if not data: error_cb("No BOLT/STUD lines found in PDF.\nCheck PDF format."); return

        progress_cb("Step 4/6 — Building summary...")
        summary = _build_summary(data)

        # Group bolts
        bolt_rows = [d for d in data if "BOLT" in d["FASTENER"]]
        stud_rows = [d for d in data if "STUD" in d["FASTENER"]]

        bolt_grouped = defaultdict(int)
        for r in bolt_rows:
            bolt_grouped[_round_size_material(r["SIZE_MATERIAL"])] += r["QTY"]

        stud_grouped = defaultdict(int)
        for r in stud_rows:
            stud_grouped[_round_size_material(r["SIZE_MATERIAL"])] += r["QTY"]

        progress_cb("Step 5/6 — Writing Excel...")
        work_order  = header["WORK_ORDER"] or "UNKNOWN"
        folder_name = f"{work_order}_Hardware_Output"
        folder_path = os.path.join(out_folder, folder_name)
        os.makedirs(folder_path, exist_ok=True)

        excel_path = os.path.join(folder_path, f"Hardware_Output_{work_order}.xlsx")
        pdf_out    = os.path.join(folder_path, f"Hardware_Output_{work_order}.pdf")

        wb = Workbook(); ws = wb.active; ws.title = "HARDWARE LIST"
        bold   = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center")
        thin   = Border(left=Side(style="thin"),  right=Side(style="thin"),
                        top=Side(style="thin"),   bottom=Side(style="thin"))
        thick  = Border(left=Side(style="medium"),right=Side(style="medium"),
                        top=Side(style="medium"), bottom=Side(style="medium"))

        for col, w in [("A",26),("B",16),("C",16),("D",24),("E",14),("F",14),("G",14)]:
            ws.column_dimensions[col].width = w

        # Logo / Header
        ws.merge_cells("A1:B3"); ws.merge_cells("C1:E3")
        try:
            # Use same method as original working code — direct path
            if os.path.exists(LOGO_PATH):
                xi = XLImage(LOGO_PATH); xi.width=240; xi.height=55
                ws.add_image(xi,"A1")
            else:
                progress_cb(f"WARNING: Logo not found at {LOGO_PATH}")
        except Exception as _logo_err:
            progress_cb(f"WARNING: Logo skipped ({_logo_err})")
        ws["C1"]="HARDWARE LIST (FINAL QUANTITY)"
        ws["C1"].font=Font(size=15,bold=True); ws["C1"].alignment=center
        ws["F1"]="FORMAT NO"; ws["G1"]=CONFIG.get("format_number","F/DSN/54")
        ws["F2"]="Version";   ws["G2"]="1.3"
        ws["F3"]="DATE";      ws["G3"]=datetime.now().strftime("%d.%m.%y")
        for r in ws["A1:G3"]:
            for c in r: c.border=thick; c.alignment=center

        # Project details
        ws["A4"]="WORK ORDER"; ws["B4"]=header["WORK_ORDER"]
        ws["A5"]="REVISION";   ws["B5"]=header["REVISION"]
        ws["C4"]="CUSTOMER";   ws["D4"]=header["CUSTOMER"]
        ws["C5"]="RATING";     ws["D5"]=header["RATING"]
        for r in ws["A4:G5"]:
            for c in r: c.border=thin; c.alignment=center

        row = 7

        # Summary table
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=7)
        ws.cell(row=row,column=1,value="SUMMARY").font=Font(size=14,bold=True)
        ws.cell(row=row,column=1).alignment=center; row+=1
        for col,head in enumerate(["SIZE","FULL NUT","LOCK NUT","PLAIN WASHER","MATERIAL","SPRING WASHER-SS"],1):
            c=ws.cell(row=row,column=col,value=head); c.font=bold; c.border=thick; c.alignment=center
        row+=1
        for (size,mat),vals in sorted(summary.items()):
            ws.cell(row=row,column=1,value=size).alignment=center
            ws.cell(row=row,column=2,value=vals["FULL_NUT"]).alignment=center
            ws.cell(row=row,column=3,value=vals["LOCK_NUT"]).alignment=center
            ws.cell(row=row,column=4,value=vals["PLAIN_WASHER"]).alignment=center
            ws.cell(row=row,column=5,value=mat).alignment=center
            ws.cell(row=row,column=6,value=vals["SPRING_WASHER"]).alignment=center
            for col in range(1,7): ws.cell(row=row,column=col).border=thin
            row+=1
        row+=2

        # Bolt table
        for col,head in enumerate(["BOLT SIZE AND MATERIAL","BOLT QUANTITY","PCS / KG","TOTAL WEIGHT (KG)"],1):
            c=ws.cell(row=row,column=col,value=head); c.font=Font(bold=True); c.border=thick; c.alignment=center
        row+=1
        total_bolt_wt = 0.0
        for sm, qty in sorted(bolt_grouped.items()):
            pcs, wt = ("N/A","N/A") if _is_below_m10(sm) else _get_weight(sm, qty)
            ws.cell(row=row,column=1,value=sm).alignment=center
            ws.cell(row=row,column=2,value=qty).alignment=center
            ws.cell(row=row,column=3,value=pcs).alignment=center
            ws.cell(row=row,column=4,value=wt).alignment=center
            for col in range(1,5): ws.cell(row=row,column=col).border=thin
            if isinstance(wt,float): total_bolt_wt+=wt
            row+=1
        row+=2

        # Stud table
        for col,head in enumerate(["STUD SIZE AND MATERIAL","STUD QUANTITY","PCS / KG","TOTAL WEIGHT (KG)"],1):
            c=ws.cell(row=row,column=col,value=head); c.font=Font(bold=True); c.border=thick; c.alignment=center
        row+=1
        total_stud_wt = 0.0
        for sm, qty in sorted(stud_grouped.items()):
            pcs, wt = ("N/A","N/A") if _is_below_m10(sm) else _get_weight(sm, qty)
            ws.cell(row=row,column=1,value=sm).alignment=center
            ws.cell(row=row,column=2,value=qty).alignment=center
            ws.cell(row=row,column=3,value=pcs).alignment=center
            ws.cell(row=row,column=4,value=wt).alignment=center
            for col in range(1,5): ws.cell(row=row,column=col).border=thin
            if isinstance(wt,float): total_stud_wt+=wt
            row+=1

        # Footer
        row+=2
        gf = PatternFill(start_color="C4D79B",end_color="C4D79B",fill_type="solid")
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=7)
        for col in range(1,8): ws.cell(row=row,column=col).fill=gf
        row+=2
        ws.cell(row=row,column=1,value="PREPARED").alignment=center
        ws.cell(row=row,column=4,value="CHECKED").alignment=center
        ws.cell(row=row,column=7,value="APPROVED").alignment=center
        row+=1
        ws.cell(row=row,column=1,value=CONFIG.get("prepared_by","RG")).alignment=center
        ws.cell(row=row,column=4,value=CONFIG.get("checked_by","JN")).alignment=center
        ws.cell(row=row,column=7,value=CONFIG.get("approved_by","JN")).alignment=center

        wb.save(excel_path)

        # Export PDF
        progress_cb("Step 6/6 — Exporting PDF...")
        pdf_out_final = ""
        if WIN32_OK:
            try:
                excel_com = win32com.client.Dispatch("Excel.Application")
                excel_com.Visible = False
                wb_com = excel_com.Workbooks.Open(os.path.abspath(excel_path))
                sheet  = wb_com.Worksheets(1)
                sheet.Columns.AutoFit(); sheet.Rows.AutoFit()
                sheet.PageSetup.Zoom=False; sheet.PageSetup.FitToPagesWide=1
                sheet.PageSetup.FitToPagesTall=1
                sheet.PageSetup.LeftMargin=excel_com.InchesToPoints(0.3)
                sheet.PageSetup.RightMargin=excel_com.InchesToPoints(0.3)
                sheet.PageSetup.TopMargin=excel_com.InchesToPoints(0.4)
                sheet.PageSetup.BottomMargin=excel_com.InchesToPoints(0.4)
                sheet.ExportAsFixedFormat(0, os.path.abspath(pdf_out))
                wb_com.Close(False); excel_com.Quit()
                pdf_out_final = pdf_out
            except Exception as e:
                progress_cb(f"PDF export failed: {e} — Excel is OK")
        else:
            progress_cb("win32com not installed — PDF skipped. Excel saved OK.")

        total_weight = round(total_bolt_wt + total_stud_wt, 2)

        stats = {
            "work_order"  : work_order,
            "customer"    : header["CUSTOMER"],
            "bolt_types"  : len(bolt_grouped),
            "stud_types"  : len(stud_grouped),
            "total_weight": total_weight,
            "excel_path"  : excel_path,
            "pdf_path"    : pdf_out_final,
            "folder_path" : folder_path,
        }

        if open_after:
            if pdf_out_final and os.path.exists(pdf_out_final):
                webbrowser.open(pdf_out_final)
            elif os.path.exists(excel_path):
                os.startfile(excel_path)

        done_cb(stats)

    except Exception as e:
        import traceback
        error_cb(f"{e}\n{traceback.format_exc()}")


# ============================================================
# SPLASH SCREEN
# ============================================================
def show_splash():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.configure(bg="#0A0C10")
    sw, sh = splash.winfo_screenwidth(), splash.winfo_screenheight()
    W, H   = 680, 380
    splash.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")
    splash.attributes("-alpha", 0.0)

    cv = tk.Canvas(splash, width=W, height=H, bg="#0A0C10", highlightthickness=0)
    cv.pack()
    for i in range(0,W,40): cv.create_line(i,0,i,H, fill="#0F1218", width=1)
    for i in range(0,H,40): cv.create_line(0,i,W,i, fill="#0F1218", width=1)
    for pts in [(0,0,55,0),(0,0,0,55),(W,H,W-55,H),(W,H,W,H-55)]:
        cv.create_line(*pts, fill=ACCENT, width=3)

    try:
        from PIL import Image, ImageTk
        import io as _io
        with open(LOGO_PATH,"rb") as _f: _b=_f.read()
        img = Image.open(_io.BytesIO(_b)).convert("RGBA").resize((190,50),Image.LANCZOS)
        photo = ImageTk.PhotoImage(img); cv._photo=photo
        cv.create_image(W//2,90,image=photo,anchor="center")
    except:
        cv.create_text(W//2,75,  text="INDO",font=("Consolas",42,"bold"),fill=WHITE, anchor="center")
        cv.create_text(W//2,125, text="TECH",font=("Consolas",42,"bold"),fill=ACCENT,anchor="center")

    cv.create_line(W//2-120,162,W//2+120,162, fill="#22283A",width=1)
    cv.create_text(W//2,178, text="HARDWARE AUTOMATION PLATFORM",
                   font=("Consolas",10,"bold"),fill="#6B7799",anchor="center")
    cv.create_rectangle(W//2-28,192,W//2+28,210, outline=ACCENT, fill="")
    cv.create_text(W//2,201, text="v 1.3",font=("Consolas",8),fill=ACCENT,anchor="center")

    bx1,bx2,by = 110,W-110,305
    cv.create_rectangle(bx1,by,bx2,by+4, fill="#181C24",outline="#22283A")
    st = cv.create_text(W//2,326, text="INITIALIZING...",font=("Consolas",8),fill="#3A4060",anchor="center")
    for i in range(11):
        splash.attributes("-alpha",i*0.1); splash.update(); time.sleep(0.04)
    try:
        import winsound
        winsound.PlaySound(WAV_PATH, winsound.SND_ASYNC)
    except: pass
    bf = None
    for prog, msg in [(0.2,"LOADING WEIGHT TABLE ENGINE..."),(0.4,"READING HARDWARE DATA..."),
                      (0.6,"INITIALIZING PDF PARSER..."),(0.8,"PREPARING EXCEL ENGINE..."),(1.0,"SYSTEM READY")]:
        if bf: cv.delete(bf)
        fx = bx1+int((bx2-bx1)*prog)
        bf = cv.create_rectangle(bx1,by,fx,by+4, fill=ACCENT,outline="")
        cv.itemconfig(st, text=msg); splash.update(); time.sleep(0.24)
    for i in range(10,-1,-1):
        splash.attributes("-alpha",i*0.1); splash.update(); time.sleep(0.03)
    splash.destroy()

# ============================================================
# WIDGETS
# ============================================================
class GlowButton(tk.Canvas):
    def __init__(self, parent, text, command=None,
                 width=240, height=48, color=ACCENT,
                 hover_color=ACCENT_GLOW, text_color=WHITE, bg=None, **kwargs):
        _bg = bg or BG_PANEL()
        super().__init__(parent, width=width, height=height,
                         bg=_bg, highlightthickness=0, **kwargs)
        self.text=text; self.command=command; self.color=color
        self.hover=hover_color; self.tc=text_color; self.w=width; self.h=height; self._bg=_bg
        self._draw(self.color)
        self.bind("<Enter>",          self._on_enter)
        self.bind("<Leave>",          self._on_leave)
        self.bind("<Button-1>",       self._on_press)
        self.bind("<ButtonRelease-1>",self._on_release)

    def _draw(self, bg, pressed=False):
        self.delete("all"); w,h,off = self.w,self.h,2 if pressed else 0
        self.create_rectangle(0,0,w,h, fill=bg,outline=bg)
        self.create_rectangle(0,0,4,h, fill=ACCENT_DARK if pressed else WHITE,outline="")
        self.create_text(w//2+4,h//2+off, text=self.text.upper(),
                         font=("Consolas",11,"bold"),fill=self.tc,anchor="center")
        self.create_text(w-16,h//2+off, text="›",font=("Consolas",18,"bold"),fill=self.tc,anchor="center")

    def _on_enter(self,e):  self._draw(self.hover); self.config(cursor="hand2")
    def _on_leave(self,e):  self._draw(self.color)
    def _on_press(self,e):  self._draw(ACCENT_DARK,True)
    def _on_release(self,e):self._draw(self.hover); self.command and self.command()

def get_coating_color(coating, all_coatings):
    try: idx=sorted(all_coatings).index(coating)
    except ValueError: idx=0
    return BADGE_PALETTE[idx % len(BADGE_PALETTE)]

class CoatingBadge(tk.Canvas):
    def __init__(self, parent, coating, active=False, all_coatings=None, badge_bg=None, **kwargs):
        all_coatings=all_coatings or [coating]; badge_bg=badge_bg or BG_PANEL()
        color = get_coating_color(coating,all_coatings) if active else FG_MUTED()
        label = f"{coating}: {'ON' if active else 'N/A'}"
        w = max(len(label)*8+28,76)
        super().__init__(parent, width=w, height=22, bg=badge_bg, highlightthickness=0, **kwargs)
        self.create_rectangle(0,0,w,22, fill=BG_DEEP(),outline=color,width=1)
        self.create_oval(6,7,14,15, fill=color,outline="")
        self.create_text(w//2+4,11, text=label.upper(),
                         font=("Consolas",8,"bold"),fill=color,anchor="center")

def make_stat_card(parent, label, value, color=ACCENT):
    f = tk.Frame(parent, bg=BG_CARD(), width=158, height=86)
    f.pack_propagate(False)
    tk.Frame(f, bg=color, height=3).pack(fill="x")
    val_lbl = tk.Label(f, text=value, font=("Consolas",22,"bold"), fg=color, bg=BG_CARD())
    val_lbl.pack(pady=(7,0))
    tk.Label(f, text=label.upper(), font=("Consolas",8), fg=FG2(), bg=BG_CARD()).pack()
    return f, val_lbl  # ← return label widget so we can update it

def add_log(widget, tag, message):
    widget.config(state="normal")
    ts = time.strftime("%H:%M:%S")
    prefix={"success":"✔ OK   ","error":"✖ ERR  ","warn":"⚠ WARN ",
            "info":"● INFO ","accent":"► SYS  "}.get(tag,"  ")
    widget.insert("end", f"[{ts}] {prefix} {message}\n", tag)
    widget.see("end"); widget.config(state="disabled")

def make_log_widget(parent):
    txt = tk.Text(parent, bg=BG_DEEP(), fg=FG2(), font=("Consolas",9),
                  relief="flat", insertbackground=ACCENT, state="normal", wrap="word", padx=10, pady=10)
    txt.tag_config("success",foreground=GREEN); txt.tag_config("error",foreground=RED)
    txt.tag_config("warn",foreground=YELLOW);  txt.tag_config("info",foreground=BLUE)
    txt.tag_config("accent",foreground=ACCENT);txt.tag_config("muted",foreground=FG_MUTED())
    return txt

def section_header(parent, title, color=ACCENT):
    row = tk.Frame(parent, bg=BG_CARD())
    row.pack(fill="x", padx=16, pady=(14,0))
    tk.Frame(row, bg=color, width=3, height=16).pack(side="left")
    tk.Label(row, text=f"  {title}", font=("Consolas",10,"bold"),
             fg=FG(), bg=BG_CARD()).pack(side="left")
    tk.Frame(parent, bg=BORDER(), height=1).pack(fill="x", pady=8)
    return row

# ============================================================
# MAIN APPLICATION
# ============================================================
def launch_main():
    root = TkinterDnD.Tk()
    root.title("IndoTech Hardware Platform  v1.3")
    root.geometry("1160x700")
    root.resizable(False, False)
    root.configure(bg=BG_DEEP())
    try: root.iconbitmap(ICO_PATH)
    except: pass

    # ── SIDEBAR ──────────────────────────────────────────────
    sidebar = tk.Frame(root, bg=BG_SIDEBAR(), width=220)
    sidebar.pack(side="left", fill="y"); sidebar.pack_propagate(False)

    logo_block = tk.Frame(sidebar, bg=ACCENT, height=72)
    logo_block.pack(fill="x"); logo_block.pack_propagate(False)
    try:
        from PIL import Image, ImageTk
        import io as _io
        with open(LOGO_PATH,"rb") as _f: _b=_f.read()
        img=Image.open(_io.BytesIO(_b)).convert("RGBA").resize((160,44),Image.LANCZOS)
        photo=ImageTk.PhotoImage(img); logo_block._photo=photo
        tk.Label(logo_block,image=photo,bg=ACCENT).place(relx=0.5,rely=0.5,anchor="center")
    except:
        tk.Label(logo_block,text="INDO TECH",font=("Consolas",14,"bold"),
                 fg=WHITE,bg=ACCENT).place(relx=0.5,rely=0.38,anchor="center")
        tk.Label(logo_block,text="HARDWARE PLATFORM",font=("Consolas",7),
                 fg="#FFE0C0",bg=ACCENT).place(relx=0.5,rely=0.75,anchor="center")

    tk.Frame(sidebar, bg=BORDER(), height=1).pack(fill="x")

    content_area = tk.Frame(root, bg=BG_DEEP())
    content_area.pack(side="right", fill="both", expand=True)

    pages = {}; coating_badge_frames = {}; _active_nav = [None]

    def make_topbar(parent, title):
        tb = tk.Frame(parent, bg=BG_PANEL(), height=52)
        tb.pack(fill="x"); tb.pack_propagate(False)
        tk.Label(tb, text=title, font=("Consolas",12,"bold"),
                 fg=FG(), bg=BG_PANEL()).pack(side="left", padx=20, pady=14)
        badge_row = tk.Frame(tb, bg=BG_PANEL())
        badge_row.pack(side="right", padx=20, pady=14)
        coating_badge_frames[title] = badge_row
        return badge_row

    def refresh_badges():
        coatings = sorted(app_state.get("coatings_loaded", set()))
        display  = coatings if coatings else ["NO TABLE"]
        for _, frame in coating_badge_frames.items():
            for w in frame.winfo_children(): w.destroy()
            for c in display:
                CoatingBadge(frame, c, active=(c in app_state["coatings_loaded"]),
                             all_coatings=display, badge_bg=BG_PANEL()).pack(side="left",padx=3)

    def navigate(page_name):
        app_state["current_page"] = page_name
        for name,(frame,_) in pages.items(): frame.pack_forget()
        if page_name in pages: pages[page_name][0].pack(fill="both",expand=True)

    def make_nav_item(icon, label, page_name):
        row = tk.Frame(sidebar, bg=BG_SIDEBAR(), height=46)
        row.pack(fill="x", pady=1); row.pack_propagate(False)
        bar = tk.Frame(row, bg=BG_SIDEBAR(), width=3); bar.pack(side="left",fill="y")
        ic  = tk.Label(row, text=icon, font=("Consolas",12), fg=FG_MUTED(), bg=BG_SIDEBAR(), width=3)
        ic.pack(side="left", padx=(8,4))
        lb  = tk.Label(row, text=label.upper(), font=("Consolas",9), fg=FG2(), bg=BG_SIDEBAR())
        lb.pack(side="left")
        def click(e=None):
            if _active_nav[0] and _active_nav[0] is not row:
                prev=_active_nav[0]
                for w in prev.winfo_children():
                    w.config(bg=BG_SIDEBAR())
                    if isinstance(w,tk.Label): w.config(fg=FG2() if w.cget("width")==0 else FG_MUTED())
                prev.config(bg=BG_SIDEBAR())
            row.config(bg=BG_CARD()); bar.config(bg=ACCENT)
            ic.config(fg=ACCENT,bg=BG_CARD()); lb.config(fg=WHITE,bg=BG_CARD())
            _active_nav[0]=row; navigate(page_name)
        row.bind("<Button-1>",click)
        for w in [bar,ic,lb]: w.bind("<Button-1>",click)
        row.bind("<Enter>",lambda e: row.config(bg=BG_CARD()) if _active_nav[0] is not row else None)
        row.bind("<Leave>",lambda e: row.config(bg=BG_CARD() if _active_nav[0] is row else BG_SIDEBAR()))
        row.config(cursor="hand2")
        return row, click

    tk.Frame(sidebar, bg=BORDER(), height=1).pack(fill="x", pady=4)
    nav_dash, click_dash = make_nav_item("⬡","Dashboard",    "dashboard")
    nav_gen,  click_gen  = make_nav_item("⬢","Generate List","generate")
    nav_wt,   click_wt   = make_nav_item("⬡","Weight Table", "weight")
    nav_hist, click_hist = make_nav_item("⬢","History",      "history")
    nav_act,  click_act  = make_nav_item("⬡","Activity Log", "activity")
    nav_set,  click_set  = make_nav_item("⬢","Settings",     "settings")

    tk.Frame(sidebar, bg=BORDER(), height=1).pack(side="bottom",fill="x")
    bot = tk.Frame(sidebar, bg=BG_SIDEBAR(), height=56)
    bot.pack(side="bottom",fill="x"); bot.pack_propagate(False)
    tk.Label(bot, text="◉  SYSTEM ACTIVE",font=("Consolas",8),fg=GREEN,bg=BG_SIDEBAR()).pack(pady=(10,2))
    tk.Label(bot, text="F/DSN/54  •  v1.3",font=("Consolas",7),fg=FG_MUTED(),bg=BG_SIDEBAR()).pack()

    _log_widgets = []
    def log(tag, message):
        app_state.setdefault("log_entries",[]).append((tag,message))
        for w in _log_widgets:
            try: add_log(w,tag,message)
            except: pass

    # ============================================================
    # PAGE: DASHBOARD
    # ============================================================
    dash_frame = tk.Frame(content_area, bg=BG_DEEP())
    pages["dashboard"] = (dash_frame, nav_dash)
    make_topbar(dash_frame, "HARDWARE LIST GENERATOR")
    tk.Frame(dash_frame, bg=BORDER(), height=1).pack(fill="x")

    dash_body = tk.Frame(dash_frame, bg=BG_DEEP())
    dash_body.pack(fill="both", expand=True, padx=20, pady=14)

    # ── Stat cards — keep references so we can UPDATE them ──
    stats_row = tk.Frame(dash_body, bg=BG_DEEP())
    stats_row.pack(fill="x", pady=(0,12))

    card_bolt, lbl_bolt  = make_stat_card(stats_row, "BOLT TYPES",   "—",    ACCENT)
    card_stud, lbl_stud  = make_stat_card(stats_row, "STUD TYPES",   "—",    BLUE)
    card_wt,   lbl_wt    = make_stat_card(stats_row, "TOTAL WEIGHT", "— KG", GREEN)
    card_wo,   lbl_wo    = make_stat_card(stats_row, "WORK ORDER",   "—",    YELLOW)
    for card in [card_bolt, card_stud, card_wt, card_wo]:
        card.pack(side="left", padx=(0,10))

    def update_stat_cards(stats):
        lbl_bolt.config(text=str(stats["bolt_types"]))
        lbl_stud.config(text=str(stats["stud_types"]))
        lbl_wt.config(text=f"{stats['total_weight']} KG" if stats["total_weight"] > 0 else "N/A KG")
        lbl_wo.config(text=f"WO-{stats['work_order']}")

    mid = tk.Frame(dash_body, bg=BG_DEEP())
    mid.pack(fill="both", expand=True)

    # ── Left panel ──
    left_p = tk.Frame(mid, bg=BG_CARD(), width=400)
    left_p.pack(side="left", fill="y", padx=(0,12))
    left_p.pack_propagate(False)
    section_header(left_p, "INPUT FILE", ACCENT)

    dz = tk.Canvas(left_p, width=368, height=148, bg=BG_DEEP(), highlightthickness=0)
    dz.pack(padx=16)

    def draw_dz(state="idle"):
        dz.delete("all"); W2,H2=368,148
        c={"idle":BORDER(),"hover":ACCENT,"loaded":GREEN,"error":RED}.get(state,BORDER())
        seg=14
        for x in range(0,W2,seg*2):
            dz.create_line(x,0,min(x+seg,W2),0,fill=c,width=2)
            dz.create_line(x,H2,min(x+seg,W2),H2,fill=c,width=2)
        for y in range(0,H2,seg*2):
            dz.create_line(0,y,0,min(y+seg,H2),fill=c,width=2)
            dz.create_line(W2,y,W2,min(y+seg,H2),fill=c,width=2)
        if state=="loaded":
            fname=os.path.basename(app_state["selected_pdf"])
            dz.create_text(W2//2,H2//2-18,text="✔",font=("Consolas",28,"bold"),fill=GREEN,anchor="center")
            dz.create_text(W2//2,H2//2+14,text=fname,font=("Consolas",9,"bold"),fill=GREEN,anchor="center")
            dz.create_text(W2//2,H2//2+32,text="click to change file",font=("Consolas",8),fill=FG_MUTED(),anchor="center")
        else:
            dz.create_text(W2//2,H2//2-18,text="⬡",font=("Consolas",26),
                           fill=ACCENT if state=="hover" else FG_MUTED(),anchor="center")
            dz.create_text(W2//2,H2//2+14,text="DROP PDF HERE",font=("Consolas",9,"bold"),
                           fill=ACCENT if state=="hover" else FG2(),anchor="center")
            dz.create_text(W2//2,H2//2+32,text="or click to browse",font=("Consolas",8),fill=FG_MUTED(),anchor="center")

    draw_dz("idle")

    def on_pdf_selected(path):
        path = path.strip().strip("{}")
        if not path.lower().endswith(".pdf"):
            draw_dz("error"); file_lbl.config(text="✖  Not a PDF file",fg=RED)
            log("error",f"Rejected (not PDF): {os.path.basename(path)}"); return
        app_state["selected_pdf"] = path
        draw_dz("loaded"); file_lbl.config(text=f"✔  {os.path.basename(path)}",fg=GREEN)
        log("success",f"PDF selected: {os.path.basename(path)}")

    def _open_pdf_dialog():
        path = filedialog.askopenfilename(parent=root, title="Select Hardware PDF",
                                          filetypes=[("PDF Files","*.pdf"),("All Files","*.*")])
        if path: on_pdf_selected(path)

    def browse_pdf(event=None): root.after(10, _open_pdf_dialog)

    dz.bind("<ButtonRelease-1>", browse_pdf)
    dz.bind("<Enter>", lambda e: draw_dz("hover") if not app_state["selected_pdf"] else None)
    dz.bind("<Leave>", lambda e: draw_dz("idle")  if not app_state["selected_pdf"] else None)
    dz.config(cursor="hand2")
    dz.drop_target_register(DND_FILES)
    dz.dnd_bind("<<Drop>>",      lambda e: on_pdf_selected(e.data))
    dz.dnd_bind("<<DragEnter>>", lambda e: draw_dz("hover"))
    dz.dnd_bind("<<DragLeave>>", lambda e: draw_dz("idle") if not app_state["selected_pdf"] else None)

    file_lbl = tk.Label(left_p, text="No file selected — drop or click above",
                        font=("Consolas",8), fg=FG_MUTED(), bg=BG_CARD())
    file_lbl.pack(pady=(4,0))

    tk.Button(left_p, text="⊕  BROWSE FOR PDF", font=("Consolas",9,"bold"),
              bg=BG_DEEP(), fg=ACCENT, activebackground=ACCENT, activeforeground=WHITE,
              relief="flat", cursor="hand2", bd=0, command=_open_pdf_dialog).pack(pady=(2,6))

    tk.Frame(left_p, bg=BORDER(), height=1).pack(fill="x", padx=16)

    opts = tk.Frame(left_p, bg=BG_CARD())
    opts.pack(fill="x", padx=20, pady=8)
    open_var     = tk.BooleanVar(value=CONFIG.get("open_after_gen",True))
    autosave_var = tk.BooleanVar(value=CONFIG.get("autosave_folder",False))
    cb_kw = dict(bg=BG_CARD(),fg=FG2(),selectcolor=BG_DEEP(),
                 activebackground=BG_CARD(),activeforeground=ACCENT,
                 font=("Consolas",9),highlightthickness=0,bd=0)
    tk.Checkbutton(opts, text="Open output after generation", variable=open_var, **cb_kw).pack(anchor="w")
    tk.Checkbutton(opts, text="Remember output folder (skip prompt next time)",
                   variable=autosave_var, **cb_kw).pack(anchor="w",pady=3)

    tk.Frame(left_p, bg=BORDER(), height=1).pack(fill="x", padx=16)

    pf = tk.Frame(left_p, bg=BG_CARD())
    pf.pack(fill="x", padx=20, pady=8)
    prog_lbl = tk.Label(pf, text="READY — select a PDF to begin",
                        font=("Consolas",8), fg=FG_MUTED(), bg=BG_CARD())
    prog_lbl.pack(anchor="w")
    style = ttk.Style(); style.theme_use("clam")
    style.configure("IT.Horizontal.TProgressbar",
                    background=ACCENT, troughcolor=BG_DEEP(),
                    bordercolor=BG_DEEP(), lightcolor=ACCENT, darkcolor=ACCENT_DARK)
    pbar = ttk.Progressbar(pf, style="IT.Horizontal.TProgressbar",
                           length=360, mode="indeterminate")
    pbar.pack(fill="x", pady=3)

    # ── GENERATE BUTTON — REAL IMPLEMENTATION ──────────────
    def do_generate():
        if not app_state["selected_pdf"]:
            messagebox.showwarning("No PDF","Please select an input PDF first.",parent=root); return
        if not PDF_OK:
            messagebox.showerror("Missing Library",
                "pdfplumber is not installed.\n\nRun: pip install pdfplumber",parent=root); return
        if not app_state.get("weight_table"):
            if not messagebox.askyesno("No Weight Table",
                "No weight table loaded.\n\nWeights will show N/A.\n\nContinue anyway?",parent=root):
                return

        if autosave_var.get() and CONFIG.get("last_output_folder"):
            out_folder = CONFIG["last_output_folder"]
            if not os.path.exists(out_folder):
                out_folder = filedialog.askdirectory(parent=root, title="Select Output Folder")
                if not out_folder: return
        else:
            out_folder = filedialog.askdirectory(parent=root, title="Select Output Folder")
            if not out_folder: return

        CONFIG["last_output_folder"] = out_folder
        save_config(CONFIG)

        # Lock UI
        prog_lbl.config(text="PROCESSING...", fg=ACCENT)
        pbar.start(10)

        def _progress(msg):
            root.after(0, lambda m=msg: (prog_lbl.config(text=m, fg=ACCENT), log("info", m)))

        def _done(stats):
            def _ui():
                pbar.stop()
                prog_lbl.config(text=f"✔  DONE — {stats['work_order']}", fg=GREEN)
                # ── Update stat cards ──────────────────────────
                update_stat_cards(stats)
                # ── Log results ───────────────────────────────
                log("success", f"Generated: WO-{stats['work_order']} | {stats['customer']}")
                log("success", f"Bolt types: {stats['bolt_types']}  Stud types: {stats['stud_types']}")
                log("success", f"Total weight: {stats['total_weight']} KG")
                log("success", f"Saved to: {stats['folder_path']}")
                if stats["pdf_path"]:
                    log("success", f"PDF: {os.path.basename(stats['pdf_path'])}")
                else:
                    log("warn", "PDF not created (win32com unavailable). Excel saved OK.")
                # ── Add to history ────────────────────────────
                app_state["history"].append(stats)
                refresh_history()
                # ── Show success popup ────────────────────────
                messagebox.showinfo("✔  Generation Complete",
                    f"Work Order : WO-{stats['work_order']}\n"
                    f"Customer   : {stats['customer']}\n"
                    f"Bolt types : {stats['bolt_types']}\n"
                    f"Stud types : {stats['stud_types']}\n"
                    f"Total wt   : {stats['total_weight']} KG\n\n"
                    f"Saved to:\n{stats['folder_path']}", parent=root)
            root.after(0, _ui)

        def _error(msg):
            def _ui():
                pbar.stop()
                prog_lbl.config(text="✖  ERROR — see log", fg=RED)
                log("error", f"Generation failed: {msg.splitlines()[0]}")
                messagebox.showerror("Generation Failed",
                    f"Error during generation:\n\n{msg.splitlines()[0]}\n\nSee Activity Log for details.",
                    parent=root)
            root.after(0, _ui)

        threading.Thread(
            target=run_export,
            args=(app_state["selected_pdf"], out_folder, open_var.get(),
                  _progress, _done, _error),
            daemon=True
        ).start()

    GlowButton(left_p, text="Generate Hardware List",
               command=do_generate, width=368, height=50, bg=BG_CARD()).pack(padx=16,pady=(0,14))

    # ── Right panel: log ──
    right_p = tk.Frame(mid, bg=BG_CARD())
    right_p.pack(side="right", fill="both", expand=True)
    section_header(right_p, "RECENT ACTIVITY", BLUE)
    dash_log = make_log_widget(right_p)
    dash_log.pack(fill="both", expand=True, padx=16, pady=(0,16))
    _log_widgets.append(dash_log)

    # ============================================================
    # PAGE: WEIGHT TABLE
    # ============================================================
    wt_frame = tk.Frame(content_area, bg=BG_DEEP())
    pages["weight"] = (wt_frame, nav_wt)
    make_topbar(wt_frame, "WEIGHT TABLE MANAGER")
    tk.Frame(wt_frame, bg=BORDER(), height=1).pack(fill="x")

    wt_body = tk.Frame(wt_frame, bg=BG_DEEP())
    wt_body.pack(fill="both", expand=True, padx=20, pady=14)

    up_card = tk.Frame(wt_body, bg=BG_CARD()); up_card.pack(fill="x", pady=(0,12))
    section_header(up_card, "WEIGHT TABLE EXCEL FILE", YELLOW)

    up_info = tk.Frame(up_card, bg=BG_CARD()); up_info.pack(fill="x", padx=16, pady=(0,10))
    wt_path_lbl = tk.Label(up_info,
                           text="No file loaded — click Upload or auto-loads from config",
                           font=("Consolas",9), fg=FG_MUTED(), bg=BG_CARD(), anchor="w")
    wt_path_lbl.pack(fill="x")
    wt_saved_lbl = tk.Label(up_info, text="", font=("Consolas",8), fg=BLUE, bg=BG_CARD(), anchor="w")
    wt_saved_lbl.pack(fill="x")

    tk.Label(tk.Frame(up_card, bg=BG_DEEP()).pack(fill="x",padx=16,pady=(0,12)) or up_card,
             text="ℹ  Upload once, click Save as Default — auto-loads every startup.",
             font=("Consolas",8), fg=BLUE, bg=BG_DEEP(), justify="left").pack(padx=10, pady=6, anchor="w") if False else None

    info_box = tk.Frame(up_card, bg=BG_DEEP()); info_box.pack(fill="x", padx=16, pady=(0,12))
    tk.Label(info_box,
             text="ℹ  Upload once → click Save as Default → auto-loads every startup.\n"
                  "   Re-upload only when weight rates change.",
             font=("Consolas",8), fg=BLUE, bg=BG_DEEP(), justify="left", wraplength=640).pack(padx=10, pady=8, anchor="w")

    wt_log_card = tk.Frame(wt_body, bg=BG_CARD()); wt_log_card.pack(fill="both", expand=True)
    section_header(wt_log_card, "LOAD LOG", GREEN)
    wt_log = make_log_widget(wt_log_card)
    wt_log.pack(fill="both", expand=True, padx=16, pady=(0,16))
    _log_widgets.append(wt_log)

    def do_auto_load_weight(path, silent=False):
        if not path or not os.path.exists(path):
            if not silent: add_log(wt_log,"warn","Saved weight table path not found.")
            return
        wt, coatings, msgs = load_weight_table_from_excel(path)
        if wt is None:
            add_log(wt_log,"error",f"Auto-load failed: {msgs[0][1] if msgs else 'unknown'}"); return
        app_state["weight_table"]=wt; app_state["weight_table_path"]=path
        app_state["coatings_loaded"]=coatings
        for tag,msg in msgs: add_log(wt_log,tag,msg)
        wt_path_lbl.config(text=f"✔  {os.path.basename(path)}",fg=GREEN)
        wt_saved_lbl.config(text=f"Saved default: {path}",fg=BLUE)
        refresh_badges(); log("success",f"Weight table loaded: {os.path.basename(path)}")

    def do_upload_weight():
        path = filedialog.askopenfilename(parent=root, title="Select Weight Table Excel",
                                          filetypes=[("Excel Files","*.xlsx *.xls")])
        if not path: return
        add_log(wt_log,"info",f"Loading: {os.path.basename(path)}")
        wt_path_lbl.config(text=f"Loading: {os.path.basename(path)}...",fg=ACCENT); root.update()
        wt, coatings, msgs = load_weight_table_from_excel(path)
        for tag,msg in msgs: add_log(wt_log,tag,msg)
        if wt is None:
            wt_path_lbl.config(text="✖  Invalid file — see log",fg=RED)
            messagebox.showerror("Invalid Weight Table",msgs[0][1] if msgs else "Error",parent=root); return
        app_state["weight_table"]=wt; app_state["weight_table_path"]=path
        app_state["coatings_loaded"]=coatings
        wt_path_lbl.config(text=f"✔  {os.path.basename(path)}",fg=GREEN)
        refresh_badges(); log("success",f"Weight table uploaded: {os.path.basename(path)}")
        if messagebox.askyesno("Save as Default?",
            f"Save as permanent default?\n\n{path}\n\nAuto-loads every startup.",parent=root):
            CONFIG["weight_table_path"]=path; save_config(CONFIG)
            wt_saved_lbl.config(text=f"✔  Saved: {os.path.basename(path)}",fg=GREEN)
            log("success","Weight table saved as default")

    btn_row = tk.Frame(up_card, bg=BG_CARD()); btn_row.pack(padx=16, pady=(0,14), anchor="w")
    GlowButton(btn_row, text="Upload Weight Table Excel",
               command=do_upload_weight, width=280, height=44, bg=BG_CARD()).pack(side="left",padx=(0,10))

    # ============================================================
    # PAGE: HISTORY (dynamic, updates after each generation)
    # ============================================================
    hist_frame = tk.Frame(content_area, bg=BG_DEEP())
    pages["history"] = (hist_frame, nav_hist)
    make_topbar(hist_frame, "GENERATION HISTORY")
    tk.Frame(hist_frame, bg=BORDER(), height=1).pack(fill="x")

    hist_body = tk.Frame(hist_frame, bg=BG_DEEP())
    hist_body.pack(fill="both", expand=True, padx=20, pady=14)
    hist_card = tk.Frame(hist_body, bg=BG_CARD()); hist_card.pack(fill="both", expand=True)
    section_header(hist_card, "PREVIOUSLY GENERATED WORK ORDERS", PURPLE)

    hist_hdr = tk.Frame(hist_card, bg=BG_DEEP()); hist_hdr.pack(fill="x", padx=16, pady=(0,4))
    for col in ["WORK ORDER","CUSTOMER","DATE","TOTAL WT","OPEN FOLDER"]:
        tk.Label(hist_hdr, text=col, font=("Consolas",8,"bold"),
                 fg=FG_MUTED(), bg=BG_DEEP(), width=16, anchor="w").pack(side="left",padx=4)
    tk.Frame(hist_card, bg=BORDER(), height=1).pack(fill="x", padx=16)

    hist_list = tk.Frame(hist_card, bg=BG_CARD()); hist_list.pack(fill="both", expand=True, padx=16)

    def refresh_history():
        for w in hist_list.winfo_children(): w.destroy()
        entries = app_state["history"]
        if not entries:
            tk.Label(hist_list, text="No generations yet this session.",
                     font=("Consolas",9), fg=FG_MUTED(), bg=BG_CARD()).pack(pady=20)
            return
        for stats in reversed(entries):
            row = tk.Frame(hist_list, bg=BG_CARD(), height=36)
            row.pack(fill="x", pady=1); row.pack_propagate(False)
            vals = [
                f"WO-{stats['work_order']}",
                stats["customer"][:18] if stats["customer"] else "—",
                datetime.now().strftime("%d.%m.%y"),
                f"{stats['total_weight']} KG" if stats["total_weight"]>0 else "N/A",
                "📂 Open",
            ]
            for i, val in enumerate(vals):
                clr = ACCENT if i==4 else (GREEN if i==0 else FG2())
                lbl = tk.Label(row, text=val, font=("Consolas",9),
                               fg=clr, bg=BG_CARD(), width=16, anchor="w")
                lbl.pack(side="left", padx=4, pady=8)
                if i == 4:
                    fp = stats["folder_path"]
                    lbl.config(cursor="hand2")
                    lbl.bind("<Button-1>", lambda e, p=fp: os.startfile(p) if os.path.exists(p) else None)
            row.bind("<Enter>", lambda e,r=row: r.config(bg=BG_DEEP()))
            row.bind("<Leave>", lambda e,r=row: r.config(bg=BG_CARD()))

    refresh_history()

    # ============================================================
    # PAGE: ACTIVITY LOG
    # ============================================================
    act_frame = tk.Frame(content_area, bg=BG_DEEP())
    pages["activity"] = (act_frame, nav_act)
    make_topbar(act_frame, "ACTIVITY LOG")
    tk.Frame(act_frame, bg=BORDER(), height=1).pack(fill="x")

    act_body = tk.Frame(act_frame, bg=BG_DEEP()); act_body.pack(fill="both",expand=True,padx=20,pady=14)
    act_card = tk.Frame(act_body, bg=BG_CARD()); act_card.pack(fill="both", expand=True)

    act_hdr = section_header(act_card, "SYSTEM LOG — ALL EVENTS", ACCENT)
    tk.Button(act_hdr, text="CLEAR LOG", font=("Consolas",8), bg=ACCENT_DARK, fg=WHITE,
              relief="flat", cursor="hand2",
              command=lambda: [act_log.config(state="normal"),
                               act_log.delete("1.0","end"),
                               act_log.config(state="disabled")]).pack(side="right")
    act_log = make_log_widget(act_card)
    act_log.pack(fill="both", expand=True, padx=16, pady=(0,16))
    _log_widgets.append(act_log)

    # ============================================================
    # PAGE: SETTINGS
    # ============================================================
    set_frame = tk.Frame(content_area, bg=BG_DEEP())
    pages["settings"] = (set_frame, nav_set)
    make_topbar(set_frame, "SETTINGS")
    tk.Frame(set_frame, bg=BORDER(), height=1).pack(fill="x")

    set_body = tk.Frame(set_frame, bg=BG_DEEP())
    set_body.pack(fill="both", expand=True, padx=20, pady=14)

    def settings_section(parent, title, color=ACCENT):
        card = tk.Frame(parent, bg=BG_CARD()); card.pack(fill="x", pady=(0,10))
        section_header(card, title, color)
        body = tk.Frame(card, bg=BG_CARD()); body.pack(fill="x", padx=16, pady=(0,12))
        return body

    def settings_row(parent, label, var_key, default=""):
        row = tk.Frame(parent, bg=BG_CARD()); row.pack(fill="x", pady=3)
        tk.Label(row, text=label, font=("Consolas",9), fg=FG2(), bg=BG_CARD(), width=24, anchor="w").pack(side="left")
        e = tk.Entry(row, font=("Consolas",9), bg=BG_DEEP(), fg=FG(),
                     insertbackground=ACCENT, relief="flat", width=36)
        e.insert(0, CONFIG.get(var_key, default)); e.pack(side="left", padx=(8,0), ipady=4)
        return e

    body_doc = settings_section(set_body, "DOCUMENT / APPROVAL", ACCENT)
    e_prep   = settings_row(body_doc, "Prepared By",   "prepared_by",  "RG")
    e_chk    = settings_row(body_doc, "Checked By",    "checked_by",   "JN")
    e_app    = settings_row(body_doc, "Approved By",   "approved_by",  "JN")
    e_fmt    = settings_row(body_doc, "Format Number", "format_number","F/DSN/54")

    body_out = settings_section(set_body, "OUTPUT FOLDER", BLUE)
    e_out    = settings_row(body_out, "Remembered Folder","last_output_folder","")

    body_theme = settings_section(set_body, "APPEARANCE", PURPLE)
    theme_lbl  = tk.Label(body_theme, text=f"Current theme: {CONFIG.get('theme','dark').upper()}",
                          font=("Consolas",9), fg=FG2(), bg=BG_CARD())
    theme_lbl.pack(anchor="w")

    def toggle_theme():
        new = "light" if CONFIG.get("theme","dark")=="dark" else "dark"
        apply_theme(new); theme_lbl.config(text=f"Current theme: {new.upper()}")
        messagebox.showinfo("Theme Changed",f"Theme set to {new.upper()}.\nRestart for full effect.",parent=root)

    GlowButton(body_theme, text="Toggle Dark / Light Theme",
               command=toggle_theme, width=260, height=40, bg=BG_CARD()).pack(anchor="w", pady=6)

    def save_settings():
        CONFIG["prepared_by"]=e_prep.get(); CONFIG["checked_by"]=e_chk.get()
        CONFIG["approved_by"]=e_app.get();  CONFIG["format_number"]=e_fmt.get()
        CONFIG["last_output_folder"]=e_out.get(); save_config(CONFIG)
        messagebox.showinfo("Saved","Settings saved ✔",parent=root); log("success","Settings saved")

    GlowButton(set_body, text="Save Settings", command=save_settings,
               width=200, height=44, bg=BG_DEEP()).pack(anchor="w", pady=4)

    # ============================================================
    # PAGE: GENERATE LIST (mirrors dashboard)
    # ============================================================
    gen_frame = tk.Frame(content_area, bg=BG_DEEP())
    pages["generate"] = (gen_frame, nav_gen)
    make_topbar(gen_frame, "GENERATE HARDWARE LIST")
    tk.Frame(gen_frame, bg=BORDER(), height=1).pack(fill="x")
    tk.Label(gen_frame,
             text="Use the Dashboard panel to generate.\n"
                  "All output appears here in Recent Activity and History.",
             font=("Consolas",10), fg=FG2(), bg=BG_DEEP()).pack(pady=40)
    GlowButton(gen_frame, text="Go to Dashboard", command=click_dash,
               width=220, height=44, bg=BG_DEEP()).pack()

    # ── STATUS BAR ───────────────────────────────────────────
    tk.Frame(root, bg=BORDER(), height=1).pack(side="bottom", fill="x")
    sb = tk.Frame(root, bg=BG_PANEL(), height=26)
    sb.pack(side="bottom", fill="x"); sb.pack_propagate(False)
    tk.Label(sb, text="◉  INDOTECH HARDWARE PLATFORM  v1.3",
             font=("Consolas",8), fg=GREEN, bg=BG_PANEL()).pack(side="left", padx=14, pady=5)
    sb_wt = tk.Label(sb, text="⚠  No weight table loaded",
                     font=("Consolas",8), fg=YELLOW, bg=BG_PANEL())
    sb_wt.pack(side="right", padx=14, pady=5)

    # ── STARTUP ──────────────────────────────────────────────
    saved_wt = CONFIG.get("weight_table_path","")
    if saved_wt and os.path.exists(saved_wt):
        root.after(400, lambda: do_auto_load_weight(saved_wt, silent=True))
        root.after(700, lambda: sb_wt.config(
            text=f"✔  Weight table: {os.path.basename(saved_wt)}", fg=GREEN))
    else:
        log("warn","No weight table loaded. Go to Weight Table → Upload.")

    log("accent","IndoTech Hardware Platform v1.3 — session started")
    log("info",  f"Config: {CONFIG_PATH}")
    log("info",  f"Base dir: {BASE_DIR}")

    click_dash()
    refresh_badges()
    root.mainloop()

# ============================================================
# ENTRY POINT
# ============================================================
show_splash()
launch_main()