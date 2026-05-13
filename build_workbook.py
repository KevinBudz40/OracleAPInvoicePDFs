"""
Creates the AP_Invoice_Downloader.xlsm skeleton with formatted sheets.
Run once; then open the file and import VBA_Module.bas via
  Alt+F11 → File → Import File.
Requires:  pip install openpyxl
"""
import openpyxl
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
import os

OUT = r"D:\Nuke\AP_Invoice_Downloader.xlsm"

# ── helpers ───────────────────────────────────────────────────────────────────
ORACLE_BLUE   = "00457C"
ORACLE_RED    = "C74634"
HEADER_GREY   = "F2F2F2"
WHITE         = "FFFFFF"
LIGHT_BLUE    = "D9E8F5"

def hdr_font(size=11, bold=True, color=WHITE):
    return Font(name="Calibri", size=size, bold=bold, color=color)

def body_font(size=10, bold=False, color="000000"):
    return Font(name="Calibri", size=size, bold=bold, color=color)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def set_col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width

# ── workbook ──────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# ═══════════════════════════════════════════════════════════════════════════════
#  Sheet 1 — Instructions
# ═══════════════════════════════════════════════════════════════════════════════
ws_inst = wb.active
ws_inst.title = "Instructions"
ws_inst.sheet_view.showGridLines = False

# banner
ws_inst.merge_cells("A1:G1")
ws_inst["A1"] = "AP Invoice PDF Downloader"
ws_inst["A1"].font      = Font(name="Calibri", size=20, bold=True, color=WHITE)
ws_inst["A1"].fill      = fill(ORACLE_BLUE)
ws_inst["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_inst.row_dimensions[1].height = 40

ws_inst.merge_cells("A2:G2")
ws_inst["A2"] = "Downloads scanned invoice PDFs from Oracle Fusion Cloud for a given project and date range"
ws_inst["A2"].font      = Font(name="Calibri", size=11, color=WHITE)
ws_inst["A2"].fill      = fill(ORACLE_RED)
ws_inst["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws_inst.row_dimensions[2].height = 22

steps = [
    ("", ""),
    ("Step 1", "Go to the Config sheet and fill in your Oracle Cloud URL, username, password,\n"
               "project number, date range, and the folder where PDFs should be saved."),
    ("", ""),
    ("Step 2 — Option A (Auto)", "Click  Get Invoice List from BIP  to automatically retrieve all\n"
               "AP invoices charged to the project for the date range.\n"
               "Requires a BIP report saved at the path shown in Config."),
    ("", ""),
    ("Step 2 — Option B (Manual)", "Go to the  Invoice_List  sheet and paste invoice numbers\n"
               "in column A (one per row, no header).  Use this if you already\n"
               "have a list from OTBI or another report."),
    ("", ""),
    ("Step 3", "Click  Download Invoice PDFs  on the Config sheet.\n"
               "Progress is written to the  Invoice_Log  sheet in real time.\n"
               "Green = saved successfully.  Red = error.  Yellow = no attachment found."),
    ("", ""),
    ("Notes", "• Only attachments with category  SCANNED_INVOICE_IMAGE  are downloaded.\n"
               "• If a PDF already exists in the save folder it will be overwritten.\n"
               "• Credentials are stored in the workbook — keep the file secure.\n"
               "• The workbook requires macros to be enabled."),
]

for r, (label, text) in enumerate(steps, start=4):
    cell_l = ws_inst.cell(row=r, column=2, value=label)
    cell_t = ws_inst.cell(row=r, column=3, value=text)
    if label.startswith("Step"):
        cell_l.font = Font(name="Calibri", size=10, bold=True, color=ORACLE_BLUE)
    elif label == "Notes":
        cell_l.font = Font(name="Calibri", size=10, bold=True, color=ORACLE_RED)
    else:
        cell_l.font = body_font()
    cell_t.font      = body_font()
    cell_t.alignment = Alignment(wrap_text=True, vertical="top")
    ws_inst.row_dimensions[r].height = 42 if "\n" in text else 18

set_col_width(ws_inst, "B", 28)
set_col_width(ws_inst, "C", 70)
for c in ["A","D","E","F","G"]:
    set_col_width(ws_inst, c, 4)

# ═══════════════════════════════════════════════════════════════════════════════
#  Sheet 2 — Config
# ═══════════════════════════════════════════════════════════════════════════════
ws_cfg = wb.create_sheet("Config")
ws_cfg.sheet_view.showGridLines = False

# banner
ws_cfg.merge_cells("A1:F1")
ws_cfg["A1"] = "AP Invoice PDF Downloader — Configuration"
ws_cfg["A1"].font      = Font(name="Calibri", size=16, bold=True, color=WHITE)
ws_cfg["A1"].fill      = fill(ORACLE_BLUE)
ws_cfg["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
ws_cfg.row_dimensions[1].height = 36

# section header helper
def section(ws, row, title):
    ws.merge_cells(f"B{row}:E{row}")
    c = ws.cell(row=row, column=2, value=title)
    c.font      = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c.fill      = fill("5B9BD5")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 18

# label+input helper
def param_row(ws, row, label, default="", note=""):
    ws.row_dimensions[row].height = 20
    lbl = ws.cell(row=row, column=2, value=label)
    lbl.font      = body_font(bold=True)
    lbl.fill      = fill(HEADER_GREY)
    lbl.alignment = Alignment(horizontal="right", vertical="center")
    lbl.border    = thin_border()

    inp = ws.cell(row=row, column=3, value=default)
    inp.font      = body_font()
    inp.fill      = fill(WHITE)
    inp.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    inp.border    = thin_border()

    if note:
        n = ws.cell(row=row, column=5, value=note)
        n.font      = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
        n.alignment = Alignment(vertical="center")

# Connection
section(ws_cfg, 3, "Oracle Cloud Connection")
param_row(ws_cfg, 4,  "Host URL",
          "https://eese.fa.us8.oraclecloud.com",
          "No trailing slash")
param_row(ws_cfg, 5,  "Username",   "Kevin.a.Budziszewski")
param_row(ws_cfg, 6,  "Password",   "",
          "Stored in workbook — keep file secure")

# Filter
section(ws_cfg, 8, "Filter Parameters")
param_row(ws_cfg, 9,  "Project Number", "22HALOPS")
param_row(ws_cfg, 10, "From Date",      "2026-04-01",  "YYYY-MM-DD")
param_row(ws_cfg, 11, "To Date",        "2026-04-30",  "YYYY-MM-DD")

# Output
section(ws_cfg, 13, "Output")
param_row(ws_cfg, 14, "PDF Save Folder",
          r"D:\Nuke\invoice_pdfs",
          "Folder is created if it does not exist")

# BIP
section(ws_cfg, 16, "BIP Report Path (for Auto mode)")
param_row(ws_cfg, 17, "BIP Report Path",
          "/Custom/AP_Invoice_Attachments/AP_Invoice_Attachments_DM",
          "Leave blank to use Manual mode (Invoice_List sheet)")

# Button placeholders (actual buttons added via VBA at runtime or manually)
ws_cfg.merge_cells("B19:C19")
btn1 = ws_cfg.cell(row=19, column=2,
                   value="▶  Get Invoice List from BIP  (click after enabling macros)")
btn1.font      = Font(name="Calibri", size=10, bold=True, color=WHITE)
btn1.fill      = fill("70AD47")
btn1.alignment = Alignment(horizontal="center", vertical="center")
ws_cfg.row_dimensions[19].height = 24

ws_cfg.merge_cells("B21:C21")
btn2 = ws_cfg.cell(row=21, column=2,
                   value="▶  Download Invoice PDFs  (click after enabling macros)")
btn2.font      = Font(name="Calibri", size=10, bold=True, color=WHITE)
btn2.fill      = fill(ORACLE_RED)
btn2.alignment = Alignment(horizontal="center", vertical="center")
ws_cfg.row_dimensions[21].height = 24

ws_cfg.merge_cells("B23:E24")
note = ws_cfg.cell(row=23, column=2,
    value="After enabling macros, use  Insert → Button  to link the green button to "
          "macro  GetInvoiceListFromBIP  and the red button to  DownloadInvoicePDFs.")
note.font      = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
note.alignment = Alignment(wrap_text=True, vertical="top")

set_col_width(ws_cfg, "A", 3)
set_col_width(ws_cfg, "B", 22)
set_col_width(ws_cfg, "C", 38)
set_col_width(ws_cfg, "D", 4)
set_col_width(ws_cfg, "E", 45)
set_col_width(ws_cfg, "F", 3)

# ═══════════════════════════════════════════════════════════════════════════════
#  Sheet 3 — Invoice_List
# ═══════════════════════════════════════════════════════════════════════════════
ws_inv = wb.create_sheet("Invoice_List")
ws_inv.sheet_view.showGridLines = False

ws_inv.merge_cells("A1:C1")
ws_inv["A1"] = "Invoice List  (auto-populated by BIP, or paste manually)"
ws_inv["A1"].font      = Font(name="Calibri", size=12, bold=True, color=WHITE)
ws_inv["A1"].fill      = fill(ORACLE_BLUE)
ws_inv["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
ws_inv.row_dimensions[1].height = 28

headers = ["Invoice Number", "Invoice Date", "Supplier Name"]
for col, h in enumerate(headers, start=1):
    c = ws_inv.cell(row=2, column=col, value=h)
    c.font      = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c.fill      = fill("5B9BD5")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = thin_border()
ws_inv.row_dimensions[2].height = 18

# Sample row (will be cleared / overwritten by macro)
sample = ["5523448494", "2026-03-15", "Airgas USA LLC"]
for col, v in enumerate(sample, start=1):
    c = ws_inv.cell(row=3, column=col, value=v)
    c.font   = body_font()
    c.fill   = fill(LIGHT_BLUE)
    c.border = thin_border()

set_col_width(ws_inv, "A", 20)
set_col_width(ws_inv, "B", 16)
set_col_width(ws_inv, "C", 35)

# ═══════════════════════════════════════════════════════════════════════════════
#  Sheet 4 — Invoice_Log
# ═══════════════════════════════════════════════════════════════════════════════
ws_log = wb.create_sheet("Invoice_Log")
ws_log.sheet_view.showGridLines = False

ws_log.merge_cells("A1:G1")
ws_log["A1"] = "Download Log"
ws_log["A1"].font      = Font(name="Calibri", size=12, bold=True, color=WHITE)
ws_log["A1"].fill      = fill(ORACLE_BLUE)
ws_log["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
ws_log.row_dimensions[1].height = 28

log_headers = ["Invoice Number", "Invoice ID", "File Name",
               "Status", "Saved Path", "File Size", "Timestamp"]
for col, h in enumerate(log_headers, start=1):
    c = ws_log.cell(row=2, column=col, value=h)
    c.font      = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c.fill      = fill("5B9BD5")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = thin_border()
ws_log.row_dimensions[2].height = 18

log_widths = [20, 16, 40, 32, 55, 14, 20]
for i, w in enumerate(log_widths, start=1):
    ws_log.column_dimensions[get_column_letter(i)].width = w

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save(OUT)
print(f"Workbook created: {OUT}")
print()
print("Next steps:")
print("  1. Open AP_Invoice_Downloader.xlsm and ENABLE MACROS.")
print("  2. Press Alt+F11 to open the VBA editor.")
print("  3. In the editor: Insert → Module")
print("  4. Paste the entire contents of VBA_Module.bas into the new module.")
print("  5. Close the VBA editor (Alt+Q).")
print("  6. Add form buttons on the Config sheet linked to:")
print("       GetInvoiceListFromBIP   (green cell B19)")
print("       DownloadInvoicePDFs     (red cell B21)")
print("  7. Fill in your password in Config row 6, then click a button.")
