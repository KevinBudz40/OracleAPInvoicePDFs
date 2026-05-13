"""
setup_workbook.py  —  Fully automates the AP Invoice PDF Downloader workbook setup.

What it does:
  1. Opens PayablesStandardInvoiceImportTemplate_unprotected.xlsm via Excel COM
  2. Renames the 2nd sheet  → "Parameters"
  3. Renames the 3rd sheet  → "Invoice_Log"
  4. Clears both sheets and adds all labels / headers
  5. Fills in default config values (host, project, dates, folder)
  6. Drops one test invoice number in A17 of Parameters
  7. Replaces Module1 code with NewModule1.bas
  8. Changes CommandButton1_Click on Sheet1 to call DownloadInvoicePDFs
  9. Saves as AP_Invoice_Downloader.xlsm

Requirements:
  pip install pywin32
  Excel must be installed on this machine.

  IMPORTANT — before running, enable VBA object model access in Excel:
    File → Options → Trust Center → Trust Center Settings
    → Macro Settings → [x] Trust access to the VBA project object model
"""

import os, sys, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
import win32com.client as win32

# ── file paths ────────────────────────────────────────────────────────────────
WORKBOOK  = r"D:\Nuke\PayablesStandardInvoiceImportTemplate_unprotected.xlsm"
OUT_FILE  = r"D:\Nuke\AP_Invoice_Downloader.xlsm"
VBA_FILE  = r"D:\Nuke\NewModule1.bas"
FORM_FILE = r"D:\Nuke\PasswordForm.frm"

# ── Parameters sheet contents ─────────────────────────────────────────────────
# Column A labels
PARAM_LABELS = {
    "A1":  "AP Invoice PDF Downloader — Configuration",
    "A3":  "Host URL",
    "A4":  "Username",
    # A5 intentionally omitted — password is prompted at runtime (masked)
    "A7":  "─── Date Range ───",
    "A8":  "Project Number",
    "A9":  "From Date",
    "A10": "To Date",
    "A12": "─── Output ───",
    "A13": "PDF Save Folder",
    "A15": "─── Invoice Numbers ───",
    "A16": "Invoice Number",          # header row
}

# Column B default values  (fill in user/pass manually)
PARAM_DEFAULTS = {
    "B3":  "https://eese.fa.us8.oraclecloud.com",
    "B8":  "22HALOPS",
    "B9":  "2026-04-01",
    "B10": "2026-04-30",
    "B13": r"D:\Nuke\invoice_pdfs",
}

TEST_INVOICE = "5523448494"   # placed in A17

# ── Invoice_Log headers (row 2) ───────────────────────────────────────────────
LOG_HEADERS = [
    "Invoice Number", "Invoice ID", "File Name",
    "Status", "Saved Path", "Timestamp"
]

# ── Replacement CommandButton1_Click ─────────────────────────────────────────
BUTTON_CODE = (
    "Private Sub CommandButton1_Click()\r\n"
    "    DownloadInvoicePDFs\r\n"
    "End Sub\r\n"
)


# ─────────────────────────────────────────────────────────────────────────────
def load_new_module_code(path):
    """Read NewModule1.bas, strip the Attribute VB_Name header line."""
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines()
    # COM CodeModule doesn't want the Attribute VB_Name line
    if lines and lines[0].startswith("Attribute VB_Name"):
        lines = lines[1:]
    return "\r\n".join(lines)


def rename_sheet_by_index(wb, one_based_index, new_name):
    """Rename sheet at 1-based index, skip if it already has the target name."""
    sh = wb.Sheets(one_based_index)
    if sh.Name != new_name:
        print(f"  Renaming sheet {one_based_index} '{sh.Name}' → '{new_name}'")
        sh.Name = new_name
    else:
        print(f"  Sheet {one_based_index} already named '{new_name}' — skipped")


def populate_parameters(wb):
    ws = wb.Sheets("Parameters")
    ws.Cells.ClearContents()

    for addr, val in PARAM_LABELS.items():
        ws.Range(addr).Value = val

    for addr, val in PARAM_DEFAULTS.items():
        ws.Range(addr).Value = val

    ws.Range("A17").Value = TEST_INVOICE
    print("  Parameters sheet populated.")


def populate_log(wb):
    ws = wb.Sheets("Invoice_Log")
    ws.Range("A2:F2").ClearContents()
    for col, hdr in enumerate(LOG_HEADERS, start=1):
        ws.Cells(2, col).Value = hdr
    print("  Invoice_Log headers set.")


def replace_module1(wb, new_code):
    """
    Inject Module1 VBA.  The DPx= unprotect trick causes Excel to silently
    wipe all VBA on open, so Module1 likely no longer exists — we create it.
    """
    vbp = wb.VBProject

    # Try to find an existing Module1; if absent, create a new standard module
    comp = None
    for i in range(1, vbp.VBComponents.Count + 1):
        c = vbp.VBComponents.Item(i)
        if c.Name == "Module1":
            comp = c
            break

    if comp is None:
        comp = vbp.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
        comp.Name = "Module1"
        print("  Module1 did not exist (VBA was wiped on open) — created fresh.")
    else:
        print("  Found existing Module1 — replacing code.")

    cm    = comp.CodeModule
    total = cm.CountOfLines
    if total > 0:
        cm.DeleteLines(1, total)
    cm.InsertLines(1, new_code)
    print("  Module1 code injected successfully.")


def import_password_form(wb):
    """
    Import PasswordForm.frm into the VBA project.
    Removes any existing PasswordForm first so re-running setup is safe.
    The .frm builds its controls entirely at runtime — no companion .frx needed.
    """
    vbp = wb.VBProject

    # Remove stale copy if present
    for i in range(1, vbp.VBComponents.Count + 1):
        comp = vbp.VBComponents.Item(i)
        if comp.Name == "PasswordForm":
            vbp.VBComponents.Remove(comp)
            print("  Removed existing PasswordForm.")
            break

    vbp.VBComponents.Import(os.path.abspath(FORM_FILE))
    print("  PasswordForm imported.")


def replace_button_handler(wb):
    """
    Find the Sheet1 VBComponent and swap CommandButton1_Click so it calls
    DownloadInvoicePDFs instead of GenCSV.
    """
    vbp = wb.VBProject

    # Sheet1 may be named "Sheet1" or have a display name like "Instructions"
    sheet1_comp = None
    for i in range(1, vbp.VBComponents.Count + 1):
        comp = vbp.VBComponents.Item(i)
        # Class modules for sheets have Type == 100
        if comp.Type == 100 and comp.Name in ("Sheet1", "Ark1"):
            sheet1_comp = comp
            break
    if sheet1_comp is None:
        # fall back: first class module (Type 100)
        for i in range(1, vbp.VBComponents.Count + 1):
            comp = vbp.VBComponents.Item(i)
            if comp.Type == 100:
                sheet1_comp = comp
                break

    if sheet1_comp is None:
        print("  WARNING: Could not locate Sheet1 class module — button not updated.")
        return

    cm    = sheet1_comp.CodeModule
    total = cm.CountOfLines

    # Search for the Sub that contains CommandButton1_Click
    start_line = 0
    end_line   = 0
    for i in range(1, total + 1):
        line_text = cm.Lines(i, 1)
        if "CommandButton1_Click" in line_text:
            start_line = i
        if start_line and i >= start_line and "End Sub" in line_text:
            end_line = i
            break

    if start_line and end_line:
        cm.DeleteLines(start_line, end_line - start_line + 1)
        cm.InsertLines(start_line, BUTTON_CODE)
        print(f"  CommandButton1_Click replaced (was lines {start_line}–{end_line}).")
    else:
        # No existing handler — append one
        insert_at = cm.CountOfLines + 1
        cm.InsertLines(insert_at, "\r\n" + BUTTON_CODE)
        print("  CommandButton1_Click appended (no existing handler found).")


# ─────────────────────────────────────────────────────────────────────────────
def main():
    for path, label in [(WORKBOOK, "source workbook"), (VBA_FILE, "VBA module"),
                        (FORM_FILE, "password form")]:
        if not os.path.isfile(path):
            sys.exit(f"ERROR: {label} not found:\n  {path}")

    # Unblock the XLSM file — removes the Zone.Identifier ADS that causes
    # Excel to open it in Protected View, which blocks COM .Open()
    import subprocess
    try:
        subprocess.run(
            ["powershell", "-Command",
             f"Unblock-File -Path '{os.path.abspath(WORKBOOK)}'"],
            check=True, capture_output=True
        )
        print("File unblocked (Zone.Identifier removed).")
    except Exception as e:
        print(f"Note: Unblock-File failed ({e}) — continuing anyway.")

    new_code = load_new_module_code(VBA_FILE)
    print(f"Loaded {len(new_code.splitlines())} lines from {os.path.basename(VBA_FILE)}")

    # DispatchEx always spawns a brand-new isolated Excel process,
    # so it never reuses (or disrupts) an Excel window you already have open.
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible       = False
    xl.DisplayAlerts = False

    # 3 = msoAutomationSecurityForceDisable — prevents any auto-open macro
    # from crashing the invisible session before we can set up the workbook.
    xl.AutomationSecurity = 3

    wb = None
    try:
        src = os.path.abspath(WORKBOOK)
        print(f"\nOpening: {src}")
        wb = xl.Workbooks.Open(src, UpdateLinks=0, ReadOnly=False)
        time.sleep(0.5)

        print("\n[1/6] Renaming sheets …")
        rename_sheet_by_index(wb, 2, "Parameters")
        rename_sheet_by_index(wb, 3, "Invoice_Log")

        print("\n[2/6] Populating Parameters sheet …")
        populate_parameters(wb)

        print("\n[3/6] Setting Invoice_Log headers …")
        populate_log(wb)

        print("\n[4/6] Replacing Module1 VBA …")
        try:
            replace_module1(wb, new_code)
        except Exception as e:
            print(f"  ERROR replacing Module1: {e}")
            print("  → Make sure 'Trust access to the VBA project object model' is ON")
            print("    Excel › File › Options › Trust Center › Trust Center Settings")
            print("    › Macro Settings › [x] Trust access to the VBA project object model")
            raise

        print("\n[5/7] Importing PasswordForm …")
        import_password_form(wb)

        print("\n[6/7] Updating CommandButton1_Click and button caption …")
        replace_button_handler(wb)

        # Set button caption directly via the OLEObjects COM interface
        try:
            btn = wb.Sheets(1).OLEObjects("CommandButton1").Object
            btn.Caption = "Download Invoice PDFs"
            print("  Button caption set to 'Download Invoice PDFs'")
        except Exception as e:
            print(f"  Note: could not rename button caption: {e}")

        print(f"\n[7/7] Saving as {OUT_FILE} …")
        # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        wb.SaveAs(OUT_FILE, FileFormat=52)
        wb.Close(False)
        wb = None

        print(f"\n✓  Done!  Open this file in Excel:\n   {OUT_FILE}")
        print("\nNext step: fill in Username (B4) on the Parameters sheet,")
        print("then click Download Invoice PDFs — you will be prompted for the password.")

    except Exception as e:
        print(f"\nFATAL: {e}")
        sys.exit(1)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        xl.DisplayAlerts = True
        xl.Quit()


if __name__ == "__main__":
    main()
