"""Rename the CommandButton1 caption and also update the Instructions sheet
description text to match the new purpose."""
import os, sys, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32com.client as win32

PATH = r"D:\Nuke\AP_Invoice_Downloader.xlsm"

xl = win32.DispatchEx("Excel.Application")
xl.Visible       = False
xl.DisplayAlerts = False
xl.AutomationSecurity = 3

wb = xl.Workbooks.Open(os.path.abspath(PATH), UpdateLinks=0, ReadOnly=False)
xl.AutomationSecurity = 1

ws = wb.Sheets(1)

# The ActiveX CommandButton caption lives in the VBProject as a property.
# Easiest route: set it via the VBA object model.
vbp = wb.VBProject
# Find Sheet1's class module and inject a caption assignment into Workbook_Open,
# OR just set it directly through the OLE object's property bag.
# Simpler: run a tiny macro that sets the caption.

# Add a helper sub, run it, then remove it.
mod = vbp.VBComponents.Add(1)   # standard module
mod.Name = "TempCapFix"
mod.CodeModule.InsertLines(1,
    "Sub FixCaption()\r\n"
    "    ThisWorkbook.Sheets(1).OLEObjects(\"CommandButton1\").Object.Caption"
    " = \"Download Invoice PDFs\"\r\n"
    "End Sub\r\n"
)

xl.AutomationSecurity = 1
xl.Application.Run("TempCapFix.FixCaption")
print("Caption updated.")

# Remove the temp module
vbp.VBComponents.Remove(mod)

wb.Save()
wb.Close(False)
xl.Quit()
print("Done.")
