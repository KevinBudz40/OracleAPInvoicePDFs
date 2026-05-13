# Oracle AP Invoice PDF Downloader

Downloads **scanned vendor invoice PDFs** (`SCANNED_INVOICE_IMAGE` attachments) from Oracle Fusion Cloud 26A for a given project and date range.

## Two delivery formats

| Format | File | Description |
|---|---|---|
| Python script | `AP_Inv.py` | Command-line, schedulable via Task Scheduler |
| Excel macro workbook | `NewModule1.bas` | FBDI-style .xlsm, runs from a button |

---

## Python script (`AP_Inv.py`)

### Requirements
```
pip install requests
```

### Configuration
Edit the constants at the top of `AP_Inv.py`:

```python
BASE      = "https://<your-pod>.fa.us8.oraclecloud.com/fscmRestApi/resources/11.13.18.05"
AUTH      = HTTPBasicAuth("username", "password")   # move to env vars for production
DEST      = r"C:\path\to\pdf_output"
PROJECT   = "22HALOPS"
DAYS_BACK = 30
```

### Usage
```
python AP_Inv.py
```

PDFs are saved to `DEST` with the original attachment filename.

### How the project filter works
Oracle Fusion does not expose a direct REST endpoint to query AP invoices by project.
The script uses the `projectExpenditureItems` and `invoices` FSCM REST endpoints
(Basic Auth, `/fscmRestApi/resources/11.13.18.05`).

The physical join confirmed from the OTBI server log:
```sql
PJC_EXP_ITEMS_ALL.ORIGINAL_HEADER_ID = AP_INVOICES_ALL.INVOICE_ID
```

---

## Excel macro workbook

### Setup
1. Run `python unprotect_vba.py PayablesStandardInvoiceImportTemplate.xlsm`  
   to remove the VBA project password from Oracle's FBDI template.
2. Open the unprotected file and rename tabs:
   - Sheet2 → `Parameters`
   - Sheet3 → `Invoice_Log`
3. `Alt+F11` → Module1 → replace all code with contents of `NewModule1.bas`
4. In Sheet1's code change `GenCSV` → `DownloadInvoicePDFs`
5. Fill in the Parameters sheet (see layout below)

### Parameters sheet layout

| Row | Column A | Column B |
|---|---|---|
| 3 | Host URL | `https://eese.fa.us8.oraclecloud.com` |
| 4 | Username | |
| 5 | Password | |
| 8 | Project Number | `22HALOPS` |
| 9 | From Date | `2026-04-01` |
| 10 | To Date | `2026-04-30` |
| 13 | PDF Save Folder | `C:\invoice_pdfs` |
| 17+ | Invoice Numbers | *(one per row, paste from OTBI)* |

### Invoice_Log headers (row 2)
`Invoice Number | Invoice ID | File Name | Status | Saved Path | Timestamp`

---

## Utility scripts

| Script | Purpose |
|---|---|
| `crack_vba.py` | Extract VBA source from any password-protected .xlsm (read-only) |
| `unprotect_vba.py` | Remove VBA project password so the file is editable in Excel |
| `build_workbook.py` | Builds a standalone formatted .xlsm skeleton (alternative to FBDI approach) |

```
pip install oletools openpyxl
python crack_vba.py     path\to\file.xlsm
python unprotect_vba.py path\to\file.xlsm
python build_workbook.py
```

---

## Security note
Credentials are currently hardcoded in `AP_Inv.py` and stored in the Excel workbook.  
For production use, move credentials to environment variables or Windows Credential Manager.
