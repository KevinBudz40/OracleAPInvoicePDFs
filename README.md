# Oracle AP Invoice PDF Downloader

Downloads **scanned vendor invoice PDFs** from Oracle Fusion Cloud 26A Payables via REST API.
The tool is an Excel macro workbook — paste invoice numbers in, click a button, PDFs land in a folder.

---

## Quick start

1. **Download** `AP_Invoice_Downloader.xlsm` from this repo.
2. **Unblock** the file before opening: right-click → Properties → check **Unblock** → OK.
   *(Skipping this causes Excel to show a red "SECURITY RISK" bar that silently blocks all macros.)*
3. Open the workbook and **Enable Content** when prompted.
4. On the **Parameters** sheet fill in:
   - **B3** Host URL — e.g. `https://eese.fa.us8.oraclecloud.com`
   - **B4** Username
   - **B13** PDF Save Folder — e.g. `C:\invoice_pdfs`
5. Paste invoice numbers into **column A starting at row 17** (one per row).
6. Click **Download Invoice PDFs**.
7. Enter your Oracle password in the masked prompt that appears.
8. PDFs are saved to the folder; results are colour-coded in the **Invoice_Log** sheet.

---

## Parameters sheet layout

| Cell | Field | Notes |
|---|---|---|
| B3 | Host URL | Your Oracle Cloud pod base URL |
| B4 | Username | Oracle username |
| B8 | Project Number | Informational only — not used for filtering yet |
| B9 | From Date | `YYYY-MM-DD` |
| B10 | To Date | `YYYY-MM-DD` |
| B13 | PDF Save Folder | Created automatically if it doesn't exist |
| A17+ | Invoice Numbers | One per row; paste from OTBI or enter manually |

**Password** is never stored in the workbook. A masked prompt appears each time you click the button.

---

## Invoice_Log sheet

Each downloaded attachment gets one row:

| Column | Content |
|---|---|
| A | Invoice Number |
| B | Invoice ID (Oracle internal key) |
| C | Attachment file name |
| D | Status (colour-coded) |
| E | Full saved path |
| F | Timestamp |

Status colours: 🟢 green = saved successfully · 🔴 red = error · 🟡 yellow = no scanned image found

---

## How it works

### Step 1 — Resolve invoice number → invoice ID
Oracle's REST API identifies invoices internally by a numeric `InvoiceId`, not the human-readable
`InvoiceNumber`. For each number in the list the macro calls:

```
GET /fscmRestApi/resources/11.13.18.05/invoices
    ?q=InvoiceNumber='NNNN'&fields=InvoiceId,InvoiceNumber&limit=1
```

### Step 2 — Fetch attachments
```
GET /fscmRestApi/resources/11.13.18.05/invoices/{InvoiceId}/child/attachments
```

A `CategoryName` query filter on this endpoint returns HTTP 400, so all attachments are fetched
and filtered client-side by `Category == "Scanned Invoice Image"`.

### Step 3 — Download the binary
Each matching attachment has a `links` array. The entry with `name = "FileContents"` carries the
direct download URL (`rel = "enclosure"`). The macro streams that URL to an ADODB.Stream and
writes the raw bytes to disk.

### HTTP client
All requests use **`MSXML2.XMLHTTP.6.0`** (not `WinHttp`). MSXML2 inherits the Windows / IE proxy
settings automatically, which is required on corporate networks where WinHttp's separate proxy
configuration is not populated.

### JSON parsing
Oracle Fusion returns pretty-printed JSON with spaces around colons: `"key" : "value"`.
The workbook contains a hand-rolled parser (`JVal`, `ParseItems`) that handles:
- Whitespace around colons in key/value pairs
- Brace characters inside string values — the `DownloadInfo` field embeds a JSON blob as a
  string literal, which would otherwise fool a naïve depth counter into extracting items early.

### Password form
The password prompt is a VBA **UserForm** (`PasswordForm`) created programmatically at workbook
build time. Its `TextBox` uses `PasswordChar = "*"` for masking. The password is held in a local
variable for the duration of one button-click and is never written to any cell or file.

---

## Rebuilding the workbook from source

The `.xlsm` is pre-built and ready to use. If you want to rebuild it from scratch (e.g. after
editing `NewModule1.bas`):

**Prerequisites**
- Python 3.x
- `pip install pywin32`
- Excel installed on the same machine
- Excel setting: File → Options → Trust Center → Trust Center Settings → Macro Settings →
  ☑ **Trust access to the VBA project object model**
- The original Oracle FBDI template `PayablesStandardInvoiceImportTemplate_unprotected.xlsm`
  in `D:\Nuke\` (not in this repo — Oracle IP)

**Run**
```
python setup_workbook.py
```

The script:
1. Opens the FBDI template in a hidden, isolated Excel instance
2. Renames sheets → `Parameters`, `Invoice_Log`
3. Populates the Parameters sheet with labels and default values
4. Replaces Module1 code with `NewModule1.bas`
5. Creates the `PasswordForm` UserForm via `VBComponents.Add(3)`
   *(Using `Add(3)` rather than `Import()` is important — `Import()` registers a `.frm` file as
   a standard module, causing a compile error on the `Begin...End` designer block.)*
6. Updates the button caption and click handler
7. Saves as `AP_Invoice_Downloader.xlsm`

---

## Repository files

| File | Purpose |
|---|---|
| `AP_Invoice_Downloader.xlsm` | Ready-to-use workbook (no credentials stored) |
| `NewModule1.bas` | VBA source — all REST, JSON parsing, and download logic |
| `PasswordForm.frm` | UserForm source for reference (not imported directly; see setup notes above) |
| `setup_workbook.py` | Builds the `.xlsm` from scratch via Excel COM automation |
| `inject_diag.py` | Dev tool — injects a `DiagTest` sub for troubleshooting HTTP calls |
| `debug_jval.py` | Dev tool — Python port of `JVal`/`ParseItems` for offline JSON testing |
| `fix_button.py` | Dev tool — patches the button caption on an existing workbook |

---

## Known limitations / future work

- **Project-number auto-lookup** — the goal of populating invoice numbers automatically from a
  project number is deferred. The data exists in the `invoiceLineProjectDff` child resource
  (`_PROJECT_ID_Display` field) but there is no top-level endpoint that accepts a project filter.
  Candidate paths forward: a BIP report via SOAP (`ReportService`), or a new BIP data model
  with a SQL join on `PJC_EXP_ITEMS_ALL.ORIGINAL_HEADER_ID = AP_INVOICES_ALL.INVOICE_ID`.

- **SSO / token auth** — currently uses HTTP Basic Auth. Oracle's SSO stack (SAML / IDCS)
  requires a browser flow and cannot be driven from VBA directly.
