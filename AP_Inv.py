import requests, os, json
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta

# ── Config ────────────────────────────────────────────────────────────────────
BASE       = "https://eese.fa.us8.oraclecloud.com/fscmRestApi/resources/11.13.18.05"
AUTH       = HTTPBasicAuth("Kevin.a.Budziszewski", "GaryTheCat2929!")   # TODO: env vars
DEST       = r"D:\Nuke\invoice_pdfs"
PROJECT    = "22HALOPS"
DAYS_BACK  = 30
START_DATE = (datetime.now() - timedelta(days=DAYS_BACK)).strftime("%Y-%m-%d")

os.makedirs(DEST, exist_ok=True)

# ── Step 1 — get invoice IDs via projectExpenditureItems FSCM REST endpoint ───
#
# OTBI REST API (/analytics/pub/v1/) requires browser-session auth; Basic Auth
# is only accepted by the FSCM REST API (/fscmRestApi/).
#
# PJC_EXP_ITEMS_ALL.ORIGINAL_HEADER_ID = AP_INVOICES_ALL.INVOICE_ID (confirmed
# from the OTBI server log).  The FSCM REST resource exposes this as
# OriginalHeaderId (or similar) on projectExpenditureItems.
#
# Phase A: probe the endpoint with a tiny result set to discover field names.
# Phase B: full query to collect all OriginalHeaderIds for the project / date.

# ── Probe both candidate endpoints to find usable field names ─────────────────
for endpoint in ("projectCosts", "projectExpenditureItems"):
    print(f"\n{'='*60}")
    print(f"Probing /{endpoint} (no filter, limit=1) ...")
    resp = requests.get(f"{BASE}/{endpoint}", auth=AUTH, params={"limit": 1})
    print(f"HTTP {resp.status_code}")
    if resp.status_code != 200:
        print(resp.text[:400])
        continue
    items = resp.json().get("items", [])
    if not items:
        print("  → 0 items returned")
        continue
    print(f"  → {len(items)} item(s).  Fields:")
    for k, v in items[0].items():
        if k != "links":          # skip the noisy links blob
            print(f"    {k!r}: {v!r}")

exit(0)   # ← diagnostic exit — remove once we identify the right fields

if not invoice_nums:
    print("No invoices for this project / date range.")
    exit(0)

# ── Step 2 / 3 — fetch attachments and download ───────────────────────────────

def get_attachments(invoice_id):
    r = requests.get(
        f"{BASE}/invoices/{invoice_id}/child/attachments",
        auth=AUTH,
        params={"q": "CategoryName='SCANNED_INVOICE_IMAGE'"},
    )
    if r.status_code != 200:
        print(f"  Attachments HTTP {r.status_code}: {r.text[:200]}")
        return []
    return r.json().get("items", [])

def download_attachment(att, inv_num):
    file_url = next(
        (l["href"] for l in att.get("links", []) if l.get("name") == "FileContents"),
        None,
    )
    if not file_url:
        print("  No FileContents link — skipping")
        return
    filename = att.get("FileName") or f"{inv_num}_{att.get('AttachedDocumentId')}.pdf"
    safe     = filename.replace("/", "_").replace("\\", "_")
    path     = os.path.join(DEST, safe)
    content  = requests.get(file_url, auth=AUTH).content
    with open(path, "wb") as f:
        f.write(content)
    print(f"  Saved: {safe}  ({len(content):,} bytes)")

# ── Main loop ─────────────────────────────────────────────────────────────────

for inv_id in invoice_ids:
    inv_num = resolve_invoice_num(inv_id)
    print(f"\nInvoice {inv_num}  (ID {inv_id})")
    atts = get_attachments(inv_id)
    if not atts:
        print("  No SCANNED_INVOICE_IMAGE attachments")
        continue
    for att in atts:
        download_attachment(att, inv_num)

print("\nDone.")
