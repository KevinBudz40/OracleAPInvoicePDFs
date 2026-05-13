"""
Removes the VBA project password from an .xlsm / .xlam so you can
open and edit the code directly in the Excel VBA editor.

How it works:
  .xlsm files are ZIP archives.  Inside is xl/vbaProject.bin (a compound
  OLE document).  The password is stored as  DPB=<hash>  in the PROJECT
  stream.  Changing  DPB=  to  DPx=  makes Excel treat it as corrupted,
  prompt you to reset (just click OK / Yes to any dialogs), and give you
  full access to the modules.

Usage:
  python unprotect_vba.py  original.xlsm          → creates original_unprotected.xlsm
"""

import sys, os, zipfile, shutil, re, struct

def unprotect(src_path):
    if not os.path.isfile(src_path):
        print(f"File not found: {src_path}")
        sys.exit(1)

    base, ext = os.path.splitext(src_path)
    out_path  = base + "_unprotected" + ext
    shutil.copy2(src_path, out_path)

    BIN_NAME = "xl/vbaProject.bin"

    # ── work on a temp copy of the zip ───────────────────────────────────────
    tmp_zip = out_path + ".tmp.zip"
    shutil.copy2(out_path, tmp_zip)

    patched = False
    with zipfile.ZipFile(tmp_zip, "r") as zin, \
         zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == BIN_NAME:
                data, patched = patch_vba_bin(data)
            zout.writestr(item, data)

    os.remove(tmp_zip)

    if patched:
        print(f"Password removed.  Open this file in Excel:")
        print(f"  {out_path}")
        print()
        print("Excel will warn about a corrupted VBA project — click through")
        print("any dialogs (Yes / OK / Continue).  The password will be gone.")
    else:
        print("DPB record not found — file may not be password-protected,")
        print("or the vbaProject.bin is in a non-standard location.")
        print(f"Output (unchanged): {out_path}")


def patch_vba_bin(bin_data: bytes) -> tuple:
    """
    Scan the raw bytes of vbaProject.bin for the PROJECT stream and
    replace  DPB=  with  DPx=  to invalidate the password hash.
    Returns (patched_bytes, was_patched).
    """
    # The PROJECT stream contains ASCII text like:
    #   CMG=...
    #   DPB=<long hex string>
    #   GC=...
    # We patch the key name only — the hash value stays (won't validate).

    needle = b"DPB="
    if needle not in bin_data:
        return bin_data, False

    patched = bin_data.replace(needle, b"DPx=", 1)
    return patched, True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:  python unprotect_vba.py  path\\to\\FBDI.xlsm")
    else:
        unprotect(sys.argv[1])
