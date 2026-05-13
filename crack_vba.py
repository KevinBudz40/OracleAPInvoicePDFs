"""
Extracts all VBA module source code from a password-protected .xlsm / .xls
without needing the password.

Requires:  pip install oletools
Usage:     python crack_vba.py  path\to\FBDI.xlsm
"""
import sys, os, zipfile, shutil, tempfile

def extract_vba(xlsm_path):
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        print("oletools not found.  Run:  pip install oletools")
        sys.exit(1)

    print(f"Extracting VBA from: {xlsm_path}\n")
    vba = VBA_Parser(xlsm_path)

    if not vba.detect_vba_macros():
        print("No VBA macros found in this file.")
        return

    out_dir = os.path.splitext(xlsm_path)[0] + "_VBA"
    os.makedirs(out_dir, exist_ok=True)

    for (filename, stream_path, vba_filename, vba_code) in vba.extract_macros():
        safe = vba_filename.replace("/", "_").replace("\\", "_")
        out_path = os.path.join(out_dir, safe + ".bas")
        with open(out_path, "w", encoding="utf-8", errors="replace") as f:
            f.write(f"' Source file : {filename}\n")
            f.write(f"' Stream      : {stream_path}\n")
            f.write(f"' Module      : {vba_filename}\n\n")
            f.write(vba_code)
        print(f"  Wrote: {out_path}")

    vba.close()
    print(f"\nAll modules saved to: {out_dir}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # Try to find any FBDI .xlsm in common locations
        candidates = []
        for root in [r"C:\Users", r"D:\\"]:
            for dirpath, _, files in os.walk(root):
                for f in files:
                    if "fbdi" in f.lower() and f.lower().endswith(".xlsm"):
                        candidates.append(os.path.join(dirpath, f))
                if len(candidates) >= 5:
                    break

        if candidates:
            print("Found possible FBDI files:")
            for c in candidates:
                print(f"  {c}")
            print("\nRun:  python crack_vba.py  <path>")
        else:
            print("Usage:  python crack_vba.py  path\\to\\FBDI.xlsm")
    else:
        extract_vba(sys.argv[1])
