"""
apply_v85_to_v841.py
====================
Injects the v8.5 VBA modules into the v8.4.1-with-BUG.xlsm workbook.

Standard modules  → removed and re-imported from build/v8_5_vba/
Document modules  → code replaced in-place via CodeModule API
                    (Sheet8, Sheet11, Sheet12 — the three that differ)

Output: Global-Smile_2026-v8.5-applied.xlsm  (sibling of the source file)
"""

import os, sys, time, threading
import win32com.client, win32con, win32gui, win32process
from pywintypes import com_error

ROOT   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_WB = os.path.join(ROOT, "Global-Smile_2026-v8.4.1-with-BUG.xlsm")
OUT_WB = os.path.join(ROOT, "Global-Smile_2026-v8.5-applied.xlsm")
VBA85  = os.path.join(ROOT, "build", "v8_5_vba")

# Standard modules to fully replace (remove + import)
STANDARD_MODULES = [
    "modGLGate",
    "modGLRefresh",
    "modGLScanEngine",
    "modGLWorkbookMap",
    "modEngine",
    "modPJSync",
    "modGraphify",
    "ImportSuppliersFromJSON",
]

# Document modules: codename -> .txt filename (only the ones that changed)
DOCUMENT_MODULES = {
    "Sheet8":  "Sheet8.txt",
    "Sheet11": "Sheet11.txt",
    "Sheet12": "Sheet12.txt",
}

LOG = []

def log(msg):
    print(msg)
    LOG.append(msg)

def read_src(filename):
    path = os.path.join(VBA85, filename)
    with open(path, "rb") as f:
        raw = f.read()
    text = raw.decode("utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()

# ── watchdog: close any Excel modal dialog that blocks compile ──────────────
_done   = threading.Event()
_pid    = None

def _watchdog():
    if _done.wait(90):
        return
    log("WATCHDOG: compile timed out — attempting to close dialogs")
    def closer(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            cls = win32gui.GetClassName(hwnd)
            if cls in ("#32770", "ThunderDFrame"):
                _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                if wpid == _pid:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        return True
    try:
        win32gui.EnumWindows(closer, None)
    except Exception as e:
        log(f"  watchdog enum error: {e}")

def main():
    global _pid

    if not os.path.isfile(SRC_WB):
        sys.exit(f"Source workbook not found: {SRC_WB}")

    # Verify all source files exist
    missing = []
    for m in STANDARD_MODULES:
        p = os.path.join(VBA85, m + ".bas")
        if not os.path.isfile(p):
            missing.append(p)
    for cn, fname in DOCUMENT_MODULES.items():
        p = os.path.join(VBA85, fname)
        if not os.path.isfile(p):
            missing.append(p)
    if missing:
        sys.exit("Missing v8.5 source files:\n  " + "\n  ".join(missing))

    log(f"Source : {SRC_WB}")
    log(f"Output : {OUT_WB}")
    log(f"VBA src: {VBA85}")

    excel = win32com.client.DispatchEx("Excel.Application")
    _pid  = excel.Application.Hwnd  # used by watchdog
    wb    = None

    try:
        excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
        excel.DisplayAlerts      = False
        excel.Visible            = False
        excel.EnableEvents       = False
        excel.ScreenUpdating     = False

        log("\nOpening workbook …")
        wb = excel.Workbooks.Open(
            SRC_WB, ReadOnly=False, UpdateLinks=0,
            IgnoreReadOnlyRecommended=True
        )
        vbp = wb.VBProject

        # ── 1. Replace standard modules ─────────────────────────────────────
        log("\n── Standard modules ──────────────────────────────")
        existing = {c.Name: c for c in vbp.VBComponents}

        for mod in STANDARD_MODULES:
            bas = mod + ".bas"
            src_path = os.path.join(VBA85, bas)
            # Remove old component if it exists
            if mod in existing:
                vbp.VBComponents.Remove(existing[mod])
                log(f"  removed : {mod}")
            # Import fresh from v8.5
            vbp.VBComponents.Import(src_path)
            log(f"  imported: {mod}  ({os.path.getsize(src_path):,} bytes)")

        # ── 2. Replace document module code (in-place) ──────────────────────
        log("\n── Document modules ──────────────────────────────")
        # Rebuild component map after imports
        comp_map = {c.Name: c for c in vbp.VBComponents}

        for codename, fname in DOCUMENT_MODULES.items():
            if codename not in comp_map:
                log(f"  SKIP {codename}: not found in workbook")
                continue
            cm   = comp_map[codename].CodeModule
            new_code = read_src(fname)
            # Strip Attribute VB_Name header if present — Excel owns it in doc modules
            lines = new_code.splitlines()
            lines = [l for l in lines if not l.startswith("Attribute VB_Name")]
            new_code = "\n".join(lines).strip()
            # Clear existing code and write new
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.InsertLines(1, new_code)
            log(f"  replaced: {codename} ({fname})  {cm.CountOfLines} lines written")

        # ── 3. Compile ───────────────────────────────────────────────────────
        log("\n── Compile ───────────────────────────────────────")
        wt = threading.Thread(target=_watchdog, daemon=True)
        wt.start()
        try:
            excel.Application.VBE.CommandBars("Menu Bar").Controls("Debug").Controls("Compile VBAProject").Execute()
            _done.set()
            log("  Compile: OK (no dialog appeared)")
        except com_error as e:
            _done.set()
            log(f"  Compile: COM error — {e}")
            log("  (A compile error dialog may have appeared; check the workbook in Excel)")

        # ── 4. Save as new file ──────────────────────────────────────────────
        log(f"\n── Saving → {os.path.basename(OUT_WB)} ─────────────")
        if os.path.isfile(OUT_WB):
            os.remove(OUT_WB)
        wb.SaveAs(OUT_WB, FileFormat=52)   # xlOpenXMLWorkbookMacroEnabled
        log("  Saved OK")

    except Exception as e:
        log(f"\nFATAL: {e}")
        raise

    finally:
        _done.set()
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        try:
            excel.DisplayAlerts = True
            excel.Quit()
        except Exception:
            pass

    # ── 5. Summary ───────────────────────────────────────────────────────────
    size = os.path.getsize(OUT_WB) if os.path.isfile(OUT_WB) else 0
    log(f"\n{'='*55}")
    log(f"DONE — {os.path.basename(OUT_WB)}  ({size:,} bytes)")
    log(f"{'='*55}")
    log(f"\nModules replaced ({len(STANDARD_MODULES)} standard + {len(DOCUMENT_MODULES)} document):")
    for m in STANDARD_MODULES:
        log(f"  {m}.bas")
    for cn, fn in DOCUMENT_MODULES.items():
        log(f"  {fn}  → {cn}")

    # Write log
    log_path = os.path.join(ROOT, "build", "apply_v85_log.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(LOG))
    print(f"\nLog written → {log_path}")


if __name__ == "__main__":
    main()
