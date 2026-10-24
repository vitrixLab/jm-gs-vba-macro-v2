import pathlib
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
SCAN = (ROOT / "vba" / "modGLScanEngine.bas").read_text(encoding="utf-8-sig")
MAP = (ROOT / "vba" / "modGLWorkbookMap.bas").read_text(encoding="utf-8-sig")
APPLY = (ROOT / "scripts" / "apply_full_fix.ps1").read_text(encoding="utf-8-sig")


class ScopeAndFailClosedTests(unittest.TestCase):
    def test_scan_engine_does_not_read_or_scan_gj(self):
        self.assertNotIn('Worksheets("GJ")', SCAN)
        self.assertNotIn("Private Sub ScanGJ(", SCAN)
        self.assertIn("ScanCDJ d, yearNumber", SCAN)
        self.assertIn("ScanCRJ d, yearNumber", SCAN)

    def test_scan_does_not_write_calc_sheet(self):
        self.assertNotIn("WriteScannedCalc d", SCAN)
        self.assertNotIn("Worksheets.Add", SCAN)

    def test_unique_account_guard_runs_before_writeback(self):
        guard = SCAN.index("If Not ValidateGLAccountTitles(detail) Then inv = inv + 1")
        cdj = SCAN.index("ScanCDJ d, yearNumber")
        gate = SCAN.index("If inv > 0 Or um > TOLERANCE")
        writeback = SCAN.index("WriteScannedGL ws, d")
        self.assertLess(guard, cdj)
        self.assertLess(cdj, gate)
        self.assertLess(gate, writeback)
        self.assertIn("duplicate account title", SCAN)
        self.assertIn("missing account title", SCAN)

    def test_resolver_is_not_named_for_excluded_journal(self):
        self.assertNotIn("V8_GJMap", SCAN)
        self.assertNotIn("V8_GJMap", MAP)
        self.assertIn("V8_ResolveAccount", SCAN)
        self.assertIn("V8_ResolveAccount", MAP)

    def test_old_mapping_writer_is_disabled(self):
        self.assertIn("DISABLED", APPLY)
        self.assertNotIn("$mapUpdates", APPLY)
        self.assertNotIn('Value2 = "MAPPED"', APPLY)


if __name__ == "__main__":
    unittest.main()
