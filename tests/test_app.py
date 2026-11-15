"""Smoke tests for the Streamlit demo UI. Skipped when Streamlit is not installed."""
import io
import tempfile
import unittest
import zipfile
from pathlib import Path

try:
    from streamlit.testing.v1 import AppTest
except ImportError:  # core tests must not depend on Streamlit
    AppTest = None

if AppTest is not None:
    import app


@unittest.skipIf(AppTest is None, "streamlit not installed (use .venv)")
class AppSmokeTest(unittest.TestCase):
    def test_one_click_sample_demo(self):
        at = AppTest.from_file("../app.py", default_timeout=30).run()
        self.assertFalse(at.exception)
        self.assertTrue(at.button(key="analyze").disabled)  # nothing uploaded yet

        at.button(key="demo").click().run()
        self.assertFalse(at.exception)
        metrics = {m.label: m.value for m in at.metric}
        self.assertEqual(metrics, {"Safe updates": "12", "Blocked for review": "5",
                                   "Warnings": "8", "Unchanged": "1"})
        result = at.session_state["result"]
        self.assertEqual(sorted(result["files"]), ["Products_update_price_and_stock.csv",
                                                   "Products_update_price_only.csv",
                                                   "Products_update_stock_only.csv"])
        self.assertIn("Supplier feed: `supplier_feed.csv`", result["report"])  # no temp paths
        self.assertNotIn(tempfile.gettempdir(), result["report"])
        self.assertTrue(any("USD" in w.value for w in at.warning))


@unittest.skipIf(AppTest is None, "streamlit not installed (use .venv)")
class AppLimitsTest(unittest.TestCase):
    def test_staged_files_are_deleted(self):
        with app.staged({"supplier": ("feed.csv", b"SKU;Price;Stock\nA;1;2\n")}) as (tmp, paths):
            self.assertTrue(paths["supplier"].exists())
        self.assertFalse(Path(tmp).exists())

    def test_too_many_rows_rejected(self):
        data = ("SKU;Price;Stock\n" + "A;1;2\n" * (app.MAX_ROWS + 1)).encode()
        with self.assertRaisesRegex(ValueError, "rows"):
            with app.staged({"supplier": ("feed.csv", data)}):
                pass

    def test_too_large_rejected(self):
        with self.assertRaisesRegex(ValueError, "MB"):
            app.check_size("feed.csv", b"x" * (app.MAX_UPLOAD_MB * 1024 * 1024 + 1))

    def test_xlsx_bomb_rejected(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("xl/worksheets/sheet1.xml", b"0" * (app.MAX_XLSX_UNPACKED_MB * 1024 * 1024 + 1))
        self.assertLess(len(buf.getvalue()), app.MAX_UPLOAD_MB * 1024 * 1024)  # small when packed
        with self.assertRaisesRegex(ValueError, "unpacks"):
            app.check_size("export.xlsx", buf.getvalue())


if __name__ == "__main__":
    unittest.main()
