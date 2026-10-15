import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
import tracing_sqlite_cache as cache


class SQLiteCacheTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.cache_patch = patch.object(cache, "CACHE_ROOT", self.root / "cache")
        self.cache_patch.start()
        self.path = self.root / "source.xlsx"
        self.write_source()

    def tearDown(self):
        self.cache_patch.stop()
        self.temp.cleanup()

    def write_source(self, extra=False):
        wb = Workbook()
        ws = wb.active
        ws.title = "JE"
        ws.append(["Report title"])
        ws.append(["GLA", "Amount", "Date", "Flag", "Description"])
        ws.append([100.0, 12.5, datetime(2026, 1, 2, 3, 4), True, "á test"])
        ws.append([200, -2, None, False, "other"])
        ws.append(["100.0", -5, None, False, "second"])
        if extra:
            ws.append([300, 9, None, False, "new"])
        wb.save(self.path)
        wb.close()

    def read(self, account):
        return cache.read_excel_filtered_by_gla(self.path, "JE", 2, "GLA", [], account)

    def test_cached_accounts_never_reopen_excel_and_preserve_values(self):
        with patch.object(cache, "BATCH_SIZE", 1):
            result = self.read("100")
        self.assertEqual(result["Amount"].tolist(), [12.5, -5])
        self.assertEqual(result.iloc[0]["Date"], datetime(2026, 1, 2, 3, 4))
        self.assertEqual(result.iloc[0]["Description"], "á test")
        self.assertTrue(result.iloc[0]["Flag"])
        with patch.object(cache, "load_workbook", side_effect=AssertionError("Excel reopened")):
            self.assertEqual(self.read(200)["Amount"].tolist(), [-2])
            empty = self.read("missing")
            self.assertTrue(empty.empty)
            self.assertEqual(empty.columns.tolist(), result.columns.tolist())

    def test_source_change_invalidates_only_its_database(self):
        self.read(100)
        previous = cache.cache_path_for(self.path, "JE", 2, "GLA")
        self.write_source(extra=True)
        self.assertIsNone(cache.load_gla_index(self.path, "JE", 2, "GLA"))
        self.assertEqual(self.read(300)["Amount"].tolist(), [9])
        self.assertNotEqual(previous, cache.cache_path_for(self.path, "JE", 2, "GLA"))

    def test_failed_import_is_not_published(self):
        def fail(*args):
            raise RuntimeError("interrupted")
        with self.assertRaises(RuntimeError):
            cache.build_gla_index(self.path, "JE", 2, "GLA", fail)
        self.assertIsNone(cache.load_gla_index(self.path, "JE", 2, "GLA"))
        self.assertEqual(list((self.root / "cache").iterdir()), [])
        self.assertEqual(len(self.read(100)), 2)

    def test_folder_selection_is_required_and_controls_storage(self):
        cache.CACHE_ROOT = None
        with self.assertRaisesRegex(ValueError, "Select a cache folder"):
            self.read(100)
        selected = self.root / "chosen"
        selected.mkdir()
        self.assertEqual(cache.set_cache_directory(selected), selected.resolve())
        self.read(100)
        self.assertEqual(len(list(selected.glob("*.sqlite3"))), 1)
        with self.assertRaises(ValueError):
            cache.set_cache_directory(self.root / "missing")
        self.assertEqual(cache.CACHE_ROOT, selected.resolve())


if __name__ == "__main__":
    unittest.main()
