import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

from tracing_validator_FINAL import App
from tracing_sqlite_cache import OperationCancelled


class ExportTests(unittest.TestCase):
    def test_cancelled_export_preserves_existing_destination(self):
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory) / "report.xlsx"
            destination.write_bytes(b"existing report")
            event = threading.Event()
            def write(path, *args):
                Path(path).write_bytes(b"new report")
                event.set()
            app = SimpleNamespace(cancel_event=event, _write_working_paper=write)
            with self.assertRaises(OperationCancelled):
                App.write_working_paper(app, destination, "100", None, None, None, None, None, 0.01)
            self.assertEqual(destination.read_bytes(), b"existing report")
            self.assertEqual(list(Path(directory).iterdir()), [destination])


    def test_completed_export_replaces_destination(self):
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory) / "report.xlsx"
            app = SimpleNamespace(
                cancel_event=threading.Event(),
                _write_working_paper=lambda path, *args: Path(path).write_bytes(b"completed report"))
            App.write_working_paper(app, destination, "100", None, None, None, None, None, 0.01)
            self.assertEqual(destination.read_bytes(), b"completed report")
            self.assertEqual(list(Path(directory).iterdir()), [destination])


class Value:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class Widget(dict):
    def configure(self, **kwargs):
        self.update(kwargs)


class ConfigurationTests(unittest.TestCase):
    def test_saved_configuration_restores_files_sheets_headers_and_mappings(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.xlsx"
            wb = Workbook()
            wb.active.title = "Data"
            wb.active.append(["GLA", "BU", "Amount", "Beginning", "Variance", "Ending"])
            wb.save(source)
            wb.close()
            errors = []
            app = SimpleNamespace(
                worker_running=False, je_files=[str(source)], je_sheets={str(source): "Data"},
                je_label=Widget(), rf_level_combo={"values": ["GL Account + Business Unit"]},
                rf_level_var=Value("GL Account + Business Unit"), tolerance_var=Value("0.02"),
                account_var=Value("100"), account_combo=Widget(), cache_directory_var=Value(str(root)),
                on_rollforward_level_change=lambda: None,
                show_message=lambda *args: None, show_error=lambda *args: errors.append(args))
            for k in ("je", "tb_beg", "tb_end", "rf"):
                setattr(app, k + "_header_var", Value("1"))
                roles = ["GLA", "Business Unit", "Beginning", "Variance", "Ending"] if k == "rf" else ["GLA", "Business Unit", "Amount"]
                box = SimpleNamespace(vars={r: Value("fixed BU" if r == "Business Unit" else r) for r in roles},
                                      set_columns=lambda columns: None)
                setattr(app, k + "_map", box)
                if k != "je":
                    setattr(app, k + "_path", str(source))
                    setattr(app, k + "_sheet", "Data")
                    setattr(app, k + "_label", Widget())
                if k in ("tb_beg", "tb_end"):
                    setattr(app, k + "_manual_source_var", Value("manual source"))
            app._configuration = lambda: App._configuration(app)
            config = root / "settings.json"
            with patch("tracing_validator_FINAL.filedialog.asksaveasfilename", return_value=str(config)):
                App.save_configuration(app)
            app.account_var.set("")
            app.tb_beg_map.vars["Business Unit"].set("")
            with patch("tracing_validator_FINAL.filedialog.askopenfilename", return_value=str(config)), patch("tracing_validator_FINAL.set_cache_directory"):
                App.load_configuration(app)
            self.assertEqual(errors, [])
            self.assertEqual(app.account_var.get(), "100")
            self.assertEqual(app.tb_beg_map.vars["Business Unit"].get(), "fixed BU")
            self.assertEqual(app.tb_end_manual_source_var.get(), "manual source")
            self.assertEqual(app.rf_sheet, "Data")
            self.assertEqual(app.je_columns[0], "GLA")



if __name__ == "__main__":
    unittest.main()
