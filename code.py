import re
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import pandas as pd
import ttkbootstrap as tb
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_TITLE = "Tracing Validator"
DEFAULT_TOLERANCE = 0.01

# UI / workbook palette
YELLOW = "FFE600"
BLACK = "111111"
DARK = "222222"
LIGHT_GRAY = "E7E6E6"
MID_GRAY = "BFBFBF"
CYAN = "DDEBF7"
GREEN = "C6EFCE"
RED = "FFC7CE"
BEG_FILL = "FFF2CC"
VAR_FILL = "DDEBF7"
END_FILL = "E2F0D9"

thin_gray = Side(style="thin", color="BFBFBF")


def norm_key(series):
    return series.astype("string").fillna("").str.strip().str.upper()


def to_number(series):
    s = series.astype("string").fillna("").str.strip()
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace("$", "", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def safe_sheet_name(value):
    value = re.sub(r'[:\\/?*\[\]]+', "_", str(value)).strip()
    return (value or "Tracing")[:31]


def safe_filename(value):
    return re.sub(r'[\\/:*?"<>|]+', "_", str(value)).strip() or "Tracing"


def read_excel(path, sheet):
    return pd.read_excel(path, sheet_name=sheet, dtype=object)


class MappingBox(ttk.LabelFrame):
    def __init__(self, parent, title, roles):
        super().__init__(parent, text=title, padding=10)
        self.vars, self.combos = {}, {}
        for i, role in enumerate(roles):
            ttk.Label(self, text=role).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=4)
            var = tk.StringVar()
            combo = ttk.Combobox(self, textvariable=var, state="readonly", width=32)
            combo.grid(row=i, column=1, sticky="ew", pady=4)
            self.vars[role] = var
            self.combos[role] = combo
        self.columnconfigure(1, weight=1)

    def set_columns(self, columns):
        vals = [""] + [str(c) for c in columns]
        for c in self.combos.values():
            c["values"] = vals

    def get(self, role):
        return self.vars[role].get()


class App(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title(APP_TITLE)
        self.geometry("1260x900")
        self.minsize(1100, 780)

        self.style.configure(
            "Accent.TButton", background="#FFE600", foreground="#111111",
            font=("Segoe UI Semibold", 10), padding=(14, 8)
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", "#FFD400")],
            foreground=[("active", "#111111")]
        )
        self.style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(11, 7))

        self.je_files = []
        self.tb_beg_path = ""
        self.tb_end_path = ""
        self.rf_path = ""

        self.je_df = self.tb_beg_df = self.tb_end_df = self.rf_df = None
        self._load_queue = queue.Queue()
        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=(24, 20))
        outer.pack(fill="both", expand=True)

        # Header
        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 16))
        bar = tk.Frame(header, bg="#FFE600", width=7, height=58)
        bar.pack(side="left", padx=(0, 14))
        bar.pack_propagate(False)

        titles = ttk.Frame(header)
        titles.pack(side="left")
        ttk.Label(titles, text="TRACING", font=("Segoe UI", 23, "bold")).pack(anchor="w")
        ttk.Label(titles, text="VALIDATOR", foreground="#FFE600",
                  font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(
            header, text="GLA + BUSINESS UNIT  â¢  WORKING PAPER GENERATOR",
            foreground="#999999", font=("Segoe UI", 9, "bold")
        ).pack(side="right", anchor="s", pady=6)

        # Files
        files = ttk.LabelFrame(outer, text="  01  SOURCE FILES  ", padding=12)
        files.pack(fill="x", pady=(0, 10))
        files.columnconfigure(1, weight=1)

        def file_row(row, label, cmd, attr):
            ttk.Label(files, text=label, font=("Segoe UI Semibold", 10)).grid(
                row=row, column=0, sticky="w", padx=(0, 12), pady=4
            )
            status = ttk.Label(files, text="Not selected", foreground="#999999")
            status.grid(row=row, column=1, sticky="w", pady=4)
            setattr(self, attr, status)
            ttk.Button(files, text="Browse", style="Secondary.TButton", command=cmd).grid(
                row=row, column=2, sticky="e", pady=4
            )

        file_row(0, "Journal Entries (one or many)", self.add_je, "je_label")
        file_row(1, "TB Beginning", lambda: self.pick_single("tb_beg"), "tb_beg_label")
        file_row(2, "TB Ending", lambda: self.pick_single("tb_end"), "tb_end_label")
        file_row(3, "Rollforward", lambda: self.pick_single("rf"), "rf_label")

        self.load_button = ttk.Button(
            files, text="LOAD WORKBOOKS  â", style="Accent.TButton", command=self.load_files
        )
        self.load_button.grid(row=4, column=2, sticky="e", pady=(8, 0))

        # Mappings
        wrap = ttk.LabelFrame(outer, text="  02  FIELD MAPPING  ", padding=10)
        wrap.pack(fill="both", expand=True, pady=(0, 10))
        maps = ttk.Frame(wrap)
        maps.pack(fill="both", expand=True)

        self.je_map = MappingBox(maps, "JOURNAL ENTRIES", ["GLA", "Business Unit", "Amount"])
        self.tb_beg_map = MappingBox(maps, "TB BEGINNING", ["GLA", "Business Unit", "Amount"])
        self.tb_end_map = MappingBox(maps, "TB ENDING", ["GLA", "Business Unit", "Amount"])
        self.rf_map = MappingBox(
            maps, "ROLLFORWARD",
            ["GLA", "Business Unit", "Beginning", "Variance", "Ending"]
        )

        self.je_map.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
        self.tb_beg_map.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6))
        self.tb_end_map.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(6, 0))
        self.rf_map.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(6, 0))
        maps.columnconfigure(0, weight=1)
        maps.columnconfigure(1, weight=1)

        # Run
        run = ttk.LabelFrame(outer, text="  03  RUN TRACING  ", padding=12)
        run.pack(fill="x")

        ttk.Label(run, text="GLA ACCOUNT").grid(row=0, column=0, sticky="w")
        self.account_var = tk.StringVar()
        self.account_combo = ttk.Combobox(run, textvariable=self.account_var, width=28)
        self.account_combo.grid(row=1, column=0, padx=(0, 12), pady=(4, 0))

        ttk.Label(run, text="TOLERANCE").grid(row=0, column=1, sticky="w")
        self.tolerance_var = tk.StringVar(value=str(DEFAULT_TOLERANCE))
        ttk.Entry(run, textvariable=self.tolerance_var, width=14).grid(
            row=1, column=1, padx=(0, 12), pady=(4, 0)
        )

        ttk.Button(
            run, text="REFRESH GLA", style="Secondary.TButton", command=self.refresh_accounts
        ).grid(row=1, column=2, padx=(0, 10), pady=(4, 0))
        ttk.Button(
            run, text="GENERATE WORKING PAPER  â", style="Accent.TButton", command=self.generate
        ).grid(row=1, column=3, pady=(4, 0))

        statusbar = ttk.Frame(outer, padding=(10, 7))
        statusbar.pack(fill="x", pady=(10, 0))
        ttk.Label(statusbar, text="â", foreground="#FFE600").pack(side="left", padx=(0, 6))
        self.status = ttk.Label(statusbar, text="Ready", foreground="#BBBBBB")
        self.status.pack(side="left")

    def add_je(self):
        paths = filedialog.askopenfilenames(
            title="Select one or more JE files",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        if paths:
            self.je_files = list(paths)
            names = ", ".join(Path(p).name for p in self.je_files[:3])
            if len(self.je_files) > 3:
                names += "..."
            self.je_label.config(text=f"{len(self.je_files)} file(s): {names}")

    def pick_single(self, kind):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        if not path:
            return
        if kind == "tb_beg":
            self.tb_beg_path = path
            self.tb_beg_label.config(text=Path(path).name)
        elif kind == "tb_end":
            self.tb_end_path = path
            self.tb_end_label.config(text=Path(path).name)
        else:
            self.rf_path = path
            self.rf_label.config(text=Path(path).name)

    def choose_sheet(self, path, title):
        sheets = pd.ExcelFile(path).sheet_names
        if len(sheets) == 1:
            return sheets[0]

        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.grab_set()
        var = tk.StringVar(value=sheets[0])
        ttk.Label(win, text=Path(path).name).pack(padx=20, pady=(15, 5))
        ttk.Combobox(win, values=sheets, textvariable=var, state="readonly", width=42).pack(
            padx=20, pady=5
        )
        result = {"sheet": None}

        def ok():
            result["sheet"] = var.get()
            win.destroy()

        ttk.Button(win, text="OK", style="Accent.TButton", command=ok).pack(pady=(5, 15))
        self.wait_window(win)
        return result["sheet"]

    def load_files(self):
        """
        Select sheets on the UI thread, then read large workbooks in a worker
        thread so Tkinter does not freeze/crash while pandas/openpyxl is busy.
        """
        if not self.je_files or not self.tb_beg_path or not self.tb_end_path or not self.rf_path:
            messagebox.showerror(
                "Missing files",
                "Select JE(s), TB Beginning, TB Ending and Rollforward."
            )
            return

        try:
            selections = {"je": []}

            # Dialogs must stay on Tkinter's main thread.
            for path in self.je_files:
                sh = self.choose_sheet(path, "JE sheet")
                if not sh:
                    return
                selections["je"].append((path, sh))

            sh = self.choose_sheet(self.tb_beg_path, "TB Beginning sheet")
            if not sh:
                return
            selections["tb_beg"] = (self.tb_beg_path, sh)

            sh = self.choose_sheet(self.tb_end_path, "TB Ending sheet")
            if not sh:
                return
            selections["tb_end"] = (self.tb_end_path, sh)

            sh = self.choose_sheet(self.rf_path, "Rollforward sheet")
            if not sh:
                return
            selections["rf"] = (self.rf_path, sh)

            self.load_button.configure(state="disabled")
            self.status.config(text="Loading Excel filesâ¦ the window will remain responsive.")
            self.update_idletasks()

            worker = threading.Thread(
                target=self._load_files_worker,
                args=(selections,),
                daemon=True
            )
            worker.start()
            self.after(150, self._poll_load_queue)

        except Exception as exc:
            self.load_button.configure(state="normal")
            messagebox.showerror("Load error", str(exc))

    def _load_files_worker(self, selections):
        try:
            je_parts = []
            expected_cols = None

            for path, sh in selections["je"]:
                df = read_excel(path, sh)
                original_cols = list(df.columns)

                if expected_cols is None:
                    expected_cols = original_cols
                elif original_cols != expected_cols:
                    raise ValueError(
                        "The selected JE files do not have the same column structure.\n"
                        "All JE files must use the same columns in this version."
                    )

                df["Source_File"] = Path(path).name
                je_parts.append(df)

            # concat(copy=False) reduces avoidable memory duplication.
            je_df = pd.concat(je_parts, ignore_index=True, copy=False)

            path, sh = selections["tb_beg"]
            tb_beg_df = read_excel(path, sh)

            path, sh = selections["tb_end"]
            tb_end_df = read_excel(path, sh)

            path, sh = selections["rf"]
            rf_df = read_excel(path, sh)

            self._load_queue.put(("ok", (je_df, tb_beg_df, tb_end_df, rf_df)))

        except Exception as exc:
            self._load_queue.put(("error", str(exc)))

    def _poll_load_queue(self):
        try:
            status, payload = self._load_queue.get_nowait()
        except queue.Empty:
            self.after(150, self._poll_load_queue)
            return

        self.load_button.configure(state="normal")

        if status == "error":
            self.status.config(text="Load failed.")
            messagebox.showerror("Load error", payload)
            return

        self.je_df, self.tb_beg_df, self.tb_end_df, self.rf_df = payload

        self.je_map.set_columns([c for c in self.je_df.columns if c != "Source_File"])
        self.tb_beg_map.set_columns(self.tb_beg_df.columns)
        self.tb_end_map.set_columns(self.tb_end_df.columns)
        self.rf_map.set_columns(self.rf_df.columns)

        self.status.config(
            text=f"Loaded: {len(self.je_df):,} JE rows | "
                 f"{len(self.tb_beg_df):,} TB Beg | "
                 f"{len(self.tb_end_df):,} TB End | "
                 f"{len(self.rf_df):,} RF"
        )
        messagebox.showinfo(
            "Loaded",
            "Files loaded successfully.\n\n"
            f"JE rows: {len(self.je_df):,}\n"
            f"TB Beginning rows: {len(self.tb_beg_df):,}\n"
            f"TB Ending rows: {len(self.tb_end_df):,}\n"
            f"Rollforward rows: {len(self.rf_df):,}\n\n"
            "Now map the required fields."
        )

    def validate_mapping(self):
        checks = [
            (self.je_map, ["GLA", "Business Unit", "Amount"], "JE"),
            (self.tb_beg_map, ["GLA", "Business Unit", "Amount"], "TB Beginning"),
            (self.tb_end_map, ["GLA", "Business Unit", "Amount"], "TB Ending"),
            (self.rf_map, ["GLA", "Business Unit", "Beginning", "Variance", "Ending"], "Rollforward"),
        ]
        for box, roles, name in checks:
            for role in roles:
                if not box.get(role):
                    raise ValueError(f"Select '{role}' for {name}.")

    def refresh_accounts(self):
        try:
            if self.rf_df is None:
                raise ValueError("Load the workbooks first.")
            gla_col = self.rf_map.get("GLA")
            if not gla_col:
                raise ValueError("Map the Rollforward GLA field first.")
            vals = sorted(x for x in norm_key(self.rf_df[gla_col]).unique().tolist() if x)
            self.account_combo["values"] = vals
            if vals and not self.account_var.get():
                self.account_var.set(vals[0])
            self.status.config(text=f"{len(vals)} GLA account(s) available.")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _prepared(self, df, mapping):
        out = df.copy()
        out["__GLA"] = norm_key(out[mapping.get("GLA")])
        out["__BU"] = norm_key(out[mapping.get("Business Unit")])
        out["__AMOUNT"] = to_number(out[mapping.get("Amount")])
        return out

    def generate(self):
        try:
            self.validate_mapping()
            account = self.account_var.get().strip().upper()
            if not account:
                raise ValueError("Select or enter a GLA.")
            tolerance = float(self.tolerance_var.get())

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"Tracing_{safe_filename(account)}.xlsx",
                filetypes=[("Excel workbook", "*.xlsx")]
            )
            if not save_path:
                return

            je = self._prepared(self.je_df, self.je_map)
            beg = self._prepared(self.tb_beg_df, self.tb_beg_map)
            end = self._prepared(self.tb_end_df, self.tb_end_map)

            rf = self.rf_df.copy()
            rf["__GLA"] = norm_key(rf[self.rf_map.get("GLA")])
            rf["__BU"] = norm_key(rf[self.rf_map.get("Business Unit")])
            rf["__BEG"] = to_number(rf[self.rf_map.get("Beginning")])
            rf["__VAR"] = to_number(rf[self.rf_map.get("Variance")])
            rf["__END"] = to_number(rf[self.rf_map.get("Ending")])

            je = je[je["__GLA"] == account].copy()
            beg = beg[beg["__GLA"] == account].copy()
            end = end[end["__GLA"] == account].copy()
            rf = rf[rf["__GLA"] == account].copy()

            if rf.empty:
                raise ValueError(f"GLA '{account}' was not found in Rollforward.")

            # Build BU-level validation with FULL OUTER behavior.
            rf_s = rf.groupby(["__GLA", "__BU"], dropna=False).agg(
                RF_Beginning=("__BEG", "sum"),
                RF_Variance=("__VAR", "sum"),
                RF_Ending=("__END", "sum"),
            ).reset_index()
            beg_s = beg.groupby(["__GLA", "__BU"], dropna=False)["__AMOUNT"].sum().reset_index(name="TB_Beginning_Total")
            je_s = je.groupby(["__GLA", "__BU"], dropna=False)["__AMOUNT"].sum().reset_index(name="JE_Total")
            end_s = end.groupby(["__GLA", "__BU"], dropna=False)["__AMOUNT"].sum().reset_index(name="TB_Ending_Total")

            val = rf_s.merge(beg_s, on=["__GLA", "__BU"], how="outer")
            val = val.merge(je_s, on=["__GLA", "__BU"], how="outer")
            val = val.merge(end_s, on=["__GLA", "__BU"], how="outer")
            nums = ["RF_Beginning", "RF_Variance", "RF_Ending",
                    "TB_Beginning_Total", "JE_Total", "TB_Ending_Total"]
            val[nums] = val[nums].fillna(0.0)

            val["Beginning_Difference"] = val["TB_Beginning_Total"] - val["RF_Beginning"]
            val["Activity_Difference"] = val["JE_Total"] - val["RF_Variance"]
            val["Ending_Difference"] = val["TB_Ending_Total"] - val["RF_Ending"]
            val["Beginning_Result"] = val["Beginning_Difference"].abs().le(tolerance).map({True: "PASS", False: "FAIL"})
            val["Activity_Result"] = val["Activity_Difference"].abs().le(tolerance).map({True: "PASS", False: "FAIL"})
            val["Ending_Result"] = val["Ending_Difference"].abs().le(tolerance).map({True: "PASS", False: "FAIL"})
            val.rename(columns={"__GLA": "GLA", "__BU": "Business Unit"}, inplace=True)

            # Strip helper columns for original-detail sections.
            helper_je = ["__GLA", "__BU", "__AMOUNT"]
            helper_rf = ["__GLA", "__BU", "__BEG", "__VAR", "__END"]
            je_original = je.drop(columns=helper_je, errors="ignore")
            beg_original = beg.drop(columns=helper_je, errors="ignore")
            end_original = end.drop(columns=helper_je, errors="ignore")
            rf_original = rf.drop(columns=helper_rf, errors="ignore")

            self.write_working_paper(
                save_path, account, rf_original, beg_original, end_original, je_original,
                val, tolerance
            )
            self.status.config(text=f"Working paper generated: {Path(save_path).name}")
            messagebox.showinfo("Done", f"Tracing working paper created:\n{save_path}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def write_working_paper(self, path, account, rf_df, beg_df, end_df, je_df, val_df, tolerance):
        # Create blank workbook via pandas/openpyxl.
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            pd.DataFrame().to_excel(writer, sheet_name=safe_sheet_name(account), index=False)

        wb = load_workbook(path)
        ws = wb[safe_sheet_name(account)]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A3"

        def section_title(row, text):
            ws.cell(row=row, column=1, value=text)
            ws.cell(row=row, column=1).font = Font(bold=True, color="FFFFFF", size=11)
            ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor=BLACK)
            # extend visual title band
            for c in range(1, 8):
                ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=BLACK)
            return row + 1

        def write_df(start_row, df, header_fill=LIGHT_GRAY):
            if df is None:
                df = pd.DataFrame()
            headers = list(df.columns)
            for j, h in enumerate(headers, 1):
                cell = ws.cell(start_row, j, h)
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor=header_fill)
                cell.border = Border(bottom=thin_gray)
            for i, row in enumerate(df.itertuples(index=False, name=None), start_row + 1):
                for j, value in enumerate(row, 1):
                    ws.cell(i, j, value)
            return start_row + 1 + len(df), headers

        # Header
        ws["A1"] = "TRACING WORKING PAPER"
        ws["A1"].font = Font(bold=True, size=16, color=BLACK)
        ws["A2"] = f"GLA: {account}"
        ws["A2"].font = Font(bold=True, color="666666")

        row = 4

        # Rollforward original
        row = section_title(row, "ROLLFORWARD")
        rf_header_row = row
        row, rf_headers = write_df(row, rf_df, "D9EAD3")
        rf_beg_col = self.rf_map.get("Beginning")
        rf_var_col = self.rf_map.get("Variance")
        rf_end_col = self.rf_map.get("Ending")
        for header, color in [(rf_beg_col, BEG_FILL), (rf_var_col, VAR_FILL), (rf_end_col, END_FILL)]:
            if header in rf_headers:
                col = rf_headers.index(header) + 1
                for r in range(rf_header_row, row):
                    ws.cell(r, col).fill = PatternFill("solid", fgColor=color)

        row += 2

        # Validation block
        row = section_title(row, "VALIDATION BY GLA + BUSINESS UNIT")
        validation_cols = [
            "GLA", "Business Unit",
            "RF_Beginning", "TB_Beginning_Total", "Beginning_Difference", "Beginning_Result",
            "RF_Variance", "JE_Total", "Activity_Difference", "Activity_Result",
            "RF_Ending", "TB_Ending_Total", "Ending_Difference", "Ending_Result"
        ]
        v = val_df[validation_cols].copy()
        validation_header = row
        row, vheaders = write_df(row, v, "D9EAF7")

        # Color validation cells
        result_names = ["Beginning_Result", "Activity_Result", "Ending_Result"]
        diff_names = ["Beginning_Difference", "Activity_Difference", "Ending_Difference"]
        for name in result_names:
            c = vheaders.index(name) + 1
            for r in range(validation_header + 1, row):
                cell = ws.cell(r, c)
                cell.fill = PatternFill("solid", fgColor=GREEN if cell.value == "PASS" else RED)
                cell.font = Font(bold=True)
        for name in diff_names:
            c = vheaders.index(name) + 1
            for r in range(validation_header + 1, row):
                cell = ws.cell(r, c)
                try:
                    ok = abs(float(cell.value or 0)) <= tolerance
                    cell.fill = PatternFill("solid", fgColor=GREEN if ok else RED)
                except Exception:
                    pass

        row += 2

        # Beginning detail
        row = section_title(row, "BEGINNING BALANCE â TB DETAIL")
        beg_header = row
        row, beg_headers = write_df(row, beg_df, "FFF2CC")
        beg_amount_col = self.tb_beg_map.get("Amount")
        if beg_amount_col in beg_headers:
            c = beg_headers.index(beg_amount_col) + 1
            total_row = row
            ws.cell(total_row, max(1, c - 1), "TOTAL")
            ws.cell(total_row, max(1, c - 1)).font = Font(bold=True)
            ws.cell(total_row, c, f"=SUM({get_column_letter(c)}{beg_header+1}:{get_column_letter(c)}{row-1})" if len(beg_df) else 0)
            ws.cell(total_row, c).font = Font(bold=True)
            ws.cell(total_row, c).fill = PatternFill("solid", fgColor=CYAN)
            row += 1

        row += 2

        # Ending detail
        row = section_title(row, "ENDING BALANCE â TB DETAIL")
        end_header = row
        row, end_headers = write_df(row, end_df, "E2F0D9")
        end_amount_col = self.tb_end_map.get("Amount")
        if end_amount_col in end_headers:
            c = end_headers.index(end_amount_col) + 1
            total_row = row
            ws.cell(total_row, max(1, c - 1), "TOTAL")
            ws.cell(total_row, max(1, c - 1)).font = Font(bold=True)
            ws.cell(total_row, c, f"=SUM({get_column_letter(c)}{end_header+1}:{get_column_letter(c)}{row-1})" if len(end_df) else 0)
            ws.cell(total_row, c).font = Font(bold=True)
            ws.cell(total_row, c).fill = PatternFill("solid", fgColor=CYAN)
            row += 1

        row += 2

        # JE detail
        row = section_title(row, "JOURNAL ENTRIES â DETAIL")
        je_header = row
        row, je_headers = write_df(row, je_df, "D9EAF7")
        je_amount_col = self.je_map.get("Amount")
        if je_amount_col in je_headers:
            c = je_headers.index(je_amount_col) + 1
            total_row = row
            ws.cell(total_row, max(1, c - 1), "TOTAL")
            ws.cell(total_row, max(1, c - 1)).font = Font(bold=True)
            ws.cell(total_row, c, f"=SUM({get_column_letter(c)}{je_header+1}:{get_column_letter(c)}{row-1})" if len(je_df) else 0)
            ws.cell(total_row, c).font = Font(bold=True)
            ws.cell(total_row, c).fill = PatternFill("solid", fgColor=CYAN)

        # Formatting
        max_col = max(ws.max_column, 1)
        for col in range(1, max_col + 1):
            letter = get_column_letter(col)
            max_len = 0
            for r in range(1, min(ws.max_row, 250) + 1):
                value = ws.cell(r, col).value
                if value is not None:
                    max_len = max(max_len, len(str(value)))
            ws.column_dimensions[letter].width = min(max(max_len + 2, 11), 32)

        # Accounting-style number format where values/formulas are numeric.
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith("=")):
                    cell.number_format = '#,##0.00;[Red](#,##0.00);-'

        # Auto-filter is intentionally not set across the whole sheet because it has multiple tables.
        wb.save(path)


if __name__ == "__main__":
    App().mainloop()