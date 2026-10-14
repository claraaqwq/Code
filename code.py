import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as tb
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


APP_TITLE = "Tracing Validator"
TOLERANCE = 0.01

# Excel colors
FILL_BEG = PatternFill("solid", fgColor="FFF2CC")
FILL_VAR = PatternFill("solid", fgColor="DDEBF7")
FILL_END = PatternFill("solid", fgColor="E2F0D9")
FILL_PASS = PatternFill("solid", fgColor="C6EFCE")
FILL_FAIL = PatternFill("solid", fgColor="FFC7CE")
FILL_HEADER = PatternFill("solid", fgColor="D9EAF7")


def norm_key(series):
    """Normalize join keys without destroying leading zeros."""
    return (
        series.astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )


def to_number(series):
    """Convert common Excel accounting formats to numeric."""
    s = series.astype("string").fillna("").str.strip()
    # (1,234.56) -> -1,234.56
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace("$", "", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def safe_filename(value):
    value = re.sub(r'[\\/:*?"<>|]+', "_", str(value))
    return value.strip() or "Tracing"


def read_excel(path, sheet_name):
    return pd.read_excel(path, sheet_name=sheet_name, dtype=object)


def sheet_names(path):
    return pd.ExcelFile(path).sheet_names


class MappingBox(ttk.LabelFrame):
    def __init__(self, parent, title, roles):
        super().__init__(parent, text=title, padding=8)
        self.roles = roles
        self.vars = {}
        self.combos = {}
        for r, label in enumerate(roles):
            ttk.Label(self, text=label).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
            var = tk.StringVar()
            combo = ttk.Combobox(self, textvariable=var, state="readonly", width=35)
            combo.grid(row=r, column=1, sticky="ew", pady=3)
            self.vars[label] = var
            self.combos[label] = combo
        self.columnconfigure(1, weight=1)

    def set_columns(self, columns):
        values = [""] + [str(c) for c in columns]
        for combo in self.combos.values():
            combo["values"] = values

    def get(self, label):
        return self.vars[label].get()


class App(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title(APP_TITLE)
        self.geometry("1220x860")
        self.minsize(1080, 760)
        self.style.configure("Accent.TButton", background="#FFE600", foreground="#111111",
                             font=("Segoe UI Semibold", 10), padding=(14, 8))
        self.style.map("Accent.TButton", background=[("active", "#FFD400")],
                       foreground=[("active", "#111111")])
        self.style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(12, 7))

        self.je_files = []
        self.tb_beg_path = ""
        self.tb_end_path = ""
        self.rf_path = ""

        self.je_df = None
        self.tb_beg_df = None
        self.tb_end_df = None
        self.rf_df = None

        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=(24, 20))
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 18))
        accent = tk.Frame(header, width=7, height=58, bg="#FFE600")
        accent.pack(side="left", padx=(0, 14)); accent.pack_propagate(False)
        titlebox = ttk.Frame(header); titlebox.pack(side="left")
        ttk.Label(titlebox, text="TRACING", font=("Segoe UI", 23, "bold")).pack(anchor="w")
        ttk.Label(titlebox, text="VALIDATOR", foreground="#FFE600", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(header, text="GLA + BUSINESS UNIT  â¢  RECONCILIATION WORKSPACE",
                  foreground="#999999", font=("Segoe UI", 9, "bold")).pack(side="right", anchor="s", pady=6)

        files = ttk.LabelFrame(outer, text="  01  SOURCE FILES  ", padding=14)
        files.pack(fill="x", pady=(0, 12)); files.columnconfigure(1, weight=1)
        def frow(r,label,cmd,attr):
            ttk.Label(files,text=label,font=("Segoe UI Semibold",10)).grid(row=r,column=0,sticky="w",padx=(0,12),pady=5)
            w=ttk.Label(files,text="Not selected",foreground="#999999"); w.grid(row=r,column=1,sticky="w",pady=5)
            setattr(self,attr,w)
            ttk.Button(files,text="Browse",style="Secondary.TButton",command=cmd).grid(row=r,column=2,sticky="e",pady=5)
        frow(0,"Journal Entries",self.add_je,"je_label")
        frow(1,"TB Beginning",lambda:self.pick_single("tb_beg"),"tb_beg_label")
        frow(2,"TB Ending",lambda:self.pick_single("tb_end"),"tb_end_label")
        frow(3,"Rollforward",lambda:self.pick_single("rf"),"rf_label")
        ttk.Button(files,text="LOAD WORKBOOKS  â",style="Accent.TButton",command=self.load_files).grid(row=4,column=2,sticky="e",pady=(10,0))

        wrap=ttk.LabelFrame(outer,text="  02  FIELD MAPPING  ",padding=12); wrap.pack(fill="both",expand=True,pady=(0,12))
        mappings=ttk.Frame(wrap); mappings.pack(fill="both",expand=True)
        self.je_map=MappingBox(mappings,"JOURNAL ENTRIES",["GLA","Business Unit","Amount"])
        self.tb_beg_map=MappingBox(mappings,"TB BEGINNING",["GLA","Business Unit","Amount"])
        self.tb_end_map=MappingBox(mappings,"TB ENDING",["GLA","Business Unit","Amount"])
        self.rf_map=MappingBox(mappings,"ROLLFORWARD",["GLA","Business Unit","Beginning","Variance","Ending"])
        self.je_map.grid(row=0,column=0,sticky="nsew",padx=(0,6),pady=(0,6)); self.tb_beg_map.grid(row=0,column=1,sticky="nsew",padx=(6,0),pady=(0,6))
        self.tb_end_map.grid(row=1,column=0,sticky="nsew",padx=(0,6),pady=(6,0)); self.rf_map.grid(row=1,column=1,sticky="nsew",padx=(6,0),pady=(6,0))
        mappings.columnconfigure(0,weight=1); mappings.columnconfigure(1,weight=1)

        run=ttk.LabelFrame(outer,text="  03  RUN TRACING  ",padding=14); run.pack(fill="x")
        ttk.Label(run,text="GLA ACCOUNT",font=("Segoe UI Semibold",9)).grid(row=0,column=0,sticky="w")
        self.account_var=tk.StringVar(); self.account_combo=ttk.Combobox(run,textvariable=self.account_var,width=28)
        self.account_combo.grid(row=1,column=0,padx=(0,12),pady=(4,0))
        ttk.Label(run,text="TOLERANCE",font=("Segoe UI Semibold",9)).grid(row=0,column=1,sticky="w")
        self.tolerance_var=tk.StringVar(value=str(TOLERANCE)); ttk.Entry(run,textvariable=self.tolerance_var,width=14).grid(row=1,column=1,padx=(0,12),pady=(4,0))
        ttk.Button(run,text="REFRESH GLA",style="Secondary.TButton",command=self.refresh_accounts).grid(row=1,column=2,padx=(0,10),pady=(4,0))
        ttk.Button(run,text="GENERATE TRACING  â",style="Accent.TButton",command=self.generate).grid(row=1,column=3,pady=(4,0))
        statusbar=ttk.Frame(outer,padding=(10,7)); statusbar.pack(fill="x",pady=(12,0))
        ttk.Label(statusbar,text="â",foreground="#FFE600").pack(side="left",padx=(0,6))
        self.status=ttk.Label(statusbar,text="Ready",foreground="#BBBBBB"); self.status.pack(side="left")
        ttk.Label(statusbar,text="Functional tracing  |  GLA + Business Unit",foreground="#777777").pack(side="right")

    def add_je(self):
        paths = filedialog.askopenfilenames(
            title="Selecciona uno o varios JE",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        if paths:
            self.je_files = list(paths)
            self.je_label.config(text=f"{len(self.je_files)} archivo(s): " + ", ".join(Path(p).name for p in self.je_files[:3]) +
                                      ("..." if len(self.je_files) > 3 else ""))

    def pick_single(self, kind):
        path = filedialog.askopenfilename(
            title="Selecciona archivo",
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
        sheets = sheet_names(path)
        if len(sheets) == 1:
            return sheets[0]

        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.grab_set()
        var = tk.StringVar(value=sheets[0])

        ttk.Label(win, text=f"{Path(path).name}\nSelecciona la hoja:").pack(padx=20, pady=(15, 5))
        combo = ttk.Combobox(win, values=sheets, textvariable=var, state="readonly", width=40)
        combo.pack(padx=20, pady=5)

        result = {"value": None}

        def ok():
            result["value"] = var.get()
            win.destroy()

        ttk.Button(win, text="OK", command=ok).pack(pady=(5, 15))
        self.wait_window(win)
        return result["value"]

    def load_files(self):
        if not self.je_files or not self.tb_beg_path or not self.tb_end_path or not self.rf_path:
            messagebox.showerror("Faltan archivos", "Selecciona JE(s), TB Beginning, TB Ending y Rollforward.")
            return

        try:
            # JE files: this version assumes a common column structure.
            je_parts = []
            first_columns = None
            for path in self.je_files:
                sh = self.choose_sheet(path, "Hoja JE")
                if not sh:
                    return
                df = read_excel(path, sh)
                df["__Source_File"] = Path(path).name
                if first_columns is None:
                    first_columns = [c for c in df.columns if c != "__Source_File"]
                je_parts.append(df)
            self.je_df = pd.concat(je_parts, ignore_index=True, sort=False)

            sh = self.choose_sheet(self.tb_beg_path, "Hoja TB Beginning")
            if not sh: return
            self.tb_beg_df = read_excel(self.tb_beg_path, sh)

            sh = self.choose_sheet(self.tb_end_path, "Hoja TB Ending")
            if not sh: return
            self.tb_end_df = read_excel(self.tb_end_path, sh)

            sh = self.choose_sheet(self.rf_path, "Hoja Rollforward")
            if not sh: return
            self.rf_df = read_excel(self.rf_path, sh)

            self.je_map.set_columns([c for c in self.je_df.columns if c != "__Source_File"])
            self.tb_beg_map.set_columns(self.tb_beg_df.columns)
            self.tb_end_map.set_columns(self.tb_end_df.columns)
            self.rf_map.set_columns(self.rf_df.columns)

            self.status.config(text="Archivos leÃ­dos. Selecciona los mappings y luego actualiza la lista de GLA.")
            messagebox.showinfo("Listo", "Archivos cargados. Ahora selecciona las columnas de cada archivo.")
        except Exception as e:
            messagebox.showerror("Error al leer archivos", str(e))

    def validate_mapping(self):
        groups = [
            (self.je_map, ["GLA", "Business Unit", "Amount"], "JE"),
            (self.tb_beg_map, ["GLA", "Business Unit", "Amount"], "TB Beginning"),
            (self.tb_end_map, ["GLA", "Business Unit", "Amount"], "TB Ending"),
            (self.rf_map, ["GLA", "Business Unit", "Beginning", "Variance", "Ending"], "Rollforward"),
        ]
        for box, roles, name in groups:
            for role in roles:
                if not box.get(role):
                    raise ValueError(f"Falta seleccionar '{role}' en {name}.")

    def refresh_accounts(self):
        try:
            if self.rf_df is None:
                raise ValueError("Primero lee los archivos.")
            gla = self.rf_map.get("GLA")
            if not gla:
                raise ValueError("Selecciona primero el campo GLA del Rollforward.")
            vals = norm_key(self.rf_df[gla])
            accounts = sorted([x for x in vals.unique().tolist() if x])
            self.account_combo["values"] = accounts
            if accounts and not self.account_var.get():
                self.account_var.set(accounts[0])
            self.status.config(text=f"{len(accounts)} GLA disponibles en Rollforward.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _prepare_source(self, df, mapping, amount_role="Amount"):
        out = df.copy()
        out["__GLA"] = norm_key(out[mapping.get("GLA")])
        out["__BU"] = norm_key(out[mapping.get("Business Unit")])
        out["__AMOUNT"] = to_number(out[mapping.get(amount_role)])
        return out

    def generate(self):
        try:
            self.validate_mapping()
            account = self.account_var.get().strip().upper()
            if not account:
                raise ValueError("Indica el GLA que quieres trazar.")
            tolerance = float(self.tolerance_var.get())

            save_path = filedialog.asksaveasfilename(
                title="Guardar tracing",
                defaultextension=".xlsx",
                initialfile=f"Tracing_{safe_filename(account)}.xlsx",
                filetypes=[("Excel workbook", "*.xlsx")]
            )
            if not save_path:
                return

            # Prepare / filter sources
            je = self._prepare_source(self.je_df, self.je_map)
            tb_beg = self._prepare_source(self.tb_beg_df, self.tb_beg_map)
            tb_end = self._prepare_source(self.tb_end_df, self.tb_end_map)

            rf = self.rf_df.copy()
            rf["__GLA"] = norm_key(rf[self.rf_map.get("GLA")])
            rf["__BU"] = norm_key(rf[self.rf_map.get("Business Unit")])
            rf["__RF_BEG"] = to_number(rf[self.rf_map.get("Beginning")])
            rf["__RF_VAR"] = to_number(rf[self.rf_map.get("Variance")])
            rf["__RF_END"] = to_number(rf[self.rf_map.get("Ending")])

            je_a = je[je["__GLA"] == account].copy()
            beg_a = tb_beg[tb_beg["__GLA"] == account].copy()
            end_a = tb_end[tb_end["__GLA"] == account].copy()
            rf_a = rf[rf["__GLA"] == account].copy()

            if rf_a.empty:
                raise ValueError(f"El GLA '{account}' no existe en el Rollforward.")

            # Aggregate each source at rollforward level: GLA + BU
            je_sum = je_a.groupby(["__GLA", "__BU"], dropna=False)["__AMOUNT"].sum().reset_index(name="JE_Sum")
            beg_sum = beg_a.groupby(["__GLA", "__BU"], dropna=False)["__AMOUNT"].sum().reset_index(name="TB_Beg_Amount")
            end_sum = end_a.groupby(["__GLA", "__BU"], dropna=False)["__AMOUNT"].sum().reset_index(name="TB_End_Amount")
            rf_sum = rf_a.groupby(["__GLA", "__BU"], dropna=False).agg(
                RF_Beginning=("__RF_BEG", "sum"),
                RF_Variance=("__RF_VAR", "sum"),
                RF_Ending=("__RF_END", "sum"),
            ).reset_index()

            # Full outer merge so unmatched BUs are visible.
            summary = rf_sum.merge(beg_sum, on=["__GLA", "__BU"], how="outer")
            summary = summary.merge(je_sum, on=["__GLA", "__BU"], how="outer")
            summary = summary.merge(end_sum, on=["__GLA", "__BU"], how="outer")

            amount_cols = ["RF_Beginning", "RF_Variance", "RF_Ending", "TB_Beg_Amount", "JE_Sum", "TB_End_Amount"]
            summary[amount_cols] = summary[amount_cols].fillna(0.0)

            summary["Beg_Difference"] = summary["TB_Beg_Amount"] - summary["RF_Beginning"]
            summary["JE_Difference"] = summary["JE_Sum"] - summary["RF_Variance"]
            summary["End_Difference"] = summary["TB_End_Amount"] - summary["RF_Ending"]
            summary["RF_Equation_Difference"] = (
                summary["RF_Beginning"] + summary["RF_Variance"] - summary["RF_Ending"]
            )

            summary["Beg_Result"] = summary["Beg_Difference"].abs().le(tolerance).map({True: "PASS", False: "FAIL"})
            summary["JE_Result"] = summary["JE_Difference"].abs().le(tolerance).map({True: "PASS", False: "FAIL"})
            summary["End_Result"] = summary["End_Difference"].abs().le(tolerance).map({True: "PASS", False: "FAIL"})
            summary["RF_Result"] = summary["RF_Equation_Difference"].abs().le(tolerance).map({True: "PASS", False: "FAIL"})

            summary.rename(columns={"__GLA": "GLA", "__BU": "Business Unit"}, inplace=True)

            # Add total line, useful but never a substitute for BU-level validation.
            total = {
                "GLA": account,
                "Business Unit": "TOTAL",
            }
            numeric_cols = [
                "RF_Beginning", "TB_Beg_Amount", "Beg_Difference",
                "RF_Variance", "JE_Sum", "JE_Difference",
                "RF_Ending", "TB_End_Amount", "End_Difference",
                "RF_Equation_Difference"
            ]
            for c in numeric_cols:
                total[c] = summary[c].sum()

            total["Beg_Result"] = "PASS" if abs(total["Beg_Difference"]) <= tolerance else "FAIL"
            total["JE_Result"] = "PASS" if abs(total["JE_Difference"]) <= tolerance else "FAIL"
            total["End_Result"] = "PASS" if abs(total["End_Difference"]) <= tolerance else "FAIL"
            total["RF_Result"] = "PASS" if abs(total["RF_Equation_Difference"]) <= tolerance else "FAIL"

            summary = pd.concat([summary, pd.DataFrame([total])], ignore_index=True)

            # Remove helper columns from evidence sheets but keep original columns + source file.
            je_export = je_a.drop(columns=["__GLA", "__BU", "__AMOUNT"], errors="ignore")
            beg_export = beg_a.drop(columns=["__GLA", "__BU", "__AMOUNT"], errors="ignore")
            end_export = end_a.drop(columns=["__GLA", "__BU", "__AMOUNT"], errors="ignore")
            rf_export = rf_a.drop(columns=["__GLA", "__BU", "__RF_BEG", "__RF_VAR", "__RF_END"], errors="ignore")

            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                summary.to_excel(writer, sheet_name="Tracing_Summary", index=False)
                rf_export.to_excel(writer, sheet_name="Rollforward", index=False)
                beg_export.to_excel(writer, sheet_name="TB_Beginning", index=False)
                je_export.to_excel(writer, sheet_name="Journal_Entries", index=False)
                end_export.to_excel(writer, sheet_name="TB_Ending", index=False)

            self.format_workbook(save_path, tolerance)

            self.status.config(text=f"Tracing generado: {save_path}")
            messagebox.showinfo("Tracing generado", f"Archivo creado correctamente:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def format_workbook(self, path, tolerance):
        wb = load_workbook(path)

        # General formatting
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = FILL_HEADER
                cell.alignment = Alignment(horizontal="center")

            # Sensible widths
            for col_cells in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col_cells[0].column)
                for cell in col_cells[:200]:
                    try:
                        max_len = max(max_len, len(str(cell.value or "")))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 11), 35)

        # Highlight original RF mapped columns.
        ws = wb["Rollforward"]
        header_map = {str(c.value): c.column for c in ws[1]}
        rf_colors = {
            self.rf_map.get("Beginning"): FILL_BEG,
            self.rf_map.get("Variance"): FILL_VAR,
            self.rf_map.get("Ending"): FILL_END,
        }
        for header, fill in rf_colors.items():
            if header in header_map:
                col = header_map[header]
                for row in range(1, ws.max_row + 1):
                    ws.cell(row=row, column=col).fill = fill
                    if row == 1:
                        ws.cell(row=row, column=col).font = Font(bold=True)

        # Summary colors
        ws = wb["Tracing_Summary"]
        headers = {str(c.value): c.column for c in ws[1]}
        for name, fill in [
            ("RF_Beginning", FILL_BEG),
            ("TB_Beg_Amount", FILL_BEG),
            ("RF_Variance", FILL_VAR),
            ("JE_Sum", FILL_VAR),
            ("RF_Ending", FILL_END),
            ("TB_End_Amount", FILL_END),
        ]:
            if name in headers:
                for row in range(1, ws.max_row + 1):
                    ws.cell(row=row, column=headers[name]).fill = fill

        for result_col in ["Beg_Result", "JE_Result", "End_Result", "RF_Result"]:
            if result_col in headers:
                col = headers[result_col]
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = FILL_PASS if cell.value == "PASS" else FILL_FAIL
                    cell.font = Font(bold=True)

        # Difference cells also visually flag exceptions.
        for diff_col in ["Beg_Difference", "JE_Difference", "End_Difference", "RF_Equation_Difference"]:
            if diff_col in headers:
                col = headers[diff_col]
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col)
                    try:
                        cell.fill = FILL_PASS if abs(float(cell.value or 0)) <= tolerance else FILL_FAIL
                    except Exception:
                        pass

        # Bold total row.
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=headers.get("Business Unit", 2)).value == "TOTAL":
                for cell in ws[row]:
                    cell.font = Font(bold=True)

        wb.save(path)


if __name__ == "__main__":
    app = App()
    app.mainloop()