"""
================================================================================
MODULE NAME   : HPLC_summary_GUI.py
DESCRIPTION   : Tkinter GUI for collection of HPLC data and analysis to calculate yield, selectivity and conversion.  
================================================================================

AUTHOR        : Dr. Sushila Marlow
EMAIL         : s.marlow@imperial.ac.uk
AFFILIATION   : DIGIBAT / Chemical Engineering / Imperial College London
DATE CREATED  : 2026-01-15
LAST MODIFIED : 2026-05-26
VERSION       : 1.3.0

LICENSE       : BSD 3-clause licence
                Copyright © 2026, Imperial College, London. All rights reserved.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os

# ─────────────────────────────────────────────
# CHEMICAL NAME NORMALISATION MAP
# ─────────────────────────────────────────────
CHEM_ALIASES = {
    "glycerol": "Glycerol",
    "formic acid": "Formic Acid",
    "fa": "Formic Acid",
    "glyceric acid": "Glyceric Acid",
    "gca": "Glyceric Acid",
    "glycolic acid": "Glycolic Acid",
    "glc": "Glycolic Acid",
    "oxalic acid": "Oxalic Acid",
    "oa": "Oxalic Acid",
    "tartronic acid": "Tartronic Acid",
    "ta": "Tartronic Acid",
    "lactic acid": "Lactic Acid",
    "la": "Lactic Acid",
}


def normalise_chemical(name: str):
    return CHEM_ALIASES.get(name.strip().lower(), None)


# ─────────────────────────────────────────────
# RETENTION TIME ROW ORDER
# ─────────────────────────────────────────────
CHEM_ORDER = [
    "Oxalic Acid",
    "Tartronic Acid",
    "Glyceric Acid",
    "Glycolic Acid",
    "Lactic Acid",
    "Formic Acid",
    "Glycerol",
]

# ─────────────────────────────────────────────
# ELECTROCHEMICAL CONSTANTS PER PRODUCT
# ─────────────────────────────────────────────
FARADAY = 96485  # C/mol

# {canonical name: (electrons transferred, carbon atoms)}
CHEM_PROPERTIES = {
    "Oxalic Acid":    (10,      2),
    "Tartronic Acid": (8,       3),
    "Glyceric Acid":  (4,       3),
    "Glycolic Acid":  (10/3,    2),
    "Lactic Acid":    (2,       3),
    "Formic Acid":    (2 + 2/3, 1),
    "Glycerol":       (0,       3),
}

# ─────────────────────────────────────────────
# SHEET SANITISATION
# ─────────────────────────────────────────────
def sanitise_sheet(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Strip all non-textual / non-numeric content from a raw sheet DataFrame.
    Compatible with all pandas versions.
    """
    def clean_cell(val):
        if val is None:
            return None
        if isinstance(val, bool):
            return None
        if isinstance(val, (int, float)):
            if isinstance(val, float) and pd.isna(val):
                return None
            return val
        if isinstance(val, str):
            stripped = val.strip()
            if stripped == "":
                return None
            try:
                return float(stripped)
            except ValueError:
                return stripped
        return None

    cleaned = raw_df.apply(lambda col: col.map(clean_cell))
    cleaned.dropna(how="all", inplace=True)
    cleaned.dropna(axis=1, how="all", inplace=True)
    cleaned.reset_index(drop=True, inplace=True)
    cleaned.columns = range(len(cleaned.columns))
    return cleaned


# ─────────────────────────────────────────────
# EXCEL PARSING
# ─────────────────────────────────────────────
def find_amount_column(data_sheet: pd.DataFrame):
    """
    Scan rows to find the header row containing 'amount' and 'mol'.
    Returns (header_row_index, amount_col_index) or (None, None).
    """
    for row_idx, row in data_sheet.iterrows():
        for col_idx, cell in enumerate(row):
            if isinstance(cell, str):
                cell_lower = cell.lower()
                if "amount" in cell_lower and "mol" in cell_lower:
                    return row_idx, col_idx
    return None, None


def parse_hplc_excel(filepath: str):
    """
    Returns:
        df           – consolidated DataFrame (chemicals × samples)
        clean_sheets – dict of sanitised sheet DataFrames for preview
    """
    ext = os.path.splitext(filepath)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"

    raw_sheets = pd.read_excel(
        filepath, sheet_name=None, header=None, engine=engine
    )
    sheet_names = list(raw_sheets.keys())

    clean_sheets = {
        name: sanitise_sheet(df) for name, df in raw_sheets.items()
    }

    sample_data: dict[str, dict[str, list]] = {}

    for i in range(0, len(sheet_names) - 1, 2):
        name_sheet = clean_sheets[sheet_names[i]]
        data_sheet = clean_sheets[sheet_names[i + 1]]

        # ── Extract sample names ──────────────────────────────────────────
        sample_names = []
        for _, row in name_sheet.iterrows():
            for j, cell in enumerate(row):
                if not isinstance(cell, str):
                    continue
                cell_lower = cell.lower()

                if re.fullmatch(r"sample\s*name\s*:?", cell_lower):
                    for k in range(j + 1, len(row)):
                        candidate = row.iloc[k]
                        if isinstance(candidate, str) and candidate.strip():
                            sample_names.append(candidate.strip())
                            break

                elif "sample name" in cell_lower and ":" in cell:
                    parts = cell.split(":", 1)
                    if len(parts) == 2 and parts[1].strip():
                        sample_names.append(parts[1].strip())

        if not sample_names:
            continue

        # ── Find Amount [mol/l] column ────────────────────────────────────
        header_row_idx, amount_col_idx = find_amount_column(data_sheet)

        if amount_col_idx is None:
            messagebox.showwarning(
                "Column Not Found",
                f"Could not find an 'Amount [mol/l]' column in sheet "
                f"'{sheet_names[i + 1]}'.\nSkipping this sheet pair."
            )
            continue

        # ── Find chemical name column ─────────────────────────────────────
        header_row = data_sheet.loc[header_row_idx]
        name_col_idx = None
        for col_idx, cell in enumerate(header_row):
            if isinstance(cell, str):
                if any(k in cell.lower() for k in ["name", "compound", "substance"]):
                    name_col_idx = col_idx
                    break
        if name_col_idx is None:
            name_col_idx = 0

        # ── Extract values below header ───────────────────────────────────
        chem_rows: dict[str, list] = {}
        data_below = data_sheet[
            data_sheet.index > header_row_idx
        ].reset_index(drop=True)

        for _, row in data_below.iterrows():
            name_cell = row.iloc[name_col_idx] if name_col_idx < len(row) else None
            if not isinstance(name_cell, str):
                continue
            canonical = normalise_chemical(name_cell)
            if canonical is None:
                continue

            amount_cell = row.iloc[amount_col_idx] if amount_col_idx < len(row) else None
            try:
                amount = float(amount_cell)
            except (TypeError, ValueError):
                amount = None

            chem_rows.setdefault(canonical, []).append(amount)

        # Average duplicate chemical entries within this sheet
        averaged_chem_rows: dict[str, float | None] = {}
        for chem, vals in chem_rows.items():
            numeric = [v for v in vals if v is not None]
            averaged_chem_rows[chem] = (
                round(sum(numeric) / len(numeric), 4) if numeric else None
            )

        # ── Map to sample names ───────────────────────────────────────────
        for sname in sample_names:
            if sname not in sample_data:
                sample_data[sname] = {}
            for chem, val in averaged_chem_rows.items():
                sample_data[sname].setdefault(chem, []).append(val)

    # ── Final averaging across duplicate sample entries ───────────────────
    final: dict[str, dict[str, float | None]] = {}
    for sname, chems in sample_data.items():
        final[sname] = {}
        for chem, vals in chems.items():
            numeric = [v for v in vals if v is not None]
            final[sname][chem] = (
                round(sum(numeric) / len(numeric), 4) if numeric else None
            )

    # ── Build DataFrame ───────────────────────────────────────────────────
    all_chems = list(
        dict.fromkeys(c for sd in final.values() for c in sd)
    )
    df = pd.DataFrame(
        index=all_chems, columns=list(final.keys()), dtype=object
    )
    df.index.name = "Chemical"
    for sname, chems in final.items():
        for chem in all_chems:
            df.loc[chem, sname] = chems.get(chem)

    # ── Sort rows by retention time order ────────────────────────────────
    # Known chemicals sorted by retention time; any unknowns go to the bottom
    known_present   = [c for c in CHEM_ORDER if c in df.index]
    unknown_present = [c for c in df.index if c not in CHEM_ORDER]
    df = df.reindex(known_present + unknown_present)

    return df, clean_sheets

# ─────────────────────────────────────────────
# RAW SHEET PREVIEW WINDOW
# ─────────────────────────────────────────────
class SheetPreviewWindow(tk.Toplevel):
    def __init__(self, parent, clean_sheets: dict):
        super().__init__(parent)
        self.title("Raw Sheet Preview (Post-Sanitisation)")
        self.geometry("900x500")
        self.configure(bg="#f0f4f8")

        tk.Label(
            self,
            text="Sheets after all formatting has been stripped. "
                 "Use this to verify the parser sees your data correctly.",
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9, "italic"),
            wraplength=880
        ).pack(pady=(8, 2), padx=10, anchor="w")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=6)

        for sheet_name, df in clean_sheets.items():
            frame = tk.Frame(nb, bg="white")
            nb.add(frame, text=str(sheet_name)[:20])

            xscroll = tk.Scrollbar(frame, orient="horizontal")
            yscroll = tk.Scrollbar(frame, orient="vertical")
            tree = ttk.Treeview(
                frame,
                xscrollcommand=xscroll.set,
                yscrollcommand=yscroll.set,
                show="headings"
            )
            xscroll.config(command=tree.xview)
            yscroll.config(command=tree.yview)
            xscroll.pack(side="bottom", fill="x")
            yscroll.pack(side="right", fill="y")
            tree.pack(fill="both", expand=True)

            if df.empty:
                tree["columns"] = ["(empty)"]
                tree.heading("(empty)", text="(empty sheet)")
                tree.column("(empty)", width=200)
                continue

            cols = [str(c) for c in df.columns]
            tree["columns"] = cols
            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=120, anchor="center", minwidth=60)

            for _, row in df.iterrows():
                tree.insert(
                    "", "end",
                    values=[
                        "" if (v is None or (isinstance(v, float) and pd.isna(v)))
                        else str(v)
                        for v in row
                    ]
                )



# ─────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────
class HPLCApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HPLC Data Summary")
        self.geometry("1100x700")
        self.resizable(True, True)
        self.configure(bg="#f0f4f8")

        self.df: pd.DataFrame | None = None
        self.clean_sheets: dict = {}
        self.display_order: list[str] = []

        self._build_ui()

    # ── UI Construction ───────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        top = tk.Frame(self, bg="#2c3e50", pady=8)
        top.pack(fill="x")

        tk.Label(
            top, text="⚗  HPLC Data Summary",
            font=("Helvetica", 16, "bold"),
            bg="#2c3e50", fg="white"
        ).pack(side="left", padx=16)

        tk.Button(
            top, text="🔍  Preview Raw Sheets",
            command=self._preview_sheets,
            bg="#16a085", fg="white",
            font=("Helvetica", 10, "bold"),
            relief="flat", padx=8, pady=4
        ).pack(side="right", padx=6)

        tk.Button(
            top, text="📂  Open Excel File",
            command=self._open_file,
            bg="#27ae60", fg="white",
            font=("Helvetica", 11, "bold"),
            relief="flat", padx=10, pady=4
        ).pack(side="right", padx=10)

        # File path label
        self.path_var = tk.StringVar(value="No file loaded.")
        tk.Label(
            self, textvariable=self.path_var,
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9, "italic")
        ).pack(anchor="w", padx=12, pady=(4, 0))

        # Main paned area
        paned = tk.PanedWindow(self, orient="horizontal", bg="#f0f4f8")
        paned.pack(fill="both", expand=True, padx=10, pady=6)

        left = tk.LabelFrame(
            paned, text="  Column Order  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=6, pady=6
        )
        paned.add(left, width=260)
        self._build_left_panel(left)

        right = tk.LabelFrame(
            paned, text="  Parsed HPLC Data  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50"
        )
        paned.add(right)
        self._build_table(right)

        # Bottom bar
        bottom = tk.Frame(self, bg="#dce3ea", pady=6)
        bottom.pack(fill="x", side="bottom")

        btn_cfg = [
            ("⤒  Top",       lambda: self._move_column("top"),    "#2c3e50"),
            ("↑  Up",        lambda: self._move_column("up"),     "#2980b9"),
            ("↓  Down",      lambda: self._move_column("down"),   "#2980b9"),
            ("⤓  Bottom",    lambda: self._move_column("bottom"), "#2c3e50"),
        ]
        for text, cmd, colour in btn_cfg:
            tk.Button(
                bottom, text=text, command=cmd,
                bg=colour, fg="white", font=("Helvetica", 10),
                relief="flat", padx=10, pady=2
            ).pack(side="left", padx=4)

        tk.Button(
            bottom, text="💾  Export CSV",
            command=self._export_csv,
            bg="#e67e22", fg="white", font=("Helvetica", 10),
            relief="flat", padx=10, pady=2
        ).pack(side="right", padx=10)

        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(
            bottom, textvariable=self.status_var,
            bg="#dce3ea", fg="#333", font=("Helvetica", 9)
        ).pack(side="right", padx=14)

    # ── Left Panel ────────────────────────────────────────────────────────
    def _build_left_panel(self, parent):
        tk.Label(
            parent,
            text="Select a column then use the\nbuttons below to reorder:",
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9)
        ).pack(anchor="w", pady=(0, 4))

        frame = tk.Frame(parent, bg="#f0f4f8")
        frame.pack(fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical")
        self.col_listbox = tk.Listbox(
            frame,
            yscrollcommand=sb.set,
            selectmode="single",
            font=("Courier", 10),
            bg="white",
            fg="#2c3e50",
            selectbackground="#2980b9",
            selectforeground="white",
            activestyle="none",
            height=28
        )
        sb.config(command=self.col_listbox.yview)
        sb.pack(side="right", fill="y")
        self.col_listbox.pack(side="left", fill="both", expand=True)

    # ── Data Table ────────────────────────────────────────────────────────
    def _build_table(self, parent):
        container = tk.Frame(parent, bg="#f0f4f8")
        container.pack(fill="both", expand=True, padx=4, pady=4)

        xscroll = tk.Scrollbar(container, orient="horizontal")
        yscroll = tk.Scrollbar(container, orient="vertical")

        self.tree = ttk.Treeview(
            container,
            xscrollcommand=xscroll.set,
            yscrollcommand=yscroll.set,
            show="headings"
        )
        xscroll.config(command=self.tree.xview)
        yscroll.config(command=self.tree.yview)
        xscroll.pack(side="bottom", fill="x")
        yscroll.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview.Heading",
            background="#2c3e50", foreground="white",
            font=("Helvetica", 10, "bold")
        )
        style.configure("Treeview", rowheight=24, font=("Courier", 10))
        style.map("Treeview", background=[("selected", "#2980b9")])

    # ── File Loading ──────────────────────────────────────────────────────
    def _open_file(self):
        filepath = filedialog.askopenfilename(
            title="Select HPLC Excel File",
            filetypes=[
                ("Excel Files", "*.xls *.xlsx *.xlsm"),
                ("All Files", "*.*")
            ]
        )
        if not filepath:
            return

        self.path_var.set(f"Loaded: {filepath}")
        self.status_var.set("Sanitising and parsing file…")
        self.update_idletasks()

        try:
            self.df, self.clean_sheets = parse_hplc_excel(filepath)

            if self.df.empty:
                messagebox.showwarning(
                    "No Data Found",
                    "No recognisable HPLC data was found after sanitisation.\n\n"
                    "Click '🔍 Preview Raw Sheets' to inspect what was read."
                )
                self.status_var.set("No data found — check raw sheet preview.")
                return

            self.display_order = list(self.df.columns)
            self._refresh_listbox()
            self._refresh_table()
            self.status_var.set(
                f"✔  {len(self.display_order)} samples · "
                f"{len(self.df)} chemicals detected."
            )
        except Exception as e:
            messagebox.showerror("Parse Error", f"Failed to parse file:\n\n{e}")
            self.status_var.set("Error — see parse error dialog.")

    # ── Sheet Preview ─────────────────────────────────────────────────────
    def _preview_sheets(self):
        if not self.clean_sheets:
            messagebox.showinfo("No File", "Load a file first.")
            return
        SheetPreviewWindow(self, self.clean_sheets)

    # ── Listbox Refresh ───────────────────────────────────────────────────
    def _refresh_listbox(self, keep_selection: int = None):
        """Repopulate the column listbox, optionally restoring a selection."""
        self.col_listbox.delete(0, "end")
        for col in self.display_order:
            self.col_listbox.insert("end", col)

        if keep_selection is not None:
            # Clamp to valid range
            idx = max(0, min(keep_selection, len(self.display_order) - 1))
            self.col_listbox.selection_set(idx)
            self.col_listbox.activate(idx)
            self.col_listbox.see(idx)

    # ── Table Refresh ─────────────────────────────────────────────────────
    def _refresh_table(self):
        """Redraw the Treeview with the current display_order."""
        self.tree.delete(*self.tree.get_children())

        cols = ["Chemical"] + self.display_order
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor="center", minwidth=80)
        self.tree.column("Chemical", width=170, anchor="w")

        for chem in self.df.index:
            row = [chem] + [
                self.df.loc[chem, s] if s in self.df.columns else ""
                for s in self.display_order
            ]
            row = [
                f"{v:.4f}" if isinstance(v, float)
                else ("" if v is None else str(v))
                for v in row
            ]
            self.tree.insert("", "end", values=row)

    # ── Column Reordering ─────────────────────────────────────────────────
    def _move_column(self, direction: str):
        """
        Move the selected column.
        direction: "up" | "down" | "top" | "bottom"
        """
        if self.df is None:
            return

        sel = self.col_listbox.curselection()
        if not sel:
            messagebox.showinfo(
                "No Selection",
                "Please click a column name in the list first."
            )
            return

        idx = sel[0]
        n = len(self.display_order)

        if direction == "up" and idx > 0:
            new_idx = idx - 1
        elif direction == "down" and idx < n - 1:
            new_idx = idx + 1
        elif direction == "top":
            new_idx = 0
        elif direction == "bottom":
            new_idx = n - 1
        else:
            return  # Already at boundary

        # Perform the move
        col = self.display_order.pop(idx)
        self.display_order.insert(new_idx, col)

        # Refresh both widgets, restoring selection at new position
        self._refresh_listbox(keep_selection=new_idx)
        self._refresh_table()
        self.status_var.set(f"Moved '{col}' → position {new_idx + 1} of {n}")

    # ── Export ────────────────────────────────────────────────────────────
    def _export_csv(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load a file first.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            title="Export as CSV"
        )
        if not path:
            return

        export_df = self.df[self.display_order].copy()
        export_df.index.name = "Chemical"
        export_df.to_csv(path)
        self.status_var.set(f"✔  Exported → {os.path.basename(path)}")
        messagebox.showinfo("Exported", f"Saved to:\n{path}")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = HPLCApp()
    app.mainloop()