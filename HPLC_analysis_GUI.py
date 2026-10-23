"""
================================================================================
MODULE NAME   : HPLC_analysis_GUI.py
DESCRIPTION   : Tkinter GUI for collection of HPLC data and calculation of 
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
    "Glycerol"
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
        real_mask    – matching boolean DataFrame: True = real value, False = zero-filled
        clean_sheets – dict of sanitised sheet DataFrames for preview
    """
    ext    = os.path.splitext(filepath)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"

    raw_sheets = pd.read_excel(
        filepath, sheet_name=None, header=None, engine=engine
    )
    sheet_names = list(raw_sheets.keys())

    clean_sheets = {
        name: sanitise_sheet(df) for name, df in raw_sheets.items()
    }

    sample_data:      dict[str, dict[str, list]]       = {}
    sample_data_real: dict[str, dict[str, list[bool]]] = {}  # ← tracks realness

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

        # ── If no Amount column found at all, zero-fill entire sheet pair ──
        if amount_col_idx is None:
            for sname in sample_names:
                if sname not in sample_data:
                    sample_data[sname]      = {}
                    sample_data_real[sname] = {}
                for chem in CHEM_ORDER:
                    sample_data[sname].setdefault(chem, []).append(0.0)
                    sample_data_real[sname].setdefault(chem, []).append(False)
            continue

        # ── Find chemical name column ─────────────────────────────────────
        header_row   = data_sheet.loc[header_row_idx]
        name_col_idx = None
        for col_idx, cell in enumerate(header_row):
            if isinstance(cell, str):
                if any(k in cell.lower() for k in ["name", "compound", "substance"]):
                    name_col_idx = col_idx
                    break
        if name_col_idx is None:
            name_col_idx = 0

        # ── Extract values below header ───────────────────────────────────
        chem_rows:      dict[str, list]       = {}
        chem_rows_real: dict[str, list[bool]] = {}

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
                amount   = float(amount_cell)
                is_real  = True
            except (TypeError, ValueError):
                amount   = 0.0   # ← missing value → zero-fill
                is_real  = False

            chem_rows.setdefault(canonical, []).append(amount)
            chem_rows_real.setdefault(canonical, []).append(is_real)

        # Average duplicate chemical entries; real only if ALL were real
        averaged_chem_rows:      dict[str, float] = {}
        averaged_chem_rows_real: dict[str, bool]  = {}

        for chem, vals in chem_rows.items():
            numeric = [v for v in vals if v is not None]
            averaged_chem_rows[chem]      = (
                round(sum(numeric) / len(numeric), 4) if numeric else 0.0
            )
            averaged_chem_rows_real[chem] = all(chem_rows_real.get(chem, [False]))

        # ── For chemicals entirely absent from this sheet → zero-fill ─────
        for chem in CHEM_ORDER:
            if chem not in averaged_chem_rows:
                averaged_chem_rows[chem]      = 0.0
                averaged_chem_rows_real[chem] = False

        # ── Map to sample names ───────────────────────────────────────────
        for sname in sample_names:
            if sname not in sample_data:
                sample_data[sname]      = {}
                sample_data_real[sname] = {}
            for chem, val in averaged_chem_rows.items():
                sample_data[sname].setdefault(chem, []).append(val)
                sample_data_real[sname].setdefault(chem, []).append(
                    averaged_chem_rows_real[chem]
                )

    # ── Final averaging across duplicate sample entries ───────────────────
    final:      dict[str, dict[str, float]] = {}
    final_real: dict[str, dict[str, bool]]  = {}

    for sname, chems in sample_data.items():
        final[sname]      = {}
        final_real[sname] = {}
        for chem, vals in chems.items():
            numeric = [v for v in vals if v is not None]
            final[sname][chem]      = (
                round(sum(numeric) / len(numeric), 4) if numeric else 0.0
            )
            # Real only if every contributing entry was real
            final_real[sname][chem] = all(
                sample_data_real[sname].get(chem, [False])
            )

    # ── Build DataFrames ──────────────────────────────────────────────────
    all_chems = list(
        dict.fromkeys(c for sd in final.values() for c in sd)
    )

    df        = pd.DataFrame(index=all_chems, columns=list(final.keys()), dtype=object)
    real_mask = pd.DataFrame(index=all_chems, columns=list(final.keys()), dtype=bool)
    df.index.name        = "Chemical"
    real_mask.index.name = "Chemical"

    for sname, chems in final.items():
        for chem in all_chems:
            df.loc[chem, sname]        = chems.get(chem, 0.0)
            real_mask.loc[chem, sname] = final_real[sname].get(chem, False)

    # ── Sort rows by retention time order ────────────────────────────────
    known_present   = [c for c in CHEM_ORDER if c in df.index]
    unknown_present = [c for c in df.index   if c not in CHEM_ORDER]
    df        = df.reindex(known_present + unknown_present)
    real_mask = real_mask.reindex(known_present + unknown_present)

    return df, real_mask, clean_sheets
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
# ANALYSIS WINDOW
# ─────────────────────────────────────────────
# ─────────────────────────────────────────────
# ANALYSIS WINDOW
# ─────────────────────────────────────────────
class AnalysisWindow(tk.Toplevel):
    def __init__(self, parent, df: pd.DataFrame):
        super().__init__(parent)
        self.title("Electrochemical Analysis")
        self.geometry("1050x780")
        self.configure(bg="#f0f4f8")
        self.df = df.copy()
        self.real_mask = real_mask

        # ── State ─────────────────────────────────────────────────────────
        self.std_vars:    dict[str, tk.BooleanVar] = {}  # standard checkboxes
        self.keep_vars:   dict[str, tk.BooleanVar] = {}  # include checkboxes
        self.pair_rows:   list[dict]               = []  # list of pair-row widgets
        self.avg_glycerol: float                   = 0.0
        self.pairs:       list[tuple[str, str]]    = []  # confirmed (A, C) pairs
        self.results:     dict                     = {}
        self.sample_widgets: dict[str, dict]       = {}

        self._build_ui()

    # ─────────────────────────────────────────────
    # TOP-LEVEL UI
    # ─────────────────────────────────────────────
    def _build_ui(self):
        top = tk.Frame(self, bg="#2c3e50", pady=8)
        top.pack(fill="x")
        tk.Label(
            top, text="⚡  Electrochemical Analysis",
            font=("Helvetica", 15, "bold"),
            bg="#2c3e50", fg="white"
        ).pack(side="left", padx=16)
        tk.Button(
            top, text="💾  Export to Excel",
            command=self._export_excel,
            bg="#e67e22", fg="white",
            font=("Helvetica", 10, "bold"),
            relief="flat", padx=8, pady=3
        ).pack(side="right", padx=8)
        tk.Button(
            top, text="▶  Run Calculations",
            command=self._run_calculations,
            bg="#27ae60", fg="white",
            font=("Helvetica", 10, "bold"),
            relief="flat", padx=8, pady=3
        ).pack(side="right", padx=4)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=6)
        self._build_setup_tab()

    # ─────────────────────────────────────────────
    # SETUP TAB
    # ─────────────────────────────────────────────
    def _build_setup_tab(self):
        outer = tk.Frame(self.nb, bg="#f0f4f8")
        self.nb.add(outer, text="  ⚙ Setup  ")

        canvas = tk.Canvas(outer, bg="#f0f4f8", highlightthickness=0)
        sb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        frame = tk.Frame(canvas, bg="#f0f4f8")
        cw = canvas.create_window((0, 0), window=frame, anchor="nw")

        frame.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(cw, width=e.width))

        # ── Section 1: Column selection ───────────────────────────────────
        s1 = tk.LabelFrame(
            frame,
            text="  Step 1 — Select Columns & Mark Standards  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=8, pady=8
        )
        s1.pack(fill="x", padx=12, pady=(10, 4))

        legend_f = tk.Frame(s1, bg="#f0f4f8")
        legend_f.pack(anchor="w", pady=(0, 6))
        for colour, meaning in [
            ("#27ae60", "✔ Included in analysis"),
            ("#f39c12", "★ Marked as Glycerol Standard"),
            ("#c0392b", "✘ Excluded from analysis"),
        ]:
            tk.Label(
                legend_f, text=meaning,
                bg="#f0f4f8", fg=colour,
                font=("Helvetica", 9, "bold")
            ).pack(side="left", padx=12)

        tk.Label(
            s1,
            text=(
                "Check 'Std' on exactly two columns to mark them as "
                "Glycerol Standards.\n"
                "Uncheck 'Use' to exclude a column from analysis entirely."
            ),
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9)
        ).pack(anchor="w", pady=(0, 6))

        # Header row
        hdr = tk.Frame(s1, bg="#dce3ea")
        hdr.pack(fill="x")
        for txt, w in [("Column Name", 22), ("Use", 5), ("Standard", 8)]:
            tk.Label(
                hdr, text=txt, width=w,
                bg="#2c3e50", fg="white",
                font=("Helvetica", 9, "bold"),
                anchor="center", relief="flat", pady=3
            ).pack(side="left", padx=1)

        # One row per column
        self.col_row_frames: dict[str, tk.Frame] = {}
        for col in self.df.columns:
            row_f = tk.Frame(s1, bg="white", pady=2)
            row_f.pack(fill="x", pady=1)
            self.col_row_frames[col] = row_f

            tk.Label(
                row_f, text=col, width=22,
                bg="white", fg="#2c3e50",
                font=("Helvetica", 10), anchor="w"
            ).pack(side="left", padx=(4, 0))

            use_var = tk.BooleanVar(value=True)
            self.keep_vars[col] = use_var
            tk.Checkbutton(
                row_f, variable=use_var,
                bg="white", activebackground="white",
                selectcolor="white", width=4,
                command=lambda c=col: self._on_use_toggle(c)
            ).pack(side="left")

            std_var = tk.BooleanVar(value=False)
            self.std_vars[col] = std_var
            tk.Checkbutton(
                row_f, variable=std_var,
                bg="white", activebackground="white",
                selectcolor="white", width=6,
                command=lambda c=col: self._on_std_toggle(c)
            ).pack(side="left")

        # ── Section 2: Standard average ───────────────────────────────────
        s2 = tk.LabelFrame(
            frame,
            text="  Step 2 — Glycerol Standard Average  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=8, pady=6
        )
        s2.pack(fill="x", padx=12, pady=4)

        self.std_avg_var = tk.StringVar(value="— (mark 2 standards above first)")
        tk.Label(
            s2,
            text="Average Glycerol from standards (mol/L):",
            bg="#f0f4f8", fg="#333", font=("Helvetica", 10)
        ).pack(side="left")
        tk.Label(
            s2, textvariable=self.std_avg_var,
            bg="#f0f4f8", fg="#f39c12",
            font=("Helvetica", 11, "bold")
        ).pack(side="left", padx=10)
        tk.Button(
            s2, text="🔄 Calculate",
            command=self._calculate_standard_avg,
            bg="#2980b9", fg="white",
            font=("Helvetica", 9), relief="flat", padx=6
        ).pack(side="left")

        # ── Section 3: Pair matching ──────────────────────────────────────
        s3 = tk.LabelFrame(
            frame,
            text="  Step 3 — Match Anodic / Cathodic Pairs  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=8, pady=8
        )
        s3.pack(fill="x", padx=12, pady=4)

        tk.Label(
            s3,
            text=(
                "Each analysis sample = one Anodic reading + one Cathodic "
                "reading.\nProduct amounts will be SUMMED across the pair.\n"
                "Click 'Auto-Pair' to pair them in order, or set each "
                "dropdown manually."
            ),
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9)
        ).pack(anchor="w", pady=(0, 6))

        btn_row = tk.Frame(s3, bg="#f0f4f8")
        btn_row.pack(anchor="w", pady=(0, 6))
        tk.Button(
            btn_row, text="➕  Add Pair Row",
            command=self._add_pair_row,
            bg="#2980b9", fg="white",
            font=("Helvetica", 9), relief="flat", padx=6
        ).pack(side="left", padx=4)
        tk.Button(
            btn_row, text="⚡  Auto-Pair in Order",
            command=self._auto_pair,
            bg="#16a085", fg="white",
            font=("Helvetica", 9), relief="flat", padx=6
        ).pack(side="left", padx=4)
        tk.Button(
            btn_row, text="🗑  Clear All Pairs",
            command=self._clear_pairs,
            bg="#c0392b", fg="white",
            font=("Helvetica", 9), relief="flat", padx=6
        ).pack(side="left", padx=4)

        # Pair header
        pair_hdr = tk.Frame(s3, bg="#2c3e50")
        pair_hdr.pack(fill="x", pady=(0, 2))
        for txt, w in [
            ("#", 3), ("Sample Name", 16),
            ("Anodic Reading", 20), ("Cathodic Reading", 20), ("", 4)
        ]:
            tk.Label(
                pair_hdr, text=txt, width=w,
                bg="#2c3e50", fg="white",
                font=("Helvetica", 9, "bold"),
                anchor="center", pady=3
            ).pack(side="left", padx=1)

        # Scrollable pair list
        pair_outer = tk.Frame(s3, bg="#f0f4f8")
        pair_outer.pack(fill="x")
        pair_canvas = tk.Canvas(
            pair_outer, bg="#f0f4f8",
            highlightthickness=0, height=200
        )
        pair_sb = tk.Scrollbar(
            pair_outer, orient="vertical",
            command=pair_canvas.yview
        )
        pair_canvas.configure(yscrollcommand=pair_sb.set)
        pair_sb.pack(side="right", fill="y")
        pair_canvas.pack(side="left", fill="both", expand=True)

        self.pair_list_frame = tk.Frame(pair_canvas, bg="#f0f4f8")
        pair_cw = pair_canvas.create_window(
            (0, 0), window=self.pair_list_frame, anchor="nw"
        )
        self.pair_list_frame.bind(
            "<Configure>",
            lambda e: pair_canvas.configure(
                scrollregion=pair_canvas.bbox("all"))
        )
        pair_canvas.bind(
            "<Configure>",
            lambda e: pair_canvas.itemconfig(pair_cw, width=e.width)
        )

        # ── Section 4: Shared parameters ─────────────────────────────────
        s4 = tk.LabelFrame(
            frame,
            text="  Step 4 — Shared Experimental Parameters  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=8, pady=8
        )
        s4.pack(fill="x", padx=12, pady=4)

        tk.Label(
            s4,
            text="Default values applied to all samples (overridable per tab).",
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9)
        ).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 4))

        self.shared_inputs: dict[str, tk.StringVar] = {}
        fields = [
            ("Current Density (mA/cm²):", "current_density"),
            ("Electrolyte Volume (mL):",  "volume"),
            ("Time Elapsed (s):",         "time"),
        ]
        for i, (label, key) in enumerate(fields):
            tk.Label(
                s4, text=label,
                bg="#f0f4f8", fg="#333", font=("Helvetica", 10)
            ).grid(row=1, column=i * 2, sticky="e", padx=(10, 2))
            var = tk.StringVar()
            self.shared_inputs[key] = var
            tk.Entry(
                s4, textvariable=var,
                width=10, font=("Helvetica", 10)
            ).grid(row=1, column=i * 2 + 1, sticky="w", padx=(0, 10))

        # ── Confirm button ────────────────────────────────────────────────
        tk.Button(
            frame,
            text="▶  Confirm Setup & Build Sample Tabs",
            command=self._build_sample_tabs,
            bg="#8e44ad", fg="white",
            font=("Helvetica", 11, "bold"),
            relief="flat", padx=12, pady=6
        ).pack(pady=10)

        self.setup_status = tk.StringVar(value="")
        tk.Label(
            frame, textvariable=self.setup_status,
            bg="#f0f4f8", fg="#c0392b",
            font=("Helvetica", 9, "italic")
        ).pack()

    # ─────────────────────────────────────────────
    # COLUMN TOGGLE CALLBACKS
    # ─────────────────────────────────────────────
    def _on_use_toggle(self, col: str):
        """Grey out a row when excluded."""
        included = self.keep_vars[col].get()
        bg = "white" if included else "#f5f5f5"
        fg = "#2c3e50" if included else "#aaa"
        f = self.col_row_frames[col]
        for widget in f.winfo_children():
            try:
                widget.configure(bg=bg, fg=fg)
            except tk.TclError:
                try:
                    widget.configure(bg=bg)
                except tk.TclError:
                    pass

    def _on_std_toggle(self, col: str):
        """
        Highlight standard rows in orange.
        Enforce maximum of 2 standards.
        """
        marked = [c for c, v in self.std_vars.items() if v.get()]
        if len(marked) > 2:
            # Uncheck the one just clicked
            self.std_vars[col].set(False)
            messagebox.showwarning(
                "Too Many Standards",
                "You can only mark exactly 2 columns as Glycerol Standards."
            )
            return

        for c, f in self.col_row_frames.items():
            if self.std_vars[c].get():
                bg, fg = "#fff8e7", "#f39c12"
            elif not self.keep_vars[c].get():
                bg, fg = "#f5f5f5", "#aaa"
            else:
                bg, fg = "white", "#2c3e50"
            for widget in f.winfo_children():
                try:
                    widget.configure(bg=bg, fg=fg)
                except tk.TclError:
                    try:
                        widget.configure(bg=bg)
                    except tk.TclError:
                        pass

    # ─────────────────────────────────────────────
    # STANDARD AVERAGE
    # ─────────────────────────────────────────────
    def _calculate_standard_avg(self):
        std_cols = [c for c, v in self.std_vars.items() if v.get()]
        if len(std_cols) != 2:
            self.std_avg_var.set("⚠  Mark exactly 2 standard columns first.")
            return

        vals = []
        for col in std_cols:
            v = self.df.loc["Glycerol", col] \
                if "Glycerol" in self.df.index else None
            try:
                vals.append(float(v))
            except (TypeError, ValueError):
                pass

        if vals:
            self.avg_glycerol = round(sum(vals), 6)
            self.std_avg_var.set(f"{self.avg_glycerol}  mol/L")
        else:
            self.std_avg_var.set("No Glycerol data in selected standards.")

    # ─────────────────────────────────────────────
    # PAIR MANAGEMENT
    # ─────────────────────────────────────────────
    def _sample_col_options(self) -> list[str]:
        """Columns that are included and NOT marked as standards."""
        return [
            c for c in self.df.columns
            if self.keep_vars[c].get() and not self.std_vars[c].get()
        ]

    def _add_pair_row(self, anodic: str = "", cathodic: str = "",
                      name: str = ""):
        options = self._sample_col_options()
        if not options:
            messagebox.showwarning(
                "No Columns",
                "No usable columns found.\n"
                "Make sure columns are checked and standards are marked."
            )
            return

        idx = len(self.pair_rows) + 1
        row_f = tk.Frame(
            self.pair_list_frame,
            bg="#eaf4fb" if idx % 2 == 0 else "white",
            pady=3
        )
        row_f.pack(fill="x", pady=1)

        # Row number
        tk.Label(
            row_f, text=str(idx), width=3,
            bg=row_f["bg"], fg="#888",
            font=("Helvetica", 9)
        ).pack(side="left", padx=2)

        # Sample name entry
        name_var = tk.StringVar(value=name if name else f"Sample {idx}")
        tk.Entry(
            row_f, textvariable=name_var,
            width=16, font=("Helvetica", 10)
        ).pack(side="left", padx=4)

        # Anodic dropdown
        anodic_var = tk.StringVar(
            value=anodic if anodic in options else options[0]
        )
        ttk.Combobox(
            row_f, textvariable=anodic_var,
            values=options, width=18,
            font=("Helvetica", 10), state="readonly"
        ).pack(side="left", padx=4)

        # Cathodic dropdown
        cathodic_var = tk.StringVar(
            value=cathodic if cathodic in options else options[0]
        )
        ttk.Combobox(
            row_f, textvariable=cathodic_var,
            values=options, width=18,
            font=("Helvetica", 10), state="readonly"
        ).pack(side="left", padx=4)

        # Remove button
        def remove(f=row_f, entry=None):
            f.destroy()
            self.pair_rows = [
                p for p in self.pair_rows if p["frame"] != f
            ]
        tk.Button(
            row_f, text="✕",
            command=remove,
            bg="#c0392b", fg="white",
            font=("Helvetica", 9), relief="flat",
            width=3
        ).pack(side="left", padx=2)

        self.pair_rows.append({
            "frame":    row_f,
            "name":     name_var,
            "anodic":   anodic_var,
            "cathodic": cathodic_var,
        })

    def _auto_pair(self):
        """Pair available columns sequentially: 1+2, 3+4, 5+6 …"""
        self._clear_pairs()
        options = self._sample_col_options()
        if len(options) < 2:
            messagebox.showwarning(
                "Not Enough Columns",
                "Need at least 2 non-standard columns to auto-pair."
            )
            return
        for i in range(0, len(options) - 1, 2):
            self._add_pair_row(
                anodic=options[i],
                cathodic=options[i + 1],
                name=f"Sample {i // 2 + 1}"
            )

    def _clear_pairs(self):
        for p in self.pair_rows:
            p["frame"].destroy()
        self.pair_rows = []

    # ─────────────────────────────────────────────
    # BUILD SAMPLE TABS
    # ─────────────────────────────────────────────
    def _build_sample_tabs(self):
        # Validate shared inputs
        for key, var in self.shared_inputs.items():
            try:
                float(var.get())
            except ValueError:
                self.setup_status.set(
                    f"⚠  Enter a valid number for: {key}"
                )
                return

        if self.avg_glycerol == 0.0:
            self._calculate_standard_avg()

        if not self.pair_rows:
            self.setup_status.set("⚠  Add at least one pair first.")
            return

        # Collect confirmed pairs
        self.pairs = []
        for p in self.pair_rows:
            a = p["anodic"].get()
            c = p["cathodic"].get()
            n = p["name"].get().strip() or f"Sample {len(self.pairs)+1}"
            self.pairs.append((n, a, c))

        # Remove old sample / result tabs
        tabs = list(self.nb.tabs())
        for tab in tabs[1:]:
            txt = self.nb.tab(tab, "text").strip()
            if txt not in ("📊 Summary", "🧪 Lactic Acid"):
                self.nb.forget(tab)

        self.sample_widgets = {}
        for name, anodic, cathodic in self.pairs:
            self._build_sample_tab(name, anodic, cathodic)

        self.setup_status.set(
            f"✔  {len(self.pairs)} sample tabs built."
        )

    # ─────────────────────────────────────────────
    # SINGLE SAMPLE TAB
    # ─────────────────────────────────────────────
    def _build_sample_tab(self, name: str, anodic: str, cathodic: str):
        outer = tk.Frame(self.nb, bg="#f0f4f8")
        self.nb.add(outer, text=f"  {name}  ")

        canvas = tk.Canvas(outer, bg="#f0f4f8", highlightthickness=0)
        sb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        frame = tk.Frame(canvas, bg="#f0f4f8")
        cw = canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(cw, width=e.width))

        widgets = {}

        # ── Pair info banner ──────────────────────────────────────────────
        banner = tk.Frame(frame, bg="#eaf4fb", pady=4)
        banner.pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(
            banner,
            text=(
                f"Anodic:  {anodic}     ➕     "
                f"Cathodic:  {cathodic}"
            ),
            bg="#eaf4fb", fg="#2980b9",
            font=("Helvetica", 10, "bold")
        ).pack()
        tk.Label(
            banner,
            text="Product amounts = Anodic value + Cathodic value",
            bg="#eaf4fb", fg="#555",
            font=("Helvetica", 9, "italic")
        ).pack()


        # ── Initial glycerol ──────────────────────────────────────────────
        gly_lf = tk.LabelFrame(
            frame, text="  Glycerol Concentrations  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=8, pady=6
        )
        gly_lf.pack(fill="x", padx=12, pady=4)

        # Initial glycerol (editable)
        tk.Label(
            gly_lf,
            text="Initial [Glycerol] (mol/L) — from standards:",
            bg="#f0f4f8", fg="#333", font=("Helvetica", 10)
        ).grid(row=0, column=0, sticky="e", padx=(4, 4), pady=4)
        init_gly_var = tk.StringVar(value=str(self.avg_glycerol))
        tk.Entry(
            gly_lf, textvariable=init_gly_var,
            width=14, font=("Helvetica", 10)
        ).grid(row=0, column=1, sticky="w", padx=(0, 20), pady=4)
        widgets["init_glycerol"] = init_gly_var

        # Sample glycerol (read-only, populated after calculation)
        tk.Label(
            gly_lf,
            text="Sample [Glycerol] (mol/L) — anodic + cathodic:",
            bg="#f0f4f8", fg="#333", font=("Helvetica", 10)
        ).grid(row=0, column=2, sticky="e", padx=(4, 4), pady=4)
        sample_gly_lbl = tk.Label(
            gly_lf, text="—",
            bg="#e8f8f5", fg="#16a085",
            font=("Helvetica", 10, "bold"),
            width=14, relief="sunken", anchor="center"
        )
        sample_gly_lbl.grid(row=0, column=3, sticky="w", padx=(0, 4), pady=4)
        widgets["sample_glycerol_lbl"] = sample_gly_lbl

        # ── Charge parameters ─────────────────────────────────────────────
        chg_lf = tk.LabelFrame(
            frame, text="  Charge Parameters  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=8, pady=6
        )
        chg_lf.pack(fill="x", padx=12, pady=4)
        tk.Label(
            chg_lf,
            text="Override shared values for this sample if needed:",
            bg="#f0f4f8", fg="#777", font=("Helvetica", 9, "italic")
        ).grid(row=0, column=0, columnspan=6, sticky="w")

        for i, (label, key) in enumerate([
            ("Current Density (mA/cm²):", "current_density"),
            ("Electrolyte Volume (mL):",  "volume"),
            ("Time Elapsed (s):",         "time"),
        ]):
            tk.Label(
                chg_lf, text=label,
                bg="#f0f4f8", fg="#333", font=("Helvetica", 10)
            ).grid(row=1, column=i * 2, sticky="e", padx=(10, 2), pady=4)
            var = tk.StringVar(value=self.shared_inputs[key].get())
            tk.Entry(
                chg_lf, textvariable=var,
                width=10, font=("Helvetica", 10)
            ).grid(row=1, column=i * 2 + 1, sticky="w", padx=(0, 10))
            widgets[key] = var

        # ── Results grid ──────────────────────────────────────────────────
        res_lf = tk.LabelFrame(
            frame, text="  Results  ",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4f8", fg="#2c3e50", padx=8, pady=6
        )
        res_lf.pack(fill="both", expand=True, padx=12, pady=4)

        headers = [
            "Chemical", "Anodic\n(mol/L)", "Cathodic\n(mol/L)",
            "Total\n(mol/L)", "Carbon\n(mol/L)",
            "Yield\n(%)", "Selectivity\n(%)", "Faradaic Eff.\n(%)"
        ]
        for c, h in enumerate(headers):
            tk.Label(
                res_lf, text=h,
                bg="#2c3e50", fg="white",
                font=("Helvetica", 9, "bold"),
                relief="flat", padx=4, pady=4,
                anchor="center", width=13
            ).grid(row=0, column=c, padx=1, pady=1, sticky="nsew")

        result_labels = {}
        display_chems = [c for c in CHEM_ORDER if c != "Glycerol"]

        for r, chem in enumerate(display_chems, start=1):
            bg = "#eaf4fb" if r % 2 == 0 else "white"
            tk.Label(
                res_lf, text=chem,
                bg=bg, fg="#2c3e50",
                font=("Helvetica", 10), anchor="w",
                padx=6, width=13
            ).grid(row=r, column=0, padx=1, pady=1, sticky="nsew")

            row_labels = {}
            for c, key in enumerate(
                ["anodic", "cathodic", "total", "carbon",
                 "yield", "selectivity", "faradaic"],
                start=1
            ):
                lbl = tk.Label(
                    res_lf, text="—",
                    bg=bg, fg="#333",
                    font=("Courier", 10),
                    anchor="center", width=13
                )
                lbl.grid(row=r, column=c, padx=1, pady=1, sticky="nsew")
                row_labels[key] = lbl
            result_labels[chem] = row_labels

        widgets["result_labels"] = result_labels

        # ── Summary metrics ───────────────────────────────────────────────
        sum_f = tk.Frame(frame, bg="#f0f4f8")
        sum_f.pack(fill="x", padx=12, pady=6)
        for i, (label, key) in enumerate([
            ("Conversion (%):",      "conversion"),
            ("Carbon Balance (%):",  "carbon_balance"),
            ("Total Charge (C):",    "charge"),
        ]):
            tk.Label(
                sum_f, text=label,
                bg="#f0f4f8", fg="#2c3e50",
                font=("Helvetica", 10, "bold")
            ).grid(row=0, column=i * 2, sticky="e", padx=(16, 4))
            lbl = tk.Label(
                sum_f, text="—",
                bg="#f0f4f8", fg="#e67e22",
                font=("Helvetica", 11, "bold")
            )
            lbl.grid(row=0, column=i * 2 + 1, sticky="w")
            widgets[key + "_lbl"] = lbl

        # Store anodic/cathodic column names for calculation
        widgets["anodic_col"]  = anodic
        widgets["cathodic_col"] = cathodic
        self.sample_widgets[name] = widgets

    # ─────────────────────────────────────────────
    # CALCULATIONS
    # ─────────────────────────────────────────────
    def _run_calculations(self):
        if not self.sample_widgets:
            messagebox.showinfo("Setup First", "Complete setup first.")
            return

        self.results = {}
        errors = []

        for sname, widgets in self.sample_widgets.items():
            try:
                init_glycerol    = float(widgets["init_glycerol"].get())
                current_density  = float(widgets["current_density"].get())
                volume           = float(widgets["volume"].get())
                time_s           = float(widgets["time"].get())
                anodic_col       = widgets["anodic_col"]
                cathodic_col     = widgets["cathodic_col"]

                # Charge passed
                charge = (current_density * time_s * 5 * 2)/1000

                # Sample glycerol = sum of anodic + cathodic glycerol
                def get_val(col, chem):
                    v = self.df.loc[chem, col] \
                        if chem in self.df.index else None
                    try:
                        return float(v)
                    except (TypeError, ValueError):
                        return 0.0

                sample_glycerol = (
                    get_val(anodic_col,   "Glycerol")
                    + get_val(cathodic_col, "Glycerol")
                )

                sample_results = {
                    "charge":          charge,
                    "init_glycerol":   init_glycerol,
                    "sample_glycerol": sample_glycerol,
                    "products":        {}
                }

                display_chems = [c for c in CHEM_ORDER if c != "Glycerol"]

                for chem in display_chems:
                    electrons, carbons = CHEM_PROPERTIES.get(chem, (1, 1))

                    # Sum anodic + cathodic amounts
                    val_a = get_val(anodic_col,   chem)
                    val_c = get_val(cathodic_col, chem)
                    amount = val_a + val_c

                    # Yield %
                    yld = (
                        (carbons * amount * 100) / (init_glycerol * 3)
                        if init_glycerol > 0 else 0.0
                    )

                    # Selectivity %
                    denom_sel = init_glycerol - sample_glycerol
                    sel = (
                        (amount * 100) / denom_sel
                        if denom_sel > 0 else 0.0
                    )

                    # Faradaic efficiency %
                    fe = (
                        (amount * (volume / 1000) * 2 * FARADAY * electrons*100)
                        / (charge)
                        if charge > 0 else 0.0
                    )
                    
                    # Explicit carbon content per product (mol/L)
                    CARBON_MULTIPLIERS = {
                        "Oxalic Acid":    2,
                        "Tartronic Acid": 3,
                        "Glyceric Acid":  3,
                        "Glycolic Acid":  2,
                        "Lactic Acid":    3,
                        "Formic Acid":    1,
                        "Glycerol":       3,
                    }
                    carbon_content = amount * CARBON_MULTIPLIERS.get(chem, carbons)

                    sample_results["products"][chem] = {
                        "anodic":      round(val_a,          6),
                        "cathodic":    round(val_c,          6),
                        "amount":      round(amount,         6),
                        "carbon":      round(carbon_content, 6),  # ✅ new
                        "yield":       round(yld,            2),
                        "selectivity": round(sel,            2),
                        "faradaic":    round(fe,             2),
                        "electrons":   electrons,
                        "carbons":     carbons,
                    }

                # Conversion %
                conversion = (
                    ((init_glycerol - sample_glycerol) / init_glycerol) * 100
                    if init_glycerol > 0 else 0.0
                )

                # Carbon balance % — uses explicit per-product carbon values
                carbon_sum = sum(
                    p["carbon"]
                    for p in sample_results["products"].values()
                )
                carbon_balance = (
                    (carbon_sum * 100) / (init_glycerol * 3)
                    if init_glycerol > 0 else 0.0
                )

                sample_results["conversion"]     = round(conversion,     2)
                sample_results["carbon_balance"] = round(carbon_balance, 2)
                self.results[sname] = sample_results

                # ── Update result labels ──────────────────────────────────
                for chem, prod in sample_results["products"].items():
                    rl = widgets["result_labels"].get(chem, {})

                    for key, fmt in [
                        ("anodic",      f"{prod['anodic']:.4e}"),
                        ("cathodic",    f"{prod['cathodic']:.4e}"),
                        ("total",       f"{prod['amount']:.4e}"),
                        ("carbon",      f"{prod['carbon']:.4e}"),
                        ("yield",       f"{prod['yield']:.2f}"),
                        ("selectivity", f"{prod['selectivity']:.2f}"),
                        ("faradaic",    f"{prod['faradaic']:.2f}"),
                    ]:
                        if key in rl:
                            rl[key].config(text=fmt)

                # Update sample glycerol display
                widgets["sample_glycerol_lbl"].config(
                    text=f"{sample_glycerol:.6f}")
                widgets["conversion_lbl"].config(
                    text=f"{sample_results['conversion']:.2f} %")
                widgets["carbon_balance_lbl"].config(
                    text=f"{sample_results['carbon_balance']:.2f} %")
                widgets["charge_lbl"].config(
                    text=f"{sample_results['charge']:.4f} C")

            except Exception as e:
                errors.append(f"{sname}: {e}")

        if errors:
            messagebox.showwarning(
                "Calculation Errors",
                "Errors in:\n" + "\n".join(errors)
            )

        self._build_summary_tab()
        self._build_lactic_acid_tab()

    # ─────────────────────────────────────────────
    # SUMMARY TAB
    # ─────────────────────────────────────────────
    def _build_summary_tab(self):
        for tab in self.nb.tabs():
            if self.nb.tab(tab, "text").strip() == "📊 Summary":
                self.nb.forget(tab)
                break

        frame = tk.Frame(self.nb, bg="#f0f4f8")
        self.nb.add(frame, text="  📊 Summary  ")

        tk.Label(
            frame,
            text="Full results for all samples",
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9, "italic")
        ).pack(anchor="w", padx=10, pady=(6, 2))

        container = tk.Frame(frame, bg="#f0f4f8")
        container.pack(fill="both", expand=True, padx=8, pady=4)

        xscroll = tk.Scrollbar(container, orient="horizontal")
        yscroll = tk.Scrollbar(container, orient="vertical")
        tree = ttk.Treeview(
            container,
            xscrollcommand=xscroll.set,
            yscrollcommand=yscroll.set,
            show="headings"
        )
        xscroll.config(command=tree.xview)
        yscroll.config(command=tree.yview)
        xscroll.pack(side="bottom", fill="x")
        yscroll.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)

        display_chems = [c for c in CHEM_ORDER if c != "Glycerol"]
        metric_cols = []
        for chem in display_chems:
            short = chem.split()[0]
            for m in ["Total(mol/L)", "Yld%", "Sel%", "FE%"]:
                metric_cols.append(f"{short} {m}")

        cols = (
            ["Sample", "Charge(C)", "Conv%", "CarbBal%"] + metric_cols
        )
        tree["columns"] = cols
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=90, anchor="center", minwidth=60)
        tree.column("Sample", width=120, anchor="w")

        for sname, res in self.results.items():
            row = [
                sname,
                f"{res['charge']:.2f}",
                f"{res['conversion']:.2f}",
                f"{res['carbon_balance']:.2f}",
            ]
            for chem in display_chems:
                p = res["products"].get(chem, {})
                row += [
                    f"{p.get('amount', 0):.3e}",
                    f"{p.get('yield', 0):.2f}",
                    f"{p.get('selectivity', 0):.2f}",
                    f"{p.get('faradaic', 0):.2f}",
                ]
            tree.insert("", "end", values=row)

    # ─────────────────────────────────────────────
    # LACTIC ACID TAB
    # ─────────────────────────────────────────────
    def _build_lactic_acid_tab(self):
        for tab in self.nb.tabs():
            if self.nb.tab(tab, "text").strip() == "🧪 Lactic Acid":
                self.nb.forget(tab)
                break

        frame = tk.Frame(self.nb, bg="#f0f4f8")
        self.nb.add(frame, text="  🧪 Lactic Acid  ")

        tk.Label(
            frame,
            text="Lactic Acid — Conversion, Selectivity, Yield & "
                 "Faradaic Efficiency per sample",
            bg="#f0f4f8", fg="#555", font=("Helvetica", 9, "italic")
        ).pack(anchor="w", padx=10, pady=(6, 2))

        container = tk.Frame(frame, bg="#f0f4f8")
        container.pack(fill="both", expand=True, padx=8, pady=4)

        xscroll = tk.Scrollbar(container, orient="horizontal")
        yscroll = tk.Scrollbar(container, orient="vertical")
        tree = ttk.Treeview(
            container,
            xscrollcommand=xscroll.set,
            yscrollcommand=yscroll.set,
            show="headings"
        )
        xscroll.config(command=tree.xview)
        yscroll.config(command=tree.yview)
        xscroll.pack(side="bottom", fill="x")
        yscroll.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)

        cols = [
            "Sample", "Conversion (%)", "Selectivity (%)",
            "Yield (%)", "Faradaic Efficiency (%)"
        ]
        tree["columns"] = cols
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=170, anchor="center", minwidth=100)
        tree.column("Sample", width=130, anchor="w")

        for sname, res in self.results.items():
            la = res["products"].get("Lactic Acid", {})
            tree.insert("", "end", values=[
                sname,
                f"{res['conversion']:.2f}",
                f"{la.get('selectivity', 0):.2f}",
                f"{la.get('yield', 0):.2f}",
                f"{la.get('faradaic', 0):.2f}",
            ])

    # ─────────────────────────────────────────────
    # EXPORT
    # ─────────────────────────────────────────────
    def _export_excel(self):
        if not self.results:
            messagebox.showwarning("No Results", "Run calculations first.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Export Analysis to Excel"
        )
        if not path:
            return

        display_chems = [c for c in CHEM_ORDER if c != "Glycerol"]

        with pd.ExcelWriter(path, engine="openpyxl") as writer:

            # Full summary sheet
            rows = []
            for sname, res in self.results.items():
                row = {
                    "Sample":             sname,
                    "Charge (C)":         res["charge"],
                    "Conversion (%)":     res["conversion"],
                    "Carbon Balance (%)": res["carbon_balance"],
                }
                for chem in display_chems:
                    p = res["products"].get(chem, {})
                    s = chem.split()[0]
                    row[f"{s} Anodic (mol/L)"]      = p.get("anodic",      "")
                    row[f"{s} Cathodic (mol/L)"]     = p.get("cathodic",    "")
                    row[f"{s} Total (mol/L)"]        = p.get("amount",      "")
                    row[f"{s} Carbon (mol/L)"]       = p.get("carbon",      "")  # ✅ new
                    row[f"{s} Yield (%)"]            = p.get("yield",       "")
                    row[f"{s} Selectivity (%)"]      = p.get("selectivity", "")
                    row[f"{s} Faradaic Eff. (%)"]    = p.get("faradaic",    "")
                rows.append(row)
            pd.DataFrame(rows).to_excel(
                writer, sheet_name="Full Summary", index=False
            )

            # Lactic acid sheet
            la_rows = []
            for sname, res in self.results.items():
                la = res["products"].get("Lactic Acid", {})
                la_rows.append({
                    "Sample":                  sname,
                    "Conversion (%)":          res["conversion"],
                    "Selectivity (%)":         la.get("selectivity", ""),
                    "Yield (%)":               la.get("yield",       ""),
                    "Faradaic Efficiency (%)": la.get("faradaic",    ""),
                })
            pd.DataFrame(la_rows).to_excel(
                writer, sheet_name="Lactic Acid", index=False
            )

            # Per-sample sheets
            for sname, res in self.results.items():
                s_rows = []
                for chem in display_chems:
                    p = res["products"].get(chem, {})
                    s_rows.append({
                        "Chemical":           chem,
                        "Anodic (mol/L)":     p.get("anodic",      ""),
                        "Cathodic (mol/L)":   p.get("cathodic",    ""),
                        "Total (mol/L)":      p.get("amount",      ""),
                        "Carbon (mol/L)":     p.get("carbon",      ""),  # ✅ new
                        "Yield (%)":          p.get("yield",       ""),
                        "Selectivity (%)":    p.get("selectivity", ""),
                        "Faradaic Eff. (%)":  p.get("faradaic",    ""),
                    })
                summary_row = pd.DataFrame([{
                    "Chemical":         "— SUMMARY —",
                    "Anodic (mol/L)":   f"Conversion: {res['conversion']:.2f}%",
                    "Cathodic (mol/L)": f"Carbon Balance: {res['carbon_balance']:.2f}%",
                    "Total (mol/L)":    f"Charge: {res['charge']:.4f} C",
                    "Yield (%)":        "",
                    "Selectivity (%)":  "",
                    "Faradaic Eff. (%)": "",
                }])
                sheet_df = pd.concat(
                    [pd.DataFrame(s_rows), summary_row],
                    ignore_index=True
                )
                safe = re.sub(r'[\\/*?:\[\]]', '_', sname)[:31]
                sheet_df.to_excel(writer, sheet_name=safe, index=False)

        messagebox.showinfo("Exported", f"Saved to:\n{path}")

# ─────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────
class HPLCApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HPLC Data Analyser")
        self.geometry("1100x700")
        self.resizable(True, True)
        self.configure(bg="#f0f4f8")

        self.df:            pd.DataFrame | None = None
        self.real_mask:     pd.DataFrame | None = None   # ← NEW
        self.clean_sheets:  dict                = {}
        self.display_order: list[str]           = []

        self._build_ui()

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
            # ← now unpacks 3 values
            self.df, self.real_mask, self.clean_sheets = parse_hplc_excel(filepath)

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

    # ── UI Construction ───────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        top = tk.Frame(self, bg="#2c3e50", pady=8)
        top.pack(fill="x")

        tk.Label(
            top, text="⚗  HPLC Data Analyser",
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

        tk.Button(
            top, text="⚡  Analysis",
            command=self._open_analysis,
            bg="#8e44ad", fg="white",
            font=("Helvetica", 10, "bold"),
            relief="flat", padx=8, pady=4
        ).pack(side="right", padx=6)

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
        """
        Redraw the Treeview with the current display_order.
        Real values from the Excel file are shown in bold.
        Auto-filled zeros are shown in normal weight to distinguish them visually.
        """
        self.tree.delete(*self.tree.get_children())

        cols = ["Chemical"] + self.display_order
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor="center", minwidth=80)
        self.tree.column("Chemical", width=170, anchor="w")

        # ── Configure bold / normal tags ──────────────────────────────────────
        self.tree.tag_configure(
            "real",
            font=("Courier", 10, "bold"),
            foreground="#1a252f"
        )
        self.tree.tag_configure(
            "zero_fill",
            font=("Courier", 10),
            foreground="#aaaaaa"   # greyed out to further distinguish
        )
        self.tree.tag_configure(
            "mixed",               # row has both real and zero-filled values
            font=("Courier", 10),
            foreground="#333333"
        )

        for chem in self.df.index:
            row_values  = [chem]
            all_real    = True
            any_real    = False

            for s in self.display_order:
                v = self.df.loc[chem, s] if s in self.df.columns else None

                # Determine if this cell was a real value
                is_real = False
                if self.real_mask is not None and s in self.real_mask.columns:
                    is_real = bool(self.real_mask.loc[chem, s])

                if is_real:
                    any_real = True
                    cell_str = f"{float(v):.4f}" if isinstance(v, (int, float)) else str(v)
                else:
                    all_real = False
                    # Always display as 0.0000 for zero-filled cells
                    cell_str = "0.0000"

                row_values.append(cell_str)

            # Assign row tag based on composition
            if all_real:
                tag = "real"
            elif any_real:
                tag = "mixed"
            else:
                tag = "zero_fill"

            self.tree.insert("", "end", values=row_values, tags=(tag,))

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

    def _open_analysis(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No Data", "Load an HPLC file first.")
            return
        AnalysisWindow(self, self.df, self.real_mask)


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = HPLCApp()
    app.mainloop()