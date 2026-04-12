# ─────────────────────────────────────────────────────────────────────────────
# DBF ↔ Excel Converter
# Developer: Ahmed Abdul Ammer Al-Saatci
#
# Two conversion modes:
#   1. DBF → CSV  : Read dBASE III/IV files and export to comma-separated values.
#   2. Excel → DBF: Read .xlsx / .xls worksheets and write dBASE III binary files.
#
# Arabic encoding reference:
#   cp720  = DOS Arabic           (legacy DBF files created by Arabic DOS programs)
#   cp1256 = Windows Arabic       (modern Arabic Office / Windows applications)
#   cp864  = DOS Arabic alternate (some older Arabic printers / ERP systems)
#   cp1252 = Western European     (last-resort fallback for mixed-encoding files)
# ─────────────────────────────────────────────────────────────────────────────

import datetime
import os
import re
import struct
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from dbfread import DBF


# ─── Constants ────────────────────────────────────────────────────────────────

# Encodings tried in priority order when the source DBF encoding is unknown.
ENCODINGS_TO_TRY = ["cp720", "cp1256", "cp864", "cp1252", "utf-8"]


# ─── DBF → CSV helpers ────────────────────────────────────────────────────────

def read_dbf(path: str) -> pd.DataFrame:
    """
    Read a dBASE (.dbf) file into a pandas DataFrame.

    Iterates through ENCODINGS_TO_TRY until one succeeds, so the function
    handles both legacy DOS-Arabic files (cp720) and modern Windows-Arabic
    files (cp1256) transparently.  Raises RuntimeError if every encoding fails.
    """
    last_exc = None
    for enc in ENCODINGS_TO_TRY:
        try:
            table = DBF(path, encoding=enc, load=True, ignore_missing_memofile=True)
            df = pd.DataFrame(iter(table))
            return df
        except Exception as exc:
            last_exc = exc
    raise RuntimeError(f"Could not read '{os.path.basename(path)}': {last_exc}")


def _decode_bytes(v):
    """
    Decode a raw bytes / bytearray object from a DBF memo or character field.

    Tries several encodings in order (UTF-8 first, then Arabic, then Western
    European).  Falls back to cp1256 with replacement characters so that the
    caller always receives a plain Python str — never raw bytes.
    Non-bytes values are returned unchanged.
    """
    if not isinstance(v, (bytes, bytearray)):
        return v
    for enc in ("utf-8", "cp1256", "cp720", "cp864", "cp1252"):
        try:
            return v.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return v.decode("cp1256", errors="replace")


def convert_dbf_to_csv(dbf_path: str, out_dir: str) -> str:
    """
    Convert a single DBF file to a CSV file saved inside out_dir.

    Steps:
      1. Read the DBF file (auto-detects encoding via read_dbf).
      2. Decode any raw-bytes columns to Unicode strings.
      3. Write the result as a cp1256-encoded CSV so Arabic-locale Excel
         can open the file directly without a BOM or import wizard.

    Returns the absolute path of the newly created CSV file.
    """
    df = read_dbf(dbf_path)

    # Decode byte-type columns to proper Unicode strings
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(_decode_bytes)

    base_name = os.path.splitext(os.path.basename(dbf_path))[0]
    out_path = os.path.join(out_dir, base_name + ".csv")

    # cp1256 = Windows Arabic; matches what Arabic-locale Excel expects
    # when opening a CSV without a BOM.
    df.to_csv(out_path, index=False, encoding="cp1256", errors="replace")
    return out_path


# ─── Excel → DBF helpers ──────────────────────────────────────────────────────

def _sanitize_field_name(name: str) -> str:
    """
    Convert an arbitrary Excel column header to a valid dBASE field name.

    dBASE III field names must be:
      • At most 10 ASCII characters
      • Only letters, digits, and underscores (spaces are replaced with '_')
      • Upper-case (convention, not a hard requirement, but aids compatibility)

    An empty result after stripping is replaced with the placeholder 'FIELD'.
    """
    clean = re.sub(r"[^A-Za-z0-9_]", "", str(name).replace(" ", "_"))
    clean = clean.upper()[:10]
    return clean if clean else "FIELD"


def _infer_dbf_field(series: pd.Series, encoding: str) -> tuple:
    """
    Infer the dBASE III field descriptor (type, length, decimal_count) for a
    pandas Series, honouring the target encoding for character-field sizing.

    Mapping rules:
      bool / boolean  →  'L'  logical,  length 1,  decimals 0
      datetime64      →  'D'  date,     length 8,  decimals 0  (YYYYMMDD)
      integer dtype   →  'N'  numeric,  width of widest value, decimals 0
      float dtype     →  'N'  numeric,  auto-width; if all values are whole
                              numbers the field is written as integer ('N',w,0)
                              otherwise real decimal precision is measured
      anything else   →  'C'  character, max encoded byte-length (≤ 254)

    Returns a 3-tuple: (ftype: str, flen: int, fdecimals: int)
    """
    dtype = series.dtype
    non_null = series.dropna()

    # ── Logical ───────────────────────────────────────────────────────────────
    if pd.api.types.is_bool_dtype(dtype):
        return "L", 1, 0

    # ── Date ──────────────────────────────────────────────────────────────────
    if pd.api.types.is_datetime64_any_dtype(dtype):
        return "D", 8, 0

    # ── Integer ───────────────────────────────────────────────────────────────
    if pd.api.types.is_integer_dtype(dtype):
        if non_null.empty:
            return "N", 10, 0
        vals = [int(v) for v in non_null]
        max_width = max(len(str(abs(v))) for v in vals)
        if min(vals) < 0:
            max_width += 1          # room for the minus sign
        return "N", min(max(max_width, 1), 19), 0

    # ── Float ─────────────────────────────────────────────────────────────────
    if pd.api.types.is_float_dtype(dtype):
        if non_null.empty:
            return "N", 10, 2
        float_vals = [float(v) for v in non_null]

        # If every value is a whole number treat it as an integer field so
        # columns that pandas promoted to float64 due to NaN are stored cleanly.
        if all(abs(v - round(v)) < 1e-9 for v in float_vals):
            max_width = max(len(str(int(abs(v)))) for v in float_vals)
            if min(float_vals) < 0:
                max_width += 1
            return "N", min(max(max_width, 1), 19), 0

        # Count the significant decimal places actually present in the data.
        def _dec_places(v: float) -> int:
            """Count non-zero decimal digits (strips trailing zeros)."""
            s = f"{v:.10f}".rstrip("0")
            return len(s.split(".")[1]) if "." in s else 0

        decimals  = min(max((_dec_places(v) for v in float_vals), default=2), 10)
        int_width = max(len(str(int(abs(v)))) for v in float_vals)
        if min(float_vals) < 0:
            int_width += 1          # minus sign
        # total width = integer digits + decimal point + decimal digits
        total = min(int_width + 1 + decimals, 19)
        return "N", max(total, decimals + 2), decimals

    # ── Character (fallback for object / mixed / string columns) ──────────────
    if non_null.empty:
        return "C", 10, 0
    # Measure the byte-length in the target encoding, not unicode char count,
    # because dBASE stores fixed-width byte arrays (not code-points).
    max_len = max(len(str(v).encode(encoding, errors="replace")) for v in non_null)
    return "C", min(max(max_len, 1), 254), 0


def _encode_dbf_value(val, ftype: str, flen: int, fdecimals: int, encoding: str) -> bytes:
    """
    Encode a single Python value to a fixed-width bytes object suitable for
    writing into a dBASE III data record field.

    NULL / NaN handling:
      • All field types → space-padded blank field  (dBASE convention)
      • Logical ('L')   → b'?' (undefined logical value per dBASE spec)

    Padding conventions:
      'C'       → left-justified, right-padded with spaces
      'N' / 'F' → right-justified, left-padded with spaces
      'D'       → YYYYMMDD, right-padded with spaces if short
      'L'       → single byte b'T' or b'F'
    """
    # ── NULL / NaN detection ──────────────────────────────────────────────────
    is_null = val is None
    if not is_null:
        try:
            is_null = bool(pd.isna(val))
        except (TypeError, ValueError):
            is_null = False

    if is_null:
        return b"?" if ftype == "L" else b" " * flen

    # ── Character ─────────────────────────────────────────────────────────────
    if ftype == "C":
        encoded = str(val).encode(encoding, errors="replace")
        if len(encoded) > flen:
            encoded = encoded[:flen]
        return encoded.ljust(flen, b" ")

    # ── Numeric ───────────────────────────────────────────────────────────────
    if ftype in ("N", "F"):
        try:
            s = str(int(round(float(val)))) if fdecimals == 0 else f"{float(val):.{fdecimals}f}"
        except (ValueError, TypeError):
            s = ""
        raw = s.encode("ascii", errors="replace")
        return raw.rjust(flen, b" ")[:flen]

    # ── Date ──────────────────────────────────────────────────────────────────
    if ftype == "D":
        if isinstance(val, (datetime.datetime, datetime.date)):
            s = val.strftime("%Y%m%d")
        else:
            s = str(val).replace("-", "")[:8]
        return s.encode("ascii", errors="replace").ljust(8, b" ")[:8]

    # ── Logical ───────────────────────────────────────────────────────────────
    if ftype == "L":
        return b"T" if val else b"F"

    # ── Fallback (should not be reached) ──────────────────────────────────────
    return b" " * flen


def excel_to_dbf(excel_path: str, out_dir: str, sheet=0, encoding: str = "cp1256") -> str:
    """
    Convert a single Excel worksheet to a dBASE III compatible .dbf file.

    Algorithm:
      1. Read the worksheet with pandas (openpyxl for .xlsx, xlrd for .xls).
      2. Sanitize column headers to valid 10-char ASCII dBASE field names and
         deduplicate any collisions by appending a numeric suffix.
      3. Infer the best matching dBASE field type / length / decimal_count for
         each column by inspecting the pandas dtype and the actual data values.
      4. Write the dBASE III binary structure:
             32-byte file header
           + 32-byte field descriptor × n_fields
           + 0x0D  (header terminator byte)
           + [ 0x20 deletion_flag + fixed-width field bytes ] × n_rows
           + 0x1A  (EOF marker)
      5. Return the absolute path of the newly created .dbf file.

    Parameters
    ----------
    excel_path : str
        Full path to the source Excel workbook.
    out_dir : str
        Directory where the output .dbf file will be saved.
    sheet : str or int, optional
        Sheet name (str) or 0-based sheet index (int).  Defaults to 0
        (first sheet).
    encoding : str, optional
        Character encoding used to store string data inside the .dbf file.
        Defaults to 'cp1256' (Windows Arabic).
    """
    # ── 1. Read Excel worksheet ───────────────────────────────────────────────
    df = pd.read_excel(excel_path, sheet_name=sheet)

    # Flatten multi-level column headers (e.g. from merged header rows)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ["_".join(str(c) for c in col).strip() for col in df.columns]

    # ── 2. Build field descriptor list ───────────────────────────────────────
    # Each entry: (sanitized_name, ftype, flen, fdecimals)
    fields = []
    seen_names: dict = {}

    for col in df.columns:
        raw_name = _sanitize_field_name(col)

        # Deduplicate colliding names by appending an incrementing counter
        if raw_name in seen_names:
            seen_names[raw_name] += 1
            suffix = str(seen_names[raw_name])
            raw_name = raw_name[: 10 - len(suffix)] + suffix
        else:
            seen_names[raw_name] = 0

        ftype, flen, fdecimals = _infer_dbf_field(df[col], encoding)
        fields.append((raw_name, ftype, flen, fdecimals))

    # ── 3. Compute binary layout sizes ───────────────────────────────────────
    n_fields    = len(fields)
    record_size = 1 + sum(f[2] for f in fields)    # 1 byte deletion flag  + data
    header_size = 32 + 32 * n_fields + 1            # file header + descriptors + 0x0D

    # ── 4. Write the binary .dbf file ─────────────────────────────────────────
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    out_path  = os.path.join(out_dir, base_name + ".dbf")
    today     = datetime.date.today()

    with open(out_path, "wb") as fh:

        # ── File header (exactly 32 bytes) ────────────────────────────────────
        fh.write(struct.pack("<B",   0x03))                                     # dBASE III version
        fh.write(struct.pack("<BBB", today.year % 100, today.month, today.day)) # last-update date
        fh.write(struct.pack("<I",   len(df)))                                  # total record count
        fh.write(struct.pack("<H",   header_size))                              # bytes in header
        fh.write(struct.pack("<H",   record_size))                              # bytes per record
        fh.write(b"\x00" * 20)                                                  # reserved bytes

        # ── Field descriptors (32 bytes each) ────────────────────────────────
        for fname, ftype, flen, fdecimals in fields:
            # Field name: null-terminated, padded to 11 bytes
            name_bytes = fname.encode("ascii", errors="replace")[:11].ljust(11, b"\x00")
            fh.write(name_bytes)
            fh.write(ftype.encode("ascii"))   # field type character
            fh.write(b"\x00" * 4)             # reserved (field data address in memory)
            fh.write(struct.pack("<B", flen))
            fh.write(struct.pack("<B", fdecimals))
            fh.write(b"\x00" * 14)            # reserved

        fh.write(b"\x0D")                     # header terminator

        # ── Data records ──────────────────────────────────────────────────────
        for _, row in df.iterrows():
            fh.write(b"\x20")                 # deletion flag: 0x20 = not deleted
            for (_, ftype, flen, fdecimals), val in zip(fields, row):
                fh.write(_encode_dbf_value(val, ftype, flen, fdecimals, encoding))

        fh.write(b"\x1A")                     # EOF marker

    return out_path


# ─── GUI ─────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    """
    Main window of the DBF ↔ Excel Converter application.

    Uses a two-tab ttk.Notebook layout:
      • Tab 1  "DBF → CSV"   – batch-convert dBASE files to comma-separated values
      • Tab 2  "Excel → DBF" – batch-convert Excel worksheets to dBASE III files

    All file I/O is executed on background daemon threads so the UI stays
    responsive during long conversions.  Progress and status updates are
    dispatched back to the main thread via Tk's thread-safe .after() mechanism.
    """

    # ── Colour palette (shared across all widgets) ────────────────────────────
    BG     = "#1e1e2e"   # window / outer background
    PANEL  = "#2a2a3e"   # list-panel background
    ACCENT = "#7c6aff"   # purple accent (buttons, selection, progress bar)
    FG     = "#cdd6f4"   # primary foreground text
    MUTED  = "#6c7086"   # secondary / caption text
    ENTRY  = "#12121e"   # text-entry / listbox background

    def __init__(self):
        """Initialise the Tk root window, apply the colour scheme, and build the UI."""
        super().__init__()
        self.title("DBF ↔ Excel Converter")
        self.resizable(True, True)
        self.minsize(720, 540)
        self.configure(bg=self.BG)

        # Per-tab file queues
        self._dbf_files:   list = []
        self._excel_files: list = []

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        """
        Create and arrange every UI widget in three layers:
          1. Header  – application title, subtitle and author credit
          2. Notebook – Tab 1 (DBF→CSV) and Tab 2 (Excel→DBF)
          3. Footer   – shared progress bar, status label, copyright
        All ttk styles are configured here so every tab inherits them.
        """
        PAD = 10

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame",         background=self.BG)
        style.configure("Panel.TFrame",   background=self.PANEL)
        style.configure("TNotebook",      background=self.BG,    borderwidth=0)
        style.configure("TNotebook.Tab",  background=self.PANEL, foreground=self.FG,
                         font=("Segoe UI", 10), padding=(14, 6))
        style.map("TNotebook.Tab",
                  background=[("selected", self.ACCENT)],
                  foreground=[("selected", "white")])
        style.configure("Accent.TButton", background=self.ACCENT, foreground="white",
                         font=("Segoe UI", 10, "bold"), padding=6)
        style.map("Accent.TButton",       background=[("active", "#6455d4")])
        style.configure("Muted.TButton",  background=self.PANEL, foreground=self.FG,
                         font=("Segoe UI", 10), padding=6)
        style.map("Muted.TButton",        background=[("active", "#3a3a55")])
        style.configure("TProgressbar",   troughcolor=self.PANEL,
                         background=self.ACCENT, thickness=8)
        style.configure("TLabel",         background=self.BG, foreground=self.FG,
                         font=("Segoe UI", 10))
        style.configure("Header.TLabel",  background=self.BG, foreground=self.FG,
                         font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel",     background=self.BG, foreground=self.MUTED,
                         font=("Segoe UI", 9))

        # ── Header ────────────────────────────────────────────────────────────
        hdr = ttk.Frame(self, style="TFrame", padding=(PAD * 2, PAD * 2, PAD * 2, PAD))
        hdr.pack(fill="x")
        ttk.Label(hdr, text="DBF ↔ Excel Converter",       style="Header.TLabel").pack(anchor="w")
        ttk.Label(hdr, text="Convert dBASE files to CSV — or Excel sheets to DBF",
                  style="Sub.TLabel").pack(anchor="w")
        ttk.Label(hdr, text="Developer: Ahmed Abdul Ammer Al-Saatci",
                  style="Sub.TLabel").pack(anchor="w")

        # ── Notebook (two tabs) ───────────────────────────────────────────────
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=PAD * 2, pady=(0, PAD))

        tab1 = ttk.Frame(self.notebook, style="TFrame", padding=PAD)
        tab2 = ttk.Frame(self.notebook, style="TFrame", padding=PAD)
        self.notebook.add(tab1, text="  DBF → CSV  ")
        self.notebook.add(tab2, text="  Excel → DBF  ")

        self._build_dbf_tab(tab1, PAD)
        self._build_excel_tab(tab2, PAD)

        # ── Shared progress bar + status label ────────────────────────────────
        prog_frame = ttk.Frame(self, style="TFrame", padding=(PAD * 2, 0, PAD * 2, 0))
        prog_frame.pack(fill="x")
        self.progress = ttk.Progressbar(prog_frame, mode="determinate", style="TProgressbar")
        self.progress.pack(fill="x", pady=(0, 4))
        self.status_var = tk.StringVar(value="Ready — select a tab to get started.")
        ttk.Label(prog_frame, textvariable=self.status_var, style="Sub.TLabel").pack(anchor="w")

        # ── Footer ────────────────────────────────────────────────────────────
        foot = ttk.Frame(self, style="TFrame", padding=(PAD * 2, PAD, PAD * 2, PAD * 2))
        foot.pack(fill="x")
        ttk.Label(foot, text="© Ahmed Abdul Ammer Al-Saatci", style="Sub.TLabel").pack(side="left")

    def _build_dbf_tab(self, parent: ttk.Frame, pad: int):
        """
        Populate the 'DBF → CSV' notebook tab.

        Builds (top to bottom):
          • Labelled panel with a scrollable file listbox and Add/Clear buttons
          • Output-folder selector (entry + Browse button)
          • 'Convert All  DBF → CSV' button (right-aligned)

        Widget handles stored on self:
          self.dbf_listbox  – Listbox showing queued .dbf files
          self.dbf_out_var  – StringVar bound to the output folder Entry
          self.dbf_conv_btn – ttk.Button that triggers conversion
        """
        PANEL, FG, ACCENT = self.PANEL, self.FG, self.ACCENT

        # File list panel ──────────────────────────────────────────────────────
        list_frame = ttk.Frame(parent, style="Panel.TFrame", padding=pad)
        list_frame.pack(fill="both", expand=True, pady=(0, pad))

        list_top = ttk.Frame(list_frame, style="Panel.TFrame")
        list_top.pack(fill="x", pady=(0, 6))
        ttk.Label(list_top, text="Selected DBF Files",
                  background=PANEL, foreground=FG,
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        btn_row = ttk.Frame(list_top, style="Panel.TFrame")
        btn_row.pack(side="right")
        ttk.Button(btn_row, text="+ Add Files", style="Accent.TButton",
                   command=self._dbf_add_files).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Clear", style="Muted.TButton",
                   command=self._dbf_clear_files).pack(side="left")

        lb_frame = ttk.Frame(list_frame, style="Panel.TFrame")
        lb_frame.pack(fill="both", expand=True)
        sb = tk.Scrollbar(lb_frame, orient="vertical", bg=PANEL)
        sb.pack(side="right", fill="y")
        self.dbf_listbox = tk.Listbox(
            lb_frame, yscrollcommand=sb.set,
            bg=self.ENTRY, fg=FG, selectbackground=ACCENT, selectforeground="white",
            font=("Segoe UI", 9), activestyle="none", borderwidth=0, highlightthickness=0,
        )
        self.dbf_listbox.pack(fill="both", expand=True)
        sb.config(command=self.dbf_listbox.yview)
        self.dbf_listbox.bind("<Delete>", self._dbf_remove_selected)

        # Output folder ────────────────────────────────────────────────────────
        out_frame = ttk.Frame(parent, style="TFrame")
        out_frame.pack(fill="x", pady=(0, pad))
        ttk.Label(out_frame, text="Output Folder:").pack(side="left", padx=(0, 6))
        self.dbf_out_var = tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        tk.Entry(
            out_frame, textvariable=self.dbf_out_var,
            bg=self.ENTRY, fg=FG, insertbackground=FG,
            font=("Segoe UI", 9), relief="flat", borderwidth=4,
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(out_frame, text="Browse", style="Muted.TButton",
                   command=self._dbf_browse_out).pack(side="left")

        # Convert button ───────────────────────────────────────────────────────
        self.dbf_conv_btn = ttk.Button(
            parent, text="Convert All  DBF → CSV",
            style="Accent.TButton", command=self._dbf_start_conversion,
        )
        self.dbf_conv_btn.pack(anchor="e")

    def _build_excel_tab(self, parent: ttk.Frame, pad: int):
        """
        Populate the 'Excel → DBF' notebook tab.

        Builds (top to bottom):
          • Labelled panel with a scrollable file listbox and Add/Clear buttons
          • Options row: sheet-name entry + DBF encoding combobox
          • Output-folder selector (entry + Browse button)
          • 'Convert All  Excel → DBF' button (right-aligned)

        Widget handles stored on self:
          self.xl_listbox   – Listbox showing queued Excel files
          self.xl_out_var   – StringVar for the output folder Entry
          self.xl_sheet_var – StringVar for sheet name / index (blank = first sheet)
          self.xl_enc_var   – StringVar for the DBF output encoding
          self.xl_conv_btn  – ttk.Button that triggers conversion
        """
        PANEL, FG, ACCENT = self.PANEL, self.FG, self.ACCENT

        # File list panel ──────────────────────────────────────────────────────
        list_frame = ttk.Frame(parent, style="Panel.TFrame", padding=pad)
        list_frame.pack(fill="both", expand=True, pady=(0, pad))

        list_top = ttk.Frame(list_frame, style="Panel.TFrame")
        list_top.pack(fill="x", pady=(0, 6))
        ttk.Label(list_top, text="Selected Excel Files",
                  background=PANEL, foreground=FG,
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        btn_row = ttk.Frame(list_top, style="Panel.TFrame")
        btn_row.pack(side="right")
        ttk.Button(btn_row, text="+ Add Files", style="Accent.TButton",
                   command=self._xl_add_files).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Clear", style="Muted.TButton",
                   command=self._xl_clear_files).pack(side="left")

        lb_frame = ttk.Frame(list_frame, style="Panel.TFrame")
        lb_frame.pack(fill="both", expand=True)
        sb = tk.Scrollbar(lb_frame, orient="vertical", bg=PANEL)
        sb.pack(side="right", fill="y")
        self.xl_listbox = tk.Listbox(
            lb_frame, yscrollcommand=sb.set,
            bg=self.ENTRY, fg=FG, selectbackground=ACCENT, selectforeground="white",
            font=("Segoe UI", 9), activestyle="none", borderwidth=0, highlightthickness=0,
        )
        self.xl_listbox.pack(fill="both", expand=True)
        sb.config(command=self.xl_listbox.yview)
        self.xl_listbox.bind("<Delete>", self._xl_remove_selected)

        # Options row: sheet selector + encoding picker ────────────────────────
        opt_frame = ttk.Frame(parent, style="TFrame")
        opt_frame.pack(fill="x", pady=(0, pad))

        ttk.Label(opt_frame, text="Sheet (name or blank for first):").pack(side="left", padx=(0, 4))
        self.xl_sheet_var = tk.StringVar(value="")
        tk.Entry(
            opt_frame, textvariable=self.xl_sheet_var, width=16,
            bg=self.ENTRY, fg=FG, insertbackground=FG,
            font=("Segoe UI", 9), relief="flat", borderwidth=4,
        ).pack(side="left", padx=(0, 12))

        ttk.Label(opt_frame, text="DBF Encoding:").pack(side="left", padx=(0, 4))
        self.xl_enc_var = tk.StringVar(value="cp1256")
        ttk.Combobox(
            opt_frame, textvariable=self.xl_enc_var, state="readonly", width=10,
            values=["cp1256", "cp720", "cp864", "cp1252", "utf-8"],
        ).pack(side="left")

        # Output folder ────────────────────────────────────────────────────────
        out_frame = ttk.Frame(parent, style="TFrame")
        out_frame.pack(fill="x", pady=(0, pad))
        ttk.Label(out_frame, text="Output Folder:").pack(side="left", padx=(0, 6))
        self.xl_out_var = tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        tk.Entry(
            out_frame, textvariable=self.xl_out_var,
            bg=self.ENTRY, fg=FG, insertbackground=FG,
            font=("Segoe UI", 9), relief="flat", borderwidth=4,
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(out_frame, text="Browse", style="Muted.TButton",
                   command=self._xl_browse_out).pack(side="left")

        # Convert button ───────────────────────────────────────────────────────
        self.xl_conv_btn = ttk.Button(
            parent, text="Convert All  Excel → DBF",
            style="Accent.TButton", command=self._xl_start_conversion,
        )
        self.xl_conv_btn.pack(anchor="e")

    # ── DBF → CSV event handlers ──────────────────────────────────────────────

    def _dbf_add_files(self):
        """
        Open a native file-picker dialog filtered to .dbf files and add the
        chosen paths to the DBF conversion queue (duplicates are silently skipped).
        """
        paths = filedialog.askopenfilenames(
            title="Select DBF files",
            filetypes=[("dBASE files", "*.dbf *.DBF"), ("All files", "*.*")],
        )
        added = 0
        for p in paths:
            if p not in self._dbf_files:
                self._dbf_files.append(p)
                self.dbf_listbox.insert("end", os.path.basename(p) + "   ←   " + p)
                added += 1
        if added:
            self.status_var.set(f"{len(self._dbf_files)} DBF file(s) queued.")

    def _dbf_clear_files(self):
        """Clear the entire DBF file queue and reset the progress bar."""
        self._dbf_files.clear()
        self.dbf_listbox.delete(0, "end")
        self.progress["value"] = 0
        self.status_var.set("DBF file list cleared.")

    def _dbf_remove_selected(self, _event=None):
        """Remove only the currently highlighted rows from the DBF queue (Delete key)."""
        for idx in reversed(self.dbf_listbox.curselection()):
            self._dbf_files.pop(idx)
            self.dbf_listbox.delete(idx)
        self.status_var.set(f"{len(self._dbf_files)} DBF file(s) queued.")

    def _dbf_browse_out(self):
        """Open a folder-picker dialog and update the DBF output-folder entry."""
        d = filedialog.askdirectory(title="Select output folder",
                                    initialdir=self.dbf_out_var.get())
        if d:
            self.dbf_out_var.set(d)

    def _dbf_start_conversion(self):
        """
        Validate the DBF queue and output folder, then spawn a background thread
        that calls convert_dbf_to_csv() for every queued file.
        """
        if not self._dbf_files:
            messagebox.showwarning("No files", "Please add at least one DBF file.")
            return
        out_dir = self.dbf_out_var.get().strip()
        if not out_dir:
            messagebox.showwarning("No output folder", "Please select an output folder.")
            return
        os.makedirs(out_dir, exist_ok=True)
        self.dbf_conv_btn.state(["disabled"])
        threading.Thread(
            target=self._dbf_run_conversion,
            args=(list(self._dbf_files), out_dir),
            daemon=True,
        ).start()

    def _dbf_run_conversion(self, files: list, out_dir: str):
        """
        Worker thread: convert every DBF file in the queue to CSV.

        Updates the shared progress bar and status label after each file via
        the thread-safe _set_status / _set_progress helpers.  Errors are
        collected and passed to _on_dbf_done for display in a summary dialog.
        """
        total = len(files)
        errors: list = []
        converted = 0
        self.progress["maximum"] = total
        self.progress["value"]   = 0

        for i, path in enumerate(files, 1):
            name = os.path.basename(path)
            self._set_status(f"Converting {i}/{total}: {name} …")
            try:
                out = convert_dbf_to_csv(path, out_dir)
                converted += 1
                self._set_status(f"Done {i}/{total}: {name} → {os.path.basename(out)}")
            except Exception as exc:
                errors.append(f"{name}: {exc}")
                self._set_status(f"Error {i}/{total}: {name} — {exc}")
            self._set_progress(i)

        self.after(0, self._on_dbf_done, converted, errors, out_dir)

    def _on_dbf_done(self, converted: int, errors: list, out_dir: str):
        """
        Called on the main thread after DBF→CSV conversion finishes.
        Re-enables the convert button and shows a success or warning dialog.
        """
        self.dbf_conv_btn.state(["!disabled"])
        if errors:
            messagebox.showwarning(
                "Completed with errors",
                f"Converted {converted} file(s).\n\nErrors:\n" + "\n".join(errors),
            )
            self.status_var.set(f"Done: {converted} converted, {len(errors)} error(s).")
        else:
            messagebox.showinfo(
                "Success",
                f"All {converted} file(s) converted successfully!\n\nSaved to:\n{out_dir}",
            )
            self.status_var.set(f"All {converted} file(s) converted successfully.")
        self.progress["value"] = self.progress["maximum"]

    # ── Excel → DBF event handlers ────────────────────────────────────────────

    def _xl_add_files(self):
        """
        Open a native file-picker dialog filtered to Excel formats and add the
        chosen paths to the Excel conversion queue (duplicates are silently skipped).
        """
        paths = filedialog.askopenfilenames(
            title="Select Excel files",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm"),
                ("All files",   "*.*"),
            ],
        )
        added = 0
        for p in paths:
            if p not in self._excel_files:
                self._excel_files.append(p)
                self.xl_listbox.insert("end", os.path.basename(p) + "   ←   " + p)
                added += 1
        if added:
            self.status_var.set(f"{len(self._excel_files)} Excel file(s) queued.")

    def _xl_clear_files(self):
        """Clear the entire Excel file queue and reset the progress bar."""
        self._excel_files.clear()
        self.xl_listbox.delete(0, "end")
        self.progress["value"] = 0
        self.status_var.set("Excel file list cleared.")

    def _xl_remove_selected(self, _event=None):
        """Remove only the currently highlighted rows from the Excel queue (Delete key)."""
        for idx in reversed(self.xl_listbox.curselection()):
            self._excel_files.pop(idx)
            self.xl_listbox.delete(idx)
        self.status_var.set(f"{len(self._excel_files)} Excel file(s) queued.")

    def _xl_browse_out(self):
        """Open a folder-picker dialog and update the Excel-to-DBF output-folder entry."""
        d = filedialog.askdirectory(title="Select output folder",
                                    initialdir=self.xl_out_var.get())
        if d:
            self.xl_out_var.set(d)

    def _xl_start_conversion(self):
        """
        Validate the Excel queue and options, then spawn a background thread
        that calls excel_to_dbf() for every queued file.

        Sheet parsing logic:
          • Blank entry  → sheet index 0 (first sheet)
          • Digit string → converted to int (0-based index)
          • Other string → used as the literal sheet name
        """
        if not self._excel_files:
            messagebox.showwarning("No files", "Please add at least one Excel file.")
            return
        out_dir = self.xl_out_var.get().strip()
        if not out_dir:
            messagebox.showwarning("No output folder", "Please select an output folder.")
            return

        sheet_raw = self.xl_sheet_var.get().strip()
        if not sheet_raw:
            sheet: object = 0
        elif sheet_raw.isdigit():
            sheet = int(sheet_raw)
        else:
            sheet = sheet_raw

        encoding = self.xl_enc_var.get()
        os.makedirs(out_dir, exist_ok=True)
        self.xl_conv_btn.state(["disabled"])
        threading.Thread(
            target=self._xl_run_conversion,
            args=(list(self._excel_files), out_dir, sheet, encoding),
            daemon=True,
        ).start()

    def _xl_run_conversion(self, files: list, out_dir: str, sheet, encoding: str):
        """
        Worker thread: convert every Excel file in the queue to DBF.

        Updates the shared progress bar and status label after each file.
        Errors are collected and passed to _on_xl_done for summary display.
        """
        total = len(files)
        errors: list = []
        converted = 0
        self.progress["maximum"] = total
        self.progress["value"]   = 0

        for i, path in enumerate(files, 1):
            name = os.path.basename(path)
            self._set_status(f"Converting {i}/{total}: {name} …")
            try:
                out = excel_to_dbf(path, out_dir, sheet=sheet, encoding=encoding)
                converted += 1
                self._set_status(f"Done {i}/{total}: {name} → {os.path.basename(out)}")
            except Exception as exc:
                errors.append(f"{name}: {exc}")
                self._set_status(f"Error {i}/{total}: {name} — {exc}")
            self._set_progress(i)

        self.after(0, self._on_xl_done, converted, errors, out_dir)

    def _on_xl_done(self, converted: int, errors: list, out_dir: str):
        """
        Called on the main thread after Excel→DBF conversion finishes.
        Re-enables the convert button and shows a success or warning dialog.
        """
        self.xl_conv_btn.state(["!disabled"])
        if errors:
            messagebox.showwarning(
                "Completed with errors",
                f"Converted {converted} file(s).\n\nErrors:\n" + "\n".join(errors),
            )
            self.status_var.set(f"Done: {converted} converted, {len(errors)} error(s).")
        else:
            messagebox.showinfo(
                "Success",
                f"All {converted} file(s) converted successfully!\n\nSaved to:\n{out_dir}",
            )
            self.status_var.set(f"All {converted} file(s) converted successfully.")
        self.progress["value"] = self.progress["maximum"]

    # ── Shared thread-safe helpers ────────────────────────────────────────────

    def _set_status(self, msg: str):
        """
        Thread-safe wrapper: schedule a status-label text update on the main
        Tk event loop using .after(0, ...) so it is safe to call from any thread.
        """
        self.after(0, self.status_var.set, msg)

    def _set_progress(self, value: int):
        """
        Thread-safe wrapper: schedule a progress-bar value update on the main
        Tk event loop using .after(0, ...) so it is safe to call from any thread.
        """
        self.after(0, self.progress.__setitem__, "value", value)


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
