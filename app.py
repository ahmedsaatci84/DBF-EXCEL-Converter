import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from dbfread import DBF


# ─── Conversion logic ────────────────────────────────────────────────────────

# Arabic encoding priority:
#   cp1256 = Windows Arabic (most common in modern Arabic DBF files)
#   cp720  = DOS Arabic (legacy)
#   cp864  = DOS Arabic alternative (used by some older systems)
#   cp1252 = Western European (last resort for mixed files)
ENCODINGS_TO_TRY = ["cp720"]


def read_dbf(path: str) -> pd.DataFrame:
    """Read a DBF file and return a DataFrame, trying multiple encodings."""
    last_exc = None
    for enc in ENCODINGS_TO_TRY:
        try:
            table = DBF(path, encoding=enc, load=True, ignore_missing_memofile=True)
            df = pd.DataFrame(iter(table))
            return df
        except Exception as exc:
            last_exc = exc
    raise RuntimeError(f"Could not read '{os.path.basename(path)}': {last_exc}")


def convert_file(dbf_path: str, out_dir: str) -> str:
    """Convert a single DBF file to a CSV file."""
    df = read_dbf(dbf_path)

    # Decode bytes columns to proper Unicode strings
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(_decode_bytes)

    base_name = os.path.splitext(os.path.basename(dbf_path))[0]
    out_path = os.path.join(out_dir, base_name + ".csv")

    # cp1256 = Windows Arabic; matches what Arabic-locale Excel expects
    # when opening a CSV without a BOM.
    df.to_csv(out_path, index=False, encoding="cp1256", errors="replace")

    return out_path


def _decode_bytes(v):
    """Decode raw bytes from a DBF field, trying Arabic encodings first."""
    if not isinstance(v, (bytes, bytearray)):
        return v
    for enc in ("utf-8", "cp1256", "cp720", "cp864", "cp1252"):
        try:
            return v.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return v.decode("cp1256", errors="replace")


# ─── GUI ─────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DBF → CSV Converter")
        self.resizable(True, True)
        self.minsize(700, 480)
        self.configure(bg="#1e1e2e")
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        PAD = 10
        BG = "#1e1e2e"
        PANEL = "#2a2a3e"
        ACCENT = "#7c6aff"
        FG = "#cdd6f4"
        MUTED = "#6c7086"

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("Panel.TFrame", background=PANEL)
        style.configure(
            "Accent.TButton",
            background=ACCENT, foreground="white",
            font=("Segoe UI", 10, "bold"), padding=6,
        )
        style.map("Accent.TButton", background=[("active", "#6455d4")])
        style.configure(
            "Muted.TButton",
            background=PANEL, foreground=FG,
            font=("Segoe UI", 10), padding=6,
        )
        style.map("Muted.TButton", background=[("active", "#3a3a55")])
        style.configure("TProgressbar", troughcolor=PANEL, background=ACCENT, thickness=8)
        style.configure("TLabel", background=BG, foreground=FG, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=BG, foreground=FG, font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))

        # ── Header ────────────────────────────────────────────────────────────
        hdr = ttk.Frame(self, style="TFrame", padding=(PAD * 2, PAD * 2, PAD * 2, PAD))
        hdr.pack(fill="x")
        ttk.Label(hdr, text="DBF → CSV Converter", style="Header.TLabel").pack(anchor="w")
        ttk.Label(hdr, text="Batch-convert dBASE files to .csv without losing data", style="Sub.TLabel").pack(anchor="w")
        ttk.Label(hdr, text="Developer: Ahmed Abdul Ammer Al-Saatci", style="Sub.TLabel").pack(anchor="w")

        # ── File list panel ───────────────────────────────────────────────────
        list_frame = ttk.Frame(self, style="Panel.TFrame", padding=PAD)
        list_frame.pack(fill="both", expand=True, padx=PAD * 2, pady=(0, PAD))

        list_top = ttk.Frame(list_frame, style="Panel.TFrame")
        list_top.pack(fill="x", pady=(0, 6))
        ttk.Label(list_top, text="Selected DBF Files", background=PANEL, foreground=FG,
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        btn_row = ttk.Frame(list_top, style="Panel.TFrame")
        btn_row.pack(side="right")
        ttk.Button(btn_row, text="+ Add Files", style="Accent.TButton",
                   command=self._add_files).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Clear", style="Muted.TButton",
                   command=self._clear_files).pack(side="left")

        # Listbox with scrollbar
        lb_frame = ttk.Frame(list_frame, style="Panel.TFrame")
        lb_frame.pack(fill="both", expand=True)

        sb = tk.Scrollbar(lb_frame, orient="vertical", bg=PANEL)
        sb.pack(side="right", fill="y")
        self.listbox = tk.Listbox(
            lb_frame, yscrollcommand=sb.set,
            bg="#12121e", fg=FG, selectbackground=ACCENT, selectforeground="white",
            font=("Segoe UI", 9), activestyle="none", borderwidth=0, highlightthickness=0,
        )
        self.listbox.pack(fill="both", expand=True)
        sb.config(command=self.listbox.yview)
        self.listbox.bind("<Delete>", self._remove_selected)

        # ── Output directory ──────────────────────────────────────────────────
        out_frame = ttk.Frame(self, style="TFrame", padding=(PAD * 2, 0, PAD * 2, PAD))
        out_frame.pack(fill="x")
        ttk.Label(out_frame, text="Output Folder:").pack(side="left", padx=(0, 6))
        self.out_var = tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        out_entry = tk.Entry(
            out_frame, textvariable=self.out_var, width=52,
            bg="#12121e", fg=FG, insertbackground=FG,
            font=("Segoe UI", 9), relief="flat", borderwidth=4,
        )
        out_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(out_frame, text="Browse", style="Muted.TButton",
                   command=self._browse_out).pack(side="left")

        # ── Progress ──────────────────────────────────────────────────────────
        prog_frame = ttk.Frame(self, style="TFrame", padding=(PAD * 2, 0, PAD * 2, 0))
        prog_frame.pack(fill="x")
        self.progress = ttk.Progressbar(prog_frame, mode="determinate", style="TProgressbar")
        self.progress.pack(fill="x", pady=(0, 4))
        self.status_var = tk.StringVar(value="Ready — add DBF files to get started.")
        ttk.Label(prog_frame, textvariable=self.status_var, style="Sub.TLabel").pack(anchor="w")

        # ── Convert button ────────────────────────────────────────────────────
        btn_frame = ttk.Frame(self, style="TFrame", padding=(PAD * 2, PAD, PAD * 2, PAD * 2))
        btn_frame.pack(fill="x")
        self.convert_btn = ttk.Button(
            btn_frame, text="Convert All to CSV",
            style="Accent.TButton", command=self._start_conversion,
        )
        self.convert_btn.pack(side="right")
        ttk.Label(btn_frame, text="© Ahmed Abdul Ammer Al-Saatci", style="Sub.TLabel").pack(side="left")

        self._files: list[str] = []

    # ── Event handlers ────────────────────────────────────────────────────────

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select DBF files",
            filetypes=[("dBASE files", "*.dbf *.DBF"), ("All files", "*.*")],
        )
        added = 0
        for p in paths:
            if p not in self._files:
                self._files.append(p)
                self.listbox.insert("end", os.path.basename(p) + "   ←   " + p)
                added += 1
        if added:
            self.status_var.set(f"{len(self._files)} file(s) queued.")

    def _clear_files(self):
        self._files.clear()
        self.listbox.delete(0, "end")
        self.progress["value"] = 0
        self.status_var.set("File list cleared.")

    def _remove_selected(self, _event=None):
        for idx in reversed(self.listbox.curselection()):
            self._files.pop(idx)
            self.listbox.delete(idx)
        self.status_var.set(f"{len(self._files)} file(s) queued.")

    def _browse_out(self):
        d = filedialog.askdirectory(title="Select output folder", initialdir=self.out_var.get())
        if d:
            self.out_var.set(d)

    # ── Conversion ────────────────────────────────────────────────────────────

    def _start_conversion(self):
        if not self._files:
            messagebox.showwarning("No files", "Please add at least one DBF file.")
            return
        out_dir = self.out_var.get().strip()
        if not out_dir:
            messagebox.showwarning("No output folder", "Please select an output folder.")
            return
        os.makedirs(out_dir, exist_ok=True)
        self.convert_btn.state(["disabled"])
        threading.Thread(target=self._run_conversion, args=(list(self._files), out_dir), daemon=True).start()

    def _run_conversion(self, files: list[str], out_dir: str):
        total = len(files)
        errors: list[str] = []
        converted = 0

        self.progress["maximum"] = total
        self.progress["value"] = 0

        for i, path in enumerate(files, 1):
            name = os.path.basename(path)
            self._set_status(f"Converting {i}/{total}: {name} …")
            try:
                out = convert_file(path, out_dir)
                converted += 1
                self._set_status(f"Done {i}/{total}: {name} → {os.path.basename(out)}")
            except Exception as exc:
                errors.append(f"{name}: {exc}")
                self._set_status(f"Error {i}/{total}: {name} — {exc}")
            self._set_progress(i)

        # Summary
        self.after(0, self._on_done, converted, errors, out_dir)

    def _on_done(self, converted: int, errors: list[str], out_dir: str):
        self.convert_btn.state(["!disabled"])
        if errors:
            detail = "\n".join(errors)
            messagebox.showwarning(
                "Completed with errors",
                f"Converted {converted} file(s).\n\nErrors:\n{detail}",
            )
            self.status_var.set(f"Done: {converted} converted, {len(errors)} error(s). See warning dialog.")
        else:
            messagebox.showinfo(
                "Success",
                f"All {converted} file(s) converted successfully!\n\nSaved to:\n{out_dir}",
            )
            self.status_var.set(f"All {converted} file(s) converted successfully.")
        self.progress["value"] = self.progress["maximum"]

    def _set_status(self, msg: str):
        self.after(0, self.status_var.set, msg)

    def _set_progress(self, value: int):
        self.after(0, self.progress.__setitem__, "value", value)


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
