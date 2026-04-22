"""
SBA File Renamer — Desktop (Tkinter)
Packages to .exe via:  pyinstaller --onefile --windowed desktop_app.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading

# Allow running from repo root or from the desktop/ folder
import sys, os
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.engine import (
    CSEC_SUBJECTS, CSEC_SUBJECT_NAMES, FILE_TYPES, DOC_NUMBERS,
    extract_title_page_text, extract_candidate_info,
    load_master_list, process_job, RenameJob,
    validate_candidate_number, build_filename,
)

ACCENT   = "#1a5276"
BG       = "#f4f6f7"
WHITE    = "#ffffff"
SUCCESS  = "#1e8449"
ERROR    = "#922b21"
DISABLED = "#abb2b9"


class SBARenamerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SBA File Renamer — CUC Secondary")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(640, 560)

        # ── State ──
        self.source_files: list[Path] = []
        self.output_dir: Path | None = None
        self.master_list: dict[str, str] = {}
        self.current_index = 0

        self._build_ui()
        self._update_preview()

    # ── UI Construction ─────────────────────────────────────────────────────

    def _build_ui(self):
        pad = dict(padx=12, pady=6)

        # Title bar
        header = tk.Frame(self, bg=ACCENT)
        header.pack(fill="x")
        tk.Label(header, text="SBA File Renamer", bg=ACCENT, fg=WHITE,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=16, pady=10)
        tk.Label(header, text="Caribbean Union College Secondary",
                 bg=ACCENT, fg="#aed6f1", font=("Segoe UI", 9)).pack(side="right", padx=16)

        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)

        # ── Row 0: File selection ──
        self._section(main, "1. Select SBA File(s)")
        fr0 = tk.Frame(main, bg=BG)
        fr0.pack(fill="x", **pad)
        self.file_label = tk.Label(fr0, text="No files selected", bg=BG,
                                   fg=DISABLED, font=("Segoe UI", 9), anchor="w")
        self.file_label.pack(side="left", fill="x", expand=True)
        tk.Button(fr0, text="Browse…", command=self._browse_files,
                  bg=ACCENT, fg=WHITE, relief="flat", padx=10).pack(side="right")

        # ── Row 1: Output folder ──
        self._section(main, "2. Output Folder")
        fr1 = tk.Frame(main, bg=BG)
        fr1.pack(fill="x", **pad)
        self.out_label = tk.Label(fr1, text="Same folder as source files (default)",
                                  bg=BG, fg=DISABLED, font=("Segoe UI", 9), anchor="w")
        self.out_label.pack(side="left", fill="x", expand=True)
        tk.Button(fr1, text="Browse…", command=self._browse_output,
                  bg=ACCENT, fg=WHITE, relief="flat", padx=10).pack(side="right")

        # ── Row 2: Master list (optional) ──
        self._section(main, "3. Master List (optional — for candidate number lookup)")
        fr2 = tk.Frame(main, bg=BG)
        fr2.pack(fill="x", **pad)
        self.ml_label = tk.Label(fr2, text="No master list loaded",
                                 bg=BG, fg=DISABLED, font=("Segoe UI", 9), anchor="w")
        self.ml_label.pack(side="left", fill="x", expand=True)
        tk.Button(fr2, text="Load…", command=self._load_master_list,
                  bg="#5d6d7e", fg=WHITE, relief="flat", padx=10).pack(side="right")

        # ── Row 3: Candidate info ──
        self._section(main, "4. Candidate Details")
        grid = tk.Frame(main, bg=BG)
        grid.pack(fill="x", **pad)

        tk.Label(grid, text="Candidate Number (10 digits):", bg=BG,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=3)
        self.cand_var = tk.StringVar()
        self.cand_var.trace_add("write", lambda *_: self._update_preview())
        self.cand_entry = tk.Entry(grid, textvariable=self.cand_var, width=16,
                                   font=("Consolas", 10))
        self.cand_entry.grid(row=0, column=1, sticky="w", padx=8)
        self.cand_name_lbl = tk.Label(grid, text="", bg=BG, fg="#5d6d7e",
                                      font=("Segoe UI", 9, "italic"))
        self.cand_name_lbl.grid(row=0, column=2, sticky="w")

        tk.Label(grid, text="Subject:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=3)
        self.subj_var = tk.StringVar()
        self.subj_combo = ttk.Combobox(grid, textvariable=self.subj_var,
                                       values=CSEC_SUBJECT_NAMES, width=36, state="readonly")
        self.subj_combo.grid(row=1, column=1, columnspan=2, sticky="w", padx=8)
        self.subj_combo.bind("<<ComboboxSelected>>", lambda _: self._update_preview())

        tk.Label(grid, text="Moderation Code:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=3)
        self.mod_var = tk.StringVar()
        mod_entry = tk.Entry(grid, textvariable=self.mod_var, width=12,
                             font=("Consolas", 10), state="readonly")
        mod_entry.grid(row=2, column=1, sticky="w", padx=8)

        tk.Label(grid, text="File Type:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=3, column=0, sticky="w", pady=3)
        self.ftype_var = tk.StringVar(value="SBA")
        ftype_frame = tk.Frame(grid, bg=BG)
        ftype_frame.grid(row=3, column=1, columnspan=2, sticky="w", padx=8)
        for ft in FILE_TYPES:
            tk.Radiobutton(ftype_frame, text=ft, variable=self.ftype_var, value=ft,
                           bg=BG, font=("Segoe UI", 9),
                           command=self._update_preview).pack(side="left", padx=6)

        tk.Label(grid, text="Document #:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w", pady=3)
        self.docnum_var = tk.IntVar(value=1)
        dn_frame = tk.Frame(grid, bg=BG)
        dn_frame.grid(row=4, column=1, columnspan=2, sticky="w", padx=8)
        for n in DOC_NUMBERS:
            tk.Radiobutton(dn_frame, text=str(n), variable=self.docnum_var, value=n,
                           bg=BG, font=("Segoe UI", 9),
                           command=self._update_preview).pack(side="left", padx=6)

        # Watch subject change to auto-fill mod code
        self.subj_var.trace_add("write", self._on_subject_change)

        # ── Preview ──
        self._section(main, "Preview")
        self.preview_var = tk.StringVar(value="—")
        tk.Label(main, textvariable=self.preview_var, bg="#eaf2ff",
                 font=("Consolas", 10), anchor="w", relief="flat",
                 padx=10, pady=6).pack(fill="x", padx=12)

        # ── Progress / status ──
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(main, textvariable=self.status_var, bg=BG, fg="#5d6d7e",
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", padx=12, pady=(4, 0))

        # ── Action buttons ──
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(fill="x", padx=16, pady=(0, 14))

        self.scan_btn = tk.Button(btn_frame, text="⟳ Scan PDF for Candidate Number",
                                  command=self._scan_pdf,
                                  bg="#5d6d7e", fg=WHITE, relief="flat", padx=12, pady=6,
                                  font=("Segoe UI", 9))
        self.scan_btn.pack(side="left", padx=(0, 8))

        self.rename_btn = tk.Button(btn_frame, text="Rename & Copy →",
                                    command=self._rename,
                                    bg=SUCCESS, fg=WHITE, relief="flat", padx=16, pady=6,
                                    font=("Segoe UI", 10, "bold"))
        self.rename_btn.pack(side="right")

    def _section(self, parent, text):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(f, text=text.upper(), bg=BG, fg=ACCENT,
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        tk.Frame(f, bg="#d5d8dc", height=1).pack(side="left", fill="x",
                                                   expand=True, padx=(8, 0))

    # ── Event handlers ───────────────────────────────────────────────────────

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select SBA file(s)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not paths:
            return
        self.source_files = [Path(p) for p in paths]
        count = len(self.source_files)
        label = self.source_files[0].name if count == 1 else f"{count} files selected"
        self.file_label.config(text=label, fg="#1a252f")
        self.current_index = 0
        self._update_preview()

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.output_dir = Path(d)
            self.out_label.config(text=str(self.output_dir), fg="#1a252f")

    def _load_master_list(self):
        path = filedialog.askopenfilename(
            title="Select master candidate list",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            self.master_list = load_master_list(path)
            self.ml_label.config(
                text=f"{Path(path).name} — {len(self.master_list)} candidates",
                fg=SUCCESS,
            )
            self.status_var.set(f"Master list loaded: {len(self.master_list)} entries.")
        except Exception as e:
            messagebox.showerror("Error loading master list", str(e))

    def _on_subject_change(self, *_):
        subj = self.subj_var.get()
        code = CSEC_SUBJECTS.get(subj, "")
        self.mod_var.set(code)
        self._update_preview()

    def _update_preview(self, *_):
        cand = self.cand_var.get().strip()
        mod  = self.mod_var.get().strip()
        ft   = self.ftype_var.get()
        dn   = self.docnum_var.get()

        if len(cand) == 10 and cand.isdigit() and len(mod) == 8 and mod.isdigit():
            name = build_filename(cand, mod, ft, dn, ".pdf")
            self.preview_var.set(name)
        else:
            self.preview_var.set("— fill in candidate number and subject —")

        # Lookup name in master list
        if cand in self.master_list:
            self.cand_name_lbl.config(text=f"✓ {self.master_list[cand]}", fg=SUCCESS)
        elif len(cand) == 10 and self.master_list:
            self.cand_name_lbl.config(text="✗ not in master list", fg=ERROR)
        else:
            self.cand_name_lbl.config(text="")

    def _scan_pdf(self):
        if not self.source_files:
            messagebox.showwarning("No file", "Please select a PDF file first.")
            return

        self.status_var.set("Scanning PDF for candidate number…")
        self.scan_btn.config(state="disabled")

        def run():
            path = self.source_files[self.current_index]
            text = extract_title_page_text(path)
            info = extract_candidate_info(text)
            self.after(0, lambda: self._on_scan_done(info))

        threading.Thread(target=run, daemon=True).start()

    def _on_scan_done(self, info: dict):
        self.scan_btn.config(state="normal")
        num = info.get("number", "")
        name = info.get("name", "")
        if num:
            self.cand_var.set(num)
            self.status_var.set(
                f"Found candidate number: {num}" + (f"  ({name})" if name else "")
            )
        else:
            self.status_var.set("Could not detect candidate number. Please enter manually.")

    def _rename(self):
        if not self.source_files:
            messagebox.showwarning("No file", "Please select at least one file.")
            return

        cand = self.cand_var.get().strip()
        mod  = self.mod_var.get().strip()
        ft   = self.ftype_var.get()
        dn   = self.docnum_var.get()

        ok, err = validate_candidate_number(cand)
        if not ok:
            messagebox.showerror("Invalid input", err)
            return
        if not mod:
            messagebox.showerror("Invalid input", "Please select a subject.")
            return

        results = []
        for i, src in enumerate(self.source_files):
            # For multiple files, auto-increment doc number
            doc_num = dn + i if ft in ("SBA", "Mark Scheme") else dn
            job = RenameJob(
                source_path=src,
                candidate_number=cand,
                moderation_code=mod,
                file_type=ft,
                doc_number=doc_num,
                output_dir=self.output_dir,
            )
            results.append(process_job(job))

        ok_count = sum(1 for r in results if r.ok)
        fail_count = len(results) - ok_count

        msg_lines = [f"✓ {r.new_name}" for r in results if r.ok]
        if fail_count:
            msg_lines += [f"✗ {r.source.name}: {r.error}" for r in results if not r.ok]

        title = "Done" if not fail_count else f"Done with {fail_count} error(s)"
        msg = f"{ok_count} file(s) renamed.\n\n" + "\n".join(msg_lines)
        messagebox.showinfo(title, msg)
        self.status_var.set(f"Renamed {ok_count} file(s). Output: {self.output_dir or 'same folder'}")


if __name__ == "__main__":
    app = SBARenamerApp()
    app.mainloop()
