#!/usr/bin/env python3
"""SheetFix GUI — pick one or more workbooks, click Run, get the fixed
files + re-check reports (plus a batch_summary.md when several are queued).
Everything runs locally; no data leaves the machine."""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sheet_fix import run  # noqa: E402
from sheet_batch import run_batch  # noqa: E402


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SheetFix — Excel Audit, Fix & Re-Check")
        self.geometry("640x360")
        self.resizable(False, False)

        self.file_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.status = tk.StringVar(value="Choose one or more .xlsx / .xlsm workbooks to begin (Ctrl/Shift-click for several).")
        self.summary = None
        self.files = []
        self.report_path = None

        pad = {"padx": 12, "pady": 6}
        ttk.Label(self, text="Workbook(s):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.file_var, width=52).grid(row=0, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self.pick_file).grid(row=0, column=2, **pad)

        ttk.Label(self, text="Output folder:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.out_var, width=52).grid(row=1, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self.pick_out).grid(row=1, column=2, **pad)

        self.run_btn = ttk.Button(self, text="Audit, Fix & Re-Check", command=self.start)
        self.run_btn.grid(row=2, column=1, pady=10)

        ttk.Label(self, textvariable=self.status, wraplength=600, justify="left").grid(
            row=3, column=0, columnspan=3, sticky="w", **pad)

        btns = ttk.Frame(self)
        btns.grid(row=4, column=0, columnspan=3, pady=8)
        self.report_btn = ttk.Button(btns, text="Open report", command=self.open_report, state="disabled")
        self.report_btn.pack(side="left", padx=6)
        self.folder_btn = ttk.Button(btns, text="Open output folder", command=self.open_folder, state="disabled")
        self.folder_btn.pack(side="left", padx=6)

        ttk.Label(self, text="Local only — your workbook never leaves this computer. "
                             "Mechanical fixes are applied and re-verified; judgment calls are reported, never silently changed.",
                  wraplength=600, foreground="#666").grid(row=5, column=0, columnspan=3, sticky="w", **pad)

    def pick_file(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if paths:
            self.files = list(paths)
            if len(self.files) == 1:
                self.file_var.set(self.files[0])
            else:
                self.file_var.set(f"{len(self.files)} workbooks selected")
            if not self.out_var.get():
                self.out_var.set(os.path.join(os.path.dirname(self.files[0]), "SheetFix_output"))

    def pick_out(self):
        p = filedialog.askdirectory()
        if p:
            self.out_var.set(p)

    def start(self):
        files = list(self.files)
        if not files and self.file_var.get().strip() and os.path.isfile(self.file_var.get().strip()):
            files = [self.file_var.get().strip()]  # path typed/pasted by hand
        out = self.out_var.get().strip()
        if not files:
            messagebox.showerror("SheetFix", "Please choose at least one workbook file first.")
            return
        if not out:
            messagebox.showerror("SheetFix", "Please choose an output folder.")
            return
        self.run_btn.config(state="disabled")
        self.report_btn.config(state="disabled")
        self.folder_btn.config(state="disabled")
        label = "workbook" if len(files) == 1 else f"{len(files)} workbooks"
        self.status.set(f"Running audit → fix → re-check on {label}… this can take a few seconds.")
        threading.Thread(target=self._work, args=(files, out), daemon=True).start()

    def _work(self, files, out):
        try:
            if len(files) == 1:
                self.summary = run(files[0], out)
                s = self.summary
                self.report_path = s["report_path"]
                text = (f"Done. Findings before: {len(s['before'])} · auto-fixed: {len(s['fixes'])} · "
                        f"resolved on re-check: {len(s['resolved'])} · remaining (need your judgment): {len(s['remaining'])} · "
                        f"new issues introduced: {len(s['new_issues'])}.\nFixed file: {s['fixed_path']}")
            else:
                batch = run_batch(files, out)
                self.summary = batch
                self.report_path = batch["summary_path"]
                t = batch["totals"]
                text = (f"Done — batch of {len(batch['entries'])} workbooks ({batch['failed']} failed). "
                        f"Findings before: {t['before']} · auto-fixed: {t['fixes']} · "
                        f"resolved on re-check: {t['resolved']} · remaining (need your judgment): {t['remaining']} · "
                        f"new issues introduced: {t['new_issues']}.\nBatch summary: {batch['summary_path']}")
        except Exception as exc:  # surface the real error, don't swallow it
            text = f"Something went wrong: {exc}"
            self.summary = None
            self.report_path = None
        self.after(0, self._done, text)

    def _done(self, text):
        self.status.set(text)
        self.run_btn.config(state="normal")
        if self.summary:
            self.report_btn.config(state="normal")
            self.folder_btn.config(state="normal")

    def _open(self, path):
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606 - opening a local file the user generated
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)

    def open_report(self):
        if self.report_path:
            self._open(self.report_path)

    def open_folder(self):
        if self.report_path:
            self._open(os.path.dirname(self.report_path))


if __name__ == "__main__":
    App().mainloop()
