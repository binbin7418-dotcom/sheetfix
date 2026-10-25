#!/usr/bin/env python3
"""SheetFix GUI — pick a workbook, click Run, get the fixed file + re-check report.
Everything runs locally; no data leaves the machine."""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sheet_fix import run  # noqa: E402


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SheetFix — Excel Audit, Fix & Re-Check")
        self.geometry("640x360")
        self.resizable(False, False)

        self.file_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.status = tk.StringVar(value="Choose an .xlsx or .xlsm workbook to begin.")
        self.summary = None

        pad = {"padx": 12, "pady": 6}
        ttk.Label(self, text="Workbook:").grid(row=0, column=0, sticky="w", **pad)
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
        p = filedialog.askopenfilename(filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if p:
            self.file_var.set(p)
            if not self.out_var.get():
                self.out_var.set(os.path.join(os.path.dirname(p), "SheetFix_output"))

    def pick_out(self):
        p = filedialog.askdirectory()
        if p:
            self.out_var.set(p)

    def start(self):
        src, out = self.file_var.get().strip(), self.out_var.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showerror("SheetFix", "Please choose a workbook file first.")
            return
        if not out:
            messagebox.showerror("SheetFix", "Please choose an output folder.")
            return
        self.run_btn.config(state="disabled")
        self.report_btn.config(state="disabled")
        self.folder_btn.config(state="disabled")
        self.status.set("Running audit → fix → re-check… this can take a few seconds.")
        threading.Thread(target=self._work, args=(src, out), daemon=True).start()

    def _work(self, src, out):
        try:
            self.summary = run(src, out)
            s = self.summary
            text = (f"Done. Findings before: {len(s['before'])} · auto-fixed: {len(s['fixes'])} · "
                    f"resolved on re-check: {len(s['resolved'])} · remaining (need your judgment): {len(s['remaining'])} · "
                    f"new issues introduced: {len(s['new_issues'])}.\nFixed file: {s['fixed_path']}")
        except Exception as exc:  # surface the real error, don't swallow it
            text = f"Something went wrong: {exc}"
            self.summary = None
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
        if self.summary:
            self._open(self.summary["report_path"])

    def open_folder(self):
        if self.summary:
            self._open(os.path.dirname(self.summary["report_path"]))


if __name__ == "__main__":
    App().mainloop()
