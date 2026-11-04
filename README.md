# SheetFix — Excel Audit, Fix & Re-Check

Local Windows tool: audits an .xlsx/.xlsm workbook, mechanically fixes what is
safely fixable (numbers stored as text, padded keys, drifted formulas, SUM ranges
that stop a row short), then re-audits the repaired file and reports what remains.

- Engine: `sheet_fix.py` (CLI: `python sheet_fix.py file.xlsx outdir`)
- Batch: `sheet_batch.py` — queue several workbooks (files and/or folders) in
  one run; each gets its own output subfolder plus a one-page
  `batch_summary.md` roll-up (per-workbook fixed/remaining counts + totals).
  CLI: `python sheet_batch.py --out outdir file1.xlsx file2.xlsx folder/`
  (equivalently `python sheet_fix.py file1.xlsx file2.xlsx outdir`).
  The single-file engine `sheet_fix.run()` is unchanged; the batch layer
  simply loops over it.
- GUI: `gui.py` — file picker accepts multiple workbooks at once; a
  multi-workbook run shows batch totals and opens `batch_summary.md`.
- Windows build: GitHub Actions workflow included (see the Building section below)
- Audit engine: spreadsheet-auditor by Pete Hottelet (MIT, see AUDITOR_LICENSE.txt)

Everything runs locally; workbooks never leave the user's machine.

## Building the Windows package
PyInstaller cannot cross-compile; build on Windows or via the included GitHub
Actions workflow (Actions tab -> build-windows-exe -> Run workflow; the
SheetFix-Windows.zip artifact is the shippable package). On a Windows PC:
`pip install "spreadsheet-auditor[all]" openpyxl pyinstaller` then
`pyinstaller sheetfix.spec`, and zip `dist/SheetFix/`.
