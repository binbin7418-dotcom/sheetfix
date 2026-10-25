# SheetFix — Excel Audit, Fix & Re-Check

Local Windows tool: audits an .xlsx/.xlsm workbook, mechanically fixes what is
safely fixable (numbers stored as text, padded keys, drifted formulas, SUM ranges
that stop a row short), then re-audits the repaired file and reports what remains.

- Engine: `sheet_fix.py` (CLI: `python sheet_fix.py file.xlsx outdir`)
- GUI: `gui.py`
- Windows build: GitHub Actions workflow included (see the Building section below)
- Audit engine: spreadsheet-auditor by Pete Hottelet (MIT, see AUDITOR_LICENSE.txt)

Everything runs locally; workbooks never leave the user's machine.

## Building the Windows package
PyInstaller cannot cross-compile; build on Windows or via the included GitHub
Actions workflow (Actions tab -> build-windows-exe -> Run workflow; the
SheetFix-Windows.zip artifact is the shippable package). On a Windows PC:
`pip install "spreadsheet-auditor[all]" openpyxl pyinstaller` then
`pyinstaller sheetfix.spec`, and zip `dist/SheetFix/`.
