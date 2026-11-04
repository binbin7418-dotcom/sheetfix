# SheetFix report — audit, fix & re-check

Workbook: bluepine_sample.xlsx
- Findings before: 13  ·  auto-fixed: 5  ·  resolved on re-check: 8  ·  remaining: 5  ·  new issues introduced: 0

## Fixes applied (each one re-verified by re-running the audit)
- [Budget!B21] NUMBERS_STORED_AS_TEXT: `2,400` -> `2400` (text number converted to a real number)
- [Budget!A19] WHITESPACE_KEY: ` North ` -> `North` (leading/trailing spaces trimmed)
- [Budget!B11] FORMULA_DRIFT: `=SUM(B8:B9)` -> `=SUM(B8:B10)` (formula restored to the block's dominant pattern (translated from C11))
- [Budget!B7] FORMULA_DRIFT: `=SUM(B5:B6)+B6` -> `=SUM(B5:B6)` (formula restored to the block's dominant pattern (translated from C7))
- [Budget!E7] FORMULA_DRIFT: `=SUM(B7:C7)` -> `=SUM(B7:D7)` (formula restored to the block's dominant pattern (translated from E6))

## Remaining findings — need your judgment, not a mechanical fix
- [Critical] [Budget!B15] LIVE_ERROR: Cell contains live spreadsheet error. Suggestion: Trace the formula precedent chain and resolve the underlying spreadsheet error.
- [High] [Budget!B15] BROKEN_REFERENCE: Formula contains deleted reference. Suggestion: Restore the deleted reference or rebuild the formula from intended source cells.
- [High] [Budget!E10] HARDCODE_IN_FORMULA_BLOCK: Hardcoded value inside formula block. Suggestion: Confirm whether this is an intentional plug. If not, restore the formula pattern.
- [Medium] [Budget!A18, Budget!A19] DUPLICATE_KEY: Duplicate key in lookup range. Suggestion: Make the keys unique or confirm the lookup is meant to return the first match.
- [Medium] [Budget!B14] HIDDEN_STRUCTURE_IN_TOTAL: Formula inputs include hidden structure. Suggestion: Confirm hidden inputs are intentional and disclosed in visible workbook documentation.

Scope note: fixes are mechanical and evidence-backed (text numbers, padded keys, drifted formulas restored to the block pattern, sum ranges missing an adjacent row). Anything that changes business meaning is reported, never silently altered. All processing is local; the workbook never leaves this machine.
