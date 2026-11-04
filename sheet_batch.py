#!/usr/bin/env python3
"""
SheetFix batch layer — multi-workbook queue on top of the single-file engine.

The engine core (sheet_fix.run) is unchanged: this module only collects
inputs (files and/or folders), calls run() once per workbook into its own
subfolder, and writes one batch_summary.md covering the whole queue.

Usage:
    python3 sheet_batch.py --out <outdir> <workbook-or-folder> [<workbook-or-folder> ...]
    python3 sheet_fix.py  <workbook-or-folder> [<workbook-or-folder> ...] <outdir>

Outputs in <outdir>:
    batch_summary.md          one-page roll-up (per-workbook counts + totals)
    <workbook-stem>/          that workbook's normal SheetFix outputs
                              (<stem>_fixed.xlsx, fix_report.md, findings_*.json)
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sheet_fix import run  # noqa: E402

WORKBOOK_EXTS = (".xlsx", ".xlsm")


def collect_workbooks(inputs):
    """Expand files/folders into an ordered, de-duplicated workbook list.

    Folders are scanned (top level only) for .xlsx/.xlsm files, sorted by
    name. Excel lock files (~$*) and SheetFix's own *_fixed outputs are
    skipped so re-running a queue over an output folder stays clean.
    """
    found, seen = [], set()
    for item in inputs:
        if os.path.isdir(item):
            candidates = [os.path.join(item, name) for name in sorted(os.listdir(item))
                          if name.lower().endswith(WORKBOOK_EXTS)]
        else:
            candidates = [item]
        for path in candidates:
            base = os.path.basename(path)
            stem_lower = os.path.splitext(base)[0].lower()
            if base.startswith("~$") or stem_lower.endswith("_fixed"):
                continue
            real = os.path.abspath(path)
            if real not in seen and os.path.isfile(path):
                seen.add(real)
                found.append(path)
    return found


def _verdict(entry):
    if entry.get("error"):
        return "FAILED — could not be processed"
    if entry["new_issues"]:
        return "Review — fixing introduced new findings"
    if entry["remaining"] == 0:
        return "Clean — passes the full audit set after fixing"
    return f"Fixed — {entry['remaining']} finding(s) still need your judgment"


def run_batch(inputs, outdir, cli="spreadsheet-auditor"):
    """Run the single-file engine once per workbook; write batch_summary.md.

    Returns a dict with per-workbook entries and totals. One bad workbook
    never stops the queue: its error is recorded and the batch continues.
    """
    workbooks = collect_workbooks(inputs)
    if not workbooks:
        raise ValueError("no .xlsx/.xlsm workbooks found in the given inputs")
    os.makedirs(outdir, exist_ok=True)

    entries, used_stems = [], set()
    for src in workbooks:
        stem = os.path.splitext(os.path.basename(src))[0]
        sub = stem
        n = 2
        while sub in used_stems:  # same filename from two folders
            sub = f"{stem}-{n}"
            n += 1
        used_stems.add(sub)
        subdir = os.path.join(outdir, sub)
        try:
            s = run(src, subdir, cli)
            entries.append({
                "workbook": os.path.basename(src), "subdir": sub,
                "before": len(s["before"]), "fixes": len(s["fixes"]),
                "resolved": len(s["resolved"]), "remaining": len(s["remaining"]),
                "new_issues": len(s["new_issues"]),
                "fixed_path": s["fixed_path"], "report_path": s["report_path"],
            })
        except Exception as exc:  # keep the queue moving; record, don't hide
            entries.append({"workbook": os.path.basename(src), "subdir": sub,
                            "before": 0, "fixes": 0, "resolved": 0, "remaining": 0,
                            "new_issues": 0, "error": str(exc)})

    totals = {k: sum(e[k] for e in entries)
              for k in ("before", "fixes", "resolved", "remaining", "new_issues")}
    failed = sum(1 for e in entries if e.get("error"))

    lines = ["# SheetFix batch summary — audit, fix & re-check", ""]
    lines.append(f"Workbooks processed: {len(entries)}  ·  failed: {failed}")
    lines.append("")
    lines.append("| Workbook | Findings before | Auto-fixed | Resolved on re-check | Remaining (your judgment) | New issues | Verdict |")
    lines.append("|---|---:|---:|---:|---:|---:|---|")
    for e in entries:
        lines.append(f"| {e['workbook']} | {e['before']} | {e['fixes']} | {e['resolved']} | "
                     f"{e['remaining']} | {e['new_issues']} | {_verdict(e)} |")
    lines.append(f"| **Total ({len(entries)} workbooks)** | **{totals['before']}** | **{totals['fixes']}** | "
                 f"**{totals['resolved']}** | **{totals['remaining']}** | **{totals['new_issues']}** | |")
    lines.append("")
    lines.append("## Per-workbook detail")
    for e in entries:
        if e.get("error"):
            lines.append(f"- {e['workbook']}: FAILED — {e['error']}")
        else:
            lines.append(f"- {e['workbook']}: report `{e['subdir']}/fix_report.md` · fixed file `{e['subdir']}/{os.path.basename(e['fixed_path'])}`")
    lines.append("")
    lines.append("Scope note: each workbook went through the same audit → fix → re-check pipeline as a single SheetFix run. Fixes are mechanical and evidence-backed; anything that changes business meaning is reported, never silently altered. All processing is local.")
    summary_path = os.path.join(outdir, "batch_summary.md")
    with open(summary_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    return {"entries": entries, "totals": totals, "failed": failed,
            "summary_path": summary_path}


def main(argv=None):
    ap = argparse.ArgumentParser(description="SheetFix batch: audit, fix & re-check many workbooks in one run.")
    ap.add_argument("inputs", nargs="+", help="workbook files (.xlsx/.xlsm) and/or folders containing them")
    ap.add_argument("-o", "--out", required=True, help="output folder (batch_summary.md lands here)")
    ap.add_argument("--auditor-cli", default="spreadsheet-auditor",
                    help="fallback auditor CLI if the in-process engine is unavailable")
    args = ap.parse_args(argv)
    result = run_batch(args.inputs, args.out, args.auditor_cli)
    t = result["totals"]
    print(f"workbooks={len(result['entries'])} failed={result['failed']} "
          f"before={t['before']} fixes={t['fixes']} resolved={t['resolved']} "
          f"remaining={t['remaining']} new={t['new_issues']} summary={result['summary_path']}")
    return 1 if result["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
