"""Validate the geometry oracle against captured DPI outputs and optional corpora."""

from __future__ import annotations

import argparse
import json
import statistics
from collections import defaultdict
from pathlib import Path
from typing import Any

from form_geometry import (
    canonicalized_geometry_signature,
    canonicalize_path,
    canonicalize_text,
    compare_signatures,
)


HERE = Path(__file__).resolve().parent
RESULTS = HERE / "results"


def result_runs() -> list[Path]:
    runs = []
    for path in sorted(RESULTS.iterdir()):
        if not path.is_dir():
            continue
        if (path / "metadata.json").is_file() and (path / "design-save").is_dir():
            if path.name.startswith("canonical-"):
                continue
            runs.append(path)
    return runs


def unify_folder(folder_name: str) -> dict[str, Any]:
    """Canonicalize the same fixture across every captured run and compare."""
    by_form: dict[str, list[tuple[str, dict[str, dict[str, int]], list[str]]]] = defaultdict(list)
    for run in result_runs():
        folder = run / folder_name
        if not folder.is_dir():
            continue
        for path in sorted(folder.glob("*.form")):
            text, result = canonicalize_path(path)
            by_form[path.name].append(
                (
                    run.name,
                    canonicalized_geometry_signature(
                        text, result, include_pages=False
                    ),
                    result.skipped_groups,
                )
            )

    report = []
    for name, entries in sorted(by_form.items()):
        baseline = entries[0][1]
        diffs = []
        skipped = [
            f"{run_name}: {item}"
            for run_name, _signature, skipped_groups in entries
            for item in skipped_groups
        ]
        for run_name, signature, _skipped_groups in entries[1:]:
            comparison = compare_signatures(baseline, signature)
            if not comparison["identical"]:
                diffs.append({"run": run_name, "diffCount": comparison["diffCount"]})
        report.append(
            {
                "form": name,
                "runs": [entry[0] for entry in entries],
                "identical": not diffs and not skipped,
                "mismatchedRuns": diffs,
                "skipped": skipped,
            }
        )
    return {
        "folder": folder_name,
        "forms": report,
        "allIdentical": all(item["identical"] for item in report),
    }


def corpus_report(folder: Path) -> dict[str, Any]:
    shifts: list[int] = []
    skipped = []
    max_shift = 0
    max_form = ""
    changed_forms = 0
    files = sorted(folder.glob("*.form"))
    for path in files:
        _text, result = canonicalize_path(path)
        if result.shifts:
            form_max = max(result.shifts)
            shifts.append(form_max)
            if form_max > max_shift:
                max_shift = form_max
                max_form = path.name
        if result.changed_properties:
            changed_forms += 1
        for item in result.skipped_groups:
            skipped.append(f"{path.name}: {item}")
    return {
        "folder": str(folder),
        "forms": len(files),
        "changedForms": changed_forms,
        "maxAbsShift": max_shift,
        "maxForm": max_form,
        "p95Shift": int(statistics.quantiles(shifts, n=20)[18]) if len(shifts) >= 20 else max_shift,
        "skippedGroups": skipped[:40],
        "skippedCount": len(skipped),
    }


def idempotence_check(paths: list[Path]) -> dict[str, Any]:
    failures = []
    for path in paths:
        first, _ = canonicalize_path(path)
        second, _ = canonicalize_text(first)
        if first != second:
            failures.append(path.name)
    return {"checked": len(paths), "failures": failures}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--corpus",
        type=Path,
        help="Optional extra folder containing .form files.",
    )
    args = parser.parse_args()

    report: dict[str, Any] = {
        "designSave": unify_folder("design-save"),
        "plain": unify_folder("plain"),
    }
    fixture_dir = HERE / "fixtures"
    if fixture_dir.is_dir():
        report["idempotence"] = idempotence_check(sorted(fixture_dir.glob("*.form")))
        report["fixtureShifts"] = corpus_report(fixture_dir)
    if args.corpus and args.corpus.is_dir():
        report["corpus"] = corpus_report(args.corpus)
        report["corpusIdempotence"] = idempotence_check(sorted(args.corpus.glob("*.form")))

    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
