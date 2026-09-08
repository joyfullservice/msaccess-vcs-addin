"""Check N(P_d(C)) = C for a completed probe run of canonical fixtures."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from form_geometry import (
    canonicalized_geometry_signature,
    canonicalize_path,
    canonicalize_text,
    compare_signatures,
    geometry_signature,
    read_form,
)


def control_names(text: str) -> list[str]:
    return [
        line.split("=", 1)[1].strip().strip('"')
        for line in text.replace("\r\n", "\n").split("\n")
        if line.strip().startswith("Name =")
    ]


def empty_cells(text: str) -> int:
    return text.count("Begin EmptyCell")


def read_metadata(run_dir: Path) -> dict:
    path = run_dir / "metadata.json"
    if not path.is_file():
        return {}
    metadata = json.loads(path.read_text(encoding="utf-8-sig"))
    return {
        "label": metadata.get("label"),
        "accessEffectiveDpi": metadata.get("accessEffectiveDpi"),
        "accessEffectiveScale": metadata.get("accessEffectiveScale"),
        "selectedDisplay": metadata.get("selectedDisplay"),
        "fixtureDirectory": metadata.get("fixtureDirectory"),
    }


def compare_run(run_dir: Path) -> dict:
    input_paths = sorted((run_dir / "input").glob("*.form"))
    metadata = read_metadata(run_dir)
    metadata_valid = metadata.get("accessEffectiveDpi") is not None
    report = {
        "run": str(run_dir),
        "metadata": metadata,
        "metadataValid": metadata_valid,
        "inputCount": len(input_paths),
        "forms": [],
        "ok": bool(input_paths) and metadata_valid,
    }
    if not input_paths:
        report["error"] = "No .form inputs found"
        return report
    if not metadata_valid:
        report["error"] = "metadata.json has no measured Access DPI"

    for path in input_paths:
        canonical = read_form(path)
        normalized_input, input_result = canonicalize_text(canonical)
        input_canonical = canonical == normalized_input
        canonical_signature = canonicalized_geometry_signature(
            normalized_input, input_result, include_pages=False
        )
        entry = {
            "form": path.name,
            "inputCanonical": input_canonical,
            "paths": {},
        }
        if not input_canonical:
            report["ok"] = False
        for folder in ("plain", "design-save"):
            exported = run_dir / folder / path.name
            projected = read_form(exported)
            renormalized, result = canonicalize_path(exported)
            invariant = compare_signatures(
                canonical_signature,
                canonicalized_geometry_signature(
                    renormalized, result, include_pages=False
                ),
            )
            projection = compare_signatures(
                canonical_signature,
                geometry_signature(projected, include_pages=False),
            )
            twice, _ = canonicalize_text(renormalized)
            folder_ok = (
                input_canonical
                and invariant["identical"]
                and renormalized == twice
                and control_names(canonical) == control_names(projected)
                and empty_cells(canonical) == empty_cells(projected)
                and not result.skipped_groups
            )
            projection_shifts = [
                abs(item["left"] - item["right"])
                for item in projection["diffs"]
                if item.get("left") is not None and item.get("right") is not None
            ]
            entry["paths"][folder] = {
                "invariantIdentical": invariant["identical"],
                "invariantDiffCount": invariant["diffCount"],
                "invariantDiffs": invariant["diffs"][:12],
                "projectionIdentical": projection["identical"],
                "projectionDiffCount": projection["diffCount"],
                "projectionMaxAbsShift": max(projection_shifts) if projection_shifts else 0,
                "projectionDiffs": projection["diffs"][:12],
                "idempotent": renormalized == twice,
                "namesMatch": control_names(canonical) == control_names(projected),
                "emptyCellsMatch": empty_cells(canonical) == empty_cells(projected),
                "skipped": result.skipped_groups,
                "ok": folder_ok,
            }
            if not folder_ok:
                report["ok"] = False
        report["forms"].append(entry)
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("run_dir", type=Path)
    args = parser.parse_args()
    report = compare_run(args.run_dir)
    print(json.dumps(report, indent=2))
    raise SystemExit(0 if report["ok"] else 1)


if __name__ == "__main__":
    main()
