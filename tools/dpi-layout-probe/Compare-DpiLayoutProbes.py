"""Compare Microsoft Access form geometry across DPI probe runs."""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
GEOMETRY = ("Left", "Top", "Width", "Height")
DERIVED = (
    "LayoutCachedLeft",
    "LayoutCachedTop",
    "LayoutCachedWidth",
    "LayoutCachedHeight",
)
NUMERIC_PROPERTY = re.compile(r"^(\w+)\s*=\s*(-?\d+)$")
NAME_PROPERTY = re.compile(r'^Name\s*=\s*"([^"]*)"')
INLINE_BINARY = re.compile(r"^\w+\s*=\s*Begin$")


def read_form(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe"):
        return raw[2:].decode("utf-16-le", errors="replace")
    return raw.decode("utf-8-sig", errors="replace")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def parse_controls(path: Path) -> dict[str, dict[str, Any]]:
    """Parse named SaveAsText blocks without mistaking binary blobs for blocks."""
    lines = read_form(path).replace("\r\n", "\n").split("\n")
    controls: dict[str, dict[str, Any]] = {}

    def parse_block(index: int, block_type: str) -> tuple[dict[str, Any], int]:
        props: dict[str, Any] = {"_type": block_type}
        child_number = 0

        while index < len(lines):
            text = lines[index].strip()
            if text == "CodeBehindForm":
                return props, len(lines)

            if INLINE_BINARY.match(text):
                index += 1
                while index < len(lines) and lines[index].strip() != "End":
                    index += 1
                index += 1
                continue

            if text.startswith("Begin"):
                child_type = text[5:].strip() or "Block"
                child_number += 1
                child, index = parse_block(index + 1, child_type)
                child["_ordinal"] = child_number
                name = child.get("_name")
                if name:
                    controls[name] = child
                continue

            if text == "End":
                return props, index + 1

            match = NUMERIC_PROPERTY.match(text)
            if match:
                props[match.group(1)] = int(match.group(2))

            match = NAME_PROPERTY.match(text)
            if match:
                props["_name"] = match.group(1)

            index += 1

        return props, index

    index = 0
    while index < len(lines):
        text = lines[index].strip()
        if text == "CodeBehindForm":
            break
        if text.startswith("Begin"):
            block_type = text[5:].strip() or "Block"
            root, index = parse_block(index + 1, block_type)
            name = root.get("_name")
            if name:
                controls[name] = root
            continue
        index += 1

    for props in controls.values():
        props["_layout"] = "LayoutGroup" in props or "GroupTable" in props
    return controls


def layout_gaps(controls: dict[str, dict[str, Any]]) -> dict[str, int]:
    """Return a histogram of gaps between consecutive explicit layout columns."""
    rows: dict[tuple[int, int, int], dict[int, dict[str, Any]]] = (
        collections.defaultdict(dict)
    )
    for props in controls.values():
        if not props.get("_layout"):
            continue
        if not all(prop in props for prop in ("Left", "Width")):
            continue
        group = props.get("GroupTable", props.get("LayoutGroup", 0))
        row = props.get("RowStart", 0)
        top = props.get("Top", 0)
        column = props.get("ColumnStart", 0)
        rows[(group, row, top)][column] = props

    gaps: collections.Counter[int] = collections.Counter()
    for cells in rows.values():
        for column, left in cells.items():
            right = cells.get(column + 1)
            if right:
                gaps[right["Left"] - (left["Left"] + left["Width"])] += 1
    return {str(gap): count for gap, count in sorted(gaps.items())}


def compare_forms(base_path: Path, target_path: Path) -> dict[str, Any]:
    base = parse_controls(base_path)
    target = parse_controls(target_path)
    base_names = set(base)
    target_names = set(target)
    common = base_names & target_names

    changed_controls: set[str] = set()
    changed_layout_controls: set[str] = set()
    property_changes: collections.Counter[str] = collections.Counter()
    property_deltas: dict[str, collections.Counter[int]] = collections.defaultdict(
        collections.Counter
    )
    missing_properties: collections.Counter[str] = collections.Counter()
    added_properties: collections.Counter[str] = collections.Counter()

    for name in common:
        before = base[name]
        after = target[name]
        is_layout = bool(before.get("_layout") or after.get("_layout"))

        for prop in GEOMETRY + DERIVED:
            if prop in before and prop not in after:
                missing_properties[prop] += 1
                changed_controls.add(name)
                if is_layout:
                    changed_layout_controls.add(name)
            elif prop not in before and prop in after:
                added_properties[prop] += 1
                changed_controls.add(name)
                if is_layout:
                    changed_layout_controls.add(name)
            elif prop in before and prop in after and before[prop] != after[prop]:
                property_changes[prop] += 1
                property_deltas[prop][after[prop] - before[prop]] += 1
                changed_controls.add(name)
                if is_layout:
                    changed_layout_controls.add(name)

    return {
        "basePath": str(base_path),
        "targetPath": str(target_path),
        "exactMatch": sha256(base_path) == sha256(target_path),
        "baseControlCount": len(base),
        "targetControlCount": len(target),
        "lostControls": sorted(base_names - target_names),
        "addedControls": sorted(target_names - base_names),
        "changedControlCount": len(changed_controls),
        "changedLayoutControlCount": len(changed_layout_controls),
        "propertyChanges": dict(sorted(property_changes.items())),
        "propertyDeltas": {
            prop: {str(delta): count for delta, count in sorted(deltas.items())}
            for prop, deltas in sorted(property_deltas.items())
        },
        "missingProperties": dict(sorted(missing_properties.items())),
        "addedProperties": dict(sorted(added_properties.items())),
        "baseGaps": layout_gaps(base),
        "targetGaps": layout_gaps(target),
    }


def load_run(path: Path) -> dict[str, Any]:
    metadata_path = path / "metadata.json"
    if not metadata_path.is_file():
        raise ValueError(f"No metadata.json found in {path}")
    metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    return {"path": path.resolve(), "metadata": metadata}


def fixture_paths(run: dict[str, Any], folder: str) -> dict[str, Path]:
    path = run["path"] / folder
    return {item.name: item for item in sorted(path.glob("*.form"))}


def comparison_line(name: str, mode: str, result: dict[str, Any]) -> str:
    if result["exactMatch"]:
        verdict = "byte-identical"
    elif (
        not result["lostControls"]
        and not result["addedControls"]
        and result["changedControlCount"] == 0
    ):
        verdict = "geometry-identical"
    else:
        verdict = (
            f"{result['changedLayoutControlCount']} layout controls changed, "
            f"{len(result['lostControls'])} lost"
        )
    return f"  {name:<22} {mode:<12} {verdict}"


def compare_runs(runs: list[dict[str, Any]]) -> tuple[dict[str, Any], str]:
    baseline = runs[0]
    report: dict[str, Any] = {
        "baseline": baseline["metadata"]["label"],
        "runs": [
            {
                "label": run["metadata"]["label"],
                "path": str(run["path"]),
                "dpi": run["metadata"].get("accessEffectiveDpi"),
                "scale": run["metadata"].get("accessEffectiveScale"),
                "display": run["metadata"].get("selectedDisplay"),
            }
            for run in runs
        ],
        "withinRuns": {},
        "crossRun": {},
    }
    lines = ["Access DPI layout probe comparison", ""]

    for run in runs:
        label = run["metadata"]["label"]
        dpi = run["metadata"].get("accessEffectiveDpi")
        scale = run["metadata"].get("accessEffectiveScale")
        lines.append(f"{label}: effective DPI {dpi}, scale {scale}%")
    lines.append("")

    lines.append("Within-run import behavior")
    for run in runs:
        label = run["metadata"]["label"]
        report["withinRuns"][label] = {}
        inputs = fixture_paths(run, "input")
        for mode, folder in (("plain", "plain"), ("design-save", "design-save")):
            outputs = fixture_paths(run, folder)
            for filename in sorted(inputs.keys() & outputs.keys()):
                key = f"{Path(filename).stem}:{mode}"
                result = compare_forms(inputs[filename], outputs[filename])
                report["withinRuns"][label][key] = result
                lines.append(comparison_line(Path(filename).stem, f"{label}/{mode}", result))
    lines.append("")

    lines.append(f"Cross-run behavior (baseline: {baseline['metadata']['label']})")
    for target in runs[1:]:
        target_label = target["metadata"]["label"]
        report["crossRun"][target_label] = {}
        lines.append(f"Against {target_label}:")
        for mode in ("plain", "design-save"):
            base_files = fixture_paths(baseline, mode)
            target_files = fixture_paths(target, mode)
            for filename in sorted(base_files.keys() & target_files.keys()):
                key = f"{Path(filename).stem}:{mode}"
                result = compare_forms(base_files[filename], target_files[filename])
                report["crossRun"][target_label][key] = result
                lines.append(comparison_line(Path(filename).stem, mode, result))
        lines.append("")

    return report, "\n".join(lines).rstrip() + "\n"


def discover_runs(results_directory: Path) -> list[Path]:
    return sorted(
        item.parent
        for item in results_directory.glob("*/metadata.json")
        if item.parent.is_dir()
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compare two or more Run-DpiLayoutProbe.ps1 result folders."
    )
    parser.add_argument(
        "runs",
        nargs="*",
        type=Path,
        help="Run folders in baseline-first order. Defaults to all completed runs.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Output directory. Defaults to results/comparison-<labels>.",
    )
    args = parser.parse_args()

    run_paths = args.runs or discover_runs(HERE / "results")
    if len(run_paths) < 2:
        raise SystemExit("At least two completed probe runs are required.")

    runs = [load_run(path) for path in run_paths]
    labels = [run["metadata"]["label"] for run in runs]
    output = args.output or HERE / "results" / (
        "comparison-" + "-vs-".join(labels)
    )
    output.mkdir(parents=True, exist_ok=False)

    report, text = compare_runs(runs)
    (output / "comparison.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8"
    )
    (output / "comparison.txt").write_text(text, encoding="utf-8")
    print(text, end="")
    print(f"\nSaved report: {output}")


if __name__ == "__main__":
    main()
