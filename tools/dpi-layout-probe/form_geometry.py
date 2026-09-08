"""Structure-aware Access form geometry canonicalizer.

This is the reference oracle for the VBA port in clsFormGeometryCanonicalizer.
The algorithm is intentionally straightforward so the two implementations can
stay aligned: parse named blocks, recover omitted sizes from LayoutCached*
edges, snap layout-group tracks and boundary pitches to a 60-twip lattice,
then derive Left/Top and spanning bounds. Free-positioned controls are left
alone. Underdetermined groups are reported and left unchanged.
"""

from __future__ import annotations

import argparse
import collections
import json
import re
import statistics
from pathlib import Path
from typing import Any


GRID_TWIPS = 60
GEOMETRY = ("Left", "Top", "Width", "Height")
CACHE_EDGE = {
    "Left": "LayoutCachedLeft",
    "Top": "LayoutCachedTop",
    "Width": "LayoutCachedWidth",
    "Height": "LayoutCachedHeight",
}
NUMERIC_PROPERTY = re.compile(r"^(\w+)\s*=\s*(-?\d+)$")
NAME_PROPERTY = re.compile(r'^Name\s*=\s*"([^"]*)"')
INLINE_BINARY = re.compile(r"^\w+\s*=\s*Begin$")
SECTION_TYPES = {"Section", "FormHeader", "FormFooter", "PageHeader", "PageFooter"}
CONTAINER_TYPES = SECTION_TYPES | {"Form", "Report", "Tab", "Page"}


def snap(value: int, grid: int = GRID_TWIPS) -> int:
    """Banker's-round to the lattice, matching VBA Round() / Python round()."""
    return int(round(value / grid)) * grid


def snap_up(value: int, grid: int = GRID_TWIPS) -> int:
    """Round up to the next lattice point so envelope growth never clips."""
    snapped = snap(value, grid)
    return snapped if snapped >= value else snapped + grid


def read_form(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe"):
        return raw[2:].decode("utf-16-le", errors="replace")
    return raw.decode("utf-8-sig", errors="replace")


def normalize_newlines(text: str) -> str:
    normalized = text.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\n", "\r\n")
    if not normalized.endswith("\r\n"):
        normalized += "\r\n"
    return normalized


def write_form(path: Path, text: str, encoding: str = "utf-16-le") -> None:
    normalized = normalize_newlines(text)
    if encoding == "utf-16-le":
        path.write_bytes(b"\xff\xfe" + normalized.encode("utf-16-le"))
    else:
        path.write_bytes(b"\xef\xbb\xbf" + normalized.encode("utf-8"))


def drop_layout_cached(text: str) -> str:
    """Remove LayoutCached* lines after they have been used as edge data."""
    newline = "\r\n" if "\r\n" in text else "\n"
    lines = text.replace("\r\r\n", "\n").replace("\r\n", "\n").split("\n")
    kept = [line for line in lines if not line.strip().startswith("LayoutCached")]
    body = newline.join(kept)
    if text.endswith(("\r\n", "\n")) and not body.endswith(newline):
        body += newline
    return body


class Block:
    def __init__(self, block_type: str, start_line: int) -> None:
        self.block_type = block_type
        self.start_line = start_line
        self.end_line = start_line
        self.props: dict[str, int] = {}
        self.prop_lines: dict[str, int] = {}
        self.name: str | None = None
        self.parent: Block | None = None
        self.children: list[Block] = []
        self.section: Block | None = None
        self.tab: Block | None = None

    @property
    def is_layout(self) -> bool:
        return "LayoutGroup" in self.props or "GroupTable" in self.props

    @property
    def group_key(self) -> tuple[int | None, int | None]:
        return (self.props.get("GroupTable"), self.props.get("LayoutGroup"))

    def get_size(self, axis: str) -> int | None:
        if axis == "x":
            if "Width" in self.props:
                return self.props["Width"]
            if "LayoutCachedWidth" in self.props and "Left" in self.props:
                return self.props["LayoutCachedWidth"] - self.props["Left"]
        else:
            if "Height" in self.props:
                return self.props["Height"]
            if "LayoutCachedHeight" in self.props and "Top" in self.props:
                return self.props["LayoutCachedHeight"] - self.props["Top"]
        return None

    def get_pos(self, axis: str) -> int | None:
        key = "Left" if axis == "x" else "Top"
        if key in self.props:
            return self.props[key]
        cache = CACHE_EDGE[key]
        if cache in self.props:
            return self.props[cache]
        return None


class FormDocument:
    def __init__(self, text: str) -> None:
        self.original = text
        self.newline = "\r\n" if "\r\n" in text else "\n"
        normalized = text.replace("\r\r\n", "\n").replace("\r\n", "\n")
        self.lines = normalized.split("\n")
        self.blocks: list[Block] = []
        self.named: dict[str, Block] = {}
        self.root: Block | None = None
        self.warnings: list[str] = []
        self._parse()

    def _parse(self) -> None:
        index = 0
        root = None
        while index < len(self.lines):
            text = self.lines[index].strip()
            if text == "CodeBehindForm":
                break
            if text.startswith("Begin"):
                block_type = text[5:].strip() or "Block"
                root, index = self._parse_block(index + 1, block_type, index, None)
                self.root = root
                continue
            index += 1

    def _parse_block(
        self,
        index: int,
        block_type: str,
        start_line: int,
        parent: Block | None,
    ) -> tuple[Block, int]:
        block = Block(block_type, start_line)
        block.parent = parent
        if parent is not None:
            parent.children.append(block)
            block.section = parent if parent.block_type in SECTION_TYPES else parent.section
            block.tab = parent if parent.block_type == "Tab" else parent.tab
        self.blocks.append(block)

        while index < len(self.lines):
            raw = self.lines[index]
            text = raw.strip()
            if text == "CodeBehindForm":
                block.end_line = index
                return block, len(self.lines)

            if INLINE_BINARY.match(text):
                index += 1
                while index < len(self.lines) and self.lines[index].strip() != "End":
                    index += 1
                index += 1
                continue

            if text.startswith("Begin"):
                child_type = text[5:].strip() or "Block"
                _, index = self._parse_block(index + 1, child_type, index, block)
                continue

            if text == "End":
                block.end_line = index
                return block, index + 1

            match = NUMERIC_PROPERTY.match(text)
            if match:
                block.props[match.group(1)] = int(match.group(2))
                block.prop_lines[match.group(1)] = index

            match = NAME_PROPERTY.match(text)
            if match:
                block.name = match.group(1)
                self.named[block.name] = block

            index += 1

        block.end_line = index
        return block, index

    def rewrite_prop(self, block: Block, name: str, value: int) -> bool:
        if name not in block.prop_lines:
            return False
        line_index = block.prop_lines[name]
        indent = re.match(r"^(\s*)", self.lines[line_index]).group(1)
        self.lines[line_index] = f"{indent}{name} ={value}"
        block.props[name] = value
        return True

    def to_text(self) -> str:
        body = self.newline.join(self.lines)
        if self.original.endswith(("\r\n", "\n")) and not body.endswith(self.newline):
            body += self.newline
        return body


class CanonicalizeResult:
    def __init__(self) -> None:
        self.changed_properties = 0
        self.changed_controls: set[str] = set()
        self.skipped_groups: list[str] = []
        self.warnings: list[str] = []
        self.max_abs_shift = 0
        self.shifts: list[int] = []
        self.unsupported = 0
        self.logical_geometry: dict[str, dict[str, int]] = {}
        self.logical_types: dict[str, str] = {}

    def record_shift(self, before: int, after: int, control: str | None) -> None:
        delta = abs(after - before)
        self.shifts.append(delta)
        self.max_abs_shift = max(self.max_abs_shift, delta)
        if delta and control:
            self.changed_controls.add(control)
        if delta:
            self.changed_properties += 1

    def record_logical(self, block: Block, name: str, value: int) -> None:
        """Record planned geometry even when Access omitted its source line."""
        if not block.name:
            return
        self.logical_geometry.setdefault(block.name, {})[name] = value
        self.logical_types[block.name] = block.block_type


def _median_int(values: list[int]) -> int:
    return int(round(statistics.median(values)))


def _track_sizes(
    cells: list[Block],
    axis: str,
    start_name: str,
    end_name: str,
) -> tuple[dict[int, int], list[str]]:
    """Return snapped single-span track sizes and any failure reasons."""
    singles: dict[int, list[int]] = collections.defaultdict(list)
    reasons: list[str] = []
    for cell in cells:
        start = cell.props.get(start_name, 0)
        end = cell.props.get(end_name, start)
        size = cell.get_size(axis)
        if start == end and size is not None:
            singles[start].append(size)
    sizes: dict[int, int] = {}
    for index, observed in singles.items():
        unique = sorted(set(observed))
        if len(unique) > 1 and max(unique) - min(unique) > GRID_TWIPS:
            reasons.append(
                f"{axis}-track {index} has inconsistent sizes {unique}"
            )
        sizes[index] = snap(_median_int(observed))
    return sizes, reasons


def _boundary_positions(
    cells: list[Block],
    axis: str,
    start_name: str,
) -> tuple[dict[int, int], list[str]]:
    """Snap the group origin and each observed pitch, then accumulate."""
    observed: dict[int, list[int]] = collections.defaultdict(list)
    for cell in cells:
        pos = cell.get_pos(axis)
        if pos is None:
            continue
        observed[cell.props.get(start_name, 0)].append(pos)
    if not observed:
        return {}, [f"no {axis} positions"]

    starts = sorted(observed)
    raw = {index: _median_int(values) for index, values in observed.items()}
    positions = {starts[0]: snap(raw[starts[0]])}
    for previous, current in zip(starts, starts[1:]):
        positions[current] = positions[previous] + snap(raw[current] - raw[previous])
    return positions, []


def _derive_span(
    start: int, end: int, positions: dict[int, int], sizes: dict[int, int]
) -> int | None:
    if start not in positions or end not in positions or end not in sizes:
        return None
    return positions[end] + sizes[end] - positions[start]


def _infer_end_sizes(
    cells: list[Block],
    axis: str,
    start_name: str,
    end_name: str,
    positions: dict[int, int],
    sizes: dict[int, int],
) -> None:
    """Fill missing end-track sizes from spanning controls with known bounds."""
    inferred: dict[int, list[int]] = collections.defaultdict(list)
    for cell in cells:
        start = cell.props.get(start_name, 0)
        end = cell.props.get(end_name, start)
        size = cell.get_size(axis)
        if size is None or start == end:
            continue
        if start not in positions or end not in positions:
            continue
        inferred[end].append(size - (positions[end] - positions[start]))
    for index, values in inferred.items():
        if index not in sizes:
            sizes[index] = snap(_median_int(values))


def _plan_axis(
    cells: list[Block],
    axis: str,
    start_name: str,
    end_name: str,
    pos_name: str,
    size_name: str,
) -> tuple[
    list[tuple[Block, str, int, int]],
    list[tuple[Block, str, int]],
    list[str],
]:
    """Return in-place rewrites for one axis, or reasons the group is underdetermined."""
    sizes, reasons = _track_sizes(cells, axis, start_name, end_name)
    positions, pos_reasons = _boundary_positions(cells, axis, start_name)
    reasons.extend(pos_reasons)
    if pos_reasons:
        return [], [], reasons
    _infer_end_sizes(cells, axis, start_name, end_name, positions, sizes)

    planned: list[tuple[Block, str, int, int]] = []
    logical: list[tuple[Block, str, int]] = []
    for cell in cells:
        start = cell.props.get(start_name, 0)
        end = cell.props.get(end_name, start)
        if start in positions:
            logical.append((cell, pos_name, positions[start]))
        if pos_name in cell.props:
            if start not in positions:
                reasons.append(
                    f"{cell.name or cell.block_type}: no {axis} position for track {start}"
                )
                continue
            planned.append((cell, pos_name, cell.props[pos_name], positions[start]))
        derived = (
            _derive_span(start, end, positions, sizes)
            if start != end
            else sizes.get(start)
        )
        if derived is not None:
            logical.append((cell, size_name, derived))
        if size_name in cell.props:
            if derived is None:
                reasons.append(
                    f"{cell.name or cell.block_type}: no {axis} size for tracks {start}-{end}"
                )
                continue
            planned.append((cell, size_name, cell.props[size_name], derived))
    if reasons:
        return [], [], reasons
    return planned, logical, []


def _canonicalize_group(
    cells: list[Block],
    document: FormDocument,
    result: CanonicalizeResult,
    label: str,
) -> None:
    if not cells:
        return
    planned, logical, reasons = _plan_axis(
        cells, "x", "ColumnStart", "ColumnEnd", "Left", "Width"
    )
    if reasons:
        result.skipped_groups.append(f"{label} (horizontal): {'; '.join(reasons)}")
        return

    by_section: dict[int | None, list[Block]] = collections.defaultdict(list)
    for cell in cells:
        section_id = id(cell.section) if cell.section is not None else None
        by_section[section_id].append(cell)
    for section_cells in by_section.values():
        section_plan, section_logical, section_reasons = _plan_axis(
            section_cells, "y", "RowStart", "RowEnd", "Top", "Height"
        )
        if section_reasons:
            result.skipped_groups.append(
                f"{label} (vertical): {'; '.join(section_reasons)}"
            )
            return
        planned.extend(section_plan)
        logical.extend(section_logical)

    for cell, name, before, after in planned:
        if document.rewrite_prop(cell, name, after):
            result.record_shift(before, after, cell.name)
    for cell, name, value in logical:
        result.record_logical(cell, name, value)


def _canonicalize_pages(document: FormDocument, result: CanonicalizeResult) -> None:
    """Snap Tab Width/Height; leave Page insets alone.

    Tab header insets are DPI-dependent and straddle the 60-twip midpoint, so
    Page Left/Top/Width/Height do not unify across standard scaling steps.
    Residual Page drift is documented and accepted. Layout groups on a Page
    are still canonicalized as groups.
    """
    for block in document.blocks:
        if block.block_type != "Tab" or block.is_layout:
            continue
        for name in ("Width", "Height"):
            if name not in block.props:
                continue
            before = block.props[name]
            after = snap(before)
            if document.rewrite_prop(block, name, after):
                result.record_shift(before, after, block.name)


def _child_extent(block: Block, pos_name: str, size_name: str) -> int:
    if pos_name in block.props and size_name in block.props:
        return block.props[pos_name] + block.props[size_name]
    return 0


def _grow_envelopes(document: FormDocument, result: CanonicalizeResult) -> None:
    """Snap form Width and section Height without clipping canonical children."""
    if document.root is None:
        return
    has_layout = any(block.is_layout for block in document.blocks)
    if not has_layout:
        return

    max_right = 0
    for block in document.blocks:
        if block.block_type in {"Form", "Report", "Tab", "Page"}:
            continue
        max_right = max(max_right, _child_extent(block, "Left", "Width"))
    if "Width" in document.root.props:
        before = document.root.props["Width"]
        grown = max(snap(before), snap_up(max_right))
        if document.rewrite_prop(document.root, "Width", grown):
            result.record_shift(before, grown, document.root.name or "Form")

    for section in document.blocks:
        if section.block_type not in SECTION_TYPES or "Height" not in section.props:
            continue
        max_bottom = 0
        for child in document.blocks:
            if child.section is not section:
                continue
            if child.block_type in {"Tab", "Page"} or child.tab is not None:
                continue
            max_bottom = max(max_bottom, _child_extent(child, "Top", "Height"))
        before = section.props["Height"]
        grown = max(snap(before), snap_up(max_bottom)) if max_bottom else snap(before)
        if document.rewrite_prop(section, "Height", grown):
            result.record_shift(before, grown, section.name)


def canonicalize_document(document: FormDocument) -> CanonicalizeResult:
    result = CanonicalizeResult()
    groups: dict[tuple[int | None, int | None], list[Block]] = collections.defaultdict(list)
    for block in document.blocks:
        if block.is_layout:
            groups[block.group_key].append(block)

    for key, cells in groups.items():
        _canonicalize_group(cells, document, result, f"GroupTable={key[0]}/LayoutGroup={key[1]}")

    _canonicalize_pages(document, result)
    _grow_envelopes(document, result)
    result.warnings.extend(document.warnings)
    return result


def canonicalize_text(
    text: str, *, drop_cached: bool = True
) -> tuple[str, CanonicalizeResult]:
    document = FormDocument(text)
    result = canonicalize_document(document)
    output = document.to_text()
    if drop_cached:
        output = drop_layout_cached(output)
    return output, result


def canonicalize_path(
    path: Path, *, drop_cached: bool = True
) -> tuple[str, CanonicalizeResult]:
    return canonicalize_text(read_form(path), drop_cached=drop_cached)


def geometry_signature(
    text: str, *, include_pages: bool = False
) -> dict[str, dict[str, int]]:
    document = FormDocument(text)
    signature: dict[str, dict[str, int]] = {}
    section_ordinals: dict[str, int] = collections.defaultdict(int)
    for block in document.blocks:
        if not include_pages and block.block_type == "Page":
            continue
        if block is document.root:
            name = "$Form"
        elif block.block_type in SECTION_TYPES:
            section_ordinals[block.block_type] += 1
            name = f"${block.block_type}[{section_ordinals[block.block_type]}]"
        elif block.name:
            name = block.name
        else:
            continue
        if block is document.root:
            props = (
                {"Width": block.props["Width"]}
                if "Width" in block.props
                else {}
            )
        elif block.block_type in SECTION_TYPES:
            props = (
                {"Height": block.props["Height"]}
                if "Height" in block.props
                else {}
            )
        else:
            props = {
                key: block.props[key] for key in GEOMETRY if key in block.props
            }
            width = block.get_size("x")
            height = block.get_size("y")
            if width is not None:
                props["Width"] = width
            if height is not None:
                props["Height"] = height
        if props:
            signature[name] = props
    return signature


def canonicalized_geometry_signature(
    text: str,
    result: CanonicalizeResult,
    *,
    include_pages: bool = False,
) -> dict[str, dict[str, int]]:
    """Return canonical geometry, including planned values for omitted lines."""
    signature = geometry_signature(text, include_pages=include_pages)
    for name, props in result.logical_geometry.items():
        if not include_pages and result.logical_types.get(name) == "Page":
            continue
        signature.setdefault(name, {}).update(props)
    return signature


def compare_signatures(
    left: dict[str, dict[str, int]],
    right: dict[str, dict[str, int]],
) -> dict[str, Any]:
    names = set(left) | set(right)
    diffs: list[dict[str, Any]] = []
    for name in sorted(names):
        if name not in left or name not in right:
            diffs.append({"name": name, "reason": "missing"})
            continue
        for prop in GEOMETRY:
            left_value = left[name].get(prop)
            right_value = right[name].get(prop)
            if left_value is None and right_value is None:
                continue
            # Access may add or omit a redundant size line. Canonicalized
            # layout signatures overlay the planned value when it can be
            # inferred; a remaining one-sided Width/Height is treated as an
            # equivalent default-size representation. Positions are required.
            if not name.startswith("$") and prop in {"Width", "Height"} and (
                left_value is None or right_value is None
            ):
                continue
            if left_value != right_value:
                diff = {
                    "name": name,
                    "prop": prop,
                    "left": left_value,
                    "right": right_value,
                }
                if left_value is None or right_value is None:
                    diff["reason"] = "missing-property"
                diffs.append(diff)
    return {"identical": not diffs, "diffs": diffs, "diffCount": len(diffs)}


def main() -> None:
    parser = argparse.ArgumentParser(description="Canonicalize Access .form geometry.")
    parser.add_argument("paths", nargs="+", type=Path, help="Form files or directories.")
    parser.add_argument("--write", action="store_true", help="Rewrite files in place.")
    parser.add_argument("--json", action="store_true", help="Print a machine-readable report.")
    args = parser.parse_args()

    files: list[Path] = []
    for path in args.paths:
        if path.is_dir():
            files.extend(sorted(path.glob("*.form")))
        else:
            files.append(path)

    report = []
    for path in files:
        text, result = canonicalize_path(path)
        entry = {
            "path": str(path),
            "changedProperties": result.changed_properties,
            "changedControls": sorted(result.changed_controls),
            "maxAbsShift": result.max_abs_shift,
            "skippedGroups": result.skipped_groups,
            "warnings": result.warnings,
        }
        report.append(entry)
        if args.write and result.changed_properties:
            encoding = "utf-16-le" if path.read_bytes().startswith(b"\xff\xfe") else "utf-8"
            write_form(path, text, encoding)
        if not args.json:
            print(
                f"{path.name}: {result.changed_properties} props, "
                f"max shift {result.max_abs_shift}, "
                f"skipped {len(result.skipped_groups)}"
            )
            for warning in result.warnings[:8]:
                print(f"  {warning}")

    if args.json:
        print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
