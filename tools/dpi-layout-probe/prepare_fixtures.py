"""Build code-free DPI layout probe fixtures from known exported forms.

This is a maintainer utility. Normal probe runs consume the checked-in files
under fixtures/ and do not need the original source forms.

Always write UTF-16-LE with a single CRLF so LoadFromText sees the same line
structure Access writes. Do not inject anchors without an Access design-save;
RightAnchored is produced from a real Access-saved copy when available.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from form_geometry import write_form


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
ACCESS_SAVED = HERE / "results" / "2304x1536-100pct" / "design-save"
HARVESTED_SOURCES = HERE / "fixture-sources.json"
PLACEHOLDER = "DPI Layout Probe"
TEXT_PROPERTIES = frozenset({"Caption", "ControlTipText", "StatusBarText"})

DROP_PROPERTIES = {
    "RecordSource",
    "ControlSource",
    "RowSource",
    "RowSourceType",
    "OrderBy",
    "Filter",
    "SourceObject",
    "LinkChildFields",
    "LinkMasterFields",
    "InputParameters",
}

HYPERLINK_PROPERTIES = {
    "HyperlinkAddress": "",
    "HyperlinkSubAddress": "Macro macProbeNoOp",
}

# Add-in forms only. Private source paths and source-specific substitutions
# live in gitignored fixture-sources.json and never enter the repository.
SOURCES = {
    "MultiRowGrid.form": (
        REPO / "Version Control.accda.src" / "forms" / "frmVCSConflictList.form"
    ),
    "StretchAnchored.form": (
        REPO / "Version Control.accda.src" / "forms" / "frmVCSTableData.form"
    ),
}


def read_form(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe"):
        return raw[2:].decode("utf-16-le")
    return raw.decode("utf-8-sig", errors="replace")


def load_config() -> tuple[dict[str, Path], dict[str, str]]:
    sources = dict(SOURCES)
    replacements: dict[str, str] = {}
    if HARVESTED_SOURCES.is_file():
        config = json.loads(HARVESTED_SOURCES.read_text(encoding="utf-8"))
        for name, raw_path in config.get("sources", {}).items():
            sources[name] = Path(raw_path)
        replacements.update(config.get("replacements", {}))
    return sources, replacements


def sanitize(text: str, replacements: dict[str, str] | None = None) -> str:
    """Remove code, bindings, and private source details while keeping layout."""
    lines = text.replace("\r\r\n", "\n").replace("\r\n", "\n").split("\n")
    output: list[str] = []
    index = 0

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if stripped == "CodeBehindForm":
            break

        match = re.match(r"^(\s*)(\w+)(\s*=.*)$", line)
        if match:
            indent, prop, _rest = match.group(1), match.group(2), match.group(3)
            if prop == "HasModule":
                line = f"{indent}HasModule =0"
            elif prop in DROP_PROPERTIES or prop.startswith(("On", "Before", "After")):
                if stripped.endswith("= Begin"):
                    depth = 1
                    index += 1
                    while index < len(lines) and depth:
                        nested = lines[index].strip()
                        if nested.startswith("Begin") or nested.endswith("= Begin"):
                            depth += 1
                        elif nested == "End":
                            depth -= 1
                        index += 1
                else:
                    index += 1
                    while index < len(lines) and lines[index].strip().startswith('"'):
                        index += 1
                continue
            elif prop in TEXT_PROPERTIES:
                line = f'{indent}{prop} ="{PLACEHOLDER}"'
                index += 1
                while index < len(lines) and lines[index].strip().startswith('"'):
                    index += 1
                output.append(line)
                continue
            elif prop == "Tag":
                line = f'{indent}Tag ="probe"'
                index += 1
                while index < len(lines) and lines[index].strip().startswith('"'):
                    index += 1
                output.append(line)
                continue
            elif prop in HYPERLINK_PROPERTIES:
                line = f'{indent}{prop} ="{HYPERLINK_PROPERTIES[prop]}"'
                index += 1
                while index < len(lines) and lines[index].strip().startswith('"'):
                    index += 1
                output.append(line)
                continue
        output.append(line)
        index += 1

    while output and not output[-1]:
        output.pop()
    sanitized = "\r\n".join(output) + "\r\n"
    for old, new in sorted((replacements or {}).items(), key=lambda item: len(item[0]), reverse=True):
        sanitized = sanitized.replace(old, new)
    return sanitized


def normalize_existing_fixtures(replacements: dict[str, str] | None = None) -> None:
    """Re-sanitize checked-in fixtures in place (single CRLF, generic text)."""
    fixture_dir = HERE / "fixtures"
    for path in sorted(fixture_dir.glob("*.form")):
        write_form(path, sanitize(read_form(path), replacements))
        print(f"normalized {path.name}")


def main() -> None:
    fixture_dir = HERE / "fixtures"
    fixture_dir.mkdir(parents=True, exist_ok=True)
    sources, replacements = load_config()

    missing = [str(source) for source in sources.values() if not source.is_file()]
    if missing:
        print("Source forms are not available; normalizing existing fixtures only.")
        normalize_existing_fixtures(replacements)
        return

    for output_name, source in sources.items():
        content = sanitize(read_form(source), replacements)
        write_form(fixture_dir / output_name, content)
        print(f"{output_name}: {source}")

    saved_right = ACCESS_SAVED / "RightAnchored.form"
    if saved_right.is_file() and "HorizontalAnchor" in read_form(saved_right):
        write_form(
            fixture_dir / "RightAnchored.form",
            sanitize(read_form(saved_right), replacements),
        )
        print(f"RightAnchored.form: {saved_right} (Access-saved)")
    else:
        print(
            "RightAnchored.form left unchanged: inject anchors in Access, "
            "design-save, then copy the result into fixtures/."
        )


if __name__ == "__main__":
    main()
