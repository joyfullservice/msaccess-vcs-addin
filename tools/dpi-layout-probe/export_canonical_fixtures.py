"""Write oracle-canonical copies of the probe fixtures for Access round-trips."""

from __future__ import annotations

from pathlib import Path

from form_geometry import canonicalize_path, write_form


HERE = Path(__file__).resolve().parent
SOURCE = HERE / "fixtures"
DEST = HERE / "canonical-fixtures"
# Compact fixtures are oracle-only; they are not complete enough for LoadFromText.
ACCESS_PROOF = {
    "TabularGrid.form",
    "RightAnchored.form",
    "MultiRowGrid.form",
    "StretchAnchored.form",
    "SpacerGrid.form",
    "StackedGrid.form",
}


def main() -> None:
    DEST.mkdir(parents=True, exist_ok=True)
    proof = DEST / "access-proof"
    proof.mkdir(parents=True, exist_ok=True)
    for path in sorted(SOURCE.glob("*.form")):
        text, result = canonicalize_path(path)
        write_form(DEST / path.name, text)
        if path.name in ACCESS_PROOF:
            write_form(proof / path.name, text)
        print(
            f"{path.name}: {result.changed_properties} props, "
            f"max shift {result.max_abs_shift}, skipped {len(result.skipped_groups)}"
        )


if __name__ == "__main__":
    main()
