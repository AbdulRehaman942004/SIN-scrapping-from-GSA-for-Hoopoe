import argparse
import csv
import re
from pathlib import Path
from typing import Iterable, List, Tuple

# Terms to strip when determining the root name. These can appear anywhere.
REMOVABLE_TERMS = {
    # corporate forms
    "inc", "incorporated", "corp", "corporation", "co", "company", "llc", "l.l.c",
    "ltd", "limited", "gmbh", "s.a.", "s.a", "s.p.a.", "spa", "ag", "kg", "nv",
    "plc", "pty", "pte", "sro", "s.r.o", "srl", "lp", "llp", "pc",
    # common descriptors
    "products", "product", "brands", "brand", "group", "international", "industries",
    "industry", "mfg", "manufacturing", "manufacturers", "division", "div",
    # geography / noise
    "usa", "u.s.a", "u.s.", "us", "america", "american", "north", "south",
    "europe", "european", "asia", "pacific",
}

# Pattern to collapse all non-alphanumeric characters to space for tokenization.
NON_ALNUM_TO_SPACE = re.compile(r"[^0-9a-z]+")

# Pattern to strip all non-alphanumeric characters entirely for the final root token
STRIP_NON_ALNUM = re.compile(r"[^0-9a-z]")


def normalize_to_root(name: str) -> str:
    """Convert a manufacturer display name to a simple root form.

    Steps:
    - lowercase
    - replace non-alphanumeric with spaces for tokenization
    - remove removable terms (corporate forms, descriptors)
    - take the first remaining token; if none, fall back to first alnum chunk
    - strip all non-alphanumeric characters from the chosen token
    """
    if not name:
        return ""

    lower = name.lower()

    # Tokenize by converting non-alnum to spaces, then splitting
    tokens = NON_ALNUM_TO_SPACE.sub(" ", lower).split()

    filtered: List[str] = []
    for token in tokens:
        # ignore pure removable terms
        if token in REMOVABLE_TERMS:
            continue
        filtered.append(token)

    chosen = ""
    if filtered:
        chosen = filtered[0]
    else:
        # Fallback: take first alphanumeric run from original
        alnum_runs = re.findall(r"[0-9a-z]+", lower)
        chosen = alnum_runs[0] if alnum_runs else ""

    # Remove any lingering non-alphanumeric chars (e.g., from hyphen merges)
    root = STRIP_NON_ALNUM.sub("", chosen)
    return root


def read_unique_list(input_path: Path) -> List[str]:
    text = input_path.read_text(encoding="utf-8")
    # Keep order as-is while removing empty lines
    lines = [line.strip() for line in text.splitlines()]
    return [line for line in lines if line]


def build_mapping(names: Iterable[str]) -> List[Tuple[str, str]]:
    mapping: List[Tuple[str, str]] = []
    for original in names:
        root = normalize_to_root(original)
        mapping.append((original, root))
    return mapping


def write_mapping_csv(rows: List[Tuple[str, str]], output_csv: Path) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["original", "root"])
        writer.writerows(rows)


def parse_args() -> argparse.Namespace:
    project_dir = Path(__file__).resolve().parents[1]

    default_input = project_dir / "indentifying unique mfr in the excel" / "unique_manufacturers.txt"
    default_output = Path(__file__).parent / "original_to_root.csv"

    parser = argparse.ArgumentParser(description="Normalize manufacturer names to root form.")
    parser.add_argument("--input", type=Path, default=default_input, help="Path to input unique list (txt)")
    parser.add_argument("--output", type=Path, default=default_output, help="Path to output CSV mapping")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if not args.input.exists():
        print(f"Error: input file not found: {args.input}")
        return

    names = read_unique_list(args.input)
    mapping = build_mapping(names)
    write_mapping_csv(mapping, args.output)

    # Summary
    non_empty_roots = sum(1 for _, r in mapping if r)
    print(f"Processed {len(mapping)} names; {non_empty_roots} produced non-empty roots.")
    print(f"Wrote mapping to: {args.output}")


if __name__ == "__main__":
    main()

