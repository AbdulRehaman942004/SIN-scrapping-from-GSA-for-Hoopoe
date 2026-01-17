import argparse
from pathlib import Path
from typing import List

import pandas as pd


DEFAULT_COLUMN_NAME: str = "Manufacturer Long Name"


def load_unique_manufacturers(
    excel_path: Path, column_name: str = DEFAULT_COLUMN_NAME
) -> List[str]:
    """
    Load the specified Excel file, extract the manufacturer column, and return a
    sorted list of unique manufacturer names.

    The function trims surrounding whitespace and removes empty/null entries.
    """
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    # Read the Excel file; openpyxl handles .xlsx reliably
    dataframe = pd.read_excel(excel_path, engine="openpyxl")

    if column_name not in dataframe.columns:
        available = ", ".join([str(c) for c in dataframe.columns])
        raise KeyError(
            f"Column '{column_name}' not found. Available columns: {available}"
        )

    column_series = dataframe[column_name]

    # Normalize values: keep as strings, strip whitespace, drop null/empty
    # Use pandas StringDtype to preserve NaN semantics while applying .str ops
    column_series = column_series.astype("string").str.strip()

    # Remove nulls and empty strings
    column_series = column_series[column_series.notna() & (column_series != "")]

    # Drop duplicates and sort for deterministic output
    unique_names = sorted(column_series.drop_duplicates().tolist())

    return unique_names


def write_list_to_file(values: List[str], output_path: Path) -> None:
    """Write each value on its own line to the output file."""
    output_path.write_text("\n".join(values), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    project_dir = Path(__file__).parent

    default_excel = project_dir / "essendant-product-list (1).xlsx"
    default_output = project_dir / "unique_manufacturers.txt"

    parser = argparse.ArgumentParser(
        description=(
            "Extract unique manufacturer names from an Excel file column and "
            "print/save the list."
        )
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=default_excel,
        help=f"Path to the Excel file (default: {default_excel.name})",
    )
    parser.add_argument(
        "--column",
        type=str,
        default=DEFAULT_COLUMN_NAME,
        help=f"Column name to extract (default: '{DEFAULT_COLUMN_NAME}')",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=default_output,
        help=f"Output text file for unique names (default: {default_output.name})",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        unique_manufacturers = load_unique_manufacturers(args.excel, args.column)
    except Exception as exc:  # Only at the top level to show a helpful error
        print(f"Error: {exc}")
        return

    # Write to file and print to console
    write_list_to_file(unique_manufacturers, args.output)

    print(f"Found {len(unique_manufacturers)} unique manufacturers in '{args.column}'.")
    print(f"Saved list to: {args.output}")
    print("\nUnique manufacturers:")
    for name in unique_manufacturers:
        print(name)


if __name__ == "__main__":
    main()
