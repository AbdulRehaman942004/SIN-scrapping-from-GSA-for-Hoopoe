from pathlib import Path
import csv


def csv_to_txt(csv_path: Path, txt_path: Path) -> None:
    with csv_path.open("r", encoding="utf-8", newline="") as f_in:
        reader = list(csv.DictReader(f_in))

    # Compute max width of the original column for alignment
    max_original = 0
    for row in reader:
        original = (row.get("original") or "").strip()
        if len(original) > max_original:
            max_original = len(original)

    with txt_path.open("w", encoding="utf-8", newline="") as f_out:
        for row in reader:
            original = (row.get("original") or "").strip()
            root = (row.get("root") or "").strip()
            f_out.write(f"{original.ljust(max_original)}   ->   {root}\n")


def main() -> None:
    base = Path(__file__).parent
    csv_path = base / "original_to_root.csv"
    txt_path = base / "original_to_root.txt"

    if not csv_path.exists():
        print(f"Error: {csv_path} not found")
        return

    csv_to_txt(csv_path, txt_path)
    print(f"Wrote: {txt_path}")


if __name__ == "__main__":
    main()
