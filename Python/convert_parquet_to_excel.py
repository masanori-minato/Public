from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

import pandas as pd


DEFAULT_INPUT_DIR = Path.home() / "Downloads" / "Parquet"
DEFAULT_OUTPUT_DIR = DEFAULT_INPUT_DIR / "excel_out"
DEFAULT_SHEET_NAME = "data"
EXCEL_MAX_ROWS = 1_048_576


def ensure_parquet_engine() -> None:
    if importlib.util.find_spec("pyarrow") is None:
        raise RuntimeError(
            "pyarrow is not installed. Run `pip install pyarrow` first."
        )


def ensure_excel_engine() -> None:
    if importlib.util.find_spec("openpyxl") is None:
        raise RuntimeError(
            "openpyxl is not installed. Run `pip install openpyxl` first."
        )


def convert_parquet_to_excel(
    input_dir: Path,
    output_dir: Path,
    sheet_name: str,
    overwrite: bool,
) -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder was not found: {input_dir}")

    parquet_files = sorted(input_dir.glob("*.parquet")) + sorted(
        input_dir.glob("*.PARQUET")
    )

    if not parquet_files:
        print(f"No Parquet files found: {input_dir}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)

    success_count = 0
    failure_count = 0

    for parquet_path in parquet_files:
        excel_path = output_dir / f"{parquet_path.stem}.xlsx"

        if excel_path.exists() and not overwrite:
            print(f"SKIP: {excel_path.name} already exists")
            continue

        try:
            df = pd.read_parquet(parquet_path, engine="pyarrow")
            if len(df) > EXCEL_MAX_ROWS:
                raise ValueError(
                    f"Excel supports up to {EXCEL_MAX_ROWS:,} rows, "
                    f"but this file has {len(df):,} rows."
                )

            df.to_excel(
                excel_path,
                index=False,
                sheet_name=sheet_name,
                engine="openpyxl",
            )
            success_count += 1
            print(f"OK: {parquet_path.name} -> {excel_path.name} ({len(df):,} rows)")
        except Exception as error:
            failure_count += 1
            print(f"NG: {parquet_path.name} ({error})")

    print(
        f"Done: success {success_count} / failed {failure_count} / "
        f"total {len(parquet_files)}"
    )
    print(f"Output folder: {output_dir}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Parquet files in Downloads\\Parquet to Excel."
    )
    parser.add_argument(
        "-i",
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help=f"Parquet input folder. Default: {DEFAULT_INPUT_DIR}",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Excel output folder. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--sheet-name",
        default=DEFAULT_SHEET_NAME,
        help=f"Excel sheet name. Default: {DEFAULT_SHEET_NAME}",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing Excel files.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    ensure_parquet_engine()
    ensure_excel_engine()
    convert_parquet_to_excel(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        sheet_name=args.sheet_name,
        overwrite=args.overwrite,
    )


if __name__ == "__main__":
    main()
