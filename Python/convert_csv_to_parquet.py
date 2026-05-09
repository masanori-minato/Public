from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

import pandas as pd


DEFAULT_INPUT_DIR = Path.home() / "Downloads" / "CSV"
DEFAULT_OUTPUT_DIR = DEFAULT_INPUT_DIR / "parquet_out"
DEFAULT_ENCODINGS = ("utf-8-sig", "utf-8", "cp932")


def ensure_parquet_engine() -> None:
    if importlib.util.find_spec("pyarrow") is None:
        raise RuntimeError(
            "pyarrow is not installed. Run `pip install pyarrow` first."
        )


def read_csv_with_fallback(
    csv_path: Path,
    encodings: tuple[str, ...],
    all_string: bool,
) -> pd.DataFrame:
    """Try common encodings used in Japanese Windows CSV files."""
    last_error: Exception | None = None
    dtype = str if all_string else None

    for encoding in encodings:
        try:
            return pd.read_csv(csv_path, encoding=encoding, dtype=dtype)
        except UnicodeDecodeError as error:
            last_error = error

    raise RuntimeError(
        f"Could not read CSV. Tried encodings: {', '.join(encodings)}"
    ) from last_error


def convert_csv_to_parquet(
    input_dir: Path,
    output_dir: Path,
    encodings: tuple[str, ...],
    all_string: bool,
    overwrite: bool,
) -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder was not found: {input_dir}")

    csv_files = sorted(input_dir.glob("*.csv")) + sorted(input_dir.glob("*.CSV"))

    if not csv_files:
        print(f"No CSV files found: {input_dir}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)

    success_count = 0
    failure_count = 0

    for csv_path in csv_files:
        parquet_path = output_dir / f"{csv_path.stem}.parquet"

        if parquet_path.exists() and not overwrite:
            print(f"SKIP: {parquet_path.name} already exists")
            continue

        try:
            df = read_csv_with_fallback(csv_path, encodings, all_string)
            df.to_parquet(parquet_path, index=False, engine="pyarrow")
            success_count += 1
            print(f"OK: {csv_path.name} -> {parquet_path.name} ({len(df):,} rows)")
        except Exception as error:
            failure_count += 1
            print(f"NG: {csv_path.name} ({error})")

    print(
        f"Done: success {success_count} / failed {failure_count} / "
        f"total {len(csv_files)}"
    )
    print(f"Output folder: {output_dir}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert CSV files in Downloads\\CSV to Parquet."
    )
    parser.add_argument(
        "-i",
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help=f"CSV input folder. Default: {DEFAULT_INPUT_DIR}",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Parquet output folder. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--encoding",
        action="append",
        dest="encodings",
        help=(
            "CSV encoding to try. Can be specified multiple times. "
            "Default: utf-8-sig, utf-8, cp932"
        ),
    )
    parser.add_argument(
        "--all-string",
        action="store_true",
        help="Read all columns as strings. Useful for preserving leading zeros.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing Parquet files.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    encodings = tuple(args.encodings) if args.encodings else DEFAULT_ENCODINGS

    ensure_parquet_engine()
    convert_csv_to_parquet(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        encodings=encodings,
        all_string=args.all_string,
        overwrite=args.overwrite,
    )


if __name__ == "__main__":
    main()
