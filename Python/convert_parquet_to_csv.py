"""
===============================================================================
 Parquetファイル一括CSV変換プログラム
===============================================================================

指定フォルダー内のParquetファイルを読み込み、ファイルごとにCSV形式へ
変換します。出力するCSVの文字コードはオプションで指定できます。

デフォルトの入出力先:
    入力: ~/Downloads/Parquet
    出力: ~/Downloads/Parquet/csv_out

主な機能:
    - 複数のParquetファイルを一括変換
    - CSVの出力文字コードを指定可能
    - 既存ファイルのスキップまたは上書き

実行例:
    python convert_parquet_to_csv.py
    python convert_parquet_to_csv.py --encoding cp932 --overwrite
    python convert_parquet_to_csv.py -i "C:\\input" -o "C:\\output"

必要なライブラリ:
    pandas、pyarrow
===============================================================================
"""

from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

import pandas as pd


DEFAULT_INPUT_DIR = Path.home() / "Downloads" / "Parquet"
DEFAULT_OUTPUT_DIR = DEFAULT_INPUT_DIR / "csv_out"


def ensure_parquet_engine() -> None:
    if importlib.util.find_spec("pyarrow") is None:
        raise RuntimeError(
            "pyarrow is not installed. Run `pip install pyarrow` first."
        )


def convert_parquet_to_csv(
    input_dir: Path,
    output_dir: Path,
    encoding: str,
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
        csv_path = output_dir / f"{parquet_path.stem}.csv"

        if csv_path.exists() and not overwrite:
            print(f"SKIP: {csv_path.name} already exists")
            continue

        try:
            df = pd.read_parquet(parquet_path, engine="pyarrow")
            df.to_csv(csv_path, index=False, encoding=encoding)
            success_count += 1
            print(f"OK: {parquet_path.name} -> {csv_path.name} ({len(df):,} rows)")
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
        description="Convert Parquet files in Downloads\\Parquet to CSV."
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
        help=f"CSV output folder. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8-sig",
        help="CSV output encoding. Default: utf-8-sig",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing CSV files.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    ensure_parquet_engine()
    convert_parquet_to_csv(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        encoding=args.encoding,
        overwrite=args.overwrite,
    )


if __name__ == "__main__":
    main()
