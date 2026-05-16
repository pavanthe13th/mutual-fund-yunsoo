#!/usr/bin/env python3
"""Interview parser (parse-only)."""

import argparse
import logging
import json
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from sheet_processor import ExcelFileProcessor


def fix_fund_names(df: pd.DataFrame) -> pd.DataFrame:
    """Fix invalid fund names by looking up better names for the same fund_code."""
    if "fund_code" not in df.columns or "fund_name" not in df.columns:
        return df

    df = df.copy()
    lookup: Dict[str, str] = {}

    for _, row in df.iterrows():
        fund_code = row["fund_code"]
        fund_name = row["fund_name"]

        if pd.isna(fund_code) or pd.isna(fund_name):
            continue

        fund_code_str = str(fund_code).strip()
        fund_name_str = str(fund_name).strip()

        if not fund_name_str or fund_name_str == fund_code_str:
            continue

        if fund_code_str not in lookup or len(fund_name_str) > len(lookup[fund_code_str]):
            lookup[fund_code_str] = fund_name_str

    for idx, row in df.iterrows():
        fund_code = row["fund_code"]
        fund_name = row["fund_name"]

        if pd.isna(fund_code):
            continue

        fund_code_str = str(fund_code).strip()
        fund_name_str = str(fund_name).strip() if pd.notna(fund_name) else ""
        needs_fix = (
            fund_name_str == fund_code_str
            or not fund_name_str
            or fund_name_str.startswith("Portfolio as on")
        )
        if needs_fix and fund_code_str in lookup:
            df.at[idx, "fund_name"] = lookup[fund_code_str]

    return df


def clean_dataframe(df: pd.DataFrame, threshold: float = 0.0001) -> pd.DataFrame:
    """Normalize parser output data types and clean common invalid values."""
    import re

    df = df.copy()
    df = fix_fund_names(df)

    numeric_cols = ["percent_to_nav", "market_value", "quantity", "notional_value", "ytm", "ytc"]
    for col in numeric_cols:
        if col in df.columns and df[col].dtype == "object":
            df[col] = df[col].astype(str).str.replace(r"[$%]", "", regex=True).str.strip()
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in df.select_dtypes(include=[np.number]).columns:
        mask = df[col].notna() & (df[col].abs() < threshold)
        df.loc[mask, col] = 0.0

    if "isin" in df.columns:
        isin_pattern = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$")
        fixed_deposit_pattern = re.compile(r"^IN[A-Z0-9]{10}$")

        def validate_isin(val: Any) -> Optional[str]:
            if pd.isna(val):
                return None
            val_str = str(val).strip().upper()
            if isin_pattern.match(val_str):
                return val_str
            if len(val_str) == 12 and fixed_deposit_pattern.match(val_str):
                return val_str
            return None

        df["isin"] = df["isin"].apply(validate_isin)

    if "maturity_date" in df.columns:
        def format_maturity_date(val: Any) -> Any:
            if pd.isna(val):
                return None
            try:
                dt = pd.to_datetime(val)
                return dt.strftime("%Y-%m-%d")
            except (ValueError, TypeError):
                return val

        df["maturity_date"] = df["maturity_date"].apply(format_maturity_date)

    return df


def discover_excel_files(folder: Path) -> List[Path]:
    files = []
    patterns = ["*.xlsx", "*.xls", "*.xlsb", "*.XLSX", "*.XLS", "*.XLSB"]
    for pattern in patterns:
        files.extend(folder.glob(pattern))
    return sorted([f for f in files if not f.name.startswith("~$")], key=lambda p: p.name.lower())


def build_skip_config(global_config: Dict[str, Any], amc_config: Any) -> Dict[str, List[str]]:
    result = {
        "exact": list(global_config.get("exact", [])),
        "contains": list(global_config.get("contains", [])),
        "regex": list(global_config.get("regex", [])),
    }

    if not isinstance(amc_config, dict):
        return result

    if "skip_sheets" in amc_config:
        return amc_config["skip_sheets"]

    if "skip_sheets_add" in amc_config:
        add = amc_config["skip_sheets_add"]
        result["exact"].extend(add.get("exact", []))
        result["contains"].extend(add.get("contains", []))
        result["regex"].extend(add.get("regex", []))

    if "skip_sheets_remove" in amc_config:
        remove = amc_config["skip_sheets_remove"]
        result["exact"] = [x for x in result["exact"] if x.lower() not in {r.lower() for r in remove.get("exact", [])}]
        result["contains"] = [x for x in result["contains"] if x.lower() not in {r.lower() for r in remove.get("contains", [])}]
        result["regex"] = [x for x in result["regex"] if x not in set(remove.get("regex", []))]

    return result


def detect_metadata_for_file(file_path: Path, amc_config: Any, metadata_dir: Path) -> Optional[Path]:
    if isinstance(amc_config, str):
        return metadata_dir / amc_config

    if not isinstance(amc_config, dict):
        return None

    if "metadata_by_format" not in amc_config:
        return metadata_dir / amc_config.get("metadata", "")

    file_ext = file_path.suffix.lower()
    for fmt in amc_config["metadata_by_format"]:
        if "file_extension" in fmt and fmt["file_extension"].lower() == file_ext:
            return metadata_dir / fmt["metadata"]

    try:
        try:
            xl = pd.ExcelFile(file_path, engine="openpyxl")
        except (ValueError, ImportError, KeyError, OSError, zipfile.BadZipFile):
            try:
                xl = pd.ExcelFile(file_path, engine="xlrd")
            except Exception:
                xl = pd.ExcelFile(file_path, engine="calamine")

        for sheet_name in xl.sheet_names[:5]:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=50)
                for idx in range(len(df)):
                    row = df.iloc[idx].tolist()
                    for col_idx, val in enumerate(row):
                        if str(val).strip() == "ISIN":
                            for fmt in amc_config["metadata_by_format"]:
                                if fmt.get("isin_col") == col_idx:
                                    return metadata_dir / fmt["metadata"]
                            break
            except Exception:
                continue
    except Exception:
        pass

    return metadata_dir / amc_config["metadata_by_format"][0]["metadata"]


def write_parsed_output(records: List[Dict[str, Any]], output_file: Path) -> None:
    df = pd.DataFrame(records)
    df = clean_dataframe(df)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_output = df.drop(columns=["table_type", "as_on_date"], errors="ignore")
        df_output.to_excel(writer, sheet_name="All Data", index=False)

        if "table_type" in df.columns:
            for table_type in df["table_type"].unique():
                df_table = df[df["table_type"] == table_type].drop(
                    columns=["table_type", "as_on_date"], errors="ignore"
                ).dropna(axis=1, how="all")
                if "instrument_name" in df_table.columns:
                    df_table = df_table[df_table["instrument_name"].notna()]
                sheet_name = table_type.replace("_", " ").title()[:31]
                df_table.to_excel(writer, sheet_name=sheet_name, index=False)


def parse_one_file(excel_path: Path, metadata_path: Path, skip_sheets: Dict[str, Any], output_file: Path) -> None:
    records = ExcelFileProcessor(str(excel_path), str(metadata_path), skip_sheets=skip_sheets).process_all_sheets()
    if records:
        write_parsed_output(records, output_file)
    elif output_file.exists():
        output_file.unlink()


def main() -> int:
    parser = argparse.ArgumentParser(description="Interview parser (parse-only)")
    parser.add_argument("--amc", required=True, help="AMC name from config")
    parser.add_argument("--input-dir", type=Path, help="Input directory override")
    parser.add_argument("--metadata", type=Path, help="Metadata file override")
    parser.add_argument("--output-dir", type=Path, default=SCRIPT_DIR.parent / "4_output_interview", help="Generated output directory")
    parser.add_argument(
        "--config",
        type=Path,
        default=SCRIPT_DIR.parent / "Header_Verification_Historical" / "config.json",
        help="Config file path",
    )
    args = parser.parse_args()

    logging.getLogger().setLevel(logging.ERROR)

    if not args.config.exists():
        return 1

    with open(args.config) as f:
        config = json.load(f)

    if args.amc not in config.get("amcs", {}):
        return 1

    amc_config = config["amcs"][args.amc]
    metadata_dir = args.config.parent / config["metadata_dir"]

    if isinstance(amc_config, str):
        metadata_filename = amc_config
    elif isinstance(amc_config, dict) and "metadata_by_format" in amc_config:
        metadata_filename = amc_config["metadata_by_format"][0]["metadata"]
    elif isinstance(amc_config, dict):
        metadata_filename = amc_config.get("metadata", "")
    else:
        metadata_filename = str(amc_config)

    metadata_path = args.metadata or (metadata_dir / metadata_filename)
    if not metadata_path.exists():
        return 1

    input_folder_name = args.amc
    if isinstance(amc_config, dict) and "input_folder" in amc_config:
        input_folder_name = amc_config["input_folder"]
    input_dir = args.input_dir or (args.config.parent / config["input_dir"] / input_folder_name)
    if not input_dir.exists():
        return 1

    files = discover_excel_files(input_dir)
    if not files:
        return 1

    skip_sheets = build_skip_config(config.get("skip_sheets", {}), amc_config)
    output_amc_dir = args.output_dir / args.amc
    output_amc_dir.mkdir(parents=True, exist_ok=True)

    print(f"AMC: {args.amc}")
    print(f"Input: {input_dir} ({len(files)} files)")

    parse_failed = False

    total = len(files)
    for idx, file_path in enumerate(files, 1):
        print(f"\rParsing: {idx}/{total}  ", end="", flush=True)
        file_metadata_path = metadata_path
        if isinstance(amc_config, dict) and "metadata_by_format" in amc_config:
            detected = detect_metadata_for_file(file_path, amc_config, metadata_dir)
            if detected:
                file_metadata_path = detected
        output_file = output_amc_dir / f"{file_path.stem}_parsed.xlsx"
        try:
            parse_one_file(file_path, file_metadata_path, skip_sheets=skip_sheets, output_file=output_file)
        except Exception:
            parse_failed = True

    print(f"\rParsed: {total} files   ")

    return 1 if parse_failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
