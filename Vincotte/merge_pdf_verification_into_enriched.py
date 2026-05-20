from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


def _load_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    raise ValueError(f"Unsupported verification file format: {path}")


def _derive_match_flag(status: object) -> str:
    token = str(status).strip().upper() if status is not None else ""
    if token == "MATCH":
        return "Y"
    if token == "MISMATCH":
        return "N"
    return ""


def merge_pdf_verification_into_enriched(
    *,
    enriched_path: Path,
    verification_path: Path,
    output_path: Path,
) -> tuple[int, int]:
    enriched_df = pd.read_excel(enriched_path)
    verification_df = _load_table(verification_path)

    required_cols = [
        "row_index",
        "label_from_pdf",
        "verify_status",
        "inspection_date_pdf",
        "inspection_date_pdf_fallback",
        "date_verify_status",
        "overall_verify_status",
    ]
    missing = [c for c in required_cols if c not in verification_df.columns]
    if missing:
        raise ValueError(f"Verification file is missing required columns: {missing}")

    verification_df = verification_df.copy()
    verification_df["row_index"] = pd.to_numeric(verification_df["row_index"], errors="coerce")
    verification_df = verification_df.dropna(subset=["row_index"])
    verification_df["row_index"] = verification_df["row_index"].astype(int)

    # Keep the latest row when duplicates exist for the same Excel row index.
    verification_df = verification_df.drop_duplicates(subset=["row_index"], keep="last")

    keep_cols = [
        "row_index",
        "label_from_pdf",
        "verify_status",
        "inspection_date_pdf",
        "inspection_date_pdf_fallback",
        "date_verify_status",
        "overall_verify_status",
    ]
    optional_cols = [
        "confidence",
        "reason",
        "date_confidence",
        "date_reason",
    ]
    for col in optional_cols:
        if col in verification_df.columns:
            keep_cols.append(col)

    verification_subset = verification_df[keep_cols].copy()
    verification_subset = verification_subset.rename(
        columns={
            "label_from_pdf": "verified_verdict",
            "inspection_date_pdf": "verified_date",
            "inspection_date_pdf_fallback": "verified_date_fallback",
            "verify_status": "verify_status_raw",
            "date_verify_status": "date_verify_status_raw",
            "confidence": "verify_confidence",
            "reason": "verify_reason",
        }
    )

    verification_subset["verdict_match"] = verification_subset["verify_status_raw"].map(_derive_match_flag)
    verification_subset["date_match"] = verification_subset["date_verify_status_raw"].map(_derive_match_flag)

    enriched_df = enriched_df.copy()
    enriched_df["_merge_row_index"] = enriched_df.index + 2

    merged_df = enriched_df.merge(
        verification_subset,
        how="left",
        left_on="_merge_row_index",
        right_on="row_index",
    )

    merged_df = merged_df.drop(columns=["_merge_row_index", "row_index"])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    merged_df.to_excel(output_path, index=False)

    csv_output = output_path.with_suffix(".csv")
    merged_df.to_csv(csv_output, index=False)

    matched_rows = int(merged_df["verified_verdict"].notna().sum()) if "verified_verdict" in merged_df.columns else 0
    return len(merged_df), matched_rows


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Merge PDF verification results back into an enriched Excel file."
    )
    parser.add_argument(
        "--enriched",
        required=True,
        type=Path,
        help="Path to enriched Excel input file.",
    )
    parser.add_argument(
        "--verification",
        required=True,
        type=Path,
        help="Path to verification output file (.xlsx or .csv from verify_pdf_contents_result_and_date.py).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output path for merged Excel. Default: <enriched>.with-pdf-verification.xlsx",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    enriched_path = args.enriched.resolve()
    verification_path = args.verification.resolve()

    if not enriched_path.exists():
        print(f"ERROR: Enriched file not found: {enriched_path}")
        return 1
    if not verification_path.exists():
        print(f"ERROR: Verification file not found: {verification_path}")
        return 1

    output_path = (
        args.output.resolve()
        if args.output is not None
        else enriched_path.with_name(f"{enriched_path.stem}.with-pdf-verification.xlsx")
    )

    try:
        total_rows, matched_rows = merge_pdf_verification_into_enriched(
            enriched_path=enriched_path,
            verification_path=verification_path,
            output_path=output_path,
        )
    except Exception as exc:
        print(f"ERROR: Merge failed: {exc}")
        return 2

    print(f"Merged file written: {output_path}")
    print(f"CSV written: {output_path.with_suffix('.csv')}")
    print(f"Rows in output: {total_rows}")
    print(f"Rows with verification data: {matched_rows}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

