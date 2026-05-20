from __future__ import annotations

import argparse
import logging
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)

DEFAULT_INPUT_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")
DEFAULT_DETAIL_GLOB = "*.infractions-detail.csv"

_STOPWORDS = {
    "aan",
    "aangebracht",
    "aangesloten",
    "alle",
    "basis",
    "bij",
    "boek",
    "code",
    "codex",
    "de",
    "deze",
    "die",
    "dient",
    "dienen",
    "dit",
    "door",
    "een",
    "en",
    "enkel",
    "geplaatst",
    "het",
    "hiervan",
    "in",
    "installatie",
    "installaties",
    "is",
    "konden",
    "met",
    "moet",
    "moeten",
    "naar",
    "niet",
    "of",
    "om",
    "ons",
    "op",
    "te",
    "titel",
    "uit",
    "uw",
    "van",
    "voorzien",
    "voorgelegd",
    "volgens",
    "werden",
    "werd",
    "worden",
    "wordt",
    "zijn",
}


@dataclass(frozen=True)
class SimilarityCandidate:
    score: float
    jaccard: float
    sequence: float
    family: str
    raw_code: str
    text: str
    count: int


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _canonical_text(value: Any) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"[^a-z0-9à-ÿ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _tokens(value: Any) -> list[str]:
    tokens = []
    for token in _canonical_text(value).split():
        if len(token) <= 2:
            continue
        if token in _STOPWORDS:
            continue
        if token.isdigit():
            continue
        tokens.append(token)
    return tokens


def _similarity(a: str, b: str) -> tuple[float, float, float]:
    ta = set(_tokens(a))
    tb = set(_tokens(b))
    jaccard = len(ta & tb) / len(ta | tb) if (ta or tb) else 0.0
    sequence = SequenceMatcher(None, _canonical_text(a), _canonical_text(b)).ratio()
    score = 0.65 * jaccard + 0.35 * sequence
    return score, jaccard, sequence


def _find_detail_input(path_arg: Path | None) -> Path:
    if path_arg is not None:
        path = path_arg.resolve()
        if not path.exists():
            raise FileNotFoundError(f"Detailbestand bestaat niet: {path}")
        return path

    matches = sorted(DEFAULT_INPUT_DIR.glob(DEFAULT_DETAIL_GLOB))
    if not matches:
        raise FileNotFoundError(
            f"Geen detailbestand gevonden in {DEFAULT_INPUT_DIR} met patroon {DEFAULT_DETAIL_GLOB}"
        )
    return matches[0].resolve()


def analyze_unmapped_infractions(
    *,
    detail_csv: Path,
    output_prefix: Path,
    top_n: int = 3,
    min_score: float = 0.18,
) -> dict[str, int]:
    detail_df = pd.read_csv(detail_csv)
    if detail_df.empty:
        raise ValueError("Detailbestand is leeg")

    unmapped_df = detail_df[detail_df["infraction_family"] == "UNMAPPED"].copy()
    mapped_df = detail_df[detail_df["infraction_family"] != "UNMAPPED"].copy()

    if unmapped_df.empty:
        raise ValueError("Geen UNMAPPED inbreuken gevonden in detailbestand")
    if mapped_df.empty:
        raise ValueError("Geen gemapte inbreuken gevonden om mee te vergelijken")

    mapped_groups = (
        mapped_df.groupby(["infraction_family", "infraction_code_raw", "infraction_text_raw"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values(["count", "infraction_family", "infraction_code_raw"], ascending=[False, True, True])
    )

    summary_groups = (
        unmapped_df.assign(_canonical_text=unmapped_df["infraction_text_raw"].map(_canonical_text))
        .groupby(["infraction_code_raw", "_canonical_text"], dropna=False)
        .agg(
            occurrences=("infraction_code_raw", "size"),
            asset_count=("asset_key", "nunique"),
            sample_text=("infraction_text_raw", "first"),
            sample_filename=("filename", "first"),
        )
        .reset_index()
        .sort_values(["occurrences", "asset_count", "infraction_code_raw"], ascending=[False, False, True])
    )

    summary_rows: list[dict[str, Any]] = []
    for _, row in summary_groups.iterrows():
        best: list[SimilarityCandidate] = []
        source_text = _normalize_text(row["sample_text"])

        for _, mapped in mapped_groups.iterrows():
            score, jaccard, sequence = _similarity(source_text, _normalize_text(mapped["infraction_text_raw"]))
            if score < min_score:
                continue
            best.append(
                SimilarityCandidate(
                    score=score,
                    jaccard=jaccard,
                    sequence=sequence,
                    family=_normalize_text(mapped["infraction_family"]),
                    raw_code=_normalize_text(mapped["infraction_code_raw"]),
                    text=_normalize_text(mapped["infraction_text_raw"]),
                    count=int(mapped["count"]),
                )
            )

        best.sort(key=lambda item: (item.score, item.jaccard, item.count), reverse=True)
        top = best[:top_n]

        record: dict[str, Any] = {
            "infraction_code_raw": _normalize_text(row["infraction_code_raw"]),
            "occurrences": int(row["occurrences"]),
            "asset_count": int(row["asset_count"]),
            "sample_text": source_text,
            "sample_filename": _normalize_text(row["sample_filename"]),
            "suggested_family": top[0].family if top else "",
            "suggested_raw_code": top[0].raw_code if top else "",
            "suggested_score": round(top[0].score, 4) if top else "",
            "suggested_text": top[0].text if top else "",
        }
        for i, candidate in enumerate(top, start=1):
            record[f"candidate_{i}_family"] = candidate.family
            record[f"candidate_{i}_raw_code"] = candidate.raw_code
            record[f"candidate_{i}_score"] = round(candidate.score, 4)
            record[f"candidate_{i}_jaccard"] = round(candidate.jaccard, 4)
            record[f"candidate_{i}_sequence"] = round(candidate.sequence, 4)
            record[f"candidate_{i}_count"] = candidate.count
            record[f"candidate_{i}_text"] = candidate.text
        summary_rows.append(record)

    summary_df = pd.DataFrame(summary_rows)

    summary_df["suggested_score"] = pd.to_numeric(summary_df["suggested_score"], errors="coerce")
    summary_df["weighted_suggestion_score"] = summary_df["occurrences"] * summary_df["suggested_score"].fillna(0)

    code_summary_rows: list[dict[str, Any]] = []
    for raw_code, group in summary_df.groupby("infraction_code_raw", dropna=False):
        group = group.sort_values(
            ["weighted_suggestion_score", "occurrences", "suggested_score"],
            ascending=[False, False, False],
        )
        best = group.iloc[0]
        code_summary_rows.append(
            {
                "infraction_code_raw": raw_code,
                "total_occurrences": int(group["occurrences"].sum()),
                "text_variants": int(group["sample_text"].nunique()),
                "sample_text": _normalize_text(best["sample_text"]),
                "suggested_family": _normalize_text(best.get("suggested_family")),
                "suggested_raw_code": _normalize_text(best.get("suggested_raw_code")),
                "suggested_score": best.get("suggested_score"),
                "weighted_suggestion_score": best.get("weighted_suggestion_score"),
            }
        )

    code_summary_df = pd.DataFrame(code_summary_rows).sort_values(
        ["total_occurrences", "infraction_code_raw"], ascending=[False, True]
    )

    detail_unmapped_df = unmapped_df.merge(
        summary_df[
            [
                "infraction_code_raw",
                "sample_text",
                "suggested_family",
                "suggested_raw_code",
                "suggested_score",
                "suggested_text",
            ]
        ],
        how="left",
        left_on=["infraction_code_raw", "infraction_text_raw"],
        right_on=["infraction_code_raw", "sample_text"],
    ).drop(columns=["sample_text"])

    summary_csv = output_prefix.with_suffix(".infractions-unmapped-summary.csv")
    summary_xlsx = output_prefix.with_suffix(".infractions-unmapped-summary.xlsx")
    code_csv = output_prefix.with_suffix(".infractions-unmapped-codes.csv")
    code_xlsx = output_prefix.with_suffix(".infractions-unmapped-codes.xlsx")
    row_csv = output_prefix.with_suffix(".infractions-unmapped-rows.csv")
    row_xlsx = output_prefix.with_suffix(".infractions-unmapped-rows.xlsx")

    summary_df.to_csv(summary_csv, index=False)
    summary_df.to_excel(summary_xlsx, index=False)
    code_summary_df.to_csv(code_csv, index=False)
    code_summary_df.to_excel(code_xlsx, index=False)
    detail_unmapped_df.to_csv(row_csv, index=False)
    detail_unmapped_df.to_excel(row_xlsx, index=False)

    logger.info("UNMAPPED summary CSV: %s", summary_csv)
    logger.info("UNMAPPED summary XLSX: %s", summary_xlsx)
    logger.info("UNMAPPED code CSV: %s", code_csv)
    logger.info("UNMAPPED code XLSX: %s", code_xlsx)
    logger.info("UNMAPPED rows CSV: %s", row_csv)
    logger.info("UNMAPPED rows XLSX: %s", row_xlsx)

    return {
        "unmapped_rows": int(len(unmapped_df)),
        "unmapped_code_count": int(unmapped_df["infraction_code_raw"].nunique()),
        "unmapped_text_variants": int(len(summary_df)),
    }


def analyze_mapped_infractions(
    *,
    detail_csv: Path,
    output_prefix: Path,
) -> dict[str, int]:
    """Analyseer de gemapte inbreuken per familie en code."""
    detail_df = pd.read_csv(detail_csv)
    if detail_df.empty:
        raise ValueError("Detailbestand is leeg")

    mapped_df = detail_df[detail_df["infraction_family"] != "UNMAPPED"].copy()

    if mapped_df.empty:
        raise ValueError("Geen gemapte inbreuken gevonden in detailbestand")

    # Overzicht per famiglia + code + tekst
    family_groups = (
        mapped_df.groupby(["infraction_family", "infraction_code_raw", "infraction_text_raw"], dropna=False)
        .agg(
            occurrences=("infraction_code_raw", "size"),
            asset_count=("asset_key", "nunique"),
            province_count=("province", "nunique"),
            sample_filename=("filename", "first"),
        )
        .reset_index()
        .sort_values(["occurrences", "asset_count"], ascending=[False, False])
    )

    # Samenvatting per familie + code (aggregatie over textvarianten)
    code_groups = (
        mapped_df.groupby(["infraction_family", "infraction_code_raw"], dropna=False)
        .agg(
            total_occurrences=("infraction_code_raw", "size"),
            text_variants=("infraction_text_raw", "nunique"),
            asset_count=("asset_key", "nunique"),
            province_count=("province", "nunique"),
            sample_text=("infraction_text_raw", "first"),
            sample_filename=("filename", "first"),
        )
        .reset_index()
        .sort_values(["total_occurrences", "asset_count"], ascending=[False, False])
    )

    # Samenvatting per familie
    family_summary = (
        mapped_df.groupby("infraction_family", dropna=False)
        .agg(
            total_occurrences=("infraction_code_raw", "size"),
            code_count=("infraction_code_raw", "nunique"),
            text_variants=("infraction_text_raw", "nunique"),
            asset_count=("asset_key", "nunique"),
            province_count=("province", "nunique"),
        )
        .reset_index()
        .sort_values(["total_occurrences", "asset_count"], ascending=[False, False])
    )

    family_csv = output_prefix.with_suffix(".infractions-mapped-family.csv")
    family_xlsx = output_prefix.with_suffix(".infractions-mapped-family.xlsx")
    code_csv = output_prefix.with_suffix(".infractions-mapped-codes.csv")
    code_xlsx = output_prefix.with_suffix(".infractions-mapped-codes.xlsx")
    detail_csv_out = output_prefix.with_suffix(".infractions-mapped-detail.csv")
    detail_xlsx_out = output_prefix.with_suffix(".infractions-mapped-detail.xlsx")

    family_summary.to_csv(family_csv, index=False)
    family_summary.to_excel(family_xlsx, index=False)
    code_groups.to_csv(code_csv, index=False)
    code_groups.to_excel(code_xlsx, index=False)
    family_groups.to_csv(detail_csv_out, index=False)
    family_groups.to_excel(detail_xlsx_out, index=False)

    logger.info("Mapped family summary CSV: %s", family_csv)
    logger.info("Mapped family summary XLSX: %s", family_xlsx)
    logger.info("Mapped codes CSV: %s", code_csv)
    logger.info("Mapped codes XLSX: %s", code_xlsx)
    logger.info("Mapped detail CSV: %s", detail_csv_out)
    logger.info("Mapped detail XLSX: %s", detail_xlsx_out)

    return {
        "mapped_rows": int(len(mapped_df)),
        "mapped_code_count": int(mapped_df["infraction_code_raw"].nunique()),
        "mapped_family_count": int(mapped_df["infraction_family"].nunique()),
        "mapped_text_variants": int(len(family_groups)),
    }


def parse_args() -> argparse.Namespace:
     parser = argparse.ArgumentParser(
         description="Analyseer UNMAPPED of MAPPED inbreuken."
     )
     parser.add_argument(
         "--detail-csv",
         type=Path,
         default=None,
         help="Pad naar *.infractions-detail.csv. Default: eerste match in Keuringen-map.",
     )
     parser.add_argument(
         "--output-prefix",
         type=Path,
         default=None,
         help="Output prefix. Default: gebaseerd op detail-csv pad.",
     )
     parser.add_argument(
         "--top-n",
         type=int,
         default=3,
         help="Aantal gelijkaardige gemapte kandidaten per UNMAPPED-variant.",
     )
     parser.add_argument(
         "--min-score",
         type=float,
         default=0.18,
         help="Minimale similariteitsscore om kandidaat op te nemen.",
     )
     parser.add_argument(
         "--analyze-mapped",
         action="store_true",
         help="Analyseer gemapte inbreuken in plaats van unmapped.",
     )
     parser.add_argument(
         "--analyze-both",
         action="store_true",
         help="Analyseer zowel unmapped als gemapte inbreuken.",
     )
     return parser.parse_args()


def main() -> int:
     logging.basicConfig(
         level=logging.INFO,
         format="%(asctime)s %(levelname)-8s %(message)s",
         datefmt="%Y-%m-%d %H:%M:%S",
     )

     args = parse_args()
     try:
         detail_csv = _find_detail_input(args.detail_csv)
     except Exception as exc:
         logger.error("Input zoeken mislukt: %s", exc)
         return 1

     output_prefix = (
         args.output_prefix.resolve()
         if args.output_prefix is not None
         else detail_csv.with_name(detail_csv.stem.replace(".infractions-detail", "") or detail_csv.stem)
     )

     all_stats: dict[str, dict[str, int]] = {}
     
     # Bepaal welke analyses uit te voeren
     do_unmapped = not args.analyze_mapped or args.analyze_both
     do_mapped = args.analyze_mapped or args.analyze_both
     
     if do_unmapped:
         try:
             logger.info("UNMAPPED inbreuken analyseren...")
             stats = analyze_unmapped_infractions(
                 detail_csv=detail_csv,
                 output_prefix=output_prefix,
                 top_n=args.top_n,
                 min_score=args.min_score,
             )
             all_stats["unmapped"] = stats
         except Exception as exc:
             logger.error("UNMAPPED analyse mislukt: %s", exc)
             return 2
     
     if do_mapped:
         try:
             logger.info("Gemapte inbreuken analyseren...")
             stats = analyze_mapped_infractions(
                 detail_csv=detail_csv,
                 output_prefix=output_prefix,
             )
             all_stats["mapped"] = stats
         except Exception as exc:
             logger.error("Gemapte analyse mislukt: %s", exc)
             return 3

     logger.info("Klaar. Stats: %s", all_stats)
     return 0


if __name__ == "__main__":
    raise SystemExit(main())



