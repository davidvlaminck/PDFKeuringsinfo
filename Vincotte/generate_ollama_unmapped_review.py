from __future__ import annotations

import argparse
import json
import logging
import re
import time
import urllib.request
from pathlib import Path
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)

DEFAULT_INPUT_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")
DEFAULT_DETAIL_GLOB = "*.infractions-detail.csv"
DEFAULT_AREI_PATH = Path(
    "/home/davidlinux/PycharmProjects/PDFKeuringsinfo/Vincotte/AREI Codes_ F-codes en Vinçotte Gebruik (1).xlsx"
)


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


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _canonical_text(value: Any) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"[^a-z0-9à-ÿ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _extract_family_catalog(arei_xlsx: Path) -> dict[str, str]:
    if not arei_xlsx.exists():
        raise FileNotFoundError(f"AREI workbook niet gevonden: {arei_xlsx}")

    df = pd.read_excel(arei_xlsx)
    if df.empty:
        raise ValueError("AREI workbook is leeg")

    fcol = next(
        (
            c
            for c in df.columns
            if "f-code" in str(c).strip().lower() or str(c).strip().lower() == "fcode"
        ),
        df.columns[0],
    )
    ccol = next((c for c in df.columns if "context" in str(c).strip().lower()), df.columns[min(2, len(df.columns) - 1)])

    rows = (
        df[[fcol, ccol]]
        .dropna(subset=[fcol])
        .assign(**{fcol: lambda d: d[fcol].astype(str).str.strip()})
    )
    rows = rows[rows[fcol].str.match(r"^F\d+(?:-\d+)?$", na=False)]
    rows = rows.drop_duplicates(subset=[fcol])

    if rows.empty:
        raise ValueError("Geen F-code families gevonden in AREI workbook")

    catalog: dict[str, str] = {}
    for _, row in rows.iterrows():
        catalog[str(row[fcol]).strip()] = _normalize_text(row.get(ccol, ""))
    return catalog


def _extract_json_payload(response_text: str) -> dict[str, Any]:
    text = response_text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)

    match = re.search(r"\{.*}", text, flags=re.S)
    if match:
        text = match.group(0)

    try:
        return json.loads(text)
    except Exception:
        return {"family": "", "confidence": 0.0, "reason": f"parse_error: {response_text[:200]}"}


def _ollama_generate(*, model: str, prompt: str, timeout_seconds: int = 120) -> str:
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": 0.1},
    }
    req = urllib.request.Request(
        "http://127.0.0.1:11434/api/generate",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=timeout_seconds) as response:
        data = json.loads(response.read().decode("utf-8"))
    return str(data.get("response", ""))


def generate_ollama_unmapped_review(
    *,
    detail_csv: Path,
    arei_xlsx: Path,
    output_xlsx: Path,
    output_csv: Path,
    model: str,
    max_rows: int | None = None,
    sleep_seconds: float = 0.1,
) -> dict[str, int]:
    detail_df = pd.read_csv(detail_csv)
    if detail_df.empty:
        raise ValueError("Detailbestand is leeg")

    required = {"infraction_family", "infraction_code_raw", "infraction_text_raw", "asset_key", "filename"}
    missing = sorted(required - set(detail_df.columns))
    if missing:
        raise ValueError(f"Detailbestand mist vereiste kolommen: {missing}")

    unmapped_df = detail_df[detail_df["infraction_family"] == "UNMAPPED"].copy()
    if unmapped_df.empty:
        raise ValueError("Geen UNMAPPED rijen gevonden in detailbestand")

    grouped = (
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

    if max_rows is not None:
        grouped = grouped.head(max_rows)

    family_catalog = _extract_family_catalog(arei_xlsx)
    families = sorted(family_catalog.keys(), key=lambda x: (len(x), x))
    family_lines = "\n".join([f"- {fam}: {family_catalog[fam]}" for fam in families])

    prompt_version = "ollama_v1_strict_json_family_only"
    rows: list[dict[str, Any]] = []

    logger.info("Start Ollama suggesties voor %s UNMAPPED varianten", len(grouped))
    for idx, row in enumerate(grouped.itertuples(index=False), start=1):
        code = _normalize_text(row.infraction_code_raw)
        text = _normalize_text(row.sample_text)

        prompt = (
            "Je bent een strikte classifier voor AREI-inbreuken. "
            "Kies EXACT EEN familie uit de lijst en geef een confidence tussen 0 en 1. "
            "Gebruik geen extra categorieen. Antwoord ALLEEN als JSON object met keys family, confidence, reason.\n\n"
            f"Beschikbare families:\n{family_lines}\n\n"
            f"Input code: {code}\n"
            f"Input tekst: {text}\n\n"
            "JSON formaat exact: {\"family\":\"F?\",\"confidence\":0.0,\"reason\":\"...\"}"
        )

        raw_response = _ollama_generate(model=model, prompt=prompt)
        parsed = _extract_json_payload(raw_response)

        family = _normalize_text(parsed.get("family", ""))
        confidence_raw = parsed.get("confidence", 0.0)
        try:
            confidence = float(confidence_raw)
        except Exception:
            confidence = 0.0
        confidence = max(0.0, min(1.0, confidence))

        if family not in family_catalog:
            family = ""
            confidence = 0.0

        rows.append(
            {
                "infraction_code_raw": code,
                "infraction_text_raw": text,
                "occurrences": int(row.occurrences),
                "asset_count": int(row.asset_count),
                "sample_filename": _normalize_text(row.sample_filename),
                "ollama_family_suggestion": family,
                "ollama_confidence": round(confidence, 4),
                "ollama_reason": _normalize_text(parsed.get("reason", "")),
                "model": model,
                "prompt_version": prompt_version,
                "verified": 0,
                "verified_family": "",
                "verified_note": "",
            }
        )

        if idx % 20 == 0 or idx == len(grouped):
            logger.info("Verwerkt %s/%s", idx, len(grouped))
        if sleep_seconds > 0:
            time.sleep(sleep_seconds)

    out_df = pd.DataFrame(rows).sort_values(["occurrences", "asset_count"], ascending=[False, False])
    out_df.to_csv(output_csv, index=False)
    out_df.to_excel(output_xlsx, index=False)

    logger.info("Review CSV: %s", output_csv)
    logger.info("Review XLSX: %s", output_xlsx)

    return {
        "unmapped_variants": int(len(out_df)),
        "total_unmapped_occurrences": int(out_df["occurrences"].sum()),
        "suggested_non_empty": int((out_df["ollama_family_suggestion"] != "").sum()),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genereer reviewbestand met ALLEEN Ollama suggesties voor UNMAPPED inbreuken."
    )
    parser.add_argument(
        "--detail-csv",
        type=Path,
        default=None,
        help="Pad naar *.infractions-detail.csv. Default: eerste match in Keuringen-map.",
    )
    parser.add_argument(
        "--arei-xlsx",
        type=Path,
        default=DEFAULT_AREI_PATH,
        help="Pad naar AREI workbook (enkel voor bestaande familie-lijst).",
    )
    parser.add_argument(
        "--output-prefix",
        type=Path,
        default=None,
        help="Output prefix. Default: gebaseerd op detail-csv naam.",
    )
    parser.add_argument(
        "--model",
        type=str,
        default="qwen2.5-coder:7b",
        help="Ollama modelnaam.",
    )
    parser.add_argument(
        "--max-rows",
        type=int,
        default=None,
        help="Optioneel: beperk aantal UNMAPPED varianten (voor snelle test).",
    )
    parser.add_argument(
        "--sleep-seconds",
        type=float,
        default=0.1,
        help="Wachtijd tussen requests om Ollama niet te overbelasten.",
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
        else detail_csv.with_name(
            (detail_csv.stem.replace(".infractions-detail", "") or detail_csv.stem)
            + ".infractions-ollama-review"
        )
    )
    output_csv = Path(f"{output_prefix}.csv")
    output_xlsx = Path(f"{output_prefix}.xlsx")

    try:
        stats = generate_ollama_unmapped_review(
            detail_csv=detail_csv,
            arei_xlsx=args.arei_xlsx.resolve(),
            output_xlsx=output_xlsx,
            output_csv=output_csv,
            model=args.model,
            max_rows=args.max_rows,
            sleep_seconds=args.sleep_seconds,
        )
    except Exception as exc:
        logger.error("Generatie mislukt: %s", exc)
        return 2

    logger.info("Klaar. Stats: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())




