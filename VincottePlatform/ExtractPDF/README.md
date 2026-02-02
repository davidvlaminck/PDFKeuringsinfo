ExtractPDF utilities

This folder contains tools to extract keuringsdata from PDF reports.

extract_pdf_reports.py
- Purpose: scan PDFs in ./pdfs and extract:
  - pdf_filename, pdf_readable, installation, mapping_diagnostics, keuringsdatum_raw, keuringsdatum_iso, resultaat_raw, resultaat_normalized, inbreuken
- Usage:
  python extract_pdf_reports.py --pdf-dir ./pdfs --out extraction_results.csv
- Dependencies: pdfplumber or PyPDF2 (install in the repo venv)

Spec
- See spec.md for parsing rules and mapping heuristics.


## Templates clustering & template metadata

Two helper scripts exist to group similar PDFs into template-like clusters and export per-cluster metadata used for building parsers.

Files
- `run_cluster.py` — cluster PDFs (text shingle Jaccard), create `templates/cluster_XXX/` folders and write `templates/clusters_mapping.csv`.
- `make_templates.py` — read `templates/clusters_mapping.csv`, compute first-page logo signatures (centers) and pHash, and write `templates/cluster_XXX/template.json`.

Quick start (terminal)
1. Activate venv and install deps if needed:

```bash
. .venv/bin/activate
pip install -r requirements.txt
```

2. Build clusters and write mapping:

```bash
cd VincottePlatform/ExtractPDF
python run_cluster.py --pdf-dir pdfs --out-dir templates --min-sim 0.25
```

3. Generate template metadata:

```bash
python make_templates.py --templates-dir templates --pdf-dir pdfs
```

Notes
- `--min-sim` controls clustering sensitivity (lower merges more).
- `make_templates.py` stores `logo_phash` only if `Pillow` and `ImageHash` are available; otherwise `logo_phash` will be null and the script uses the center-based signature as a fallback.
- If you want copies instead of symlinks, run `run_cluster.py` with `--no-symlinks`.

Next steps
- Generate a review CSV of sample PDFs per template for manual validation (I can add this script).
- After validation, implement per-template extraction rules to reliably parse `keuringsdatum`, `resultaat`, and `inbreuken`.
