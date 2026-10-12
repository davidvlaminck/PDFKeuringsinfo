# 📊 Vincotte Verification & Infraction Analysis Workflow

**Complete pipeline** to enrich Vincotte inspection data, verify PDF contents against Excel, and extract infraction families.

---

## 🎯 Quick Start (Complete Flow)

```bash
# Step 1: Enrich with Arango data
uv run python Vincotte/enrich_vincotte_assets.py

# Step 2: Verify PDF contents (verdict + date)
uv run python Vincotte/verify_pdf_contents_result_and_date.py

# Step 2b: Merge verification results back into enriched Excel
uv run python Vincotte/merge_pdf_verification_into_enriched.py

# Step 3: (One-time / on update) Rebuild the combined text-family master mapping
uv run python Vincotte/build_combined_text_family_mapping.py

# Step 4: Merge verified infractions into enriched Excel (summary pivot)
uv run python Vincotte/merge_infractions_to_enriched.py

# Step 5: Build the per-asset detail report from the combined master mapping
uv run python Vincotte/build_final_detaillijst.py
```

**Result**: `detaillijst-latest-keuring.xlsx` with per-asset infraction details and family pivot.

---

## 📋 Detailed Workflow

### **STEP 1️⃣: Enrich Excel with ArangoDB Asset Info**

**Purpose**: Add ArangoDB asset UUIDs to Vincotte inspection records

**Input**:
- `MyVinotte Min.Vl.Gemeenschap Reports list.xlsx`
  - Columns: Site ID, Label, Visit Date, Visiting place, Description, FILENAME, City, etc.

**Process**:
1. Parse installation codes from Visiting place / Description
2. Match against ArangoDB `onderdeel#Laagspanningsbord` assets
3. Smart matching: installation code + naampad first part (fallback: naampad only)
4. Supports: WO/WW, KO/KW, XT-sites, OZ (Zone30), MET, MIV variants, etc.

**Output**:
- `MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx`
- **New columns**: Arango UUID, Arango Naam, Arango Naampad, Uitkomst enrich script

**Command**:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/enrich_vincotte_assets.py
```

---

### **STEP 2️⃣: Verify PDF Contents (Verdict + Date Extraction)**

**Purpose**: Validate that PDF BESLUIT verdict and inspection date match enriched Excel

**Input**:
- `MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx`
- PDF files in `<Site ID>/` subdirectories

**Process**:

1. **Extract BESLUIT section** from PDF
2. **Classify verdict** into: RED | ORANGE | GREEN | GEEN_KEURING
   - RED: "voldoet niet aan", "niet in overeenstemming", "inbreuken aanwezig"
   - ORANGE: "voldoet aan" + remarks (Opmerkingen/Aanbevelingen)
   - GREEN: "voldoet aan" without remarks
   - GEEN_KEURING: summary docs, no inspection verdict
   
3. **Extract inspection date** from:
   - Interventiegegevens section (with date range support)
   - "Uitgevoerd op" section
   - BESLUIT section
   - Fallback: all document dates (if unique)
   
4. **Compare**:
   - PDF verdict vs Excel `Label` → `verdict_match` (Y/N)
   - PDF date vs Excel `Visit Date` → `date_match` (Y/N)
   
5. **Detect duplicates**: SHA256 hash of normalized PDF text

**Output**:
- `MyVinotte Min.Vl.Gemeenschap Reports list.enriched.pdf-contents-check.xlsx`
- Columns: verify_status, confidence, reason, besluit_excerpt, inspection_date_pdf, date_verify_status, overall_verify_status

**Command**:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/verify_pdf_contents_result_and_date.py \
  --input "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx"
```

**Step 2b command (merge-back into enriched.xlsx):**
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/merge_pdf_verification_into_enriched.py \
  --enriched "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx" \
  --verification "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.pdf-contents-check.xlsx"
```

New columns added to enriched output:
- `verified_verdict` (RED/ORANGE/GREEN/GEEN_KEURING from PDF)
- `verdict_match` (`Y`/`N` from `verify_status`)
- `verified_date` (inspection date from PDF)
- `verified_date_fallback` (fallback date from PDF range)
- `date_match` (`Y`/`N` from `date_verify_status`)
- `overall_verify_status` (combined match status)
- `verify_status_raw`, `date_verify_status_raw` (original status values)

---

### **STEP 3️⃣: Build / Refresh Combined Text-Family Master Mapping**

**Purpose**: Combine all infraction sources into a single text-first mapping file.
Run this whenever source files are updated (e.g. after manual review).

**Input** (3 sources, read-only):
| File | Role |
|------|------|
| `final_corrected_with_review.xlsx` | mapped_master — highest priority |
| `unmapped-ollama-review-template_verwerkt.xlsx` | unmapped_master — second priority |
| `MyVinotte...all-infractions-with-ollama-suggestions.xlsx` | suggestions — fallback |
| `MyVinotte...all-infractions-text-review.xlsx` | used to backfill `sample_filename` |

**Family priority within each row** (highest → lowest):
1. `corrected_family`
2. `verified_family`
3. `infraction_family`
4. `master_combined_families`
5. `ollama_family_suggestion`

**Source priority** (highest → lowest):
1. `mapped_master`
2. `unmapped_master`
3. `suggestions`

**Output**:
- `MyVinotte Min.Vl.Gemeenschap Reports list.combined-text-family-mapping.xlsx`
  - Sheet: `combined_text_mapping`
  - One row per unique normalised infraction text
  - Columns: `infraction_text_raw`, `infraction_code_raw`, `sample_filename`, `occurrences`, `family`, `verified`, `text_norm`, `family_source`, `sources_present`, `source_rows`

**Command**:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/build_combined_text_family_mapping.py
```

---

### **STEP 4️⃣: Map Infractions to Enriched Excel (summary pivot)**

**Purpose**: Add family-level infraction counts as pivot columns to the enriched Excel

**Input**:
- `MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx` (with verification results from Step 2)
- **COMBINED MASTER** (single source of truth):
  - `MyVinotte Min.Vl.Gemeenschap Reports list.combined-text-family-mapping.xlsx`

**Command**:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/merge_infractions_to_enriched.py \
  --enriched "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx"
```

---

### **STEP 5️⃣: Build Per-Asset Detail Report**

**Purpose**: Generate the final per-asset infraction detail list for the latest inspection of each asset.

**Input**:
- `MyVinotte...combined-text-family-mapping.xlsx` — single source of truth for text→family
- `MyVinotte...all-infractions-text-review.xlsx` — all extracted infraction rows

**Process**:
1. Load combined mapping (already resolved `family` column — no priority logic needed here)
2. Load all infractions and normalise text
3. Join on `text_norm`
4. Keep only the **most recent inspection** per asset
5. Aggregate per asset: infraction count, categories, full texts
6. Pivot per-category counts (one `cat_F##` column per family present)

**Output**:
- `MyVinotte...detaillijst-latest-keuring.xlsx` with sheets:
  - `combined_text_mapping` — the mapping used (for auditability)
  - `latest_asset_summary` — one row per asset, all family pivot columns
  - `latest_asset_infractions` — one row per infraction in the latest inspection
  - `summary` — high-level metrics
- `...detaillijst-latest-keuring.latest_asset_summary.csv`

**Command**:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/build_final_detaillijst.py
```

Optional arguments:
```
--combined-mapping  Path to combined-text-family-mapping.xlsx
--all-infractions   Path to all-infractions-text-review.xlsx
--output            Output Excel path
```

---

## 📁 Key Files

| File | Purpose |
|------|---------|
| `enrich_vincotte_assets.py` | Step 1: Add Arango UUIDs |
| `verify_pdf_contents_result_and_date.py` | Step 2: Verify PDF verdict + date |
| `merge_pdf_verification_into_enriched.py` | Step 2b: Merge verification back into enriched Excel |
| `merge_infractions_to_enriched.py` | Step 4: Add family pivot counts to enriched Excel |
| `build_combined_text_family_mapping.py` | Step 3: Rebuild combined master mapping from all sources |
| `build_final_detaillijst.py` | Step 5: Build per-asset detail report |
| `AREI Codes_ F-codes...xlsx` | F-code reference (F1-F40 ranges) |
| `arango_settings.sample.json` | ArangoDB connection template |

### Key Data Files

| File | Purpose |
|------|---------|
| `...combined-text-family-mapping.xlsx` | **MASTER** — single source of truth: text → family |
| `...all-infractions-text-review.xlsx` | All extracted infraction rows (input to detaillijst) |
| `...detaillijst-latest-keuring.xlsx` | Final output: per-asset detail report |
| `...enriched.xlsx` | Asset metadata + verification results |

---

## 📌 Installation Code Patterns (Supported by enrich_vincotte_assets.py)

```
WO/WW####         — e.g., WO0215, WW0001
KO/KW#### (±suffix) — KW0515, KW0515PS → KW0515 (normalized)
XT### (+ 3 letters) — XTAAL, XTCDE
OZ#### (Zone30)   — OZ0594
MET##### (Meteo)  — MET71001, MET 32007 (space support)
A/C/G#### or ###A/C/G#  — 013A4, C0518A
M IV variants     — MIV004 LVE, MIV004 SS1/SS2, MIV 110 AN LVE
```

---

## ⚙️ Setup

```bash
# Install dependencies
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv sync

# (Optional) Create local ArangoDB config
cp Vincotte/arango_settings.sample.json Vincotte/arango_settings.local.json
# Edit with your credentials
```

---

## 🔄 Data Flow Diagram

```
┌─────────────────────────────────────────────┐
│  MyVinotte...xlsx (raw inspection data)     │
│  Site ID, Label, Visit Date, Filename, ...  │
└──────────────┬──────────────────────────────┘
               │
        Step 1: ENRICH ↓
               │
┌──────────────▼──────────────────────────────┐
│ enriched.xlsx                               │
│ + Arango UUID, Naam, Naampad               │
└──────────────┬──────────────────────────────┘
               │
     Step 2: VERIFY PDF CONTENTS ↓
               │
   ┌───────────▼────────────┐
   │ pdf-contents-check.xlsx │──Step 2b──▶ enriched.xlsx updated
   │ verdict_match=Y/N       │             + verified_verdict
   │ date_match=Y/N          │             + verdict_match
   └─────────────────────────┘             + verified_date
                                           + date_match
                                           + overall_status

        Step 4: MERGE INFRACTIONS (summary pivot) ↓

┌────────────────────────────────────────────────────────────────────┐
│                    COMBINED MASTER MAPPING                         │
│  combined-text-family-mapping.xlsx  (text_norm → family)          │
│  Built once by build_combined_text_family_mapping.py (Step 3)     │
│  Sources: final_corrected_with_review +                            │
│           unmapped-ollama-review-template_verwerkt +               │
│           all-infractions-with-ollama-suggestions                  │
└─────────────────────────────┬──────────────────────────────────────┘
                              │
              ┌───────────────┴────────────────────────┐
              │                                        │
              ▼                                        ▼
  Step 4: merge_infractions_to_enriched          Step 5: build_final_detaillijst
  → enriched.xlsx                                → detaillijst-latest-keuring.xlsx
    + infractions_F1_count                         + latest_asset_summary
    + infractions_F2_count                         + latest_asset_infractions
    + ...                                          + cat_F1, cat_F2, ...
    + infractions_UNMAPPED_count
    + infractions_total
```

---

## 📊 Expected Output Summary

After completing all steps, the **detaillijst-latest-keuring.xlsx** contains:

| Sheet | Content |
|-------|---------|
| `combined_text_mapping` | The mapping used (text → family, for audit) |
| `latest_asset_summary` | One row per asset, family pivot columns (`cat_F##`) |
| `latest_asset_infractions` | One row per infraction in the latest inspection |
| `summary` | High-level counts (unique texts, assets, infraction rows) |

> The **combined master mapping** (`combined-text-family-mapping.xlsx`) is the single edit point.
> After updating it (or its source files), re-run Step 3 then Step 5 to regenerate the report.
