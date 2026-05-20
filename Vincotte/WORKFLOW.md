# 📊 Vincotte Verification & Infraction Analysis Workflow

**Complete pipeline** to enrich Vincotte inspection data, verify PDF contents against Excel, and extract infraction families.

---

## 🎯 Quick Start (Complete Flow)

```bash
# Step 1: Enrich with Arango data
uv run python Vincotte/enrich_vincotte_assets.py

# Step 2: Verify PDF contents (verdict + date)
uv run python Vincotte/verify_pdf_contents_result_and_date.py

# Step 3: Merge verified infractions (MASTER files) into enriched Excel
uv run python Vincotte/merge_infractions_to_enriched.py
```

**Result**: `enriched.xlsx` with all asset info, verification results, and infraction counts.

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

**➡️ TODO: Merge verification results back into enriched.xlsx**
```
New columns to add:
- verified_verdict        (RED/ORANGE/GREEN/GEEN_KEURING from PDF)
- verdict_match           (Y/N if matches enriched Label)
- verified_date           (inspection_date from PDF)
- date_match              (Y/N if matches enriched Visit Date)
- overall_verify_status   (combined match status)
```

---

### **STEP 3️⃣: Map Infractions to Enriched Excel**

**Purpose**: Load verified infraction mappings (MASTER files) and add family counts to enriched Excel

**Input**:
- `MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx` (with verification results from Step 2)
- **MASTER files** (finalized, read-only):
  - `unmapped-ollama-review-template_verwerkt.xlsx` (formerly unmapped codes, now mapped via Ollama + user review)
  - `final_corrected_with_review.xlsx` (all mapped infractions, verified)

**These MASTER files are FINAL** — DO NOT edit. They contain all infraction mappings with:
- infraction_code_raw
- infraction_text_raw
- infraction_family (F1-F40 or UNMAPPED)
- All user reviews and corrections baked in

**Process**:

1. **Load both MASTER files**:
   - Combine `unmapped-ollama-review-template_verwerkt` + `final_corrected_with_review`
   - Deduplicate by (asset, code, family) if overlap exists
   
2. **Map infractions to PDFs**:
   - Create mapping: infraction_code_raw → infraction_family
   - For each PDF in enriched.xlsx, use its filename to find infractions
   
3. **Pivot family counts per asset**:
   - Group by `Arango UUID` (asset identifier)
   - Count occurrences per family (F1-F40)
   - Include `infractions_total` and `infractions_UNMAPPED_count`
   
4. **Add to enriched.xlsx**:
   - Join on `Arango UUID`
   - Add columns: `infractions_F1_count`, `infractions_F2_count`, ..., `infractions_F40_count`, `infractions_UNMAPPED_count`, `infractions_total`
   - Zero if no infractions for that asset/family

**Output**:
- Updated `enriched.xlsx` with infraction family pivot columns
- Side effect: `infractions-mapped-to-assets.csv|xlsx` (detail: asset, filename, code, family, text)

**Command**:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/merge_infractions_to_enriched.py \
  --enriched "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx" \
  --mapped-master "/home/davidlinux/Documenten/AWV/Keuringen/final_corrected_with_review.xlsx" \
  --unmapped-master "/home/davidlinux/Documenten/AWV/Keuringen/unmapped-ollama-review-template_verwerkt.xlsx"
```

**➡️ TODO (to implement)**: Create `merge_infractions_to_enriched.py` script
```python
# Logic:
# 1. Load two MASTER files
# 2. Filter verified rows (verified=1) from both
# 3. Extract: infraction_code_raw, infraction_family, filename
# 4. Match filename to enriched.xlsx records
# 5. For each Arango UUID: pivot_table(count by infraction_family)
# 6. Add columns to enriched.xlsx
```

This enables rapid analysis: how many assets have infractions in F5 (electrical safety), etc.

---

---

### **STEP 4️⃣ (Archive): Extract Infractions Script (NO LONGER USED)**

⚠️ **DEPRECATED**: Historical script `extract_red_infractions.py` extracted infractions from PDFs directly.

Now we use **MASTER files** instead:
- `final_corrected_with_review.xlsx` (verified mapped infractions)
- `unmapped-ollama-review-template_verwerkt.xlsx` (verified unmapped → mapped)

These are FINAL and immutable. See Step 3 for merging into enriched Excel.

---

## 📁 Key Files

| File | Purpose |
|------|---------|
| `enrich_vincotte_assets.py` | Step 1: Add Arango UUIDs |
| `verify_pdf_contents_result_and_date.py` | Step 2: Verify PDF verdict + date |
| `extract_red_infractions.py` | Step 3: Extract infraction codes |
| `analyze_unmapped_infractions.py` | Step 4: Quality analysis |
| `AREI Codes_ F-codes...xlsx` | Mapping reference (F1-F40 ranges) |
| `arango_settings.sample.json` | ArangoDB connection template |

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
┌─────────────────────────────────────────┐
│  MyVinotte...xlsx (raw inspection data) │
│  + Site ID, Label, Visit Date, Filename │
└──────────────┬──────────────────────────┘
               │
        Step 1: ENRICH ↓
               │
┌──────────────▼──────────────────────────┐
│ enriched.xlsx                           │
│ + Arango UUID (asset identifier)        │
│ + Arango Naam, Arango Naampad           │
└──────────────┬──────────────────────────┘
               │
     Step 2: VERIFY PDF CONTENTS ↓
               │                    (TODO: merge back)
    ┌──────────▼──────────┐
    │ pdf-contents-check  │─────────────┐
    │ verdict_match=Y/N   │             │
    │ date_match=Y/N      │             │
    └─────────────────────┘             ▼
               │         ┌──────────────────────┐
               │         │ enriched.xlsx updated│
               │         │ + verified_verdict   │
               │         │ + verdict_match      │
               │         │ + verified_date      │
               │         │ + date_match         │
               │         │ + overall_status     │
               │         └──────────┬───────────┘
               │                    │
         Step 3: MERGE INFRACTIONS (MASTER files) ↓
               │                    │
    ┌──────────┴────────────────────┼──────────────────────────┐
    │                                │                          │
    │    MASTER FILES (FINAL)        │                          │
    │    ┌────────────────────────────┐                          │
    │    │ final_corrected_with_review│                          │
    │    │ (all mapped infrastructures)                          │
    │    └────────┬───────────────────┘                          │
    │             │                                              │
    │    ┌────────▼─────────────────────────────────┐            │
    │    │ unmapped-ollama-...._verwerkt            │            │
    │    │ (formerly unmapped, now mapped + review) │            │
    │    └────────┬─────────────────────────────────┘            │
    │             │                                              │
    │             └──────────────┬──────────────────┬────────────┘
    │                            │                  │
    │  1. Load MASTER files      │                  │
    │  2. Combine (deduplicate)  │                  │
    │  3. Match filename→         │                  │
    │     infraction_code→        │                  │
    │     infraction_family       │                  │
    │  4. Pivot per asset        │                  │
    │                            ▼                  ▼
    └────────────────────────────────────────────────────┐
                                    │
                                    ↓
                    ┌──────────────────────────────┐
                    │ enriched.xlsx FINAL          │
                    │ + infractions_F1_count       │
                    │ + infractions_F2_count       │
                    │ + ... F40_count              │
                    │ + infractions_UNMAPPED_count │
                    │ + infractions_total          │
                    └──────────────────────────────┘
```

---

## 📊 Expected Output Summary

After completing all steps, the **final enriched.xlsx** contains:

| Column | Source | Purpose |
|--------|--------|---------|
| Site ID, Label, Visit Date, FILENAME, City | Original | Inspection metadata |
| Arango UUID, Arango Naam, Arango Naampad | Step 1 | Asset identification |
| verified_verdict, verdict_match | Step 2 | PDF verdict validation |
| verified_date, date_match, overall_verify_status | Step 2 | PDF date validation |
| infractions_F1_count ... infractions_F40_count | Step 3 | Infraction family distributions |
| infractions_UNMAPPED_count, infractions_total | Step 3 | Unmapped + total counts |

**This single Excel file** now contains everything needed for compliance analysis:
- Which assets are verified (verdict + date matches)
- Which assets have infractions in which families (F1-F40)
- Ready for pivoting by province, asset type, infraction family, etc.

> All enrichment → verification → infraction merging is **cumulative** — each step's output feeds into the next.







