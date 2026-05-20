# 📊 Vincotte Data Verification & Enrichment

**⚡️ START HERE**: See [**WORKFLOW.md**](WORKFLOW.md) for complete step-by-step instructions.

> **Key info**: Infraction data comes from pre-verified **MASTER files** (not extracted from PDFs), ensuring consistency and allowing corrections before integration.

## 📋 Workflow Summary (3 Steps)

| Step | Input | Process | Output |
|------|-------|---------|--------|
| **1️⃣ ENRICH** | Vincotte Excel + ArangoDB | Match install codes, add asset UUIDs | `enriched.xlsx` |
| **2️⃣ VERIFY + MERGE BACK** | `enriched.xlsx` + PDFs | Extract verdict + date, compare to Excel, merge report back | enriched + `verified_verdict`, `date_match`, `overall_verify_status` |
| **3️⃣ MERGE INFRACTIONS** | enriched + **MASTER files** | Load verified mappings, pivot F1-F40 counts | enriched + `infractions_F*_count` |

**MASTER Infraction Files** (FINAL, immutable):
- `final_corrected_with_review.xlsx` — all mapped infractions
- `unmapped-ollama-review-template_verwerkt.xlsx` — formerly unmapped, now mapped + verified

## 📁 This Folder

| File | Purpose | Status |
|------|---------|--------|
| `enrich_vincotte_assets.py` | Step 1: Add Arango UUIDs | ✅ Active |
| `verify_pdf_contents_result_and_date.py` | Step 2: PDF validation | ✅ Active |
| `merge_pdf_verification_into_enriched.py` | Step 2b: Merge verification report into enriched | ✅ Active |
| `merge_infractions_to_enriched.py` | Step 3: Merge infraction counts | 🚧 TODO: Implement |
| `AREI Codes_ F-codes...xlsx` | Mapping reference | 📚 Reference |
| `arango_settings.sample.json` | ArangoDB config | 📚 Template |
| ~~`extract_red_infractions.py`~~ | Extract from PDFs (OLD) | ⛔ Replaced by MASTER files |
| ~~`analyze_unmapped_infractions.py`~~ | Infraction analysis (OLD) | ⛔ Replaced by MASTER files |

## Aanvullende context

- De Excel bevat een code in de kolom `Label`.
- Kleurbetekenis van `Label`:
  - Rood = niet conform
  - Oranje = conform met opmerkingen
  - Groen = conform
- Er bestaan meerdere opmaakvarianten van keuringsdocumenten.
- Voorlopig zijn enkel keuringen sinds `2021-01-01` relevant.

## Matchlogica (nu geïmplementeerd)

Per record wordt eerst `Visiting place` gebruikt en, als daar geen geldig installatienummer in zit, `Description`.

Een waarde is een geldig installatienummer als die matcht op:

- `^(WO|WW)\d{4}$`
- `^K[OW]\d{4}$`
- `^XT[A-Z]{3}$`
- `^OZ\d{4}$`
- `^MET\d{5}$`
- `^[ACG]\d{4}$`
- `^WA[DO]\d{3}$`
- `^\d{3}[ACG]\d$`

Extra parsingregels:

- KW/KO-codes met 4 cijfers en suffix worden genormaliseerd naar de basiscode.
  Voorbeelden: `KW0515PS` -> `KW0515`, `KW0017BR` / `KW0017BR1` -> `KW0017`, `KW0557/TU` / `KW0557TU` -> `KW0557`.
- `MET 32007` wordt behandeld als `MET32007`.
- `Installatie nr: ...` (met variaties in leestekens/spaties) wordt ondersteund.
- Hardcoded correctie: `0134A4` -> `013A4`.
- Hardcoded correctie: `WO0206` -> `111A8.WO0206`.
- Hardcoded naampad-matches:
  - `KW0060TU` -> `WEV.Tunnel.KW0060/KW0060.HSCab/LSGEDEELTE`
  - `010A8` -> `K511.010A8/K511.010A8.K/K511.010A8.LsDeel`
  - `239A8` -> `K307.239A8/K307.239A8.K/K307.239A8.LSDeel`
  - `C0518A` of `C0518 A` -> `C0518/KAST_A/KAST_A.LSDeel`
  - `C0518B` of `C0518 B` -> `C0518/KAST_B/KAST_B.LSDeel`
- ASW-fallback: `<code>.ASW` wordt omgezet naar `<code>.GEB.LSDeel` voor naam-matching.
- MIV-varianten worden ondersteund:
  - `MIV004 LVE` -> module `MIV004`
  - `MIV 173 LVE` -> module `MIV173`
  - `MIV004 SS1` / `MIV004 SS2` -> module `MIV004.1` / `MIV004.2`
  - `MIV 110 AN LVE` -> module `MIV110` (provinciecode verwijderd)
- XT-sites worden ondersteund:
  - patroon `XT` + 3 letters (bv. `XTAAL`)
  - match via naampad op uniek laagspanningsbord (bv. `XTAAL/XTAAL/AAL.LSDeel`)
- Zone30-sites worden ondersteund:
  - patroon `OZ` + 4 cijfers (bv. `OZ0594`)
  - match via tweede deel van het naampad (bv. `N9X45.35/OZ0594/N9P45.35.K/N9P45.35.K.LSDeel`)
- Meteosites worden ondersteund:
  - patroon `MET` + 5 cijfers (bv. `MET71001` of `MET 32007`)
  - match op uniek laagspanningsbord (bv. `MET71001/MET71001_Neerpelt/MET71001.K/MET71001.LSDeel`)

Voor een **100% zekere** match moet exact 1 asset voldoen aan beide voorwaarden:

1. `AIMNaamObject_naam == "<installatienummer>.LSDeel"`
2. eerste deel van `naampad_parts` (fallback: eerste segment van `NaampadObject_naampad`) == `<installatienummer>`

Als die strikte naamvoorwaarde niet lukt, is er nu ook een minder zekere fallback:

- unieke match enkel op eerste deel van `naampad_parts` == `<installatienummer>`.

## Outputkolommen

Het script voegt deze kolommen toe in de enriched Excel:

- `Arango UUID`
- `Arango Naam`
- `Arango Naampad`
- `Uitkomst enrich script`

`Uitkomst enrich script` bevat nu duidelijke tekstwaarden:

- `Match gevonden via Visiting place`
- `Match gevonden via Description`
- `Match gevonden via Visiting place/Description, enkel via naampad`
- `Match via OZ-zone30 naampad via Visiting place/Description`
- `Match via XT-site naampad via Visiting place/Description`
- `Match via MIV in boomstructuur via Visiting place/Description`
- `Match gevonden via ... na manuele correctie`
- `Geen geldig installatienummer in Visiting place of Description`
- `Geen match gevonden via Visiting place/Description`
- `Meerdere matches op naam via Visiting place/Description; geen unieke asset`

Onbekende provinciecodes in MIV-tekst worden gelogd als warning zodat je die kan valideren.

## Gebruik met uv
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv sync
```
Maak eerst een lokale settings file (niet committen), bijvoorbeeld:
```bash
cp Vincotte/arango_settings.sample.json Vincotte/arango_settings.local.json
```
Run daarna:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/enrich_vincotte_assets.py \
  --input "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.xlsx" \
  --settings /home/davidlinux/Documenten/AWV/resources/settings_SyncToArangoDB.json \
  --env prd \
  --installation-col E
```
Optioneel output pad:
```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/enrich_vincotte_assets.py \
  --input "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.xlsx" \
  --output "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx" \
  --settings /home/davidlinux/Documenten/AWV/resources/settings_SyncToArangoDB.json
```

Standaard run (zonder argumenten):

- input: `/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.xlsx`
- settings: `/home/davidlinux/Documenten/AWV/resources/settings_SyncToArangoDB.json`
- output: zelfde map als input met suffix `.enriched.xlsx`

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/enrich_vincotte_assets.py
```

## Verificatie Label vs PDF-besluit

Gebruik `verify_enriched_labels_from_pdf.py` om het resultaat in kolom `Label` te controleren tegen de besluittekst in het PDF-keuringsdocument.

Aannames in de mapstructuur:

- de enriched Excel staat in de `Keuringen`-map,
- per record staat de submap in kolom `Site ID` (kolom A),
- de PDF-bestandsnaam staat in kolom `FILENAME` (kolom L),
- enkel records met `Visit date > 2021-01-01` worden meegenomen (standaard).

Labelmapping:

- `Red`/`Rood` -> niet conform
- `Orange`/`Oranje` -> conform met opmerkingen
- `Green`/`Groen` -> conform

Extra beslissingsregel in PDF-parser:

- `Vastgestelde inbreuken: - Nihil` wordt beschouwd als conform-signaal.
- Als er tegelijk ook opmerkingen/aanbevelingen gedetecteerd worden, wordt dit `Orange`.
- Zonder opmerkingen wordt dit `Green`.
- Als `BESLUIT` conform is maar er elders in het document (bv. sectie `Opmerkingen`) opmerkingen staan,
  wordt dit ook als `Orange` geclassificeerd.
- Teksten zoals `niet teruggevonden`, `niet gevonden`, `niet worden gecontroleerd` of
  `niet gecontroleerd` worden als `Red` (niet conform) geclassificeerd.
- Teksten zoals `geen besluit genomen` of `geen controle uitgevoerd` worden ook als
  `Red` (niet conform) geclassificeerd.
- Teksten zoals `geen controle mogelijk` of `controle niet mogelijk` worden ook als
  `Red` (niet conform) geclassificeerd.
- Teksten zoals `niet bereikbaar` (bv. "niet bereikbaar door werken") worden ook als
  `Red` (niet conform) geclassificeerd.
- Tekst zoals `geen kast aanwezig` (bv. "Er was geen kast aanwezig op de aangegeven locatie")
  wordt ook als `Red` (niet conform) geclassificeerd.
- Teksten zoals `geen keuring uitgevoerd` worden als `Orange` geclassificeerd wanneer er
  effectief opmerkingen aanwezig zijn; zonder opmerkingen vallen ze terug op `Red`.
- Teksten zoals `nog niets om te keuren`, `niet toegankelijk`, `bord leeg`, `geen bord gevonden`,
  `kon niet uitgevoerd worden` en `geen ingevulde gegevens` worden als `Red` geclassificeerd.
- Tekst zoals `cabine is reeds afgebroken` (in opmerkingen) wordt als
  `Orange` (conform met opmerkingen) geclassificeerd.
- `Opmerkingen: Nihil` (ook in gecodeerde PDF-tekstextractie) telt als **geen** opmerkingen,
  dus dit blijft `Green` tenzij er elders effectieve opmerkingen gevonden worden.
- Als zulke niet-conform-signalen enkel buiten `BESLUIT` staan (bv. in `Opmerkingen`),
  worden ze ook meegenomen als `Red`.
- `documentatie AREI` en gelijkaardige documenten zonder keuringsbesluit worden als
  `GEEN_KEURING` geclassificeerd.
- In gecodeerde PDF-varianten geldt: `Vastgestelde inbreuken` zonder `Nihil` wordt als
  `Red` geïnterpreteerd.

Datumverificatie:

- De parser zoekt een inspectiedatum in de PDF (o.a. `datum van controle`, `keuringsdatum`,
  `inspectiedatum`, `... uitgevoerd op ...`).
- Gevonden inspectiedatum wordt vergeleken met `Visit date` uit Excel.
- Uitkomst per rij staat in `date_verify_status` (`MATCH`, `MISMATCH`, `UNKNOWN`).
- Voor oude expliciete datumvelden met twee-cijferig jaar (`DD/MM/YY`) wordt het jaar
  indien nodig bijgestuurd naar het jaar van `Visit date` wanneer dag/maand gelijk zijn en
  het verschil precies 1 jaar is.
- Als meerdere verschillende datums in de PDF staan zonder duidelijke inspectiedatum,
  wordt de status `UNKNOWN` met reden in `date_reason`.
- `overall_verify_status` combineert labelcontrole + datumcontrole in 1 veld.
- **Checkbox-template** (HS-verslagvorm): sommige documenten tonen in het BESLUIT altijd
  beide opties (`VOLDOET AAN` én `VOLDOET NIET AAN`) als aanvinkvakjes.  Omdat tekst-
  extractie niet ziet welk vakje aangevinkt is, wordt de aanwezigheid van `VOLDOET NIET AAN`
  genegeerd als er ook `VOLDOET AAN` staat.  In plaats daarvan wordt de sectie
  `INBREUKEN / OPMERKINGEN / NOTA'S` (vóór BESLUIT) gebruikt:
  - `Nihil` → `Green`
  - Inbreuken aanwezig (geen Nihil) → `Red`
- Documenten zonder keuringsbesluit (bv. `Synthesedocument ... risicoanalyse`,
  samenvatting/documentatie) worden geclassificeerd als `GEEN_KEURING`.
- Records met `Geen leesbare tekst in PDF` worden ook als `GEEN_KEURING` behandeld,
  omdat dit in de huidige dataset doorgaans geen keuringsverslagen zijn.

Run:

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/verify_enriched_labels_from_pdf.py \
  --input "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx"
```

Optioneel:

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/verify_enriched_labels_from_pdf.py \
  --input "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx" \
  --output "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.pdf-label-check.xlsx" \
  --since-date 2021-01-01
```

Output:

- Excel-rapport met verificatiestatus per rij (`MATCH`, `MISMATCH`, `UNKNOWN`, `PDF_MISSING`, `GEEN_KEURING`)
- CSV met dezelfde resultaten (zelfde bestandsnaam, extensie `.csv`)
- Extra kolommen voor inhoudelijke duplicaten: `pdf_content_hash`, `content_duplicate`,
  `content_duplicate_count` en `duplicate_filenames`.
- Extra datumkolommen: `inspection_date_pdf`, `date_verify_status`, `date_confidence`, `date_reason`.
- Gecombineerde statuskolom: `overall_verify_status`.

## Inbreuken extractie (RED-labels)

Gebruik `extract_red_infractions.py` om per **asset** enkel de **meest recente** keuring te nemen, daaruit de inbreuken uit de PDF te extraheren, en een provinciale draaitabel te maken.

Wat het script doet:

- filter op `Visit Date > --since-date` (standaard `2021-01-01`)
- houdt per asset alleen de meest recente keuring (`Arango UUID` als primaire assetsleutel)
- neemt enkel records met label `Red`/`Rood`
- extraheert inbreuken uit de sectie `III. INBREUKEN`
- mapt codes naar families met `AREI Codes_ F-codes ...xlsx`
- schrijft 3 outputs:
  - `<prefix>.infractions-detail.(csv|xlsx)`
  - `<prefix>.infractions-per-asset.(csv|xlsx)`
  - `<prefix>.infractions-pivot-province.(csv|xlsx)`

Run:

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/extract_red_infractions.py \
  --input "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx"
```

Optioneel:

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/extract_red_infractions.py \
  --input "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx" \
  --output-prefix "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.red-infractions" \
  --arei-codes "/home/davidlinux/PycharmProjects/PDFKeuringsinfo/Vincotte/AREI Codes_ F-codes en Vinçotte Gebruik (1).xlsx" \
  --since-date 2021-01-01
```

> Opmerking: provincie wordt afgeleid op basis van postcode in `City`, met fallback op provincie-tekst in `Description`. Niet-herkende gevallen krijgen `UNKNOWN`.

## Analyse van UNMAPPED inbreuken

Gebruik `analyze_unmapped_infractions.py` op het gegenereerde detailbestand om alle niet-gemapte inbreuken te groeperen en per variant gelijkaardige **wel** gemapte stukken te tonen.

Outputs:

- `<prefix>.infractions-unmapped-summary.(csv|xlsx)`
  - per unieke `UNMAPPED` code + tekstvariant
  - met top gelijkaardige gemapte kandidaten
- `<prefix>.infractions-unmapped-codes.(csv|xlsx)`
  - samenvatting per rauwe code
- `<prefix>.infractions-unmapped-rows.(csv|xlsx)`
  - rij-per-rij `UNMAPPED` detail met voorgestelde match

Run:

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/analyze_unmapped_infractions.py
```

Optioneel:

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python Vincotte/analyze_unmapped_infractions.py \
  --detail-csv "/home/davidlinux/Documenten/AWV/Keuringen/MyVinotte Min.Vl.Gemeenschap Reports list.enriched.infractions-detail.csv" \
  --top-n 5 \
  --min-score 0.18
```
