# BTV detectie (hoe werkt dit script?)

Dit project bevat een simpel uitvoerbaar script: `detect_btv.py`.
Het doel daarvan is **per PDF een best-effort inschatting** te maken of het document (waarschijnlijk) een BTV-keuringsdocument is.

> Belangrijk: dit is heuristiek (regelgebaseerd) en dus geen 100% sluitend bewijs.

## Excel aanvullen (Lijst miv.xlsx)

Naast de detectie-mode kan `detect_btv.py` ook een Excel-bestand aanvullen via:

- `--update-excel <pad/naar/Lijst miv.xlsx>`
- `--pdf-dir <map/met/pdfs>`
- `--excel-out <output.xlsx>` (optioneel; default overschrijven)

### Wat wordt aangevuld?

Het script voegt (of vult) de volgende kolommen toe:

- `PDF-bestand`: bestandsnaam van de gemapte PDF
- `PDF inleesbaar`: `JA` als er genoeg tekst uit de PDF kan geëxtraheerd worden (geen pure scan), anders `NEE`
- `Keuringsdatum`: uit de PDF-tekst (indien inleesbaar)
- `Keuringsresultaat`: uit de PDF-tekst (indien inleesbaar)
- `Vastgestelde inbreuken`: uit de PDF-tekst (indien inleesbaar)

### PDF → Excel mapping (MIV)

Mapping gebeurt op basis van het MIV-nummer dat uit de PDF bestandsnaam gehaald wordt.

Voorbeelden:

- `MIV418 SAT1` in Excel → key `MIV418.1`
- `MIV041 SAT` in Excel → key `MIV041.1`
- `MIV041` in Excel → key `MIV041`

Wanneer er meerdere Excel-rijen dezelfde key hebben, kiest het script een “beste” rij (bv. `.1` PDF kiest liefst een `SAT`-rij).

### Keuring: datum/resultaat/inbreuken (anchors)

De extractie is regelgebaseerd en zoekt o.a. naar deze tekstankers:

- Datum:
  - `Datum(s) van controle: 18/06/2020`
- Resultaat (voorbeelden):
  - `DE INSTALLATIE IS NIET CONFORM`
  - `conform`
  - `conform met opmerkingen`
  - `niet conform`
  - of zinnen met `voldoet` in de buurt van `Besluit van de controle:`
- Inbreuken:
  - sectie `vastgestelde inbreuken` / `vastgestelde inbreuk`

> Tip: zet `--max-pages 2` of meer als de relevante secties (bv. inbreuken) pas op pagina 2+ staan.

---

## Overzicht

Input → (tekst + metadata) → regels scoren → confidence berekenen → vergelijken met threshold → output.

1. **PDF inlezen** met `pypdf.PdfReader`.
2. **Metadata lezen** (Producer/Creator/Title/etc. indien aanwezig).
3. **Tekst extracten** uit de eerste `--max-pages` pagina’s (default: 3).
4. **Regels toepassen** op tekst + metadata en punten optellen.
5. **Confidence** berekenen (0.0–1.0) en vergelijken met `--threshold`.
6. Resultaat printen als simpele tabel (`yes/no`) of als JSON (`--json`).

## 1) Tekstextractie

Het script gebruikt de tekstlaag in de PDF.

- Per pagina: `page.extract_text()`.
- Het scant maximaal `--max-pages` pagina’s (default 3) om snel te blijven.
- Alle pagina-tekst wordt samengevoegd tot één tekststring.

### Gevolgen
- PDF’s die enkel een scan zijn (geen tekstlaag) geven vaak **lege/arme tekst**, waardoor de score 0 blijft.
- Er zit **geen OCR** in dit project.

## 2) Detectieregels (token scoring)

Het script zoekt naar typische woorden/tokens die in BTV-documenten voorkomen.
Elke match geeft punten.

**Sterke signalen** (meer punten):
- `btv` → +3.0
- `keuringsverslag` → +2.0
- `keuringsattest` → +2.0

**Zwakker/ondersteunend**:
- `keuring` → +0.5
- `miv` → +1.0
- `inbreuk(en)` → +0.5
- `opmerking(en)` → +0.5

De matches worden opgeslagen in `matched_rules`, zodat je kan zien waarom iets scoorde.

## 3) Metadata scoring

Metadata is vaak leeg of onbetrouwbaar, daarom heeft dit een lage weging.

- Als in metadata-velden (bv. Producer/Creator/Title) het woord `btv` voorkomt: +1.0

## 4) Confidence en threshold

De totale score (tekst + metadata) wordt omgezet naar een **confidence** tussen 0 en 1:

- `confidence = min(raw_score / 5.0, 1.0)`

Daarna beslist het script:

- `is_btv = confidence >= threshold`
- default threshold is `0.7`.

### Praktisch voorbeeld
- Alleen `btv` gevonden: score 3.0 → confidence 0.6 → bij threshold 0.7: **no**
- `btv` + `keuringsverslag`: 3.0 + 2.0 = 5.0 → confidence 1.0 → **yes**

## 5) Output

Standaard (zonder `--json`) output per PDF:

- `yes|no` (is BTV?)
- confidence (2 decimalen)
- pad naar de PDF

Met `--json` krijg je een JSON array met o.a.:

- `is_btv`, `confidence`
- `matched_rules`
- `errors` (bv. corrupte PDF)
- `metadata` en `text_stats`

## Beperkingen en tips

- **False negatives**: scan-PDF zonder tekstlaag (geen OCR) → score 0.
- **False positives**: als een niet-BTV document toevallig dezelfde woorden bevat.
- Als je detectie “te streng” is: verlaag `--threshold` (bv. 0.5).
- Als je detectie “te los” is: verhoog `--threshold` of voeg extra regels toe (bv. unieke footer/URL/BTW-nummer als die terugkomt in jouw BTV templates).

## Waar de regels staan

De volledige implementatie staat in:
- `BTV MIV Keuringsverslagen WVB/detect_btv.py`

Zoek daar naar:
- `_TOKEN_RULES` (tekstregels)
- `_META_RULES` (metadata regels)
- `confidence_from_score()`
