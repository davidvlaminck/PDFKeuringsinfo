# Vincotte mapping scaffold
Deze map bevat de voorbereiding om een Vincotte Excel aan te vullen met asset UUIDs uit ArangoDB.
## Wat staat er klaar
- `enrich_vincotte_assets.py`: uitvoerbaar script met `if __name__ == "__main__"`.
- `arango_settings.sample.json`: voorbeeldconfig voor ArangoDB connectie.
## Huidige status

- Assettype staat op `onderdeel#Laagspanningsbord`.
- Matching gebeurt nu met strikte regels op `AIMNaamObject_naam` en het eerste deel van `naampad_parts`.
- Het script overschrijft de bron-Excel nooit; output gaat altijd naar een apart `.enriched.xlsx` bestand.
- De kolommen `Arango UUID`, `Arango Naam` en `Arango Naampad` worden aangevuld bij unieke match.

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
