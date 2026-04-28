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

## Matchlogica (nu geïmplementeerd)

Per record wordt eerst `Visiting place` gebruikt en, als daar geen geldig installatienummer in zit, `Description`.

Een waarde is een geldig installatienummer als die matcht op:

- `^(WO|WW)\d{4}$`
- `^K[OW]\d{4}$`
- `^[ACG]\d{4}$`
- `^WAD\d{3}$`
- `^\d{3}[ACG]\d$`

Extra parsingregels:

- `KW0515PS` wordt behandeld als `KW0515`.
- `Installatie nr: ...` (met variaties in leestekens/spaties) wordt ondersteund.
- Hardcoded correctie: `0134A4` -> `013A4`.
- ASW-fallback: `<code>.ASW` wordt omgezet naar `<code>.GEB.LSDeel` voor naam-matching.
- MIV-varianten worden ondersteund:
  - `MIV004 LVE` -> module `MIV004`
  - `MIV 173 LVE` -> module `MIV173`
  - `MIV004 SS1` / `MIV004 SS2` -> module `MIV004.1` / `MIV004.2`
  - `MIV 110 AN LVE` -> module `MIV110` (provinciecode verwijderd)

Voor een **100% zekere** match moet exact 1 asset voldoen aan beide voorwaarden:

1. `AIMNaamObject_naam == "<installatienummer>.LSDeel"`
2. eerste deel van `naampad_parts` (fallback: eerste segment van `NaampadObject_naampad`) == `<installatienummer>`

## Outputkolommen

Het script voegt deze kolommen toe in de enriched Excel:

- `Arango UUID`
- `Arango Naam`
- `Arango Naampad`
- `Uitkomst enrich script`

`Uitkomst enrich script` bevat nu duidelijke tekstwaarden:

- `Match gevonden via Visiting place`
- `Match gevonden via Description`
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
