# AI Handoff Notes

Deze notities geven vaste context voor volgende AI-sessies.

## Rebuild-aanpak

- Werk direct op `main`, maar plaats nieuwe code eerst onder `rebuild/` of afgesproken doelmap.
- Legacy blijft beschikbaar voor referentie.
- Archive snapshot: `archive/main-20260428` (tag: `archive-main-2026-04-28`).

## Belangrijke domeinregels (samenvatting)

- Excel en PDF mapping gebeurt op installatienummer met edge-cases (suffixen, SAT-varianten).
- Resultaat keuring wordt gemapt naar toegelaten waarden van het API-selection veld.
- Records met "buiten gebruik" niet patchen als de waarde niet in de API-selectielijst voorkomt.
- Alleen updaten als datum in bron recenter is dan bestaande datum (of bestaande datum ontbreekt).

## Scope discipline

- Focus op `VincottePlatform/ExtractPDF` wanneer gevraagd.
- Vermijd onnodige wijzigingen in andere mappen.

## Praktische workflow

1. Kleine wijziging.
2. Lokale run met duidelijke output.
3. Korte documentatie-update in nabije README/spec.

