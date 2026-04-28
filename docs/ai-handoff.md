# AI Handoff Notes

Deze notities geven vaste context voor volgende AI-sessies.

## Generiek

- De code zal door andere partijen (mogelijks ook AI) gereviewd worden 
- De code moet best practices uit de softwareontwikkeling toepassen 
- De code moet clean zijn en duidelijk leesbaar voor een mens 
- De code moet uitvoerbaar en testbaar zijn 

## Rebuild-aanpak

- Werk direct op `main`
- Legacy blijft beschikbaar voor referentie.
- Archive snapshot: `archive/main-20260428` (tag: `archive-main-2026-04-28`).

## Scope discipline

- Vermijd onnodige wijzigingen in andere mappen.

## Omgeving

- Gebruik `uv` als standaard voor environment, dependency sync en script uitvoering.
- Vermijd `pip`/`venv` commando's tenzij expliciet gevraagd.
- Gebruik bij voorkeur `uv run ...` voor alle run-commando's.
