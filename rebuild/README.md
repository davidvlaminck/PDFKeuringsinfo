# Rebuild Workspace (on `main`)

Deze map is de schone werkzone voor de nieuwe codebasis, terwijl legacy code in de bestaande mappen blijft staan.

## Doel

- Nieuwe implementatie opbouwen zonder eerst legacy code te verplaatsen of te verwijderen.
- Iteratief functionaliteit overnemen vanuit de archive-snapshot.
- Duidelijke scheiding tussen legacy en nieuwe code.

## Werkwijze

1. Voeg nieuwe scripts/modules eerst hier toe.
2. Reuse van legacy logica gebeurt expliciet (copy/refactor per feature).
3. Houd elke stap klein en verifieerbaar.

## Scope voor eerste iteraties

- PDF template clustering en template metadata.
- Template-gebaseerde extractie (datum, resultaat, inbreuken).
- Mapping naar Excel en API-update flows.

