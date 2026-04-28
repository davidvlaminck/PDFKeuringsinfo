# Rebuild Spec (v2)

## Functionele doelen

1. PDF's groeperen op template en toewijzen aan template-mappen.
2. Per template parserregels definiëren.
3. Keuringsdatum, keuringsresultaat en inbreuken extraheren.
4. Mapping naar installaties/LSDeel behouden via Excel- en Arango-flow.
5. Voor API-update alleen patchen wanneer brondata recenter is.

## Niet-doelen (initieel)

- Geen brede refactor van alle legacy mappen.
- Geen wijzigingen buiten afgesproken scope zonder expliciete vraag.

## Kwaliteitscriteria

- Scripts rechtstreeks runbaar in PyCharm (`if __name__ == "__main__"`).
- Heldere logging per record (skip/fill/error).
- Documentatie naast scripts voor run-volgorde.

