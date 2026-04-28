# Archive Procedure for `main`

Dit document beschrijft hoe de huidige codebasis van `main` gearchiveerd wordt voor referentie tijdens een clean rebuild.

## Doel

- Een volledige, herbruikbare snapshot bewaren van de huidige `main`.
- Verder ontwikkelen op een schone rebuild-branch zonder historiek te verliezen.
- Oude implementaties beschikbaar houden voor vergelijking en cherry-picks.

## Branch conventie

Gebruik een datumgebaseerde branchnaam:

- `archive/main-YYYYMMDD`

Voor deze snapshot:

- `archive/main-20260428`

## Stappen

```bash
git checkout main
git pull --ff-only
git checkout -b archive/main-20260428
```

Optioneel (aanbevolen) tag op de archiefsnapshot:

```bash
git tag archive-main-2026-04-28
git push origin archive/main-20260428
git push origin archive-main-2026-04-28
```

## Verder werken

Na archiveren kan je clean starten op een nieuwe branch, bv.:

```bash
git checkout main
git checkout -b rebuild/v2-clean-start
```

## Herstellen of vergelijken

- Bekijk oud bestand: `git show archive/main-20260428:path/to/file.py`
- Cherry-pick commit: `git cherry-pick <commit>`
- Vergelijk branches: `git diff archive/main-20260428..rebuild/v2-clean-start`

