# PDFKeuringsinfo

Deze repo gebruikt `uv` als standaard voor Python dependencies en run-commando's.

## 1) Eenmalige setup

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv python install 3.12
uv sync
```

## 2) Extra dependency-set voor geo functies

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv sync --extra geo
```

## 3) Scripts runnen

Voorbeelden:

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python API/debug.py
```

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python API/OneDriveClient.py
```

## 4) Setup check

```bash
cd /home/davidlinux/PycharmProjects/PDFKeuringsinfo
uv run python utils/check_setup.py
```

## 5) Lokale config

- Gebruik `API/settings_sample.json` als basis voor je lokale `API/settings.json`.
- Zet geen credentials in git.

