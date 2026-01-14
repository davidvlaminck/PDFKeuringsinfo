# PDFKeuringsinfo

Tools om keuringsverslag-PDF's te analyseren.

## Run (simpel script)

Dit project gebruikt één script dat je rechtstreeks kan uitvoeren:

```bash
python3 "BTV MIV Keuringsverslagen WVB/detect_btv.py" "BTV MIV Keuringsverslagen WVB/pdfs/*.pdf"
```

JSON output:

```bash
python3 "BTV MIV Keuringsverslagen WVB/detect_btv.py" "BTV MIV Keuringsverslagen WVB/pdfs/*.pdf" --json
```

## Excel aanvullen (mapping + extractie)

Je kan de bron-Excel (`Lijst miv.xlsx`) automatisch aanvullen op basis van de PDF-bestandsnamen.
Het script:

- mapt PDF's naar een record in de Excel (op basis van MIV-nummer + suffix zoals `.1`)
- schrijft de PDF-bestandsnaam weg
- detecteert of de PDF **inleesbaar** is (tekstlaag aanwezig, dus geen pure scan)
- als de PDF inleesbaar is: probeert het ook velden uit de PDF te halen:
  - `Keuringsdatum`
  - `Keuringsresultaat` (conform / conform met opmerkingen / niet conform)
  - `Vastgestelde inbreuken`

Voorbeeld:

```bash
python3 "BTV MIV Keuringsverslagen WVB/detect_btv.py" \
  --update-excel "BTV MIV Keuringsverslagen WVB/Lijst miv.xlsx" \
  --pdf-dir "BTV MIV Keuringsverslagen WVB/pdfs" \
  --excel-out "BTV MIV Keuringsverslagen WVB/Lijst miv.filled.xlsx" \
  --max-pages 2
```

> Tip: gebruik `--max-pages 2` als je ook de inbreuken wil mee opnemen (die staan vaak op p2).

## PyCharm runner

Er staat een convenience runner in de project-root:

- `run_update_excel.py`

Die voert dezelfde Excel-update uit, maar is handig om rechtstreeks vanuit PyCharm te starten.

## Dependencies

Het script gebruikt `pypdf`.

Aanrader (met venv):

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -U pip
python -m pip install -r requirements.txt
```
