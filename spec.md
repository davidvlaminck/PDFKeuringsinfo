# Intro
Dit project zal trachten pdf's van keuringsverslagen in te lezen en hieruit te halen of een keuring geslaagd is, wanneer, en welke inbreuken of opmerkingen er waren.

# Vereisten
Per aannemer zal er een andere map worden aangemaakt, aangezien deze onderling niet compatibel zijn. Per aannemer zal er ook een spec.md file zijn waarin beschreven staat hoe de pdf er uit ziet en hoe de tekst er moet worden uitgehaald.

## Status (BTV MIV Keuringsverslagen WVB)

Voor BTV is er een eerste implementatie die:

- PDF's mapt naar de Excel `Lijst miv.xlsx` op basis van MIV-code in bestandsnaam
- detecteert of PDF's inleesbaar zijn (tekstlaag aanwezig; geen OCR)
- extra velden uit de PDF-tekst probeert te halen en te bewaren in Excel:
  - keuringsdatum
  - keuringsresultaat
  - vastgestelde inbreuken

Zie `BTV MIV Keuringsverslagen WVB/BTV_spec.md` voor de concrete anchors/heuristieken.
