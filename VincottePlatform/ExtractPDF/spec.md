# Doel

Dit document beschrijft het doel en de aanpak voor het automatisch uitlezen van keuringverslagen (PDF) in de map `VincottePlatform/ExtractPDF`.
Het uiteindelijke resultaat is een verrijkte Excel (of meerdere Excels) met per PDF: bestandsnaam, of het document tekstueel uitleesbaar is (dus geen scan), de keuringsdatum, het resultaat van de keuring en een eventuele lijst met vastgestelde inbreuken.

Samengevat:
- Classificeer PDF's per template / formaat.
- Detecteer of een PDF tekst bevat of een scan is (OCR nodig).
- Haal uit tekstuele PDF's (of na OCR) de keuringsdatum, het resultaat en eventuele inbreuken.
- Koppel PDFs aan installatie-nummers volgens bestaande mapping-regels en vul een Excel aan.
- Maak een overzicht van niet-gemapte records en foutgevallen.

Belangrijke velden om te extraheren
- Datum van de keuring: zoek naar lijnen zoals
  "Datum(s) van controle:18/06/2020" of met spaties: "Datum(s) van controle: 18/06/2020".
  Formaat: DD/MM/YYYY (standaard); normaliseer naar ISO YYYY-MM-DD in de output.

- Resultaat van de keuring: tekstblok dat begint met of bevat "Besluit van de controle:" gevolgd door een conclusie. Mogelijke gerelateerde waarden (normaliseer):
  - "conform"
  - "conform met opmerkingen"
  - "niet-conform"
  - "niet-conform met inbreuken"
  - andere (bewaar originele tekst)

- Vastgestelde inbreuken: wanneer het rapport "NIET CONFORM" is, staat er vaak een paragraaf of lijst met "vastgestelde inbreuken" of "vastgestelde inbreuken, die ..." of "vastgestelde inbreuken: ...". Verzamel deze tekstblokken en zet ze als komma- of puntgescheiden string in een aparte kolom.

Mapping- en bestandsnaamregels
- Bestandsnamen bevatten meestal het installatie-nummer: voorbeelden: "MIV004.pdf", "Keuring MIV004.1.pdf", "MIV418 SAT1.pdf".
- Trim alle naamvelden (gebruik `.strip()` op strings).
- Wanneer een naam iets als "MIV418 SAT1" bevat, normalizeer naar installatie-nummer "MIV418.1" (regel: SAT of SAT1 -> suffix `.1`).
- Wanneer er een spatie in de installatie-naam staat, probeer:
  1) full stripped name
  2) split op spatie en neem het eerste token (en strip dat)
  3) controleer of dat token een geldige installatie is (zie validatieregels hieronder)
  Als dat token valide is, gebruik het.
- Specifieke mapping: als er twee records bestaan: "MIV041" en "MIV041 SAT" — map "MIV041" naar MIV041, en "MIV041 SAT" naar MIV041.1.
- Als kolom G (Keuringsattest beschikbaar?) in de Excel 'NEE' bevat, mag je deze record in principe negeren (geen match verwacht). Alleen wanneer er per toeval toch een match wordt gevonden kun je deze invullen, anders negeren.

Validatie van installatie-nummers
- Er zijn meerdere geldige vormen. Gebruik regex-matching na strip():
  1) Hoofdvorm (veel gebruikt): (A|C|G|WO|WW) gevolgd door 4 cijfers. Voorbeeld: A1234, WO1234
     Regex (conceptueel): r'^(?:A|C|G|WO|WW)\s*\d{4}$'  (sta ook spaties toe tussen prefix en cijfers)

  2) Alternatieve vorm gevonden in data: 3 cijfers gevolgd door A, C of G gevolgd door 1 cijfer. Er kunnen spaties voor en na de letter staan.
     Voorbeelden: "123A4", "123 A 4", "123 A4".
     Regex (conceptueel): r'^\d{3}\s*[ACG]\s*\d$'

  3) Voor mapping op Arango (Vincotte enrichment) mogen witruimtes genegeerd worden, normaliseer altijd (maak hoofdletters waar relevant).

PDF detectie en OCR
- Detecteer of PDF tekst bevat (bijv. met pdfminer/pdfplumber of PyPDF2: proberen extraheertekst). Als geen bruikbare tekst gevonden wordt, markeer als "scan":
  - Flag: readable = False
  - Optie: run Tesseract OCR en probeer dezelfde extractiestappen op de OCR-output.
- Bij tekstuele PDF's (readable = True): parseer met reguliere expressies en tekstblok-analyse.

Parsing-strategie (algemeen)
1) Extract volledige tekst van PDF (of OCR output) en normaliseer whitespace.
2) Zoek eerst naar keuringsdatum:
   - Gebruik regex die zoekt naar 'Datum' of 'Datum(s) van controle' gevolgd door een DD/MM/YYYY.
   - Normaliseer naar YYYY-MM-DD.
3) Zoek resultaat:
   - Zoek het blok na 'Besluit van de controle:'; neem de eerstvolgende zin of de eerste regel met een duidelijke conclusie.
   - Normaliseer naar één van de bekende waarden waar mogelijk (vertaal/verkort).
4) Als het resultaat 'niet conform' of 'niet-conform met inbreuken' of er expliciet 'inbreuken' genoemd worden:
   - Zoek naar kopjes of paragrafen met woorden 'vastgestelde inbreuken', 'vastgestelde inbreuken:' of 'vastgestelde inbreuken, die'
   - Verzamel de tekst van de paragraaf of de opsomming en zet in kolom 'inbreuken'.

Output / Excel-kolommen
Voor elk record/regel in de bestaande Excel (Lijst miv.xlsx of verrijkte variant) willen we toevoegen:
- pdf_filename -> bestandsnaam
- pdf_readable -> True/False (tekst of scan)
- keuringsdatum_raw -> originele gevonden string
- keuringsdatum -> genormaliseerde ISO YYYY-MM-DD (of leeg)
- resultaat_raw -> originele conclusie
- resultaat_normalized -> een van [conform, conform met opmerkingen, niet-conform, niet-conform met inbreuken, niet van toepassing, niet gekend]
- inbreuken -> samengevoegde tekst van vastgestelde inbreuken (of leeg)
- mapping_diagnostics -> korte boodschap waarom gemapt/niet gemapt (bv. "mapped by filename MIV041", "no match", "ignored: keuringsattest NEE")

Logging en foutafhandeling
- Voor alle PDF's die niet automatisch gemapt konden worden, maak een lijst met:
  - bestandsnaam
  - reden (geen installatie nummer gevonden / meerdere matches / installatie ongeldig / scan en geen OCR)
- Schrijf diagnostische logs tijdens verwerking: hoeveel PDFs verwerkt, hoeveel readable, hoeveel gemapt, hoeveel niet-gemapte regels.

Implementatie-notities / suggesties
- Gebruik bestaande scripts als referentie (`detect_btv.py`) voor bestandsnaam-parsing en mapping-logica.
- Gebruik `pdfplumber` of `pdfminer.six` voor tekst-extractie; gebruik `pytesseract` + `Pillow` voor OCR wanneer nodig.
- Houd de extractie-regels per PDF-template uitbreidbaar: identificeer templates (bv. op basis van aanwezigheid van specifieke zinnen of lay-out) en implementeer per-template parsers.
- Testcases: maak kleine test-PDF's voor elk geval:
  - tekstuele PDF met datum en resultaat
  - gescande PDF zonder OCR
  - PDF met meerdere mogelijke installatie-nummers in tekst
  - PDF met "inbreuken"

Deliverables
- Script(en) in `VincottePlatform/ExtractPDF` die:
  1) alle PDF's in `VincottePlatform/pdfs` doorlopen
  2) per PDF het hierboven beschrevenen velden proberen te extraheren
  3) de oorspronkelijke Excel (Lijst miv.xlsx) aanvullen met de nieuwe kolommen en een nieuwe file schrijven `Lijst miv.filled.xlsx`
  4) een aparte CSV of sheet met "not mapped" / diagnostics


Versiebeheer-documentatie
- Documenteer wijzigingen en belangrijke parsing-regels in dit `spec.md` en maak per-template kort extra md-bestanden (bv. `template_miv041.md`) als er speciale uitzonderingen worden gevonden.

Einde document
