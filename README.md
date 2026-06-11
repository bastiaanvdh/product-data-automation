# Product Data Automatisering

Een Flask webapplicatie die het aanmaken van productimportbestanden voor ERP-systemen automatiseert. Op basis van een eenvoudige spreadsheet met basisproductinformatie genereert de tool een compleet, importklaar Excel-bestand met AI-geschreven content, gescrapede leveranciersdata, verwerkte productafbeeldingen en links naar technische documenten.

Gebouwd om het repetitieve handmatige werk bij het aanmaken van nieuwe productlijsten te elimineren — wat vroeger 30 tot 60 minuten per product kostte, duurt nu enkele seconden.

---

## Wat het doet

Upload een spreadsheet met artikelnummer, omschrijving, merk en optioneel een leveranciers-URL. De tool regelt de rest:

1. **Scrapet de leverancierspagina** — haalt producttekst op en zoekt links naar technische PDF-documenten (PIB/VIB)
2. **Downloadt en slaat PDF's op** — weggeschreven naar de juiste mappenstructuur op het netwerk
3. **Downloadt productafbeeldingen** — filtert logo's en sfeerfoto's op basis van bestandsnaam en alt-tekst, schaalt naar exact 1600×1600 px met witte achtergrond (zowel omhoog als omlaag), slaat op op de NAS
4. **Genereert content via OpenAI (gpt-4o-mini)** — webtitel, factuurnaam, URL-slug, metatitel (max 70 tekens), metabeschrijving (max 160 tekens), volledige HTML-productomschrijving (~300 woorden)
5. **Classificeert het product** — selecteert het juiste attribuutset-code en categorie-ID's uit referentiebestanden
6. **Genereert een complete import-Excel** — 42 kolommen in het exacte ERP-importformaat, kleurgecodeerd op betrouwbaarheid en databron

---

## Technische stack

| | |
|---|---|
| **Backend** | Python, Flask |
| **Scraping** | httpx, BeautifulSoup4 |
| **AI** | OpenAI API (gpt-4o-mini) |
| **Beeldverwerking** | Pillow |
| **Excel** | openpyxl, pandas |
| **Concurrency** | ThreadPoolExecutor |

---

## Belangrijkste kenmerken

- **Testmodus** — één config-vlag stuurt alle bestandsschrijfacties naar een lokale map; geen risico op aanpassen van productiedata tijdens ontwikkeling
- **SSL-proxy compatibiliteit** — werkt op bedrijfsnetwerken met onderscheppende proxies (`httpx verify=False`)
- **URL-fallback voor PDF's** — als een leveranciers-URL wordt meegegeven in plaats van een lokaal pad, downloadt de tool de PDF automatisch; valt graceful terug als de download mislukt
- **Afbeeldingsfilter** — controleert bestandsnaam en alt-tekst op merk- en productnaam (niet de volledige URL, om false positives op domeinnamen te voorkomen)
- **Betrouwbaarheidskleur** — AI-classificatieresultaten worden groen/geel/rood gekleurd op basis van de confidence score
- **Voortgangspolling** — langlopende taken streamen hun status terug naar de UI; geen pagina-verversingen nodig

---

## Installatie

```bash
pip install flask httpx beautifulsoup4 openai openpyxl pandas pillow
```

Stel je OpenAI API-sleutel in:

```bash
export OPENAI_API_KEY=sk-...
```

Plaats de referentiebestanden (attribuutset en categoriestructuur) in `referentiedata/` of configureer de netwerkpaden in `config.py`.

```bash
python app.py
# → http://localhost:5002
```

---

## Projectstructuur

```
├── app.py                        # Flask routes
├── nieuw_product_generator.py    # Kernlogica: scraping, AI, beeldverwerking, Excel-output
├── config.py                     # Paden, API-instellingen, testmodus-vlag
├── templates/
│   └── index.html                # Upload-UI en voortgangsweergave
└── referentiedata/
    ├── attribuutset_v4.xlsx      # Productattribuut-referentie
    ├── categorieindeling.xlsx    # Categoriestructuur-referentie
    └── merk_domeinen.json        # Merk-naar-domein mapping (handmatig bijgehouden)
```

---

## Output

De gegenereerde Excel bevat 42 kolommen in het exacte formaat voor ERP-import:

| Kolomgroep | Bron |
|---|---|
| Vaste waarden (actief, shop, merk) | Config |
| Taalcodes (titel, slug, meta, HTML-omschrijving) | OpenAI |
| Attribuutset + categorie-ID's | OpenAI + referentiebestanden |
| Afbeeldingspaden (hoofd + max 5 extra's) | Gescraped + verwerkt |
| Technische documentpaden (PIB/VIB) | Gescraped of handmatig meegegeven |
| Commerciële velden (prijs, EAN) | Doorgegeven vanuit invoer |