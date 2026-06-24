# Product Data Automatisering

Een Flask-webapplicatie die het aanmaken van productcontent voor meerdere webshops volledig automatiseert. Op basis van een spreadsheet met basisproductinformatie genereert de tool een importklare Excel met AI-geschreven content in drie talen en toonsoorten, geselecteerde productafbeeldingen en directe push naar het ERP-systeem.

Gebouwd om handmatig productwerk te elimineren: wat vroeger 30 tot 60 minuten per product kostte, duurt nu enkele seconden — voor drie webshops tegelijk.

---

## Wat het doet

Upload een spreadsheet met artikelnummer, omschrijving en merk. De tool regelt de rest:

1. **Scrapet de leverancierspagina** — haalt producttekst op en zoekt links naar technische PDF-documenten (PIB/VIB); valt terug op DuckDuckGo-zoekopdrachten als er geen URL beschikbaar is
2. **Downloadt en slaat PDF's op** — productinformatiebladen en veiligheidsbladen weggeschreven naar de juiste mappenstructuur op het netwerk, met fuzzy filename-matching
3. **Selecteert de beste productafbeelding via GPT-4o Vision** — verzamelt tot 5 kandidaatafbeeldingen van concurrerende webshops via DuckDuckGo, laat GPT-4o Vision de beste kiezen (witachtergrond, geen kleurstaal, geen sfeerfoto), schaalt naar 1600×1600 px en slaat op op de NAS
4. **Genereert content voor drie webshops via OpenAI:**
   - **FOV NL** (1N*) — altijd, B2B zakelijk Nederlands
   - **FOV DE** (4D*) — altijd, vertaling + herschrijving in B2B Duits voor fovfarbe.de
   - **Jachtlakken.nl** (2N*) — optioneel, B2C informeel Nederlands met eigen HTML-layout en USP-structuur
5. **Classificeert het product** — selecteert attribuutset-code en categorie-ID's uit referentiebestanden met confidence scoring
6. **Genereert een complete import-Excel** — kleurgecodeerd op databron (handmatig/AI/mapping)
7. **Pusht rechtstreeks naar KING ERP** — via de REST API, met dry-run modus en sessiegeschiedenis

---

## Multi-webshop content

De generator behandelt drie webshops als aparte tenants met eigen taalcodes, tone of voice en HTML-templates:

| Webshop | Codes | Toon | HTML |
|---|---|---|---|
| FOV (fov.nl) | 1N1/1N2/1NF/1NH/1NL/1NT | B2B, zakelijk | Tab-layout |
| FOV DE (fovfarbe.de) | 4D1/4D2/4DF/4DT/4DU/4DL | B2B, Duits | Tab-layout (vertaald) |
| Jachtlakken.nl | 2N1/2N2/2NF/2NT/2NH/2NL | B2C, informeel "je" | Scroll-layout (5 blokken) |

Elke tenant heeft een eigen module (`de_taalcodes.py`, `jl_taalcodes.py`) met afzonderlijke prompts, USP-sets, merkenlijsten en CSS-templates. FOV NL en DE worden altijd gegenereerd; Jachtlakken.nl is aan/uit via een checkbox in de UI.

---

## Technische stack

| | |
|---|---|
| **Backend** | Python, Flask |
| **Scraping** | httpx, BeautifulSoup4 |
| **Beeldzoeken** | DuckDuckGo Search (ddgs) |
| **AI — tekst** | OpenAI API: gpt-4o-mini (korte velden), gpt-4o (HTML) |
| **AI — beeld** | OpenAI GPT-4o Vision (kandidaatselectie) |
| **Beeldverwerking** | Pillow |
| **PDF-verwerking** | pypdf |
| **Excel** | openpyxl, pandas |
| **ERP** | KING REST API (directe push) |

---

## Kernkenmerken

- **Vision-gebaseerde afbeeldingsselectie** — GPT-4o beoordeelt meerdere kandidaatafbeeldingen van concurrent-webshops en kiest de meest geschikte productfoto (geen kleurstalen, geen lifestyle)
- **Multi-tenant contentgeneratie** — drie volledig afzonderlijke prompts, USP-sets en HTML-templates per webshop
- **B2B vs. B2C differentiatie** — FOV NL/DE in zakelijke toon, Jachtlakken.nl in informeel "je" met nautische context
- **KING ERP directe push** — geen tussenstap: gegenereerde data gaat rechtstreeks via REST API het ERP in, met dry-run modus en sessiegeschiedenis
- **Afbeelding RAL/kleurcode-stripping** — kleurvarianten delen één gedeelde productfoto; bestandsnaam wordt automatisch gestript van RAL/JTN/NCS-codes en basisaanduidingen
- **Testmodus** — één config-vlag stuurt alle schrijfacties naar een lokale map; geen risico op aanpassen van productiedata
- **SSL-proxy compatibiliteit** — werkt op bedrijfsnetwerken met onderscheppende proxies
- **Voortgangspolling** — langlopende taken streamen hun status live naar de UI

---

## Installatie

```bash
pip install flask httpx beautifulsoup4 openai openpyxl pandas pillow pypdf ddgs
```

```bash
export OPENAI_API_KEY=sk-...
```

Plaats referentiebestanden in `referentiedata/` of configureer netwerkpaden in `config.py`.

```bash
python app.py
# → http://localhost:5002
```

---

## Projectstructuur

```
├── app.py                        # Flask routes + KING push endpoints
├── nieuw_product_generator.py    # Kernlogica: scraping, Vision, AI, Excel-output
├── de_taalcodes.py               # FOV DE tenant — prompts, slug-normalisatie, HTML-vertaling
├── jl_taalcodes.py               # Jachtlakken.nl tenant — B2C prompts, scroll-layout HTML
├── king_artikel_push.py          # KING ERP REST API integratie
├── config.py                     # Paden, API-instellingen, testmodus-vlag
├── templates/
│   ├── index.html                # Hoofd-UI: upload, instellingen, voortgang, push
│   └── formulier.html            # Direct invoerformulier (alternatief voor Excel)
└── referentiedata/
    ├── attribuutset_v4.xlsx      # Productattribuut-referentie
    ├── categorieindeling.xlsx    # Categoriestructuur-referentie
    └── merk_domeinen.json        # Merk-naar-domein mapping
```

---

## Output

De gegenereerde Excel bevat kolommen per webshop, kleurgecodeerd op databron:

| Kolomgroep | Bron | Kleur |
|---|---|---|
| FOV NL taalcodes (1N*) — titel, slug, meta, HTML | OpenAI | Groen |
| FOV DE taalcodes (4D*) — Duits | OpenAI | Groen |
| Jachtlakken.nl taalcodes (2N*) — B2C NL | OpenAI | Lichtblauw |
| Attribuutset + categorie-ID's | OpenAI + referentie | Geel |
| Afbeeldingspaden (hoofd + max 5 extra's) | Vision + scraping | — |
| Technische documentpaden (PIB/VIB) | Gescraped | — |
| Commerciële velden (prijs, EAN, leverancier) | Invoer/handmatig | Oranje |
