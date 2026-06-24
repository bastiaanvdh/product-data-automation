"""
DE Taalcodes Generator
======================
Genereert 4D1/4D2/4DF/4DT/4DU/4DL op basis van de NL-resultaten
uit nieuw_product_generator. Zelfstandige module, geen externe config nodig.
"""

import json
import re
import time

# ─── Constanten ──────────────────────────────────────────────────────────────

_BESCHERMDE_MERKEN = [
    "Tercoo", "FEIN", "Hempel", "Jotun", "Hempatex", "Copaint",
    "Workman", "Festool", "FOV", "Perago", "Multi-8",
    "International", "Sika", "Anza", "Festool", "Motip",
]

_USPS_DE_PAREN = [
    ("Schnelle Lieferung", "günstige Preise"),
    ("Schnelle Lieferung", "breites Sortiment"),
    ("Günstige Preise", "persönliche Beratung"),
    ("Breites Sortiment", "persönliche Beratung"),
]

_FOV_HTML_CSS = """<style>
  .tab-style { display: block; padding: 1rem; border: 1px solid #ccc; background: white; margin-bottom: 1.5rem; border-radius: 5px; }
  .tab-style h2 { margin-top: 0; color: #103a5d; }
  .tab-style strong { color: #103a5d; }
  .tabs { margin-top: 2rem; }
  .tabs input[type="radio"] { display: none; }
  .tabs label { padding: 0.5rem 1rem; background: #eee; margin-right: 0.2rem; cursor: pointer; border-top-left-radius: 5px; border-top-right-radius: 5px; font-weight: bold; color: #103a5d; }
  .tabs label:hover { background: #ddd; }
  .tabs .tab-content { display: none; border: 1px solid #ccc; padding: 1rem; background: white; border-top: none; border-radius: 0 0 5px 5px; }
  .tabs input[type="radio"]:checked + label { background: #103a5d; color: white; }
  .tabs input[type="radio"]:checked + label + .tab-content { display: block; }
</style>"""

_MODEL_KORT = "gpt-4o-mini"
_MODEL_HTML = "gpt-4o"

_PROMPT_DE_KORT = """\
Je bent een professionele productcontent schrijver voor fovfarbe.de (de Duitse webshop van FOV, een Nederlandse verfgroothandel).

Genereer de volgende DE taalcodes op basis van de NL-informatie hieronder.
Geef je antwoord UITSLUITEND als geldig JSON object, geen uitleg, geen markdown.

Artikelinformatie:
- Artikelnaam: {omschrijving}
- Categorie: {opbrengstgroep}
- NL 1N1: {n1}
- NL 1N2: {n2}
- NL 1NF: {nf}

Genereer dit JSON:
{{
  "4D1": "Meta titel in het Duits, exact dit formaat: [Duitstalige productnaam] | [VERTAAL '{opbrengstgroep}' naar het Duits, bijv. 'Jotun - Maritiem' → 'Jotun - Meerestechnik'] | FOV",
  "4D2": "Commerciële zin IN HET DUITS: {de_naam} [korte Duitstalige toepassing/kernspec]. {usps_de} bei FOV! — max 160 tekens totaal. Geen pipes.",
  "4DF": "Duitstalige factuurtitel: alleen de schone productnaam in het Duits, max 50 tekens",
  "4DT": "Exact gelijk aan 4DF",
  "4DU": "URL-slug van 4DF: ALLEEN kleine letters a-z, cijfers en koppeltekens. Geen hoofdletters, umlauts omzetten (ä→ae, ö→oe, ü→ue, ß→ss). 'edition' blijft 'edition'. Voorbeeld: 'jotun-hardtop-flexi', 'hempel-hempadur-mastic'"
}}

Strikte vertaalregels:
- Merknamen NOOIT vertalen: {merken}
- "Farbe bei FOV!" (niet "Farbe ist FOV!")
- "Edition" blijft altijd "Edition", NOOIT "Ausgabe"
- 4DT = exact gelijk aan 4DF, kopieer letterlijk
- Enamel ≠ Email (pas producttypes zorgvuldig aan)"""

_PROMPT_DE_HTML_VERTALING = """\
Je bent een professionele productcontent schrijver voor fovfarbe.de.

Vertaal de onderstaande Nederlandse HTML productpagina naar het Duits voor 4DL.
Behoud EXACT dezelfde HTML-structuur, CSS en opmaak. Vertaal alleen de tekst.

Strikte regels:
- Merknamen NOOIT vertalen: {merken}
- "Edition" blijft altijd "Edition", nooit "Ausgabe"
- Gebruik formele schrijfstijl (Sie-form)
- Behoud alle HTML-tags, class-namen, style-attributen intact
- Vervang tab radio-button IDs: tab1_xxx → tab1_de_xxx (ter vermijding van conflicten)
- Umlauts in tekst zijn OK (ä, ö, ü, ß), maar NIET in URL-slugs
- Geef ALLEEN de HTML terug, geen uitleg, geen markdown code blocks

Nederlandse HTML (1NL):
{nl_html}"""


# ─── Hulpfuncties ─────────────────────────────────────────────────────────────

def _kies_usps(artikelnummer: str) -> str:
    som = sum(int(c) for c in str(artikelnummer) if c.isdigit())
    u1, u2 = _USPS_DE_PAREN[som % len(_USPS_DE_PAREN)]
    return f"{u1}, {u2}"


def _normaliseer_slug(tekst: str) -> str:
    s = tekst.strip()
    # Strip RAL/JTN/NCS kleurcodes en basisaanduidingen (B1/B3) — zelfde logica als 1NH
    s = re.sub(r"\s+(?:RAL|JTN|NCS)\s*[-/]?\s*\d+\S*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*[-–]?\s*\bB[0-9]\b", "", s, flags=re.IGNORECASE).strip()
    s = s.lower()
    for umlaut, vervang in [("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")]:
        s = s.replace(umlaut, vervang)
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"[^a-z0-9\-]", "", s)
    s = re.sub(r"-+", "-", s)
    return s.strip("-")


def _strip_code_fence(tekst: str) -> str:
    tekst = tekst.strip()
    if tekst.startswith("```"):
        eerste_newline = tekst.find("\n")
        if eerste_newline != -1:
            tekst = tekst[eerste_newline + 1:].strip()
        if tekst.endswith("```"):
            tekst = tekst[:-3].rstrip()
    return tekst


# ─── Generatiefuncties ────────────────────────────────────────────────────────

def _genereer_de_kort(product: dict, nl_ai: dict, client) -> dict | None:
    """Genereert 4D1/4D2/4DF/4DT/4DU via gpt-4o-mini."""
    artikelnummer = str(product.get("Artikelnummer", "0"))
    omschrijving  = str(product.get("Omschrijving", ""))
    opbrengstgroep = str(product.get("Opbrengstgroep", ""))
    n1 = nl_ai.get("1N1", "")
    nf = nl_ai.get("1NF", "")
    de_naam = nf or omschrijving

    prompt = _PROMPT_DE_KORT.format(
        omschrijving=omschrijving,
        opbrengstgroep=opbrengstgroep,
        n1=n1,
        n2=nl_ai.get("1N2", ""),
        nf=nf,
        de_naam=de_naam,
        usps_de=_kies_usps(artikelnummer),
        merken=", ".join(_BESCHERMDE_MERKEN),
    )

    try:
        response = client.chat.completions.create(
            model=_MODEL_KORT,
            temperature=0.3,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": prompt}],
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"  [DE] Korte velden mislukt: {e}")
        return None


def _genereer_de_html(nl_ai: dict, client) -> str | None:
    """Genereert 4DL via gpt-4o (vertaling van 1NL)."""
    nl_html = nl_ai.get("1NL", "").strip()
    if not nl_html:
        return None

    prompt = _PROMPT_DE_HTML_VERTALING.format(
        nl_html=nl_html,
        merken=", ".join(_BESCHERMDE_MERKEN),
    )

    try:
        response = client.chat.completions.create(
            model=_MODEL_HTML,
            temperature=0.4,
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}],
        )
        return _strip_code_fence(response.choices[0].message.content)
    except Exception as e:
        print(f"  [DE] HTML mislukt: {e}")
        return None


# ─── Publieke interface ───────────────────────────────────────────────────────

DE_KOLOMMEN = [
    "4D1_FOV_DE_META_DATA_1",
    "4D2_FOV_DE_META_DATA_2",
    "4DF_FOV_DE_TITLE_FACTUUR",
    "4DT_FOV_DE_TITLE_WEB",
    "4DU_FOV_DE_URL",
    "4DL_FOV_DE_LANGE_OMSCHRIJVING",
]

DE_VELDEN = set(DE_KOLOMMEN)


def genereer_de_taalcodes(product: dict, nl_ai: dict, client) -> dict:
    """
    Genereert alle DE taalcodes op basis van het NL-AI-resultaat.
    Retourneert dict met DE_KOLOMMEN als keys (lege strings bij fout).
    """
    leeg = {k: "" for k in DE_KOLOMMEN}

    kort = _genereer_de_kort(product, nl_ai, client)
    if not kort:
        return leeg

    time.sleep(0.5)

    # Normaliseer outputs
    for tc in ["4DF", "4DT"]:
        if kort.get(tc):
            kort[tc] = kort[tc].split("\n")[0].strip()[:50]
    if kort.get("4DU"):
        kort["4DU"] = _normaliseer_slug(kort["4DU"])
    if kort.get("4DT") and kort.get("4DF"):
        kort["4DT"] = kort["4DF"]

    de_html = _genereer_de_html(nl_ai, client)

    return {
        "4D1_FOV_DE_META_DATA_1":       kort.get("4D1", ""),
        "4D2_FOV_DE_META_DATA_2":       kort.get("4D2", ""),
        "4DF_FOV_DE_TITLE_FACTUUR":     kort.get("4DF", ""),
        "4DT_FOV_DE_TITLE_WEB":         kort.get("4DT", ""),
        "4DU_FOV_DE_URL":               kort.get("4DU", ""),
        "4DL_FOV_DE_LANGE_OMSCHRIJVING": de_html or "",
    }
