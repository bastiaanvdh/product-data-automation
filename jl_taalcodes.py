"""
Jachtlakken.nl Taalcodes Generator
====================================
Genereert 2N1/2N2/2NF/2NT/2NH/2NL op basis van de NL-resultaten
uit nieuw_product_generator. Zelfstandige module, geen externe config nodig.

Toon: B2C, informeel "je/jij", geen em-dashes, geen emoji's.
"""

import json
import re
import time

# ─── Constanten ──────────────────────────────────────────────────────────────

_BESCHERMDE_MERKEN = [
    "Epifanes", "Hempel", "International", "Jotun", "West Systems", "Esdé",
    "Anza", "Staalmeester", "Hydrant", "Nelf", "Aemme", "Hempatex",
    "Sika", "Tec7", "Festool", "Motip",
]

_USPS = [
    "Hoge kortingen",
    "Voor 15:00 besteld, morgen in huis",
    "Gratis verzending v.a. €75",
    "14 dagen ruilen en retourneren",
    "Persoonlijk advies van verfexperts",
]

_JL_HTML_CSS = """\
<style>
  .jl-block { padding: 1.25rem 1.5rem; border: 1px solid #dee2e6; background: #fff; border-radius: 6px; margin-bottom: 1.25rem; }
  .jl-block h2 { margin-top: 0; color: #003a60; font-size: 1.1rem; }
  .jl-intro { background: #eaf7f8; border-color: #c9eff1; }
  .jl-steps ol { padding-left: 1.25rem; margin: 0; }
  .jl-steps li { margin-bottom: 0.5rem; line-height: 1.6; }
  .jl-checks ul { list-style: none; padding: 0; margin: 0; }
  .jl-checks ul li::before { content: "✓ "; color: #30b7be; font-weight: bold; }
  .jl-checks ul li { margin-bottom: 0.4rem; }
  .jl-specs table { width: 100%; border-collapse: collapse; font-size: 0.93rem; }
  .jl-specs td { padding: 0.45rem 0.6rem; border-bottom: 1px solid #dee2e6; }
  .jl-specs tr:last-child td { border-bottom: none; }
  .jl-specs td:first-child { color: #003a60; font-weight: 600; width: 45%; }
  .jl-usps { background: #003a60; color: #fff; border-color: #003a60; display: flex; flex-wrap: wrap; gap: 0.75rem 2rem; align-items: center; }
  .jl-usps span::before { content: "✓ "; color: #30b7be; font-weight: bold; }
</style>"""

_MODEL_KORT = "gpt-4o-mini"
_MODEL_HTML  = "gpt-4o"

_PROMPT_JL_KORT = """\
Je bent een professionele productcontent schrijver voor Jachtlakken.nl (een B2C webshop voor bootverf en jachtlakken, gericht op recreatieve boteneigenaren en DIY-schilders).

Genereer de volgende taalcodes op basis van de productinformatie hieronder.
Geef je antwoord UITSLUITEND als geldig JSON object, geen uitleg, geen markdown.

Artikelinformatie:
- Artikelnaam: {omschrijving}
- Categorie: {opbrengstgroep}
- FOV metatitel (1N1): {n1}
- FOV metabeschrijving (1N2): {n2}
- FOV factuurtitel (1NF): {nf}

Genereer dit JSON:
{{
  "2N1": "[Productnaam] | [Categorie] | Jachtlakken",
  "2N2": "Informele metabeschrijving (max 160 tekens). Productnaam + wat het doet + USPs. Gebruik 'je/jij'. Geen em-dashes. Geen specifieke kortingspercentages (schrijf 'hoge kortingen' niet '20% korting'). Eindig met een of twee USPs. Voorbeeld: 'Hempel Hempaspeed TF 77222: biocidevrije antifouling voor alle wateren en aluminium. {usps} bij Jachtlakken!'",
  "2NF": "Schone factuurtitel: alleen de productnaam, max 50 tekens",
  "2NT": "Exact gelijk aan 2NF",
  "2NH": "URL-slug van 2NF: ALLEEN kleine letters a-z, cijfers en koppeltekens. Geen RAL/JTN/NCS codes. Voorbeeld: 'hempel-hempaspeed-tf-77222', 'epifanes-cr-antifouling-primer'"
}}

Strikte regels:
- Merknamen NOOIT aanpassen of vertalen: {merken}
- Aanspreekvorm: altijd "je/jij", nooit "u"
- Geen em-dashes (— of –) in enig veld
- Geen emoji's
- 2N1 eindigt altijd op "| Jachtlakken" (zonder .nl)
- 2NT = exact gelijk aan 2NF, kopieer letterlijk
- 2NH: RAL/JTN/NCS codes en basisaanduidingen (B1/B3) weglaten"""

_PROMPT_JL_HTML = """\
Je bent een professionele productcontent schrijver voor Jachtlakken.nl (B2C webshop voor bootverf, gericht op recreatieve boteneigenaren).

Schrijf een volledige HTML productpagina (2NL) in scroll-layout voor dit product.
Gebruik de informatie uit de FOV teksten hieronder als basis, maar herschrijf alles in een informele B2C stijl.

Artikelinformatie:
- Artikelnaam: {omschrijving}
- Categorie: {opbrengstgroep}
- FOV metatitel: {n1}
- FOV metabeschrijving: {n2}

FOV HTML productpagina (ter referentie, NIET kopiëren):
{nl_html}

Strikte schrijfregels:
- Aanspreekvorm: altijd "je/jij", nooit "u"
- Geen em-dashes (— of –)
- Geen emoji's
- Informele, toegankelijke toon. Nautische context mag.
- Problem-solution opbouw: begin met waarom dit product het probleem van de schipper oplost
- Merknamen NOOIT aanpassen: {merken}

Gebruik EXACT deze HTML-structuur en CSS (voeg geen andere stijlen toe):

{css}

Bouw de pagina op uit precies deze 5 blokken in deze volgorde:

1. <div class="jl-block jl-intro"> met <h2> en een inleidende alinea (2-3 zinnen, waarom dit product)
2. <div class="jl-block jl-checks"> met <h2>Waarom de [productnaam]?</h2> en <ul> met 4-5 <li> kernvoordelen
3. <div class="jl-block jl-steps"> met <h2>Hoe breng je het aan?</h2> en <ol> met genummerde stappen. Splits per situatie (bijv. nieuwbouw vs. bestaande coating) als dat van toepassing is.
4. <div class="jl-block jl-specs"> met <h2>Technische specificaties</h2> en een <table> met relevante specs als rijen
5. <div class="jl-block jl-usps"> met precies deze 5 <span>-elementen (in deze volgorde):
   <span>Hoge kortingen</span>
   <span>Voor 15:00 besteld, morgen in huis</span>
   <span>Gratis verzending v.a. €75</span>
   <span>14 dagen ruilen en retourneren</span>
   <span>Persoonlijk advies van verfexperts</span>

Geef ALLEEN de HTML terug (beginnend met <style>), geen uitleg, geen markdown code blocks."""


# ─── Hulpfuncties ─────────────────────────────────────────────────────────────

def _normaliseer_slug(tekst: str) -> str:
    s = tekst.strip()
    s = re.sub(r"\s+(?:RAL|JTN|NCS)\s*[-/]?\s*\d+\S*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*[-–]?\s*\bB[0-9]\b", "", s, flags=re.IGNORECASE).strip()
    s = s.lower()
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


def _kies_usps(artikelnummer: str) -> str:
    som = sum(int(c) for c in str(artikelnummer) if c.isdigit())
    u1 = _USPS[som % len(_USPS)]
    u2 = _USPS[(som + 2) % len(_USPS)]
    if u1 == u2:
        u2 = _USPS[(som + 1) % len(_USPS)]
    return f"{u1}, {u2}"


# ─── Generatiefuncties ────────────────────────────────────────────────────────

def _genereer_jl_kort(product: dict, nl_ai: dict, client) -> dict | None:
    artikelnummer  = str(product.get("Artikelnummer", "0"))
    omschrijving   = str(product.get("Omschrijving", ""))
    opbrengstgroep = str(product.get("Opbrengstgroep", ""))

    prompt = _PROMPT_JL_KORT.format(
        omschrijving=omschrijving,
        opbrengstgroep=opbrengstgroep,
        n1=nl_ai.get("1N1", ""),
        n2=nl_ai.get("1N2", ""),
        nf=nl_ai.get("1NF", ""),
        usps=_kies_usps(artikelnummer),
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
        print(f"  [JL] Korte velden mislukt: {e}")
        return None


def _genereer_jl_html(product: dict, nl_ai: dict, client) -> str | None:
    nl_html = nl_ai.get("1NL", "").strip()

    prompt = _PROMPT_JL_HTML.format(
        omschrijving=str(product.get("Omschrijving", "")),
        opbrengstgroep=str(product.get("Opbrengstgroep", "")),
        n1=nl_ai.get("1N1", ""),
        n2=nl_ai.get("1N2", ""),
        nl_html=nl_html or "(niet beschikbaar)",
        merken=", ".join(_BESCHERMDE_MERKEN),
        css=_JL_HTML_CSS,
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
        print(f"  [JL] HTML mislukt: {e}")
        return None


# ─── Publieke interface ───────────────────────────────────────────────────────

JL_KOLOMMEN = [
    "2N1_JL_NL_META_DATA_1",
    "2N2_JL_NL_META_DATA_2",
    "2NF_JL_NL_TITLE_FACTUUR",
    "2NT_JL_NL_TITLE_WEB",
    "2NH_JL_NL_URL",
    "2NL_JL_NL_LANGE_OMSCHRIJVING",
]

JL_VELDEN = set(JL_KOLOMMEN)


def genereer_jl_taalcodes(product: dict, nl_ai: dict, client) -> dict:
    """
    Genereert alle Jachtlakken.nl taalcodes op basis van het NL-AI-resultaat.
    Retourneert dict met JL_KOLOMMEN als keys (lege strings bij fout).
    """
    leeg = {k: "" for k in JL_KOLOMMEN}

    kort = _genereer_jl_kort(product, nl_ai, client)
    if not kort:
        return leeg

    time.sleep(0.5)

    # Normaliseer outputs
    for tc in ["2NF", "2NT"]:
        if kort.get(tc):
            kort[tc] = kort[tc].split("\n")[0].strip()[:50]
    if kort.get("2N1"):
        kort["2N1"] = re.sub(r"^(?:meta\s*titel\s*:\s*)", "", kort["2N1"], flags=re.IGNORECASE).strip()
    if kort.get("2NH"):
        kort["2NH"] = _normaliseer_slug(kort["2NH"])
    if kort.get("2NF"):
        kort["2NT"] = kort["2NF"]

    jl_html = _genereer_jl_html(product, nl_ai, client)

    return {
        "2N1_JL_NL_META_DATA_1":       kort.get("2N1", ""),
        "2N2_JL_NL_META_DATA_2":       kort.get("2N2", ""),
        "2NF_JL_NL_TITLE_FACTUUR":     kort.get("2NF", ""),
        "2NT_JL_NL_TITLE_WEB":         kort.get("2NT", ""),
        "2NH_JL_NL_URL":               kort.get("2NH", ""),
        "2NL_JL_NL_LANGE_OMSCHRIJVING": jl_html or "",
    }
