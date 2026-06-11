"""
Product Import Generator - KING ERP Import
==========================================
Genereert een KING-importexcel voor nieuwe producten.

Input:  Excel/CSV met minimaal: Artikelnummer, Omschrijving, Merk
Output: KING import Excel klaar voor Excel2King

Gebruik:
    python nieuw_product_generator.py   (via app.py of standalone)
"""

import io
import json
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import httpx
import openpyxl
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openai import OpenAI
from PIL import Image

import config

# Doelmap productafbeeldingen — NAS in productie, lokale testmap in testmodus
AFBEELDING_BASE_DIR = (
    config._TEST_DIR / "Productafbeeldingen"
    if config.TEST_MODUS
    else config.AFBEELDING_NAS_PAD
)
AFBEELDING_GROOTTE  = (1600, 1600)
# Minimale afmeting om logo's/iconen te filteren
AFBEELDING_MIN_PX   = 200

# ─────────────────────────────────────────────
#  LEVERANCIERSPAGINA OPHALEN
# ─────────────────────────────────────────────

_PIB_TREFWOORDEN = ["productinformatieblad", "informatieblad", "tds", "technical data", "productblad", "pib"]
_VIB_TREFWOORDEN = ["veiligheidsblad", "safety data", "sds", "msds", "vib"]
_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


def _schone_bestandsnaam(tekst: str) -> str:
    """Verwijder tekens die niet in bestandsnamen mogen."""
    return re.sub(r'[\\/:*?"<>|]', "", tekst).strip()


def _download_pdf(pdf_url: str, save_pad: Path) -> bool:
    """Download een PDF naar save_pad. Retourneert True bij succes."""
    try:
        save_pad.parent.mkdir(parents=True, exist_ok=True)
        with httpx.stream("GET", pdf_url, timeout=30, follow_redirects=True,
                          verify=False, headers=_HEADERS) as r:
            r.raise_for_status()
            with open(save_pad, "wb") as f:
                for chunk in r.iter_bytes(chunk_size=8192):
                    f.write(chunk)
        return True
    except Exception:
        return False


def _verwerk_afbeelding(img_bytes: bytes, save_pad: Path) -> bool:
    """
    Schaal afbeelding naar 1600x1600 met witte achtergrond (padding).
    Slaat op als JPEG. Retourneert True bij succes.
    """
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
        if img.width < AFBEELDING_MIN_PX or img.height < AFBEELDING_MIN_PX:
            return False
        # Schaal naar maximaal 1600x1600 met behoud van verhouding (omhoog én omlaag)
        schaal = min(AFBEELDING_GROOTTE[0] / img.width, AFBEELDING_GROOTTE[1] / img.height)
        nieuw_w = round(img.width * schaal)
        nieuw_h = round(img.height * schaal)
        img = img.resize((nieuw_w, nieuw_h), Image.LANCZOS)
        canvas = Image.new("RGB", AFBEELDING_GROOTTE, (255, 255, 255))
        offset = ((AFBEELDING_GROOTTE[0] - nieuw_w) // 2,
                  (AFBEELDING_GROOTTE[1] - nieuw_h) // 2)
        canvas.paste(img, offset, mask=img.split()[3] if img.mode == "RGBA" else None)
        save_pad.parent.mkdir(parents=True, exist_ok=True)
        canvas.save(save_pad, "JPEG", quality=92)
        return True
    except Exception:
        return False


def _download_afbeeldingen(soup: BeautifulSoup, base_url: str,
                            merk: str, bestandsnaam_basis: str) -> list[str]:
    """
    Zoek productafbeeldingen op de pagina, download en sla op als 1600x1600 JPEG.
    Filtert op alt-tekst of bestandsnaam die merknaam of productnaam bevat.
    Retourneert lijst van lokale paden (strings).
    """
    merk_slug = merk.lower().strip()
    merk_dir  = AFBEELDING_BASE_DIR / merk_slug
    paden: list[str] = []
    gezien: set[str] = set()

    # Zoekwoorden waarop gefilterd wordt (merk + losse woorden uit omschrijving)
    basis_slug   = re.sub(r"[^a-z0-9]", "", bestandsnaam_basis.lower())
    zoekwoorden  = [re.sub(r"[^a-z0-9]", "", merk_slug)]
    for woord in bestandsnaam_basis.lower().split():
        slug = re.sub(r"[^a-z0-9]", "", woord)
        if len(slug) > 2:
            zoekwoorden.append(slug)

    def _is_productafbeelding(src: str, alt: str) -> bool:
        # Alleen bestandsnaam checken (niet domein — dat geeft valse matches)
        bestandsnaam = Path(urlparse(src).path).name
        combined = re.sub(r"[^a-z0-9]", "", (bestandsnaam + " " + alt).lower())
        return any(w in combined for w in zoekwoorden)

    # Verzamel kandidaat-afbeelding-URLs
    kandidaten: list[tuple[str, str]] = []  # (full_url, alt)
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or img.get("data-lazy-src") or ""
        if not src:
            continue
        full = urljoin(base_url, src)
        if full in gezien:
            continue
        gezien.add(full)
        if any(x in full.lower() for x in [".svg", ".gif", "data:"]):
            continue
        alt = img.get("alt", "")
        if _is_productafbeelding(full, alt):
            kandidaten.append((full, alt))

    teller = 0
    for img_url, alt in kandidaten:
        if teller >= 6:  # max 6 afbeeldingen per product
            break
        try:
            r = httpx.get(img_url, timeout=15, follow_redirects=True,
                          verify=False, headers=_HEADERS)
            r.raise_for_status()
        except Exception:
            continue

        suffix = ".jpg" if teller == 0 else f" {teller + 1}.jpg"
        save_pad = merk_dir / f"{bestandsnaam_basis}{suffix}"
        if _verwerk_afbeelding(r.content, save_pad):
            paden.append(str(save_pad))
            teller += 1

    return paden


def haal_leverancier_pagina(url: str, merk: str = "", omschrijving: str = "") -> dict:
    """
    Haalt leverancierspagina op en retourneert:
      tekst     - leesbare paginatekst voor AI-prompt
      pib_pad   - lokaal K-schijf pad van gedownload PIB (of "")
      vib_pad   - lokaal K-schijf pad van gedownload VIB (of "")
    """
    if not url or not url.startswith("http"):
        return {"tekst": "", "pib_pad": "", "vib_pad": ""}

    try:
        resp = httpx.get(url, timeout=15, follow_redirects=True,
                         verify=False, headers=_HEADERS)
        resp.raise_for_status()
    except Exception as e:
        return {"tekst": f"[Pagina ophalen mislukt: {e}]", "pib_pad": "", "vib_pad": ""}

    soup = BeautifulSoup(resp.text, "html.parser")

    for tag in soup(["script", "style", "nav", "footer", "header"]):
        tag.decompose()
    tekst = " ".join(soup.get_text(separator=" ").split())[:4000]

    # PDF-links zoeken
    pib_url = vib_url = ""
    base = f"{urlparse(url).scheme}://{urlparse(url).netloc}"

    for a in soup.find_all("a", href=True):
        href = a["href"].lower()
        combined = href + " " + a.get_text().lower()
        if not pib_url and any(t in combined for t in _PIB_TREFWOORDEN):
            pib_url = urljoin(base, a["href"])
        if not vib_url and any(t in combined for t in _VIB_TREFWOORDEN):
            vib_url = urljoin(base, a["href"])

    # PDFs downloaden naar K:\ structuur (zelfde als FLEX Excel)
    # K:\1. Productinformatiebladen\{Merk}\NL\PIB {Merk} {Omschrijving}-NL.pdf
    pib_pad = vib_pad = ""
    merk = merk or ""
    omschrijving = omschrijving or ""
    merk_clean = _schone_bestandsnaam(merk)
    omschr_stripped = omschrijving.strip()
    if merk and omschr_stripped.lower().startswith(merk.lower()):
        omschr_stripped = omschr_stripped[len(merk):].strip()
    omschr_clean = _schone_bestandsnaam(omschr_stripped).strip(" -")
    merk_dir = config.PIB_BASE_DIR / merk_clean / "NL"

    if pib_url:
        pib_bestand = merk_dir / f"PIB {merk_clean} {omschr_clean}-NL.pdf"
        if _download_pdf(pib_url, pib_bestand):
            pib_pad = str(pib_bestand)

    if vib_url:
        vib_bestand = merk_dir / f"VIB {merk_clean} {omschr_clean}-NL.pdf"
        if _download_pdf(vib_url, vib_bestand):
            vib_pad = str(vib_bestand)

    # Afbeeldingen downloaden en opslaan (map wordt aangemaakt indien nodig)
    afbeelding_basis = f"{merk_clean.title()} {omschr_clean}" if omschr_clean else merk_clean.title()
    afbeeldingen = _download_afbeeldingen(soup, url, merk_clean, afbeelding_basis)

    return {"tekst": tekst, "pib_pad": pib_pad, "vib_pad": vib_pad, "afbeeldingen": afbeeldingen}

# ─────────────────────────────────────────────
#  EXCEL KOLOMDEFINITIE (zelfde structuur als FLEX import)
# ─────────────────────────────────────────────

KOLOMMEN = [
    # ── Kolommen 1-42: exact FLEX Import structuur ───────────────
    "Artikelnummer",                        # 1
    "Eenheid",                              # 2
    "Zoekcode",                             # 3  ← Merk
    "Omschrijving",                         # 4
    "Opbrengstgroep",                       # 5
    "WebArtikel",                           # 6  = 1
    "TekstOpFactuur",                       # 7  ← Omschrijving (invoer)
    "AfbeeldingKlein",                      # 8
    "AfbeeldingGroot",                      # 9
    "Leveranciernummer",                    # 10
    "Leveranciernaam",                      # 11 ← Merk
    "ArtikelOmschrijvingLeverancier",       # 12 ← Omschrijving
    "ArtikelNummerBijLeverancier",          # 13
    "EanCode",                              # 14
    "VR_ART_Magentotype",                   # 15 = "Simpel"
    "VR_ART_Zichtbaarheid",                 # 16 = "Catalogus, zoeken"
    "VR_ART_Actief_in_shop",               # 17 = 1
    "VR_ART_F-Merk",                       # 18 ← Merk
    "VR_ART_Extra_afbeelding_1",           # 19
    "VR_ART_Extra_afbeelding_2",           # 20
    "VR_ART_Extra_afbeelding_3",           # 21
    "VR_ART_Extra_afbeelding_4",           # 22
    "VR_ART_Extra_afbeelding_5",           # 23
    "VR_ART_Productinformatieblad_NL",     # 24
    "VR_ART_Productveiligheidsblad_NL_A",  # 25
    "VR_ART_Productveiligheidsblad_NL_B",  # 26
    "VR_ART_Productveiligheidsblad_NL_C",  # 27
    "VR_ART_Productveiligheidsblad_ENG_A", # 28
    "VR_ART_Productveiligheidsblad_ENG_B", # 29
    "VR_ART_Productveiligheidsblad_ENG_C", # 30
    "VR_ART_Productveiligheidsblad_DE_A",  # 31
    "VR_ART_Productveiligheidsblad_DE_B",  # 32
    "VR_ART_Productveiligheidsblad_DE_C",  # 33
    "VR_ART_Explosietekening_UNI",         # 34
    "VR_ART_Producthandleiding_NL",        # 35
    "VR_ART_Attribuutset_V4",              # 36 ← AI
    "1NT_FOV_NL_TITLE_WEB",               # 37 ← AI
    "1NF_FOV_NL_TITLE_FACTUUR",           # 38 ← AI
    "1NH_FOV_NL_URL",                     # 39 ← AI
    "1N1_FOV_NL_META_DATA_1",             # 40 ← AI
    "1N2_FOV_NL_META_DATA_2",             # 41 ← AI
    "1NL_FOV_NL_LANGE_OMSCHRIJVING",      # 42 ← AI
    # ── Extra controle-kolommen (niet in FLEX, voor review) ──────
    "Attribuutset_Confidence_%",
    "Attribuutset_Label",
    "VR_ART_Webcategorie_ID._V2",
    "Webcategorie_Confidence_%",
]

HANDMATIG    = {"Leveranciernummer", "VerkoopPrijsExBTW", "AdviesPrijsExBTW"}
AI_VELDEN    = {
    "1N1_FOV_NL_META_DATA_1", "1N2_FOV_NL_META_DATA_2",
    "1NL_FOV_NL_LANGE_OMSCHRIJVING", "1NT_FOV_NL_TITLE_WEB",
    "1NF_FOV_NL_TITLE_FACTUUR", "1NH_FOV_NL_URL",
}
MAPPING_VELDEN = {
    "VR_ART_Attribuutset_V4", "Attribuutset_Confidence_%", "Attribuutset_Label",
    "VR_ART_Webcategorie_ID._V2", "Webcategorie_Confidence_%",
}


# ─────────────────────────────────────────────
#  REFERENTIEDATA LADEN
# ─────────────────────────────────────────────

def laad_attribuutset(pad: Path) -> list[dict]:
    """
    Laad attribuutset V4 uit het BRON-bestand.
    Kolom 0 (Keuzelijst) = attribuutset code.
    Kolom 6 (Kolom3)     = taxonomy-omschrijving als label.
    """
    if not pad.exists():
        return []
    df = pd.read_excel(pad, header=0)
    SKIP = {"nan", "-", "keuzelijst", ""}
    result = []
    for _, row in df.iterrows():
        code = str(row.iloc[0]).strip()
        if code.lower() in SKIP:
            continue
        # Kolom 6 = taxonomy-pad (bijv. "Bouwmaterialen - Gereedschap - ...")
        label = ""
        if len(row) > 6:
            raw = str(row.iloc[6]).strip()
            if raw not in ("nan", "#N/A", ""):
                label = raw
        if not label:
            label = code.replace("_", " ").title()
        result.append({"code": code, "label": label})
    return result


# Kolomgroepen in categorieindeling: (naam-kolom, id-kolom) paren (0-gebaseerd)
_CAT_GROEPEN = [(1, 3), (4, 6), (7, 9), (10, 12), (13, 15)]

# Rij-waarden die header zijn (overslaan)
_CAT_HEADERS = {
    "nl", "de", "type", "typ", "soort", "toepassing", "anwendung",
    "ondergrond", "substrat", "merken", "merke", "merk", "marke",
    "eigenschappen", "eigenschaft", "toepasing",
}


def laad_categorieindeling(pad: Path) -> list[dict]:
    """
    Laad categorieindeling uit het meerdere-sheets hiërarchische bestand.
    Elke sheet heeft kolomgroepen (naam, DE-naam, ID) op vaste posities.
    Geeft een platte lijst met id, naam en sectie.
    """
    if not pad.exists():
        return []

    skip_sheets = {"Legenda", "Acties"}
    result = []
    seen_ids: set[str] = set()

    xl = pd.ExcelFile(pad)
    for sheet_name in xl.sheet_names:
        if sheet_name in skip_sheets:
            continue
        df = pd.read_excel(pad, sheet_name=sheet_name, header=None)

        for _, row in df.iterrows():
            for naam_col, id_col in _CAT_GROEPEN:
                if naam_col >= len(row) or id_col >= len(row):
                    continue
                naam   = row.iloc[naam_col]
                id_val = row.iloc[id_col]

                if pd.isna(naam) or pd.isna(id_val):
                    continue
                naam   = str(naam).strip()
                id_str = str(id_val).strip()

                if naam.lower() in _CAT_HEADERS or not naam:
                    continue

                try:
                    id_int = int(float(id_str))
                    id_str = str(id_int)
                except (ValueError, OverflowError):
                    continue

                if id_str in seen_ids:
                    continue
                seen_ids.add(id_str)
                result.append({"id": id_str, "naam": naam, "sectie": sheet_name})

    return result


def _compacte_attribuutset_lijst(attribuutsets: list[dict]) -> str:
    return "\n".join(f"  {a['code']}: {a['label']}" for a in attribuutsets)


def _compacte_categorie_lijst(categorieen: list[dict]) -> str:
    by_sec: dict[str, list[dict]] = {}
    for c in categorieen:
        by_sec.setdefault(c.get("sectie", "Overig"), []).append(c)
    lines = []
    for sec, items in by_sec.items():
        lines.append(f"  [{sec}]")
        for c in items:
            lines.append(f"    {c['id']}: {c['naam']}")
    return "\n".join(lines)


# ─────────────────────────────────────────────
#  AI PROMPT
# ─────────────────────────────────────────────

def maak_prompt(product: dict, attribuutsets: list[dict], categorieen: list[dict], idx: int,
                pagina_tekst: str = "") -> str:
    omschrijving = product.get("Omschrijving", "")
    merk         = product.get("Merk", "")
    artikelnr    = product.get("Artikelnummer", "")
    extra_info   = " | ".join(
        f"{k}: {v}" for k, v in product.items()
        if k not in ("Artikelnummer", "Omschrijving", "Merk",
                     "Productinformatieblad_pad", "Productveiligheidsblad_pad",
                     "URL_leverancier")
        and v and str(v) not in ("nan", "None")
    )
    if pagina_tekst:
        extra_info = (extra_info + " | " if extra_info else "") + f"Leverancierspagina: {pagina_tekst}"

    usp = config.SHOP_USPS[idx % len(config.SHOP_USPS)]

    attrib_sectie = (
        _compacte_attribuutset_lijst(attribuutsets)
        if attribuutsets
        else "  (geen referentiebestand gevonden – gebruik beste schatting)"
    )
    cat_sectie = (
        _compacte_categorie_lijst(categorieen)
        if categorieen
        else "  (geen referentiebestand gevonden – gebruik beste schatting)"
    )

    return f"""Je bent een productspecialist en SEO-expert voor {config.BEDRIJF_NAAM}. {config.BEDRIJF_NAAM} is {config.BEDRIJF_OMSCHRIJVING}.

PRODUCT:
Artikelnummer: {artikelnr}
Omschrijving: {omschrijving}
Merk: {merk}
Extra: {extra_info or "geen"}

TAAK: Genereer alle KING-velden voor dit product als JSON.

ATTRIBUUTSET V4 – kies de BEST passende code uit deze lijst:
{attrib_sectie}

CATEGORIEINDELING – kies 1 tot 3 best passende IDs uit deze lijst:
{cat_sectie}

TAALCODE REGELS:
1N1 (meta titel):
  - Structuur: [MERK] [VOLLEDIGE PRODUCTNAAM + SPECIFICATIES] | [USP] | FOV
  - MINIMAAL 60, MAXIMAAL 70 tekens – tel exact
  - Gebruik bij voorkeur deze USP: "{usp}"
  - Merk altijd vooraan, GEEN CTA-woorden (kopen/bestellen)
  - Eindigt exact op " | {config.BEDRIJF_NAAM}"

1N2 (meta description):
  - MINIMAAL 150, MAXIMAAL 160 tekens
  - Bevat merk + productnaam + minimaal 2 concrete eigenschappen
  - Sluit af met "bij {config.BEDRIJF_NAAM}" of variatie
  - GEEN "u" of "uw"

1NT (webtitel):
  - Volledige productnaam, leesbaar voor website
  - Zonder merk-prefix als merk al duidelijk is

1NF (factuuromschrijving):
  - Maximaal 50 tekens
  - Korte, duidelijke naam voor op de factuur

1NH (URL-slug):
  - Alles lowercase, woorden gescheiden door koppeltekens
  - Geen special characters, geen merk-prefix

1NL (lange omschrijving HTML):
  Gebruik EXACT deze structuur. Kopieer de style-tags letterlijk, vul de placeholders in.

  <style>
    .tab-style {{ display: block; padding: 1rem; border: 1px solid #ccc; background: white; margin-bottom: 1.5rem; border-radius: 5px; }} .tab-style h2 {{ margin-top: 0; color: #103a5d; }} .tab-style strong {{ color: #103a5d; }}
  </style>
  <style>
    .tabs {{ margin-top: 2rem; }} .tabs input[type="radio"] {{ display: none; }} .tabs label {{ padding: 0.5rem 1rem; background: #eee; margin-right: 0.2rem; cursor: pointer; border-top-left-radius: 5px; border-top-right-radius: 5px; font-weight: bold; color: #103a5d; }} .tabs label:hover {{ background: #ddd; }} .tabs .tab-content {{ display: none; border: 1px solid #ccc; padding: 1rem; background: white; border-top: none; border-radius: 0 0 5px 5px; }} .tabs input[type="radio"]:checked + label {{ background: #103a5d; color: white; }} .tabs input[type="radio"]:checked + label + .tab-content {{ display: block; }}
  </style>

  <h2>[Merk] [Productnaam] – [Korte ondertitel/toepassing]</h2>
  <p><strong>[Één zin: voor wie is dit product / hoofdtoepassing]</strong></p>
  <p>[Inleidende beschrijving in 2-3 zinnen over het product]</p>
  <ul>
    <li><span style="color:#c3a923;"><i class="check"></i></span> <strong>[Voordeel 1]</strong> – [toelichting]</li>
    <li><span style="color:#c3a923;"><i class="check"></i></span> <strong>[Voordeel 2]</strong> – [toelichting]</li>
    <li><span style="color:#c3a923;"><i class="check"></i></span> <strong>[Voordeel 3]</strong> – [toelichting]</li>
    <li><span style="color:#c3a923;"><i class="check"></i></span> <strong>[Voordeel 4]</strong> – [toelichting]</li>
    <li><span style="color:#c3a923;"><i class="check"></i></span> <strong>[Voordeel 5]</strong> – [toelichting]</li>
  </ul>
  <section style="margin-top: 3rem;">
    <h3>Waarom kiezen voor [productnaam]?</h3>
    <div style="display: flex; flex-wrap: wrap; gap: 1rem;">
      <div style="flex: 1; background-color: #f5f5f5; padding: 1rem;"><strong>[Kenmerk 1]</strong><p>[Toelichting]</p></div>
      <div style="flex: 1; background-color: #f5f5f5; padding: 1rem;"><strong>[Kenmerk 2]</strong><p>[Toelichting]</p></div>
      <div style="flex: 1; background-color: #f5f5f5; padding: 1rem;"><strong>[Kenmerk 3]</strong><p>[Toelichting]</p></div>
    </div>
  </section>
  <div style="display:flex;flex-wrap:wrap;gap:1rem;margin-top:2rem;">
    <div style="flex:1;background:#f5f5f5;padding:1rem;"><strong>Verbruik</strong><p>[verbruik alleen als bekend uit productdata, anders weglaten]</p></div>
    <div style="flex:1;background:#f5f5f5;padding:1rem;"><strong>Droogtijd</strong><p>[droogtijd alleen als bekend uit productdata, anders weglaten]</p></div>
    <div style="flex:1;background:#f5f5f5;padding:1rem;"><strong>Toepassing</strong><p>[verwerkingswijze alleen als bekend uit productdata, anders weglaten]</p></div>
  </div>
  BELANGRIJK: laat een heel spec-blok weg als de waarde niet uit de productdata bekend is. Verzin NOOIT verbruik, droogtijd of technische specs.
  <div style="display:flex;flex-wrap:wrap;gap:1rem;margin-top:1rem;">
    <div style="flex:1;background:#f5f5f5;padding:1rem;">
      <strong>Waar toepassen?</strong>
      <ul><li>[toepassingsgebied 1]</li><li>[toepassingsgebied 2]</li><li>[toepassingsgebied 3]</li></ul>
    </div>
    <div style="flex:1;background:#f5f5f5;padding:1rem;">
      <strong>Gebruikstips</strong>
      <ul><li>[tip 1]</li><li>[tip 2]</li><li>[tip 3]</li></ul>
    </div>
  </div>
  <div class="tabs">
    <input type="radio" id="tab1_{artikelnr}" name="tabgroup_{artikelnr}" checked />
    <label for="tab1_{artikelnr}">Technische gegevens</label>
    <div class="tab-content"><p><strong>[spec label]:</strong> [waarde]</p><p><strong>[spec label]:</strong> [waarde]</p></div>
    <input type="radio" id="tab2_{artikelnr}" name="tabgroup_{artikelnr}" />
    <label for="tab2_{artikelnr}">Kenmerken</label>
    <div class="tab-content"><p>- [kenmerk 1]</p><p>- [kenmerk 2]</p><p>- [kenmerk 3]</p></div>
    <input type="radio" id="tab3_{artikelnr}" name="tabgroup_{artikelnr}" />
    <label for="tab3_{artikelnr}">Verwerkingstips</label>
    <div class="tab-content"><p>- [tip 1]</p><p>- [tip 2]</p><p>- [tip 3]</p></div>
  </div>

  Regels:
  - Minimaal 200 woorden totaal
  - Vul ALLE placeholders in met product-specifieke inhoud
  - Gebruik het artikelnummer {artikelnr} letterlijk in de tab id/name attributen
  - Schrijf in het Nederlands, geen "u"/"uw"
  - Geen verwijzing naar FOV aan het einde
  - NOOIT technische specs verzinnen (verbruik, droogtijd, treksterkte, temperatuur etc.) — alleen opnemen als ze expliciet in de productdata staan. Laat het blok of de waarde weg als de info ontbreekt.

GEEF ALLEEN geldige JSON (geen markdown, geen uitleg):
{{
  "attribuutset_code":       "...",
  "attribuutset_confidence": 0.85,
  "attribuutset_label":      "...",
  "categorie_ids":           "...,",
  "categorie_confidence":    0.80,
  "1N1": "...",
  "1N2": "...",
  "1NT": "...",
  "1NF": "...",
  "1NH": "...",
  "1NL": "..."
}}""".strip()


# ─────────────────────────────────────────────
#  OPENAI AANROEP
# ─────────────────────────────────────────────

def vraag_openai(prompt: str, client: OpenAI, model: str) -> dict:
    response = client.chat.completions.create(
        model=model,
        temperature=0.3,
        messages=[
            {
                "role": "system",
                "content": (
                    f"Je bent een {config.BEDRIJF_NAAM} productexpert en SEO-specialist. "
                    "Je schrijft altijd in het Nederlands. "
                    "Je telt tekens exact en houdt je aan alle lengteregels. "
                    "Antwoord ALLEEN in geldig JSON zonder markdown."
                )
            },
            {"role": "user", "content": prompt}
        ],
        max_tokens=2000,
        response_format={"type": "json_object"}
    )
    tekst = response.choices[0].message.content.strip()
    if tekst.startswith("```"):
        tekst = re.sub(r"^```json?\s*|\s*```$", "", tekst).strip()
    return json.loads(tekst)


def _fallback_waarden(product: dict) -> dict:
    omschrijving = product.get("Omschrijving", "")
    merk         = product.get("Merk", "FOV")
    slug = re.sub(r"[^a-z0-9]+", "-", f"{merk}-{omschrijving}".lower()).strip("-")
    return {
        "attribuutset_code":       "e_gereedschap_overig",
        "attribuutset_confidence": 0.10,
        "attribuutset_label":      "Handmatig invullen",
        "categorie_ids":           "",
        "categorie_confidence":    0.10,
        "1N1": f"{merk} {omschrijving} | {config.BEDRIJF_NAAM}"[:70],
        "1N2": f"{merk} {omschrijving}. Bestel bij {config.BEDRIJF_NAAM}.",
        "1NT": omschrijving,
        "1NF": f"{merk} {omschrijving}"[:50],
        "1NH": slug[:80],
        "1NL": f"<p>{omschrijving}</p>",
    }


def _conf_kleur(c: float) -> str:
    if c >= 0.75: return "D4EDDA"
    if c >= 0.55: return "FFF3CD"
    return "F8D7DA"


# ─────────────────────────────────────────────
#  VERWERK EEN PRODUCT
# ─────────────────────────────────────────────

def verwerk_product(
    product: dict,
    idx: int,
    attribuutsets: list[dict],
    categorieen: list[dict],
    client: OpenAI,
    model: str,
    log_func,
) -> dict:
    artikelnr = product.get("Artikelnummer", f"#{idx+1}")
    pagina    = {"tekst": "", "pib_url": "", "vib_url": ""}

    url = str(product.get("URL_leverancier", "") or "").strip()
    if url.startswith("http"):
        log_func(f"  Pagina ophalen: {url[:70]}", "info")
        pagina = haal_leverancier_pagina(
            url,
            merk=str(product.get("Merk") or ""),
            omschrijving=str(product.get("Omschrijving") or ""),
        )
        if pagina["pib_pad"]:
            log_func(f"  PIB opgeslagen: {pagina['pib_pad']}", "success")
        if pagina["vib_pad"]:
            log_func(f"  VIB opgeslagen: {pagina['vib_pad']}", "success")

    try:
        prompt = maak_prompt(product, attribuutsets, categorieen, idx, pagina["tekst"])
        ai     = vraag_openai(prompt, client, model)

        # Vervang em/en-dashes door gewoon koppelteken (voorkomt encoding-problemen in KING)
        for k, v in ai.items():
            if isinstance(v, str):
                ai[k] = v.replace("—", " - ").replace("–", " - ")

        # Controleer 1N1 lengte en corrigeer lichte overschrijding
        n1 = ai.get("1N1", "")
        if len(n1) > 70:
            ai["1N1"] = n1[:67].rsplit(" ", 1)[0] + " | FOV"

        log_func(f"✓ {artikelnr} – 1N1: {len(ai.get('1N1',''))}t, 1N2: {len(ai.get('1N2',''))}t", "success")
        return {**product, "_ai": ai, "_pagina": pagina, "_status": "success"}

    except Exception as e:
        log_func(f"✗ {artikelnr} – {str(e)[:80]}", "error")
        return {**product, "_ai": _fallback_waarden(product), "_pagina": pagina, "_status": "error"}


# ─────────────────────────────────────────────
#  EXCEL OPBOUWEN
# ─────────────────────────────────────────────

def bouw_excel(
    producten: list[dict],
    output_pad: Path,
    client: OpenAI,
    model: str,
    attribuutsets: list[dict],
    categorieen: list[dict],
    log_func,
    progress_func,
) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KING Import"

    hdr_font  = Font(name="Arial", bold=True, color=config.CLR_WIT, size=10)
    hdr_fill  = PatternFill("solid", fgColor=config.CLR_BLAUW)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font  = Font(name="Arial", size=9)
    data_align = Alignment(vertical="top", wrap_text=True)
    rand = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    for ci, kol in enumerate(KOLOMMEN, 1):
        c = ws.cell(row=1, column=ci, value=kol)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = hdr_align; c.border = rand
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"

    verwerkte_producten = []

    for i, product in enumerate(producten):
        progress_func(i + 1, len(producten))
        result = verwerk_product(product, i, attribuutsets, categorieen, client, model, log_func)
        verwerkte_producten.append(result)

        if i < len(producten) - 1:
            time.sleep(0.5)

    for ri, result in enumerate(verwerkte_producten, 2):
        ai  = result.get("_ai", _fallback_waarden(result))
        conf = float(ai.get("attribuutset_confidence", 0.10))
        cf   = PatternFill("solid", fgColor=_conf_kleur(conf))

        pagina  = result.get("_pagina", {})

        # Fix 1: placeholder-waarden uit template negeren
        def _echte_pad(waarde):
            v = str(waarde or "").strip()
            return "" if v.startswith(r"C:\paden") or "leverancierspagina" in v else v

        def _lokaal_pad(waarde, prefix: str) -> str:
            """Als waarde een URL is, download dan naar lokale mapstructuur."""
            v = _echte_pad(waarde)
            if not v or not v.startswith("http"):
                return v
            merk_c = _schone_bestandsnaam(str(result.get("Merk") or ""))
            omschr = str(result.get("Omschrijving") or "")
            if omschr.lower().startswith(merk_c.lower()):
                omschr = omschr[len(merk_c):].strip()
            omschr_c = _schone_bestandsnaam(omschr).strip(" -")
            merk_dir = config.PIB_BASE_DIR / merk_c / "NL"
            bestand = merk_dir / f"{prefix} {merk_c} {omschr_c}-NL.pdf"
            return str(bestand) if _download_pdf(v, bestand) else v

        pib_pad = _lokaal_pad(result.get("Productinformatieblad_pad"), "PIB") or pagina.get("pib_pad", "")
        vib_pad = _lokaal_pad(result.get("Productveiligheidsblad_pad"), "VIB") or pagina.get("vib_pad", "")
        afbeeldingen = pagina.get("afbeeldingen", [])

        waarden = {
            "Artikelnummer":                        result.get("Artikelnummer", ""),
            "Eenheid":                              result.get("Eenheid", "Stuk"),
            "Zoekcode":                             result.get("Merk", ""),
            "Omschrijving":                         result.get("Omschrijving", ""),
            "Opbrengstgroep":                       result.get("Opbrengstgroep", ""),
            "WebArtikel":                           config.WEBART,
            "TekstOpFactuur":                       result.get("Omschrijving", ""),
            "AfbeeldingKlein":                      afbeeldingen[0] if afbeeldingen else "",
            "AfbeeldingGroot":                      afbeeldingen[0] if afbeeldingen else "",
            "Leveranciernummer":                    result.get("Leveranciernummer", ""),
            "Leveranciernaam":                      result.get("Merk", ""),
            "ArtikelOmschrijvingLeverancier":       result.get("Omschrijving", ""),
            "ArtikelNummerBijLeverancier":          result.get("ArtikelNummerBijLeverancier", "") or "",
            "EanCode":                              result.get("EanCode", "") or "",
            "VR_ART_Magentotype":                   "Simpel",
            "VR_ART_Zichtbaarheid":                 "Catalogus, zoeken",
            "VR_ART_Actief_in_shop":                config.ACTIEF_IN_SHOP,
            "VR_ART_F-Merk":                        result.get("Merk", ""),
            "VR_ART_Extra_afbeelding_1":            afbeeldingen[1] if len(afbeeldingen) > 1 else "",
            "VR_ART_Extra_afbeelding_2":            afbeeldingen[2] if len(afbeeldingen) > 2 else "",
            "VR_ART_Extra_afbeelding_3":            afbeeldingen[3] if len(afbeeldingen) > 3 else "",
            "VR_ART_Extra_afbeelding_4":            afbeeldingen[4] if len(afbeeldingen) > 4 else "",
            "VR_ART_Extra_afbeelding_5":            afbeeldingen[5] if len(afbeeldingen) > 5 else "",
            "VR_ART_Productinformatieblad_NL":      pib_pad,
            "VR_ART_Productveiligheidsblad_NL_A":   vib_pad,
            "VR_ART_Productveiligheidsblad_NL_B":   "",
            "VR_ART_Productveiligheidsblad_NL_C":   "",
            "VR_ART_Productveiligheidsblad_ENG_A":  "",
            "VR_ART_Productveiligheidsblad_ENG_B":  "",
            "VR_ART_Productveiligheidsblad_ENG_C":  "",
            "VR_ART_Productveiligheidsblad_DE_A":   "",
            "VR_ART_Productveiligheidsblad_DE_B":   "",
            "VR_ART_Productveiligheidsblad_DE_C":   "",
            "VR_ART_Explosietekening_UNI":          "",
            "VR_ART_Producthandleiding_NL":         "",
            "VR_ART_Attribuutset_V4":               ai.get("attribuutset_code", ""),
            "1NT_FOV_NL_TITLE_WEB":                 ai.get("1NT", ""),
            "1NF_FOV_NL_TITLE_FACTUUR":             ai.get("1NF", ""),
            "1NH_FOV_NL_URL":                       ai.get("1NH", ""),
            "1N1_FOV_NL_META_DATA_1":               ai.get("1N1", ""),
            "1N2_FOV_NL_META_DATA_2":               ai.get("1N2", ""),
            "1NL_FOV_NL_LANGE_OMSCHRIJVING":        ai.get("1NL", ""),
            "Attribuutset_Confidence_%":            f"{conf*100:.0f}%",
            "Attribuutset_Label":                   ai.get("attribuutset_label", ""),
            "VR_ART_Webcategorie_ID._V2":           ai.get("categorie_ids", ""),
            "Webcategorie_Confidence_%":            f"{float(ai.get('categorie_confidence', 0))*100:.0f}%",
        }

        for ci, kol in enumerate(KOLOMMEN, 1):
            cel = ws.cell(row=ri, column=ci, value=waarden.get(kol, ""))
            cel.font = data_font; cel.alignment = data_align; cel.border = rand
            if kol in HANDMATIG:
                cel.fill = PatternFill("solid", fgColor=config.CLR_ORANJE)
            elif kol in AI_VELDEN:
                cel.fill = PatternFill("solid", fgColor=config.CLR_GROEN)
            elif kol in MAPPING_VELDEN:
                cel.fill = cf

    # Kolombreedte
    breedte = {
        "1N1_FOV_NL_META_DATA_1": 44, "1N2_FOV_NL_META_DATA_2": 58,
        "1NF_FOV_NL_TITLE_FACTUUR": 34, "1NH_FOV_NL_URL": 38,
        "1NL_FOV_NL_LANGE_OMSCHRIJVING": 88, "1NT_FOV_NL_TITLE_WEB": 28,
        "Omschrijving": 34, "Artikelnummer": 18,
        "VR_ART_Productinformatieblad_NL": 64, "VR_ART_Productveiligheidsblad_NL_A": 64,
        "VR_ART_Attribuutset_V4": 30, "Attribuutset_Confidence_%": 14,
        "Attribuutset_Label": 26, "VR_ART_Webcategorie_ID._V2": 24,
        "Webcategorie_Confidence_%": 14,
    }
    for ci, kol in enumerate(KOLOMMEN, 1):
        ws.column_dimensions[get_column_letter(ci)].width = breedte.get(kol, 17)

    _voeg_legenda_toe(wb, attribuutsets)
    wb.save(output_pad)
    log_func(f"Klaar — {len(producten)} producten → {output_pad.name}", "success")
    return output_pad


def _voeg_legenda_toe(wb: openpyxl.Workbook, attribuutsets: list[dict]):
    ws2 = wb.create_sheet("Legenda")
    ws2.merge_cells("A1:B1")
    ws2["A1"].value = "Legenda – Product Import"
    ws2["A1"].font  = Font(name="Arial", bold=True, size=13, color=config.CLR_WIT)
    ws2["A1"].fill  = PatternFill("solid", fgColor=config.CLR_BLAUW)
    ws2.row_dimensions[1].height = 24

    items = [
        ("KLEURCODERING", None),
        ("Oranje", "Handmatig invullen (prijs, leveranciernummer)"),
        ("Lichtgroen", "AI-gegenereerde teksten – altijd controleren vóór publicatie"),
        ("Groen (mapping)", "Hoge confidence ≥ 75% – waarschijnlijk correct"),
        ("Geel (mapping)", "Matige confidence 55-74% – controleer de waarde"),
        ("Rood (mapping)", "Lage confidence < 55% – handmatig aanpassen"),
        ("", None),
        ("INPUTKOLOMMEN (verplicht)", None),
        ("Artikelnummer", "KING artikelnummer"),
        ("Omschrijving", "Productnaam / omschrijving"),
        ("Merk", "F-merk (bijv. Jotun, Hempel, Sika)"),
        ("", None),
        ("INPUTKOLOMMEN (optioneel)", None),
        ("EanCode", "EAN/barcode"),
        ("Eenheid", "Stuk, Liter, etc."),
        ("Gewicht", "In kg"),
        ("VerkoopPrijsExBTW / AdviesPrijsExBTW", "Prijzen (oranje = handmatig invullen)"),
        ("Leveranciernummer", "KING leveranciernummer"),
        ("ArtikelNummerBijLeverancier", "Bestelnummer bij leverancier"),
        ("Productinformatieblad_pad", r"K:\pad\naar\PIB.pdf"),
        ("Productveiligheidsblad_pad", r"K:\pad\naar\VIB.pdf"),
        ("", None),
        ("AI-VELDEN (lichtgroen)", None),
        ("1N1", "Meta titel: Merk Product | USP | [BEDRIJF] (60-70 tekens)"),
        ("1N2", "Meta description (150-160 tekens)"),
        ("1NT", "Webtitel (zichtbare naam op website)"),
        ("1NF", "Factuurnaam (max 50 tekens)"),
        ("1NH", "URL-slug (alles na fov.nl/)"),
        ("1NL", "Lange HTML omschrijving"),
        ("", None),
        ("AUTO-MAPPING (kleurgecodeerd)", None),
        ("VR_ART_Attribuutset_V4", "Uit attribuutset_v4.xlsx – zie referentiedata map"),
        ("VR_ART_Webcategorie_ID._V2", "Uit categorieindeling.xlsx – kommagescheiden IDs"),
        ("Confidence_%", "< 55% = rood = handmatig controleren"),
        ("", None),
        ("VASTE WAARDEN (automatisch)", None),
        ("WebArtikel", "1 (live in shop)"),
        ("VR_ART_Actief_in_shop", "1 (online)"),
        ("VR_ART_F-Merk", "Overgenomen uit Merk kolom"),
    ]

    for i, (k, v) in enumerate(items, 3):
        ck = ws2.cell(row=i, column=1, value=k)
        ck.font = Font(name="Arial", size=9,
                       bold=(v is None and bool(k)),
                       color=config.CLR_BLAUW if v is None and k else "000000")
        if v is None and k:
            ck.fill = PatternFill("solid", fgColor="E8EFF7")
        if v:
            ws2.cell(row=i, column=2, value=v).font = Font(name="Arial", size=9)

    ws2.column_dimensions["A"].width = 40
    ws2.column_dimensions["B"].width = 72

    # Attribuutset referentie op apart tabblad
    if attribuutsets:
        ws3 = wb.create_sheet("Attribuutsets")
        ws3.cell(1, 1, "Code").font = Font(bold=True, name="Arial")
        ws3.cell(1, 2, "Label").font = Font(bold=True, name="Arial")
        for i, a in enumerate(attribuutsets, 2):
            ws3.cell(i, 1, a["code"])
            ws3.cell(i, 2, a["label"])
        ws3.column_dimensions["A"].width = 35
        ws3.column_dimensions["B"].width = 40


# ─────────────────────────────────────────────
#  INVOER INLEZEN
# ─────────────────────────────────────────────

def lees_invoer(pad: Path) -> list[dict]:
    """Lees Excel of CSV invoerbestand naar lijst van dicts."""
    suffix = pad.suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(pad, encoding="utf-8-sig")
    elif suffix in (".xlsx", ".xls"):
        df = pd.read_excel(pad)
    else:
        raise ValueError(f"Niet-ondersteund bestandsformaat: {suffix}")

    import math
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(subset=["Artikelnummer"])
    records = df.to_dict("records")
    return [
        {k: (None if (v is None or (isinstance(v, float) and math.isnan(v))) else v)
         for k, v in row.items()}
        for row in records
    ]


# ─────────────────────────────────────────────
#  STANDALONE GEBRUIK
# ─────────────────────────────────────────────

if __name__ == "__main__":
    from openai import OpenAI

    if not config.OPENAI_API_KEY:
        print("Stel OPENAI_API_KEY in als omgevingsvariabele.")
        raise SystemExit(1)

    bron_bestanden = list(config.BRON_DIR.glob("*.xlsx")) + list(config.BRON_DIR.glob("*.csv"))
    if not bron_bestanden:
        print(f"Geen invoerbestanden gevonden in {config.BRON_DIR}")
        raise SystemExit(1)

    invoer_pad = bron_bestanden[0]
    print(f"Invoer: {invoer_pad.name}")

    producten    = lees_invoer(invoer_pad)
    attribuutsets = laad_attribuutset(config.ATTRIBUUTSET_FILE)
    categorieen   = laad_categorieindeling(config.CATEGORIE_FILE)
    client       = OpenAI(api_key=config.OPENAI_API_KEY)

    print(f"Producten: {len(producten)}")
    print(f"Attribuutsets geladen: {len(attribuutsets)}")
    print(f"Categorieën geladen:   {len(categorieen)}")

    datum      = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_pad = config.OUTPUT_DIR / f"KING_import_{datum}.xlsx"

    totaal = [0]

    def log(msg, t="info"):
        print(msg)

    def progress(h, t):
        if h != totaal[0]:
            totaal[0] = h
            print(f"  [{h}/{t}]")

    bouw_excel(producten, output_pad, client, config.OPENAI_MODEL,
               attribuutsets, categorieen, log, progress)
    print(f"\nOutput: {output_pad}")

