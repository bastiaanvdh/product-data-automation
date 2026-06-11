"""
Configuratie — Product Import Generator
"""

import os
from pathlib import Path

BASE_DIR      = Path(__file__).resolve().parent
BRON_DIR      = BASE_DIR / "bron"
OUTPUT_DIR    = BASE_DIR / "output"
REF_DIR       = BASE_DIR / "referentiedata"
UPLOAD_DIR    = BASE_DIR / "uploads"
TEMPLATES_DIR = BASE_DIR / "templates"

for _d in [BRON_DIR, OUTPUT_DIR, REF_DIR, UPLOAD_DIR, TEMPLATES_DIR]:
    _d.mkdir(exist_ok=True)

# OpenAI
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL   = "gpt-4o-mini"
MAX_WORKERS    = 5
SAVE_ELKE_N    = 10

# -- Bedrijf / webshop instellingen -----------------------------------------
# Pas aan voor jouw situatie of stel in via omgevingsvariabelen
BEDRIJF_NAAM         = os.getenv("BEDRIJF_NAAM", "MijnShop")
BEDRIJF_OMSCHRIJVING = os.getenv("BEDRIJF_OMSCHRIJVING", "een Nederlandse webshop")
SHOP_DOMEIN          = os.getenv("SHOP_DOMEIN", "mijnshop.nl")

# -- Testmodus ---------------------------------------------------------------
# True  -> schrijft naar lokale test_output/ map (geen NAS, geen netwerkschijf)
# False -> schrijft naar productie-paden
TEST_MODUS = False

_TEST_DIR = BASE_DIR / "test_output"

# -- Opslag: Productinformatiebladen -----------------------------------------
# Structuur: {PIB_BASE_DIR}/{Merk}/NL/PIB {Merk} {Omschrijving}-NL.pdf
# Stel in via omgevingsvariabele PIB_BASE_DIR of pas het pad hieronder aan
_pib_prod    = Path(os.getenv("PIB_BASE_DIR", r"\\server\share\Productinformatiebladen"))
PIB_BASE_DIR = (_TEST_DIR / "Productinformatiebladen") if TEST_MODUS else _pib_prod

# -- Opslag: Productafbeeldingen ---------------------------------------------
# Structuur: {AFBEELDING_NAS_PAD}/{Merk}/{Merk} {Omschrijving}.jpg
# Stel in via omgevingsvariabele AFBEELDING_NAS_PAD of pas het pad hieronder aan
AFBEELDING_NAS_PAD = Path(os.getenv("AFBEELDING_NAS_PAD", r"\\server\share\Productafbeeldingen"))

# KING vaste waarden
WEBART         = "1"
ACTIEF_IN_SHOP = "1"
SHOP           = "SHOP"

# -- Referentiebestanden -----------------------------------------------------
# Zoekt eerst in referentiedata/, dan op het geconfigureerde netwerkpad
_attrib_local   = REF_DIR / "attribuutset_v4.xlsx"
_attrib_netwerk = Path(os.getenv("ATTRIBUUTSET_PAD", r"\\server\share\attribuutset_v4.xlsx"))
ATTRIBUUTSET_FILE = _attrib_local if _attrib_local.exists() else _attrib_netwerk

_cat_local   = REF_DIR / "categorieindeling.xlsx"
_cat_netwerk = Path(os.getenv("CATEGORIEINDELING_PAD", r"\\server\share\categorieindeling.xlsx"))
CATEGORIE_FILE = _cat_local if _cat_local.exists() else _cat_netwerk

# -- USPs voor meta titels ---------------------------------------------------
SHOP_USPS = [
    "Morgen in huis",
    "Gratis verzending",
    "Betaal op rekening",
    "Scherpe prijzen",
    "Persoonlijk advies",
    "Direct leverbaar",
    "Professioneel",
]

# -- Excel kleuren -----------------------------------------------------------
CLR_BLAUW  = "002B5C"
CLR_TEAL   = "00A3A1"
CLR_WIT    = "FFFFFF"
CLR_BG     = "F5F7FA"
CLR_ORANJE = "FFE5D0"
CLR_GROEN  = "E6F7F7"
