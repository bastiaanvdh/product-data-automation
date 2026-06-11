# Product Data Automation

A Flask web application that automates the creation of product import files for ERP systems. Given a simple spreadsheet with basic product info, it generates a complete, ready-to-import Excel file with AI-written content, scraped supplier data, processed product images, and technical document links.

Built to eliminate the repetitive manual work of creating new product listings — what used to take 30–60 minutes per product now takes seconds.

---

## What it does

Upload a spreadsheet with `Article number`, `Description`, `Brand`, and optionally a supplier URL. The tool handles the rest:

1. **Scrapes the supplier page** — extracts product text, finds PIB/VIB technical PDF links
2. **Downloads and stores PDFs** — saves to the correct network folder structure automatically
3. **Downloads product images** — filters out logos and lifestyle photos using filename/alt-text matching, scales to exactly 1600×1600 px with white-background padding (up- and downscaling), saves to NAS
4. **Generates content via OpenAI (gpt-4o-mini)** — web title, invoice title, URL slug, meta title (≤70 chars), meta description (≤160 chars), full HTML product description (~300 words)
5. **Classifies the product** — selects the right attribute set and category IDs from reference files
6. **Outputs a complete import Excel** — 42 columns matching the exact ERP import format, color-coded by confidence and data source

---

## Tech stack

| | |
|---|---|
| **Backend** | Python, Flask |
| **Scraping** | httpx, BeautifulSoup4 |
| **AI** | OpenAI API (gpt-4o-mini) |
| **Image processing** | Pillow |
| **Excel I/O** | openpyxl, pandas |
| **Concurrency** | ThreadPoolExecutor |

---

## Key features

- **Test mode** — single config flag redirects all file writes to a local folder; no risk of touching production storage during development
- **SSL proxy compatibility** — works on corporate networks with intercepting proxies (`httpx verify=False`)
- **URL fallback for PDFs** — if a supplier URL is provided instead of a local path, the PDF is downloaded automatically; falls back gracefully if the download fails
- **Image filter** — checks filename and alt text against brand/product keywords (not the full URL, to avoid domain name false positives)
- **Confidence coloring** — AI classification results are color-coded green/yellow/red based on confidence score
- **Progress polling** — long-running jobs stream status back to the UI; no page refreshes needed

---

## Setup

```bash
pip install flask httpx beautifulsoup4 openai openpyxl pandas pillow
```

Set your OpenAI API key:

```bash
export OPENAI_API_KEY=sk-...
```

Place reference files (attribute set + category tree) in `referentiedata/` or configure network paths in `config.py`.

```bash
python app.py
# → http://localhost:5002
```

---

## Project structure

```
├── app.py                    # Flask routes
├── nieuw_product_generator.py  # Core logic: scraping, AI, image processing, Excel output
├── config.py                 # Paths, API settings, test mode flag
├── templates/
│   └── index.html            # Upload UI + progress view
└── referentiedata/
    ├── attribuutset_v4.xlsx  # Product attribute reference
    ├── categorieindeling.xlsx  # Category tree reference
    └── merk_domeinen.json    # Brand → domain mapping (manually maintained)
```

---

## Output

The generated Excel contains 42 columns in the exact format required for ERP import:

| Column group | Source |
|---|---|
| Fixed values (active, shop, brand) | Hardcoded config |
| Language codes (title, slug, meta, HTML description) | OpenAI |
| Attribute set + category IDs | OpenAI + reference files |
| Image paths (main + up to 5 extras) | Scraped + processed |
| Technical document paths (PIB/VIB) | Scraped or manually provided |
| Commercial fields (price, EAN) | Passed through from input |