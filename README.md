# Venus Health Care product scraper

## Setup

Install Python 3.10 or newer, then run:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

## Extract the supplied product

```powershell
py scrape_product.py
```

Save the result as JSON:

```powershell
py scrape_product.py -o standard-f200.json
```

Scrape every product linked from the website and save JSON plus Excel:

```powershell
py -m pip install -r requirements.txt
py scrape_product.py --all -o products.json --excel products.xlsx
```

The Excel workbook has one row per product and preserves the original columns for product name, brand, description, URLs, features, applications, and related products. It also includes manufacturer/importer/origin fields, supplier/company profile and contact details, brochure links, machine usage, and tests performed. The crawler follows product links from the product index and linked product pages, removes duplicates, and reports failed URLs in `products.json`.

Blank importer, country-of-manufacture, and brochure fields mean that the website did not state those details; the scraper does not invent them.

## Multi-source medical equipment collection

`medical_equipment_scraper.py` creates one JSON file per supplied listing URL. It follows product-shaped links, extracts only stated values, keeps the product source URL and source website, and removes duplicate product/company pairs. It does not treat a brand as a manufacturer, and it leaves importer and contact fields blank unless they are explicitly present.

Run it with the five supplied sources:

```powershell
py medical_equipment_scraper.py `
	"https://www.tradeindia.com/manufacturers/clinical-laboratory-equipment.html" `
	"https://www.tradeindia.com/manufacturers/medical-laboratory-equipment.html" `
	"https://www.medicalexpo.com/medical-manufacturer/eeg-system-2716.html" `
	"https://www.tradeindia.com/manufacturers/hospital-laboratory-equipment.html" `
	"https://www.tradeindia.com/manufacturers/scientific-lab-equipment.html" `
	--output-dir source_outputs --limit 100
```

This writes `source_01.json` through `source_05.json`. Marketplace results are supporting sources only; for verified manufacturer phone, email, address, importer, and brochure values, add the official manufacturer or Indian distributor product URL as another input. HTTP failures are retained in each file's `errors` array. Use `--timeout` to bound slow or blocked sources.

## Indian manufacturer sites

The same script also scrapes manufacturers' own sites directly, which state phone, email, and address themselves (via schema.org data or the page footer) rather than gating them behind a marketplace form:

```powershell
py medical_equipment_scraper.py `
	"https://www.trivitron.com/products/in-vitro-diagnostics/clia" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/clinical-chemistry" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/elisa" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/genomics-ngs" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/hematology" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/histopathology" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/hplc" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/molecular-diagnostics" `
	"https://www.trivitron.com/products/in-vitro-diagnostics/point-of-care-test" `
	"https://accurex.net/" `
	"https://accurex.net/product-categories/" `
	"https://biosystems.in/clinical-analysis/" `
	--output-dir source_outputs_manufacturers --limit 100
```

This writes `source_01.json` through `source_12.json` into `source_outputs_manufacturers/`, covering three Indian in-vitro-diagnostics manufacturers:

| Manufacturer | Website | Products found |
| --- | --- | --- |
| Trivitron Healthcare | trivitron.com | 62 |
| Accurex Biomedical | accurex.net | 8 |
| BioSystems Diagnostics (Trivitron/BioSystems SA joint venture) | biosystems.in | 15 |

Candidates checked but not included: **Coral Clinical Systems** (coralclinicalsystems.com) has no structured product data and its product names/labels sit in generic HTML the scraper can't reliably tell apart from navigation text, so scraping it produced mislabeled records; it needs bespoke per-site extraction, not attempted here. **Transasia Bio-Medicals / Erba Mannheim** (transasia.co.in) and **Agappe Diagnostics** (agappe.com) load their product listings via JavaScript, so static requests return no product links; they would need Playwright rendering. **Robonik India** has no resolvable standalone website.

## Combined output (all sources in one file)

Each scraper above writes its own file with its own field names. To get every product from every source in one place, with one set of columns:

```powershell
py consolidate_products.py
```

This reads `products.json`, `tradeindia_products.json`, `tradeindia_kanad.json`, `standard-f200.json`, `source_outputs/`, and `source_outputs_manufacturers/`, maps them onto a single schema, removes duplicates, and writes **`all_products.json`** and **`all_products.xlsx`**. The `data_source` column says which scrape each row came from. Re-run it after any scraper run to refresh the combined file.

## TradeIndia clinical analyzers

Scrape the supplied TradeIndia clinical-analyzer search results, including the Kanad product URL, into the same 21-field JSON and Excel schema:

```powershell
py tradeindia_scraper.py --all -o tradeindia_products.json --excel tradeindia_products.xlsx
```

Scrape only the supplied product:

```powershell
py tradeindia_scraper.py -o tradeindia_kanad.json --excel tradeindia_kanad.xlsx
```

The search crawler follows all TradeIndia pagination pages. The current search returned 61 unique products. The Excel output includes the existing fields plus `seller_phone`.

TradeIndia hides seller numbers behind its `VIEW NUMBER` form. To use the optional browser-assisted reveal, install Playwright and provide your own buyer details:

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
py tradeindia_scraper.py --all --reveal-number --buyer-name "Your Name" --buyer-mobile "Your 10-digit mobile" --quantity 1 -o tradeindia_products_with_numbers.json --excel tradeindia_products_with_numbers.xlsx
```

The browser opens TradeIndia's own form for each product. If TradeIndia requests OTP or another verification, complete it in the browser window; the revealed number is saved in `seller_phone`. The scraper does not bypass verification or invent seller numbers.

Use another product page:

```powershell
py scrape_product.py https://www.venushealthcare.net/products/standard-f2400-analyzer
```

The output contains the product name, description, image, brand, feature bullets, supported applications, and related product links. The site states that product specifications are confirmed at quotation, so this script does not invent missing specifications.