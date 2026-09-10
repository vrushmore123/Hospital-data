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