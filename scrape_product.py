"""Extract product information from Venus Health Care product pages."""

from __future__ import annotations

import argparse
import json
import re
import time
from typing import Any
from urllib.parse import urljoin

from openpyxl import Workbook
from openpyxl.styles import Alignment
import requests
from bs4 import BeautifulSoup, Tag


DEFAULT_URL = "https://www.venushealthcare.net/products/standard-f200-analyzer"
PRODUCT_ROOT = "https://www.venushealthcare.net/products/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; ProductDataExtractor/1.0)"
}


def clean_text(value: str) -> str:
    """Collapse HTML whitespace while preserving readable text."""
    return re.sub(r"\s+", " ", value).strip()


def unique(values: list[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))


def section_items(soup: BeautifulSoup, heading_text: str) -> list[str]:
    """Collect list items until the next heading at the same or higher level."""
    heading = next(
        (
            element
            for element in soup.find_all(re.compile(r"^h[1-6]$"))
            if clean_text(element.get_text(" ")).lower() == heading_text.lower()
        ),
        None,
    )
    if heading is None:
        return []

    level = int(heading.name[1])
    results: list[str] = []
    for sibling in heading.find_all_next():
        if sibling is heading:
            continue
        if isinstance(sibling, Tag) and re.fullmatch(r"h[1-6]", sibling.name or ""):
            if int(sibling.name[1]) <= level:
                break
        if isinstance(sibling, Tag) and sibling.name == "li":
            text = clean_text(sibling.get_text(" "))
            if text:
                results.append(text.lstrip("✓•- "))
    return unique(results)


def json_ld_product(soup: BeautifulSoup) -> dict[str, Any]:
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            data = json.loads(script.string or script.get_text())
        except json.JSONDecodeError:
            continue
        candidates = data if isinstance(data, list) else data.get("@graph", [data])
        for candidate in candidates:
            if isinstance(candidate, dict) and candidate.get("@type") == "Product":
                return candidate
    return {}


def json_ld_business(soup: BeautifulSoup) -> dict[str, Any]:
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            data = json.loads(script.string or script.get_text())
        except json.JSONDecodeError:
            continue
        candidates = data if isinstance(data, list) else data.get("@graph", [data])
        for candidate in candidates:
            if not isinstance(candidate, dict):
                continue
            types = candidate.get("@type", [])
            types = types if isinstance(types, list) else [types]
            if any(value in {"LocalBusiness", "MedicalOrganization"} for value in types):
                return candidate
    return {}


def related_products(soup: BeautifulSoup, page_url: str) -> list[dict[str, str]]:
    products: list[dict[str, str]] = []
    for link in soup.select('a[href*="/products/"]'):
        href = urljoin(page_url, link.get("href", ""))
        name = clean_text(link.get_text(" "))
        if not name or href.rstrip("/") == page_url.rstrip("/"):
            continue
        if not href.startswith(PRODUCT_ROOT):
            continue
        if name.lower().startswith("view details"):
            continue
        products.append({"name": name, "url": href})
    return unique_dicts(products)


def application_tags(soup: BeautifulSoup) -> list[str]:
    tags = [clean_text(tag.get_text(" ")) for tag in soup.select(".application-tags span")]
    return unique(tags)


def brochure_links(soup: BeautifulSoup, page_url: str) -> list[str]:
    links = []
    for link in soup.find_all("a", href=True):
        href = urljoin(page_url, link["href"])
        label = clean_text(link.get_text(" ")).lower()
        if any(term in href.lower() or term in label for term in (".pdf", "brochure", "download brochure")):
            links.append(href)
    return unique(links)


def business_profile(business: dict[str, Any], page_url: str) -> dict[str, Any]:
    address = business.get("address", {})
    if isinstance(address, dict):
        address_text = ", ".join(
            clean_text(str(address.get(key, "")))
            for key in ("streetAddress", "addressLocality", "addressRegion", "postalCode", "addressCountry")
            if address.get(key)
        )
    else:
        address_text = clean_text(str(address))
    email = business.get("email", "")
    phone = business.get("telephone", "")
    if isinstance(email, list):
        email = ", ".join(str(value) for value in email)
    if isinstance(phone, list):
        phone = ", ".join(str(value) for value in phone)
    return {
        "company_name": clean_text(str(business.get("name", ""))),
        "company_profile": clean_text(str(business.get("description", ""))),
        "phone": clean_text(str(phone)),
        "email": clean_text(str(email)),
        "website": urljoin(page_url, str(business.get("url", ""))),
        "address": address_text,
    }


def labeled_value(text: str, labels: tuple[str, ...]) -> str:
    pattern = r"(?:" + "|".join(re.escape(label) for label in labels) + r")\s*[:\-]\s*([^\n|;]+)"
    match = re.search(pattern, text, re.IGNORECASE)
    return clean_text(match.group(1)) if match else ""


def unique_dicts(values: list[dict[str, str]]) -> list[dict[str, str]]:
    seen: set[tuple[str, str]] = set()
    result = []
    for value in values:
        key = (value["name"], value["url"])
        if key not in seen:
            seen.add(key)
            result.append(value)
    return result


def extract_product(url: str, timeout: int = 30) -> dict[str, Any]:
    response = requests.get(url, headers=HEADERS, timeout=timeout)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    structured = json_ld_product(soup)
    business = business_profile(json_ld_business(soup), url)

    title = clean_text(structured.get("name", ""))
    if not title:
        title_node = soup.find("h1")
        title = clean_text(title_node.get_text(" ")) if title_node else ""

    description = clean_text(structured.get("description", ""))
    if not description:
        description_node = soup.select_one('meta[name="description"]')
        description = clean_text(description_node.get("content", "")) if description_node else ""

    image = structured.get("image", "")
    if isinstance(image, list):
        image = image[0] if image else ""
    if image:
        image = urljoin(url, str(image))

    tests = application_tags(soup)
    if not tests:
        tests = section_items(soup, "Designed to support")
    manufacturer = structured.get("brand", {})
    manufacturer = manufacturer.get("name", "") if isinstance(manufacturer, dict) else manufacturer
    usage = description or clean_text(structured.get("category", ""))
    visible_text = soup.get_text("\n", strip=True)
    importer = labeled_value(visible_text, ("Indian importer", "Importer company", "Importer"))
    country = labeled_value(visible_text, ("Country of manufacture", "Country of origin", "Made in"))

    return {
        "url": url,
        "name": title,
        "description": description,
        "image": image,
        "brand": structured.get("brand", {}).get("name") if isinstance(structured.get("brand"), dict) else structured.get("brand", ""),
        "features": section_items(soup, "At a glance"),
        "applications": section_items(soup, "Designed to support"),
        "related_products": related_products(soup, url),
        "manufacturer_company": clean_text(str(manufacturer)),
        "importer_company": importer,
        "country_of_manufacture": country,
        "supplier_company": business["company_name"],
        "company_name": business["company_name"],
        "company_profile": business["company_profile"],
        "phone": business["phone"],
        "email": business["email"],
        "website": business["website"],
        "company_address": business["address"],
        "brochure_links": brochure_links(soup, url),
        "machine_usage": usage,
        "tests_performed": tests,
    }


def product_links(soup: BeautifulSoup, page_url: str) -> list[str]:
    """Return unique product-page URLs found on a page."""
    urls = []
    for link in soup.find_all("a", href=True):
        href = urljoin(page_url, link["href"]).split("#", 1)[0].rstrip("/")
        if href.startswith(PRODUCT_ROOT) and href != PRODUCT_ROOT.rstrip("/"):
            urls.append(href)
    return list(dict.fromkeys(urls))


def discover_product_urls(timeout: int = 30) -> list[str]:
    """Discover product pages from the product index and linked product pages."""
    session = requests.Session()
    session.headers.update(HEADERS)
    pending = [PRODUCT_ROOT.rstrip("/")]
    discovered: set[str] = set()

    while pending:
        page_url = pending.pop(0)
        if page_url in discovered:
            continue
        response = session.get(page_url, timeout=timeout)
        response.raise_for_status()
        discovered.add(page_url)
        soup = BeautifulSoup(response.text, "html.parser")
        for product_url in product_links(soup, page_url):
            if product_url not in discovered and product_url not in pending:
                pending.append(product_url)

    return sorted(url for url in discovered if url != PRODUCT_ROOT.rstrip("/"))


def save_products_to_excel(products: list[dict[str, Any]], output_path: str) -> None:
    """Save scraped products as one row per product in an Excel workbook."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Products"
    columns = [
        "name", "brand", "description", "url", "image", "features", "applications", "related_products",
        "manufacturer_company", "importer_company", "country_of_manufacture", "supplier_company",
        "company_name", "company_profile", "phone", "email", "website", "company_address",
        "brochure_links", "machine_usage", "tests_performed",
    ]
    sheet.append(columns)

    for product in products:
        row = dict(product)
        row["features"] = "\n".join(row.get("features", []))
        row["applications"] = "\n".join(row.get("applications", []))
        row["related_products"] = "\n".join(
            f"{item['name']}: {item['url']}" for item in row.get("related_products", [])
        )
        row["brochure_links"] = "\n".join(row.get("brochure_links", []))
        row["tests_performed"] = "\n".join(row.get("tests_performed", []))
        sheet.append([row.get(column, "") for column in columns])

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    widths = {"A": 30, "B": 20, "C": 65, "D": 55, "E": 55, "F": 45, "G": 45, "H": 65}
    widths.update({column: 35 for column in ("I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")})
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    workbook.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("url", nargs="?", default=DEFAULT_URL, help="Product page URL")
    parser.add_argument("-o", "--output", help="Write JSON to this file instead of stdout")
    parser.add_argument("--all", action="store_true", help="Scrape every product discovered on the website")
    parser.add_argument("--excel", help="Write scraped products to an Excel workbook")
    parser.add_argument("--delay", type=float, default=0.2, help="Seconds to wait between product requests")
    args = parser.parse_args()

    if args.all:
        urls = discover_product_urls()
        products = []
        errors = []
        for index, url in enumerate(urls, start=1):
            try:
                print(f"[{index}/{len(urls)}] Scraping {url}")
                products.append(extract_product(url))
            except requests.RequestException as error:
                errors.append({"url": url, "error": str(error)})
            if index < len(urls):
                time.sleep(max(args.delay, 0))

        output = json.dumps({"products": products, "errors": errors}, indent=2, ensure_ascii=False)
        output_path = args.output or "products.json"
        with open(output_path, "w", encoding="utf-8") as file:
            file.write(output + "\n")
        if args.excel:
            save_products_to_excel(products, args.excel)
        print(f"Scraped {len(products)} products; {len(errors)} errors")
        print(f"JSON: {output_path}")
        if args.excel:
            print(f"Excel: {args.excel}")
    else:
        product = extract_product(args.url)
        output = json.dumps(product, indent=2, ensure_ascii=False)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as file:
                file.write(output + "\n")
        else:
            print(output)
        if args.excel:
            save_products_to_excel([product], args.excel)


if __name__ == "__main__":
    main()