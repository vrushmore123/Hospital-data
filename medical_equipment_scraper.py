"""Collect verified medical-equipment product records from supplied source pages."""

from __future__ import annotations

import argparse
import json
import re
import time
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlsplit, urlunsplit

import requests
from bs4 import BeautifulSoup


HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; MedicalEquipmentDataExtractor/1.0)"}
FIELDS = [
    "machine_product_name", "company_name", "manufacturer_name", "importer_name",
    "used_in_which_test", "company_phone_number", "company_email_id",
    "company_location_address", "company_profile_description", "product_brochure_link",
    "product_source_link", "source_website",
]


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def canonical(url: str) -> str:
    parts = urlsplit(url)
    return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), parts.path.rstrip("/"), "", ""))


def empty_record(source_url: str) -> dict[str, str]:
    return {field: "" for field in FIELDS} | {
        "product_source_link": source_url,
        "source_website": urlsplit(source_url).netloc,
    }


def json_ld(soup: BeautifulSoup) -> list[dict[str, Any]]:
    objects: list[dict[str, Any]] = []
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            value = json.loads(script.string or script.get_text())
        except (json.JSONDecodeError, TypeError):
            continue
        candidates = value if isinstance(value, list) else value.get("@graph", [value]) if isinstance(value, dict) else []
        objects.extend(item for item in candidates if isinstance(item, dict))
    return objects


def first_object(objects: list[dict[str, Any]], types: tuple[str, ...]) -> dict[str, Any]:
    for item in objects:
        item_types = item.get("@type", [])
        item_types = item_types if isinstance(item_types, list) else [item_types]
        if any(item_type in types for item_type in item_types):
            return item
    return {}


def address_text(value: Any) -> str:
    if isinstance(value, dict):
        return clean(", ".join(str(value.get(key, "")) for key in (
            "streetAddress", "addressLocality", "addressRegion", "postalCode", "addressCountry"
        ) if value.get(key)))
    return clean(value)


def labeled_text(soup: BeautifulSoup, labels: tuple[str, ...]) -> str:
    pattern = re.compile(r"^(?:" + "|".join(re.escape(label) for label in labels) + r")\s*:?$", re.I)
    for node in soup.find_all(string=pattern):
        parent = node.parent
        if not parent:
            continue
        sibling = parent.find_next_sibling()
        if sibling and clean(sibling.get_text(" ")):
            return clean(sibling.get_text(" "))
        container = parent.parent.get_text(" ") if parent.parent else ""
        value = re.sub(pattern, "", clean(container)).strip(" :|- ")
        if value:
            return value
    return ""


def links_matching(soup: BeautifulSoup, page_url: str, words: tuple[str, ...]) -> list[str]:
    values = []
    for link in soup.find_all("a", href=True):
        href = urljoin(page_url, link["href"])
        label = clean(link.get_text(" ")).lower()
        if any(word in href.lower() or word in label for word in words):
            values.append(href)
    return list(dict.fromkeys(values))


def product_links(soup: BeautifulSoup, page_url: str) -> list[str]:
    base = urlsplit(page_url).netloc
    links = []
    for link in soup.find_all("a", href=True):
        href = urljoin(page_url, link["href"]).split("#", 1)[0]
        parsed = urlsplit(href)
        label = clean(link.get_text(" "))
        product_path = "/products/" in parsed.path or "/prod/" in parsed.path
        if parsed.netloc == base and product_path and label:
            links.append(href)
    return list(dict.fromkeys(links))


def is_product_page(url: str, soup: BeautifulSoup) -> bool:
    path = urlsplit(url).path.lower()
    return "/products/" in path or "/prod/" in path or bool(first_object(json_ld(soup), ("Product",)))


def next_data_company(soup: BeautifulSoup) -> dict[str, str]:
    """Read seller details TradeIndia embeds as page state (__NEXT_DATA__) rather than visible text."""
    script = soup.find("script", id="__NEXT_DATA__", type="application/json")
    if not script or not script.string:
        return {}
    try:
        data = json.loads(script.string)
        details = data["props"]["pageProps"]["initialState"]["product"]["PDP_page"]["PDP_page_res"]["company_details"]
    except (json.JSONDecodeError, KeyError, TypeError):
        return {}
    business_type = (details.get("business_details") or {}).get("business_type") or []
    profile_html = details.get("company_details") or ""
    name = clean(details.get("co_name"))
    return {
        "company_name": name,
        "company_location_address": clean(details.get("address")),
        "company_profile_description": clean(BeautifulSoup(profile_html, "html.parser").get_text(" ")),
        "manufacturer_name": name if "Manufacturer" in business_type else "",
    }


def extract_product(page_url: str, soup: BeautifulSoup) -> dict[str, str]:
    record = empty_record(page_url)
    objects = json_ld(soup)
    product = first_object(objects, ("Product",))
    business = first_object(objects, ("Organization", "LocalBusiness", "MedicalOrganization"))
    brand = product.get("brand", {})
    manufacturer = product.get("manufacturer", {})
    record["machine_product_name"] = clean(product.get("name")) or clean(soup.find("h1").get_text(" ") if soup.find("h1") else "")
    record["manufacturer_name"] = clean(manufacturer.get("name") if isinstance(manufacturer, dict) else manufacturer)
    record["company_name"] = clean(business.get("name"))
    record["company_phone_number"] = clean(business.get("telephone")) or labeled_text(soup, ("Phone", "Telephone", "Mobile", "Contact Number"))
    record["company_email_id"] = clean(business.get("email")) or labeled_text(soup, ("Email", "Email ID", "Email Id"))
    record["company_location_address"] = address_text(business.get("address")) or labeled_text(soup, ("Address", "Location"))
    record["company_profile_description"] = clean(business.get("description")) or clean(product.get("description"))
    record["used_in_which_test"] = labeled_text(soup, ("Used In", "Test", "Application", "Applications", "Clinical Application"))
    record["importer_name"] = labeled_text(soup, ("Indian Importer", "Importer Name", "Importer", "Distributor"))
    brochures = links_matching(soup, page_url, ("brochure", ".pdf"))
    record["product_brochure_link"] = " | ".join(brochures)
    for field, value in next_data_company(soup).items():
        if value and not record[field]:
            record[field] = value
    return record


def scrape_source(source_url: str, session: requests.Session, limit: int, delay: float, timeout: int) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    errors: list[dict[str, str]] = []
    try:
        response = session.get(source_url, timeout=timeout)
        response.raise_for_status()
    except requests.RequestException as error:
        return [], [{"url": source_url, "error": str(error)}]
    soup = BeautifulSoup(response.text, "html.parser")
    urls = ([source_url] if is_product_page(response.url, soup) else []) + product_links(soup, response.url)
    records = []
    for index, url in enumerate(urls[:limit], 1):
        try:
            page = session.get(url, timeout=timeout)
            page.raise_for_status()
            records.append(extract_product(canonical(page.url), BeautifulSoup(page.text, "html.parser")))
        except requests.RequestException as error:
            errors.append({"url": url, "error": str(error)})
        if index < min(len(urls), limit):
            time.sleep(max(delay, 0))
    return records, errors


def deduplicate(records: list[dict[str, str]]) -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for record in records:
        key = (clean(record["machine_product_name"]).lower(), clean(record["company_name"]).lower())
        if key == ("", "") or key in seen:
            continue
        seen.add(key)
        result.append(record)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("urls", nargs="+", help="Listing or product URLs")
    parser.add_argument("--output-dir", default="source_outputs")
    parser.add_argument("--limit", type=int, default=100)
    parser.add_argument("--delay", type=float, default=0.4)
    parser.add_argument("--timeout", type=int, default=10)
    args = parser.parse_args()
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update(HEADERS)
    for index, source_url in enumerate(args.urls, 1):
        records, errors = scrape_source(source_url, session, args.limit, args.delay, args.timeout)
        payload = {"source_url": source_url, "products": deduplicate(records), "errors": errors}
        output_path = output_dir / f"source_{index:02d}.json"
        output_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        print(f"{output_path}: {len(payload['products'])} products, {len(errors)} errors")


if __name__ == "__main__":
    main()