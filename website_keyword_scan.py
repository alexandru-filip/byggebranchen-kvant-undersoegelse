#!/usr/bin/env python3
"""Find company websites and scan them for selected keywords.

VS Code default:
    Put this script next to companies_with_public_financials.csv and press Run.

Default input:
    companies_with_public_financials.csv

Default output:
    website_keyword_hits.csv

Optional:
    Create keywords.txt next to this script, one keyword per line.
"""

from __future__ import annotations

import argparse
import csv
import gzip
import html
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import urllib.robotparser
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "companies_with_public_financials.csv"
DEFAULT_OUTPUT = SCRIPT_DIR / "website_keyword_hits.csv"
DEFAULT_KEYWORDS_FILE = SCRIPT_DIR / "keywords.txt"
APICVR_BY_CVR_URL = "https://apicvr.dk/api/v1/{cvr}"
APICVR_SEARCH_URL = "https://apicvr.dk/api/v1/search/company/{company_name}"
USER_AGENT = "company-website-keyword-scan/0.1"

DEFAULT_KEYWORDS = [
    "nybygget",
    "ny bygget",
    "nybyggeri",
    "nybyg",
    "familiehus",
    "familiehuse",
    "typehus",
    "parcelhus",
    "enfamiliehus",
    "husbyggeri",
    "boligbyggeri",
]

OUTPUT_COLUMNS = [
    "input_company_name",
    "cvr_number",
    "website",
    "matched_keywords",
    "keyword_count",
    "matched_pages",
    "pages_checked",
    "status",
    "notes",
]

NAME_COLUMNS = ["input_company_name", "resolved_company_name", "company_name", "name", "Navn"]
CVR_COLUMNS = ["resolved_cvr", "cvr_number", "cvr", "CVR", "vat"]
WEBSITE_COLUMNS = ["website", "Website", "url", "URL", "web"]

XLSX_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass
class Company:
    name: str
    cvr: str = ""
    website: str = ""


class HttpClient:
    def __init__(self, sleep_seconds: float) -> None:
        self.sleep_seconds = sleep_seconds
        self.last_request_at = 0.0

    def get_json(self, url: str) -> Any:
        data = self.get_bytes(url)
        return json.loads(data.decode("utf-8"))

    def get_text(self, url: str) -> str:
        data = self.get_bytes(url)
        for encoding in ("utf-8", "iso-8859-1"):
            try:
                return data.decode(encoding)
            except UnicodeDecodeError:
                continue
        return data.decode("utf-8", errors="replace")

    def get_bytes(self, url: str) -> bytes:
        self.wait()
        request = urllib.request.Request(
            url,
            headers={
                "Accept": "text/html,application/json,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Encoding": "gzip, identity",
                "User-Agent": USER_AGENT,
            },
        )
        with urllib.request.urlopen(request, timeout=20) as response:
            data = response.read()
            encoding = response.headers.get("Content-Encoding", "")
        self.last_request_at = time.monotonic()

        if encoding.lower() == "gzip" or data[:2] == b"\x1f\x8b":
            data = gzip.decompress(data)
        return data

    def wait(self) -> None:
        elapsed = time.monotonic() - self.last_request_at
        if elapsed < self.sleep_seconds:
            time.sleep(self.sleep_seconds - elapsed)


def main() -> int:
    args = parse_args()
    input_path = resolve_path(args.input, must_exist=True)
    output_path = resolve_path(args.output, must_exist=False)
    keywords = load_keywords(resolve_path(args.keywords_file, must_exist=False))
    companies = read_companies(input_path, args.name_column, args.cvr_column, args.website_column)
    if args.limit:
        companies = companies[: args.limit]

    print(f"Input:    {input_path}")
    print(f"Output:   {output_path}")
    print(f"Keywords: {', '.join(keywords)}")
    print(f"Loaded {len(companies)} companies")

    if args.dry_run:
        for company in companies[:10]:
            print(f"- {company.name} | CVR: {company.cvr or '-'} | website: {company.website or '-'}")
        return 0

    client = HttpClient(args.sleep)
    results = []
    for index, company in enumerate(companies, start=1):
        print(f"[{index}/{len(companies)}] {company.name}")
        domain_guess_limit = 0 if args.no_domain_guess else args.domain_guess_limit
        row = scan_company(company, keywords, client, args.max_pages, args.include_non_hits, domain_guess_limit)
        if row:
            results.append(row)
            write_rows(output_path, results)

    write_rows(output_path, results)
    print(f"Wrote {len(results)} rows to {output_path}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Find company websites and scan for keywords.")
    parser.add_argument(
        "input",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help="Input .csv or .xlsx. Defaults to companies_with_public_financials.csv next to this script.",
    )
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output CSV path.")
    parser.add_argument("--keywords-file", default=str(DEFAULT_KEYWORDS_FILE), help="Optional keywords.txt path.")
    parser.add_argument("--name-column", default="", help="Override company-name column.")
    parser.add_argument("--cvr-column", default="", help="Override CVR column.")
    parser.add_argument("--website-column", default="", help="Override website column if input already has one.")
    parser.add_argument("--max-pages", type=int, default=12, help="Maximum pages to scan per website.")
    parser.add_argument(
        "--domain-guess-limit",
        type=int,
        default=5,
        help="Try up to N likely .dk domains if CVR has no website.",
    )
    parser.add_argument("--no-domain-guess", action="store_true", help="Only use websites from input/CVR.")
    parser.add_argument("--sleep", type=float, default=0.5, help="Delay between HTTP requests.")
    parser.add_argument("--limit", type=int, default=0, help="Only process first N companies.")
    parser.add_argument("--include-non-hits", action="store_true", help="Also save companies with no keyword hits.")
    parser.add_argument("--dry-run", action="store_true", help="Preview input rows without using the network.")
    return parser.parse_args()


def scan_company(
    company: Company,
    keywords: list[str],
    client: HttpClient,
    max_pages: int,
    include_non_hits: bool,
    domain_guess_limit: int,
) -> dict[str, str] | None:
    website, website_note = find_website(company, client, domain_guess_limit)
    if not website:
        row = result_row(company, "", [], [], 0, "no_website_found", website_note)
        return row if include_non_hits else None

    website = normalize_website_url(website)
    try:
        if not robots_allows(website):
            row = result_row(company, website, [], [], 0, "blocked_by_robots", "")
            return row

        pages = collect_pages_to_scan(website, client, max_pages)
        matched_keywords: set[str] = set()
        matched_pages: list[str] = []

        for page_url, page_text in pages:
            found = find_keywords(page_text, keywords)
            if found:
                matched_keywords.update(found)
                matched_pages.append(page_url)

        if matched_keywords:
            return result_row(
                company,
                website,
                sorted(matched_keywords),
                list(dict.fromkeys(matched_pages)),
                len(pages),
                "keyword_hit",
                website_note,
            )

        row = result_row(company, website, [], [], len(pages), "no_keyword_hit", website_note)
        return row if include_non_hits else None
    except urllib.error.HTTPError as error:
        row = result_row(company, website, [], [], 0, "http_error", f"{error.code} {error.reason}")
        return row if include_non_hits else None
    except urllib.error.URLError as error:
        row = result_row(company, website, [], [], 0, "url_error", str(error.reason))
        return row if include_non_hits else None
    except Exception as error:
        row = result_row(company, website, [], [], 0, "error", f"{type(error).__name__}: {error}")
        return row if include_non_hits else None


def find_website(company: Company, client: HttpClient, domain_guess_limit: int) -> tuple[str, str]:
    if company.website:
        return company.website, "website came from input"

    if company.cvr:
        try:
            data = client.get_json(APICVR_BY_CVR_URL.format(cvr=urllib.parse.quote(company.cvr)))
            website = clean_website_value(data.get("website", ""))
            if website:
                return website, "website found by CVR through apicvr.dk"
        except urllib.error.HTTPError:
            pass

    if company.name:
        for query in company_search_queries(company.name):
            try:
                url = APICVR_SEARCH_URL.format(company_name=urllib.parse.quote(query, safe=""))
                results = client.get_json(url)
            except urllib.error.HTTPError:
                continue
            if not isinstance(results, list):
                continue
            for result in results:
                if company.cvr and str(result.get("vat", "")) != company.cvr:
                    continue
                website = clean_website_value(result.get("website", ""))
                if website:
                    return website, f"website found by name through apicvr.dk; query={query}"

    if domain_guess_limit:
        guessed = try_guessed_websites(company, client, domain_guess_limit)
        if guessed[0]:
            return guessed

    return "", "apicvr.dk did not return a website"


def try_guessed_websites(company: Company, client: HttpClient, limit: int) -> tuple[str, str]:
    for url in candidate_website_urls(company.name, limit):
        try:
            page_text = client.get_text(url)
        except Exception:
            continue
        if site_looks_like_company(company.name, url, page_text):
            return url, "website guessed from company name"
    return "", "no likely .dk domain responded"


def candidate_website_urls(company_name: str, limit: int) -> list[str]:
    words = company_words(company_name)
    if not words:
        return []

    stems = []
    joined = "".join(words)
    dashed = "-".join(words)
    stems.extend([joined, dashed])
    if len(words) > 1:
        stems.append(words[0] + words[-1])
    if len(words) == 1:
        stems.append(words[0])

    unique_stems = []
    seen = set()
    for stem in stems:
        if stem and stem not in seen:
            seen.add(stem)
            unique_stems.append(stem)

    urls = []
    for stem in unique_stems[:limit]:
        domain = stem + ".dk"
        urls.extend([
            f"https://www.{domain}",
            f"https://{domain}",
            f"http://www.{domain}",
            f"http://{domain}",
        ])
    return urls


def site_looks_like_company(company_name: str, url: str, page_text: str) -> bool:
    searchable = normalize_text(url + " " + page_text)
    if any(bad in searchable for bad in ["domain til salg", "domaene til salg", "domain is parked", "buy this domain"]):
        return False

    all_words = company_words(company_name)
    words = [word for word in all_words if len(word) > 2] or all_words
    if not words:
        return True
    if len(words) == 1:
        return words[0] in searchable
    matches = sum(1 for word in words if word in searchable)
    return matches >= min(2, len(words))


def company_words(company_name: str) -> list[str]:
    name = re.sub(
        r"\b(a/s|aps|ivs|i/s|amba|s\.m\.b\.a\.?|smb[a.]*)\b",
        " ",
        company_name,
        flags=re.IGNORECASE,
    )
    name = name.replace("&", " og ")
    name = transliterate_danish(name.casefold())
    words = re.findall(r"[a-z0-9]+", name)
    return [word for word in words if word not in {"og", "the", "as", "aps", "byg", "as"}]


def transliterate_danish(value: str) -> str:
    return (
        value.replace("æ", "ae")
        .replace("ø", "oe")
        .replace("å", "aa")
    )


def collect_pages_to_scan(
    website: str,
    client: HttpClient,
    max_pages: int,
) -> list[tuple[str, str]]:
    queue = [website]
    seen = set()
    pages: list[tuple[str, str]] = []

    while queue and len(pages) < max_pages:
        url = queue.pop(0)
        if url in seen:
            continue
        seen.add(url)

        html_text = client.get_text(url)
        text = html_to_text(html_text)
        pages.append((url, text))

        links = extract_same_site_links(url, html_text)
        for link in prioritize_links(links):
            if link not in seen and link not in queue:
                queue.append(link)
            if len(queue) + len(pages) >= max_pages * 2:
                break

    return pages


def find_keywords(text: str, keywords: list[str]) -> list[str]:
    normalized_text = normalize_text(text)
    found = []
    for keyword in keywords:
        normalized_keyword = normalize_text(keyword)
        if normalized_keyword and normalized_keyword in normalized_text:
            found.append(keyword)
    return found


def robots_allows(url: str) -> bool:
    parsed = urllib.parse.urlparse(url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    parser = urllib.robotparser.RobotFileParser()
    parser.set_url(robots_url)
    try:
        parser.read()
    except Exception:
        return True
    return parser.can_fetch(USER_AGENT, url)


def extract_same_site_links(base_url: str, html_text: str) -> list[str]:
    base = urllib.parse.urlparse(base_url)
    links = []
    for match in re.finditer(r"""href=["']([^"']+)["']""", html_text, flags=re.IGNORECASE):
        href = html.unescape(match.group(1)).strip()
        if not href or href.startswith(("mailto:", "tel:", "javascript:", "#")):
            continue
        url = urllib.parse.urljoin(base_url, href)
        parsed = urllib.parse.urlparse(url)
        if normalized_host(parsed.netloc) != normalized_host(base.netloc):
            continue
        if re.search(r"\.(pdf|jpg|jpeg|png|gif|svg|webp|zip|docx?|xlsx?)$", parsed.path, re.IGNORECASE):
            continue
        clean = urllib.parse.urlunparse((parsed.scheme, parsed.netloc, parsed.path.rstrip("/") or "/", "", "", ""))
        links.append(clean)
    return list(dict.fromkeys(links))


def prioritize_links(links: list[str]) -> list[str]:
    priority_words = [
        "nybyg",
        "bolig",
        "huse",
        "hus",
        "familie",
        "typehus",
        "projekter",
        "referencer",
        "cases",
        "ydelser",
    ]
    return sorted(
        links,
        key=lambda url: (
            0 if any(word in url.casefold() for word in priority_words) else 1,
            len(url),
            url,
        ),
    )


def result_row(
    company: Company,
    website: str,
    matched_keywords: list[str],
    matched_pages: list[str],
    pages_checked: int,
    status: str,
    notes: str,
) -> dict[str, str]:
    return {
        "input_company_name": company.name,
        "cvr_number": company.cvr,
        "website": website,
        "matched_keywords": "; ".join(matched_keywords),
        "keyword_count": str(len(matched_keywords)),
        "matched_pages": "; ".join(matched_pages),
        "pages_checked": str(pages_checked),
        "status": status,
        "notes": notes,
    }


def read_companies(
    path: Path,
    name_column: str = "",
    cvr_column: str = "",
    website_column: str = "",
) -> list[Company]:
    rows = read_csv_rows(path) if path.suffix.lower() == ".csv" else read_xlsx_rows(path)
    if not rows:
        return []

    headers = list(rows[0].keys())
    name_column = name_column or first_matching_column(headers, NAME_COLUMNS)
    cvr_column = cvr_column or first_matching_column(headers, CVR_COLUMNS)
    website_column = website_column or first_matching_column(headers, WEBSITE_COLUMNS)
    if not name_column:
        raise SystemExit(f"Could not find company-name column. Available columns: {', '.join(headers)}")

    companies = []
    seen = set()
    for row in rows:
        name = str(row.get(name_column, "")).strip()
        cvr = digits_only(str(row.get(cvr_column, "")).strip()) if cvr_column else ""
        website = clean_website_value(str(row.get(website_column, "")).strip()) if website_column else ""
        if not name:
            continue
        key = (name.casefold(), cvr)
        if key in seen:
            continue
        seen.add(key)
        companies.append(Company(name=name, cvr=cvr, website=website))
    return companies


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as csvfile:
        return list(csv.DictReader(csvfile))


def read_xlsx_rows(path: Path) -> list[dict[str, str]]:
    rows = read_first_sheet_rows(path)
    if not rows:
        return []
    headers = [str(value or "").strip() for value in rows[0]]
    records = []
    for row in rows[1:]:
        records.append({header: cell(row, index) for index, header in enumerate(headers) if header})
    return records


def read_first_sheet_rows(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as xlsx:
        shared_strings = read_shared_strings(xlsx)
        sheet_path = first_sheet_path(xlsx)
        sheet_xml = ET.fromstring(xlsx.read(sheet_path))

    rows = []
    for row in sheet_xml.findall(".//main:sheetData/main:row", XLSX_NS):
        values = []
        for cell_element in row.findall("main:c", XLSX_NS):
            column = column_number(cell_element.attrib["r"]) - 1
            while len(values) < column:
                values.append("")
            values.append(cell_value(cell_element, shared_strings))
        rows.append(values)
    return rows


def read_shared_strings(xlsx: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in xlsx.namelist():
        return []
    root = ET.fromstring(xlsx.read("xl/sharedStrings.xml"))
    return ["".join(text.text or "" for text in item.findall(".//main:t", XLSX_NS)) for item in root.findall("main:si", XLSX_NS)]


def first_sheet_path(xlsx: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(xlsx.read("xl/workbook.xml"))
    first_sheet = workbook.find(".//main:sheets/main:sheet", XLSX_NS)
    if first_sheet is None:
        raise SystemExit("Workbook does not contain any sheets.")

    rel_id = first_sheet.attrib[f"{{{XLSX_NS['rel']}}}id"]
    rels = ET.fromstring(xlsx.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall("pkgrel:Relationship", XLSX_NS):
        if rel.attrib["Id"] == rel_id:
            return "xl/" + rel.attrib["Target"].lstrip("/")
    raise SystemExit("Could not locate first worksheet XML.")


def cell_value(cell_element: ET.Element, shared_strings: list[str]) -> str:
    value_element = cell_element.find("main:v", XLSX_NS)
    if value_element is None:
        inline_text = cell_element.find(".//main:t", XLSX_NS)
        return inline_text.text if inline_text is not None and inline_text.text else ""
    raw_value = value_element.text or ""
    if cell_element.attrib.get("t") == "s":
        return shared_strings[int(raw_value)]
    return raw_value


def column_number(cell_reference: str) -> int:
    letters = "".join(char for char in cell_reference if char.isalpha())
    number = 0
    for char in letters:
        number = number * 26 + ord(char.upper()) - ord("A") + 1
    return number


def cell(row: list[str], index: int) -> str:
    return str(row[index] or "") if index < len(row) else ""


def write_rows(path: Path, rows: list[dict[str, str]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def load_keywords(path: Path) -> list[str]:
    if path.exists():
        keywords = [line.strip() for line in path.read_text(encoding="utf-8").splitlines()]
        keywords = [line for line in keywords if line and not line.startswith("#")]
        if keywords:
            return unique_strings(keywords)
    return unique_strings(DEFAULT_KEYWORDS)


def unique_strings(values: list[str]) -> list[str]:
    unique = []
    seen = set()
    for value in values:
        key = value.casefold()
        if key in seen:
            continue
        seen.add(key)
        unique.append(value)
    return unique


def resolve_path(path_value: str | Path, must_exist: bool) -> Path:
    path = Path(path_value).expanduser()
    if not path.is_absolute():
        cwd_candidate = (Path.cwd() / path).resolve()
        script_candidate = (SCRIPT_DIR / path).resolve()
        path = cwd_candidate if cwd_candidate.exists() else script_candidate
    else:
        path = path.resolve()
    if must_exist and not path.exists():
        raise SystemExit(f"Could not find file: {path}")
    return path


def clean_website_value(value: Any) -> str:
    if value is None:
        return ""
    value = str(value).strip()
    if not value or value.lower() in {"none", "nan", "null"}:
        return ""
    return value


def normalize_website_url(value: str) -> str:
    value = value.strip()
    if not value.startswith(("http://", "https://")):
        value = "https://" + value
    return value.rstrip("/")


def html_to_text(html_text: str) -> str:
    html_text = re.sub(r"(?is)<script.*?</script>", " ", html_text)
    html_text = re.sub(r"(?is)<style.*?</style>", " ", html_text)
    html_text = re.sub(r"(?s)<[^>]+>", " ", html_text)
    return html.unescape(re.sub(r"\s+", " ", html_text))


def normalize_text(value: str) -> str:
    value = value.casefold()
    value = re.sub(r"[^0-9a-zæøå]+", " ", value)
    return " ".join(value.split())


def normalized_host(host: str) -> str:
    return host.casefold().removeprefix("www.")


def digits_only(value: str) -> str:
    return re.sub(r"\D+", "", value)


def first_matching_column(headers: list[str], candidates: list[str]) -> str:
    lower_to_actual = {header.casefold(): header for header in headers}
    for candidate in candidates:
        if candidate.casefold() in lower_to_actual:
            return lower_to_actual[candidate.casefold()]
    return ""


def company_search_queries(name: str) -> list[str]:
    without_suffix = re.sub(
        r"\s+(a/s|aps|ivs|i/s|amba|s\.m\.b\.a\.?|smb[a.]*)\.?$",
        "",
        name,
        flags=re.IGNORECASE,
    ).strip()
    variants = [name, without_suffix, without_suffix.replace("/", " ")]
    words = without_suffix.split()
    if len(words) > 2:
        variants.append(" ".join(words[:3]))
        variants.append(" ".join(words[:2]))
    elif len(words) == 2:
        variants.append(words[0])

    unique = []
    seen = set()
    for variant in variants:
        variant = " ".join(variant.split()).strip()
        key = variant.casefold()
        if variant and key not in seen:
            seen.add(key)
            unique.append(variant)
    return unique


if __name__ == "__main__":
    raise SystemExit(main())
