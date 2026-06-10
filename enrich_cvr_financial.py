#!/usr/bin/env python3
"""Enrich a company list with public Danish annual-report XBRL financials.

The script reads company names, optionally resolves CVR numbers through
apicvr.dk, finds the latest public annual-report XML in Virk's publication
index, and extracts both:

    - nettoomsætning: usually XBRL tag Revenue
    - bruttofortjeneste: usually XBRL tag GrossProfitLoss

It writes a CSV and saves progress after every company.

Usage:
    Open this folder in VS Code and press Run on this file.

Defaults when run without arguments:
    input:  names.xlsx in the same folder as this script
    output: companies_with_public_financials.csv in the same folder
"""

from __future__ import annotations

import argparse
import csv
import gzip
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "names.xlsx"
DEFAULT_OUTPUT = SCRIPT_DIR / "companies_with_public_financials.csv"
DEFAULT_CACHE_DIR = SCRIPT_DIR / ".annual_report_cache"
APICVR_SEARCH_URL = "https://apicvr.dk/api/v1/search/company/{company_name}"
VIRK_PUBLICATION_SEARCH_URL = "http://distribution.virk.dk/offentliggoerelser/_search"
USER_AGENT = "cvr-public-financial-enrichment/0.1"

BRANCH_CODE = "410000"
NET_REVENUE_TAGS = {
    "Revenue",
    "NetRevenue",
    "RevenueFromSaleOfGoodsAndServices",
    "RevenueFromContractsWithCustomers",
}
GROSS_PROFIT_TAGS = {
    "GrossProfitLoss",
    "GrossProfit",
    "GrossResult",
    "GrossProfitOrLoss",
}
EMPLOYEE_TAGS = {
    "AverageNumberOfEmployees",
    "AverageNumberOfFullTimeEmployees",
    "AverageNumberOfEmployeesDuringTheFinancialYear",
}

OUTPUT_COLUMNS = [
    "input_company_name",
    "resolved_cvr",
    "resolved_company_name",
    "match_status",
    "match_reason",
    "industry_code",
    "industry_description",
    "company_status",
    "employees_from_cvr",
    "annual_report_start_date",
    "annual_report_end_date",
    "annual_report_publication_date",
    "net_revenue_dkk",
    "net_revenue_tag",
    "gross_profit_dkk",
    "gross_profit_tag",
    "average_employees_xbrl",
    "financial_status",
    "annual_report_xml_url",
    "notes",
]

XLSX_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass
class Company:
    input_company_name: str
    cvr_number: str = ""


@dataclass
class Candidate:
    value: str
    tag: str
    context_ref: str
    period_end: date | None
    has_dimensions: bool


class HttpClient:
    def __init__(self, cache_dir: Path, sleep_seconds: float, refresh_cache: bool = False) -> None:
        self.cache_dir = cache_dir
        self.sleep_seconds = sleep_seconds
        self.refresh_cache = refresh_cache
        self.last_request_at = 0.0
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def get_json(self, url: str, cache_key: str) -> Any:
        data = self.get_bytes(url, cache_key)
        return json.loads(data.decode("utf-8"))

    def get_xml_root(self, url: str, cache_key: str) -> ET.Element:
        data = self.get_bytes(url, cache_key)
        return ET.fromstring(data)

    def get_bytes(self, url: str, cache_key: str) -> bytes:
        cache_path = self.cache_dir / safe_filename(cache_key)
        if cache_path.exists() and not self.refresh_cache:
            return cache_path.read_bytes()

        self.wait()
        request = urllib.request.Request(
            url,
            headers={
                "Accept": "application/json, application/xml, text/xml, */*",
                "Accept-Encoding": "gzip, identity",
                "User-Agent": USER_AGENT,
            },
        )
        with urllib.request.urlopen(request, timeout=30) as response:
            data = response.read()
            encoding = response.headers.get("Content-Encoding", "")
        self.last_request_at = time.monotonic()

        if encoding.lower() == "gzip" or data[:2] == b"\x1f\x8b":
            data = gzip.decompress(data)

        cache_path.write_bytes(data)
        return data

    def wait(self) -> None:
        elapsed = time.monotonic() - self.last_request_at
        if elapsed < self.sleep_seconds:
            time.sleep(self.sleep_seconds - elapsed)


def main() -> int:
    args = parse_args()
    input_path = resolve_path(args.input, must_exist=True)
    output_path = resolve_path(args.output, must_exist=False)
    cache_dir = resolve_path(args.cache_dir, must_exist=False)

    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")

    companies = read_companies(input_path, args.name_column, args.cvr_column)
    if args.limit:
        companies = companies[: args.limit]
    if not companies:
        print("No companies found.", file=sys.stderr)
        return 1

    completed = read_completed_rows(output_path) if args.resume else {}
    rows = unique_completed_rows(completed)
    client = HttpClient(cache_dir, args.sleep, args.refresh_cache)

    print(f"Loaded {len(companies)} companies")
    print(f"Already completed: {len(completed)}")

    for index, company in enumerate(companies, start=1):
        key = row_key(company.input_company_name, company.cvr_number)
        name_key = row_key(company.input_company_name)
        if key in completed or name_key in completed:
            continue

        print(f"[{index}/{len(companies)}] {company.input_company_name}")
        row = enrich_company(company, client, args.max_reports, args.branch_code)
        rows.append(row)
        write_rows(output_path, rows)

    write_rows(output_path, rows)
    print(f"Wrote {len(rows)} rows to {output_path}")
    return 0


def resolve_path(path_value: str | Path, must_exist: bool) -> Path:
    path = Path(path_value).expanduser()
    if not path.is_absolute():
        cwd_candidate = (Path.cwd() / path).resolve()
        script_candidate = (SCRIPT_DIR / path).resolve()
        path = cwd_candidate if cwd_candidate.exists() else script_candidate
    else:
        path = path.resolve()

    if must_exist and not path.exists():
        raise SystemExit(
            f"Could not find input file: {path}\n"
            f"Put names.xlsx in the same folder as enrich_cvr_financials.py, "
            f"or pass the full path to your file."
        )
    return path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Enrich companies with public annual-report net revenue and gross profit."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help="Input .xlsx or .csv file. Defaults to names.xlsx next to this script.",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help="Output CSV path. Defaults to this script folder.",
    )
    parser.add_argument("--name-column", default="Navn", help="Column containing company names.")
    parser.add_argument("--cvr-column", default="", help="Optional column containing CVR numbers.")
    parser.add_argument(
        "--branch-code",
        default=BRANCH_CODE,
        help="Preferred industry code when resolving names through apicvr.dk.",
    )
    parser.add_argument(
        "--cache-dir",
        default=str(DEFAULT_CACHE_DIR),
        help="Cache directory. Defaults to this script folder.",
    )
    parser.add_argument("--refresh-cache", action="store_true", help="Ignore cached HTTP responses.")
    parser.add_argument("--sleep", type=float, default=0.5, help="Delay between HTTP requests.")
    parser.add_argument("--max-reports", type=int, default=8, help="Annual reports to inspect per company.")
    parser.add_argument("--limit", type=int, default=0, help="Only process the first N companies.")
    parser.add_argument("--no-resume", dest="resume", action="store_false", help="Do not skip existing rows.")
    parser.set_defaults(resume=True)
    return parser.parse_args()


def enrich_company(
    company: Company,
    client: HttpClient,
    max_reports: int,
    branch_code: str,
) -> dict[str, Any]:
    base_row = blank_row(company.input_company_name)
    try:
        resolved = resolve_company(company, client, branch_code)
        base_row.update(resolved)
        cvr_number = str(base_row.get("resolved_cvr") or "")
        if not cvr_number:
            base_row["financial_status"] = "no_cvr"
            return base_row

        report = latest_report_with_xml(client, cvr_number, max_reports)
        if not report:
            base_row["financial_status"] = "no_xml_annual_report"
            return base_row

        base_row.update(report["metadata"])
        root = client.get_xml_root(report["xml_url"], f"xbrl_{cvr_number}_{report['report_id']}.xml")
        facts = extract_financial_facts(root)
        base_row.update(facts)

        has_net_revenue = bool(base_row["net_revenue_dkk"])
        has_gross_profit = bool(base_row["gross_profit_dkk"])
        if has_net_revenue and has_gross_profit:
            base_row["financial_status"] = "net_revenue_and_gross_profit_found"
        elif has_net_revenue:
            base_row["financial_status"] = "net_revenue_found"
        elif has_gross_profit:
            base_row["financial_status"] = "gross_profit_found"
        else:
            base_row["financial_status"] = "no_revenue_or_gross_profit_tags"
        return base_row
    except urllib.error.HTTPError as error:
        base_row["financial_status"] = "http_error"
        base_row["notes"] = f"{error.code} {error.reason}"
        return base_row
    except urllib.error.URLError as error:
        base_row["financial_status"] = "url_error"
        base_row["notes"] = str(error.reason)
        return base_row
    except ET.ParseError as error:
        base_row["financial_status"] = "xml_parse_error"
        base_row["notes"] = str(error)
        return base_row
    except Exception as error:
        base_row["financial_status"] = "error"
        base_row["notes"] = f"{type(error).__name__}: {error}"
        return base_row


def blank_row(input_company_name: str) -> dict[str, Any]:
    return {column: "" for column in OUTPUT_COLUMNS} | {"input_company_name": input_company_name}


def resolve_company(company: Company, client: HttpClient, branch_code: str) -> dict[str, Any]:
    if company.cvr_number:
        return {
            "resolved_cvr": company.cvr_number,
            "match_status": "provided_cvr",
            "match_reason": "CVR came from input file",
        }

    results = []
    successful_query = ""
    for query in company_search_queries(company.input_company_name):
        encoded_name = urllib.parse.quote(query, safe="")
        url = APICVR_SEARCH_URL.format(company_name=encoded_name)
        try:
            results = client.get_json(url, f"apicvr_search_{query}.json")
        except urllib.error.HTTPError as error:
            if error.code == 404:
                continue
            raise
        if isinstance(results, list) and results:
            successful_query = query
            break

    if not isinstance(results, list) or not results:
        return {"match_status": "not_found", "match_reason": "No apicvr.dk results"}

    best, reason = choose_best_company_match(company.input_company_name, results, branch_code)
    return {
        "resolved_cvr": best.get("vat", ""),
        "resolved_company_name": best.get("name", ""),
        "match_status": "matched",
        "match_reason": f"{reason}; query={successful_query}",
        "industry_code": best.get("industrycode", ""),
        "industry_description": best.get("industrydesc", ""),
        "company_status": best.get("status", ""),
        "employees_from_cvr": best.get("employees", ""),
    }


def company_search_queries(name: str) -> list[str]:
    without_suffix = strip_company_suffix(name)
    slash_as_space = name.replace("/", " ")
    variants = [
        name,
        without_suffix,
        strip_company_suffix(slash_as_space),
    ]

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


def strip_company_suffix(name: str) -> str:
    return re.sub(
        r"\s+(a/s|aps|ivs|i/s|amba|s\.m\.b\.a\.?|smb[a.]*)\.?$",
        "",
        name,
        flags=re.IGNORECASE,
    ).strip()


def choose_best_company_match(
    input_name: str,
    results: list[dict[str, Any]],
    branch_code: str,
) -> tuple[dict[str, Any], str]:
    normalized_input = normalize_name(input_name)

    def is_active(result: dict[str, Any]) -> bool:
        return str(result.get("status", "")).upper() in {"NORMAL", "AKTIV"}

    def industry_matches(result: dict[str, Any]) -> bool:
        return str(result.get("industrycode", "")) == str(branch_code)

    def exact_name(result: dict[str, Any]) -> bool:
        return normalize_name(str(result.get("name", ""))) == normalized_input

    rules = [
        (lambda item: exact_name(item) and industry_matches(item) and is_active(item), "exact_name_branch_active"),
        (lambda item: exact_name(item) and industry_matches(item), "exact_name_branch"),
        (lambda item: industry_matches(item) and is_active(item), "branch_active"),
        (lambda item: exact_name(item) and is_active(item), "exact_name_active"),
        (lambda item: exact_name(item), "exact_name"),
        (lambda item: is_active(item), "first_active"),
    ]
    for predicate, reason in rules:
        for result in results:
            if predicate(result):
                return result, reason
    return results[0], "first_result"


def latest_report_with_xml(client: HttpClient, cvr_number: str, max_reports: int) -> dict[str, Any] | None:
    params = {
        "q": f"cvrNummer:{cvr_number}",
        "size": str(max_reports),
        "sort": "regnskab.regnskabsperiode.slutDato:desc",
    }
    url = VIRK_PUBLICATION_SEARCH_URL + "?" + urllib.parse.urlencode(params)
    data = client.get_json(url, f"virk_reports_{cvr_number}.json")

    for hit in data.get("hits", {}).get("hits", []):
        source = hit.get("_source", {})
        if source.get("offentliggoerelsestype") != "regnskab":
            continue
        if source.get("omgoerelse") is True:
            continue

        xml_url = annual_report_xml_url(source.get("dokumenter", []))
        if not xml_url:
            continue

        period = (source.get("regnskab") or {}).get("regnskabsperiode") or {}
        return {
            "report_id": hit.get("_id", "latest").replace(":", "_").replace("/", "_"),
            "xml_url": xml_url,
            "metadata": {
                "annual_report_start_date": period.get("startDato", ""),
                "annual_report_end_date": period.get("slutDato", ""),
                "annual_report_publication_date": source.get("offentliggoerelsesTidspunkt", ""),
                "annual_report_xml_url": xml_url,
            },
        }
    return None


def annual_report_xml_url(documents: list[dict[str, Any]]) -> str:
    for document in documents:
        if document.get("dokumentType") != "AARSRAPPORT":
            continue
        mime_type = str(document.get("dokumentMimeType", "")).lower()
        url = str(document.get("dokumentUrl", ""))
        if mime_type == "application/xml" or url.lower().endswith(".xml"):
            return url
    return ""


def extract_financial_facts(root: ET.Element) -> dict[str, Any]:
    contexts = parse_contexts(root)
    facts_by_tag = collect_numeric_candidates(root, contexts)

    net_revenue = choose_fact(facts_by_tag, NET_REVENUE_TAGS)
    gross_profit = choose_fact(facts_by_tag, GROSS_PROFIT_TAGS)
    average_employees = choose_fact(facts_by_tag, EMPLOYEE_TAGS)

    return {
        "net_revenue_dkk": net_revenue.value if net_revenue else "",
        "net_revenue_tag": net_revenue.tag if net_revenue else "",
        "gross_profit_dkk": gross_profit.value if gross_profit else "",
        "gross_profit_tag": gross_profit.tag if gross_profit else "",
        "average_employees_xbrl": average_employees.value if average_employees else "",
    }


def parse_contexts(root: ET.Element) -> dict[str, tuple[date | None, bool]]:
    contexts = {}
    for context in root.iter():
        if local_name(context.tag) != "context":
            continue

        context_id = context.attrib.get("id")
        if not context_id:
            continue

        period_end = None
        for child in context.iter():
            name = local_name(child.tag)
            if name in {"endDate", "instant"} and child.text:
                period_end = parse_date(child.text.strip())
                if period_end:
                    break

        has_dimensions = any(
            local_name(child.tag) in {"explicitMember", "typedMember", "segment", "scenario"}
            for child in context.iter()
            if child is not context
        )
        contexts[context_id] = (period_end, has_dimensions)
    return contexts


def collect_numeric_candidates(
    root: ET.Element,
    contexts: dict[str, tuple[date | None, bool]],
) -> dict[str, list[Candidate]]:
    facts: dict[str, list[Candidate]] = {}
    interesting_tags = NET_REVENUE_TAGS | GROSS_PROFIT_TAGS | EMPLOYEE_TAGS
    for element in root.iter():
        tag = local_name(element.tag)
        if tag not in interesting_tags:
            continue

        value = clean_number(element.text)
        if value == "":
            continue

        context_ref = element.attrib.get("contextRef", "")
        period_end, has_dimensions = contexts.get(context_ref, (None, False))
        facts.setdefault(tag, []).append(
            Candidate(
                value=value,
                tag=tag,
                context_ref=context_ref,
                period_end=period_end,
                has_dimensions=has_dimensions,
            )
        )
    return facts


def choose_fact(
    facts_by_tag: dict[str, list[Candidate]],
    allowed_tags: set[str],
) -> Candidate | None:
    candidates = [
        candidate
        for tag in allowed_tags
        for candidate in facts_by_tag.get(tag, [])
    ]
    if not candidates:
        return None

    return sorted(
        candidates,
        key=lambda candidate: (
            not candidate.has_dimensions,
            candidate.period_end or date.min,
            tag_priority(candidate.tag, allowed_tags),
        ),
        reverse=True,
    )[0]


def tag_priority(tag: str, allowed_tags: set[str]) -> int:
    preferred = list(allowed_tags)
    return len(preferred) - preferred.index(tag) if tag in preferred else 0


def clean_number(value: str | None) -> str:
    if value is None:
        return ""
    value = value.strip().replace("\u00a0", "")
    if not value:
        return ""
    if re.fullmatch(r"-?\d+(\.\d+)?", value):
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number)
    return ""


def read_companies(path: Path, name_column: str, cvr_column: str = "") -> list[Company]:
    if path.suffix.lower() == ".csv":
        return read_companies_from_csv(path, name_column, cvr_column)
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        return read_companies_from_xlsx(path, name_column, cvr_column)
    raise SystemExit(f"Unsupported input type: {path.suffix}")


def read_companies_from_csv(path: Path, name_column: str, cvr_column: str = "") -> list[Company]:
    with path.open(newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        if not reader.fieldnames or name_column not in reader.fieldnames:
            available = ", ".join(reader.fieldnames or [])
            raise SystemExit(f"Could not find name column '{name_column}'. Available columns: {available}")
        return unique_companies(
            Company(
                input_company_name=str(row.get(name_column, "")).strip(),
                cvr_number=str(row.get(cvr_column, "")).strip() if cvr_column else "",
            )
            for row in reader
        )


def read_companies_from_xlsx(path: Path, name_column: str, cvr_column: str = "") -> list[Company]:
    rows = read_first_sheet_rows(path)
    if not rows:
        return []

    headers = [str(value or "").strip() for value in rows[0]]
    name_index = column_index(headers, name_column)
    cvr_index = column_index(headers, cvr_column) if cvr_column else None
    return unique_companies(
        Company(
            input_company_name=cell(row, name_index).strip(),
            cvr_number=cell(row, cvr_index).strip() if cvr_index is not None else "",
        )
        for row in rows[1:]
    )


def column_index(headers: list[str], column: str) -> int:
    try:
        return headers.index(column)
    except ValueError:
        available = ", ".join(header for header in headers if header)
        raise SystemExit(f"Could not find column '{column}'. Available columns: {available}")


def unique_companies(companies: Any) -> list[Company]:
    unique = []
    seen = set()
    for company in companies:
        if not company.input_company_name:
            continue
        key = row_key(company.input_company_name, company.cvr_number)
        if key in seen:
            continue
        seen.add(key)
        unique.append(company)
    return unique


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
    strings = []
    for item in root.findall("main:si", XLSX_NS):
        strings.append("".join(text.text or "" for text in item.findall(".//main:t", XLSX_NS)))
    return strings


def first_sheet_path(xlsx: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(xlsx.read("xl/workbook.xml"))
    first_sheet = workbook.find(".//main:sheets/main:sheet", XLSX_NS)
    if first_sheet is None:
        raise SystemExit("Workbook does not contain any sheets.")

    rel_id = first_sheet.attrib[f"{{{XLSX_NS['rel']}}}id"]
    rels = ET.fromstring(xlsx.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall("pkgrel:Relationship", XLSX_NS):
        if rel.attrib["Id"] == rel_id:
            target = rel.attrib["Target"]
            return "xl/" + target.lstrip("/")
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


def cell(row: list[str], index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    return str(row[index] or "")


def unique_completed_rows(completed: dict[tuple[str, str], dict[str, str]]) -> list[dict[str, str]]:
    rows = []
    seen = set()
    for row in completed.values():
        key = (
            row.get("input_company_name", ""),
            row.get("resolved_cvr", ""),
            row.get("annual_report_xml_url", ""),
        )
        if key in seen:
            continue
        seen.add(key)
        rows.append(row)
    return rows


def read_completed_rows(path: Path) -> dict[tuple[str, str], dict[str, str]]:
    if not path.exists():
        return {}
    with path.open(newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        completed = {}
        for row in reader:
            name = row.get("input_company_name", "")
            completed[row_key(name)] = row
            completed[row_key(name, row.get("resolved_cvr", ""))] = row
        return completed


def write_rows(path: Path, rows: list[dict[str, Any]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def safe_filename(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value)[:180]


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def parse_date(value: str) -> date | None:
    try:
        return date.fromisoformat(value[:10])
    except ValueError:
        return None


def normalize_name(value: str) -> str:
    value = value.casefold()
    value = re.sub(r"\b(a/s|aps|ivs|i/s|amba|s\.m\.b\.a\.?)\b", "", value)
    value = re.sub(r"[^0-9a-zæøå]+", " ", value)
    return " ".join(value.split())


def row_key(name: str, cvr_number: str = "") -> tuple[str, str]:
    return (name.casefold().strip(), str(cvr_number or "").strip())


if __name__ == "__main__":
    raise SystemExit(main())
