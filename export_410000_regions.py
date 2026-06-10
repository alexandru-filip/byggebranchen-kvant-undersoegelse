
#!/usr/bin/env python3
"""Export CVR companies in selected Danish regions for industry code 410000.

Requires System-til-System access to CVR:

    export API_TOKEN="base64-encoded-basic-auth-token"
    python export_cvr_410000_regions.py

The API_TOKEN value is the same token format used by apicvr.dk.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
from collections.abc import Iterable
from typing import Any


CVR_SEARCH_URL = "http://distribution.virk.dk/cvr-permanent/virksomhed/_search"
CVR_SCROLL_URL = "http://distribution.virk.dk/_search/scroll"

INDUSTRY_CODE = 410000
DEFAULT_OUTPUT = "companies_410000_nord_mid_syd.csv"

REGION_MUNICIPALITIES = {
    "Nordjylland": {
        773: "Morsø",
        787: "Thisted",
        810: "Brønderslev",
        813: "Frederikshavn",
        820: "Vesthimmerlands",
        825: "Læsø",
        840: "Rebild",
        846: "Mariagerfjord",
        849: "Jammerbugt",
        851: "Aalborg",
        860: "Hjørring",
    },
    "Midtjylland": {
        615: "Horsens",
        657: "Herning",
        661: "Holstebro",
        665: "Lemvig",
        671: "Struer",
        706: "Syddjurs",
        707: "Norddjurs",
        710: "Favrskov",
        727: "Odder",
        730: "Randers",
        740: "Silkeborg",
        741: "Samsø",
        746: "Skanderborg",
        751: "Aarhus",
        756: "Ikast-Brande",
        760: "Ringkøbing-Skjern",
        766: "Hedensted",
        779: "Skive",
        791: "Viborg",
    },
    "Syddanmark": {
        410: "Middelfart",
        420: "Assens",
        430: "Faaborg-Midtfyn",
        440: "Kerteminde",
        450: "Nyborg",
        461: "Odense",
        479: "Svendborg",
        480: "Nordfyns",
        482: "Langeland",
        492: "Ærø",
        510: "Haderslev",
        530: "Billund",
        540: "Sønderborg",
        550: "Tønder",
        561: "Esbjerg",
        563: "Fanø",
        573: "Varde",
        575: "Vejen",
        580: "Aabenraa",
        607: "Fredericia",
        621: "Kolding",
        630: "Vejle",
    },
}

CSV_COLUMNS = [
    "cvr_number",
    "name",
    "industry_code",
    "industry_text",
    "region",
    "municipality_code",
    "municipality_name",
    "address",
    "zipcode",
    "city",
    "company_form_code",
    "company_form",
    "status",
    "start_date",
]


def main() -> int:
    args = parse_args()
    token = os.getenv("API_TOKEN")
    if not token:
        print("Missing API_TOKEN environment variable.", file=sys.stderr)
        return 2

    import requests

    municipality_to_region = build_municipality_region_lookup(args.regions)
    municipality_codes = sorted(municipality_to_region)

    session = requests.Session()
    session.headers.update(
        {
            "Authorization": f"Basic {token}",
            "Content-Type": "application/json",
        }
    )

    rows = []
    for company in iter_companies(
        session=session,
        industry_code=args.industry_code,
        municipality_codes=municipality_codes,
        batch_size=args.batch_size,
    ):
        row = company_to_row(company, municipality_to_region)
        if not row:
            continue
        if args.active_only and row["status"] != "NORMAL":
            continue
        rows.append(row)

    rows.sort(key=lambda row: (row["region"], row["municipality_name"], row["name"] or ""))
    write_csv(args.output, rows)
    print(f"Wrote {len(rows)} companies to {args.output}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export CVR companies for industry code 410000 in selected Danish regions."
    )
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="CSV output path.")
    parser.add_argument(
        "--industry-code",
        type=int,
        default=INDUSTRY_CODE,
        help="CVR main industry/branchekode to export.",
    )
    parser.add_argument(
        "--regions",
        nargs="+",
        default=["Nordjylland", "Midtjylland", "Syddanmark"],
        choices=sorted(REGION_MUNICIPALITIES),
        help="Region names to include.",
    )
    parser.add_argument(
        "--active-only",
        action="store_true",
        help='Only keep companies whose sammensatStatus is "NORMAL".',
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=1000,
        help="Elasticsearch scroll batch size.",
    )
    return parser.parse_args()


def build_municipality_region_lookup(regions: Iterable[str]) -> dict[int, tuple[str, str]]:
    lookup = {}
    for region in regions:
        for code, name in REGION_MUNICIPALITIES[region].items():
            lookup[code] = (region, name)
    return lookup


def iter_companies(
    session: requests.Session,
    industry_code: int,
    municipality_codes: list[int],
    batch_size: int,
) -> Iterable[dict[str, Any]]:
    payload = {
        "_source": ["Vrvirksomhed"],
        "query": {
            "bool": {
                "filter": [
                    {
                        "term": {
                            "Vrvirksomhed.virksomhedMetadata.nyesteHovedbranche.branchekode": industry_code
                        }
                    },
                    {
                        "terms": {
                            "Vrvirksomhed.virksomhedMetadata.nyesteBeliggenhedsadresse.kommune.kommuneKode": municipality_codes
                        }
                    },
                ]
            }
        },
        "size": batch_size,
    }

    response = session.post(CVR_SEARCH_URL, params={"scroll": "2m"}, json=payload, timeout=30)
    response.raise_for_status()
    data = response.json()
    scroll_id = data.get("_scroll_id")

    try:
        while True:
            hits = data.get("hits", {}).get("hits", [])
            if not hits:
                break
            for hit in hits:
                company = hit.get("_source", {}).get("Vrvirksomhed")
                if company:
                    yield company

            if not scroll_id:
                break

            data = next_scroll_page(session, scroll_id)
            scroll_id = data.get("_scroll_id", scroll_id)
    finally:
        if scroll_id:
            clear_scroll(session, scroll_id)


def next_scroll_page(session: requests.Session, scroll_id: str) -> dict[str, Any]:
    response = session.post(
        CVR_SCROLL_URL,
        json={"scroll": "2m", "scroll_id": scroll_id},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def clear_scroll(session: requests.Session, scroll_id: str) -> None:
    try:
        session.delete(CVR_SCROLL_URL, json={"scroll_id": [scroll_id]}, timeout=10)
    except Exception:
        pass


def company_to_row(
    company: dict[str, Any],
    municipality_to_region: dict[int, tuple[str, str]],
) -> dict[str, Any] | None:
    metadata = company.get("virksomhedMetadata") or {}
    branch = metadata.get("nyesteHovedbranche") or {}
    address = metadata.get("nyesteBeliggenhedsadresse") or {}
    municipality = address.get("kommune") or {}
    company_form = metadata.get("nyesteVirksomhedsform") or {}

    municipality_code = municipality.get("kommuneKode")
    if municipality_code not in municipality_to_region:
        return None

    region, fallback_municipality_name = municipality_to_region[municipality_code]
    return {
        "cvr_number": company.get("cvrNummer"),
        "name": nested_get(metadata, "nyesteNavn", "navn"),
        "industry_code": branch.get("branchekode"),
        "industry_text": branch.get("branchetekst"),
        "region": region,
        "municipality_code": municipality_code,
        "municipality_name": municipality.get("kommuneNavn") or fallback_municipality_name,
        "address": format_address(address),
        "zipcode": address.get("postnummer"),
        "city": address.get("postdistrikt"),
        "company_form_code": company_form.get("virksomhedsformkode"),
        "company_form": company_form.get("langBeskrivelse"),
        "status": metadata.get("sammensatStatus"),
        "start_date": metadata.get("stiftelsesDato"),
    }


def nested_get(data: dict[str, Any], *keys: str) -> Any:
    value: Any = data
    for key in keys:
        if not isinstance(value, dict):
            return None
        value = value.get(key)
    return value


def format_address(address: dict[str, Any]) -> str | None:
    if address.get("adressebetegnelse"):
        return address["adressebetegnelse"]

    street = address.get("vejnavn")
    house_number = address.get("husnummerFra")
    if not street or house_number is None:
        return None

    house = str(house_number)
    if address.get("husnummerTil"):
        house += f"-{address['husnummerTil']}"
    if address.get("bogstavFra"):
        house += str(address["bogstavFra"])
    if address.get("bogstavTil"):
        house += f"-{address['bogstavTil']}"

    parts = [f"{street} {house}"]
    if address.get("etage"):
        parts.append(str(address["etage"]))
    if address.get("sidedoer"):
        parts.append(str(address["sidedoer"]))
    return ", ".join(parts)


def write_csv(path: str, rows: list[dict[str, Any]]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    raise SystemExit(main())
