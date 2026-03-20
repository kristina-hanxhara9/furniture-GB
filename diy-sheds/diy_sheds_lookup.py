#!/usr/bin/env python3
"""
DIY Sheds — Companies House Lookup Script

Reads an Excel file with company names (column B) and postcodes (column G),
looks up each company in the Companies House API, enriches with full profile
data (SIC codes, status, officers, charges, accounts, etc.), and exports
an enriched Excel + JSON for the analysis script.

Usage:
    python diy_sheds_lookup.py <path-to-excel-file>

Output:
    diy_sheds_enriched.xlsx  — original data + all Companies House columns
    diy_sheds_enriched.json  — structured JSON for the analysis script
"""

import os
import sys
import json
import re
from time import sleep
from pathlib import Path
from collections import Counter

import pandas as pd
import requests
from dotenv import load_dotenv

# Load .env from this folder or parent folder
load_dotenv()
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

API_KEY = os.getenv("COMPANIES_HOUSE_API_KEY")
BASE_URL = "https://api.company-information.service.gov.uk"

REQUEST_DELAY = 0.5  # seconds between API calls


# ═══════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def normalize_postcode(pc):
    """Normalize UK postcode for comparison."""
    if not pc or not isinstance(pc, str):
        return ""
    return re.sub(r"\s+", "", pc.strip().upper())


def normalize_name(name):
    """Normalize company name for matching."""
    if not name or not isinstance(name, str):
        return ""
    name = name.strip().upper()
    for suffix in [" LTD", " LIMITED", " PLC", " LLP", " INC", " & CO", " UK"]:
        if name.endswith(suffix):
            name = name[:-len(suffix)].strip()
    return name


def api_get(url, params=None, retries=3):
    """Authenticated GET with retry logic."""
    for attempt in range(retries):
        try:
            r = requests.get(url, auth=(API_KEY, ""), params=params, timeout=30)
            if r.status_code == 200:
                return r.json()
            if r.status_code == 429:
                wait = 2 ** (attempt + 2)
                print(f"    Rate limited, waiting {wait}s...")
                sleep(wait)
                continue
            if r.status_code == 404:
                return None
            print(f"    API {r.status_code} for {url}")
            return None
        except requests.RequestException as e:
            print(f"    Request error: {e}")
            if attempt < retries - 1:
                sleep(2 ** attempt)
    return None


def search_company(name):
    """Search Companies House by company name."""
    return api_get(
        f"{BASE_URL}/search/companies",
        params={"q": name, "items_per_page": 20},
    )


def get_company_profile(company_number):
    """Get full company profile."""
    return api_get(f"{BASE_URL}/company/{company_number}") or {}


def get_officer_count(company_number):
    """Get officer count for a company."""
    data = api_get(f"{BASE_URL}/company/{company_number}/officers")
    if data:
        return data.get("total_results", 0)
    return 0


def get_filing_history_count(company_number):
    """Get filing history count."""
    data = api_get(f"{BASE_URL}/company/{company_number}/filing-history",
                   params={"items_per_page": 0})
    if data:
        return data.get("total_count", 0)
    return 0


def match_by_postcode(search_results, target_postcode):
    """Find best matching company from search results based on postcode."""
    if not search_results or "items" not in search_results:
        return None

    target_pc = normalize_postcode(target_postcode)

    # Pass 1: exact postcode match
    for item in search_results["items"]:
        address = item.get("address", {})
        result_pc = normalize_postcode(address.get("postal_code", ""))
        if target_pc and result_pc and target_pc == result_pc:
            return item

    # Pass 2: partial match (outward code)
    if target_pc and len(target_pc) >= 3:
        prefix = target_pc[:4] if len(target_pc) > 5 else target_pc[:3]
        for item in search_results["items"]:
            address = item.get("address", {})
            result_pc = normalize_postcode(address.get("postal_code", ""))
            if result_pc and result_pc.startswith(prefix):
                return item

    # Pass 3: first active result
    for item in search_results["items"]:
        if item.get("company_status") == "active":
            return item

    # Fallback: first result
    return search_results["items"][0] if search_results["items"] else None


def extract_keywords(name):
    """Extract meaningful keywords from a company name."""
    if not name:
        return []
    name = name.upper()
    # Remove common legal suffixes and noise
    for remove in ["LTD", "LIMITED", "PLC", "LLP", "INC", "UK", "THE",
                    "AND", "&", "OF", "IN", "AT", "FOR"]:
        name = re.sub(rf"\b{remove}\b", "", name)
    # Split into words, keep only meaningful ones (2+ chars)
    words = re.findall(r"[A-Z]{2,}", name)
    return words


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN LOOKUP
# ═══════════════════════════════════════════════════════════════════════════

def run_lookup(excel_path):
    if not API_KEY or API_KEY == "your_api_key_here":
        print("Error: Set COMPANIES_HOUSE_API_KEY in your .env file")
        print("Get your key: https://developer.company-information.service.gov.uk/")
        sys.exit(1)

    if not os.path.exists(excel_path):
        print(f"Error: File not found: {excel_path}")
        sys.exit(1)

    # Read Excel
    print(f"Reading: {excel_path}")
    df = pd.read_excel(excel_path)
    cols = list(df.columns)
    print(f"Found {len(df)} rows, {len(cols)} columns")
    print(f"Columns: {cols}")

    # Column B = index 1, Column G = index 6
    name_col = cols[1] if len(cols) > 1 else cols[0]
    postcode_col = cols[6] if len(cols) > 6 else cols[-1]

    print(f"Using name column: '{name_col}' (column B)")
    print(f"Using postcode column: '{postcode_col}' (column G)")
    print(f"\nFirst 5 entries:")
    for i, row in df.head().iterrows():
        print(f"  {row[name_col]} | {row[postcode_col]}")

    # Lookup each company
    results = []
    cache = {}
    total = len(df)

    print(f"\nLooking up {total} companies...\n")

    for idx, row in df.iterrows():
        name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
        postcode = str(row[postcode_col]).strip() if pd.notna(row[postcode_col]) else ""

        print(f"[{idx+1}/{total}] {name} ({postcode})")

        if not name or name.lower() == "nan":
            results.append({
                "original_name": name,
                "original_postcode": postcode,
                "ch_found": False,
                "error": "Empty name",
            })
            continue

        # Cache check
        cache_key = f"{normalize_name(name)}|{normalize_postcode(postcode)}"
        if cache_key in cache:
            print(f"  [cached]")
            cached = cache[cache_key].copy()
            cached["original_name"] = name
            cached["original_postcode"] = postcode
            results.append(cached)
            continue

        # Search
        sleep(REQUEST_DELAY)
        search_data = search_company(name)

        if not search_data or not search_data.get("items"):
            result = {
                "original_name": name, "original_postcode": postcode,
                "ch_found": False, "error": "No results found",
            }
            results.append(result)
            cache[cache_key] = result
            continue

        # Match
        match = match_by_postcode(search_data, postcode)
        if not match:
            result = {
                "original_name": name, "original_postcode": postcode,
                "ch_found": False, "error": "No postcode match",
            }
            results.append(result)
            cache[cache_key] = result
            continue

        # Get full profile
        company_number = match.get("company_number", "")
        sleep(REQUEST_DELAY)
        profile = get_company_profile(company_number)

        # Get officer count
        sleep(REQUEST_DELAY)
        officer_count = get_officer_count(company_number)

        address = profile.get("registered_office_address", {})
        sic_codes = profile.get("sic_codes", [])
        accounts = profile.get("accounts", {})
        conf_stmt = profile.get("confirmation_statement", {})

        result = {
            "original_name": name,
            "original_postcode": postcode,
            "ch_found": True,
            "company_name": profile.get("company_name", ""),
            "company_number": company_number,
            "company_status": profile.get("company_status", "unknown"),
            "company_type": profile.get("type", ""),
            "date_of_creation": profile.get("date_of_creation", ""),
            "date_of_cessation": profile.get("date_of_cessation", ""),
            "sic_codes": ", ".join(sic_codes) if sic_codes else "",
            "sic_code_list": sic_codes,
            "registered_address": ", ".join(filter(None, [
                address.get("address_line_1", ""),
                address.get("address_line_2", ""),
                address.get("locality", ""),
                address.get("region", ""),
                address.get("postal_code", ""),
                address.get("country", ""),
            ])),
            "postal_code": address.get("postal_code", ""),
            "locality": address.get("locality", ""),
            "region": address.get("region", ""),
            "country": address.get("country", ""),
            "has_charges": profile.get("has_charges", False),
            "has_insolvency_history": profile.get("has_insolvency_history", False),
            "has_been_liquidated": profile.get("has_been_liquidated", False),
            "accounts_overdue": accounts.get("overdue", False),
            "last_accounts_type": accounts.get("last_accounts", {}).get("type", ""),
            "last_accounts_date": accounts.get("last_accounts", {}).get("made_up_to", ""),
            "next_accounts_due": accounts.get("next_due", ""),
            "confirmation_stmt_overdue": conf_stmt.get("overdue", False),
            "next_confirmation_due": conf_stmt.get("next_due", ""),
            "officer_count": officer_count,
            "name_keywords": extract_keywords(profile.get("company_name", "")),
        }

        results.append(result)
        cache[cache_key] = result

        status_icon = "✓" if result["company_status"] == "active" else "✗"
        print(f"  {status_icon} {result['company_name']} | "
              f"{result['company_status']} | SIC: {result['sic_codes']} | "
              f"Officers: {officer_count}")

    # ── Save outputs ────────────────────────────────────────────────────────
    output_dir = Path(__file__).resolve().parent
    enriched_xlsx = output_dir / "diy_sheds_enriched.xlsx"
    enriched_json = output_dir / "diy_sheds_enriched.json"

    # Excel output
    enriched_df = df.copy()
    export_fields = [
        "ch_found", "company_name", "company_number", "company_status",
        "company_type", "date_of_creation", "date_of_cessation",
        "sic_codes", "registered_address", "postal_code", "locality",
        "region", "country", "has_charges", "has_insolvency_history",
        "has_been_liquidated", "accounts_overdue", "last_accounts_type",
        "last_accounts_date", "next_accounts_due",
        "confirmation_stmt_overdue", "next_confirmation_due",
        "officer_count",
    ]
    for field in export_fields:
        enriched_df[f"CH_{field}"] = [r.get(field, "") for r in results]

    enriched_df.to_excel(enriched_xlsx, index=False)
    print(f"\nEnriched Excel saved: {enriched_xlsx}")

    # JSON output (for analysis script)
    # Convert non-serializable types
    json_results = []
    for r in results:
        jr = {}
        for k, v in r.items():
            if isinstance(v, bool):
                jr[k] = v
            elif isinstance(v, (list, dict)):
                jr[k] = v
            else:
                jr[k] = str(v) if v is not None else ""
        json_results.append(jr)

    with open(enriched_json, "w", encoding="utf-8") as f:
        json.dump(json_results, f, indent=2, ensure_ascii=False)
    print(f"Enriched JSON saved: {enriched_json}")

    # Quick summary
    found = [r for r in results if r.get("ch_found")]
    statuses = Counter(r.get("company_status") for r in found)
    print(f"\n{'─' * 50}")
    print(f"QUICK SUMMARY")
    print(f"{'─' * 50}")
    print(f"  Total rows:    {total}")
    print(f"  Found in CH:   {len(found)}")
    print(f"  Not found:     {total - len(found)}")
    for status, count in statuses.most_common():
        print(f"  {status}: {count}")
    print(f"\nRun the analysis script next:")
    print(f"  python diy_sheds_analysis.py")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python diy_sheds_lookup.py <path-to-excel-file>")
        sys.exit(1)
    run_lookup(sys.argv[1])
