#!/usr/bin/env python3
"""
Companies House Lookup & Analysis Script

Reads an Excel file with company names (column A) and postcodes (column E),
looks up each company in the Companies House API, enriches the data,
and produces an analysis report.

Usage:
    python companies_house_lookup.py <path-to-excel-file>
"""

import os
import sys
import time
import json
import re
from collections import Counter
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

API_KEY = os.getenv("COMPANIES_HOUSE_API_KEY")
BASE_URL = "https://api.company-information.service.gov.uk"

# Rate limiting: 600 requests per 5 minutes = 2 per second max
REQUEST_DELAY = 0.5  # seconds between requests

# SIC code descriptions (common ones for furniture/retail)
SIC_DESCRIPTIONS = {
    "31010": "Manufacture of office and shop furniture",
    "31020": "Manufacture of kitchen furniture",
    "31090": "Manufacture of other furniture",
    "46470": "Wholesale of furniture, carpets and lighting equipment",
    "47591": "Retail sale of furniture and lighting equipment in specialised stores",
    "47599": "Retail of furniture/furnishings (other) in specialised stores",
    "47540": "Retail sale of electrical household appliances in specialised stores",
    "47190": "Other retail sale in non-specialised stores",
    "47910": "Retail sale via mail order houses or via Internet",
    "47990": "Other retail sale not in stores, stalls or markets",
    "82990": "Other business support service activities n.e.c.",
    "68100": "Buying and selling of own real estate",
    "68209": "Other letting and operating of own or leased real estate",
    "70100": "Activities of head offices",
    "70210": "Public relations and communication activities",
    "74909": "Other professional, scientific and technical activities n.e.c.",
    "96090": "Other personal service activities n.e.c.",
}


def normalize_postcode(pc):
    """Normalize a UK postcode for comparison."""
    if not pc or not isinstance(pc, str):
        return ""
    return re.sub(r"\s+", "", pc.strip().upper())


def normalize_name(name):
    """Normalize a company name for comparison and grouping."""
    if not name or not isinstance(name, str):
        return ""
    name = name.strip().upper()
    # Remove common suffixes for grouping
    for suffix in [" LTD", " LIMITED", " PLC", " LLP", " INC", " CO", " & CO"]:
        if name.endswith(suffix):
            name = name[: -len(suffix)].strip()
    return name


def search_company(name, api_key):
    """Search Companies House for a company by name."""
    url = f"{BASE_URL}/search/companies"
    params = {"q": name, "items_per_page": 20}
    try:
        resp = requests.get(url, params=params, auth=(api_key, ""), timeout=30)
        if resp.status_code == 429:
            print("  Rate limited, waiting 60 seconds...")
            time.sleep(60)
            resp = requests.get(url, params=params, auth=(api_key, ""), timeout=30)
        if resp.status_code == 200:
            return resp.json()
        else:
            print(f"  Search API returned {resp.status_code} for '{name}'")
            return None
    except requests.RequestException as e:
        print(f"  Request error searching '{name}': {e}")
        return None


def get_company_profile(company_number, api_key):
    """Get full company profile from Companies House."""
    url = f"{BASE_URL}/company/{company_number}"
    try:
        resp = requests.get(url, auth=(api_key, ""), timeout=30)
        if resp.status_code == 429:
            print("  Rate limited, waiting 60 seconds...")
            time.sleep(60)
            resp = requests.get(url, auth=(api_key, ""), timeout=30)
        if resp.status_code == 200:
            return resp.json()
        else:
            print(f"  Profile API returned {resp.status_code} for '{company_number}'")
            return None
    except requests.RequestException as e:
        print(f"  Request error for company {company_number}: {e}")
        return None


def match_by_postcode(search_results, target_postcode):
    """Find the best matching company from search results based on postcode."""
    if not search_results or "items" not in search_results:
        return None

    target_pc = normalize_postcode(target_postcode)

    # First pass: exact postcode match
    for item in search_results["items"]:
        address = item.get("address", {})
        result_pc = normalize_postcode(address.get("postal_code", ""))
        if target_pc and result_pc and target_pc == result_pc:
            return item

    # Second pass: partial postcode match (first half)
    if target_pc and len(target_pc) >= 3:
        target_prefix = target_pc[:4] if len(target_pc) > 5 else target_pc[:3]
        for item in search_results["items"]:
            address = item.get("address", {})
            result_pc = normalize_postcode(address.get("postal_code", ""))
            if result_pc and result_pc.startswith(target_prefix):
                return item

    # Fallback: return first active result, or first result
    for item in search_results["items"]:
        if item.get("company_status") == "active":
            return item

    return search_results["items"][0] if search_results["items"] else None


def get_sic_description(code):
    """Get SIC code description."""
    return SIC_DESCRIPTIONS.get(code, f"SIC {code}")


def lookup_companies(df, name_col, postcode_col):
    """Look up all companies in the DataFrame via Companies House API."""
    results = []
    cache = {}  # Cache by normalized name to avoid duplicate API calls
    total = len(df)

    print(f"\nLooking up {total} companies in Companies House API...\n")

    for idx, row in df.iterrows():
        name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
        postcode = str(row[postcode_col]).strip() if pd.notna(row[postcode_col]) else ""

        progress = idx + 1
        print(f"[{progress}/{total}] Looking up: {name} ({postcode})")

        if not name:
            results.append({
                "original_name": name,
                "original_postcode": postcode,
                "ch_found": False,
                "error": "Empty name",
            })
            continue

        # Check cache
        cache_key = f"{normalize_name(name)}|{normalize_postcode(postcode)}"
        if cache_key in cache:
            print(f"  Using cached result")
            results.append(cache[cache_key].copy())
            results[-1]["original_name"] = name
            results[-1]["original_postcode"] = postcode
            continue

        # Search Companies House
        time.sleep(REQUEST_DELAY)
        search_data = search_company(name, API_KEY)

        if not search_data or not search_data.get("items"):
            result = {
                "original_name": name,
                "original_postcode": postcode,
                "ch_found": False,
                "error": "No results found",
            }
            results.append(result)
            cache[cache_key] = result
            continue

        # Match by postcode
        match = match_by_postcode(search_data, postcode)

        if not match:
            result = {
                "original_name": name,
                "original_postcode": postcode,
                "ch_found": False,
                "error": "No postcode match",
            }
            results.append(result)
            cache[cache_key] = result
            continue

        # Get full company profile
        company_number = match.get("company_number", "")
        time.sleep(REQUEST_DELAY)
        profile = get_company_profile(company_number, API_KEY)

        if profile:
            sic_codes = profile.get("sic_codes", [])
            sic_descriptions = [get_sic_description(c) for c in sic_codes]
            address = profile.get("registered_office_address", {})

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
                "sic_codes": ", ".join(sic_codes),
                "sic_descriptions": ", ".join(sic_descriptions),
                "registered_address": ", ".join(
                    filter(None, [
                        address.get("address_line_1", ""),
                        address.get("address_line_2", ""),
                        address.get("locality", ""),
                        address.get("region", ""),
                        address.get("postal_code", ""),
                    ])
                ),
                "has_charges": profile.get("has_charges", False),
                "has_insolvency_history": profile.get("has_insolvency_history", False),
                "accounts_overdue": profile.get("accounts", {}).get("overdue", False),
                "last_accounts_type": profile.get("accounts", {}).get("last_accounts", {}).get("type", ""),
                "confirmation_statement_overdue": profile.get("confirmation_statement", {}).get("overdue", False),
            }
        else:
            # Use search result data as fallback
            address = match.get("address", {})
            result = {
                "original_name": name,
                "original_postcode": postcode,
                "ch_found": True,
                "company_name": match.get("title", ""),
                "company_number": company_number,
                "company_status": match.get("company_status", "unknown"),
                "company_type": match.get("company_type", ""),
                "date_of_creation": match.get("date_of_creation", ""),
                "date_of_cessation": "",
                "sic_codes": "",
                "sic_descriptions": "",
                "registered_address": ", ".join(
                    filter(None, [
                        address.get("address_line_1", ""),
                        address.get("address_line_2", ""),
                        address.get("locality", ""),
                        address.get("region", ""),
                        address.get("postal_code", ""),
                    ])
                ),
                "has_charges": False,
                "has_insolvency_history": False,
                "accounts_overdue": False,
                "last_accounts_type": "",
                "confirmation_statement_overdue": False,
            }

        results.append(result)
        cache[cache_key] = result
        print(f"  -> {result['company_name']} | Status: {result['company_status']} | SIC: {result['sic_codes']}")

    return results


def detect_chains(results):
    """Detect chain stores - companies appearing multiple times."""
    name_counts = Counter()
    company_number_counts = Counter()

    for r in results:
        if r.get("ch_found"):
            norm = normalize_name(r.get("company_name", "") or r.get("original_name", ""))
            if norm:
                name_counts[norm] += 1
            cn = r.get("company_number", "")
            if cn:
                company_number_counts[cn] += 1

    chain_names = {name for name, count in name_counts.items() if count >= 2}
    chain_numbers = {num for num, count in company_number_counts.items() if count >= 2}

    for r in results:
        norm = normalize_name(r.get("company_name", "") or r.get("original_name", ""))
        cn = r.get("company_number", "")
        r["is_chain"] = norm in chain_names or cn in chain_numbers
        if norm in chain_names:
            r["chain_count"] = name_counts[norm]
        elif cn in chain_numbers:
            r["chain_count"] = company_number_counts[cn]
        else:
            r["chain_count"] = 1

    return results


def print_analysis(results):
    """Print comprehensive analysis of the results."""
    total = len(results)
    found = [r for r in results if r.get("ch_found")]
    not_found = [r for r in results if not r.get("ch_found")]

    print("\n" + "=" * 80)
    print("                    COMPANIES HOUSE ANALYSIS REPORT")
    print("=" * 80)

    # --- Overview ---
    print(f"\n{'─' * 40}")
    print("OVERVIEW")
    print(f"{'─' * 40}")
    print(f"  Total entries in file:        {total}")
    print(f"  Found in Companies House:     {len(found)}")
    print(f"  Not found:                    {len(not_found)}")
    print(f"  Match rate:                   {len(found)/total*100:.1f}%" if total > 0 else "")

    # --- Status Breakdown ---
    print(f"\n{'─' * 40}")
    print("COMPANY STATUS BREAKDOWN")
    print(f"{'─' * 40}")
    status_counts = Counter(r.get("company_status", "unknown") for r in found)
    for status, count in status_counts.most_common():
        pct = count / len(found) * 100 if found else 0
        print(f"  {status:<30} {count:>5}  ({pct:.1f}%)")

    active = status_counts.get("active", 0)
    dissolved = status_counts.get("dissolved", 0)
    liquidation = status_counts.get("liquidation", 0) + status_counts.get("voluntary-arrangement", 0)

    print(f"\n  >> ACTIVE shops:              {active}")
    print(f"  >> CLOSED/DISSOLVED shops:    {dissolved}")
    print(f"  >> IN LIQUIDATION/TROUBLE:    {liquidation}")

    # --- SIC Code Analysis ---
    print(f"\n{'─' * 40}")
    print("SIC CODE ANALYSIS")
    print(f"{'─' * 40}")
    all_sic = []
    for r in found:
        codes = r.get("sic_codes", "")
        if codes:
            all_sic.extend([c.strip() for c in codes.split(",")])

    sic_counts = Counter(all_sic)
    print(f"  Total unique SIC codes found: {len(sic_counts)}")
    print(f"\n  Most common SIC codes:")
    for code, count in sic_counts.most_common(15):
        desc = get_sic_description(code)
        pct = count / len(found) * 100 if found else 0
        print(f"    {code}: {desc}")
        print(f"         Used by {count} companies ({pct:.1f}%)")

    # --- Name Analysis ---
    print(f"\n{'─' * 40}")
    print("MOST COMMON COMPANY NAMES")
    print(f"{'─' * 40}")
    name_counts = Counter()
    for r in results:
        name = r.get("company_name") or r.get("original_name", "")
        norm = normalize_name(name)
        if norm:
            name_counts[norm] += 1

    for name, count in name_counts.most_common(20):
        print(f"  {name:<45} {count:>3} location(s)")

    # --- Chain Analysis ---
    print(f"\n{'─' * 40}")
    print("CHAIN STORE ANALYSIS")
    print(f"{'─' * 40}")
    chains = [r for r in results if r.get("is_chain")]
    independents = [r for r in results if not r.get("is_chain")]
    print(f"  Chain store entries:          {len(chains)}")
    print(f"  Independent store entries:    {len(independents)}")

    chain_names_detail = {}
    for r in chains:
        name = normalize_name(r.get("company_name") or r.get("original_name", ""))
        if name not in chain_names_detail:
            chain_names_detail[name] = {"count": 0, "statuses": []}
        chain_names_detail[name]["count"] += 1
        chain_names_detail[name]["statuses"].append(r.get("company_status", "unknown"))

    if chain_names_detail:
        print(f"\n  Chain stores identified:")
        for name, info in sorted(chain_names_detail.items(), key=lambda x: -x[1]["count"]):
            status_summary = Counter(info["statuses"])
            status_str = ", ".join(f"{s}: {c}" for s, c in status_summary.most_common())
            print(f"    {name}: {info['count']} locations ({status_str})")

    # --- Company Type ---
    print(f"\n{'─' * 40}")
    print("COMPANY TYPE BREAKDOWN")
    print(f"{'─' * 40}")
    type_counts = Counter(r.get("company_type", "unknown") for r in found)
    for ctype, count in type_counts.most_common():
        print(f"  {ctype:<35} {count:>5}")

    # --- Insolvency & Charges ---
    print(f"\n{'─' * 40}")
    print("FINANCIAL HEALTH INDICATORS")
    print(f"{'─' * 40}")
    insolvency = sum(1 for r in found if r.get("has_insolvency_history"))
    charges = sum(1 for r in found if r.get("has_charges"))
    overdue_accounts = sum(1 for r in found if r.get("accounts_overdue"))
    overdue_cs = sum(1 for r in found if r.get("confirmation_statement_overdue"))

    print(f"  With insolvency history:      {insolvency}")
    print(f"  With charges (loans/security):{charges}")
    print(f"  Accounts overdue:             {overdue_accounts}")
    print(f"  Confirmation stmt overdue:    {overdue_cs}")

    # --- Incorporation Timeline ---
    print(f"\n{'─' * 40}")
    print("INCORPORATION TIMELINE")
    print(f"{'─' * 40}")
    years = []
    for r in found:
        doc = r.get("date_of_creation", "")
        if doc and len(doc) >= 4:
            try:
                years.append(int(doc[:4]))
            except ValueError:
                pass

    if years:
        year_counts = Counter(years)
        decades = Counter()
        for y, c in year_counts.items():
            decade = (y // 10) * 10
            decades[decade] += c

        for decade in sorted(decades.keys()):
            bar = "█" * decades[decade]
            print(f"  {decade}s: {decades[decade]:>4} companies  {bar}")

        print(f"\n  Oldest company: incorporated {min(years)}")
        print(f"  Newest company: incorporated {max(years)}")

    # --- Closure Analysis ---
    print(f"\n{'─' * 40}")
    print("CLOSURE ANALYSIS")
    print(f"{'─' * 40}")
    dissolved_companies = [r for r in found if r.get("company_status") == "dissolved"]
    if dissolved_companies:
        cessation_years = []
        for r in dissolved_companies:
            doc = r.get("date_of_cessation", "")
            if doc and len(doc) >= 4:
                try:
                    cessation_years.append(int(doc[:4]))
                except ValueError:
                    pass

        if cessation_years:
            year_counts = Counter(cessation_years)
            print(f"  Closures by year:")
            for year in sorted(year_counts.keys()):
                bar = "█" * year_counts[year]
                print(f"    {year}: {year_counts[year]:>4} closures  {bar}")
        else:
            print(f"  {len(dissolved_companies)} dissolved companies (cessation dates not available)")
    else:
        print("  No dissolved companies found")

    print(f"\n{'=' * 80}")
    print("                         END OF REPORT")
    print(f"{'=' * 80}\n")


def main():
    if len(sys.argv) < 2:
        print("Usage: python companies_house_lookup.py <path-to-excel-file>")
        print("\nExample: python companies_house_lookup.py data/furniture_shops.xlsx")
        sys.exit(1)

    excel_path = sys.argv[1]

    if not os.path.exists(excel_path):
        print(f"Error: File not found: {excel_path}")
        sys.exit(1)

    if not API_KEY or API_KEY == "your_api_key_here":
        print("Error: Please set COMPANIES_HOUSE_API_KEY in your .env file")
        print("Get your API key from: https://developer.company-information.service.gov.uk/")
        sys.exit(1)

    # Read Excel file
    print(f"Reading Excel file: {excel_path}")
    df = pd.read_excel(excel_path)

    print(f"Found {len(df)} rows")
    print(f"Columns: {list(df.columns)}")

    # Determine column names (A = first column, E = fifth column)
    all_cols = list(df.columns)
    name_col = all_cols[0]  # Column A
    postcode_col = all_cols[4] if len(all_cols) > 4 else all_cols[-1]  # Column E

    print(f"Using name column: '{name_col}' (column A)")
    print(f"Using postcode column: '{postcode_col}' (column E)")

    # Preview first few rows
    print(f"\nFirst 5 entries:")
    for i, row in df.head().iterrows():
        print(f"  {row[name_col]} | {row[postcode_col]}")

    # Lookup companies
    results = lookup_companies(df, name_col, postcode_col)

    # Detect chains
    results = detect_chains(results)

    # Print analysis
    print_analysis(results)

    # Save enriched data
    output_dir = Path(excel_path).parent
    output_filename = Path(excel_path).stem + "_enriched.xlsx"
    output_path = output_dir / output_filename

    # Create enriched DataFrame
    enriched_df = df.copy()
    result_fields = [
        "ch_found", "company_name", "company_number", "company_status",
        "company_type", "date_of_creation", "date_of_cessation",
        "sic_codes", "sic_descriptions", "registered_address",
        "has_charges", "has_insolvency_history", "accounts_overdue",
        "is_chain", "chain_count",
    ]

    for field in result_fields:
        enriched_df[f"CH_{field}"] = [r.get(field, "") for r in results]

    enriched_df.to_excel(output_path, index=False)
    print(f"\nEnriched data saved to: {output_path}")

    # Also save analysis as CSV
    analysis_path = output_dir / (Path(excel_path).stem + "_analysis.csv")
    pd.DataFrame(results).to_csv(analysis_path, index=False)
    print(f"Raw results saved to: {analysis_path}")


if __name__ == "__main__":
    main()
