#!/usr/bin/env python3
"""
Fast Fit Car Specialists — Companies House Discovery Script

Searches Companies House for Fast Fit Car Specialists (Autocenters) using
SIC codes, keyword scoring, and exclusion filters to separate genuine fast-fit
operations from noise (tyre-only shops, independents, body shops, etc.).

Definition:
  Fast Fit Car Specialists generate >50% of turnover from quick automotive
  repairs/services (exhaust, brakes, oil, filters). They do NOT offer full
  car service/paintwork. May also sell spare parts and accessories.
  Only organized retailers — independents are classified as Car Garages.

Excluded: Tyre Specialists, Car Accessory Retailers, Body Shops, Dealerships.

Usage:
    pip install -r requirements.txt
    # Set COMPANIES_HOUSE_API_KEY in ../.env or ./env
    python fast_fit_lookup.py
"""

import os
import sys
import csv
import json
from time import sleep
from pathlib import Path

import requests
from dotenv import load_dotenv

# Load .env from this folder or parent folder
load_dotenv()
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

API_KEY = os.getenv("COMPANIES_HOUSE_API_KEY")
BASE_URL = "https://api.company-information.service.gov.uk"

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  SIC CODE STRATEGY                                                      ║
# ║                                                                          ║
# ║  45200 — Maintenance and repair of motor vehicles                        ║
# ║          This is the PRIMARY code. Most fast-fit centres register here.  ║
# ║          Covers: exhaust, brakes, oil change, MOT, general servicing.   ║
# ║          Noise: also used by full garages, independents, body shops.    ║
# ║                                                                          ║
# ║  45319 — Wholesale of motor vehicle parts and accessories (other)        ║
# ║          Some fast-fit chains also wholesale parts. Secondary signal.    ║
# ║                                                                          ║
# ║  45320 — Retail trade of motor vehicle parts and accessories             ║
# ║          Fast-fit shops that also sell parts/accessories on a shop       ║
# ║          floor or online. BUT this code alone = likely tyre/accessory   ║
# ║          retailer, so we penalise tyre-only SIC profiles.               ║
# ║                                                                          ║
# ║  NOT USED (and why):                                                     ║
# ║  45310 — Wholesale of parts (too broad, mostly wholesalers)              ║
# ║  45400 — Sale/maintenance of motorcycles (wrong segment)                 ║
# ║  45111 — Sale of cars (dealerships — excluded)                           ║
# ║  45190 — Sale of other motor vehicles (commercial — excluded)            ║
# ╚══════════════════════════════════════════════════════════════════════════╝

SIC_CODES = ["45200", "45319", "45320"]

# ── HARD EXCLUDES ───────────────────────────────────────────────────────────
# If ANY of these appear in the company name → discard immediately.
# These are business types that share SIC codes but are NOT fast-fit.

HARD_EXCLUDES = [
    # Body work / paint — full service, not fast-fit
    "body shop", "bodywork", "panel beating", "respray", "paintwork",
    "coachwork", "spray", "refinish",
    # Car sales / dealerships
    "car sales", "used cars", "forecourt", "dealership", "dealer",
    "motor trade", "car supermarket",
    # Car wash / valeting
    "car wash", "valeting", "detailing", "hand wash",
    # Recovery / breakdown
    "recovery", "breakdown", "rescue", "towing",
    # Scrap / salvage
    "scrap", "salvage", "dismantler", "breaker",
    # Rental / hire
    "rental", "hire", "leasing", "rent a car",
    # Tuning / performance (niche, not fast-fit)
    "tuning", "performance custom", "motorsport", "racing",
    # Driving schools
    "driving school", "driving instructor",
    # Insurance / claims
    "insurance", "claims",
    # Parking
    "parking", "car park",
]

# ── POSITIVE KEYWORDS ──────────────────────────────────────────────────────
# Score companies higher if their name contains these signals.
# Higher points = stronger signal that this is a fast-fit operation.

POSITIVE_KEYWORDS = {
    # +3 — Brand names and concept terms (very strong signal)
    "fast fit": 3, "fastfit": 3, "kwik fit": 3, "kwikfit": 3,
    "autocentre": 3, "auto centre": 3, "autocenter": 3, "auto center": 3,
    "pit stop": 3, "pitstop": 3,
    "midas": 3, "speedy": 3, "norauto": 3, "halfords autocentre": 3,
    "national tyres": 3, "mr clutch": 3, "formula one autocentres": 3,
    "ats euromaster": 3,

    # +2 — Service-type signals (strong indicator of fast-fit work)
    "exhaust": 2, "brake": 2, "brakes": 2, "lube": 2, "oil change": 2,
    "fitting centre": 2, "fitting center": 2, "mot centre": 2, "mot center": 2,
    "service centre": 2, "service center": 2, "quick fit": 2, "quickfit": 2,
    "tyre exhaust": 2, "tyre brake": 2, "tyres exhaust": 2, "tyres brake": 2,
    "clutch": 2, "muffler": 2, "silencer": 2,
    "auto repair": 2, "car servicing": 2,

    # +1 — Generic signals (weak, need corroboration)
    "tyre": 1, "tyres": 1, "repair": 1, "servicing": 1,
    "workshop": 1, "automotive": 1, "motor": 1, "garage": 1,
    "mot": 1, "car care": 1,
}

# ── PENALTY KEYWORDS ───────────────────────────────────────────────────────
# Reduce score if name suggests non-fast-fit or independent operation.

PENALTY_KEYWORDS = {
    "independent": -2,
    "mobile": -2,       # mobile mechanics are not fast-fit centres
    "roadside": -2,
    "accessories": -1,  # might be accessory-only retailer
    "wholesale": -1,
    "distribution": -1,
    "parts only": -1,
    "consultant": -2,
    "training": -2,
}

# ── SCORING THRESHOLDS ─────────────────────────────────────────────────────
SCORE_KEEP = 2      # net score >= 2 → auto-keep
SCORE_REVIEW = 0    # net score 0–1 → flag for manual review
                    # net score < 0 → discard


# ═══════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════

def name_lower(company):
    """Get lowercased company name from search result."""
    return company.get("title", company.get("company_name", "")).lower()


def hard_exclude(company):
    """Return True if this company should be excluded outright."""
    name = name_lower(company)

    # Only keep active companies
    status = company.get("company_status", "").lower()
    if status != "active":
        return True

    # Check all exclusion phrases
    for phrase in HARD_EXCLUDES:
        if phrase in name:
            return True

    return False


def positive_score(company):
    """Calculate positive score from name keywords."""
    name = name_lower(company)
    score = 0
    reasons = []
    for kw, pts in POSITIVE_KEYWORDS.items():
        if kw in name:
            score += pts
            reasons.append(f"+{pts}:{kw}")
    return score, reasons


def penalty_score(company, sic_codes, officer_count):
    """Calculate penalty score from name, SIC profile, and officers."""
    name = name_lower(company)
    score = 0
    reasons = []

    # Tyre-only SIC penalty — if the ONLY SIC code is 45320 (retail parts)
    # this is likely a tyre/accessory retailer, not a fast-fit centre
    if set(sic_codes) == {"45320"}:
        score -= 1
        reasons.append("-1:tyre-only-sic")

    # Keyword penalties
    for kw, pts in PENALTY_KEYWORDS.items():
        if kw in name:
            score += pts
            reasons.append(f"{pts}:{kw}")

    # Single officer → likely independent garage, not organized chain
    # Fast-fit chains typically have multiple directors/secretaries
    if officer_count <= 1:
        score -= 2
        reasons.append("-2:single-officer(likely-independent)")

    return score, reasons


def api_get(url, params=None, retries=3):
    """Make an authenticated GET request with retry logic."""
    for attempt in range(retries):
        try:
            r = requests.get(url, auth=(API_KEY, ""), params=params, timeout=30)
            if r.status_code == 200:
                return r.json()
            if r.status_code == 429:
                wait = 2 ** (attempt + 1)
                print(f"  Rate limited, waiting {wait}s...")
                sleep(wait)
                continue
            if r.status_code == 404:
                return None
            print(f"  API returned {r.status_code} for {url}")
            return None
        except requests.RequestException as e:
            print(f"  Request error: {e}")
            if attempt < retries - 1:
                sleep(2 ** attempt)
    return None


def get_officer_count(company_number):
    """Get the number of officers for a company."""
    data = api_get(f"{BASE_URL}/company/{company_number}/officers")
    if data:
        return data.get("total_results", 0)
    return 0


def fetch_full_profile(company_number):
    """Get the full company profile."""
    return api_get(f"{BASE_URL}/company/{company_number}") or {}


def search_by_sic(sic_code, start_index=0, items_per_page=100):
    """Search for companies by SIC code using advanced search."""
    data = api_get(
        f"{BASE_URL}/advanced-search/companies",
        params={
            "sic_codes": sic_code,
            "company_status": "active",
            "size": items_per_page,
            "start_index": start_index,
        },
    )
    if data:
        return data.get("items", [])
    return []


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN PIPELINE
# ═══════════════════════════════════════════════════════════════════════════

def run():
    if not API_KEY or API_KEY == "your_api_key_here":
        print("Error: Set COMPANIES_HOUSE_API_KEY in your .env file")
        print("Get your key from: https://developer.company-information.service.gov.uk/")
        sys.exit(1)

    seen = set()
    results = []
    stats = {
        "total_searched": 0,
        "hard_excluded": 0,
        "score_discarded": 0,
        "kept": 0,
        "review": 0,
    }

    print("=" * 70)
    print("  Fast Fit Car Specialists — Companies House Discovery")
    print("=" * 70)
    print(f"\nSIC codes to search: {', '.join(SIC_CODES)}")
    print(f"Score threshold: KEEP >= {SCORE_KEEP}, REVIEW >= {SCORE_REVIEW}")
    print()

    for sic in SIC_CODES:
        print(f"\n{'─' * 50}")
        print(f"Searching SIC code: {sic}")
        print(f"{'─' * 50}")

        start = 0
        page = 0

        while True:
            companies = search_by_sic(sic, start_index=start)
            if not companies:
                print(f"  No more results for SIC {sic} (page {page})")
                break

            page += 1
            print(f"  Page {page}: fetched {len(companies)} companies (start={start})")

            for company in companies:
                number = company.get("company_number")
                if not number or number in seen:
                    continue
                seen.add(number)
                stats["total_searched"] += 1

                name = name_lower(company)

                # Step 1: Hard exclude
                if hard_exclude(company):
                    stats["hard_excluded"] += 1
                    continue

                # Step 2: Positive scoring
                pos, pos_reasons = positive_score(company)

                # Step 3: Get officer count for penalty scoring
                officer_count = get_officer_count(number)
                sleep(0.3)

                # Extract SIC codes from the search result
                sic_list = []
                sic_field = company.get("sic_codes", [])
                if isinstance(sic_field, list):
                    for s in sic_field:
                        if isinstance(s, dict):
                            sic_list.append(s.get("sic_code", ""))
                        elif isinstance(s, str):
                            sic_list.append(s)

                # Step 4: Penalty scoring
                neg, neg_reasons = penalty_score(company, sic_list, officer_count)

                net = pos + neg

                # Step 5: Threshold check
                if net < SCORE_REVIEW:
                    stats["score_discarded"] += 1
                    continue

                flag = "KEEP" if net >= SCORE_KEEP else "REVIEW"

                if flag == "KEEP":
                    stats["kept"] += 1
                else:
                    stats["review"] += 1

                # Step 6: Fetch full profile for enrichment
                profile = fetch_full_profile(number)
                sleep(0.3)

                address = profile.get("registered_office_address", {})
                address_str = ", ".join(filter(None, [
                    address.get("address_line_1", ""),
                    address.get("address_line_2", ""),
                    address.get("locality", ""),
                    address.get("region", ""),
                    address.get("postal_code", ""),
                    address.get("country", ""),
                ]))

                profile_sic = profile.get("sic_codes", sic_list)
                if isinstance(profile_sic, list):
                    profile_sic_str = ", ".join(str(s) for s in profile_sic)
                else:
                    profile_sic_str = str(profile_sic)

                results.append({
                    "company_number": number,
                    "company_name": profile.get("company_name", company.get("title", "")),
                    "status": profile.get("company_status", ""),
                    "company_type": profile.get("type", ""),
                    "sic_codes": profile_sic_str,
                    "registered_address": address_str,
                    "postal_code": address.get("postal_code", ""),
                    "locality": address.get("locality", ""),
                    "region": address.get("region", ""),
                    "date_created": profile.get("date_of_creation", ""),
                    "officer_count": officer_count,
                    "has_charges": profile.get("has_charges", False),
                    "has_insolvency_history": profile.get("has_insolvency_history", False),
                    "accounts_overdue": profile.get("accounts", {}).get("overdue", False),
                    "net_score": net,
                    "score_reasons": " | ".join(pos_reasons + neg_reasons),
                    "flag": flag,
                })

                print(f"    [{flag}] {profile.get('company_name', name)} "
                      f"(score={net}, officers={officer_count})")

            start += items_per_page
            if len(companies) < items_per_page:
                break

    # ── Sort by score and export ────────────────────────────────────────────
    results.sort(key=lambda x: x["net_score"], reverse=True)

    output_file = Path(__file__).resolve().parent / "fast_fit_companies.csv"

    if results:
        with open(output_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=results[0].keys())
            writer.writeheader()
            writer.writerows(results)

    # ── Print summary ───────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("  RESULTS SUMMARY")
    print("=" * 70)
    print(f"  Total unique companies searched: {stats['total_searched']}")
    print(f"  Hard excluded (inactive/wrong type): {stats['hard_excluded']}")
    print(f"  Score discarded (net < {SCORE_REVIEW}): {stats['score_discarded']}")
    print(f"  ─────────────────────────────────")
    print(f"  KEEP (score >= {SCORE_KEEP}):   {stats['kept']}")
    print(f"  REVIEW (score {SCORE_REVIEW}-{SCORE_KEEP-1}): {stats['review']}")
    print(f"  Total output:           {len(results)}")
    print(f"\n  Output file: {output_file}")

    # ── Quick analysis of results ───────────────────────────────────────────
    if results:
        print(f"\n{'─' * 50}")
        print("  TOP COMPANIES BY SCORE")
        print(f"{'─' * 50}")
        for r in results[:20]:
            print(f"  [{r['flag']}] {r['company_name']}")
            print(f"        Score: {r['net_score']} | SIC: {r['sic_codes']} | "
                  f"Officers: {r['officer_count']}")
            print(f"        Reasons: {r['score_reasons']}")

        # Regional breakdown
        from collections import Counter
        regions = Counter(r.get("region", "Unknown") or "Unknown" for r in results)
        print(f"\n{'─' * 50}")
        print("  REGIONAL BREAKDOWN")
        print(f"{'─' * 50}")
        for region, count in regions.most_common(15):
            print(f"  {region:<30} {count:>4}")

    print(f"\n{'=' * 70}")
    print("  Done!")
    print(f"{'=' * 70}\n")


if __name__ == "__main__":
    run()
