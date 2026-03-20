#!/usr/bin/env python3
"""
Fast Fit Car Specialists — Noise Detection Script

Reads fast_fit_companies.csv from fast_fit_lookup.py, deduplicates by
company number, then classifies every company as KEEP / REVIEW / NOISE
based on SIC codes, name keywords, and structural signals.

Definition reminder:
  Fast Fit = >50% turnover from quick automotive repairs (exhaust, brakes,
  oil, filters). NOT full-service garages, NOT body shops, NOT tyre-only,
  NOT dealerships, NOT independents. Only organized retailers.

Outputs:
  - Console report showing what's noise and why
  - fast_fit_noise_report.xlsx with sheets: Clean, Review, Noise, Summary

Usage:
    python fast_fit_noise_check.py                          # default CSV
    python fast_fit_noise_check.py custom_results.csv       # custom file
"""

import csv
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd


# ═══════════════════════════════════════════════════════════════════════════
#  SIC CODE CLASSIFICATION
# ═══════════════════════════════════════════════════════════════════════════

# GREEN — SIC codes that strongly indicate fast-fit / auto repair
SIC_GREEN = {
    "45200",  # Maintenance and repair of motor vehicles — PRIMARY
}

# AMBER — could be fast-fit but needs keyword confirmation
SIC_AMBER = {
    "45320",  # Retail trade of motor vehicle parts/accessories
    "45319",  # Wholesale of motor vehicle parts (other)
    "45310",  # Wholesale of motor vehicle parts
    "45400",  # Sale/maintenance of motorcycles (edge case)
}

# RED — SIC codes that are NOT fast-fit
SIC_RED = {
    # Car sales / dealerships
    "45111",  # Sale of cars and light motor vehicles
    "45112",  # Agents for sale of motor vehicles
    "45190",  # Sale of other motor vehicles
    # Non-automotive retail
    "47190",  # Other retail sale in non-specialised stores
    "47520",  # Retail: hardware, paints, glass
    "47910",  # Retail via mail order / internet
    "47990",  # Other retail not in stores
    # Real estate
    "68100",  # Buying and selling of own real estate
    "68209",  # Other letting of own/leased real estate
    "68320",  # Management of real estate
    # Professional / Office
    "69201",  # Accounting and auditing
    "70100",  # Activities of head offices
    "70229",  # Other management consultancy
    "74990",  # Other professional/scientific activities
    "82990",  # Other business support
    # Transport
    "49100",  # Passenger rail transport
    "49320",  # Taxi operation
    "49390",  # Other passenger land transport
    "49410",  # Freight transport by road
    "52210",  # Service activities incidental to land transport
    # Recovery / breakdown (not fast-fit)
    "52219",  # Other service activities incidental to land transport
    # Construction
    "41100",  # Development of building projects
    "41202",  # Construction of domestic buildings
    "43999",  # Other specialised construction
    # Cleaning
    "81210",  # General cleaning of buildings
    "81300",  # Landscape service activities
    "96010",  # Washing and dry cleaning
    # IT
    "62012",  # Business/domestic software development
    "62020",  # IT consultancy
    "62090",  # Other IT service activities
    # Manufacturing (non-automotive)
    "25620",  # Machining
    "28110",  # Manufacture of engines
    "29100",  # Manufacture of motor vehicles
    "29320",  # Manufacture of other parts for motor vehicles
    # Food
    "56101",  # Licenced restaurants
    "56103",  # Take away food shops
    # Finance
    "64209",  # Activities of other holding companies
    "66300",  # Fund management activities
    # Education
    "85530",  # Driving school activities
    # Dormant
    "98000",  # Residents property management
    "99999",  # Dormant company / no SIC code
}


# ═══════════════════════════════════════════════════════════════════════════
#  NAME KEYWORDS
# ═══════════════════════════════════════════════════════════════════════════

# Keywords that confirm this IS a fast-fit operation
FASTFIT_POSITIVE = [
    # Brand / concept names
    "fast fit", "fastfit", "kwik fit", "kwikfit", "quick fit", "quickfit",
    "autocentre", "auto centre", "autocenter", "auto center",
    "pit stop", "pitstop",
    "halfords", "ats euromaster", "formula one autocentres",
    "mr clutch", "national tyres",
    # Service signals
    "exhaust", "brake", "brakes", "clutch", "muffler", "silencer",
    "oil change", "lube", "mot centre", "mot center",
    "service centre", "service center",
    "fitting centre", "fitting center",
    "auto repair", "car servicing",
    "tyre and exhaust", "tyre and brake", "tyres and exhaust", "tyres and brake",
    "tyre & exhaust", "tyre & brake", "tyres & exhaust", "tyres & brake",
]

# Keywords that suggest this is NOT fast-fit
FASTFIT_NOISE = [
    # Body work — excluded from fast-fit definition
    "body shop", "bodywork", "panel beat", "respray", "paintwork",
    "coachwork", "spray", "refinish", "accident repair",
    # Dealerships
    "car sales", "used cars", "forecourt", "dealership", "dealer",
    "motor trade", "car supermarket", "prestige",
    # Car wash / valeting
    "car wash", "valeting", "detailing", "hand wash", "polish",
    # Recovery / breakdown
    "recovery", "breakdown", "rescue", "towing", "roadside",
    # Scrap / salvage
    "scrap", "salvage", "dismantler", "breaker", "recycl",
    # Rental / hire
    "rental", "hire", "leasing", "rent a car",
    # Tuning / performance — niche, not fast-fit
    "tuning", "performance", "custom", "motorsport", "racing", "rally",
    # Driving school
    "driving school", "driving instructor",
    # Insurance
    "insurance", "claims",
    # Parking
    "parking", "car park",
    # Mobile mechanic — not a fast-fit centre
    "mobile mechanic", "mobile repair",
    # Tyre-only signals (tyre specialist, not fast-fit)
    "tyre centre", "tyre center", "tyre specialist", "tyre depot",
    "tyre warehouse", "tyre wholesale",
    # Manufacturing
    "manufacture", "engineering", "fabricat",
    # Wholesale-only
    "wholesale", "distribut", "parts supply", "trade only",
]


# ═══════════════════════════════════════════════════════════════════════════
#  CLASSIFICATION
# ═══════════════════════════════════════════════════════════════════════════

def classify(company):
    """
    Classify a fast-fit company as KEEP / REVIEW / NOISE.
    Returns (classification, score, reasons[])
    """
    name = company.get("company_name", "").lower()
    sic_raw = company.get("sic_codes", "")
    sic_codes = set(s.strip() for s in sic_raw.split(",") if s.strip())
    status = company.get("status", "").lower()
    officer_count = int(company.get("officer_count", 0) or 0)

    score = 0
    reasons = []

    # ── Status ──────────────────────────────────────────────────────────
    if status != "active":
        score -= 2
        reasons.append(f"-2:status={status}")

    # ── SIC scoring ─────────────────────────────────────────────────────
    green = sic_codes & SIC_GREEN
    amber = sic_codes & SIC_AMBER
    red = sic_codes & SIC_RED

    if green:
        score += 3
        reasons.append(f"+3:green-sic({','.join(green)})")
    if amber:
        score += 1
        reasons.append(f"+1:amber-sic({','.join(amber)})")
    if red:
        pen = len(red) * -2
        score += pen
        reasons.append(f"{pen}:red-sic({','.join(red)})")
    if sic_codes and not green and not amber:
        score -= 2
        reasons.append("-2:no-relevant-sic")

    # Tyre-only SIC (45320 only, no 45200)
    if sic_codes == {"45320"}:
        score -= 1
        reasons.append("-1:tyre-only-sic(45320)")

    # ── Name: positive keywords ─────────────────────────────────────────
    pos_hits = [kw for kw in FASTFIT_POSITIVE if kw in name]
    if pos_hits:
        pts = min(len(pos_hits) * 2, 8)
        score += pts
        reasons.append(f"+{pts}:fast-fit-keywords({','.join(pos_hits[:3])})")

    # ── Name: noise keywords ───────────────────────────────────────────
    noise_hits = [kw for kw in FASTFIT_NOISE if kw in name]
    if noise_hits:
        pen = len(noise_hits) * -3
        score += pen
        reasons.append(f"{pen}:noise-keywords({','.join(noise_hits[:3])})")

    # ── Structural: single officer = likely independent ─────────────────
    if officer_count <= 1:
        score -= 1
        reasons.append("-1:single-officer(likely-independent)")

    # ── Generic garage name with no fast-fit signal ─────────────────────
    generic_terms = ["garage", "motors", "autos", "motor services",
                     "auto services", "vehicle services"]
    is_generic = any(g in name for g in generic_terms)
    has_fastfit_signal = bool(pos_hits)
    if is_generic and not has_fastfit_signal:
        score -= 1
        reasons.append("-1:generic-garage-no-fastfit-signal")

    # ── Classify ────────────────────────────────────────────────────────
    if score >= 4:
        classification = "KEEP"
    elif score >= 0:
        classification = "REVIEW"
    else:
        classification = "NOISE"

    return classification, score, reasons


# ═══════════════════════════════════════════════════════════════════════════
#  REPORTING
# ═══════════════════════════════════════════════════════════════════════════

def print_report(classified):
    keep = [c for c in classified if c["classification"] == "KEEP"]
    review = [c for c in classified if c["classification"] == "REVIEW"]
    noise = [c for c in classified if c["classification"] == "NOISE"]
    total = len(classified)

    print("\n" + "=" * 80)
    print("  FAST FIT — NOISE DETECTION REPORT")
    print("=" * 80)

    print(f"""
  ┌──────────────────────────────────────────────┐
  │  Total unique companies:     {total:>5}            │
  │                                              │
  │  KEEP  (score >= 4):         {len(keep):>5}  ({len(keep)/total*100:.1f}%)   │
  │  REVIEW (score 0-3):         {len(review):>5}  ({len(review)/total*100:.1f}%)   │
  │  NOISE (score < 0):          {len(noise):>5}  ({len(noise)/total*100:.1f}%)   │
  └──────────────────────────────────────────────┘
""")

    # ── NOISE detail ────────────────────────────────────────────────────
    print(f"{'─' * 70}")
    print(f"  NOISE ({len(noise)}) — Why they don't belong")
    print(f"{'─' * 70}")

    noise_groups = defaultdict(list)
    for c in noise:
        reasons = c["reasons"]
        if any("noise-keywords" in r for r in reasons):
            group = "Name contains non-fast-fit keywords"
        elif any("red-sic" in r or "no-relevant-sic" in r for r in reasons):
            group = "Wrong SIC codes"
        elif any("single-officer" in r for r in reasons):
            group = "Likely independent (single officer)"
        elif any("generic-garage" in r for r in reasons):
            group = "Generic garage, no fast-fit signal"
        elif any("status=" in r for r in reasons):
            group = "Not active"
        else:
            group = "Other / combined reasons"
        noise_groups[group].append(c)

    for group, items in sorted(noise_groups.items(), key=lambda x: -len(x[1])):
        print(f"\n  [{group}] — {len(items)} companies:")
        for c in items[:10]:
            print(f"    NOISE  {c['company_name']:<40} SIC: {c['sic_codes']:<16} "
                  f"Score: {c['score']}")
            print(f"           {' | '.join(c['reasons'])}")
        if len(items) > 10:
            print(f"    ... and {len(items) - 10} more")

    # ── REVIEW detail ───────────────────────────────────────────────────
    print(f"\n{'─' * 70}")
    print(f"  REVIEW ({len(review)}) — Need manual check")
    print(f"{'─' * 70}")
    for c in sorted(review, key=lambda x: x["score"]):
        print(f"    REVIEW {c['company_name']:<40} SIC: {c['sic_codes']:<16} "
              f"Score: {c['score']}")
        print(f"           {' | '.join(c['reasons'])}")

    # ── KEEP summary ────────────────────────────────────────────────────
    print(f"\n{'─' * 70}")
    print(f"  KEEP ({len(keep)}) — Confirmed fast-fit")
    print(f"{'─' * 70}")
    for c in sorted(keep, key=lambda x: -x["score"])[:30]:
        print(f"    KEEP   {c['company_name']:<40} SIC: {c['sic_codes']:<16} "
              f"Score: {c['score']}")

    # ── What SIC codes are causing noise ────────────────────────────────
    print(f"\n{'─' * 70}")
    print(f"  SIC CODES CAUSING THE MOST NOISE")
    print(f"{'─' * 70}")
    noise_sics = Counter()
    for c in noise:
        for s in c.get("sic_codes", "").split(","):
            s = s.strip()
            if s:
                noise_sics[s] += 1

    for sic, count in noise_sics.most_common(15):
        tag = "RED" if sic in SIC_RED else ("AMBER" if sic in SIC_AMBER else "GREEN")
        print(f"    {sic} [{tag}]: appeared in {count} noise companies")


def export_excel(classified, output_path):
    keep = [c for c in classified if c["classification"] == "KEEP"]
    review = [c for c in classified if c["classification"] == "REVIEW"]
    noise = [c for c in classified if c["classification"] == "NOISE"]

    def to_rows(items):
        return [{
            "Classification": c["classification"],
            "Score": c["score"],
            "Reasons": " | ".join(c["reasons"]),
            "Company Name": c.get("company_name", ""),
            "Company Number": c.get("company_number", ""),
            "Status": c.get("status", ""),
            "SIC Codes": c.get("sic_codes", ""),
            "Company Type": c.get("company_type", ""),
            "Address": c.get("registered_address", ""),
            "Postcode": c.get("postal_code", ""),
            "Locality": c.get("locality", ""),
            "Region": c.get("region", ""),
            "Date Created": c.get("date_created", ""),
            "Officer Count": c.get("officer_count", ""),
            "Original Score": c.get("net_score", ""),
            "Original Flag": c.get("flag", ""),
        } for c in items]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(to_rows(keep)).to_excel(
            writer, sheet_name="Clean (KEEP)", index=False)
        pd.DataFrame(to_rows(review)).to_excel(
            writer, sheet_name="Review", index=False)
        pd.DataFrame(to_rows(noise)).to_excel(
            writer, sheet_name="Noise", index=False)
        all_sorted = sorted(classified, key=lambda x: -x["score"])
        pd.DataFrame(to_rows(all_sorted)).to_excel(
            writer, sheet_name="All (by score)", index=False)

        # Summary sheet
        total = len(classified)
        noise_sics = Counter()
        for c in noise:
            for s in c.get("sic_codes", "").split(","):
                s = s.strip()
                if s:
                    noise_sics[s] += 1

        summary = [
            {"Metric": "Total unique companies", "Value": total},
            {"Metric": "KEEP (confirmed fast-fit)", "Value": len(keep)},
            {"Metric": "REVIEW (manual check)", "Value": len(review)},
            {"Metric": "NOISE (remove)", "Value": len(noise)},
            {"Metric": "KEEP %", "Value": f"{len(keep)/total*100:.1f}%"},
            {"Metric": "NOISE %", "Value": f"{len(noise)/total*100:.1f}%"},
            {"Metric": "", "Value": ""},
            {"Metric": "Top noise SIC codes:", "Value": ""},
        ]
        for sic, count in noise_sics.most_common(10):
            summary.append({"Metric": f"  SIC {sic}", "Value": count})
        pd.DataFrame(summary).to_excel(
            writer, sheet_name="Summary", index=False)


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) > 1:
        csv_path = sys.argv[1]
    else:
        csv_path = Path(__file__).resolve().parent / "fast_fit_companies.csv"

    if not Path(csv_path).exists():
        print(f"Error: File not found: {csv_path}")
        print("Run fast_fit_lookup.py first to generate the data.")
        sys.exit(1)

    # Load CSV
    print(f"Loading: {csv_path}")
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        companies = list(reader)
    print(f"Loaded {len(companies)} entries")

    # Deduplicate by company number
    seen = set()
    deduped = []
    dupes = 0
    for c in companies:
        cn = c.get("company_number", "")
        if cn and cn in seen:
            dupes += 1
            continue
        if cn:
            seen.add(cn)
        deduped.append(c)
    companies = deduped
    print(f"After dedup: {len(companies)} unique ({dupes} duplicates removed)")

    # Classify
    classified = []
    for c in companies:
        classification, score, reasons = classify(c)
        entry = dict(c)
        entry["classification"] = classification
        entry["score"] = score
        entry["reasons"] = reasons
        classified.append(entry)

    # Report
    print_report(classified)

    # Excel
    output_path = Path(__file__).resolve().parent / "fast_fit_noise_report.xlsx"
    export_excel(classified, output_path)
    print(f"\nExcel report saved: {output_path}")
    print(f"  'Clean (KEEP)'   — confirmed fast-fit companies")
    print(f"  'Review'         — needs manual check")
    print(f"  'Noise'          — not fast-fit, remove")
    print(f"  'All (by score)' — everything ranked")
    print(f"  'Summary'        — overview stats")


if __name__ == "__main__":
    main()
