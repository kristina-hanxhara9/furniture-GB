#!/usr/bin/env python3
"""
Fast Fit Car Specialists — Noise Detection Script (v2 — strict gating)

Reads fast_fit_companies.csv from fast_fit_lookup.py, deduplicates by
company number, then applies a two-stage filter:

  Stage 1 — GATE: Company must contain a Tier 1 (brand/explicit fast-fit)
            or Tier 2 (multi-service combo) keyword in its name.
            No signal → GATED OUT, never scored, never enriched.

  Stage 2 — SCORE: Weighted keyword scoring + SIC check + structural
            signals.  KEEP >= 4, REVIEW >= 2, below 2 → DISCARD.

Definition:
  Fast Fit = >50% turnover from quick automotive repairs (exhaust, brakes,
  oil, filters). NOT full-service garages, NOT body shops, NOT tyre-only,
  NOT dealerships, NOT independents. Only organized multi-service retailers.

Outputs:
  - Console report with gate stats + scored breakdown
  - fast_fit_noise_report.xlsx (Clean / Review / Gated Out / Discarded / Summary)

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
#  TIER 1 & TIER 2 — GATE KEYWORDS
# ═══════════════════════════════════════════════════════════════════════════
#
#  A company MUST match at least one keyword here to even be scored.
#  This is the single biggest noise reducer.

# Tier 1: Known brands / explicit "fast fit" language — instant qualify
TIER_1 = [
    "fast fit", "fastfit", "kwik fit", "kwikfit", "quick fit", "quickfit",
    "autocentre", "auto centre", "autocenter", "auto center",
    "pit stop", "pitstop",
    "midas", "speedy", "norauto",
    "halfords autocentre", "national autocentre", "national exhausts",
    "formula one autocentres", "ats euromaster",
]

# Tier 2: Multi-service combos — strong signal of fast-fit operation
TIER_2 = [
    "exhausts and brakes", "brakes and exhausts",
    "tyres exhausts brakes", "tyres brakes exhausts",
    "tyre exhaust", "tyres exhaust",
    "tyre brake", "tyres brake",
    "exhaust brake", "exhausts brakes",
    "tyre fitting", "tyres fitting",
    "brake fitting", "exhausts fitting",
    "oil and tyres", "lube centre", "lube center",
    "service and mot", "mot and service",
    "tyre and exhaust centre", "exhaust centre", "exhaust center",
    "brake centre", "brake center",
    "tyre and exhaust", "tyre and brake",
    "tyres and exhaust", "tyres and brake",
    "tyre & exhaust", "tyre & brake",
    "tyres & exhaust", "tyres & brake",
    "tyres & exhausts", "exhaust & brake",
    "exhausts & brakes",
    "mr clutch",
]


def has_fast_fit_signal(name):
    """
    Gate check: does this company name contain a Tier 1 or Tier 2 keyword?
    Returns (passed, tier_label) — e.g. (True, "tier1") or (False, None).
    """
    name_lower = name.lower()
    for kw in TIER_1:
        if kw in name_lower:
            return True, "tier1"
    for kw in TIER_2:
        if kw in name_lower:
            return True, "tier2"
    return False, None


# ═══════════════════════════════════════════════════════════════════════════
#  SIC CODE CLASSIFICATION
# ═══════════════════════════════════════════════════════════════════════════

SIC_GREEN = {
    "45200",  # Maintenance and repair of motor vehicles — PRIMARY
}

SIC_AMBER = {
    "45320",  # Retail trade of motor vehicle parts/accessories
    "45319",  # Wholesale of motor vehicle parts (other)
    "45310",  # Wholesale of motor vehicle parts
    "45400",  # Sale/maintenance of motorcycles (edge case)
}

SIC_RED = {
    "45111", "45112", "45190",        # Car sales / dealerships
    "47190", "47520", "47910", "47990",  # Non-automotive retail
    "68100", "68209", "68320",        # Real estate
    "69201", "70100", "70229", "74990", "82990",  # Professional / Office
    "49100", "49320", "49390", "49410", "52210", "52219",  # Transport
    "41100", "41202", "43999",        # Construction
    "81210", "81300", "96010",        # Cleaning
    "62012", "62020", "62090",        # IT
    "25620", "28110", "29100", "29320",  # Manufacturing
    "56101", "56103",                 # Food
    "64209", "66300",                 # Finance
    "85530",                          # Education
    "98000", "99999",                 # Dormant
}


# ═══════════════════════════════════════════════════════════════════════════
#  SCORING KEYWORDS (weighted)
# ═══════════════════════════════════════════════════════════════════════════

POSITIVE_KEYWORDS = {
    # Tier 1 brands — reward heavily (they already gated in)
    "fast fit": 5, "fastfit": 5, "kwik fit": 5, "kwikfit": 5,
    "quick fit": 5, "quickfit": 5,
    "autocentre": 5, "auto centre": 5, "autocenter": 5, "auto center": 5,
    "pit stop": 5, "pitstop": 5,
    "midas": 5, "speedy": 5,
    # Tier 2 combos — high confidence
    "tyre exhaust": 4, "tyres exhaust": 4,
    "tyre brake": 4, "tyres brake": 4,
    "exhaust brake": 4, "exhausts brakes": 4,
    "tyre and exhaust": 4, "tyres and exhaust": 4,
    "tyre & exhaust": 4, "tyres & exhaust": 4,
    "tyre and brake": 4, "tyres and brake": 4,
    "tyre & brake": 4, "tyres & brake": 4,
    "tyres & exhausts": 4, "exhaust & brake": 4, "exhausts & brakes": 4,
    "lube centre": 4, "lube center": 4,
    "exhaust centre": 3, "exhaust center": 3,
    "brake centre": 3, "brake center": 3,
    "tyre fitting": 3, "tyres fitting": 3,
    "brake fitting": 3, "exhausts fitting": 3,
    "mr clutch": 4,
    # Supplementary boosters (add points but can't gate alone)
    "mot": 1, "servicing": 1,
}

PENALTY_KEYWORDS = {
    # Explicitly excluded by fast-fit definition
    "tyre specialist": -5, "tyre specialists": -5,
    "national tyres": -5,
    # Body work — not fast-fit
    "body shop": -5, "bodywork": -5, "panel beat": -5,
    "respray": -5, "coachwork": -5, "accident repair": -5,
    "spray": -3, "refinish": -3, "paintwork": -3,
    # Independent signals
    "independent": -3, "mobile": -3, "mobile mechanic": -5,
    "mobile repair": -5,
    # Not a repairer
    "accessories": -2, "wholesale": -3, "parts supply": -3,
    "parts": -2, "trade only": -3,
    # Dealerships
    "car sales": -5, "used cars": -5, "dealership": -5, "dealer": -3,
    "forecourt": -3, "prestige": -3, "car supermarket": -5,
    "motor trade": -3,
    # Car wash / valeting
    "car wash": -3, "valeting": -3, "detailing": -3, "hand wash": -3,
    # Recovery / scrap
    "recovery": -3, "breakdown": -3, "towing": -3,
    "scrap": -5, "salvage": -5, "dismantler": -5, "breaker": -5,
    # Rental / hire
    "rental": -3, "hire": -2, "leasing": -3, "rent a car": -5,
    # Tuning / performance
    "tuning": -3, "performance": -2, "custom": -2,
    "motorsport": -5, "racing": -5,
    # Manufacturing / engineering
    "manufacture": -3, "engineering": -2, "fabricat": -3,
    # Distribution
    "distribut": -3,
    # Parking
    "parking": -3, "car park": -3,
    # Insurance
    "insurance": -3, "claims": -2,
    # Tyre-only outlets
    "tyre centre": -3, "tyre center": -3, "tyre depot": -3,
    "tyre warehouse": -5, "tyre wholesale": -5,
}

# Thresholds
SCORE_KEEP = 4
SCORE_REVIEW = 2


# ═══════════════════════════════════════════════════════════════════════════
#  CLASSIFICATION
# ═══════════════════════════════════════════════════════════════════════════

def classify(company):
    """
    Two-stage classification:
      1. Gate: must have Tier 1 or Tier 2 keyword
      2. Score: weighted keywords + SIC + structural

    Returns (classification, score, reasons[], tier)
    """
    name = company.get("company_name", "").lower()

    # ── Stage 1: GATE ──────────────────────────────────────────────────
    passed_gate, tier = has_fast_fit_signal(name)
    if not passed_gate:
        return "GATED_OUT", 0, ["no-tier1-or-tier2-keyword"], None

    sic_raw = company.get("sic_codes", "")
    sic_codes = set(s.strip() for s in sic_raw.split(",") if s.strip())
    status = company.get("status", "").lower()
    officer_count = int(company.get("officer_count", 0) or 0)

    score = 0
    reasons = [f"gate:{tier}"]

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

    # SIC combination gate: 45320-only needs Tier 1 to pass
    if sic_codes == {"45320"} and tier != "tier1":
        score -= 3
        reasons.append("-3:tyre-only-sic(45320)+no-tier1-brand")

    # ── Name: positive keywords (weighted) ─────────────────────────────
    pos_total = 0
    pos_hits = []
    for kw, pts in POSITIVE_KEYWORDS.items():
        if kw in name:
            pos_total += pts
            pos_hits.append(f"{kw}(+{pts})")

    if pos_total > 0:
        score += pos_total
        reasons.append(f"+{pos_total}:positive({','.join(pos_hits[:4])})")

    # ── Name: penalty keywords (weighted) ──────────────────────────────
    neg_total = 0
    neg_hits = []
    for kw, pen in PENALTY_KEYWORDS.items():
        if kw in name:
            neg_total += pen
            neg_hits.append(f"{kw}({pen})")

    if neg_total < 0:
        score += neg_total
        reasons.append(f"{neg_total}:penalty({','.join(neg_hits[:4])})")

    # ── Structural: single officer = likely independent ─────────────────
    if officer_count <= 1:
        score -= 1
        reasons.append("-1:single-officer(likely-independent)")

    # ── Generic garage name with no strong fast-fit signal ──────────────
    generic_terms = ["garage", "motors", "autos", "motor services",
                     "auto services", "vehicle services"]
    is_generic = any(g in name for g in generic_terms)
    has_brand = tier == "tier1"
    if is_generic and not has_brand:
        score -= 1
        reasons.append("-1:generic-garage-no-brand")

    # ── Classify ────────────────────────────────────────────────────────
    if score >= SCORE_KEEP:
        classification = "KEEP"
    elif score >= SCORE_REVIEW:
        classification = "REVIEW"
    else:
        classification = "DISCARD"

    return classification, score, reasons, tier


# ═══════════════════════════════════════════════════════════════════════════
#  REPORTING
# ═══════════════════════════════════════════════════════════════════════════

def print_report(all_companies, classified, gated_out):
    keep = [c for c in classified if c["classification"] == "KEEP"]
    review = [c for c in classified if c["classification"] == "REVIEW"]
    discard = [c for c in classified if c["classification"] == "DISCARD"]
    total_input = len(all_companies)
    total_gated = len(gated_out)
    total_scored = len(classified)

    print("\n" + "=" * 80)
    print("  FAST FIT — NOISE DETECTION REPORT (v2 strict)")
    print("=" * 80)

    pct = lambda n, d: f"{n/d*100:.1f}%" if d else "0%"

    print(f"""
  ┌────────────────────────────────────────────────────────────┐
  │  Total input (after dedup):          {total_input:>5}                  │
  │                                                            │
  │  GATED OUT (no Tier 1/2 keyword):    {total_gated:>5}  ({pct(total_gated, total_input)})   │
  │  Passed gate → scored:               {total_scored:>5}  ({pct(total_scored, total_input)})   │
  │                                                            │
  │  Of scored:                                                │
  │    KEEP     (score >= {SCORE_KEEP}):             {len(keep):>5}  ({pct(len(keep), total_input)})   │
  │    REVIEW   (score {SCORE_REVIEW}-{SCORE_KEEP - 1}):             {len(review):>5}  ({pct(len(review), total_input)})   │
  │    DISCARD  (score < {SCORE_REVIEW}):             {len(discard):>5}  ({pct(len(discard), total_input)})   │
  └────────────────────────────────────────────────────────────┘
""")

    # ── KEEP detail ─────────────────────────────────────────────────────
    print(f"{'─' * 80}")
    print(f"  KEEP ({len(keep)}) — Confirmed fast-fit")
    print(f"{'─' * 80}")
    for c in sorted(keep, key=lambda x: -x["score"]):
        print(f"    KEEP   {c['company_name']:<45} SIC: {c['sic_codes']:<16} "
              f"Score: {c['score']}")
        print(f"           {' | '.join(c['reasons'])}")

    # ── REVIEW detail ───────────────────────────────────────────────────
    if review:
        print(f"\n{'─' * 80}")
        print(f"  REVIEW ({len(review)}) — Need manual check")
        print(f"{'─' * 80}")
        for c in sorted(review, key=lambda x: x["score"]):
            print(f"    REVIEW {c['company_name']:<45} SIC: {c['sic_codes']:<16} "
                  f"Score: {c['score']}")
            print(f"           {' | '.join(c['reasons'])}")

    # ── DISCARD detail ──────────────────────────────────────────────────
    if discard:
        print(f"\n{'─' * 80}")
        print(f"  DISCARD ({len(discard)}) — Passed gate but scored too low")
        print(f"{'─' * 80}")
        for c in sorted(discard, key=lambda x: x["score"])[:20]:
            print(f"    DISCARD {c['company_name']:<44} SIC: {c['sic_codes']:<16} "
                  f"Score: {c['score']}")
            print(f"           {' | '.join(c['reasons'])}")
        if len(discard) > 20:
            print(f"    ... and {len(discard) - 20} more")

    # ── GATED OUT sample ───────────────────────────────────────────────
    print(f"\n{'─' * 80}")
    print(f"  GATED OUT ({total_gated}) — No Tier 1/2 keyword, never scored")
    print(f"{'─' * 80}")
    print(f"  Showing first 20:")
    for c in gated_out[:20]:
        name = c.get("company_name", "")
        sic = c.get("sic_codes", "")
        print(f"    GATED  {name:<45} SIC: {sic}")
    if total_gated > 20:
        print(f"    ... and {total_gated - 20} more")

    # ── SIC codes in KEEP vs everything else ───────────────────────────
    print(f"\n{'─' * 80}")
    print(f"  SIC CODE DISTRIBUTION")
    print(f"{'─' * 80}")

    keep_sics = Counter()
    for c in keep:
        for s in c.get("sic_codes", "").split(","):
            s = s.strip()
            if s:
                keep_sics[s] += 1

    gated_sics = Counter()
    for c in gated_out:
        for s in c.get("sic_codes", "").split(","):
            s = s.strip()
            if s:
                gated_sics[s] += 1

    print(f"\n  In KEEP companies:")
    for sic, count in keep_sics.most_common(10):
        print(f"    {sic}: {count}")

    print(f"\n  In GATED OUT companies (top 10):")
    for sic, count in gated_sics.most_common(10):
        tag = "RED" if sic in SIC_RED else ("AMBER" if sic in SIC_AMBER else
              "GREEN" if sic in SIC_GREEN else "???")
        print(f"    {sic} [{tag}]: {count}")


def export_excel(all_companies, classified, gated_out, output_path):
    keep = [c for c in classified if c["classification"] == "KEEP"]
    review = [c for c in classified if c["classification"] == "REVIEW"]
    discard = [c for c in classified if c["classification"] == "DISCARD"]

    def to_rows(items, include_reasons=True):
        rows = []
        for c in items:
            row = {
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
            }
            if include_reasons:
                row["Classification"] = c.get("classification", "")
                row["Score"] = c.get("score", "")
                row["Tier"] = c.get("tier", "")
                row["Reasons"] = " | ".join(c.get("reasons", []))
            return rows.append(row) or rows  # append and continue
        return rows

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Clean list
        pd.DataFrame(to_rows(keep)).to_excel(
            writer, sheet_name="Clean (KEEP)", index=False)

        # Review
        if review:
            pd.DataFrame(to_rows(review)).to_excel(
                writer, sheet_name="Review", index=False)

        # Gated out (no reasons needed — they all failed the same gate)
        pd.DataFrame(to_rows(gated_out, include_reasons=False)).to_excel(
            writer, sheet_name="Gated Out", index=False)

        # Discarded (passed gate but scored too low)
        if discard:
            pd.DataFrame(to_rows(discard)).to_excel(
                writer, sheet_name="Discarded", index=False)

        # Summary
        total = len(all_companies)
        summary = [
            {"Metric": "Total unique companies", "Value": total},
            {"Metric": "", "Value": ""},
            {"Metric": "GATED OUT (no Tier 1/2 keyword)", "Value": len(gated_out)},
            {"Metric": "Passed gate", "Value": len(classified)},
            {"Metric": "", "Value": ""},
            {"Metric": f"KEEP (score >= {SCORE_KEEP})", "Value": len(keep)},
            {"Metric": f"REVIEW (score {SCORE_REVIEW}-{SCORE_KEEP-1})", "Value": len(review)},
            {"Metric": f"DISCARD (score < {SCORE_REVIEW})", "Value": len(discard)},
            {"Metric": "", "Value": ""},
            {"Metric": "KEEP % of total", "Value": f"{len(keep)/total*100:.1f}%"},
            {"Metric": "Noise removed %", "Value": f"{(len(gated_out)+len(discard))/total*100:.1f}%"},
        ]
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

    # Stage 1: Gate
    gated_out = []
    passed_gate = []
    for c in companies:
        name = c.get("company_name", "")
        passed, tier = has_fast_fit_signal(name)
        if passed:
            c["tier"] = tier
            passed_gate.append(c)
        else:
            gated_out.append(c)

    print(f"\nGate results:")
    print(f"  Passed (Tier 1/2 keyword found): {len(passed_gate)}")
    print(f"  Gated out (no signal):           {len(gated_out)}")

    # Stage 2: Score those that passed the gate
    classified = []
    for c in passed_gate:
        classification, score, reasons, tier = classify(c)
        entry = dict(c)
        entry["classification"] = classification
        entry["score"] = score
        entry["reasons"] = reasons
        entry["tier"] = tier or c.get("tier", "")
        classified.append(entry)

    # Report
    print_report(companies, classified, gated_out)

    # Excel
    output_path = Path(__file__).resolve().parent / "fast_fit_noise_report.xlsx"
    export_excel(companies, classified, gated_out, output_path)
    print(f"\nExcel report saved: {output_path}")
    print(f"  'Clean (KEEP)'  — confirmed fast-fit companies")
    print(f"  'Review'        — borderline, needs manual check")
    print(f"  'Gated Out'     — no Tier 1/2 keyword, removed pre-scoring")
    print(f"  'Discarded'     — passed gate but scored too low")
    print(f"  'Summary'       — overview stats")


if __name__ == "__main__":
    main()
