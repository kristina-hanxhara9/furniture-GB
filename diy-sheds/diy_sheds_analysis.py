#!/usr/bin/env python3
"""
DIY Sheds — Companies House Analysis Script

Reads the enriched JSON from diy_sheds_lookup.py and performs deep analysis:
  - SIC code frequency and distribution
  - Keyword frequency in company names
  - SIC code + keyword combinations
  - Co-occurrence patterns (which SIC codes appear together)
  - Status breakdown (active, dissolved, liquidation)
  - Chain detection
  - Regional analysis
  - Financial health indicators
  - Timeline analysis

Usage:
    python diy_sheds_analysis.py                          # uses default diy_sheds_enriched.json
    python diy_sheds_analysis.py custom_enriched.json     # use custom file
"""

import json
import re
import sys
from collections import Counter, defaultdict
from itertools import combinations
from pathlib import Path

import pandas as pd


# ═══════════════════════════════════════════════════════════════════════════
#  SIC CODE REFERENCE
# ═══════════════════════════════════════════════════════════════════════════

SIC_DESCRIPTIONS = {
    # Retail - DIY / Hardware
    "47521": "Retail sale of hardware, paints and glass in specialised stores",
    "47523": "Retail sale of hardware, paints and glass in specialised stores",
    "47529": "Retail sale of hardware, paints and glass in specialised stores (other)",
    "47599": "Retail sale of furniture, lighting and household articles n.e.c.",
    "47591": "Retail sale of furniture and lighting equipment",
    "47190": "Other retail sale in non-specialised stores",
    "47910": "Retail sale via mail order houses or via Internet",
    "47990": "Other retail sale not in stores, stalls or markets",
    # Manufacturing - Wood / Sheds
    "16100": "Sawmilling and planing of wood",
    "16210": "Manufacture of veneer sheets and wood-based panels",
    "16230": "Manufacture of other builders' carpentry and joinery",
    "16240": "Manufacture of wooden containers",
    "16290": "Manufacture of other products of wood; articles of cork, straw",
    "25110": "Manufacture of metal structures and parts of structures",
    "25120": "Manufacture of doors and windows of metal",
    # Construction
    "41100": "Development of building projects",
    "41201": "Construction of commercial buildings",
    "41202": "Construction of domestic buildings",
    "43320": "Joinery installation",
    "43290": "Other construction installation",
    "43390": "Other building completion and finishing",
    "43999": "Other specialised construction activities n.e.c.",
    # Wholesale
    "46130": "Agents involved in the sale of timber and building materials",
    "46730": "Wholesale of wood, construction materials and sanitary equipment",
    "46499": "Wholesale of other household goods",
    "46900": "Non-specialised wholesale trade",
    # Other relevant
    "68100": "Buying and selling of own real estate",
    "68209": "Other letting and operating of own or leased real estate",
    "70100": "Activities of head offices",
    "77390": "Renting and leasing of other machinery and equipment n.e.c.",
    "82990": "Other business support service activities n.e.c.",
    "01300": "Plant propagation (garden centres)",
    "01190": "Growing of other non-perennial crops",
    "47761": "Retail sale of flowers, plants, seeds and fertilisers",
    "47762": "Retail sale of pet animals and pet food",
}


# ═══════════════════════════════════════════════════════════════════════════
#  DIY / SHED KEYWORD CATEGORIES
# ═══════════════════════════════════════════════════════════════════════════

KEYWORD_CATEGORIES = {
    "diy": ["diy", "do it yourself", "do-it-yourself"],
    "shed": ["shed", "sheds", "garden building", "garden buildings",
             "summerhouse", "summer house", "log cabin", "outhouse"],
    "garden": ["garden", "gardens", "gardening", "garden centre", "garden center"],
    "building_supplies": ["building supplies", "building materials", "builders merchant",
                          "builders merchants", "building merchant"],
    "timber": ["timber", "wood", "wooden", "lumber", "fencing", "fence"],
    "hardware": ["hardware", "ironmongery", "tools", "fixings", "fasteners"],
    "home_improvement": ["home improvement", "home improvements", "homebase",
                         "wickes", "b&q", "screwfix"],
    "paint": ["paint", "paints", "painting", "decorator", "decorating"],
    "plumbing": ["plumbing", "plumber", "bathroom", "bathrooms", "heating"],
    "electrical": ["electrical", "electrician", "lighting", "lights"],
    "flooring": ["flooring", "floor", "carpet", "tile", "tiles", "tiling"],
    "roofing": ["roof", "roofing", "guttering"],
    "landscaping": ["landscape", "landscaping", "paving", "decking", "deck"],
    "brand_chain": ["wickes", "b&q", "homebase", "screwfix", "toolstation",
                    "travis perkins", "jewson", "selco", "howdens"],
    "online": ["online", "direct", "warehouse", "depot", "outlet"],
}


def get_sic_description(code):
    """Get SIC code description, with fallback."""
    return SIC_DESCRIPTIONS.get(code, f"SIC {code} (unlisted)")


def extract_keywords_from_name(name):
    """Extract meaningful lowercase keywords from company name."""
    if not name:
        return []
    name = name.upper()
    for remove in ["LTD", "LIMITED", "PLC", "LLP", "INC", "UK", "THE",
                    "AND", "&", "OF", "IN", "AT", "FOR", "CO"]:
        name = re.sub(rf"\b{remove}\b", "", name)
    return [w.lower() for w in re.findall(r"[A-Z]{2,}", name)]


def categorize_name(name):
    """Assign keyword categories to a company name."""
    if not name:
        return []
    name_lower = name.lower()
    categories = []
    for cat, keywords in KEYWORD_CATEGORIES.items():
        for kw in keywords:
            if kw in name_lower:
                categories.append(cat)
                break
    return categories


# ═══════════════════════════════════════════════════════════════════════════
#  ANALYSIS FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════

def analyze_sic_codes(companies):
    """Analyze SIC code frequency and distribution."""
    print("\n" + "=" * 80)
    print("  1. SIC CODE ANALYSIS")
    print("=" * 80)

    all_sic = []
    sic_per_company = []

    for c in companies:
        codes = [s.strip() for s in c.get("sic_codes", "").split(",") if s.strip()]
        all_sic.extend(codes)
        sic_per_company.append(len(codes))

    sic_counts = Counter(all_sic)
    total_companies = len(companies)

    print(f"\n  Total unique SIC codes found: {len(sic_counts)}")
    print(f"  Average SIC codes per company: {sum(sic_per_company)/len(sic_per_company):.1f}" if sic_per_company else "")
    print(f"  Max SIC codes on one company: {max(sic_per_company)}" if sic_per_company else "")

    print(f"\n  {'SIC Code':<10} {'Count':>6} {'%':>7}  Description")
    print(f"  {'─'*10} {'─'*6} {'─'*7}  {'─'*45}")
    for code, count in sic_counts.most_common(30):
        pct = count / total_companies * 100
        desc = get_sic_description(code)
        bar = "█" * int(pct / 2)
        print(f"  {code:<10} {count:>6} {pct:>6.1f}%  {desc}")
        if bar:
            print(f"  {'':10} {'':6} {'':7}  {bar}")

    return sic_counts


def analyze_keywords(companies):
    """Analyze keyword frequency in company names."""
    print("\n" + "=" * 80)
    print("  2. KEYWORD ANALYSIS (Company Names)")
    print("=" * 80)

    all_keywords = []
    for c in companies:
        name = c.get("company_name", "") or c.get("original_name", "")
        keywords = extract_keywords_from_name(name)
        all_keywords.extend(keywords)

    keyword_counts = Counter(all_keywords)

    # Filter out very common but meaningless words
    noise = {"company", "group", "holdings", "services", "enterprises",
             "trading", "international", "management", "solutions", "properties"}
    filtered = {k: v for k, v in keyword_counts.items() if k.lower() not in noise}
    filtered_counts = Counter(filtered)

    print(f"\n  Total unique keywords: {len(keyword_counts)}")
    print(f"  After noise removal:  {len(filtered_counts)}")

    print(f"\n  TOP 40 KEYWORDS:")
    print(f"  {'Keyword':<25} {'Count':>6} {'%':>7}")
    print(f"  {'─'*25} {'─'*6} {'─'*7}")
    total = len(companies)
    for kw, count in filtered_counts.most_common(40):
        pct = count / total * 100
        print(f"  {kw:<25} {count:>6} {pct:>6.1f}%")

    return keyword_counts


def analyze_keyword_categories(companies):
    """Analyze which keyword categories companies fall into."""
    print("\n" + "=" * 80)
    print("  3. KEYWORD CATEGORY ANALYSIS")
    print("=" * 80)

    category_counts = Counter()
    category_companies = defaultdict(list)

    for c in companies:
        name = c.get("company_name", "") or c.get("original_name", "")
        cats = categorize_name(name)
        for cat in cats:
            category_counts[cat] += 1
            category_companies[cat].append(name)

    total = len(companies)
    print(f"\n  {'Category':<25} {'Count':>6} {'%':>7}")
    print(f"  {'─'*25} {'─'*6} {'─'*7}")
    for cat, count in category_counts.most_common():
        pct = count / total * 100
        print(f"  {cat:<25} {count:>6} {pct:>6.1f}%")

    # Show example companies for each category
    print(f"\n  EXAMPLES PER CATEGORY:")
    for cat, count in category_counts.most_common():
        examples = category_companies[cat][:5]
        print(f"\n  [{cat}] ({count} companies):")
        for ex in examples:
            print(f"    - {ex}")

    return category_counts


def analyze_sic_keyword_combinations(companies):
    """Analyze combinations of SIC codes and name keywords."""
    print("\n" + "=" * 80)
    print("  4. SIC CODE + KEYWORD COMBINATIONS")
    print("=" * 80)

    combo_counts = Counter()
    sic_keyword_map = defaultdict(Counter)

    for c in companies:
        name = c.get("company_name", "") or c.get("original_name", "")
        sic_codes = [s.strip() for s in c.get("sic_codes", "").split(",") if s.strip()]
        categories = categorize_name(name)

        for sic in sic_codes:
            for cat in categories:
                combo = f"{sic} + {cat}"
                combo_counts[combo] += 1
                sic_keyword_map[sic][cat] += 1

    print(f"\n  Total unique SIC+Keyword combinations: {len(combo_counts)}")

    print(f"\n  TOP 30 COMBINATIONS:")
    print(f"  {'SIC + Keyword Category':<45} {'Count':>6}")
    print(f"  {'─'*45} {'─'*6}")
    for combo, count in combo_counts.most_common(30):
        sic_part = combo.split(" + ")[0]
        desc = get_sic_description(sic_part)
        print(f"  {combo:<45} {count:>6}")
        print(f"    ({desc})")

    # Cross-tabulation: for each SIC code, which keywords appear
    print(f"\n  SIC CODE → KEYWORD CATEGORY HEATMAP:")
    print(f"  (showing top 10 SIC codes and their keyword associations)\n")

    top_sics = Counter()
    for c in companies:
        for s in [s.strip() for s in c.get("sic_codes", "").split(",") if s.strip()]:
            top_sics[s] += 1

    for sic, _ in top_sics.most_common(10):
        if sic in sic_keyword_map:
            kw_dist = sic_keyword_map[sic]
            desc = get_sic_description(sic)
            print(f"  SIC {sic} ({desc}):")
            for cat, count in kw_dist.most_common(5):
                bar = "█" * count
                print(f"    {cat:<20} {count:>4} {bar}")
            print()

    return combo_counts


def analyze_sic_cooccurrence(companies):
    """Analyze which SIC codes appear together on the same company."""
    print("\n" + "=" * 80)
    print("  5. SIC CODE CO-OCCURRENCE (which codes appear together)")
    print("=" * 80)

    pair_counts = Counter()
    multi_sic_companies = 0

    for c in companies:
        codes = sorted(set(s.strip() for s in c.get("sic_codes", "").split(",") if s.strip()))
        if len(codes) >= 2:
            multi_sic_companies += 1
            for pair in combinations(codes, 2):
                pair_counts[pair] += 1

    total = len(companies)
    print(f"\n  Companies with 2+ SIC codes: {multi_sic_companies} ({multi_sic_companies/total*100:.1f}%)" if total else "")

    print(f"\n  TOP 25 SIC CODE PAIRS:")
    print(f"  {'SIC Pair':<25} {'Count':>6}")
    print(f"  {'─'*25} {'─'*6}")
    for (a, b), count in pair_counts.most_common(25):
        desc_a = get_sic_description(a)[:30]
        desc_b = get_sic_description(b)[:30]
        print(f"  {a} + {b:<14} {count:>6}")
        print(f"    {desc_a}")
        print(f"    {desc_b}")

    return pair_counts


def analyze_status(companies):
    """Analyze company status breakdown."""
    print("\n" + "=" * 80)
    print("  6. COMPANY STATUS BREAKDOWN")
    print("=" * 80)

    found = [c for c in companies if c.get("ch_found")]
    not_found = [c for c in companies if not c.get("ch_found")]
    total = len(companies)

    print(f"\n  Total entries:        {total}")
    print(f"  Found in CH:         {len(found)}")
    print(f"  Not found in CH:     {len(not_found)}")

    status_counts = Counter(c.get("company_status", "unknown") for c in found)
    print(f"\n  {'Status':<30} {'Count':>6} {'%':>7}")
    print(f"  {'─'*30} {'─'*6} {'─'*7}")
    for status, count in status_counts.most_common():
        pct = count / len(found) * 100 if found else 0
        print(f"  {status:<30} {count:>6} {pct:>6.1f}%")

    active = sum(1 for c in found if c.get("company_status") == "active")
    dissolved = sum(1 for c in found if c.get("company_status") == "dissolved")
    other_inactive = len(found) - active

    print(f"\n  ACTIVE:              {active}")
    print(f"  DISSOLVED/CLOSED:    {dissolved}")
    print(f"  OTHER INACTIVE:      {other_inactive - dissolved}")

    # SIC distribution for active vs dissolved
    active_sics = Counter()
    dissolved_sics = Counter()
    for c in found:
        codes = [s.strip() for s in c.get("sic_codes", "").split(",") if s.strip()]
        if c.get("company_status") == "active":
            active_sics.update(codes)
        elif c.get("company_status") == "dissolved":
            dissolved_sics.update(codes)

    if active_sics and dissolved_sics:
        print(f"\n  SIC CODES: ACTIVE vs DISSOLVED COMPANIES")
        all_sics_here = set(list(active_sics.keys()) + list(dissolved_sics.keys()))
        print(f"  {'SIC':<10} {'Active':>8} {'Dissolved':>10}")
        print(f"  {'─'*10} {'─'*8} {'─'*10}")
        for sic in sorted(all_sics_here, key=lambda x: -(active_sics.get(x, 0) + dissolved_sics.get(x, 0))):
            a = active_sics.get(sic, 0)
            d = dissolved_sics.get(sic, 0)
            print(f"  {sic:<10} {a:>8} {d:>10}")

    return status_counts


def analyze_chains(companies):
    """Detect and analyze chain stores."""
    print("\n" + "=" * 80)
    print("  7. CHAIN DETECTION")
    print("=" * 80)

    from collections import defaultdict

    def norm(name):
        if not name:
            return ""
        name = name.strip().upper()
        for suffix in [" LTD", " LIMITED", " PLC", " LLP", " INC", " UK", " & CO"]:
            if name.endswith(suffix):
                name = name[:-len(suffix)].strip()
        return name

    name_groups = defaultdict(list)
    number_groups = defaultdict(list)

    for c in companies:
        cname = norm(c.get("company_name") or c.get("original_name", ""))
        if cname:
            name_groups[cname].append(c)
        cn = c.get("company_number", "")
        if cn:
            number_groups[cn].append(c)

    chains_by_name = {k: v for k, v in name_groups.items() if len(v) >= 2}
    chains_by_number = {k: v for k, v in number_groups.items() if len(v) >= 2}

    chain_entries = sum(len(v) for v in chains_by_name.values())
    total = len(companies)

    print(f"\n  Unique chain names (2+ locations): {len(chains_by_name)}")
    print(f"  Total chain entries:               {chain_entries}")
    print(f"  Independent entries:               {total - chain_entries}")
    if total:
        print(f"  Chain %:                           {chain_entries/total*100:.1f}%")

    if chains_by_name:
        print(f"\n  CHAINS BY NAME (sorted by location count):")
        for name, entries in sorted(chains_by_name.items(), key=lambda x: -len(x[1])):
            statuses = Counter(e.get("company_status", "?") for e in entries)
            status_str = ", ".join(f"{s}:{c}" for s, c in statuses.most_common())
            regions = Counter(e.get("region", "?") or "?" for e in entries)
            region_str = ", ".join(f"{r}" for r, _ in regions.most_common(3))
            print(f"    {name}: {len(entries)} locations ({status_str})")
            print(f"      Regions: {region_str}")

    if chains_by_number:
        print(f"\n  SAME COMPANY NUMBER appearing multiple times:")
        for num, entries in sorted(chains_by_number.items(), key=lambda x: -len(x[1])):
            name = entries[0].get("company_name", "")
            print(f"    {num}: {name} ({len(entries)} entries)")


def analyze_regions(companies):
    """Regional breakdown."""
    print("\n" + "=" * 80)
    print("  8. REGIONAL ANALYSIS")
    print("=" * 80)

    found = [c for c in companies if c.get("ch_found")]
    region_counts = Counter(c.get("region", "Unknown") or "Unknown" for c in found)
    locality_counts = Counter(c.get("locality", "Unknown") or "Unknown" for c in found)

    total = len(found)
    print(f"\n  BY REGION:")
    print(f"  {'Region':<30} {'Count':>6} {'%':>7}")
    print(f"  {'─'*30} {'─'*6} {'─'*7}")
    for region, count in region_counts.most_common():
        pct = count / total * 100 if total else 0
        print(f"  {region:<30} {count:>6} {pct:>6.1f}%")

    print(f"\n  TOP 20 LOCALITIES:")
    print(f"  {'Locality':<30} {'Count':>6}")
    print(f"  {'─'*30} {'─'*6}")
    for loc, count in locality_counts.most_common(20):
        print(f"  {loc:<30} {count:>6}")


def analyze_financial_health(companies):
    """Financial health indicators."""
    print("\n" + "=" * 80)
    print("  9. FINANCIAL HEALTH INDICATORS")
    print("=" * 80)

    found = [c for c in companies if c.get("ch_found")]
    total = len(found)

    insolvency = sum(1 for c in found if c.get("has_insolvency_history"))
    liquidated = sum(1 for c in found if c.get("has_been_liquidated"))
    charges = sum(1 for c in found if c.get("has_charges"))
    acct_overdue = sum(1 for c in found if c.get("accounts_overdue"))
    cs_overdue = sum(1 for c in found if c.get("confirmation_stmt_overdue"))

    print(f"\n  {'Indicator':<35} {'Count':>6} {'%':>7}")
    print(f"  {'─'*35} {'─'*6} {'─'*7}")
    for label, val in [
        ("Has insolvency history", insolvency),
        ("Has been liquidated", liquidated),
        ("Has charges (loans/security)", charges),
        ("Accounts overdue", acct_overdue),
        ("Confirmation statement overdue", cs_overdue),
    ]:
        pct = val / total * 100 if total else 0
        print(f"  {label:<35} {val:>6} {pct:>6.1f}%")

    # Officer count distribution
    officer_counts = [c.get("officer_count", 0) for c in found
                      if isinstance(c.get("officer_count"), (int, float))]
    if officer_counts:
        print(f"\n  OFFICER COUNT DISTRIBUTION:")
        oc = Counter(officer_counts)
        for count in sorted(oc.keys()):
            bar = "█" * oc[count]
            print(f"    {count} officers: {oc[count]:>4} companies {bar}")


def analyze_timeline(companies):
    """Incorporation and cessation timeline."""
    print("\n" + "=" * 80)
    print("  10. TIMELINE ANALYSIS")
    print("=" * 80)

    found = [c for c in companies if c.get("ch_found")]

    # Incorporation years
    inc_years = []
    for c in found:
        doc = str(c.get("date_of_creation", ""))
        if doc and len(doc) >= 4:
            try:
                inc_years.append(int(doc[:4]))
            except ValueError:
                pass

    if inc_years:
        decade_counts = Counter((y // 10) * 10 for y in inc_years)
        print(f"\n  INCORPORATIONS BY DECADE:")
        for decade in sorted(decade_counts.keys()):
            bar = "█" * decade_counts[decade]
            print(f"    {decade}s: {decade_counts[decade]:>4} {bar}")
        print(f"\n    Oldest: {min(inc_years)}")
        print(f"    Newest: {max(inc_years)}")

    # Cessation years (dissolved companies)
    cess_years = []
    for c in found:
        doc = str(c.get("date_of_cessation", ""))
        if doc and len(doc) >= 4:
            try:
                cess_years.append(int(doc[:4]))
            except ValueError:
                pass

    if cess_years:
        year_counts = Counter(cess_years)
        print(f"\n  CLOSURES BY YEAR:")
        for year in sorted(year_counts.keys()):
            bar = "█" * year_counts[year]
            print(f"    {year}: {year_counts[year]:>4} {bar}")


def analyze_name_patterns(companies):
    """Deep dive into name patterns — bigrams, trigrams."""
    print("\n" + "=" * 80)
    print("  11. NAME PATTERN ANALYSIS (Bigrams & Trigrams)")
    print("=" * 80)

    bigrams = Counter()
    trigrams = Counter()

    for c in companies:
        name = c.get("company_name", "") or c.get("original_name", "")
        words = extract_keywords_from_name(name)

        for i in range(len(words) - 1):
            bigrams[f"{words[i]} {words[i+1]}"] += 1
        for i in range(len(words) - 2):
            trigrams[f"{words[i]} {words[i+1]} {words[i+2]}"] += 1

    print(f"\n  TOP 30 BIGRAMS (two-word patterns):")
    print(f"  {'Bigram':<30} {'Count':>6}")
    print(f"  {'─'*30} {'─'*6}")
    for bg, count in bigrams.most_common(30):
        print(f"  {bg:<30} {count:>6}")

    print(f"\n  TOP 20 TRIGRAMS (three-word patterns):")
    print(f"  {'Trigram':<35} {'Count':>6}")
    print(f"  {'─'*35} {'─'*6}")
    for tg, count in trigrams.most_common(20):
        print(f"  {tg:<35} {count:>6}")


def print_summary_dashboard(companies):
    """Print a compact summary dashboard."""
    print("\n" + "=" * 80)
    print("  SUMMARY DASHBOARD")
    print("=" * 80)

    total = len(companies)
    found = [c for c in companies if c.get("ch_found")]
    active = sum(1 for c in found if c.get("company_status") == "active")
    dissolved = sum(1 for c in found if c.get("company_status") == "dissolved")

    all_sic = []
    for c in found:
        codes = [s.strip() for s in c.get("sic_codes", "").split(",") if s.strip()]
        all_sic.extend(codes)
    top_sic = Counter(all_sic).most_common(1)

    all_kw = []
    for c in companies:
        name = c.get("company_name", "") or c.get("original_name", "")
        all_kw.extend(extract_keywords_from_name(name))
    noise = {"company", "group", "holdings", "services", "enterprises",
             "trading", "international", "management", "solutions", "properties"}
    kw_counts = Counter(k for k in all_kw if k not in noise)
    top_kw = kw_counts.most_common(1)

    print(f"""
  ┌─────────────────────────────────────────────────┐
  │  Total entries:           {total:>6}                │
  │  Found in Companies House:{len(found):>6}                │
  │  Not found:               {total - len(found):>6}                │
  │                                                 │
  │  ACTIVE:                  {active:>6}                │
  │  DISSOLVED/CLOSED:        {dissolved:>6}                │
  │  OTHER:                   {len(found) - active - dissolved:>6}                │
  │                                                 │
  │  Most common SIC:  {(top_sic[0][0] + ' (' + str(top_sic[0][1]) + ')') if top_sic else 'N/A':>22}   │
  │  Most common keyword: {(top_kw[0][0] + ' (' + str(top_kw[0][1]) + ')') if top_kw else 'N/A':>19}   │
  └─────────────────────────────────────────────────┘
""")


# ═══════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════════

def export_to_excel(companies, sic_counts, keyword_counts, category_counts,
                    combo_counts, pair_counts, output_path):
    """Export all analysis results to a multi-sheet Excel file."""

    found = [c for c in companies if c.get("ch_found")]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # Sheet 1: All companies (raw enriched data)
        rows = []
        for c in companies:
            name = c.get("company_name") or c.get("original_name", "")
            cats = categorize_name(name)
            rows.append({
                "Original Name": c.get("original_name", ""),
                "CH Company Name": c.get("company_name", ""),
                "Company Number": c.get("company_number", ""),
                "Status": c.get("company_status", ""),
                "SIC Codes": c.get("sic_codes", ""),
                "Company Type": c.get("company_type", ""),
                "Date Created": c.get("date_of_creation", ""),
                "Date Ceased": c.get("date_of_cessation", ""),
                "Postcode": c.get("postal_code", ""),
                "Locality": c.get("locality", ""),
                "Region": c.get("region", ""),
                "Address": c.get("registered_address", ""),
                "Officer Count": c.get("officer_count", ""),
                "Has Charges": c.get("has_charges", ""),
                "Insolvency History": c.get("has_insolvency_history", ""),
                "Accounts Overdue": c.get("accounts_overdue", ""),
                "Keyword Categories": ", ".join(cats),
                "Found in CH": c.get("ch_found", False),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name="All Companies", index=False)

        # Sheet 2: SIC Code frequency
        sic_rows = [{"SIC Code": code, "Count": count,
                      "% of Companies": round(count / len(found) * 100, 1) if found else 0,
                      "Description": get_sic_description(code)}
                     for code, count in sic_counts.most_common()]
        pd.DataFrame(sic_rows).to_excel(writer, sheet_name="SIC Codes", index=False)

        # Sheet 3: Keyword frequency
        noise = {"company", "group", "holdings", "services", "enterprises",
                 "trading", "international", "management", "solutions", "properties"}
        kw_rows = [{"Keyword": kw, "Count": count,
                     "% of Companies": round(count / len(companies) * 100, 1)}
                    for kw, count in keyword_counts.most_common()
                    if kw not in noise]
        pd.DataFrame(kw_rows).to_excel(writer, sheet_name="Keywords", index=False)

        # Sheet 4: Keyword categories
        cat_rows = [{"Category": cat, "Count": count,
                      "% of Companies": round(count / len(companies) * 100, 1)}
                     for cat, count in category_counts.most_common()]
        pd.DataFrame(cat_rows).to_excel(writer, sheet_name="Keyword Categories", index=False)

        # Sheet 5: SIC + Keyword combinations
        combo_rows = [{"SIC + Keyword": combo, "Count": count,
                        "SIC Code": combo.split(" + ")[0],
                        "SIC Description": get_sic_description(combo.split(" + ")[0]),
                        "Keyword Category": combo.split(" + ")[1] if " + " in combo else ""}
                       for combo, count in combo_counts.most_common()]
        pd.DataFrame(combo_rows).to_excel(writer, sheet_name="SIC+Keyword Combos", index=False)

        # Sheet 6: SIC co-occurrence pairs
        pair_rows = [{"SIC Code A": a, "SIC Code B": b, "Count": count,
                       "Description A": get_sic_description(a),
                       "Description B": get_sic_description(b)}
                      for (a, b), count in pair_counts.most_common()]
        pd.DataFrame(pair_rows).to_excel(writer, sheet_name="SIC Co-occurrence", index=False)

        # Sheet 7: Status breakdown
        status_rows = []
        for c in found:
            codes = [s.strip() for s in c.get("sic_codes", "").split(",") if s.strip()]
            status_rows.append({
                "Company Name": c.get("company_name", ""),
                "Status": c.get("company_status", ""),
                "SIC Codes": c.get("sic_codes", ""),
                "Date Created": c.get("date_of_creation", ""),
                "Date Ceased": c.get("date_of_cessation", ""),
            })
        pd.DataFrame(status_rows).to_excel(writer, sheet_name="Status Detail", index=False)

        # Sheet 8: Chain detection
        name_groups = defaultdict(list)
        for c in companies:
            def norm(n):
                if not n: return ""
                n = n.strip().upper()
                for s in [" LTD", " LIMITED", " PLC", " LLP"]:
                    if n.endswith(s): n = n[:-len(s)].strip()
                return n
            cname = norm(c.get("company_name") or c.get("original_name", ""))
            if cname:
                name_groups[cname].append(c)

        chain_rows = []
        for name, entries in sorted(name_groups.items(), key=lambda x: -len(x[1])):
            if len(entries) < 2:
                continue
            statuses = Counter(e.get("company_status", "?") for e in entries)
            chain_rows.append({
                "Chain Name": name,
                "Locations": len(entries),
                "Active": statuses.get("active", 0),
                "Dissolved": statuses.get("dissolved", 0),
                "Company Number": entries[0].get("company_number", ""),
            })
        pd.DataFrame(chain_rows).to_excel(writer, sheet_name="Chains", index=False)

        # Sheet 9: Regional breakdown
        region_counts = Counter(c.get("region", "Unknown") or "Unknown" for c in found)
        region_rows = [{"Region": r, "Count": count,
                         "% of Found": round(count / len(found) * 100, 1) if found else 0}
                        for r, count in region_counts.most_common()]
        pd.DataFrame(region_rows).to_excel(writer, sheet_name="Regions", index=False)


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    # Determine input file
    if len(sys.argv) > 1:
        json_path = sys.argv[1]
    else:
        json_path = Path(__file__).resolve().parent / "diy_sheds_enriched.json"

    if not Path(json_path).exists():
        print(f"Error: File not found: {json_path}")
        print("Run diy_sheds_lookup.py first to generate the enriched data.")
        sys.exit(1)

    # Load data
    print(f"Loading: {json_path}")
    with open(json_path, "r", encoding="utf-8") as f:
        companies = json.load(f)

    print(f"Loaded {len(companies)} companies")

    # Run all analyses
    print_summary_dashboard(companies)
    sic_counts = analyze_sic_codes(companies)
    keyword_counts = analyze_keywords(companies)
    category_counts = analyze_keyword_categories(companies)
    combo_counts = analyze_sic_keyword_combinations(companies)
    pair_counts = analyze_sic_cooccurrence(companies)
    status_counts = analyze_status(companies)
    analyze_chains(companies)
    analyze_regions(companies)
    analyze_financial_health(companies)
    analyze_timeline(companies)
    analyze_name_patterns(companies)

    print("\n" + "=" * 80)
    print("  END OF ANALYSIS")
    print("=" * 80 + "\n")

    # ── Export to Excel ─────────────────────────────────────────────────────
    output_path = Path(__file__).resolve().parent / "diy_sheds_analysis.xlsx"
    export_to_excel(companies, sic_counts, keyword_counts, category_counts,
                    combo_counts, pair_counts, output_path)
    print(f"\nExcel report saved: {output_path}")


if __name__ == "__main__":
    main()
