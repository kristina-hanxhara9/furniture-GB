# Fast Fit Car Specialists — Companies House Discovery

## Definition

Fast Fit Car Specialists (Autocenters) generate >50% of turnover from quick
automotive repairs and services (exhaust, brakes, oil, filters). They do NOT
offer full car service/paintwork. Only organized retailers — independents are
classified as Car Garages.

**Examples:** Kwik Fit, Halfords Autocentre, ATS Euromaster, Mr Clutch, Formula One Autocentres

**Excluded:** Tyre Specialists, Car Accessory Retailers, Body Shops, Dealerships, Independent Garages

## SIC Code Strategy

| SIC Code | Description | Why |
|----------|-------------|-----|
| **45200** | Maintenance and repair of motor vehicles | PRIMARY — most fast-fit centres register here |
| **45319** | Wholesale of motor vehicle parts (other) | Some chains also wholesale parts |
| **45320** | Retail trade of motor vehicle parts/accessories | Shops with parts counter/online. Penalised if sole SIC |

### SIC Codes NOT Used

| SIC Code | Description | Why Excluded |
|----------|-------------|-------------|
| 45310 | Wholesale of parts (broad) | Too many pure wholesalers |
| 45400 | Motorcycles | Wrong segment |
| 45111 | Sale of cars | Dealerships |
| 45190 | Sale of other vehicles | Commercial vehicles |

## Scoring Logic

### Positive Keywords (+1 to +3)
- **+3:** Brand/concept names (fast fit, kwik fit, autocentre, pit stop, midas, speedy, norauto)
- **+2:** Service signals (exhaust, brake, oil change, MOT centre, clutch, muffler)
- **+1:** Generic signals (tyre, repair, servicing, garage, automotive)

### Penalty Keywords (-1 to -2)
- **-2:** independent, mobile, roadside, consultant, training
- **-1:** accessories, wholesale, distribution, parts only

### Structural Penalties
- **-1:** Tyre-only SIC (only 45320 registered)
- **-2:** Single officer (likely independent, not chain)

### Thresholds
- **KEEP** (score >= 2): Auto-include
- **REVIEW** (score 0-1): Needs manual review
- **DISCARD** (score < 0): Filtered out

## Hard Excludes (name contains)
Body shops, car sales, dealerships, car wash, valeting, recovery, breakdown,
scrap, salvage, rental, hire, tuning, motorsport, driving school, insurance, parking

## Usage

```bash
pip install requests python-dotenv
# Set COMPANIES_HOUSE_API_KEY in ../.env
python fast_fit_lookup.py
```

## Output
- `fast_fit_companies.csv` — all matched companies with scores, SIC codes, addresses, officer counts
