# Vantage Sustainability Dashboard

An interactive, **single-file** dashboard for Vantage Specialty Chemicals covering
GHG emissions (Scopes 1–3) and water withdrawal, by site and company, 2021–2025 —
plus forward-looking **project scenario modeling** against internal targets.

See [`PRODUCT_SPEC.md`](PRODUCT_SPEC.md) for the full product specification.

## Open it

Double-click **`dashboard.html`** — it opens in any modern browser. No server, no
install, no internet required. All data is embedded in the file, so it stays on your
machine and is safe to email or drop on SharePoint.

The dashboard has five tabs:
- **Overview** — headline KPIs, Scope 1+2 trend, water trend.
- **Trends** — emissions by scope (Scope 3 toggle), top sites, Scope 3 categories,
  water; switch Absolute ⇄ Per-ton intensity.
- **Scenario Modeling** — toggle the initiative pipeline to see how projects bend the
  curve toward target, with CapEx / OpEx / payback and "gap-to-target closed."
- **Targets & Progress** — actuals vs. internal targets to 2030.
- **Water Risk** — site withdrawal vs. WWF water stress, with a sortable table.
- **Site Map** — global markers sized by emissions/water, colored by facility type (offline map).

## Refresh it each year

The dashboard is generated from the source Excel workbooks by one script.

```
pip install openpyxl
python build/build_data.py
```

This reads the workbooks in the repo root, writes the consolidated **`data.json`**,
and regenerates a self-contained **`dashboard.html`** (data + logo inlined). It prints
a validation report — confirm the Scope 1+2 totals still match the anchor numbers.

To add the new year:
1. Drop the new annual workbook(s) into the repo root and update the filenames in the
   `FILES` map at the top of `build/build_data.py`.
2. If a new site appears, add it to `SITE_MASTER` (with its aliases / water stress).
3. Run the script; check the validation report; commit the refreshed files.

## What's in here

| Path | Purpose |
|---|---|
| `dashboard.html` | **The deliverable** — open this. Self-contained. |
| `data.json` | Consolidated dataset (source of truth), generated. |
| `build/build_data.py` | ETL: workbooks → `data.json` → `dashboard.html`. |
| `build/dashboard_template.html` | App shell (HTML/CSS/JS, custom SVG charts). |
| `build/worldmap.txt` | Offline world-map SVG path (from `world-atlas`), inlined for the map. |
| `assets/vantage-logo.jpg` | Official logo, inlined into the dashboard. |
| `PRODUCT_SPEC.md` | Product specification. |
| `*.xlsx` | Source inventories (read-only inputs). |

## Data notes

- Emissions in **tCO₂e**; water in **megalitres (ML)**; production in **metric tons**.
- **Scope 2** defaults to **market-based** (toggle to location-based where available).
- **Scope 3** by category covers **2021–2024** (2021–22 from RouteZero; 2023–24 from the summary
  workbooks). 2025 Scope 3 isn't inventoried yet. Category 1 ≈ 81%.
- **Intensity** (per ton) covers **2022–2025** for the 7 manufacturing sites (no 2021
  production data).
- **2025 water** is converted from inventory gallons → ML; the large drop vs. 2024
  reflects reduced farm irrigation (Dateland).
- **Targets are internal** (not SBTi-validated); baseline year **2021**.
- Closed sites (e.g., Granollers) are hidden by default; use "Show closed sites."
