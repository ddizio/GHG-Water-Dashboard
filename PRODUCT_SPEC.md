# Vantage Sustainability Dashboard — Product Specification

**Owner:** Sustainability Director (ESG strategy, Vantage Specialty Chemicals)
**Status:** Draft v1.0 — for review before build
**Last updated:** 2026-06-17

---

## 1. Summary

An interactive, **single-file HTML dashboard** that turns Vantage's annual GHG (Scope 1–3)
and water spreadsheets into one place to **see trends, track progress against targets, and
model how past, current, and future reduction projects bend the curve.**

Today the data lives in separate Excel workbooks per year with no trending and no way to
test "what if we implement these projects?" This dashboard fixes that for a small internal
team, with no server, no logins, and no IT dependency — it opens in any browser and the
data stays local.

### Why these choices (plain language)
- **One HTML file** → you can email it, drop it on SharePoint, or open it offline. Sensitive
  figures (revenue, site detail) never leave your machine. No software to install.
- **One consolidated dataset** → the dashboard reads from a single clean data file we build
  from your spreadsheets. This solves the inconsistent site names and units across workbooks
  and makes the yearly refresh a 30-minute job instead of a rebuild.
- **All four capabilities in v1** → trends, scenario modeling, targets, and water risk, because
  they reinforce each other (a project's value is only clear against the trend and the target).

---

## 2. Goals & Non-Goals

### Goals
1. Show GHG (Scope 1, 2, 3) and water withdrawal **trends 2021→2025**, company-wide and by site.
2. Let the user **toggle reduction initiatives** (implemented, in progress, proposed) and see the
   projected impact on emissions/water vs. target — including cost (CapEx, OpEx savings, payback).
3. Track **actuals vs. reduction targets** through 2030, with gap-to-target at a glance.
4. Surface **water stress/risk by site** alongside withdrawal, so attention goes where it matters.
5. Be **reusable**: a documented annual refresh, not a throwaway prototype.

### Non-Goals (v1)
- Not a GHG **calculation engine** — it consumes inventory results already produced in the
  workbooks; it does not recompute emission factors.
- Not a **data-entry system** — editing happens in the source spreadsheets / consolidated file.
- No live database, user accounts, or multi-user editing.
- No automated pull from accounting/utility systems (future consideration).
- Not external/CDP disclosure-grade publishing in v1 (internal tool first; CDP-friendly exports later).

---

## 3. Audience & Use Cases

**Primary audience:** internal sustainability team (you + small team). Practical working tool.

**Representative use cases**
- *Trend check:* "How have Scope 1+2 emissions moved since 2021, and which sites drive it?"
- *Target gap:* "Are we on track to the 2030 target? How big is the gap if we do nothing?"
- *Scenario:* "If we implement only the *committed* projects, where do we land in 2027? What if we
  add the *under-investigation* ones?"
- *Capital case:* "What CapEx and payback does the committed project portfolio represent?"
- *Water risk:* "Which sites with high WWF water stress also have the largest withdrawal?"
- *Board prep:* "Export a clean year-over-year summary and a target-progress chart."

---

## 4. Source Data Inventory

All files are in the project folder. The dashboard does **not** read these directly at runtime;
they are the inputs to the consolidated dataset (Section 6).

| File | What it provides | Years | Notes |
|---|---|---|---|
| `Copy of RouteZero_results_template_v3_VantageFY21-22.xlsx` | Scope 1 & 2 results (market + location) | FY21–FY22 | RouteZero export; `RESULTS` sheet has scope/source rollups |
| `VSC_ FY23_Scope1&2_Inventory_v3_8192024.xlsx` | Scope 1 & 2 by site + water report | FY23 | `Overview`, `Pivot`, `Results by Site`, `RA- Water Report` |
| `VSC_ FY24_Scope1&2_Inventory_v2_06252025.xlsx` | Scope 1 & 2 by site + water report | FY24 | same structure as FY23 |
| `VSC_ FY25_Scope1&2_Inventory_v2_5-20-2026 (1).xlsx` | Scope 1 & 2 by site + water + fertilizer | FY25 | adds `Fertilizer` sheet; water in **gallons** here |
| `VSC_Scope 3_FY2023_Summary Workbook.xlsx` | Scope 3 by category (1–12) | FY23 | category sheets C1…C12 + `Scope 3 Summary` |
| `VSC_S1 S2 and S3_FY2024_Inventory Summary_v2.xlsx` | **Full S1+S2+S3** rollup by category | FY24 | best single-year total view (`Total Inventory Summary`) |
| `CDP Question List (2025 Update).xlsx` | Consolidated trends, water, targets, **initiatives** | 2021–2030 | richest summary file — see below |
| `Production Volumes - Vantage.xlsx` | **Site-level production** (lbs + tons), monthly + annual | 2022–2025 | 7 manufacturing sites; enables intensity. No 2021. `Prod Volumes` sheet |

### Key sheets inside `CDP Question List (2025 Update).xlsx`
- **`Data 2021 to 2023`** — Scope 1 & 2 (market) by site for 2021/2022/2023 + production tons. The
  cleanest multi-year, site-level trend we have.
- **`Summary Water 2024`** — withdrawal/discharge by site with facility type, river basin, and
  **WWF water stress / risk level**.
- **`Target Tracking`** — S1+2 emissions vs. target (2021–2030); water use vs. 25%/50% targets
  (2021–2030); and **initiative savings by year** (GHG MT CO2e, water ML).
- **`GHG-Water Initiatives` / `(2)`** — the project pipeline: project, site, **stage**, year, type,
  GHG savings, scope split, water savings, energy savings, incentive, CapEx, OpEx savings, payback.

### Data coverage matrix (what's actually available)
| Metric | 2021 | 2022 | 2023 | 2024 | 2025 |
|---|:--:|:--:|:--:|:--:|:--:|
| Scope 1 (site-level) | ✅ | ✅ | ✅ | ✅ | ✅ |
| Scope 2 market + location | ✅ | ✅ | ✅ | ✅ | ✅ |
| Scope 3 (by category) | ⚠️ partial | ⚠️ partial | ✅ | ✅ | ❌ not yet |
| Water withdrawal (company) | ✅ | ✅ | ✅ | ✅ | ✅ |
| Water by site + stress/risk | — | — | — | ✅ | ✅ |
| Production tons (site-level) | ❌ | ✅ | ✅ | ✅ | ✅ |
| Initiatives / targets | — | — | — | ✅ 2024–2030 | ✅ |

**Design implications:**
- Scope 3 trends render for **2023–2024 only** with a clear "limited history" note; Scope 1/2 and
  water carry the full **2021–2025** trend.
- **Intensity (per ton) is available 2022–2025** for the 7 manufacturing sites; 2021 intensity is
  not shown (no 2021 production). Company-wide intensity uses the production total for those years.
- **2025 water is final** (confirmed); the build converts it from gallons to ML (Section 7).

### Anchor numbers (sanity checks for the build)
- **Scope 1+2 (market), tCO2e:** 2021 ≈ 107,249 · 2022 ≈ 106,088 · 2023 ≈ 88,585 · 2024 ≈ 82,258 · 2025 ≈ 80,144
- **2024 total incl. Scope 3:** ≈ 891,695 tCO2e; Scope 3 ≈ 809,400, of which **Category 1 (Purchased
  goods & services) ≈ 726,053 = ~81%**.
- **Water withdrawal (ML):** 2021 ≈ 16,522 · 2022 ≈ 15,202 · 2023 ≈ 14,434 · 2024 ≈ 13,155 · 2025 ≈ TBD
  (converted from gallons during build).
- **Production (tons):** 2022 ≈ 367,234 · 2023 ≈ 328,709 · 2024 ≈ 363,159 · 2025 ≈ 335,868 (7
  manufacturing sites: Gurnee, Chicago/Oleo, Leuna, Carnegie/Mallet, Englewood, Spain/Les Borges, Tucson).
- **Top emitting sites:** Chicago, Gurnee, Leuna (together the majority of Scope 1+2).
- **Targets (internal, not SBTi-validated):** S1+2 target ≈ **85,799 tCO2e** (~20% below 2021);
  water targets **25% = 12,392 ML**, **50% = 8,261 ML**.
- **~29 operating sites** across US, Europe (Leuna, Barcelona, Granollers, Les Borges, Paris),
  Latin America, India, China, South Africa.

---

## 5. Scope of v1 (the four modules)

1. **Trends** — GHG by scope and water, over time, company and per-site.
2. **Scenario Modeling** — toggle initiatives → projected trajectory vs. target + cost summary.
3. **Targets & Progress** — actuals vs. target lines through 2030, gap-to-target.
4. **Water Stress / Risk** — site table/map of withdrawal × WWF stress/risk.

A persistent **global filter bar** (year range, site(s), scope basis) ties them together.

---

## 6. Consolidated Data Model (single source of truth)

We build one normalized data file from the spreadsheets. Recommended format: a single
`data.json` (embedded into the HTML at build time) generated by a small, documented Python
script (`build_data.py`) so the yearly refresh is repeatable. Mirror to CSVs for transparency.

### Dimension: `sites`
| field | type | example | notes |
|---|---|---|---|
| `site_id` | string | `chicago_il` | stable slug, the join key |
| `name` | string | `Chicago, IL` | canonical display name |
| `aliases` | string[] | `["Vantage_Chicago","Chicago"]` | resolves naming variants across files |
| `country` / `region` | string | `USA` / `North America` | |
| `facility_type` | enum | `Factory` | Factory / Farm / Warehouse / Office |
| `status` | enum | `active` | `active` / `closed` (closed = all-zero across years, or tapered to zero) |
| `river_basin` | string | `Michigan` | from water summary |
| `wwf_water_stress` | enum | `High` | High / Medium / Low / NA |
| `wwf_water_risk` | enum | `High` | High / Medium / Low / NA |
| `lat` / `lon` | number | optional | only if we add a map |

### Fact: `emissions`
| field | type | notes |
|---|---|---|
| `site_id` | string | `company` for rollup-only rows |
| `year` | int | 2021–2025 |
| `scope` | enum | `1` / `2` / `3` |
| `category` | string | S1/S2 source (Natural Gas, Electric Power, Steam…) or S3 cat # + name |
| `basis` | enum | `market` / `location` (S2 only; null otherwise) |
| `value_tco2e` | number | metric tons CO2e |

### Fact: `water`
| field | type | notes |
|---|---|---|
| `site_id`, `year` | | |
| `withdrawal_ml` | number | normalized to **megaliters (ML)** |
| `discharge_ml` | number | |
| `consumption_ml` | number | withdrawal − discharge |
| `source_breakdown` | object | surface / groundwater / third-party / recycled (ML) |

### Fact: `production`
`site_id`, `year`, `production_tons` — enables **intensity** metrics (tCO2e / ton, ML / ton).

### Fact: `targets`
| field | type | notes |
|---|---|---|
| `metric` | enum | `ghg_s1s2` / `water` |
| `year` | int | 2021–2030 |
| `target_value` | number | tCO2e or ML |
| `scenario` | string | e.g. `base`, `water_25`, `water_50` |

### Fact: `initiatives`
| field | type | notes |
|---|---|---|
| `project_id` | string | |
| `name` | string | |
| `site_id` | string | |
| `type` | enum | `GHG` / `Water` |
| `stage` | enum | `Under Investigation` / `To be implemented` / `Implementation commenced` / `Implemented` / `Not to be implemented` |
| `year` | int | expected/actual impact year |
| `ghg_savings_tco2e` | number | |
| `scope1_savings` / `scope2_savings` | number | split where known |
| `water_savings_ml` | number | |
| `energy_savings_mwh` | number | |
| `capex_usd` / `opex_savings_usd` / `incentive_usd` | number | economics |
| `payback_years` | number | simple payback with incentive |

---

## 7. Data Quality & Normalization (must-resolve before/while building)

These come straight from inspecting the files. The build script handles each explicitly and logs
what it did, so numbers are auditable.

1. **Site naming differs across files** — `Vantage_Chicago` vs `Chicago, IL` vs `Chicago`; the
   production file adds business-unit labels: **`Chicago (Oleo)` → Chicago**, **`Carnegie (Mallet
   Ingredients)` → Carnegie**, and **`Spain` → Les Borges** (the Spanish *manufacturing* site; not
   the Barcelona/Granollers offices). Resolved via the `sites.aliases` map; any unmapped name is
   flagged, not silently dropped.
2. **Units vary.** Emissions are tCO2e throughout (good). Water is **ML in summaries but gallons in
   the FY25 inventory and m³/L in some site source rows.** Normalize everything to **ML**; keep a
   documented conversion table (gal→ML, m³→ML, L→ML). **2025 water is confirmed final** — the build
   converts the FY25 gallons figures to ML and adds 2025 to the water trend and site-level
   stress/risk view. Production is in lbs + tons; **use tons** (1 ton = 2,000 lbs) for intensity.
3. **Scope 2 dual basis.** Every S2 figure has market- and location-based values. Default the UI to
   **market-based**, with a toggle to location-based. Never mix the two in one total.
4. **Scope 3 history is thin.** Full category data exists only for **2023 and 2024**; 2021–22 is
   partial (RouteZero) and 2025 isn't done. Show Scope 3 with an explicit coverage caveat; don't
   imply a 5-year Scope 3 trend.
5. **Error cells & blanks** (`#DIV/0!`, empty production for some sites/years) → coerce to null and
   exclude from intensity math; never render as 0 where it means "missing."
6. **Methodology changes** noted in the FY24 summary (e.g., Cat 5 waste, Cat 12 EOL emission-factor
   updates) cause large year-over-year swings. Carry a `notes` field and surface it on hover so a
   −87% change reads as "methodology change," not a real drop.
7. **Production coverage is 2022–2025, 7 sites** — intensity (tCO2e/ton, ML/ton) renders for those
   site/years only. **2021 has no production**, so 2021 intensity is hidden and the intensity x-axis
   starts at 2022. Non-manufacturing sites (offices/warehouses) have no production and show absolute
   values only.
8. **Closed sites** — `status=closed` when a site is zero across all metrics/years, or once it
   tapers to zero (e.g., Granollers, Delta, Linden, Warren). Closed sites are hidden from the default
   "active" views but retained in the dataset; a "show closed" toggle exposes their history so a
   company-total drop isn't mistaken for a reduction.

---

## 8. Functional Requirements by Module

### Global filter bar (persistent)
- Year range (2021–2025; targets extend the x-axis to 2030).
- Site selector: All / region / multi-select individual sites.
- Scope-2 basis toggle: Market (default) / Location.
- Absolute vs. intensity toggle (per ton) where production exists.
- Units note always visible (tCO2e, ML).

### Module 1 — Trends
- **GHG over time:** stacked area/bar of Scope 1 / 2 / 3 by year; toggle scopes on/off.
- **Scope 1+2 line** (the headline metric) with target overlay.
- **Per-site breakdown:** ranked bar ("top emitters") and small-multiples or a site drill-down.
- **Scope 3 composition:** category bar for 2023/2024 highlighting Category 1 dominance.
- **Water trend:** withdrawal by year, with discharge/consumption secondary; per-site ranked bar.
- Hover tooltips show exact value, YoY %, and any methodology note.

### Module 2 — Scenario Modeling (the differentiator)
- List initiatives grouped by **stage**, each with a checkbox; quick presets:
  - *Implemented only* (what's banked) · *Committed* (implemented + commenced + to-be) ·
    *All including under-investigation* (best case) · *Custom*.
- Selecting initiatives recomputes the **forward trajectory**:
  `projected[year] = baseline[year] − Σ savings of selected initiatives active by that year`
  (cumulative; GHG in tCO2e, water in ML), plotted against the **target line**.
- **Baseline (BAU) for 2026→2030 is held flat at the last actual** (confirmed). Stated on-chart as
  "Baseline held flat at 2025 actual." (A growth/decline assumption can be added later in Phase 3.)
- **Cost panel** for the selected set: total CapEx, total annual OpEx savings, total incentives,
  blended simple payback, and resulting **gap-to-target closed (%)**.
- Filter initiatives by site and type (GHG/Water).
- Clear labeling that future years are **projections**, with the baseline assumption stated.

### Module 3 — Targets & Progress
- S1+2 actuals (2021–2025) vs. target trajectory to 2030; gap-to-target callout.
- Water actuals vs. 25% and 50% reduction scenarios.
- Progress indicators (e.g., "2024: 4% below target" / "X ML to 25% goal").
- Optional baseline-year reference line (2021).

### Module 4 — Water Stress / Risk
- Site table: name, facility type, river basin, **WWF stress**, **WWF risk**, withdrawal (ML),
  discharge, consumption — sortable, with stress/risk color-coded.
- Emphasis view: sites that are **High stress AND high withdrawal** (the priority quadrant) — a
  scatter (withdrawal × stress) or a filtered list.
- Optional choropleth/marker map if `lat/lon` added (phase 2).

### Cross-cutting
- **Export:** download current chart as PNG and underlying data as CSV.
- **"Data as of" stamp** and source-file provenance footer.
- **Empty/limited-data states** handled gracefully (esp. Scope 3 history and 2021 intensity).

---

## 9. Information Architecture

Single page, top nav switches modules; global filter bar stays fixed.

```
┌──────────────────────────────────────────────────────────────┐
│  ◴ Vantage logo     Sustainability Dashboard      Data: FY2025 │
│  [ Year ▸ ] [ Site ▸ ] [ S2: Market ▸ ] [ Abs/Intensity ]      │
├──────────────────────────────────────────────────────────────┤
│  Overview │ Trends │ Scenario Modeling │ Targets │ Water Risk  │
├──────────────────────────────────────────────────────────────┤
│  KPI cards: S1+2 (tCO2e) │ Total incl. S3 │ Water (ML) │ vs Tgt│
│                                                                │
│  [ main chart area for the active module ]                     │
│                                                                │
│  [ secondary: per-site table / cost panel / category bar ]     │
└──────────────────────────────────────────────────────────────┘
```

An **Overview** landing tab shows the four KPI cards + headline trend so the most common question
("where are we vs. target?") is answered in one glance.

---

## 10. Visual & Brand Design System

From the 2026 Vantage Style Guide. Use these exactly.

### Color palette
| Role | Name | HEX |
|---|---|---|
| **Primary 1** (brand) | Fresh Green | `#00B388` |
| **Primary 2** | Medium Green | `#008264` |
| **Primary blue** | New Dark Blue | `#015B9A` |
| **Accent green** | New Bright Green | `#97C93C` |
| Text-only green | Darker Green | `#038264` |
| Accent 1 | Blue | `#1AB9DE` (75% `#57CBE7`, 50% `#BFDCEF`) |
| Accent 2 | Orange Pop | `#F15F22` (75% `#F58558`, 50% `#F8AE90`) |
| Accent 3 | Sunny Yellow | `#F3B64F` (75% `#F6C878`, 50% `#F9DAA7`) |
| Text 1 | Dark Charcoal | `#161617` |
| Text 2 | Dark Grey | `#3D3B3E` |
| Text 3 | Medium Grey | `#4D4B4B` |
| Neutral | Grey | `#70706F` |
| Background 1 | White | `#FFFFFF` |
| Background 2 | Web Blue Grey 50% | `#E6E4E6` |
| Background 3 | Web Blue Grey 25% | `#F2F1F2` |

**Data-encoding conventions (consistent everywhere):**
- **Scope 1** → Fresh Green `#00B388` · **Scope 2** → New Dark Blue `#015B9A` ·
  **Scope 3** → Grey `#70706F` (it dwarfs 1&2; keep it neutral so 1&2 stay legible).
- **Water** → Accent Blue `#1AB9DE`.
- **Target lines** → Orange Pop `#F15F22` (dashed).
- **Initiative savings / positive progress** → New Bright Green `#97C93C`.
- **Warnings / over-target / high water stress** → Orange Pop / Sunny Yellow ramp.
- Use accent tints (75/50%) for category sub-breakdowns.

### Typography
- **Headings:** Poppins (brand font; load via Google Fonts, weights 500/600).
- **Body/UI:** Poppins or system sans fallback; **Century Gothic** is the MS-default equivalent for
  any exported Office collateral. Numbers in tabular figures for alignment.
- Fallback stack: `Poppins, "Century Gothic", "Open Sans", system-ui, sans-serif`.

### Components
- **Buttons:** Primary = Fresh Green fill, white text; Secondary = New Dark Blue. Rounded
  (~6–8px), matching the style guide's pill buttons. Active states darken one step.
- **KPI cards:** white on Background 2/3, dark-charcoal numerals, small trend sparkline + YoY chip.
- **Logo:** official Vantage wordmark (`assets/vantage-logo.jpg`, water-drop mark) top-left —
  thematically perfect for a water + emissions tool. Inlined as base64 in the HTML.
- Generous white space, minimal gridlines, accessible contrast (see Section 13).

---

## 11. Technical Architecture

- **Deliverable:** one `dashboard.html` — self-contained, opens via double-click, works offline.
- **Charts:** a single lightweight library. **Recommendation: Chart.js** (small, MIT, easy
  line/bar/area/scatter; good tooltips). Plotly is the alternative if we later want the map/zoom.
  Pinned, vendored locally (no CDN) so it runs offline and is reproducible.
- **Data embedding:** consolidated `data.json` inlined into the HTML at build time → no fetch, no
  CORS, no server. Keeps sensitive figures local.
- **Build script:** `build_data.py` (pandas/openpyxl) reads the workbooks → normalizes per Sections
  6–7 → writes `data.json` → injects into `dashboard.html`. Documented and rerunnable each year.
- **No backend, no tracking, no external calls** at runtime.
- **Size/perf:** dataset is small (hundreds of rows); everything renders client-side instantly.
- **Browsers:** current Chrome/Edge/Firefox/Safari.

### Repository layout (proposed)
```
/data_sources/            # original Excel workbooks (read-only inputs)
/build/build_data.py      # Excel → normalized data.json
/build/sites.yml          # canonical site map + aliases + stress/risk
/data.json                # generated consolidated dataset (source of truth)
/dashboard.html           # the deliverable (data + logo inlined at build)
/assets/vantage-logo.jpg  # official Vantage logo (inlined as base64)
/vendor/chart.min.js      # pinned chart library
PRODUCT_SPEC.md           # this file
README.md                 # how to refresh + open
```

---

## 12. Metric Definitions & Units (glossary)

- **Scope 1** — direct emissions from owned/controlled sources (natural gas, diesel, propane,
  refrigerants, fertilizer). Unit: tCO2e.
- **Scope 2** — indirect from purchased electricity/steam. **Market-based** (contractual/supplier
  factors, default) vs **location-based** (grid-average). Unit: tCO2e.
- **Scope 3** — value-chain emissions across 15 categories; Vantage's material categories are 1, 3,
  4, 5, 6, 7, 8, 11, 12, with **Category 1 (purchased goods & services)** dominant. Unit: tCO2e.
- **Water withdrawal** — total water taken from all sources (surface, ground, third-party,
  recycled). Unit: **megaliters (ML)** in the dashboard.
- **Discharge / Consumption** — water returned vs. withdrawal − discharge. Unit: ML.
- **Intensity** — emissions or water per metric ton of production (where production exists).
- **Initiative / project** — a discrete action reducing GHG and/or water, with a lifecycle stage and
  economics (CapEx, OpEx savings, incentive, simple payback).
- **Target** — **internal** (not SBTi-validated) reduction goals: S1+2 ≈ 85,799 tCO2e (~20% vs.
  2021); water 25%/50% reduction milestones. Label as "Internal target" in the UI.

---

## 13. Accessibility & Usability

- WCAG AA contrast for text and key data marks; never rely on color alone (use labels, patterns,
  direct value annotations) — important because the green/blue brand palette can be hard for some
  color-vision types.
- Keyboard-navigable controls; visible focus states.
- Tooltips and legends in plain language; jargon defined on hover/info icons.
- Sensible defaults so a first-time viewer sees something meaningful with zero clicks.
- Responsive down to a laptop screen (desktop-first; mobile is best-effort).

---

## 14. Annual Refresh Workflow (keep it reusable)

1. Drop the new year's workbook(s) into `/data_sources/`.
2. Add any new site to `/build/sites.yml` (name, aliases, region, stress/risk).
3. Run `python build/build_data.py` → it validates, normalizes, logs unmapped names/units, and
   regenerates `data.json` + `dashboard.html`.
4. Eyeball the anchor numbers (Section 4) against the workbook rollups.
5. Commit + share the refreshed `dashboard.html`.

Target: **≤ 30 minutes** per yearly update once site mapping is stable.

---

## 15. Phasing

- **Phase 1 (v1):** consolidated dataset + Overview, Trends, Targets, Water Risk; Scenario Modeling
  with stage presets and cost panel. All four capabilities, market-based default.
- **Phase 2:** site map (lat/lon), intensity views everywhere production allows, CSV/PNG export
  polish, location-based comparisons.
- **Phase 3:** CDP-friendly export pack; per-initiative drill-down with documents; optional
  what-if editor to add a hypothetical project on the fly.

---

## 16. Success Criteria

- A team member can answer "are we on track to target, and which projects close the gap?" in
  **under a minute**, unaided.
- Every headline number reconciles to the source workbooks (auditable).
- Annual refresh done by the team without rebuilding the app.
- Used in real planning/board prep within one quarter of launch.

---

## 17. Decisions Confirmed & Remaining Questions

### Confirmed (2026-06-17)
- ✅ **2025 water is final** → convert gallons→ML and include 2025 in water trend + site stress/risk.
- ✅ **Target is internal** (not SBTi-validated) → label "Internal target."
- ✅ **Target baseline year = 2021** (107,249 tCO2e S1+2; water 16,522 ML). Reductions measured vs. 2021.
- ✅ **Scenario baseline** holds **flat from the last actual** for 2026–2030.
- ✅ **Production data provided** → site-level tons 2022–2025 (7 manufacturing sites); intensity
  enabled for those years. **No 2021 production** → 2021 intensity not shown.
- ✅ **Closed-site rule** → any site that is **zero across all metrics and all years** is marked
  `closed` and excluded from active views (kept in data for auditability). Sites that taper to zero
  mid-period (e.g., **Granollers — no longer operating**; Delta, Linden, Warren) keep their history
  and are marked `closed` from their first all-zero year.
- ✅ **Spain = Les Borges** (manufacturing). **Barcelona = office** (no production). Confirmed.
- ✅ **Official logo incorporated** → `assets/vantage-logo.jpg` (the real Vantage wordmark, JPEG
  415×155, white background). Embedded as a base64 data-URI in the single-file HTML so it stays
  self-contained. Header sits on a white/Background-1 surface so the white logo background is seamless.

### Still open
- None blocking. Future input welcome on a growth/decline BAU assumption (Phase 3) and lat/lon for a
  site map (Phase 2).
