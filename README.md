# MacDougald Lab — Mouse Colony Dashboard

Live dashboard of the MacDougald Lab mouse colony — inventory, breeding, cohort planning, and cull decisions.

**🌐 Live site:** https://culturedhilfolk.github.io/macdougald-colony-dashboard/

The shared URL is permanent — it never changes when the data is refreshed.

## Tabs

| Tab | What's there |
|---|---|
| **Overview** | Headline KPIs, strain / age / sex / use distributions, strain summary table |
| **Inventory** | Every mouse, with full filtering (search, strain, sex, use, age) and sortable columns |
| **Cages** | Visual rack 7A grid (rows 1–7 × cols A–J) — click any cage for occupants. Plus the unassigned-cages table |
| **Breeding** | Active pairs, performance bars/lines, full pair table |
| **Cohort Planning** | Each active research cohort: cross used, scientific purpose, expected Mendelian ratios, target genotype |
| **Cull Candidates** | Pre-defined filtered pools (aged > 12mo non-breeder, > 18mo any, sex unknown, etc.) — table is sortable, multi-selectable, and exportable to CSV |
| **Walkthrough** | Reconciliation tab for physical inventory walks (placeholder until next walkthrough) |
| **Experiment Details** | Colony purpose, husbandry standards, open data gaps |
| **Statistics** | Descriptive stats: age summary by strain, sex ratios, breeding productivity |

Top-right buttons:
- **ⓘ Info** — overall how-to-navigate
- **⚡ Summary** — auto-generated "what's interesting right now"
- **↻ How to Update** — instructions for refreshing the data

## Repo contents

| Path | Purpose |
|---|---|
| `index.html` | The dashboard (served by GitHub Pages) — single self-contained file with embedded data |
| `data/colony.json` | Aggregated colony data feeding the dashboard |
| `data/cohort_plans.json` | Hand-authored research-cohort goals + expected ratios (edit this to change cohort cards) |
| `build_dashboard_data.py` | Reads raw exports → writes `data/colony.json` |
| `build_dashboard_html.py` | Embeds `data/colony.json` into `index.html` |
| `mouse_colony_analysis.ipynb` | Notebook source-of-truth — runs the full pipeline end-to-end |
| `build_blank_rack_template.py` | Builds the printable physical-walkthrough form |

## Source data (kept local, NOT in this repo)

Raw colony exports live in the local OneDrive folder and are excluded by `.gitignore`:
- `CageListExcel.xlsx`
- `MacLab - Brian_Mice.xlsx`
- `StrainList.xlsx`
- `Breedings.xlsx`

The build pipeline reads them locally, derives aggregated/anonymized data, and writes the
artifacts that get committed.

## Updating the dashboard

1. Pull the latest exports from the colony software into the local folder (overwrite the four `.xlsx` above).
2. Open `mouse_colony_analysis.ipynb` and run all cells. (Or run the two `build_*.py` scripts from a terminal.)
3. Commit and push:
   ```bash
   git add data/colony.json index.html
   git commit -m "Refresh colony data <YYYY-MM-DD>"
   git push
   ```
4. GitHub Pages re-deploys within ~30 seconds. The shared URL stays the same.

To update **cohort plans** (target genotypes, ratios, scientific purpose), edit
`data/cohort_plans.json` and rerun the notebook.

## Contact

Brian Desrosiers — MacDougald Lab, University of Michigan
