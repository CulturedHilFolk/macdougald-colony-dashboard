# MacDougald Lab — Mouse Colony Dashboard

Living dashboard of the MacDougald Lab mouse colony (Brian Desrosiers).

The dashboard summarizes:
- Strain inventory and demographics
- Active breedings and cohort plans
- Aging, sex balance, and culling priorities
- Walkthrough reconciliations vs. the lab's colony software

Live site: _set after GitHub Pages is enabled_

## Repo contents

| File | Purpose |
|---|---|
| `index.html` | The dashboard (served by GitHub Pages) |
| `data/colony.json` | Aggregated colony data feeding the charts |
| `build_blank_rack_template.py` | Generates a printable walkthrough form for physical inventory |
| `mouse_colony_analysis.ipynb` | Source notebook that processes raw exports → dashboard data |

## Source data (kept local, not in this repo)

Raw colony exports from the lab's colony software live in the local OneDrive folder and are excluded by `.gitignore`. The notebook reads them, derives aggregated/anonymized data, and writes `data/colony.json` which is then committed.

## Update workflow

1. Pull latest exports from the colony software into the local folder.
2. Run the notebook to regenerate `data/colony.json`.
3. Commit `data/colony.json` (and any HTML changes) and push.
4. GitHub Pages auto-rebuilds within ~30 seconds.
