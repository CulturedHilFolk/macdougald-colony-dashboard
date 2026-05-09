"""
Build data/colony.json from the latest colony exports.

Reads:
  - CageListExcel.xlsx
  - MacLab - Brian_Mice.xlsx
  - StrainList.xlsx
  - Breedings.xlsx
  - data/cohort_plans.json   (user-authored research goals; preserved)

Writes:
  - data/colony.json          (everything the dashboard renders)
"""

from pathlib import Path
from datetime import date, datetime
import json
import re
import pandas as pd

DIR = Path("/Users/briandes/Library/CloudStorage/OneDrive-MichiganMedicine/Desktop/MacDougald Lab/Mouse Colony")
DATA = DIR / "data"
DATA.mkdir(exist_ok=True)

TODAY = date.today()


def clean_strain_label(s):
    if pd.isna(s):
        return ""
    return str(s).replace("|~|", "").strip().strip("|").strip() or "(blank)"


def parse_pos(p):
    if pd.isna(p) or p == "Unassigned":
        return None
    m = re.match(r"^(\d+)([A-Z])$", str(p).strip())
    if not m:
        return None
    return {"row": int(m.group(1)), "col": m.group(2)}


def age_bucket(days):
    if pd.isna(days):
        return "Unknown"
    d = float(days)
    if d < 60: return "0–2 mo"
    if d < 120: return "2–4 mo"
    if d < 180: return "4–6 mo"
    if d < 365: return "6–12 mo"
    return ">12 mo"


# ---------- Load ----------
cages = pd.read_excel(DIR / "CageListExcel.xlsx")
mice = pd.read_excel(DIR / "MacLab - Brian_Mice.xlsx")
strains = pd.read_excel(DIR / "StrainList.xlsx")
breedings = pd.read_excel(DIR / "Breedings.xlsx")

# Strip header rows
mice = mice[mice["Strain"].notna()].copy()
mice["Cage ID"] = mice["Cage ID"].astype(str)
cages["Name"] = cages["Name"].astype(str)

# ---------- Mouse-level enrichment ----------
mice["age_bucket"] = mice["Age"].apply(age_bucket)
mice["age_months"] = (mice["Age"].astype(float) / 30.0).round(1)
mice["DOB_str"] = pd.to_datetime(mice["DOB"]).dt.strftime("%Y-%m-%d")
mice["genotype_clean"] = mice["Genotype"].astype(str).str.replace("Adipoq-cre", "Aq-cre", regex=False)

# ---------- Cull-candidate flags ----------
def cull_flags(row):
    flags = []
    if row["Age"] and row["Age"] >= 365:
        flags.append("aged_>12mo")
    if row["Age"] and row["Age"] >= 540:
        flags.append("aged_>18mo")
    if row["Sex"] == "Unknown":
        flags.append("sex_unknown")
    return flags

mice["cull_flags"] = mice.apply(cull_flags, axis=1)

# ---------- Strain filtering helpers ----------
def is_strain(s, target):
    return clean_strain_label(s).lower() == target.lower()

# ---------- Aggregations ----------
strain_order = ["AdipoGlo", "AdipoGlo+", "Dendra2", "Marrow Glo", "mTmG", "Adipoq-Cre"]
strain_order = [s for s in strain_order if s in mice["Strain"].unique()] + \
               [s for s in mice["Strain"].unique() if s not in strain_order]

strain_counts = (mice.groupby("Strain").size()
                 .reindex(strain_order, fill_value=0).to_dict())

sex_by_strain = {}
for s in strain_order:
    sub = mice[mice["Strain"] == s]
    sex_by_strain[s] = {
        "Female": int((sub["Sex"] == "Female").sum()),
        "Male": int((sub["Sex"] == "Male").sum()),
        "Unknown": int((sub["Sex"] == "Unknown").sum()),
    }

age_buckets = ["0–2 mo", "2–4 mo", "4–6 mo", "6–12 mo", ">12 mo", "Unknown"]
age_dist_total = {b: int((mice["age_bucket"] == b).sum()) for b in age_buckets}
age_dist_by_strain = {}
for s in strain_order:
    sub = mice[mice["Strain"] == s]
    age_dist_by_strain[s] = {b: int((sub["age_bucket"] == b).sum()) for b in age_buckets}

use_by_strain = {}
for s in strain_order:
    sub = mice[mice["Strain"] == s]
    use_by_strain[s] = {
        "Available": int((sub["Use"] == "Available").sum()),
        "Breeding": int((sub["Use"] == "Breeding").sum()),
    }

# Top genotypes by strain
geno_by_strain = {}
for s in strain_order:
    sub = mice[mice["Strain"] == s]
    counts = sub["genotype_clean"].value_counts().head(8)
    geno_by_strain[s] = [{"genotype": g, "count": int(c)} for g, c in counts.items()]

# ---------- Cull candidates summary ----------
cull_groups = {
    "aged_>12mo_available_nonbreeder": mice[
        (mice["Age"] >= 365) & (mice["Use"] == "Available")],
    "aged_>18mo_any": mice[mice["Age"] >= 540],
    "sex_unknown": mice[mice["Sex"] == "Unknown"],
    "available_no_breeders_strain": mice[
        (mice["Use"] == "Available") & (mice["Strain"] == "Adipoq-Cre")],
}
cull_summary = {k: {"count": int(len(v)), "ids": v["Mouse ID"].astype(str).tolist()[:50]}
                for k, v in cull_groups.items()}

# ---------- Cage-level inventory ----------
def summarize_cage(group):
    return pd.Series({
        "n_mice": int(len(group)),
        "ids": group["Mouse ID"].astype(str).tolist(),
        "sexes": group["Sex"].fillna("Unknown").tolist(),
        "ages_months": group["age_months"].tolist(),
        "genotypes": group["genotype_clean"].tolist(),
        "strain": group["Strain"].iloc[0],
        "breeders_in_cage": int((group["Use"] == "Breeding").sum()),
    })

per_cage_records = mice.groupby("Cage ID").apply(summarize_cage, include_groups=False)
per_cage_records = per_cage_records.reset_index()

cages_full = cages.merge(per_cage_records, how="left", left_on="Name", right_on="Cage ID")
cages_full["strain_clean"] = cages_full["Strains"].apply(clean_strain_label)
cages_full["position"] = cages_full["Rack Location"].apply(parse_pos)
cages_full["n_mice"] = cages_full["n_mice"].fillna(0).astype(int)

cage_records = []
for _, r in cages_full.iterrows():
    cage_records.append({
        "cage_id": str(r["Name"]),
        "strain": r["strain_clean"],
        "n_mice": int(r["n_mice"]),
        "expected_count": int(r.get("Mice", 0) or 0),
        "rack": "7A" if r["Rack Location"] != "Unassigned" else None,
        "room": r["Room Name"] if not pd.isna(r["Room Name"]) else None,
        "position": r["position"],
        "has_breeders": r["Has Breeders"] == "Yes",
        "ids": r["ids"] if isinstance(r["ids"], list) else [],
        "sexes": r["sexes"] if isinstance(r["sexes"], list) else [],
        "ages_months": r["ages_months"] if isinstance(r["ages_months"], list) else [],
        "genotypes": r["genotypes"] if isinstance(r["genotypes"], list) else [],
        "breeders_in_cage": int(r["breeders_in_cage"]) if not pd.isna(r.get("breeders_in_cage")) else 0,
        "date_created": pd.Timestamp(r["Date Created"]).strftime("%Y-%m-%d") if not pd.isna(r["Date Created"]) else None,
    })

# ---------- Inventory (mouse-level) ----------
inventory = []
for _, m in mice.iterrows():
    inventory.append({
        "mouse_id": str(m["Mouse ID"]),
        "use": m["Use"],
        "strain": m["Strain"],
        "sex": m["Sex"],
        "age_days": int(m["Age"]) if not pd.isna(m["Age"]) else None,
        "age_months": float(m["age_months"]) if not pd.isna(m["age_months"]) else None,
        "dob": m["DOB_str"] if not pd.isna(m["DOB_str"]) else None,
        "genotype": m["Genotype"] if not pd.isna(m["Genotype"]) else "",
        "cage_id": str(m["Cage ID"]),
        "cull_flags": m["cull_flags"],
    })

# ---------- Breedings ----------
breeding_records = []
for _, b in breedings.iterrows():
    active_date = pd.Timestamp(b["Active Date"]) if not pd.isna(b["Active Date"]) else None
    days_active = (datetime.now() - active_date).days if active_date is not None else None
    breeding_records.append({
        "name": b["Name"],
        "status": b["Status"],
        "offspring_strain": b["Offspring Strain"],
        "mother": str(b["Mother"]),
        "father": str(b["Father"]),
        "active_date": active_date.strftime("%Y-%m-%d") if active_date is not None else None,
        "litters": int(b["Litters"]) if not pd.isna(b["Litters"]) else 0,
        "born": int(b["Born"]) if not pd.isna(b["Born"]) else 0,
        "avg_litter_size": round(float(b["Born"]) / float(b["Litters"]), 1) if b["Litters"] else None,
        "last_weaned": str(b["Last Weaned (wk)"]) if not pd.isna(b["Last Weaned (wk)"]) else "—",
        "mother_geno": b["Mother Genotype"],
        "father_geno": b["Father Genotype"],
        "days_active": days_active,
        "months_active": round(days_active / 30.0, 1) if days_active else None,
    })

# ---------- Strain summary table ----------
strain_summary = []
for _, s in strains.iterrows():
    strain_summary.append({
        "strain": s["Strain Name"],
        "active_breedings": int(s["Active Breedings"] or 0),
        "active_cages": int(s["Active Cages"] or 0),
        "active_litters": int(s["Active Litters"] or 0),
        "alive_mice": int(s["Alive Mice"] or 0),
        "dead_mice": int(s["Dead Mice"] or 0),
        "n_mutations": int(s["No. of Mutations"] or 0),
    })

# ---------- Load preserved cohort plans ----------
with open(DATA / "cohort_plans.json") as f:
    cohort_plans = json.load(f)

# ---------- Top-level summary KPIs ----------
total_mice = int(len(mice))
total_cages = int(len(cages))
unassigned_cages = int((cages["Rack Location"] == "Unassigned").sum())
located_cages = total_cages - unassigned_cages
active_breeders = int((mice["Use"] == "Breeding").sum())

summary = {
    "as_of": TODAY.isoformat(),
    "total_mice": total_mice,
    "total_cages": total_cages,
    "located_cages": located_cages,
    "unassigned_cages": unassigned_cages,
    "active_breeders": active_breeders,
    "active_breeding_pairs": int((breedings["Status"] == "Active").sum()),
    "available_mice": int((mice["Use"] == "Available").sum()),
    "aged_over_12mo": int((mice["Age"] >= 365).sum()),
    "aged_over_18mo": int((mice["Age"] >= 540).sum()),
    "unknown_sex": int((mice["Sex"] == "Unknown").sum()),
    "n_strains": int(mice["Strain"].nunique()),
}

# ---------- Cage costs (UMich ULAM CY26 rates, effective 04/01/2026) ----------
RATE_VENT_STD = 1.19
RATE_VENT_BREED = 2.16

def cage_rate(c):
    return RATE_VENT_BREED if c["has_breeders"] else RATE_VENT_STD

cost_total_day = 0.0
cost_breakdown = {
    "by_strain": {},
    "by_breeding": {"breeding": 0.0, "non_breeding": 0.0},
    "by_housing": {"single": 0.0, "group": 0.0, "empty": 0.0},
    "by_room": {},
    "by_rack_status": {"located": 0.0, "unassigned": 0.0},
    "by_sex_dominant": {"female": 0.0, "male": 0.0, "mixed": 0.0, "empty": 0.0},
}
strain_lines = {}
strain_geno_lines = {}
cost_per_cage = []

for c in cage_records:
    rate = cage_rate(c)
    cost_total_day += rate
    strain = c["strain"] or "(blank)"
    cost_breakdown["by_strain"][strain] = cost_breakdown["by_strain"].get(strain, 0.0) + rate
    cost_breakdown["by_breeding"]["breeding" if c["has_breeders"] else "non_breeding"] += rate
    if c["n_mice"] == 1:
        cost_breakdown["by_housing"]["single"] += rate
    elif c["n_mice"] >= 2:
        cost_breakdown["by_housing"]["group"] += rate
    else:
        cost_breakdown["by_housing"]["empty"] += rate
    cost_breakdown["by_rack_status"]["located" if c["position"] else "unassigned"] += rate
    rm = c["room"] or "Unassigned room"
    cost_breakdown["by_room"][rm] = cost_breakdown["by_room"].get(rm, 0.0) + rate
    sx = c["sexes"]
    if not sx:
        cost_breakdown["by_sex_dominant"]["empty"] += rate
    elif all(s == "Female" for s in sx):
        cost_breakdown["by_sex_dominant"]["female"] += rate
    elif all(s == "Male" for s in sx):
        cost_breakdown["by_sex_dominant"]["male"] += rate
    else:
        cost_breakdown["by_sex_dominant"]["mixed"] += rate

    strain_lines.setdefault(strain, {"cages": 0, "mice": 0, "day": 0.0, "breeding_cages": 0})
    strain_lines[strain]["cages"] += 1
    strain_lines[strain]["mice"] += c["n_mice"]
    strain_lines[strain]["day"] += rate
    if c["has_breeders"]:
        strain_lines[strain]["breeding_cages"] += 1

    if c["genotypes"]:
        dom = max(set(c["genotypes"]), key=c["genotypes"].count)
    else:
        dom = "(empty)"
    k = (strain, dom)
    strain_geno_lines.setdefault(k, {"cages": 0, "mice": 0, "day": 0.0})
    strain_geno_lines[k]["cages"] += 1
    strain_geno_lines[k]["mice"] += c["n_mice"]
    strain_geno_lines[k]["day"] += rate

    cost_per_cage.append({
        "cage_id": c["cage_id"],
        "strain": strain,
        "n_mice": c["n_mice"],
        "has_breeders": c["has_breeders"],
        "rate_per_day": rate,
    })

# ---------- Post-cull projection ----------
# Recompute the 173 killed today using the same smart-cull logic the dashboard ships,
# then project costs assuming cages with zero surviving mice retire (standard practice).
import re as _re_post

_priority_post = {"AdipoGlo": "maintain", "AdipoGlo+": "maintain", "Marrow Glo": "expand",
                  "Adipoq-Cre": "winddown", "mTmG": "winddown", "Dendra2": "winddown"}
_KEEP_FLOOR_POST = 4
_DONOR_N_POST = 4

def _is_adipoq_donor(g):
    g = (g or "").lower().replace(" ", "")
    return "adipoq-cre" in g and "dendra2<-/->" in g and "mtmg<mtmg" not in g
def _is_dendra2_donor(g):
    g = (g or "").lower().replace(" ", "")
    return bool(_re_post.search(r"dendra2<\+/[+-]>", g)) and "adipoq-cre" not in g and "mtmg<mtmg" not in g
def _is_mtmg_donor(g):
    g = (g or "").lower().replace(" ", "")
    return "mtmg<mtmg/mtmg>" in g and "adipoq-cre" not in g and not _re_post.search(r"dendra2<\+/[+-]>", g)

_breeders_post = set()
for b in breeding_records:
    for k in ("mother", "father"):
        v = str(b.get(k, ""))
        s = v.split("|")[0].strip() if "|" in v else v.strip()
        if s:
            _breeders_post.add(s)
for r in inventory:
    if r["use"] == "Breeding":
        _breeders_post.add(str(r["mouse_id"]))

inv_df = pd.DataFrame(inventory)
inv_df["mouse_id"] = inv_df["mouse_id"].astype(str)
inv_df["genotype"] = inv_df["genotype"].fillna("").astype(str)

_rescue_donors = set()
for target, fn in [("Adipoq-Cre", _is_adipoq_donor), ("Dendra2", _is_dendra2_donor), ("mTmG", _is_mtmg_donor)]:
    if _priority_post.get(target) != "winddown":
        continue
    cands = inv_df[(inv_df["strain"] != target) & (inv_df["use"] == "Available") &
                   (~inv_df["mouse_id"].isin(_breeders_post)) & (inv_df["sex"] != "Unknown")].copy()
    cands = cands[cands["genotype"].apply(fn)]
    for sex in ("Female", "Male"):
        sub = cands[cands["sex"] == sex].dropna(subset=["age_months"]).sort_values("age_months").head(_DONOR_N_POST)
        for mid in sub["mouse_id"]:
            _rescue_donors.add(mid)

_inv_sorted = inv_df.sort_values("age_months", na_position="last")
_cell_rank = {}
for (_s, _g, _sx), grp in _inv_sorted.groupby(["strain", "genotype", "sex"], sort=False):
    for i, mid in enumerate(grp["mouse_id"]):
        _cell_rank[mid] = (i, len(grp))
_geno_size = inv_df.groupby(["strain", "genotype"]).size().to_dict()
_RESCUABLE_POST = {"Adipoq-Cre", "mTmG", "Dendra2"}

def _is_culled(m):
    mid = m["mouse_id"]
    if m["use"] == "Breeding" or mid in _breeders_post: return False
    if m["sex"] == "Unknown": return False
    if mid in _rescue_donors: return False
    if _geno_size.get((m["strain"], m["genotype"]), 0) == 1 and m["strain"] not in _RESCUABLE_POST: return False
    p = _priority_post.get(m["strain"], "maintain")
    if p == "expand": return False
    am = m.get("age_months")
    if p == "winddown": return not (pd.notna(am) and am < 6)
    idx, total = _cell_rank.get(mid, (999, 0))
    if idx < min(_KEEP_FLOOR_POST, total): return False
    if pd.notna(am) and am < 12: return False
    return True

_UNFOUND = {"0023","100","101","102","130","132","179","207","293","294","295",
            "48","68","85","87","92","141","166","225",
            "00001","00002","00008","00016","0417","202","204"}
inv_df["plan_cull"] = inv_df.apply(_is_culled, axis=1)
killed_ids = set(inv_df[(inv_df["plan_cull"]) & (~inv_df["mouse_id"].isin(_UNFOUND))]["mouse_id"])

# Cage-level survivors
cage_survivors = {}
for r in inventory:
    cid = str(r["cage_id"])
    cage_survivors.setdefault(cid, 0)
    if str(r["mouse_id"]) not in killed_ids:
        cage_survivors[cid] += 1

post_cull_day = 0.0
post_cull_active_cages = 0
post_cull_breakdown = {
    "by_strain": {},
    "by_breeding": {"breeding": 0.0, "non_breeding": 0.0},
    "by_housing": {"single": 0.0, "group": 0.0, "empty": 0.0},
}
for c in cage_records:
    survivors_n = cage_survivors.get(c["cage_id"], 0)
    if survivors_n == 0:
        continue  # cage retired
    rate = cage_rate(c)
    post_cull_day += rate
    post_cull_active_cages += 1
    strain = c["strain"] or "(blank)"
    post_cull_breakdown["by_strain"][strain] = post_cull_breakdown["by_strain"].get(strain, 0.0) + rate
    post_cull_breakdown["by_breeding"]["breeding" if c["has_breeders"] else "non_breeding"] += rate
    if survivors_n == 1:
        post_cull_breakdown["by_housing"]["single"] += rate
    else:
        post_cull_breakdown["by_housing"]["group"] += rate

cost_summary = {
    "rate_vent_standard": RATE_VENT_STD,
    "rate_vent_breeding": RATE_VENT_BREED,
    "rates_effective": "04/01/2026 (CY26)",
    "rate_source": "UMich ULAM rate schedule",
    "total_cages": len(cage_records),
    "per_day": round(cost_total_day, 2),
    "per_month": round(cost_total_day * 30, 2),
    "per_year": round(cost_total_day * 365, 2),
    "post_cull": {
        "killed_count": int(len(killed_ids)),
        "active_cages": post_cull_active_cages,
        "retired_cages": len(cage_records) - post_cull_active_cages,
        "per_day": round(post_cull_day, 2),
        "per_month": round(post_cull_day * 30, 2),
        "per_year": round(post_cull_day * 365, 2),
        "savings_per_day": round(cost_total_day - post_cull_day, 2),
        "savings_per_month": round((cost_total_day - post_cull_day) * 30, 2),
        "savings_per_year": round((cost_total_day - post_cull_day) * 365, 2),
        "savings_pct": round(100 * (cost_total_day - post_cull_day) / cost_total_day, 1) if cost_total_day else 0,
        "breakdown": {
            k: ({kk: round(vv, 2) for kk, vv in v.items()} if isinstance(v, dict) else round(v, 2))
            for k, v in post_cull_breakdown.items()
        },
    },
    "breakdown": {
        k: ({kk: round(vv, 2) for kk, vv in v.items()} if isinstance(v, dict) else round(v, 2))
        for k, v in cost_breakdown.items()
    },
    "by_strain_table": [
        {"strain": s, "cages": d["cages"], "mice": d["mice"],
         "breeding_cages": d["breeding_cages"],
         "day": round(d["day"], 2),
         "month": round(d["day"] * 30, 2),
         "year": round(d["day"] * 365, 2)}
        for s, d in sorted(strain_lines.items(), key=lambda x: -x[1]["day"])
    ],
    "by_strain_geno_table": [
        {"strain": s, "genotype": g, "cages": d["cages"], "mice": d["mice"],
         "day": round(d["day"], 2),
         "month": round(d["day"] * 30, 2),
         "year": round(d["day"] * 365, 2)}
        for (s, g), d in sorted(strain_geno_lines.items(), key=lambda x: -x[1]["day"])
    ],
}

# ---------- Final output ----------
out = {
    "generated_at": datetime.now().isoformat(timespec="seconds"),
    "as_of": TODAY.isoformat(),
    "summary": summary,
    "strain_order": strain_order,
    "strain_counts": strain_counts,
    "sex_by_strain": sex_by_strain,
    "age_buckets": age_buckets,
    "age_dist_total": age_dist_total,
    "age_dist_by_strain": age_dist_by_strain,
    "use_by_strain": use_by_strain,
    "geno_by_strain": geno_by_strain,
    "strain_summary": strain_summary,
    "breedings": breeding_records,
    "cohort_plans": cohort_plans,
    "cages": cage_records,
    "inventory": inventory,
    "cull_summary": cull_summary,
    "cost_summary": cost_summary,
}

with open(DATA / "colony.json", "w") as f:
    json.dump(out, f, indent=2, default=str)

print(f"Wrote {DATA / 'colony.json'}")
print(f"  total_mice={total_mice}, cages={total_cages}, breedings={len(breedings)}")
print(f"  located={located_cages}, unassigned={unassigned_cages}")
print(f"  cull aged_>12mo_available_nonbreeder: {cull_summary['aged_>12mo_available_nonbreeder']['count']}")
print(f"  cull aged_>18mo_any: {cull_summary['aged_>18mo_any']['count']}")
print(f"  cull sex_unknown: {cull_summary['sex_unknown']['count']}")
