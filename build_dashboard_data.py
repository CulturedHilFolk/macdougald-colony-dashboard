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
}

with open(DATA / "colony.json", "w") as f:
    json.dump(out, f, indent=2, default=str)

print(f"Wrote {DATA / 'colony.json'}")
print(f"  total_mice={total_mice}, cages={total_cages}, breedings={len(breedings)}")
print(f"  located={located_cages}, unassigned={unassigned_cages}")
print(f"  cull aged_>12mo_available_nonbreeder: {cull_summary['aged_>12mo_available_nonbreeder']['count']}")
print(f"  cull aged_>18mo_any: {cull_summary['aged_>18mo_any']['count']}")
print(f"  cull sex_unknown: {cull_summary['sex_unknown']['count']}")
