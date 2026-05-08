"""
Build index.html — a polished, self-contained colony dashboard.

Reads:  data/colony.json
Writes: index.html  (data embedded; single file)
"""
from pathlib import Path
import json

DIR = Path("/Users/briandes/Library/CloudStorage/OneDrive-MichiganMedicine/Desktop/MacDougald Lab/Mouse Colony")
DATA = DIR / "data" / "colony.json"
OUT = DIR / "index.html"

with open(DATA) as f:
    colony_json_str = f.read()

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1.0" />
<title>MacDougald Lab — Mouse Colony Dashboard</title>

<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

<style>
* { box-sizing: border-box; }
:root {
  --bg: #f7f8fa;
  --panel: #ffffff;
  --ink: #0f172a;
  --ink-muted: #475569;
  --ink-soft: #94a3b8;
  --line: #e5e9f0;
  --line-soft: #f1f5f9;
  --brand: #1e3a8a;
  --brand-soft: #e0e7ff;
  --accent: #2563eb;
  --ok: #15803d;
  --ok-soft: #dcfce7;
  --warn: #b45309;
  --warn-soft: #fef3c7;
  --danger: #b91c1c;
  --danger-soft: #fee2e2;
  --neutral: #64748b;
  --neutral-soft: #f1f5f9;
  --shadow: 0 1px 2px rgba(15,23,42,.04), 0 4px 12px rgba(15,23,42,.04);
  --shadow-lg: 0 4px 24px rgba(15,23,42,.08), 0 12px 40px rgba(15,23,42,.06);
  --radius: 10px;
  --radius-lg: 14px;
}

html, body {
  margin: 0; padding: 0; background: var(--bg); color: var(--ink);
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
  font-size: 14px; line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}

.app { min-height: 100vh; display: flex; flex-direction: column; }

/* ---- Header ---- */
.header {
  background: linear-gradient(135deg, #1e293b 0%, #1e3a8a 100%);
  color: #fff; padding: 18px 32px; box-shadow: 0 1px 0 rgba(0,0,0,.05);
}
.header-row { display: flex; align-items: center; justify-content: space-between; gap: 24px; flex-wrap: wrap; }
.header-title { display: flex; align-items: center; gap: 14px; }
.header-logo {
  width: 38px; height: 38px; border-radius: 9px;
  background: rgba(255,255,255,.12);
  display: flex; align-items: center; justify-content: center;
  font-size: 20px;
}
.header-title h1 { font-size: 1.1rem; font-weight: 700; letter-spacing: -0.01em; margin: 0; }
.header-title p { font-size: .78rem; opacity: .75; margin: 2px 0 0; }
.header-meta { display: flex; align-items: center; gap: 12px; font-size: .78rem; opacity: .85; }
.header-meta .pill {
  padding: 5px 12px; border-radius: 999px;
  background: rgba(255,255,255,.1); border: 1px solid rgba(255,255,255,.15);
}
.header-actions { display: flex; gap: 8px; }
.header-btn {
  background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.18);
  color: #fff; padding: 7px 14px; border-radius: 8px; font-size: .78rem; font-weight: 600;
  cursor: pointer; display: inline-flex; align-items: center; gap: 6px;
  transition: background .15s, border-color .15s;
}
.header-btn:hover { background: rgba(255,255,255,.18); border-color: rgba(255,255,255,.3); }

/* ---- Tab bar ---- */
.tabs {
  background: #fff; border-bottom: 1px solid var(--line);
  padding: 0 32px; display: flex; gap: 0;
  position: sticky; top: 0; z-index: 50;
  overflow-x: auto;
}
.tab {
  background: none; border: none; padding: 14px 18px;
  font-size: .85rem; font-weight: 600; color: var(--ink-muted);
  cursor: pointer; border-bottom: 2px solid transparent;
  transition: color .12s, border-color .12s;
  white-space: nowrap; font-family: inherit;
}
.tab:hover { color: var(--ink); }
.tab.active { color: var(--brand); border-bottom-color: var(--brand); }

/* ---- Main ---- */
.main { flex: 1; padding: 24px 32px; max-width: 1600px; width: 100%; margin: 0 auto; }
.tab-pane { display: none; }
.tab-pane.active { display: block; animation: fadein .25s; }
@keyframes fadein { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: none; } }

/* ---- Cards/grids ---- */
.kpi-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 14px; margin-bottom: 24px;
}
.kpi {
  background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
  padding: 16px 18px; box-shadow: var(--shadow);
}
.kpi-label { font-size: .72rem; font-weight: 600; color: var(--ink-soft); text-transform: uppercase; letter-spacing: .04em; }
.kpi-value { font-size: 1.7rem; font-weight: 700; color: var(--ink); margin-top: 4px; line-height: 1.1; }
.kpi-sub { font-size: .75rem; color: var(--ink-muted); margin-top: 4px; }
.kpi.alert .kpi-value { color: var(--danger); }
.kpi.warn .kpi-value { color: var(--warn); }
.kpi.ok .kpi-value { color: var(--ok); }

.row { display: grid; gap: 16px; margin-bottom: 16px; }
.row.cols-2 { grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); }
.row.cols-3 { grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); }

.panel {
  background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
  padding: 18px 20px; box-shadow: var(--shadow);
}
.panel h3 {
  margin: 0 0 14px; font-size: .9rem; font-weight: 700; color: var(--ink);
  display: flex; align-items: center; gap: 8px; justify-content: space-between;
}
.panel h3 .info-btn {
  background: var(--neutral-soft); border: none; cursor: pointer;
  width: 22px; height: 22px; border-radius: 50%;
  font-size: .72rem; font-weight: 700; color: var(--ink-muted);
  display: inline-flex; align-items: center; justify-content: center;
  transition: background .12s, color .12s;
}
.panel h3 .info-btn:hover { background: var(--brand-soft); color: var(--brand); }

.chart { width: 100%; height: 280px; }
.chart-tall { height: 360px; }
.chart-short { height: 220px; }

/* ---- Tables ---- */
.table-wrap {
  background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
  box-shadow: var(--shadow); overflow: hidden;
}
.table-controls {
  display: flex; gap: 10px; padding: 14px 16px;
  border-bottom: 1px solid var(--line); flex-wrap: wrap; align-items: center;
}
.table-controls input[type="text"], .table-controls select {
  padding: 8px 12px; border: 1px solid var(--line); border-radius: 8px;
  font-size: .82rem; font-family: inherit; background: #fff; color: var(--ink);
}
.table-controls input[type="text"] { min-width: 200px; }
.table-controls input[type="text"]:focus, .table-controls select:focus {
  outline: none; border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(37,99,235,.15);
}
.table-controls .count { margin-left: auto; font-size: .8rem; color: var(--ink-muted); font-weight: 500; }

.scroll-y { max-height: 600px; overflow-y: auto; }
.scroll-x { overflow-x: auto; }
table.data {
  width: 100%; border-collapse: collapse; font-size: .82rem;
}
table.data thead th {
  position: sticky; top: 0; z-index: 1;
  background: var(--neutral-soft); padding: 10px 12px;
  font-weight: 600; color: var(--ink-muted); text-align: left;
  border-bottom: 1px solid var(--line);
  text-transform: uppercase; letter-spacing: .03em; font-size: .72rem;
  cursor: pointer; user-select: none;
  white-space: nowrap;
}
table.data thead th:hover { background: #e5e9f0; color: var(--ink); }
table.data tbody td {
  padding: 9px 12px; border-bottom: 1px solid var(--line-soft); color: var(--ink);
  vertical-align: middle;
}
table.data tbody tr:hover { background: var(--neutral-soft); }
table.data tbody tr:last-child td { border-bottom: none; }
.id-mono { font-family: 'JetBrains Mono', monospace; font-size: .78rem; }

.tag {
  display: inline-block; padding: 2px 8px; border-radius: 999px;
  font-size: .7rem; font-weight: 600; line-height: 1.5;
}
.tag.brand { background: var(--brand-soft); color: var(--brand); }
.tag.ok { background: var(--ok-soft); color: var(--ok); }
.tag.warn { background: var(--warn-soft); color: var(--warn); }
.tag.danger { background: var(--danger-soft); color: var(--danger); }
.tag.neutral { background: var(--neutral-soft); color: var(--neutral); }

/* ---- Cull tab ---- */
.cull-list {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 14px; margin-bottom: 18px;
}
.cull-card {
  background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
  padding: 16px 18px; box-shadow: var(--shadow);
}
.cull-card.alert { border-left: 4px solid var(--danger); }
.cull-card.warn { border-left: 4px solid var(--warn); }
.cull-card.info { border-left: 4px solid var(--accent); }
.cull-card .count { font-size: 2rem; font-weight: 800; color: var(--ink); }
.cull-card .label { font-size: .82rem; color: var(--ink-muted); margin-top: 2px; }
.cull-card .desc { font-size: .75rem; color: var(--ink-soft); margin-top: 8px; }

/* ---- Cohort plans ---- */
.cohort-card {
  background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
  padding: 18px 20px; box-shadow: var(--shadow); margin-bottom: 14px;
}
.cohort-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 14px; margin-bottom: 8px; }
.cohort-head h4 { font-size: 1rem; font-weight: 700; margin: 0 0 2px; color: var(--brand); }
.cohort-cross { font-family: 'JetBrains Mono', monospace; font-size: .82rem; color: var(--ink-muted); }
.cohort-purpose { font-size: .85rem; color: var(--ink); margin: 6px 0 12px; padding-left: 12px; border-left: 3px solid var(--brand-soft); }
.cohort-target {
  background: var(--ok-soft); color: var(--ok); padding: 8px 12px; border-radius: 8px;
  font-size: .82rem; font-weight: 600; margin: 10px 0;
}
.cohort-ratios { display: flex; flex-direction: column; gap: 5px; }
.cohort-ratio {
  display: grid; grid-template-columns: 1fr 60px 90px; gap: 8px;
  align-items: center; padding: 5px 10px; background: var(--neutral-soft);
  border-radius: 6px; font-size: .8rem;
}
.cohort-ratio.target { background: var(--ok-soft); }
.cohort-ratio .pct { font-weight: 700; }
.cohort-ratio .bar {
  height: 6px; background: rgba(0,0,0,.06); border-radius: 3px; overflow: hidden;
}
.cohort-ratio .bar > span { height: 100%; display: block; background: var(--accent); }
.cohort-ratio.target .bar > span { background: var(--ok); }

/* ---- Cohort builder ---- */
.cb-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
@media (max-width: 980px) { .cb-grid { grid-template-columns: 1fr; } }
.cb-side {
  background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
  padding: 16px 18px; box-shadow: var(--shadow);
}
.cb-side h4 { margin: 0 0 10px; font-size: .95rem; font-weight: 700; color: var(--ink); display: flex; align-items: center; gap: 8px; justify-content: space-between; }
.cb-side h4 .badge { background: var(--brand); color: #fff; padding: 2px 9px; border-radius: 999px; font-size: .72rem; font-weight: 700; }
.cb-filter-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 8px; margin-bottom: 12px; }
.cb-filter-row input, .cb-filter-row select {
  padding: 7px 10px; border: 1px solid var(--line); border-radius: 7px;
  font-size: .8rem; font-family: inherit; background: #fff; color: var(--ink); width: 100%;
}
.cb-filter-row input:focus, .cb-filter-row select:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px rgba(37,99,235,.15); }
.cb-filter-row label { display: flex; align-items: center; gap: 6px; font-size: .78rem; color: var(--ink-muted); }
.cb-action-row { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 10px; align-items: center; }
.cb-table-wrap { max-height: 380px; overflow-y: auto; overflow-x: auto; border: 1px solid var(--line-soft); border-radius: 7px; }
.cb-empty {
  padding: 24px; text-align: center; color: var(--ink-soft); font-size: .85rem; font-style: italic;
}
.btn {
  background: var(--brand); color: #fff; border: none; padding: 8px 14px; border-radius: 7px;
  font-size: .8rem; font-weight: 600; cursor: pointer; font-family: inherit;
  display: inline-flex; align-items: center; gap: 6px; transition: background .12s;
}
.btn:hover { background: #1e40af; }
.btn.ghost { background: var(--neutral-soft); color: var(--ink); }
.btn.ghost:hover { background: #e5e9f0; }
.btn.ok { background: var(--ok); }
.btn.ok:hover { background: #166534; }
.btn.small { padding: 4px 9px; font-size: .72rem; }
.btn.danger { background: var(--danger); }
.btn.danger:hover { background: #991b1b; }
.cohort-summary {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(80px, 1fr));
  gap: 8px; padding: 10px; background: var(--neutral-soft); border-radius: 7px;
  margin-bottom: 10px;
}
.cohort-summary .item { text-align: center; }
.cohort-summary .item .v { font-size: 1.05rem; font-weight: 700; color: var(--ink); }
.cohort-summary .item .l { font-size: .68rem; color: var(--ink-muted); text-transform: uppercase; letter-spacing: .04em; }
.role-tag { display: inline-block; padding: 1px 7px; border-radius: 4px; font-size: .68rem; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; }
.role-tag.exp { background: var(--brand-soft); color: var(--brand); }
.role-tag.ctrl { background: var(--ok-soft); color: var(--ok); }
.control-suggestions {
  background: #fafbfd; border: 1px solid var(--line); border-radius: 7px;
  padding: 12px 14px; margin-top: 10px;
}
.control-suggestions h5 {
  margin: 0 0 8px; font-size: .82rem; color: var(--brand); font-weight: 700;
  display: flex; justify-content: space-between; align-items: center;
}
.control-suggestions h5 .target {
  font-family: 'JetBrains Mono', monospace; font-size: .72rem; color: var(--ink-muted); font-weight: 500;
}
.control-suggestions table { width: 100%; font-size: .76rem; }
.control-suggestions table td { padding: 5px 8px; border-bottom: 1px solid var(--line-soft); }
.age-diff { font-size: .68rem; color: var(--ink-soft); margin-left: 4px; }
.age-diff.close { color: var(--ok); font-weight: 600; }

/* ---- Modal ---- */
.modal-bg {
  position: fixed; inset: 0; background: rgba(15,23,42,.6); z-index: 100;
  display: none; align-items: center; justify-content: center;
  padding: 24px; overflow-y: auto;
}
.modal-bg.show { display: flex; }
.modal {
  background: #fff; border-radius: var(--radius-lg); max-width: 720px;
  width: 100%; padding: 28px 32px; box-shadow: var(--shadow-lg);
  max-height: 90vh; overflow-y: auto;
}
.modal h2 { margin: 0 0 14px; font-size: 1.2rem; color: var(--ink); }
.modal h3 { margin: 18px 0 6px; font-size: .92rem; color: var(--brand); }
.modal p, .modal li { font-size: .88rem; color: var(--ink); margin: 6px 0; }
.modal ul { padding-left: 20px; margin: 6px 0; }
.modal-close {
  background: var(--neutral-soft); border: none; padding: 6px 14px; border-radius: 8px;
  font-weight: 600; cursor: pointer; font-size: .82rem; color: var(--ink-muted);
  margin-top: 16px;
}
.modal-close:hover { background: #e5e9f0; color: var(--ink); }

/* ---- Cage grid ---- */
.cage-grid {
  display: grid; grid-template-columns: 40px repeat(10, 1fr);
  gap: 4px; margin-top: 14px;
}
.cage-grid .corner { background: transparent; }
.cage-grid .col-head, .cage-grid .row-head {
  background: var(--neutral-soft); color: var(--ink-muted);
  font-size: .72rem; font-weight: 700; text-align: center; padding: 5px 0;
  border-radius: 4px;
}
.cage-cell {
  border: 1px solid var(--line); border-radius: 5px; min-height: 60px;
  padding: 5px 6px; font-size: .68rem; background: #fff; cursor: pointer;
  transition: box-shadow .12s, transform .12s, border-color .12s;
  position: relative;
}
.cage-cell.empty { background: var(--neutral-soft); border-style: dashed; }
.cage-cell:hover { box-shadow: var(--shadow); transform: translateY(-1px); border-color: var(--accent); }
.cage-cell .pos { font-weight: 700; color: var(--brand); font-size: .68rem; }
.cage-cell .cid { font-family: 'JetBrains Mono', monospace; font-size: .62rem; color: var(--ink-muted); margin-top: 1px; }
.cage-cell .n { font-size: .82rem; font-weight: 700; margin-top: 2px; color: var(--ink); }
.cage-cell .strain { font-size: .62rem; color: var(--ink-soft); }
.cage-cell.has-breeders { border-left: 3px solid var(--warn); }

.cage-list {
  background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
  box-shadow: var(--shadow); padding: 16px 18px; margin-top: 16px;
}

/* ---- Misc ---- */
.section-head { font-size: 1.05rem; font-weight: 700; color: var(--ink); margin: 6px 0 14px; }
.muted { color: var(--ink-muted); }
.divider { border: none; border-top: 1px solid var(--line); margin: 18px 0; }
.bullet-list { padding-left: 20px; margin: 8px 0; }
.bullet-list li { margin: 4px 0; font-size: .85rem; }
.code { font-family: 'JetBrains Mono', monospace; font-size: .82rem; background: var(--neutral-soft); padding: 1px 6px; border-radius: 4px; }

@media (max-width: 720px) {
  .header, .tabs, .main { padding-left: 16px; padding-right: 16px; }
  .header-row { flex-direction: column; align-items: flex-start; }
  .cage-grid { grid-template-columns: 30px repeat(10, minmax(50px, 1fr)); overflow-x: auto; }
}
</style>
</head>
<body>

<div class="app">

<!-- HEADER -->
<header class="header">
  <div class="header-row">
    <div class="header-title">
      <div class="header-logo">🧬</div>
      <div>
        <h1>MacDougald Lab — Mouse Colony</h1>
        <p>Inventory · Breeding · Cohort Planning · Cull Decisions</p>
      </div>
    </div>
    <div class="header-meta">
      <span class="pill" id="hdr-asof">As of —</span>
      <span class="pill" id="hdr-mice">— mice</span>
      <span class="pill" id="hdr-cages">— cages</span>
    </div>
    <div class="header-actions">
      <button class="header-btn" onclick="openModal('info')">ⓘ Info</button>
      <button class="header-btn" onclick="openModal('summary')">⚡ Summary</button>
      <button class="header-btn" onclick="openModal('update')">↻ How to Update</button>
    </div>
  </div>
</header>

<!-- TABS -->
<nav class="tabs">
  <button class="tab active" data-pane="overview">Overview</button>
  <button class="tab" data-pane="inventory">Inventory</button>
  <button class="tab" data-pane="cages">Cages</button>
  <button class="tab" data-pane="breeding">Breeding</button>
  <button class="tab" data-pane="cohort">Cohort Planning</button>
  <button class="tab" data-pane="cull">Cull Candidates</button>
  <button class="tab" data-pane="smartcull">Smart Cull Plan</button>
  <button class="tab" data-pane="walkthrough">Walkthrough</button>
  <button class="tab" data-pane="experiment">Experiment Details</button>
  <button class="tab" data-pane="stats">Statistics</button>
</nav>

<main class="main">

<!-- ============= OVERVIEW ============= -->
<section class="tab-pane active" id="pane-overview">
  <div class="kpi-grid" id="kpi-grid"></div>
  <div class="row cols-2">
    <div class="panel">
      <h3>Mice by strain
        <button class="info-btn" onclick="openModal('chart-strain')">i</button>
      </h3>
      <div id="chart-strain" class="chart"></div>
    </div>
    <div class="panel">
      <h3>Age distribution (all strains)
        <button class="info-btn" onclick="openModal('chart-age')">i</button>
      </h3>
      <div id="chart-age" class="chart"></div>
    </div>
  </div>
  <div class="row cols-2">
    <div class="panel">
      <h3>Sex by strain
        <button class="info-btn" onclick="openModal('chart-sex')">i</button>
      </h3>
      <div id="chart-sex" class="chart"></div>
    </div>
    <div class="panel">
      <h3>Use status by strain
        <button class="info-btn" onclick="openModal('chart-use')">i</button>
      </h3>
      <div id="chart-use" class="chart"></div>
    </div>
  </div>
  <div class="row">
    <div class="panel">
      <h3>Strain summary
        <button class="info-btn" onclick="openModal('strain-summary')">i</button>
      </h3>
      <div class="scroll-x">
        <table class="data" id="strain-summary-table">
          <thead><tr>
            <th>Strain</th><th>Active Breedings</th><th>Active Cages</th><th>Active Litters</th>
            <th>Alive</th><th>Dead</th><th>Mutations</th>
          </tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>
</section>

<!-- ============= INVENTORY ============= -->
<section class="tab-pane" id="pane-inventory">
  <div class="table-wrap">
    <div class="table-controls">
      <input type="text" id="inv-search" placeholder="Search ID, genotype, cage…" />
      <select id="inv-strain"><option value="">All strains</option></select>
      <select id="inv-sex">
        <option value="">All sexes</option>
        <option value="Female">Female</option>
        <option value="Male">Male</option>
        <option value="Unknown">Unknown</option>
      </select>
      <select id="inv-use">
        <option value="">All use</option>
        <option value="Available">Available</option>
        <option value="Breeding">Breeding</option>
      </select>
      <select id="inv-age">
        <option value="">All ages</option>
        <option value="0-2">0–2 mo</option>
        <option value="2-4">2–4 mo</option>
        <option value="4-6">4–6 mo</option>
        <option value="6-12">6–12 mo</option>
        <option value="12+">>12 mo</option>
      </select>
      <span class="count" id="inv-count">— mice</span>
    </div>
    <div class="scroll-y scroll-x">
      <table class="data" id="inv-table">
        <thead><tr>
          <th data-sort="mouse_id">ID</th>
          <th data-sort="strain">Strain</th>
          <th data-sort="sex">Sex</th>
          <th data-sort="age_months">Age (mo)</th>
          <th data-sort="dob">DOB</th>
          <th data-sort="genotype">Genotype</th>
          <th data-sort="cage_id">Cage</th>
          <th data-sort="use">Use</th>
          <th>Flags</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</section>

<!-- ============= CAGES ============= -->
<section class="tab-pane" id="pane-cages">
  <div class="panel">
    <h3>Rack 7A (Room 1) — physical layout
      <button class="info-btn" onclick="openModal('cage-grid')">i</button>
    </h3>
    <div class="cage-grid" id="cage-grid"></div>
    <p class="muted" style="margin-top: 12px; font-size: .82rem;">
      Click any cage for details. Orange left border = breeders inside. Gray dashed = empty slot.
    </p>
  </div>
  <div class="cage-list">
    <h3>Unassigned cages — physical location unknown
      <button class="info-btn" onclick="openModal('unassigned')">i</button>
    </h3>
    <div class="scroll-y scroll-x" style="max-height: 500px;">
      <table class="data" id="unassigned-table">
        <thead><tr>
          <th>Cage #</th><th>Strain</th><th>n Mice</th><th>Has Breeders</th><th>Mouse IDs</th><th>Created</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</section>

<!-- ============= BREEDING ============= -->
<section class="tab-pane" id="pane-breeding">
  <div class="row cols-3" id="breeding-kpis"></div>
  <div class="row cols-2">
    <div class="panel">
      <h3>Active breeding pairs — performance
        <button class="info-btn" onclick="openModal('breeding-perf')">i</button>
      </h3>
      <div id="chart-breeding-perf" class="chart"></div>
    </div>
    <div class="panel">
      <h3>Months active by pair
        <button class="info-btn" onclick="openModal('breeding-age')">i</button>
      </h3>
      <div id="chart-breeding-age" class="chart"></div>
    </div>
  </div>
  <div class="table-wrap">
    <div class="table-controls">
      <span class="count" id="breeding-count">— pairs</span>
    </div>
    <div class="scroll-y scroll-x">
      <table class="data" id="breeding-table">
        <thead><tr>
          <th>Pair</th><th>Status</th><th>Strain</th><th>Mother</th><th>Father</th>
          <th>Mother Geno</th><th>Father Geno</th>
          <th>Active Since</th><th>Mo Active</th><th>Litters</th><th>Born</th><th>Avg Litter</th><th>Last Weaned</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</section>

<!-- ============= COHORT PLANNING ============= -->
<section class="tab-pane" id="pane-cohort">
  <div class="panel" style="margin-bottom: 18px;">
    <h3>What is this tab?</h3>
    <p class="muted" style="font-size: .88rem;">
      Each card describes an active research cohort: the cross used to generate it, why we need it,
      the target genotype, and the expected Mendelian ratios per cross. Use this to plan how many
      breedings you need to produce the experimental n.
    </p>
  </div>
  <div id="cohort-list"></div>

  <!-- ===== Cohort Builder ===== -->
  <div class="panel" style="margin-top: 24px;">
    <h3>Build a cohort
      <button class="info-btn" onclick="openModal('cohort-builder')">i</button>
    </h3>
    <p class="muted" style="font-size: .85rem; margin: 0 0 16px;">
      Filter the live colony, hand-pick experimental mice, surface genetic controls
      best matched by sex and age, and export the assembled cohort to Excel.
    </p>

    <div class="cb-grid">
      <!-- LEFT: filter + candidates -->
      <div class="cb-side">
        <h4>1 · Available mice <span class="badge" id="cb-cand-count">0</span></h4>
        <div class="cb-filter-row">
          <select id="cb-strain"><option value="">All strains</option></select>
          <select id="cb-sex">
            <option value="">All sexes</option>
            <option value="Female">Female</option>
            <option value="Male">Male</option>
          </select>
          <select id="cb-genotype"><option value="">All genotypes</option></select>
          <input id="cb-age-min" type="number" min="0" step="0.1" placeholder="Min age (mo)" />
          <input id="cb-age-max" type="number" min="0" step="0.1" placeholder="Max age (mo)" />
          <select id="cb-use">
            <option value="Available">Only Available</option>
            <option value="">Any use status</option>
            <option value="Breeding">Only Breeding</option>
          </select>
        </div>
        <div class="cb-table-wrap">
          <table class="data" id="cb-cand-table">
            <thead><tr>
              <th></th><th>ID</th><th>Strain</th><th>Sex</th><th>Age</th>
              <th>Genotype</th><th>Cage</th>
            </tr></thead>
            <tbody></tbody>
          </table>
        </div>
        <div class="cb-action-row">
          <button class="btn small ghost" onclick="cbAddAllVisible('experimental')">+ Add all visible as experimental</button>
          <button class="btn small ghost" onclick="cbAddAllVisible('control')">+ Add all visible as control</button>
        </div>
      </div>

      <!-- RIGHT: cohort + controls + export -->
      <div class="cb-side">
        <h4>2 · My cohort <span class="badge" id="cb-cohort-count">0</span></h4>
        <div class="cohort-summary" id="cb-cohort-summary">
          <div class="item"><div class="v">0</div><div class="l">Total</div></div>
          <div class="item"><div class="v">0</div><div class="l">Experimental</div></div>
          <div class="item"><div class="v">0</div><div class="l">Controls</div></div>
          <div class="item"><div class="v">—</div><div class="l">F : M</div></div>
          <div class="item"><div class="v">—</div><div class="l">Mean age (mo)</div></div>
        </div>
        <div class="cb-table-wrap">
          <table class="data" id="cb-cohort-table">
            <thead><tr>
              <th>Role</th><th>ID</th><th>Strain</th><th>Sex</th><th>Age</th>
              <th>Genotype</th><th>Cage</th><th></th>
            </tr></thead>
            <tbody></tbody>
          </table>
        </div>
        <div class="cb-action-row">
          <button class="btn" onclick="cbSuggestControls()">🔍 Suggest controls</button>
          <button class="btn ok" onclick="cbExportXlsx()">⬇ Export to Excel</button>
          <button class="btn ghost" onclick="cbClear()">Clear cohort</button>
        </div>
        <div id="cb-control-area"></div>
      </div>
    </div>
  </div>
</section>

<!-- ============= CULL CANDIDATES ============= -->
<section class="tab-pane" id="pane-cull">
  <div class="panel" style="margin-bottom: 18px;">
    <h3>How to use this tab</h3>
    <p class="muted" style="font-size: .88rem;">
      The cards below show pre-defined cull queries. Click any card to filter the table.
      The table is sortable and exportable to CSV. Always verify against the live cage tags
      before euthanizing. Breeders and mice on active projects are <strong>excluded</strong>
      from "available, aged" by default.
    </p>
  </div>
  <div class="cull-list" id="cull-cards"></div>
  <div class="table-wrap">
    <div class="table-controls">
      <select id="cull-filter">
        <option value="all">Show all candidates</option>
      </select>
      <button class="header-btn" style="background: var(--brand); border-color: var(--brand);"
              onclick="exportCullCSV()">⬇ Export CSV</button>
      <span class="count" id="cull-count">— candidates</span>
    </div>
    <div class="scroll-y scroll-x">
      <table class="data" id="cull-table">
        <thead><tr>
          <th><input type="checkbox" id="cull-all" /></th>
          <th>ID</th><th>Strain</th><th>Sex</th><th>Age (mo)</th>
          <th>Genotype</th><th>Cage</th><th>Use</th><th>Flags</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</section>

<!-- ============= SMART CULL PLAN ============= -->
<section class="tab-pane" id="pane-smartcull">
  <div class="panel" style="margin-bottom: 16px;">
    <h3>What this tab does
      <button class="info-btn" onclick="openModal('smart-cull')">i</button>
    </h3>
    <p style="font-size: .9rem; margin: 0 0 6px;">
      Generates a preserve / cull recommendation by combining (a) per-strain priority,
      (b) cohort-rebuild insurance per (strain × genotype × sex) cell, and
      (c) automatic rescue-donor preservation for orphan strains that can be regenerated
      from compound strains.
    </p>
    <p class="muted" style="font-size: .82rem; margin: 6px 0 0;">
      Defaults: <strong>Expand → Marrow Glo</strong> ·
      <strong>Wind down → Adipoq-Cre, mTmG, Dendra2</strong> (rescuable from AdipoGlo / AdipoGlo+) ·
      <strong>Maintain → AdipoGlo, AdipoGlo+</strong>. Adjust below.
    </p>
  </div>

  <!-- Strain priorities -->
  <div class="panel" style="margin-bottom: 16px;">
    <h3>1 · Strain priorities</h3>
    <div id="sc-priority-grid"></div>
    <div class="cb-action-row" style="margin-top: 14px;">
      <label style="display:flex;align-items:center;gap:8px;font-size:.85rem;">
        Keep-floor (≥ youngest mice per strain × genotype × sex cell):
        <input id="sc-keepfloor" type="number" min="1" max="20" value="4" style="width:60px;padding:5px 8px;border:1px solid var(--line);border-radius:6px;font-family:inherit;">
      </label>
      <label style="display:flex;align-items:center;gap:6px;font-size:.85rem;">
        <input id="sc-rescue-on" type="checkbox" checked> Auto-preserve rescue donors
      </label>
      <label style="display:flex;align-items:center;gap:6px;font-size:.85rem;">
        Donors per sex per orphan strain:
        <input id="sc-donor-n" type="number" min="1" max="10" value="4" style="width:50px;padding:5px 8px;border:1px solid var(--line);border-radius:6px;font-family:inherit;">
      </label>
      <button class="btn" onclick="scRecompute()">↻ Recompute plan</button>
    </div>
  </div>

  <!-- Plan summary -->
  <div class="row cols-3" id="sc-summary" style="margin-bottom: 14px;"></div>

  <!-- Recommended actions -->
  <div class="panel" id="sc-actions" style="margin-bottom: 14px; display:none;">
    <h3>2 · Recommended actions before culling</h3>
    <div id="sc-actions-body"></div>
  </div>

  <!-- Cull list -->
  <div class="table-wrap" style="margin-bottom: 14px;">
    <div class="table-controls">
      <strong style="font-size:.92rem; color: var(--ink);">3 · Cull list</strong>
      <button class="btn ok" style="padding:6px 12px;font-size:.78rem" onclick="scExportCull()">⬇ Export to Excel</button>
      <button class="btn ghost" style="padding:6px 12px;font-size:.78rem" onclick="scExportCullCSV()">⬇ CSV</button>
      <span class="count" id="sc-cull-count">— mice</span>
    </div>
    <div class="scroll-y scroll-x" style="max-height: 500px;">
      <table class="data" id="sc-cull-table">
        <thead><tr>
          <th>ID</th><th>Strain</th><th>Sex</th><th>Age</th><th>Genotype</th>
          <th>Cage</th><th>Use</th><th>Reason</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <!-- Preserve list -->
  <details class="table-wrap" style="margin-bottom: 14px;">
    <summary style="padding: 14px 16px; cursor: pointer; font-weight: 700; font-size: .92rem; color: var(--ink); display: flex; align-items: center; gap: 10px;">
      4 · Preserved (do NOT cull)
      <span class="count" id="sc-preserve-count" style="font-weight: 500; color: var(--ink-muted);">— mice</span>
    </summary>
    <div class="scroll-y scroll-x" style="max-height: 500px; border-top: 1px solid var(--line);">
      <table class="data" id="sc-preserve-table">
        <thead><tr>
          <th>ID</th><th>Strain</th><th>Sex</th><th>Age</th><th>Genotype</th>
          <th>Cage</th><th>Use</th><th>Reason(s)</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </details>
</section>

<!-- ============= WALKTHROUGH ============= -->
<section class="tab-pane" id="pane-walkthrough">
  <div class="panel">
    <h3>Physical walkthrough reconciliation</h3>
    <p>
      This tab is a placeholder until the next physical walkthrough is logged.
      The lab uses a printed rack template (Rack_Template_<em>YYYY-MM-DD</em>.xlsx) to record
      observed cages, mouse counts, and IDs. After the walkthrough is digitized, this tab will show:
    </p>
    <ul class="bullet-list">
      <li><strong>Discrepancies</strong> — observed vs. recorded counts; missing or extra mice</li>
      <li><strong>Newly located</strong> — cages found on a rack that were "Unassigned" in the colony software</li>
      <li><strong>Decommissioned</strong> — cages that no longer exist in the room</li>
      <li><strong>Sex / ID corrections</strong> — observed sex differs from records, or ear tag missing</li>
    </ul>
    <hr class="divider" />
    <h3>Last completed walkthrough</h3>
    <p class="muted" id="walkthrough-status">No completed walkthrough on record.</p>
  </div>
</section>

<!-- ============= EXPERIMENT DETAILS ============= -->
<section class="tab-pane" id="pane-experiment">
  <div class="panel" style="margin-bottom: 14px;">
    <h3>Colony purpose</h3>
    <p>
      The MacDougald Lab maintains this colony to study adipocyte biology and mitochondrial
      transfer in adipose tissue. The active reporter and Cre lines are used to:
    </p>
    <ul class="bullet-list">
      <li>Track adipocyte mitochondrial dynamics in vivo (Dendra2 photoconversion)</li>
      <li>Mark and lineage-trace adipocytes specifically (Adipoq-Cre)</li>
      <li>Visualize membrane / cellular contributions across donor-host pairs (mTmG)</li>
      <li>Generate triple-reporter (mTmG; Adipoq-Cre; Dendra2) cohorts for combined imaging</li>
    </ul>
  </div>
  <div class="panel" style="margin-bottom: 14px;">
    <h3>Strains in colony</h3>
    <ul class="bullet-list" id="exp-strain-list"></ul>
  </div>
  <div class="panel" style="margin-bottom: 14px;">
    <h3>Standard practices</h3>
    <ul class="bullet-list">
      <li><strong>Genotyping:</strong> tail-snip → 96-well plate → Transnetyx dropbox → genotypes returned in ~3 days. Mice are genotyped at weaning, and again if an ear tag is lost.</li>
      <li><strong>Husbandry:</strong> ULAM-managed barrier facility. Standard chow unless otherwise noted.</li>
      <li><strong>Aging cohorts:</strong> mice are kept past 12 months only when aging is part of the experimental design. Otherwise, available mice over 12 months are flagged for cull.</li>
      <li><strong>Records of truth:</strong> primary inventory lives in the lab's colony software (the Excel exports under <span class="code">CageListExcel.xlsx</span> and <span class="code">MacLab - Brian_Mice.xlsx</span>). Periodic physical walkthroughs reconcile the records.</li>
    </ul>
  </div>
  <div class="panel">
    <h3>Open data gaps (please flag if you know the answer)</h3>
    <ul class="bullet-list" id="exp-gaps">
      <li>Specific room / building location of "Unassigned" cages — likely on additional racks not yet entered into the colony software.</li>
      <li>Sex of 5 mice currently recorded as "Unknown" — to be determined at next walkthrough.</li>
      <li>Whether any of the Adipoq-Cre-only mice (n=1) are still actively used vs. residual.</li>
    </ul>
  </div>
</section>

<!-- ============= STATISTICS ============= -->
<section class="tab-pane" id="pane-stats">
  <div class="panel" style="margin-bottom: 14px;">
    <h3>What this tab covers</h3>
    <p class="muted" style="font-size: .88rem;">
      This is a colony-management dashboard, so the "statistics" here are <em>descriptive</em> —
      demographic breakdowns, distributions, and breeding-performance metrics. No hypothesis tests
      are run on the colony itself. (Inferential statistics are reserved for downstream
      assays — e.g. Seahorse, microscopy, flow — and live in those project-specific dashboards.)
    </p>
  </div>
  <div class="row cols-2">
    <div class="panel">
      <h3>Age statistics by strain</h3>
      <div class="scroll-x">
        <table class="data" id="age-stats-table">
          <thead><tr>
            <th>Strain</th><th>n</th><th>Median (mo)</th><th>Mean (mo)</th><th>Min</th><th>Max</th>
          </tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
    <div class="panel">
      <h3>Sex ratio by strain</h3>
      <div class="scroll-x">
        <table class="data" id="sex-ratio-table">
          <thead><tr>
            <th>Strain</th><th>F</th><th>M</th><th>F:M ratio</th><th>% Female</th>
          </tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>
  <div class="row">
    <div class="panel">
      <h3>Breeding performance</h3>
      <div class="scroll-x">
        <table class="data" id="breeding-stats-table">
          <thead><tr>
            <th>Pair</th><th>Strain</th><th>Litters</th><th>Born</th>
            <th>Avg litter size</th><th>Months active</th><th>Productivity (born/mo)</th>
          </tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>
  <div class="panel" style="margin-top: 16px;">
    <h3>Methodology notes</h3>
    <ul class="bullet-list">
      <li>Age computed as <span class="code">today - DOB</span> in days; rendered as months (days / 30).</li>
      <li>Cull queries are <strong>filters</strong>, not predictions — they surface candidates that meet the rule, but final cull decisions are made by the experimenter.</li>
      <li>Avg litter size = (total born / total litters); excludes pairs with zero litters from the average.</li>
      <li>Productivity (born/mo) = born / months_active; for active pairs only.</li>
    </ul>
  </div>
</section>

</main>

</div><!-- /.app -->

<!-- MODAL -->
<div class="modal-bg" id="modal-bg" onclick="closeModalIfBg(event)">
  <div class="modal" id="modal">
    <h2 id="modal-title"></h2>
    <div id="modal-body"></div>
    <button class="modal-close" onclick="closeModal()">Close</button>
  </div>
</div>

<script>
// ============= DATA =============
const COLONY = __COLONY_JSON__;

// ============= UTILS =============
const $ = (s, root=document) => root.querySelector(s);
const $$ = (s, root=document) => Array.from(root.querySelectorAll(s));

const STRAIN_COLORS = {
  "AdipoGlo": "#1e40af",
  "AdipoGlo+": "#7c3aed",
  "Dendra2": "#0d9488",
  "Marrow Glo": "#ca8a04",
  "mTmG": "#db2777",
  "Adipoq-Cre": "#64748b",
};
const STRAIN_COLOR = s => STRAIN_COLORS[s] || "#64748b";
const SEX_COLOR = { "Female": "#db2777", "Male": "#2563eb", "Unknown": "#64748b" };
const USE_COLOR = { "Available": "#15803d", "Breeding": "#b45309" };
const AGE_COLOR = {
  "0–2 mo": "#16a34a", "2–4 mo": "#22c55e", "4–6 mo": "#84cc16",
  "6–12 mo": "#eab308", ">12 mo": "#dc2626", "Unknown": "#94a3b8"
};

function fmt(n) { return new Intl.NumberFormat().format(n); }

function tagFor(value, type) {
  if (type === "use") {
    if (value === "Breeding") return `<span class="tag warn">Breeding</span>`;
    return `<span class="tag ok">Available</span>`;
  }
  if (type === "sex") {
    const cls = value === "Female" ? "tag" : value === "Male" ? "tag brand" : "tag neutral";
    return `<span class="${cls}" style="${value==='Female'?'background:#fce7f3;color:#9d174d':''}">${value}</span>`;
  }
  if (type === "strain") {
    const c = STRAIN_COLOR(value);
    return `<span class="tag" style="background:${c}20; color:${c}">${value}</span>`;
  }
  return value;
}

function flagsTags(flags) {
  if (!flags || !flags.length) return "";
  return flags.map(f => {
    if (f.includes("18mo")) return `<span class="tag danger">>18mo</span>`;
    if (f.includes("12mo")) return `<span class="tag warn">>12mo</span>`;
    if (f.includes("sex_unknown")) return `<span class="tag neutral">sex?</span>`;
    return `<span class="tag neutral">${f}</span>`;
  }).join(" ");
}

// ============= TABS =============
$$(".tab").forEach(btn => {
  btn.addEventListener("click", () => {
    $$(".tab").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    $$(".tab-pane").forEach(p => p.classList.remove("active"));
    $("#pane-" + btn.dataset.pane).classList.add("active");
    // Force chart resize on tab show
    setTimeout(() => window.dispatchEvent(new Event("resize")), 50);
  });
});

// ============= HEADER =============
$("#hdr-asof").textContent = "As of " + COLONY.as_of;
$("#hdr-mice").textContent = fmt(COLONY.summary.total_mice) + " mice";
$("#hdr-cages").textContent = fmt(COLONY.summary.total_cages) + " cages";

// ============= KPIs =============
const KPIS = [
  { label: "Total mice", val: COLONY.summary.total_mice, sub: COLONY.summary.n_strains + " strains" },
  { label: "Total cages", val: COLONY.summary.total_cages,
    sub: COLONY.summary.located_cages + " located · " + COLONY.summary.unassigned_cages + " unassigned",
    cls: COLONY.summary.unassigned_cages > 50 ? "warn" : "" },
  { label: "Active breeders (mice)", val: COLONY.summary.active_breeders,
    sub: COLONY.summary.active_breeding_pairs + " pairs" },
  { label: "Available", val: COLONY.summary.available_mice, sub: "no project assignment", cls: "ok" },
  { label: "Aged > 12 mo", val: COLONY.summary.aged_over_12mo,
    sub: COLONY.summary.aged_over_18mo + " over 18 mo", cls: "warn" },
  { label: "Sex unknown", val: COLONY.summary.unknown_sex, sub: "ID at next walkthrough",
    cls: COLONY.summary.unknown_sex > 0 ? "alert" : "ok" },
];
$("#kpi-grid").innerHTML = KPIS.map(k => `
  <div class="kpi ${k.cls || ''}">
    <div class="kpi-label">${k.label}</div>
    <div class="kpi-value">${fmt(k.val)}</div>
    <div class="kpi-sub">${k.sub}</div>
  </div>
`).join("");

// ============= CHARTS =============
const charts = {};

// Strain bar
charts.strain = echarts.init($("#chart-strain"));
charts.strain.setOption({
  grid: { left: 10, right: 24, top: 10, bottom: 30, containLabel: true },
  xAxis: { type: "category", data: COLONY.strain_order, axisLine: { lineStyle: { color: "#cbd5e1" } } },
  yAxis: { type: "value", splitLine: { lineStyle: { color: "#f1f5f9" } } },
  tooltip: { trigger: "axis" },
  series: [{
    type: "bar", barMaxWidth: 60,
    data: COLONY.strain_order.map(s => ({ value: COLONY.strain_counts[s], itemStyle: { color: STRAIN_COLOR(s) } })),
    label: { show: true, position: "top", color: "#0f172a", fontWeight: 600 }
  }]
});

// Age dist pie
charts.age = echarts.init($("#chart-age"));
charts.age.setOption({
  tooltip: { trigger: "item" },
  legend: { orient: "vertical", right: 10, top: "middle", textStyle: { fontSize: 11 } },
  series: [{
    type: "pie", radius: ["45%", "75%"], avoidLabelOverlap: true,
    center: ["38%", "50%"],
    label: { show: false },
    data: COLONY.age_buckets.map(b => ({
      name: b, value: COLONY.age_dist_total[b], itemStyle: { color: AGE_COLOR[b] }
    })).filter(d => d.value > 0)
  }]
});

// Sex by strain stacked bar
charts.sex = echarts.init($("#chart-sex"));
charts.sex.setOption({
  grid: { left: 10, right: 24, top: 30, bottom: 30, containLabel: true },
  xAxis: { type: "category", data: COLONY.strain_order },
  yAxis: { type: "value" },
  tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
  legend: { top: 0, textStyle: { fontSize: 11 } },
  series: ["Female", "Male", "Unknown"].map(sex => ({
    name: sex, type: "bar", stack: "total",
    itemStyle: { color: SEX_COLOR[sex] },
    data: COLONY.strain_order.map(s => COLONY.sex_by_strain[s][sex] || 0)
  }))
});

// Use status by strain stacked bar
charts.use = echarts.init($("#chart-use"));
charts.use.setOption({
  grid: { left: 10, right: 24, top: 30, bottom: 30, containLabel: true },
  xAxis: { type: "category", data: COLONY.strain_order },
  yAxis: { type: "value" },
  tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
  legend: { top: 0, textStyle: { fontSize: 11 } },
  series: ["Available", "Breeding"].map(use => ({
    name: use, type: "bar", stack: "total",
    itemStyle: { color: USE_COLOR[use] },
    data: COLONY.strain_order.map(s => COLONY.use_by_strain[s][use] || 0)
  }))
});

// Strain summary table
$("#strain-summary-table tbody").innerHTML = COLONY.strain_summary.map(s => `
  <tr>
    <td>${tagFor(s.strain, "strain")}</td>
    <td>${s.active_breedings}</td>
    <td>${s.active_cages}</td>
    <td>${s.active_litters}</td>
    <td>${s.alive_mice}</td>
    <td>${s.dead_mice}</td>
    <td>${s.n_mutations}</td>
  </tr>
`).join("");

// ============= INVENTORY =============
const INV = COLONY.inventory.slice();

function renderInventory(rows) {
  const html = rows.map(m => `
    <tr>
      <td class="id-mono">${m.mouse_id}</td>
      <td>${tagFor(m.strain, "strain")}</td>
      <td>${tagFor(m.sex, "sex")}</td>
      <td>${m.age_months ?? "—"}</td>
      <td>${m.dob || "—"}</td>
      <td class="id-mono" style="font-size:.75rem">${m.genotype || ""}</td>
      <td class="id-mono">${m.cage_id}</td>
      <td>${tagFor(m.use, "use")}</td>
      <td>${flagsTags(m.cull_flags)}</td>
    </tr>
  `).join("");
  $("#inv-table tbody").innerHTML = html || `<tr><td colspan="9" class="muted" style="text-align:center; padding:24px">No matches.</td></tr>`;
  $("#inv-count").textContent = rows.length + " mice";
}

// Populate strain dropdown
$("#inv-strain").innerHTML = `<option value="">All strains</option>` +
  COLONY.strain_order.map(s => `<option>${s}</option>`).join("");

let invSort = { key: "mouse_id", dir: 1 };
function applyInvFilter() {
  const q = $("#inv-search").value.trim().toLowerCase();
  const strain = $("#inv-strain").value;
  const sex = $("#inv-sex").value;
  const use = $("#inv-use").value;
  const age = $("#inv-age").value;
  let rows = INV.filter(m => {
    if (strain && m.strain !== strain) return false;
    if (sex && m.sex !== sex) return false;
    if (use && m.use !== use) return false;
    if (age) {
      const am = m.age_months ?? -1;
      if (age === "0-2" && !(am >= 0 && am < 2)) return false;
      if (age === "2-4" && !(am >= 2 && am < 4)) return false;
      if (age === "4-6" && !(am >= 4 && am < 6)) return false;
      if (age === "6-12" && !(am >= 6 && am < 12)) return false;
      if (age === "12+" && !(am >= 12)) return false;
    }
    if (q) {
      const blob = (m.mouse_id + " " + m.genotype + " " + m.cage_id + " " + m.strain).toLowerCase();
      if (!blob.includes(q)) return false;
    }
    return true;
  });
  rows.sort((a, b) => {
    const av = a[invSort.key], bv = b[invSort.key];
    if (av == null) return 1;
    if (bv == null) return -1;
    if (av < bv) return -invSort.dir;
    if (av > bv) return invSort.dir;
    return 0;
  });
  renderInventory(rows);
}
["inv-search", "inv-strain", "inv-sex", "inv-use", "inv-age"].forEach(id =>
  $("#" + id).addEventListener("input", applyInvFilter));
$$("#inv-table thead th[data-sort]").forEach(th => th.addEventListener("click", () => {
  const key = th.dataset.sort;
  if (invSort.key === key) invSort.dir *= -1; else invSort = { key, dir: 1 };
  applyInvFilter();
}));
applyInvFilter();

// ============= CAGES =============
function renderCageGrid() {
  const grid = $("#cage-grid");
  let html = `<div class="corner"></div>`;
  for (const c of "ABCDEFGHIJ".split("")) html += `<div class="col-head">${c}</div>`;
  for (let r = 1; r <= 7; r++) {
    html += `<div class="row-head">${r}</div>`;
    for (const c of "ABCDEFGHIJ".split("")) {
      const cage = COLONY.cages.find(x =>
        x.position && x.position.row === r && x.position.col === c);
      if (cage) {
        html += `
          <div class="cage-cell ${cage.has_breeders ? "has-breeders" : ""}" onclick='showCage(${JSON.stringify(cage.cage_id)})'>
            <div class="pos">${r}${c}</div>
            <div class="cid">${cage.cage_id}</div>
            <div class="n">${cage.n_mice} mice</div>
            <div class="strain">${cage.strain || ""}</div>
          </div>`;
      } else {
        html += `<div class="cage-cell empty"><div class="pos muted">${r}${c}</div><div class="cid muted">—</div></div>`;
      }
    }
  }
  grid.innerHTML = html;
}
renderCageGrid();

window.showCage = function(cageId) {
  const cage = COLONY.cages.find(c => c.cage_id === cageId);
  if (!cage) return;
  let body = `
    <p><strong>Cage ${cage.cage_id}</strong> · ${tagFor(cage.strain || "—", "strain")}
       · ${cage.has_breeders ? '<span class="tag warn">Has breeders</span>' : ''}</p>
    <p class="muted">Position: ${cage.position ? cage.position.row + cage.position.col : "Unassigned"}
       · ${cage.n_mice} mice · created ${cage.date_created || "—"}</p>
    <h3>Mice in this cage</h3>
    <div class="scroll-x"><table class="data">
      <thead><tr><th>ID</th><th>Sex</th><th>Age (mo)</th><th>Genotype</th></tr></thead>
      <tbody>${cage.ids.map((id, i) => `
        <tr>
          <td class="id-mono">${id}</td>
          <td>${tagFor(cage.sexes[i] || "Unknown", "sex")}</td>
          <td>${cage.ages_months[i] ?? "—"}</td>
          <td class="id-mono" style="font-size:.78rem">${cage.genotypes[i] || ""}</td>
        </tr>`).join("")}
      </tbody>
    </table></div>
  `;
  openModal(null, "Cage " + cage.cage_id, body);
};

// Unassigned cages
const UNASSIGNED = COLONY.cages.filter(c => !c.position);
$("#unassigned-table tbody").innerHTML = UNASSIGNED.map(c => `
  <tr onclick='showCage(${JSON.stringify(c.cage_id)})' style="cursor:pointer">
    <td class="id-mono">${c.cage_id}</td>
    <td>${tagFor(c.strain || "—", "strain")}</td>
    <td>${c.n_mice}</td>
    <td>${c.has_breeders ? '<span class="tag warn">Yes</span>' : '<span class="tag neutral">No</span>'}</td>
    <td class="id-mono" style="font-size:.72rem">${c.ids.join(", ")}</td>
    <td>${c.date_created || "—"}</td>
  </tr>
`).join("");

// ============= BREEDING =============
const BR = COLONY.breedings;
const brKpis = [
  { label: "Active pairs", val: BR.filter(b => b.status === "Active").length, cls: "" },
  { label: "Total litters", val: BR.reduce((a,b) => a + (b.litters||0), 0), cls: "" },
  { label: "Total pups born", val: BR.reduce((a,b) => a + (b.born||0), 0), cls: "ok" },
];
$("#breeding-kpis").innerHTML = brKpis.map(k => `
  <div class="kpi ${k.cls}">
    <div class="kpi-label">${k.label}</div>
    <div class="kpi-value">${k.val}</div>
  </div>
`).join("");

const brLabels = BR.map((b, i) => (b.offspring_strain || "?") + " #" + (i+1));
charts.brPerf = echarts.init($("#chart-breeding-perf"));
charts.brPerf.setOption({
  grid: { left: 10, right: 24, top: 30, bottom: 60, containLabel: true },
  xAxis: { type: "category", data: brLabels, axisLabel: { rotate: 30, fontSize: 10 } },
  yAxis: [{ type: "value", name: "Born", nameTextStyle: { fontSize: 11 } },
          { type: "value", name: "Litters", nameTextStyle: { fontSize: 11 } }],
  tooltip: { trigger: "axis" },
  legend: { top: 0, textStyle: { fontSize: 11 } },
  series: [
    { name: "Born", type: "bar", data: BR.map(b => b.born || 0),
      itemStyle: { color: "#1e40af" }, barMaxWidth: 30 },
    { name: "Litters", type: "line", yAxisIndex: 1, data: BR.map(b => b.litters || 0),
      itemStyle: { color: "#dc2626" }, lineStyle: { color: "#dc2626", width: 2 },
      symbolSize: 8 }
  ]
});

charts.brAge = echarts.init($("#chart-breeding-age"));
charts.brAge.setOption({
  grid: { left: 10, right: 24, top: 10, bottom: 60, containLabel: true },
  xAxis: { type: "category", data: brLabels, axisLabel: { rotate: 30, fontSize: 10 } },
  yAxis: { type: "value", name: "Months active", nameTextStyle: { fontSize: 11 } },
  tooltip: { trigger: "axis" },
  series: [{
    type: "bar", data: BR.map(b => b.months_active || 0),
    itemStyle: { color: "#7c3aed" }, barMaxWidth: 30,
    label: { show: true, position: "top", fontSize: 10 }
  }]
});

$("#breeding-count").textContent = BR.length + " pairs";
$("#breeding-table tbody").innerHTML = BR.map((b, i) => `
  <tr>
    <td>#${i+1}</td>
    <td>${b.status === "Active" ? '<span class="tag ok">Active</span>' : '<span class="tag neutral">'+b.status+'</span>'}</td>
    <td>${tagFor(b.offspring_strain || "—", "strain")}</td>
    <td class="id-mono">${b.mother}</td>
    <td class="id-mono">${b.father}</td>
    <td class="id-mono" style="font-size:.74rem">${b.mother_geno}</td>
    <td class="id-mono" style="font-size:.74rem">${b.father_geno}</td>
    <td>${b.active_date || "—"}</td>
    <td>${b.months_active ?? "—"}</td>
    <td>${b.litters}</td>
    <td>${b.born}</td>
    <td>${b.avg_litter_size ?? "—"}</td>
    <td>${b.last_weaned}</td>
  </tr>
`).join("");

// ============= COHORT PLANNING =============
$("#cohort-list").innerHTML = COLONY.cohort_plans.map(c => `
  <div class="cohort-card">
    <div class="cohort-head">
      <div>
        <h4>${c.strain} — ${c.target_label || c.target_geno}</h4>
        <div class="cohort-cross">${c.cross_desc}</div>
      </div>
    </div>
    <div class="cohort-purpose">${c.purpose}</div>
    <div class="cohort-target">🎯 Experimental target: <span class="id-mono">${c.target_geno}</span></div>
    <div class="muted" style="font-size:.78rem; margin-bottom:6px;">Expected Mendelian ratios per cross:</div>
    <div class="cohort-ratios">
      ${(c.expected_ratios || []).map(r => `
        <div class="cohort-ratio ${r.target ? 'target' : ''}">
          <span class="id-mono" style="font-size:.78rem">${r.genotype}</span>
          <span class="pct">${r.pct}%</span>
          <div class="bar"><span style="width:${Math.min(100, r.pct*1.33)}%"></span></div>
        </div>
      `).join("")}
    </div>
    ${c.currently_available != null ? `
      <p class="muted" style="font-size:.82rem; margin-top:10px;">
        Currently available toward target: <strong style="color:var(--ok)">${c.currently_available}</strong> mice
      </p>` : ''}
  </div>
`).join("");

// ============= CULL CANDIDATES =============
const CULL_DEFINITIONS = [
  { key: "aged_>12mo_available_nonbreeder",
    label: "Aged > 12 mo, available, not breeder",
    desc: "Default first-pass cull pool. Excludes breeders and project-assigned mice.",
    cls: "warn" },
  { key: "aged_>18mo_any",
    label: "Aged > 18 mo (any use status)",
    desc: "Mice past 18 months — review individually.",
    cls: "alert" },
  { key: "sex_unknown",
    label: "Sex unknown",
    desc: "Identify sex during the next walkthrough; do not cull these.",
    cls: "info" },
  { key: "available_no_breeders_strain",
    label: "Available, strain has no breeders",
    desc: "Strain has no active breeding pairs and these mice are unassigned.",
    cls: "info" },
];

$("#cull-cards").innerHTML = CULL_DEFINITIONS.map(d => `
  <div class="cull-card ${d.cls}" onclick="showCull('${d.key}')">
    <div class="count">${COLONY.cull_summary[d.key].count}</div>
    <div class="label">${d.label}</div>
    <div class="desc">${d.desc}</div>
  </div>
`).join("");

// Populate filter
$("#cull-filter").innerHTML =
  `<option value="all">Show all candidates</option>` +
  CULL_DEFINITIONS.map(d => `<option value="${d.key}">${d.label}</option>`).join("");

function getCullRows(key) {
  if (key === "all") {
    const set = new Set();
    Object.keys(COLONY.cull_summary).forEach(k =>
      COLONY.cull_summary[k].ids.forEach(id => set.add(id)));
    return INV.filter(m => set.has(m.mouse_id));
  }
  const ids = new Set(COLONY.cull_summary[key].ids);
  return INV.filter(m => ids.has(m.mouse_id));
}

function renderCullTable(key) {
  const rows = getCullRows(key);
  $("#cull-count").textContent = rows.length + " candidates";
  $("#cull-table tbody").innerHTML = rows.map(m => `
    <tr>
      <td><input type="checkbox" class="cull-cb" data-id="${m.mouse_id}" /></td>
      <td class="id-mono">${m.mouse_id}</td>
      <td>${tagFor(m.strain, "strain")}</td>
      <td>${tagFor(m.sex, "sex")}</td>
      <td>${m.age_months ?? "—"}</td>
      <td class="id-mono" style="font-size:.75rem">${m.genotype}</td>
      <td class="id-mono">${m.cage_id}</td>
      <td>${tagFor(m.use, "use")}</td>
      <td>${flagsTags(m.cull_flags)}</td>
    </tr>
  `).join("") || `<tr><td colspan="9" class="muted" style="text-align:center; padding:24px">No candidates.</td></tr>`;
}

window.showCull = function(key) {
  $("#cull-filter").value = key;
  renderCullTable(key);
  document.getElementById("cull-table").scrollIntoView({behavior: "smooth", block: "start"});
};

$("#cull-filter").addEventListener("change", () => renderCullTable($("#cull-filter").value));
$("#cull-all").addEventListener("change", e => {
  $$(".cull-cb").forEach(cb => cb.checked = e.target.checked);
});
renderCullTable("all");

window.exportCullCSV = function() {
  const rows = $$(".cull-cb:checked");
  let toExport;
  if (rows.length === 0) {
    toExport = getCullRows($("#cull-filter").value);
  } else {
    const ids = new Set(rows.map(r => r.dataset.id));
    toExport = INV.filter(m => ids.has(m.mouse_id));
  }
  const headers = ["mouse_id","strain","sex","age_months","dob","genotype","cage_id","use","flags"];
  const csv = [headers.join(",")].concat(
    toExport.map(m => [
      m.mouse_id, m.strain, m.sex, m.age_months, m.dob, '"'+(m.genotype||'').replace(/"/g,'""')+'"',
      m.cage_id, m.use, (m.cull_flags||[]).join(";")
    ].join(","))
  ).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "cull_candidates_" + COLONY.as_of + ".csv"; a.click();
  URL.revokeObjectURL(url);
};

// ============= STATISTICS =============
function median(arr) {
  if (!arr.length) return null;
  const s = arr.slice().sort((a,b) => a-b);
  const m = Math.floor(s.length/2);
  return s.length % 2 ? s[m] : (s[m-1] + s[m]) / 2;
}
function mean(arr) { return arr.length ? arr.reduce((a,b) => a+b, 0) / arr.length : null; }

const ageStatsRows = COLONY.strain_order.map(s => {
  const ages = INV.filter(m => m.strain === s && m.age_months != null).map(m => m.age_months);
  if (!ages.length) return null;
  return {
    strain: s, n: ages.length,
    median: median(ages).toFixed(1),
    mean: mean(ages).toFixed(1),
    min: Math.min(...ages).toFixed(1),
    max: Math.max(...ages).toFixed(1)
  };
}).filter(Boolean);
$("#age-stats-table tbody").innerHTML = ageStatsRows.map(r => `
  <tr>
    <td>${tagFor(r.strain, "strain")}</td>
    <td>${r.n}</td><td>${r.median}</td><td>${r.mean}</td><td>${r.min}</td><td>${r.max}</td>
  </tr>
`).join("");

const sexRatioRows = COLONY.strain_order.map(s => {
  const f = COLONY.sex_by_strain[s].Female || 0;
  const m = COLONY.sex_by_strain[s].Male || 0;
  const tot = f + m;
  return {
    strain: s, f, m,
    ratio: m === 0 ? "—" : (f/m).toFixed(2),
    pctF: tot ? Math.round(100*f/tot) : 0
  };
});
$("#sex-ratio-table tbody").innerHTML = sexRatioRows.map(r => `
  <tr>
    <td>${tagFor(r.strain, "strain")}</td>
    <td>${r.f}</td><td>${r.m}</td><td>${r.ratio}</td><td>${r.pctF}%</td>
  </tr>
`).join("");

$("#breeding-stats-table tbody").innerHTML = BR.map((b, i) => {
  const prod = (b.months_active && b.born) ? (b.born / b.months_active).toFixed(2) : "—";
  return `
    <tr>
      <td>#${i+1}</td>
      <td>${tagFor(b.offspring_strain || "—", "strain")}</td>
      <td>${b.litters}</td><td>${b.born}</td>
      <td>${b.avg_litter_size ?? "—"}</td>
      <td>${b.months_active ?? "—"}</td>
      <td>${prod}</td>
    </tr>
  `;
}).join("");

// ============= COHORT BUILDER =============
const CB = {
  selected: new Map(),  // mouse_id -> { mouse, role: "experimental" | "control" }
  candidates: [],       // current filtered list
};

function normalizeGeno(g) {
  return (g || "").toLowerCase().replace(/\s+/g, "").replace(/;+/g, ";");
}

function cbInit() {
  // Populate strain dropdown
  $("#cb-strain").innerHTML = `<option value="">All strains</option>` +
    COLONY.strain_order.map(s => `<option>${s}</option>`).join("");

  // Populate genotype dropdown — grouped by strain
  cbPopulateGenotypes($("#cb-strain").value);

  // Re-populate genotype options whenever strain changes
  $("#cb-strain").addEventListener("change", () => {
    cbPopulateGenotypes($("#cb-strain").value);
    $("#cb-genotype").value = "";  // reset to "all genotypes"
    cbApplyFilter();
  });

  ["cb-strain", "cb-sex", "cb-genotype", "cb-age-min", "cb-age-max", "cb-use"]
    .forEach(id => $("#" + id).addEventListener("input", cbApplyFilter));
  cbApplyFilter();
  cbRenderCohort();
}

function cbPopulateGenotypes(strainFilter) {
  const sel = $("#cb-genotype");
  sel.innerHTML = "";
  const allOpt = document.createElement("option");
  allOpt.value = "";
  allOpt.textContent = "All genotypes";
  sel.appendChild(allOpt);

  // Build (strain → Map<genotype, count>)
  const byStrain = new Map();
  for (const m of INV) {
    if (strainFilter && m.strain !== strainFilter) continue;
    const g = (m.genotype || "").trim();
    if (!g) continue;
    if (!byStrain.has(m.strain)) byStrain.set(m.strain, new Map());
    const inner = byStrain.get(m.strain);
    inner.set(g, (inner.get(g) || 0) + 1);
  }
  // Order strains using COLONY.strain_order, then any extras
  const orderedStrains = COLONY.strain_order.filter(s => byStrain.has(s))
    .concat(Array.from(byStrain.keys()).filter(s => !COLONY.strain_order.includes(s)));

  for (const strain of orderedStrains) {
    const inner = byStrain.get(strain);
    const og = document.createElement("optgroup");
    og.label = strain + (strainFilter ? "" : "");
    const sorted = Array.from(inner.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    for (const [g, n] of sorted) {
      const opt = document.createElement("option");
      opt.value = g;
      opt.textContent = `${g}  (n=${n})`;
      og.appendChild(opt);
    }
    sel.appendChild(og);
  }
}

function cbApplyFilter() {
  const strain = $("#cb-strain").value;
  const sex = $("#cb-sex").value;
  const geno = $("#cb-genotype").value.trim();
  const ageMin = parseFloat($("#cb-age-min").value);
  const ageMax = parseFloat($("#cb-age-max").value);
  const use = $("#cb-use").value;

  CB.candidates = INV.filter(m => {
    if (strain && m.strain !== strain) return false;
    if (sex && m.sex !== sex) return false;
    if (use && m.use !== use) return false;
    if (geno && (m.genotype || "").trim() !== geno) return false;
    if (!isNaN(ageMin) && (m.age_months ?? -1) < ageMin) return false;
    if (!isNaN(ageMax) && (m.age_months ?? 9999) > ageMax) return false;
    return true;
  });

  $("#cb-cand-count").textContent = CB.candidates.length;
  const tbody = $("#cb-cand-table tbody");
  if (!CB.candidates.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="cb-empty">No mice match these filters.</td></tr>`;
    return;
  }
  tbody.innerHTML = CB.candidates.slice(0, 200).map(m => {
    const inCohort = CB.selected.has(m.mouse_id);
    return `
      <tr>
        <td>
          ${inCohort
            ? `<span class="role-tag ${CB.selected.get(m.mouse_id).role === 'experimental' ? 'exp' : 'ctrl'}">in</span>`
            : `<button class="btn small" onclick="cbAdd('${m.mouse_id}','experimental')">+ exp</button>
               <button class="btn small ghost" onclick="cbAdd('${m.mouse_id}','control')">+ ctrl</button>`
          }
        </td>
        <td class="id-mono">${m.mouse_id}</td>
        <td>${tagFor(m.strain, "strain")}</td>
        <td>${tagFor(m.sex, "sex")}</td>
        <td>${m.age_months ?? "—"}</td>
        <td class="id-mono" style="font-size:.74rem">${m.genotype || ""}</td>
        <td class="id-mono">${m.cage_id}</td>
      </tr>`;
  }).join("");
  if (CB.candidates.length > 200) {
    tbody.innerHTML += `<tr><td colspan="7" class="cb-empty">…showing first 200 of ${CB.candidates.length}. Narrow filters to see more.</td></tr>`;
  }
}

window.cbAdd = function(mouseId, role) {
  const m = INV.find(x => x.mouse_id === mouseId);
  if (!m) return;
  CB.selected.set(mouseId, { mouse: m, role });
  cbRenderCohort();
  cbApplyFilter();
};

window.cbRemove = function(mouseId) {
  CB.selected.delete(mouseId);
  cbRenderCohort();
  cbApplyFilter();
};

window.cbAddAllVisible = function(role) {
  CB.candidates.slice(0, 200).forEach(m => {
    if (!CB.selected.has(m.mouse_id)) CB.selected.set(m.mouse_id, { mouse: m, role });
  });
  cbRenderCohort();
  cbApplyFilter();
};

window.cbClear = function() {
  if (CB.selected.size === 0) return;
  if (!confirm("Clear the cohort? (" + CB.selected.size + " mice)")) return;
  CB.selected.clear();
  cbRenderCohort();
  cbApplyFilter();
  $("#cb-control-area").innerHTML = "";
};

function cbRenderCohort() {
  const all = Array.from(CB.selected.values());
  const exp = all.filter(x => x.role === "experimental");
  const ctrl = all.filter(x => x.role === "control");
  const f = all.filter(x => x.mouse.sex === "Female").length;
  const ml = all.filter(x => x.mouse.sex === "Male").length;
  const ages = all.map(x => x.mouse.age_months).filter(a => a != null);
  const meanAge = ages.length ? (ages.reduce((s,a)=>s+a,0)/ages.length).toFixed(1) : "—";

  $("#cb-cohort-count").textContent = all.length;
  $("#cb-cohort-summary").innerHTML = `
    <div class="item"><div class="v">${all.length}</div><div class="l">Total</div></div>
    <div class="item"><div class="v" style="color:var(--brand)">${exp.length}</div><div class="l">Experimental</div></div>
    <div class="item"><div class="v" style="color:var(--ok)">${ctrl.length}</div><div class="l">Controls</div></div>
    <div class="item"><div class="v">${f}:${ml}</div><div class="l">F : M</div></div>
    <div class="item"><div class="v">${meanAge}</div><div class="l">Mean age (mo)</div></div>
  `;
  const tbody = $("#cb-cohort-table tbody");
  if (!all.length) {
    tbody.innerHTML = `<tr><td colspan="8" class="cb-empty">Pick mice from the left to build a cohort.</td></tr>`;
    return;
  }
  tbody.innerHTML = all.map(({mouse: m, role}) => `
    <tr>
      <td><span class="role-tag ${role==='experimental'?'exp':'ctrl'}">${role==='experimental'?'Exp':'Ctrl'}</span></td>
      <td class="id-mono">${m.mouse_id}</td>
      <td>${tagFor(m.strain, "strain")}</td>
      <td>${tagFor(m.sex, "sex")}</td>
      <td>${m.age_months ?? "—"}</td>
      <td class="id-mono" style="font-size:.74rem">${m.genotype || ""}</td>
      <td class="id-mono">${m.cage_id}</td>
      <td><button class="btn small danger" onclick="cbRemove('${m.mouse_id}')">×</button></td>
    </tr>`).join("");
}

window.cbSuggestControls = function() {
  const exp = Array.from(CB.selected.values()).filter(x => x.role === "experimental");
  if (!exp.length) {
    $("#cb-control-area").innerHTML = `<div class="cb-empty" style="margin-top:10px;">Add at least one experimental mouse first, then click "Suggest controls".</div>`;
    return;
  }
  // Build map: strain -> set of "control genotypes" from cohort plans (target=false)
  const ctrlGenosByStrain = {};
  COLONY.cohort_plans.forEach(c => {
    const set = new Set((c.expected_ratios || [])
      .filter(r => !r.target)
      .map(r => normalizeGeno(r.genotype)));
    ctrlGenosByStrain[c.strain] = set;
  });

  let html = `<h4 style="margin: 18px 0 10px; font-size:.92rem; color: var(--ink);">Suggested genetic controls</h4>`;
  exp.forEach(({mouse: m}) => {
    const targetGenoNorm = normalizeGeno(m.genotype);
    const candidates = INV.filter(x => {
      if (x.strain !== m.strain) return false;
      if (x.sex !== m.sex) return false;
      if (x.mouse_id === m.mouse_id) return false;
      if (CB.selected.has(x.mouse_id)) return false;  // already in cohort
      if (normalizeGeno(x.genotype) === targetGenoNorm) return false;  // same geno = not control
      if (x.use === "Breeding") return false;
      return true;
    });

    // Score: prefer mice whose genotype is in the cohort plan's "control" list,
    // then by absolute age difference.
    const ctrlSet = ctrlGenosByStrain[m.strain] || new Set();
    candidates.forEach(c => {
      const inPlan = ctrlSet.has(normalizeGeno(c.genotype));
      const ageDiff = (c.age_months != null && m.age_months != null)
        ? Math.abs(c.age_months - m.age_months) : 999;
      c._score = (inPlan ? 0 : 100) + ageDiff;
      c._inPlan = inPlan;
      c._ageDiff = ageDiff;
    });
    candidates.sort((a,b) => a._score - b._score);
    const top = candidates.slice(0, 5);

    html += `
      <div class="control-suggestions">
        <h5>
          <span>For <span class="id-mono">${m.mouse_id}</span> · ${m.strain} · ${m.sex} · ${m.age_months}mo</span>
          <span class="target">${m.genotype || ""}</span>
        </h5>
        ${top.length ? `
        <table>
          <thead><tr>
            <th>ID</th><th>Sex</th><th>Age</th><th>Δ age</th><th>Genotype</th><th>Cage</th><th></th>
          </tr></thead>
          <tbody>
            ${top.map(c => `
              <tr>
                <td class="id-mono">${c.mouse_id}${c._inPlan ? ' <span class="role-tag ctrl" style="margin-left:4px">plan</span>' : ''}</td>
                <td>${tagFor(c.sex, "sex")}</td>
                <td>${c.age_months ?? "—"}</td>
                <td><span class="age-diff ${c._ageDiff < 1 ? 'close' : ''}">±${c._ageDiff.toFixed(1)} mo</span></td>
                <td class="id-mono" style="font-size:.72rem">${c.genotype || ""}</td>
                <td class="id-mono">${c.cage_id}</td>
                <td><button class="btn small ok" onclick="cbAdd('${c.mouse_id}','control'); cbSuggestControls();">+ add</button></td>
              </tr>`).join("")}
          </tbody>
        </table>` : `<p class="muted" style="font-size:.8rem">No matching controls found in the colony.</p>`}
      </div>`;
  });
  $("#cb-control-area").innerHTML = html;
};

window.cbExportXlsx = function() {
  if (CB.selected.size === 0) {
    alert("Cohort is empty. Add some mice first.");
    return;
  }
  const rows = Array.from(CB.selected.values()).map(({mouse: m, role}) => ({
    role,
    mouse_id: m.mouse_id,
    strain: m.strain,
    sex: m.sex,
    age_months: m.age_months,
    age_days: m.age_days,
    dob: m.dob,
    genotype: m.genotype,
    cage_id: m.cage_id,
    use: m.use,
    flags: (m.cull_flags || []).join("; "),
  }));

  // Group summary
  const exp = rows.filter(r => r.role === "experimental");
  const ctrl = rows.filter(r => r.role === "control");
  const summary = [
    ["Cohort exported", new Date().toISOString().slice(0,19).replace("T", " ")],
    ["Source dashboard", "MacDougald Lab Mouse Colony Dashboard"],
    ["Data as of", COLONY.as_of],
    [],
    ["Total mice in cohort", rows.length],
    ["  Experimental", exp.length],
    ["  Controls", ctrl.length],
    [],
    ["Filters used at time of export"],
    ["  Strain", $("#cb-strain").value || "(all)"],
    ["  Sex", $("#cb-sex").value || "(all)"],
    ["  Genotype contains", $("#cb-genotype").value || "(any)"],
    ["  Age min (mo)", $("#cb-age-min").value || "(any)"],
    ["  Age max (mo)", $("#cb-age-max").value || "(any)"],
    ["  Use", $("#cb-use").value || "(any)"],
  ];

  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.json_to_sheet(rows, {
    header: ["role", "mouse_id", "strain", "sex", "age_months", "age_days", "dob", "genotype", "cage_id", "use", "flags"]
  });
  // Auto column widths
  const cols = Object.keys(rows[0] || {});
  ws1["!cols"] = cols.map(k => {
    const max = Math.max(k.length, ...rows.map(r => String(r[k] ?? "").length));
    return { wch: Math.min(40, Math.max(8, max + 2)) };
  });
  XLSX.utils.book_append_sheet(wb, ws1, "Cohort");

  const ws2 = XLSX.utils.aoa_to_sheet(summary);
  ws2["!cols"] = [{ wch: 30 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Summary");

  const fname = `Cohort_${COLONY.as_of}_${rows.length}mice.xlsx`;
  XLSX.writeFile(wb, fname);
};

cbInit();

// ============= SMART CULL PLAN =============
const SC = {
  // Defaults derived from user's biology:
  //   Marrow Glo → expand
  //   Adipoq-Cre, mTmG, Dendra2 → wind-down (rescuable)
  //   AdipoGlo, AdipoGlo+ → maintain
  defaultPriority: {
    "AdipoGlo": "maintain",
    "AdipoGlo+": "maintain",
    "Marrow Glo": "expand",
    "Adipoq-Cre": "winddown",
    "mTmG": "winddown",
    "Dendra2": "winddown",
  },
  priority: {},
  keepFloor: 4,
  rescueOn: true,
  donorN: 4,
  plan: null,
};

// Rescue-donor predicates: identify mice in OTHER strains that can serve as
// breeding stock to regenerate an orphan single-transgene line.
const RESCUE_RULES = {
  "Adipoq-Cre": {
    label: "Adipoq-Cre alone",
    note: "Find offspring carrying only the Adipoq-Cre transgene (Dendra2 null, no mTmG).",
    candidate: m => {
      const g = (m.genotype || "").toLowerCase().replace(/\s+/g, "");
      return /adipoq-cre/.test(g) && /dendra2<-\/->/.test(g) && !/mtmg<mtmg/.test(g);
    },
  },
  "Dendra2": {
    label: "Dendra2 alone",
    note: "Find offspring carrying Dendra2 (+/- or +/+) without Adipoq-Cre and without mTmG.",
    candidate: m => {
      const g = (m.genotype || "").toLowerCase().replace(/\s+/g, "");
      return /dendra2<\+\/[+-]>/.test(g) && !/adipoq-cre/.test(g) && !/mtmg<mtmg/.test(g);
    },
  },
  "mTmG": {
    label: "mTmG alone",
    note: "Find offspring carrying mTmG hom without Adipoq-Cre and without Dendra2 transgene.",
    candidate: m => {
      const g = (m.genotype || "").toLowerCase().replace(/\s+/g, "");
      return /mtmg<mtmg\/mtmg>/.test(g) && !/adipoq-cre/.test(g) && !/dendra2<\+\/[+-]>/.test(g);
    },
  },
};

function scInit() {
  // Initialize priorities from defaults (only for strains present in the colony)
  COLONY.strain_order.forEach(s => {
    SC.priority[s] = SC.defaultPriority[s] || "maintain";
  });

  renderPriorityGrid();

  // Wire inputs
  $("#sc-keepfloor").addEventListener("input", () => {
    SC.keepFloor = parseInt($("#sc-keepfloor").value) || 4;
  });
  $("#sc-rescue-on").addEventListener("change", () => { SC.rescueOn = $("#sc-rescue-on").checked; });
  $("#sc-donor-n").addEventListener("input", () => {
    SC.donorN = parseInt($("#sc-donor-n").value) || 4;
  });

  // Compute initial plan
  scRecompute();
}

function renderPriorityGrid() {
  const html = COLONY.strain_order.map(s => {
    const n = COLONY.strain_counts[s] || 0;
    const rescueNote = RESCUE_RULES[s]
      ? `<div class="muted" style="font-size:.74rem;margin-top:4px;">↪︎ ${RESCUE_RULES[s].note}</div>`
      : "";
    const cur = SC.priority[s];
    return `
      <div class="cull-card ${cur === 'expand' ? 'info' : cur === 'winddown' ? 'warn' : ''}"
           style="cursor:default;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;">
          <div>
            <div class="label" style="font-weight:700;color:var(--ink)">${s}</div>
            <div class="desc" style="margin-top:2px">${n} mice in colony</div>
          </div>
          <select onchange="scSetPriority('${s.replace(/'/g, "\\'")}', this.value)"
                  style="padding:6px 10px;border:1px solid var(--line);border-radius:7px;font-size:.82rem;font-family:inherit;background:#fff;">
            <option value="expand" ${cur === 'expand' ? 'selected' : ''}>Expand</option>
            <option value="maintain" ${cur === 'maintain' ? 'selected' : ''}>Maintain</option>
            <option value="winddown" ${cur === 'winddown' ? 'selected' : ''}>Wind down</option>
          </select>
        </div>
        ${rescueNote}
      </div>`;
  }).join("");
  $("#sc-priority-grid").innerHTML = `<div class="cull-list">${html}</div>`;
}

window.scSetPriority = function(strain, value) {
  SC.priority[strain] = value;
  renderPriorityGrid();
  scRecompute();
};

// Cohort plan that produced this strain (for "rescuable" tag in UI)
function isRescuable(strain) {
  return RESCUE_RULES[strain] != null;
}

// Active-breeder mouse_ids parsed from breedings
function getActiveBreederIds() {
  const ids = new Set();
  COLONY.breedings.forEach(b => {
    if (b.status !== "Active") return;
    [b.mother, b.father].forEach(s => {
      if (!s) return;
      const id = String(s).split("|")[0].trim();
      if (id) ids.add(id);
    });
  });
  // Also use anyone marked as Use=Breeding in inventory
  INV.forEach(m => { if (m.use === "Breeding") ids.add(m.mouse_id); });
  return ids;
}

window.scRecompute = function() {
  const breederIds = getActiveBreederIds();

  // Step 1: identify rescue donors for each wind-down rescuable strain
  const rescueDonors = new Map(); // mouseId -> targetStrain
  if (SC.rescueOn) {
    Object.entries(RESCUE_RULES).forEach(([target, rule]) => {
      if (SC.priority[target] !== "winddown") return;
      const candidates = INV.filter(m =>
        m.strain !== target &&
        rule.candidate(m) &&
        m.use === "Available" &&
        !breederIds.has(m.mouse_id) &&
        m.sex !== "Unknown"
      );
      ["Female", "Male"].forEach(sex => {
        candidates
          .filter(c => c.sex === sex)
          .sort((a, b) => (a.age_months ?? 999) - (b.age_months ?? 999))
          .slice(0, SC.donorN)
          .forEach(m => rescueDonors.set(m.mouse_id, target));
      });
    });
  }

  // Step 2: precompute (strain × genotype × sex) cell rankings
  const cellMembers = new Map();
  for (const m of INV) {
    const key = `${m.strain}||${(m.genotype || "").trim()}||${m.sex}`;
    if (!cellMembers.has(key)) cellMembers.set(key, []);
    cellMembers.get(key).push(m);
  }
  cellMembers.forEach(arr => arr.sort((a, b) => (a.age_months ?? 999) - (b.age_months ?? 999)));
  const cellSizesGeno = new Map();
  for (const m of INV) {
    const key = `${m.strain}||${(m.genotype || "").trim()}`;
    cellSizesGeno.set(key, (cellSizesGeno.get(key) || 0) + 1);
  }

  // Step 3: classify
  const preserve = [];
  const cull = [];

  for (const m of INV) {
    const reasons = [];
    let preserveDecision = false;
    const priority = SC.priority[m.strain] || "maintain";

    if (m.use === "Breeding" || breederIds.has(m.mouse_id)) {
      preserveDecision = true; reasons.push("active breeder");
    }
    if (m.sex === "Unknown") {
      preserveDecision = true; reasons.push("sex unknown — ID first");
    }
    if (rescueDonors.has(m.mouse_id)) {
      preserveDecision = true;
      reasons.push(`rescue donor for ${rescueDonors.get(m.mouse_id)}`);
    }

    // Singleton genotype protection (unless rescuable)
    const genoCellSize = cellSizesGeno.get(`${m.strain}||${(m.genotype || "").trim()}`) || 0;
    if (genoCellSize === 1 && !isRescuable(m.strain)) {
      preserveDecision = true;
      reasons.push("singleton genotype (last of kind)");
    }

    // Strain-priority rules
    if (!preserveDecision) {
      if (priority === "expand") {
        preserveDecision = true; reasons.push("strain priority: EXPAND");
      } else if (priority === "winddown") {
        // Cull most; only preserve if young (<6mo) or if singleton (handled above)
        if ((m.age_months ?? 999) < 6) {
          preserveDecision = true; reasons.push("under 6 mo — preserve in wind-down");
        }
      } else if (priority === "maintain") {
        // Keep ≥keepFloor youngest of (strain × genotype × sex)
        const cellKey = `${m.strain}||${(m.genotype || "").trim()}||${m.sex}`;
        const cellArr = cellMembers.get(cellKey) || [];
        const idx = cellArr.findIndex(x => x.mouse_id === m.mouse_id);
        const floor = Math.min(SC.keepFloor, cellArr.length);
        if (idx < floor) {
          preserveDecision = true; reasons.push(`youngest in ${m.strain}/${m.sex} for this genotype (rank ${idx + 1}/${cellArr.length})`);
        } else if ((m.age_months ?? 999) < 12) {
          preserveDecision = true; reasons.push("under 12 mo — too young to cull");
        }
      }
    }

    const row = { mouse: m, reasons };
    if (preserveDecision) preserve.push(row);
    else {
      // default cull reason
      if (!reasons.length) reasons.push("aged surplus");
      else reasons.push("aged surplus");
      cull.push(row);
    }
  }

  SC.plan = { preserve, cull, rescueDonors };

  // Render summary, actions, tables
  scRenderSummary();
  scRenderActions();
  scRenderTables();
};

function scRenderSummary() {
  const { preserve, cull } = SC.plan;
  const breederCount = preserve.filter(p => p.reasons.some(r => r.startsWith("active"))).length;
  const rescueCount = preserve.filter(p => p.reasons.some(r => r.startsWith("rescue"))).length;
  const sexUnkCount = preserve.filter(p => p.reasons.some(r => r.startsWith("sex unknown"))).length;
  const expandCount = preserve.filter(p => p.reasons.some(r => r.includes("EXPAND"))).length;
  const insuranceCount = preserve.filter(p => p.reasons.some(r => r.startsWith("youngest"))).length;
  const youngCount = preserve.filter(p => p.reasons.some(r => r.startsWith("under"))).length;
  const singletonCount = preserve.filter(p => p.reasons.some(r => r.startsWith("singleton"))).length;

  $("#sc-summary").innerHTML = `
    <div class="kpi alert">
      <div class="kpi-label">Cull</div>
      <div class="kpi-value">${cull.length}</div>
      <div class="kpi-sub">${Math.round(100 * cull.length / INV.length)}% of colony</div>
    </div>
    <div class="kpi ok">
      <div class="kpi-label">Preserve</div>
      <div class="kpi-value">${preserve.length}</div>
      <div class="kpi-sub">colony goes to ${preserve.length} mice</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Preservation reasons</div>
      <div class="kpi-sub" style="margin-top:8px; line-height:1.7;">
        <span class="tag warn">${breederCount}</span> breeders ·
        <span class="tag neutral">${sexUnkCount}</span> sex unknown ·
        <span class="tag ok">${rescueCount}</span> rescue donors ·
        <span class="tag brand">${expandCount}</span> expand-strain ·
        <span class="tag">${insuranceCount}</span> cohort insurance ·
        <span class="tag">${youngCount}</span> too young ·
        <span class="tag danger">${singletonCount}</span> singleton geno
      </div>
    </div>
  `;
}

function scRenderActions() {
  const actions = [];
  // Expand strains: recommend breedings if no active breeders exist
  Object.entries(SC.priority).forEach(([strain, p]) => {
    if (p !== "expand") return;
    const strainMice = INV.filter(m => m.strain === strain);
    const breeders = strainMice.filter(m => m.use === "Breeding");
    if (breeders.length === 0) {
      // Recommend a pair from young mice
      const candidates = strainMice.filter(m => m.use === "Available");
      const f = candidates.filter(m => m.sex === "Female").sort((a, b) => (a.age_months ?? 999) - (b.age_months ?? 999));
      const ml = candidates.filter(m => m.sex === "Male").sort((a, b) => (a.age_months ?? 999) - (b.age_months ?? 999));
      const suggF = f[0], suggM = ml[0];
      actions.push(`
        <div class="control-suggestions" style="background:#f0f9ff;border-color:#bae6fd;">
          <h5 style="color:#0369a1">⚙ Set up breeding for <strong>${strain}</strong> (priority: EXPAND, no active breeders)</h5>
          <p style="font-size:.82rem;margin:0 0 8px;color:var(--ink-muted)">
            ${strain} has ${strainMice.length} mice but no active breeding pairs. Suggested pair (youngest available):
          </p>
          ${(suggF && suggM) ? `
            <table>
              <tr><td>♀ Female</td><td class="id-mono">${suggF.mouse_id}</td><td>${suggF.age_months}mo</td><td class="id-mono" style="font-size:.74rem">${suggF.genotype || ""}</td></tr>
              <tr><td>♂ Male</td><td class="id-mono">${suggM.mouse_id}</td><td>${suggM.age_months}mo</td><td class="id-mono" style="font-size:.74rem">${suggM.genotype || ""}</td></tr>
            </table>
          ` : `<p class="muted">Not enough Available mice of both sexes to suggest a pair.</p>`}
        </div>`);
    } else {
      actions.push(`
        <div class="control-suggestions" style="background:#f0fdf4;border-color:#bbf7d0;">
          <h5 style="color:#15803d">✓ <strong>${strain}</strong> (EXPAND) — already has ${breeders.length} active breeder(s)</h5>
          <p style="font-size:.82rem;margin:0;color:var(--ink-muted)">No new breeding setup required; preserving all ${strainMice.length} mice in this strain.</p>
        </div>`);
    }
  });
  // Rescue donor announcements
  Object.entries(RESCUE_RULES).forEach(([target, rule]) => {
    if (SC.priority[target] !== "winddown" || !SC.rescueOn) return;
    const donors = Array.from(SC.plan.rescueDonors.entries())
      .filter(([id, t]) => t === target)
      .map(([id]) => INV.find(m => m.mouse_id === id))
      .filter(Boolean);
    if (donors.length === 0) {
      actions.push(`
        <div class="control-suggestions" style="background:#fef2f2;border-color:#fecaca;">
          <h5 style="color:var(--danger)">⚠ No rescue donors found for <strong>${target}</strong></h5>
          <p style="font-size:.82rem;margin:0;color:var(--ink-muted)">
            ${rule.note} The colony does not currently contain candidate donors.
            Consider sourcing fresh ${target} stock or cryopreserving any ${target} mouse before culling.
          </p>
        </div>`);
    } else {
      const f = donors.filter(d => d.sex === "Female"), ml = donors.filter(d => d.sex === "Male");
      actions.push(`
        <div class="control-suggestions" style="background:#f0fdf4;border-color:#bbf7d0;">
          <h5 style="color:#15803d">✓ Rescue donors preserved for <strong>${target}</strong></h5>
          <p style="font-size:.82rem;margin:0 0 6px;color:var(--ink-muted)">
            ${rule.note} ${donors.length} donor(s) preserved (${f.length} F · ${ml.length} M):
          </p>
          <table>
            ${donors.map(d => `<tr>
              <td class="id-mono">${d.mouse_id}</td>
              <td>${tagFor(d.strain, "strain")}</td>
              <td>${tagFor(d.sex, "sex")}</td>
              <td>${d.age_months}mo</td>
              <td class="id-mono" style="font-size:.72rem">${d.genotype || ""}</td>
              <td class="id-mono">${d.cage_id}</td>
            </tr>`).join("")}
          </table>
        </div>`);
    }
  });

  if (actions.length) {
    $("#sc-actions").style.display = "block";
    $("#sc-actions-body").innerHTML = actions.join("");
  } else {
    $("#sc-actions").style.display = "none";
  }
}

function scRenderTables() {
  const cull = SC.plan.cull.slice().sort((a, b) =>
    a.mouse.strain.localeCompare(b.mouse.strain) ||
    (b.mouse.age_months ?? 0) - (a.mouse.age_months ?? 0));
  const preserve = SC.plan.preserve.slice().sort((a, b) =>
    a.mouse.strain.localeCompare(b.mouse.strain) ||
    (a.mouse.age_months ?? 0) - (b.mouse.age_months ?? 0));

  $("#sc-cull-count").textContent = cull.length + " mice";
  $("#sc-preserve-count").textContent = preserve.length + " mice";

  $("#sc-cull-table tbody").innerHTML = cull.map(({mouse: m, reasons}) => `
    <tr>
      <td class="id-mono">${m.mouse_id}</td>
      <td>${tagFor(m.strain, "strain")}</td>
      <td>${tagFor(m.sex, "sex")}</td>
      <td>${m.age_months ?? "—"}</td>
      <td class="id-mono" style="font-size:.74rem">${m.genotype || ""}</td>
      <td class="id-mono">${m.cage_id}</td>
      <td>${tagFor(m.use, "use")}</td>
      <td style="font-size:.78rem;color:var(--ink-muted)">${reasons.join("; ")}</td>
    </tr>`).join("") || `<tr><td colspan="8" class="cb-empty">No mice flagged for cull with current settings.</td></tr>`;

  $("#sc-preserve-table tbody").innerHTML = preserve.map(({mouse: m, reasons}) => `
    <tr>
      <td class="id-mono">${m.mouse_id}</td>
      <td>${tagFor(m.strain, "strain")}</td>
      <td>${tagFor(m.sex, "sex")}</td>
      <td>${m.age_months ?? "—"}</td>
      <td class="id-mono" style="font-size:.74rem">${m.genotype || ""}</td>
      <td class="id-mono">${m.cage_id}</td>
      <td>${tagFor(m.use, "use")}</td>
      <td style="font-size:.78rem;color:var(--ink-muted)">${reasons.join("; ")}</td>
    </tr>`).join("");
}

window.scExportCullCSV = function() {
  if (!SC.plan || !SC.plan.cull.length) { alert("Cull list is empty."); return; }
  const headers = ["mouse_id","strain","sex","age_months","dob","genotype","cage_id","use","reason"];
  const csv = [headers.join(",")].concat(
    SC.plan.cull.map(({mouse: m, reasons}) => [
      m.mouse_id, m.strain, m.sex, m.age_months ?? "", m.dob ?? "",
      '"'+(m.genotype||'').replace(/"/g,'""')+'"', m.cage_id, m.use,
      '"'+reasons.join("; ")+'"'
    ].join(","))
  ).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `Smart_Cull_Plan_${COLONY.as_of}.csv`; a.click();
  URL.revokeObjectURL(url);
};

window.scExportCull = function() {
  if (!SC.plan || !SC.plan.cull.length) { alert("Cull list is empty."); return; }
  const cullRows = SC.plan.cull.map(({mouse: m, reasons}) => ({
    mouse_id: m.mouse_id, strain: m.strain, sex: m.sex,
    age_months: m.age_months, dob: m.dob, genotype: m.genotype,
    cage_id: m.cage_id, use: m.use, reason: reasons.join("; ")
  }));
  const preserveRows = SC.plan.preserve.map(({mouse: m, reasons}) => ({
    mouse_id: m.mouse_id, strain: m.strain, sex: m.sex,
    age_months: m.age_months, dob: m.dob, genotype: m.genotype,
    cage_id: m.cage_id, use: m.use, reason: reasons.join("; ")
  }));
  const summary = [
    ["Smart Cull Plan generated", new Date().toISOString().slice(0,19).replace("T", " ")],
    ["Colony data as of", COLONY.as_of],
    [],
    ["Total mice", INV.length],
    ["  Cull", cullRows.length],
    ["  Preserve", preserveRows.length],
    [],
    ["Strain priorities used"],
    ...COLONY.strain_order.map(s => [`  ${s}`, SC.priority[s], `${COLONY.strain_counts[s] || 0} mice`]),
    [],
    ["Settings"],
    ["  Keep-floor (per strain × genotype × sex cell)", SC.keepFloor],
    ["  Auto-preserve rescue donors", SC.rescueOn ? "Yes" : "No"],
    ["  Donors per sex per orphan strain", SC.donorN],
  ];
  const wb = XLSX.utils.book_new();
  const wsCull = XLSX.utils.json_to_sheet(cullRows);
  const wsPreserve = XLSX.utils.json_to_sheet(preserveRows);
  const wsSum = XLSX.utils.aoa_to_sheet(summary);
  [wsCull, wsPreserve].forEach(ws => {
    ws["!cols"] = ["mouse_id","strain","sex","age_months","dob","genotype","cage_id","use","reason"]
      .map(k => ({ wch: k === "genotype" ? 36 : k === "reason" ? 40 : 14 }));
  });
  wsSum["!cols"] = [{ wch: 40 }, { wch: 22 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsCull, "Cull");
  XLSX.utils.book_append_sheet(wb, wsPreserve, "Preserve");
  XLSX.utils.book_append_sheet(wb, wsSum, "Summary");
  XLSX.writeFile(wb, `Smart_Cull_Plan_${COLONY.as_of}.xlsx`);
};

scInit();

// ============= EXPERIMENT =============
$("#exp-strain-list").innerHTML = COLONY.strain_summary.map(s => `
  <li><strong>${s.strain}</strong> — ${s.alive_mice} alive · ${s.active_breedings} active breeding(s) · ${s.active_cages} cages</li>
`).join("");

// ============= MODAL =============
const MODAL_CONTENT = {
  "info": {
    title: "About this dashboard",
    body: `
      <p>This dashboard summarizes the live state of the MacDougald Lab mouse colony, generated from
         the lab's colony software exports.</p>
      <h3>How to navigate</h3>
      <ul>
        <li><strong>Overview</strong> — top-line numbers and demographic charts</li>
        <li><strong>Inventory</strong> — every mouse, filterable and searchable</li>
        <li><strong>Cages</strong> — physical layout of Rack 7A + list of unassigned cages</li>
        <li><strong>Breeding</strong> — active pairs, performance metrics, full pair table</li>
        <li><strong>Cohort Planning</strong> — research cohorts, target genotypes, expected ratios</li>
        <li><strong>Cull Candidates</strong> — filtered cull pools with CSV export</li>
        <li><strong>Walkthrough</strong> — most recent physical inventory reconciliation</li>
        <li><strong>Experiment Details</strong> — colony purpose, methods, husbandry</li>
        <li><strong>Statistics</strong> — descriptive stats on demographics + breeding performance</li>
      </ul>
    ` },
  "summary": {
    title: "What's interesting right now",
    body: `<div id="modal-summary-body">Loading…</div>`,
    onShow: renderInsightSummary,
  },
  "update": {
    title: "How to update this dashboard",
    body: `
      <ol>
        <li>Export the latest data from the colony software into the local <span class="code">Mouse Colony</span> folder, replacing:
          <ul>
            <li><span class="code">CageListExcel.xlsx</span></li>
            <li><span class="code">MacLab - Brian_Mice.xlsx</span></li>
            <li><span class="code">StrainList.xlsx</span></li>
            <li><span class="code">Breedings.xlsx</span></li>
          </ul>
        </li>
        <li>Run from a terminal in that folder:
          <pre class="code" style="display:block; padding:10px; margin-top:6px;">python build_dashboard_data.py
python build_dashboard_html.py</pre>
        </li>
        <li>Commit and push:
          <pre class="code" style="display:block; padding:10px; margin-top:6px;">git add data/colony.json index.html
git commit -m "Update colony data"
git push</pre>
        </li>
        <li>GitHub Pages re-deploys within ~30 seconds. The shared dashboard URL doesn't change.</li>
      </ol>
      <p>To update cohort plans (target genotypes, ratios, purpose), edit
         <span class="code">data/cohort_plans.json</span> directly and rebuild.</p>
    ` },
  "chart-strain": {
    title: "Mice by strain",
    body: `<p>Bar chart counts every mouse in the inventory grouped by strain.
              Strains are sorted by colony size.</p>` },
  "chart-age": {
    title: "Age distribution",
    body: `<p>All mice grouped into 5 age brackets. <em>Aged &gt; 12 mo</em> is the default cull-candidate
              threshold for available, non-breeder mice.</p>` },
  "chart-sex": {
    title: "Sex by strain",
    body: `<p>Stacked counts of Female / Male / Unknown per strain.
              "Unknown" mice need to be sexed at the next walkthrough.</p>` },
  "chart-use": {
    title: "Use status by strain",
    body: `<p>"Available" means the mouse has no project assignment in the colony software.
              "Breeding" means the mouse is in an active breeding cage.</p>` },
  "strain-summary": {
    title: "Strain summary",
    body: `<p>Pulled directly from the colony software's <span class="code">StrainList.xlsx</span>.
              Includes deceased counts as a record-of-history.</p>` },
  "breeding-perf": {
    title: "Breeding pair performance",
    body: `<p>Each pair's pup count (bars) plotted against litter count (line). Pairs with low
              productivity may need to be retired or replaced.</p>` },
  "breeding-age": {
    title: "Breeder age",
    body: `<p>Months elapsed since the breeding cage went active. Older breedings sometimes
              need rotation due to declining productivity.</p>` },
  "cage-grid": {
    title: "Rack 7A layout",
    body: `<p>Visual layout of the rack: rows 1–7 × columns A–J. Click a cage for details.
              The orange left border indicates a breeding cage.</p>` },
  "unassigned": {
    title: "Unassigned cages",
    body: `<p>Cages without a recorded rack position. These need to be located physically and
              their position entered in the colony software.</p>` },
  "smart-cull": {
    title: "How the Smart Cull Plan works",
    body: `
      <h3>Inputs</h3>
      <ul>
        <li><strong>Strain priority</strong> — set per strain to <em>Expand</em> (preserve all),
            <em>Maintain</em> (keep youngest insurance, cull aged surplus), or
            <em>Wind down</em> (cull aggressively, but auto-preserve rescue donors).</li>
        <li><strong>Keep-floor</strong> — for each (strain × genotype × sex) cell, the algorithm
            preserves at least the N youngest mice as cohort-rebuild insurance. Default 4.</li>
        <li><strong>Rescue donors</strong> — for orphan strains that can be regenerated by selecting
            offspring from compound crosses, the algorithm finds 4 donors per sex from other strains
            and preserves them.</li>
      </ul>
      <h3>Built-in rescue rules</h3>
      <ul>
        <li><strong>Adipoq-Cre</strong> rescue donors: any mouse with <span class="id-mono">Adipoq-cre; Dendra2&lt;-/-&gt;</span>
            (no mTmG) — found primarily in AdipoGlo strain.</li>
        <li><strong>Dendra2</strong> rescue donors: any mouse with <span class="id-mono">WT; Dendra2&lt;+/-&gt;</span> or
            <span class="id-mono">WT; Dendra2&lt;+/+&gt;</span> (no Adipoq-Cre, no mTmG).</li>
        <li><strong>mTmG</strong> rescue donors: any mouse with <span class="id-mono">mTmG&lt;mTmG/mTmG&gt;; WT</span>
            (no Adipoq-Cre, no Dendra2 transgene).</li>
      </ul>
      <h3>Always preserved (regardless of strain priority)</h3>
      <ul>
        <li>Active breeders (parsed from the Breedings sheet + use=Breeding flag)</li>
        <li>Sex-unknown mice (until ID'd at the next walkthrough)</li>
        <li>Singleton genotypes — the only mouse of its (strain × genotype) cell —
            unless that strain is rescuable, in which case the rescue rule supersedes the singleton rule</li>
      </ul>
      <h3>Outputs</h3>
      <ul>
        <li>Cull list with the reason for each mouse</li>
        <li>Preserve list with the reason for each mouse</li>
        <li>Recommended actions: breeding setups for Expand strains; rescue-donor manifest for Wind-down strains</li>
        <li>Excel export: 3-sheet workbook (Cull, Preserve, Summary)</li>
      </ul>
    ` },
  "cohort-builder": {
    title: "How the cohort builder works",
    body: `
      <h3>Steps</h3>
      <ol>
        <li><strong>Filter</strong> the colony by strain, sex, genotype text, age range, and use status.</li>
        <li><strong>Add</strong> mice as <em>experimental</em> or <em>control</em>. Mice already in the cohort are tagged.</li>
        <li>Click <strong>"Suggest controls"</strong> — for each experimental mouse, the dashboard surfaces
            up to 5 best-matched controls by:
          <ul>
            <li>Same strain</li>
            <li>Same sex</li>
            <li>Different genotype than the experimental mouse</li>
            <li>Closest age (Δ months shown)</li>
            <li>Bonus: prioritizes genotypes listed as non-target in that strain's cohort plan
                (marked with the <span class="role-tag ctrl">plan</span> tag)</li>
          </ul>
        </li>
        <li>Click <strong>"Export to Excel"</strong> — produces a 2-sheet workbook (Cohort + Summary)
            ready to share or import into Prism / R / Excel directly.</li>
      </ol>
      <h3>Notes</h3>
      <ul>
        <li>Breeders are excluded from control suggestions automatically.</li>
        <li>Mice already in the cohort cannot be re-suggested.</li>
        <li>The Δ age is shown in months; differences under 1 month are highlighted green.</li>
      </ul>
    ` },
};

function renderInsightSummary() {
  const s = COLONY.summary;
  const insights = [];
  if (s.unassigned_cages > 50) insights.push(`<li><strong>${s.unassigned_cages}</strong> cages have no rack location — biggest inventory drift to fix.</li>`);
  if (s.aged_over_12mo > 100) insights.push(`<li><strong>${s.aged_over_12mo}</strong> mice are over 12 months — large cull pool available.</li>`);
  if (s.aged_over_18mo > 0) insights.push(`<li><strong>${s.aged_over_18mo}</strong> mice are over 18 months — review individually.</li>`);
  if (s.unknown_sex > 0) insights.push(`<li><strong>${s.unknown_sex}</strong> mice have unknown sex — ID at next walkthrough.</li>`);
  const top = COLONY.strain_summary.slice().sort((a,b) => (b.alive_mice||0)-(a.alive_mice||0))[0];
  if (top) insights.push(`<li>Largest strain: <strong>${top.strain}</strong> (${top.alive_mice} alive, ${top.active_breedings} active breeding(s)).</li>`);
  const perfBR = BR.filter(b => b.months_active && b.born).map(b => ({i: BR.indexOf(b), prod: b.born / b.months_active}));
  perfBR.sort((a,b) => b.prod - a.prod);
  if (perfBR.length) insights.push(`<li>Most productive pair: <strong>#${perfBR[0].i+1}</strong> at ${perfBR[0].prod.toFixed(1)} pups/mo.</li>`);
  if (perfBR.length > 1) insights.push(`<li>Least productive active pair: <strong>#${perfBR[perfBR.length-1].i+1}</strong> at ${perfBR[perfBR.length-1].prod.toFixed(1)} pups/mo — consider rotation.</li>`);
  $("#modal-summary-body").innerHTML = `<ul>${insights.join("")}</ul>`;
}

function openModal(key, title, body) {
  if (key && MODAL_CONTENT[key]) {
    const m = MODAL_CONTENT[key];
    $("#modal-title").textContent = m.title;
    $("#modal-body").innerHTML = m.body;
    if (m.onShow) m.onShow();
  } else {
    $("#modal-title").textContent = title || "";
    $("#modal-body").innerHTML = body || "";
  }
  $("#modal-bg").classList.add("show");
}
function closeModal() { $("#modal-bg").classList.remove("show"); }
function closeModalIfBg(e) { if (e.target.id === "modal-bg") closeModal(); }
window.openModal = openModal;
window.closeModal = closeModal;
window.closeModalIfBg = closeModalIfBg;

// Resize charts on window resize
window.addEventListener("resize", () => {
  Object.values(charts).forEach(c => c.resize());
});

// Initial chart resize after layout settles
setTimeout(() => Object.values(charts).forEach(c => c.resize()), 100);
</script>
</body>
</html>
"""

OUT.write_text(HTML.replace("__COLONY_JSON__", colony_json_str))
size_kb = OUT.stat().st_size / 1024
print(f"Wrote {OUT}  ({size_kb:.0f} KB)")
