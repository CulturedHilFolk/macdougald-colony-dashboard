// Build cull-summary PowerPoint deck.
// Style: matches the live dashboard — navy + slate + accent blue, Cambria/Calibri.
//
// Reads:  /tmp/ppt_data.json
// Writes: Mouse_Colony_Cull_Summary_<DATE>.pptx

const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const DATA = JSON.parse(fs.readFileSync("/tmp/ppt_data.json", "utf8"));
const OUT_DIR = "/Users/briandes/Library/CloudStorage/OneDrive-MichiganMedicine/Desktop/MacDougald Lab/Mouse Colony";
const TODAY = DATA.as_of;

// ---- Color palette ----
const C = {
  navy: "1E3A8A",         // brand
  navyDark: "0F172A",     // ink dark
  slate: "475569",        // body text
  slateLight: "94A3B8",   // muted
  slateBg: "F8FAFC",      // light bg
  white: "FFFFFF",
  accent: "2563EB",       // accent blue
  ok: "15803D",
  okSoft: "DCFCE7",
  warn: "B45309",
  warnSoft: "FEF3C7",
  danger: "B91C1C",
  dangerSoft: "FEE2E2",
  brandSoft: "E0E7FF",
  line: "E5E9F0",
  // Strain colors (matching dashboard)
  AdipoGlo: "1E40AF",
  "AdipoGlo+": "7C3AED",
  Dendra2: "0D9488",
  "Marrow Glo": "CA8A04",
  mTmG: "DB2777",
  "Adipoq-Cre": "64748B",
};

const FONT_HEADER = "Cambria";
const FONT_BODY = "Calibri";

const fmt = n => new Intl.NumberFormat("en-US").format(n);
const fmt$ = n => "$" + new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(Math.round(n));
const fmt$2 = n => "$" + new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";  // 13.3 × 7.5
pres.author = "Brian Desrosiers · MacDougald Lab";
pres.title = "Mouse Colony Cull Summary — " + TODAY;

const W = 13.3, H = 7.5;

// ---- Helpers ----
function addPageHeader(slide, title, accent="") {
  // Dark slim header bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.55,
    fill: { color: C.navy }, line: { type: "none" },
  });
  slide.addText(title, {
    x: 0.5, y: 0.07, w: W - 1, h: 0.42,
    fontFace: FONT_HEADER, fontSize: 18, bold: true, color: C.white,
    align: "left", valign: "middle", margin: 0,
  });
  if (accent) {
    slide.addText(accent, {
      x: W - 4, y: 0.07, w: 3.5, h: 0.42,
      fontFace: FONT_BODY, fontSize: 11, color: "BFDBFE",
      align: "right", valign: "middle", italic: true, margin: 0,
    });
  }
  // Footer
  slide.addText(`MacDougald Lab · Brian Desrosiers · ${TODAY}`, {
    x: 0.5, y: H - 0.35, w: W - 1, h: 0.25,
    fontFace: FONT_BODY, fontSize: 9, color: C.slateLight,
    align: "left", valign: "middle", margin: 0,
  });
}

function addStatCard(slide, opts) {
  const { x, y, w, h, value, label, sub, valueColor, accentColor } = opts;
  // Card
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white }, line: { color: C.line, width: 0.75 },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 1, angle: 90, opacity: 0.06 },
  });
  // Optional thin colored top bar (motif)
  if (accentColor) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 0.08,
      fill: { color: accentColor }, line: { type: "none" },
    });
  }
  // Label
  slide.addText(label, {
    x: x + 0.2, y: y + 0.18, w: w - 0.4, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, bold: true, color: C.slate,
    align: "left", valign: "top", charSpacing: 1, margin: 0,
  });
  // Value
  slide.addText(value, {
    x: x + 0.2, y: y + 0.55, w: w - 0.4, h: h - 1.0,
    fontFace: FONT_HEADER, fontSize: 40, bold: true, color: valueColor || C.navyDark,
    align: "left", valign: "middle", margin: 0,
  });
  // Sub
  if (sub) {
    slide.addText(sub, {
      x: x + 0.2, y: y + h - 0.5, w: w - 0.4, h: 0.3,
      fontFace: FONT_BODY, fontSize: 10, color: C.slate,
      align: "left", valign: "middle", italic: true, margin: 0,
    });
  }
}

// =========================================================
// Slide 1 — Title
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navyDark };

  // Side accent strip (motif)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.4, h: H,
    fill: { color: C.accent }, line: { type: "none" },
  });

  // Tag
  s.addText("COLONY MANAGEMENT REPORT", {
    x: 0.9, y: 1.6, w: 8, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, bold: true, color: "BFDBFE",
    charSpacing: 4, align: "left", margin: 0,
  });

  // Title
  s.addText("Mouse Colony Cull", {
    x: 0.9, y: 2.0, w: 11.5, h: 1.2,
    fontFace: FONT_HEADER, fontSize: 56, bold: true, color: C.white,
    align: "left", valign: "top", margin: 0,
  });
  s.addText("Pre/post inventory · cost analysis · cohort outlook", {
    x: 0.9, y: 3.2, w: 11.5, h: 0.55,
    fontFace: FONT_HEADER, fontSize: 22, italic: true, color: "CBD5E1",
    align: "left", margin: 0,
  });

  // Hero stat
  s.addText("173", {
    x: 0.9, y: 4.4, w: 4, h: 1.6,
    fontFace: FONT_HEADER, fontSize: 110, bold: true, color: C.accent,
    align: "left", valign: "middle", margin: 0,
  });
  s.addText("mice culled today", {
    x: 0.9, y: 5.95, w: 5, h: 0.45,
    fontFace: FONT_BODY, fontSize: 16, color: "CBD5E1",
    align: "left", italic: true, margin: 0,
  });

  // Right side: secondary numbers stacked
  const yBase = 4.4;
  [
    { v: `${DATA.pre_total} → ${DATA.post_total}`, l: "colony size" },
    { v: fmt$(DATA.cost_savings), l: "annual savings" },
    { v: `${DATA.cost_savings_pct}%`, l: "cost reduction" },
  ].forEach((d, i) => {
    s.addText(d.v, {
      x: 6.2, y: yBase + i * 0.6, w: 6, h: 0.55,
      fontFace: FONT_HEADER, fontSize: 28, bold: true, color: C.white,
      align: "left", margin: 0,
    });
    s.addText(d.l, {
      x: 9.0, y: yBase + i * 0.6 + 0.12, w: 4, h: 0.4,
      fontFace: FONT_BODY, fontSize: 13, color: "CBD5E1",
      align: "left", italic: true, margin: 0,
    });
  });

  // Author block
  s.addText([
    { text: "Brian Desrosiers", options: { bold: true, breakLine: true } },
    { text: "MacDougald Lab · University of Michigan", options: {} },
  ], {
    x: 0.9, y: H - 1.05, w: 8, h: 0.6,
    fontFace: FONT_BODY, fontSize: 13, color: "CBD5E1",
    align: "left", margin: 0,
  });

  s.addText(TODAY, {
    x: W - 2.5, y: H - 0.8, w: 2, h: 0.4,
    fontFace: FONT_BODY, fontSize: 13, color: "BFDBFE",
    align: "right", italic: true, margin: 0,
  });
}

// =========================================================
// Slide 2 — Executive Summary (4 big stat cards)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "Executive summary", "Headline numbers");

  s.addText("Today's cull was a big inventory cleanup with a meaningful recurring cost reduction.", {
    x: 0.6, y: 0.85, w: W - 1.2, h: 0.4,
    fontFace: FONT_BODY, fontSize: 14, color: C.slate, italic: true,
    align: "left", margin: 0,
  });

  const cardY = 1.6, cardH = 2.1;
  const cardW = (W - 1.2 - 0.6) / 4;
  const cards = [
    { value: fmt(DATA.killed),         label: "MICE CULLED",       sub: `out of ${DATA.pre_total} pre-cull`, accentColor: C.danger,  valueColor: C.danger },
    { value: fmt(DATA.post_total),     label: "MICE REMAINING",    sub: `${Math.round(100*(DATA.pre_total-DATA.post_total)/DATA.pre_total)}% reduction`, accentColor: C.accent,  valueColor: C.navyDark },
    { value: fmt$(DATA.cost_savings),  label: "ANNUAL SAVINGS",    sub: `was ${fmt$(DATA.cost_pre)} · now ${fmt$(DATA.cost_post)}`, accentColor: C.ok, valueColor: C.ok },
    { value: `${DATA.cost_savings_pct}%`, label: "COST REDUCTION", sub: "recurring; renews each fiscal year", accentColor: C.ok,  valueColor: C.ok },
  ];
  cards.forEach((c, i) => {
    addStatCard(s, {
      x: 0.6 + i * (cardW + 0.2), y: cardY, w: cardW, h: cardH,
      ...c,
    });
  });

  // Below cards: framework summary box
  const boxY = cardY + cardH + 0.4;
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: boxY, w: W - 1.2, h: 1.7,
    fill: { color: C.white }, line: { color: C.line, width: 0.75 },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 1, angle: 90, opacity: 0.06 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: boxY, w: 0.08, h: 1.7,
    fill: { color: C.navy }, line: { type: "none" },
  });
  s.addText("Strategy", {
    x: 0.85, y: boxY + 0.18, w: 6, h: 0.35,
    fontFace: FONT_HEADER, fontSize: 14, bold: true, color: C.navy,
    align: "left", margin: 0,
  });
  s.addText([
    { text: "Marrow Glo", options: { bold: true, color: C.warn } },
    { text: " preserved entirely (priority strain to expand) · ", options: { color: C.slate } },
    { text: "Adipoq-Cre, mTmG, Dendra2", options: { bold: true, color: C.warn } },
    { text: " wound down — recoverable from compound-strain donors · ", options: { color: C.slate } },
    { text: "AdipoGlo, AdipoGlo+", options: { bold: true, color: C.accent } },
    { text: " maintained with cohort-rebuild insurance.", options: { color: C.slate } },
  ], {
    x: 0.85, y: boxY + 0.55, w: W - 1.6, h: 0.7,
    fontFace: FONT_BODY, fontSize: 12, color: C.slate,
    align: "left", valign: "top", margin: 0,
  });
  s.addText([
    { text: "19 ", options: { bold: true, color: C.ok } },
    { text: "rescue donors auto-preserved across compound strains. ", options: { color: C.slate } },
    { text: "26 ", options: { bold: true, color: C.warn } },
    { text: "mice on the cull list could not be located physically and remain to be reconciled.", options: { color: C.slate } },
  ], {
    x: 0.85, y: boxY + 1.15, w: W - 1.6, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, color: C.slate,
    align: "left", margin: 0,
  });
}

// =========================================================
// Slide 3 — Strain priority framework
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "Strain priorities — the framework", "How cull decisions were made");

  const items = [
    { strain: "Marrow Glo",  pre: DATA.pre_strain["Marrow Glo"]||0, post: DATA.post_strain["Marrow Glo"]||0,
      tag: "EXPAND",   tagColor: C.warn, why: "Brian's research priority. Preserve all mice; new breeding pair recommended."  },
    { strain: "AdipoGlo",    pre: DATA.pre_strain["AdipoGlo"]||0,    post: DATA.post_strain["AdipoGlo"]||0,
      tag: "MAINTAIN", tagColor: C.accent, why: "Largest strain; cull aged surplus while keeping ≥4 youngest per (genotype × sex) for cohort rebuilds." },
    { strain: "AdipoGlo+",   pre: DATA.pre_strain["AdipoGlo+"]||0,   post: DATA.post_strain["AdipoGlo+"]||0,
      tag: "MAINTAIN", tagColor: C.accent, why: "Triple-reporter line; same cohort-insurance rule applied per genotype × sex cell." },
    { strain: "Dendra2",     pre: DATA.pre_strain["Dendra2"]||0,     post: DATA.post_strain["Dendra2"]||0,
      tag: "WIND DOWN", tagColor: C.danger, why: "Aging out; rescuable from AdipoGlo offspring carrying Dendra2 alone." },
    { strain: "mTmG",        pre: DATA.pre_strain["mTmG"]||0,        post: DATA.post_strain["mTmG"]||0,
      tag: "WIND DOWN", tagColor: C.danger, why: "Rescuable from AdipoGlo+ offspring with mTmG hom + WT Cre + null Dendra2." },
    { strain: "Adipoq-Cre",  pre: DATA.pre_strain["Adipoq-Cre"]||0,  post: DATA.post_strain["Adipoq-Cre"]||0,
      tag: "WIND DOWN", tagColor: C.danger, why: "Rescuable from AdipoGlo Adipoq-cre; Dendra2<-/-> donors." },
  ];

  const startY = 1.05;
  const rowH = 0.92;
  // Headers
  const cols = [
    { x: 0.6,  w: 1.6, label: "STRAIN",        align: "left" },
    { x: 2.3,  w: 1.5, label: "PRIORITY",      align: "center" },
    { x: 3.9,  w: 0.8, label: "PRE",           align: "right" },
    { x: 4.8,  w: 0.5, label: "→",             align: "center" },
    { x: 5.4,  w: 0.8, label: "POST",          align: "right" },
    { x: 6.4,  w: 6.3, label: "RATIONALE",     align: "left" },
  ];
  cols.forEach(c => {
    s.addText(c.label, {
      x: c.x, y: startY, w: c.w, h: 0.32,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      align: c.align, charSpacing: 2, margin: 0,
    });
  });
  // Header underline
  s.addShape(pres.shapes.LINE, {
    x: 0.6, y: startY + 0.34, w: W - 1.2, h: 0,
    line: { color: C.line, width: 1 },
  });

  items.forEach((it, i) => {
    const rY = startY + 0.5 + i * rowH;
    if (i % 2 === 0) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.55, y: rY - 0.05, w: W - 1.1, h: rowH - 0.1,
        fill: { color: C.white }, line: { type: "none" },
      });
    }
    // Strain (with color dot)
    s.addShape(pres.shapes.OVAL, {
      x: 0.6, y: rY + 0.15, w: 0.18, h: 0.18,
      fill: { color: C[it.strain] || C.slate }, line: { type: "none" },
    });
    s.addText(it.strain, {
      x: 0.85, y: rY, w: cols[0].w - 0.3, h: 0.5,
      fontFace: FONT_HEADER, fontSize: 14, bold: true, color: C.navyDark,
      align: "left", valign: "middle", margin: 0,
    });
    // Priority pill
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 2.4, y: rY + 0.1, w: 1.3, h: 0.32,
      fill: { color: it.tagColor }, line: { type: "none" },
      rectRadius: 0.05,
    });
    s.addText(it.tag, {
      x: 2.4, y: rY + 0.11, w: 1.3, h: 0.3,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.white,
      align: "center", valign: "middle", charSpacing: 2, margin: 0,
    });
    // Pre / arrow / post
    s.addText(String(it.pre), {
      x: cols[2].x, y: rY, w: cols[2].w, h: 0.5,
      fontFace: FONT_HEADER, fontSize: 18, bold: true, color: C.slate,
      align: "right", valign: "middle", margin: 0,
    });
    s.addText("→", {
      x: cols[3].x, y: rY, w: cols[3].w, h: 0.5,
      fontFace: FONT_BODY, fontSize: 16, color: C.slateLight,
      align: "center", valign: "middle", margin: 0,
    });
    const postColor = it.post < it.pre ? (it.post === 0 ? C.danger : C.warn) : C.ok;
    s.addText(String(it.post), {
      x: cols[4].x, y: rY, w: cols[4].w, h: 0.5,
      fontFace: FONT_HEADER, fontSize: 18, bold: true, color: postColor,
      align: "right", valign: "middle", margin: 0,
    });
    // Rationale
    s.addText(it.why, {
      x: cols[5].x, y: rY, w: cols[5].w, h: 0.55,
      fontFace: FONT_BODY, fontSize: 11, color: C.slate,
      align: "left", valign: "middle", margin: 0,
    });
  });
}

// =========================================================
// Slide 4 — Cost analysis
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "Cage cost — before & after", "ULAM CY26 ventilated rate · effective 04/01/2026");

  // Big savings number left side
  s.addText("ANNUAL SAVINGS", {
    x: 0.6, y: 1.0, w: 6, h: 0.32,
    fontFace: FONT_BODY, fontSize: 11, bold: true, color: C.slate,
    charSpacing: 3, margin: 0,
  });
  s.addText(fmt$(DATA.cost_savings), {
    x: 0.6, y: 1.35, w: 6, h: 1.7,
    fontFace: FONT_HEADER, fontSize: 96, bold: true, color: C.ok,
    align: "left", valign: "middle", margin: 0,
  });
  s.addText(`${DATA.cost_savings_pct}% reduction in monthly recurring cost`, {
    x: 0.6, y: 3.1, w: 6, h: 0.45,
    fontFace: FONT_BODY, fontSize: 14, italic: true, color: C.slate,
    align: "left", margin: 0,
  });

  // Three time-horizon stat cards on the right
  const cardX = 7.0, cardW = 5.7, cardH = 1.55;
  const horizons = [
    { label: "PER DAY",   pre: DATA.cost_per_day_pre,   post: DATA.cost_per_day_post },
    { label: "PER MONTH", pre: DATA.cost_per_month_pre, post: DATA.cost_per_month_post },
    { label: "PER YEAR",  pre: DATA.cost_pre,           post: DATA.cost_post },
  ];
  horizons.forEach((h, i) => {
    const y = 1.0 + i * (cardH + 0.15);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cardX, y, w: cardW, h: cardH,
      fill: { color: C.white }, line: { color: C.line, width: 0.75 },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 1, angle: 90, opacity: 0.05 },
    });
    s.addText(h.label, {
      x: cardX + 0.25, y: y + 0.12, w: 2.5, h: 0.3,
      fontFace: FONT_BODY, fontSize: 10, bold: true, color: C.slateLight,
      charSpacing: 2, margin: 0,
    });
    s.addText("Before:", {
      x: cardX + 0.25, y: y + 0.5, w: 1.4, h: 0.3,
      fontFace: FONT_BODY, fontSize: 11, color: C.slate, margin: 0,
    });
    s.addText(i === 0 ? fmt$2(h.pre) : fmt$(h.pre), {
      x: cardX + 1.6, y: y + 0.45, w: 1.6, h: 0.4,
      fontFace: FONT_HEADER, fontSize: 18, bold: true, color: C.slate,
      align: "left", margin: 0,
    });
    s.addText("After:", {
      x: cardX + 0.25, y: y + 0.85, w: 1.4, h: 0.3,
      fontFace: FONT_BODY, fontSize: 11, bold: true, color: C.navyDark, margin: 0,
    });
    s.addText(i === 0 ? fmt$2(h.post) : fmt$(h.post), {
      x: cardX + 1.6, y: y + 0.8, w: 1.6, h: 0.4,
      fontFace: FONT_HEADER, fontSize: 22, bold: true, color: C.ok,
      align: "left", margin: 0,
    });
    // Delta
    const delta = h.pre - h.post;
    s.addText(`-${i === 0 ? fmt$2(delta) : fmt$(delta)}`, {
      x: cardX + 3.5, y: y + 0.6, w: 2, h: 0.5,
      fontFace: FONT_HEADER, fontSize: 22, bold: true, color: C.ok,
      align: "right", valign: "middle", margin: 0,
    });
    s.addText("savings", {
      x: cardX + 3.5, y: y + 1.0, w: 2, h: 0.3,
      fontFace: FONT_BODY, fontSize: 9, italic: true, color: C.slateLight,
      align: "right", margin: 0,
    });
  });

  // Bottom note
  s.addText([
    { text: "Methodology: ", options: { bold: true } },
    { text: "$1.19/day per ventilated cage, $2.16/day per breeding-vented cage. Post-cull projection assumes cages with zero surviving mice are returned to ULAM and stop billing. Husbandry tech labor is billed separately and is not included.", options: {} },
  ], {
    x: 0.6, y: 6.4, w: W - 1.2, h: 0.55,
    fontFace: FONT_BODY, fontSize: 10, color: C.slate, italic: true,
    align: "left", margin: 0,
  });
}

// =========================================================
// Slide 5 — Pre vs Post by strain (chart)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "Inventory — pre vs post-cull, by strain", `${DATA.pre_total} → ${DATA.post_total} mice`);

  const strainOrder = ["AdipoGlo", "AdipoGlo+", "Dendra2", "Marrow Glo", "mTmG", "Adipoq-Cre"];
  const chartData = [
    { name: "Pre-cull",  labels: strainOrder, values: strainOrder.map(s => DATA.pre_strain[s] || 0) },
    { name: "Post-cull", labels: strainOrder, values: strainOrder.map(s => DATA.post_strain[s] || 0) },
  ];

  s.addChart(pres.charts.BAR, chartData, {
    x: 0.6, y: 1.0, w: 8.5, h: 5.5,
    chartColors: [C.slateLight, C.accent],
    barDir: "col",
    barGapWidthPct: 40,
    catAxisLabelFontSize: 12,
    catAxisLabelFontFace: FONT_BODY,
    catAxisLabelColor: C.slate,
    valAxisLabelFontSize: 11,
    valAxisLabelFontFace: FONT_BODY,
    valAxisLabelColor: C.slateLight,
    showLegend: true,
    legendPos: "t",
    legendFontFace: FONT_BODY,
    legendFontSize: 11,
    showValue: true,
    dataLabelFontFace: FONT_BODY,
    dataLabelFontSize: 9,
    dataLabelColor: C.navyDark,
    dataLabelPosition: "outEnd",
    plotArea: { fill: { color: "FFFFFF" } },
  });

  // Right side: callouts
  const calloutX = 9.4;
  const calloutW = 3.4;
  const callouts = [
    { strain: "AdipoGlo",   delta: (DATA.post_strain["AdipoGlo"]||0) - (DATA.pre_strain["AdipoGlo"]||0),   action: "Maintained — aged surplus removed" },
    { strain: "Dendra2",    delta: (DATA.post_strain["Dendra2"]||0) - (DATA.pre_strain["Dendra2"]||0),    action: "Wound down — restart from rescue donors" },
    { strain: "mTmG",       delta: (DATA.post_strain["mTmG"]||0) - (DATA.pre_strain["mTmG"]||0),       action: "Wound down — restart from rescue donors" },
    { strain: "Marrow Glo", delta: (DATA.post_strain["Marrow Glo"]||0) - (DATA.pre_strain["Marrow Glo"]||0), action: "Preserved entirely; expansion plan recommended" },
  ];
  s.addText("Strain notes", {
    x: calloutX, y: 1.0, w: calloutW, h: 0.4,
    fontFace: FONT_HEADER, fontSize: 14, bold: true, color: C.navy,
    margin: 0,
  });
  callouts.forEach((c, i) => {
    const yC = 1.5 + i * 1.15;
    s.addShape(pres.shapes.RECTANGLE, {
      x: calloutX, y: yC, w: calloutW, h: 1.0,
      fill: { color: C.white }, line: { color: C.line, width: 0.75 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: calloutX, y: yC, w: 0.06, h: 1.0,
      fill: { color: C[c.strain] || C.slate }, line: { type: "none" },
    });
    s.addText(c.strain, {
      x: calloutX + 0.2, y: yC + 0.12, w: calloutW - 0.4, h: 0.3,
      fontFace: FONT_HEADER, fontSize: 12, bold: true, color: C.navyDark,
      margin: 0,
    });
    const deltaStr = c.delta >= 0 ? `+${c.delta}` : `${c.delta}`;
    const deltaColor = c.delta < 0 ? C.danger : c.delta > 0 ? C.ok : C.slate;
    s.addText(`${deltaStr} mice`, {
      x: calloutX + 0.2, y: yC + 0.38, w: calloutW - 0.4, h: 0.3,
      fontFace: FONT_HEADER, fontSize: 14, bold: true, color: deltaColor,
      margin: 0,
    });
    s.addText(c.action, {
      x: calloutX + 0.2, y: yC + 0.65, w: calloutW - 0.4, h: 0.3,
      fontFace: FONT_BODY, fontSize: 9, italic: true, color: C.slate,
      margin: 0,
    });
  });
}

// =========================================================
// Slide 6 — Rescue donor strategy
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "Rescue donors — preserving wound-down strains", `${DATA.rescue_donors_count} donors flagged across compound strains`);

  s.addText("Adipoq-Cre, mTmG, and Dendra2 standalone strains were aged out — but each can be regenerated from offspring of compound strains where the target transgene is present alone.", {
    x: 0.6, y: 0.95, w: W - 1.2, h: 0.6,
    fontFace: FONT_BODY, fontSize: 13, color: C.slate, italic: true,
    align: "left", margin: 0,
  });

  const rescues = [
    { target: "Adipoq-Cre",  donors: "Adipoq-cre; Dendra2<-/->",   src: "AdipoGlo offspring", note: "Cre present, no other transgene." },
    { target: "Dendra2",     donors: "WT; Dendra2<+/+ or +/->",     src: "AdipoGlo offspring", note: "Dendra2 alone, no Cre." },
    { target: "mTmG",        donors: "mTmG<mTmG/mTmG>; WT",         src: "AdipoGlo+ offspring", note: "mTmG hom alone, no Cre, no Dendra2." },
  ];
  const cardX = 0.6, cardY = 1.95, cardW = (W - 1.2 - 0.4) / 3, cardH = 4.6;
  rescues.forEach((r, i) => {
    const x = cardX + i * (cardW + 0.2);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: cardH,
      fill: { color: C.white }, line: { color: C.line, width: 0.75 },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 1, angle: 90, opacity: 0.06 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: 0.4,
      fill: { color: C.warn }, line: { type: "none" },
    });
    s.addText(`Rescue: ${r.target}`, {
      x: x + 0.25, y: cardY + 0.05, w: cardW - 0.5, h: 0.32,
      fontFace: FONT_BODY, fontSize: 11, bold: true, color: C.white,
      charSpacing: 1, margin: 0,
    });
    s.addText("Donor genotype", {
      x: x + 0.25, y: cardY + 0.7, w: cardW - 0.5, h: 0.3,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      charSpacing: 1, margin: 0,
    });
    s.addText(r.donors, {
      x: x + 0.25, y: cardY + 1.0, w: cardW - 0.5, h: 0.55,
      fontFace: "Consolas", fontSize: 12, color: C.navyDark, bold: true,
      align: "left", margin: 0,
    });
    s.addText("Source", {
      x: x + 0.25, y: cardY + 1.7, w: cardW - 0.5, h: 0.3,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      charSpacing: 1, margin: 0,
    });
    s.addText(r.src, {
      x: x + 0.25, y: cardY + 2.0, w: cardW - 0.5, h: 0.4,
      fontFace: FONT_BODY, fontSize: 13, color: C.navyDark,
      align: "left", margin: 0,
    });
    s.addText(r.note, {
      x: x + 0.25, y: cardY + 2.55, w: cardW - 0.5, h: 0.7,
      fontFace: FONT_BODY, fontSize: 11, color: C.slate, italic: true,
      align: "left", margin: 0,
    });
    // Donor count placeholder (we track 4 per sex per strain by default)
    s.addText("PRESERVED THIS CULL", {
      x: x + 0.25, y: cardY + 3.4, w: cardW - 0.5, h: 0.3,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      charSpacing: 1, margin: 0,
    });
    s.addText("up to 4 ♀ + 4 ♂", {
      x: x + 0.25, y: cardY + 3.7, w: cardW - 0.5, h: 0.5,
      fontFace: FONT_HEADER, fontSize: 18, bold: true, color: C.ok,
      align: "left", margin: 0,
    });
    s.addText("youngest in cell", {
      x: x + 0.25, y: cardY + 4.2, w: cardW - 0.5, h: 0.3,
      fontFace: FONT_BODY, fontSize: 9, italic: true, color: C.slate,
      margin: 0,
    });
  });
}

// =========================================================
// Slide 7 — Available cohorts (post-cull)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "What experiments are still possible", "Available cohort cells (post-cull)");

  s.addText("Top (strain × genotype × sex) cells with ≥3 surviving mice — these are the experimental groups you can still assemble without new breeding.", {
    x: 0.6, y: 0.95, w: W - 1.2, h: 0.5,
    fontFace: FONT_BODY, fontSize: 12, color: C.slate, italic: true,
    margin: 0,
  });

  const top = (DATA.cohort_cells_top || []).slice(0, 8);
  const startY = 1.7;
  const rowH = 0.55;
  // Header
  const hdr = [
    { x: 0.6, w: 1.4,  label: "STRAIN" },
    { x: 2.05, w: 4.5, label: "GENOTYPE" },
    { x: 6.6, w: 0.7,  label: "SEX" },
    { x: 7.4, w: 0.8,  label: "n" },
    { x: 8.3, w: 1.4,  label: "MEAN AGE" },
    { x: 9.8, w: 1.3,  label: "MIN" },
    { x: 11.2, w: 1.3, label: "MAX" },
  ];
  hdr.forEach(c => {
    s.addText(c.label, {
      x: c.x, y: startY, w: c.w, h: 0.32,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      align: "left", charSpacing: 2, margin: 0,
    });
  });
  s.addShape(pres.shapes.LINE, {
    x: 0.6, y: startY + 0.32, w: W - 1.2, h: 0,
    line: { color: C.line, width: 1 },
  });

  top.forEach((row, i) => {
    const y = startY + 0.45 + i * rowH;
    if (i % 2 === 0) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.55, y: y - 0.05, w: W - 1.1, h: rowH - 0.05,
        fill: { color: C.white }, line: { type: "none" },
      });
    }
    // Strain pill
    s.addShape(pres.shapes.OVAL, {
      x: 0.6, y: y + 0.13, w: 0.18, h: 0.18,
      fill: { color: C[row.strain] || C.slate }, line: { type: "none" },
    });
    s.addText(row.strain, {
      x: 0.85, y, w: hdr[0].w - 0.3, h: rowH,
      fontFace: FONT_BODY, fontSize: 11, bold: true, color: C.navyDark,
      align: "left", valign: "middle", margin: 0,
    });
    s.addText(row.genotype, {
      x: hdr[1].x, y, w: hdr[1].w, h: rowH,
      fontFace: "Consolas", fontSize: 9, color: C.slate,
      align: "left", valign: "middle", margin: 0,
    });
    s.addText(row.sex === "Female" ? "♀" : "♂", {
      x: hdr[2].x, y, w: hdr[2].w, h: rowH,
      fontFace: FONT_BODY, fontSize: 16, color: row.sex === "Female" ? "DB2777" : "2563EB",
      align: "left", valign: "middle", bold: true, margin: 0,
    });
    s.addText(String(row.n), {
      x: hdr[3].x, y, w: hdr[3].w, h: rowH,
      fontFace: FONT_HEADER, fontSize: 14, bold: true, color: C.navyDark,
      align: "left", valign: "middle", margin: 0,
    });
    s.addText(row.mean_age != null ? `${row.mean_age} mo` : "—", {
      x: hdr[4].x, y, w: hdr[4].w, h: rowH,
      fontFace: FONT_BODY, fontSize: 11, color: C.slate,
      align: "left", valign: "middle", margin: 0,
    });
    s.addText(row.min_age != null ? `${row.min_age}` : "—", {
      x: hdr[5].x, y, w: hdr[5].w, h: rowH,
      fontFace: FONT_BODY, fontSize: 11, color: C.slate,
      align: "left", valign: "middle", margin: 0,
    });
    s.addText(row.max_age != null ? `${row.max_age}` : "—", {
      x: hdr[6].x, y, w: hdr[6].w, h: rowH,
      fontFace: FONT_BODY, fontSize: 11, color: C.slate,
      align: "left", valign: "middle", margin: 0,
    });
  });
}

// =========================================================
// Slide 8 — Cohort plans status
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "Active cohort plans", "Research goals + currently available toward target");

  const plans = (DATA.cohort_plans || []).slice(0, 3);
  const cardW = (W - 1.2 - 0.4) / plans.length;
  const cardY = 1.0, cardH = 5.6;

  plans.forEach((p, i) => {
    const x = 0.6 + i * (cardW + 0.2);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: cardH,
      fill: { color: C.white }, line: { color: C.line, width: 0.75 },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 1, angle: 90, opacity: 0.06 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: 0.5,
      fill: { color: C[p.strain] || C.navy }, line: { type: "none" },
    });
    s.addText(p.strain, {
      x: x + 0.25, y: cardY + 0.08, w: cardW - 0.5, h: 0.4,
      fontFace: FONT_HEADER, fontSize: 18, bold: true, color: C.white,
      align: "left", margin: 0,
    });
    // Cross
    s.addText("CROSS", {
      x: x + 0.25, y: cardY + 0.65, w: cardW - 0.5, h: 0.25,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      charSpacing: 2, margin: 0,
    });
    s.addText(p.cross_desc || "—", {
      x: x + 0.25, y: cardY + 0.9, w: cardW - 0.5, h: 0.7,
      fontFace: "Consolas", fontSize: 10, color: C.navyDark,
      align: "left", margin: 0,
    });
    // Purpose
    s.addText("PURPOSE", {
      x: x + 0.25, y: cardY + 1.7, w: cardW - 0.5, h: 0.25,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      charSpacing: 2, margin: 0,
    });
    s.addText(p.purpose || "—", {
      x: x + 0.25, y: cardY + 1.95, w: cardW - 0.5, h: 1.2,
      fontFace: FONT_BODY, fontSize: 11, color: C.slate, italic: true,
      align: "left", margin: 0,
    });
    // Target
    s.addText("EXPERIMENTAL TARGET", {
      x: x + 0.25, y: cardY + 3.25, w: cardW - 0.5, h: 0.25,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
      charSpacing: 2, margin: 0,
    });
    s.addText(p.target_geno || "—", {
      x: x + 0.25, y: cardY + 3.5, w: cardW - 0.5, h: 0.45,
      fontFace: "Consolas", fontSize: 10, color: C.ok, bold: true,
      align: "left", margin: 0,
    });
    // Currently available
    if (p.currently_available != null) {
      s.addText("AVAILABLE TOWARD TARGET", {
        x: x + 0.25, y: cardY + 4.15, w: cardW - 0.5, h: 0.25,
        fontFace: FONT_BODY, fontSize: 9, bold: true, color: C.slateLight,
        charSpacing: 2, margin: 0,
      });
      s.addText(String(p.currently_available), {
        x: x + 0.25, y: cardY + 4.45, w: cardW - 0.5, h: 0.7,
        fontFace: FONT_HEADER, fontSize: 36, bold: true, color: C.ok,
        align: "left", margin: 0,
      });
      s.addText("mice (pre-cull data)", {
        x: x + 0.25, y: cardY + 5.15, w: cardW - 0.5, h: 0.3,
        fontFace: FONT_BODY, fontSize: 9, italic: true, color: C.slate,
        margin: 0,
      });
    }
  });
}

// =========================================================
// Slide 9 — Open items / next steps
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.slateBg };
  addPageHeader(s, "Open items and next steps", "What's left after today's cull");

  const items = [
    {
      icon: "🔍", color: C.warn,
      title: "Walkthrough — locate 26 unfound mice",
      detail: "26 mice on the cull list could not be physically located today. They are still recorded as Alive in the colony software. Tomorrow's walkthrough should account for each — confirm dead, locate in unlabeled cage, or flag as missing.",
    },
    {
      icon: "⚧", color: C.warn,
      title: "Identify sex of 5 unknown mice",
      detail: "Five mice in the colony software have Sex=Unknown. They were preserved automatically. ID by physical exam at the next walkthrough and update the records.",
    },
    {
      icon: "🐭", color: C.accent,
      title: "Set up Marrow Glo breeding pair",
      detail: "Marrow Glo is the priority strain to expand. Check the dashboard's Smart Cull Plan tab → Recommended Actions for the suggested pair (youngest available F + M).",
    },
    {
      icon: "🧬", color: C.ok,
      title: "Initiate rescue-line breedings as needed",
      detail: "If Adipoq-Cre, mTmG, or Dendra2 standalone lines are needed in the next 6–12 months, set up breedings from the preserved donor mice (see slide 6). Otherwise the donors continue to live in their compound strains.",
    },
    {
      icon: "📋", color: C.navy,
      title: "Update Transnetyx — manual mark-deceased",
      detail: `Use the Cull_Checklist_${TODAY}.xlsx workbook to track 173 manual updates in the Transnetyx UI (sorted by cage for batch processing). Bulk upload was deferred due to risk to historical records.`,
    },
    {
      icon: "🔄", color: C.accent,
      title: "Refresh dashboard once colony software is updated",
      detail: "After Transnetyx and the colony software are reconciled, re-export the four .xlsx files and run the notebook → push. The shared dashboard URL stays the same.",
    },
  ];

  const startY = 1.0;
  const rowH = 0.95;
  items.forEach((it, i) => {
    const y = startY + i * rowH;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y, w: W - 1.2, h: rowH - 0.1,
      fill: { color: C.white }, line: { color: C.line, width: 0.75 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y, w: 0.08, h: rowH - 0.1,
      fill: { color: it.color }, line: { type: "none" },
    });
    // Icon circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.85, y: y + 0.18, w: 0.5, h: 0.5,
      fill: { color: it.color }, line: { type: "none" },
    });
    s.addText(it.icon, {
      x: 0.85, y: y + 0.15, w: 0.5, h: 0.5,
      fontFace: FONT_BODY, fontSize: 18, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(it.title, {
      x: 1.55, y: y + 0.1, w: W - 2.5, h: 0.35,
      fontFace: FONT_HEADER, fontSize: 14, bold: true, color: C.navyDark,
      align: "left", margin: 0,
    });
    s.addText(it.detail, {
      x: 1.55, y: y + 0.42, w: W - 2.5, h: 0.45,
      fontFace: FONT_BODY, fontSize: 10, color: C.slate,
      align: "left", margin: 0,
    });
  });
}

// =========================================================
// Slide 10 — Methodology / appendix
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navyDark };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.4, h: H,
    fill: { color: C.accent }, line: { type: "none" },
  });

  s.addText("APPENDIX", {
    x: 0.9, y: 0.5, w: 8, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, bold: true, color: "BFDBFE",
    charSpacing: 4, margin: 0,
  });
  s.addText("Methodology, sources, links", {
    x: 0.9, y: 0.85, w: 11.5, h: 0.7,
    fontFace: FONT_HEADER, fontSize: 32, bold: true, color: C.white,
    margin: 0,
  });

  // Two-column body
  const colY = 2.0, colH = 5.0;
  const colsAppendix = [
    {
      title: "Smart Cull Plan algorithm",
      lines: [
        ["For each strain, set a priority: Expand, Maintain, or Wind down.", false],
        ["For 'Maintain' strains, the youngest N (default 4) per", false],
        ["(strain × genotype × sex) cell are preserved as cohort-rebuild", false],
        ["insurance. The rest, if ≥12 mo, become cull candidates.", false],
        ["", false],
        ["For 'Wind down' strains rescuable via compound crosses, the", false],
        ["algorithm flags up to 4 donors per sex in the source strain", false],
        ["(youngest, available, non-breeder) and preserves them.", false],
        ["", false],
        ["Always preserved: active breeders, sex-unknown mice, and", false],
        ["singleton genotypes in non-rescuable strains.", false],
      ],
    },
    {
      title: "Sources & references",
      lines: [
        ["ULAM CY26 rate schedule:", true],
        ["animalcare.umich.edu/business-services/rates", false],
        ["Effective: 04/01/2026  ·  Vented mouse: $1.19/cage/day", false],
        ["Vented breeding: $2.16/cage/day", false],
        ["", false],
        ["Live colony dashboard:", true],
        ["culturedhilfolk.github.io/macdougald-colony-dashboard", false],
        ["URL is permanent; updates push automatically.", false],
        ["", false],
        ["Generated artifacts (in Mouse Colony folder):", true],
        ["Cull_Checklist_" + TODAY + ".xlsx", false],
        ["MacLab_Brian_Mice_Transnetyx_Upload_" + TODAY + ".xlsx", false],
        ["Smart_Cull_Plan_" + TODAY + ".xlsx (from dashboard)", false],
      ],
    },
  ];
  const colW = (W - 0.4 - 1.2 - 0.4) / 2;
  colsAppendix.forEach((c, i) => {
    const x = 0.9 + i * (colW + 0.4);
    s.addText(c.title, {
      x, y: colY, w: colW, h: 0.4,
      fontFace: FONT_HEADER, fontSize: 16, bold: true, color: "BFDBFE",
      margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x, y: colY + 0.42, w: colW, h: 0,
      line: { color: C.accent, width: 1 },
    });
    const text = c.lines.map(([line, bold]) => ({
      text: line,
      options: { bold, color: bold ? C.white : "CBD5E1", breakLine: true, fontSize: 11.5 }
    }));
    s.addText(text, {
      x, y: colY + 0.55, w: colW, h: colH,
      fontFace: FONT_BODY, color: "CBD5E1",
      align: "left", valign: "top", margin: 0,
    });
  });

  // Footer
  s.addText("Brian Desrosiers · MacDougald Lab · " + TODAY, {
    x: 0.9, y: H - 0.5, w: 11.5, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, italic: true, color: "BFDBFE",
    align: "left", margin: 0,
  });
}

// ---- Save ----
const outPath = path.join(OUT_DIR, `Mouse_Colony_Cull_Summary_${TODAY}.pptx`);
pres.writeFile({ fileName: outPath }).then(f => console.log("Saved:", f));
