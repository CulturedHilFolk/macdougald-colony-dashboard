// Build the cull-summary PowerPoint in Brian's actual style
// (matched against /Experiments/LD Immunization/LD Immunization.pptx):
//   - White background, thin light-blue accent line across the top
//   - Big left-aligned serif title in deep navy (Cambria), with optional
//     italic gray subtitle inline ("Title · subtitle")
//   - Bold blue section sub-headers (Calibri)
//   - Dark gray body text (Calibri), simple bullet dots, generous whitespace
//   - Navy header tables with zebra alternating rows
//   - Muted scientific color palette for charts (gray / blue / green)
//   - Italic notes; small footer "MacDougald Lab · Brian Desrosiers · DATE"
//
// Reads:  /tmp/ppt_data.json
// Writes: Mouse_Colony_Cull_Summary_<DATE>.pptx

const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const D = JSON.parse(fs.readFileSync("/tmp/ppt_data.json", "utf8"));
const OUT_DIR = "/Users/briandes/Library/CloudStorage/OneDrive-MichiganMedicine/Desktop/MacDougald Lab/Mouse Colony";
const TODAY = D.as_of;

// ---- Palette (sampled from LD Immunization deck) ----
const C = {
  bg:           "FFFFFF",
  topAccent:    "BDD7EE",  // light blue line at top
  titleNavy:    "1F4E79",  // big serif titles
  sectionBlue:  "2E75B6",  // bold sub-headers
  warnRed:      "A0292B",  // limitations / warnings
  bodyDark:     "333333",  // primary body text
  muted:        "6F6F6F",  // secondary / footnotes
  caption:      "8C8C8C",  // captions
  divider:      "D9D9D9",
  altRow:       "F2F6FA",
  tableHdrBg:   "1F4E79",
  tableHdrText: "FFFFFF",
  // Group colors (matching the LD deck's gray/blue/green motif)
  groupGray:    "6F6F6F",
  groupBlue:    "2F5996",
  groupGreen:   "5B8C3D",
  groupAmber:   "BF8B3F",
  // Strain accents (kept subtle so text stays readable)
  AdipoGlo:     "2F5996",
  "AdipoGlo+":  "5B7AB5",
  Dendra2:      "1F8470",
  "Marrow Glo": "BF8B3F",
  mTmG:         "9C4978",
  "Adipoq-Cre": "6F6F6F",
};

const FONT_HEAD = "Cambria";
const FONT_BODY = "Calibri";

const fmt   = n => new Intl.NumberFormat("en-US").format(n);
const fmt$  = n => "$" + new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(Math.round(n));
const fmt$2 = n => "$" + new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.author = "Brian Desrosiers · MacDougald Lab";
pres.title  = "Mouse Colony Cull Summary — " + TODAY;

const W = 13.3, H = 7.5;
const MARGIN_X = 0.7;

// ---- Slide chrome (top accent line + footer) ----
function addChrome(slide) {
  // Thin top accent line, full width
  slide.addShape(pres.shapes.LINE, {
    x: 0, y: 0.06, w: W, h: 0,
    line: { color: C.topAccent, width: 1.5 },
  });
  // Footer
  slide.addText(`MacDougald Lab  ·  Brian Desrosiers  ·  ${TODAY}`, {
    x: MARGIN_X, y: H - 0.35, w: W - 2 * MARGIN_X, h: 0.25,
    fontFace: FONT_BODY, fontSize: 9, italic: true, color: C.caption,
    align: "left", valign: "middle", margin: 0,
  });
}

// ---- Standard slide title (left-aligned serif, navy) ----
function addTitle(slide, title, subtitle) {
  if (subtitle) {
    slide.addText([
      { text: title, options: { bold: true, color: C.titleNavy, fontFace: FONT_HEAD } },
      { text: "  ·  ",   options: { color: C.muted, fontFace: FONT_HEAD } },
      { text: subtitle,  options: { italic: true, color: C.muted, fontFace: FONT_HEAD } },
    ], {
      x: MARGIN_X, y: 0.45, w: W - 2 * MARGIN_X, h: 0.85,
      fontSize: 30, align: "left", valign: "middle", margin: 0,
    });
  } else {
    slide.addText(title, {
      x: MARGIN_X, y: 0.45, w: W - 2 * MARGIN_X, h: 0.85,
      fontFace: FONT_HEAD, fontSize: 30, bold: true, color: C.titleNavy,
      align: "left", valign: "middle", margin: 0,
    });
  }
}

// ---- Bold blue sub-header ----
function addSubHeader(slide, x, y, w, text, color) {
  slide.addText(text, {
    x, y, w, h: 0.32,
    fontFace: FONT_BODY, fontSize: 14, bold: true,
    color: color || C.sectionBlue,
    align: "left", valign: "top", margin: 0,
  });
}

// ---- Bullet list ----
function addBullets(slide, x, y, w, h, items, opts) {
  const o = opts || {};
  const arr = items.map((it, i) => {
    const isStr = typeof it === "string";
    const text = isStr ? it : it.text;
    const bold = isStr ? false : !!it.bold;
    const last = i === items.length - 1;
    return {
      text,
      options: {
        bullet: { code: "25CF", indent: 14 }, // small filled circle
        bold,
        breakLine: !last,
        paraSpaceAfter: 6,
      },
    };
  });
  slide.addText(arr, {
    x, y, w, h,
    fontFace: FONT_BODY, fontSize: o.size || 12,
    color: C.bodyDark, valign: "top", margin: 0,
  });
}

// ---- Table helper that mirrors the LD deck look ----
function addStyledTable(slide, x, y, w, headers, rows, opts) {
  opts = opts || {};
  const colWPct = opts.colWPct || headers.map(() => 1 / headers.length);
  const colW = colWPct.map(p => +(p * w).toFixed(3));
  const headerRow = headers.map(h => ({
    text: h,
    options: {
      bold: true, color: C.tableHdrText, fill: { color: C.tableHdrBg },
      align: opts.headerAlign || "left", valign: "middle",
      fontFace: FONT_BODY, fontSize: 11,
    },
  }));
  const dataRows = rows.map((r, ri) => r.map((cell, ci) => {
    const text = typeof cell === "object" ? cell.text : cell;
    const cellOpts = (typeof cell === "object" ? cell.options : {}) || {};
    return {
      text: String(text),
      options: Object.assign({
        fill: { color: ri % 2 === 0 ? "FFFFFF" : C.altRow },
        align: opts.colAlign ? opts.colAlign[ci] : "left",
        valign: "middle",
        color: C.bodyDark,
        fontFace: FONT_BODY, fontSize: 11,
      }, cellOpts),
    };
  }));
  slide.addTable([headerRow, ...dataRows], {
    x, y, w, h: opts.h,
    colW, rowH: opts.rowH || 0.32,
    border: { type: "solid", color: C.divider, pt: 0.5 },
  });
}

// =========================================================
// Slide 1 — Title
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);

  // Centered serif title (two lines if needed)
  s.addText("Mouse Colony Cull", {
    x: 0, y: 2.6, w: W, h: 1.4,
    fontFace: FONT_HEAD, fontSize: 56, bold: true, color: C.titleNavy,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("Pre/Post Inventory  ·  Cost Analysis  ·  Cohort Outlook", {
    x: 0, y: 4.0, w: W, h: 0.55,
    fontFace: FONT_HEAD, fontSize: 22, italic: true, color: C.titleNavy,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("MacDougald Lab cull walkthrough  ·  173 mice culled  ·  n = 397 → 224", {
    x: 0, y: 4.55, w: W, h: 0.4,
    fontFace: FONT_BODY, fontSize: 13, color: C.muted,
    align: "center", valign: "middle", margin: 0,
  });

  // Centered author block lower
  s.addText([
    { text: "Brian Desrosiers", options: { bold: true, color: C.bodyDark, breakLine: true } },
    { text: "MacDougald Lab — University of Michigan", options: { color: C.muted, breakLine: true } },
    { text: TODAY, options: { italic: true, color: C.muted } },
  ], {
    x: 0, y: 5.7, w: W, h: 1.1,
    fontFace: FONT_BODY, fontSize: 12, align: "center", margin: 0,
  });
}

// =========================================================
// Slide 2 — Cull overview (study-design style)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Cull Overview");

  let y = 1.55;
  addSubHeader(s, MARGIN_X, y, 6, "Objective");
  y += 0.4;
  s.addText("Reduce per-diem cost burden while preserving every active strain and the genetic stock needed to rebuild any cohort within 6–12 months. Cull was driven by a Smart Cull Plan combining strain priority, cohort-rebuild insurance, and rescue-donor preservation for orphan strains.", {
    x: MARGIN_X, y, w: W - 2 * MARGIN_X, h: 1.0,
    fontFace: FONT_BODY, fontSize: 12, color: C.bodyDark,
    align: "left", valign: "top", margin: 0,
  });

  y = 2.85;
  addSubHeader(s, MARGIN_X, y, 6, "Strategy");
  y += 0.4;
  addBullets(s, MARGIN_X, y, W - 2 * MARGIN_X, 1.7, [
    { text: "Marrow Glo — preserve all mice; recommended new breeding pair (priority strain to expand).", bold: false },
    { text: "Adipoq-Cre, mTmG, Dendra2 — wind down and rely on rescue donors in compound strains to rebuild if needed.", bold: false },
    { text: "AdipoGlo, AdipoGlo+ — maintained with cohort-rebuild insurance (≥4 youngest mice per genotype × sex cell preserved).", bold: false },
    { text: "Always preserved regardless of priority: active breeders, sex-unknown mice, and singleton genotypes in non-rescuable strains.", bold: false },
  ]);

  y = 5.0;
  addSubHeader(s, MARGIN_X, y, 6, "Outcome");
  y += 0.4;
  s.addText([
    { text: `${D.killed} mice culled`,                    options: { bold: true, color: C.titleNavy } },
    { text: `   ·   `,                                    options: { color: C.muted } },
    { text: `${D.pre_total} → ${D.post_total} colony size`, options: { color: C.bodyDark } },
    { text: `   ·   `,                                    options: { color: C.muted } },
    { text: `${fmt$(D.cost_savings)}/yr saved`,           options: { bold: true, color: C.groupGreen } },
    { text: `   ·   `,                                    options: { color: C.muted } },
    { text: `${D.cost_savings_pct}% cost reduction`,       options: { bold: true, color: C.groupGreen } },
    { text: `   ·   `,                                    options: { color: C.muted } },
    { text: `${D.rescue_donors_count} rescue donors flagged`, options: { color: C.bodyDark } },
  ], {
    x: MARGIN_X, y, w: W - 2 * MARGIN_X, h: 0.45,
    fontFace: FONT_BODY, fontSize: 13, valign: "middle", margin: 0,
  });

  // Caveat note
  s.addText(`Note: 26 mice on the cull list could not be located physically and remain to be reconciled at the next walkthrough.`, {
    x: MARGIN_X, y: 5.85, w: W - 2 * MARGIN_X, h: 0.4,
    fontFace: FONT_BODY, fontSize: 11, italic: true, color: C.warnRed,
    valign: "middle", margin: 0,
  });
}

// =========================================================
// Slide 3 — Strain priorities (table)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Strain Priorities", "framework used to drive cull decisions");

  const rows = [
    ["Marrow Glo",  "EXPAND",    String(D.pre_strain["Marrow Glo"]||0),  String(D.post_strain["Marrow Glo"]||0),  "Priority strain to expand. All mice preserved; new breeding pair recommended."],
    ["AdipoGlo",    "MAINTAIN",  String(D.pre_strain["AdipoGlo"]||0),    String(D.post_strain["AdipoGlo"]||0),    "Largest strain. Aged surplus removed; ≥4 youngest per (genotype × sex) preserved as cohort insurance."],
    ["AdipoGlo+",   "MAINTAIN",  String(D.pre_strain["AdipoGlo+"]||0),   String(D.post_strain["AdipoGlo+"]||0),   "Triple-reporter line. Same cohort-insurance rule applied per genotype × sex cell."],
    ["Dendra2",     "WIND DOWN", String(D.pre_strain["Dendra2"]||0),     String(D.post_strain["Dendra2"]||0),     "Aging out; rescuable from AdipoGlo offspring carrying Dendra2 alone."],
    ["mTmG",        "WIND DOWN", String(D.pre_strain["mTmG"]||0),        String(D.post_strain["mTmG"]||0),        "Rescuable from AdipoGlo+ offspring with mTmG hom + WT Cre + null Dendra2."],
    ["Adipoq-Cre",  "WIND DOWN", String(D.pre_strain["Adipoq-Cre"]||0),  String(D.post_strain["Adipoq-Cre"]||0),  "Rescuable from AdipoGlo Adipoq-cre; Dendra2<-/-> donors."],
  ];

  // Add color-coded priority pills via custom cell options
  const styledRows = rows.map(r => {
    const tag = r[1];
    const tagColor = tag === "EXPAND" ? C.groupAmber : tag === "MAINTAIN" ? C.groupBlue : C.warnRed;
    return [
      { text: r[0], options: { bold: true, color: C.bodyDark } },
      { text: tag,  options: { color: tagColor, bold: true, align: "center" } },
      { text: r[2], options: { align: "right", color: C.muted } },
      { text: r[3], options: { align: "right", bold: true, color: C.bodyDark } },
      { text: r[4], options: { color: C.bodyDark } },
    ];
  });

  addStyledTable(s, MARGIN_X, 1.7, W - 2 * MARGIN_X,
    ["Strain", "Priority", "Pre", "Post", "Rationale"],
    styledRows,
    {
      colWPct: [0.13, 0.12, 0.07, 0.07, 0.61],
      colAlign: ["left", "center", "right", "right", "left"],
      headerAlign: "left",
      rowH: 0.55,
      h: 4.0,
    });

  // Caption
  s.addText("Pre / Post counts reflect the inventory immediately before and after today's cull (excludes the 26 mice that could not be located physically).", {
    x: MARGIN_X, y: 5.85, w: W - 2 * MARGIN_X, h: 0.4,
    fontFace: FONT_BODY, fontSize: 10, italic: true, color: C.caption,
    valign: "middle", margin: 0,
  });
}

// =========================================================
// Slide 4 — Inventory pre vs post-cull (table + bar chart)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Inventory by Strain", "pre- and post-cull mouse counts");

  // Table left
  const strainOrder = ["AdipoGlo", "AdipoGlo+", "Dendra2", "Marrow Glo", "mTmG", "Adipoq-Cre"];
  s.addText("Table 1.  Mouse counts by strain", {
    x: MARGIN_X, y: 1.55, w: 5, h: 0.3,
    fontFace: FONT_BODY, fontSize: 11, italic: true, color: C.muted, margin: 0,
  });
  const rows = strainOrder.map(name => {
    const pre = D.pre_strain[name] || 0;
    const post = D.post_strain[name] || 0;
    const delta = post - pre;
    const deltaStr = delta === 0 ? "—" : (delta > 0 ? "+" + delta : String(delta));
    const deltaColor = delta < 0 ? C.warnRed : delta > 0 ? C.groupGreen : C.muted;
    return [
      { text: name, options: { bold: true } },
      { text: String(pre),  options: { align: "right" } },
      { text: String(post), options: { align: "right", bold: true } },
      { text: deltaStr,     options: { align: "right", color: deltaColor } },
    ];
  });
  // Total row
  const totalDelta = D.post_total - D.pre_total;
  rows.push([
    { text: "Total", options: { bold: true, color: C.titleNavy, fill: { color: "EAF1F8" } } },
    { text: String(D.pre_total),  options: { align: "right", bold: true, fill: { color: "EAF1F8" } } },
    { text: String(D.post_total), options: { align: "right", bold: true, fill: { color: "EAF1F8" } } },
    { text: (totalDelta < 0 ? "" : "+") + totalDelta,
      options: { align: "right", bold: true, color: C.warnRed, fill: { color: "EAF1F8" } } },
  ]);
  addStyledTable(s, MARGIN_X, 1.85, 5.6,
    ["Strain", "Pre", "Post", "Δ"],
    rows,
    { colWPct: [0.45, 0.18, 0.18, 0.19], colAlign: ["left","right","right","right"], rowH: 0.36, h: 3.0 });

  // Chart right
  const chartX = 6.6, chartY = 1.55;
  s.addText("Figure 1.  Pre vs post-cull, by strain", {
    x: chartX, y: chartY, w: 6.0, h: 0.3,
    fontFace: FONT_BODY, fontSize: 11, italic: true, color: C.muted, margin: 0,
  });
  s.addChart(pres.charts.BAR, [
    { name: "Pre-cull",  labels: strainOrder, values: strainOrder.map(n => D.pre_strain[n]  || 0) },
    { name: "Post-cull", labels: strainOrder, values: strainOrder.map(n => D.post_strain[n] || 0) },
  ], {
    x: chartX, y: chartY + 0.35, w: 6.0, h: 3.5,
    chartColors: [C.groupGray, C.groupBlue],
    barDir: "col",
    barGapWidthPct: 50,
    catAxisLabelFontFace: FONT_BODY, catAxisLabelFontSize: 10, catAxisLabelColor: C.bodyDark,
    valAxisLabelFontFace: FONT_BODY, valAxisLabelFontSize: 9,  valAxisLabelColor: C.muted,
    showLegend: true, legendPos: "t",
    legendFontFace: FONT_BODY, legendFontSize: 10,
    showValue: true,
    dataLabelFontFace: FONT_BODY, dataLabelFontSize: 9, dataLabelColor: C.bodyDark,
    dataLabelPosition: "outEnd",
    plotArea: { fill: { color: "FFFFFF" } },
  });

  // Bottom narrative — placed below the chart end (y = chartY + 0.35 + 3.5 = 5.40)
  addSubHeader(s, MARGIN_X, 5.65, 6, "Notes");
  addBullets(s, MARGIN_X, 6.0, W - 2 * MARGIN_X, 1.05, [
    "AdipoGlo and AdipoGlo+ retained ≈50% and ≈75% of pre-cull mice respectively — largely the youngest in each genotype × sex cell.",
    "Dendra2, mTmG, and Adipoq-Cre standalone strains were aged out today; rescue donors preserved across compound strains (slide 6).",
    "Marrow Glo unchanged (n = 8) — recommended for breeding expansion (slide 9).",
  ], { size: 10 });
}

// =========================================================
// Slide 5 — Cage cost analysis (table + bar chart)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Cage Cost — Before & After", "ULAM CY26 ventilated rate");

  s.addText("Table 1.  Per-diem and run-rate cage costs", {
    x: MARGIN_X, y: 1.55, w: 6, h: 0.3,
    fontFace: FONT_BODY, fontSize: 11, italic: true, color: C.muted, margin: 0,
  });
  const rows = [
    ["Per day",   fmt$2(D.cost_per_day_pre),  fmt$2(D.cost_per_day_post),  "−" + fmt$2(D.cost_per_day_pre - D.cost_per_day_post)],
    ["Per month", fmt$(D.cost_per_month_pre), fmt$(D.cost_per_month_post), "−" + fmt$(D.cost_per_month_pre - D.cost_per_month_post)],
    ["Per year",  fmt$(D.cost_pre),           fmt$(D.cost_post),           "−" + fmt$(D.cost_savings)],
  ];
  const styledRows = rows.map(r => [
    { text: r[0], options: { bold: true } },
    { text: r[1], options: { align: "right" } },
    { text: r[2], options: { align: "right", bold: true } },
    { text: r[3], options: { align: "right", bold: true, color: C.groupGreen } },
  ]);
  addStyledTable(s, MARGIN_X, 1.85, 6.0,
    ["Horizon", "Pre-cull", "Post-cull", "Savings"],
    styledRows,
    { colWPct: [0.28, 0.24, 0.24, 0.24], colAlign: ["left","right","right","right"], rowH: 0.45, h: 1.9 });

  // Highlight box for annual savings
  s.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN_X, y: 4.05, w: 6.0, h: 1.45,
    fill: { color: "F4F8F1" }, line: { color: C.groupGreen, width: 0.75 },
  });
  s.addText("Annual savings", {
    x: MARGIN_X + 0.25, y: 4.15, w: 5.5, h: 0.32,
    fontFace: FONT_BODY, fontSize: 11, bold: true, color: C.groupGreen,
    margin: 0, charSpacing: 1,
  });
  s.addText(fmt$(D.cost_savings), {
    x: MARGIN_X + 0.25, y: 4.45, w: 3.8, h: 0.85,
    fontFace: FONT_HEAD, fontSize: 44, bold: true, color: C.groupGreen,
    align: "left", valign: "middle", margin: 0,
  });
  s.addText(`${D.cost_savings_pct}% reduction`, {
    x: MARGIN_X + 4.05, y: 4.5, w: 1.9, h: 0.45,
    fontFace: FONT_BODY, fontSize: 14, bold: true, color: C.bodyDark,
    align: "right", valign: "middle", margin: 0,
  });
  s.addText("recurring run-rate", {
    x: MARGIN_X + 4.05, y: 4.95, w: 1.9, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, italic: true, color: C.muted,
    align: "right", valign: "middle", margin: 0,
  });

  // Chart right side
  const cx = 7.5, cy = 1.55;
  s.addText("Figure 1.  Annual cage cost — pre vs post-cull", {
    x: cx, y: cy, w: 5.2, h: 0.3,
    fontFace: FONT_BODY, fontSize: 11, italic: true, color: C.muted, margin: 0,
  });
  s.addChart(pres.charts.BAR, [
    { name: "Cost ($/yr)", labels: ["Pre-cull", "Post-cull"],
      values: [Math.round(D.cost_pre), Math.round(D.cost_post)] },
  ], {
    x: cx, y: cy + 0.3, w: 5.2, h: 4.7,
    chartColors: [C.groupBlue, C.groupGreen],
    chartColorsOpacity: 100,
    barDir: "col",
    barGapWidthPct: 60,
    barColors: [C.groupBlue, C.groupGreen],
    catAxisLabelFontFace: FONT_BODY, catAxisLabelFontSize: 11, catAxisLabelColor: C.bodyDark,
    valAxisLabelFontFace: FONT_BODY, valAxisLabelFontSize: 9,  valAxisLabelColor: C.muted,
    showLegend: false,
    showValue: true,
    valueFormatCode: "$#,##0",
    dataLabelFontFace: FONT_BODY, dataLabelFontSize: 11, dataLabelColor: C.bodyDark,
    dataLabelPosition: "outEnd",
    plotArea: { fill: { color: "FFFFFF" } },
  });

  // Methodology note
  s.addText([
    { text: "Methodology: ", options: { bold: true, color: C.bodyDark } },
    { text: "$1.19/day per ventilated cage, $2.16/day per breeding-vented cage. Post-cull projection assumes cages with zero surviving mice are returned to ULAM and stop billing. Husbandry tech labor billed separately and not included.",
      options: { color: C.muted } },
  ], {
    x: MARGIN_X, y: 6.55, w: W - 2 * MARGIN_X, h: 0.5,
    fontFace: FONT_BODY, fontSize: 10, italic: true, valign: "top", margin: 0,
  });
}

// =========================================================
// Slide 6 — Rescue donor strategy
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Rescue Donor Strategy", `${D.rescue_donors_count} donors flagged`);

  s.addText("Adipoq-Cre, mTmG, and Dendra2 standalone strains were aged out, but each can be regenerated from offspring of compound strains where the target transgene is present alone. The Smart Cull Plan automatically preserves the youngest donors.", {
    x: MARGIN_X, y: 1.5, w: W - 2 * MARGIN_X, h: 0.6,
    fontFace: FONT_BODY, fontSize: 12, color: C.bodyDark,
    align: "left", valign: "top", margin: 0,
  });

  const rescues = [
    { target: "Adipoq-Cre", donors: "Adipoq-cre; Dendra2<-/->",  src: "AdipoGlo offspring",  note: "Cre present, no Dendra2 transgene, no mTmG." },
    { target: "Dendra2",    donors: "WT; Dendra2<+/+ or +/->",   src: "AdipoGlo offspring",  note: "Dendra2 present, no Cre, no mTmG." },
    { target: "mTmG",       donors: "mTmG<mTmG/mTmG>; WT",        src: "AdipoGlo+ offspring", note: "mTmG hom only, no Cre, no Dendra2 transgene." },
  ];

  const styledRows = rescues.map(r => [
    { text: r.target,  options: { bold: true, color: C.titleNavy } },
    { text: r.donors,  options: { fontFace: "Consolas", fontSize: 10, color: C.bodyDark } },
    { text: r.src,     options: { color: C.bodyDark } },
    { text: "up to 4 ♀ + 4 ♂", options: { align: "center", bold: true, color: C.groupGreen } },
    { text: r.note,    options: { italic: true, color: C.muted, fontSize: 10 } },
  ]);
  addStyledTable(s, MARGIN_X, 2.3, W - 2 * MARGIN_X,
    ["Target", "Donor genotype", "Source strain", "Preserved this cull", "Note"],
    styledRows,
    { colWPct: [0.13, 0.30, 0.18, 0.16, 0.23], colAlign: ["left","left","left","center","left"], rowH: 0.5, h: 2.5 });

  addSubHeader(s, MARGIN_X, 5.05, 6, "Why this works");
  addBullets(s, MARGIN_X, 5.4, W - 2 * MARGIN_X, 1.5, [
    "Donors carry only the target transgene (or have null forms of the others), so breeding a donor to a WT mouse yields offspring that are effectively the orphan strain — no extra back-crossing needed.",
    "Donors preserved are the youngest available, sex-balanced (up to 4 per sex per orphan strain) — typical fertility and a usable runway of breeding age.",
    "If any orphan strain is needed within the next 6–12 months, set up the corresponding rescue cross. Otherwise donors continue to live productively in their native compound strain.",
  ], { size: 11 });
}

// =========================================================
// Slide 7 — Available cohort cells
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Available Cohort Cells", "post-cull groups with n ≥ 3");

  s.addText("Each row is an experimental group that can be assembled from surviving mice without new breeding. Sorted by group size descending.", {
    x: MARGIN_X, y: 1.55, w: W - 2 * MARGIN_X, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, color: C.muted, italic: true,
    align: "left", valign: "top", margin: 0,
  });

  const top = (D.cohort_cells_top || []).slice(0, 8);
  const styledRows = top.map(r => [
    { text: r.strain, options: { bold: true } },
    { text: r.genotype, options: { fontFace: "Consolas", fontSize: 10, color: C.bodyDark } },
    { text: r.sex === "Female" ? "F" : "M", options: { align: "center", color: r.sex === "Female" ? "9C4978" : C.groupBlue, bold: true } },
    { text: String(r.n), options: { align: "right", bold: true } },
    { text: r.mean_age != null ? r.mean_age + " mo" : "—", options: { align: "right" } },
    { text: r.min_age != null ? r.min_age + " – " + r.max_age + " mo" : "—", options: { align: "right", color: C.muted } },
  ]);

  addStyledTable(s, MARGIN_X, 2.1, W - 2 * MARGIN_X,
    ["Strain", "Genotype", "Sex", "n", "Mean age", "Range"],
    styledRows,
    { colWPct: [0.12, 0.46, 0.07, 0.08, 0.13, 0.14],
      colAlign: ["left","left","center","right","right","right"],
      rowH: 0.4, h: 4.0 });

  addSubHeader(s, MARGIN_X, 6.25, 6, "How to read");
  s.addText("Use these groups for end-point assays (Seahorse, microscopy, flow). For matched controls of a different genotype, use the Cohort Builder on the live dashboard — it surfaces best-matched controls by sex and age automatically.", {
    x: MARGIN_X, y: 6.6, w: W - 2 * MARGIN_X, h: 0.55,
    fontFace: FONT_BODY, fontSize: 11, color: C.bodyDark,
    align: "left", valign: "top", margin: 0,
  });
}

// =========================================================
// Slide 8 — Active cohort plans (study-design style)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Active Cohort Plans", "research goals + available counts");

  const plans = (D.cohort_plans || []).slice(0, 3);
  let y = 1.5;
  plans.forEach((p, i) => {
    addSubHeader(s, MARGIN_X, y, 6, p.strain);
    s.addText([
      { text: "Cross — ",   options: { bold: true, color: C.bodyDark } },
      { text: p.cross_desc || "—", options: { fontFace: "Consolas", fontSize: 10, color: C.bodyDark } },
    ], {
      x: MARGIN_X + 0.0, y: y + 0.4, w: W - 2 * MARGIN_X, h: 0.4,
      fontFace: FONT_BODY, fontSize: 11, valign: "top", margin: 0,
    });
    s.addText([
      { text: "Purpose — ", options: { bold: true, color: C.bodyDark } },
      { text: p.purpose || "—", options: { color: C.bodyDark } },
    ], {
      x: MARGIN_X, y: y + 0.75, w: W - 2 * MARGIN_X, h: 0.45,
      fontFace: FONT_BODY, fontSize: 11, italic: false, valign: "top", margin: 0,
    });
    s.addText([
      { text: "Target — ",  options: { bold: true, color: C.bodyDark } },
      { text: p.target_geno || "—", options: { fontFace: "Consolas", fontSize: 10, color: C.groupGreen, bold: true } },
      { text: "    ·    Available toward target — ", options: { color: C.bodyDark } },
      { text: String(p.currently_available != null ? p.currently_available : "—") + " mice (pre-cull)",
        options: { bold: true, color: C.titleNavy } },
    ], {
      x: MARGIN_X, y: y + 1.2, w: W - 2 * MARGIN_X, h: 0.4,
      fontFace: FONT_BODY, fontSize: 11, valign: "top", margin: 0,
    });

    // Divider
    if (i < plans.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: MARGIN_X, y: y + 1.7, w: W - 2 * MARGIN_X, h: 0,
        line: { color: C.divider, width: 0.5 },
      });
    }
    y += 1.75;
  });
}

// =========================================================
// Slide 9 — Findings & open items (two-column)
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Findings & Open Items");

  const colY = 1.55;
  // Findings (left)
  addSubHeader(s, MARGIN_X, colY, 5.5, "Findings");
  addBullets(s, MARGIN_X, colY + 0.4, 5.5, 5.0, [
    { text: "173 mice culled — 78% of the planned 199-mouse cull list executed; remaining 26 deferred to walkthrough.", bold: false },
    { text: "Annual cage cost reduced 35.1% (≈$27,798/yr). 64 cages projected to retire as their occupants are removed.", bold: false },
    { text: "Marrow Glo entirely preserved (n = 8) — no cull, expansion plan recommended.", bold: false },
    { text: "Adipoq-Cre, mTmG, Dendra2 standalone strains aged out; 19 rescue donors preserved across AdipoGlo / AdipoGlo+ to enable rebuild.", bold: false },
    { text: "AdipoGlo and AdipoGlo+ cohort-rebuild floors held — every (genotype × sex) cell retains its 4 youngest mice.", bold: false },
  ], { size: 11 });

  // Open items (right) — header in red like the LD deck
  addSubHeader(s, 7.1, colY, 5.5, "Open Items", C.warnRed);
  addBullets(s, 7.1, colY + 0.4, 5.5, 5.0, [
    { text: "Walkthrough — locate / reconcile 26 mice marked Alive but not found on cull day.", bold: false },
    { text: "Identify sex of 5 sex-unknown mice and update colony software.", bold: false },
    { text: "Set up recommended Marrow Glo breeding pair (see dashboard → Smart Cull Plan).", bold: false },
    { text: "Initiate rescue-line breedings if Adipoq-Cre / mTmG / Dendra2 needed in the next 6–12 mo.", bold: false },
    { text: "Manually mark 173 mice deceased in Transnetyx (Cull_Checklist_" + TODAY + ".xlsx).", bold: false },
    { text: "Refresh dashboard once colony software is updated — same URL, auto-deploys.", bold: false },
  ], { size: 11 });

  // Note at bottom
  s.addText("Bulk Transnetyx upload was deferred because Hercules import behavior could not be confirmed against historical genotyping records — manual updates are safer.", {
    x: MARGIN_X, y: 6.85, w: W - 2 * MARGIN_X, h: 0.4,
    fontFace: FONT_BODY, fontSize: 10, italic: true, color: C.warnRed,
    valign: "middle", margin: 0,
  });
}

// =========================================================
// Slide 10 — Methodology & sources
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addChrome(s);
  addTitle(s, "Methodology & Sources");

  let y = 1.55;
  addSubHeader(s, MARGIN_X, y, 6, "Smart Cull Plan algorithm");
  y += 0.4;
  addBullets(s, MARGIN_X, y, W - 2 * MARGIN_X, 1.4, [
    "For each strain, set a priority — Expand, Maintain, or Wind down.",
    "For Maintain strains, the youngest 4 per (strain × genotype × sex) cell are preserved as cohort-rebuild insurance; the rest, if ≥12 months old, become cull candidates.",
    "For Wind down strains rescuable via compound crosses, up to 4 donors per sex are flagged in the source strain (youngest, available, non-breeder) and preserved.",
    "Always preserved: active breeders, sex-unknown mice, and singleton genotypes in non-rescuable strains.",
  ], { size: 11 });

  y = 3.5;
  addSubHeader(s, MARGIN_X, y, 6, "Cost model");
  y += 0.4;
  addBullets(s, MARGIN_X, y, W - 2 * MARGIN_X, 0.9, [
    "ULAM CY26 ventilated rate (effective 04/01/2026): $1.19/day non-breeding cage; $2.16/day breeding cage.",
    "Source: animalcare.umich.edu/business-services/rates  ·  rate schedule revised annually.",
    "Post-cull projection assumes cages with zero surviving mice are returned to ULAM and stop billing.",
  ], { size: 11 });

  y = 5.2;
  addSubHeader(s, MARGIN_X, y, 6, "Live colony dashboard");
  y += 0.35;
  s.addText("https://culturedhilfolk.github.io/macdougald-colony-dashboard/", {
    x: MARGIN_X, y, w: W - 2 * MARGIN_X, h: 0.35,
    fontFace: "Consolas", fontSize: 11, color: C.sectionBlue,
    align: "left", valign: "middle", margin: 0,
  });
  s.addText("URL is permanent; pushes auto-deploy in ~30 seconds.", {
    x: MARGIN_X, y: y + 0.35, w: W - 2 * MARGIN_X, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, italic: true, color: C.muted,
    align: "left", valign: "middle", margin: 0,
  });

  y = 6.15;
  addSubHeader(s, MARGIN_X, y, 6, "Generated artifacts (in Mouse Colony folder)");
  s.addText([
    { text: `Cull_Checklist_${TODAY}.xlsx`, options: { fontFace: "Consolas", fontSize: 10, color: C.bodyDark, breakLine: true } },
    { text: `MacLab_Brian_Mice_Transnetyx_Upload_${TODAY}.xlsx`, options: { fontFace: "Consolas", fontSize: 10, color: C.bodyDark, breakLine: true } },
    { text: `Mouse_Colony_Cull_Summary_${TODAY}.pptx`, options: { fontFace: "Consolas", fontSize: 10, color: C.bodyDark } },
  ], {
    x: MARGIN_X, y: y + 0.35, w: W - 2 * MARGIN_X, h: 0.7,
    valign: "top", margin: 0,
  });
}

// ---- Save ----
const outPath = path.join(OUT_DIR, `Mouse_Colony_Cull_Summary_${TODAY}.pptx`);
pres.writeFile({ fileName: outPath }).then(f => console.log("Saved:", f));
