// McDonald's Azerbaijan – Meta performance review (BEFORE Apr–Jun vs AFTER Jul–Sep 2026)
// Every number on the slides is read from metrics.json (built by analysis.py from the two Ads Manager exports).
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { setTheme } = require("./set_theme.js");

const ROOT = __dirname;
const M = JSON.parse(fs.readFileSync(path.join(ROOT, "..", "analysis", "metrics.json"), "utf8"));
const OUT = process.argv[2] || path.join(ROOT, "..", "McDonalds_AZ_Meta_Performance_Review_Q2_vs_Q3_2026.pptx");
const ASSETS = process.env.ASSETS_DIR || path.join(ROOT, "..", "assets");
const findAsset = (base) => ["png", "jpg", "jpeg"].map((e) => path.join(ASSETS, `${base}.${e}`)).find((p) => fs.existsSync(p));
const LOGO = findAsset("mcd_logo");          // official logo, supplied by client
const HERO = findAsset("product_hero");      // official product image, supplied by client
const DECOR = findAsset("decor_arches");     // Higgsfield-generated abstract visual (decorative only)

const THEME = {
  name: "McDonalds AZ QBR",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "27251F", lt1: "FFFFFF", dk2: "4A4A4A", lt2: "F6F5F2",
    accent1: "DA291C", accent2: "FFBC0D", accent3: "C9C9C9", accent4: "2E8540",
    accent5: "6E6E6E", accent6: "FFF1C7", hlink: "DA291C", folHlink: "9E1B10",
  },
};
const H = THEME.colors; // hex (charts, shadows)
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "McDonald's Azerbaijan – Meta Performance Review";
pres.subject = "Meta Ads: April–June vs July–September 2026";
pres.author = "Meta media agency";
const C = pres.SchemeColor;
const RED = C.accent1, GOLD = C.accent2, DARK = C.text1, MID = C.text2, GREY = C.accent5, LIGHT = C.background2, GREEN = C.accent4, PALE = C.accent6;
const BEFORE_HEX = "BDBDBD", AFTER_HEX = H.accent2, GOOD_HEX = H.accent4, BAD_HEX = H.accent1, NEUTRAL_HEX = "9A9A9A";

// ---------- formatting ----------
const T = M.total, B = T.BEFORE, A = T.AFTER;
const ch = (b, a) => ((a - b) / b) * 100;
const num = (v, d = 0) => v.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const usd = (v, d = 2) => "$" + num(v, d);
const big = (v, d = 1) => (Math.abs(v) >= 1e6 ? num(v / 1e6, d) + "M" : Math.abs(v) >= 1e3 ? num(v / 1e3, d) + "K" : num(v, 0));
const sgn = (p, d = 0) => (p > 0 ? "+" : p < 0 ? "−" : "") + num(Math.abs(p), d) + "%";
const pa = (p, d = 0) => num(Math.abs(p), d) + "%"; // magnitude only, for sentences with a verb
const arrow = (p) => (p > 0 ? "▲ " : p < 0 ? "▼ " : "");

// ---------- layouts ----------
const FOOT = "McDonald's Azerbaijan  ·  Meta performance review  ·  Apr–Jun vs Jul–Sep 2026";
pres.defineSlideMaster({
  title: "MCD_COVER",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 2.05, w: 6.9, h: 0.8, fontSize: 38, bold: true, color: DARK, margin: 0, valign: "top", align: "left" }, text: "" } },
  ],
});
const contentObjects = [
  { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.42, w: 10.5, h: 0.3, fontSize: 12, bold: true, color: RED, charSpacing: 1.5, margin: 0, valign: "middle" }, text: "" } },
  { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.76, w: 11.3, h: 0.95, fontSize: 26, bold: true, color: DARK, margin: 0, valign: "top", align: "left" }, text: "" } },
  { text: { text: FOOT, options: { x: 0.6, y: 7.02, w: 9, h: 0.25, fontSize: 9, color: GREY, margin: 0 } } },
];
if (LOGO) contentObjects.push({ image: { path: LOGO, x: 12.18, y: 0.4, w: 0.55, h: 0.48, sizing: { type: "contain", w: 0.55, h: 0.48 } } });
pres.defineSlideMaster({
  title: "MCD_CONTENT",
  background: { color: "FFFFFF" },
  objects: contentObjects,
  slideNumber: { x: 12.23, y: 7.02, w: 0.5, h: 0.25, fontSize: 9, color: GREY, align: "right" },
});

// ---------- helpers ----------
let n = 0;
const name = (s) => `${s}-${++n}`;
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: DARK, valign: "top", objectName: name("text") }, o));
}
function motif(slide) {
  // small golden half-ring: the deck's recurring arch cue (replaced by the official logo when supplied)
  if (LOGO) return;
  slide.addShape(pres.shapes.BLOCK_ARC, { x: 12.18, y: 0.42, w: 0.56, h: 0.56, fill: { color: GOLD }, line: { type: "none" }, angleRange: [180, 0], arcThicknessRatio: 0.32, objectName: name("arch-motif") });
}
function content(section, kicker, title) {
  const s = pres.addSlide({ masterName: "MCD_CONTENT", sectionTitle: section });
  s.addText(kicker.toUpperCase(), { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  motif(s);
  return s;
}
function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill || LIGHT }, line: { type: "none" }, objectName: name("card") });
}
function badge(slide, x, y, w, p, verdict, d = 0) {
  const fill = verdict === "good" ? GREEN : verdict === "bad" ? RED : GREY;
  slide.addText(arrow(p) + num(Math.abs(p), d) + "%", { isTextBox: true, x, y, w, h: 0.32, margin: 0, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.16, fill: { color: fill }, objectName: name("badge") });
}
function source(slide, t) {
  txt(slide, t, { x: 0.6, y: 6.68, w: 12.1, h: 0.26, fontSize: 9, color: GREY, valign: "bottom" });
}
function numCircle(slide, x, y, d, label, fill, color) {
  slide.addText(String(label), { isTextBox: true, x, y, w: d, h: d, margin: 0, shape: pres.shapes.OVAL, fill: { color: fill || GOLD }, color: color || DARK, fontSize: 14, bold: true, align: "center", valign: "middle", objectName: name("num") });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt" };
function chartBase(extra) {
  return Object.assign({
    catAxisLabelColor: "4A4A4A", valAxisLabelColor: "6E6E6E", catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
    dataLabelFontSize: 11, dataLabelColor: "27251F", legendFontSize: 11, legendColor: "4A4A4A",
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLineShow: false, valAxisLineShow: false,
    showValue: true, valAxisHidden: true, showTitle: false,
  }, chartText, extra);
}
function chartTitle(slide, t, x, y, w) {
  txt(slide, t, { x, y, w, h: 0.3, fontSize: 13, bold: true, color: DARK });
}

const META_SRC = "Source: Meta Ads Manager ad-level exports – BEFORE 1 Apr–30 Jun 2026 (77 ads), AFTER 1 Jul–30 Sep 2026 (106 ads). Rates recalculated from totals (weighted), not averaged.";
const obj = M.objective;
const reachB = obj.Reach.BEFORE, reachA = obj.Reach.AFTER;

// =====================================================================================
// 1. COVER
// =====================================================================================
pres.addSection({ title: "Cover" });
{
  const s = pres.addSlide({ masterName: "MCD_COVER", sectionTitle: "Cover" });
  txt(s, "QUARTERLY BUSINESS REVIEW  ·  META ADVERTISING", { x: 0.7, y: 1.55, w: 6.6, h: 0.3, fontSize: 12, bold: true, color: RED, charSpacing: 2 });
  s.addText("McDonald's Azerbaijan", { placeholder: "title" });
  txt(s, "Meta Performance Review", { x: 0.7, y: 2.9, w: 6.6, h: 0.6, fontSize: 30, bold: true, color: RED });
  txt(s, "April–June vs July–September 2026", { x: 0.7, y: 3.6, w: 6.6, h: 0.45, fontSize: 20, color: MID });
  // period chips
  s.addText([{ text: "BEFORE  ", options: { bold: true } }, { text: "1 Apr – 30 Jun · previous agency" }], { isTextBox: true, x: 0.7, y: 4.45, w: 3.55, h: 0.42, margin: 0, fontSize: 12, color: DARK, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.21, fill: { color: LIGHT }, objectName: name("chip") });
  s.addText([{ text: "AFTER  ", options: { bold: true } }, { text: "1 Jul – 30 Sep · current agency" }], { isTextBox: true, x: 4.4, y: 4.45, w: 3.4, h: 0.42, margin: 0, fontSize: 12, color: DARK, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.21, fill: { color: GOLD }, objectName: name("chip") });
  // logo slot
  if (LOGO) {
    s.addImage({ path: LOGO, x: 0.7, y: 0.5, w: 0.95, h: 0.84, sizing: { type: "contain", w: 0.95, h: 0.84 }, objectName: name("logo") });
  } else {
    s.addText("Official McDonald's logo\n(placeholder – add assets/mcd_logo.png)", { isTextBox: true, x: 0.7, y: 0.5, w: 2.6, h: 0.75, margin: 0.05, fontSize: 10, color: GREY, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, line: { color: C.accent3, width: 1, dashType: "dash" }, objectName: name("logo-placeholder") });
  }
  s.addText("Agency logo\n(placeholder)", { isTextBox: true, x: 0.7, y: 6.15, w: 1.8, h: 0.6, margin: 0.05, fontSize: 10, color: GREY, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, line: { color: C.accent3, width: 1, dashType: "dash" }, objectName: name("agency-placeholder") });
  txt(s, "Prepared October 2026", { x: 2.75, y: 6.3, w: 3, h: 0.3, fontSize: 11, color: GREY, valign: "middle" });
  // right visual
  if (HERO) {
    s.addImage({ path: HERO, x: 8.1, y: 0.6, w: 4.65, h: 6.3, sizing: { type: "cover", w: 4.65, h: 6.3 }, objectName: name("hero") });
  } else if (DECOR) {
    s.addImage({ path: DECOR, x: 8.1, y: 0.0, w: 5.23, h: 7.5, sizing: { type: "cover", w: 5.23, h: 7.5 }, objectName: name("decor") });
  } else {
    card(s, 8.1, 0.6, 4.65, 6.3, PALE);
    const cx = 8.1 + 2.325, base = 6.9;
    [[4.1, GOLD, 0.2], [2.95, RED, 0.24], [1.8, GOLD, 0.34]].forEach(([d, col, r]) => {
      s.addShape(pres.shapes.BLOCK_ARC, { x: cx - d / 2, y: base - d / 2, w: d, h: d, fill: { color: col }, line: { type: "none" }, angleRange: [180, 0], arcThicknessRatio: r, objectName: name("cover-arch") });
    });
    s.addShape(pres.shapes.OVAL, { x: 11.55, y: 1.1, w: 0.55, h: 0.55, fill: { color: RED }, line: { type: "none" }, objectName: name("dot") });
    s.addShape(pres.shapes.OVAL, { x: 8.6, y: 1.6, w: 0.3, h: 0.3, fill: { color: GOLD }, line: { type: "none" }, objectName: name("dot") });
  }
  s.addNotes("Cover. BEFORE = 1 Apr–30 Jun 2026 (previous agency). AFTER = 1 Jul–30 Sep 2026 (current agency, took over 1 July). Logo and imagery: official assets only – placeholders are marked where files were not supplied.");
}

// =====================================================================================
// 2. EXECUTIVE SUMMARY
// =====================================================================================
pres.addSection({ title: "Executive summary" });
{
  const s = content("Executive summary", "Executive summary", "Q3 bought broader reach and much stronger clicks; media cost and engagement efficiency are the next battleground");
  const exposureCampaigns = ["EntryLevel_Reach_Sep_26", "HM_Reach_Sep'26"];
  const campA = Object.fromEntries(M.campaigns.AFTER.map((c) => [c.campaign, c]));
  const inter = M.result_type.Interactions.AFTER;
  const pocket = exposureCampaigns.reduce((t, c) => t + campA[c].spend, 0) + inter.spend;
  const pocketShare = (pocket / A.spend) * 100;
  const peGoal = M.result_type["Post engagements"].AFTER.cost_per_post_eng;
  const ratioLo = campA["EntryLevel_Reach_Sep_26"].cost_per_1k_reach / reachA.cost_per_1k_reach;
  const ratioHi = inter.cost_per_post_eng / peGoal;
  const rows = [
    ["Scale: budget " + sgn(ch(B.spend, A.spend)) + ", reach " + sgn(ch(B.reach_sum, A.reach_sum)),
      `Spend rose from ${usd(B.spend / 1000, 1)}K to ${usd(A.spend / 1000, 1)}K. Summed reach grew almost in line while frequency fell from ${num(B.freq, 2)} to ${num(A.freq, 2)} – more people, fewer repeats.`,
      sgn(ch(B.reach_sum, A.reach_sum)), "reach", "neutral"],
    ["Traffic is the clearest efficiency win",
      `Link clicks ${sgn(ch(B.link_clicks, A.link_clicks))}, link CTR ${sgn(ch(B.link_ctr, A.link_ctr))} and CPC down from ${usd(B.cpc)} to ${usd(A.cpc)}. Reach campaigns alone produced ${num(M.clicks_by_obj.AFTER.Reach / M.clicks_by_obj.BEFORE.Reach, 1)}× more link clicks.`,
      sgn(ch(B.cpc, A.cpc)), "CPC", "good"],
    ["Reach campaigns reach people more cheaply",
      `In Reach-objective campaigns, cost per 1,000 people reached fell from ${usd(reachB.cost_per_1k_reach)} to ${usd(reachA.cost_per_1k_reach)}: lower frequency more than offset a ${sgn(ch(reachB.cpm, reachA.cpm))} CPM increase.`,
      sgn(ch(reachB.cost_per_1k_reach, reachA.cost_per_1k_reach)), "cost / 1,000 reached", "good"],
    ["Media cost and engagement efficiency declined",
      `Account CPM rose ${pa(ch(B.cpm, A.cpm))}. Post engagements (${sgn(ch(B.post_eng, A.post_eng))}) and 3-second plays (${sgn(ch(B.v3s, A.v3s))}) stayed flat on the bigger budget, so cost per engagement rose ${pa(ch(B.cpe, A.cpe))}.`,
      sgn(ch(B.cpm, A.cpm)), "CPM", "bad"],
    [`${num(pocketShare, 0)}% of Q3 spend sits in three fixable pockets`,
      `September Happy Meal and Entry Level reach campaigns (frequency 2.3–2.6) and Interactions-optimised engagement ads cost ${num(ratioLo, 1)}–${num(ratioHi, 0)}× more per outcome than the Q3 norm.`,
      usd(pocket / 1000, 1) + "K", "of AFTER spend", "bad"],
  ];
  rows.forEach((r, i) => {
    const y = 1.92 + i * 0.95;
    card(s, 0.6, y, 12.13, 0.82);
    numCircle(s, 0.8, y + 0.16, 0.5, i + 1, i < 3 ? GOLD : RED, i < 3 ? DARK : C.background1);
    txt(s, r[0], { x: 1.5, y: y + 0.1, w: 8.6, h: 0.3, fontSize: 15, bold: true });
    txt(s, r[1], { x: 1.5, y: y + 0.4, w: 8.7, h: 0.4, fontSize: 11.5, color: MID });
    const col = r[4] === "good" ? GREEN : r[4] === "bad" ? RED : DARK;
    txt(s, r[2], { x: 10.35, y: y + 0.08, w: 2.2, h: 0.45, fontSize: 24, bold: true, color: col, align: "right" });
    txt(s, r[3], { x: 10.35, y: y + 0.52, w: 2.2, h: 0.25, fontSize: 10, color: GREY, align: "right" });
  });
  source(s, META_SRC);
  s.addNotes(`Pocket = EntryLevel_Reach_Sep_26 + HM_Reach_Sep'26 + all Interactions-optimised ads = ${usd(pocket)} (${num(pocketShare, 1)}% of AFTER spend). Entry Level cost per 1,000 reached ${usd(campA["EntryLevel_Reach_Sep_26"].cost_per_1k_reach)} vs Q3 Reach average ${usd(reachA.cost_per_1k_reach)} (${num(ratioLo, 2)}×); Interactions-goal cost per post engagement ${usd(inter.cost_per_post_eng, 4)} vs Post-engagement goal ${usd(peGoal, 4)} (${num(ratioHi, 1)}×).`);
}

// =====================================================================================
// 3. SNAPSHOT
// =====================================================================================
pres.addSection({ title: "Before vs after" });
{
  const s = content("Before vs after", "Before vs after · performance snapshot",
    `At a glance: ${sgn(ch(B.spend, A.spend))} budget bought ${sgn(ch(B.reach_sum, A.reach_sum))} reach and ${sgn(ch(B.link_clicks, A.link_clicks))} link clicks`);
  const K = [
    ["Amount spent", "spend", (v) => usd(v / 1000, 1) + "K", "neutral"],
    ["Reach*", "reach_sum", (v) => big(v), "up"],
    ["Impressions", "impr", (v) => big(v), "up"],
    ["Frequency*", "freq", (v) => num(v, 2), "neutral"],
    ["CPM", "cpm", (v) => usd(v), "down"],
    ["Link clicks", "link_clicks", (v) => big(v), "up"],
    ["Link CTR", "link_ctr", (v) => num(v, 3) + "%", "up"],
    ["CPC (link)", "cpc", (v) => usd(v), "down"],
    ["Post engagements", "post_eng", (v) => big(v, 2), "up"],
    ["3-sec video plays", "v3s", (v) => big(v, 2), "up"],
  ];
  const w = 2.25, gap = 0.22, h = 2.02;
  K.forEach(([lab, k, f, good], i) => {
    const x = 0.6 + (i % 5) * (w + gap), y = 1.88 + Math.floor(i / 5) * (h + 0.2);
    const p = ch(B[k], A[k]);
    const verdict = good === "neutral" ? "neutral" : (good === "up" ? p > 0 : p < 0) ? "good" : "bad";
    card(s, x, y, w, h);
    txt(s, lab.toUpperCase(), { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.25, fontSize: 10.5, bold: true, color: GREY, charSpacing: 1 });
    txt(s, f(A[k]), { x: x + 0.2, y: y + 0.5, w: w - 0.4, h: 0.55, fontSize: 28, bold: true, color: DARK });
    txt(s, "AFTER (Jul–Sep)", { x: x + 0.2, y: y + 1.05, w: w - 0.4, h: 0.22, fontSize: 9.5, color: GREY });
    txt(s, [{ text: "Before: ", options: { color: GREY } }, { text: f(B[k]), options: { bold: true, color: MID } }], { x: x + 0.2, y: y + 1.32, w: w - 0.4, h: 0.25, fontSize: 12 });
    badge(s, x + 0.2, y + 1.65, 1.05, p, verdict);
  });
  txt(s, [
    { text: "▲▼ ", options: { bold: true } }, { text: "Green = improvement, red = decline, grey = context (budget, frequency).  " },
    { text: "*Reach is the sum of ad-level reach (Meta does not de-duplicate across ads in this export); frequency = impressions ÷ summed reach." },
  ], { x: 0.6, y: 6.2, w: 12.1, h: 0.4, fontSize: 9.5, color: GREY });
  source(s, META_SRC);
  s.addNotes("All values are period totals across every ad in each export. CPM = spend ÷ impressions × 1,000; Link CTR = link clicks ÷ impressions; CPC = spend ÷ link clicks; frequency = impressions ÷ summed reach.");
}

// =====================================================================================
// 4. TRANSITION – account structure
// =====================================================================================
{
  const s = content("Before vs after", "Transition · 1 July 2026",
    "The new setup moved from monthly objective buckets to product-led campaigns with audience-based ad sets");
  // spend mix chart
  chartTitle(s, "Share of spend by optimisation objective", 0.6, 1.9, 5.6);
  const objs = ["Reach", "Engagement", "Traffic", "Video views", "App installs"];
  const objColors = [H.accent2, H.accent1, "27251F", "9A9A9A", "D9D9D9"];
  const data = objs.map((o) => ({
    name: o, labels: ["BEFORE", "AFTER"],
    values: ["BEFORE", "AFTER"].map((p) => (obj[o] && obj[o][p] ? (obj[o][p].spend / T[p].spend) * 100 : 0)),
  }));
  s.addChart(pres.charts.BAR, data, chartBase({
    x: 0.5, y: 2.25, w: 5.8, h: 3.0, barDir: "bar", barGrouping: "percentStacked", chartColors: objColors,
    dataLabelPosition: "ctr", dataLabelFormatCode: '[>=5]0"%";;;', dataLabelColor: "FFFFFF", dataLabelFontSize: 10, dataLabelFontBold: true,
    showLegend: true, legendPos: "b", catAxisLabelFontSize: 12, catAxisLabelFontBold: true, barGapWidthPct: 45, objectName: name("chart-mix"),
  }));
  txt(s, `Reach stays the backbone (${num(obj.Reach.BEFORE.spend / B.spend * 100, 0)}% → ${num(obj.Reach.AFTER.spend / A.spend * 100, 0)}%); engagement and video goals took a larger share, and an app-install test was added.`,
    { x: 0.6, y: 5.4, w: 5.6, h: 0.75, fontSize: 12, color: MID });
  // structure comparison
  const rows = [
    ["", "BEFORE · Apr–Jun", "AFTER · Jul–Sep"],
    ["Campaigns", `${B.campaigns} – by objective × month (e.g. Reach_April'26)`, `${A.campaigns} – by product / initiative (FIFA, Mixology, Happy Meal, Chicken…)`],
    ["Ad sets", `${B.adsets}, each holding one ad and named after the creative`, `${A.adsets}, named by audience (GenZ, Parents of Gen A, Core 18-34…)`],
    ["Ads", `${B.ads} (${num(B.ads / B.adsets, 1)} per ad set)`, `${A.ads} (${num(A.ads / A.adsets, 1)} per ad set)`],
    ["Result types", `${Object.keys(M.result_type).filter((k) => M.result_type[k].BEFORE).length}: Reach, IG profile visits, Post engagements, ThruPlay`, `7: adds Link clicks, Interactions, 2-sec video views, app installs (iOS / Android)`],
  ];
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 11, color: DARK, valign: "middle", margin: [4, 6, 4, 6] }, o) });
  s.addTable(rows.map((r, i) => r.map((t, j) => cell(t, {
    bold: i === 0 || j === 0, fill: { color: i === 0 ? (j === 2 ? H.accent2 : j === 1 ? "E4E4E4" : "FFFFFF") : (j === 2 ? "FFF8E1" : j === 1 ? H.lt2 : "FFFFFF") },
    color: DARK, fontSize: i === 0 ? 12 : 11,
  }))), { x: 6.75, y: 1.95, w: 5.98, colW: [1.25, 2.3, 2.43], rowH: [0.42, 0.78, 0.78, 0.5, 0.78], border: { type: "solid", pt: 2, color: "FFFFFF" }, objectName: name("structure-table") });
  card(s, 6.75, 5.45, 5.98, 0.85, PALE);
  const co = M.carryover.ag;
  txt(s, [{ text: "Carry-over: ", options: { bold: true } }, { text: `${M.carryover.ads.length} ads from the previous agency's Reach_June'26 campaign kept running in July (${usd(co.spend, 0)}, ${num(co.spend / A.spend * 100, 1)}% of AFTER spend). They are included in AFTER totals.` }],
    { x: 6.95, y: 5.55, w: 5.6, h: 0.65, fontSize: 11, color: DARK, valign: "middle" });
  source(s, META_SRC);
  s.addNotes("Structure counted from the exports: unique campaign names, unique (campaign, ad set) pairs and ad rows. AFTER contains one ad name that appears twice in the same ad set (Spiderman_HM_20.08) with different delivery – treated as two ads.");
}

// =====================================================================================
// 5. WHAT CHANGED – efficiency movements
// =====================================================================================
{
  const s = content("Before vs after", "What changed after July 1",
    `Click efficiency improved sharply (link CTR ${sgn(ch(B.link_ctr, A.link_ctr))}); the cost of impressions and attention rose`);
  const items = [
    ["Link CTR", ch(B.link_ctr, A.link_ctr), "good"],
    ["CTR (all)", ch(B.ctr_all, A.ctr_all), "good"],
    ["CPC (link)", ch(B.cpc, A.cpc), "good"],
    ["Frequency", ch(B.freq, A.freq), "neutral"],
    ["Cost per 1,000 reached", ch(B.cost_per_1k_reach, A.cost_per_1k_reach), "bad"],
    ["Engagement rate", ch(B.eng_rate, A.eng_rate), "bad"],
    ["3s plays per impression", ch(B.v3s_rate, A.v3s_rate), "bad"],
    ["CPM", ch(B.cpm, A.cpm), "bad"],
    ["Cost per engagement", ch(B.cpe, A.cpe), "bad"],
    ["Cost per 1,000 3s plays", ch(B.cost_per_1k_v3s, A.cost_per_1k_v3s), "bad"],
  ].reverse();
  const labels = items.map((i) => i[0]);
  const ser = (v) => items.map((i) => (i[2] === v ? i[1] : 0));
  chartTitle(s, "% change AFTER vs BEFORE – efficiency KPIs", 0.6, 1.9, 7.5);
  s.addChart(pres.charts.BAR, [
    { name: "Improved", labels, values: ser("good") },
    { name: "Declined", labels, values: ser("bad") },
    { name: "Context", labels, values: ser("neutral") },
  ], chartBase({
    x: 0.5, y: 2.2, w: 7.7, h: 4.35, barDir: "bar", barGrouping: "clustered", barOverlapPct: 100, barGapWidthPct: 35,
    chartColors: [GOOD_HEX, BAD_HEX, NEUTRAL_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '+0"%";-0"%";;', dataLabelFontBold: true,
    catAxisLabelPos: "low", valAxisMinVal: -40, valAxisMaxVal: 60, showLegend: true, legendPos: "b", objectName: name("chart-changes"),
  }));
  card(s, 8.6, 1.95, 4.13, 2.15);
  txt(s, "BIGGEST GAIN", { x: 8.85, y: 2.1, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, charSpacing: 1 });
  txt(s, sgn(ch(B.link_ctr, A.link_ctr)) + " link CTR", { x: 8.85, y: 2.38, w: 3.7, h: 0.5, fontSize: 24, bold: true, color: GREEN });
  txt(s, `${num(B.link_ctr, 3)}% → ${num(A.link_ctr, 3)}%. Ads now earn the click: CPC fell ${pa(ch(B.cpc, A.cpc))} even though impressions cost more.`, { x: 8.85, y: 2.95, w: 3.7, h: 1.05, fontSize: 12, color: MID });
  card(s, 8.6, 4.3, 4.13, 2.15);
  txt(s, "BIGGEST DRAG", { x: 8.85, y: 4.45, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: RED, charSpacing: 1 });
  txt(s, sgn(ch(B.cpm, A.cpm)) + " CPM", { x: 8.85, y: 4.73, w: 3.7, h: 0.5, fontSize: 24, bold: true, color: RED });
  txt(s, `${usd(B.cpm)} → ${usd(A.cpm)} per 1,000 impressions. It flows straight into cost per engagement (${sgn(ch(B.cpe, A.cpe))}) and per 1,000 3s plays (${sgn(ch(B.cost_per_1k_v3s, A.cost_per_1k_v3s))}).`, { x: 8.85, y: 5.3, w: 3.7, h: 1.05, fontSize: 12, color: MID });
  source(s, META_SRC + " Lower frequency shown as context: it means wider distribution, not automatically better.");
  s.addNotes(items.slice().reverse().map((i) => `${i[0]}: ${sgn(i[1], 1)}`).join("; "));
}

// =====================================================================================
// 6. INVESTMENT VS PERFORMANCE – scale
// =====================================================================================
{
  const sp = ch(B.spend, A.spend);
  const s = content("Before vs after", "Investment vs performance",
    `Budget grew ${pa(sp)}: link clicks outpaced it, reach kept pace, impressions and engagement did not`);
  const items = [
    ["Amount spent", sp, "budget"],
    ["Link clicks", ch(B.link_clicks, A.link_clicks)],
    ["Reach*", ch(B.reach_sum, A.reach_sum)],
    ["Impressions", ch(B.impr, A.impr)],
    ["Post engagements", ch(B.post_eng, A.post_eng)],
    ["3-sec video plays", ch(B.v3s, A.v3s)],
  ];
  const cls = (i) => (i[2] ? "budget" : i[1] > sp ? "ahead" : i[1] > sp - 5 ? "inline" : "behind");
  const labels = items.map((i) => i[0]);
  const ser = (c) => items.map((i) => (cls(i) === c ? i[1] : 0));
  chartTitle(s, "% change AFTER vs BEFORE – volume metrics vs budget", 0.6, 1.9, 7.6);
  s.addChart(pres.charts.BAR, [
    { name: "Budget", labels, values: ser("budget") },
    { name: "Outpaced budget", labels, values: ser("ahead") },
    { name: "Roughly in line", labels, values: ser("inline") },
    { name: "Lagged budget", labels, values: ser("behind") },
  ], chartBase({
    x: 0.5, y: 2.2, w: 7.7, h: 4.3, barDir: "col", barGrouping: "clustered", barOverlapPct: 100, barGapWidthPct: 45,
    chartColors: ["27251F", GOOD_HEX, AFTER_HEX, BAD_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '+0"%";-0"%";;',
    dataLabelFontBold: true, dataLabelFontSize: 12, catAxisLabelPos: "low", valAxisMinVal: -10, valAxisMaxVal: 65, showLegend: true, legendPos: "b", objectName: name("chart-scale"),
  }));
  const ipd = (p) => T[p].impr / T[p].spend;
  card(s, 8.6, 1.95, 4.13, 1.45, PALE);
  txt(s, "SCALE", { x: 8.85, y: 2.08, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: DARK, charSpacing: 1 });
  txt(s, `More money bought ${sgn(ch(B.reach_sum, A.reach_sum))} reach and ${sgn(ch(B.link_clicks, A.link_clicks))} link clicks – growth that comes partly from budget alone.`, { x: 8.85, y: 2.38, w: 3.7, h: 0.95, fontSize: 12, color: MID });
  card(s, 8.6, 3.55, 4.13, 1.45);
  txt(s, "EFFICIENCY", { x: 8.85, y: 3.68, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, charSpacing: 1 });
  txt(s, `Only link clicks grew faster than spend – the genuine efficiency gain (CPC ${sgn(ch(B.cpc, A.cpc))}).`, { x: 8.85, y: 3.98, w: 3.7, h: 0.95, fontSize: 12, color: MID });
  card(s, 8.6, 5.15, 4.13, 1.4);
  txt(s, "WHY IMPRESSIONS LAGGED", { x: 8.85, y: 5.28, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: RED, charSpacing: 1 });
  txt(s, `Each $1 bought ${num(ipd("AFTER"), 0)} impressions vs ${num(ipd("BEFORE"), 0)} before (${sgn(ch(ipd("BEFORE"), ipd("AFTER")))}), because CPM rose ${pa(ch(B.cpm, A.cpm))}.`, { x: 8.85, y: 5.58, w: 3.7, h: 0.9, fontSize: 12, color: MID });
  source(s, META_SRC + " *Summed ad-level reach.");
  s.addNotes(items.map((i) => `${i[0]}: ${sgn(i[1], 1)}`).join("; ") + `. "Roughly in line" = within 5 points of budget growth.`);
}

// =====================================================================================
// 7. REACH & VISIBILITY
// =====================================================================================
pres.addSection({ title: "Performance deep-dive" });
{
  const s = content("Performance deep-dive", "Reach & visibility",
    `Reach grew ${pa(ch(B.reach_sum, A.reach_sum))} at lower frequency; in Reach campaigns, cost per 1,000 reached fell ${sgn(ch(reachB.cost_per_1k_reach, reachA.cost_per_1k_reach)).replace("−", "")} despite higher CPM`);
  chartTitle(s, "Media cost, USD – BEFORE vs AFTER", 0.6, 1.9, 7.5);
  const labels = ["CPM\nall campaigns", "CPM\nReach campaigns", "Cost / 1,000 reached\nall campaigns", "Cost / 1,000 reached\nReach campaigns"];
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE (Apr–Jun)", labels, values: [B.cpm, reachB.cpm, B.cost_per_1k_reach, reachB.cost_per_1k_reach] },
    { name: "AFTER (Jul–Sep)", labels, values: [A.cpm, reachA.cpm, A.cost_per_1k_reach, reachA.cost_per_1k_reach] },
  ], chartBase({
    x: 0.5, y: 2.2, w: 7.7, h: 3.35, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60, chartColors: [BEFORE_HEX, AFTER_HEX],
    dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00', showLegend: true, legendPos: "t", valAxisMinVal: 0, valAxisMaxVal: 1.0, objectName: name("chart-cpm"),
  }));
  const changes = [[B.cpm, A.cpm], [reachB.cpm, reachA.cpm], [B.cost_per_1k_reach, A.cost_per_1k_reach], [reachB.cost_per_1k_reach, reachA.cost_per_1k_reach]];
  changes.forEach(([b, a], i) => {
    const p = ch(b, a);
    badge(s, 1.05 + i * 1.79, 5.62, 0.95, p, p < 0 ? "good" : "bad");
  });
  const ex = M.reach_after_ex_sept;
  txt(s, `In Reach campaigns, lower frequency (${num(reachB.freq, 2)} → ${num(reachA.freq, 2)}) means fewer paid repeats per person, so cost per person reached fell even as CPM rose. Excluding the two September campaigns with frequency above 2.3, Q3 Reach cost per 1,000 reached was ${usd(ex.cost_per_1k_reach)} (${sgn(ch(reachB.cost_per_1k_reach, ex.cost_per_1k_reach))} vs Q2).`,
    { x: 0.6, y: 6.0, w: 7.6, h: 0.65, fontSize: 11, color: MID });
  const stats = [
    ["REACH*", big(B.reach_sum), big(A.reach_sum), ch(B.reach_sum, A.reach_sum), "good"],
    ["IMPRESSIONS", big(B.impr), big(A.impr), ch(B.impr, A.impr), "good"],
    ["FREQUENCY*", num(B.freq, 2), num(A.freq, 2), ch(B.freq, A.freq), "neutral"],
  ];
  stats.forEach(([l, b, a, p, v], i) => {
    const y = 1.95 + i * 1.5;
    card(s, 8.6, y, 4.13, 1.35);
    txt(s, l, { x: 8.85, y: y + 0.15, w: 2.4, h: 0.25, fontSize: 10.5, bold: true, color: GREY, charSpacing: 1 });
    badge(s, 11.5, y + 0.13, 1.0, p, v);
    txt(s, [{ text: b, options: { color: GREY, fontSize: 20 } }, { text: "  →  ", options: { color: GREY, fontSize: 18 } }, { text: a, options: { bold: true, color: DARK, fontSize: 28 } }], { x: 8.85, y: y + 0.5, w: 3.7, h: 0.6, valign: "middle" });
  });
  source(s, META_SRC + " *Summed ad-level reach, not de-duplicated.");
  s.addNotes(`Reach-objective campaigns: spend ${usd(reachB.spend, 0)} → ${usd(reachA.spend, 0)}; reach results ${big(reachB.reach_sum, 2)} → ${big(reachA.reach_sum, 2)}; CPM ${usd(reachB.cpm, 3)} → ${usd(reachA.cpm, 3)}; cost per 1,000 reached ${usd(reachB.cost_per_1k_reach, 3)} → ${usd(reachA.cost_per_1k_reach, 3)}; excluding EntryLevel_Reach_Sep_26 and HM_Reach_Sep'26: ${usd(ex.cost_per_1k_reach, 3)}.`);
}

// =====================================================================================
// 8. TRAFFIC EFFICIENCY
// =====================================================================================
{
  const cb = M.clicks_by_obj;
  const s = content("Performance deep-dive", "Traffic efficiency",
    `Link clicks ${sgn(ch(B.link_clicks, A.link_clicks))} at ${sgn(ch(B.cpc, A.cpc)).replace("−", "")} lower CPC – Reach ads became a traffic driver`);
  chartTitle(s, "Link clicks by campaign objective", 0.6, 1.9, 6.5);
  const labels = ["BEFORE (Apr–Jun)", "AFTER (Jul–Sep)"];
  const g = (p, o) => cb[p][o] || 0;
  s.addChart(pres.charts.BAR, [
    { name: "Reach campaigns", labels, values: [g("BEFORE", "Reach"), g("AFTER", "Reach")] },
    { name: "Traffic campaigns", labels, values: [g("BEFORE", "Traffic"), g("AFTER", "Traffic")] },
    { name: "Engagement, video & app", labels, values: [g("BEFORE", "Engagement") + g("BEFORE", "Video views"), g("AFTER", "Engagement") + g("AFTER", "Video views") + g("AFTER", "App installs")] },
  ], chartBase({
    x: 0.5, y: 2.2, w: 6.6, h: 3.75, barDir: "col", barGrouping: "stacked", barGapWidthPct: 70, chartColors: [AFTER_HEX, H.accent1, "4A4A4A"],
    dataLabelPosition: "ctr", dataLabelFormatCode: '[>=2000]#,##0;;;', dataLabelColor: "FFFFFF", dataLabelFontBold: true, showLegend: true, legendPos: "b", catAxisLabelFontSize: 12, catAxisLabelFontBold: true, objectName: name("chart-clicks"),
  }));
  txt(s, [{ text: big(B.link_clicks, 1) + " → " + big(A.link_clicks, 1), options: { bold: true } }, { text: " total link clicks" }], { x: 0.6, y: 6.0, w: 6.5, h: 0.3, fontSize: 12, color: DARK });
  const stats = [
    ["LINK CTR", num(B.link_ctr, 3) + "%", num(A.link_ctr, 3) + "%", ch(B.link_ctr, A.link_ctr), "good"],
    ["CPC (LINK)", usd(B.cpc), usd(A.cpc), ch(B.cpc, A.cpc), "good"],
    ["CTR (ALL CLICKS)", num(B.ctr_all, 3) + "%", num(A.ctr_all, 3) + "%", ch(B.ctr_all, A.ctr_all), "good"],
  ];
  stats.forEach(([l, b, a, p, v], i) => {
    const y = 1.95 + i * 1.12;
    card(s, 7.5, y, 5.23, 1.0);
    txt(s, l, { x: 7.72, y: y + 0.13, w: 2.5, h: 0.25, fontSize: 10.5, bold: true, color: GREY, charSpacing: 1 });
    txt(s, [{ text: b, options: { color: GREY, fontSize: 16 } }, { text: "  →  ", options: { color: GREY, fontSize: 14 } }, { text: a, options: { bold: true, color: DARK, fontSize: 22 } }], { x: 7.72, y: y + 0.42, w: 3.6, h: 0.45, valign: "middle" });
    badge(s, 11.5, y + 0.34, 1.0, p, v);
  });
  const pv = M.result_type["Instagram profile visits"].BEFORE, lc = M.result_type["Link clicks"].AFTER;
  card(s, 7.5, 5.35, 5.23, 1.25, PALE);
  txt(s, [{ text: "Not like-for-like: ", options: { bold: true } }, { text: `Q2 Traffic campaigns optimised for Instagram profile visits (${num(pv.results)} visits at ${usd(pv.cpr, 3)}). Q3's Happy Meal Traffic campaign optimised for link clicks (${num(lc.results)} at ${usd(lc.cpr)}), so their results are not compared directly.` }],
    { x: 7.72, y: 5.45, w: 4.85, h: 1.05, fontSize: 11, color: DARK, valign: "middle" });
  source(s, META_SRC);
  s.addNotes(`Link clicks by objective – BEFORE: ${JSON.stringify(cb.BEFORE)}; AFTER: ${JSON.stringify(cb.AFTER)}. Reach-objective link CTR ${num(reachB.link_ctr, 4)}% → ${num(reachA.link_ctr, 4)}%. CTR (all) is rebuilt as Σ(CTR(all) × impressions) ÷ Σ impressions because the export has no 'clicks (all)' column.`);
}

// =====================================================================================
// 9. ENGAGEMENT & VIDEO
// =====================================================================================
{
  const s = content("Performance deep-dive", "Engagement & video",
    `Engagement and video volume stayed flat on a ${pa(ch(B.spend, A.spend))} bigger budget, so the cost of attention rose`);
  chartTitle(s, "Volume, millions", 0.6, 1.9, 4.0);
  const lab1 = ["Post engagements", "3-sec video plays"];
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE", labels: lab1, values: [B.post_eng / 1e6, B.v3s / 1e6] },
    { name: "AFTER", labels: lab1, values: [A.post_eng / 1e6, A.v3s / 1e6] },
  ], chartBase({
    x: 0.5, y: 2.2, w: 4.2, h: 3.0, barDir: "col", barGrouping: "clustered", barGapWidthPct: 55, chartColors: [BEFORE_HEX, AFTER_HEX],
    dataLabelPosition: "outEnd", dataLabelFormatCode: '0.00"M"', showLegend: true, legendPos: "t", valAxisMinVal: 0, valAxisMaxVal: 11, objectName: name("chart-eng-vol"),
  }));
  const rt = M.result_type;
  chartTitle(s, "Cost per 1,000 post engagements, USD – by optimisation goal", 4.95, 1.9, 7.8);
  const lab2 = ["All campaigns", "Post-engagement goal", "Interactions goal (new in Q3)"];
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE", labels: lab2, values: [B.cpe * 1000, rt["Post engagements"].BEFORE.cost_per_post_eng * 1000, 0] },
    { name: "AFTER", labels: lab2, values: [A.cpe * 1000, rt["Post engagements"].AFTER.cost_per_post_eng * 1000, rt.Interactions.AFTER.cost_per_post_eng * 1000] },
  ], chartBase({
    x: 4.85, y: 2.2, w: 7.9, h: 3.0, barDir: "col", barGrouping: "clustered", barGapWidthPct: 55, chartColors: [BEFORE_HEX, AFTER_HEX],
    dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00;;;', showLegend: true, legendPos: "t", valAxisMinVal: 0, valAxisMaxVal: 19, objectName: name("chart-eng-cost"),
  }));
  const ins = [
    [`${num(A.v3s_share_of_eng, 0)}–${num(B.v3s_share_of_eng, 0)}%`, "of post engagements are 3-second video plays, so engagement here is essentially video consumption."],
    [`${num(B.video_v3s_rate, 1)}% → ${num(A.video_v3s_rate, 1)}%`, "3s-play rate on video ads: Q3 videos hooked viewers slightly better (" + sgn(ch(B.video_v3s_rate, A.video_v3s_rate)) + ")."],
    [`${num(B.video_spend_share, 0)}% → ${num(A.video_spend_share, 0)}%`, `video's share of spend fell; with higher CPM, cost per 1,000 3s plays on video ads rose ${pa(ch(B.video_cost_1k_v3s, A.video_cost_1k_v3s))}.`],
  ];
  ins.forEach(([bigt, t], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 5.4, 3.93, 1.2, i === 2 ? PALE : LIGHT);
    txt(s, bigt, { x: x + 0.2, y: 5.5, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: i === 1 ? GREEN : i === 2 ? RED : DARK });
    txt(s, t, { x: x + 0.2, y: 5.9, w: 3.6, h: 0.65, fontSize: 10.5, color: MID });
  });
  source(s, META_SRC + " Interactions is a different Meta result type; compared here on post engagements, the metric both goals report.");
  s.addNotes(`Post-engagement goal: BEFORE ${usd(rt["Post engagements"].BEFORE.cost_per_post_eng, 5)} / AFTER ${usd(rt["Post engagements"].AFTER.cost_per_post_eng, 5)} per engagement. Interactions goal (AFTER, ${rt.Interactions.AFTER.ads} ads, ${usd(rt.Interactions.AFTER.spend, 0)}): ${num(rt.Interactions.AFTER.results)} interactions at ${usd(rt.Interactions.AFTER.cpr, 3)}; ${num(rt.Interactions.AFTER.post_eng)} post engagements at ${usd(rt.Interactions.AFTER.cost_per_post_eng, 4)}. ThruPlay: ${usd(rt.ThruPlay.BEFORE.cpr, 4)} (1 ad) → ${usd(rt.ThruPlay.AFTER.cpr, 4)} (3 ads).`);
}

// =====================================================================================
// 10. CAMPAIGN PERFORMANCE
// =====================================================================================
pres.addSection({ title: "Campaign, ad set and creative" });
const SHORT = {
  "Mixology_Reach_Aug_26": "Mixology (Aug)", "Mixology_Reach_Jul_26": "Mixology (Jul)", "Happy Meal_July'26": "Happy Meal FIFA (Jul)",
  "Generic_Reach_July'2026": "Generic (Jul)", "MCD_Cheddar_sauce_Reach_July'26": "Cheddar Sauce (Jul)", "Generic_Reach_Sep'2026": "Generic (Sep)",
  "MCD Chicken-reach-september": "Chicken influencer (Sep)", "FIFA_Reach_July'26": "FIFA (Jul)", "Generic_Reach_Aug'2026": "Generic (Aug)",
  "MCD&Spiderman_HM_Reach_Aug_26": "Spiderman Happy Meal (Aug)", "MCD_Grimace_Aug_26": "Grimace (Aug)", "Opening_Nizami Mall": "Nizami Mall opening",
  "MCD&SpidermanGenZ_Reach_Aug_26": "Spiderman GenZ (Aug)", "McCafe_Reach_Sep_26": "McCafé (Sep)", "McChicken_Reach_July'2026": "McChicken (Jul)",
  "AOS_McChicken_Reach_AUG'2026_13043682564": "McChicken Android (Aug)", "EntryLevel_Reach_Sep_26": "Entry Level (Sep)",
  "iOS_McChicken_Reach_AUG'2026_13043686683": "McChicken iOS (Aug)", "HM_Reach_Sep'26": "Happy Meal (Sep)",
};
{
  const bench = reachB.cost_per_1k_reach;
  const rc = M.campaigns.AFTER.filter((c) => c.objective === "Reach" && c.spend >= 500 && !c.campaign.includes("June'26"))
    .sort((a, b) => b.cost_per_1k_reach - a.cost_per_1k_reach); // bottom→top in bar chart = worst at bottom
  const s = content("Campaign, ad set and creative", "Campaign performance",
    "FIFA and Mixology set the efficiency bar; September Happy Meal and Entry Level pulled it down");
  chartTitle(s, `Q3 Reach campaigns (≥$500 spend): cost per 1,000 reached, USD`, 0.6, 1.85, 7.8);
  const labels = rc.map((c) => SHORT[c.campaign] || c.campaign);
  s.addChart(pres.charts.BAR, [
    { name: `Cheaper than Q2 average (${usd(bench)})`, labels, values: rc.map((c) => (c.cost_per_1k_reach < bench ? c.cost_per_1k_reach : 0)) },
    { name: "Costlier than Q2 average", labels, values: rc.map((c) => (c.cost_per_1k_reach >= bench ? c.cost_per_1k_reach : 0)) },
  ], chartBase({
    x: 0.5, y: 2.12, w: 7.9, h: 4.55, barDir: "bar", barGrouping: "clustered", barOverlapPct: 100, barGapWidthPct: 30,
    chartColors: [AFTER_HEX, BAD_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00;;;', dataLabelFontSize: 10,
    catAxisLabelFontSize: 10, showLegend: true, legendPos: "b", legendFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 2.2, objectName: name("chart-campaigns"),
  }));
  const cA = Object.fromEntries(M.campaigns.AFTER.map((c) => [c.campaign, c]));
  const hmA = cA["MCD&Spiderman_HM_Reach_Aug_26"], ios = cA["iOS_McChicken_Reach_AUG'2026_13043686683"];
  const fifa = cA["FIFA_Reach_July'26"], mixJ = cA["Mixology_Reach_Jul_26"], mixA = cA["Mixology_Reach_Aug_26"], hmS = cA["HM_Reach_Sep'26"], el = cA["EntryLevel_Reach_Sep_26"];
  card(s, 8.75, 1.95, 3.98, 2.2);
  txt(s, "DRIVERS", { x: 8.97, y: 2.07, w: 3.5, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, charSpacing: 1 });
  txt(s, [
    { text: "FIFA (Jul): ", options: { bold: true } }, { text: `${big(fifa.reach_sum)} reach for ${usd(fifa.spend / 1000, 1)}K at ${usd(fifa.cost_per_1k_reach)} per 1,000, frequency ${num(fifa.freq, 2)}, ${num(fifa.eng_rate, 0)}% engagement rate.`, options: { breakLine: true } },
    { text: "Mixology (Jul + Aug): ", options: { bold: true } }, { text: `${big(mixJ.reach_sum + mixA.reach_sum)} reach at ${usd(mixA.cost_per_1k_reach)}–${usd(mixJ.cost_per_1k_reach)} per 1,000 – the cheapest reach in the account.`, options: { breakLine: true } },
    { text: "Spiderman Happy Meal (Aug): ", options: { bold: true } }, { text: `${big(hmA.reach_sum)} reach at ${usd(hmA.cost_per_1k_reach)} per 1,000, frequency ${num(hmA.freq, 2)}.` },
  ], { x: 8.97, y: 2.37, w: 3.6, h: 1.72, fontSize: 11, color: MID, paraSpaceAfter: 6 });
  card(s, 8.75, 4.3, 3.98, 2.2, PALE);
  txt(s, "DRAGS", { x: 8.97, y: 4.42, w: 3.5, h: 0.25, fontSize: 10.5, bold: true, color: RED, charSpacing: 1 });
  txt(s, [
    { text: "Happy Meal (Sep): ", options: { bold: true } }, { text: `${usd(hmS.spend / 1000, 1)}K at ${usd(hmS.cost_per_1k_reach)} per 1,000, frequency ${num(hmS.freq, 1)}.`, options: { breakLine: true } },
    { text: "Entry Level (Sep): ", options: { bold: true } }, { text: `${usd(el.spend / 1000, 1)}K at ${usd(el.cost_per_1k_reach)} per 1,000, frequency ${num(el.freq, 1)}, ${num(el.eng_rate, 2)}% engagement rate.`, options: { breakLine: true } },
    { text: "McChicken iOS (Aug): ", options: { bold: true } }, { text: `${usd(ios.spend / 1000, 1)}K at ${usd(ios.cost_per_1k_reach)} per 1,000, frequency ${num(ios.freq, 1)}.` },
  ], { x: 8.97, y: 4.72, w: 3.6, h: 1.72, fontSize: 11, color: MID, paraSpaceAfter: 6 });
  source(s, `Q2's three monthly Reach campaigns ran at $0.68–0.72 per 1,000 reached. Campaign names shortened; Reach_June'26 carry-over excluded. Full names and all campaigns: analysis workbook.`);
  s.addNotes(rc.slice().reverse().map((c) => `${c.campaign}: spend ${usd(c.spend, 0)}, cost/1k reached ${usd(c.cost_per_1k_reach, 3)}, freq ${num(c.freq, 2)}`).join("\n"));
}

// =====================================================================================
// 11. LIKE-FOR-LIKE INITIATIVES
// =====================================================================================
{
  const th = M.theme_reach;
  const order = ["Grimace", "Restaurant openings", "Mixology", "FIFA / World Cup", "Chicken", "McCafé"];
  const s = content("Campaign, ad set and creative", "Like-for-like initiatives",
    "Like-for-like: FIFA, Grimace, Mixology and openings got cheaper to reach; Chicken and McCafé got costlier");
  chartTitle(s, "Cost per 1,000 reached, USD – Reach-objective ads, matched initiatives", 0.6, 1.9, 7.4);
  const labels = order.slice().reverse().map((t) => t === "Restaurant openings" ? "Restaurant openings" : t);
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE (Apr–Jun)", labels, values: order.slice().reverse().map((t) => th[t].BEFORE.cost_per_1k_reach) },
    { name: "AFTER (Jul–Sep)", labels, values: order.slice().reverse().map((t) => th[t].AFTER.cost_per_1k_reach) },
  ], chartBase({
    x: 0.5, y: 2.2, w: 7.5, h: 4.4, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 45, chartColors: [BEFORE_HEX, AFTER_HEX],
    dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00', dataLabelFontSize: 10, showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 1.1,
    catAxisLabelFontSize: 11, objectName: name("chart-themes"),
  }));
  // right: change list
  const hdr = (t, x, w, al) => txt(s, t, { x, y: 1.95, w, h: 0.25, fontSize: 9.5, bold: true, color: GREY, align: al || "left" });
  hdr("INITIATIVE", 8.45, 1.6); hdr("SPEND B → A", 10.0, 1.3); hdr("CHANGE", 11.55, 1.1, "center");
  order.forEach((t, i) => {
    const y = 2.3 + i * 0.6, b = th[t].BEFORE, a = th[t].AFTER, p = ch(b.cost_per_1k_reach, a.cost_per_1k_reach);
    card(s, 8.35, y, 4.38, 0.5, i % 2 ? "FFFFFF" : LIGHT);
    txt(s, t, { x: 8.45, y: y + 0.12, w: 1.6, h: 0.28, fontSize: 11, bold: true });
    txt(s, `${usd(b.spend / 1000, 1)}K → ${usd(a.spend / 1000, 1)}K`, { x: 10.0, y: y + 0.12, w: 1.5, h: 0.28, fontSize: 10.5, color: MID });
    badge(s, 11.6, y + 0.09, 1.0, p, p < 0 ? "good" : "bad");
  });
  card(s, 8.35, 5.95, 4.38, 0.68, PALE);
  txt(s, "Q3 Chicken and McCafé ran at higher CPM but also earned link clicks (CTR ~0.09%) that Q2's versions did not.", { x: 8.5, y: 6.0, w: 4.1, h: 0.58, fontSize: 10, color: DARK, valign: "middle" });
  source(s, "Matched only where the product keyword appears in campaign / ad set / ad names in both periods (fifa|world_cup, mixology, toyuq|chicken, cafe, grimace, opening); Reach-objective ads only. Mixology and openings rest on 1 Q2 ad each.");
  s.addNotes(order.map((t) => { const b = th[t].BEFORE, a = th[t].AFTER; return `${t}: ads ${b.ads}→${a.ads}; spend ${usd(b.spend, 0)}→${usd(a.spend, 0)}; CPM ${usd(b.cpm, 3)}→${usd(a.cpm, 3)}; freq ${num(b.freq, 2)}→${num(a.freq, 2)}; cost/1k reached ${usd(b.cost_per_1k_reach, 3)}→${usd(a.cost_per_1k_reach, 3)}; link CTR ${num(b.link_ctr, 3)}%→${num(a.link_ctr, 3)}%`; }).join("\n") + "\nUnmatched work (e.g. Stranger Things, Bizim Burger, McAvto 5.95 in Q2; Spiderman, Entry Level, Generic in Q3) is excluded from this view.");
}

// =====================================================================================
// 12. AD SET PERFORMANCE
// =====================================================================================
{
  const as = M.adsets.AFTER;
  const find = (c, a) => as.find((x) => x.campaign === c && x.adset === a);
  const genAug = find("MCD&Spiderman_HM_Reach_Aug_26", "HappyMeal_Aug_26 _ GenA_interest based");
  const genSep = find("HM_Reach_Sep'26", "HappyMeal_Aug_26 _ GenA_interest based");
  const parAug = find("MCD&Spiderman_HM_Reach_Aug_26", "HappyMeal_Aug_26_Parents of Gen A");
  const parSep = find("HM_Reach_Sep'26", "HappyMeal_Aug_26_Parents");
  const rLo = parSep.cost_per_1k_reach / parAug.cost_per_1k_reach, rHi = genSep.cost_per_1k_reach / genAug.cost_per_1k_reach;
  const s = content("Campaign, ad set and creative", "Ad set performance",
    `The same Happy Meal audiences cost ${num(rLo, 1)}–${num(rHi, 1)}× more to reach in September as frequency climbed`);
  chartTitle(s, "Happy Meal ad sets: cost per 1,000 reached, USD", 0.6, 1.9, 5.8);
  const labels = ["Gen A interest-based", "Parents of Gen A"];
  s.addChart(pres.charts.BAR, [
    { name: "August (Spiderman HM)", labels, values: [genAug.cost_per_1k_reach, parAug.cost_per_1k_reach] },
    { name: "September (One Piece × SpongeBob HM)", labels, values: [genSep.cost_per_1k_reach, parSep.cost_per_1k_reach] },
  ], chartBase({
    x: 0.5, y: 2.2, w: 5.9, h: 3.1, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60, chartColors: [AFTER_HEX, H.accent1],
    dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00', showLegend: true, legendPos: "b", legendFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 2.6, catAxisLabelFontSize: 12, objectName: name("chart-hm"),
  }));
  txt(s, [{ text: "Frequency  ", options: { bold: true } }, { text: `Gen A ${num(genAug.freq, 2)} → ${num(genSep.freq, 2)}   ·   Parents ${num(parAug.freq, 2)} → ${num(parSep.freq, 2)}` }],
    { x: 0.6, y: 5.38, w: 5.8, h: 0.3, fontSize: 12, color: DARK });
  txt(s, "The Gen A ad set name is identical in both months; the September parents ad set is similarly named (\"HappyMeal_Aug_26_Parents\").", { x: 0.6, y: 5.7, w: 5.8, h: 0.5, fontSize: 10, color: GREY });
  // leaderboard
  const big500 = as.filter((x) => x.spend >= 500);
  const reachSets = big500.filter((x) => x.objective === "Reach");
  const bestCtr = reachSets.slice().sort((a, b) => b.link_ctr - a.link_ctr)[0];
  const cheapest = reachSets.slice().sort((a, b) => a.cost_per_1k_reach - b.cost_per_1k_reach)[0];
  const fifa = find("FIFA_Reach_July'26", "FIFA_Reach_July");
  const entry = find("EntryLevel_Reach_Sep_26", "Primary Audience");
  const lal = find("Happy Meal Traffic Aug'26 13043682564", "HM_Look_alike Recent Videos");
  const rowsL = [
    ["good", "Best click-through", bestCtr.adset, `${num(bestCtr.link_ctr, 2)}% link CTR, ${usd(bestCtr.cpc)} CPC (${num(bestCtr.link_ctr / A.link_ctr, 1)}× account CTR)`],
    ["good", "Cheapest reach at scale", cheapest.adset, `${usd(cheapest.cost_per_1k_reach)} per 1,000 reached across ${big(cheapest.reach_sum)} reach`],
    ["good", "Reach + attention", fifa.adset, `${usd(fifa.cost_per_1k_reach)} per 1,000 reached with a ${num(fifa.eng_rate, 0)}% engagement rate`],
    ["bad", "Weakest reach set", `${entry.adset} (Entry Level)`, `${usd(entry.cost_per_1k_reach)} per 1,000 reached, frequency ${num(entry.freq, 2)}, ${num(entry.eng_rate, 2)}% engagement`],
    ["bad", "Over-served audience", lal.adset, `frequency ${num(lal.freq, 1)} and ${usd(lal.cpm)} CPM (${num(lal.cpm / A.cpm, 1)}× account) on the link-click campaign`],
  ];
  rowsL.forEach(([v, k, nm, d], i) => {
    const y = 1.95 + i * 0.78;
    card(s, 6.8, y, 5.93, 0.68, v === "good" ? LIGHT : PALE);
    numCircle(s, 6.95, y + 0.17, 0.34, v === "good" ? "✓" : "!", v === "good" ? GREEN : RED, C.background1);
    txt(s, [{ text: k + ": ", options: { bold: true, color: v === "good" ? GREEN : RED } }, { text: nm, options: { bold: true } }], { x: 7.42, y: y + 0.07, w: 5.2, h: 0.27, fontSize: 11 });
    txt(s, d, { x: 7.42, y: y + 0.35, w: 5.2, h: 0.27, fontSize: 10.5, color: MID });
  });
  const ig = M.platform_before.ig, fb = M.platform_before.fb;
  card(s, 6.8, 5.88, 5.93, 0.75, "FFFFFF");
  txt(s, [{ text: "Q2 ad sets ", options: { bold: true } }, { text: `held one ad each and carried creative names, so audience comparison is only possible for Q3. Their _ig / _fb suffixes do show Facebook Reach ads cost ${sgn(ch(ig.cpm, fb.cpm))} more CPM than Instagram (${usd(fb.cpm)} vs ${usd(ig.cpm)}).` }],
    { x: 6.95, y: 5.9, w: 5.7, h: 0.72, fontSize: 10, color: DARK, valign: "middle" });
  source(s, META_SRC + " Leaderboard: Q3 ad sets with ≥$500 spend.");
  s.addNotes(`Gen A interest-based: Aug ${usd(genAug.spend, 0)} spend, ${usd(genAug.cost_per_1k_reach, 3)}/1k reached, freq ${num(genAug.freq, 2)}; Sep ${usd(genSep.spend, 0)}, ${usd(genSep.cost_per_1k_reach, 3)}, freq ${num(genSep.freq, 2)}. Parents: Aug ${usd(parAug.spend, 0)}, ${usd(parAug.cost_per_1k_reach, 3)}, freq ${num(parAug.freq, 2)}; Sep ${usd(parSep.spend, 0)}, ${usd(parSep.cost_per_1k_reach, 3)}, freq ${num(parSep.freq, 2)}. Best CTR set campaign: ${bestCtr.campaign}; cheapest: ${cheapest.campaign}.`);
}

// =====================================================================================
// 13. CREATIVE / AD PERFORMANCE
// =====================================================================================
{
  const ads = M.ads.AFTER, adsB = M.ads.BEFORE;
  const big300 = ads.filter((a) => a.impr >= 300000);
  const ctrTop = big300.slice().sort((a, b) => b.link_ctr - a.link_ctr).slice(0, 4);
  const engTop = big300.slice().sort((a, b) => b.eng_rate - a.eng_rate).slice(0, 4);
  const pick = (nm, c) => ads.find((a) => a.ad === nm && (!c || a.campaign === c));
  const under = [pick("Entry_Level_#1"), pick("OnePieceXSpongebob_launch", "HM_Reach_Sep'26") && ads.filter((a) => a.ad === "OnePieceXSpongebob_launch").sort((a, b) => b.spend - a.spend)[0], pick("McDonald's Sea Breeze", "FIFA_Reach_July'26") && ads.filter((a) => a.ad === "McDonald's Sea Breeze").sort((a, b) => b.spend - a.spend)[0], pick("visa_17.07")];
  const bBestCtr = adsB.filter((a) => a.impr >= 300000).sort((a, b) => b.link_ctr - a.link_ctr)[0];
  const bBestEng = adsB.filter((a) => a.impr >= 300000).sort((a, b) => b.eng_rate - a.eng_rate)[0];
  const s = content("Campaign, ad set and creative", "Creative / ad performance",
    "Spiderman menu creatives won the click; Mixology and FIFA videos won attention");
  const cols = [
    ["STRONGEST LINK CTR", GREEN, ctrTop.map((a) => [a.ad, `${num(a.link_ctr, 2)}% CTR · ${usd(a.cpc_calc)} CPC${a.objective !== "Reach" ? " · " + a.objective + " goal" : ""}`]), `Q2 best: ${bBestCtr.ad} ${num(bBestCtr.link_ctr, 2)}% (profile-visit goal)`],
    ["STRONGEST ENGAGEMENT & VIDEO", GREEN, engTop.map((a) => [a.ad, `${num(a.eng_rate, 1)}% engagement · ${num(a.v3_rate, 1)}% 3s plays`]), `Q2 best: ${bBestEng.ad} ${num(bBestEng.eng_rate, 1)}% engagement`],
    ["UNDERPERFORMERS", RED, under.map((a) => [a.ad, a.ad === "OnePieceXSpongebob_launch" ? `${usd(a.spend, 0)} · ${usd(a.c1kr)} per 1,000 reached · freq ${num(a.freq, 1)}` : `${usd(a.spend, 0)} · ${num(a.eng_rate, 2)}% engagement · ${num(a.link_clicks || 0)} link clicks`]), "Static reach posts with near-zero response"],
  ];
  cols.forEach(([h, col, items, foot], ci) => {
    const x = 0.6 + ci * 4.1, w = 3.93;
    card(s, x, 1.9, w, 3.62, ci === 2 ? PALE : LIGHT);
    txt(s, h, { x: x + 0.2, y: 2.02, w: w - 0.4, h: 0.25, fontSize: 10.5, bold: true, color: col, charSpacing: 1 });
    items.forEach(([nm, d], i) => {
      const y = 2.38 + i * 0.7;
      numCircle(s, x + 0.2, y + 0.04, 0.32, i + 1, ci === 2 ? RED : GOLD, ci === 2 ? C.background1 : DARK);
      txt(s, nm, { x: x + 0.62, y, w: w - 0.8, h: 0.26, fontSize: 11, bold: true, fit: "shrink" });
      txt(s, d, { x: x + 0.62, y: y + 0.27, w: w - 0.8, h: 0.3, fontSize: 10, color: MID });
    });
    txt(s, foot, { x: x + 0.2, y: 5.18, w: w - 0.4, h: 0.28, fontSize: 9.5, color: GREY, italic: true });
  });
  const cr = M.creator_reach;
  card(s, 0.6, 5.68, 12.13, 0.95, "FFFFFF");
  s.addShape(pres.shapes.BLOCK_ARC, { x: 0.75, y: 5.85, w: 0.6, h: 0.6, fill: { color: GOLD }, line: { type: "none" }, angleRange: [180, 0], arcThicknessRatio: 0.32, objectName: name("arch-icon") });
  txt(s, [
    { text: "Pattern in both periods: ", options: { bold: true } },
    { text: `creator / influencer content earns ~3× the engagement rate of other Reach ads – Q2 ${num(cr.BEFORE.creator.eng_rate, 1)}% vs ${num(cr.BEFORE.other.eng_rate, 1)}%, Q3 ${num(cr.AFTER.creator.eng_rate, 1)}% vs ${num(cr.AFTER.other.eng_rate, 1)}%. Common thread of the click winners: menu / offer-led creatives aimed at GenZ.` },
  ], { x: 1.5, y: 5.72, w: 11.1, h: 0.88, fontSize: 11.5, color: DARK, valign: "middle" });
  source(s, META_SRC + " Rankings use ads with ≥300K impressions. Creator = ad names containing collab / inf / influencer / blogger.");
  s.addNotes("CTR leaders: " + ctrTop.map((a) => `${a.ad} [${a.campaign}] ${num(a.link_ctr, 3)}%`).join("; ") + "\nEngagement leaders: " + engTop.map((a) => `${a.ad} [${a.campaign}] ${num(a.eng_rate, 2)}%`).join("; ") + "\nUnderperformers: " + under.map((a) => `${a.ad} [${a.campaign}] spend ${usd(a.spend)}`).join("; ") + `\nCreator ads Q2: ${cr.BEFORE.ads.join(", ")}; Q3: ${cr.AFTER.ads.join(", ")}.`);
}

// =====================================================================================
// 14. KEY WINS vs OPPORTUNITIES
// =====================================================================================
pres.addSection({ title: "Impact and next steps" });
{
  const cA = Object.fromEntries(M.campaigns.AFTER.map((c) => [c.campaign, c]));
  const th = M.theme_reach, rt = M.result_type;
  const s = content("Impact and next steps", "Key wins & opportunities",
    "Scorecard: what the new setup delivered – and what it has not fixed yet");
  const wins = [
    [sgn(ch(B.link_clicks, A.link_clicks)), "link clicks", `Link CTR ${sgn(ch(B.link_ctr, A.link_ctr))}, CPC ${usd(B.cpc)} → ${usd(A.cpc)}`],
    [sgn(ch(reachB.cost_per_1k_reach, reachA.cost_per_1k_reach)), "cost / 1,000 reached", `Reach campaigns: ${usd(reachB.cost_per_1k_reach)} → ${usd(reachA.cost_per_1k_reach)} at ${sgn(ch(reachB.reach_sum, reachA.reach_sum))} reach`],
    [sgn(ch(B.reach_sum, A.reach_sum)), "reach", `Wider spread: frequency ${num(B.freq, 2)} → ${num(A.freq, 2)}`],
    [sgn(ch(th["FIFA / World Cup"].BEFORE.cost_per_1k_reach, th["FIFA / World Cup"].AFTER.cost_per_1k_reach)), "FIFA reach cost", `Grimace ${sgn(ch(th.Grimace.BEFORE.cost_per_1k_reach, th.Grimace.AFTER.cost_per_1k_reach))}, Mixology ${sgn(ch(th.Mixology.BEFORE.cost_per_1k_reach, th.Mixology.AFTER.cost_per_1k_reach))}`],
    [`${A.campaigns}`, "product-led campaigns", `${A.adsets} audience-named ad sets; new link-click and app-install tests`],
  ];
  const opps = [
    [sgn(ch(B.cpm, A.cpm)), "CPM", `${usd(B.cpm)} → ${usd(A.cpm)}; Reach-campaign CPM ${sgn(ch(reachB.cpm, reachA.cpm))}`],
    [sgn(ch(B.cpe, A.cpe)), "cost / engagement", `Engagements ${sgn(ch(B.post_eng, A.post_eng))} and 3s plays ${sgn(ch(B.v3s, A.v3s))} on more budget`],
    [usd((cA["EntryLevel_Reach_Sep_26"].spend + cA["HM_Reach_Sep'26"].spend) / 1000, 1) + "K", "high-frequency reach", `Sept Entry Level + Happy Meal at ${usd(cA["EntryLevel_Reach_Sep_26"].cost_per_1k_reach)}–${usd(cA["HM_Reach_Sep'26"].cost_per_1k_reach)} per 1,000 reached`],
    [`${num(rt.Interactions.AFTER.cost_per_post_eng / rt["Post engagements"].AFTER.cost_per_post_eng, 1)}×`, "cost / engagement", `Interactions goal vs Post-engagement goal (${usd(rt.Interactions.AFTER.spend / 1000, 1)}K spend)`],
    [`${sgn(ch(th.Chicken.BEFORE.cost_per_1k_reach, th.Chicken.AFTER.cost_per_1k_reach))} / ${sgn(ch(th["McCafé"].BEFORE.cost_per_1k_reach, th["McCafé"].AFTER.cost_per_1k_reach))}`, "reach cost", "Chicken / McCafé vs their Q2 equivalents"],
  ];
  [[wins, "KEY WINS AFTER JULY 1", GREEN, LIGHT, 0.6], [opps, "OPPORTUNITIES & LEARNINGS", RED, PALE, 6.82]].forEach(([list, h, col, fill, x]) => {
    card(s, x, 1.9, 5.91, 4.72, fill);
    txt(s, h, { x: x + 0.25, y: 2.03, w: 5.4, h: 0.28, fontSize: 11, bold: true, color: col, charSpacing: 1 });
    list.forEach(([bn, lab, d], i) => {
      const y = 2.42 + i * 0.83;
      txt(s, bn, { x: x + 0.25, y, w: 1.8, h: 0.42, fontSize: bn.length > 6 ? 17 : 22, bold: true, color: col, valign: "bottom" });
      txt(s, lab, { x: x + 0.25, y: y + 0.42, w: 1.75, h: 0.25, fontSize: 9.5, color: GREY });
      txt(s, d, { x: x + 2.1, y: y + 0.05, w: 3.6, h: 0.62, fontSize: 11.5, color: DARK, valign: "middle" });
    });
  });
  source(s, META_SRC);
}

// =====================================================================================
// 15. NEXT STEPS
// =====================================================================================
{
  const cA = Object.fromEntries(M.campaigns.AFTER.map((c) => [c.campaign, c]));
  const rt = M.result_type, th = M.theme_reach, cr = M.creator_reach;
  const s = content("Impact and next steps", "Next steps · Q4 2026",
    "Q4 priorities: keep the click and reach gains, bring down the cost of attention");
  const recs = [
    ["Put frequency guardrails on Reach flights", `Refresh creative or audience once an ad set passes ~2.0 frequency. September Happy Meal and Entry Level ran at 2.3–2.6 and cost ${usd(cA["EntryLevel_Reach_Sep_26"].cost_per_1k_reach)}–${usd(cA["HM_Reach_Sep'26"].cost_per_1k_reach)} per 1,000 reached vs ${usd(reachA.cost_per_1k_reach)} average.`],
    ["Re-test the Interactions goal", `${usd(rt.Interactions.AFTER.spend / 1000, 1)}K bought engagement at ${usd(rt.Interactions.AFTER.cost_per_post_eng * 1000)} per 1,000 vs ${usd(rt["Post engagements"].AFTER.cost_per_post_eng * 1000)} with the Post-engagement goal. Run a head-to-head before scaling.`],
    ["Scale the click winners", "Spiderman GenZ menu creatives delivered 0.32–0.55% link CTR. Brief Chicken, McCafé and value offers in the same menu / offer-led style."],
    ["Make creators a standing format", `Creator content earned ~3× the engagement rate in both quarters (Q3 ${num(cr.AFTER.creator.eng_rate, 0)}% vs ${num(cr.AFTER.other.eng_rate, 0)}%). Plan a creator slot for each product launch.`],
    ["Win back CPM on Chicken & McCafé", `Reach cost rose ${pa(ch(th.Chicken.BEFORE.cost_per_1k_reach, th.Chicken.AFTER.cost_per_1k_reach))} and ${pa(ch(th["McCafé"].BEFORE.cost_per_1k_reach, th["McCafé"].AFTER.cost_per_1k_reach))}. Test broader audiences and placements; in Q2, Facebook Reach CPM ran ~30% above Instagram.`],
    ["Upgrade measurement", "Monthly exports with time breakdown, de-duplicated campaign reach, placement split and app-install results; one primary KPI per objective."],
  ];
  recs.forEach(([h, d], i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.95 + Math.floor(i / 3) * 2.38, w = 3.93;
    card(s, x, y, w, 2.22, i < 2 ? PALE : LIGHT);
    numCircle(s, x + 0.22, y + 0.22, 0.5, i + 1, i < 2 ? RED : GOLD, i < 2 ? C.background1 : DARK);
    txt(s, h, { x: x + 0.85, y: y + 0.2, w: w - 1.05, h: 0.58, fontSize: 14, bold: true, valign: "middle" });
    txt(s, d, { x: x + 0.22, y: y + 0.9, w: w - 0.44, h: 1.25, fontSize: 11, color: MID });
  });
  source(s, "Red = fix first (largest efficiency gap); gold = build on what works. Frequency threshold is a proposed starting guardrail, to be validated in Q4.");
}

// =====================================================================================
// 16. APPENDIX – methodology
// =====================================================================================
pres.addSection({ title: "Appendix" });
{
  const s = content("Appendix", "Appendix · data & methodology", "How the numbers were built – and what the exports cannot show");
  const mis = M.missing;
  const left = [
    ["Sources", `Two Meta Ads Manager exports, one row per ad: BEFORE ${B.ads} ads (1 Apr–30 Jun 2026), AFTER ${A.ads} ads (1 Jul–30 Sep 2026). Same 19 columns in both.`],
    ["Weighted rates", "CPM, CPC, CTR, frequency, cost per result and engagement rates are recalculated from summed spend, impressions, clicks and reach – never averaged across rows."],
    ["CTR (all)", "The export has no 'clicks (all)' column, so clicks were rebuilt as CTR(all) × impressions per ad, then summed."],
    ["Reach", "Period reach is the sum of ad-level reach. People reached by several ads are counted more than once, so it overstates unique reach; the method is the same for both periods."],
  ];
  const right = [
    ["Results", "Result types differ by objective (Reach, profile visits, link clicks, post engagements, interactions, ThruPlay, 2-sec views) and are never added together. Cost per result is shown per result type."],
    ["Blanks", `Blank link clicks / 3s plays are treated as zero (BEFORE ${mis.BEFORE.link_clicks_blank} / ${mis.BEFORE.v3s_blank} rows, AFTER ${mis.AFTER.link_clicks_blank} / ${mis.AFTER.v3s_blank}). 3 AFTER app-install ads ($95) report no results.`],
    ["Period overlap", `BEFORE includes 2 March-named ads that delivered in April (${usd(M.march_in_before.spend, 0)}); AFTER includes ${M.carryover.ads.length} June-named carry-over ads (${usd(M.carryover.ag.spend, 0)}).`],
    ["Not available", "No daily / monthly breakdown (so no trend lines), no placement split, no conversion or sales data. Campaign names were not assumed equivalent; initiatives matched by product keywords only."],
  ];
  [[left, 0.6], [right, 6.82]].forEach(([list, x]) => {
    list.forEach(([h, d], i) => {
      const y = 1.92 + i * 1.18;
      card(s, x, y, 5.91, 1.06, i % 2 ? "FFFFFF" : LIGHT);
      txt(s, h, { x: x + 0.22, y: y + 0.1, w: 5.5, h: 0.26, fontSize: 12, bold: true, color: RED });
      txt(s, d, { x: x + 0.22, y: y + 0.38, w: 5.5, h: 0.65, fontSize: 10.5, color: MID });
    });
  });
  source(s, "Full calculation tables (KPIs, objectives, result types, campaigns, ad sets, ads, initiatives): analysis_tables.xlsx.");
}

pres.writeFile({ fileName: OUT }).then(async () => {
  await setTheme(OUT, THEME);
  console.log("wrote", OUT, "| logo:", !!LOGO, "| hero:", !!HERO, "| decor:", !!DECOR);
});
