// McDonald's Azerbaijan – TikTok performance review (BEFORE Apr–Jun vs AFTER Jul–Sep 2026)
// Every number on the slides is read from analysis/tiktok/metrics.json (built by tiktok_analysis.py from the two TikTok Ads exports).
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { setTheme } = require("./set_theme.js");

const ROOT = __dirname;
const M = JSON.parse(fs.readFileSync(path.join(ROOT, "..", "analysis", "tiktok", "metrics.json"), "utf8"));
const OUT = process.argv[2] || path.join(ROOT, "..", "tiktok", "McDonalds_AZ_TikTok_Performance_Review_Q2_vs_Q3_2026.pptx");
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
pres.title = "McDonald's Azerbaijan – TikTok Performance Review";
pres.subject = "TikTok Ads: April–June vs July–September 2026";
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
const FOOT = "McDonald's Azerbaijan  ·  TikTok performance review  ·  Apr–Jun vs Jul–Sep 2026";
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

const SRC = `Source: TikTok Ads Manager ad-level exports – BEFORE 1 Apr–30 Jun 2026 (${B.ads} active ads), AFTER 1 Jul–30 Sep 2026 (${A.ads} active ads). Rates recalculated from totals (weighted), not averaged.`;
const obj = M.objective;
const rB = obj.Reach.BEFORE, rA = obj.Reach.AFTER, vvA = obj["Video views"].AFTER;
const adsA = M.ads.AFTER, adsB = M.ads.BEFORE;
const cA = Object.fromEntries(M.campaigns.AFTER.map((c) => [c.campaign, c]));
const sum = (arr, k) => arr.reduce((t, a) => t + (a[k] || 0), 0);
// the three highest-spend AFTER Reach ads (all with 2-sec view rates below 10%)
const BIG3 = adsA.filter((a) => a.objective === "Reach").sort((a, b) => b.spend - a.spend).slice(0, 3);
const big3 = { spend: sum(BIG3, "spend"), v2s: sum(BIG3, "v2s"), impr: sum(BIG3, "impr") };
const AG = (c, a) => M.adgroups.AFTER.find((x) => x.campaign === c && x.adgroup === a);
const AOS_AG = AG("AOS_McChicken_Reach_AUG'2026_13043676216", "Chicken Menu Value"), IOS_AG = AG("IOS_McChicken_Reach_AUG'2026_13043676216", "Chicken Menu Value");
const restA = adsA.filter((a) => !BIG3.includes(a));
const rest = { spend: sum(restA, "spend"), v2s: sum(restA, "v2s"), impr: sum(restA, "impr") };
// ads that recorded 0–1 clicks outside Community interaction (which has no click destination)
const LOWCLICK = adsA.filter((a) => a.clicks <= 1 && a.objective !== "Community interaction");
const lowClickSpend = sum(LOWCLICK, "spend");
const SHORT = {
  "MCchicken_Reach_July'26": "McChicken value (Jul)", "FIFA_Reach_July'26": "FIFA (Jul)", "McCafe_Reach_Sep_26": "McCafé (Sep)",
  "Mixology_Reach_July'26": "Mixology launch (Jul)", "HM_Reach_Sep'26": "Happy Meal (Sep)", "MCD&Spiderman_HM_Reach_Aug_26": "Spiderman Happy Meal (Aug)",
  "Mixology_Reach_Aug'26": "Mixology (Aug)", "MCD&SpidermanGenZ_Reach_Aug_26": "Spiderman GenZ (Aug)", "MCchicken_Reach_Sep'26": "McChicken influencer (Sep)",
  "Grimace_Reach_July'26": "Grimace (Jul)", "AOS_McChicken_Reach_AUG'2026_13043676216": "McChicken Android (Aug)",
  "IOS_McChicken_Reach_AUG'2026_13043676216": "McChicken iOS (Aug)", "MCD_Cheddar_sauce_Reach_July'26": "Cheddar Sauce (Jul)",
};

// =====================================================================================
// 1. COVER
// =====================================================================================
pres.addSection({ title: "Cover" });
{
  const s = pres.addSlide({ masterName: "MCD_COVER", sectionTitle: "Cover" });
  txt(s, "QUARTERLY BUSINESS REVIEW  ·  TIKTOK ADVERTISING", { x: 0.7, y: 1.55, w: 6.9, h: 0.3, fontSize: 12, bold: true, color: RED, charSpacing: 2 });
  s.addText("McDonald's Azerbaijan", { placeholder: "title" });
  txt(s, "TikTok Performance Review", { x: 0.7, y: 2.9, w: 6.9, h: 0.6, fontSize: 30, bold: true, color: RED });
  txt(s, "April–June vs July–September 2026", { x: 0.7, y: 3.6, w: 6.6, h: 0.45, fontSize: 20, color: MID });
  s.addText([{ text: "BEFORE  ", options: { bold: true } }, { text: "1 Apr – 30 Jun · previous agency" }], { isTextBox: true, x: 0.7, y: 4.45, w: 3.55, h: 0.42, margin: 0, fontSize: 12, color: DARK, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.21, fill: { color: LIGHT }, objectName: name("chip") });
  s.addText([{ text: "AFTER  ", options: { bold: true } }, { text: "1 Jul – 30 Sep · current agency" }], { isTextBox: true, x: 4.4, y: 4.45, w: 3.4, h: 0.42, margin: 0, fontSize: 12, color: DARK, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.21, fill: { color: GOLD }, objectName: name("chip") });
  if (LOGO) {
    s.addImage({ path: LOGO, x: 0.7, y: 0.5, w: 0.95, h: 0.84, sizing: { type: "contain", w: 0.95, h: 0.84 }, objectName: name("logo") });
  } else {
    s.addText("Official McDonald's logo\n(placeholder – add assets/mcd_logo.png)", { isTextBox: true, x: 0.7, y: 0.5, w: 2.6, h: 0.75, margin: 0.05, fontSize: 10, color: GREY, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, line: { color: C.accent3, width: 1, dashType: "dash" }, objectName: name("logo-placeholder") });
  }
  s.addText("Agency logo\n(placeholder)", { isTextBox: true, x: 0.7, y: 6.15, w: 1.8, h: 0.6, margin: 0.05, fontSize: 10, color: GREY, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, line: { color: C.accent3, width: 1, dashType: "dash" }, objectName: name("agency-placeholder") });
  txt(s, "Prepared October 2026", { x: 2.75, y: 6.3, w: 3, h: 0.3, fontSize: 11, color: GREY, valign: "middle" });
  if (HERO) {
    s.addImage({ path: HERO, x: 8.1, y: 0.6, w: 4.65, h: 6.3, sizing: { type: "cover", w: 4.65, h: 6.3 }, objectName: name("hero") });
  } else if (DECOR) {
    s.addImage({ path: DECOR, x: 8.1, y: 0.0, w: 5.23, h: 7.5, sizing: { type: "cover", w: 5.23, h: 7.5 }, objectName: name("decor") });
  } else {
    card(s, 8.1, 0.6, 4.65, 6.3, PALE);
    const cx = 8.1 + 2.325, base = 6.9;
    [[4.1, RED, 0.2], [2.95, GOLD, 0.24], [1.8, RED, 0.34]].forEach(([d, col, r]) => {
      s.addShape(pres.shapes.BLOCK_ARC, { x: cx - d / 2, y: base - d / 2, w: d, h: d, fill: { color: col }, line: { type: "none" }, angleRange: [180, 0], arcThicknessRatio: r, objectName: name("cover-arch") });
    });
    s.addShape(pres.shapes.OVAL, { x: 11.55, y: 1.1, w: 0.55, h: 0.55, fill: { color: GOLD }, line: { type: "none" }, objectName: name("dot") });
    s.addShape(pres.shapes.OVAL, { x: 8.6, y: 1.6, w: 0.3, h: 0.3, fill: { color: RED }, line: { type: "none" }, objectName: name("dot") });
  }
  s.addNotes("Cover. BEFORE = 1 Apr–30 Jun 2026 (previous agency). AFTER = 1 Jul–30 Sep 2026 (current agency, took over TikTok on 1 July). Logo and imagery: official assets only – placeholders are marked where files were not supplied.");
}

// =====================================================================================
// 2. EXECUTIVE SUMMARY
// =====================================================================================
pres.addSection({ title: "Executive summary" });
{
  const s = content("Executive summary", "Executive summary",
    `Q3 reached ${pa(ch(B.unique_reach, A.unique_reach))} more people on ${pa(ch(B.spend, A.spend))} less budget; clicks and video views did not keep pace`);
  const rows = [
    ["Reach efficiency is the headline win",
      `Unique reach rose from ${big(B.unique_reach, 2)} to ${big(A.unique_reach, 2)} while spend fell from ${usd(B.spend / 1000, 1)}K to ${usd(A.spend / 1000, 1)}K – cost per 1,000 people reached dropped from ${usd(B.cost_1k_unique)} to ${usd(A.cost_1k_unique)}.`,
      sgn(ch(B.cost_1k_unique, A.cost_1k_unique)), "cost / 1,000 unique reached", "good"],
    ["Less over-exposure, stable media cost",
      `Impressions per person fell from ${num(B.freq_unique, 1)} to ${num(A.freq_unique, 1)} over the quarter, while CPM held almost flat (${usd(B.cpm, 3)} → ${usd(A.cpm, 3)}).`,
      sgn(ch(B.freq_unique, A.freq_unique)), "frequency", "neutral"],
    ["Clicks fell faster than budget",
      `Clicks ${sgn(ch(B.clicks, A.clicks))} vs spend ${sgn(ch(B.spend, A.spend))}: CTR ${num(B.ctr, 3)}% → ${num(A.ctr, 3)}%, CPC ${usd(B.cpc, 3)} → ${usd(A.cpc, 3)}. ${LOWCLICK.length} Q3 ads (${usd(lowClickSpend / 1000, 1)}K) recorded 0–1 clicks.`,
      sgn(ch(B.cpc, A.cpc)), "CPC", "bad"],
    ["Video views are the weakest area",
      `2-second views ${sgn(ch(B.v2s, A.v2s))}; view rate ${num(B.v2s_rate, 1)}% → ${num(A.v2s_rate, 1)}%. The three largest Q3 Reach ads (${num(big3.spend / A.spend * 100, 0)}% of spend) ran at a ${num(big3.v2s / big3.impr * 100, 1)}% view rate.`,
      sgn(ch(B.cost_1k_v2s, A.cost_1k_v2s)), "cost / 1,000 2-sec views", "bad"],
    ["New Video views campaigns are the most efficient video source",
      `${usd(vvA.spend / 1000, 1)}K delivered ${big(vvA.v2s, 1)} 2-sec views at ${usd(vvA.cost_1k_v2s)} per 1,000 – about half the ${usd(rA.cost_1k_v2s)} of Q3 Reach campaigns – with a ${num(vvA.v2s_rate, 0)}% view rate.`,
      num(vvA.v2s_rate, 0) + "%", "view rate, Video views goal", "good"],
  ];
  rows.forEach((r, i) => {
    const y = 1.92 + i * 0.95;
    card(s, 0.6, y, 12.13, 0.82);
    const good = r[4] !== "bad";
    numCircle(s, 0.8, y + 0.16, 0.5, i + 1, good ? GOLD : RED, good ? DARK : C.background1);
    txt(s, r[0], { x: 1.5, y: y + 0.1, w: 8.6, h: 0.3, fontSize: 15, bold: true });
    txt(s, r[1], { x: 1.5, y: y + 0.4, w: 8.7, h: 0.4, fontSize: 11.5, color: MID });
    const col = r[4] === "good" ? GREEN : r[4] === "bad" ? RED : DARK;
    txt(s, r[2], { x: 10.35, y: y + 0.08, w: 2.2, h: 0.45, fontSize: 24, bold: true, color: col, align: "right" });
    txt(s, r[3], { x: 10.15, y: y + 0.52, w: 2.4, h: 0.25, fontSize: 10, color: GREY, align: "right" });
  });
  source(s, SRC + " Unique reach and frequency from the de-duplicated Total row of each export.");
  s.addNotes(`Big-3 Reach ads: ${BIG3.map((a) => `${a.ad} (${usd(a.spend)}, view rate ${num(a.v2s_rate, 1)}%)`).join("; ")}. All other AFTER ads: view rate ${num(rest.v2s / rest.impr * 100, 1)}%, cost per 1,000 2-sec views ${usd(rest.spend / rest.v2s * 1000, 3)}. Low-click ads: ${LOWCLICK.map((a) => `${a.ad} [${a.campaign}] ${usd(a.spend)} – ${a.clicks} clicks`).join("; ")}.`);
}

// =====================================================================================
// 3. KPI SNAPSHOT
// =====================================================================================
pres.addSection({ title: "Before vs after" });
{
  const s = content("Before vs after", "Before vs after · KPI snapshot",
    `At a glance: ${sgn(ch(B.spend, A.spend))} budget, ${sgn(ch(B.unique_reach, A.unique_reach))} unique reach, ${sgn(ch(B.v2s, A.v2s))} video views`);
  const K = [
    ["Amount spent", "spend", (v) => usd(v / 1000, 1) + "K", "neutral"],
    ["Unique reach*", "unique_reach", (v) => big(v, 2), "up"],
    ["Cost / 1K reached*", "cost_1k_unique", (v) => usd(v), "down"],
    ["Frequency*", "freq_unique", (v) => num(v, 1), "neutral"],
    ["Impressions", "impr", (v) => big(v, 1), "up"],
    ["CPM", "cpm", (v) => usd(v, 3), "down"],
    ["Clicks", "clicks", (v) => big(v, 1), "up"],
    ["CTR", "ctr", (v) => num(v, 3) + "%", "up"],
    ["CPC", "cpc", (v) => usd(v, 3), "down"],
    ["2-sec video views", "v2s", (v) => big(v, 1), "up"],
    ["2-sec view rate", "v2s_rate", (v) => num(v, 1) + "%", "up"],
    ["Cost / 1K views", "cost_1k_v2s", (v) => usd(v), "down"],
  ];
  const w = 1.87, gap = 0.18, h = 2.0;
  K.forEach(([lab, k, fmt, good], i) => {
    const x = 0.6 + (i % 6) * (w + gap), y = 1.88 + Math.floor(i / 6) * (h + 0.2);
    const p = ch(B[k], A[k]);
    const verdict = good === "neutral" ? "neutral" : (good === "up" ? p > 0 : p < 0) ? "good" : "bad";
    card(s, x, y, w, h);
    txt(s, lab.toUpperCase(), { x: x + 0.17, y: y + 0.18, w: w - 0.3, h: 0.25, fontSize: 9.5, bold: true, color: GREY, charSpacing: 0.5 });
    txt(s, fmt(A[k]), { x: x + 0.17, y: y + 0.48, w: w - 0.3, h: 0.5, fontSize: 24, bold: true, color: DARK });
    txt(s, "AFTER (Jul–Sep)", { x: x + 0.17, y: y + 0.98, w: w - 0.3, h: 0.22, fontSize: 9, color: GREY });
    txt(s, [{ text: "Before: ", options: { color: GREY } }, { text: fmt(B[k]), options: { bold: true, color: MID } }], { x: x + 0.17, y: y + 1.22, w: w - 0.3, h: 0.25, fontSize: 11 });
    badge(s, x + 0.17, y + 1.56, 1.0, p, verdict);
  });
  txt(s, [
    { text: "Green = improvement, red = decline, grey = context (budget, frequency).  " },
    { text: "*Unique reach, frequency and cost per 1,000 reached use the de-duplicated account totals TikTok reports in each export's Total row." },
  ], { x: 0.6, y: 6.2, w: 12.1, h: 0.4, fontSize: 9.5, color: GREY });
  source(s, SRC);
  s.addNotes("All values are period totals across every active ad (spend > 0). CPM = spend ÷ impressions × 1,000; CTR = clicks ÷ impressions; CPC = spend ÷ clicks; 2-sec view rate = 2-sec views ÷ impressions; frequency = impressions ÷ unique reach. Ad-row sums match each export's Total row exactly for spend, impressions, clicks and 2-sec views.");
}

// =====================================================================================
// 4. TRANSITION – structure
// =====================================================================================
{
  const s = content("Before vs after", "Transition · 1 July 2026",
    "The account moved from monthly Reach buckets to product-led campaigns with audience-named ad groups");
  chartTitle(s, "Share of spend by campaign objective", 0.6, 1.9, 5.6);
  const objs = ["Reach", "Video views", "Community interaction", "Unlabelled (results ≠ reach)", "App promotion"];
  const objLabels = ["Reach", "Video views", "Community interaction", "Opening (objective not labelled)", "App promotion"];
  const data = objs.map((o, i) => ({
    name: objLabels[i], labels: ["BEFORE", "AFTER"],
    values: ["BEFORE", "AFTER"].map((p) => (obj[o] && obj[o][p] ? obj[o][p].share : 0)),
  }));
  s.addChart(pres.charts.BAR, data, chartBase({
    x: 0.5, y: 2.25, w: 5.8, h: 3.0, barDir: "bar", barGrouping: "percentStacked", chartColors: [H.accent2, H.accent1, "27251F", "9A9A9A", "D9D9D9"],
    dataLabelPosition: "ctr", dataLabelFormatCode: '[>=5]0"%";;;', dataLabelColor: "FFFFFF", dataLabelFontSize: 10, dataLabelFontBold: true,
    showLegend: true, legendPos: "b", legendFontSize: 10, catAxisLabelFontSize: 12, catAxisLabelFontBold: true, barGapWidthPct: 45, objectName: name("chart-mix"),
  }));
  txt(s, `Reach stays the backbone (${num(rB.share, 0)}% → ${num(rA.share, 0)}%). Community interaction was cut to ${num(obj["Community interaction"].AFTER.share, 1)}%; dedicated Video views campaigns took ${num(vvA.share, 0)}%.`,
    { x: 0.6, y: 5.45, w: 5.6, h: 0.7, fontSize: 12, color: MID });
  const rowsT = [
    ["", "BEFORE · Apr–Jun", "AFTER · Jul–Sep"],
    ["Campaigns", `${B.campaigns} – Reach by month (Reach_April'26…) plus community interaction`, `${A.campaigns} – by product / initiative (FIFA, Mixology, McChicken, McCafé, Happy Meal…)`],
    ["Ad groups", `${B.adgroups}, one per campaign, named "all" (up to 24 ads each)`, `${A.adgroups}, named by audience (Parents, GenZ with GenA interest, 18-34, 18-50…) or OS`],
    ["Active ads", `${B.ads} (${num(B.ads / B.adgroups, 1)} per ad group)`, `${A.ads} (${num(A.ads / A.adgroups, 1)} per ad group)`],
    ["Objectives", "Reach, Community interaction", "Adds Video views, App promotion (test) and an opening campaign"],
  ];
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 11, color: DARK, valign: "middle", margin: [4, 6, 4, 6] }, o) });
  s.addTable(rowsT.map((r, i) => r.map((t, j) => cell(t, {
    bold: i === 0 || j === 0, fill: { color: i === 0 ? (j === 2 ? H.accent2 : j === 1 ? "E4E4E4" : "FFFFFF") : (j === 2 ? "FFF8E1" : j === 1 ? H.lt2 : "FFFFFF") }, fontSize: i === 0 ? 12 : 11,
  }))), { x: 6.75, y: 1.95, w: 5.98, colW: [1.25, 2.3, 2.43], rowH: [0.42, 0.78, 0.78, 0.5, 0.6], border: { type: "solid", pt: 2, color: "FFFFFF" }, objectName: name("structure-table") });
  card(s, 6.75, 5.3, 5.98, 1.0, PALE);
  txt(s, [{ text: "Clean cut-over: ", options: { bold: true } }, { text: `no previous-agency ad spent in July–September. Each export also lists zero-spend legacy ads (${B.zero_spend_rows} BEFORE, ${A.zero_spend_rows} AFTER); they are excluded from all calculations.` }],
    { x: 6.95, y: 5.38, w: 5.6, h: 0.85, fontSize: 11, color: DARK, valign: "middle" });
  source(s, SRC);
  s.addNotes("Structure counted from active rows (spend > 0): unique campaign names, unique (campaign, ad group) pairs and ad rows. The Opening_Nizami Mall campaign reports results (219,241) that differ from its reach, so its objective cannot be confirmed from the export; it is shown separately.");
}

// =====================================================================================
// 5. WHAT CHANGED – efficiency
// =====================================================================================
{
  const s = content("Before vs after", "What changed after July 1",
    `Cost to reach a person fell ${pa(ch(B.cost_1k_unique, A.cost_1k_unique))}; the cost of a click and of a video view rose`);
  const items = [
    ["Cost / 1,000 unique reached", ch(B.cost_1k_unique, A.cost_1k_unique), "good"],
    ["Cost / 1,000 reached (ad level)", ch(B.cost_1k_reach, A.cost_1k_reach), "good"],
    ["Frequency (per person)", ch(B.freq_unique, A.freq_unique), "neutral"],
    ["CPM", ch(B.cpm, A.cpm), "bad"],
    ["CTR", ch(B.ctr, A.ctr), "bad"],
    ["CPC", ch(B.cpc, A.cpc), "bad"],
    ["2-sec view rate", ch(B.v2s_rate, A.v2s_rate), "bad"],
    ["Cost / 1,000 2-sec views", ch(B.cost_1k_v2s, A.cost_1k_v2s), "bad"],
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
    chartColors: [GOOD_HEX, BAD_HEX, NEUTRAL_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '+0.0"%";-0.0"%";;', dataLabelFontBold: true,
    catAxisLabelPos: "low", valAxisMinVal: -35, valAxisMaxVal: 40, showLegend: true, legendPos: "b", objectName: name("chart-changes"),
  }));
  card(s, 8.6, 1.95, 4.13, 2.15);
  txt(s, "BIGGEST GAIN", { x: 8.85, y: 2.1, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, charSpacing: 1 });
  txt(s, sgn(ch(B.cost_1k_unique, A.cost_1k_unique)) + " cost / 1K reached", { x: 8.85, y: 2.38, w: 3.7, h: 0.5, fontSize: 22, bold: true, color: GREEN });
  txt(s, `${usd(B.cost_1k_unique)} → ${usd(A.cost_1k_unique)} per 1,000 unique people. With CPM flat, the gain comes from spreading impressions across more people.`, { x: 8.85, y: 2.95, w: 3.7, h: 1.05, fontSize: 12, color: MID });
  card(s, 8.6, 4.3, 4.13, 2.15);
  txt(s, "BIGGEST DRAG", { x: 8.85, y: 4.45, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: RED, charSpacing: 1 });
  txt(s, sgn(ch(B.cost_1k_v2s, A.cost_1k_v2s)) + " cost / 1K views", { x: 8.85, y: 4.73, w: 3.7, h: 0.5, fontSize: 22, bold: true, color: RED });
  txt(s, `${usd(B.cost_1k_v2s)} → ${usd(A.cost_1k_v2s)} per 1,000 2-sec views, because fewer impressions turned into views (${num(B.v2s_rate, 1)}% → ${num(A.v2s_rate, 1)}%).`, { x: 8.85, y: 5.3, w: 3.7, h: 1.05, fontSize: 12, color: MID });
  source(s, SRC + " Frequency shown as context: lower means less repetition per person.");
  s.addNotes(items.slice().reverse().map((i) => `${i[0]}: ${sgn(i[1], 1)}`).join("; "));
}

// =====================================================================================
// 6. INVESTMENT VS PERFORMANCE
// =====================================================================================
{
  const sp = ch(B.spend, A.spend);
  const s = content("Before vs after", "Investment vs performance",
    `Budget fell ${pa(sp)}: unique reach still grew ${pa(ch(B.unique_reach, A.unique_reach))}, but clicks and views fell further than spend`);
  const items = [
    ["Amount spent", sp, "budget"],
    ["Unique reach", ch(B.unique_reach, A.unique_reach)],
    ["Impressions", ch(B.impr, A.impr)],
    ["Clicks", ch(B.clicks, A.clicks)],
    ["2-sec video views", ch(B.v2s, A.v2s)],
  ];
  const cls = (i) => (i[2] ? "budget" : i[1] > sp + 5 ? "ahead" : i[1] >= sp - 5 ? "inline" : "behind");
  const labels = items.map((i) => i[0]);
  const ser = (c) => items.map((i) => (cls(i) === c ? i[1] : 0));
  chartTitle(s, "% change AFTER vs BEFORE – volume vs budget", 0.6, 1.9, 7.6);
  s.addChart(pres.charts.BAR, [
    { name: "Budget", labels, values: ser("budget") },
    { name: "Outperformed budget", labels, values: ser("ahead") },
    { name: "In line with budget", labels, values: ser("inline") },
    { name: "Fell more than budget", labels, values: ser("behind") },
  ], chartBase({
    x: 0.5, y: 2.2, w: 7.7, h: 4.3, barDir: "col", barGrouping: "clustered", barOverlapPct: 100, barGapWidthPct: 45,
    chartColors: ["27251F", GOOD_HEX, AFTER_HEX, BAD_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '+0"%";-0"%";;',
    dataLabelFontBold: true, dataLabelFontSize: 12, catAxisLabelPos: "low", valAxisMinVal: -40, valAxisMaxVal: 25, showLegend: true, legendPos: "b", objectName: name("chart-scale"),
  }));
  const rpd = (p) => T[p].unique_reach / T[p].spend;
  card(s, 8.6, 1.95, 4.13, 1.45);
  txt(s, "GENUINE EFFICIENCY", { x: 8.85, y: 2.08, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, charSpacing: 1 });
  txt(s, `Each $1 reached ${num(rpd("AFTER"), 0)} unique people vs ${num(rpd("BEFORE"), 0)} before (${sgn(ch(rpd("BEFORE"), rpd("AFTER")))}) – a gain made with less money, not more.`, { x: 8.85, y: 2.38, w: 3.7, h: 0.95, fontSize: 12, color: MID });
  card(s, 8.6, 3.55, 4.13, 1.45, PALE);
  txt(s, "SCALE EFFECT", { x: 8.85, y: 3.68, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: DARK, charSpacing: 1 });
  txt(s, `Impressions (${sgn(ch(B.impr, A.impr))}) moved in line with the budget cut, as expected with a flat CPM.`, { x: 8.85, y: 3.98, w: 3.7, h: 0.95, fontSize: 12, color: MID });
  card(s, 8.6, 5.15, 4.13, 1.4);
  txt(s, "BEYOND THE BUDGET CUT", { x: 8.85, y: 5.28, w: 3.7, h: 0.25, fontSize: 10.5, bold: true, color: RED, charSpacing: 1 });
  txt(s, `Clicks (${sgn(ch(B.clicks, A.clicks))}) and 2-sec views (${sgn(ch(B.v2s, A.v2s))}) fell more than spend – a real efficiency loss.`, { x: 8.85, y: 5.58, w: 3.7, h: 0.9, fontSize: 12, color: MID });
  source(s, SRC + " \"In line\" = within 5 points of the budget change.");
  s.addNotes(items.map((i) => `${i[0]}: ${sgn(i[1], 1)}`).join("; "));
}

// =====================================================================================
// 7. REACH & VISIBILITY
// =====================================================================================
pres.addSection({ title: "Performance deep-dive" });
{
  const s = content("Performance deep-dive", "Reach & visibility",
    `Unique reach grew ${pa(ch(B.unique_reach, A.unique_reach))} with ${pa(ch(B.freq_unique, A.freq_unique))} fewer impressions per person, so each person reached cost ${pa(ch(B.cost_1k_unique, A.cost_1k_unique))} less`);
  chartTitle(s, "Cost per 1,000 unique people reached, USD", 0.6, 1.9, 3.9);
  s.addChart(pres.charts.BAR, [{ name: "Cost / 1,000 reached", labels: ["BEFORE (Apr–Jun)", "AFTER (Jul–Sep)"], values: [B.cost_1k_unique, A.cost_1k_unique] }], chartBase({
    x: 0.5, y: 2.2, w: 3.9, h: 3.4, barDir: "col", barGapWidthPct: 55, chartColors: [BEFORE_HEX, AFTER_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00',
    dataLabelFontSize: 14, dataLabelFontBold: true, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 6.5, catAxisLabelFontSize: 11, objectName: name("chart-c1k"),
  }));
  chartTitle(s, "CPM, USD", 4.65, 1.9, 3.6);
  s.addChart(pres.charts.BAR, [{ name: "CPM", labels: ["BEFORE (Apr–Jun)", "AFTER (Jul–Sep)"], values: [B.cpm, A.cpm] }], chartBase({
    x: 4.55, y: 2.2, w: 3.7, h: 3.4, barDir: "col", barGapWidthPct: 55, chartColors: [BEFORE_HEX, AFTER_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.000',
    dataLabelFontSize: 14, dataLabelFontBold: true, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 0.35, catAxisLabelFontSize: 11, objectName: name("chart-cpm"),
  }));
  badge(s, 1.95, 5.68, 1.0, ch(B.cost_1k_unique, A.cost_1k_unique), "good");
  badge(s, 5.9, 5.68, 1.0, ch(B.cpm, A.cpm), "bad", 1);
  txt(s, "CPM barely moved, so the gain in cost per person comes from distribution: the same budget was spread across more people instead of repeating to the same audience.",
    { x: 0.6, y: 6.08, w: 7.6, h: 0.55, fontSize: 11, color: MID });
  const stats = [
    ["UNIQUE REACH", big(B.unique_reach, 2), big(A.unique_reach, 2), ch(B.unique_reach, A.unique_reach), "good"],
    ["IMPRESSIONS", big(B.impr, 1), big(A.impr, 1), ch(B.impr, A.impr), "neutral"],
    ["FREQUENCY (PER PERSON)", num(B.freq_unique, 1), num(A.freq_unique, 1), ch(B.freq_unique, A.freq_unique), "neutral"],
  ];
  stats.forEach(([l, b, a, p, v], i) => {
    const y = 1.95 + i * 1.5;
    card(s, 8.6, y, 4.13, 1.35);
    txt(s, l, { x: 8.85, y: y + 0.15, w: 2.6, h: 0.25, fontSize: 10.5, bold: true, color: GREY, charSpacing: 1 });
    badge(s, 11.5, y + 0.13, 1.0, p, v);
    txt(s, [{ text: b, options: { color: GREY, fontSize: 20 } }, { text: "  →  ", options: { color: GREY, fontSize: 18 } }, { text: a, options: { bold: true, color: DARK, fontSize: 28 } }], { x: 8.85, y: y + 0.5, w: 3.7, h: 0.6, valign: "middle" });
  });
  source(s, SRC + " Unique reach is TikTok's de-duplicated account reach (Total row), not a sum of ad-level reach.");
  s.addNotes(`Ad-level reach (summed, not de-duplicated) for reference: ${big(B.reach_sum, 1)} → ${big(A.reach_sum, 1)}; cost per 1,000 ad-level reached ${usd(B.cost_1k_reach, 3)} → ${usd(A.cost_1k_reach, 3)} (${sgn(ch(B.cost_1k_reach, A.cost_1k_reach), 1)}). Reach-objective campaigns: CPM ${usd(rB.cpm, 3)} → ${usd(rA.cpm, 3)}, ad-level frequency ${num(rB.freq_adlevel, 2)} → ${num(rA.freq_adlevel, 2)}.`);
}

// =====================================================================================
// 8. TRAFFIC & CLICKS
// =====================================================================================
{
  const s = content("Performance deep-dive", "Traffic & click efficiency",
    `Click efficiency slipped: CTR ${sgn(ch(B.ctr, A.ctr), 1)} and CPC ${sgn(ch(B.cpc, A.cpc), 1)}, with ${usd(lowClickSpend / 1000, 1)}K spent on ads that recorded 0–1 clicks`);
  const labels = ["All campaigns", "Reach campaigns", "Video views (new)"];
  chartTitle(s, "CTR, %", 0.6, 1.9, 3.8);
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE", labels, values: [B.ctr, rB.ctr, 0] },
    { name: "AFTER", labels, values: [A.ctr, rA.ctr, vvA.ctr] },
  ], chartBase({
    x: 0.5, y: 2.2, w: 3.95, h: 3.3, barDir: "col", barGapWidthPct: 55, chartColors: [BEFORE_HEX, AFTER_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '0.000"%";;;',
    dataLabelFontSize: 10, showLegend: true, legendPos: "t", valAxisMinVal: 0, valAxisMaxVal: 0.3, catAxisLabelFontSize: 10, objectName: name("chart-ctr"),
  }));
  chartTitle(s, "CPC, USD", 4.65, 1.9, 3.8);
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE", labels, values: [B.cpc, rB.cpc, 0] },
    { name: "AFTER", labels, values: [A.cpc, rA.cpc, vvA.cpc] },
  ], chartBase({
    x: 4.55, y: 2.2, w: 3.95, h: 3.3, barDir: "col", barGapWidthPct: 55, chartColors: [BEFORE_HEX, AFTER_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.000;;;',
    dataLabelFontSize: 10, showLegend: true, legendPos: "t", valAxisMinVal: 0, valAxisMaxVal: 0.25, catAxisLabelFontSize: 10, objectName: name("chart-cpc"),
  }));
  txt(s, [{ text: big(B.clicks, 1) + " → " + big(A.clicks, 1), options: { bold: true } }, { text: ` total clicks (${sgn(ch(B.clicks, A.clicks), 1)}) on ${sgn(ch(B.spend, A.spend), 1)} spend. Q2 Community interaction ads also record no clicks; they are part of both totals.` }],
    { x: 0.6, y: 5.65, w: 7.9, h: 0.6, fontSize: 11.5, color: DARK });
  card(s, 8.85, 1.95, 3.88, 4.6, PALE);
  txt(s, "0–1 CLICKS IN Q3", { x: 9.08, y: 2.08, w: 3.4, h: 0.25, fontSize: 10.5, bold: true, color: RED, charSpacing: 1 });
  txt(s, usd(lowClickSpend, 0), { x: 9.08, y: 2.36, w: 3.4, h: 0.5, fontSize: 26, bold: true, color: RED });
  txt(s, `${num(lowClickSpend / A.spend * 100, 1)}% of AFTER spend`, { x: 9.08, y: 2.86, w: 3.4, h: 0.25, fontSize: 10.5, color: GREY });
  const lc = LOWCLICK.filter((a) => a.spend >= 1).sort((a, b) => b.spend - a.spend);
  txt(s, lc.map((a, i) => ({ text: `${a.ad} – ${usd(a.spend, 0)}, ${a.clicks} click${a.clicks === 1 ? "" : "s"}`, options: { bullet: true, breakLine: i < lc.length - 1 } })),
    { x: 9.08, y: 3.25, w: 3.45, h: 2.1, fontSize: 10.5, color: DARK, paraSpaceAfter: 4 });
  txt(s, "Whether these ads carried a click destination cannot be seen in the export – worth checking before Q4.", { x: 9.08, y: 5.5, w: 3.45, h: 0.9, fontSize: 10, color: MID, italic: true });
  source(s, SRC + " Clicks = TikTok 'Clicks (destination)'.");
  s.addNotes(`Reach campaigns: CTR ${num(rB.ctr, 3)}% → ${num(rA.ctr, 3)}%, CPC ${usd(rB.cpc, 4)} → ${usd(rA.cpc, 4)}. Video views campaigns (AFTER only): CTR ${num(vvA.ctr, 3)}%, CPC ${usd(vvA.cpc, 4)}.`);
}

// =====================================================================================
// 9. VIDEO VIEWS
// =====================================================================================
{
  const s = content("Performance deep-dive", "Video views",
    `Video efficiency fell: cost per 1,000 2-sec views rose ${pa(ch(B.cost_1k_v2s, A.cost_1k_v2s))}, driven by three large low-view Reach ads`);
  chartTitle(s, "Cost per 1,000 2-sec views, USD – by objective", 0.6, 1.9, 6.0);
  const lab = ["Reach campaigns", "Video views (new in Q3)", "Community interaction"];
  const ci = obj["Community interaction"];
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE", labels: lab, values: [rB.cost_1k_v2s, 0, ci.BEFORE.cost_1k_v2s] },
    { name: "AFTER", labels: lab, values: [rA.cost_1k_v2s, vvA.cost_1k_v2s, ci.AFTER.cost_1k_v2s] },
  ], chartBase({
    x: 0.5, y: 2.2, w: 6.1, h: 3.1, barDir: "col", barGapWidthPct: 55, chartColors: [BEFORE_HEX, AFTER_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00;;;',
    showLegend: true, legendPos: "t", valAxisMinVal: 0, valAxisMaxVal: 3.0, catAxisLabelFontSize: 11, objectName: name("chart-cpv"),
  }));
  chartTitle(s, "2-sec view rate (views ÷ impressions)", 6.95, 1.9, 5.8);
  const vl = ["BEFORE – all ads", "AFTER – 3 largest Reach ads", "AFTER – all other ads"];
  s.addChart(pres.charts.BAR, [{ name: "View rate", labels: vl, values: [B.v2s_rate, big3.v2s / big3.impr * 100, rest.v2s / rest.impr * 100] }], chartBase({
    x: 6.85, y: 2.2, w: 5.95, h: 3.1, barDir: "bar", barGapWidthPct: 45, chartColors: [BEFORE_HEX, H.accent1, AFTER_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"%"',
    dataLabelFontSize: 12, dataLabelFontBold: true, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 30, catAxisLabelFontSize: 11, objectName: name("chart-vr"),
  }));
  const ins = [
    [`${num(big3.spend / A.spend * 100, 0)}% of spend`, `went to ${BIG3.map((a) => a.ad.replace("Xususi toyuq kombolari_03.06", "McChicken value")).join(", ")} – at ${num(big3.v2s / big3.impr * 100, 1)}% view rate and ${usd(big3.spend / big3.v2s * 1000)} per 1,000 views.`, RED],
    [`${num(rest.v2s / rest.impr * 100, 1)}% vs ${num(B.v2s_rate, 1)}%`, `Every other Q3 ad held close to the Q2 view rate, at ${usd(rest.spend / rest.v2s * 1000)} per 1,000 views (Q2: ${usd(B.cost_1k_v2s)}).`, DARK],
    [`${usd(vvA.cost_1k_v2s)} / 1K views`, `New Video views campaigns: ${big(vvA.v2s, 1)} views at a ${num(vvA.v2s_rate, 0)}% view rate – the cheapest video in the account.`, GREEN],
  ];
  ins.forEach(([bigt, t, col], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 5.45, 3.93, 1.15, i === 0 ? PALE : LIGHT);
    txt(s, bigt, { x: x + 0.2, y: 5.53, w: 3.6, h: 0.38, fontSize: 17, bold: true, color: col });
    txt(s, t, { x: x + 0.2, y: 5.9, w: 3.6, h: 0.66, fontSize: 10, color: MID });
  });
  source(s, SRC + " The exports contain 2-sec views only – no 6-sec views, likes, comments, shares or engagement rate – so engagement cannot be assessed.");
  s.addNotes(BIG3.map((a) => `${a.ad} [${a.campaign}]: ${usd(a.spend)}, view rate ${num(a.v2s_rate, 2)}%, frequency ${num(a.freq, 2)}`).join("\n") + `\nVideo views campaign 'Results' (${num(obj["Video views"].AFTER.results)}) are a result type the export does not name, so they are not used.`);
}

// =====================================================================================
// 10. CAMPAIGN PERFORMANCE
// =====================================================================================
pres.addSection({ title: "Campaign, ad group and creative" });
{
  const bench = rB.cost_1k_reach;
  const rc = M.campaigns.AFTER.filter((c) => c.objective === "Reach" && c.spend >= 400).sort((a, b) => b.cost_1k_reach - a.cost_1k_reach);
  const s = content("Campaign, ad group and creative", "Campaign performance",
    "Grimace, Cheddar and Mixology set the efficiency bar; the McChicken value and McCafé campaigns pulled it down");
  chartTitle(s, "Q3 Reach campaigns (≥$400 spend): cost per 1,000 reached (ad level), USD", 0.6, 1.85, 7.8);
  const labels = rc.map((c) => SHORT[c.campaign] || c.campaign);
  s.addChart(pres.charts.BAR, [
    { name: `Cheaper than Q2 Reach average (${usd(bench)})`, labels, values: rc.map((c) => (c.cost_1k_reach < bench ? c.cost_1k_reach : 0)) },
    { name: "Costlier than Q2 average", labels, values: rc.map((c) => (c.cost_1k_reach >= bench ? c.cost_1k_reach : 0)) },
  ], chartBase({
    x: 0.5, y: 2.15, w: 7.9, h: 4.5, barDir: "bar", barGrouping: "clustered", barOverlapPct: 100, barGapWidthPct: 30,
    chartColors: [AFTER_HEX, BAD_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00;;;', dataLabelFontSize: 10,
    catAxisLabelFontSize: 10, showLegend: true, legendPos: "b", legendFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 1.25, objectName: name("chart-campaigns"),
  }));
  const g = cA["Grimace_Reach_July'26"], ch2 = cA["MCD_Cheddar_sauce_Reach_July'26"], mj = cA["Mixology_Reach_July'26"], fifa = cA["FIFA_Reach_July'26"];
  const mcv = cA["MCchicken_Reach_July'26"], caf = cA["McCafe_Reach_Sep_26"], ios = cA["IOS_McChicken_Reach_AUG'2026_13043676216"];
  card(s, 8.75, 1.95, 3.98, 2.3);
  txt(s, "DRIVERS", { x: 8.97, y: 2.07, w: 3.5, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, charSpacing: 1 });
  txt(s, [
    { text: "Grimace / Cheddar (Jul): ", options: { bold: true } }, { text: `${usd(g.cost_1k_reach)} / ${usd(ch2.cost_1k_reach)} per 1,000 at frequency ~1.1.`, options: { breakLine: true } },
    { text: "Mixology launch (Jul): ", options: { bold: true } }, { text: `${big(mj.reach_sum, 1)} reach for ${usd(mj.spend / 1000, 1)}K at ${usd(mj.cost_1k_reach)}.`, options: { breakLine: true } },
    { text: "FIFA (Jul): ", options: { bold: true } }, { text: `largest reach (${big(fifa.reach_sum, 1)}) with a ${num(fifa.v2s_rate, 0)}% view rate.` },
  ], { x: 8.97, y: 2.37, w: 3.6, h: 1.82, fontSize: 10.5, color: MID, paraSpaceAfter: 5 });
  card(s, 8.75, 4.4, 3.98, 2.25, PALE);
  txt(s, "DRAGS", { x: 8.97, y: 4.52, w: 3.5, h: 0.25, fontSize: 10.5, bold: true, color: RED, charSpacing: 1 });
  txt(s, [
    { text: "McChicken value (Jul): ", options: { bold: true } }, { text: `${usd(mcv.spend / 1000, 1)}K in one ad at ${usd(mcv.cost_1k_reach)} per 1,000, frequency ${num(mcv.freq_adlevel, 1)}.`, options: { breakLine: true } },
    { text: "McCafé (Sep): ", options: { bold: true } }, { text: `${usd(caf.spend / 1000, 1)}K at ${usd(caf.cost_1k_reach)}, frequency ${num(caf.freq_adlevel, 1)}, ${num(caf.v2s_rate, 1)}% view rate.`, options: { breakLine: true } },
    { text: "McChicken iOS (Aug): ", options: { bold: true } }, { text: `${usd(ios.cost_1k_reach)} per 1,000 at ${usd(ios.cpm)} CPM.` },
  ], { x: 8.97, y: 4.82, w: 3.6, h: 1.78, fontSize: 10.5, color: MID, paraSpaceAfter: 5 });
  source(s, "Q2 ran three monthly Reach campaigns at $0.36–0.47 per 1,000 reached (ad level). Campaign names shortened; full list in the analysis workbook.");
  s.addNotes(rc.slice().reverse().map((c) => `${c.campaign}: spend ${usd(c.spend, 0)}, cost/1k reached ${usd(c.cost_1k_reach, 3)}, ad-level freq ${num(c.freq_adlevel, 2)}, CTR ${num(c.ctr, 3)}%, view rate ${num(c.v2s_rate, 1)}%`).join("\n"));
}

// =====================================================================================
// 11. LIKE-FOR-LIKE
// =====================================================================================
{
  const th = M.theme_reach;
  const order = ["Mixology", "Grimace", "FIFA / World Cup", "Happy Meal", "Chicken", "McCafé"];
  const s = content("Campaign, ad group and creative", "Like-for-like initiatives",
    "Like-for-like: Mixology and Grimace got cheaper to reach, FIFA held steady, Chicken and McCafé got costlier");
  chartTitle(s, "Cost per 1,000 reached (ad level), USD – Reach-objective ads, matched initiatives", 0.6, 1.9, 7.6);
  const rev = order.slice().reverse();
  s.addChart(pres.charts.BAR, [
    { name: "BEFORE (Apr–Jun)", labels: rev, values: rev.map((t) => th[t].BEFORE.cost_1k_reach) },
    { name: "AFTER (Jul–Sep)", labels: rev, values: rev.map((t) => th[t].AFTER.cost_1k_reach) },
  ], chartBase({
    x: 0.5, y: 2.2, w: 7.5, h: 4.4, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 45, chartColors: [BEFORE_HEX, AFTER_HEX],
    dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00', dataLabelFontSize: 10, showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 0.8,
    catAxisLabelFontSize: 11, objectName: name("chart-themes"),
  }));
  const hdr = (t, x, w, al) => txt(s, t, { x, y: 1.95, w, h: 0.25, fontSize: 9, bold: true, color: GREY, align: al || "left" });
  hdr("INITIATIVE", 8.4, 1.4); hdr("COST/1K", 9.75, 0.95, "center"); hdr("CTR", 10.72, 0.95, "center"); hdr("VIEW RATE", 11.7, 1.0, "center");
  order.forEach((t, i) => {
    const y = 2.3 + i * 0.6, b = th[t].BEFORE, a = th[t].AFTER;
    card(s, 8.3, y, 4.43, 0.5, i % 2 ? "FFFFFF" : LIGHT);
    txt(s, t.replace(" / World Cup", ""), { x: 8.4, y: y + 0.12, w: 1.4, h: 0.28, fontSize: 11, bold: true });
    const p1 = ch(b.cost_1k_reach, a.cost_1k_reach), p2 = ch(b.ctr, a.ctr), p3 = ch(b.v2s_rate, a.v2s_rate);
    badge(s, 9.8, y + 0.09, 0.85, p1, p1 < 0 ? "good" : "bad");
    badge(s, 10.77, y + 0.09, 0.85, p2, p2 > 0 ? "good" : "bad");
    badge(s, 11.77, y + 0.09, 0.85, p3, p3 > 0 ? "good" : "bad");
  });
  card(s, 8.3, 5.95, 4.43, 0.68, PALE);
  txt(s, "Chicken and McCafé costs rose with frequency (Q3 ad-level 2.3 and 2.7); Q3 CTR improved in four of six initiatives.", { x: 8.45, y: 6.0, w: 4.15, h: 0.58, fontSize: 10, color: DARK, valign: "middle" });
  source(s, "Matched only where the product keyword appears in both periods (fifa, mixology, happy meal / HM, chicken / toyuq, cafe / cofe, grimace); Reach-objective ads only. Mixology and Happy Meal rest on one Q2 ad each.");
  s.addNotes(order.map((t) => { const b = th[t].BEFORE, a = th[t].AFTER; return `${t}: ads ${b.ads}→${a.ads}; spend ${usd(b.spend, 0)}→${usd(a.spend, 0)}; cost/1k reached ${usd(b.cost_1k_reach, 3)}→${usd(a.cost_1k_reach, 3)}; CPM ${usd(b.cpm, 3)}→${usd(a.cpm, 3)}; CTR ${num(b.ctr, 3)}%→${num(a.ctr, 3)}%; view rate ${num(b.v2s_rate, 1)}%→${num(a.v2s_rate, 1)}%; freq ${num(b.freq_adlevel, 2)}→${num(a.freq_adlevel, 2)}`; }).join("\n") + "\nUnmatched: Stranger Things, Bizim Burger/Roll, Squishmallows (Q2 only); Spiderman GenZ, Cheddar, McFlurry, Generic (Q3 only).");
}

// =====================================================================================
// 12. AD GROUP PERFORMANCE
// =====================================================================================
{
  const ag = M.adgroups.AFTER;
  const find = (c, a) => ag.find((x) => x.campaign === c && x.adgroup === a);
  const aos = find("AOS_McChicken_Reach_AUG'2026_13043676216", "Chicken Menu Value"), ios = find("IOS_McChicken_Reach_AUG'2026_13043676216", "Chicken Menu Value");
  const hmAugPar = find("MCD&Spiderman_HM_Reach_Aug_26", "HappyMeal_Aug_26"), hmAugGz = find("MCD&Spiderman_HM_Reach_Aug_26", "HappyMeal_GenZ(with GenA interest)_Aug_26");
  const hmSepPar = find("HM_Reach_Sep'26", "Parents"), hmSepGz = find("HM_Reach_Sep'26", "GenZ(with GenA interest)");
  const s = content("Campaign, ad group and creative", "Ad group performance",
    `The same chicken creative cost ${num(ios.cost_1k_reach / aos.cost_1k_reach, 1)}× more per person reached on iOS; GenZ-with-GenA-interest ad groups out-clicked Parents`);
  chartTitle(s, "Chicken_offer_inf_16.08.26 – Android vs iOS ad group", 0.6, 1.9, 5.8);
  const l1 = ["Cost / 1,000 reached ($)", "CPM ($)"];
  s.addChart(pres.charts.BAR, [
    { name: "Android (AOS)", labels: l1, values: [aos.cost_1k_reach, aos.cpm] },
    { name: "iOS", labels: l1, values: [ios.cost_1k_reach, ios.cpm] },
  ], chartBase({
    x: 0.5, y: 2.2, w: 5.9, h: 2.15, barDir: "bar", barGapWidthPct: 40, chartColors: [AFTER_HEX, H.accent1], dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00',
    showLegend: true, legendPos: "r", legendFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 1.0, catAxisLabelFontSize: 11, objectName: name("chart-os"),
  }));
  chartTitle(s, "Happy Meal ad groups – CTR, %", 0.6, 4.45, 5.8);
  const l2 = ["August (Spiderman HM)", "September (One Piece × SpongeBob)"];
  s.addChart(pres.charts.BAR, [
    { name: "Parents / HappyMeal", labels: l2, values: [hmAugPar.ctr, hmSepPar.ctr] },
    { name: "GenZ with GenA interest", labels: l2, values: [hmAugGz.ctr, hmSepGz.ctr] },
  ], chartBase({
    x: 0.5, y: 4.72, w: 5.9, h: 1.95, barDir: "bar", barGapWidthPct: 40, chartColors: [BEFORE_HEX, AFTER_HEX], dataLabelPosition: "outEnd", dataLabelFormatCode: '0.00"%"',
    showLegend: true, legendPos: "r", legendFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 0.55, catAxisLabelFontSize: 10, objectName: name("chart-hm"),
  }));
  const big400 = ag.filter((x) => x.spend >= 400 && x.objective === "Reach");
  const best = big400.slice().sort((a, b) => a.cost_1k_reach - b.cost_1k_reach)[0];
  const fifa = find("FIFA_Reach_July'26", "Fifa_reach_18-50");
  const mcs = find("MCchicken_Reach_Sep'26", "MCchicken_Reach_Sep'26");
  const mcv = find("MCchicken_Reach_July'26", "MC_Value_03.06"), core = find("McCafe_Reach_Sep_26", "Core TA");
  const rowsL = [
    ["good", "Cheapest reach", `${best.adgroup} (${SHORT[best.campaign] || best.campaign})`, `${usd(best.cost_1k_reach)} per 1,000 reached, frequency ${num(best.freq_adlevel, 2)}`],
    ["good", "Reach + video", `${fifa.adgroup}`, `${big(fifa.reach_sum, 1)} reach at ${usd(fifa.cost_1k_reach)}, ${num(fifa.v2s_rate, 1)}% view rate`],
    ["good", "Best video value", `${mcs.adgroup}`, `${num(mcs.v2s_rate, 1)}% view rate at ${usd(mcs.cost_1k_v2s)} per 1,000 views`],
    ["bad", "Weakest reach set", `${mcv.adgroup} (McChicken value)`, `${usd(mcv.cost_1k_reach)} per 1,000, frequency ${num(mcv.freq_adlevel, 1)}, ${num(mcv.v2s_rate, 1)}% view rate`],
    ["bad", "Over-served audience", `${core.adgroup} (McCafé)`, `frequency ${num(core.freq_adlevel, 1)}, ${num(core.v2s_rate, 1)}% view rate on ${usd(core.spend / 1000, 1)}K`],
  ];
  rowsL.forEach(([v, k, nm, d], i) => {
    const y = 1.95 + i * 0.78;
    card(s, 6.8, y, 5.93, 0.68, v === "good" ? LIGHT : PALE);
    numCircle(s, 6.95, y + 0.17, 0.34, v === "good" ? "✓" : "!", v === "good" ? GREEN : RED, C.background1);
    txt(s, [{ text: k + ": ", options: { bold: true, color: v === "good" ? GREEN : RED } }, { text: nm, options: { bold: true } }], { x: 7.42, y: y + 0.07, w: 5.2, h: 0.27, fontSize: 11 });
    txt(s, d, { x: 7.42, y: y + 0.35, w: 5.2, h: 0.27, fontSize: 10.5, color: MID });
  });
  card(s, 6.8, 5.88, 5.93, 0.75, "FFFFFF");
  txt(s, [{ text: "Q2 ad groups ", options: { bold: true } }, { text: "were one per campaign, named \"all\" and holding up to 24 ads, so audience-level comparison is only possible for Q3." }],
    { x: 6.95, y: 5.9, w: 5.7, h: 0.72, fontSize: 10.5, color: DARK, valign: "middle" });
  source(s, SRC + " Leaderboard: Q3 Reach ad groups with ≥$400 spend.");
  s.addNotes(`Android: spend ${usd(aos.spend)}, CPM ${usd(aos.cpm, 3)}, cost/1k reached ${usd(aos.cost_1k_reach, 3)}, view rate ${num(aos.v2s_rate, 1)}%, CTR ${num(aos.ctr, 3)}%. iOS: spend ${usd(ios.spend)}, CPM ${usd(ios.cpm, 3)}, cost/1k ${usd(ios.cost_1k_reach, 3)}, view rate ${num(ios.v2s_rate, 1)}%, CTR ${num(ios.ctr, 3)}%. HM CTR – Aug: HappyMeal_Aug_26 ${num(hmAugPar.ctr, 3)}% vs GenZ(GenA) ${num(hmAugGz.ctr, 3)}%; Sep: Parents ${num(hmSepPar.ctr, 3)}% vs GenZ(GenA) ${num(hmSepGz.ctr, 3)}%. The August 'HappyMeal_Aug_26' ad group is not explicitly named Parents; it is the non-GenZ ad group of that campaign.`);
}

// =====================================================================================
// 13. CREATIVE / AD PERFORMANCE
// =====================================================================================
{
  const bigA = adsA.filter((a) => a.impr >= 1e6), bigB = adsB.filter((a) => a.impr >= 1e6);
  const ctrTop = bigA.slice().sort((a, b) => b.ctr_pct - a.ctr_pct).slice(0, 4);
  const vrTop = bigA.slice().sort((a, b) => b.v2s_rate - a.v2s_rate).slice(0, 4);
  const pick = (nm) => adsA.filter((a) => a.ad === nm).sort((a, b) => b.spend - a.spend)[0];
  const under = [pick("Xususi toyuq kombolari_03.06"), pick("Cofe_#1"), pick("opening_nizami_mall"), pick("Spiderman_launch_03.08")];
  const bCtr = bigB.slice().sort((a, b) => b.ctr_pct - a.ctr_pct)[0], bVr = bigB.slice().sort((a, b) => b.v2s_rate - a.v2s_rate)[0];
  const s = content("Campaign, ad group and creative", "Creative / ad performance",
    "FIFA and chicken-offer ads won the click; dedicated Video views ads won attention");
  const tag = (a) => (a.campaign.startsWith("AOS") ? " (Android)" : a.campaign.startsWith("IOS") ? " (iOS)" : "");
  const cols = [
    ["STRONGEST CTR", GREEN, ctrTop.map((a) => [a.ad + tag(a), `${num(a.ctr_pct, 2)}% CTR · ${usd(a.cpc_calc, 3)} CPC · ${a.objective}`]), `Q2 best: ${bCtr.ad} ${num(bCtr.ctr_pct, 2)}%`],
    ["STRONGEST 2-SEC VIEW RATE", GREEN, vrTop.map((a) => [a.ad + tag(a), `${num(a.v2s_rate, 1)}% view rate · ${usd(a.cost_1k_v2s)} / 1K views · ${a.objective}`]), `Q2 best: ${bVr.ad} ${num(bVr.v2s_rate, 1)}%`],
    ["UNDERPERFORMERS", RED, under.map((a) => [a.ad, a.clicks <= 1 ? `${usd(a.spend, 0)} · ${a.clicks} click${a.clicks === 1 ? "" : "s"} · ${num(a.v2s_rate, 1)}% view rate` : `${usd(a.spend, 0)} · ${num(a.v2s_rate, 1)}% view rate · freq ${num(a.freq, 1)}`]), "High spend with weak views or no clicks"],
  ];
  cols.forEach(([h, col, items, foot], ci) => {
    const x = 0.6 + ci * 4.1, w = 3.93;
    card(s, x, 1.9, w, 3.62, ci === 2 ? PALE : LIGHT);
    txt(s, h, { x: x + 0.2, y: 2.02, w: w - 0.4, h: 0.25, fontSize: 10.5, bold: true, color: col, charSpacing: 1 });
    items.forEach(([nm, d], i) => {
      const y = 2.38 + i * 0.7;
      numCircle(s, x + 0.2, y + 0.04, 0.32, i + 1, ci === 2 ? RED : GOLD, ci === 2 ? C.background1 : DARK);
      txt(s, nm, { x: x + 0.62, y, w: w - 0.8, h: 0.26, fontSize: 10.5, bold: true, fit: "shrink" });
      txt(s, d, { x: x + 0.62, y: y + 0.27, w: w - 0.8, h: 0.3, fontSize: 9.5, color: MID });
    });
    txt(s, foot, { x: x + 0.2, y: 5.18, w: w - 0.4, h: 0.28, fontSize: 9.5, color: GREY, italic: true });
  });
  const inf = adsA.filter((a) => /_inf_/i.test(a.ad));
  const infR = { spend: sum(inf, "spend"), v2s: sum(inf, "v2s"), impr: sum(inf, "impr"), reach: sum(inf, "reach") };
  card(s, 0.6, 5.68, 12.13, 0.95, "FFFFFF");
  s.addShape(pres.shapes.BLOCK_ARC, { x: 0.75, y: 5.85, w: 0.6, h: 0.6, fill: { color: GOLD }, line: { type: "none" }, angleRange: [180, 0], arcThicknessRatio: 0.32, objectName: name("arch-icon") });
  txt(s, [
    { text: "Pattern: ", options: { bold: true } },
    { text: `ads with "inf" (influencer) in the name – ${inf.length} Q3 chicken-offer ads, ${usd(infR.spend / 1000, 1)}K – averaged a ${num(infR.v2s / infR.impr * 100, 1)}% view rate vs ${num(A.v2s_rate, 1)}% for the account. The biggest single-ad budgets (McChicken value, McCafé) had the weakest view rates.` },
  ], { x: 1.5, y: 5.72, w: 11.1, h: 0.88, fontSize: 11.5, color: DARK, valign: "middle" });
  source(s, SRC + " Rankings use ads with ≥1M impressions. Creative content is not inferred from ad names.");
  s.addNotes("CTR leaders: " + ctrTop.map((a) => `${a.ad} [${a.campaign}] ${num(a.ctr_pct, 3)}%`).join("; ") + "\nView-rate leaders: " + vrTop.map((a) => `${a.ad} [${a.campaign}] ${num(a.v2s_rate, 2)}%`).join("; ") + "\nUnderperformers: " + under.map((a) => `${a.ad} [${a.campaign}] ${usd(a.spend)}`).join("; ") + "\nInfluencer-named ads: " + inf.map((a) => `${a.ad} [${a.campaign}]`).join("; "));
}

// =====================================================================================
// 14. KEY WINS & OPPORTUNITIES
// =====================================================================================
pres.addSection({ title: "Impact and next steps" });
{
  const th = M.theme_reach;
  const s = content("Impact and next steps", "Key wins & opportunities",
    "Scorecard: what improved after July 1 – and what needs fixing in Q4");
  const wins = [
    [sgn(ch(B.unique_reach, A.unique_reach)), "unique reach", `${big(B.unique_reach, 2)} → ${big(A.unique_reach, 2)} people on ${sgn(ch(B.spend, A.spend))} spend`],
    [sgn(ch(B.cost_1k_unique, A.cost_1k_unique)), "cost / 1,000 reached", `${usd(B.cost_1k_unique)} → ${usd(A.cost_1k_unique)} per 1,000 unique people`],
    [sgn(ch(B.freq_unique, A.freq_unique)), "frequency", `${num(B.freq_unique, 1)} → ${num(A.freq_unique, 1)} impressions per person – less repetition`],
    [usd(vvA.cost_1k_v2s), "per 1,000 views", `New Video views campaigns: ${num(vvA.v2s_rate, 0)}% view rate, ${big(vvA.v2s, 1)} views`],
    [sgn(ch(th.Grimace.BEFORE.cost_1k_reach, th.Grimace.AFTER.cost_1k_reach)), "Grimace reach cost", `Mixology ${sgn(ch(th.Mixology.BEFORE.cost_1k_reach, th.Mixology.AFTER.cost_1k_reach))}; product-led structure, ${A.adgroups} audience ad groups`],
  ];
  const opps = [
    [sgn(ch(B.cost_1k_v2s, A.cost_1k_v2s)), "cost / 1K views", `View rate ${num(B.v2s_rate, 1)}% → ${num(A.v2s_rate, 1)}%; 2-sec views ${sgn(ch(B.v2s, A.v2s))}`],
    [usd(big3.spend / 1000, 1) + "K", "in 3 low-view ads", `McChicken value + 2 McCafé ads: ${num(big3.v2s / big3.impr * 100, 1)}% view rate, frequency 2.7–3.3`],
    [sgn(ch(B.cpc, A.cpc)), "CPC", `CTR ${num(B.ctr, 3)}% → ${num(A.ctr, 3)}%; clicks ${sgn(ch(B.clicks, A.clicks))}`],
    [usd(lowClickSpend / 1000, 1) + "K", "with 0–1 clicks", `${LOWCLICK.length} ads incl. Spiderman launch, Generic video, Nizami Mall opening`],
    [`${num(IOS_AG.cost_1k_reach / AOS_AG.cost_1k_reach, 1)}×`, "iOS vs Android", "Cost per person reached for the same chicken creative"],
  ];
  [[wins, "KEY WINS AFTER JULY 1", GREEN, LIGHT, 0.6], [opps, "OPPORTUNITIES & LEARNINGS", RED, PALE, 6.82]].forEach(([list, h, col, fill, x]) => {
    card(s, x, 1.9, 5.91, 4.72, fill);
    txt(s, h, { x: x + 0.25, y: 2.03, w: 5.4, h: 0.28, fontSize: 11, bold: true, color: col, charSpacing: 1 });
    list.forEach(([bn, lab, d], i) => {
      const y = 2.42 + i * 0.83;
      txt(s, bn, { x: x + 0.25, y, w: 1.8, h: 0.42, fontSize: bn.length > 6 ? 18 : 22, bold: true, color: col, valign: "bottom" });
      txt(s, lab, { x: x + 0.25, y: y + 0.42, w: 1.8, h: 0.25, fontSize: 9.5, color: GREY });
      txt(s, d, { x: x + 2.1, y: y + 0.05, w: 3.6, h: 0.62, fontSize: 11.5, color: DARK, valign: "middle" });
    });
  });
  source(s, SRC);
}

// =====================================================================================
// 15. Q4 RECOMMENDATIONS
// =====================================================================================
{
  const th = M.theme_reach;
  const s = content("Impact and next steps", "Q4 recommendations · next steps",
    "Q4 priorities: keep the reach efficiency, win back video views and clicks");
  const mcs = cA["MCchicken_Reach_Sep'26"];
  const recs = [
    ["Fix the low-view, high-budget Reach ads", `Three ads took ${num(big3.spend / A.spend * 100, 0)}% of Q3 spend at a ${num(big3.v2s / big3.impr * 100, 1)}% view rate (rest of account ${num(rest.v2s / rest.impr * 100, 1)}%). Rotate in stronger creatives and cap single-ad budgets.`],
    ["Cap frequency on Reach flights", `Across Q3 Reach ads, higher frequency went with higher cost per person reached. McChicken value (3.3) and McCafé (2.7) cost ${usd(cA["MCchicken_Reach_July'26"].cost_1k_reach)} and ${usd(cA["McCafe_Reach_Sep_26"].cost_1k_reach)} per 1,000 vs ${usd(rA.cost_1k_reach)} average.`],
    ["Audit click destinations", `${usd(lowClickSpend / 1000, 1)}K went to ads that recorded 0–1 clicks. Confirm each Reach and Video views ad carries a CTA and landing page before launch.`],
    ["Scale dedicated Video views", `At ${usd(vvA.cost_1k_v2s)} per 1,000 2-sec views and a ${num(vvA.v2s_rate, 0)}% view rate, this objective beat Reach (${usd(rA.cost_1k_v2s)}) for video delivery. Use it for launch videos.`],
    ["Rebalance iOS and Android", `The same chicken creative cost ${usd(IOS_AG.cost_1k_reach)} per 1,000 reached on iOS vs ${usd(AOS_AG.cost_1k_reach)} on Android. Weight reach budgets toward the cheaper OS unless iOS users are a priority.`],
    ["Repeat what worked", `Cheddar, Grimace and Mixology formats (${usd(cA["MCD_Cheddar_sauce_Reach_July'26"].cost_1k_reach)}–${usd(th.Mixology.AFTER.cost_1k_reach)} per 1,000) and influencer-led offers (McChicken Sep: ${num(mcs.v2s_rate, 0)}% view rate). Add 6-sec views and engagement to exports.`],
  ];
  recs.forEach(([h, d], i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.95 + Math.floor(i / 3) * 2.38, w = 3.93;
    card(s, x, y, w, 2.22, i < 3 ? PALE : LIGHT);
    numCircle(s, x + 0.22, y + 0.22, 0.5, i + 1, i < 3 ? RED : GOLD, i < 3 ? C.background1 : DARK);
    txt(s, h, { x: x + 0.85, y: y + 0.2, w: w - 1.05, h: 0.58, fontSize: 14, bold: true, valign: "middle" });
    txt(s, d, { x: x + 0.22, y: y + 0.9, w: w - 0.44, h: 1.25, fontSize: 11, color: MID });
  });
  source(s, "Red = fix first (largest efficiency loss); gold = build on what works. Frequency and budget caps are proposed starting points, to be validated in Q4.");
  s.addNotes(`Correlation between ad-level frequency and cost per 1,000 reached across Q3 Reach ads with ≥$200 spend: r = 0.75 (Q2: 0.65). This is an association, not proof of cause.`);
}

// =====================================================================================
// 16. APPENDIX
// =====================================================================================
pres.addSection({ title: "Appendix" });
{
  const s = content("Appendix", "Appendix · data & methodology", "How the numbers were built – and what the exports cannot show");
  const left = [
    ["Sources", `Two TikTok Ads Manager exports, one row per ad: BEFORE ${B.rows} rows (1 Apr–30 Jun 2026), AFTER ${A.rows} rows (1 Jul–30 Sep 2026), same 15 columns, all in USD.`],
    ["Active ads only", `${B.zero_spend_rows} BEFORE and ${A.zero_spend_rows} AFTER rows have zero spend and zero delivery (legacy ads). They are excluded, leaving ${B.ads} and ${A.ads} active ads.`],
    ["Validation", "Active-row sums match each export's Total row exactly for spend, impressions, clicks and 2-sec views. No duplicate or missing values among active ads."],
    ["Weighted rates", "CPM, CTR, CPC, view rate and cost per view are recalculated from summed spend, impressions, clicks and views – never averaged across rows."],
  ];
  const right = [
    ["Reach & frequency", "Account reach and frequency use TikTok's de-duplicated Total row. Campaign, ad group and ad comparisons use ad-level reach, which counts a person once per ad."],
    ["Results", "The 'Results' column mixes types (reach, community interactions, video-view results, one app result) without naming them, so results are not summed or compared across objectives."],
    ["Objectives", "Inferred from campaign names and results. Opening_Nizami Mall reports results that differ from reach and is shown as 'objective not labelled'."],
    ["Not available", "No 6-sec views, likes, comments, shares, engagement rate, conversions, time breakdown or placement split. Creative content is not inferred from ad names."],
  ];
  [[left, 0.6], [right, 6.82]].forEach(([list, x]) => {
    list.forEach(([h, d], i) => {
      const y = 1.92 + i * 1.18;
      card(s, x, y, 5.91, 1.06, i % 2 ? "FFFFFF" : LIGHT);
      txt(s, h, { x: x + 0.22, y: y + 0.1, w: 5.5, h: 0.26, fontSize: 12, bold: true, color: RED });
      txt(s, d, { x: x + 0.22, y: y + 0.38, w: 5.5, h: 0.65, fontSize: 10.5, color: MID });
    });
  });
  source(s, "Full calculation tables with live formulas: McDonalds_AZ_TikTok_Analysis_Q2_vs_Q3_2026.xlsx.");
}

fs.mkdirSync(path.dirname(OUT), { recursive: true });
pres.writeFile({ fileName: OUT }).then(async () => {
  await setTheme(OUT, THEME);
  console.log("wrote", OUT, "| logo:", !!LOGO, "| hero:", !!HERO, "| decor:", !!DECOR);
});
