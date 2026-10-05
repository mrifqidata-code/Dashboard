// Monthly Review September 2026 · Center Manager Meeting
// Jalankan: NODE_PATH=<folder node_modules berisi pptxgenjs, react, react-dom, react-icons, sharp> node build_deck.js [output.pptx]
// Semua angka September ada di blok DATA di bawah (sumber: Exboard › 02 Acquisition, 08 Class Utilization,
// deck SA League September per 4 Okt 2026; Agustus dari Center Performance Agustus 2026).

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || path.join(__dirname, "Monthly_Review_September_2026_CM.pptx");

// ---------------------------------------------------------------- THEME
const THEME = {
  name: "Sparks Swim Monthly Review",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "13242E", lt1: "FFFFFF", dk2: "0A3D4A", lt2: "EEF4F5",
    accent1: "0B6475", accent2: "35B6B4", accent3: "D9482B", accent4: "E39B2D",
    accent5: "2E9E6A", accent6: "647A82", hlink: "0B6475", folHlink: "647A82",
  },
};
const HEX = THEME.colors; // hex dipakai di opsi yang hanya menerima hex (chart, shadow, tabel)
const TINT = {
  good: "DFF2E8", goodInk: "1C6542", warn: "FCEFD6", warnInk: "7A5200", bad: "FBE3DD", badInk: "9B2C1D",
  teal: "DCEEF1", line: "D5E0E3", grid: "E3EAEC", axis: "4A5C63",
};

// ---------------------------------------------------------------- DATA
const CENTERS = ["KLM", "KWC", "PML", "TMP", "BTU", "HIB"];
const TARGET = 600; // Rp jt per center, sama dengan September
const D = {
  KLM: { name: "Kalimalang", rev: 651, revAug: 677.6, fp: 72, leads: 402, leadsAug: 589, su: 187, suAug: 294, paid: 74, sur: 46.5, cvr: 39.6, cvrD: 3.9, fpp: 89.2, fppD: 0.6, dp: 8, sunp: 142, top: "Diskusi dengan keluarga", renew: 108, util: 58.9, seats: 644, zero: 62, empty: 56, isNew: false },
  KWC: { name: "Karawaci", rev: 473, revAug: 612.4, fp: 63, leads: 374, leadsAug: 468, su: 178, suAug: 234, paid: 70, sur: 47.6, cvr: 39.3, cvrD: -4.5, fpp: 61.4, fppD: -1.7, dp: 27, sunp: 107, top: "Diskusi dengan keluarga", renew: 77, util: 38.2, seats: 1042, zero: 120, empty: 35, isNew: false },
  PML: { name: "Pamulang", rev: 311, revAug: 375.5, fp: 42, leads: 250, leadsAug: 476, su: 130, suAug: 230, paid: 49, sur: 52.0, cvr: 37.7, cvrD: 8.6, fpp: 73.5, fppD: -4.1, dp: 13, sunp: 93, top: "Tidak tertarik / coba saja", renew: 37, util: 38.7, seats: 913, zero: 120, empty: 33, isNew: false },
  TMP: { name: "Taman Palem", rev: 221, revAug: 330.4, fp: 43, leads: 228, leadsAug: 473, su: 113, suAug: 225, paid: 62, sur: 49.6, cvr: 54.9, cvrD: 19.3, fpp: 43.5, fppD: -17.7, dp: 35, sunp: 75, top: "Tidak tertarik / coba saja", renew: 0, util: 8.8, seats: 1303, zero: 269, empty: 0, isNew: true },
  BTU: { name: "Bintaro U-Town", rev: 393, revAug: 392.4, fp: 85, leads: 395, leadsAug: 526, su: 184, suAug: 268, paid: 82, sur: 46.6, cvr: 44.6, cvrD: 3.5, fpp: 79.3, fppD: 20.2, dp: 17, sunp: 135, top: "Jadwal kelas tidak cocok", renew: 1, util: 10.4, seats: 1756, zero: 277, empty: 84, isNew: true },
  HIB: { name: "Harapan Indah", rev: 403, revAug: 325.3, fp: 82, leads: 367, leadsAug: 565, su: 201, suAug: 292, paid: 74, sur: 54.8, cvr: 36.8, cvrD: 12.8, fpp: 79.7, fppD: 2.6, dp: 15, sunp: 156, top: "Diskusi dengan keluarga", renew: 1, util: 9.8, seats: 1768, zero: 257, empty: 84, isNew: true },
};
const AREA = { rev: 2452, revAug: 2713.6, target: 3600, fp: 387, leads: 2016, leadsAug: 3097, su: 993, suAug: 1543, paid: 411, paidAug: 532, sur: 49.3, cvr: 41.4, cvrD: 6.7, fpp: 72.0, fppD: 1.4, dp: 115, sunp: 708, renew: 224, util: 26.4, seats: 7426 };
const SUNP_REASONS = [
  { r: "Diskusi dengan keluarga", n: 131, p: 18.5, c: "HIB (54)" },
  { r: "Harga / kemahalan", n: 100, p: 14.1, c: "BTU (28)" },
  { r: "Jadwal kelas tidak cocok", n: 95, p: 13.4, c: "BTU (50)" },
  { r: "Tidak tertarik / coba saja", n: 73, p: 10.3, c: "KLM (17)" },
  { r: "Rumah jauh / out of area", n: 69, p: 9.7, c: "KLM (18)" },
];

// Asumsi model komitmen (bisa direvisi CM di meeting)
const ASSUME = { dpToFp: 0.6, sunpClose: 0.1 };
// Target rasio per center untuk Oktober
const GOAL = {
  KLM: { cvr: 45, fpp: 90, sur: 48 },
  KWC: { cvr: 45, fpp: 75, sur: 49 },
  PML: { cvr: 43, fpp: 80, sur: 52 },
  TMP: { cvr: 55, fpp: 70, sur: 50 },
  BTU: { cvr: 47, fpp: 85, sur: 48 },
  HIB: { cvr: 42, fpp: 85, sur: 55 },
};
const ceil5 = (n) => Math.ceil(n / 5) * 5;
const PLAN = {};
for (const c of CENTERS) {
  const d = D[c], g = GOAL[c];
  const ticket = d.rev / d.fp; // Rp jt revenue per FP September
  const fpNeed = Math.ceil(TARGET / ticket);
  const fromDP = Math.round(d.dp * ASSUME.dpToFp);
  const fromSUNP = Math.round(d.sunp * ASSUME.sunpClose);
  const fromTrial = fpNeed - fromDP - fromSUNP;
  const suNeed = Math.ceil(fromTrial / ((g.cvr / 100) * (g.fpp / 100)));
  const suCommit = ceil5(Math.max(suNeed, d.su));
  const leadsCommit = ceil5(Math.max(suCommit / (g.sur / 100), d.leads));
  const fpCommit = Math.max(fpNeed, d.fp);
  const fpNeedUp20 = Math.ceil(TARGET / (ticket * 1.2));
  PLAN[c] = { ticket, fpNeed, fromDP, fromSUNP, fromTrial, suNeed, suCommit, leadsCommit, fpCommit, fpNeedUp20 };
}
const FPP_T = Math.round(CENTERS.reduce((a, c) => a + GOAL[c].fpp * D[c].paid, 0) / CENTERS.reduce((a, c) => a + D[c].paid, 0));
const TOT = {
  leads: CENTERS.reduce((s, c) => s + PLAN[c].leadsCommit, 0),
  su: CENTERS.reduce((s, c) => s + PLAN[c].suCommit, 0),
  fp: CENTERS.reduce((s, c) => s + PLAN[c].fpCommit, 0),
};

// ---------------------------------------------------------------- FORMAT
const id = (n) => Math.round(n).toLocaleString("id-ID");
const dec = (n, k = 1) => n.toFixed(k).replace(".", ",");
const pct = (n, k = 1) => `${dec(n, k)}%`;
const chg = (now, before) => (now / before - 1) * 100;
const arrowPct = (v, k = 0) => (Math.abs(v) < 0.5 ? "● 0%" : `${v > 0 ? "▲" : "▼"} ${dec(Math.abs(v), k)}%`);
const arrowPp = (v) => `${v > 0 ? "▲" : "▼"} ${dec(Math.abs(v))} pp`;
const pctTarget = (c) => Math.round((D[c].rev / TARGET) * 100);

// ---------------------------------------------------------------- PRES
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Monthly Review September 2026 - Center Manager Meeting";
pres.subject = "Achievement September 2026, center yang perlu improvement, strategi dan komitmen Oktober";
pres.author = "Area Manager Sparks Swim";
const C = pres.SchemeColor;

const W = 13.333, X0 = 0.6, CW = W - 2 * X0;
const FOOT = "Monthly Review September 2026 · Center Manager Meeting · Sparks Swim";

pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 2.05, w: 7.4, h: 2.0, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: X0, y: 4.25, w: 7.0, h: 0.9, fontSize: 18, color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: X0, y: 0.38, w: CW, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 1.5, valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 0.7, w: CW, h: 0.75, fontSize: 32, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: X0, y: 7.02, w: 9, h: 0.3, fontSize: 10, color: C.accent6, margin: 0, valign: "middle" } } },
  ],
  slideNumber: { x: W - X0 - 0.6, y: 7.02, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 1.2, w: CW, h: 1.2, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});

// ---------------------------------------------------------------- HELPERS
async function icon(Comp, hex) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: `#${hex}`, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return `image/png;base64,${buf.toString("base64")}`;
}
const ICONS = {};
async function loadIcons() {
  const list = ["FaChartLine", "FaFilter", "FaHandshake", "FaMoneyBillWave", "FaChair", "FaBullseye", "FaUsers", "FaCalendarCheck", "FaPhoneAlt", "FaSchool", "FaUserFriends", "FaSwimmer", "FaTrophy", "FaStar", "FaClock", "FaClipboardCheck", "FaSearchDollar", "FaLayerGroup", "FaWallet", "FaRedo", "FaCheckCircle", "FaExclamationTriangle", "FaBullhorn", "FaUserCheck", "FaArrowUp", "FaMedal", "FaCalendarAlt", "FaTasks", "FaComments"];
  for (const n of list) ICONS[n] = await icon(fa[n], "FFFFFF");
}

let objN = 0;
const oname = (p) => `${p}-${++objN}`;

function newSlide(kicker, title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function source(s, text) {
  s.addText(text, { x: X0, y: 6.64, w: CW, h: 0.3, fontSize: 10, color: C.accent6, italic: true, margin: 0, valign: "top", isTextBox: true, objectName: oname("source") });
}
function card(s, x, y, w, h, fill = C.background2) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, rectRadius: 0.12, objectName: oname("card") });
}
function iconDot(s, name, x, y, d = 0.5, fill = C.accent1) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: oname("icon-bg") });
  const p = d * 0.25;
  s.addImage({ data: ICONS[name], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: oname("icon"), altText: name.replace(/^Fa/, "") });
}
function chip(s, text, x, y, w, kind) {
  const map = { good: [C.accent5, C.background1], warn: [C.accent4, C.text1], bad: [C.accent3, C.background1], teal: [C.accent1, C.background1], dark: [C.text2, C.background1] };
  const [fill, ink] = map[kind] || map.teal;
  s.addText(text, { x, y, w, h: 0.3, fontSize: 11, bold: true, color: ink, align: "center", valign: "middle", margin: 0, isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, fill: { color: fill }, line: { type: "none" }, objectName: oname("chip") });
}
function kpiTile(s, x, y, w, h, { label, value, sub, status, statusText }) {
  card(s, x, y, w, h);
  s.addText(label.toUpperCase(), { x: x + 0.25, y: y + 0.2, w: w - 1.9, h: 0.3, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1, margin: 0, isTextBox: true });
  if (status) chip(s, statusText, x + w - 1.55, y + 0.2, 1.3, status);
  s.addText(value, { x: x + 0.25, y: y + 0.55, w: w - 0.5, h: 0.75, fontSize: 36, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(sub, { x: x + 0.25, y: y + 1.35, w: w - 0.5, h: h - 1.5, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
}
const chartText = () => ({
  catAxisLabelColor: TINT.axis, valAxisLabelColor: TINT.axis, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  dataLabelFontSize: 11, dataLabelColor: HEX.dk1, legendFontSize: 12, legendColor: TINT.axis,
  valGridLine: { color: TINT.grid, size: 0.75 }, catGridLine: { style: "none" },
  catAxisLineColor: TINT.line, valAxisLineShow: false,
});
// status helpers (ambang SA League: CVR hijau >=45, kuning 35-45; FP Progress hijau >=85, kuning 75-85)
const stTarget = (p) => (p >= 100 ? "good" : p >= 75 ? "warn" : "bad");
const stCvr = (v) => (v >= 45 ? "good" : v >= 35 ? "warn" : "bad");
const stFpp = (v) => (v >= 85 ? "good" : v >= 75 ? "warn" : "bad");
const stMoM = (v) => (v >= 0 ? "good" : v > -15 ? "warn" : "bad");
const stUtil = (v) => (v >= 50 ? "good" : v >= 30 ? "warn" : "bad");
const cell = (text, opt = {}) => ({ text: String(text), options: { valign: "middle", margin: [0.04, 0.08, 0.04, 0.08], ...opt } });
const stCell = (text, st, opt = {}) => cell(text, { fill: { color: TINT[st] }, color: TINT[`${st}Ink`], bold: true, align: "center", ...opt });
const hdr = (text, opt = {}) => cell(text, { fill: { color: HEX.dk2 }, color: "FFFFFF", bold: true, fontSize: 12, align: "center", ...opt });
const BORDER = { type: "solid", pt: 1, color: "FFFFFF" };

// ---------------------------------------------------------------- SLIDES
async function build() {
  await loadIcons();

  // 1 · COVER -------------------------------------------------------
  pres.addSection({ title: "Pembuka" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Pembuka" });
    s.addText("MONTHLY REVIEW · CENTER MANAGER MEETING", { x: X0, y: 1.45, w: 7.4, h: 0.35, fontSize: 14, bold: true, color: C.accent2, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText("Achievement September 2026 & Komitmen Oktober", { placeholder: "title" });
    s.addText("Apa yang terjadi, center mana yang perlu improvement, strategi kita, dan komitmen per center", { placeholder: "body" });
    s.addText("KLM · KWC · PML · TMP · BTU · HIB", { x: X0, y: 5.6, w: 7.0, h: 0.4, fontSize: 16, bold: true, color: C.accent2, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText("Data per 4 Oktober 2026 · sumber Exboard", { x: X0, y: 6.05, w: 7.0, h: 0.35, fontSize: 14, color: C.background2, margin: 0, isTextBox: true });
    // Donut pencapaian area
    const ach = Math.round((AREA.rev / AREA.target) * 100);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Revenue", labels: ["Tercapai", "Gap"], values: [AREA.rev, AREA.target - AREA.rev] }], {
      x: 8.55, y: 1.35, w: 4.2, h: 4.2, holeSize: 72, chartColors: [HEX.accent2, "1E5563"], showLegend: false, showValue: false, showPercent: false, showLabel: false,
      dataBorder: { pt: 0, color: HEX.dk2 }, objectName: "donut-pencapaian", altText: `Revenue area ${ach}% dari target`,
    });
    s.addText([
      { text: `${ach}%`, options: { fontSize: 60, bold: true, color: C.background1, breakLine: true } },
      { text: "dari target area", options: { fontSize: 14, color: C.background2 } },
    ], { x: 9.15, y: 2.75, w: 3.0, h: 1.4, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(`Rp ${dec(AREA.rev / 1000, 2)} M dari Rp ${dec(AREA.target / 1000, 1)} M`, { x: 8.55, y: 5.65, w: 4.2, h: 0.4, fontSize: 16, bold: true, color: C.background1, align: "center", margin: 0, isTextBox: true });
    s.addNotes("Pembuka. Hari ini kita review September: apa yang terjadi, center mana yang perlu improvement, strategi yang akan kita jalankan, lalu tiap CM menyepakati komitmen Oktober. Angka utama: revenue area Rp 2,45 M = 68% dari target Rp 3,6 M.");
  }

  // 2 · AGENDA ------------------------------------------------------
  {
    const s = newSlide("AGENDA", "Empat hal yang kita bahas hari ini", "Pembuka");
    const items = [
      ["01", "Apa yang terjadi di September", "Revenue, funnel, alasan belum closing, dan yang berjalan baik", "FaChartLine"],
      ["02", "Center yang perlu improvement", "Scorecard per center dan urutan prioritasnya", "FaExclamationTriangle"],
      ["03", "Apa yang akan kita lakukan", "Deep dive strategi: volume, closing, cash & ticket, kapasitas", "FaBullseye"],
      ["04", "Komitmen per center", "Target Oktober Rp 600 jt per center dan 3 aksi utama", "FaHandshake"],
    ];
    const gap = 0.3, w = (CW - 3 * gap) / 4, y = 1.95, h = 4.3;
    items.forEach(([n, t, d, ic], i) => {
      const x = X0 + i * (w + gap);
      card(s, x, y, w, h);
      iconDot(s, ic, x + 0.3, y + 0.35, 0.7, i === 1 ? C.accent3 : C.accent1);
      s.addText(n, { x: x + w - 1.3, y: y + 0.3, w: 1.0, h: 0.8, fontSize: 40, bold: true, color: C.accent2, align: "right", margin: 0, isTextBox: true });
      s.addText(t, { x: x + 0.3, y: y + 1.4, w: w - 0.6, h: 1.1, fontSize: 22, bold: true, color: C.text1, valign: "top", margin: 0, isTextBox: true });
      s.addText(d, { x: x + 0.3, y: y + 2.6, w: w - 0.6, h: 1.4, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    });
    s.addNotes("Urutan sesuai alur meeting: (1) fakta September, (2) center yang perlu improvement, (3) strategi, (4) komitmen tiap center. Komitmen di bagian 4 adalah draft; angkanya kita sepakati bersama di meeting.");
  }

  // ===== SECTION 1 =====
  const S1 = "01 Apa yang terjadi di September";
  pres.addSection({ title: S1 });
  const K1 = "01 · APA YANG TERJADI DI SEPTEMBER";

  // 3 · Ringkasan -------------------------------------------------------
  {
    const s = newSlide(K1, "Volume turun, closing membaik, uang belum lunas", S1);
    const gap = 0.3, w = (CW - 2 * gap) / 3, h = 1.95;
    const onTarget = CENTERS.filter((c) => D[c].rev >= TARGET);
    const tiles = [
      { label: "Revenue area", value: `Rp ${dec(AREA.rev / 1000, 2)} M`, sub: `${Math.round((AREA.rev / AREA.target) * 100)}% dari target Rp 3,6 M · ${arrowPct(chg(AREA.rev, AREA.revAug))} vs Agustus`, status: "bad", statusText: "68% target" },
      { label: "Gap ke target", value: `– Rp ${id(AREA.target - AREA.rev)} jt`, sub: "Gap terbesar di TMP (Rp 379 jt) dan PML (Rp 289 jt)", status: "bad", statusText: "Gap" },
      { label: "Center ≥ target", value: `${onTarget.length} dari 6`, sub: `Hanya KLM (${pctTarget("KLM")}%). Center lain 37–79%`, status: "warn", statusText: "Perlu naik" },
      { label: "Show up trial", value: id(AREA.su), sub: `${arrowPct(chg(AREA.su, AREA.suAug))} vs Agustus (${id(AREA.suAug)}). Leads ${arrowPct(chg(AREA.leads, AREA.leadsAug))}`, status: "bad", statusText: "Volume turun" },
      { label: "CVR show up → paid", value: pct(AREA.cvr), sub: `${arrowPp(AREA.cvrD)} vs Agustus · target 45%`, status: "good", statusText: "Membaik" },
      { label: "FP progress", value: pct(AREA.fpp), sub: `${arrowPp(AREA.fppD)} · ambang sehat 85% · ${AREA.dp} DP belum lunas`, status: "warn", statusText: "Belum sehat" },
    ];
    tiles.forEach((t, i) => kpiTile(s, X0 + (i % 3) * (w + gap), 1.7 + Math.floor(i / 3) * (h + 0.2), w, h, t));
    card(s, X0, 5.98, CW, 0.5, C.background2);
    s.addText([
      { text: "Konteks: ", options: { bold: true } },
      { text: "TMP, BTU, HIB baru bulan ke-2 setelah grand opening (11–12 Agt). Lonjakan leads Agustus tidak berulang di September." },
    ], { x: X0 + 0.25, y: 5.98, w: CW - 0.5, h: 0.5, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    source(s, "Sumber: Exboard › 02 Acquisition (deck SA League September), data per 4 Okt 2026. Agustus: Center Performance Agustus 2026.");
    s.addNotes("Pesan utama September: (1) volume turun tajam, show up -36% dan leads -35% vs Agustus, sebagian karena efek grand opening 3 center baru sudah habis; (2) kualitas closing justru membaik, CVR naik 6,7 pp ke 41,4%; (3) uang belum masuk penuh, FP progress 72% dan 115 murid masih DP. Hasilnya revenue Rp 2,45 M, 68% dari target, gap Rp 1,15 M.");
  }

  // 4 · Revenue per center ---------------------------------------------
  {
    const s = newSlide(K1, "Hanya KLM yang tembus target Rp 600 jt", S1);
    s.addChart(
      [
        { type: pres.charts.BAR, data: [
          { name: "Agustus", labels: CENTERS, values: CENTERS.map((c) => Math.round(D[c].revAug)) },
          { name: "September", labels: CENTERS, values: CENTERS.map((c) => D[c].rev) },
        ], options: { barGrouping: "clustered", chartColors: ["A9CDD3", HEX.accent1], barGapWidthPct: 60 } },
        { type: pres.charts.LINE, data: [{ name: "Target Rp 600 jt", labels: CENTERS, values: CENTERS.map(() => TARGET) }], options: { chartColors: [HEX.accent3], lineSize: 2, lineDash: "dash", lineDataSymbol: "none", showValue: false } },
      ],
      { x: X0, y: 1.65, w: 7.6, h: 4.85, ...chartText(), showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 800, valAxisMajorUnit: 200, valAxisLabelFormatCode: "#,##0",
        showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", showTitle: true, title: "Revenue per center (Rp jt)", titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-revenue" }
    );
    const rows = [[hdr("Center"), hdr("Sep (Rp jt)"), hdr("% target"), hdr("Δ vs Agt")]];
    for (const c of CENTERS) {
      const p = pctTarget(c), m = chg(D[c].rev, D[c].revAug);
      rows.push([cell(c, { bold: true }), cell(id(D[c].rev), { align: "right" }), stCell(`${p}%`, stTarget(p)), stCell(arrowPct(m), stMoM(m))]);
    }
    rows.push([cell("Area", { bold: true, fill: { color: TINT.teal } }), cell(id(AREA.rev), { align: "right", bold: true, fill: { color: TINT.teal } }), stCell("68%", "bad"), stCell(arrowPct(chg(AREA.rev, AREA.revAug)), "warn")]);
    s.addTable(rows, { x: 8.5, y: 1.75, w: 4.23, colW: [0.95, 1.18, 1.0, 1.1], rowH: 0.46, fontSize: 14, color: C.text1, border: BORDER, objectName: "table-revenue" });
    s.addText([
      { text: "HIB satu-satunya yang tumbuh ", options: { bold: true } },
      { text: "(▲24%). TMP turun paling dalam (▼33%) dan paling jauh dari target." },
    ], { x: 8.5, y: 5.6, w: 4.23, h: 0.9, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    source(s, "Target September Rp 600 jt per center (total Rp 3,6 M). Sumber: Exboard › 02 Acquisition, per 4 Okt 2026; Agustus: Center Performance Agustus 2026.");
    s.addNotes("KLM 109% dan jadi Center of the Month. KWC 79%, HIB 67%, BTU 66%, PML 52%, TMP 37%. Gap per center ke Rp 600 jt: KWC -127, PML -289, TMP -379, BTU -207, HIB -197 (Rp jt). HIB satu-satunya yang revenue-nya naik dari Agustus.");
  }

  // 5 · Funnel / show up -------------------------------------------------
  {
    const s = newSlide(K1, "Masalah utama di volume: show up turun 36%", S1);
    s.addChart(pres.charts.BAR, [
      { name: "Agustus", labels: CENTERS, values: CENTERS.map((c) => D[c].suAug) },
      { name: "September", labels: CENTERS, values: CENTERS.map((c) => D[c].su) },
    ], { x: X0, y: 1.65, w: 7.6, h: 4.85, barGrouping: "clustered", barGapWidthPct: 60, chartColors: ["A9CDD3", HEX.accent1], ...chartText(), showLegend: true, legendPos: "b",
      valAxisMinVal: 0, valAxisMaxVal: 350, valAxisMajorUnit: 50, showValue: true, dataLabelPosition: "outEnd", showTitle: true, title: "Show up trial per center", titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-showup" });
    const x = 8.5, w = 4.23;
    const rows = [
      ["Leads", `${id(AREA.leadsAug)} → ${id(AREA.leads)}`, arrowPct(chg(AREA.leads, AREA.leadsAug)), "bad"],
      ["Show up", `${id(AREA.suAug)} → ${id(AREA.su)}`, arrowPct(chg(AREA.su, AREA.suAug)), "bad"],
      ["Paid (DP+FP)", `${id(AREA.paidAug)} → ${id(AREA.paid)}`, arrowPct(chg(AREA.paid, AREA.paidAug)), "warn"],
      ["CVR SU → paid", `${pct(AREA.cvr - AREA.cvrD)} → ${pct(AREA.cvr)}`, arrowPp(AREA.cvrD), "good"],
    ];
    s.addText("FUNNEL AREA · AGT → SEP", { x, y: 1.75, w, h: 0.3, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1, margin: 0, isTextBox: true });
    rows.forEach(([l, v, d, st], i) => {
      const y = 2.15 + i * 0.78;
      card(s, x, y, w, 0.66);
      s.addText(l, { x: x + 0.2, y, w: 1.45, h: 0.66, fontSize: 14, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(v, { x: x + 1.6, y, w: 1.45, h: 0.66, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
      chip(s, d, x + w - 1.15, y + 0.18, 1.0, st);
    });
    s.addText([
      { text: "Rasio stabil, jumlahnya yang hilang. ", options: { bold: true } },
      { text: "Show-up rate tetap ±49% dan CVR naik — penurunan revenue datang dari sedikitnya orang yang masuk funnel." },
    ], { x, y: 5.3, w, h: 1.2, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    source(s, "Sumber: Exboard › 02 Acquisition (SA League September), per 4 Okt 2026; Agustus: Center Performance Agustus 2026. CVR = total paid ÷ show up.");
    s.addNotes("Show up turun di semua center: TMP -50%, PML -43%, KLM -36%, BTU -31%, HIB -31%, KWC -24%. Show-up rate area stabil 49%, CVR naik ke 41,4%. Artinya tim SA closing lebih baik; yang kurang adalah volume leads yang masuk. Ini jadi Strategi 1.");
  }

  // 6 · SUNP reasons ---------------------------------------------------
  {
    const s = newSlide(K1, "708 trial belum closing, 2/3 karena 5 alasan", S1);
    s.addChart(pres.charts.BAR, [{ name: "Kasus SUNP", labels: SUNP_REASONS.map((r) => r.r), values: SUNP_REASONS.map((r) => r.n) }], {
      x: X0, y: 1.65, w: 7.6, h: 4.85, barDir: "bar", chartColors: [HEX.accent1], ...chartText(), catAxisOrientation: "maxMin", catAxisLabelFontSize: 14,
      valAxisHidden: true, valGridLine: { style: "none" }, showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 13, barGapWidthPct: 45, showLegend: false,
      showTitle: true, title: "Top 5 alasan Show Up Not Paid (kasus)", titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-sunp" });
    const x = 8.5, w = 4.23;
    s.addText("ALASAN #1 PER CENTER", { x, y: 1.75, w, h: 0.3, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1, margin: 0, isTextBox: true });
    CENTERS.forEach((c, i) => {
      const y = 2.15 + i * 0.62;
      card(s, x, y, w, 0.52);
      chip(s, c, x + 0.15, y + 0.11, 0.75, D[c].isNew ? "dark" : "teal");
      s.addText(D[c].top, { x: x + 1.05, y, w: w - 1.95, h: 0.52, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(id(D[c].sunp), { x: x + w - 0.85, y, w: 0.7, h: 0.52, fontSize: 14, bold: true, color: C.accent6, align: "right", valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("Angka kanan = total SUNP center itu", { x, y: 5.9, w, h: 0.3, fontSize: 11, color: C.accent6, margin: 0, isTextBox: true });
    source(s, "Terbanyak per alasan: keluarga → HIB (54) · harga → BTU (28) · jadwal → BTU (50) · coba saja → KLM (17) · jauh → KLM (18). 5 alasan = 66% dari 708 kasus. Sumber: Exboard › 02 Acquisition.");
    s.addNotes("SUNP = Show Up Not Paid. Lima alasan ini 66% dari 708 kasus. Dua pola penting: (1) 'diskusi dengan keluarga' = pengambil keputusan tidak ikut trial, paling banyak di HIB; (2) 'jadwal tidak cocok' 50 dari 95 kasus ada di BTU, padahal utilisasi BTU baru 10% — jamnya yang belum pas. Playbook jawabannya ada di Strategi 2.");
  }

  // 7 · Highlights -------------------------------------------------------
  {
    const s = newSlide(K1, "Yang berjalan baik di September", S1);
    const items = [
      ["FaTrophy", "Center of the Month", "KLM", "109% target · Rp 651 jt · FP progress 89%"],
      ["FaMedal", "SA of the Month", "Agung · HIB", "30 FP dan revenue Rp 145 jt, juara di 3 kategori"],
      ["FaBullseye", "CVR tertinggi", "TMP 54,9%", "▲19,3 pp vs Agustus · satu-satunya center zona hijau"],
      ["FaChartLine", "Satu-satunya yang tumbuh", "HIB ▲24%", "Revenue Rp 325 jt → Rp 403 jt"],
      ["FaUsers", "Full payment terbanyak", "BTU 85 FP", "FP progress ▲20,2 pp ke 79,3%"],
      ["FaStar", "Most improved SA", "Ety · TMP", "CVR ▲25,9 pp · disusul Vina (HIB) dan Sari (TMP)"],
    ];
    const gap = 0.3, w = (CW - 2 * gap) / 3, h = 2.25;
    items.forEach(([ic, lbl, big, sub], i) => {
      const x = X0 + (i % 3) * (w + gap), y = 1.7 + Math.floor(i / 3) * (h + 0.3);
      card(s, x, y, w, h);
      iconDot(s, ic, x + 0.25, y + 0.25, 0.6, C.accent5);
      s.addText(lbl.toUpperCase(), { x: x + 1.0, y: y + 0.25, w: w - 1.2, h: 0.6, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(big, { x: x + 0.25, y: y + 0.95, w: w - 0.5, h: 0.6, fontSize: 28, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(sub, { x: x + 0.25, y: y + 1.55, w: w - 0.5, h: 0.6, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    });
    source(s, "Sumber: deck SA League Monthly September 2026 (Exboard, per 4 Okt 2026).");
    s.addNotes("Sebelum masuk ke yang perlu diperbaiki, apresiasi dulu. KLM Center of the Month. Agung (HIB) SA of the Month. TMP punya CVR tertinggi, HIB satu-satunya yang revenue-nya naik, BTU paling banyak full payment. Praktik baik ini yang kita replikasi ke center lain.");
  }

  // ===== SECTION 2 =====
  const S2 = "02 Center yang perlu improvement";
  pres.addSection({ title: S2 });
  const K2 = "02 · CENTER YANG PERLU IMPROVEMENT";

  // 8 · Scorecard heatmap -------------------------------------------------
  {
    const s = newSlide(K2, "Scorecard: di mana tiap center tertinggal", S2);
    const head = ["Center", "Revenue\n(Rp jt)", "% target", "Δ revenue\nvs Agt", "Δ show up\nvs Agt", "CVR\nSU → paid", "FP\nprogress", "DP belum\nlunas", "SUNP", "Utilisasi\nkelas"];
    const rows = [head.map((h) => hdr(h, { fontSize: 12 }))];
    for (const c of CENTERS) {
      const d = D[c], p = pctTarget(c), rm = chg(d.rev, d.revAug), sm = chg(d.su, d.suAug);
      rows.push([
        cell(`${c}${d.isNew ? " *" : ""}`, { bold: true, fontSize: 15 }),
        cell(id(d.rev), { align: "center" }),
        stCell(`${p}%`, stTarget(p)),
        stCell(arrowPct(rm), stMoM(rm)),
        stCell(arrowPct(sm), sm > -25 ? "warn" : "bad"),
        stCell(pct(d.cvr), stCvr(d.cvr)),
        stCell(pct(d.fpp), stFpp(d.fpp)),
        cell(id(d.dp), { align: "center", bold: d.dp >= 25 }),
        cell(id(d.sunp), { align: "center" }),
        stCell(pct(d.util), stUtil(d.util)),
      ]);
    }
    rows.push([
      cell("Area", { bold: true, fill: { color: TINT.teal }, fontSize: 15 }), cell(id(AREA.rev), { align: "center", bold: true, fill: { color: TINT.teal } }),
      stCell("68%", "bad"), stCell(arrowPct(chg(AREA.rev, AREA.revAug)), "warn"), stCell(arrowPct(chg(AREA.su, AREA.suAug)), "bad"),
      stCell(pct(AREA.cvr), "warn"), stCell(pct(AREA.fpp), "bad"), cell(id(AREA.dp), { align: "center", bold: true, fill: { color: TINT.teal } }),
      cell(id(AREA.sunp), { align: "center", bold: true, fill: { color: TINT.teal } }), stCell(pct(AREA.util), "bad"),
    ]);
    s.addTable(rows, { x: X0, y: 1.7, w: CW, colW: [1.13, 1.2, 1.15, 1.25, 1.25, 1.25, 1.25, 1.2, 1.1, 1.353], rowH: [0.62, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5], fontSize: 14, color: C.text1, border: BORDER, objectName: "table-scorecard" });
    const ly = 6.0;
    [["good", "Sehat"], ["warn", "Waspada"], ["bad", "Perlu tindakan"]].forEach(([st, t], i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X0 + i * 1.95, y: ly + 0.05, w: 0.3, h: 0.25, fill: { color: TINT[st] }, line: { color: TINT[`${st}Ink`], width: 0.75 }, rectRadius: 0.05, objectName: oname("legend") });
      s.addText(t, { x: X0 + 0.4 + i * 1.95, y: ly, w: 1.5, h: 0.35, fontSize: 12, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("* center baru (grand opening 11–12 Agt) — utilisasi rendah masih wajar di bulan ke-2", { x: X0 + 5.9, y: ly, w: CW - 5.9, h: 0.35, fontSize: 12, color: C.accent6, valign: "middle", align: "right", margin: 0, isTextBox: true });
    source(s, "Ambang: % target ≥100 hijau, 75–99 kuning · CVR ≥45% hijau, 35–45% kuning · FP progress ≥85% hijau, 75–85% kuning · utilisasi ≥50% hijau, 30–50% kuning. Sumber: Exboard 02 & 08.");
    s.addNotes("Baca per baris. TMP merah di hampir semua kolom kecuali CVR. PML merah di revenue dan volume. KWC tertahan di FP progress 61%. BTU dan HIB kuning: BTU bocor di jadwal, HIB bocor di CVR. KLM hijau di revenue dan FP progress tapi volumenya juga turun 36%.");
  }

  // 9 · Prioritas ----------------------------------------------------------
  {
    const s = newSlide(K2, "Prioritas improvement Oktober", S2);
    const cols = [
      { t: "Prioritas 1 · intervensi penuh", st: "bad", items: [
        ["TMP", "Revenue terendah. FP progress 43,5%, 35 DP belum lunas", "Leads −52% setelah grand opening"],
        ["PML", "Leads −47%, show up −43%", "SUNP #1 “coba saja” → kualitas lead"],
      ] },
      { t: "Prioritas 2 · tutup satu kebocoran", st: "warn", items: [
        ["KWC", "FP progress 61,4%, 27 DP belum lunas", "Bocor di pelunasan"],
        ["BTU", "50 kasus “jadwal tidak cocok”", "Bocor di jadwal & harga"],
        ["HIB", "CVR 36,8% terendah, SUNP 156", "Bocor di “diskusi keluarga”"],
      ] },
      { t: "Jaga momentum", st: "good", items: [
        ["KLM", "Show up −36%, CVR 39,6% < 45%", "Jaga di atas Rp 600 jt"],
      ] },
    ];
    const gap = 0.3, w = (CW - 2 * gap) / 3, top = 1.7, bottom = 6.6;
    cols.forEach((col, i) => {
      const x = X0 + i * (w + gap);
      chip(s, col.t, x, top, w, col.st);
      const n = col.items.length, cg = 0.15, ch = n === 1 ? 1.75 : (bottom - (top + 0.5) - (n - 1) * cg) / n;
      col.items.forEach(([c, issue, cause], j) => {
        const y = top + 0.5 + j * (ch + cg);
        card(s, x, y, w, ch);
        s.addText(c, { x: x + 0.25, y: y + 0.15, w: 1.2, h: 0.5, fontSize: 24, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
        chip(s, `${pctTarget(c)}% target`, x + w - 1.55, y + 0.25, 1.3, stTarget(pctTarget(c)));
        s.addText(issue, { x: x + 0.25, y: y + 0.62, w: w - 0.5, h: ch - 1.04, fontSize: 14, bold: true, color: C.text1, valign: "top", margin: 0, isTextBox: true });
        s.addText(cause, { x: x + 0.25, y: y + ch - 0.42, w: w - 0.5, h: 0.32, fontSize: 14, color: C.accent6, valign: "top", margin: 0, isTextBox: true });
      });
    });
    {
      const x = X0 + 2 * (w + gap), y = 1.7 + 0.5 + 1.75 + 0.15;
      card(s, x, y, w, 6.6 - y, TINT.teal);
      s.addText("CARA MENENTUKAN PRIORITAS", { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 1, margin: 0, isTextBox: true });
      s.addText([
        { text: "Prioritas 1: gap > Rp 250 jt dan bocor di lebih dari satu titik funnel.", options: { bullet: true, breakLine: true } },
        { text: "Prioritas 2: satu kebocoran utama yang jelas.", options: { bullet: true, breakLine: true } },
        { text: "Gap: TMP 379 · PML 289 · BTU 207 · HIB 197 · KWC 127 (Rp jt).", options: { bullet: true } },
      ], { x: x + 0.25, y: y + 0.6, w: w - 0.5, h: 6.6 - y - 0.8, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true });
    }
    s.addNotes("Prioritas 1: TMP dan PML, gap terbesar (Rp 379 jt dan Rp 289 jt) dan masalahnya di lebih dari satu titik funnel. Prioritas 2: KWC, BTU, HIB, masing-masing punya satu kebocoran utama yang jelas. KLM tetap dijaga karena volumenya juga turun.");
  }

  // ===== SECTION 3 =====
  const S3 = "03 Apa yang akan kita lakukan";
  pres.addSection({ title: S3 });
  const K3 = "03 · APA YANG AKAN KITA LAKUKAN · DEEP DIVE STRATEGI";

  // 10 · Lima tuas -----------------------------------------------------------
  {
    const s = newSlide(K3, "Lima tuas revenue harus bergerak bersama", S3);
    const lev = [
      ["FaBullhorn", "Leads", id(AREA.leads), id(TOT.leads), "Strategi 1"],
      ["FaUserCheck", "Show-up rate", pct(AREA.sur), "50%", "Strategi 1"],
      ["FaHandshake", "CVR SU → paid", pct(AREA.cvr), "45%", "Strategi 2"],
      ["FaWallet", "FP progress", pct(AREA.fpp), `≥ ${FPP_T}%`, "Strategi 3"],
      ["FaMoneyBillWave", "Revenue per FP", `Rp ${dec(AREA.rev / AREA.fp)} jt`, "Naik", "Strategi 3"],
    ];
    const gap = 0.32, w = (CW - 4 * gap) / 5, y = 1.75, h = 3.2;
    lev.forEach(([ic, l, now, goal, st], i) => {
      const x = X0 + i * (w + gap);
      card(s, x, y, w, h);
      iconDot(s, ic, x + 0.25, y + 0.25, 0.6, C.accent1);
      s.addText(l, { x: x + 0.25, y: y + 0.95, w: w - 0.4, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      s.addText("SEP", { x: x + 0.25, y: y + 1.4, w: w - 0.4, h: 0.25, fontSize: 11, bold: true, color: C.accent6, margin: 0, isTextBox: true });
      s.addText(now, { x: x + 0.25, y: y + 1.62, w: w - 0.4, h: 0.5, fontSize: 24, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      s.addText("TARGET OKT", { x: x + 0.25, y: y + 2.2, w: w - 0.4, h: 0.25, fontSize: 11, bold: true, color: C.accent1, margin: 0, isTextBox: true });
      s.addText(goal, { x: x + 0.25, y: y + 2.42, w: w - 0.4, h: 0.5, fontSize: 24, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
      if (i < lev.length - 1) s.addText("×", { x: x + w, y: y + 1.3, w: gap, h: 0.5, fontSize: 20, bold: true, color: C.accent6, align: "center", valign: "middle", margin: 0, isTextBox: true });
      chip(s, st, x + 0.25, y + h - 0.02 + 0.12, w - 0.5, "dark");
    });
    card(s, X0, 5.65, CW, 0.85, C.background2);
    s.addText([
      { text: "Revenue Rp 2,45 M → target Rp 3,6 M (+47%). ", options: { bold: true } },
      { text: "Tidak cukup menggerakkan satu tuas: volume harus kembali mendekati Agustus, sambil CVR, pelunasan, dan ukuran paket ikut naik. Strategi 4 (kapasitas) memastikan jadwal yang dijual memang tersedia." },
    ], { x: X0 + 0.25, y: 5.65, w: CW - 0.5, h: 0.85, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    source(s, `Target Okt = gabungan komitmen 6 center (slide 16–21). Revenue per FP = revenue ÷ jumlah FP September (proxy ukuran paket). FP progress target = rata-rata target center, dibobot total paid.`);
    s.addNotes(`Kerangka strategi: revenue digerakkan oleh lima tuas berurutan. September kita kuat di CVR tapi lemah di volume dan pelunasan. Target Oktober Rp 3,6 M butuh semua tuas naik: leads ke ±${id(TOT.leads)}, show-up rate 50%, CVR 45%, FP progress minimal ${FPP_T}%, dan paket yang lebih panjang supaya revenue per FP naik.`);
  }

  // 11 · Strategi 1 Volume -------------------------------------------------------
  {
    const s = newSlide(K3, "Strategi 1 · Kembalikan volume funnel", S3);
    const rows = [[hdr("Center"), hdr("Leads Agt"), hdr("Leads Sep"), hdr("Δ"), hdr("Target Okt")]];
    for (const c of CENTERS) {
      const m = chg(D[c].leads, D[c].leadsAug);
      rows.push([cell(c, { bold: true }), cell(id(D[c].leadsAug), { align: "center" }), cell(id(D[c].leads), { align: "center" }), stCell(arrowPct(m), m > -30 ? "warn" : "bad"), cell(id(PLAN[c].leadsCommit), { align: "center", bold: true, color: HEX.accent1 })]);
    }
    rows.push([cell("Area", { bold: true, fill: { color: TINT.teal } }), cell(id(AREA.leadsAug), { align: "center", fill: { color: TINT.teal } }), cell(id(AREA.leads), { align: "center", fill: { color: TINT.teal } }), stCell(arrowPct(chg(AREA.leads, AREA.leadsAug)), "bad"), cell(id(TOT.leads), { align: "center", bold: true, color: HEX.accent1, fill: { color: TINT.teal } })]);
    s.addTable(rows, { x: X0, y: 1.75, w: 5.3, colW: [0.95, 1.1, 1.1, 1.0, 1.15], rowH: 0.5, fontSize: 14, color: C.text1, border: BORDER, objectName: "table-leads" });
    const acts = [
      ["FaSchool", "Aktivasi lokal 2× per minggu per center", "Sekolah/TK, komunitas perumahan, event weekend. Prioritas TMP, PML, HIB yang leads-nya turun ≥35%."],
      ["FaUserFriends", "Referral di setiap closing", "Diskon uang pangkal Rp 250 rb untuk 3 kontak referral (sudah ada di price list) — wajib ditawarkan."],
      ["FaCalendarCheck", "Jaga show-up rate ≥ 50%", "Konfirmasi H-1 dan pagi H-0; yang batal langsung di-reschedule di hari yang sama."],
      ["FaRedo", "Reaktivasi 708 SUNP September", "Undang trial ulang saat ada slot/jadwal baru yang cocok dengan alasan mereka."],
    ];
    const x = 6.25, w = W - X0 - x;
    acts.forEach(([ic, t, d], i) => {
      const y = 1.75 + i * 1.05;
      iconDot(s, ic, x, y + 0.05, 0.6, C.accent1);
      s.addText(t, { x: x + 0.8, y, w: w - 0.8, h: 0.35, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      s.addText(d, { x: x + 0.8, y: y + 0.38, w: w - 0.8, h: 0.65, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    card(s, x, 6.02, w, 0.5, C.background2);
    s.addText([{ text: "Ask ke Marketing: ", options: { bold: true } }, { text: "geser alokasi ads ke TMP (−52%) dan PML (−47%)." }], { x: x + 0.2, y: 6.02, w: w - 0.4, h: 0.5, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    source(s, "Target leads Okt = show up yang dibutuhkan ÷ target show-up rate (lihat slide 15). Leads Agt: Center Performance Agustus 2026; Sep: Exboard › 02 Acquisition.");
    s.addNotes(`Volume adalah tuas terbesar. Total leads perlu naik dari 2.016 ke sekitar ${id(TOT.leads)}, masih di bawah Agustus (3.097). Empat aksi: aktivasi lokal, referral di setiap closing, konfirmasi H-1/H-0 untuk jaga show-up rate, dan reaktivasi pool SUNP. Plus satu permintaan ke Marketing untuk geser budget ads ke TMP dan PML.`);
  }

  // 12 · Strategi 2 Closing playbook ------------------------------------------------
  {
    const s = newSlide(K3, "Strategi 2 · Playbook 5 alasan belum closing", S3);
    const play = [
      ["Diskusi dengan keluarga", "131", "HIB · KLM · KWC", "Pastikan pengambil keputusan ikut trial (konfirmasi saat booking). Kalau tidak hadir: video call di tempat. Follow-up H+1 dengan video progres anak."],
      ["Harga / kemahalan", "100", "BTU", "Value dulu (air hangat, maks 1:4, lifeguard, make-up class), baru harga. Tawarkan cicilan Indodana — bayar mulai bulan depan. Jangan turun di bawah bottom."],
      ["Jadwal tidak cocok", "95", "BTU (50)", "Catat jam yang diminta → bahas di slot map meeting Senin → buka/alih slot → hubungi ulang parents."],
      ["Tidak tertarik / coba saja", "73", "PML · TMP", "Kualifikasi saat booking: tujuan, usia, jadwal, budget. Trial diarahkan ke paket yang cocok, bukan sekadar coba."],
      ["Rumah jauh / out of area", "69", "KLM (18)", "Cek jarak saat booking. Tawarkan jadwal weekend atau center Sparks yang lebih dekat."],
    ];
    const rows = [[hdr("Alasan SUNP", { align: "left" }), hdr("Kasus"), hdr("Fokus center"), hdr("Jawaban & aksi SA", { align: "left" })]];
    play.forEach(([r, n, c, a]) => rows.push([cell(r, { bold: true }), cell(n, { align: "center", bold: true, color: HEX.accent1 }), cell(c, { align: "center" }), cell(a)]));
    s.addTable(rows, { x: X0, y: 1.7, w: CW, colW: [2.6, 0.9, 1.75, 6.883], rowH: [0.42, 0.78, 0.78, 0.6, 0.6, 0.6], fontSize: 14, color: C.text1, border: { type: "solid", pt: 1, color: TINT.line }, fill: { color: "FFFFFF" }, objectName: "table-playbook" });
    card(s, X0, 5.95, CW, 0.6, C.background2);
    s.addText([{ text: "Target CVR area 45% (Sep 41,4%). ", options: { bold: true } }, { text: "Alat: Pitching Script v2.0, Pricing War Battlecard, SUNP & DP Tracker. CM role-play 15 menit tiap pagi untuk SA di P3 (8 SA)." }], { x: X0 + 0.25, y: 5.95, w: CW - 0.5, h: 0.6, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    source(s, "Jumlah kasus: Exboard › 02 Acquisition (blok SUNP per center), September 2026. P-band CVR September: P1 13 SA, P2 5 SA, P3 8 SA.");
    s.addNotes("Setiap alasan SUNP punya jawaban standar. Yang paling besar dampaknya: 'diskusi keluarga' (131 kasus), bisa dicegah sejak booking dengan memastikan pengambil keputusan hadir. Untuk harga, jual value dulu baru angka; cicilan Indodana menjawab keberatan cashflow tanpa menurunkan harga. Untuk jadwal, solusinya di slot map, bukan di pitching.");
  }

  // 13 · Strategi 3 Cash & ticket -------------------------------------------------
  {
    const s = newSlide(K3, "Strategi 3 · DP jadi FP, ticket size naik", S3);
    // kiri: FP progress
    const xl = X0, wl = 5.6;
    card(s, xl, 1.7, wl, 4.85);
    s.addText("PELUNASAN", { x: xl + 0.3, y: 1.9, w: wl - 0.6, h: 0.3, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1, margin: 0, isTextBox: true });
    s.addText([{ text: `${AREA.dp} murid `, options: { bold: true, fontSize: 32 } }, { text: "masih DP", options: { fontSize: 18 } }], { x: xl + 0.3, y: 2.2, w: wl - 0.6, h: 0.65, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
    s.addText(`TMP ${D.TMP.dp} · KWC ${D.KWC.dp} · BTU ${D.BTU.dp} · HIB ${D.HIB.dp} · PML ${D.PML.dp} · KLM ${D.KLM.dp}`, { x: xl + 0.3, y: 2.88, w: wl - 0.6, h: 0.35, fontSize: 14, color: C.accent6, margin: 0, isTextBox: true });
    const la = [
      ["FaTasks", "Antrian DP harian di SUNP & DP Tracker", "CM cek 3 angka tiap pagi: DP overdue, SUNP overdue, belum pernah di-FU."],
      ["FaWallet", "Closing langsung FP dengan Indodana", "Sparks terima penuh, parents cicil — bayar mulai bulan depan."],
      ["FaClock", "Target: 60% DP lunas s.d. 20 Okt", `≈ ${Math.round(AREA.dp * ASSUME.dpToFp)} FP tambahan dari pool September.`],
    ];
    la.forEach(([ic, t, d], i) => {
      const y = 3.4 + i * 1.02;
      iconDot(s, ic, xl + 0.3, y + 0.03, 0.55, C.accent1);
      s.addText(t, { x: xl + 1.05, y, w: wl - 1.3, h: 0.35, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      s.addText(d, { x: xl + 1.05, y: y + 0.36, w: wl - 1.3, h: 0.6, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    // kanan: revenue per FP chart + priority package
    const xr = X0 + wl + 0.35, wr = W - X0 - xr;
    s.addChart(pres.charts.BAR, [{ name: "Revenue per FP", labels: CENTERS, values: CENTERS.map((c) => Math.round((D[c].rev / D[c].fp) * 10) / 10) }], {
      x: xr, y: 1.65, w: wr, h: 2.75, chartColors: CENTERS.map((c) => (D[c].isNew ? HEX.accent4 : HEX.accent1)), ...chartText(), showLegend: false,
      valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 11, showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 12,
      showTitle: true, title: "Revenue per FP September (Rp jt) · center baru oranye", titleFontSize: 13, titleColor: HEX.dk1, barGapWidthPct: 50, objectName: "chart-ticket" });
    const pk = [["Advanced 12 bulan", "Rp 9,35 jt", "6+6 bulan · cicilan Rp 896 rb/bln"], ["Basic 6 bulan", "Rp 5,32 jt", "3+3 bulan"], ["Lite 3 bulan", "Rp 4,5 jt", "starter, tanpa bonus"]];
    s.addText("PROMO OKTOBER · PITCH DARI PAKET TERPANJANG", { x: xr, y: 4.5, w: wr, h: 0.3, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1, margin: 0, isTextBox: true });
    pk.forEach(([n, p, d], i) => {
      const y = 4.85 + i * 0.57;
      card(s, xr, y, wr, 0.5, i === 0 ? TINT.teal : C.background2);
      s.addText(`${i + 1}`, { x: xr + 0.15, y, w: 0.3, h: 0.5, fontSize: 16, bold: true, color: C.accent1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(n, { x: xr + 0.5, y, w: 2.0, h: 0.5, fontSize: 14, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(p, { x: xr + 2.45, y, w: 1.1, h: 0.5, fontSize: 14, bold: true, color: C.accent1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: xr + 3.55, y, w: wr - 3.7, h: 0.5, fontSize: 12, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    });
    source(s, "Revenue per FP = revenue ÷ FP September (proxy ukuran paket). Harga: deck Priority Package & New Pitching Style (promo FP hari ini, belum termasuk uang pangkal & weekend).");
    s.addNotes("Dua hal: pelunasan dan ukuran paket. 115 murid closing dengan DP; kalau 60% lunas sebelum 20 Okt itu sekitar 69 FP tambahan. Cicilan Indodana membuat closing bisa langsung FP. Lalu ukuran paket: revenue per FP center baru Rp 4,6–5,1 jt, center matang Rp 7,4–9,0 jt. Pitch selalu mulai dari Advanced 12 bulan (Rp 9,35 jt, 6+6 gratis).");
  }

  // 14 · Strategi 4 Kapasitas ---------------------------------------------------
  {
    const s = newSlide(K3, "Strategi 4 · Isi kursi kosong yang sudah ada", S3);
    s.addChart(pres.charts.BAR, [{ name: "Utilisasi", labels: CENTERS, values: CENTERS.map((c) => D[c].util / 100) }], {
      x: X0, y: 1.65, w: 6.0, h: 3.6, chartColors: CENTERS.map((c) => (D[c].isNew ? HEX.accent4 : HEX.accent1)), ...chartText(), showLegend: false,
      valAxisMinVal: 0, valAxisMaxVal: 0.7, valAxisMajorUnit: 0.1, valAxisLabelFormatCode: "0%", showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%", dataLabelFontSize: 12,
      showTitle: true, title: "Utilisasi kelas reguler (kursi terisi ÷ kapasitas)", titleFontSize: 13, titleColor: HEX.dk1, barGapWidthPct: 50, objectName: "chart-util" });
    const st = [["Kursi kosong", id(AREA.seats)], ["Kelas dibuka, 0 murid", id(CENTERS.reduce((a, c) => a + D[c].zero, 0))], ["Utilisasi area", pct(AREA.util)]];
    st.forEach(([l, v], i) => {
      const w = (6.0 - 0.4) / 3, x = X0 + i * (w + 0.2);
      card(s, x, 5.4, w, 1.15);
      s.addText(v, { x: x + 0.15, y: 5.48, w: w - 0.3, h: 0.55, fontSize: 26, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      s.addText(l, { x: x + 0.15, y: 6.03, w: w - 0.3, h: 0.4, fontSize: 13, color: C.accent6, margin: 0, valign: "top", isTextBox: true });
    });
    const acts = [
      ["FaCalendarAlt", "Slot map meeting tiap Senin (CM × Senior Coach)", "Pakai working sheet mingguan: putuskan kelas 0 murid mana yang dijual minggu ini, coach PIC, dan target murid."],
      ["FaClipboardCheck", "Daftar jual mingguan ke SA", "Kelas “Jual minggu ini” / “Buka kelas baru” jadi bahan pitching dan reaktivasi SUNP."],
      ["FaClock", "Jam yang diminta parents dulu", "Alasan “jadwal tidak cocok” (BTU 50 kasus) jadi input utama slot yang dibuka atau dialihkan."],
      ["FaSwimmer", "Center baru: jual slot jam favorit", "TMP, BTU, HIB masih ±10% — fokus isi jam sore & weekend sebelum membuka slot baru."],
    ];
    const x = 6.95, w = W - X0 - x;
    acts.forEach(([ic, t, d], i) => {
      const y = 1.75 + i * 1.2;
      iconDot(s, ic, x, y + 0.05, 0.6, C.accent1);
      s.addText(t, { x: x + 0.8, y, w: w - 0.8, h: 0.38, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      s.addText(d, { x: x + 0.8, y: y + 0.4, w: w - 0.8, h: 0.72, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    source(s, "Sumber: Exboard › 08 Class Utilization (per 4 Okt 2026); kelas 0 murid: Slot Map Weekly Meeting W40 (data 26 Sep 2026). Center baru oranye.");
    s.addNotes("Kapasitas bukan hambatan: ada 7.426 kursi kosong dan 1.105 kelas yang sudah punya age group dan coach tapi 0 murid. Masalahnya kecocokan jam. Slot map meeting tiap Senin memutuskan kelas mana yang dijual minggu itu, dan jam yang diminta parents di SUNP 'jadwal' jadi input utamanya.");
  }

  // ===== SECTION 4 =====
  const S4 = "04 Komitmen per center";
  pres.addSection({ title: S4 });
  const K4 = "04 · KOMITMEN PER CENTER";

  // 15 · Jalan ke 600 jt ------------------------------------------------------------
  {
    const s = newSlide(K4, "Jalan ke Rp 600 jt per center", S4);
    const head = ["Center", "Revenue\nSep", "Revenue\nper FP", "FP\nSep", "FP butuh\nRp 600 jt", "dari DP\n(60%)", "dari SUNP\n(10%)", "dari trial\nbaru", "Show up\nbutuh", "Show up\nSep"];
    const rows = [head.map((h) => hdr(h, { fontSize: 12 }))];
    for (const c of CENTERS) {
      const p = PLAN[c], d = D[c];
      const gapSu = p.suNeed > d.su;
      rows.push([
        cell(c, { bold: true, fontSize: 15 }), cell(id(d.rev), { align: "center" }), cell(dec(p.ticket), { align: "center" }), cell(id(d.fp), { align: "center" }),
        cell(id(p.fpNeed), { align: "center", bold: true, color: HEX.accent1 }), cell(id(p.fromDP), { align: "center" }), cell(id(p.fromSUNP), { align: "center" }), cell(id(p.fromTrial), { align: "center" }),
        stCell(id(p.suNeed), gapSu ? "bad" : "good"), cell(id(d.su), { align: "center" }),
      ]);
    }
    s.addTable(rows, { x: X0, y: 1.7, w: CW, colW: [1.13, 1.15, 1.15, 1.0, 1.4, 1.2, 1.25, 1.25, 1.3, 1.303], rowH: [0.62, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48], fontSize: 14, color: C.text1, border: { type: "solid", pt: 1, color: TINT.line }, fill: { color: "FFFFFF" }, objectName: "table-bridge" });
    const xb = X0, yb = 5.35;
    card(s, xb, yb, 7.3, 1.2, C.background2);
    s.addText([
      { text: "Center baru butuh show up kembali ke level Agustus. ", options: { bold: true } },
      { text: `TMP ${id(PLAN.TMP.suNeed)} (Agt ${D.TMP.suAug}), BTU ${id(PLAN.BTU.suNeed)} (Agt ${D.BTU.suAug}), HIB ${id(PLAN.HIB.suNeed)} (Agt ${D.HIB.suAug}). KLM dan KWC sudah cukup dengan volume September.` },
    ], { x: xb + 0.25, y: yb, w: 6.8, h: 1.2, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    const xs = xb + 7.6, ws = W - X0 - xs;
    card(s, xs, yb, ws, 1.2, TINT.teal);
    s.addText([
      { text: "Kalau revenue per FP naik 20% ", options: { bold: true } },
      { text: `(paket 6 & 12 bulan), FP butuh turun: TMP ${PLAN.TMP.fpNeed}→${PLAN.TMP.fpNeedUp20}, BTU ${PLAN.BTU.fpNeed}→${PLAN.BTU.fpNeedUp20}, HIB ${PLAN.HIB.fpNeed}→${PLAN.HIB.fpNeedUp20}.` },
    ], { x: xs + 0.25, y: yb, w: ws - 0.5, h: 1.2, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    source(s, "FP butuh = Rp 600 jt ÷ revenue per FP Sep · trial baru = FP butuh − 60% DP − 10% SUNP · show up butuh = trial baru ÷ (target CVR × target FP progress). Asumsi bisa direvisi.");
    s.addNotes("Ini dasar draft komitmen. Langkahnya: berapa FP yang dibutuhkan untuk Rp 600 jt dengan ukuran paket September; kurangi yang bisa didapat dari pool DP (asumsi 60% lunas) dan pool SUNP (asumsi 10% closing); sisanya harus dari trial baru, dibagi target CVR dan FP progress. Hasilnya: KLM dan KWC aman dengan volume September, PML dan 3 center baru harus menaikkan show up. Kalau ukuran paket naik 20%, kebutuhan FP center baru turun sekitar 17%.");
  }

  // 16–21 · Per center ------------------------------------------------------------
  const CENTER_PAGES = {
    KLM: {
      title: "KLM · Jaga di atas target, pulihkan volume",
      diag: "Satu-satunya center ≥ target (109%), tapi show up turun 36% dan CVR 39,6% masih di bawah 45%. 142 SUNP — alasan #1 diskusi keluarga; “rumah jauh” terbanyak di KLM (18).",
      acts: [
        ["Reaktivasi 142 SUNP + referral di setiap closing (diskon uang pangkal untuk 3 kontak)", "CM + SA · mulai 6 Okt"],
        ["Protokol pengambil keputusan hadir saat trial; follow-up H+1 dengan video progres", "SA · setiap trial"],
        ["108 murid di renewal window: tawarkan paket 12 bulan (6+6) sebelum sesi habis", "SA retention · s.d. 20 Okt"],
      ],
    },
    KWC: {
      title: "KWC · Sudah closing, tinggal dilunasi",
      diag: "Revenue 79% (▼23% vs Agt). FP progress 61,4% — terendah di center matang — dengan 27 DP belum lunas. CVR turun 4,5 pp ke 39,3%.",
      acts: [
        ["Antrian DP harian di SUNP & DP Tracker: 16 dari 27 DP lunas", "CM cek harian · s.d. 20 Okt"],
        ["Tawarkan cicilan Indodana sebelum opsi DP, supaya closing langsung FP", "SA · mulai 6 Okt"],
        ["Survei harga Studiorenang Gading Serpong (2,8 km, belum ada data) + isi Intel Lapangan", "CM · s.d. 11 Okt"],
      ],
    },
    PML: {
      title: "PML · Isi ulang funnel yang turun 47%",
      diag: "Revenue 52% (gap Rp 289 jt). Leads −47% dan show up −43% vs Agustus. SUNP #1 “tidak tertarik / coba saja” → kualitas lead. Harga sudah setara Studiorenang BSD, jadi tidak perlu diskon.",
      acts: [
        ["Aktivasi sekolah/TK & komunitas sekitar Pamulang 2× per minggu; minta tambahan alokasi ads", "CM + Marketing · mulai W41"],
        ["Kualifikasi saat booking (tujuan, usia, jadwal, budget) untuk kurangi trial “coba saja”", "SA / telesales · mulai 6 Okt"],
        ["Jual 120 kelas 0 murid lewat slot map Senin; pitching value (make-up class, lifeguard)", "CM × Senior Coach · tiap Senin"],
      ],
    },
    TMP: {
      title: "TMP · Prioritas #1: dari DP ke FP",
      diag: "Revenue Rp 221 jt (37%), terendah. CVR tertinggi (54,9%) tapi FP progress hanya 43,5% — 35 DP belum lunas. Leads −52% setelah grand opening; revenue per FP Rp 5,1 jt.",
      acts: [
        ["Lunasi 35 DP lewat cash bertahap / Indodana: target 21 FP", "CM cek harian · s.d. 20 Okt"],
        ["Closing FP, bukan DP: promo FP hari ini + pitch Advanced 12 bulan dulu", "SA · mulai 6 Okt"],
        ["Leads ke 460: aktivasi perumahan & sekolah Taman Palem, 2 event weekend; battlecard vs Studiorenang Taman Surya (2,4 km)", "CM + Marketing · W41–W42"],
      ],
    },
    BTU: {
      title: "BTU · Buka jadwal yang dicari parents",
      diag: "Revenue 66%, stagnan vs Agustus. 50 dari 95 kasus “jadwal tidak cocok” se-area ada di BTU, plus 28 kasus harga. Utilisasi 10,4% — kursinya ada, jamnya belum pas. Revenue per FP Rp 4,6 jt, terendah.",
      acts: [
        ["Rekap jam yang diminta 50 SUNP “jadwal” → buka/alih slot di slot map meeting → hubungi ulang", "CM × Senior Coach · s.d. 12 Okt"],
        ["Keberatan harga: battlecard + Indodana; survei harga Studiorenang Bintaro U-Town (1,5 km)", "CM · s.d. 11 Okt"],
        ["Naikkan ticket: pitch Advanced 12 bulan dan Basic 6 bulan sebelum Lite", "SA · mulai 6 Okt"],
      ],
    },
    HIB: {
      title: "HIB · Ubah “diskusi keluarga” jadi closing",
      diag: "Satu-satunya center yang tumbuh (▲24%) dengan show-up rate terbaik 54,8%. Tapi CVR terendah (36,8%) dan SUNP terbanyak (156) — 54 kasus “diskusi dengan keluarga”.",
      acts: [
        ["Konfirmasi kedua orang tua hadir saat booking; kalau tidak, video call di tempat + follow-up H+1", "SA · mulai 6 Okt"],
        ["Replikasi Agung (SA of the Month): shadowing & role-play pitching untuk SA lain", "CM · W41"],
        ["Reaktivasi 156 SUNP lewat SUNP Tracker: target 16 closing", "SA · harian"],
      ],
    },
  };
  for (const c of CENTERS) {
    const d = D[c], p = PLAN[c], g = GOAL[c], pg = CENTER_PAGES[c];
    const s = newSlide(`${K4} · ${d.name.toUpperCase()}`, pg.title, S4);
    // Snapshot September
    const snap = [
      ["Revenue", `Rp ${id(d.rev)} jt`, `${pctTarget(c)}% target`, stTarget(pctTarget(c))],
      ["Show up", id(d.su), `${arrowPct(chg(d.su, d.suAug))} vs Agt`, chg(d.su, d.suAug) > -25 ? "warn" : "bad"],
      ["CVR SU → paid", pct(d.cvr), arrowPp(d.cvrD), stCvr(d.cvr)],
      ["FP progress", pct(d.fpp), arrowPp(d.fppD), stFpp(d.fpp)],
      ["DP belum lunas", id(d.dp), "kejar → FP", d.dp >= 25 ? "bad" : "warn"],
      ["SUNP", id(d.sunp), "pool follow-up", "teal"],
    ];
    s.addText("SEPTEMBER", { x: X0, y: 1.5, w: 3, h: 0.28, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1, margin: 0, isTextBox: true });
    const tg = 0.2, tw = (CW - 5 * tg) / 6;
    snap.forEach(([l, v, sub, st], i) => {
      const x = X0 + i * (tw + tg), y = 1.82;
      card(s, x, y, tw, 1.0);
      s.addText(l, { x: x + 0.15, y: y + 0.08, w: tw - 0.3, h: 0.26, fontSize: 12, bold: true, color: C.accent6, margin: 0, isTextBox: true });
      s.addText(v, { x: x + 0.15, y: y + 0.32, w: tw - 0.3, h: 0.36, fontSize: 20, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      chip(s, sub, x + 0.15, y + 0.66, tw - 0.3, st);
    });
    // Diagnosis banner
    card(s, X0, 2.97, CW, 0.7, TINT.teal);
    s.addText([{ text: "Diagnosis: ", options: { bold: true, color: HEX.accent1 } }, { text: pg.diag }], { x: X0 + 0.2, y: 2.97, w: CW - 0.4, h: 0.7, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    // Komitmen table
    s.addText("DRAFT KOMITMEN OKTOBER", { x: X0, y: 3.8, w: 5.4, h: 0.28, fontSize: 12, bold: true, color: C.accent1, charSpacing: 1, margin: 0, isTextBox: true });
    const rows = [[hdr("KPI", { align: "left" }), hdr("Sep"), hdr("Target Okt")]];
    const tRow = (k, a, b) => [cell(k, { bold: true }), cell(a, { align: "center" }), cell(b, { align: "center", bold: true, color: HEX.accent1, fill: { color: TINT.teal } })];
    rows.push(tRow("Leads", id(d.leads), id(p.leadsCommit)));
    rows.push(tRow("Show up", id(d.su), id(p.suCommit)));
    rows.push(tRow("CVR SU → paid", pct(d.cvr), `${g.cvr}%`));
    rows.push(tRow("FP progress", pct(d.fpp), `${g.fpp}%`));
    rows.push(tRow("Full payment", id(d.fp), `≥ ${id(p.fpCommit)}`));
    rows.push(tRow("Revenue (Rp jt)", id(d.rev), `≥ ${TARGET}`));
    s.addTable(rows, { x: X0, y: 4.1, w: 5.4, colW: [2.2, 1.5, 1.7], rowH: [0.34, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35], fontSize: 14, color: C.text1, border: { type: "solid", pt: 1, color: TINT.line }, fill: { color: "FFFFFF" }, objectName: `table-commit-${c}` });
    // Aksi
    const xr = X0 + 5.75, wr = W - X0 - xr;
    s.addText("3 AKSI UTAMA", { x: xr, y: 3.8, w: wr, h: 0.28, fontSize: 12, bold: true, color: C.accent1, charSpacing: 1, margin: 0, isTextBox: true });
    pg.acts.forEach(([a, who], i) => {
      const y = 4.1 + i * 0.82;
      s.addText(String(i + 1), { x: xr, y: y + 0.02, w: 0.36, h: 0.36, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, shape: pres.shapes.OVAL, fill: { color: C.accent1 }, line: { type: "none" } });
      s.addText([
        { text: a, options: { fontSize: 14, color: C.text1, breakLine: true } },
        { text: who, options: { fontSize: 12, color: C.accent6, bold: true } },
      ], { x: xr + 0.5, y, w: wr - 0.5, h: 0.76, valign: "top", margin: 0, isTextBox: true });
    });
    source(s, `Target Okt dari model slide 15: show up = max(kebutuhan ${p.suNeed}, Sep ${d.su}); leads = show up ÷ show-up rate ${g.sur}% (min. leads Sep); dibulatkan ke 5. Revenue ≥ Rp 600 jt = target Oktober.`);
    s.addNotes(`${c} (${d.name}). September: revenue Rp ${id(d.rev)} jt (${pctTarget(c)}% target), show up ${d.su}, CVR ${pct(d.cvr)}, FP progress ${pct(d.fpp)}, ${d.dp} DP belum lunas, ${d.sunp} SUNP. Draft komitmen Oktober: leads ${p.leadsCommit}, show up ${p.suCommit}, CVR ${g.cvr}%, FP progress ${g.fpp}%, FP minimal ${p.fpCommit}, revenue minimal Rp 600 jt. Minta CM konfirmasi atau revisi angka dan PIC tiap aksi di meeting.`);
  }

  // 22 · Ringkasan komitmen ---------------------------------------------------------
  {
    const s = newSlide(K4, "Ringkasan komitmen Oktober", S4);
    const rows = [["Center", "Leads", "Show up", "CVR", "FP progress", "Full payment", "Revenue (Rp jt)", "Konfirmasi CM"].map((h) => hdr(h))];
    for (const c of CENTERS) {
      const p = PLAN[c], g = GOAL[c];
      rows.push([cell(c, { bold: true, fontSize: 15 }), cell(id(p.leadsCommit), { align: "center" }), cell(id(p.suCommit), { align: "center" }), cell(`${g.cvr}%`, { align: "center" }), cell(`${g.fpp}%`, { align: "center" }), cell(`≥ ${id(p.fpCommit)}`, { align: "center" }), cell(`≥ ${TARGET}`, { align: "center", bold: true, color: HEX.accent1 }), cell("", { fill: { color: "FFFFFF" } })]);
    }
    rows.push([cell("Area", { bold: true, fill: { color: TINT.teal }, fontSize: 15 }), cell(id(TOT.leads), { align: "center", bold: true, fill: { color: TINT.teal } }), cell(id(TOT.su), { align: "center", bold: true, fill: { color: TINT.teal } }), cell("45%", { align: "center", bold: true, fill: { color: TINT.teal } }), cell(`≥ ${FPP_T}%`, { align: "center", bold: true, fill: { color: TINT.teal } }), cell(`≥ ${id(TOT.fp)}`, { align: "center", bold: true, fill: { color: TINT.teal } }), cell(`≥ ${id(TARGET * 6)}`, { align: "center", bold: true, color: HEX.accent1, fill: { color: TINT.teal } }), cell("", { fill: { color: TINT.teal } })]);
    s.addTable(rows, { x: X0, y: 1.7, w: CW, colW: [1.2, 1.35, 1.35, 1.2, 1.5, 1.6, 1.8, 2.133], rowH: 0.52, fontSize: 14, color: C.text1, border: { type: "solid", pt: 1, color: TINT.line }, fill: { color: "F4F8F9" }, objectName: "table-summary" });
    card(s, X0, 5.95, CW, 0.6, C.background2);
    s.addText([{ text: "Dibanding September: ", options: { bold: true } }, { text: `leads +${Math.round(chg(TOT.leads, AREA.leads))}%, show up +${Math.round(chg(TOT.su, AREA.su))}%, full payment +${Math.round(chg(TOT.fp, AREA.fp))}%, revenue +${Math.round(chg(TARGET * 6, AREA.rev))}%. Kolom konfirmasi diisi CM di meeting.` }], { x: X0 + 0.25, y: 5.95, w: CW - 0.5, h: 0.6, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    source(s, "Area CVR dan FP progress = target minimum area; target per center di kolom masing-masing. Revenue target Oktober Rp 600 jt per center (total Rp 3,6 M).");
    s.addNotes(`Ringkasan seluruh draft komitmen dalam satu tabel. Di meeting, tiap CM menyatakan setuju atau mengusulkan revisi, lalu kolom konfirmasi diisi. Total: leads sekitar ${id(TOT.leads)}, show up ${id(TOT.su)}, FP minimal ${id(TOT.fp)}, revenue Rp 3,6 M.`);
  }

  // 23 · Ritme monitoring --------------------------------------------------------
  {
    const s = newSlide(K4, "Ritme monitoring Oktober", S4);
    const weeks = [
      ["W41", "5–11 Okt", "Kick-off komitmen. Antrian DP & SUNP dibersihkan. Survei harga kompetitor KWC, BTU, HIB. Slot BTU dibuka."],
      ["W42", "12–18 Okt", "Checkpoint 15 Okt: revenue ≥ 50% target (Rp 300 jt per center). Di bawah 40% → action plan bersama AM."],
      ["W43", "19–25 Okt", "Batas 20 Okt: 60% DP September lunas. Evaluasi aktivasi lokal: leads vs target."],
      ["W44", "26–31 Okt", "Closing bulan. Data per 31 Okt jadi bahan review Oktober."],
    ];
    const gap = 0.3, w = (CW - 3 * gap) / 4, y = 1.75, h = 2.55;
    weeks.forEach(([wk, dt, t], i) => {
      const x = X0 + i * (w + gap);
      card(s, x, y, w, h, i === 1 ? TINT.teal : C.background2);
      s.addText(wk, { x: x + 0.25, y: y + 0.2, w: 1.2, h: 0.5, fontSize: 26, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
      s.addText(dt, { x: x + 1.4, y: y + 0.2, w: w - 1.6, h: 0.5, fontSize: 14, bold: true, color: C.accent6, align: "right", margin: 0, valign: "middle", isTextBox: true });
      s.addText(t, { x: x + 0.25, y: y + 0.85, w: w - 0.5, h: h - 1.0, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    const rit = [
      ["FaTasks", "Harian · CM", "Cek DSR dan SUNP & DP Tracker: DP overdue, SUNP overdue, belum pernah di-FU. Naik dari kemarin = ada SA yang perlu dibantu."],
      ["FaUsers", "Senin · CM × Senior Coach", "Slot map meeting: kelas yang dijual minggu ini, coach PIC, target murid → daftar jual ke SA."],
      ["FaChartLine", "Senin · Area", "SA League Weekly: CVR, FP, revenue per SA dan per center vs komitmen."],
    ];
    rit.forEach(([ic, t, d], i) => {
      const yy = 4.6 + i * 0.66;
      iconDot(s, ic, X0, yy + 0.04, 0.5, C.accent1);
      s.addText(t, { x: X0 + 0.7, y: yy, w: 3.0, h: 0.58, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
      s.addText(d, { x: X0 + 3.75, y: yy, w: CW - 3.75, h: 0.58, fontSize: 14, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
    });
    s.addNotes("Komitmen hanya jalan kalau dimonitor. Harian CM cek tiga angka di tracker. Tiap Senin ada slot map meeting dengan Senior Coach dan SA League Weekly. Checkpoint utama 15 Oktober: revenue minimal 50% target; center di bawah 40% menyusun action plan bersama Area Manager. 20 Oktober batas pelunasan 60% DP September.");
  }

  // 24 · Closing -------------------------------------------------------------------
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: S4 });
    s.addText("Oktober: Rp 3,6 M — 6 center × Rp 600 jt", { placeholder: "title" });
    const pts = [
      ["FaBullhorn", "Volume", `Leads ${id(TOT.leads)} · show up ${id(TOT.su)}`],
      ["FaHandshake", "Closing", "CVR 45% dengan playbook 5 alasan SUNP"],
      ["FaWallet", "Cash & ticket", "60% DP lunas s.d. 20 Okt · pitch paket 12 bulan dulu"],
      ["FaCalendarAlt", "Kapasitas", "Slot map tiap Senin · jual jam yang dicari parents"],
    ];
    const gap = 0.3, w = (CW - 3 * gap) / 4;
    pts.forEach(([ic, t, d], i) => {
      const x = X0 + i * (w + gap), y = 3.0;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 2.0, fill: { color: "134C5A" }, line: { type: "none" }, rectRadius: 0.12, objectName: oname("card-dark") });
      iconDot(s, ic, x + 0.25, y + 0.25, 0.6, C.accent2);
      s.addText(t, { x: x + 0.25, y: y + 0.95, w: w - 0.5, h: 0.4, fontSize: 18, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true });
      s.addText(d, { x: x + 0.25, y: y + 1.35, w: w - 0.5, h: 0.6, fontSize: 14, color: C.background2, margin: 0, valign: "top", isTextBox: true });
    });
    s.addText([
      { text: "Langkah berikutnya: ", options: { bold: true, color: C.accent2 } },
      { text: "CM konfirmasi angka & PIC aksi paling lambat H+1 setelah meeting. Checkpoint pertama 15 Oktober.", options: { color: C.background1 } },
    ], { x: X0, y: 5.5, w: CW, h: 0.6, fontSize: 16, valign: "middle", margin: 0, isTextBox: true });
    s.addNotes("Penutup. Satu target: Rp 3,6 M, Rp 600 jt per center. Empat strategi: volume, closing, cash & ticket, kapasitas. Langkah berikutnya: CM konfirmasi komitmen H+1, checkpoint 15 Oktober.");
  }

  await pres.writeFile({ fileName: OUT });
  await writeThemeColors(OUT, THEME);
  console.log("Wrote", OUT);
  console.log(JSON.stringify(PLAN, (k, v) => (typeof v === "number" ? Math.round(v * 100) / 100 : v)));
  console.log("TOT", TOT);
}

// pptxgenjs menulis palet Office bawaan ke theme1.xml; ganti dengan palet THEME supaya warna skema sesuai.
async function writeThemeColors(file, theme) {
  const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const part = "ppt/theme/theme1.xml";
  const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const scheme = `<a:clrScheme name="${theme.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
  const xml = (await zip.file(part).async("string"))
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, h) => `${h}${theme.name}"`);
  zip.file(part, xml);
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith(".xml")) continue;
    const body = await zip.file(name).async("string");
    const bad = body.match(/srgbClr val="(?![0-9A-Fa-f]{6}")[^"]*"/);
    if (bad) throw new Error(`${name}: warna bukan hex (${bad[0]}) — opsi itu hanya menerima hex`);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

build().catch((e) => { console.error(e); process.exit(1); });
