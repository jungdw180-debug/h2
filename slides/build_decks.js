// 1단계 단계평가 발표자료 생성 스크립트 (본평가판 + 현장평가판)
// 실행: NODE_PATH=<pptxgenjs·react-icons·sharp가 설치된 node_modules> APPLY_THEME=<apply_theme.js 경로> node slides/build_decks.js
// 내용 근거: report/발표자료_슬라이드내용_v1.md (쪽수는 단계보고서 인쇄 쪽 번호)

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const { applyTheme } = require(process.env.APPLY_THEME);

const ROOT = path.join(__dirname, "..");
const FIG = (f) => (f.startsWith("fig") ? path.join(ROOT, "figures", f) : path.join(ROOT, "figures", "pdf", f));
const OUT = path.join(ROOT, "slides");

// ---------------------------------------------------------------- 테마
const THEME = {
  name: "H2 Station Stage Review",
  headFontFace: "맑은 고딕",
  bodyFontFace: "맑은 고딕",
  colors: {
    dk1: "1E2530", lt1: "FFFFFF", dk2: "16365C", lt2: "EEF3F8",
    accent1: "1F5C99", accent2: "D9822B", accent3: "1F8A7A", accent4: "6B7889",
    accent5: "A9C4DE", accent6: "B5483B", hlink: "1F5C99", folHlink: "6B4E9B",
  },
};
const HEX = THEME.colors;
const W = 13.333, MX = 0.6, CW = 12.133;

// ---------------------------------------------------------------- 이미지 크기
const SIZES = {};
for (const dir of [path.join(ROOT, "figures"), path.join(ROOT, "figures", "pdf")]) {
  for (const f of fs.readdirSync(dir)) {
    if (!/\.(png|jpg)$/.test(f)) continue;
    const buf = fs.readFileSync(path.join(dir, f));
    let w, h;
    if (f.endsWith(".png")) { w = buf.readUInt32BE(16); h = buf.readUInt32BE(20); }
    else {
      let i = 2;
      while (i < buf.length) {
        const m = buf[i + 1], len = buf.readUInt16BE(i + 2);
        if (m >= 0xc0 && m <= 0xc3) { h = buf.readUInt16BE(i + 5); w = buf.readUInt16BE(i + 7); break; }
        i += 2 + len;
      }
    }
    SIZES[f] = [w, h];
  }
}
function fit(file, x, y, w, h, align = "center", valign = "middle") {
  const [iw, ih] = SIZES[file];
  const r = Math.min(w / iw, h / ih);
  const ww = iw * r, hh = ih * r;
  const xx = align === "left" ? x : x + (w - ww) / 2;
  return { x: xx, y: valign === "top" ? y : y + (h - hh) / 2, w: ww, h: hh };
}

// ---------------------------------------------------------------- 아이콘
const ICON_CACHE = {};
async function icon(name, color = "FFFFFF") {
  const key = name + color;
  if (ICON_CACHE[key]) return ICON_CACHE[key];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(md[name], { color: "#" + color, size: 256 }));
  const png = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return (ICON_CACHE[key] = "image/png;base64," + png.toString("base64"));
}
const ICONS = ["MdPrecisionManufacturing", "MdMonitorHeart", "MdElectricBolt", "MdConstruction", "MdBuild",
  "MdLocalGasStation", "MdAcUnit", "MdHub", "MdCheckCircle", "MdWarningAmber", "MdSchedule", "MdStorage",
  "MdTrendingUp", "MdFactory", "MdPlace", "MdSensors", "MdAssignment", "MdGroups", "MdShowChart", "MdSpeed",
  "MdInsights", "MdEngineering", "MdVerified", "MdHandshake"];

// ---------------------------------------------------------------- 공통 헬퍼
function makePres(subject) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "수소충전소 핵심설비·부품 내구성 및 효율 향상 1단계 단계평가";
  pres.subject = subject;
  pres.author = "고등기술연구원";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "COVER",
    background: { color: HEX.dk2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 1.9, w: 7.2, h: 1.9, fontSize: 34, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "subtitle", type: "body", x: MX, y: 3.95, w: 7.2, h: 0.6, fontSize: 20, color: C.accent5, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "meta", type: "body", x: MX, y: 5.1, w: 7.2, h: 1.6, fontSize: 14, color: C.background1, valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: HEX.lt1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.3, w: 8.5, h: 0.32, fontSize: 12, bold: true, color: C.accent1, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.64, w: CW, h: 0.72, fontSize: 26, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "1단계 단계평가 · 수소충전소 핵심설비·부품 내구성 및 효율 향상", options: { x: MX, y: 7.05, w: 9, h: 0.28, fontSize: 10, color: C.accent4, margin: 0 } } },
    ],
    slideNumber: { x: 12.13, y: 7.05, w: 0.6, h: 0.28, fontSize: 10, color: C.accent4, align: "right" },
  });
  pres.defineSlideMaster({
    title: "DIVIDER",
    background: { color: HEX.dk2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 1.5, w: CW, h: 2.2, fontSize: 36, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: MX, y: 3.95, w: CW, h: 1.0, fontSize: 18, color: C.accent5, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  return { pres, C };
}

function badges(slide, pres, C, list) {
  const spec = {
    공인: { fillC: C.accent1, color: C.background1 }, 현장: { fillC: C.accent3, color: C.background1 },
    자체: { fillC: C.accent4, color: C.background1 }, 추정: { color: C.accent2, line: true },
    진행: { color: C.accent4, line: true },
  };
  let x = W - MX;
  for (const name of [...list].reverse()) {
    const s = spec[name];
    x -= 0.78;
    slide.addText(name, {
      shape: pres.shapes.ROUNDED_RECTANGLE, x, y: 0.32, w: 0.72, h: 0.3, rectRadius: 0.15,
      fill: { color: s.fillC || C.background1 },
      line: s.line ? { color: s.color, width: 1 } : undefined,
      fontSize: 11, bold: true, color: s.color, align: "center", valign: "middle", margin: 0, isTextBox: true,
      objectName: "badge-" + name,
    });
  }
}
function pill(slide, pres, C, text, x, y, kind) {
  const m = { 공인: C.accent1, 현장: C.accent3, 자체: C.accent4 };
  slide.addText(text, {
    shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w: 0.62, h: 0.26, rectRadius: 0.13,
    fill: m[kind] ? { color: m[kind] } : { color: C.background1 }, line: m[kind] ? undefined : { color: kind === "진행" ? C.accent4 : C.accent2, width: 1 },
    fontSize: 10, bold: true, color: m[kind] ? C.background1 : kind === "진행" ? C.accent4 : C.accent2, align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
}
function source(slide, C, text) {
  slide.addText(text, { x: MX, y: 6.68, w: CW, h: 0.28, fontSize: 10, color: C.accent4, margin: 0, isTextBox: true, objectName: "source" });
}
function card(slide, pres, C, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08, fill: { color: opts.fill || C.background2 },
    line: opts.line ? { color: opts.line, width: 1 } : undefined, objectName: opts.name || "card",
  });
}
async function iconCircle(slide, pres, C, name, x, y, d = 0.56, fill) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill || C.accent1 }, objectName: "icon-bg" });
  const pad = d * 0.2;
  slide.addImage({ data: await icon(name), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: name });
}
function contentSlide(pres, section, kicker, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
const chartBase = (C) => ({
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, catAxisLabelColor: HEX.dk1, valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: HEX.dk1,
  showTitle: true, titleFontSize: 13, titleColor: HEX.dk2, barGapWidthPct: 60, catAxisLineShow: true,
  catAxisLineColor: "C9D3DE",
});

// ================================================================ 슬라이드
async function sCover(pres, C, variant) {
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  s.addText("수소충전소 핵심설비·부품\n내구성 및 효율 향상", { placeholder: "title" });
  s.addText(variant === "field" ? "1단계(2024.04~2026.12) 현장평가" : "1단계(2024.04~2026.12) 단계평가", { placeholder: "subtitle" });
  s.addText([
    { text: "RS-2024-00419764 · 기후에너지환경부 / 한국에너지기술평가원", options: { breakLine: true } },
    { text: "주관 고등기술연구원 · 공동 11개 기관 · 지원 충주시, 충북도", options: { breakLine: true } },
    { text: "발표 [발표자] · [일자]" },
  ], { placeholder: "meta" });
  const im = fit("p049_station_overview.jpg", 8.2, 1.2, 4.55, 5.1);
  s.addImage({ path: FIG("p049_station_overview.jpg"), ...im, rounding: false, altText: "충주바이오그린수소충전소 전경과 주요설비" });
  s.addText("충주바이오그린수소충전소 전경(보고서 p.49)", { x: 8.2, y: im.y + im.h + 0.08, w: 4.55, h: 0.28, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true });
}

async function sConclusion(pres, C, section) {
  const s = contentSlide(pres, section, "1단계 결론", "부품 내구성은 공인시험으로 입증했고, 실증 인프라는 이미 착공했습니다");
  const cards = [
    ["MdPrecisionManufacturing", "내구", "1.5배+", "", "핵심부품 4종 공인 달성\n피스톤 씰은 3.3배", "공인"],
    ["MdMonitorHeart", "진단", "92일 전", "", "실제 중정비 고장을\n조기경보로 사전 감지", "현장"],
    ["MdElectricBolt", "효율", "9.5", " kWh/kg", "공인 기준선 확정\n고효율 칠러 kg당 −21.4%", "공인"],
    ["MdConstruction", "실증", "착공", "", "인허가 3건 완료\n2단계 공사 조기 착수", "현장"],
  ];
  const w = 2.883, g = 0.2, y = 1.65, h = 3.55;
  for (let i = 0; i < 4; i++) {
    const [ic, lab, big, unit, desc, b] = cards[i];
    const x = MX + i * (w + g);
    card(s, pres, C, x, y, w, h, { name: "card-" + lab });
    await iconCircle(s, pres, C, ic, x + 0.25, y + 0.25);
    s.addText(lab, { x: x + 0.95, y: y + 0.3, w: 1.6, h: 0.46, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addText([{ text: big, options: { fontSize: 36, bold: true, color: C.accent1 } }, { text: unit, options: { fontSize: 16, color: C.accent1 } }],
      { x: x + 0.25, y: y + 1.05, w: w - 0.4, h: 0.8, margin: 0, valign: "middle", isTextBox: true });
    s.addText(desc, { x: x + 0.25, y: y + 1.95, w: w - 0.4, h: 0.95, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    pill(s, pres, C, b, x + 0.25, y + 3.05, b);
  }
  s.addText([
    { text: "다음 단계  ", options: { bold: true, color: C.accent5 } },
    { text: "개선 설비를 충전소에 설치해 내구성과 에너지 절감을 현장에서 검증합니다", options: { color: C.background1 } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, x: MX, y: 5.45, w: CW, h: 0.9, rectRadius: 0.08, fill: { color: C.text2 }, fontSize: 17, margin: 14, valign: "middle", isTextBox: true });
  source(s, C, "출처: 단계보고서 p.12, p.13, p.22, p.36");
  s.addNotes("결론부터 말씀드리겠습니다. 1단계에서는 고장이 잦은 핵심부품 네 가지를 개선해 공인시험에서 내구성 1.5배 이상을 확인했습니다. 피스톤 씰은 3.3배입니다. 운전데이터로 실제 고장을 석 달 전에 감지하는 체계를 만들었고, 에너지는 공인 기준선을 확정하고 고효율 칠러를 확보했습니다. 그리고 이 개선 설비들을 설치할 실증 설비동은 2단계 공사를 앞당겨 이미 착공했습니다. 2단계는 이것을 충전소에서 증명하는 단계입니다.");
}

async function sProblem(pres, C) {
  const s = contentSlide(pres, "과제 개요", "과제 개요 | 배경", "국산 충전소 설비는 '고장·전력·운영실적' 세 가지가 약합니다");
  const cols = [
    ["MdBuild", "잦은 고장", "117 hr", "기존 다이어프램 평균 교체 운전시간(현장)\n릴리프밸브 수명 약 3개월"],
    ["MdElectricBolt", "높은 전력", "9.5 kWh/kg", "수소 1 kg당 충전소 전력 사용량\n20% 이상 절감 필요"],
    ["MdAssignment", "운영실적 부족", "Track Record", "효율·내구성 검증과 현장 운영실적 부족\n중소기업 중심 공급 구조"],
  ];
  const w = 3.91, g = 0.2, y = 1.7, h = 4.05;
  for (let i = 0; i < 3; i++) {
    const [ic, lab, big, desc] = cols[i];
    const x = MX + i * (w + g);
    card(s, pres, C, x, y, w, h);
    await iconCircle(s, pres, C, ic, x + 0.3, y + 0.3, 0.7);
    s.addText(lab, { x: x + 1.15, y: y + 0.38, w: w - 1.4, h: 0.55, fontSize: 20, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addText(big, { x: x + 0.3, y: y + 1.3, w: w - 0.6, h: 0.9, fontSize: 34, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    s.addText(desc, { x: x + 0.3, y: y + 2.35, w: w - 0.6, h: 1.4, fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  }
  s.addText("정책 배경: 美 Clean Hydrogen Strategy & Roadmap(’23.6), 日 수소기본전략 개정(’23.6), EU AFIR(2030년까지 핵심망 200 km 간격 충전소 의무화)",
    { x: MX, y: 5.95, w: CW, h: 0.5, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.2, p.26");
  s.addNotes("국내 충전소 핵심설비는 주로 중소기업이 공급합니다. 그런데 부품 교체가 잦고, 수소 1 kg에 약 9.5 kWh를 쓰며, 현장 운영 실적이 부족해 신뢰를 얻지 못하고 있습니다.");
}

async function sGoal(pres, C) {
  const s = contentSlide(pres, "과제 개요", "과제 개요 | 목표와 단계 구조", "1단계는 개발과 체계 구축, 2단계는 충전소 현장 검증입니다");
  const goals = [["MdPrecisionManufacturing", "핵심부품·설비 내구성 1.5배"], ["MdElectricBolt", "에너지 사용량 20% 이상 절감"], ["MdHub", "K-수소충전소 운영모델 제시"]];
  const gw = 3.91;
  s.addText("최종목표", { x: MX, y: 1.6, w: 3, h: 0.35, fontSize: 14, bold: true, color: C.accent4, margin: 0, isTextBox: true });
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (gw + 0.2);
    card(s, pres, C, x, 2.0, gw, 0.95, { fill: C.background1, line: HEX.accent5 });
    await iconCircle(s, pres, C, goals[i][0], x + 0.2, 2.18, 0.58);
    s.addText(goals[i][1], { x: x + 0.95, y: 2.0, w: gw - 1.05, h: 0.95, fontSize: 17, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
  }
  const st = [
    ["1단계  2024.04~2026.12", "개발·체계 구축", ["핵심부품 개선, 시제품 제작과 공인 평가", "신뢰성 관리모드(FMECA·조기경보) 구축", "실증 사이트 설계와 인허가"], C.accent1],
    ["2단계  2027.01~2029.03", "충전소 현장 검증", ["개선 설비 충전소 설치와 장기 실증", "에너지 절감량 공인기관 검증", "HyEReD DB와 K-충전소 운영모델"], C.text2],
  ];
  const sw = 5.4;
  for (let i = 0; i < 2; i++) {
    const [hd, sub, items, col] = st[i];
    const x = i === 0 ? MX : MX + CW - sw;
    card(s, pres, C, x, 3.3, sw, 3.0, { fill: col });
    s.addText(hd, { x: x + 0.3, y: 3.45, w: sw - 0.6, h: 0.4, fontSize: 14, color: C.accent5, bold: true, margin: 0, isTextBox: true });
    s.addText(sub, { x: x + 0.3, y: 3.85, w: sw - 0.6, h: 0.55, fontSize: 22, bold: true, color: C.background1, margin: 0, isTextBox: true });
    s.addText(items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })),
      { x: x + 0.3, y: 4.5, w: sw - 0.6, h: 1.65, fontSize: 15, color: C.background1, paraSpaceAfter: 6, margin: 0, valign: "top", isTextBox: true });
  }
  s.addShape(pres.shapes.RIGHT_ARROW, { x: MX + sw + 0.2, y: 4.45, w: CW - 2 * sw - 0.4, h: 0.7, fill: { color: C.accent2 }, objectName: "arrow" });
  s.addText("TRL 7 → 8", { x: MX + sw + 0.05, y: 5.25, w: CW - 2 * sw - 0.1, h: 0.35, fontSize: 13, bold: true, color: C.accent2, align: "center", margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 요약문 p.3, p.5");
  s.addNotes("과제는 단계 구조로 설계되어 있습니다. 1단계에서는 부품을 개선하고 관리체계를 만들며, 2단계에서는 이를 충전소에 설치해 내구성과 에너지 절감을 검증합니다.");
}

async function sOrg(pres, C) {
  const s = contentSlide(pres, "과제 개요", "과제 개요 | 추진체계", "12개 기관이 '부품 → 설비 → 충전소'로 연결되어 움직입니다");
  const layers = [
    ["충전소 운영모델", "고등기술연구원(총괄·진단·실증) · 고려대(경제성·탄소)\n창원산업진흥원, 코하이젠(현장 실증)", C.text2],
    ["설비", "압축기: GTC, 범한산업 · 충전기: MSENG, 제아ENG\n냉각기: 삼정이엔씨", C.accent1],
    ["부품", "피스톤 씰: GTC, 한국자동차연구원\n다이어프램: 범한산업, 한국자동차연구원\nIR 노즐: MSENG · 밸브: MT.H · PCHE: 에너진", "6E93BF"],
  ];
  const lw = 7.9, lh = 1.35;
  for (let i = 0; i < 3; i++) {
    const y = 1.7 + i * (lh + 0.17);
    card(s, pres, C, MX, y, lw, lh, { fill: C.background2 });
    s.addText(layers[i][0], { shape: pres.shapes.ROUNDED_RECTANGLE, x: MX + 0.2, y: y + 0.3, w: 1.9, h: lh - 0.6, rectRadius: 0.08, fill: { color: layers[i][2] }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(layers[i][1], { x: MX + 2.3, y, w: lw - 2.45, h: lh, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  }
  s.addText("지원: 충주시 · 충북도", { x: MX, y: 6.27, w: lw, h: 0.3, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
  const rx = MX + lw + 0.35, rw = CW - lw - 0.35;
  s.addText("데이터 환류 사례", { x: rx, y: 1.7, w: rw, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
  const cases = [
    ["다이어프램", "창원산업진흥원이 현장 교체 데이터를 분석해 범한산업 후속 설계 개선에 반영(p.26)"],
    ["조기경보", "코하이젠 운전·고장이력 데이터로 주관기관 진단 알고리즘 개발(p.28)"],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 2.2 + i * 1.75;
    card(s, pres, C, rx, y, rw, 1.6, { fill: C.background1, line: HEX.accent5 });
    await iconCircle(s, pres, C, "MdHub", rx + 0.2, y + 0.2, 0.46, C.accent3);
    s.addText(cases[i][0], { x: rx + 0.8, y: y + 0.2, w: rw - 1, h: 0.46, fontSize: 15, bold: true, color: C.accent3, valign: "middle", margin: 0, isTextBox: true });
    s.addText(cases[i][1], { x: rx + 0.2, y: y + 0.75, w: rw - 0.4, h: 0.8, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  }
  s.addText("현장 데이터 → 고장모드 정량화 → 개선 부품 → 현장 재확인", { x: rx, y: 5.8, w: rw, h: 0.5, fontSize: 13, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 표지, p.26, p.28");
  s.addNotes("현장 운영기관이 고장 데이터를 내면, 주관기관이 고장모드를 정량화해 개선 대상을 정하고, 부품 기업이 개선품을 만들어 다시 현장에서 확인합니다. 다이어프램은 이 순환을 실제로 한 바퀴 돈 사례입니다.");
}

async function sTimeline(pres, C) {
  const s = contentSlide(pres, "1단계 성과", "1단계 성과 | 추진경과", "연차별 개발 항목을 모두 완료했습니다");
  const years = [
    ["2024", "기준 수립", ["회전설비 FMEA 수립(p.10)", "피스톤 씰 기준치 300 hr 공인(p.20)", "유량조절밸브 KS 인증(2024.08, p.37)"]],
    ["2025", "시제품·공인 기준", ["에너지 기준선 9.5 KTC 성적서(p.12)", "피스톤 씰 450 hr, 150% 공인(p.21)", "이상감지 모델 2종 개발(p.12)"]],
    ["2026", "공인 목표 달성", ["피스톤 씰 1,000 hr, 330% 공인(p.22)", "IR 리시버 15,000 cycle 공인(p.30)", "인허가 완료, 설비동 기초 타설(p.13)"]],
  ];
  const lineY = 2.55;
  s.addShape(pres.shapes.LINE, { x: MX + 0.4, y: lineY, w: CW - 0.8, h: 0, line: { color: HEX.accent5, width: 3 }, objectName: "timeline-axis" });
  const w = 3.91;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (w + 0.2);
    s.addShape(pres.shapes.OVAL, { x: x + w / 2 - 0.22, y: lineY - 0.22, w: 0.44, h: 0.44, fill: { color: C.accent1 }, line: { color: C.background1, width: 3 } });
    s.addText(years[i][0], { x, y: 1.6, w, h: 0.5, fontSize: 24, bold: true, color: C.text2, align: "center", margin: 0, isTextBox: true });
    card(s, pres, C, x, 3.05, w, 3.3);
    s.addText(years[i][1], { x: x + 0.3, y: 3.2, w: w - 0.6, h: 0.5, fontSize: 18, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    s.addText(years[i][2].map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < 2 } })),
      { x: x + 0.3, y: 3.8, w: w - 0.5, h: 2.4, fontSize: 15, color: C.text1, paraSpaceAfter: 10, margin: 0, valign: "top", isTextBox: true });
  }
  source(s, C, "출처: 단계보고서 2-2절 기관별 추진 실적(p.10~39)");
  s.addNotes("1차년도에 기준을 세우고, 2차년도에 시제품을 만들고 공인 기준을 받았으며, 3차년도에 공인시험으로 목표를 확인했습니다. 연차별 개발 항목은 모두 완료했습니다.");
}

async function sDashboard(pres, C, section) {
  const s = contentSlide(pres, section, "1단계 성과 | 정량성과", "정량지표 12개 중 10개를 달성했습니다");
  const d = [
    ["SCIE 논문", "14", "5", "건"], ["국내 논문", "8", "8", "건"], ["SW 저작권", "28", "5", "건"], ["사업화", "7", "3", "건"],
    ["사업화 매출", "3,280", "2,000", "백만 원"], ["고용 창출", "35", "16", "명"], ["현장시험", "14", "10", "건"], ["시제품", "7", "5", "건"],
    ["MOU", "12", "2", "건"], ["국가인증", "2", "1", "건"], ["특허 등록", "4", "10", "건", "출원 14건"], ["기술실시", "2", "4", "회", ""],
  ];
  const cols = 6, g = 0.18, w = (CW - g * (cols - 1)) / cols, h = 2.15;
  for (let i = 0; i < d.length; i++) {
    const [lab, a, t, u, note] = d[i];
    const miss = note !== undefined;
    const x = MX + (i % cols) * (w + g), y = 1.65 + Math.floor(i / cols) * (h + g);
    card(s, pres, C, x, y, w, h, miss ? { fill: C.background1, line: HEX.accent2, name: "tile-miss" } : { name: "tile" });
    s.addText(lab, { x: x + 0.18, y: y + 0.18, w: w - 0.36, h: 0.4, fontSize: 14, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addText([{ text: a, options: { fontSize: a.length > 3 ? 28 : 34, bold: true, color: miss ? C.accent2 : C.accent1 } }],
      { x: x + 0.18, y: y + 0.65, w: w - 0.36, h: 0.75, margin: 0, valign: "middle", isTextBox: true });
    s.addText(`목표 ${t} ${u}`, { x: x + 0.18, y: y + 1.45, w: w - 0.36, h: 0.32, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
    if (miss) s.addText(note ? `미달 · ${note}` : "미달", { x: x + 0.18, y: y + 1.75, w: w - 0.36, h: 0.3, fontSize: 11, bold: true, color: C.accent2, margin: 0, isTextBox: true });
  }
  s.addText("SCIE 14건 중 IF 10 이상 4건 · 실적 기준: 정량성과표(2026.10.30 제출본)", { x: MX, y: 6.33, w: CW, h: 0.3, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.291 정량적 연구개발성과표, p.6");
  s.addNotes("정량지표 12개 중 10개를 달성했습니다. 특허 등록과 기술이전 두 가지는 미달인데, 뒤에서 원인과 만회 계획을 따로 말씀드리겠습니다. (발표 전 요약문·4장 수치를 성과표 값으로 통일할 것)");
}

async function sPiston(pres, C) {
  const s = contentSlide(pres, "기술성과", "기술성과 | 내구 ① 피스톤 압축기 (GTC · 한국자동차연구원)", "피스톤 씰 내구성 300 → 1,000 hr로 목표의 2배 이상 달성");
  badges(s, pres, C, ["공인"]);
  s.addChart(pres.charts.BAR, [{ name: "내구시험 시간", labels: ["기준치(2024)", "2025", "2026"], values: [300, 450, 1000] }], {
    x: MX, y: 1.6, w: 5.6, h: 4.85, barDir: "col", chartColors: ["9AA5B1", HEX.accent5, HEX.accent1], dataLabelFormatCode: "#,##0",
    title: "피스톤 씰 공인 내구시험 시간 (hr)", ...chartBase(C), valAxisMinVal: 0, valAxisMaxVal: 1150,
  });
  const bx = 6.5, bw = 3.55;
  const pts = [["330%", "기준 300 hr 대비\n(1단계 목표 150%)"], ["A7075 + 이중 스프링씰", "피스톤 소재와 씰 구조 개선, 특허 출원"], ["신규 압축기 제작 완료", "개선 씰 적용, 2단계 설비동 설치 예정"]];
  for (let i = 0; i < 3; i++) {
    const y = 1.7 + i * 1.55;
    card(s, pres, C, bx, y, bw, 1.4);
    s.addText(pts[i][0], { x: bx + 0.25, y: y + 0.15, w: bw - 0.5, h: 0.55, fontSize: i === 0 ? 28 : 17, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
    s.addText(pts[i][1], { x: bx + 0.25, y: y + 0.72, w: bw - 0.5, h: 0.6, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  }
  const im = fit("p155_piston_parts.jpg", 10.3, 1.7, 2.43, 3.6);
  s.addImage({ path: FIG("p155_piston_parts.jpg"), ...im, altText: "개선 피스톤 부품 사진" });
  s.addText("개선 피스톤 부품(p.155)", { x: 10.3, y: im.y + im.h + 0.06, w: 2.43, h: 0.28, fontSize: 10, color: C.accent4, align: "center", margin: 0, isTextBox: true });
  s.addText("한국자동차연구원: 450 hr 연속 내구시험 성적서, 온도계산 프로그램(p.19)", { x: 10.3, y: 5.75, w: 2.43, h: 0.75, fontSize: 11, color: C.accent4, margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.19~22 공인인증평가 성적서(2024·2025·2026)");
  s.addNotes("피스톤 씰은 수명이 짧아 수소 보충과 정비가 잦던 부품입니다. 구조를 바꾼 결과 공인시험에서 기준의 3.3배인 1,000시간을 달성했습니다. 1단계 목표 1.5배의 두 배가 넘습니다. 이 씰을 적용한 신규 압축기 제작도 마쳤습니다.");
}

async function sDiaphragm(pres, C) {
  const s = contentSlide(pres, "기술성과", "기술성과 | 내구 ② 다이어프램 압축기 (범한산업 · 한국자동차연구원 · 창원산업진흥원)", "다이어프램 압축기는 현장 교체주기가 4.2배 늘었습니다");
  badges(s, pres, C, ["공인", "현장"]);
  const cw = 5.95;
  // 시험실
  s.addText("시험실", { x: MX, y: 1.6, w: 2, h: 0.4, fontSize: 16, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  pill(s, pres, C, "공인", MX + 0.9, 1.67, "공인");
  card(s, pres, C, MX, 2.1, cw, 4.35);
  const im1 = fit("p173_diaphragm_compressor_prototype.jpg", MX + 0.25, 2.3, 2.6, 3.95);
  s.addImage({ path: FIG("p173_diaphragm_compressor_prototype.jpg"), ...im1, altText: "내구 개선 다이어프램 압축기 시제품" });
  const tx = MX + 3.05, tw = cw - 3.25;
  s.addText([{ text: "1,561 hr", options: { fontSize: 30, bold: true, color: C.accent1, breakLine: true } }, { text: "PFA 코팅 다이어프램 내구\n(목표 1,500 hr, 150%)", options: { fontSize: 14, color: C.text1 } }],
    { x: tx, y: 2.35, w: tw, h: 1.7, margin: 0, valign: "top", isTextBox: true });
  s.addText([{ text: "−22.6%", options: { fontSize: 30, bold: true, color: C.accent1, breakLine: true } }, { text: "캐비티 형상 최적화로\n최대응력 저감", options: { fontSize: 14, color: C.text1 } }],
    { x: tx, y: 4.3, w: tw, h: 1.7, margin: 0, valign: "top", isTextBox: true });
  // 현장
  const rx = MX + cw + 0.23;
  s.addText("현장(가포 충전소)", { x: rx, y: 1.6, w: 3, h: 0.4, fontSize: 16, bold: true, color: C.accent3, margin: 0, isTextBox: true });
  pill(s, pres, C, "현장", rx + 2.35, 1.67, "현장");
  card(s, pres, C, rx, 2.1, cw, 4.35);
  s.addChart(pres.charts.BAR, [{ name: "평균 교체 운전시간", labels: ["기존품", "개발품"], values: [117, 492.7] }], {
    x: rx + 0.15, y: 2.2, w: 3.4, h: 4.15, barDir: "col", chartColors: ["9AA5B1", HEX.accent3], dataLabelFormatCode: "#,##0.#",
    title: "평균 교체 운전시간 (hr)", ...chartBase(C), valAxisMinVal: 0, valAxisMaxVal: 580,
  });
  s.addText([{ text: "4.21배", options: { fontSize: 30, bold: true, color: C.accent3, breakLine: true } }, { text: "교체주기 연장", options: { fontSize: 14, color: C.text1 } }],
    { x: rx + 3.65, y: 2.35, w: 2.15, h: 1.5, margin: 0, valign: "top", isTextBox: true });
  s.addText([{ text: "3 → 8개월+", options: { fontSize: 22, bold: true, color: C.accent3, breakLine: true } }, { text: "개선 릴리프밸브\n사용 수명", options: { fontSize: 14, color: C.text1 } }],
    { x: rx + 3.65, y: 4.3, w: 2.15, h: 1.7, margin: 0, valign: "top", isTextBox: true });
  source(s, C, "출처: 단계보고서 p.19, p.24 공인인증평가 성적서, p.26 상용충전소 운영 현황 보고서");
  s.addNotes("다이어프램은 실제 상용 충전소에 먼저 적용했습니다. 교체주기가 117시간에서 493시간으로 4배 넘게 늘었고, 릴리프밸브 수명도 3개월에서 8개월 이상으로 늘었습니다. 시험실에서는 공인시험으로 1,561시간을 확인했습니다. (4.21배의 표본 수와 기간은 백업으로 준비)");
}

async function sDispenser(pres, C) {
  const s = contentSlide(pres, "기술성과", "기술성과 | 내구 ③ 충전기 (MSENG · MT.H콘트롤밸브 · 제아ENG)", "충전기 핵심부품 1.5배 달성, 현장에서 1,672회 실충전");
  badges(s, pres, C, ["공인", "현장"]);
  const items = [
    ["p194_ir_receiver.jpg", "IR 리시버 (MSENG)", "15,000 cycle", "기준 10,000 cycle 대비 150%\n−40~20 ℃ 온도별 내구시험", ["공인"]],
    ["p277_valve_field_charging.jpg", "유량조절밸브 (MT.H)", "153,030회", "96.3 MPa 작동시험 150%\n10개월 실충전 1,672회, 제조사 2곳 3기 공급", ["공인", "현장"]],
    ["p219_dual_dispenser.jpg", "듀얼 디스펜서 (제아ENG)", "60 g/s급", "동시충전 시제품 제작·시운전\n120 g/s 공인시험 2026.11 예정", ["진행"]],
  ];
  const w = 3.91, g = 0.2, y = 1.6, h = 4.9;
  for (let i = 0; i < 3; i++) {
    const [img, lab, big, desc, bs] = items[i];
    const x = MX + i * (w + g);
    card(s, pres, C, x, y, w, h);
    const im = fit(img, x + 0.2, y + 0.2, w - 0.4, 2.0);
    s.addImage({ path: FIG(img), ...im, altText: lab });
    s.addText(lab, { x: x + 0.25, y: y + 2.35, w: w - 0.5, h: 0.4, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addText(big, { x: x + 0.25, y: y + 2.8, w: w - 0.5, h: 0.65, fontSize: 28, bold: true, color: i === 2 ? C.accent4 : C.accent1, margin: 0, valign: "middle", isTextBox: true });
    s.addText(desc, { x: x + 0.25, y: y + 3.5, w: w - 0.5, h: 0.95, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    bs.forEach((b, k) => pill(s, pres, C, b, x + 0.25 + k * 0.7, y + 4.5, b));
  }
  source(s, C, "출처: 단계보고서 p.29~32, p.39 공인인증평가 성적서, p.277, p.319");
  s.addNotes("충전기 핵심부품 두 가지는 공인시험에서 1.5배를 달성했습니다. 특히 밸브는 계획보다 앞당겨 설치해 10개월 동안 실제 차량 충전 1,672회를 수행했고, 이미 충전기 제조사 두 곳에 세 기를 공급했습니다. 듀얼 디스펜서는 시제품을 만들었고 11월에 공인시험을 받습니다. (시험 결과가 나오면 이 장에 반영)");
}

async function sChiller(pres, C) {
  const s = contentSlide(pres, "기술성과", "기술성과 | 내구·효율 ④ 냉각기 (삼정이엔씨 · 에너진)", "고효율 칠러로 충전 1 kg당 전력을 21.4% 줄였습니다");
  badges(s, pres, C, ["공인"]);
  s.addChart(pres.charts.BAR, [{ name: "충전 1 kg당 칠러 전력", labels: ["기존 칠러", "개선 칠러"], values: [0.7829, 0.6155] }], {
    x: MX, y: 1.6, w: 4.3, h: 4.2, barDir: "col", chartColors: ["9AA5B1", HEX.accent1], dataLabelFormatCode: "0.0000",
    title: "충전 1 kg당 칠러 전력 (kWh/kg, KCL)", ...chartBase(C), valAxisMinVal: 0, valAxisMaxVal: 0.92,
  });
  s.addText([{ text: "−21.4%", options: { fontSize: 30, bold: true, color: C.accent1, breakLine: true } }, { text: "충전 시 kg당 전력", options: { fontSize: 14, color: C.text1 } }],
    { x: MX + 4.5, y: 1.8, w: 2.6, h: 1.3, margin: 0, isTextBox: true });
  s.addText([{ text: "22.89 → 20.30", options: { fontSize: 20, bold: true, color: C.accent1, breakLine: true } }, { text: "충전대기 1시간당 kWh(−11.3%)", options: { fontSize: 14, color: C.text1 } }],
    { x: MX + 4.5, y: 3.35, w: 2.6, h: 1.3, margin: 0, isTextBox: true });
  s.addText("대기 시 절감은 운전 로직 개선과 함께 2단계에서 확대", { x: MX, y: 5.95, w: 7, h: 0.5, fontSize: 13, color: C.accent2, bold: true, margin: 0, isTextBox: true });
  const rx = 7.95, rw = W - MX - rx;
  card(s, pres, C, rx, 1.6, rw, 2.35);
  const i1 = fit("p265_chiller_prototype.jpg", rx + 0.15, 1.72, rw - 0.3, 1.45);
  s.addImage({ path: FIG("p265_chiller_prototype.jpg"), ...i1, altText: "냉각기 시제품" });
  s.addText("대형 상용차용 냉각기 시제품(216 kg/h)", { x: rx + 0.2, y: 3.3, w: rw - 0.4, h: 0.5, fontSize: 14, bold: true, color: C.text2, margin: 0, isTextBox: true });
  card(s, pres, C, rx, 4.1, rw, 2.4);
  const i2 = fit("p239_pche_prototype.jpg", rx + 0.15, 4.22, 1.65, 2.16);
  s.addImage({ path: FIG("p239_pche_prototype.jpg"), ...i2, altText: "PCHE 시제품" });
  s.addText([{ text: "PCHE (에너진)", options: { bold: true, color: C.text2, breakLine: true } },
    { text: "130 MPa 이상 1시간 내압, 누설·균열 없음", options: { breakLine: true } },
    { text: "0~875 bar 반복가압 테스트베드 운용(목표 100,000회)" }],
  { x: rx + 2.0, y: 4.25, w: rw - 2.15, h: 2.1, fontSize: 14, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.33, p.36 공인인증평가 성적서(KCL), p.239, p.265, p.324");
  s.addNotes("충전 시 냉각 전력을 21% 줄인 고효율 칠러를 공인기관에서 확인했습니다. PCHE는 130 MPa 내압을 통과했고 반복가압 시험을 진행하고 있습니다. 이 칠러는 뒤에서 말씀드릴 충전소 에너지 병목을 겨냥한 설비입니다.");
}

async function sDiag(pres, C, section) {
  const s = contentSlide(pres, section, "기술성과 | 진단 · 신뢰성 관리체계 (고등기술연구원)", "운전데이터 기반 조기경보로 실제 고장을 92일 전에 감지했습니다");
  badges(s, pres, C, ["현장"]);
  const lx = MX, lw = 7.4;
  const tl = (y, title, span, marks, endLabel) => {
    card(s, pres, C, lx, y, lw, 2.3);
    s.addText(title, { x: lx + 0.25, y: y + 0.15, w: lw - 0.5, h: 0.4, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
    const ax = lx + 0.45, aw = lw - 1.6, ay = y + 1.45;
    s.addShape(pres.shapes.LINE, { x: ax, y: ay, w: aw, h: 0, line: { color: HEX.accent4, width: 2 } });
    s.addShape(pres.shapes.OVAL, { x: ax + aw - 0.14, y: ay - 0.14, w: 0.28, h: 0.28, fill: { color: C.accent6 } });
    s.addText(endLabel, { x: ax + aw - 0.55, y: ay + 0.22, w: 1.6, h: 0.5, fontSize: 12, bold: true, color: C.accent6, margin: 0, isTextBox: true });
    for (const [d, lab, col] of marks) {
      const mx = ax + aw * (1 - d / span);
      s.addShape(pres.shapes.OVAL, { x: mx - 0.12, y: ay - 0.12, w: 0.24, h: 0.24, fill: { color: col } });
      s.addText([{ text: `D−${d}`, options: { fontSize: 20, bold: true, color: col, breakLine: true } }, { text: lab, options: { fontSize: 12, color: C.text1 } }],
        { x: mx - 0.7, y: y + 0.55, w: 1.4, h: 0.75, align: "center", margin: 0, isTextBox: true });
    }
  };
  tl(1.6, "압축기 중정비 고장(2023.12.18) 사전 감지", 100, [[92, "통계 모델", C.accent1], [41, "AI 모델", C.accent3]], "고장 발생");
  tl(4.1, "다이어프램 사건 조기경보", 8, [[6.57, "모델 1", C.accent1], [2.72, "모델 2", C.accent3]], "사건 추정 시작");
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  const stats = [["148개", "고장모드 위험도(RPN) 정량화\n압축기 35 · 충전기 14 · 냉각기 99"], ["3종", "설비별 고장대응 매뉴얼"], ["9종", "이상감지 모델 비교 검증(4방식)"], ["7개", "고압 압축기 진동 측정점\nAlert/Trip 기준 설정"]];
  const hs = [1.3, 1.0, 1.0, 1.3];
  let y = 1.6;
  for (let i = 0; i < 4; i++) {
    card(s, pres, C, rx, y, rw, hs[i] - 0.1, { fill: C.background1, line: HEX.accent5 });
    s.addText(stats[i][0], { x: rx + 0.2, y, w: 1.2, h: hs[i] - 0.1, fontSize: 24, bold: true, color: C.accent1, valign: "middle", margin: 0, isTextBox: true });
    s.addText(stats[i][1], { x: rx + 1.45, y, w: rw - 1.6, h: hs[i] - 0.1, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    y += hs[i] + 0.06;
  }
  source(s, C, "출처: 단계보고서 p.12~13, p.70, p.74, p.315");
  s.addNotes("실제 중정비가 발생한 고장을 통계 모델은 92일 전, AI 모델은 41일 전에 감지했습니다. 운영사는 부품 조달과 정비 일정을 미리 잡을 수 있습니다. 고장모드 148개의 위험도를 정량화해 설비별 고장대응 매뉴얼로 현장에 돌려줬습니다. (정확도 1.000은 질문 시 평가 기간 108일·사례 수와 함께 답변)");
}

async function sEff1(pres, C) {
  const s = contentSlide(pres, "기술성과", "기술성과 | 효율 ① 기준선과 병목", "기준선 9.5를 공인 확정하고, 칠러 대기전력 병목을 찾았습니다");
  badges(s, pres, C, ["공인"]);
  s.addChart(pres.charts.BAR, [{ name: "SEC", labels: ["기준선\n(KTC 검증 실측)", "최종 목표\n(20% 절감)"], values: [9.5, 7.6] }], {
    x: MX, y: 1.6, w: 4.1, h: 4.3, barDir: "col", chartColors: ["9AA5B1", HEX.dk2], dataLabelFormatCode: "0.0",
    title: "충전소 SEC (kWh/kgH₂)", ...chartBase(C), valAxisMinVal: 0, valAxisMaxVal: 11,
  });
  s.addText("정밀 분석 43주 9.1 · 모니터링 전 기간 9.5\nKTC 공인 계측 체계로 검증", { x: MX, y: 5.95, w: 4.1, h: 0.6, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
  const rx = 5.0, rw = W - MX - rx;
  const im = fit("p095_main_chiller_standby.jpg", rx, 1.6, rw, 3.0);
  s.addImage({ path: FIG("p095_main_chiller_standby.jpg"), ...im, altText: "메인 칠러 시간별 전력과 시간대별 평균" });
  const st = [["12.9 kWh/h", "차량이 없는 야간에도\n메인 칠러 평균 전력"], ["38.7%", "칠러 전력량 중 야간 대기\n(약 155 kWh/일)"], ["−40 ℃", "사전 냉각 상시 유지\n(SAE J2601 T40)"]];
  const sw = (rw - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    const x = rx + i * (sw + 0.2);
    card(s, pres, C, x, 4.75, sw, 1.75);
    s.addText(st[i][0], { x: x + 0.2, y: 4.85, w: sw - 0.4, h: 0.6, fontSize: 22, bold: true, color: C.accent2, margin: 0, valign: "middle", isTextBox: true });
    s.addText(st[i][1], { x: x + 0.2, y: 5.5, w: sw - 0.4, h: 0.9, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  }
  source(s, C, "출처: 단계보고서 p.12, p.67 KTC 성적서(GT2025-00687), p.95~96 · 목표 7.6은 기준선의 20% 절감 환산값");
  s.addNotes("공인기관과 함께 이 충전소의 기준선을 9.5로 확정했습니다. 그리고 차량이 없는 밤에도 칠러가 영하 40도를 유지하느라 전력을 쓰는 구조가 가장 큰 낭비라는 것을 실측으로 확인했습니다.");
}

async function sEff2(pres, C, section) {
  const s = contentSlide(pres, section, "기술성과 | 효율 ② 절감 수단과 검증 경로", "하드웨어·운영 두 축으로 줄이고, 2단계에 충전소에서 검증합니다");
  const tw = 5.95;
  const tracks = [
    ["MdPrecisionManufacturing", "Track 1  하드웨어", C.accent1, ["고효율 칠러: 충전 kg당 −21.4%(KCL)", "압축기 동력전달 개선(플라이휠, 고효율 벨트)", "고내구 PCHE"]],
    ["MdInsights", "Track 2  운영", C.accent3, ["피크·오프피크 차등 운전(T40 / T30)", "절전 모드: Standby(10분) → Deep Sleep(60분)", "부하 패턴 기반 시간대별 제어"]],
  ];
  for (let i = 0; i < 2; i++) {
    const [ic, hd, col, items] = tracks[i];
    const x = MX + i * (tw + 0.233);
    card(s, pres, C, x, 1.6, tw, 2.6);
    await iconCircle(s, pres, C, ic, x + 0.25, 1.78, 0.56, col);
    s.addText(hd, { x: x + 0.95, y: 1.78, w: tw - 1.2, h: 0.56, fontSize: 18, bold: true, color: col, valign: "middle", margin: 0, isTextBox: true });
    s.addText(items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < 2 } })),
      { x: x + 0.3, y: 2.5, w: tw - 0.6, h: 1.6, fontSize: 15, color: C.text1, paraSpaceAfter: 6, margin: 0, valign: "top", isTextBox: true });
  }
  // 시뮬레이션(추정)
  card(s, pres, C, MX, 4.4, 4.9, 2.1, { fill: C.background1, line: HEX.accent2 });
  pill(s, pres, C, "추정", MX + 0.25, 4.55, "추정");
  s.addText("1차 시뮬레이션(메인 칠러 야간만)", { x: MX + 1.0, y: 4.5, w: 3.8, h: 0.42, fontSize: 14, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
  s.addText([{ text: "야간 대기 155 → 23 kWh/일(−85%)", options: { breakLine: true } }, { text: "최근 12개월 실측 8,629 h에 적용한 보수적 산정", options: { color: C.accent4 } }],
    { x: MX + 0.25, y: 5.05, w: 4.4, h: 1.3, fontSize: 14, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 4, isTextBox: true });
  // 검증 경로
  const vx = MX + 5.1, vw = CW - 5.1;
  s.addText("2단계 검증 방법  ·  목표 9.5 → 7.6 kWh/kg(−20%)", { x: vx, y: 4.4, w: vw, h: 0.42, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
  const steps = ["적용 전후\n설비별 비교", "처리량 보정\nSEC로 효과 분리", "외부전문가·\n공인기관 입회"];
  const sw = (vw - 0.6) / 3;
  for (let i = 0; i < 3; i++) {
    const x = vx + i * (sw + 0.3);
    s.addText([{ text: `${i + 1}  `, options: { bold: true, color: C.accent5 } }, { text: steps[i], options: { color: C.background1 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, x, y: 4.95, w: sw, h: 1.5, rectRadius: 0.08, fill: { color: C.text2 }, fontSize: 14, valign: "middle", margin: 8, isTextBox: true });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + sw + 0.04, y: 5.55, w: 0.22, h: 0.3, fill: { color: C.accent2 } });
  }
  source(s, C, "출처: 단계보고서 p.36, p.314, p.363~366 · 시뮬레이션은 보고서 기재 추정값");
  s.addNotes("절감은 고효율 설비와 운전 로직, 두 축으로 진행합니다. 칠러 야간 운전만 바꿔도 대기전력을 85% 줄일 수 있다는 것을 실측 데이터로 확인했고, 2단계에서 설비 교체와 운전 최적화를 통합해 공인기관 입회 아래 충전소 단위로 검증합니다. 적용 전후 비교는 처리량 보정 SEC로 처리량 효과를 분리해서 평가합니다.");
}

async function sInfra(pres, C, section) {
  const s = contentSlide(pres, section, "실증 인프라", "2단계 실증 공사를 1단계에 앞당겨 착공했고, 인허가를 모두 받았습니다");
  badges(s, pres, C, ["현장"]);
  const lw = 6.5;
  const i1 = fit("p093_construction_photos.jpg", MX, 1.6, lw, 2.0, "left");
  s.addImage({ path: FIG("p093_construction_photos.jpg"), ...i1, altText: "설비동 공사 사진" });
  const i2 = fit("p360_install_location.jpg", MX, i1.y + i1.h + 0.15, lw, 6.45 - (i1.y + i1.h + 0.15), "left");
  s.addImage({ path: FIG("p360_install_location.jpg"), ...i2, altText: "개선 핵심설비 설치 위치(안)" });
  const rx = MX + lw + 0.35, rw = CW - lw - 0.35;
  s.addText("충주 실증 설비동", { x: rx, y: 1.6, w: rw, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
  const checks = ["KGS 변경 기술검토 적합", "건축허가", "고압가스 제조 변경허가", "설비동 기초 타설(2단계 공사 조기 착수)"];
  for (let i = 0; i < 4; i++) {
    const y = 2.08 + i * 0.47;
    s.addImage({ data: await icon("MdCheckCircle", HEX.accent3), x: rx, y: y + 0.04, w: 0.32, h: 0.32 });
    s.addText(checks[i], { x: rx + 0.45, y, w: rw - 0.45, h: 0.4, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  }
  card(s, pres, C, rx, 4.1, rw, 1.1);
  s.addText([{ text: "설비 제작", options: { bold: true, color: C.text2, breakLine: true } }, { text: "고압 압축기 제작 완료(2026.07) · 냉동설비 도장·배치 · 듀얼 디스펜서 제작 중" }],
    { x: rx + 0.2, y: 4.15, w: rw - 0.4, h: 1.0, fontSize: 14, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  card(s, pres, C, rx, 5.35, rw, 1.1);
  s.addText([{ text: "세종대평 충전기 실증", options: { bold: true, color: C.text2, breakLine: true } }, { text: "2단계 → 1단계로 1년 앞당김 · KGS 신청(2026.09.29) · 2026.12 준공 예정" }],
    { x: rx + 0.2, y: 5.4, w: rw - 0.4, h: 1.0, fontSize: 14, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  source(s, C, "출처: 단계보고서 p.13, p.28, p.93, p.315, p.317, p.360");
  s.addNotes("상용 충전소를 운영하면서 설비를 추가해야 하는 조건이었지만, 인허가를 모두 마치고 2단계 공사를 앞당겨 착공했습니다. 고압 압축기는 이미 제작을 마쳤습니다. 세종 충전소 충전기 실증도 당초보다 1년 앞당겼습니다. 2단계가 시작되는 날부터 설치에 들어갈 수 있습니다.");
}

function statusCell(C, st) {
  const col = { 달성: C.accent1, 진행: C.accent4, 미달: C.accent2 }[st.split("(")[0]];
  return { text: "● " + st, options: { color: col, bold: true } };
}
async function sPerf(pres, C) {
  const s = contentSlide(pres, "성능지표·미달", "성능지표", "부품 내구성 지표는 모두 공인시험으로 달성했습니다");
  const H = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
  const rows = [
    [H("성능지표"), H("1단계 목표"), H("1단계 실적"), H("검증"), H("상태")],
    ["피스톤식 압축기 내구성", "150%(450 hr)", "330%(1,000 hr)", "공인", "달성"],
    ["다이어프램식 압축기 내구성", "150%(1,500 hr)", "1,561 hr", "공인", "달성"],
    ["충전기 내구성(IR 리시버)", "150%(15,000 cycle)", "15,000 cycle", "공인", "달성"],
    ["유량조절밸브 작동시험", "153,000회", "153,030회", "공인", "달성"],
    ["냉각기 에너지효율", "80%(20% 절감)", "kg당 21.4% 절감", "공인(KCL)", "달성"],
    ["운영현황·공정분석·경제성 보고서", "2 · 1 · 1건", "4 · 1 · 1건", "수요기업", "달성"],
    ["충전 유량", "120 g/s", "공인시험 2026.11", "공인", "진행"],
    ["PCHE 내구성", "(최종 100,000회)", "반복가압 시험 중", "–", "진행"],
    ["다이어프램 압축기 전력소비량", "90%(10% 절감)", "96%(4% 절감)", "공인", "미달"],
    ["충전소 소비에너지 저감", "90%(10% 절감)", "기준선 확정, 검증 2단계", "공인", "미달(이월)"],
  ].map((r, i) => i === 0 ? r : r.map((c, k) => (k === 4 ? statusCell(C, c) : { text: c, options: { color: C.text1 } })));
  s.addTable(rows, {
    x: MX, y: 1.6, w: CW, colW: [3.6, 2.4, 3.0, 1.43, 1.7], rowH: 0.42, fontSize: 14, valign: "middle",
    border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.12, 0, 0.12], fill: { color: HEX.lt1 },
  });
  source(s, C, "출처: 단계보고서 p.10~39 기관별 성능지표, p.292~293 · 제품설계 보고서 지표는 보고서 참조");
  s.addNotes("부품 내구성 관련 지표는 모두 공인시험으로 달성했습니다. 충전 유량과 PCHE는 시험을 진행하고 있고, 미달은 다이어프램 압축기 전력과 충전소 소비에너지 저감 두 가지입니다.");
}

async function sGap(pres, C) {
  const s = contentSlide(pres, "성능지표·미달", "미달 항목", "미달 항목은 원인과 만회 일정을 함께 말씀드립니다");
  const H = (t) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: C.background2 } } });
  const rows = [
    [H("항목"), H("원인"), H("만회 계획")],
    ["특허 등록 4 / 10건", "출원 14건 중 2025년 출원 10건이 심사 중", "의견제출·보정 대응 강화, 2027년 안에 등록 전환"],
    ["기술실시 2 / 4회", "수요기업이 현장 실증 결과를 도입 기준으로 요구", "실증 결과와 특허를 묶은 기술 패키지로 이전"],
    ["충전소 에너지 절감 검증", "기준선 확정 후 운전 로직 개발 중", "2026.12 로직 완료 → 2027년 안에 공인기관 검증"],
    ["다이어프램 압축기 전력", "플라이휠 단독 개선 4% 절감", "고효율 벨트(5V)로 동력전달 개선 후 재검증"],
  ].map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: { color: C.text1, bold: k === 0 } })));
  s.addTable(rows, {
    x: MX, y: 1.7, w: CW, colW: [3.0, 4.4, 4.733], rowH: [0.5, 0.85, 0.85, 0.85, 0.85], fontSize: 15, valign: "middle",
    border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.15, 0, 0.15], fill: { color: HEX.lt1 },
  });
  s.addImage({ data: await icon("MdSchedule", HEX.accent1), x: MX, y: 5.85, w: 0.4, h: 0.4 });
  s.addText("공통 원인은 '현장 실증 전 단계'이며, 네 항목 모두 2단계 1차년도에 만회 일정을 잡았습니다", { x: MX + 0.55, y: 5.82, w: CW - 0.55, h: 0.46, fontSize: 15, bold: true, color: C.accent1, valign: "middle", margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 4장 p.312~314");
  s.addNotes("미달 항목은 네 가지입니다. 공통 원인은 현장 실증 전 단계라는 점이고, 모두 2단계 1차년도에 만회 일정을 잡았습니다. (특허 등록 예상 건수는 확인 후 보완)");
}

async function sNext(pres, C, section) {
  const s = contentSlide(pres, section, "2단계 계획", "1단계 결과물이 그대로 2단계의 출발점이 됩니다");
  const pairs = [
    ["피스톤 씰 1,000 hr, 신규 압축기", "압축기 패키지 현장 장기 실증"],
    ["고효율 칠러, 절전 로직", "충전소 단위 에너지 절감 공인검증"],
    ["FMECA·조기경보", "HyEReD 신뢰성 DB, 잔여수명(RUL) 예측"],
    ["경제성 평가 툴", "K-충전소 운영모델, 탄소배출 예측"],
  ];
  const bw = 4.9, ah = 0.62, gap = 0.12;
  s.addText("1단계 결과", { x: MX, y: 1.55, w: bw, h: 0.35, fontSize: 14, bold: true, color: C.accent4, margin: 0, isTextBox: true });
  s.addText("2단계 과제", { x: MX + CW - bw, y: 1.55, w: bw, h: 0.35, fontSize: 14, bold: true, color: C.accent4, margin: 0, isTextBox: true });
  for (let i = 0; i < 4; i++) {
    const y = 1.95 + i * (ah + gap);
    s.addText(pairs[i][0], { shape: pres.shapes.ROUNDED_RECTANGLE, x: MX, y, w: bw, h: ah, rectRadius: 0.08, fill: { color: C.background2 }, fontSize: 15, color: C.text1, valign: "middle", margin: 10, isTextBox: true });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: MX + bw + 0.4, y: y + 0.16, w: CW - 2 * bw - 0.8, h: 0.3, fill: { color: C.accent2 } });
    s.addText(pairs[i][1], { shape: pres.shapes.ROUNDED_RECTANGLE, x: MX + CW - bw, y, w: bw, h: ah, rectRadius: 0.08, fill: { color: C.text2 }, fontSize: 15, bold: true, color: C.background1, valign: "middle", margin: 10, isTextBox: true });
  }
  const yrs = [["2027", "실증 설비 설치·시운전\n에너지 절감 공인검증\nHyEReD 개발"], ["2028", "장기 실증, HyEReD 유효성 검증\n조기경보 운영절차서"], ["2029.1Q", "K-충전소 운영모델\n공인 성적서·설계표준안"]];
  const yw = (CW - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (yw + 0.2);
    card(s, pres, C, x, 5.0, yw, 1.5, { fill: C.background1, line: HEX.accent5 });
    s.addText(yrs[i][0], { x: x + 0.2, y: 5.05, w: 1.3, h: 1.4, fontSize: 20, bold: true, color: C.accent1, valign: "middle", margin: 0, isTextBox: true });
    s.addText(yrs[i][1], { x: x + 1.5, y: 5.05, w: yw - 1.65, h: 1.4, fontSize: 13, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  }
  source(s, C, "출처: 단계보고서 요약문 p.4~5, 9장 p.357~366");
  s.addNotes("2단계의 모든 과제는 1단계 결과물에서 바로 이어집니다. 2027년에 실증 설비를 설치하고 에너지 절감량을 공인검증하며, 2028년 장기 실증을 거쳐 2029년 K-충전소 운영모델로 정리합니다.");
}

async function sRisk(pres, C) {
  const s = contentSlide(pres, "2단계 계획", "2단계 계획 | 리스크 관리", "실증 일정, 인허가, 데이터 리스크에 이미 대응하고 있습니다");
  const rs = [
    ["MdConstruction", "실증 설비 설치 지연", "2단계 공사를 1단계에 착공, 고압 압축기 제작 완료", "기초 타설 완료"],
    ["MdAssignment", "인허가", "충주: KGS 기술검토·건축·고압가스 변경허가 / 세종: 안전영향평가·기술검토 신청", "충주 완료 · 세종 신청"],
    ["MdStorage", "신뢰성 DB 데이터 부족", "충주 3개년 상용 운전 데이터 + 세종·가포, OREDA 구조와 미국 HyCReD 벤치마킹", "3개년 데이터 확보"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.7 + i * 1.55;
    card(s, pres, C, MX, y, CW, 1.38);
    await iconCircle(s, pres, C, rs[i][0], MX + 0.25, y + 0.38, 0.6);
    s.addText(rs[i][1], { x: MX + 1.05, y, w: 3.0, h: 1.38, fontSize: 17, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
    s.addText(rs[i][2], { x: MX + 4.15, y, w: 5.4, h: 1.38, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    s.addText(rs[i][3], { shape: pres.shapes.ROUNDED_RECTANGLE, x: MX + 9.75, y: y + 0.42, w: 2.15, h: 0.55, rectRadius: 0.27, fill: { color: C.accent3 }, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  }
  source(s, C, "출처: 단계보고서 p.13, p.28, p.93, p.315, p.358");
  s.addNotes("실증 설비 설치는 공사를 앞당겨 착공했고 고압 압축기 제작을 마쳤습니다. 인허가는 충주는 완료, 세종은 신청을 마쳤습니다. 신뢰성 DB는 3년치 상용 운전 데이터에 세종·가포 데이터를 더하고 OREDA 구조로 벤치마킹합니다.");
}

async function sBiz(pres, C) {
  const s = contentSlide(pres, "2단계 계획", "활용·사업화", "부품 국산화에서 운영모델 확산으로 사업화합니다");
  const st = [["3,280", "백만 원", "1단계 사업화 매출(목표 2,000)"], ["7", "건", "1단계 사업화(목표 3)"], ["8,000", "백만 원", "2단계 매출 목표(사업화 12건)"]];
  const w = (CW - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (w + 0.2);
    card(s, pres, C, x, 1.6, w, 1.75);
    s.addText([{ text: st[i][0], options: { fontSize: 36, bold: true, color: i === 2 ? C.accent2 : C.accent1 } }, { text: " " + st[i][1], options: { fontSize: 16, color: C.text1 } }],
      { x: x + 0.3, y: 1.7, w: w - 0.6, h: 0.9, valign: "middle", margin: 0, isTextBox: true });
    s.addText(st[i][2], { x: x + 0.3, y: 2.65, w: w - 0.6, h: 0.5, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });
  }
  const road = [["MdBuild", "부품 공급", "밸브 개선품 3기 제조사 공급\n피스톤 씰·다이어프램 적용 확대"], ["MdFactory", "설비 패키지", "압축기·냉각기·충전기\n현장 실증 후 공급"], ["MdHub", "운영모델", "K-충전소 운영절차서\n경제성·탄소배출 예측"]];
  const rw = (CW - 1.0) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (rw + 0.5);
    card(s, pres, C, x, 3.6, rw, 2.0, { fill: C.background1, line: HEX.accent5 });
    await iconCircle(s, pres, C, road[i][0], x + 0.25, 3.8, 0.56);
    s.addText(road[i][1], { x: x + 0.95, y: 3.8, w: rw - 1.1, h: 0.56, fontSize: 17, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
    s.addText(road[i][2], { x: x + 0.25, y: 4.5, w: rw - 0.5, h: 1.0, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + rw + 0.1, y: 4.45, w: 0.3, h: 0.35, fill: { color: C.accent2 } });
  }
  s.addText("기대효과: 충전소 셧다운 방지 · 운영·유지보수 비용(OPEX) 절감 · 가동률 향상", { x: MX, y: 5.85, w: CW, h: 0.45, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.6, p.39, p.291");
  s.addNotes("1단계 사업화 매출은 32억 8천만 원으로 목표를 넘었습니다. 밸브 개선품은 이미 제조사에 공급되고 있고, 2단계에는 설비 패키지와 운영모델로 사업화를 넓힙니다. (매출 품목별 내역은 백업으로 준비)");
}

async function sClosing(pres, C) {
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: "맺음말" });
  s.addText("부품은 입증했습니다.\n이제 충전소에서 증명하겠습니다", { placeholder: "title" });
  s.addText("2단계 약속: 설비 패키지 내구성 1.5배 현장 검증 · 충전소 에너지 20% 절감 공인검증 · K-충전소 운영모델", { placeholder: "body" });
  s.addNotes("1단계에서 부품과 체계를 만들었고, 실증 인프라도 준비했습니다. 2단계에서 충전소 현장에서 내구성과 에너지 절감을 증명하고 K-충전소 운영모델로 완성하겠습니다. 감사합니다.");
}

async function backups(pres, C) {
  const d = pres.addSlide({ masterName: "DIVIDER", sectionTitle: "백업" });
  d.addText("백업", { placeholder: "title" });
  d.addText("에너지 기준선 · 처리량 보정 · 절감 시뮬레이션", { placeholder: "body" });
  const two = (kicker, title, a, b, capA, capB, src, note) => {
    const s = contentSlide(pres, "백업", kicker, title);
    const w = (CW - 0.3) / 2;
    const ia = fit(a, MX, 1.7, w, 4.4, "center", "top"), ib = fit(b, MX + w + 0.3, 1.7, w, 4.4, "center", "top");
    s.addImage({ path: FIG(a), ...ia, altText: capA });
    s.addImage({ path: FIG(b), ...ib, altText: capB });
    s.addText(capA, { x: MX, y: Math.max(ia.y + ia.h, ib.y + ib.h) + 0.12, w, h: 0.3, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
    s.addText(capB, { x: MX + w + 0.3, y: Math.max(ia.y + ia.h, ib.y + ib.h) + 0.12, w, h: 0.3, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
    source(s, C, src);
    s.addNotes(note);
  };
  two("백업 B-1 | 기준선 경계", "기준선 9.5는 KTC로 검증한 계측 체계의 전 공정 실측값입니다",
    "p067_process_share.jpg", "fig15_ktc_verification.png", "공정별 에너지 비중과 SEC 비중(보고서 p.67)", "KTC 성적서 대비 모니터링 데이터 교차검증(추가 분석)",
    "출처: 단계보고서 p.12, p.67 · 교차검증은 보고서 외 추가 분석", "기준선 경계 질문 대응: 9.5는 생산·압축·충전 공정을 포함한 사업대상지 실측값이며, KTC 측정값과 모니터링 데이터의 오차는 평균 1.6%입니다.");
  two("백업 B-2 | 처리량 보정", "처리량 보정 SEC로 처리량 효과와 절감 효과를 분리합니다",
    "p068_sec_vs_throughput.jpg", "p095_full_period_sec.jpg", "처리량과 공정별 SEC의 상관(보고서 p.68)", "모니터링 전 기간 주간 공정별 SEC(2024.08~2026.08, p.95)",
    "출처: 단계보고서 p.68, p.95, p.366", "처리량 보정 질문 대응: 분기 SEC는 처리량에 따라 7.0~15.4로 변동합니다. 그래서 2단계 검증은 처리량 보정 SEC로 순수 운영 개선 효과를 분리해서 평가합니다(p.366).");
  two("백업 B-3 | 절감 시뮬레이션", "부하 패턴으로 절전 시간대를 정했고, 1차 산정은 보수적입니다",
    "p364_load_heatmap.jpg", "p365_standby_simulation.jpg", "요일·시간대별 충전소 전력 패턴(p.364)", "메인 칠러 대기전력 절감 시뮬레이션(추정, p.365)",
    "출처: 단계보고서 p.364~365 · 시뮬레이션은 추정값", "1단계 실적 질문 대응: 충전소 단위 절감량 공인검증은 2단계 1차년도로 이월했습니다. 기준선과 병목은 확정했고, 1차 시뮬레이션은 메인 칠러 야간만 대상으로 한 보수적 산정(충전소 약 2.7%)이며, 하드웨어 교체와 운영 최적화를 통합해 20%를 추진합니다.");
}

// ---------------------------------------------------------------- 현장평가 전용
async function fOverview(pres, C) {
  const s = contentSlide(pres, "과제 개요", "과제 개요", "고장·전력·실적 문제를 '부품 → 설비 → 충전소'로 풉니다");
  const probs = [["MdBuild", "잦은 고장", "다이어프램 교체주기 117 hr"], ["MdElectricBolt", "높은 전력", "충전소 SEC 약 9.5 kWh/kg"], ["MdAssignment", "실적 부족", "현장 운영실적 부족"]];
  const w = (CW - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (w + 0.2);
    card(s, pres, C, x, 1.6, w, 1.45);
    await iconCircle(s, pres, C, probs[i][0], x + 0.25, 1.85, 0.6);
    s.addText([{ text: probs[i][1], options: { bold: true, color: C.text2, fontSize: 17, breakLine: true } }, { text: probs[i][2], options: { fontSize: 14, color: C.text1 } }],
      { x: x + 1.05, y: 1.65, w: w - 1.2, h: 1.35, valign: "middle", margin: 0, isTextBox: true });
  }
  const chain = [["부품", "피스톤 씰 · 다이어프램 · IR 노즐 · 밸브 · PCHE", "6E93BF"], ["설비", "압축기 · 충전기 · 냉각기", HEX.accent1], ["충전소", "진단·실증·운영모델", HEX.dk2]];
  const cw = (CW - 1.0) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (cw + 0.5);
    s.addText([{ text: chain[i][0], options: { bold: true, fontSize: 20, breakLine: true } }, { text: chain[i][1], options: { fontSize: 14 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, x, y: 3.35, w: cw, h: 1.45, rectRadius: 0.08, fill: { color: chain[i][2] }, color: C.background1, valign: "middle", margin: 12, isTextBox: true });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + cw + 0.1, y: 3.9, w: 0.3, h: 0.35, fill: { color: C.accent2 } });
  }
  s.addText([{ text: "최종목표  ", options: { bold: true, color: C.accent2 } }, { text: "내구성 1.5배 · 에너지 20% 절감 · K-충전소 운영모델     ", options: { color: C.text2, bold: true } }, { text: "1단계 개발 → 2단계 현장 검증", options: { color: C.accent4 } }],
    { x: MX, y: 5.1, w: CW, h: 0.55, fontSize: 16, margin: 0, valign: "middle", isTextBox: true });
  s.addText("주관 고등기술연구원 · 공동 11개 기관(대학·연구기관·부품·설비 기업·운영사) · 지원 충주시, 충북도", { x: MX, y: 5.75, w: CW, h: 0.45, fontSize: 14, color: C.accent4, margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.2~5, p.26");
  s.addNotes("고장이 잦고 전력을 많이 쓰며 운영실적이 부족한 국산 충전소 설비 문제를, 부품 개선에서 설비, 충전소 운영모델까지 이어서 풉니다. 1단계는 개발, 2단계는 현장 검증입니다.");
}

async function fDurability(pres, C) {
  const s = contentSlide(pres, "기술성과", "기술성과 | 내구", "핵심부품 4종 1.5배 이상 달성, 일부는 현장 효과도 확인");
  badges(s, pres, C, ["공인", "현장"]);
  const H = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
  const rows = [
    [H("부품"), H("개선 전 → 개선 후"), H("개선율"), H("검증")],
    ["피스톤 씰(GTC)", "300 → 1,000 hr", "3.3배", "공인"],
    ["다이어프램(범한산업)", "현장 교체주기 117 → 492.7 hr", "4.21배", "공인 · 현장"],
    ["IR 리시버(MSENG)", "10,000 → 15,000 cycle", "1.5배", "공인"],
    ["유량조절밸브(MT.H)", "작동시험 153,030회 · 실충전 1,672회", "1.5배", "공인 · 현장"],
    ["고효율 칠러(삼정이엔씨)", "0.7829 → 0.6155 kWh/kg", "−21.4%", "공인"],
  ].map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: { color: k === 2 ? C.accent1 : C.text1, bold: k === 2 || k === 0 } })));
  s.addTable(rows, { x: MX, y: 1.6, w: 8.3, colW: [2.6, 3.4, 1.0, 1.3], rowH: 0.62, fontSize: 14, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.12, 0, 0.12], fill: { color: HEX.lt1 } });
  const rx = MX + 8.6, rw = CW - 8.6;
  const ims = ["p155_piston_parts.jpg", "p194_ir_receiver.jpg"];
  const i1 = fit(ims[0], rx, 1.6, rw / 2 - 0.1, 2.2), i2 = fit(ims[1], rx + rw / 2 + 0.1, 1.6, rw / 2 - 0.1, 2.2);
  s.addImage({ path: FIG(ims[0]), ...i1, altText: "개선 피스톤" });
  s.addImage({ path: FIG(ims[1]), ...i2, altText: "IR 리시버" });
  const i3 = fit("p173_diaphragm_compressor_prototype.jpg", rx, 4.0, rw, 2.0);
  s.addImage({ path: FIG("p173_diaphragm_compressor_prototype.jpg"), ...i3, altText: "다이어프램 압축기 시제품" });
  s.addText("실물은 견학 4번 지점에서 확인하실 수 있습니다", { x: MX, y: 5.6, w: 8.3, h: 0.45, fontSize: 15, bold: true, color: C.accent3, margin: 0, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.19~39 공인인증평가 성적서, p.26, p.319");
  s.addNotes("고장이 잦던 부품 네 가지는 공인시험에서 1.5배 이상을 달성했고, 다이어프램과 밸브는 실제 충전소에서도 효과를 확인했습니다. 실물은 잠시 뒤 견학에서 보여드리겠습니다.");
}

async function fEfficiency(pres, C) {
  const s = contentSlide(pres, "기술성과", "기술성과 | 효율", "기준선을 확정하고 병목을 찾았으며, 두 축의 절감 수단을 확보했습니다");
  badges(s, pres, C, ["공인"]);
  const lw = 6.9;
  const im = fit("p095_main_chiller_standby.jpg", MX, 1.6, lw, 2.75, "left");
  s.addImage({ path: FIG("p095_main_chiller_standby.jpg"), ...im, altText: "메인 칠러 대기전력" });
  const st = [["9.5", "kWh/kg 공인 기준선(KTC)"], ["38.7%", "칠러 전력량 중 야간 대기"], ["−21.4%", "고효율 칠러 kg당(KCL)"]];
  const sw = (lw - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (sw + 0.2);
    card(s, pres, C, x, 4.55, sw, 1.9);
    s.addText(st[i][0], { x: x + 0.2, y: 4.65, w: sw - 0.4, h: 0.75, fontSize: 28, bold: true, color: i === 1 ? C.accent2 : C.accent1, margin: 0, valign: "middle", isTextBox: true });
    s.addText(st[i][1], { x: x + 0.2, y: 5.45, w: sw - 0.4, h: 0.85, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  }
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  const im2 = fit("fig21_two_track_strategy.png", rx, 1.6, rw, 2.9);
  s.addImage({ path: FIG("fig21_two_track_strategy.png"), ...im2, altText: "하드웨어·운영 투트랙 전략" });
  card(s, pres, C, rx, 4.55, rw, 1.9, { fill: C.background1, line: HEX.accent5 });
  s.addText([{ text: "2단계 검증", options: { bold: true, color: C.text2, breakLine: true } }, { text: "적용 전후 비교 → 처리량 보정 SEC → 공인기관 입회", options: { breakLine: true } }, { text: "2027년 안에 충전소 단위 절감량 검증", options: { color: C.accent1, bold: true } }],
    { x: rx + 0.25, y: 4.65, w: rw - 0.5, h: 1.7, fontSize: 14, color: C.text1, margin: 0, valign: "middle", paraSpaceAfter: 4, isTextBox: true });
  source(s, C, "출처: 단계보고서 p.12, p.36, p.95, p.314, p.363~366");
  s.addNotes("공인기관과 함께 기준선을 9.5로 확정했고, 밤에도 칠러가 사전 냉각을 유지하는 대기전력이 병목이라는 것을 찾았습니다. 고효율 칠러와 절전 운전 로직, 두 축으로 줄이고 2단계에서 공인기관 입회로 검증합니다. 잠시 뒤 견학 2번 지점에서 이 칠러를 보여드립니다.");
}

async function fPerfGap(pres, C) {
  const s = contentSlide(pres, "성능지표·미달", "성능지표와 미달 항목", "부품 지표는 공인 달성, 미달 항목은 2단계 1차년도에 만회합니다");
  const sum = [["6", "달성", C.accent1], ["2", "진행", C.accent4], ["2", "미달", C.accent2]];
  for (let i = 0; i < 3; i++) {
    const y = 1.6 + i * 1.6;
    card(s, pres, C, MX, y, 3.3, 1.45);
    s.addText([{ text: sum[i][0], options: { fontSize: 40, bold: true, color: sum[i][2] } }, { text: "  " + sum[i][1], options: { fontSize: 18, bold: true, color: C.text2 } }],
      { x: MX + 0.3, y, w: 2.9, h: 1.45, valign: "middle", margin: 0, isTextBox: true });
  }
  s.addText("성능지표 10개 기준(보고서 p.292~293)", { x: MX, y: 6.4 - 0.15, w: 3.3, h: 0.3, fontSize: 11, color: C.accent4, margin: 0, isTextBox: true });
  const H = (t) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: C.background2 } } });
  const rows = [
    [H("미달 항목"), H("만회 계획")],
    ["특허 등록 4 / 10건", "심사 중 10건 대응, 2027년 안에 등록 전환"],
    ["기술실시 2 / 4회", "실증 결과와 특허를 묶어 기술 패키지 이전"],
    ["충전소 에너지 절감 검증", "2026.12 로직 완료 → 2027년 안에 공인검증"],
    ["다이어프램 압축기 전력", "고효율 벨트 적용 후 재검증"],
  ].map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: { color: C.text1, bold: k === 0 } })));
  s.addTable(rows, { x: MX + 3.6, y: 1.6, w: CW - 3.6, colW: [3.2, CW - 3.6 - 3.2], rowH: [0.5, 0.95, 0.95, 0.95, 0.95], fontSize: 15, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.15, 0, 0.15], fill: { color: HEX.lt1 } });
  source(s, C, "출처: 단계보고서 p.291~293, 4장 p.312~314");
  s.addNotes("부품 내구성 지표는 모두 공인시험으로 달성했습니다. 미달 항목은 네 가지이고, 공통 원인은 현장 실증 전 단계라는 점입니다. 모두 2단계 1차년도에 만회 일정을 잡았습니다.");
}

async function fNext(pres, C) {
  const s = contentSlide(pres, "2단계 계획", "2단계 계획", "2단계는 충전소에서 증명하는 단계입니다");
  const yrs = [["2027", "실증 설비 설치·시운전\n에너지 절감 공인검증\nHyEReD 개발"], ["2028", "장기 실증\nHyEReD 유효성 검증\n조기경보 운영절차서"], ["2029.1Q", "K-충전소 운영모델\n공인 성적서·설계표준안"]];
  const w = (CW - 1.0) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (w + 0.5);
    card(s, pres, C, x, 1.6, w, 2.45);
    s.addText(yrs[i][0], { x: x + 0.3, y: 1.72, w: w - 0.6, h: 0.65, fontSize: 28, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    s.addText(yrs[i][1], { x: x + 0.3, y: 2.45, w: w - 0.6, h: 1.5, fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + w + 0.1, y: 2.65, w: 0.3, h: 0.35, fill: { color: C.accent2 } });
  }
  s.addText([{ text: "리스크 대응  ", options: { bold: true, color: C.accent3 } }, { text: "공사 조기 착공 · 고압 압축기 제작 완료 · 충주 인허가 완료 · 세종 실증 1년 앞당김", options: { color: C.text1 } }],
    { x: MX, y: 4.25, w: CW, h: 0.5, fontSize: 15, margin: 0, valign: "middle", isTextBox: true });
  s.addText("2단계 약속", { x: MX, y: 4.95, w: 3, h: 0.4, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
  const pr = ["설비 패키지 내구성 1.5배\n현장 검증", "충전소 에너지 20% 절감\n공인검증", "K-충전소 운영모델\n운영절차서"];
  const pw = (CW - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    s.addText(pr[i], { shape: pres.shapes.ROUNDED_RECTANGLE, x: MX + i * (pw + 0.2), y: 5.4, w: pw, h: 1.05, rectRadius: 0.08, fill: { color: C.text2 }, fontSize: 15, bold: true, color: C.background1, valign: "middle", margin: 12, isTextBox: true });
  }
  source(s, C, "출처: 단계보고서 p.4~5, p.13, p.93, p.317, 9장 p.357~366");
  s.addNotes("2027년에 실증 설비를 설치하고 에너지 절감량을 공인검증하며, 2028년 장기 실증을 거쳐 2029년 K-충전소 운영모델로 정리합니다. 설치 일정 리스크는 공사를 앞당겨 착공해 줄였습니다.");
}

async function fTour(pres, C) {
  const s = contentSlide(pres, "현장 확인", "현장 확인 안내", "이제 현장에서 직접 보여드리겠습니다");
  const im = fit("p049_process_flow.jpg", MX, 1.55, CW, 2.75);
  s.addImage({ path: FIG("p049_process_flow.jpg"), ...im, altText: "충주바이오그린수소충전소 공정 흐름도" });
  const stops = [["1", "모니터링 화면", "실시간 전력 15채널 · 진동 7개 측정점 · 조기경보"], ["2", "메인 칠러·압축기", "야간 대기전력 병목, 2단계 교체 대상"], ["3", "설비동 공사 현장", "기초 타설 완료, 개선 설비 반입 예정 위치"], ["4", "개선 부품 실물", "피스톤 씰 · 다이어프램 · IR 리시버 · 밸브와 성적서"]];
  const w = (CW - 0.6) / 4;
  for (let i = 0; i < 4; i++) {
    const x = MX + i * (w + 0.2);
    card(s, pres, C, x, 4.5, w, 1.95);
    s.addText(stops[i][0], { shape: pres.shapes.OVAL, x: x + 0.2, y: 4.65, w: 0.5, h: 0.5, fill: { color: C.accent3 }, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(stops[i][1], { x: x + 0.85, y: 4.65, w: w - 1.0, h: 0.5, fontSize: 16, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
    s.addText(stops[i][2], { x: x + 0.2, y: 5.3, w: w - 0.4, h: 1.05, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  }
  source(s, C, "공정 흐름도: 단계보고서 p.49 · 견학 동선은 현장 여건에 맞춰 조정 [확인 필요]");
  s.addNotes("이제 현장으로 이동하겠습니다. 모니터링 화면, 메인 칠러와 압축기, 설비동 공사 현장, 개선 부품 실물 순서로 보여드리겠습니다.");
}

// ---------------------------------------------------------------- 빌드
async function buildMain() {
  const { pres, C } = makePres("본평가판");
  pres.addSection({ title: "표지" }); await sCover(pres, C, "main");
  pres.addSection({ title: "결론" }); await sConclusion(pres, C, "결론");
  pres.addSection({ title: "과제 개요" }); await sProblem(pres, C); await sGoal(pres, C); await sOrg(pres, C);
  pres.addSection({ title: "1단계 성과" }); await sTimeline(pres, C); await sDashboard(pres, C, "1단계 성과");
  pres.addSection({ title: "기술성과" }); await sPiston(pres, C); await sDiaphragm(pres, C); await sDispenser(pres, C); await sChiller(pres, C);
  await sDiag(pres, C, "기술성과"); await sEff1(pres, C); await sEff2(pres, C, "기술성과");
  pres.addSection({ title: "실증 인프라" }); await sInfra(pres, C, "실증 인프라");
  pres.addSection({ title: "성능지표·미달" }); await sPerf(pres, C); await sGap(pres, C);
  pres.addSection({ title: "2단계 계획" }); await sNext(pres, C, "2단계 계획"); await sRisk(pres, C); await sBiz(pres, C);
  pres.addSection({ title: "맺음말" }); await sClosing(pres, C);
  pres.addSection({ title: "백업" }); await backups(pres, C);
  const f = path.join(OUT, "1단계_단계평가_본평가판.pptx");
  await pres.writeFile({ fileName: f }); await applyTheme(f, THEME); return f;
}
async function buildField() {
  const { pres, C } = makePres("현장평가판");
  pres.addSection({ title: "표지" }); await sCover(pres, C, "field");
  pres.addSection({ title: "결론" }); await sConclusion(pres, C, "결론");
  pres.addSection({ title: "과제 개요" }); await fOverview(pres, C);
  pres.addSection({ title: "1단계 성과" }); await sDashboard(pres, C, "1단계 성과");
  pres.addSection({ title: "기술성과" }); await fDurability(pres, C); await sDiag(pres, C, "기술성과"); await fEfficiency(pres, C);
  pres.addSection({ title: "실증 인프라" }); await sInfra(pres, C, "실증 인프라");
  pres.addSection({ title: "성능지표·미달" }); await fPerfGap(pres, C);
  pres.addSection({ title: "2단계 계획" }); await fNext(pres, C);
  pres.addSection({ title: "맺음말" }); await sClosing(pres, C);
  pres.addSection({ title: "현장 확인" }); await fTour(pres, C);
  const f = path.join(OUT, "1단계_단계평가_현장평가판.pptx");
  await pres.writeFile({ fileName: f }); await applyTheme(f, THEME); return f;
}

(async () => {
  for (const n of ICONS) await icon(n);
  console.log(await buildMain());
  console.log(await buildField());
})().catch((e) => { console.error(e); process.exit(1); });
