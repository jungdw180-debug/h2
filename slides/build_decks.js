// 1단계 단계평가 발표자료 생성 스크립트 (본평가판 + 현장평가판)
// 양식: 고등기술연구원 발표 양식(번호 제목, 소제목 표시, 남색 머리띠 패널, ◆·▪ 개조식, 하단 강조 박스)
// 실행: NODE_PATH=<pptxgenjs·react-icons·sharp node_modules> APPLY_THEME=<apply_theme.js> node slides/build_decks.js
// 내용 근거: report/발표자료_슬라이드내용_v1.md (쪽수는 단계보고서 인쇄 쪽 번호)

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const { applyTheme } = require(process.env.APPLY_THEME);
const JSZip = require("jszip");

// pptxgenjs는 한 문단의 글자 조각마다 <a:pPr>를 반복해 넣는다. 문단당 첫 번째만 남겨 스키마에 맞춘다.
async function dedupeParagraphProps(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const names = Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(n));
  let fixed = 0;
  for (const n of names) {
    const xml = await zip.file(n).async("string");
    const out = xml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (para) => {
      let seen = false;
      return para.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, (m) => {
        if (!seen) { seen = true; return m; }
        fixed++; return "";
      });
    });
    if (out !== xml) zip.file(n, out);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return fixed;
}

const ROOT = path.join(__dirname, "..");
const FIG = (f) => (f.startsWith("fig") ? path.join(ROOT, "figures", f) : path.join(ROOT, "figures", "pdf", f));
const TPL = (f) => path.join(__dirname, "template", f);
const OUT = __dirname;

// ---------------------------------------------------------------- 테마 (참고 양식 색상)
const THEME = {
  name: "IAE Stage Review",
  headFontFace: "넥슨Lv2고딕",
  bodyFontFace: "넥슨Lv2고딕",
  colors: {
    dk1: "222222", lt1: "FFFFFF", dk2: "1C3B5F", lt2: "E7F2F8",
    accent1: "0A78BC", accent2: "0A2B52", accent3: "115985", accent4: "1981C1",
    accent5: "7A8796", accent6: "D9822B", hlink: "0A78BC", folHlink: "6B4E9B",
  },
};
const HEX = THEME.colors;
const W = 13.333, MX = 0.6, CW = 12.133;
const BORDER = "C9D3DD", CALLOUT_LINE = "B9D3E6", ROWBG = "F2F4F6";

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

// ---------------------------------------------------------------- 프레젠테이션과 레이아웃
function makePres(subject) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "수소충전소 핵심설비·부품 내구성 및 효율 향상 1단계 단계평가";
  pres.subject = subject;
  pres.author = "고등기술연구원";
  const C = pres.SchemeColor;
  const logoGray = { image: { path: TPL("logo_gray.png"), x: 11.0, y: 0.36, w: 1.9, h: 0.4 } };
  const logoBlue = { image: { path: TPL("logo_blue.png"), x: 11.0, y: 0.36, w: 1.9, h: 0.4 } };
  const swoosh = { image: { path: TPL("swoosh.png"), x: 0, y: 0, w: 4.0, h: 0.95 } };
  const pageNo = { x: 12.33, y: 7.08, w: 0.6, h: 0.3, fontSize: 12, color: C.text1, align: "right" };

  pres.defineSlideMaster({
    title: "COVER", background: { color: HEX.lt1 },
    objects: [
      { image: { path: TPL("logo_blue.png"), x: 0.45, y: 0.36, w: 1.9, h: 0.4 } }, logoGray,
      { placeholder: { options: { name: "eng", type: "body", x: 1.15, y: 1.65, w: 9.5, h: 0.6, fontSize: 12, color: C.accent4, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 1.0, y: 2.35, w: 10.5, h: 1.65, fontSize: 36, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "date", type: "body", x: 1.0, y: 4.35, w: 5.3, h: 0.5, fontSize: 22, color: C.text1, align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "who", type: "body", x: 6.3, y: 4.25, w: 6.6, h: 2.6, fontSize: 18, color: C.text1, valign: "top", align: "right", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT", background: { color: HEX.lt1 },
    objects: [
      swoosh, logoGray,
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.26, w: 10.2, h: 0.66, fontSize: 30, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    ],
    slideNumber: pageNo,
  });
  pres.defineSlideMaster({ title: "TOC", background: { color: HEX.lt1 }, objects: [swoosh, logoBlue] });
  pres.defineSlideMaster({
    title: "PLAIN", background: { color: HEX.lt1 },
    objects: [
      swoosh, logoGray,
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 2.6, w: CW, h: 1.6, fontSize: 60, bold: true, color: C.text1, valign: "middle", align: "center", margin: 0 }, text: "" } },
    ],
  });
  return { pres, C };
}

// ---------------------------------------------------------------- 공통 헬퍼
function contentSlide(pres, C, section, title, sub) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (sub) {
    s.addImage({ path: TPL("marker.png"), x: MX, y: 1.04, w: 0.14, h: 0.245, altText: "소제목 표시" });
    s.addText(sub, { x: MX + 0.22, y: 0.95, w: CW - 0.22, h: 0.44, fontSize: 19, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "subheading" });
  }
  return s;
}
function source(s, C, text) {
  s.addText(text, { x: MX, y: 7.1, w: 10.5, h: 0.26, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true, objectName: "source" });
}
// 하단 강조 박스: parts = [["일반 문구"], ["강조 문구", true], ...]
function callout(s, pres, C, parts, y = 6.42, h = 0.56) {
  s.addText(parts.map(([t, hl]) => ({ text: t, options: { color: hl ? C.accent1 : C.text1, bold: true } })), {
    shape: pres.shapes.RECTANGLE, x: MX, y, w: CW, h, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 },
    fontSize: 17, align: "center", valign: "middle", margin: 4, isTextBox: true, objectName: "callout",
  });
}
// 남색 머리띠 패널
function panel(s, pres, C, x, y, w, h, header, hfill) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background1 }, line: { color: BORDER, width: 0.75 }, objectName: "panel" });
  s.addText(header, {
    shape: pres.shapes.RECTANGLE, x, y, w, h: 0.46, fill: { color: hfill || C.text2 },
    fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "panel-header",
  });
}
// ◆·▪ 개조식 목록. items: [lv, text | [[t, hl], ...]]  lv 1=◆, 2=▪, 3=→
function bullets(C, items, size = 15) {
  const runs = [];
  items.forEach(([lv, content], idx) => {
    const segs = typeof content === "string" ? [[content, false]] : content;
    const last = idx === items.length - 1;
    segs.forEach(([t, hl], k) => {
      const o = {};
      if (k === 0) {
        if (lv === 1) Object.assign(o, { bullet: { code: "25C6", indent: 18 }, paraSpaceBefore: idx === 0 ? 0 : 9 });
        else if (lv === 2) Object.assign(o, { bullet: { code: "25AA", indent: 16 }, indentLevel: 1, paraSpaceBefore: 3 });
        else if (lv === 3) Object.assign(o, { bullet: { code: "2192", indent: 20 }, indentLevel: 1, paraSpaceBefore: 3 });
        else Object.assign(o, { paraSpaceBefore: idx === 0 ? 0 : 6 });
      }
      Object.assign(o, {
        fontSize: lv === 1 ? size + 1 : size,
        bold: lv === 1 || !!hl,
        color: hl ? C.accent1 : lv === 1 ? C.accent2 : C.text1,
      });
      if (k === segs.length - 1 && !last) o.breakLine = true;
      runs.push({ text: t, options: o });
    });
  });
  return runs;
}
function addBullets(s, C, items, x, y, w, h, size = 15) {
  s.addText(bullets(C, items, size), { x, y, w, h, valign: "top", margin: 0, isTextBox: true, objectName: "bullets" });
}
// 화살표형 단계 머리띠
function stepHeader(s, pres, C, x, y, w, text, fill, fs = 18) {
  s.addText(text, { shape: pres.shapes.PENTAGON, x, y, w, h: 0.55, fill: { color: fill }, fontSize: fs, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "step-header" });
}
// 남색 탭 + 체크 행
function checkRow(s, pres, C, y, h, tab, fill, text, tabW = 2.3) {
  s.addText(tab, { shape: pres.shapes.RECTANGLE, x: MX, y, w: tabW, h, fill: { color: fill }, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.RECTANGLE, { x: MX + tabW, y, w: CW - tabW, h, fill: { color: ROWBG }, line: { color: ROWBG, width: 0 } });
  s.addText(text, { x: MX + tabW + 0.25, y, w: CW - tabW - 0.4, h, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
}
function pill(s, pres, C, text, x, y, kind) {
  const solid = { 공인: C.accent1, 현장: C.accent3 }[kind];
  const lineC = kind === "추정" ? C.accent6 : C.accent5;
  s.addText(text, {
    shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w: 0.66, h: 0.28, rectRadius: 0.14,
    fill: { color: solid || C.background1 }, line: solid ? undefined : { color: lineC, width: 1 },
    fontSize: 11, bold: true, color: solid ? C.background1 : lineC, align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
}
async function iconCircle(s, pres, C, name, x, y, d = 0.56, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill || C.accent1 }, objectName: "icon-bg" });
  const pad = d * 0.2;
  s.addImage({ data: await icon(name), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: name });
}
const chartBase = () => ({
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, catAxisLabelColor: HEX.dk1, valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: HEX.dk1,
  showTitle: true, titleFontSize: 13, titleColor: HEX.dk1, barGapWidthPct: 60, catAxisLineShow: true, catAxisLineColor: "C9D3DE",
});

// ================================================================ 슬라이드
function sCover(pres, C, variant) {
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  s.addShape(pres.shapes.DIAMOND, { x: 0.95, y: 1.74, w: 0.12, h: 0.12, fill: { color: "7A2E5E" } });
  s.addText("IMPROVEMENT OF DURABILITY AND EFFICIENCY OF PRIMARY FACILITIES\nAND COMPONENTS IN HYDROGEN REFUELING STATION", { placeholder: "eng" });
  s.addText("수소충전소 핵심설비·부품\n내구성 및 효율 향상", { placeholder: "title" });
  s.addText(`1단계 ${variant === "field" ? "현장평가" : "단계평가"}  |  [일자]`, { placeholder: "date" });
  s.addText([
    { text: "[발표자]", options: { bold: true, fontSize: 22, breakLine: true } },
    { text: "주관연구개발기관  고등기술연구원", options: { fontSize: 18, color: C.accent5, breakLine: true } },
    { text: "공동연구개발기관  11개 기관 · 지원 충주시, 충북도", options: { fontSize: 16, color: C.accent5, breakLine: true } },
    { text: "RS-2024-00419764 · 기후에너지환경부 / 한국에너지기술평가원", options: { fontSize: 14, color: C.accent5 } },
  ], { placeholder: "who" });
}

function sToc(pres, C, items) {
  const s = pres.addSlide({ masterName: "TOC", sectionTitle: "목차" });
  s.addShape(pres.shapes.PARALLELOGRAM, { x: 1.95, y: 1.6, w: 1.25, h: 0.82, fill: { color: "E3EAF0" }, objectName: "toc-shadow" });
  s.addShape(pres.shapes.PARALLELOGRAM, { x: 1.8, y: 1.45, w: 1.25, h: 0.82, fill: { color: C.accent2 }, objectName: "toc-mark" });
  s.addText("목차", { x: 3.35, y: 1.38, w: 3, h: 0.95, fontSize: 48, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  s.addText("IMPROVEMENT OF DURABILITY AND EFFICIENCY OF PRIMARY FACILITIES\nAND COMPONENTS IN HYDROGEN REFUELING STATION",
    { x: 7.2, y: 1.6, w: 5.6, h: 0.5, fontSize: 9, color: "A0A8B0", align: "right", margin: 0, isTextBox: true });
  const gap = items.length > 5 ? 0.62 : 0.7;
  items.forEach((t, i) => {
    const y = 2.75 + i * gap;
    s.addText(String(i + 1), { shape: pres.shapes.DIAMOND, x: 3.4, y, w: 0.5, h: 0.5, fill: { color: C.background2 }, line: { color: "9CC3DF", width: 1 }, fontSize: 16, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: 4.1, y, w: 7, h: 0.5, fontSize: 24, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addNotes("발표 순서입니다.");
}

// 1. 과제 개요 -----------------------------------------------------
function sBackground(pres, C, t) {
  const s = contentSlide(pres, C, "과제 개요", t, "국산 수소충전소 핵심설비의 내구성·효율 확보 필요");
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w, 4.7, "국내 수소충전소 핵심설비 현황");
  addBullets(s, C, [
    [1, "잦은 고장과 부품 교체"], [2, "다이어프램 평균 교체 운전시간 약 117 hr(기존품, 현장)"], [2, "릴리프밸브 사용 수명 약 3개월"],
    [1, "높은 에너지 사용량"], [2, [["충전소 SEC 약 ", false], ["9.5 kWh/kgH₂", true]]], [2, "압축기·냉각기 등 전력 다소비 설비에 부하 집중"],
    [1, "운영실적(Track Record) 부족"], [2, "효율·내구성 검증 및 현장 운영실적 미흡"], [2, "중소기업 중심 공급 구조 → 개발·검증 인프라 필요"],
  ], MX + 0.3, 2.2, w - 0.5, 3.95);
  s.addShape(pres.shapes.LINE, { x: MX + w + 0.2, y: 1.6, w: 0, h: 4.6, line: { color: "C9D3DD", width: 1, dashType: "dash" } });
  const x2 = MX + w + 0.4;
  panel(s, pres, C, x2, 1.55, w, 4.7, "과제 추진 방향");
  addBullets(s, C, [
    [1, "핵심부품 고내구화·국산화"], [2, "압축기·충전기 소모성 부품 내구성 1.5배 이상 향상"],
    [1, "에너지 사용량 20% 이상 절감"], [2, "압축기·냉각기 효율 향상 및 운전 최적화"],
    [1, "데이터 기반 신뢰성 관리"], [2, "고장 후 정비 → 상태기반 예측정비(CBM) 전환"],
    [1, "글로벌 기술·표준 선점 대응"], [2, "美·日 수소전략 개정(’23.6), EU AFIR 등 경쟁 가속"],
  ], x2 + 0.3, 2.2, w - 0.5, 3.95);
  callout(s, pres, C, [["핵심설비 내구성·효율 향상을 통한 ", false], ["충전소 운영 경제성(OPEX) 확보", true], [" 시급", false]]);
  source(s, C, "출처: 단계보고서 p.2, p.26");
  s.addNotes("국내 충전소 핵심설비는 주로 중소기업이 공급합니다. 그런데 부품 교체가 잦고, 수소 1 kg에 약 9.5 kWh를 쓰며, 현장 운영 실적이 부족해 신뢰를 얻지 못하고 있습니다. 그래서 본 과제는 핵심부품 고내구화, 에너지 20% 절감, 데이터 기반 신뢰성 관리를 목표로 합니다.");
}

async function sGoal(pres, C, t) {
  const s = contentSlide(pres, C, "과제 개요", t, "최종목표 및 단계별 추진 구조");
  const goals = [["MdPrecisionManufacturing", "내구성 1.5배 향상"], ["MdElectricBolt", "에너지 20% 이상 절감"], ["MdHub", "K-수소충전소 운영모델"]];
  const gw = (CW - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (gw + 0.2);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w: gw, h: 0.85, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 } });
    await iconCircle(s, pres, C, goals[i][0], x + 0.2, 1.74, 0.56, C.accent1);
    s.addText(goals[i][1], { x: x + 0.9, y: 1.6, w: gw - 1.0, h: 0.85, fontSize: 18, bold: true, color: C.accent2, valign: "middle", margin: 0, isTextBox: true });
  }
  const sw = (CW - 0.4) / 2;
  const st = [
    ["1단계 (2024.04~2026.12)", C.text2, "개발·체계 구축", [[2, "핵심부품 개선, 시제품 제작 및 공인 평가"], [2, "신뢰성 관리모드(FMECA·조기경보) 구축"], [2, "실증 사이트 설계 및 인허가"]]],
    ["2단계 (2027.01~2029.03)", C.accent4, "충전소 현장 검증", [[2, "개선 설비 충전소 설치 및 장기 실증"], [2, "에너지 절감량 공인기관 검증"], [2, "HyEReD DB 및 K-충전소 운영모델"]]],
  ];
  for (let i = 0; i < 2; i++) {
    const x = MX + i * (sw + 0.4);
    stepHeader(s, pres, C, x, 2.75, sw, st[i][0], st[i][1]);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.4, w: sw, h: 2.8, fill: { color: C.background1 }, line: { color: BORDER, width: 0.75 } });
    s.addText(st[i][2], { x: x + 0.3, y: 3.55, w: sw - 0.6, h: 0.5, fontSize: 19, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    addBullets(s, C, st[i][3], x + 0.3, 4.15, sw - 0.6, 1.9, 16);
  }
  callout(s, pres, C, [["TRL 7 → 8  ·  1단계 개발 결과를 ", false], ["2단계 충전소 실증으로 검증", true], ["하는 구조", false]]);
  source(s, C, "출처: 단계보고서 요약문 p.3, p.5");
  s.addNotes("과제는 단계 구조로 설계되어 있습니다. 1단계에서는 부품을 개선하고 관리체계를 만들며, 2단계에서는 이를 충전소에 설치해 내구성과 에너지 절감을 검증합니다.");
}

function sOrg(pres, C, t) {
  const s = contentSlide(pres, C, "과제 개요", t, "부품–설비–충전소 연계 추진체계 (12개 기관)");
  const rows = [
    ["충전소 운영모델", C.accent2, "✓  고등기술연구원(총괄·진단·실증) · 고려대(경제성·탄소) · 창원산업진흥원, 코하이젠(현장 실증)"],
    ["설비", C.accent3, "✓  압축기: GTC, 범한산업 · 충전기: MSENG, 제아ENG · 냉각기: 삼정이엔씨"],
    ["부품", C.accent4, "✓  피스톤 씰: GTC, 한국자동차연구원 · 다이어프램: 범한산업, 한국자동차연구원\n✓  IR 노즐: MSENG · 유량조절밸브: MT.H콘트롤밸브 · PCHE: 에너진"],
  ];
  const hs = [0.75, 0.75, 1.0];
  let y = 1.6;
  rows.forEach(([tab, col, txt], i) => { checkRow(s, pres, C, y, hs[i], tab, col, txt); y += hs[i] + 0.12; });
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 4.45, w, 1.8, "데이터 환류 사례 ① 다이어프램");
  addBullets(s, C, [[2, "창원산업진흥원 현장 교체 데이터 분석"], [2, "범한산업 후속 설계 개선에 반영(p.26)"]], MX + 0.3, 5.05, w - 0.5, 1.1);
  panel(s, pres, C, MX + w + 0.4, 4.45, w, 1.8, "데이터 환류 사례 ② 조기경보");
  addBullets(s, C, [[2, "코하이젠 운전·고장이력 데이터 수집"], [2, "주관기관 진단 알고리즘 개발에 활용(p.28)"]], MX + w + 0.7, 5.05, w - 0.5, 1.1);
  callout(s, pres, C, [["현장 데이터 → 고장모드 정량화 → 개선 부품 → 현장 재확인의 ", false], ["순환 구조", true]]);
  source(s, C, "출처: 단계보고서 표지, p.26, p.28 · 지원: 충주시, 충북도");
  s.addNotes("현장 운영기관이 고장 데이터를 내면, 주관기관이 고장모드를 정량화해 개선 대상을 정하고, 부품 기업이 개선품을 만들어 다시 현장에서 확인합니다. 다이어프램은 이 순환을 실제로 한 바퀴 돈 사례입니다.");
}

function fOverview(pres, C, t) {
  const s = contentSlide(pres, C, "과제 개요", t, "국산 수소충전소 핵심설비의 내구성·효율 확보 및 현장 검증");
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w, 4.7, "현황 및 필요성");
  addBullets(s, C, [
    [1, "잦은 고장"], [2, "다이어프램 교체주기 약 117 hr, 릴리프밸브 수명 약 3개월"],
    [1, "높은 에너지 사용량"], [2, [["충전소 SEC 약 ", false], ["9.5 kWh/kgH₂", true]]],
    [1, "운영실적 부족"], [2, "효율·내구성 검증 및 현장 운영실적 미흡"],
  ], MX + 0.3, 2.2, w - 0.5, 3.9);
  const x2 = MX + w + 0.4;
  panel(s, pres, C, x2, 1.55, w, 4.7, "최종목표 및 추진체계");
  addBullets(s, C, [
    [1, "최종목표"], [2, "핵심부품·설비 내구성 1.5배"], [2, "에너지 사용량 20% 이상 절감"], [2, "K-수소충전소 운영모델 제시"],
    [1, "단계 구조"], [2, "1단계 개발·체계 구축 → 2단계 충전소 현장 검증"],
    [1, "추진체계"], [2, "부품 → 설비 → 충전소 운영모델 연계(12개 기관)"],
  ], x2 + 0.3, 2.2, w - 0.5, 3.9);
  callout(s, pres, C, [["주관 고등기술연구원 · 공동 11개 기관 · 지원 충주시, 충북도  →  ", false], ["부품–설비–충전소 연계 추진", true]]);
  source(s, C, "출처: 단계보고서 p.2~5, p.26");
  s.addNotes("고장이 잦고 전력을 많이 쓰며 운영실적이 부족한 국산 충전소 설비 문제를, 부품 개선에서 설비, 충전소 운영모델까지 이어서 풉니다. 1단계는 개발, 2단계는 현장 검증입니다.");
}

// 2. 1단계 추진 실적 -------------------------------------------------
async function sSummary(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 추진 실적", t, "1단계 성과 요약");
  const cards = [
    ["MdPrecisionManufacturing", "내구", "1.5배+", "", [[2, "핵심부품 4종 공인"], [2, "피스톤 씰 330%"]], "공인"],
    ["MdMonitorHeart", "진단", "92일 전", "", [[2, "중정비 고장 사전 감지"], [2, "고장모드 148개 정량화"]], "현장"],
    ["MdElectricBolt", "효율", "9.5", " kWh/kg", [[2, "공인 기준선 확정"], [2, "칠러 kg당 −21.4%"]], "공인"],
    ["MdConstruction", "실증", "착공", "", [[2, "인허가 3건 완료"], [2, "2단계 공사 조기 착수"]], "현장"],
  ];
  const w = (CW - 0.6) / 4;
  for (let i = 0; i < 4; i++) {
    const [ic, lab, big, unit, items, b] = cards[i];
    const x = MX + i * (w + 0.2);
    panel(s, pres, C, x, 1.55, w, 4.65, lab);
    await iconCircle(s, pres, C, ic, x + w / 2 - 0.35, 2.2, 0.7, C.accent1);
    s.addText([{ text: big, options: { fontSize: 34, bold: true, color: C.accent1 } }, { text: unit, options: { fontSize: 15, color: C.accent1 } }],
      { x: x + 0.15, y: 3.05, w: w - 0.3, h: 0.8, align: "center", valign: "middle", margin: 0, isTextBox: true });
    addBullets(s, C, items, x + 0.25, 4.0, w - 0.35, 1.45, 15);
    pill(s, pres, C, b, x + w / 2 - 0.33, 5.65, b);
  }
  callout(s, pres, C, [["부품 내구성 공인 입증 및 실증 인프라 조기 착공  →  ", false], ["2단계 충전소 현장 검증 준비 완료", true]]);
  source(s, C, "출처: 단계보고서 p.12, p.13, p.22, p.36");
  s.addNotes("결론부터 말씀드리겠습니다. 1단계에서는 고장이 잦은 핵심부품 네 가지를 개선해 공인시험에서 내구성 1.5배 이상을 확인했습니다. 피스톤 씰은 3.3배입니다. 운전데이터로 실제 고장을 석 달 전에 감지하는 체계를 만들었고, 에너지는 공인 기준선을 확정하고 고효율 칠러를 확보했습니다. 그리고 이 개선 설비들을 설치할 실증 설비동은 2단계 공사를 앞당겨 이미 착공했습니다.");
}

function sTimeline(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 추진 실적", t, "연차별 추진 경과: 개발 항목 전 항목 완료");
  const years = [
    ["2024 · 1차년도", C.text2, "기준 수립", [[2, "회전설비 FMEA 수립(p.10)"], [2, "피스톤 씰 기준치 300 hr 공인(p.20)"], [2, "유량조절밸브 KS 인증(2024.08, p.37)"], [2, "실증 사이트 선정(p.25, p.27)"]]],
    ["2025 · 2차년도", C.accent3, "시제품 및 공인 기준", [[2, "에너지 기준선 9.5 KTC 성적서(p.12)"], [2, "피스톤 씰 450 hr, 150% 공인(p.21)"], [2, "이상감지 모델 2종 개발(p.12)"], [2, "다이어프램 개선품 현장 설치(p.26)"]]],
    ["2026 · 3차년도", C.accent4, "공인 목표 달성", [[2, "피스톤 씰 1,000 hr, 330% 공인(p.22)"], [2, "IR 리시버 15,000 cycle 공인(p.30)"], [2, "고장모드 148개 RPN, 매뉴얼 3종(p.13)"], [2, "인허가 완료, 설비동 기초 타설(p.13)"]]],
  ];
  const w = (CW - 0.4) / 3;
  for (let i = 0; i < 3; i++) {
    const x = MX + i * (w + 0.2);
    stepHeader(s, pres, C, x, 1.6, w, years[i][0], years[i][1]);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.25, w, h: 3.95, fill: { color: C.background1 }, line: { color: BORDER, width: 0.75 } });
    s.addText(years[i][2], { x: x + 0.2, y: 2.4, w: w - 0.4, h: 0.5, fontSize: 19, bold: true, color: C.accent1, align: "center", margin: 0, isTextBox: true });
    addBullets(s, C, years[i][3], x + 0.3, 3.05, w - 0.45, 3.0, 15);
  }
  callout(s, pres, C, [["1차년도 기준 수립 → 2차년도 시제품·공인 기준 → 3차년도 ", false], ["공인 목표 달성", true]]);
  source(s, C, "출처: 단계보고서 2-2절 기관별 추진 실적(p.10~39)");
  s.addNotes("1차년도에 기준을 세우고, 2차년도에 시제품을 만들고 공인 기준을 받았으며, 3차년도에 공인시험으로 목표를 확인했습니다. 연차별 개발 항목은 모두 완료했습니다.");
}

function sDashboard(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 추진 실적", t, "정량 성과: 12개 지표 중 10개 달성");
  const d = [
    ["SCIE 논문", "14", "5", "건"], ["국내 논문", "8", "8", "건"], ["SW 저작권", "28", "5", "건"], ["사업화", "7", "3", "건"],
    ["사업화 매출", "3,280", "2,000", "백만 원"], ["고용 창출", "35", "16", "명"], ["현장시험", "14", "10", "건"], ["시제품", "7", "5", "건"],
    ["MOU", "12", "2", "건"], ["국가인증", "2", "1", "건"], ["특허 등록", "4", "10", "건", "출원 14건"], ["기술실시", "2", "4", "회", ""],
  ];
  const cols = 6, g = 0.16, w = (CW - g * (cols - 1)) / cols, h = 2.18;
  for (let i = 0; i < d.length; i++) {
    const [lab, a, tg, u, note] = d[i];
    const miss = note !== undefined;
    const x = MX + (i % cols) * (w + g), y = 1.6 + Math.floor(i / cols) * (h + g);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: miss ? C.background1 : C.background2 }, line: { color: miss ? HEX.accent6 : CALLOUT_LINE, width: miss ? 1.25 : 0.75 } });
    s.addText(lab, { shape: pres.shapes.RECTANGLE, x, y, w, h: 0.42, fill: { color: miss ? C.accent6 : C.text2 }, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(a, { x: x + 0.1, y: y + 0.55, w: w - 0.2, h: 0.85, fontSize: a.length > 3 ? 30 : 36, bold: true, color: miss ? C.accent6 : C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(`목표 ${tg} ${u}`, { x: x + 0.1, y: y + 1.45, w: w - 0.2, h: 0.3, fontSize: 13, color: C.accent5, align: "center", margin: 0, isTextBox: true });
    if (miss) s.addText(note ? `미달 · ${note}` : "미달", { x: x + 0.1, y: y + 1.78, w: w - 0.2, h: 0.28, fontSize: 12, bold: true, color: C.accent6, align: "center", margin: 0, isTextBox: true });
  }
  callout(s, pres, C, [["SCIE 14건(IF 10 이상 4건) · SW 28건 · 사업화 매출 3,280백만 원 등 ", false], ["10개 지표 목표 초과", true]]);
  source(s, C, "출처: 단계보고서 p.291 정량적 연구개발성과표(2026.10.30 제출본), p.6");
  s.addNotes("정량지표 12개 중 10개를 달성했습니다. 특허 등록과 기술이전 두 가지는 미달인데, 뒤에서 원인과 보완 계획을 따로 말씀드리겠습니다. (발표 전 요약문·4장 수치를 성과표 값으로 통일할 것)");
}

// 3. 1단계 연구 결과 -------------------------------------------------
function sPiston(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[내구] 피스톤 압축기: 씰 내구성 300 → 1,000 hr (기준 대비 330%)");
  s.addChart(pres.charts.BAR, [{ name: "내구시험 시간", labels: ["기준치(2024)", "2025", "2026"], values: [300, 450, 1000] }], {
    x: MX, y: 1.55, w: 5.2, h: 4.7, barDir: "col", chartColors: ["9AA5B1", "8DB9DC", HEX.accent1], dataLabelFormatCode: "#,##0",
    title: "피스톤 씰 공인 내구시험 시간 (hr)", ...chartBase(), valAxisMinVal: 0, valAxisMaxVal: 1150,
  });
  panel(s, pres, C, 6.0, 1.55, 4.15, 4.7, "주요 내용 (GTC · 한국자동차연구원)");
  addBullets(s, C, [
    [1, "개선 내용"], [2, "A7075 피스톤, 이중 스프링씰 적용"], [2, "보조 Energized Seal 구조 특허 출원"],
    [1, "공인 내구시험"], [2, [["300 hr → 450 hr → ", false], ["1,000 hr", true]]], [2, "450 hr 연속 내구시험 성적서(한국자동차연구원)"],
    [1, "후속 조치"], [2, "개선 씰 적용 신규 압축기 제작 완료"], [2, "2단계 설비동 설치·실증 예정"],
  ], 6.25, 2.15, 3.75, 4.0, 14);
  const im = fit("p155_piston_parts.jpg", 10.35, 1.55, 2.38, 4.3);
  s.addImage({ path: FIG("p155_piston_parts.jpg"), ...im, altText: "개선 피스톤 부품" });
  s.addText("[개선 피스톤 부품]", { x: 10.35, y: im.y + im.h + 0.05, w: 2.38, h: 0.28, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  callout(s, pres, C, [["1단계 목표(150%) 대비 ", false], ["2배 이상 초과 달성", true], [", 연차별 공인기관 성적서 확보", false]]);
  source(s, C, "출처: 단계보고서 p.19~22 공인인증평가 성적서(2024·2025·2026)");
  s.addNotes("피스톤 씰은 수명이 짧아 수소 보충과 정비가 잦던 부품입니다. 구조를 바꾼 결과 공인시험에서 기준의 3.3배인 1,000시간을 달성했습니다. 1단계 목표 1.5배의 두 배가 넘습니다. 이 씰을 적용한 신규 압축기 제작도 마쳤습니다.");
}

function sDiaphragm(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[내구] 다이어프램 압축기: 현장 교체주기 4.21배 연장");
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w, 4.7, "시험실 검증 (범한산업 · 한국자동차연구원)");
  const im = fit("p173_diaphragm_compressor_prototype.jpg", MX + 0.2, 2.15, 2.5, 3.9);
  s.addImage({ path: FIG("p173_diaphragm_compressor_prototype.jpg"), ...im, altText: "다이어프램 압축기 시제품" });
  addBullets(s, C, [
    [1, "PFA 코팅 다이어프램"], [2, [["내구성 ", false], ["1,561 hr", true], [" 공인", false]]], [2, "목표 1,500 hr(150%) 달성"],
    [1, "캐비티 형상 최적화"], [2, [["최대응력 ", false], ["22.6% 저감", true]]], [2, "형상 자동 생성 프로그램 개발"],
  ], MX + 2.95, 2.25, w - 3.15, 3.8, 14);
  s.addShape(pres.shapes.LINE, { x: MX + w + 0.2, y: 1.6, w: 0, h: 4.6, line: { color: BORDER, width: 1, dashType: "dash" } });
  const x2 = MX + w + 0.4;
  panel(s, pres, C, x2, 1.55, w, 4.7, "현장 검증 (가포 충전소 · 창원산업진흥원)", C.accent3);
  s.addChart(pres.charts.BAR, [{ name: "평균 교체 운전시간", labels: ["기존품", "개발품"], values: [117, 492.7] }], {
    x: x2 + 0.1, y: 2.1, w: 2.6, h: 4.05, barDir: "col", chartColors: ["9AA5B1", HEX.accent3], dataLabelFormatCode: "#,##0.#",
    title: "평균 교체 운전시간 (hr)", ...chartBase(), valAxisMinVal: 0, valAxisMaxVal: 580,
  });
  addBullets(s, C, [
    [1, "다이어프램 교체주기"], [2, [["117 → 492.7 hr(", false], ["4.21배", true], [")", false]]],
    [1, "개선 릴리프밸브"], [2, "사용 수명 3개월 → 8개월 이상"], [2, "응력 29.3% 개선(p.23)"],
  ], x2 + 2.8, 2.25, w - 2.95, 3.8, 14);
  callout(s, pres, C, [["개발품 상용 충전소 적용 결과 ", false], ["교체주기 4.21배 연장", true], [" 확인", false]]);
  source(s, C, "출처: 단계보고서 p.19, p.23~24 공인인증평가 성적서, p.26 상용충전소 운영 현황 보고서");
  s.addNotes("다이어프램은 실제 상용 충전소에 먼저 적용했습니다. 교체주기가 117시간에서 493시간으로 4배 넘게 늘었고, 릴리프밸브 수명도 3개월에서 8개월 이상으로 늘었습니다. 시험실에서는 공인시험으로 1,561시간을 확인했습니다. (4.21배의 표본 수와 기간은 백업으로 준비)");
}

function sDispenser(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[내구] 충전기: IR 리시버·유량조절밸브 공인 1.5배 달성");
  const items = [
    ["IR 리시버 (MSENG)", "p194_ir_receiver.jpg", [[1, [["15,000 cycle", true]]], [2, "기준 10,000 cycle 대비 150% 공인"], [2, "−40~20 ℃ 온도별 내구시험"], [2, "실차 충전 적용 및 로그 확보"]], C.text2],
    ["유량조절밸브 (MT.H)", "p277_valve_field_charging.jpg", [[1, [["153,030회", true]]], [2, "96.3 MPa 작동시험 150% 공인"], [2, "10개월 실충전 1,672회"], [2, "충전기 제조사 2곳 3기 공급"]], C.text2],
    ["듀얼 디스펜서 (제아ENG)", "p219_dual_dispenser.jpg", [[1, [["60 g/s급 시제품", true]]], [2, "동시충전 시제품 제작·시운전"], [2, "JICS 보드 전자파적합 등록"], [2, "120 g/s 공인시험 2026.11 예정"]], C.accent5],
  ];
  const w = (CW - 0.4) / 3;
  items.forEach(([hd, img, bl, col], i) => {
    const x = MX + i * (w + 0.2);
    panel(s, pres, C, x, 1.55, w, 4.7, hd, col);
    const im = fit(img, x + 0.2, 2.12, w - 0.4, 1.85);
    s.addImage({ path: FIG(img), ...im, altText: hd });
    addBullets(s, C, bl, x + 0.3, 4.1, w - 0.45, 2.05, 14);
  });
  callout(s, pres, C, [["밸브 개선품 ", false], ["10개월 실충전 1,672회", true], [" 안정 운전 및 제조사 공급", false]]);
  source(s, C, "출처: 단계보고서 p.29~32, p.39 공인인증평가 성적서, p.277, p.319");
  s.addNotes("충전기 핵심부품 두 가지는 공인시험에서 1.5배를 달성했습니다. 특히 밸브는 계획보다 앞당겨 설치해 10개월 동안 실제 차량 충전 1,672회를 수행했고, 이미 충전기 제조사 두 곳에 세 기를 공급했습니다. 듀얼 디스펜서는 시제품을 만들었고 11월에 공인시험을 받습니다.");
}

function sChiller(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[효율·내구] 냉각기: 고효율 칠러 충전 kg당 전력 21.4% 절감");
  s.addChart(pres.charts.BAR, [{ name: "충전 1 kg당 칠러 전력", labels: ["기존 칠러", "개선 칠러"], values: [0.7829, 0.6155] }], {
    x: MX, y: 1.55, w: 4.3, h: 4.7, barDir: "col", chartColors: ["9AA5B1", HEX.accent1], dataLabelFormatCode: "0.0000",
    title: "충전 1 kg당 칠러 전력 (kWh/kg, KCL)", ...chartBase(), valAxisMinVal: 0, valAxisMaxVal: 0.92,
  });
  panel(s, pres, C, 5.1, 1.55, 3.9, 4.7, "고효율 칠러 (삼정이엔씨)");
  addBullets(s, C, [
    [1, "충전 시 전력"], [2, [["0.7829 → 0.6155 kWh/kg(", false], ["−21.4%", true], [")", false]]],
    [1, "충전대기 시 전력"], [2, "22.89 → 20.30 kWh/h(−11.3%)"],
    [1, "적용 기술"], [2, "2단 압축·1단 팽창, 인버터"], [2, "대형 상용차용 216 kg/h 시제품"],
  ], 5.35, 2.15, 3.5, 4.0, 14);
  panel(s, pres, C, 9.2, 1.55, CW - 8.6, 4.7, "PCHE (에너진)", C.accent3);
  const im = fit("p239_pche_prototype.jpg", 9.4, 2.12, CW - 9.0, 1.9);
  s.addImage({ path: FIG("p239_pche_prototype.jpg"), ...im, altText: "PCHE 시제품" });
  addBullets(s, C, [[2, "130 MPa 이상 1시간 내압, 누설·균열 없음"], [2, "0~875 bar 반복가압 시험 운용(목표 100,000회)"]], 9.45, 4.15, CW - 9.1, 2.0, 14);
  callout(s, pres, C, [["충전 시 전력 절감 확인  →  ", false], ["대기 시 절감은 운전 로직 개선과 연계", true], [" 추진", false]]);
  source(s, C, "출처: 단계보고서 p.33, p.36 공인인증평가 성적서(KCL), p.239, p.324");
  s.addNotes("충전 시 냉각 전력을 21% 줄인 고효율 칠러를 공인기관에서 확인했습니다. PCHE는 130 MPa 내압을 통과했고 반복가압 시험을 진행하고 있습니다. 이 칠러는 뒤에서 말씀드릴 충전소 에너지 병목을 겨냥한 설비입니다.");
}

function sDiag(pres, C, t, section) {
  const s = contentSlide(pres, C, section, t, "[진단] 운전데이터 기반 조기경보: 실제 고장 92일 전 감지");
  const lw = 7.3;
  const tl = (y, title, span, marks, endLabel) => {
    panel(s, pres, C, MX, y, lw, 2.25, title);
    const ax = MX + 0.5, aw = lw - 1.7, ay = y + 1.62;
    s.addShape(pres.shapes.LINE, { x: ax, y: ay, w: aw, h: 0, line: { color: HEX.accent5, width: 2 } });
    s.addShape(pres.shapes.OVAL, { x: ax + aw - 0.14, y: ay - 0.14, w: 0.28, h: 0.28, fill: { color: "C0392B" } });
    s.addText(endLabel, { x: ax + aw - 0.45, y: ay + 0.2, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: "C0392B", margin: 0, isTextBox: true });
    for (const [d, lab, col] of marks) {
      const mx = ax + aw * (1 - d / span);
      s.addShape(pres.shapes.OVAL, { x: mx - 0.12, y: ay - 0.12, w: 0.24, h: 0.24, fill: { color: col } });
      s.addText([{ text: `D−${d}`, options: { fontSize: 20, bold: true, color: col, breakLine: true } }, { text: lab, options: { fontSize: 12, color: C.text1 } }],
        { x: mx - 0.7, y: y + 0.6, w: 1.4, h: 0.75, align: "center", margin: 0, isTextBox: true });
    }
  };
  tl(1.55, "압축기 중정비 고장(2023.12.18) 사전 감지", 100, [[92, "통계 모델", C.accent2], [41, "AI 모델", C.accent4]], "고장 발생");
  tl(4.0, "다이어프램 사건 조기경보", 8, [[6.57, "모델 1", C.accent2], [2.72, "모델 2", C.accent4]], "사건 추정 시작");
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  panel(s, pres, C, rx, 1.55, rw, 4.7, "신뢰성 관리체계 (고등기술연구원)");
  addBullets(s, C, [
    [1, "FMECA 정량화"], [2, [["고장모드 ", false], ["148개", true], [" RPN 산출", false]]], [2, "압축기 35 · 충전기 14 · 냉각기 99"],
    [1, "현장 환류"], [2, "설비별 고장대응 매뉴얼 3종"],
    [1, "결함진단 모델"], [2, "4방식 9종 비교 검증"], [2, "고압 압축기 진동 7개 측정점 Alert/Trip 기준"],
  ], rx + 0.3, 2.15, rw - 0.45, 4.0, 14);
  callout(s, pres, C, [["고장 후 정비  →  ", false], ["상태기반 예측정비(CBM)", true], [" 전환 기반 확보", false]]);
  source(s, C, "출처: 단계보고서 p.12~13, p.70, p.74, p.315");
  s.addNotes("실제 중정비가 발생한 고장을 통계 모델은 92일 전, AI 모델은 41일 전에 감지했습니다. 운영사는 부품 조달과 정비 일정을 미리 잡을 수 있습니다. 고장모드 148개의 위험도를 정량화해 설비별 고장대응 매뉴얼로 현장에 돌려줬습니다. (정확도 1.000은 질문 시 평가 기간 108일·사례 수와 함께 답변)");
}

function sEff1(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[효율] 공인 기준선 확정 및 충전 공정 비효율 원인 규명");
  s.addChart(pres.charts.BAR, [{ name: "SEC", labels: ["기준선\n(KTC 검증 실측)", "최종 목표\n(20% 절감)"], values: [9.5, 7.6] }], {
    x: MX, y: 1.55, w: 3.9, h: 4.15, barDir: "col", chartColors: ["9AA5B1", HEX.dk2], dataLabelFormatCode: "0.0",
    title: "충전소 SEC (kWh/kgH₂)", ...chartBase(), valAxisMinVal: 0, valAxisMaxVal: 11,
  });
  s.addText("정밀 분석 43주 9.1 · 전 기간 9.5\n목표 7.6은 기준선의 20% 절감 환산값", { x: MX, y: 5.75, w: 3.9, h: 0.55, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true });
  const rx = 4.75, rw = W - MX - rx;
  const im = fit("p095_main_chiller_standby.jpg", rx, 1.55, rw, 2.75);
  s.addImage({ path: FIG("p095_main_chiller_standby.jpg"), ...im, altText: "메인 칠러 시간별 전력" });
  s.addText("[메인 칠러 시간별 전력 및 시간대별 평균]", { x: rx, y: im.y + im.h + 0.02, w: rw, h: 0.28, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  const st = [["12.9 kWh/h", "야간(20~08시) 평균\n메인 칠러 전력"], ["38.7%", "칠러 전력량 중\n야간 대기 비중"], ["−40 ℃", "사전 냉각 상시 유지\n(SAE J2601 T40)"]];
  const sw = (rw - 0.4) / 3;
  st.forEach(([big, lab], i) => {
    const x = rx + i * (sw + 0.2);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 4.7, w: sw, h: 1.55, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 } });
    s.addText(big, { x, y: 4.78, w: sw, h: 0.6, fontSize: 24, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(lab, { x, y: 5.42, w: sw, h: 0.75, fontSize: 13, color: C.text1, align: "center", valign: "top", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["차량 수요와 무관한 ", false], ["메인 칠러 상시 대기전력", true], ["이 충전 공정 비효율의 핵심 원인", false]]);
  source(s, C, "출처: 단계보고서 p.12, p.67 KTC 성적서(GT2025-00687), p.95~96");
  s.addNotes("공인기관과 함께 이 충전소의 기준선을 9.5로 확정했습니다. 그리고 차량이 없는 밤에도 칠러가 영하 40도를 유지하느라 전력을 쓰는 구조가 가장 큰 낭비라는 것을 실측으로 확인했습니다.");
}

function sEff2(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[효율] 하드웨어·운영 투트랙 절감 수단 확보 및 검증 체계 수립");
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w, 2.3, "Track 1  하드웨어");
  addBullets(s, C, [[2, [["고효율 칠러 충전 kg당 ", false], ["−21.4%", true], ["(KCL)", false]]], [2, "압축기 동력전달 개선(플라이휠, 고효율 벨트)"], [2, "고내구 PCHE"]], MX + 0.3, 2.15, w - 0.5, 1.85, 15);
  panel(s, pres, C, MX + w + 0.4, 1.55, w, 2.3, "Track 2  운영 최적화", C.accent3);
  addBullets(s, C, [[2, "피크·오프피크 차등 운전(T40 / T30)"], [2, "절전 모드: Standby(10분) → Deep Sleep(60분)"], [2, "부하 패턴 기반 시간대별 제어"]], MX + w + 0.7, 2.15, w - 0.5, 1.85, 15);
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: 4.05, w: 4.9, h: 2.2, fill: { color: C.background1 }, line: { color: HEX.accent6, width: 1 } });
  pill(s, pres, C, "추정", MX + 0.2, 4.2, "추정");
  s.addText("1차 시뮬레이션 (메인 칠러 야간만)", { x: MX + 1.0, y: 4.15, w: 3.8, h: 0.38, fontSize: 14, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  addBullets(s, C, [[2, "야간 대기 155 → 23 kWh/일(−85%)"], [2, "최근 12개월 실측 8,629 h 적용"], [2, "칠러 1종·야간만 대상, 보수적 산정"]], MX + 0.3, 4.7, 4.45, 1.45, 14);
  const vx = MX + 5.1, vw = CW - 5.1;
  s.addText([{ text: "2단계 검증 방법  ", options: { bold: true, color: C.text1 } }, { text: "목표 9.5 → 7.6 kWh/kg(−20%)", options: { bold: true, color: C.accent1 } }],
    { x: vx, y: 4.08, w: vw, h: 0.4, fontSize: 16, margin: 0, isTextBox: true });
  const steps = ["① 전후 비교", "② 처리량 보정", "③ 공인 입회"];
  const sw = (vw - 0.3) / 3;
  steps.forEach((st, i) => stepHeader(s, pres, C, vx + i * (sw + 0.15), 4.7, sw, st, [C.text2, C.accent3, C.accent4][i], 15));
  addBullets(s, C, [[2, "적용 전후 설비별 에너지 사용량 비교"], [2, "처리량 효과와 순수 운영 개선 효과 분리 평가(p.366)"], [2, "외부전문가·공인기관 입회평가로 절감량 확정"]], vx + 0.1, 5.4, vw - 0.2, 0.95, 14);
  callout(s, pres, C, [["설비 교체 + 운전 최적화 통합  →  ", false], ["2단계 충전소 단위 20% 절감 공인검증", true]]);
  source(s, C, "출처: 단계보고서 p.36, p.314, p.363~366 · 시뮬레이션은 보고서 기재 추정값");
  s.addNotes("절감은 고효율 설비와 운전 로직, 두 축으로 진행합니다. 칠러 야간 운전만 바꿔도 대기전력을 85% 줄일 수 있다는 것을 실측 데이터로 확인했고, 2단계에서 설비 교체와 운전 최적화를 통합해 공인기관 입회 아래 충전소 단위로 검증합니다. 적용 전후 비교는 처리량 보정 SEC로 처리량 효과를 분리해서 평가합니다.");
}

function fDurability(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[내구] 핵심부품 4종 공인 1.5배 이상 달성, 일부 현장 효과 확인");
  const H = (x) => ({ text: x, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const rows = [
    [H("부품"), H("개선 전 → 개선 후"), H("개선율"), H("검증")],
    ["피스톤 씰 (GTC)", "300 → 1,000 hr", "3.3배", "공인"],
    ["다이어프램 (범한산업)", "현장 교체주기 117 → 492.7 hr", "4.21배", "공인 · 현장"],
    ["IR 리시버 (MSENG)", "10,000 → 15,000 cycle", "1.5배", "공인"],
    ["유량조절밸브 (MT.H)", "작동시험 153,030회 · 실충전 1,672회", "1.5배", "공인 · 현장"],
    ["고효율 칠러 (삼정이엔씨)", "0.7829 → 0.6155 kWh/kg", "−21.4%", "공인"],
  ].map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: { color: k === 2 ? C.accent1 : C.text1, bold: k === 0 || k === 2, align: k >= 2 ? "center" : "left" } })));
  s.addTable(rows, { x: MX, y: 1.6, w: 8.4, colW: [2.6, 3.5, 1.0, 1.3], rowH: 0.66, fontSize: 14, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.1, 0, 0.1], fill: { color: HEX.lt1 } });
  const rx = MX + 8.65, rw = CW - 8.65;
  const i1 = fit("p155_piston_parts.jpg", rx, 1.6, rw / 2 - 0.08, 2.1), i2 = fit("p194_ir_receiver.jpg", rx + rw / 2 + 0.08, 1.6, rw / 2 - 0.08, 2.1);
  s.addImage({ path: FIG("p155_piston_parts.jpg"), ...i1, altText: "개선 피스톤" });
  s.addImage({ path: FIG("p194_ir_receiver.jpg"), ...i2, altText: "IR 리시버" });
  const i3 = fit("p173_diaphragm_compressor_prototype.jpg", rx, 3.9, rw, 2.35);
  s.addImage({ path: FIG("p173_diaphragm_compressor_prototype.jpg"), ...i3, altText: "다이어프램 압축기 시제품" });
  s.addText("※ 실물은 견학 4번 지점에서 확인", { x: MX, y: 5.75, w: 8.4, h: 0.4, fontSize: 15, bold: true, color: C.accent3, margin: 0, isTextBox: true });
  callout(s, pres, C, [["핵심부품 4종 ", false], ["공인시험 1.5배 이상", true], [" · 다이어프램·밸브 ", false], ["현장 효과 확인", true]]);
  source(s, C, "출처: 단계보고서 p.19~39 공인인증평가 성적서, p.26, p.319");
  s.addNotes("고장이 잦던 부품 네 가지는 공인시험에서 1.5배 이상을 달성했고, 다이어프램과 밸브는 실제 충전소에서도 효과를 확인했습니다. 실물은 잠시 뒤 견학에서 보여드리겠습니다.");
}

function fEfficiency(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[효율] 기준선 확정·병목 규명 및 투트랙 절감 수단 확보");
  const lw = 7.0;
  const im = fit("p095_main_chiller_standby.jpg", MX, 1.55, lw, 2.75, "left");
  s.addImage({ path: FIG("p095_main_chiller_standby.jpg"), ...im, altText: "메인 칠러 대기전력" });
  const st = [["9.5", "kWh/kg 공인 기준선\n(KTC)"], ["38.7%", "칠러 전력량 중\n야간 대기"], ["−21.4%", "고효율 칠러\nkg당(KCL)"]];
  const sw = (lw - 0.4) / 3;
  st.forEach(([big, lab], i) => {
    const x = MX + i * (sw + 0.2);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 4.5, w: sw, h: 1.75, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 } });
    s.addText(big, { x, y: 4.6, w: sw, h: 0.65, fontSize: 26, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(lab, { x, y: 5.3, w: sw, h: 0.85, fontSize: 13, color: C.text1, align: "center", margin: 0, isTextBox: true });
  });
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  panel(s, pres, C, rx, 1.55, rw, 4.7, "절감 수단 및 2단계 검증");
  addBullets(s, C, [
    [1, "Track 1 하드웨어"], [2, "고효율 칠러, 압축기 동력전달 개선"],
    [1, "Track 2 운영 최적화"], [2, "피크·오프피크 차등 운전, 절전 모드"],
    [1, "2단계 검증"], [2, "적용 전후 비교 → 처리량 보정 SEC → 공인기관 입회"], [2, [["2027년 내 ", false], ["충전소 단위 절감량 검증", true]]],
  ], rx + 0.3, 2.15, rw - 0.45, 4.0, 14);
  callout(s, pres, C, [["메인 칠러 상시 대기전력", true], [" 병목 규명 → 설비 교체 + 운전 최적화로 2단계 검증", false]]);
  source(s, C, "출처: 단계보고서 p.12, p.36, p.95, p.314, p.363~366");
  s.addNotes("공인기관과 함께 기준선을 9.5로 확정했고, 밤에도 칠러가 사전 냉각을 유지하는 대기전력이 병목이라는 것을 찾았습니다. 고효율 칠러와 절전 운전 로직, 두 축으로 줄이고 2단계에서 공인기관 입회로 검증합니다. 잠시 뒤 견학 2번 지점에서 이 칠러를 보여드립니다.");
}

// 4. 실증 인프라 및 성능지표 -----------------------------------------
async function sInfra(pres, C, t, section) {
  const s = contentSlide(pres, C, section, t, "실증 설비동 조기 착공 및 인허가 완료");
  const lw = 6.6;
  const i1 = fit("p093_construction_photos.jpg", MX, 1.55, lw, 2.1, "left");
  s.addImage({ path: FIG("p093_construction_photos.jpg"), ...i1, altText: "설비동 공사 사진" });
  const y2 = i1.y + i1.h + 0.12;
  const i2 = fit("p360_install_location.jpg", MX, y2, lw, 6.25 - y2, "left");
  s.addImage({ path: FIG("p360_install_location.jpg"), ...i2, altText: "개선 핵심설비 설치 위치(안)" });
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  panel(s, pres, C, rx, 1.55, rw, 2.35, "충주 실증 설비동 인허가");
  const checks = ["KGS 변경 기술검토 적합", "건축허가", "고압가스 제조 변경허가", "설비동 기초 타설(2단계 공사 조기 착수)"];
  for (let i = 0; i < 4; i++) {
    const y = 2.12 + i * 0.43;
    s.addImage({ data: await icon("MdCheckCircle", HEX.accent1), x: rx + 0.3, y: y + 0.05, w: 0.28, h: 0.28 });
    s.addText(checks[i], { x: rx + 0.7, y, w: rw - 0.9, h: 0.38, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  }
  panel(s, pres, C, rx, 4.05, rw, 2.2, "설비 제작 및 세종대평 실증", C.accent3);
  addBullets(s, C, [
    [2, "고압 압축기 제작 완료(2026.07)"], [2, "냉동설비 도장·배치, 듀얼 디스펜서 제작 중"],
    [2, [["세종대평 충전기 실증 ", false], ["1년 앞당김", true]]], [2, "KGS 신청(2026.09.29) → 2026.12 준공 예정"],
  ], rx + 0.3, 4.62, rw - 0.45, 1.55, 14);
  callout(s, pres, C, [["2단계 착수와 동시에 ", false], ["개선 설비 설치·실증 가능", true]]);
  source(s, C, "출처: 단계보고서 p.13, p.28, p.93, p.315, p.317, p.360");
  s.addNotes("상용 충전소를 운영하면서 설비를 추가해야 하는 조건이었지만, 인허가를 모두 마치고 2단계 공사를 앞당겨 착공했습니다. 고압 압축기는 이미 제작을 마쳤습니다. 세종 충전소 충전기 실증도 당초보다 1년 앞당겼습니다.");
}

function statusCell(C, st) {
  const col = { 달성: C.accent1, 진행: C.accent5, 미달: C.accent6 }[st.split("(")[0]];
  return { text: "● " + st, options: { color: col, bold: true, align: "center" } };
}
function sPerf(pres, C, t) {
  const s = contentSlide(pres, C, "실증 인프라 및 성능지표", t, "성능지표 달성 현황: 부품 내구성 지표 전 항목 공인 달성");
  const H = (x) => ({ text: x, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
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
  ].map((r, i) => i === 0 ? r : r.map((c, k) => (k === 4 ? statusCell(C, c) : { text: c, options: { color: C.text1, align: k === 0 ? "left" : "center" } })));
  s.addTable(rows, { x: MX, y: 1.55, w: CW, colW: [3.7, 2.4, 3.0, 1.43, 1.6], rowH: 0.41, fontSize: 14, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.1, 0, 0.1], fill: { color: HEX.lt1 } });
  callout(s, pres, C, [["달성 6 · 진행 2 · 미달 2  →  ", false], ["부품 내구성 지표 전 항목 공인 달성", true]]);
  source(s, C, "출처: 단계보고서 p.10~39 기관별 성능지표, p.292~293 · 제품설계 보고서 지표는 보고서 참조");
  s.addNotes("부품 내구성 관련 지표는 모두 공인시험으로 달성했습니다. 충전 유량과 PCHE는 시험을 진행하고 있고, 미달은 다이어프램 압축기 전력과 충전소 소비에너지 저감 두 가지입니다.");
}

function sGap(pres, C, t) {
  const s = contentSlide(pres, C, "실증 인프라 및 성능지표", t, "미달 항목 원인 및 보완 계획");
  const rows = [
    ["특허 등록 4 / 10건", "출원 14건 중 2025년 출원 10건 심사 중", "의견제출·보정 대응 강화 → 2027년 내 등록 전환"],
    ["기술실시 2 / 4회", "수요기업의 현장 실증 결과 요구", "실증 결과 + 특허 연계 기술 패키지 이전"],
    ["충전소 에너지 절감 검증", "기준선 확정 후 운전 로직 개발 중", "2026.12 로직 완료 → 2027년 내 공인검증"],
    ["다이어프램 압축기 전력", "플라이휠 단독 개선 4% 절감", "고효율 벨트(5V) 동력전달 개선 후 재검증"],
  ];
  const tabs = [C.accent2, C.accent3, C.accent4, C.accent4];
  s.addText("미달 항목", { x: MX, y: 1.55, w: 2.9, h: 0.38, fontSize: 14, bold: true, color: C.accent5, align: "center", margin: 0, isTextBox: true });
  s.addText("원인", { x: MX + 3.15, y: 1.55, w: 4.2, h: 0.38, fontSize: 14, bold: true, color: C.accent5, margin: 0, isTextBox: true });
  s.addText("보완 계획", { x: MX + 7.6, y: 1.55, w: 4.4, h: 0.38, fontSize: 14, bold: true, color: C.accent5, margin: 0, isTextBox: true });
  rows.forEach(([a, b, c], i) => {
    const y = 2.0 + i * 1.03;
    s.addText(a, { shape: pres.shapes.RECTANGLE, x: MX, y, w: 2.9, h: 0.9, fill: { color: tabs[i] }, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: MX + 2.9, y, w: CW - 2.9, h: 0.9, fill: { color: ROWBG }, line: { color: ROWBG, width: 0 } });
    s.addText(b, { x: MX + 3.15, y, w: 4.2, h: 0.9, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: "✓  ", options: { color: C.accent1, bold: true } }, { text: c, options: { color: C.text1 } }], { x: MX + 7.6, y, w: 4.4, h: 0.9, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["공통 원인: 현장 실증 전 단계  →  ", false], ["2단계 1차년도 내 보완 완료", true], [" 추진", false]]);
  source(s, C, "출처: 단계보고서 4장 p.312~314");
  s.addNotes("미달 항목은 네 가지입니다. 공통 원인은 현장 실증 전 단계라는 점이고, 모두 2단계 1차년도에 보완 일정을 잡았습니다. (특허 등록 예상 건수는 확인 후 보완)");
}

function fPerfGap(pres, C, t) {
  const s = contentSlide(pres, C, "실증 인프라 및 성능지표", t, "성능지표 달성 현황 및 미달 항목 보완 계획");
  const sum = [["6", "달성", C.accent1], ["2", "진행", C.accent5], ["2", "미달", C.accent6]];
  sum.forEach(([n, lab, col], i) => {
    const y = 1.55 + i * 1.6;
    s.addShape(pres.shapes.RECTANGLE, { x: MX, y, w: 3.1, h: 1.45, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 } });
    s.addText([{ text: n, options: { fontSize: 40, bold: true, color: col } }, { text: "  " + lab, options: { fontSize: 18, bold: true, color: C.text1 } }],
      { x: MX, y, w: 3.1, h: 1.45, align: "center", valign: "middle", margin: 0, isTextBox: true });
  });
  const rows = [
    ["특허 등록 4 / 10건", "심사 중 10건 대응 → 2027년 내 등록 전환"],
    ["기술실시 2 / 4회", "실증 결과 + 특허 연계 기술 패키지 이전"],
    ["충전소 에너지 절감 검증", "2026.12 로직 완료 → 2027년 내 공인검증"],
    ["다이어프램 압축기 전력", "고효율 벨트 적용 후 재검증"],
  ];
  const tx = MX + 3.4, tw = CW - 3.4;
  rows.forEach(([a, b], i) => {
    const y = 1.55 + i * 1.18;
    s.addText(a, { shape: pres.shapes.RECTANGLE, x: tx, y, w: 3.0, h: 1.05, fill: { color: [C.accent2, C.accent3, C.accent4, C.accent4][i] }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: tx + 3.0, y, w: tw - 3.0, h: 1.05, fill: { color: ROWBG }, line: { color: ROWBG, width: 0 } });
    s.addText([{ text: "✓  ", options: { color: C.accent1, bold: true } }, { text: b }], { x: tx + 3.25, y, w: tw - 3.4, h: 1.05, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["부품 내구성 지표 공인 달성 · 미달 항목 ", false], ["2단계 1차년도 내 보완", true]]);
  source(s, C, "출처: 단계보고서 p.291~293, 4장 p.312~314 · 성능지표 10개 기준");
  s.addNotes("부품 내구성 지표는 모두 공인시험으로 달성했습니다. 미달 항목은 네 가지이고, 공통 원인은 현장 실증 전 단계라는 점입니다. 모두 2단계 1차년도에 보완 일정을 잡았습니다.");
}

// 5. 2단계 추진 계획 -------------------------------------------------
function sNext(pres, C, t) {
  const s = contentSlide(pres, C, "2단계 추진 계획", t, "1단계 결과 → 2단계 현장 검증 연계");
  const pairs = [
    ["피스톤 씰 1,000 hr, 신규 압축기", "압축기 패키지 현장 장기 실증"],
    ["고효율 칠러, 절전 로직", "충전소 단위 에너지 절감 공인검증"],
    ["FMECA·조기경보", "HyEReD 신뢰성 DB, 잔여수명(RUL) 예측"],
    ["경제성 평가 툴", "K-충전소 운영모델, 탄소배출 예측"],
  ];
  const bw = 5.0, ah = 0.55, gap = 0.1;
  s.addText("1단계 결과", { x: MX, y: 1.5, w: bw, h: 0.32, fontSize: 14, bold: true, color: C.accent5, margin: 0, isTextBox: true });
  s.addText("2단계 과제", { x: MX + CW - bw, y: 1.5, w: bw, h: 0.32, fontSize: 14, bold: true, color: C.accent5, margin: 0, isTextBox: true });
  pairs.forEach(([a, b], i) => {
    const y = 1.85 + i * (ah + gap);
    s.addText(a, { shape: pres.shapes.RECTANGLE, x: MX, y, w: bw, h: ah, fill: { color: ROWBG }, fontSize: 15, color: C.text1, valign: "middle", margin: 10, isTextBox: true });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: MX + bw + 0.45, y: y + 0.13, w: CW - 2 * bw - 0.9, h: 0.3, fill: { color: C.accent4 } });
    s.addText(b, { shape: pres.shapes.RECTANGLE, x: MX + CW - bw, y, w: bw, h: ah, fill: { color: C.text2 }, fontSize: 15, bold: true, color: C.background1, valign: "middle", margin: 10, isTextBox: true });
  });
  const yrs = [["2027 (4차년도)", C.text2, [[2, "실증 설비 설치·시운전"], [2, "에너지 절감 공인검증"], [2, "HyEReD 개발"]]], ["2028 (5차년도)", C.accent3, [[2, "장기 실증"], [2, "HyEReD 유효성 검증"], [2, "조기경보 운영절차서"]]], ["2029.1Q (6차년도)", C.accent4, [[2, "K-충전소 운영모델"], [2, "공인 성적서·설계표준안"]]]];
  const yw = (CW - 0.4) / 3;
  yrs.forEach(([hd, col, bl], i) => {
    const x = MX + i * (yw + 0.2);
    stepHeader(s, pres, C, x, 4.55, yw, hd, col);
    addBullets(s, C, bl, x + 0.3, 5.18, yw - 0.45, 1.15, 14);
  });
  callout(s, pres, C, [["2단계는 신규 연구가 아닌 ", false], ["1단계 결과물의 현장 검증·확산", true]]);
  source(s, C, "출처: 단계보고서 요약문 p.4~5, 9장 p.357~366");
  s.addNotes("2단계의 모든 과제는 1단계 결과물에서 바로 이어집니다. 2027년에 실증 설비를 설치하고 에너지 절감량을 공인검증하며, 2028년 장기 실증을 거쳐 2029년 K-충전소 운영모델로 정리합니다.");
}

function sRisk(pres, C, t) {
  const s = contentSlide(pres, C, "2단계 추진 계획", t, "2단계 실증 리스크 대응 현황");
  const rs = [
    ["실증 설비 설치", C.accent2, "✓  2단계 공사 1단계 착공 · 고압 압축기 제작 완료(2026.07)", "기초 타설 완료"],
    ["인허가", C.accent3, "✓  충주: KGS 기술검토·건축·고압가스 변경허가 완료\n✓  세종: 안전영향평가·기술검토 신청(2026.09.29)", "충주 완료 · 세종 신청"],
    ["신뢰성 DB 데이터", C.accent4, "✓  충주 3개년 상용 운전 데이터 + 세종·가포 데이터\n✓  OREDA 구조 및 미국 HyCReD 벤치마킹", "3개년 데이터 확보"],
  ];
  rs.forEach(([tab, col, txt, st], i) => {
    const y = 1.6 + i * 1.5;
    s.addText(tab, { shape: pres.shapes.RECTANGLE, x: MX, y, w: 2.6, h: 1.3, fill: { color: col }, fontSize: 17, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: MX + 2.6, y, w: CW - 2.6, h: 1.3, fill: { color: ROWBG }, line: { color: ROWBG, width: 0 } });
    s.addText(txt, { x: MX + 2.85, y, w: 6.8, h: 1.3, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    s.addText(st, { shape: pres.shapes.ROUNDED_RECTANGLE, x: MX + 9.85, y: y + 0.38, w: 2.05, h: 0.55, rectRadius: 0.27, fill: { color: C.accent1 }, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["공사 조기 착공 · 인허가 완료 · 데이터 확보로 ", false], ["2단계 일정 리스크 최소화", true]]);
  source(s, C, "출처: 단계보고서 p.13, p.28, p.93, p.315, p.358");
  s.addNotes("실증 설비 설치는 공사를 앞당겨 착공했고 고압 압축기 제작을 마쳤습니다. 인허가는 충주는 완료, 세종은 신청을 마쳤습니다. 신뢰성 DB는 3년치 상용 운전 데이터에 세종·가포 데이터를 더하고 OREDA 구조로 벤치마킹합니다.");
}

function sBiz(pres, C, t) {
  const s = contentSlide(pres, C, "2단계 추진 계획", t, "활용 및 사업화 계획");
  const st = [["3,280", "백만 원", "1단계 사업화 매출 (목표 2,000)"], ["7", "건", "1단계 사업화 (목표 3)"], ["8,000", "백만 원", "2단계 매출 목표 (사업화 12건)"]];
  const w = (CW - 0.4) / 3;
  st.forEach(([n, u, lab], i) => {
    const x = MX + i * (w + 0.2);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w, h: 2.0, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 } });
    s.addText([{ text: n, options: { fontSize: 40, bold: true, color: C.accent1 } }, { text: " " + u, options: { fontSize: 17, color: C.text1 } }],
      { x, y: 1.7, w, h: 1.05, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(lab, { x, y: 2.85, w, h: 0.45, fontSize: 15, color: C.text1, align: "center", margin: 0, isTextBox: true });
  });
  const road = [["부품 공급", C.text2, [[2, "밸브 개선품 3기 제조사 공급"], [2, "피스톤 씰·다이어프램 적용 확대"]]], ["설비 패키지", C.accent3, [[2, "압축기·냉각기·충전기"], [2, "현장 실증 후 공급"]]], ["운영모델", C.accent4, [[2, "K-충전소 운영절차서"], [2, "경제성·탄소배출 예측"]]]];
  road.forEach(([hd, col, bl], i) => {
    const x = MX + i * (w + 0.2);
    stepHeader(s, pres, C, x, 3.85, w, hd, col);
    addBullets(s, C, bl, x + 0.3, 4.6, w - 0.45, 1.6, 16);
  });
  callout(s, pres, C, [["기대효과: ", false], ["충전소 셧다운 방지 · OPEX 절감 · 가동률 향상", true]]);
  source(s, C, "출처: 단계보고서 p.6, p.39, p.291");
  s.addNotes("1단계 사업화 매출은 32억 8천만 원으로 목표를 넘었습니다. 밸브 개선품은 이미 제조사에 공급되고 있고, 2단계에는 설비 패키지와 운영모델로 사업화를 넓힙니다. (매출 품목별 내역은 백업으로 준비)");
}

function fNext(pres, C, t) {
  const s = contentSlide(pres, C, "2단계 추진 계획", t, "2단계 충전소 현장 검증 계획 및 리스크 대응");
  const yrs = [["2027 (4차년도)", C.text2, [[2, "실증 설비 설치·시운전"], [2, "에너지 절감 공인검증"], [2, "HyEReD 개발"]]], ["2028 (5차년도)", C.accent3, [[2, "장기 실증"], [2, "HyEReD 유효성 검증"], [2, "조기경보 운영절차서"]]], ["2029.1Q (6차년도)", C.accent4, [[2, "K-충전소 운영모델"], [2, "공인 성적서·설계표준안"]]]];
  const yw = (CW - 0.4) / 3;
  yrs.forEach(([hd, col, bl], i) => {
    const x = MX + i * (yw + 0.2);
    stepHeader(s, pres, C, x, 1.6, yw, hd, col);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.25, w: yw, h: 1.75, fill: { color: C.background1 }, line: { color: BORDER, width: 0.75 } });
    addBullets(s, C, bl, x + 0.3, 2.4, yw - 0.45, 1.5, 15);
  });
  checkRow(s, pres, C, 4.25, 0.85, "리스크 대응", C.accent3, "✓  공사 조기 착공 · 고압 압축기 제작 완료 · 충주 인허가 완료 · 세종 실증 1년 앞당김");
  checkRow(s, pres, C, 5.2, 0.85, "2단계 목표", C.accent2, "✓  설비 패키지 내구성 1.5배 현장 검증 · 충전소 에너지 20% 절감 공인검증 · K-충전소 운영모델");
  callout(s, pres, C, [["2단계: ", false], ["1단계 결과물의 충전소 현장 검증", true]]);
  source(s, C, "출처: 단계보고서 p.4~5, p.13, p.93, p.317, 9장 p.357~366");
  s.addNotes("2027년에 실증 설비를 설치하고 에너지 절감량을 공인검증하며, 2028년 장기 실증을 거쳐 2029년 K-충전소 운영모델로 정리합니다. 설치 일정 리스크는 공사를 앞당겨 착공해 줄였습니다.");
}

// 6. 결론 ------------------------------------------------------------
function sConclusion(pres, C, t) {
  const s = contentSlide(pres, C, "결론", t, null);
  const H = (x) => ({ text: x, options: { fontSize: 21, bold: true, color: C.text1, breakLine: true, paraSpaceBefore: 14 } });
  const line = (parts, code, lvl, last) => {
    const runs = parts.map(([x, hl]) => ({ text: x, options: { fontSize: 18, color: hl ? C.accent1 : C.text1, bold: !!hl } }));
    Object.assign(runs[0].options, { bullet: { code, indent: code === "2192" ? 24 : 18 }, indentLevel: lvl, paraSpaceBefore: 5 });
    if (!last) runs[runs.length - 1].options.breakLine = true;
    return runs;
  };
  const B = (p) => line(p, "2022", 0);
  const A = (p, last) => line(p, "2192", 1, last);
  const runs = [
    H("1단계 개요"),
    ...B([["수소충전소 핵심설비(압축기·충전기·냉각기) 부품 개선 및 신뢰성 관리체계 구축", false]]),
    ...B([["상용 충전소 운전데이터 기반 고장모드 정량화 및 에너지 기준선 확정", false]]),
    H("주요 성과"),
    ...B([["내구: 핵심부품 4종 ", false], ["공인시험 1.5배 이상", true], [" 달성(피스톤 씰 330%)", false]]),
    ...B([["진단: 실제 중정비 고장 ", false], ["92일 전 사전 감지", true], [", 고장모드 148개 RPN 정량화", false]]),
    ...B([["효율: ", false], ["공인 기준선 9.5 kWh/kg 확정", true], [", 고효율 칠러 kg당 21.4% 절감", false]]),
    ...B([["실증: 인허가 완료 및 ", false], ["2단계 실증 설비동 조기 착공", true]]),
    H("2단계 추진 방향"),
    ...B([["1단계 결과물의 충전소 현장 검증:", false]]),
    ...A([["[내구] 설비 패키지 ", false], ["내구성 1.5배 현장 검증", true]]),
    ...A([["[효율] 충전소 단위 ", false], ["에너지 20% 절감 공인검증", true]]),
    ...A([["[운영] HyEReD 기반 ", false], ["K-수소충전소 운영모델", true], [" 구축", false]], true),
  ];
  runs[0].options.paraSpaceBefore = 0;
  s.addText(runs, { x: MX, y: 1.15, w: CW, h: 5.6, valign: "top", margin: 0, isTextBox: true, objectName: "conclusion" });
  s.addText("본 과제는 기후에너지환경부와 한국에너지기술평가원의 지원을 받아 수행 중인 연구개발 과제임 (No. RS-2024-00419764)", { x: MX, y: 6.95, w: 11.5, h: 0.32, fontSize: 13, color: C.text1, margin: 0, isTextBox: true });
  s.addNotes("정리하겠습니다. 1단계에서 부품 내구성은 공인시험으로 입증했고, 고장을 미리 아는 체계와 에너지 기준선을 확보했으며, 실증 인프라도 준비했습니다. 2단계에서는 충전소 현장에서 내구성과 에너지 절감을 증명하고 K-충전소 운영모델로 완성하겠습니다.");
}

function sThanks(pres) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "맺음" });
  s.addText("감사합니다", { placeholder: "title" });
  s.addNotes("감사합니다.");
}

function fTour(pres, C, t) {
  const s = contentSlide(pres, C, "현장 확인", t, "현장 견학 동선 및 확인 사항");
  const im = fit("p049_process_flow.jpg", MX, 1.5, CW, 2.75);
  s.addImage({ path: FIG("p049_process_flow.jpg"), ...im, altText: "충주바이오그린수소충전소 공정 흐름도" });
  const stops = [["① 모니터링 화면", [[2, "실시간 전력 15채널"], [2, "진동 7개 측정점·조기경보"]]], ["② 메인 칠러·압축기", [[2, "야간 대기전력 병목"], [2, "2단계 교체 대상"]]], ["③ 설비동 공사 현장", [[2, "기초 타설 완료"], [2, "개선 설비 반입 예정 위치"]]], ["④ 개선 부품 실물", [[2, "피스톤 씰·다이어프램"], [2, "IR 리시버·밸브, 성적서"]]]];
  const w = (CW - 0.6) / 4;
  stops.forEach(([hd, bl], i) => {
    const x = MX + i * (w + 0.2);
    stepHeader(s, pres, C, x, 4.4, w, hd, [C.text2, C.accent3, C.accent4, C.accent1][i]);
    addBullets(s, C, bl, x + 0.25, 5.05, w - 0.35, 1.2, 14);
  });
  callout(s, pres, C, [["견학 동선: ", false], ["모니터링 → 칠러·압축기 → 공사 현장 → 실물 전시", true]]);
  source(s, C, "공정 흐름도: 단계보고서 p.49 · 견학 동선은 현장 여건에 맞춰 조정 [확인 필요]");
  s.addNotes("이제 현장으로 이동하겠습니다. 모니터링 화면, 메인 칠러와 압축기, 설비동 공사 현장, 개선 부품 실물 순서로 보여드리겠습니다.");
}

// Appendix ----------------------------------------------------------
function appendix(pres, C) {
  const d = pres.addSlide({ masterName: "PLAIN", sectionTitle: "Appendix" });
  d.addText("Appendix", { placeholder: "title" });
  const two = (title, sub, a, b, capA, capB, src, note) => {
    const s = contentSlide(pres, C, "Appendix", title, sub);
    const w = (CW - 0.3) / 2;
    const ia = fit(a, MX, 1.6, w, 4.6, "center", "top"), ib = fit(b, MX + w + 0.3, 1.6, w, 4.6, "center", "top");
    s.addImage({ path: FIG(a), ...ia, altText: capA });
    s.addImage({ path: FIG(b), ...ib, altText: capB });
    const cy = Math.max(ia.y + ia.h, ib.y + ib.h) + 0.08;
    s.addText(`[${capA}]`, { x: MX, y: cy, w, h: 0.3, fontSize: 12, color: C.text1, align: "center", margin: 0, isTextBox: true });
    s.addText(`[${capB}]`, { x: MX + w + 0.3, y: cy, w, h: 0.3, fontSize: 12, color: C.text1, align: "center", margin: 0, isTextBox: true });
    source(s, C, src);
    s.addNotes(note);
  };
  two("[Appendix] 기준선 경계", "기준선 9.5: KTC 검증 계측 체계의 생산·압축·충전 전 공정 실측값",
    "p067_process_share.jpg", "fig15_ktc_verification.png", "공정별 에너지 비중과 SEC 비중 (p.67)", "KTC 성적서 대비 모니터링 데이터 교차검증 (추가 분석)",
    "출처: 단계보고서 p.12, p.67 · 교차검증은 보고서 외 추가 분석", "기준선 경계 질문 대응: 9.5는 생산·압축·충전 공정을 포함한 사업대상지 실측값이며, KTC 측정값과 모니터링 데이터의 오차는 평균 1.6%입니다.");
  two("[Appendix] 처리량 보정", "처리량 보정 SEC 기반 절감 효과 분리 평가",
    "p068_sec_vs_throughput.jpg", "p095_full_period_sec.jpg", "처리량과 공정별 SEC 상관 (p.68)", "모니터링 전 기간 주간 공정별 SEC (2024.08~2026.08, p.95)",
    "출처: 단계보고서 p.68, p.95, p.366", "처리량 보정 질문 대응: 분기 SEC는 처리량에 따라 7.0~15.4로 변동합니다. 그래서 2단계 검증은 처리량 보정 SEC로 순수 운영 개선 효과를 분리해서 평가합니다(p.366).");
  two("[Appendix] 절감 시뮬레이션", "부하 패턴 기반 절전 시간대 설정 및 1차 보수적 산정",
    "p364_load_heatmap.jpg", "p365_standby_simulation.jpg", "요일·시간대별 충전소 전력 패턴 (p.364)", "메인 칠러 대기전력 절감 시뮬레이션 (추정, p.365)",
    "출처: 단계보고서 p.364~365 · 시뮬레이션은 추정값", "1단계 실적 질문 대응: 충전소 단위 절감량 공인검증은 2단계 1차년도로 이월했습니다. 기준선과 병목은 확정했고, 1차 시뮬레이션은 메인 칠러 야간만 대상으로 한 보수적 산정(충전소 약 2.7%)이며, 하드웨어 교체와 운영 최적화를 통합해 20%를 추진합니다.");
}

// ---------------------------------------------------------------- 빌드
async function buildMain() {
  const { pres, C } = makePres("본평가판");
  pres.addSection({ title: "표지" }); sCover(pres, C, "main");
  pres.addSection({ title: "목차" }); sToc(pres, C, ["과제 개요", "1단계 추진 실적", "1단계 연구 결과", "실증 인프라 및 성능지표", "2단계 추진 계획", "결론"]);
  pres.addSection({ title: "과제 개요" });
  sBackground(pres, C, "1. 과제 개요 (1/3)"); await sGoal(pres, C, "1. 과제 개요 (2/3)"); sOrg(pres, C, "1. 과제 개요 (3/3)");
  pres.addSection({ title: "1단계 추진 실적" });
  await sSummary(pres, C, "2. 1단계 추진 실적 (1/3)"); sTimeline(pres, C, "2. 1단계 추진 실적 (2/3)"); sDashboard(pres, C, "2. 1단계 추진 실적 (3/3)");
  pres.addSection({ title: "1단계 연구 결과" });
  sPiston(pres, C, "3. 1단계 연구 결과 (1/7)"); sDiaphragm(pres, C, "3. 1단계 연구 결과 (2/7)"); sDispenser(pres, C, "3. 1단계 연구 결과 (3/7)");
  sChiller(pres, C, "3. 1단계 연구 결과 (4/7)"); sDiag(pres, C, "3. 1단계 연구 결과 (5/7)", "1단계 연구 결과");
  sEff1(pres, C, "3. 1단계 연구 결과 (6/7)"); sEff2(pres, C, "3. 1단계 연구 결과 (7/7)");
  pres.addSection({ title: "실증 인프라 및 성능지표" });
  await sInfra(pres, C, "4. 실증 인프라 및 성능지표 (1/3)", "실증 인프라 및 성능지표"); sPerf(pres, C, "4. 실증 인프라 및 성능지표 (2/3)"); sGap(pres, C, "4. 실증 인프라 및 성능지표 (3/3)");
  pres.addSection({ title: "2단계 추진 계획" });
  sNext(pres, C, "5. 2단계 추진 계획 (1/3)"); sRisk(pres, C, "5. 2단계 추진 계획 (2/3)"); sBiz(pres, C, "5. 2단계 추진 계획 (3/3)");
  pres.addSection({ title: "결론" }); sConclusion(pres, C, "6. 결론");
  pres.addSection({ title: "맺음" }); sThanks(pres);
  pres.addSection({ title: "Appendix" }); appendix(pres, C);
  const f = path.join(OUT, "1단계_단계평가_본평가판.pptx");
  await pres.writeFile({ fileName: f }); console.log("pPr 정리", await dedupeParagraphProps(f)); await applyTheme(f, THEME); return f;
}
async function buildField() {
  const { pres, C } = makePres("현장평가판");
  pres.addSection({ title: "표지" }); sCover(pres, C, "field");
  pres.addSection({ title: "목차" }); sToc(pres, C, ["과제 개요", "1단계 추진 실적", "1단계 연구 결과", "실증 인프라 및 성능지표", "2단계 추진 계획", "결론 및 현장 확인"]);
  pres.addSection({ title: "과제 개요" }); fOverview(pres, C, "1. 과제 개요");
  pres.addSection({ title: "1단계 추진 실적" }); await sSummary(pres, C, "2. 1단계 추진 실적 (1/2)"); sDashboard(pres, C, "2. 1단계 추진 실적 (2/2)");
  pres.addSection({ title: "1단계 연구 결과" });
  fDurability(pres, C, "3. 1단계 연구 결과 (1/3)"); sDiag(pres, C, "3. 1단계 연구 결과 (2/3)", "1단계 연구 결과"); fEfficiency(pres, C, "3. 1단계 연구 결과 (3/3)");
  pres.addSection({ title: "실증 인프라 및 성능지표" });
  await sInfra(pres, C, "4. 실증 인프라 및 성능지표 (1/2)", "실증 인프라 및 성능지표"); fPerfGap(pres, C, "4. 실증 인프라 및 성능지표 (2/2)");
  pres.addSection({ title: "2단계 추진 계획" }); fNext(pres, C, "5. 2단계 추진 계획");
  pres.addSection({ title: "결론 및 현장 확인" }); sConclusion(pres, C, "6. 결론"); fTour(pres, C, "6. 현장 확인 안내");
  pres.addSection({ title: "맺음" }); sThanks(pres);
  const f = path.join(OUT, "1단계_단계평가_현장평가판.pptx");
  await pres.writeFile({ fileName: f }); console.log("pPr 정리", await dedupeParagraphProps(f)); await applyTheme(f, THEME); return f;
}

(async () => {
  console.log(await buildMain());
  console.log(await buildField());
})().catch((e) => { console.error(e); process.exit(1); });
