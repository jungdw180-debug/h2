// 1단계 단계평가 발표자료 생성 스크립트 (본평가판 + 현장평가판)
// 양식: 고등기술연구원 발표 양식(번호 제목, 소제목 표시, 남색 머리띠 패널, ◆·▪ 개조식, 하단 강조 박스)
// 실행: NODE_PATH=<pptxgenjs·react-icons·sharp node_modules> APPLY_THEME=<apply_theme.js> node slides/build_decks.js
// 내용 근거: report/발표자료_슬라이드내용_v1.md, report/예상질의응답_v2.md (쪽수는 단계보고서 인쇄 쪽 번호)

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

// ================================================================ 발표 대본 (노트)
// 첫 줄에 목표 시간을 붙인다. 한국어 발표 속도 약 300자/분(공백 제외) 기준.
const NOTES = {
  cover: "안녕하십니까. 고등기술연구원 [발표자]입니다. 수소충전소 핵심설비와 부품의 내구성 및 효율 향상 과제의 1단계 결과를 보고드리겠습니다.",
  toc: "발표는 과제 개요, 1단계 추진 실적과 연구 결과, 실증 인프라와 성능지표, 2단계 추진 계획, 결론 순서로 말씀드리겠습니다.",
  background: "먼저 과제 배경입니다. 국내 수소충전소 핵심설비는 대부분 중소기업이 공급하는데, 현장에서는 세 가지 문제가 반복됩니다. 첫째, 고장이 잦습니다. 저희가 분석한 상용 충전소에서 다이어프램은 평균 117시간마다 교체됐고, 릴리프밸브는 석 달을 버티지 못했습니다. 아래 사진이 실제 고장 부품입니다. 둘째, 에너지를 많이 씁니다. 수소 1킬로그램을 충전하는 데 약 9.5 kWh를 쓰고 있고, 압축기와 냉각기에 부하가 집중됩니다. 셋째, 국산 설비의 검증 실적이 부족해 현장 신뢰를 얻기 어렵습니다. 그래서 본 과제는 핵심부품을 고내구화하고, 에너지를 20% 이상 줄이며, 데이터로 고장을 미리 아는 관리체계를 만드는 것을 목표로 합니다.",
  goal: "최종목표는 세 가지입니다. 핵심부품과 설비의 내구성을 1.5배 높이고, 에너지 사용량을 20% 이상 줄이며, 이를 묶어 K-수소충전소 운영모델을 제시하는 것입니다. 과제는 두 단계로 설계되어 있습니다. 1단계에서는 부품을 개선해 공인시험으로 성능을 확인하고, 신뢰성 관리체계와 실증 기반을 만들었습니다. 2단계에서는 개선 설비를 실제 충전소에 설치해 설비 단위 내구성과 충전소 단위 에너지 절감을 현장에서 검증합니다. 기술성숙도는 TRL 7에서 8로 올리는 것이 목표입니다.",
  org: "추진체계입니다. 12개 기관이 부품, 설비, 충전소 세 층으로 연결되어 있습니다. 부품 기업이 피스톤 씰, 다이어프램, IR 노즐, 밸브, PCHE를 개선하고, 설비 기업이 이를 압축기, 충전기, 냉각기에 적용하며, 주관기관과 운영기관이 진단과 실증, 운영모델을 맡습니다. 이 층들은 데이터와 제품으로 실제 연결되어 있습니다. 창원산업진흥원이 현장 교체 데이터를 분석해 범한산업 설계에 반영했고, 코하이젠의 고장이력은 진단 알고리즘 개발에 쓰였습니다. 밸브 개선품은 컨소시엄 안의 충전기 제조사 두 곳에 공급되었고, 듀얼 디스펜서 성능시험은 삼정이엔씨와 함께 수행했습니다.",
  summary: "1단계 결과를 네 가지로 요약하겠습니다. 첫째, 내구입니다. 고장이 잦던 핵심부품 네 가지를 개선해 공인시험에서 모두 1.5배 이상을 확인했고, 피스톤 씰은 3.3배입니다. 둘째, 진단입니다. 운전데이터로 실제 중정비 고장을 92일 전에 감지했습니다. 셋째, 효율입니다. 공인기관과 함께 에너지 기준선 9.5를 확정했고, 고효율 칠러로 충전 시 전력을 21.4% 줄였습니다. 넷째, 실증입니다. 인허가 세 건을 모두 받고 2단계 공사를 앞당겨 착공했습니다. 부품은 입증했고, 이를 충전소에서 검증할 준비가 끝났습니다.",
  achieve: "계획 대비 달성도입니다. 평가위원님들께서 연구개발계획서와 비교하실 수 있도록 개발 항목 기준으로 정리했습니다. 1단계 개발 항목은 1차년도 30개, 2차년도 29개, 3차년도 33개로 모두 92개이며, 전 항목을 완료했습니다. 분야별로는 분석·진단 16개, 압축기 26개, 냉각기 5개, 충전기 33개, 현장 실증 12개입니다. 1차년도에는 고장 분석과 설계 기준을 세웠고, 2차년도에는 시제품과 공인 기준치를 확보했으며, 3차년도에는 공인시험으로 목표를 확인하고 실증 인허가를 마쳤습니다. 오른쪽에 분야별 대표 산출물을 정리했고, 이어지는 연구 결과에서 자세히 말씀드리겠습니다.",
  dashboard: "정량 성과입니다. 1단계 목표가 설정된 12개 지표 중 10개를 달성했습니다. SCIE 논문은 목표 5건 대비 14건, SW 저작권은 5건 대비 28건, 사업화 매출은 20억 원 목표 대비 32억 8천만 원을 달성했습니다. 고용은 35명, 현장시험은 14건으로 실증 중심 과제의 특성이 반영된 결과입니다. 다만 주황색으로 표시한 특허 등록과 기술실시 두 지표는 목표에 못 미쳤습니다. 특허는 14건을 출원했으나 심사 중인 건이 많고, 기술실시는 실증 결과를 기다리는 수요기업이 있기 때문입니다. 원인과 보완 계획은 뒤에서 따로 말씀드리겠습니다.",
  piston: "지금부터 분야별 연구 결과입니다. 먼저 피스톤 압축기입니다. 피스톤 씰은 마모가 빨라 수소 보충과 정비가 잦은 부품이었습니다. 오른쪽 아래처럼 마모 해석으로 취약 부위를 확인한 뒤, 피스톤 소재를 A7075로 바꾸고 이중 스프링씰과 보조 씰 구조를 적용했습니다. 그 결과 공인시험에서 기준 300시간 대비 2025년 450시간, 2026년 1,000시간을 달성했습니다. 1단계 목표가 1.5배였는데 3.3배로 두 배 이상 넘었습니다. 이 씰을 적용한 신규 압축기 제작도 마쳐 2단계에 실증 설비동에 설치합니다.",
  diaphragm: "다음은 다이어프램 압축기입니다. 왼쪽은 시험실 결과입니다. 캐비티 형상을 최적화해 최대응력을 22.6% 낮췄고, PFA 코팅 다이어프램은 공인시험에서 1,561시간을 확인했습니다. 오른쪽은 현장 결과입니다. 가포 상용 충전소에 개발품을 먼저 적용했더니, 평균 교체 운전시간이 기존 117시간에서 492.7시간으로 4.21배 늘었습니다. 최초 적용품은 1,025시간까지 운전했습니다. 함께 개선한 릴리프밸브도 수명이 3개월에서 8개월 이상으로 늘었습니다. 시험실을 넘어 상용 충전소에서 효과를 확인한 사례입니다.",
  dispenser: "충전기 핵심부품입니다. 충전노즐의 IR 리시버는 영하 40도부터 상온까지 온도별 시험을 거쳐 공인시험에서 15,000사이클, 기준 대비 150%를 달성했습니다. 유량조절밸브는 96.3 MPa 작동시험 153,030회로 역시 150%를 달성했습니다. 특히 밸브는 계획보다 앞당겨 코하이젠 충전소에 설치해 10개월 동안 실제 차량 충전 1,672회를 문제없이 수행했고, 이미 충전기 제조사 두 곳에 세 기를 공급했습니다. 동시충전 듀얼 디스펜서는 시제품을 제작해 시운전했고, 120 g/s 공인시험은 11월에 진행합니다.",
  chiller: "냉각기입니다. 삼정이엔씨가 2단 압축·1단 팽창 방식과 인버터를 적용한 고효율 칠러를 개발했습니다. 공인기관 측정 결과 충전 1킬로그램당 칠러 전력이 0.7829에서 0.6155 kWh로 21.4% 줄었고, 충전대기 시에도 시간당 전력이 줄었습니다. 대형 상용차 충전용 216 kg/h급 시제품도 제작했습니다. 에너진의 PCHE는 130 MPa 이상 내압을 누설 없이 통과했고, 0에서 875 bar 반복가압 시험을 진행하고 있습니다. 이 칠러는 뒤에서 말씀드릴 충전소 에너지 병목을 직접 겨냥한 설비입니다.",
  diag: "신뢰성 관리체계입니다. 충주 상용 충전소의 3년 운전·정비 데이터로 이상감지 모델을 만들었습니다. 2023년 12월 실제 중정비가 발생한 압축기 고장을 통계 모델은 92일 전, AI 모델은 41일 전에 감지했습니다. 아래 그림처럼 다이어프램 사건에서도 6.57일, 2.72일 전에 경보를 냈습니다. 또 압축기, 충전기, 냉각기 고장모드 148개의 위험도를 정량화해 설비별 고장대응 매뉴얼 세 종으로 현장에 돌려줬습니다. 고장 후 정비에서 미리 준비하는 정비로 바뀌는 기반입니다.",
  eff1: "효율 분야입니다. 먼저 출발점인 기준선을 확정했습니다. 공인기관 입회로 검증한 계측 체계로 사업대상지의 에너지 사용량을 측정해, 수소 1킬로그램당 9.5 kWh를 기준선으로 정했습니다. 최종 목표는 여기서 20% 줄인 7.6입니다. 그리고 병목을 찾았습니다. 오른쪽 그래프처럼 차량이 없는 밤에도 메인 칠러가 영하 40도 사전 냉각을 유지하느라 시간당 12.9 kWh를 쓰고 있었고, 칠러 전력량의 38.7%가 야간 대기에 쓰이고 있었습니다.",
  eff2: "절감 수단은 두 축으로 확보했습니다. 하드웨어 측면에서는 방금 보신 고효율 칠러와 압축기 동력전달 개선, 고내구 PCHE가 있습니다. 운영 측면에서는 피크와 오프피크를 나눈 차등 운전과, 유휴 시간에 단계적으로 들어가는 절전 모드를 설계했습니다. 칠러 야간 운전만 바꿔도 대기전력을 85% 줄일 수 있다는 것을 12개월 실측 데이터로 확인했습니다. 2단계에서는 두 축을 함께 적용하고, 처리량 영향을 보정한 SEC로 순수 절감 효과를 분리해 공인기관 입회 아래 충전소 단위로 검증하겠습니다.",
  infra: "실증 인프라입니다. 상용 충전소를 운영하면서 설비를 추가해야 하는 조건이었지만, KGS 변경 기술검토, 건축허가, 고압가스 제조 변경허가를 모두 받았고, 2단계에 하기로 했던 설비동 공사를 앞당겨 기초까지 타설했습니다. 오른쪽 배치도가 신규 압축기, 냉각기, 충전기가 들어갈 위치입니다. 고압 압축기는 이미 제작을 마쳤고, 세종대평 충전소의 충전기 실증도 1년 앞당겨 12월 준공 예정입니다. 2단계가 시작되는 날부터 개선 설비를 설치할 수 있습니다.",
  perf: "성능지표 달성 현황입니다. 연구개발계획서에 정한 성능지표를 1단계 목표와 실적, 검증 방식으로 정리했습니다. 피스톤식과 다이어프램식 압축기 내구성, 충전기 IR 리시버 내구성, 유량조절밸브 작동시험, 냉각기 에너지효율은 모두 공인기관 시험으로 달성했습니다. 충전 유량은 11월 공인시험을, PCHE는 반복가압 시험을 진행하고 있습니다. 미달은 다이어프램 압축기 전력 절감과 충전소 단위 에너지 저감 검증 두 가지이며, 바로 다음 장에서 보완 계획을 말씀드리겠습니다.",
  gap: "미달 항목과 보완 계획입니다. 특허는 14건을 출원했지만 심사 기간 때문에 등록은 4건입니다. 심사 대응을 강화해 2027년 안에 등록으로 전환하겠습니다. 기술실시는 수요기업이 현장 실증 결과를 요구해 2건에 그쳤고, 실증 결과와 특허를 묶어 이전하겠습니다. 충전소 에너지 절감은 기준선과 병목을 확정했고, 운전 로직을 12월까지 완료해 2027년 안에 공인검증을 마치겠습니다. 다이어프램 압축기 전력은 고효율 벨트로 동력전달을 개선해 다시 검증합니다. 네 항목의 공통 원인은 현장 실증 전 단계라는 점이며, 2단계 1차년도에 보완하겠습니다.",
  next: "2단계 계획입니다. 2단계는 새로운 연구가 아니라 1단계 결과물을 현장에서 검증하고 확산하는 단계입니다. 피스톤 씰 압축기는 패키지로 장기 실증하고, 고효율 칠러와 절전 로직은 충전소 단위 에너지 절감으로 검증합니다. FMECA와 조기경보는 신뢰성 데이터베이스와 잔여수명 예측으로, 경제성 평가 툴은 K-충전소 운영모델로 발전시킵니다. 2027년 설치와 공인검증, 2028년 장기 실증, 2029년 운영모델 순서로 진행합니다.",
  target: "2단계 성과 목표와 연구비입니다. 2단계는 1단계에서 논문과 소프트웨어 성과를 충분히 확보한 만큼, 실증 이후 권리화와 사업화에 무게를 두었습니다. 특허 등록 11건, 사업화 12건, 매출 80억 원이 핵심 목표입니다. 연구비는 2단계 전체 약 78억 원이며, 실증 설비 설치와 장기 실증이 집중되는 4차년도와 5차년도에 각각 약 34억 원을 배정했고, 마지막 3개월인 6차년도는 결과 정리와 운영모델 완성에 씁니다.",
  risk: "실증 단계의 리스크와 대응 현황입니다. 가장 큰 리스크는 설비 설치 지연인데, 2단계 공사를 1단계에 앞당겨 착공했고 고압 압축기는 이미 제작을 마쳤습니다. 인허가는 충주가 완료되었고, 세종은 안전영향평가와 기술검토 신청을 마쳤습니다. 신뢰성 데이터베이스는 데이터 양이 관건인데, 충주 3년치 상용 데이터에 세종과 가포 데이터를 더하고, 해외 OREDA 구조와 미국 HyCReD로 벤치마킹하겠습니다.",
  biz: "활용과 사업화입니다. 1단계 사업화 매출은 32억 8천만 원으로 목표를 넘었습니다. 지티씨와 범한산업의 압축기, MT.H의 밸브가 실제로 판매되었습니다. 시제품은 아래와 같이 압축기, 충전기, 듀얼 디스펜서, 냉각기, 밸브까지 7건을 제작했습니다. 2단계에서는 부품 공급을 넘어 실증을 마친 설비 패키지와 운영모델로 사업화를 넓혀, 충전소 셧다운 방지와 운영비 절감에 기여하겠습니다.",
  conclusion: "정리하겠습니다. 1단계에서는 핵심부품 네 가지의 내구성을 공인시험으로 1.5배 이상 입증했고, 실제 고장을 92일 전에 감지하는 관리체계와 공인 에너지 기준선을 확보했습니다. 그리고 이 성과를 검증할 실증 설비동을 이미 착공했습니다. 2단계에서는 개선 설비를 충전소에 설치해 설비 내구성 1.5배와 에너지 20% 절감을 현장에서 검증하고, 신뢰성 데이터베이스와 함께 K-수소충전소 운영모델로 완성하겠습니다. 이상으로 발표를 마치겠습니다.",
  thanks: "감사합니다. 질문 주시면 답변드리겠습니다.",
  // 현장평가판
  f_overview: "안녕하십니까. 현장에서 짧게 과제를 소개드리고 견학으로 이어가겠습니다. 국내 수소충전소 핵심설비는 고장이 잦습니다. 이 충전소에서도 다이어프램이 평균 117시간마다 교체됐고, 릴리프밸브는 석 달을 버티지 못했습니다. 왼쪽 아래 사진이 실제 고장 부품입니다. 또 수소 1킬로그램에 약 9.5 kWh를 쓰고, 국산 설비의 현장 실적이 부족합니다. 그래서 본 과제는 핵심부품과 설비 내구성 1.5배, 에너지 20% 절감, K-충전소 운영모델을 목표로 합니다. 12개 기관이 부품, 설비, 충전소로 연결되어 1단계에서는 개발과 체계 구축, 2단계에서는 현장 검증을 수행합니다.",
  f_summary: "1단계 성과입니다. 위쪽 네 가지가 핵심입니다. 첫째, 핵심부품 네 가지는 공인시험에서 1.5배 이상을 달성했고 피스톤 씰은 3.3배입니다. 둘째, 이 충전소 데이터로 만든 이상감지 모델이 실제 중정비 고장을 92일 전에 감지했습니다. 셋째, 공인기관과 함께 에너지 기준선 9.5를 확정했고 고효율 칠러로 충전 전력을 21.4% 줄였습니다. 넷째, 인허가를 모두 받고 실증 설비동을 앞당겨 착공했습니다. 아래는 정량지표입니다. 12개 중 10개를 달성했고, 특허 등록과 기술실시는 2단계에서 보완하겠습니다.",
  f_durability: "내구 성과를 표로 정리했습니다. 피스톤 씰은 공인시험에서 300시간에서 1,000시간으로 3.3배가 되었습니다. 다이어프램은 상용 충전소에서 교체주기가 117시간에서 492.7시간으로 4.21배 늘었습니다. IR 리시버는 15,000사이클, 유량조절밸브는 153,030회로 각각 1.5배를 달성했고, 밸브는 실제 충전소에서 10개월간 1,672회 충전했습니다. 고효율 칠러는 충전 1킬로그램당 전력을 21.4% 줄였습니다. 이 부품들의 실물과 성적서는 견학 4번 지점에서 직접 보여드리겠습니다.",
  f_diag: "진단 성과입니다. 이 충전소의 3년 운전·정비 데이터로 이상감지 모델을 만들었습니다. 2023년 12월 실제 중정비가 발생한 압축기 고장을 통계 모델은 92일 전, AI 모델은 41일 전에 감지했고, 다이어프램 사건도 며칠 앞서 경보를 냈습니다. 또 고장모드 148개의 위험도를 정량화해 고장대응 매뉴얼 세 종을 만들었습니다. 견학 1번 지점 모니터링 화면에서 실시간 전력과 진동 데이터를 보여드리겠습니다.",
  f_eff: "효율 성과입니다. 공인기관 입회로 이 충전소의 에너지 기준선을 수소 1킬로그램당 9.5 kWh로 확정했습니다. 그래프처럼 차량이 없는 밤에도 메인 칠러가 영하 40도를 유지하느라 전력을 쓰고 있었고, 칠러 전력량의 38.7%가 야간 대기에 쓰였습니다. 그래서 고효율 칠러로 교체하는 하드웨어 축과, 피크와 오프피크를 나눈 절전 운전 축으로 줄이겠습니다. 2단계에서 공인기관 입회로 충전소 단위 절감량을 검증합니다. 견학 2번 지점에서 이 칠러를 보여드리겠습니다.",
  f_infra: "실증 인프라입니다. 상용 운영을 유지하면서 설비를 추가해야 하는 조건이었지만, KGS 기술검토, 건축허가, 고압가스 제조 변경허가를 모두 받았고, 2단계 공사를 앞당겨 설비동 기초까지 타설했습니다. 배치도의 붉은 표시가 신규 압축기, 냉각기, 충전기가 들어갈 위치입니다. 고압 압축기는 이미 제작했고, 세종 충전기 실증도 1년 앞당겼습니다. 견학 3번 지점에서 공사 현장을 직접 보여드리겠습니다.",
  f_perfgap: "성능지표와 미달 항목입니다. 성능지표는 달성 6개, 진행 2개, 미달 2개이며, 부품 내구성 지표는 모두 공인시험으로 달성했습니다. 오른쪽은 미달 항목의 보완 계획입니다. 특허 등록은 심사 대응으로 2027년 안에 전환하고, 기술실시는 실증 결과와 특허를 묶어 이전합니다. 충전소 에너지 절감은 12월에 운전 로직을 완료하고 2027년 안에 공인검증을 마치며, 다이어프램 압축기 전력은 고효율 벨트로 다시 검증합니다.",
  f_next: "2단계 계획입니다. 2027년에는 이 설비동에 개선 설비를 설치하고 에너지 절감량을 공인검증하며, 2028년에는 장기 실증과 신뢰성 데이터베이스 검증, 2029년에는 K-충전소 운영모델로 정리합니다. 설치 일정 리스크는 공사를 앞당겨 착공하고 압축기를 미리 제작해 줄였습니다.",
  f_conclusion: "정리하면, 핵심부품 내구성은 공인시험으로 입증했고, 고장을 미리 아는 체계와 에너지 기준선을 확보했으며, 실증 설비동도 착공했습니다. 2단계에서 이 충전소에서의 현장 검증으로 완성하겠습니다.",
  f_tour: "이제 현장으로 이동하겠습니다. 모니터링 화면, 메인 칠러와 압축기, 설비동 공사 현장, 개선 부품 실물 순서로 보여드리겠습니다.",
  // Appendix 답변
  a_errata: "수치 질문 대응: 보고서 장마다 집계 기준일이 달라 생긴 차이입니다. 2026년 10월 30일 제출본 정량성과표를 공식 수치로 하며, 정오표를 제출하겠습니다.",
  a_patent: "특허 질문 대응: 14건을 출원했고 4건이 등록되었습니다. 심사 중인 10건은 의견제출과 보정 대응을 하고 있으며, 2027년 안에 등록 전환을 추진합니다.",
  a_paper: "논문 질문 대응: SCIE 14건 중 충전소 설비·운영과 직접 관련된 논문이 7건입니다. 나머지는 수소 생산·운송과 공정 경제성 분석 역량을 활용한 파생 성과입니다.",
  a_sales: "매출 질문 대응: 매출은 GTC 압축기 납품, 범한산업 압축기, MT.H 밸브 공급으로 구성됩니다. 2026년 GTC 매출 10억 8백만 원은 세금계산서 발행 후 연내 수금 예정입니다.",
  a_field: "현장 근거 질문 대응: 다이어프램은 기존품과 개발품 각각 교체 3회를 비교했습니다. 최초 적용품이 1,025시간을 운전했고, 이후 교체주기 단축도 확인되어 원인을 범한산업과 분석하고 있습니다.",
  a_budget: "연구비 질문 대응: 1단계 연구비는 약 118억 원이며 2단계는 약 78억 원입니다. 기관별 집행 현황은 보고서 7장에 정리되어 있습니다.",
  a_safety: "안전관리 질문 대응: 전 참여기관이 관련 법령에 따른 안전관리를 이행했고, 2단계 설비 설치와 시운전 단계의 안전관리 계획을 수립했습니다.",
  a_baseline: "기준선 경계 질문 대응: 9.5는 생산·압축·충전 공정을 포함한 사업대상지 실측값입니다. KTC 측정값과 모니터링 데이터의 오차는 평균 1.6%입니다.",
  a_throughput: "처리량 보정 질문 대응: 분기 SEC는 처리량에 따라 7.0에서 15.4까지 변동합니다. 그래서 2단계 검증은 처리량 보정 SEC로 순수 운영 개선 효과를 분리해 평가합니다.",
  a_sim: "1단계 실적 질문 대응: 충전소 단위 절감량 공인검증은 2단계 1차년도로 이월했습니다. 기준선과 병목은 확정했고, 1차 시뮬레이션은 메인 칠러 야간만 대상으로 한 보수적 산정이며, 하드웨어 교체와 운영 최적화를 통합해 20%를 추진합니다.",
};
function addNote(s, key) {
  const t = NOTES[key];
  const sec = Math.max(5, Math.round(t.replace(/\s/g, "").length / 5 / 5) * 5);
  s.addNotes(`⏱ 약 ${sec}초\n${t}`);
}

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
  addNote(s, "cover");
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
  addNote(s, "toc");
}

// 1. 과제 개요 -----------------------------------------------------
function sBackground(pres, C, t) {
  const s = contentSlide(pres, C, "과제 개요", t, "국산 수소충전소 핵심설비의 내구성·효율 확보 필요");
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w, 4.7, "국내 수소충전소 핵심설비 현황");
  addBullets(s, C, [
    [1, "잦은 고장과 부품 교체"], [2, "다이어프램 평균 교체 운전시간 약 117 hr(기존품)"], [2, "릴리프밸브 사용 수명 약 3개월"],
    [1, "높은 에너지 사용량"], [2, [["충전소 SEC 약 ", false], ["9.5 kWh/kgH₂", true]]],
    [1, "운영실적(Track Record) 부족"], [2, "중소기업 중심 공급 → 개발·검증 인프라 필요"],
  ], MX + 0.3, 2.15, w - 0.5, 2.65, 15);
  const im = fit("p160_relief_valve_failure.jpg", MX + 0.2, 4.85, w - 0.4, 1.05);
  s.addImage({ path: FIG("p160_relief_valve_failure.jpg"), ...im, altText: "기존 릴리프밸브 고장 부품" });
  s.addText("[기존 부품 고장 형태 (p.160)]", { x: MX, y: 5.92, w, h: 0.26, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.LINE, { x: MX + w + 0.2, y: 1.6, w: 0, h: 4.6, line: { color: "C9D3DD", width: 1, dashType: "dash" } });
  const x2 = MX + w + 0.4;
  panel(s, pres, C, x2, 1.55, w, 4.7, "과제 추진 방향");
  addBullets(s, C, [
    [1, "핵심부품 고내구화·국산화"], [2, "압축기·충전기 소모성 부품 내구성 1.5배 이상 향상"],
    [1, "에너지 사용량 20% 이상 절감"], [2, "압축기·냉각기 효율 향상 및 운전 최적화"],
    [1, "데이터 기반 신뢰성 관리"], [2, "고장 후 정비 → 상태기반 예측정비(CBM) 전환"],
    [1, "글로벌 기술·표준 선점 대응"], [2, "美·日 수소전략 개정(’23.6), EU AFIR 등 경쟁 가속"],
  ], x2 + 0.3, 2.15, w - 0.5, 4.0);
  callout(s, pres, C, [["핵심설비 내구성·효율 향상을 통한 ", false], ["충전소 운영 경제성(OPEX) 확보", true], [" 시급", false]]);
  source(s, C, "출처: 단계보고서 p.2, p.26, p.160");
  addNote(s, "background");
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
    ["1단계 (2024.04~2026.12)", C.text2, "개발·체계 구축", [[2, "핵심부품 개선, 시제품 제작 및 공인 평가"], [2, "신뢰성 관리모드(FMECA·조기경보) 구축"], [2, "실증 사이트 설계 및 인허가"]], "부품 단위 공인 1.5배 이상 달성"],
    ["2단계 (2027.01~2029.03)", C.accent4, "충전소 현장 검증", [[2, "개선 설비 충전소 설치 및 장기 실증"], [2, "에너지 절감량 공인기관 검증"], [2, "HyEReD DB 및 K-충전소 운영모델"]], "설비 1.5배 · 충전소 20% 절감 현장 검증"],
  ];
  for (let i = 0; i < 2; i++) {
    const x = MX + i * (sw + 0.4);
    stepHeader(s, pres, C, x, 2.75, sw, st[i][0], st[i][1]);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.4, w: sw, h: 2.8, fill: { color: C.background1 }, line: { color: BORDER, width: 0.75 } });
    s.addText(st[i][2], { x: x + 0.3, y: 3.52, w: sw - 0.6, h: 0.5, fontSize: 19, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    addBullets(s, C, st[i][3], x + 0.3, 4.08, sw - 0.6, 1.3, 16);
    s.addText([{ text: "목표 수준  ", options: { bold: true, color: C.accent5 } }, { text: st[i][4], options: { bold: true, color: C.accent2 } }],
      { shape: pres.shapes.RECTANGLE, x: x + 0.25, y: 5.5, w: sw - 0.5, h: 0.5, fill: { color: ROWBG }, fontSize: 14, valign: "middle", margin: 8, isTextBox: true });
  }
  callout(s, pres, C, [["TRL 7 → 8  ·  1단계 개발 결과를 ", false], ["2단계 충전소 실증으로 검증", true], ["하는 구조", false]]);
  source(s, C, "출처: 단계보고서 요약문 p.3, p.5");
  addNote(s, "goal");
}

function sOrg(pres, C, t) {
  const s = contentSlide(pres, C, "과제 개요", t, "부품–설비–충전소 연계 추진체계 (12개 기관)");
  const rows = [
    ["충전소 운영모델", C.accent2, "✓  고등기술연구원(총괄·진단·실증) · 고려대(경제성·탄소) · 창원산업진흥원, 코하이젠(현장 실증)"],
    ["설비", C.accent3, "✓  압축기: GTC, 범한산업 · 충전기: MSENG, 제아ENG · 냉각기: 삼정이엔씨"],
    ["부품", C.accent4, "✓  피스톤 씰: GTC, 한국자동차연구원 · 다이어프램: 범한산업, 한국자동차연구원\n✓  IR 노즐: MSENG · 유량조절밸브: MT.H콘트롤밸브 · PCHE: 에너진"],
  ];
  const hs = [0.7, 0.7, 0.95];
  let y = 1.55;
  rows.forEach(([tab, col, txt], i) => { checkRow(s, pres, C, y, hs[i], tab, col, txt); y += hs[i] + 0.1; });
  panel(s, pres, C, MX, 4.2, CW, 2.05, "기관 간 협력 실적");
  const w = (CW - 0.6) / 2;
  addBullets(s, C, [
    [1, "데이터 환류"], [2, "창원산업진흥원 현장 교체 데이터 → 범한산업 설계 개선"], [2, "코하이젠 고장이력 → 주관기관 진단 알고리즘"],
  ], MX + 0.3, 4.78, w, 1.4, 14);
  addBullets(s, C, [
    [1, "제품·시험 연계"], [2, "MT.H 밸브 개선품 → MSENG·제아ENG 3기 공급"], [2, "듀얼 디스펜서 성능시험 제아ENG–삼정이엔씨 협업"],
  ], MX + 0.3 + w + 0.2, 4.78, w, 1.4, 14);
  callout(s, pres, C, [["현장 데이터 → 고장모드 정량화 → 개선 부품 → 현장 재확인의 ", false], ["순환 구조", true]]);
  source(s, C, "출처: 단계보고서 표지, p.26, p.28, p.39, p.42, p.315 · 지원: 충주시, 충북도");
  addNote(s, "org");
}

function fOverview(pres, C, t) {
  const s = contentSlide(pres, C, "과제 개요", t, "국산 수소충전소 핵심설비의 내구성·효율 확보 및 현장 검증");
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w, 4.7, "현황 및 필요성");
  addBullets(s, C, [
    [1, "잦은 고장"], [2, "다이어프램 교체주기 약 117 hr"], [2, "릴리프밸브 수명 약 3개월"],
    [1, "높은 에너지 사용량"], [2, [["충전소 SEC 약 ", false], ["9.5 kWh/kgH₂", true]]],
    [1, "운영실적 부족"], [2, "효율·내구성 검증 및 현장 운영실적 미흡"],
  ], MX + 0.3, 2.15, w - 0.5, 2.7);
  const im = fit("p160_relief_valve_failure.jpg", MX + 0.2, 4.9, w - 0.4, 1.0);
  s.addImage({ path: FIG("p160_relief_valve_failure.jpg"), ...im, altText: "기존 릴리프밸브 고장 부품" });
  s.addText("[기존 부품 고장 형태]", { x: MX, y: 5.92, w, h: 0.26, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  const x2 = MX + w + 0.4;
  panel(s, pres, C, x2, 1.55, w, 4.7, "최종목표 및 추진체계");
  addBullets(s, C, [
    [1, "최종목표"], [2, "핵심부품·설비 내구성 1.5배"], [2, "에너지 사용량 20% 이상 절감"], [2, "K-수소충전소 운영모델 제시"],
    [1, "단계 구조"], [2, "1단계 개발·체계 구축 → 2단계 충전소 현장 검증"],
    [1, "추진체계"], [2, "부품 → 설비 → 충전소 운영모델 연계(12개 기관)"],
  ], x2 + 0.3, 2.15, w - 0.5, 4.0);
  callout(s, pres, C, [["주관 고등기술연구원 · 공동 11개 기관 · 지원 충주시, 충북도  →  ", false], ["부품–설비–충전소 연계 추진", true]]);
  source(s, C, "출처: 단계보고서 p.2~5, p.26, p.160");
  addNote(s, "f_overview");
}

// 2. 1단계 추진 실적 -------------------------------------------------
const SUMMARY_CARDS = [
  ["MdPrecisionManufacturing", "내구", "1.5배+", "", [[2, "핵심부품 4종 공인"], [2, "피스톤 씰 330%"]], "공인"],
  ["MdMonitorHeart", "진단", "92일 전", "", [[2, "중정비 고장 사전 감지"], [2, "고장모드 148개"]], "현장"],
  ["MdElectricBolt", "효율", "9.5", " kWh/kg", [[2, "공인 기준선 확정"], [2, "칠러 kg당 −21.4%"]], "공인"],
  ["MdConstruction", "실증", "착공", "", [[2, "인허가 3건 완료"], [2, "2단계 공사 조기 착수"]], "현장"],
];
async function sSummary(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 추진 실적", t, "1단계 성과 요약");
  const w = (CW - 0.6) / 4;
  for (let i = 0; i < 4; i++) {
    const [ic, lab, big, unit, items, b] = SUMMARY_CARDS[i];
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
  addNote(s, "summary");
}

function sAchieve(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 추진 실적", t, "1단계 개발 목표 대비 달성도: 92개 항목 전 항목 100%");
  const H = (x) => ({ text: x, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const data = [
    ["분석·진단·예측", "고등기술연구원, 고려대", 5, 5, 6, "FMECA 148개 고장모드, 조기경보 모델, 경제성 평가 보고서"],
    ["압축기", "한국자동차연구원, GTC, 범한산업", 7, 8, 11, "피스톤 씰 1,000 hr, 다이어프램 1,561 hr, 신규 압축기 제작"],
    ["냉각기", "삼정이엔씨", 2, 2, 1, "고효율 칠러(충전 kg당 −21.4%), 액·기 분리기 냉각기"],
    ["충전기", "MSENG, 제아ENG, 에너진, MT.H", 12, 10, 11, "IR 15,000 cycle, 밸브 153,030회, 듀얼 디스펜서, PCHE"],
    ["현장 실증", "창원산업진흥원, 코하이젠", 4, 4, 4, "가포 충전소 실증, 세종대평 실증 설계·인허가"],
  ];
  const rows = [[H("분야"), H("참여기관"), H("1차년도"), H("2차년도"), H("3차년도"), H("달성도"), H("대표 산출물")]];
  data.forEach(([a, b, c, d, e, f]) => rows.push([
    { text: a, options: { bold: true, color: C.accent2 } }, { text: b, options: { color: C.text1 } },
    { text: String(c), options: { align: "center", color: C.text1 } }, { text: String(d), options: { align: "center", color: C.text1 } }, { text: String(e), options: { align: "center", color: C.text1 } },
    { text: "100%", options: { align: "center", bold: true, color: C.accent1 } }, { text: f, options: { color: C.text1 } },
  ]));
  const tot = (k) => data.reduce((a, r) => a + r[k], 0);
  rows.push([
    { text: "합계", options: { bold: true, color: C.text1, fill: { color: C.background2 } } }, { text: "12개 기관", options: { bold: true, color: C.text1, fill: { color: C.background2 } } },
    ...[2, 3, 4].map((k) => ({ text: `${tot(k)}개`, options: { bold: true, align: "center", color: C.text1, fill: { color: C.background2 } } })),
    { text: "100%", options: { bold: true, align: "center", color: C.accent1, fill: { color: C.background2 } } },
    { text: `총 ${tot(2) + tot(3) + tot(4)}개 개발 항목 완료`, options: { bold: true, color: C.accent1, fill: { color: C.background2 } } },
  ]);
  s.addTable(rows, { x: MX, y: 1.6, w: CW, colW: [1.75, 2.75, 0.95, 0.95, 0.95, 0.95, 3.833], rowH: [0.45, 0.66, 0.66, 0.66, 0.66, 0.66, 0.55], fontSize: 14, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.1, 0, 0.1], fill: { color: HEX.lt1 } });
  callout(s, pres, C, [["연구개발계획서 1단계 개발 항목 ", false], ["92개 전 항목 100% 완료", true]]);
  source(s, C, "출처: 단계보고서 3장 p.40~42 연차별 개발목표 달성도 (항목 수는 기관별 개발목표 집계)");
  addNote(s, "achieve");
}

const DASH = [
  ["SCIE 논문", "14", "5", "건"], ["국내 논문", "8", "8", "건"], ["SW 저작권", "28", "5", "건"], ["사업화", "7", "3", "건"],
  ["사업화 매출", "3,280", "2,000", "백만 원"], ["고용 창출", "35", "16", "명"], ["현장시험", "14", "10", "건"], ["시제품", "7", "5", "건"],
  ["MOU", "12", "2", "건"], ["국가인증", "2", "1", "건"], ["특허 등록", "4", "10", "건", "출원 14건"], ["기술실시", "2", "4", "회", ""],
];
function sDashboard(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 추진 실적", t, "정량 성과: 12개 지표 중 10개 달성");
  const cols = 6, g = 0.16, w = (CW - g * (cols - 1)) / cols, h = 2.18;
  for (let i = 0; i < DASH.length; i++) {
    const [lab, a, tg, u, note] = DASH[i];
    const miss = note !== undefined;
    const x = MX + (i % cols) * (w + g), y = 1.6 + Math.floor(i / cols) * (h + g);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: miss ? C.background1 : C.background2 }, line: { color: miss ? HEX.accent6 : CALLOUT_LINE, width: miss ? 1.25 : 0.75 } });
    s.addText(lab, { shape: pres.shapes.RECTANGLE, x, y, w, h: 0.42, fill: { color: miss ? C.accent6 : C.text2 }, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(a, { x: x + 0.1, y: y + 0.55, w: w - 0.2, h: 0.85, fontSize: a.length > 3 ? 30 : 36, bold: true, color: miss ? C.accent6 : C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(`목표 ${tg} ${u}`, { x: x + 0.1, y: y + 1.45, w: w - 0.2, h: 0.3, fontSize: 13, color: C.accent5, align: "center", margin: 0, isTextBox: true });
    if (miss) s.addText(note ? `미달 · ${note}` : "미달", { x: x + 0.1, y: y + 1.78, w: w - 0.2, h: 0.28, fontSize: 12, bold: true, color: C.accent6, align: "center", margin: 0, isTextBox: true });
  }
  callout(s, pres, C, [["SCIE 14건 · SW 28건 · 사업화 매출 3,280백만 원 · 고용 35명 등 ", false], ["10개 지표 목표 초과", true]]);
  source(s, C, "출처: 단계보고서 p.291 정량적 연구개발성과표(2026.10.30 제출본)");
  addNote(s, "dashboard");
}

async function fSummaryDash(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 추진 실적", t, "1단계 성과 요약 및 정량 성과");
  const w = (CW - 0.6) / 4;
  for (let i = 0; i < 4; i++) {
    const [ic, lab, big, unit, items, b] = SUMMARY_CARDS[i];
    const x = MX + i * (w + 0.2);
    panel(s, pres, C, x, 1.55, w, 2.85, lab);
    await iconCircle(s, pres, C, ic, x + 0.25, 2.15, 0.6, C.accent1);
    s.addText([{ text: big, options: { fontSize: 30, bold: true, color: C.accent1 } }, { text: unit, options: { fontSize: 13, color: C.accent1 } }],
      { x: x + 0.95, y: 2.1, w: w - 1.05, h: 0.7, valign: "middle", margin: 0, isTextBox: true });
    addBullets(s, C, items, x + 0.25, 3.0, w - 0.35, 1.2, 14);
  }
  const cols = 6, g = 0.12, mw = (CW - g * (cols - 1)) / cols, mh = 0.82;
  DASH.forEach(([lab, a, tg, u, note], i) => {
    const miss = note !== undefined;
    const x = MX + (i % cols) * (mw + g), y = 4.55 + Math.floor(i / cols) * (mh + 0.1);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: mw, h: mh, fill: { color: miss ? C.background1 : C.background2 }, line: { color: miss ? HEX.accent6 : CALLOUT_LINE, width: miss ? 1.25 : 0.75 } });
    s.addText(lab, { x: x + 0.1, y: y + 0.05, w: mw - 0.2, h: 0.3, fontSize: 12, bold: true, color: miss ? C.accent6 : C.text2, margin: 0, isTextBox: true });
    s.addText([{ text: a, options: { fontSize: 20, bold: true, color: miss ? C.accent6 : C.accent1 } }, { text: ` / ${tg}`, options: { fontSize: 12, color: C.accent5 } }],
      { x: x + 0.1, y: y + 0.36, w: mw - 0.2, h: 0.42, valign: "middle", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["정량지표 12개 중 ", false], ["10개 목표 초과", true], [" · 특허 등록·기술실시는 2단계 보완", false]]);
  source(s, C, "출처: 단계보고서 p.12, p.13, p.22, p.36, p.291 · 수치는 실적 / 1단계 목표");
  addNote(s, "f_summary");
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
  const im = fit("p155_piston_parts.jpg", 10.35, 1.55, 2.38, 2.65);
  s.addImage({ path: FIG("p155_piston_parts.jpg"), ...im, altText: "개선 피스톤 부품" });
  s.addText("[개선 피스톤 부품]", { x: 10.35, y: im.y + im.h + 0.03, w: 2.38, h: 0.26, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  const im2 = fit("p133_piston_seal_wear_analysis.jpg", 10.35, 4.55, 2.38, 1.35);
  s.addImage({ path: FIG("p133_piston_seal_wear_analysis.jpg"), ...im2, altText: "피스톤 씰 마모 해석" });
  s.addText("[씰 마모 해석 (p.133)]", { x: 10.35, y: im2.y + im2.h + 0.03, w: 2.38, h: 0.26, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  callout(s, pres, C, [["1단계 목표(150%) 대비 ", false], ["2배 이상 초과 달성", true], [", 연차별 공인기관 성적서 확보", false]]);
  source(s, C, "출처: 단계보고서 p.19~22 공인인증평가 성적서(2024·2025·2026), p.133, p.155");
  addNote(s, "piston");
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
    [1, "다이어프램 교체주기"], [2, [["117 → 492.7 hr(", false], ["4.21배", true], [")", false]]], [2, "최초 적용품 1,025 hr 운전"],
    [1, "개선 릴리프밸브"], [2, "사용 수명 3개월 → 8개월 이상"], [2, "응력 29.3% 개선(p.23)"],
  ], x2 + 2.8, 2.25, w - 2.95, 3.8, 14);
  callout(s, pres, C, [["개발품 상용 충전소 적용 결과 ", false], ["교체주기 4.21배 연장", true], [" 확인", false]]);
  source(s, C, "출처: 단계보고서 p.19, p.23~24 공인인증평가 성적서, p.26 상용충전소 운영 현황 보고서");
  addNote(s, "diaphragm");
}

function sDispenser(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[내구] 충전기: IR 리시버·유량조절밸브 공인 1.5배 달성");
  const items = [
    ["IR 리시버 (MSENG)", "p194_ir_receiver.jpg", [[1, [["15,000 cycle", true]]], [2, "기준 10,000 cycle 대비 150% 공인"], [2, "−40~20 ℃ 온도별 내구시험"], [2, "실차 충전 적용 및 로그 확보"]], C.text2],
    ["유량조절밸브 (MT.H)", "p285_valve_hydrogen_test.jpg", [[1, [["153,030회", true]]], [2, "96.3 MPa 작동시험 150% 공인"], [2, "10개월 실충전 1,672회"], [2, "충전기 제조사 2곳 3기 공급"]], C.text2],
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
  source(s, C, "출처: 단계보고서 p.29~32, p.39 공인인증평가 성적서, p.285(수소 내구시험), p.319");
  addNote(s, "dispenser");
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
  addNote(s, "chiller");
}

function sDiag(pres, C, t, section, noteKey = "diag") {
  const s = contentSlide(pres, C, section, t, "[진단] 운전데이터 기반 조기경보: 실제 고장 92일 전 감지");
  const lw = 7.3;
  panel(s, pres, C, MX, 1.55, lw, 2.2, "압축기 중정비 고장(2023.12.18) 사전 감지");
  const ax = MX + 0.5, aw = lw - 1.7, ay = 3.15;
  s.addShape(pres.shapes.LINE, { x: ax, y: ay, w: aw, h: 0, line: { color: HEX.accent5, width: 2 } });
  s.addShape(pres.shapes.OVAL, { x: ax + aw - 0.14, y: ay - 0.14, w: 0.28, h: 0.28, fill: { color: "C0392B" } });
  s.addText("고장 발생", { x: ax + aw - 0.45, y: ay + 0.2, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: "C0392B", margin: 0, isTextBox: true });
  for (const [d, lab, col] of [[92, "통계 모델", C.accent2], [41, "AI 모델", C.accent4]]) {
    const mx = ax + aw * (1 - d / 100);
    s.addShape(pres.shapes.OVAL, { x: mx - 0.12, y: ay - 0.12, w: 0.24, h: 0.24, fill: { color: col } });
    s.addText([{ text: `D−${d}`, options: { fontSize: 22, bold: true, color: col, breakLine: true } }, { text: lab, options: { fontSize: 12, color: C.text1 } }],
      { x: mx - 0.7, y: 2.1, w: 1.4, h: 0.8, align: "center", margin: 0, isTextBox: true });
  }
  panel(s, pres, C, MX, 3.9, lw, 2.35, "다이어프램 사건 조기경보 (실데이터)");
  const im = fit("p074_diaphragm_early_warning.jpg", MX + 0.15, 4.42, 3.4, 1.78);
  s.addImage({ path: FIG("p074_diaphragm_early_warning.jpg"), ...im, altText: "다이어프램 사건 조기경보 결과" });
  addBullets(s, C, [[1, [["D−6.57일", true], [" 모델 1 첫 경보", false]]], [1, [["D−2.72일", true], [" 모델 2 첫 경보", false]]], [2, "사건 추정 시작 기준"]], MX + 3.75, 4.6, lw - 3.9, 1.5, 15);
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  panel(s, pres, C, rx, 1.55, rw, 4.7, "신뢰성 관리체계 (고등기술연구원)");
  addBullets(s, C, [
    [1, "FMECA 정량화"], [2, [["고장모드 ", false], ["148개", true], [" RPN 산출", false]]], [2, "압축기 35 · 충전기 14 · 냉각기 99"],
    [1, "현장 환류"], [2, "설비별 고장대응 매뉴얼 3종"],
    [1, "결함진단 모델"], [2, "4방식 9종 비교 검증"], [2, "고압 압축기 진동 7개 측정점 Alert/Trip 기준"],
  ], rx + 0.3, 2.15, rw - 0.45, 4.0, 14);
  callout(s, pres, C, [["고장 후 정비  →  ", false], ["상태기반 예측정비(CBM)", true], [" 전환 기반 확보", false]]);
  source(s, C, "출처: 단계보고서 p.12~13, p.70, p.74, p.315");
  addNote(s, noteKey);
}

function sEff1(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[효율] 공인 기준선 확정 및 충전 공정 비효율 원인 규명");
  s.addChart(pres.charts.BAR, [{ name: "SEC", labels: ["기준선\n(KTC 검증 실측)", "최종 목표\n(20% 절감)"], values: [9.5, 7.6] }], {
    x: MX, y: 1.55, w: 3.6, h: 4.15, barDir: "col", chartColors: ["9AA5B1", HEX.dk2], dataLabelFormatCode: "0.0",
    title: "충전소 SEC (kWh/kgH₂)", ...chartBase(), valAxisMinVal: 0, valAxisMaxVal: 11,
  });
  s.addText("정밀 분석 43주 9.1 · 전 기간 9.5\n목표 7.6은 기준선의 20% 절감 환산값", { x: MX, y: 5.75, w: 3.6, h: 0.55, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true });
  const rx = 4.4, rw = W - MX - rx;
  const im = fit("p095_main_chiller_standby.jpg", rx, 1.5, rw, 3.1);
  s.addImage({ path: FIG("p095_main_chiller_standby.jpg"), ...im, altText: "메인 칠러 시간별 전력" });
  s.addText("[메인 칠러 시간별 전력 및 시간대별 평균]", { x: rx, y: im.y + im.h + 0.02, w: rw, h: 0.26, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  const st = [["12.9 kWh/h", "야간(20~08시) 평균\n메인 칠러 전력"], ["38.7%", "칠러 전력량 중\n야간 대기 비중"], ["−40 ℃", "사전 냉각 상시 유지\n(SAE J2601 T40)"]];
  const sw = (rw - 0.4) / 3;
  st.forEach(([big, lab], i) => {
    const x = rx + i * (sw + 0.2);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 4.95, w: sw, h: 1.3, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 } });
    s.addText(big, { x, y: 4.99, w: sw, h: 0.55, fontSize: 24, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(lab, { x, y: 5.55, w: sw, h: 0.65, fontSize: 13, color: C.text1, align: "center", valign: "top", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["차량 수요와 무관한 ", false], ["메인 칠러 상시 대기전력", true], ["이 충전 공정 비효율의 핵심 원인", false]]);
  source(s, C, "출처: 단계보고서 p.12, p.67 KTC 성적서(GT2025-00687), p.95~96");
  addNote(s, "eff1");
}

function sEff2(pres, C, t) {
  const s = contentSlide(pres, C, "1단계 연구 결과", t, "[효율] 하드웨어·운영 투트랙 절감 수단 확보 및 검증 체계 수립");
  const w = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w, 2.3, "Track 1  하드웨어");
  addBullets(s, C, [[2, [["고효율 칠러 충전 kg당 ", false], ["−21.4%", true], ["(KCL)", false]]], [2, "압축기 동력전달 개선(플라이휠, 고효율 벨트)"], [2, "고내구 PCHE"]], MX + 0.3, 2.15, w - 0.5, 1.6, 15);
  panel(s, pres, C, MX + w + 0.4, 1.55, w, 2.3, "Track 2  운영 최적화", C.accent3);
  addBullets(s, C, [[2, "피크·오프피크 차등 운전(T40 / T30)"], [2, "절전 모드: Standby(10분) → Deep Sleep(60분)"], [2, "부하 패턴 기반 시간대별 제어"]], MX + w + 0.7, 2.15, w - 0.5, 1.6, 15);
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
  addNote(s, "eff2");
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
  s.addText("※ 실물은 견학 ④ 지점에서 확인", { x: MX, y: 5.75, w: 8.4, h: 0.4, fontSize: 15, bold: true, color: C.accent3, margin: 0, isTextBox: true });
  callout(s, pres, C, [["핵심부품 4종 ", false], ["공인시험 1.5배 이상", true], [" · 다이어프램·밸브 ", false], ["현장 효과 확인", true]]);
  source(s, C, "출처: 단계보고서 p.19~39 공인인증평가 성적서, p.26, p.319");
  addNote(s, "f_durability");
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
  addNote(s, "f_eff");
}

// 4. 실증 인프라 및 성능지표 -----------------------------------------
async function sInfra(pres, C, t, section, noteKey = "infra") {
  const s = contentSlide(pres, C, section, t, "실증 설비동 조기 착공 및 인허가 완료");
  const lw = 7.0;
  const i1 = fit("p093_construction_photos.jpg", MX, 1.5, lw, 1.6, "left");
  s.addImage({ path: FIG("p093_construction_photos.jpg"), ...i1, altText: "설비동 공사 사진" });
  const y2 = i1.y + i1.h + 0.1;
  const i2 = fit("p360_install_location.jpg", MX, y2, lw, 5.98 - y2, "left");
  s.addImage({ path: FIG("p360_install_location.jpg"), ...i2, altText: "개선 핵심설비 설치 위치(안)" });
  s.addText("[개선 압축기·냉각기·충전기 설치 위치 (p.360)]", { x: MX, y: i2.y + i2.h + 0.02, w: lw, h: 0.26, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  panel(s, pres, C, rx, 1.55, rw, 2.35, "충주 실증 설비동 인허가");
  const checks = ["KGS 변경 기술검토 적합", "건축허가", "고압가스 제조 변경허가", "설비동 기초 타설(2단계 공사 조기 착수)"];
  for (let i = 0; i < 4; i++) {
    const y = 2.12 + i * 0.43;
    s.addImage({ data: await icon("MdCheckCircle", HEX.accent1), x: rx + 0.3, y: y + 0.05, w: 0.28, h: 0.28 });
    s.addText(checks[i], { x: rx + 0.7, y, w: rw - 0.9, h: 0.38, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  }
  panel(s, pres, C, rx, 4.05, rw, 2.2, "설비 제작 및 세종대평 실증", C.accent3);
  addBullets(s, C, [
    [2, "고압 압축기 제작 완료(2026.07)"], [2, "냉동설비 도장·배치, 듀얼 디스펜서 제작 중"],
    [2, [["세종대평 충전기 실증 ", false], ["1년 앞당김", true]]], [2, "KGS 신청(2026.09.29) → 2026.12 준공 예정"],
  ], rx + 0.3, 4.62, rw - 0.45, 1.55, 14);
  callout(s, pres, C, [["2단계 착수와 동시에 ", false], ["개선 설비 설치·실증 가능", true]]);
  source(s, C, "출처: 단계보고서 p.13, p.28, p.93, p.315, p.317, p.360");
  addNote(s, noteKey);
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
  addNote(s, "perf");
}

const GAPS = [
  ["특허 등록 4 / 10건", "출원 14건 중 10건 심사 중", "의견제출·보정 대응 → 2027년 내 등록 전환(예상 [__]건)"],
  ["기술실시 2 / 4회", "수요기업의 현장 실증 결과 요구", "실증 결과 + 특허 연계 기술 패키지 이전"],
  ["충전소 에너지 절감 검증", "기준선 확정 후 운전 로직 개발 중", "2026.12 로직 완료 → 2027년 내 공인검증"],
  ["다이어프램 압축기 전력", "플라이휠 단독 개선 4% 절감", "고효율 벨트(5V) 동력전달 개선 후 재검증"],
];
function sGap(pres, C, t) {
  const s = contentSlide(pres, C, "실증 인프라 및 성능지표", t, "미달 항목 원인 및 보완 계획");
  const tabs = [C.accent2, C.accent3, C.accent4, C.accent4];
  s.addText("미달 항목", { x: MX, y: 1.55, w: 2.9, h: 0.38, fontSize: 14, bold: true, color: C.accent5, align: "center", margin: 0, isTextBox: true });
  s.addText("원인", { x: MX + 3.15, y: 1.55, w: 3.9, h: 0.38, fontSize: 14, bold: true, color: C.accent5, margin: 0, isTextBox: true });
  s.addText("보완 계획", { x: MX + 7.3, y: 1.55, w: 4.7, h: 0.38, fontSize: 14, bold: true, color: C.accent5, margin: 0, isTextBox: true });
  GAPS.forEach(([a, b, c], i) => {
    const y = 2.0 + i * 1.03;
    s.addText(a, { shape: pres.shapes.RECTANGLE, x: MX, y, w: 2.9, h: 0.9, fill: { color: tabs[i] }, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: MX + 2.9, y, w: CW - 2.9, h: 0.9, fill: { color: ROWBG }, line: { color: ROWBG, width: 0 } });
    s.addText(b, { x: MX + 3.15, y, w: 3.9, h: 0.9, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: "✓  ", options: { color: C.accent1, bold: true } }, { text: c, options: { color: C.text1 } }], { x: MX + 7.3, y, w: 4.7, h: 0.9, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["공통 원인: 현장 실증 전 단계  →  ", false], ["2단계 1차년도 내 보완 완료", true], [" 추진", false]]);
  source(s, C, "출처: 단계보고서 4장 p.312~314, 3-3절 p.297");
  addNote(s, "gap");
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
  const tx = MX + 3.4, tw = CW - 3.4;
  GAPS.forEach(([a, , c], i) => {
    const y = 1.55 + i * 1.18;
    s.addText(a, { shape: pres.shapes.RECTANGLE, x: tx, y, w: 3.0, h: 1.05, fill: { color: [C.accent2, C.accent3, C.accent4, C.accent4][i] }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: tx + 3.0, y, w: tw - 3.0, h: 1.05, fill: { color: ROWBG }, line: { color: ROWBG, width: 0 } });
    s.addText([{ text: "✓  ", options: { color: C.accent1, bold: true } }, { text: c }], { x: tx + 3.25, y, w: tw - 3.4, h: 1.05, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  });
  callout(s, pres, C, [["부품 내구성 지표 공인 달성 · 미달 항목 ", false], ["2단계 1차년도 내 보완", true]]);
  source(s, C, "출처: 단계보고서 p.291~293, 4장 p.312~314 · 성능지표 10개 기준");
  addNote(s, "f_perfgap");
}

// 5. 2단계 추진 계획 -------------------------------------------------
const YEARS = [["2027 (4차년도)", [[2, "실증 설비 설치·시운전"], [2, "에너지 절감 공인검증"], [2, "HyEReD 개발"]]], ["2028 (5차년도)", [[2, "장기 실증"], [2, "HyEReD 유효성 검증"], [2, "조기경보 운영절차서"]]], ["2029.1Q (6차년도)", [[2, "K-충전소 운영모델"], [2, "공인 성적서·설계표준안"]]]];
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
  const yw = (CW - 0.4) / 3;
  YEARS.forEach(([hd, bl], i) => {
    const x = MX + i * (yw + 0.2);
    stepHeader(s, pres, C, x, 4.55, yw, hd, [C.text2, C.accent3, C.accent4][i]);
    addBullets(s, C, bl, x + 0.3, 5.18, yw - 0.45, 1.15, 14);
  });
  callout(s, pres, C, [["2단계는 신규 연구가 아닌 ", false], ["1단계 결과물의 현장 검증·확산", true]]);
  source(s, C, "출처: 단계보고서 요약문 p.4~5, 9장 p.357~366");
  addNote(s, "next");
}

function sTarget(pres, C, t) {
  const s = contentSlide(pres, C, "2단계 추진 계획", t, "2단계 정량 성과 목표 및 연구비");
  const H = (x) => ({ text: x, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const kpi = [["특허 등록", "4", "11"], ["SCIE 논문", "14", "4"], ["국내 논문", "8", "5"], ["SW 저작권", "28", "4"], ["기술실시", "2", "2"], ["사업화", "7", "12"],
    ["사업화 매출(백만 원)", "3,280", "8,000"], ["고용 창출", "35", "3"], ["현장시험", "14", "6"], ["시제품", "7", "1"], ["국가인증", "2", "1"]];
  const rows = [[H("성과지표"), H("1단계 실적"), H("2단계 목표")]];
  kpi.forEach(([a, b, c]) => rows.push([{ text: a, options: { color: C.text1 } }, { text: b, options: { align: "center", color: C.accent5 } }, { text: c, options: { align: "center", bold: true, color: C.accent1 } }]));
  s.addTable(rows, { x: MX, y: 1.55, w: 6.2, colW: [2.8, 1.7, 1.7], rowH: 0.385, fontSize: 14, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.1, 0, 0.1], fill: { color: HEX.lt1 } });
  const rx = MX + 6.5, rw = CW - 6.5;
  panel(s, pres, C, rx, 1.55, rw, 4.7, "2단계 연구비 (단위: 천 원)");
  s.addChart(pres.charts.BAR, [{ name: "연구비", labels: ["4차년도\n(2027)", "5차년도\n(2028)", "6차년도\n(2029.1Q)"], values: [3388600, 3393000, 1041100] }], {
    x: rx + 0.15, y: 2.1, w: rw - 0.3, h: 2.9, barDir: "col", chartColors: [HEX.dk2, HEX.accent3, HEX.accent4], dataLabelFormatCode: "#,##0",
    title: "연차별 총 연구비", ...chartBase(), dataLabelFontSize: 12, valAxisMinVal: 0, valAxisMaxVal: 4200000,
  });
  addBullets(s, C, [[2, [["2단계 합계 ", false], ["7,822,700", true], [" (정부 6,700,000 · 기관부담 1,122,700)", false]]], [2, "실증 설비 설치·장기 실증 4·5차년도 집중"]], rx + 0.3, 5.1, rw - 0.45, 1.05, 14);
  callout(s, pres, C, [["실증 이후 ", false], ["사업화 12건 · 매출 8,000백만 원 · 특허 등록 11건", true], [" 목표", false]]);
  source(s, C, "출처: 단계보고서 p.291 정량적 연구개발성과표, 표지 연구개발비 표(연차별)");
  addNote(s, "target");
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
  addNote(s, "risk");
}

function sBiz(pres, C, t) {
  const s = contentSlide(pres, C, "2단계 추진 계획", t, "활용 및 사업화 계획");
  const st = [["3,280", "백만 원", "1단계 사업화 매출 (목표 2,000)"], ["7", "건", "1단계 사업화 (목표 3)"], ["8,000", "백만 원", "2단계 매출 목표 (사업화 12건)"]];
  const w = (CW - 0.4) / 3;
  st.forEach(([n, u, lab], i) => {
    const x = MX + i * (w + 0.2);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w, h: 1.55, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 } });
    s.addText([{ text: n, options: { fontSize: 36, bold: true, color: C.accent1 } }, { text: " " + u, options: { fontSize: 16, color: C.text1 } }],
      { x, y: 1.62, w, h: 0.9, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(lab, { x, y: 2.55, w, h: 0.4, fontSize: 14, color: C.text1, align: "center", margin: 0, isTextBox: true });
  });
  const road = [["부품 공급", [[2, "밸브 개선품 3기 제조사 공급"], [2, "피스톤 씰·다이어프램 적용 확대"]]], ["설비 패키지", [[2, "압축기·냉각기·충전기"], [2, "현장 실증 후 공급"]]], ["운영모델", [[2, "K-충전소 운영절차서"], [2, "경제성·탄소배출 예측"]]]];
  road.forEach(([hd, bl], i) => {
    const x = MX + i * (w + 0.2);
    stepHeader(s, pres, C, x, 3.3, w, hd, [C.text2, C.accent3, C.accent4][i]);
    addBullets(s, C, bl, x + 0.3, 3.95, w - 0.45, 0.85, 15);
  });
  checkRow(s, pres, C, 5.0, 1.2, "시제품 7건", C.accent2, "✓  압축기 2종(GTC·범한산업) · 충전기(MSENG) · 듀얼 디스펜서(제아ENG)\n✓  냉각기(삼정이엔씨) · 유량조절밸브 2종(MT.H)", 2.0);
  callout(s, pres, C, [["기대효과: ", false], ["충전소 셧다운 방지 · OPEX 절감 · 가동률 향상", true]]);
  source(s, C, "출처: 단계보고서 p.6, p.39, p.291, p.300");
  addNote(s, "biz");
}

function fNext(pres, C, t) {
  const s = contentSlide(pres, C, "2단계 추진 계획", t, "2단계 충전소 현장 검증 계획 및 리스크 대응");
  const yw = (CW - 0.4) / 3;
  YEARS.forEach(([hd, bl], i) => {
    const x = MX + i * (yw + 0.2);
    stepHeader(s, pres, C, x, 1.6, yw, hd, [C.text2, C.accent3, C.accent4][i]);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.25, w: yw, h: 1.75, fill: { color: C.background1 }, line: { color: BORDER, width: 0.75 } });
    addBullets(s, C, bl, x + 0.3, 2.4, yw - 0.45, 1.5, 15);
  });
  checkRow(s, pres, C, 4.25, 0.85, "리스크 대응", C.accent3, "✓  공사 조기 착공 · 고압 압축기 제작 완료 · 충주 인허가 완료 · 세종 실증 1년 앞당김");
  checkRow(s, pres, C, 5.2, 0.85, "2단계 목표", C.accent2, "✓  설비 패키지 내구성 1.5배 현장 검증 · 충전소 에너지 20% 절감 공인검증 · K-충전소 운영모델");
  callout(s, pres, C, [["2단계: ", false], ["1단계 결과물의 충전소 현장 검증", true]]);
  source(s, C, "출처: 단계보고서 p.4~5, p.13, p.93, p.317, 9장 p.357~366");
  addNote(s, "f_next");
}

// 6. 결론 ------------------------------------------------------------
function sConclusion(pres, C, t, noteKey = "conclusion") {
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
  addNote(s, noteKey);
}

function sThanks(pres) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "맺음" });
  s.addText("감사합니다", { placeholder: "title" });
  addNote(s, "thanks");
}

function fTour(pres, C, t) {
  const s = contentSlide(pres, C, "결론 및 현장 확인", t, "현장 견학 동선 및 확인 사항");
  const im = fit("p049_process_flow.jpg", MX, 1.5, CW, 2.75);
  s.addImage({ path: FIG("p049_process_flow.jpg"), ...im, altText: "충주바이오그린수소충전소 공정 흐름도" });
  const stops = [["① 모니터링 화면", [[2, "실시간 전력 15채널"], [2, "진동 7개 측정점·조기경보"]]], ["② 메인 칠러·압축기", [[2, "야간 대기전력 병목"], [2, "2단계 교체 대상"]]], ["③ 설비동 공사 현장", [[2, "기초 타설 완료"], [2, "개선 설비 반입 위치"]]], ["④ 개선 부품 실물", [[2, "피스톤 씰·다이어프램"], [2, "IR 리시버·밸브, 성적서"]]]];
  const w = (CW - 0.6) / 4;
  stops.forEach(([hd, bl], i) => {
    const x = MX + i * (w + 0.2);
    stepHeader(s, pres, C, x, 4.4, w, hd, [C.text2, C.accent3, C.accent4, C.accent1][i], 17);
    addBullets(s, C, bl, x + 0.25, 5.05, w - 0.35, 1.2, 14);
  });
  callout(s, pres, C, [["견학 동선: ", false], ["모니터링 → 칠러·압축기 → 공사 현장 → 실물 전시", true]]);
  source(s, C, "공정 흐름도: 단계보고서 p.49 · 견학 동선은 현장 여건에 맞춰 조정 [확인 필요]");
  addNote(s, "f_tour");
}

// Appendix ----------------------------------------------------------
function tableSlide(pres, C, title, sub, header, rows, colW, opts = {}) {
  const s = contentSlide(pres, C, "Appendix", title, sub);
  const H = (x) => ({ text: x, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const body = rows.map((r) => r.map((c, k) => (typeof c === "object" ? c : { text: String(c), options: { color: C.text1, align: (opts.center || []).includes(k) ? "center" : "left" } })));
  s.addTable([header.map(H), ...body], { x: MX, y: opts.y || 1.55, w: CW, colW, rowH: opts.rowH || 0.36, fontSize: opts.fontSize || 13, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.08, 0, 0.08], fill: { color: HEX.lt1 } });
  return s;
}
function appendix(pres, C) {
  const d = pres.addSlide({ masterName: "PLAIN", sectionTitle: "Appendix" });
  d.addText("Appendix", { placeholder: "title" });
  addNote(d, "thanks");
  const hl = (t) => ({ text: t, options: { bold: true, color: C.accent1, align: "center" } });

  // A1 정오표
  let s = tableSlide(pres, C, "[Appendix] 정량 성과 정오표", "보고서 장별 수치 차이 및 발표 기준 수치",
    ["항목", "요약문(p.6) · 4-3절(p.315)", "정량성과표(p.291) · 본문", "발표 기준"],
    [["SCIE 논문", "16건, 320%", "14건(목표 5)", hl("14건")], ["국내 논문", "9건, 112%", "8건(목표 8)", hl("8건")],
     ["SW 저작권", "29건, 580%", "28건(목표 5)", hl("28건")], ["특허 출원", "11건", "14건(p.297 목록)", hl("14건")],
     ["사업화 매출", "3,022백만 원, 151%", "3,280백만 원(목표 2,000)", hl("3,280백만 원")], ["고용 창출", "31명, 194%", "35명(목표 16)", hl("35명")],
     ["총 연구비", "21,029,300천 원(요약문)", "19,617,800천 원(표지 연차별 표)", hl("[확인 필요]")], ["KTC 측정시간", "–", "144시간(p.67) / 4,144시간(성적서)", hl("약 4,100시간")],
     ["칠러 절감률", "약 21%(p.324)", "22%(p.36 캡션) · 78(p.34)", hl("21.4%(실측)")], ["성적서 캡션", "–", "‘기준치 대비 XXX%’(p.32)", hl("수정 필요")]],
    [2.3, 3.6, 4.0, 2.233], { rowH: 0.43, fontSize: 14 });
  callout(s, pres, C, [["공식 수치: ", false], ["2026.10.30 제출본 정량성과표 기준", true], [" · 정오표 별도 제출", false]]);
  source(s, C, "출처: 단계보고서 p.6, p.12, p.32, p.34, p.36, p.67, p.291, p.297, p.315, p.324, 표지");
  addNote(s, "a_errata");

  // A2 특허
  s = tableSlide(pres, C, "[Appendix] 지식재산권(특허) 현황", "특허 출원 14건 · 등록 4건 · 심사 중 10건",
    ["No.", "특허명", "출원인", "출원일", "상태"],
    [["1", "수소 충전 시스템 및 그 제어 방법", "지티씨", "2026.05.07", hl("등록(2026.08.20)")], ["2", "수소 충전 시스템 및 그 제어 방법", "지티씨", "2026.05.07", hl("등록(2026.08.20)")],
     ["3", "듀얼 씰을 포함하는 압축 피스톤", "지티씨", "2025.01.03", hl("등록(2026.04.02)")], ["4", "유량 제어용 고압 유량조절밸브", "MT.H콘트롤밸브", "2025.11.20", hl("등록(2026.04.23)")],
     ["5", "수소 충전소 설비의 조기 이상감지 시스템 및 제어방법", "고등기술연구원", "2025.05.16", "심사 중"], ["6", "수소 생산장치 및 수소 생산방법", "고등기술연구원", "2025.09.23", "심사 중"],
     ["7", "액상연료개질장치 및 이를 포함하는 연료전지설비", "고등기술연구원", "2025.09.26", "심사 중"], ["8", "유기성 폐기물의 열분해 장치 및 그 열분해 방법", "고등기술연구원", "2025.09.29", "심사 중"],
     ["9", "수소 충전소 설비의 이상감지 시스템 및 그 동작방법", "고등기술연구원", "2025.12.31", "심사 중"], ["10", "유체 압축 장치", "범한산업", "2024.06.20", "심사 중"],
     ["11", "수소압축기", "범한산업", "2024.10.07", "심사 중"], ["12", "다이어프램 압축기", "범한산업", "2025.05.26", "심사 중"],
     ["13", "I/O 확장성을 가지는 PCB 기반 산업용 기기 제어보드", "제아이엔지", "2025.08.23", "심사 중"], ["14", "제어보드 I/O 확장성을 고려한 SW 아키텍처 플랫폼", "제아이엔지", "2025.08.23", "심사 중"]],
    [0.6, 6.2, 2.0, 1.4, 1.933], { rowH: 0.33, fontSize: 12, center: [0, 3, 4] });
  source(s, C, "출처: 단계보고서 3-3절 p.297 지식재산권 표");
  addNote(s, "a_patent");

  // A3 논문
  s = contentSlide(pres, C, "Appendix", "[Appendix] SCIE 논문과 과제 연관성", "SCIE 14건 중 수소충전소 설비·운영 직접 관련 7건");
  const pw = (CW - 0.4) / 3;
  const papers = [
    ["충전소 설비·운영 직접 관련 (7건)", C.text2, [[2, "충전소 3년 고장·신뢰성 분석(KJChE)"], [2, "충전 시스템 열유체 모델(KJChE)"], [2, "충전소 에너지 소비·효율 2건(KJChE)"], [2, "충전소 수익성 분석(IJHE)"], [2, "액화수소 충전소 공정(eTransportation)"], [2, "유도전동기 결함 분류(Electronics)"]]],
    ["수소 생산·운송 (2건)", C.accent3, [[2, "액화수소 트레일러 BOG(RSER)"], [2, "바이오가스 수소 생산·CO2 포집(PSEP)"]]],
    ["공정 경제성 등 파생 성과 (5건)", C.accent5, [[2, "플라스틱 업사이클링(Nat. Commun.)"], [2, "바이오매스·바이오연료 2건"], [2, "그린 나프타(JEM)"], [2, "LNG 냉열 활용(Global Challenges)"]]],
  ];
  papers.forEach(([hd, col, bl], i) => {
    const x = MX + i * (pw + 0.2);
    panel(s, pres, C, x, 1.55, pw, 4.7, hd, col);
    addBullets(s, C, bl, x + 0.3, 2.15, pw - 0.45, 4.0, 14);
  });
  callout(s, pres, C, [["직접 관련 7건 중심으로 답변", true], [" · 분류는 발표팀 확인 필요", false]]);
  source(s, C, "출처: 단계보고서 3-3절 p.293 논문 표 (분류는 논문 주제 기준 작성)");
  addNote(s, "a_paper");

  // A4 매출·기술실시
  s = tableSlide(pres, C, "[Appendix] 사업화 매출 및 기술실시 내역", "사업화 매출 3,280백만 원 · 기술실시 2건",
    ["구분", "내용", "기관", "연도", "금액(천 원)"],
    [["매출", "수소충전소 구축 수소압축기 납품(하이리움)", "지티씨", "2025", { text: "108,000", options: { align: "right", color: C.text1 } }],
     ["매출", "수소충전소 구축 수소압축기 납품(하이리움)", "지티씨", "2025", { text: "324,000", options: { align: "right", color: C.text1 } }],
     ["매출", "수소충전소 구축 수소압축기 납품(하이리움) · 연내 수금 예정", "지티씨", "2026", { text: "1,008,000", options: { align: "right", color: C.text1 } }],
     ["매출", "H2 Compressor 수소압축기", "범한산업", "2026", { text: "1,148,400", options: { align: "right", color: C.text1 } }],
     ["매출", "수소충전기용 유량조절밸브 제조사 공급", "MT.H콘트롤밸브", "2024~2026", { text: "692,000", options: { align: "right", color: C.text1 } }],
     [{ text: "합계", options: { bold: true, color: C.text1, fill: { color: C.background2 } } }, { text: "", options: { fill: { color: C.background2 } } }, { text: "", options: { fill: { color: C.background2 } } }, { text: "", options: { fill: { color: C.background2 } } }, { text: "3,280,400", options: { bold: true, align: "right", color: C.accent1, fill: { color: C.background2 } } }],
     ["기술실시", "AI 활용 중대형 선박용 CO2 분리 시스템 구성 기술", "고려대", "2025", { text: "50,600", options: { align: "right", color: C.text1 } }],
     ["기술실시", "바이오가스를 이용한 수소융복합 충전 시스템", "고등기술연구원", "2026", { text: "30,000", options: { align: "right", color: C.text1 } }]],
    [1.3, 6.0, 1.9, 1.3, 1.633], { rowH: 0.48, fontSize: 13, center: [0, 3] });
  callout(s, pres, C, [["압축기·밸브 등 ", false], ["개발 부품 적용 제품 매출", true], [" · 2026년 지티씨 매출 수금 예정", false]], 6.25, 0.6);
  source(s, C, "출처: 단계보고서 p.300~301 사업화 현황·매출 실적·기술실시 표");
  addNote(s, "a_sales");

  // A5 현장 근거 (4.21배, 실증지 변경)
  s = contentSlide(pres, C, "Appendix", "[Appendix] 현장 실증 근거", "다이어프램 4.21배 산출 근거 및 실증지 변경 사유");
  const w2 = (CW - 0.4) / 2;
  panel(s, pres, C, MX, 1.55, w2, 4.7, "다이어프램 교체주기 비교 (가포 충전소 HP-A)");
  const H = (x) => ({ text: x, options: { bold: true, color: C.background1, fill: { color: C.accent3 }, align: "center" } });
  s.addTable([[H("구분"), H("기존품"), H("개발품")],
    ["분석기간", "2025.07.08~09.15", "2025.10.17~2026.07.21"], ["교체이력", "3회", "3회"], ["평균 교체 운전시간", "117시간", "492.7시간(4.21배)"], ["최초 적용", "–", "1,025시간 운전"]]
    .map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: { color: k === 2 ? C.accent1 : C.text1, bold: k === 2, align: k === 0 ? "left" : "center" } }))),
    { x: MX + 0.2, y: 2.2, w: w2 - 0.4, colW: [1.5, (w2 - 1.9) * 0.42, (w2 - 1.9) * 0.58], rowH: 0.5, fontSize: 13, valign: "middle", border: { type: "solid", color: "D5DEE8", pt: 0.75 }, margin: [0, 0.08, 0, 0.08], fill: { color: HEX.lt1 } });
  addBullets(s, C, [[2, "후속 교체주기 단축 확인 → 범한산업과 원인 분석"], [2, "2단계 장기 실증으로 표본 확대"]], MX + 0.3, 4.75, w2 - 0.5, 1.3, 14);
  panel(s, pres, C, MX + w2 + 0.4, 1.55, w2, 4.7, "실증지 변경 (코하이젠)", C.accent3);
  addBullets(s, C, [
    [1, "여수주삼 → 세종대평"], [2, "수소버스 충전횟수 등 실증 운전시간 확보 여건 우수"],
    [1, "일정 영향"], [2, "충전기 실증 2단계 → 1단계로 1년 앞당김"], [2, "KGS 안전영향평가·기술검토 신청(2026.09.29)"], [2, "2026.12 준공·시운전 예정"],
  ], MX + w2 + 0.7, 2.15, w2 - 0.5, 4.0, 14);
  source(s, C, "출처: 단계보고서 p.26, p.28, p.183, p.317");
  addNote(s, "a_field");

  // A6 연구비
  s = tableSlide(pres, C, "[Appendix] 연구개발비 현황", "연차별 연구개발비 (단위: 천 원, 표지 연차별 표 기준)",
    ["단계", "연차", "정부지원", "기관부담 현금", "기관부담 현물", "지방자치단체", "합계"],
    [["1단계", "1차년도(2024)", "2,500,000", "54,800", "497,200", "–", "3,052,000"], ["1단계", "2차년도(2025)", "3,400,000", "95,200", "812,000", "–", "4,307,200"],
     ["1단계", "3차년도(2026)", "3,040,000", "94,500", "801,400", "500,000", "4,435,900"],
     [{ text: "1단계 소계", options: { bold: true, color: C.text1, fill: { color: C.background2 } } }, { text: "", options: { fill: { color: C.background2 } } }, ...["8,940,000", "244,500", "2,110,600", "500,000"].map((x) => ({ text: x, options: { bold: true, align: "right", color: C.text1, fill: { color: C.background2 } } })), { text: "11,795,100", options: { bold: true, align: "right", color: C.accent1, fill: { color: C.background2 } } }],
     ["2단계", "4차년도(2027)", "2,890,000", "49,400", "449,200", "–", "3,388,600"], ["2단계", "5차년도(2028)", "2,940,000", "44,900", "408,100", "–", "3,393,000"],
     ["2단계", "6차년도(2029.1Q)", "870,000", "17,100", "154,000", "–", "1,041,100"],
     [{ text: "2단계 소계", options: { bold: true, color: C.text1, fill: { color: C.background2 } } }, { text: "", options: { fill: { color: C.background2 } } }, ...["6,700,000", "111,400", "1,011,300", "–"].map((x) => ({ text: x, options: { bold: true, align: "right", color: C.text1, fill: { color: C.background2 } } })), { text: "7,822,700", options: { bold: true, align: "right", color: C.accent1, fill: { color: C.background2 } } }]]
      .map((r) => r.map((c, k) => (typeof c === "object" ? c : { text: c, options: { color: C.text1, align: k >= 2 ? "right" : "left" } }))),
    [1.4, 2.2, 1.75, 1.75, 1.75, 1.6, 1.683], { rowH: 0.45, fontSize: 14 });
  callout(s, pres, C, [["기관별 집행률: ", false], ["보고서 7장(p.332~356) 참조 [확인 필요]", true]]);
  source(s, C, "출처: 단계보고서 표지 연구개발비 표, 7장 연구개발비 사용실적 · 요약문 총액(21,029,300천 원)과 차이 확인 필요");
  addNote(s, "a_budget");

  // A7 안전관리
  s = contentSlide(pres, C, "Appendix", "[Appendix] 안전관리 이행 현황", "전 참여기관 안전관리 이행 및 2단계 실증 안전관리 계획");
  panel(s, pres, C, MX, 1.55, (CW - 0.4) / 2, 4.7, "1단계 이행 현황");
  addBullets(s, C, [
    [1, "위험물질 취급"], [2, "안전관리 전문기관 점검·검사 및 조치 이행"],
    [1, "실증형 과제"], [2, "실증연구 평가관리지침·위험성 평가 이행"],
    [1, "기관별 법령 준수"], [2, "산업안전보건법·연구실안전법 준수(전 기관 이행)"], [2, "일상점검·정기점검·안전교육 실시"],
  ], MX + 0.3, 2.15, (CW - 0.4) / 2 - 0.5, 4.0, 14);
  const x7 = MX + (CW - 0.4) / 2 + 0.4;
  panel(s, pres, C, x7, 1.55, (CW - 0.4) / 2, 4.7, "2단계 실증 안전관리 계획 (주관기관)", C.accent3);
  addBullets(s, C, [
    [1, "설비 설치"], [2, "고압튜빙 자체검사·기밀검사, 방폭인증 확인"], [2, "가스·불꽃감지기 설치 및 작동 확인"],
    [1, "시운전"], [2, "질소·수소 Purge, C&E 인터락 검증"],
    [1, "사업 개시 전"], [2, "KGS 완성검사, 안전관리규정 제출"], [2, "안전관리책임자 선임, 배상책임보험 가입"],
  ], x7 + 0.3, 2.15, (CW - 0.4) / 2 - 0.5, 4.0, 14);
  callout(s, pres, C, [["안전관리형 과제: ", false], ["전 기관 이행 및 2단계 설치·시운전 안전관리 계획 수립", true]]);
  source(s, C, "출처: 단계보고서 11장 p.444~445");
  addNote(s, "a_safety");

  // A8~A10 에너지
  const two = (title, sub, a, b, capA, capB, src, key) => {
    const s2 = contentSlide(pres, C, "Appendix", title, sub);
    const w = (CW - 0.3) / 2;
    const ia = fit(a, MX, 1.6, w, 4.6, "center", "top"), ib = fit(b, MX + w + 0.3, 1.6, w, 4.6, "center", "top");
    s2.addImage({ path: FIG(a), ...ia, altText: capA });
    s2.addImage({ path: FIG(b), ...ib, altText: capB });
    const cy = Math.max(ia.y + ia.h, ib.y + ib.h) + 0.08;
    s2.addText(`[${capA}]`, { x: MX, y: cy, w, h: 0.3, fontSize: 12, color: C.text1, align: "center", margin: 0, isTextBox: true });
    s2.addText(`[${capB}]`, { x: MX + w + 0.3, y: cy, w, h: 0.3, fontSize: 12, color: C.text1, align: "center", margin: 0, isTextBox: true });
    source(s2, C, src);
    addNote(s2, key);
  };
  two("[Appendix] 기준선 경계", "기준선 9.5: KTC 검증 계측 체계의 생산·압축·충전 전 공정 실측값",
    "p067_process_share.jpg", "fig15_ktc_verification.png", "공정별 에너지 비중과 SEC 비중 (p.67)", "KTC 성적서 대비 모니터링 데이터 교차검증 (추가 분석)",
    "출처: 단계보고서 p.12, p.67 · 교차검증은 보고서 외 추가 분석", "a_baseline");
  two("[Appendix] 처리량 보정", "처리량 보정 SEC 기반 절감 효과 분리 평가",
    "p068_sec_vs_throughput.jpg", "p095_full_period_sec.jpg", "처리량과 공정별 SEC 상관 (p.68)", "모니터링 전 기간 주간 공정별 SEC (2024.08~2026.08, p.95)",
    "출처: 단계보고서 p.68, p.95, p.366", "a_throughput");
  two("[Appendix] 절감 시뮬레이션", "부하 패턴 기반 절전 시간대 설정 및 1차 보수적 산정",
    "p364_load_heatmap.jpg", "p365_standby_simulation.jpg", "요일·시간대별 충전소 전력 패턴 (p.364)", "메인 칠러 대기전력 절감 시뮬레이션 (추정, p.365)",
    "출처: 단계보고서 p.364~365 · 시뮬레이션은 추정값", "a_sim");
}

// ---------------------------------------------------------------- 빌드
async function finish(pres, name) {
  const f = path.join(OUT, name);
  await pres.writeFile({ fileName: f });
  console.log(name, "pPr 정리", await dedupeParagraphProps(f));
  await applyTheme(f, THEME);
  return f;
}
async function buildMain() {
  const { pres, C } = makePres("본평가판");
  pres.addSection({ title: "표지" }); sCover(pres, C, "main");
  pres.addSection({ title: "목차" }); sToc(pres, C, ["과제 개요", "1단계 추진 실적", "1단계 연구 결과", "실증 인프라 및 성능지표", "2단계 추진 계획", "결론"]);
  pres.addSection({ title: "과제 개요" });
  sBackground(pres, C, "1. 과제 개요 (1/3)"); await sGoal(pres, C, "1. 과제 개요 (2/3)"); sOrg(pres, C, "1. 과제 개요 (3/3)");
  pres.addSection({ title: "1단계 추진 실적" });
  await sSummary(pres, C, "2. 1단계 추진 실적 (1/3)"); sAchieve(pres, C, "2. 1단계 추진 실적 (2/3)"); sDashboard(pres, C, "2. 1단계 추진 실적 (3/3)");
  pres.addSection({ title: "1단계 연구 결과" });
  sPiston(pres, C, "3. 1단계 연구 결과 (1/7)"); sDiaphragm(pres, C, "3. 1단계 연구 결과 (2/7)"); sDispenser(pres, C, "3. 1단계 연구 결과 (3/7)");
  sChiller(pres, C, "3. 1단계 연구 결과 (4/7)"); sDiag(pres, C, "3. 1단계 연구 결과 (5/7)", "1단계 연구 결과");
  sEff1(pres, C, "3. 1단계 연구 결과 (6/7)"); sEff2(pres, C, "3. 1단계 연구 결과 (7/7)");
  pres.addSection({ title: "실증 인프라 및 성능지표" });
  await sInfra(pres, C, "4. 실증 인프라 및 성능지표 (1/3)", "실증 인프라 및 성능지표"); sPerf(pres, C, "4. 실증 인프라 및 성능지표 (2/3)"); sGap(pres, C, "4. 실증 인프라 및 성능지표 (3/3)");
  pres.addSection({ title: "2단계 추진 계획" });
  sNext(pres, C, "5. 2단계 추진 계획 (1/4)"); sTarget(pres, C, "5. 2단계 추진 계획 (2/4)"); sRisk(pres, C, "5. 2단계 추진 계획 (3/4)"); sBiz(pres, C, "5. 2단계 추진 계획 (4/4)");
  pres.addSection({ title: "결론" }); sConclusion(pres, C, "6. 결론");
  pres.addSection({ title: "맺음" }); sThanks(pres);
  pres.addSection({ title: "Appendix" }); appendix(pres, C);
  return finish(pres, "1단계_단계평가_본평가판.pptx");
}
async function buildField() {
  const { pres, C } = makePres("현장평가판");
  pres.addSection({ title: "표지" }); sCover(pres, C, "field");
  pres.addSection({ title: "과제 개요" }); fOverview(pres, C, "1. 과제 개요");
  pres.addSection({ title: "1단계 추진 실적" }); await fSummaryDash(pres, C, "2. 1단계 추진 실적");
  pres.addSection({ title: "1단계 연구 결과" });
  fDurability(pres, C, "3. 1단계 연구 결과 (1/3)"); sDiag(pres, C, "3. 1단계 연구 결과 (2/3)", "1단계 연구 결과", "f_diag"); fEfficiency(pres, C, "3. 1단계 연구 결과 (3/3)");
  pres.addSection({ title: "실증 인프라 및 성능지표" });
  await sInfra(pres, C, "4. 실증 인프라 및 성능지표 (1/2)", "실증 인프라 및 성능지표", "f_infra"); fPerfGap(pres, C, "4. 실증 인프라 및 성능지표 (2/2)");
  pres.addSection({ title: "2단계 추진 계획" }); fNext(pres, C, "5. 2단계 추진 계획");
  pres.addSection({ title: "결론 및 현장 확인" }); sConclusion(pres, C, "6. 결론", "f_conclusion"); fTour(pres, C, "6. 현장 확인 안내");
  pres.addSection({ title: "맺음" }); sThanks(pres);
  return finish(pres, "1단계_단계평가_현장평가판.pptx");
}


// ---------------------------------------------------------------- 현장평가 인쇄물 (A4 가로)
async function buildPrint() {
  const pres = new pptxgen();
  pres.defineLayout({ name: "A4L", width: 11.69, height: 8.27 });
  pres.layout = "A4L";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "1단계 현장평가 인쇄물";
  const C = pres.SchemeColor;
  const PW = 11.69, M = 0.5, IW = PW - 2 * M;
  const head = (s, title, sub) => {
    s.addImage({ path: TPL("logo_blue.png"), x: PW - M - 1.6, y: 0.3, w: 1.6, h: 0.34 });
    s.addText(title, { x: M, y: 0.28, w: 8.5, h: 0.5, fontSize: 24, bold: true, color: C.text1, margin: 0, isTextBox: true });
    s.addImage({ path: TPL("marker.png"), x: M, y: 0.92, w: 0.12, h: 0.21 });
    s.addText(sub, { x: M + 0.2, y: 0.85, w: IW - 0.2, h: 0.36, fontSize: 14, bold: true, color: C.text1, margin: 0, isTextBox: true });
  };
  // 1. 견학 안내
  let s = pres.addSlide();
  head(s, "현장 견학 안내", "수소충전소 핵심설비·부품 내구성 및 효율 향상 · 1단계 현장평가 · 충주바이오그린수소충전소 [확인 필요]");
  const im = fit("p049_process_flow.jpg", M, 1.35, IW, 3.0);
  s.addImage({ path: FIG("p049_process_flow.jpg"), ...im, altText: "공정 흐름도" });
  const stops = [
    ["① 모니터링 화면", "사무동", [[2, "실시간 전력 15채널 계측"], [2, "고압 압축기 진동 7개 측정점"], [2, "조기경보: 실제 고장 92일 전 감지"]]],
    ["② 메인 칠러·압축기", "기존 설비 구역", [[2, "야간에도 −40 ℃ 사전 냉각 유지"], [2, "칠러 전력량 38.7%가 야간 대기"], [2, "2단계 고효율 칠러 교체 대상"]]],
    ["③ 설비동 공사 현장", "증축 부지", [[2, "인허가 3건 완료, 기초 타설"], [2, "신규 압축기·냉각기·충전기 위치"], [2, "2단계 착수와 동시에 반입"]]],
    ["④ 개선 부품 실물", "전시 공간", [[2, "피스톤 씰 1,000 hr(3.3배)"], [2, "다이어프램 교체주기 4.21배"], [2, "IR 리시버·밸브 1.5배, 성적서"]]],
  ];
  const w = (IW - 0.45) / 4;
  stops.forEach(([hd, place, bl], i) => {
    const x = M + i * (w + 0.15);
    s.addText(hd, { shape: pres.shapes.RECTANGLE, x, y: 4.55, w, h: 0.5, fill: { color: [HEX.dk2, HEX.accent3, HEX.accent4, HEX.accent1][i] }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 5.05, w, h: 2.25, fill: { color: C.background1 }, line: { color: BORDER, width: 0.75 } });
    s.addText(place, { x: x + 0.15, y: 5.12, w: w - 0.3, h: 0.3, fontSize: 11, bold: true, color: C.accent5, margin: 0, isTextBox: true });
    s.addText(bullets(C, bl, 12), { x: x + 0.15, y: 5.45, w: w - 0.25, h: 1.8, valign: "top", margin: 0, isTextBox: true });
  });
  s.addText("소요 시간 약 20분 · 안전모·정전기 방지 조치 등 현장 안전수칙 준수 [확인 필요]", { x: M, y: 7.5, w: IW, h: 0.35, fontSize: 12, color: C.accent5, margin: 0, isTextBox: true });

  // 2~3. 실물 전시 카드 (한 장에 2매, 가운데 절취선)
  const cards = [
    { name: "피스톤 씰", org: "GTC · 한국자동차연구원", img: "p155_piston_parts.jpg", before: "300 hr", after: "1,000 hr", rate: "3.3배", check: "공인인증평가 성적서(2024·2025·2026)", note: "A7075 피스톤, 이중 스프링씰 적용 · 특허 등록(듀얼 씰 압축 피스톤)", src: "단계보고서 p.20~22, p.297" },
    { name: "다이어프램 압축기", org: "범한산업 · 창원산업진흥원", img: "p173_diaphragm_compressor_prototype.jpg", before: "교체주기 117 hr", after: "492.7 hr", rate: "4.21배", check: "공인시험 1,561 hr · 가포 충전소 현장 실증", note: "캐비티 형상 최적화(최대응력 −22.6%), PFA 코팅 다이어프램", src: "단계보고서 p.19, p.24, p.26" },
    { name: "충전노즐 IR 리시버", org: "MSENG", img: "p194_ir_receiver.jpg", before: "10,000 cycle", after: "15,000 cycle", rate: "1.5배", check: "공인인증평가 성적서(2026)", note: "−40~20 ℃ 온도별 내구시험, 실차 충전 적용", src: "단계보고서 p.29~30" },
    { name: "유량조절밸브", org: "MT.H콘트롤밸브", img: "p285_valve_hydrogen_test.jpg", before: "102,000회", after: "153,030회", rate: "1.5배", check: "공인인증평가 성적서 · 10개월 실충전 1,672회", note: "96.3 MPa 작동시험, 충전기 제조사 2곳 3기 공급 · 특허 등록", src: "단계보고서 p.37~39, p.319" },
  ];
  for (let pg = 0; pg < 2; pg++) {
    s = pres.addSlide();
    s.addShape(pres.shapes.LINE, { x: PW / 2, y: 0.2, w: 0, h: 7.87, line: { color: "B0B8C0", width: 0.75, dashType: "dash" } });
    s.addText("✂ 절취선", { x: PW / 2 + 0.05, y: 0.15, w: 0.9, h: 0.25, fontSize: 9, color: "B0B8C0", margin: 0, isTextBox: true });
    for (let k = 0; k < 2; k++) {
      const c = cards[pg * 2 + k];
      const x0 = k === 0 ? 0.45 : PW / 2 + 0.45, cw = PW / 2 - 0.9;
      s.addImage({ path: TPL("logo_gray.png"), x: x0 + cw - 1.3, y: 0.4, w: 1.3, h: 0.27 });
      s.addText(c.name, { x: x0, y: 0.35, w: cw - 1.4, h: 0.55, fontSize: 26, bold: true, color: C.text1, margin: 0, isTextBox: true });
      s.addText(c.org, { x: x0, y: 0.92, w: cw, h: 0.32, fontSize: 13, color: C.accent5, margin: 0, isTextBox: true });
      const ii = fit(c.img, x0, 1.4, cw, 2.6);
      s.addImage({ path: FIG(c.img), ...ii, altText: c.name });
      s.addText("개선 전", { shape: pres.shapes.RECTANGLE, x: x0, y: 4.2, w: cw / 2 - 0.1, h: 0.4, fill: { color: "9AA5B1" }, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText("개선 후", { shape: pres.shapes.RECTANGLE, x: x0 + cw / 2 + 0.1, y: 4.2, w: cw / 2 - 0.1, h: 0.4, fill: { color: HEX.dk2 }, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(c.before, { x: x0, y: 4.65, w: cw / 2 - 0.1, h: 0.65, fontSize: 22, bold: true, color: C.accent5, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(c.after, { x: x0 + cw / 2 + 0.1, y: 4.65, w: cw / 2 - 0.1, h: 0.65, fontSize: 22, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText([{ text: "개선율  ", options: { fontSize: 16, bold: true, color: C.text1 } }, { text: c.rate, options: { fontSize: 30, bold: true, color: C.accent1 } }],
        { shape: pres.shapes.RECTANGLE, x: x0, y: 5.45, w: cw, h: 0.75, fill: { color: C.background2 }, line: { color: CALLOUT_LINE, width: 0.75 }, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(bullets(C, [[2, [["검증  ", true], [c.check, false]]], [2, c.note]], 12), { x: x0, y: 6.35, w: cw, h: 1.05, valign: "top", margin: 0, isTextBox: true });
      s.addText(`출처: ${c.src}`, { x: x0, y: 7.55, w: cw, h: 0.25, fontSize: 9, color: C.accent5, margin: 0, isTextBox: true });
    }
  }
  return finish(pres, "1단계_현장평가_인쇄물.pptx");
}

module.exports = { buildMain, buildField, buildPrint, NOTES };
if (require.main === module) {
  (async () => {
    console.log(await buildMain());
    console.log(await buildField());
    console.log(await buildPrint());
  })().catch((e) => { console.error(e); process.exit(1); });
}
