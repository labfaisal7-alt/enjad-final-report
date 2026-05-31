const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "جمعية إنجاد للبحث والإنقاذ";
pptx.company = "إدارة الخدمات الإسعافية";
pptx.subject = "التقرير الختامي لإدارة الخدمات الإسعافية في موسم حج 1447هـ";
pptx.title = "التقرير الختامي لإدارة الخدمات الإسعافية في موسم حج 1447هـ";
pptx.lang = "ar-SA";
pptx.rtlMode = true;
pptx.theme = {
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  lang: "ar-SA",
};

const OUT = path.join(__dirname, "enjad-hajj-1447-final-report-presentation.pptx");
const LOGO = path.join(__dirname, "enjad-logo-crop.png");
const COVER_BG = path.join(__dirname, "enjad-brand-style.png");

const C = {
  red: "9B1F1F",
  red2: "BE1E3D",
  gray: "4A4A4D",
  gray2: "666B73",
  gold: "C79249",
  goldSoft: "F3DFBD",
  beige: "FAF7F0",
  white: "FFFFFF",
  ink: "22272C",
  line: "E4DED3",
};

const W = 13.333;
const H = 7.5;

const cases = [
  ["جروح / تضميد", 226, "22.3%"],
  ["إجهاد عام", 181, "17.8%"],
  ["إجهاد حراري", 165, "16.3%"],
  ["ارتفاع / انخفاض سكر الدم", 114, "11.2%"],
  ["آلام عضلية", 112, "11.0%"],
  ["تسلخات وإصابات احتكاك", 44, "4.3%"],
  ["ارتفاع / انخفاض ضغط الدم", 42, "4.1%"],
  ["جروح قطعية", 33, "3.3%"],
  ["سقوط", 28, "2.8%"],
  ["إغماء", 27, "2.7%"],
  ["ضربة شمس", 22, "2.2%"],
  ["جفاف", 14, "1.4%"],
  ["اختناق", 7, "0.7%"],
];

const sites = [
  ["الجمرات", 337],
  ["عرفات / جبل الرحمة", 179],
  ["أجياد / المنطقة المركزية", 141],
  ["بوابة الملك عبدالعزيز", 115],
];

const procedures = [
  ["تم العلاج بالموقع", 905, "89.2%"],
  ["نقل / توجيه لمركز صحي", 67, "6.6%"],
  ["تسليم للهلال الأحمر", 27, "2.7%"],
  ["رفض النقل", 11, "1.1%"],
  ["حالات حرجة استدعت تدخل متقدم", 5, "0.4%"],
];

function addText(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    margin: opts.margin || 0.06,
    fontFace: "Arial",
    color: opts.color || C.ink,
    fontSize: opts.fontSize || 16,
    bold: opts.bold || false,
    align: opts.align || "right",
    valign: opts.valign || "mid",
    fit: opts.fit || "shrink",
    rtlMode: true,
    breakLine: opts.breakLine,
    bullet: opts.bullet,
    paraSpaceAfterPt: opts.paraSpaceAfterPt || 0,
  });
}

function addTitle(slide, title, subtitle) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.88, fill: { color: C.white, transparency: 4 }, line: { color: C.line, transparency: 30 } });
  slide.addImage({ path: LOGO, x: 11.35, y: 0.12, w: 1.45, h: 0.9 });
  addText(slide, title, 1.1, 0.18, 10.05, 0.34, { fontSize: 19, bold: true, color: C.red });
  if (subtitle) addText(slide, subtitle, 1.1, 0.53, 10.05, 0.22, { fontSize: 9.5, color: C.gray2 });
}

function addFooter(slide, num) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 6.86, w: W, h: 0.64, fill: { color: C.gray }, line: { color: C.gray } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 6.78, w: W, h: 0.08, fill: { color: C.red }, line: { color: C.red } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 6.68, w: 5.2, h: 0.07, fill: { color: C.gold }, line: { color: C.gold } });
  addText(slide, "التقرير الختامي لإدارة الخدمات الإسعافية في موسم حج 1447هـ", 4.7, 6.98, 8.05, 0.28, { fontSize: 10.5, color: C.white });
  addText(slide, String(num).padStart(2, "0"), 0.48, 6.98, 0.5, 0.25, { fontSize: 10, color: C.white, align: "left" });
}

function addBase(slide, title, num, subtitle = "جمعية إنجاد للبحث والإنقاذ – إدارة الخدمات الإسعافية") {
  slide.background = { color: C.white };
  addTitle(slide, title, subtitle);
  addFooter(slide, num);
}

function card(slide, x, y, w, h, label, value, note, accent = C.red) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color: C.beige },
    line: { color: C.line, width: 1 },
  });
  slide.addShape(pptx.ShapeType.rect, { x: x + w - 0.08, y, w: 0.08, h, fill: { color: accent }, line: { color: accent } });
  addText(slide, label, x + 0.16, y + 0.12, w - 0.32, 0.26, { fontSize: 10.5, color: C.gray2 });
  addText(slide, value, x + 0.16, y + 0.44, w - 0.32, 0.46, { fontSize: 25, bold: true, color: accent });
  addText(slide, note, x + 0.16, y + 0.96, w - 0.32, 0.28, { fontSize: 9.2, color: C.gray2 });
}

function sectionLabel(slide, text, x, y, w) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.08, fill: { color: C.gold }, line: { color: C.gold } });
  addText(slide, text, x, y + 0.12, w, 0.34, { fontSize: 17, bold: true, color: C.red });
}

function bulletBox(slide, title, bullets, x, y, w, h, accent = C.red) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.line, width: 1 } });
  slide.addShape(pptx.ShapeType.rect, { x: x + w - 0.08, y, w: 0.08, h, fill: { color: accent }, line: { color: accent } });
  addText(slide, title, x + 0.18, y + 0.16, w - 0.38, 0.34, { fontSize: 15, bold: true, color: accent });
  addText(slide, bullets.map((b) => `• ${b}`).join("\n"), x + 0.18, y + 0.58, w - 0.38, h - 0.72, { fontSize: 11.5, color: C.ink, valign: "top", fit: "shrink" });
}

function barChart(slide, data, x, y, w, h, maxValue, opts = {}) {
  const rowH = h / data.length;
  data.forEach(([label, value, pct], i) => {
    const yy = y + i * rowH;
    addText(slide, label, x + w - 2.25, yy + 0.02, 2.2, rowH - 0.04, { fontSize: opts.fontSize || 9.8, color: C.ink });
    const barW = w - 3.25;
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.72, y: yy + rowH * 0.31, w: barW, h: rowH * 0.32, fill: { color: "EEE8DE" }, line: { color: "EEE8DE" } });
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.72 + barW * (1 - value / maxValue), y: yy + rowH * 0.31, w: barW * (value / maxValue), h: rowH * 0.32, fill: { color: opts.color || C.red }, line: { color: opts.color || C.red } });
    addText(slide, pct ? `${value} | ${pct}` : String(value), x, yy, 0.66, rowH - 0.02, { fontSize: opts.fontSize || 9.5, bold: true, color: opts.color || C.red, align: "left" });
  });
}

function table(slide, rows, x, y, w, rowH, colWidths, header = true) {
  rows.forEach((row, r) => {
    let cx = x;
    row.forEach((cell, c) => {
      const cw = colWidths[c] * w;
      const fill = r === 0 && header ? C.gray : (r % 2 === 0 ? "FBF8F1" : C.white);
      const color = r === 0 && header ? C.white : C.ink;
      slide.addShape(pptx.ShapeType.rect, { x: cx, y: y + r * rowH, w: cw, h: rowH, fill: { color: fill }, line: { color: C.line, width: 0.6 } });
      addText(slide, String(cell), cx + 0.05, y + r * rowH + 0.02, cw - 0.1, rowH - 0.04, { fontSize: r === 0 && header ? 10.2 : 9.2, bold: r === 0 && header, color, align: c === 0 ? "right" : "center" });
      cx += cw;
    });
  });
}

function tryDoughnut(slide, labels, values, x, y, w, h) {
  try {
    slide.addChart(pptx.ChartType.doughnut, [{ name: "الإجراءات", labels, values }], {
      x, y, w, h,
      holeSize: 58,
      showLegend: false,
      showValue: false,
      showTitle: false,
      chartColors: [C.red, C.gold, C.gray2, "D8C7AA", "8D172E"],
      border: { color: C.white, transparency: 100 },
      showCatName: false,
    });
  } catch (e) {
    slide.addShape(pptx.ShapeType.ellipse, { x, y, w, h, fill: { color: C.red }, line: { color: C.red } });
    slide.addShape(pptx.ShapeType.ellipse, { x: x + w * 0.25, y: y + h * 0.25, w: w * 0.5, h: h * 0.5, fill: { color: C.white }, line: { color: C.white } });
  }
}

function cover() {
  const slide = pptx.addSlide();
  slide.addImage({ path: COVER_BG, x: 0, y: 0, w: W, h: H });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.25, w: 11.75, h: 2.25, fill: { color: C.white, transparency: 10 }, line: { color: C.white, transparency: 100 } });
  addText(slide, "التقرير الختامي\nلإدارة الخدمات الإسعافية\nفي موسم حج 1447هـ", 1.25, 1.34, 10.9, 1.55, { fontSize: 32, bold: true, color: C.red, align: "center", fit: "shrink" });
  addText(slide, "جمعية إنجاد للبحث والإنقاذ – إدارة الخدمات الإسعافية", 2.1, 3.03, 9.2, 0.36, { fontSize: 16, color: C.gray, align: "center", bold: true });
  addText(slide, "فترة التقرير: 18 مايو 2026 – 29 مايو 2026", 2.1, 3.45, 9.2, 0.3, { fontSize: 13, color: C.gray2, align: "center" });
  addText(slide, "نسخة عرض تنفيذية", 0.85, 6.8, 2.2, 0.3, { fontSize: 12, color: C.white, align: "left", bold: true });
}

function slideSummary() {
  const slide = pptx.addSlide();
  addBase(slide, "الملخص التنفيذي", 2);
  addText(slide, "يعرض هذا التقرير نتائج تشغيل الخدمات الإسعافية التطوعية خلال موسم حج 1447هـ، مع تحليل لطبيعة الحالات ومناطق الضغط والإجراءات الإسعافية المقدمة، إضافة إلى الجاهزية التوعوية والتوصيات التطويرية.", 0.85, 1.12, 11.8, 0.72, { fontSize: 15, color: C.gray, align: "center" });
  card(slide, 9.55, 2.0, 2.75, 1.35, "إجمالي الحالات المباشرة", "1,015", "حالة إسعافية ميدانية", C.red);
  card(slide, 6.55, 2.0, 2.75, 1.35, "المسعفون المشاركون", "79", "مسعف ومتطوع", C.gold);
  card(slide, 3.55, 2.0, 2.75, 1.35, "أعلى يوم تشغيلي", "28 مايو", "2026م", C.red);
  card(slide, 0.55, 2.0, 2.75, 1.35, "أعلى منطقة تشغيلية", "الجمرات", "337 حالة مباشرة", C.gray);
  card(slide, 5.25, 3.75, 2.85, 1.35, "الجاهزية التوعوية", "3", "محاضرات مكافحة عدوى", C.red2);
  card(slide, 2.05, 3.75, 2.85, 1.35, "المستفيدون من التوعية", "114", "عضو مشارك", C.gold);
  card(slide, 8.45, 3.75, 2.85, 1.35, "العلاج بالموقع", "89.2%", "مؤشر أداء مرتفع", C.red);
}

function slideDashboard() {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addText(slide, "لوحة المؤشرات التنفيذية", 7.45, 0.25, 4.65, 0.38, { fontSize: 22, bold: true, color: C.red });
  slide.addShape(pptx.ShapeType.line, { x: 0.38, y: 0.78, w: 11.9, h: 0, line: { color: C.gold, width: 1.4 } });

  const panelFill = "FDF9F1";
  slide.addShape(pptx.ShapeType.rect, { x: 0.38, y: 1.02, w: 5.1, h: 1.95, fill: { color: panelFill }, line: { color: C.line, width: 1 } });
  addText(slide, "الإجراءات الإسعافية", 3.28, 1.27, 1.85, 0.3, { fontSize: 16, bold: true, color: C.red });
  const procBars = [
    ["علاج بالموقع", "89.2%", 0.892],
    ["نقل / توجيه", "6.6%", 0.066],
    ["هلال أحمر", "2.7%", 0.027],
  ];
  procBars.forEach(([label, pct, ratio], i) => {
    const y = 1.72 + i * 0.43;
    addText(slide, label, 4.08, y - 0.04, 0.95, 0.18, { fontSize: 10.2, color: C.ink });
    slide.addShape(pptx.ShapeType.rect, { x: 1.45, y, w: 2.2, h: 0.12, fill: { color: "EAE4D8" }, line: { color: "EAE4D8" } });
    slide.addShape(pptx.ShapeType.rect, { x: 1.45 + 2.2 * (1 - ratio), y, w: 2.2 * ratio, h: 0.12, fill: { color: C.red }, line: { color: C.red } });
    addText(slide, pct, 0.62, y - 0.08, 0.55, 0.23, { fontSize: 11, bold: true, color: C.red, align: "left" });
  });

  slide.addShape(pptx.ShapeType.rect, { x: 0.38, y: 3.18, w: 5.1, h: 1.95, fill: { color: panelFill }, line: { color: C.line, width: 1 } });
  addText(slide, "ملخص تشغيلي سريع", 2.72, 3.45, 2.28, 0.3, { fontSize: 16, bold: true, color: C.red });
  addText(slide, "الحالات الميدانية تركزت في نطاقات المشي والكثافة البشرية، مع غلبة الحالات البسيطة والمتوسطة، وارتفاع واضح في فعالية المعالجة بالموقع.", 0.82, 3.98, 4.15, 0.78, { fontSize: 13.2, color: C.ink, align: "center", fit: "shrink" });

  slide.addShape(pptx.ShapeType.rect, { x: 5.7, y: 1.02, w: 6.58, h: 4.11, fill: { color: "912A45" }, line: { color: "912A45" } });
  slide.addShape(pptx.ShapeType.rect, { x: 5.7, y: 1.02, w: 6.58, h: 4.11, fill: { color: C.gray, transparency: 38 }, line: { color: "912A45", transparency: 100 } });
  slide.addImage({ path: LOGO, x: 6.18, y: 3.86, w: 2.0, h: 1.25, transparency: 78 });
  addText(slide, "مؤشر الأداء العام", 9.18, 1.36, 2.5, 0.32, { fontSize: 16.5, bold: true, color: C.white, align: "center" });
  addText(slide, "89.2%", 8.32, 1.87, 3.42, 0.72, { fontSize: 45, bold: true, color: C.white, align: "center" });
  addText(slide, "من الحالات تم علاجها بالموقع دون الحاجة إلى نقل أو تسليم لجهة أخرى.", 7.15, 2.66, 4.25, 0.32, { fontSize: 12.4, bold: true, color: C.white, align: "center" });

  const mini = [
    [9.93, 3.08, 1.7, "متوسط الحالات لكل مسعف", "12.8"],
    [7.75, 3.08, 1.9, "أعلى موقع ضغط", "الجمرات"],
    [5.95, 3.08, 1.55, "الحالات الحرارية", "201"],
    [9.93, 4.06, 1.7, "المستفيدون من التوعية", "114"],
  ];
  mini.forEach(([x, y, w, label, value]) => {
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.72, fill: { color: C.white, transparency: 82 }, line: { color: C.white, transparency: 55 } });
    addText(slide, label, x + 0.12, y + 0.12, w - 0.22, 0.16, { fontSize: 8.4, color: C.white, align: "right" });
    addText(slide, value, x + 0.12, y + 0.39, w - 0.22, 0.22, { fontSize: 17, bold: true, color: C.white, align: "right" });
  });
  addFooter(slide, 3);
}

function slideCasesChart() {
  const slide = pptx.addSlide();
  addBase(slide, "توزيع الحالات بعد تصنيفها", 4);
  addText(slide, "يوضح الرسم أكثر أنواع الحالات مباشرة، حيث تصدرت حالات الجروح والتضميد، ثم الإجهاد العام والإجهاد الحراري.", 0.85, 1.05, 11.8, 0.38, { fontSize: 13, color: C.gray, align: "center" });
  barChart(slide, cases.slice(0, 8), 0.8, 1.7, 11.85, 4.55, 226, { color: C.red, fontSize: 11 });
}

function slideCasesTable() {
  const slide = pptx.addSlide();
  addBase(slide, "جدول توزيع الحالات النهائي", 5);
  const rows = [["نوع الحالة", "العدد", "النسبة"], ...cases];
  table(slide, rows.slice(0, 8), 6.75, 1.22, 5.75, 0.48, [0.58, 0.2, 0.22]);
  table(slide, [rows[0], ...rows.slice(8)], 0.8, 1.22, 5.75, 0.48, [0.58, 0.2, 0.22]);
}

function slideGeo() {
  const slide = pptx.addSlide();
  addBase(slide, "التوزيع الجغرافي للحالات", 6);
  sectionLabel(slide, "المناطق التشغيلية", 7.2, 1.15, 5.1);
  card(slide, 9.85, 1.88, 2.4, 1.1, "المشاعر المقدسة", "516", "50.8% تقريباً", C.red);
  card(slide, 7.2, 1.88, 2.4, 1.1, "المنطقة المركزية", "493", "48.6% تقريباً", C.gold);
  card(slide, 8.5, 3.28, 2.4, 1.1, "مواقع أخرى", "6", "0.6% تقريباً", C.gray);
  sectionLabel(slide, "أكثر المواقع مباشرة للحالات", 0.8, 1.15, 5.75);
  barChart(slide, sites, 0.8, 1.82, 5.75, 2.7, 337, { color: C.red, fontSize: 11.5 });
  addText(slide, "تتركز الكثافة التشغيلية في نطاقات المشي والكثافة البشرية، مما يدعم توصية زيادة الفرق ونقاط التدخل في الجمرات وعرفات والمنطقة المركزية.", 1.0, 5.08, 11.3, 0.55, { fontSize: 13.2, color: C.gray, align: "center" });
}

function slideProcedures() {
  const slide = pptx.addSlide();
  addBase(slide, "الإجراءات الإسعافية المقدمة", 7);
  tryDoughnut(slide, procedures.map((p) => p[0]), procedures.map((p) => p[1]), 8.35, 1.48, 3.55, 3.1);
  addText(slide, "89.2%", 8.92, 2.55, 2.35, 0.48, { fontSize: 31, bold: true, color: C.red, align: "center" });
  addText(slide, "تم العلاج بالموقع", 8.95, 3.02, 2.3, 0.28, { fontSize: 11.5, color: C.gray, align: "center" });
  table(slide, [["الإجراء الإسعافي", "العدد", "النسبة"], ...procedures], 0.85, 1.35, 6.8, 0.55, [0.58, 0.2, 0.22]);
  addText(slide, "انخفاض نسب النقل والتسليم يعكس كفاءة الفرز الميداني وفعالية التدخلات الأولية للحالات غير الحرجة.", 1.0, 5.48, 11.2, 0.42, { fontSize: 13, color: C.gray, align: "center" });
}

function slideAnalysis() {
  const slide = pptx.addSlide();
  addBase(slide, "التحليل التشغيلي", 8);
  bulletBox(slide, "طبيعة الحالات", ["إصابات بسيطة ومتوسطة.", "إجهادات ناتجة عن المشي والحرارة.", "مشاكل سكر وضغط مرتبطة بالإجهاد البدني."], 8.55, 1.22, 3.7, 2.05, C.red);
  bulletBox(slide, "الضغط التشغيلي", ["تركز الحالات في الجمرات وعرفات والمنطقة المركزية.", "المواقع الأعلى كثافة ترتبط بمسافات المشي والكثافة البشرية.", "الحاجة إلى توزيع مرن للفرق حسب الذروة."], 4.85, 1.22, 3.45, 2.05, C.gold);
  bulletBox(slide, "كفاءة الفرق", ["معالجة 89.2% من الحالات بالموقع.", "سرعة استجابة وفرز ميداني جيد.", "تقليل الضغط على المنشآت الصحية الرسمية."], 0.85, 1.22, 3.75, 2.05, C.gray);
  bulletBox(slide, "ملاحظات تشغيلية", ["تفاوت جودة التوثيق بين بعض السجلات.", "الحاجة إلى ضبط التصنيفات الطبية وتوحيدها.", "تفاوت الخبرة بين الفرق في التفاصيل الطبية."], 0.85, 3.65, 11.4, 1.85, C.red2);
}

function slideAwareness() {
  const slide = pptx.addSlide();
  addBase(slide, "الجاهزية التوعوية قبل الموسم", 9);
  card(slide, 9.0, 1.45, 2.9, 1.35, "محاضرات مكافحة العدوى", "3", "برنامج توعوي قبل الموسم", C.red);
  card(slide, 5.7, 1.45, 2.9, 1.35, "النمط", "1 + 2", "أونلاين + حضوري", C.gold);
  card(slide, 2.4, 1.45, 2.9, 1.35, "المستفيدون", "114", "عضواً من المشاركين", C.gray);
  bulletBox(slide, "أثر الجاهزية التوعوية", [
    "رفع وعي المسعفين بإجراءات الوقاية الشخصية ومكافحة انتقال العدوى.",
    "تعزيز الالتزام بالممارسات الآمنة أثناء مباشرة الحالات في المواقع عالية الكثافة.",
    "دعم جاهزية الفرق للتعامل مع الحالات الميدانية ضمن بيئة تشغيلية تتطلب انضباطاً صحياً عالياً.",
  ], 1.1, 3.45, 11.1, 1.85, C.red);
}

function slideOperationalRecs() {
  const slide = pptx.addSlide();
  addBase(slide, "التوصيات التشغيلية", 10);
  bulletBox(slide, "أولويات التشغيل", [
    "رفع الكثافة التشغيلية في الجمرات، عرفات / جبل الرحمة، والمنطقة المركزية.",
    "تخصيص فرق للحالات الحرارية والإجهاد.",
    "إنشاء نقاط ثابتة للتدخل السريع قرب مسارات المشي ومناطق الانتظار.",
    "تعزيز مستلزمات التضميد، علاج التسلخات، التبريد، الإماهة، وأجهزة القياس.",
    "اعتماد فرز ميداني أسرع للحالات غير الحرجة.",
  ], 6.95, 1.2, 5.35, 4.6, C.red);
  bulletBox(slide, "إدارة الانتشار اليومي", [
    "توزيع الفرق حسب الوقت والموقع وساعات الذروة.",
    "تفعيل قائد ميداني لكل نطاق لمتابعة الفرق والمخزون وجودة التوثيق.",
    "استمرار محاضرات مكافحة العدوى قبل الموسم.",
    "تنفيذ مراجعة تشغيلية يومية للبيانات وتعديل توزيع الفرق وفق النتائج.",
  ], 0.85, 1.2, 5.75, 4.6, C.gold);
}

function slideAdminRecs() {
  const slide = pptx.addSlide();
  addBase(slide, "التوصيات الإدارية", 11);
  bulletBox(slide, "حوكمة الموارد والتوثيق", [
    "تحسين تصنيفات الأعضاء بين مسعف أولي، مسعف متقدم، ومسعف صحي متخصص.",
    "تفعيل مراجعة يومية للسجلات قبل إغلاق الشفت.",
    "توزيع القوى البشرية حسب البيانات الفعلية للحالات والذروة والموقع.",
    "توحيد مسميات المواقع والتصنيفات والإجراءات الإسعافية.",
  ], 6.95, 1.18, 5.35, 4.7, C.red);
  bulletBox(slide, "مؤشرات الأداء والجودة", [
    "قياس أداء الفرق بمؤشرات: عدد الحالات، جودة التوثيق، سرعة الاستجابة، العلاج بالموقع، البلاغات المكتملة.",
    "إنشاء مسار اعتماد للفرق قبل التشغيل يشمل التسجيل والتصنيف والتدريب واستلام المستلزمات.",
    "تعيين مسؤول جودة للتوثيق لرصد الأخطاء ورفع تقرير مختصر يومي للإدارة.",
  ], 0.85, 1.18, 5.75, 4.7, C.gray);
}

function slideTrainingRecs() {
  const slide = pptx.addSlide();
  addBase(slide, "التوصيات التدريبية", 12);
  card(slide, 9.2, 1.35, 2.55, 1.12, "الفرز الإسعافي", "01", "توحيد تقييم الحالات", C.red);
  card(slide, 6.25, 1.35, 2.55, 1.12, "التوثيق الطبي", "02", "رفع جودة السجلات", C.gold);
  card(slide, 3.3, 1.35, 2.55, 1.12, "العلامات الحيوية", "03", "قياس وتفسير أدق", C.gray);
  card(slide, 0.35, 1.35, 2.55, 1.12, "الإجهاد الحراري", "04", "تدخل أسرع وأكثر أماناً", C.red2);
  bulletBox(slide, "توحيد آلية اتخاذ القرار", [
    "متى يعالج بالموقع.",
    "متى يحول لمركز صحي.",
    "متى يحتاج إلى تدخل متقدم.",
    "كيف يوثق القرار والنتيجة النهائية للحالة.",
  ], 1.25, 3.28, 10.9, 1.95, C.red);
}

function slideTeam() {
  const slide = pptx.addSlide();
  addBase(slide, "فريق الإعداد والمشاركة الإدارية", 13);
  addText(slide, "تم إعداد هذا التقرير بمشاركة فريق إدارة الخدمات الإسعافية بجمعية إنجاد، وبإشراف ومتابعة الكوادر الإدارية والميدانية المشاركة في موسم حج 1447هـ.", 1.05, 1.18, 11.2, 0.62, { fontSize: 15, color: C.gray, align: "center" });

  slide.addShape(pptx.ShapeType.roundRect, { x: 8.35, y: 2.05, w: 3.85, h: 1.25, rectRadius: 0.08, fill: { color: C.beige }, line: { color: C.line, width: 1 } });
  slide.addShape(pptx.ShapeType.rect, { x: 12.08, y: 2.05, w: 0.08, h: 1.25, fill: { color: C.red }, line: { color: C.red } });
  addText(slide, "مدير إدارة الخدمات الإسعافية", 8.6, 2.24, 3.2, 0.22, { fontSize: 12, color: C.gray2 });
  addText(slide, "سلطان علي المطيري", 8.6, 2.62, 3.2, 0.32, { fontSize: 17, bold: true, color: C.red });

  slide.addShape(pptx.ShapeType.roundRect, { x: 4.55, y: 2.05, w: 3.45, h: 1.25, rectRadius: 0.08, fill: { color: C.beige }, line: { color: C.line, width: 1 } });
  slide.addShape(pptx.ShapeType.rect, { x: 7.88, y: 2.05, w: 0.08, h: 1.25, fill: { color: C.gold }, line: { color: C.gold } });
  addText(slide, "التنسيق الميداني", 4.8, 2.24, 2.8, 0.22, { fontSize: 12, color: C.gray2 });
  addText(slide, "فيصل العسيري\nسلمان القحطاني", 4.8, 2.53, 2.8, 0.5, { fontSize: 15, bold: true, color: C.ink });

  slide.addShape(pptx.ShapeType.roundRect, { x: 0.95, y: 2.05, w: 3.25, h: 1.25, rectRadius: 0.08, fill: { color: C.beige }, line: { color: C.line, width: 1 } });
  slide.addShape(pptx.ShapeType.rect, { x: 4.08, y: 2.05, w: 0.08, h: 1.25, fill: { color: C.gray }, line: { color: C.gray } });
  addText(slide, "نطاق المشاركة", 1.2, 2.24, 2.65, 0.22, { fontSize: 12, color: C.gray2 });
  addText(slide, "إدارية وميدانية\nوتوثيقية", 1.2, 2.52, 2.65, 0.52, { fontSize: 15, bold: true, color: C.ink });

  bulletBox(slide, "إعداد ومراجعة التقرير", [
    "فيصل عبدالرحمن العسيري: نائب مدير إدارة الخدمات الإسعافية.",
    "سلمان ناصر القحطاني: عضو إداري.",
    "عبدالله الحليبي: عضو إداري ومنسق إعلامي.",
    "عبدالله الشهري: عضو إداري.",
  ], 1.0, 3.85, 11.25, 1.75, C.red);
}

function slideConclusion() {
  const slide = pptx.addSlide();
  addBase(slide, "الخلاصة النهائية", 14);
  slide.addShape(pptx.ShapeType.roundRect, { x: 1.0, y: 1.32, w: 11.35, h: 4.42, rectRadius: 0.12, fill: { color: C.gray }, line: { color: C.gray } });
  slide.addImage({ path: LOGO, x: 9.95, y: 1.62, w: 1.7, h: 1.08 });
  addText(slide, "أثبتت الفرق الإسعافية التطوعية التابعة لجمعية إنجاد قدرة تشغيلية عالية خلال موسم الحج، وساهمت بفاعلية في مباشرة أكثر من ألف حالة ميدانية، مع تحقيق نسبة مرتفعة من العلاج بالموقع وتقليل الضغط على الجهات الصحية الرسمية.", 1.55, 2.1, 9.85, 0.95, { fontSize: 17, color: C.white, align: "center" });
  addText(slide, "كما دعمت إدارة الخدمات الإسعافية جاهزية المسعفين قبل الموسم من خلال ثلاث محاضرات توعوية حول مكافحة العدوى بالحج، استفاد منها 114 عضواً.", 1.55, 3.22, 9.85, 0.7, { fontSize: 15, color: C.goldSoft, align: "center", bold: true });
  addText(slide, "وترتكز فرص التطوير القادمة على جودة التوثيق، التصنيف الطبي، توزيع القوى البشرية، والتدريب الميداني المتخصص.", 1.55, 4.15, 9.85, 0.72, { fontSize: 15, color: C.white, align: "center" });
}

cover();
slideSummary();
slideDashboard();
slideCasesChart();
slideCasesTable();
slideGeo();
slideProcedures();
slideAnalysis();
slideAwareness();
slideOperationalRecs();
slideAdminRecs();
slideTrainingRecs();
slideTeam();
slideConclusion();

pptx.writeFile({ fileName: OUT });
