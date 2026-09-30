const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');

const A = (f) => path.join(__dirname, 'assets', f);
const OUT = process.argv[2] || path.join(__dirname, 'nafitha-afrad.pptx');

// ---- palette ----
const NAVY = '052977';
const GREEN = '0B9D76';
const INK = '1A2340';
const MUTED = '5B6784';
const TINT = 'EEF2FA';
const TINT_G = 'E7F5F0';
const LINE = 'D6DEEE';
const WHITE = 'FFFFFF';
// placeholder (content still to be supplied) — deliberately off-brand so it cannot be missed
const PH_FILL = 'FFF6E5';
const PH_LINE = 'D97706';
const PH_TEXT = '92400E';
const FONT = 'Cairo';
const SLOGAN = 'مدينتك بين يديك';
const SLOGAN_EN = 'Your City, within Reach';

const W = 13.333, H = 7.5, M = 0.6;

const iconCache = {};
async function icon(name, color, size = 256) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: '#' + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return (iconCache[key] = 'image/png;base64,' + buf.toString('base64'));
}

// shape naming: "anim<step>_<n>[_TODO]" marks shapes for the progressive-build post-process
let nameSeq = 0;
const an = (step, tag = '') => `anim${step == null ? 0 : step}_${nameSeq++}${tag}`;

function T(slide, text, o) {
  const opts = Object.assign(
    { fontFace: FONT, color: INK, fontSize: 16, align: 'right', valign: 'top', rtlMode: true, lang: 'ar-SA', isTextBox: true, margin: 0 },
    o
  );
  if (o.step != null) opts.objectName = an(o.step, o.tag);
  delete opts.step; delete opts.tag;
  slide.addText(text, opts);
}
function S(slide, shape, o) {
  const opts = Object.assign({}, o);
  if (o.step != null) opts.objectName = an(o.step, o.tag);
  delete opts.step; delete opts.tag;
  slide.addShape(shape, opts);
}
function I(slide, o) {
  const opts = Object.assign({ altText: 'نافذة أفراد' }, o);
  if (o.step != null) opts.objectName = an(o.step, o.tag);
  delete opts.step; delete opts.tag;
  slide.addImage(opts);
}
// dashed amber box = content still to be supplied by RCJY / Rqeem
function PH(slide, text, o) {
  const { x, y, w, h, step, fontSize = 12, align = 'right' } = o;
  S(slide, pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: PH_FILL }, line: { color: PH_LINE, width: 1, dashType: 'dash' }, step, tag: '_TODO' });
  T(slide, 'يُستكمل: ' + text, { x: x + 0.15, y, w: w - 0.3, h, fontSize, color: PH_TEXT, valign: 'middle', align, step, tag: '_TODO' });
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.rtlMode = true;
pres.title = 'نافذة أفراد';
pres.company = 'شركة رقيم للحلول الذكية';
pres.theme = { headFontFace: FONT, bodyFontFace: FONT, lang: 'ar-SA' };

const MARK = A('logo_mark.png'); // 330x333
const MARK_NAVY = A('logo_mark_navy.png');
const MARK_WHITE = A('logo_mark_white.png');
const MR = 333 / 330;
// entity logos, in the required order: RCJY, then Jabeen, then Rqeem
const LOGOS = [
  { path: A('logo_rcjy.png'), w: 1141, h: 865, alt: 'شعار الهيئة الملكية للجبيل وينبع' },
  { path: A('logo_jabeen.png'), w: 480, h: 336, alt: 'شعار جبين' },
  { path: A('logo_rqeem.png'), w: 607, h: 415, alt: 'شعار رقيم للحلول الذكية' },
];
const CITY = {
  jubail: { img: A('city_jubail.png'), w: 149, h: 148 },
  raskhair: { img: A('city_raskhair.png'), w: 228, h: 141 },
  yanbu: { img: A('city_yanbu.png'), w: 138, h: 138 },
  jazan: { img: A('city_jazan.png'), w: 132, h: 132 },
};
function cityImg(slide, key, cx, cy, box, step) {
  const c = CITY[key];
  const s = box / Math.max(c.w, c.h);
  const w = c.w * s, h = c.h * s;
  I(slide, { path: c.img, x: cx - w / 2, y: cy - h / 2, w, h, step });
}

// the four service sectors, in presentation order (right to left)
const SECTORS = [
  { key: 'health', name: 'الصحة', icon: 'FaHeartbeat' },
  { key: 'edu', name: 'التعليم', icon: 'FaGraduationCap' },
  { key: 'fun', name: 'الترفيه', icon: 'FaTheaterMasks' },
  { key: 'inv', name: 'الاستثمار', icon: 'FaHandshake' },
];

// shared chrome for content slides
function chrome(slide, n, kicker, title, color = NAVY) {
  slide.background = { color: WHITE };
  // faint logo pattern, bleeding off the top-left corner
  slide.addImage({ path: MARK_NAVY, x: -1.9, y: -1.9, w: 4.6, h: 4.6 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
  T(slide, kicker, { x: 5.2, y: 0.45, w: W - M - 5.2, h: 0.4, fontSize: 14, bold: true, color: color === NAVY ? MUTED : GREEN });
  T(slide, title, { x: 2.6, y: 0.8, w: W - M - 2.6, h: 0.85, fontSize: 30, bold: true, color });
  // brand + slogan + number
  slide.addImage({ path: MARK, x: M, y: 6.83, w: 0.34, h: 0.34 * MR, altText: 'شعار نافذة' });
  T(slide, 'نافذة أفراد  ·  ' + SLOGAN, { x: M + 0.42, y: 6.83, w: 3.2, h: 0.34, fontSize: 11, bold: true, color: NAVY, valign: 'middle', align: 'left' });
  T(slide, String(n), { x: W - M - 0.6, y: 6.83, w: 0.6, h: 0.34, fontSize: 11, color: MUTED, valign: 'middle' });
}

function circleIcon(slide, x, y, d, data, bg, step, pad = 0.26) {
  S(slide, pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: 'none' }, step });
  const s = d * (1 - 2 * pad);
  I(slide, { data, x: x + (d - s) / 2, y: y + (d - s) / 2, w: s, h: s, step });
}

// "example" pill used wherever the persona appears
function examplePill(slide, x, y, step) {
  S(slide, pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.35, h: 0.36, rectRadius: 0.18, fill: { color: WHITE }, line: { color: NAVY, width: 1 }, step });
  T(slide, 'مثال توضيحي', { x, y, w: 1.35, h: 0.36, fontSize: 11, bold: true, color: NAVY, align: 'center', valign: 'middle', step });
}

// persona strip at the bottom of story slides
async function personaStrip(slide, text, step, y = 5.85) {
  S(slide, pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: W - 2 * M, h: 0.72, rectRadius: 0.14, fill: { color: TINT }, line: { type: 'none' }, step });
  circleIcon(slide, W - M - 0.2 - 0.52, y + 0.1, 0.52, await icon('FaUserCircle', WHITE), NAVY, step, 0.14);
  T(slide, text, { x: M + 1.9, y, w: W - 2 * M - 2.75, h: 0.72, fontSize: 16, bold: true, color: NAVY, valign: 'middle', step });
  examplePill(slide, M + 0.25, y + 0.18, step);
}

(async () => {
  // ===== 1. Cover =====
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addImage({ path: MARK_NAVY, x: 9.6, y: -2.2, w: 5.6, h: 5.6 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    s.addImage({ path: MARK_NAVY, x: -2.4, y: 4.2, w: 5.6, h: 5.6 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    const lw = 1.9;
    I(s, { path: MARK, x: (W - lw) / 2, y: 0.55, w: lw, h: lw * MR, step: 1 });
    T(s, 'نافذة أفراد', { x: 2, y: 2.55, w: W - 4, h: 1.05, fontSize: 52, bold: true, color: NAVY, align: 'center', valign: 'middle', step: 2 });
    T(s, SLOGAN, { x: 2, y: 3.6, w: W - 4, h: 0.65, fontSize: 26, bold: true, color: MUTED, align: 'center', valign: 'middle', step: 3 });
    T(s, SLOGAN_EN, { x: 2, y: 4.2, w: W - 4, h: 0.45, fontSize: 15, color: MUTED, align: 'center', valign: 'middle', rtlMode: false, lang: 'en-US', step: 3 });
    // entity logos, right to left: RCJY, then Jabeen, then Rqeem
    const lh = 0.95, lgap = 1.1;
    const lws = LOGOS.map((l) => (lh * l.w) / l.h);
    const total = lws.reduce((a, b) => a + b, 0) + lgap * (LOGOS.length - 1);
    let lx = (W + total) / 2;
    LOGOS.forEach((l, i) => {
      lx -= lws[i];
      I(s, { path: l.path, x: lx, y: 5.6, w: lws[i], h: lh, altText: l.alt, step: 4 });
      if (i < LOGOS.length - 1) S(s, pres.shapes.LINE, { x: lx - lgap / 2, y: 5.7, w: 0, h: lh - 0.2, line: { color: LINE, width: 1 }, step: 4 });
      lx -= lgap;
    });
    s.addNotes('نرحب بكم في حفل تدشين تطبيق نافذة أفراد، وشعاره: مدينتك بين يديك. سنروي اليوم قصة التطبيق من خلال ساكن في مدننا: ما يواجهه اليوم، وكيف تغيّر نافذة يومه. التطبيق من الهيئة الملكية للجبيل وينبع وشركة جبين، ومن تطوير شركة رقيم للحلول الذكية.');
  }

  // ===== 2. The story: meet the resident =====
  {
    const s = pres.addSlide();
    chrome(s, 2, 'القصة', 'لنبدأ بيوم في حياة أحمد');
    // persona card (left)
    const px = M + 0.2, py = 1.95, pw = 3.6, ph = 3.55;
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, rectRadius: 0.2, fill: { color: TINT }, line: { type: 'none' }, step: 1 });
    circleIcon(s, px + (pw - 1.5) / 2, py + 0.35, 1.5, await icon('FaUserCircle', WHITE), NAVY, 1, 0.1);
    T(s, 'أحمد', { x: px + 0.2, y: py + 1.95, w: pw - 0.4, h: 0.55, fontSize: 24, bold: true, color: NAVY, align: 'center', step: 1 });
    T(s, 'يسكن في الجبيل مع أسرته', { x: px + 0.2, y: py + 2.5, w: pw - 0.4, h: 0.4, fontSize: 14, color: MUTED, align: 'center', step: 1 });
    examplePill(s, px + (pw - 1.35) / 2, py + 3.0, 1);
    // needs (right)
    const tx = 4.9, tw = W - M - tx;
    T(s, 'هذا الأسبوع يحتاج أحمد إلى:', { x: tx, y: 1.95, w: tw, h: 0.5, fontSize: 20, bold: true, color: INK, step: 2 });
    const needs = [
      ['FaClinicMedical', 'موعد في عيادة', 'الصحة'],
      ['FaSchool', 'تسجيل ابنه في مدرسة', 'التعليم'],
      ['FaTicketAlt', 'فعالية لأسرته في نهاية الأسبوع', 'الترفيه'],
    ];
    for (let i = 0; i < needs.length; i++) {
      const [ic, t, sec] = needs[i];
      const y = 2.65 + i * 0.95;
      circleIcon(s, W - M - 0.75, y, 0.75, await icon(ic, NAVY), TINT, 3 + i);
      T(s, t, { x: W - M - 0.95 - 4.4, y, w: 4.4, h: 0.75, fontSize: 19, bold: true, color: INK, valign: 'middle', step: 3 + i });
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x: W - M - 0.95 - 4.4 - 1.5, y: y + 0.19, w: 1.3, h: 0.38, rectRadius: 0.19, fill: { color: TINT }, line: { type: 'none' }, step: 3 + i });
      T(s, sec, { x: W - M - 0.95 - 4.4 - 1.5, y: y + 0.19, w: 1.3, h: 0.38, fontSize: 12, bold: true, color: NAVY, align: 'center', valign: 'middle', step: 3 + i });
    }
    // audience question
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 5.75, w: W - 2 * M, h: 0.85, rectRadius: 0.14, fill: { color: NAVY }, line: { type: 'none' }, step: 6 });
    circleIcon(s, W - M - 0.2 - 0.56, 5.75 + 0.145, 0.56, await icon('FaQuestion', NAVY), WHITE, 6, 0.27);
    T(s, 'سؤال للحضور: من أين يبدأ أحمد اليوم؟ وكم تطبيقًا أو موقعًا أو زيارة سيحتاج؟', { x: M + 0.35, y: 5.75, w: W - 2 * M - 1.3, h: 0.85, fontSize: 17, bold: true, color: WHITE, valign: 'middle', step: 6 });
    s.addNotes('لنتعرف على أحمد، وهو مثال توضيحي لساكن في الجبيل مع أسرته. هذا الأسبوع يحتاج إلى ثلاثة أشياء: موعد في عيادة، وتسجيل ابنه في مدرسة، وفعالية لأسرته في نهاية الأسبوع. سؤالي لكم: من أين يبدأ أحمد اليوم؟ وكم تطبيقًا أو موقعًا أو زيارة سيحتاج؟ (امنح الحضور لحظة للإجابة).');
  }

  // ===== 3. Current situation =====
  {
    const s = pres.addSlide();
    chrome(s, 3, 'الوضع الحالي', 'الخدمات اليوم موزعة على قنوات كثيرة');
    const cw = 2.85, gap = 0.243, y0 = 1.85;
    for (let i = 0; i < SECTORS.length; i++) {
      const sec = SECTORS[i];
      const x = W - M - cw - i * (cw + gap);
      const st = i + 1;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: 3.75, rectRadius: 0.16, fill: { color: WHITE }, line: { color: LINE, width: 1 }, step: st });
      circleIcon(s, x + cw - 0.2 - 0.62, y0 + 0.2, 0.62, await icon(sec.icon, WHITE), NAVY, st, 0.24);
      T(s, sec.name, { x: x + 0.2, y: y0 + 0.2, w: cw - 1.1, h: 0.62, fontSize: 20, bold: true, color: NAVY, valign: 'middle', step: st });
      T(s, 'القنوات الحالية', { x: x + 0.2, y: y0 + 1.0, w: cw - 0.4, h: 0.32, fontSize: 12, bold: true, color: MUTED, step: st });
      PH(s, 'القنوات الإلكترونية والحضورية', { x: x + 0.2, y: y0 + 1.35, w: cw - 0.4, h: 0.8, step: st });
      T(s, 'أبرز التحديات', { x: x + 0.2, y: y0 + 2.3, w: cw - 0.4, h: 0.32, fontSize: 12, bold: true, color: MUTED, step: st });
      PH(s, 'تحديان مدعومان ببيانات فعلية', { x: x + 0.2, y: y0 + 2.65, w: cw - 0.4, h: 0.9, step: st });
    }
    await personaStrip(s, 'أحمد اليوم: يتنقل بين أكثر من قناة ليُنجز حاجاته الثلاث', 5);
    s.addNotes('هذا واقع الخدمات اليوم في أربعة مجالات: الصحة والتعليم والترفيه والاستثمار. لكل مجال قنوات إلكترونية وحضورية مختلفة، والتحدي ليس في الخدمة نفسها بل في تفرّق الوصول إليها. [يُستكمل: القنوات الحالية وأبرز التحديات لكل مجال بالأرقام].');
  }

  // ===== 4. The solution =====
  {
    const s = pres.addSlide();
    chrome(s, 4, 'الحل', 'نافذة أفراد: ' + SLOGAN);
    // left visual: scattered channels converge into one window
    const cx = 3.35, cy = 3.85;
    const ch = [
      ['FaGlobe', cx - 2.2, cy - 1.45], ['FaMobileAlt', cx - 2.45, cy + 0.05], ['FaBuilding', cx - 2.2, cy + 1.45],
      ['FaPhoneAlt', cx - 0.9, cy - 1.95], ['FaLaptop', cx - 0.9, cy + 1.95],
    ];
    for (const [ic, x, y] of ch) circleIcon(s, x - 0.33, y - 0.33, 0.66, await icon(ic, MUTED), TINT, 1, 0.26);
    for (const [, x, y] of ch) {
      const sx = x + (cx - x) * 0.28, sy = y + (cy - y) * 0.28, ex = x + (cx - x) * 0.6, ey = y + (cy - y) * 0.6;
      S(s, pres.shapes.LINE, { x: Math.min(sx, ex), y: Math.min(sy, ey), w: Math.abs(ex - sx) || 0.001, h: Math.abs(ey - sy) || 0.001, flipH: ex < sx, flipV: ey < sy, line: { color: MUTED, width: 1, endArrowType: 'triangle' }, step: 2 });
    }
    S(s, pres.shapes.OVAL, { x: cx + 0.1 - 0.95, y: cy - 0.95, w: 1.9, h: 1.9, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: { type: 'outer', color: NAVY, opacity: 0.15, blur: 14, offset: 4, angle: 90 }, step: 3 });
    I(s, { path: MARK, x: cx + 0.1 - 0.62, y: cy - 0.62 * MR, w: 1.24, h: 1.24 * MR, step: 3 });
    // right text
    const tx = 6.4, tw = W - M - tx;
    T(s, 'تطبيق واحد يجمع خدمات المدينة في الصحة والتعليم والترفيه والاستثمار.', { x: tx, y: 2.0, w: tw, h: 1.3, fontSize: 24, bold: true, color: NAVY, step: 4 });
    T(s, 'بدل التنقل بين القنوات، يصل الساكن إلى ما يحتاجه من مكان واحد.', { x: tx, y: 3.4, w: tw, h: 0.9, fontSize: 17, color: INK, step: 5 });
    for (let i = 0; i < SECTORS.length; i++) {
      const cw = 1.38, x = W - M - cw - i * (cw + 0.13);
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.6, w: cw, h: 0.55, rectRadius: 0.27, fill: { color: TINT }, line: { type: 'none' }, step: 6 });
      I(s, { data: await icon(SECTORS[i].icon, NAVY), x: x + cw - 0.42, y: 4.74, w: 0.27, h: 0.27, step: 6 });
      T(s, SECTORS[i].name, { x: x + 0.1, y: 4.6, w: cw - 0.55, h: 0.55, fontSize: 13, bold: true, color: NAVY, valign: 'middle', step: 6 });
    }
    s.addNotes('هنا يأتي دور نافذة أفراد. تطبيق واحد يجمع خدمات المدينة في المجالات الأربعة، فيصل الساكن إلى ما يحتاجه من مكان واحد بدل التنقل بين القنوات. وهذا معنى شعارنا: مدينتك بين يديك.');
  }

  // ===== 5. How Nafithah answers each challenge =====
  {
    const s = pres.addSlide();
    chrome(s, 5, 'الحل', 'كيف تعالج نافذة هذه التحديات؟');
    const answers = [
      'خدمات الصحة في المدينة من التطبيق نفسه',
      'مدارس المدينة وخدماتها في مكان واحد',
      'فعاليات المدينة ووجهاتها في مكان واحد',
      'خدمات المستثمرين تصل إلى السكان موسومة',
    ];
    // column headers
    T(s, 'التحدي', { x: 6.85, y: 1.75, w: 4.2, h: 0.35, fontSize: 13, bold: true, color: MUTED, step: 1 });
    T(s, 'ما تقدمه نافذة', { x: M, y: 1.75, w: 5.3, h: 0.35, fontSize: 13, bold: true, color: MUTED, step: 1 });
    for (let i = 0; i < SECTORS.length; i++) {
      const sec = SECTORS[i], y = 2.15 + i * 0.9, st = i + 1;
      circleIcon(s, W - M - 0.62, y + 0.05, 0.62, await icon(sec.icon, WHITE), NAVY, st, 0.24);
      T(s, sec.name, { x: 11.05, y, w: 1.0, h: 0.72, fontSize: 15, bold: true, color: NAVY, valign: 'middle', step: st });
      PH(s, 'التحدي كما في الشريحة 3', { x: 6.85, y: y + 0.06, w: 4.1, h: 0.6, step: st });
      I(s, { data: await icon('FaArrowLeft', NAVY), x: 6.2, y: y + 0.21, w: 0.3, h: 0.3, step: st });
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x: M, y: y + 0.06, w: 5.35, h: 0.6, rectRadius: 0.1, fill: { color: TINT }, line: { type: 'none' }, step: st });
      T(s, answers[i], { x: M + 0.2, y: y + 0.06, w: 4.95, h: 0.6, fontSize: 14, bold: true, color: INK, valign: 'middle', step: st });
    }
    await personaStrip(s, 'أحمد مع نافذة: حاجاته الثلاث من تطبيق واحد', 5);
    s.addNotes('لكل تحدٍّ ذكرناه إجابة في نافذة. الصحة والتعليم والترفيه في مكان واحد، وخدمات المستثمرين تصل إلى السكان موسومة باسم مقدّمها. وأحمد الذي بدأ يومه بين قنوات كثيرة، ينجز حاجاته الثلاث من تطبيق واحد. [يُستكمل: التحديات من الشريحة 3].');
  }

  // ===== 6. Alignment =====
  {
    const s = pres.addSlide();
    chrome(s, 6, 'التوافق', 'نافذة تدعم توجهات المملكة والهيئة');
    const cols = [
      ['FaFlag', 'رؤية المملكة 2030', ['جودة الحياة: وصول أسهل إلى خدمات المدينة وترفيهها', 'تجربة أفضل للمستفيد من الخدمات'], null],
      ['FaLaptop', 'هيئة الحكومة الرقمية', ['توحيد القنوات في تطبيق واحد بدل تعددها', 'تجربة متكاملة تتمحور حول المستفيد'], 'المرجع الرسمي من وثائق الهيئة'],
      ['FaIndustry', 'الهيئة الملكية: صناعة وحياة', ['المدن الصناعية بيئة متكاملة للعيش والعمل والاستثمار', 'نافذة تخدم جانب الحياة وتربطه بالاستثمار'], 'الصياغة الرسمية للاستراتيجية'],
    ];
    const cw = 3.85, gap = 0.29, y0 = 1.9, chh = 4.55;
    for (let i = 0; i < cols.length; i++) {
      const [ic, head, pts, todo] = cols[i];
      const x = W - M - cw - i * (cw + gap), st = i + 1;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: chh, rectRadius: 0.18, fill: { color: TINT }, line: { type: 'none' }, step: st });
      circleIcon(s, x + cw - 0.3 - 0.85, y0 + 0.3, 0.85, await icon(ic, WHITE), NAVY, st, 0.26);
      T(s, head, { x: x + 0.3, y: y0 + 1.3, w: cw - 0.6, h: 0.55, fontSize: 19, bold: true, color: NAVY, step: st });
      for (let j = 0; j < pts.length; j++) {
        T(s, pts[j], { x: x + 0.3, y: y0 + 1.95 + j * 0.8, w: cw - 0.6, h: 0.72, fontSize: 14, color: INK, step: st });
      }
      if (todo) PH(s, todo, { x: x + 0.3, y: y0 + chh - 0.8, w: cw - 0.6, h: 0.52, step: st, fontSize: 11 });
    }
    s.addNotes('نافذة تنسجم مع ثلاثة توجهات. رؤية المملكة 2030 وجودة الحياة، وتوجه هيئة الحكومة الرقمية نحو توحيد القنوات وتجربة متكاملة للمستفيد، واستراتيجية الهيئة الملكية صناعة وحياة التي ترى المدن الصناعية بيئة متكاملة للعيش والعمل والاستثمار. [يُستكمل: الصياغة الرسمية لاستراتيجية صناعة وحياة ومرجع الحكومة الرقمية].');
  }

  // ===== 7-10. What the app offers, one slide per sector =====
  const offer = [
    { sec: SECTORS[0], n: 7, title: 'الصحة: ما تقدمه نافذة', persona: 'أحمد يجد خدمة العيادة التي يحتاجها من التطبيق', known: [],
      notes: 'في الصحة تجمع نافذة خدمات المدينة الصحية في مكان واحد. [يُستكمل: الخدمات الصحية الفعلية في التطبيق]. وهنا يجد أحمد خدمة العيادة التي يحتاجها دون أن يغادر التطبيق.' },
    { sec: SECTORS[1], n: 8, title: 'التعليم: ما تقدمه نافذة', persona: 'أحمد يجد مدارس المدينة وخدماتها في مكان واحد', known: [],
      notes: 'في التعليم تجمع نافذة مدارس المدينة وخدماتها. [يُستكمل: الخدمات التعليمية الفعلية في التطبيق]. وأحمد يجد ما يحتاجه لتسجيل ابنه في مكان واحد.' },
    { sec: SECTORS[2], n: 9, title: 'الترفيه: ما تقدمه نافذة', persona: 'أحمد يختار فعالية نهاية الأسبوع لأسرته من التطبيق', known: [],
      notes: 'في الترفيه تعرض نافذة فعاليات المدينة ووجهاتها. [يُستكمل: خدمات الترفيه الفعلية في التطبيق]. وأحمد يختار فعالية نهاية الأسبوع لأسرته من التطبيق نفسه.' },
    { sec: SECTORS[3], n: 10, title: 'الاستثمار: ما تقدمه نافذة', persona: 'أحمد يرى خدمات المستثمرين في مدينته موسومة باسم مقدّمها',
      known: [['FaTag', 'خدمات موسومة', 'خدمات المستثمرين تظهر للسكان موسومة باسم مقدّمها']],
      notes: 'في الاستثمار تربط نافذة السكان بخدمات المستثمرين في المدينة، وتظهر كل خدمة موسومة باسم مقدّمها. [يُستكمل: بقية خدمات الاستثمار في التطبيق]. وهكذا يرى أحمد خدمات المستثمرين في مدينته ويصل إليها مباشرة.' },
  ];
  for (const o of offer) {
    const s = pres.addSlide();
    chrome(s, o.n, 'ما تقدمه نافذة أفراد', o.title);
    // left: big sector icon
    S(s, pres.shapes.OVAL, { x: M + 0.3, y: 1.95, w: 3.3, h: 3.3, fill: { color: TINT }, line: { type: 'none' }, step: 1 });
    I(s, { data: await icon(o.sec.icon, NAVY), x: M + 0.3 + 0.95, y: 1.95 + 0.95, w: 1.4, h: 1.4, step: 1 });
    // right: 2x2 feature grid (known features first, then placeholders)
    const gx = 4.75, cw = (W - M - gx - 0.3) / 2, chh = 1.6;
    for (let k = 0; k < 4; k++) {
      const col = k % 2, row = Math.floor(k / 2);
      const x = W - M - cw - col * (cw + 0.3), y = 1.9 + row * (chh + 0.3), st = 2 + k;
      if (k < o.known.length) {
        const [ic, t, d] = o.known[k];
        S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: chh, rectRadius: 0.16, fill: { color: WHITE }, line: { color: LINE, width: 1 }, step: st });
        circleIcon(s, x + cw - 0.25 - 0.6, y + 0.2, 0.6, await icon(ic, WHITE), NAVY, st, 0.26);
        T(s, t, { x: x + 0.25, y: y + 0.25, w: cw - 1.35, h: 0.5, fontSize: 18, bold: true, color: NAVY, step: st });
        T(s, d, { x: x + 0.25, y: y + 0.95, w: cw - 0.5, h: 0.5, fontSize: 13, color: INK, step: st });
      } else {
        PH(s, `الخدمة ${k + 1}: اسمها ووصفها في سطر واحد`, { x, y, w: cw, h: chh, step: st, fontSize: 13 });
      }
    }
    await personaStrip(s, o.persona, 6);
    s.addNotes(o.notes);
  }

  // ===== 11. Identity: windows of four cities -> one window =====
  {
    const s = pres.addSlide();
    chrome(s, 11, 'الهوية', 'من نوافذ المدن الأربع إلى نافذة واحدة');
    const cities = [
      ['jubail', 'الجبيل', 'عمارة الساحل الشرقي'],
      ['raskhair', 'رأس الخير', 'عمارة الساحل الشرقي'],
      ['yanbu', 'ينبع', 'روشان من ينبع'],
      ['jazan', 'جازان', 'عمارة ساحل تهامة'],
    ];
    const cw = 1.95, gap = 0.28, y0 = 1.85, phh = cw * 6 / 5;
    for (let i = 0; i < cities.length; i++) {
      const [k, name, cap] = cities[i];
      const x = W - M - cw - i * (cw + gap), st = i + 1;
      I(s, { path: A(`photo_${k}.jpg`), x, y: y0, w: cw, h: phh, altText: cap, step: st });
      const d = 0.9;
      S(s, pres.shapes.OVAL, { x: x + (cw - d) / 2, y: y0 + phh - d / 2, w: d, h: d, fill: { color: WHITE }, line: { color: LINE, width: 1 }, step: st });
      cityImg(s, k, x + cw / 2, y0 + phh, 0.56, st);
      T(s, name, { x, y: y0 + phh + 0.55, w: cw, h: 0.45, fontSize: 18, bold: true, color: NAVY, align: 'center', step: st });
      T(s, cap, { x, y: y0 + phh + 1.0, w: cw, h: 0.32, fontSize: 11, color: MUTED, align: 'center', step: st });
    }
    // left: result
    const lx = M, lw = 3.05;
    I(s, { data: await icon('FaArrowLeft', MUTED), x: lx + lw + 0.1, y: y0 + phh / 2 - 0.15, w: 0.3, h: 0.3, step: 5 });
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: y0, w: lw, h: 4.05, rectRadius: 0.18, fill: { color: TINT }, line: { type: 'none' }, step: 5 });
    I(s, { path: MARK, x: lx + (lw - 1.55) / 2, y: y0 + 0.3, w: 1.55, h: 1.55 * MR, step: 5 });
    T(s, 'نافذة', { x: lx + 0.2, y: y0 + 2.0, w: lw - 0.4, h: 0.5, fontSize: 20, bold: true, color: NAVY, align: 'center', step: 5 });
    T(s, 'نمط جديد وُلد من المدن الأربع، في مركزه نافذة مفتوحة إلى المدينة', { x: lx + 0.25, y: y0 + 2.55, w: lw - 0.5, h: 1.2, fontSize: 13, color: INK, align: 'center', step: 5 });
    T(s, 'الصور: دليل العمارة السعودية', { x: 4.2, y: 6.35, w: W - M - 4.2, h: 0.3, fontSize: 10, color: MUTED, step: 5 });
    s.addNotes('استلهمنا هوية نافذة من نوافذ المدن الأربع وزخارفها: الساحل الشرقي للجبيل ورأس الخير، والروشان الحجازي في ينبع، وعمارة تهامة في جازان. بسّطنا عنصرًا من كل مدينة ودمجناها في نمط واحد جديد، في مركزه مُعيّن مفتوح يمثل النافذة إلى المدينة.');
  }

  // ===== 12. Identity: colours and slogan =====
  {
    const s = pres.addSlide();
    chrome(s, 12, 'الهوية', 'الألوان والشعار');
    // right: colours from the RCJY logo
    const rx = 6.9, rw = W - M - rx;
    T(s, 'ألوان من شعار الهيئة الملكية', { x: rx + 1.35, y: 1.95, w: rw - 1.35, h: 0.5, fontSize: 19, bold: true, color: INK, step: 1 });
    const rl = LOGOS[0], rh = 0.85, rlw = (rh * rl.w) / rl.h;
    I(s, { path: rl.path, x: rx, y: 1.8, w: rlw, h: rh, altText: rl.alt, step: 1 });
    const sw = [[NAVY, 'نافذة أفراد', '#052977'], [GREEN, 'نافذة أعمال', '#0B9D76']];
    for (let i = 0; i < 2; i++) {
      const [c, n, hex] = sw[i], y = 2.9 + i * 1.45, st = 2 + i;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x: W - M - 1.2, y, w: 1.2, h: 1.2, rectRadius: 0.18, fill: { color: c }, line: { type: 'none' }, step: st });
      T(s, n, { x: rx, y: y + 0.15, w: rw - 1.5, h: 0.5, fontSize: 20, bold: true, color: c, step: st });
      T(s, hex, { x: rx, y: y + 0.65, w: rw - 1.5, h: 0.4, fontSize: 15, color: MUTED, step: st });
    }
    // left: slogan card
    const lx = M, lw = 5.6;
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: 1.9, w: lw, h: 3.9, rectRadius: 0.2, fill: { color: NAVY }, line: { type: 'none' }, step: 4 });
    I(s, { path: MARK_WHITE, x: lx + 0.3, y: 1.9 + 3.9 - 1.9, w: 1.6, h: 1.6 * MR, transparency: 80, step: 4 });
    T(s, 'الشعار اللفظي', { x: lx + 0.4, y: 2.25, w: lw - 0.8, h: 0.4, fontSize: 14, bold: true, color: 'C9D3EA', step: 4 });
    T(s, SLOGAN, { x: lx + 0.4, y: 2.8, w: lw - 0.8, h: 0.9, fontSize: 36, bold: true, color: WHITE, step: 4 });
    T(s, SLOGAN_EN, { x: lx + 0.4, y: 3.7, w: lw - 0.8, h: 0.5, fontSize: 17, color: WHITE, align: 'right', rtlMode: false, lang: 'en-US', step: 4 });
    s.addNotes('الألوان مأخوذة من شعار الهيئة الملكية للجبيل وينبع: الأزرق الداكن لنافذة أفراد، والأخضر لنافذة أعمال. والشعار اللفظي يلخص الفكرة كلها: مدينتك بين يديك.');
  }

  // ===== 13. Roadmap =====
  {
    const s = pres.addSlide();
    chrome(s, 13, 'خارطة الطريق', 'من نافذة أفراد إلى نافذة أعمال');
    const stages = [
      [NAVY, TINT, 'المرحلة الأولى', 'نافذة أفراد', 'الإطلاق الحالي', 'للسكان والزوار', 'FaFlag'],
      [GREEN, TINT_G, 'المرحلة التالية', 'نافذة أعمال', 'قادمة', 'للمستثمرين وأصحاب الأعمال', 'FaBriefcase'],
    ];
    const cw = 5.0, ch = 3.6, y0 = 2.1;
    const xs = [W - M - 0.4 - cw, M + 0.4];
    S(s, pres.shapes.LINE, { x: xs[1] + cw + 0.25, y: y0 + ch / 2, w: xs[0] - xs[1] - cw - 0.5, h: 0, flipH: true, line: { color: MUTED, width: 2, endArrowType: 'triangle' }, step: 2 });
    for (let i = 0; i < 2; i++) {
      const [c, bg, lbl, name, status, who, ic] = stages[i];
      const x = xs[i], st = i === 0 ? 1 : 3;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: ch, rectRadius: 0.2, fill: { color: bg }, line: { type: 'none' }, step: st });
      circleIcon(s, x + cw - 0.4 - 1.0, y0 + 0.4, 1.0, await icon(ic, WHITE), c, st, 0.27);
      T(s, lbl, { x: x + 0.4, y: y0 + 0.45, w: cw - 2.0, h: 0.4, fontSize: 14, bold: true, color: MUTED, step: st });
      T(s, name, { x: x + 0.4, y: y0 + 1.6, w: cw - 0.8, h: 0.7, fontSize: 30, bold: true, color: c, step: st });
      T(s, who, { x: x + 0.4, y: y0 + 2.3, w: cw - 0.8, h: 0.45, fontSize: 16, color: INK, step: st });
      const pw = 1.9;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x: x + cw - 0.4 - pw, y: y0 + ch - 0.8, w: pw, h: 0.5, rectRadius: 0.25, fill: { color: c }, line: { type: 'none' }, step: st });
      T(s, status, { x: x + cw - 0.4 - pw, y: y0 + ch - 0.8, w: pw, h: 0.5, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle', step: st });
    }
    s.addNotes('نطلق اليوم المرحلة الأولى: نافذة أفراد للسكان والزوار. والمرحلة التالية نافذة أعمال للمستثمرين وأصحاب الأعمال.');
  }

  // ===== 14. Nafithah Business: the connection =====
  {
    const s = pres.addSlide();
    chrome(s, 14, 'نافذة أعمال', 'نافذة أعمال: تربط المستثمرين بالسكان', GREEN);
    T(s, 'كل ما يخص المستثمرين وخدماتهم في نافذة أعمال، وخدماتهم تصل إلى السكان عبر نافذة أفراد.', { x: M, y: 1.85, w: W - 2 * M, h: 0.55, fontSize: 18, color: INK, step: 1 });
    const nodes = [
      ['FaBriefcase', 'المستثمر', 'يقدّم خدماته', GREEN, TINT_G, 'icon'],
      [null, 'نافذة أعمال', 'يدير فيها خدماته وكل ما يخصه', GREEN, TINT_G, 'mark'],
      [null, 'نافذة أفراد', 'تظهر فيها خدماته موسومة باسمه', NAVY, TINT, 'mark'],
      ['FaUsers', 'الساكن', 'يصل إلى الخدمة مباشرة', NAVY, TINT, 'icon'],
    ];
    const nw = 2.55, gap = (W - 2 * M - 4 * nw) / 3, y0 = 2.6, nh = 2.75;
    for (let i = 0; i < nodes.length; i++) {
      const [ic, t, d, c, bg, kind] = nodes[i];
      const x = W - M - nw - i * (nw + gap), st = 2 + i;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: nw, h: nh, rectRadius: 0.18, fill: { color: bg }, line: { type: 'none' }, step: st });
      if (kind === 'icon') {
        circleIcon(s, x + (nw - 1.05) / 2, y0 + 0.3, 1.05, await icon(ic, WHITE), c, st, 0.27);
      } else {
        S(s, pres.shapes.OVAL, { x: x + (nw - 1.05) / 2, y: y0 + 0.3, w: 1.05, h: 1.05, fill: { color: WHITE }, line: { type: 'none' }, step: st });
        I(s, { path: MARK, x: x + (nw - 0.7) / 2, y: y0 + 0.3 + (1.05 - 0.7 * MR) / 2, w: 0.7, h: 0.7 * MR, step: st });
      }
      T(s, t, { x: x + 0.15, y: y0 + 1.5, w: nw - 0.3, h: 0.5, fontSize: 19, bold: true, color: c, align: 'center', step: st });
      T(s, d, { x: x + 0.2, y: y0 + 2.0, w: nw - 0.4, h: 0.6, fontSize: 13, color: INK, align: 'center', step: st });
      if (i < nodes.length - 1) I(s, { data: await icon('FaArrowLeft', MUTED), x: x - gap / 2 - 0.15, y: y0 + nh / 2 - 0.15, w: 0.3, h: 0.3, step: st });
    }
    await personaStrip(s, 'أحمد يرى خدمة المستثمر في نافذة أفراد ويصل إليها مباشرة', 6);
    s.addNotes('نافذة أعمال هي المرحلة التالية. فيها كل ما يخص المستثمرين وخدماتهم، ومنها تظهر خدماتهم في نافذة أفراد موسومة باسم مقدّمها. وبهذا تصل خدمة المستثمر إلى الساكن مباشرة، ويصبح التطبيقان جسرًا بين الاستثمار والحياة في مدننا.');
  }

  // ===== 15. Closing =====
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addImage({ path: MARK_NAVY, x: 9.3, y: -2.3, w: 6.0, h: 6.0 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    s.addImage({ path: MARK_NAVY, x: -2.2, y: 4.3, w: 5.4, h: 5.4 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    const x = 5.0, w = W - M - 0.4 - x;
    I(s, { path: MARK, x: W - M - 0.4 - 1.2, y: 0.9, w: 1.2, h: 1.2 * MR, step: 1 });
    // persona resolution
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: W - M - 0.4 - 5.3, y: 2.35, w: 5.3, h: 0.55, rectRadius: 0.27, fill: { color: TINT }, line: { type: 'none' }, step: 2 });
    T(s, 'أحمد: ثلاث حاجات، وتطبيق واحد', { x: W - M - 0.4 - 5.3 + 0.2, y: 2.35, w: 5.3 - 0.4, h: 0.55, fontSize: 15, bold: true, color: NAVY, valign: 'middle', step: 2 });
    T(s, 'نافذة أفراد', { x, y: 3.1, w, h: 0.8, fontSize: 36, bold: true, color: NAVY, step: 3 });
    T(s, SLOGAN, { x, y: 3.9, w, h: 0.7, fontSize: 28, bold: true, color: MUTED, step: 3 });
    T(s, 'شكرًا لكم', { x, y: 4.75, w, h: 0.5, fontSize: 18, color: INK, step: 4 });
    // QR slot
    PH(s, 'رمز QR لرابط تحميل التطبيق', { x: 1.2, y: 1.55, w: 3.0, h: 3.0, step: 3, fontSize: 14, align: 'center' });
    T(s, 'امسح الرمز لتحميل التطبيق', { x: 0.9, y: 4.75, w: 3.6, h: 0.45, fontSize: 15, bold: true, color: NAVY, align: 'center', step: 3 });
    const ql = LOGOS[2], qh = 0.7, qw = (qh * ql.w) / ql.h;
    I(s, { path: ql.path, x: (W - qw) / 2, y: 5.75, w: qw, h: qh, altText: ql.alt, step: 5 });
    T(s, 'تطوير شركة رقيم للحلول الذكية', { x: 3, y: 6.55, w: W - 6, h: 0.4, fontSize: 14, color: MUTED, align: 'center', valign: 'middle', step: 5 });
    s.addNotes('بدأنا بأحمد وحاجاته الثلاث، وانتهينا بتطبيق واحد يلبيها. بهذا نعلن تدشين نافذة أفراد: مدينتك بين يديك. يمكنكم تحميله الآن بمسح الرمز على الشاشة. شكرًا للهيئة الملكية للجبيل وينبع وشركة جبين، والتطبيق من تطوير شركة رقيم للحلول الذكية.');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
