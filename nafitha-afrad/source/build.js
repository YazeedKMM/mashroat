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
const FONT = 'Cairo';

const W = 13.333, H = 7.5, M = 0.6;

const iconCache = {};
async function icon(name, color, size = 256) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: '#' + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return (iconCache[key] = 'image/png;base64,' + buf.toString('base64'));
}

// shape naming: "anim<step>_<n>" marks shapes for the progressive-build post-process
let nameSeq = 0;
const an = (step) => (step == null ? undefined : `anim${step}_${nameSeq++}`);

function T(slide, text, o) {
  const opts = Object.assign(
    { fontFace: FONT, color: INK, fontSize: 16, align: 'right', valign: 'top', rtlMode: true, lang: 'ar-SA', isTextBox: true, margin: 0 },
    o
  );
  if (o.step != null) opts.objectName = an(o.step);
  delete opts.step;
  slide.addText(text, opts);
}
function S(slide, shape, o) {
  const opts = Object.assign({}, o);
  if (o.step != null) opts.objectName = an(o.step);
  delete opts.step;
  slide.addShape(shape, opts);
}
function I(slide, o) {
  const opts = Object.assign({ altText: 'نافذة أفراد' }, o);
  if (o.step != null) opts.objectName = an(o.step);
  delete opts.step;
  slide.addImage(opts);
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
const CITY = {
  jubail: { img: A('city_jubail.png'), w: 149, h: 148 },
  raskhair: { img: A('city_raskhair.png'), w: 211, h: 141 },
  yanbu: { img: A('city_yanbu.png'), w: 138, h: 138 },
  jazan: { img: A('city_jazan.png'), w: 132, h: 132 },
};
function cityImg(slide, key, cx, cy, box, step) {
  const c = CITY[key];
  const s = box / Math.max(c.w, c.h);
  const w = c.w * s, h = c.h * s;
  I(slide, { path: c.img, x: cx - w / 2, y: cy - h / 2, w, h, step });
}

// shared chrome for content slides
function chrome(slide, n, kicker, title, color = NAVY) {
  slide.background = { color: WHITE };
  // faint logo pattern, bleeding off the top-left corner
  slide.addImage({ path: MARK_NAVY, x: -1.9, y: -1.9, w: 4.6, h: 4.6 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
  T(slide, kicker, { x: 5.2, y: 0.45, w: W - M - 5.2, h: 0.4, fontSize: 14, bold: true, color: color === NAVY ? MUTED : GREEN });
  T(slide, title, { x: 3.2, y: 0.8, w: W - M - 3.2, h: 0.85, fontSize: 32, bold: true, color });
  // brand + number
  slide.addImage({ path: MARK, x: M, y: 6.83, w: 0.34, h: 0.34 * MR, altText: 'شعار نافذة' });
  T(slide, 'نافذة أفراد', { x: M + 0.42, y: 6.83, w: 1.6, h: 0.34, fontSize: 11, bold: true, color: NAVY, valign: 'middle', align: 'left' });
  T(slide, String(n), { x: W - M - 0.6, y: 6.83, w: 0.6, h: 0.34, fontSize: 11, color: MUTED, valign: 'middle' });
}

function circleIcon(slide, x, y, d, data, bg, step, pad = 0.26) {
  S(slide, pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: 'none' }, step });
  const s = d * (1 - 2 * pad);
  I(slide, { data, x: x + (d - s) / 2, y: y + (d - s) / 2, w: s, h: s, step });
}

(async () => {
  // ===== 1. Cover =====
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addImage({ path: MARK_NAVY, x: 9.6, y: -2.2, w: 5.6, h: 5.6 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    s.addImage({ path: MARK_NAVY, x: -2.4, y: 4.2, w: 5.6, h: 5.6 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    const lw = 2.2;
    I(s, { path: MARK, x: (W - lw) / 2, y: 0.7, w: lw, h: lw * MR, step: 1 });
    T(s, 'نافذة أفراد', { x: 2, y: 3.05, w: W - 4, h: 1.1, fontSize: 54, bold: true, color: NAVY, align: 'center', valign: 'middle', step: 2 });
    T(s, 'نافذة رقمية واحدة إلى خدمات المدينة وأماكنها وفعالياتها', { x: 2, y: 4.15, w: W - 4, h: 0.6, fontSize: 20, color: MUTED, align: 'center', valign: 'middle', step: 3 });
    // entity marks, right to left: RCJY, then JABEEN, then RAQEEM
    const ents = [
      ['الهيئة الملكية للجبيل وينبع', 'LOGO_RCJY'],
      ['جبين', 'LOGO_JABEEN'],
      ['رقيم للحلول الذكية', 'LOGO_RAQEEM'],
    ];
    const bw = 3.0, gap = 0.35, total = 3 * bw + 2 * gap;
    ents.forEach(([name], i) => {
      const x = (W + total) / 2 - bw - i * (bw + gap);
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.75, w: bw, h: 0.85, rectRadius: 0.12, fill: { color: WHITE }, line: { color: LINE, width: 1 }, step: 4 });
      T(s, name, { x: x + 0.1, y: 5.75, w: bw - 0.2, h: 0.85, fontSize: 14, bold: true, color: NAVY, align: 'center', valign: 'middle', step: 4 });
    });
    s.addNotes('نرحب بكم في حفل تدشين تطبيق نافذة أفراد. التطبيق نافذة رقمية واحدة يصل منها السكان والزوار إلى خدمات المدينة وأماكنها وفعالياتها. وهو ثمرة عمل مشترك بين الهيئة الملكية للجبيل وينبع، وشركة جبين، وشركة رقيم للحلول الذكية التي طوّرت التطبيق.');
  }

  // ===== 2. Logo idea =====
  {
    const s = pres.addSlide();
    chrome(s, 2, 'الهوية', 'فكرة الشعار');
    // visual on the left
    S(s, pres.shapes.OVAL, { x: 0.9, y: 1.85, w: 4.4, h: 4.4, fill: { color: TINT }, line: { type: 'none' }, step: 1 });
    I(s, { path: MARK, x: 0.9 + 0.75, y: 1.85 + 0.72, w: 2.9, h: 2.9 * MR, step: 1 });
    // statement on the right
    const tx = 6.2, tw = W - M - tx;
    T(s, 'نافذة واحدة يكتشف منها السكان والزوار المدينة', { x: tx, y: 1.95, w: tw, h: 0.95, fontSize: 22, bold: true, color: NAVY, step: 2 });
    const rows = [
      ['FaDoorOpen', 'نافذة إلى المدينة', 'المركز المفتوح بوابة إلى أماكن المدينة وفرصها'],
      ['FaUsers', 'هوية تجمع المدن', 'عناصر من الجبيل ورأس الخير وينبع وجازان'],
      ['FaCalendarAlt', 'خدمات وفعاليات في مكان واحد', 'وصول مباشر إلى المدينة من تطبيق واحد'],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [ic, t, d] = rows[i];
      const y = 3.15 + i * 1.12;
      circleIcon(s, W - M - 0.8, y, 0.8, await icon(ic, NAVY), TINT, 3 + i);
      T(s, t, { x: tx, y: y - 0.02, w: tw - 1.05, h: 0.42, fontSize: 18, bold: true, color: INK, step: 3 + i });
      T(s, d, { x: tx, y: y + 0.4, w: tw - 1.05, h: 0.38, fontSize: 14, color: MUTED, step: 3 + i });
    }
    s.addNotes('اسم نافذة يعبّر عن الفكرة ببساطة: نافذة واحدة يطل منها السكان والزوار على المدينة. من خلالها يصلون إلى الخدمات والأماكن والفعاليات دون تنقل بين قنوات متعددة.');
  }

  // ===== 3. Four cities =====
  {
    const s = pres.addSlide();
    chrome(s, 3, 'الهوية', 'المدن الأربع: من كل مدينة عنصر');
    const cards = [
      ['jubail', 'الجبيل', 'الساحل الشرقي', 'شكل هندسي متوازن\nبأربع بتلات'],
      ['raskhair', 'رأس الخير', 'الساحل الشرقي', 'تكوينات مدببة من الأوراق\nوالزخارف التقليدية'],
      ['yanbu', 'ينبع', 'الساحل الحجازي', 'نمط زهري هندسي متناظر\nيوحي بالحركة'],
      ['jazan', 'جازان', 'ساحل تهامة', 'أشكال ماسية متداخلة\nوزخارف نباتية'],
    ];
    const cw = 2.8, gap = 0.3, y0 = 1.95, ch = 4.2;
    const x0 = W - M - cw;
    cards.forEach(([k, name, region, desc], i) => {
      const x = x0 - i * (cw + gap);
      const st = i + 1;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: ch, rectRadius: 0.18, fill: { color: TINT }, line: { type: 'none' }, step: st });
      S(s, pres.shapes.OVAL, { x: x + (cw - 1.7) / 2, y: y0 + 0.35, w: 1.7, h: 1.7, fill: { color: WHITE }, line: { type: 'none' }, step: st });
      cityImg(s, k, x + cw / 2, y0 + 0.35 + 0.85, 1.05, st);
      T(s, name, { x: x + 0.2, y: y0 + 2.2, w: cw - 0.4, h: 0.55, fontSize: 22, bold: true, color: NAVY, align: 'center', step: st });
      T(s, region, { x: x + 0.2, y: y0 + 2.75, w: cw - 0.4, h: 0.35, fontSize: 12, color: MUTED, align: 'center', step: st });
      T(s, desc, { x: x + 0.25, y: y0 + 3.15, w: cw - 0.5, h: 0.85, fontSize: 14, color: INK, align: 'center', step: st });
    });
    T(s, 'المرجع المعماري: دليل العمارة السعودية (الساحل الشرقي، الساحل الحجازي، ساحل تهامة)', { x: 2.6, y: 6.3, w: W - M - 2.6, h: 0.35, fontSize: 11, color: MUTED, step: 5 });
    s.addNotes('استلهمنا الشعار من الطابع المعماري للمدن الأربع وفق دليل العمارة السعودية. أخذنا من كل مدينة عنصرًا يعبّر عنها: البتلات الأربع من الجبيل، والتكوينات المدببة من رأس الخير، والنمط الزهري من ينبع، والأشكال الماسية من جازان.');
  }

  // ===== 4. Four patterns into one =====
  {
    const s = pres.addSlide();
    chrome(s, 4, 'الهوية', 'من أربعة أنماط إلى نمط واحد');
    // diagram on the left
    const cx = 3.55, cy = 4.1;
    const pos = [
      ['jubail', 'الجبيل', cx + 2.05, cy - 1.55],
      ['raskhair', 'رأس الخير', cx - 2.05, cy - 1.55],
      ['yanbu', 'ينبع', cx + 2.05, cy + 1.35],
      ['jazan', 'جازان', cx - 2.05, cy + 1.35],
    ];
    const d = 1.05;
    pos.forEach(([k, name, px, py]) => {
      S(s, pres.shapes.OVAL, { x: px - d / 2, y: py - d / 2, w: d, h: d, fill: { color: TINT }, line: { type: 'none' }, step: 1 });
      cityImg(s, k, px, py, 0.62, 1);
      T(s, name, { x: px - 0.8, y: py + d / 2 + 0.02, w: 1.6, h: 0.32, fontSize: 12, bold: true, color: NAVY, align: 'center', step: 1 });
      // connector toward the centre
      const sx = px + (cx - px) * 0.32, sy = py + (cy - py) * 0.32;
      const ex = px + (cx - px) * 0.58, ey = py + (cy - py) * 0.58;
      S(s, pres.shapes.LINE, {
        x: Math.min(sx, ex), y: Math.min(sy, ey), w: Math.abs(ex - sx), h: Math.abs(ey - sy),
        flipH: ex < sx, flipV: ey < sy,
        line: { color: NAVY, width: 1.25, endArrowType: 'triangle', dashType: 'solid' },
        step: 2,
      });
    });
    const lw = 1.75;
    I(s, { path: MARK, x: cx - lw / 2, y: cy - (lw * MR) / 2, w: lw, h: lw * MR, step: 3 });

    // points on the right
    const tx = 7.0, tw = W - M - tx;
    const pts = [
      ['FaLayerGroup', 'الدمج', 'بسّطنا أبرز عناصر كل نمط ودمجناها في زخرفة واحدة'],
      ['FaGem', 'المركز', 'مساحة مفتوحة على شكل مُعيّن تمثل النافذة'],
      ['FaArrowsAlt', 'الاتجاهات الأربعة', 'اجتماع المدن الأربع حول نافذة رقمية واحدة'],
    ];
    for (let i = 0; i < pts.length; i++) {
      const [ic, t, dd] = pts[i];
      const y = 1.95 + i * 1.08;
      circleIcon(s, W - M - 0.72, y, 0.72, await icon(ic, NAVY), TINT, 4 + i);
      T(s, t, { x: tx, y: y - 0.04, w: tw - 0.95, h: 0.4, fontSize: 18, bold: true, color: INK, step: 4 + i });
      T(s, dd, { x: tx, y: y + 0.36, w: tw - 0.95, h: 0.36, fontSize: 14, color: MUTED, step: 4 + i });
    }
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: tx, y: 5.3, w: tw, h: 1.05, rectRadius: 0.14, fill: { color: NAVY }, line: { type: 'none' }, step: 7 });
    T(s, 'نمط جديد وُلد من هوية المدن الأربع، وليس تجميعًا لأربع زخارف', { x: tx + 0.3, y: 5.3, w: tw - 0.6, h: 1.05, fontSize: 17, bold: true, color: WHITE, valign: 'middle', step: 7 });
    s.addNotes('لم نضع الزخارف الأربع جنبًا إلى جنب، بل بسّطنا عناصرها ودمجناها في نمط واحد جديد. في المركز مُعيّن مفتوح يمثل النافذة، وحوله أربعة اتجاهات تجتمع فيها المدن الأربع. التناسق الهندسي يعبّر عن التكامل وسهولة الوصول.');
  }

  // ===== 5. Colours =====
  {
    const s = pres.addSlide();
    chrome(s, 5, 'الهوية', 'الألوان: من شعار الهيئة');
    const pw = 5.8, ph = 3.6, y = 1.95, gap = 0.5;
    const panels = [
      [NAVY, 'نافذة أفراد', '#052977', 'الثقة والطابع الرسمي'],
      [GREEN, 'نافذة أعمال', '#0B9D76', 'لون نسخة الأعمال'],
    ];
    panels.forEach(([c, name, hex, desc], i) => {
      const x = W - M - pw - i * (pw + gap);
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y, w: pw, h: ph, rectRadius: 0.2, fill: { color: c }, line: { type: 'none' }, step: i + 1 });
      I(s, { path: MARK_WHITE, x: x + 0.35, y: y + ph - 2.05, w: 1.7, h: 1.7 * MR, transparency: 80, step: i + 1 });
      T(s, name, { x: x + 0.4, y: y + 0.4, w: pw - 0.8, h: 0.75, fontSize: 30, bold: true, color: WHITE, step: i + 1 });
      T(s, hex, { x: x + 0.4, y: y + 1.2, w: pw - 0.8, h: 0.5, fontSize: 20, color: WHITE, step: i + 1 });
      T(s, desc, { x: x + 0.4, y: y + 1.75, w: pw - 0.8, h: 0.45, fontSize: 16, color: WHITE, step: i + 1 });
    });
    T(s, 'اللونان مأخوذان من شعار الهيئة الملكية للجبيل وينبع: الأزرق لنافذة أفراد، والأخضر لنافذة أعمال', { x: M, y: 5.85, w: W - 2 * M, h: 0.45, fontSize: 15, color: INK, step: 3 });
    s.addNotes('اخترنا الألوان من شعار الهيئة الملكية للجبيل وينبع لتبقى الهوية امتدادًا لهوية الهيئة. الأزرق الداكن لنافذة أفراد ويعبّر عن الثقة، والأخضر مخصص لنافذة أعمال.');
  }

  // ===== 6. What is Nafitha Afrad =====
  {
    const s = pres.addSlide();
    chrome(s, 6, 'التطبيق', 'ما هي نافذة أفراد؟');
    // app tile on the left
    const tx0 = 1.4, ty0 = 2.1, ts = 3.0;
    S(s, pres.shapes.ROUNDED_RECTANGLE, {
      x: tx0, y: ty0, w: ts, h: ts, rectRadius: 0.55, fill: { color: WHITE }, line: { color: LINE, width: 1 },
      shadow: { type: 'outer', color: '052977', opacity: 0.18, blur: 18, offset: 6, angle: 90 }, step: 1,
    });
    I(s, { path: MARK, x: tx0 + 0.55, y: ty0 + 0.52, w: 1.9, h: 1.9 * MR, step: 1 });
    T(s, 'نافذة أفراد', { x: tx0 - 0.3, y: ty0 + ts + 0.2, w: ts + 0.6, h: 0.5, fontSize: 18, bold: true, color: NAVY, align: 'center', step: 1 });

    const x = 5.0, w = W - M - x;
    T(s, 'نافذة أفراد تطبيق رقمي يجمع خدمات المدينة وأماكنها وفعالياتها في مكان واحد.', { x, y: 2.05, w, h: 1.5, fontSize: 26, bold: true, color: NAVY, step: 2 });
    T(s, 'هدفه أن يكتشف السكان والزوار المدينة، ويصلوا إلى خدماتها بسهولة.', { x, y: 3.75, w, h: 1.0, fontSize: 18, color: INK, step: 3 });
    // audience chips
    const chips = [['FaUsers', 'السكان'], ['FaUserFriends', 'الزوار']];
    for (let i = 0; i < chips.length; i++) {
      const cw = 1.9, cx = W - M - cw - i * (cw + 0.25);
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: 5.1, w: cw, h: 0.62, rectRadius: 0.31, fill: { color: TINT }, line: { type: 'none' }, step: 4 });
      I(s, { data: await icon(chips[i][0], NAVY), x: cx + cw - 0.52, y: 5.26, w: 0.3, h: 0.3, step: 4 });
      T(s, chips[i][1], { x: cx + 0.15, y: 5.1, w: cw - 0.75, h: 0.62, fontSize: 15, bold: true, color: NAVY, valign: 'middle', step: 4 });
    }
    s.addNotes('نافذة أفراد تطبيق واحد يجمع ما تحتاجه عن المدينة: خدماتها وأماكنها وفعالياتها. صُمّم للسكان والزوار معًا، ليكتشفوا المدينة ويصلوا إلى خدماتها بسهولة.');
  }

  // ===== 7. Why Nafitha =====
  {
    const s = pres.addSlide();
    chrome(s, 7, 'التطبيق', 'لماذا نافذة؟');
    const cw = 5.8, ch = 3.05, y0 = 1.9;
    const cols = [
      ['التحدي', WHITE, LINE, [
        ['FaRandom', 'معلومات المدينة موزعة على قنوات متعددة'],
        ['FaSearch', 'البحث عن الفعاليات والأماكن يأخذ وقتًا'],
        ['FaCompass', 'الزائر لا يعرف من أين يبدأ'],
      ]],
      ['ما تقدمه نافذة', TINT, null, [
        ['FaDoorOpen', 'نافذة واحدة لخدمات المدينة وفعالياتها'],
        ['FaStopwatch', 'وصول أسرع للسكان والزوار'],
        ['FaCity', 'تجربة واحدة في المدن الأربع'],
      ]],
    ];
    for (let c = 0; c < 2; c++) {
      const [head, bg, ln, items] = cols[c];
      const x = W - M - cw - c * (cw + 0.53);
      const st = c + 1;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: ch, rectRadius: 0.18, fill: { color: bg }, line: ln ? { color: ln, width: 1 } : { type: 'none' }, step: st });
      T(s, head, { x: x + 0.35, y: y0 + 0.25, w: cw - 0.7, h: 0.5, fontSize: 20, bold: true, color: NAVY, step: st });
      for (let i = 0; i < items.length; i++) {
        const y = y0 + 0.95 + i * 0.66;
        I(s, { data: await icon(items[i][0], c === 0 ? MUTED : NAVY), x: x + cw - 0.35 - 0.34, y: y + 0.06, w: 0.34, h: 0.34, step: st });
        T(s, items[i][1], { x: x + 0.3, y, w: cw - 1.15, h: 0.46, fontSize: 15, color: INK, valign: 'middle', step: st });
      }
    }
    // audience question
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 5.3, w: W - 2 * M, h: 1.15, rectRadius: 0.16, fill: { color: NAVY }, line: { type: 'none' }, step: 3 });
    circleIcon(s, W - M - 0.3 - 0.72, 5.3 + 0.215, 0.72, await icon('FaQuestion', NAVY), WHITE, 3, 0.28);
    T(s, 'سؤال للحضور: كم تطبيقًا أو موقعًا تحتاج اليوم لتعرف فعاليات مدينتك وتصل إلى خدماتها؟', { x: M + 0.4, y: 5.3, w: W - 2 * M - 1.65, h: 1.15, fontSize: 18, bold: true, color: WHITE, valign: 'middle', step: 3 });
    s.addNotes('اليوم تتوزع معلومات المدينة على قنوات متعددة، ويحتاج المستخدم وقتًا ليعرف أين يجد ما يريد. نافذة تجمع ذلك في مكان واحد. قبل أن ننتقل: كم تطبيقًا أو موقعًا تحتاجون اليوم لتعرفوا فعاليات مدينتكم وتصلوا إلى خدماتها؟ (امنح الحضور لحظة للإجابة).');
  }

  // ===== 8. Services 1: discover the city =====
  {
    const s = pres.addSlide();
    chrome(s, 8, 'خدمات نافذة أفراد', 'اكتشف المدينة');
    const items = [
      ['FaMapMarkedAlt', 'الأماكن', 'تعرّف على وجهات المدينة ومعالمها'],
      ['FaCalendarAlt', 'الفعاليات', 'تابع فعاليات المدينة في مكان واحد'],
      ['FaCompass', 'التجارب', 'تجارب المدينة للسكان والزوار'],
    ];
    const cw = 3.75, gap = 0.44, y0 = 2.05, ch = 3.55;
    for (let i = 0; i < items.length; i++) {
      const [ic, t, d] = items[i];
      const x = W - M - cw - i * (cw + gap);
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: ch, rectRadius: 0.2, fill: { color: TINT }, line: { type: 'none' }, step: i + 1 });
      circleIcon(s, x + (cw - 1.3) / 2, y0 + 0.45, 1.3, await icon(ic, WHITE), NAVY, i + 1, 0.28);
      T(s, t, { x: x + 0.3, y: y0 + 2.0, w: cw - 0.6, h: 0.6, fontSize: 24, bold: true, color: NAVY, align: 'center', step: i + 1 });
      T(s, d, { x: x + 0.3, y: y0 + 2.7, w: cw - 0.6, h: 0.8, fontSize: 15, color: INK, align: 'center', step: i + 1 });
    }
    s.addNotes('القسم الأول من الخدمات يساعدك على اكتشاف المدينة. تجد فيه أماكن المدينة وفعالياتها وتجاربها في مكان واحد، سواء كنت من السكان أو من الزوار.');
  }

  // ===== 9. Services 2: get things done =====
  {
    const s = pres.addSlide();
    chrome(s, 9, 'خدمات نافذة أفراد', 'أنجز من التطبيق');
    // visual on the left
    S(s, pres.shapes.OVAL, { x: 1.0, y: 1.95, w: 4.1, h: 4.1, fill: { color: TINT }, line: { type: 'none' }, step: 1 });
    I(s, { data: await icon('FaMobileAlt', NAVY), x: 1.0 + 1.25, y: 1.95 + 1.1, w: 1.6, h: 1.6, step: 1 });
    const items = [
      ['FaThLarge', 'خدمات المدينة', 'الوصول إلى خدمات المدينة من مكان واحد'],
      ['FaLightbulb', 'الفرص', 'اطّلع على فرص المدينة المتاحة'],
      ['FaCity', 'المدن الأربع', 'الجبيل ورأس الخير وينبع وجازان في تطبيق واحد'],
    ];
    const x = 6.0, w = W - M - x;
    for (let i = 0; i < items.length; i++) {
      const [ic, t, d] = items[i];
      const y = 1.95 + i * 1.4;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.15, rectRadius: 0.16, fill: { color: WHITE }, line: { color: LINE, width: 1 }, step: i + 2 });
      circleIcon(s, x + w - 0.25 - 0.75, y + 0.2, 0.75, await icon(ic, WHITE), NAVY, i + 2, 0.28);
      T(s, t, { x: x + 0.3, y: y + 0.14, w: w - 1.4, h: 0.45, fontSize: 19, bold: true, color: NAVY, step: i + 2 });
      T(s, d, { x: x + 0.3, y: y + 0.6, w: w - 1.4, h: 0.4, fontSize: 14, color: MUTED, step: i + 2 });
    }
    s.addNotes('القسم الثاني يوصلك إلى خدمات المدينة وفرصها مباشرة من التطبيق. والتطبيق واحد للمدن الأربع: الجبيل ورأس الخير وينبع وجازان.');
  }

  // ===== 10. User journey =====
  {
    const s = pres.addSlide();
    chrome(s, 10, 'التطبيق', 'رحلة المستخدم: من الفتح إلى الإنجاز');
    const steps = [
      ['FaMobileAlt', 'افتح التطبيق', 'ابدأ من الشاشة الرئيسية'],
      ['FaHandPointer', 'اختر الخدمة', 'حدّد الخدمة التي تحتاجها'],
      ['FaEdit', 'أكمل الطلب', 'أدخل البيانات المطلوبة'],
      ['FaCheckCircle', 'أنجز الخدمة', 'من التطبيق نفسه'],
    ];
    const n = steps.length, span = W - 2 * M - 0.6, sw = span / n;
    const cy = 3.05, d = 1.25;
    // connector behind the circles (right to left)
    S(s, pres.shapes.LINE, { x: M + 0.3 + sw / 2, y: cy + d / 2, w: sw * (n - 1), h: 0, line: { color: LINE, width: 2 }, step: 1 });
    for (let i = 0; i < n; i++) {
      const [ic, t, dsc] = steps[i];
      const ccx = W - M - 0.3 - sw / 2 - i * sw;
      const st = i + 1;
      circleIcon(s, ccx - d / 2, cy, d, await icon(ic, WHITE), NAVY, st, 0.28);
      S(s, pres.shapes.OVAL, { x: ccx + d / 2 - 0.42, y: cy - 0.12, w: 0.46, h: 0.46, fill: { color: WHITE }, line: { color: NAVY, width: 1.5 }, step: st });
      T(s, String(i + 1), { x: ccx + d / 2 - 0.42, y: cy - 0.12, w: 0.46, h: 0.46, fontSize: 14, bold: true, color: NAVY, align: 'center', valign: 'middle', step: st });
      T(s, t, { x: ccx - sw / 2 + 0.1, y: cy + d + 0.3, w: sw - 0.2, h: 0.5, fontSize: 20, bold: true, color: NAVY, align: 'center', step: st });
      T(s, dsc, { x: ccx - sw / 2 + 0.1, y: cy + d + 0.85, w: sw - 0.2, h: 0.45, fontSize: 15, color: MUTED, align: 'center', step: st });
    }
    s.addNotes('هذا مثال مبسط لرحلة المستخدم. يفتح التطبيق، ويختار الخدمة التي يحتاجها، ويكمل الطلب، وينجز الخدمة دون أن يغادر التطبيق.');
  }

  // ===== 11. Roadmap =====
  {
    const s = pres.addSlide();
    chrome(s, 11, 'المستقبل', 'خارطة الطريق');
    const stages = [
      [NAVY, TINT, 'المرحلة الأولى', 'نافذة أفراد', 'الإطلاق الحالي', 'للسكان والزوار', 'FaFlag'],
      [GREEN, TINT_G, 'المرحلة التالية', 'نافذة أعمال', 'قادمة', 'لأصحاب الأعمال', 'FaBriefcase'],
    ];
    const cw = 5.0, ch = 3.6, y0 = 2.1;
    const xs = [W - M - 0.4 - cw, M + 0.4];
    // arrow between stages (points left: reading direction)
    S(s, pres.shapes.LINE, { x: xs[1] + cw + 0.25, y: y0 + ch / 2, w: xs[0] - xs[1] - cw - 0.5, h: 0, flipH: true, line: { color: MUTED, width: 2, endArrowType: 'triangle' }, step: 2 });
    for (let i = 0; i < 2; i++) {
      const [c, bg, lbl, name, status, who, ic] = stages[i];
      const x = xs[i], st = i === 0 ? 1 : 3;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: ch, rectRadius: 0.2, fill: { color: bg }, line: { type: 'none' }, step: st });
      circleIcon(s, x + cw - 0.4 - 1.0, y0 + 0.4, 1.0, await icon(ic, WHITE), c, st, 0.27);
      T(s, lbl, { x: x + 0.4, y: y0 + 0.45, w: cw - 2.0, h: 0.4, fontSize: 14, bold: true, color: MUTED, step: st });
      T(s, name, { x: x + 0.4, y: y0 + 1.6, w: cw - 0.8, h: 0.7, fontSize: 30, bold: true, color: c, step: st });
      T(s, who, { x: x + 0.4, y: y0 + 2.3, w: cw - 0.8, h: 0.45, fontSize: 16, color: INK, step: st });
      // status pill
      const pw = 1.9;
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x: x + cw - 0.4 - pw, y: y0 + ch - 0.8, w: pw, h: 0.5, rectRadius: 0.25, fill: { color: c }, line: { type: 'none' }, step: st });
      T(s, status, { x: x + cw - 0.4 - pw, y: y0 + ch - 0.8, w: pw, h: 0.5, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle', step: st });
    }
    s.addNotes('نطلق اليوم المرحلة الأولى: نافذة أفراد للسكان والزوار. المرحلة التالية هي نافذة أعمال، وهي موجهة لأصحاب الأعمال.');
  }

  // ===== 12. Business: upcoming services =====
  {
    const s = pres.addSlide();
    chrome(s, 12, 'نافذة أعمال', 'نافذة أعمال: الخدمات القادمة', GREEN);
    const items = [
      ['FaBuilding', 'خدمات المنشآت', 'خدمات المدينة لأصحاب الأعمال من نافذة واحدة'],
      ['FaLightbulb', 'فرص الأعمال', 'الاطلاع على الفرص في المدن الأربع'],
      ['FaCalendarAlt', 'فعاليات الأعمال', 'متابعة فعاليات الأعمال في المدينة'],
    ];
    const cw = 3.75, gap = 0.44, y0 = 2.05, ch = 3.55;
    for (let i = 0; i < items.length; i++) {
      const [ic, t, d] = items[i];
      const x = W - M - cw - i * (cw + gap);
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: ch, rectRadius: 0.2, fill: { color: WHITE }, line: { color: GREEN, width: 1.25 }, step: i + 1 });
      circleIcon(s, x + (cw - 1.2) / 2, y0 + 0.5, 1.2, await icon(ic, GREEN), TINT_G, i + 1, 0.28);
      T(s, t, { x: x + 0.3, y: y0 + 2.0, w: cw - 0.6, h: 0.6, fontSize: 22, bold: true, color: GREEN, align: 'center', step: i + 1 });
      T(s, d, { x: x + 0.3, y: y0 + 2.7, w: cw - 0.6, h: 0.8, fontSize: 15, color: INK, align: 'center', step: i + 1 });
    }
    s.addNotes('المرحلة التالية هي نافذة أعمال. ستجمع لأصحاب الأعمال خدمات المدينة وفرصها وفعالياتها في مكان واحد، بالفكرة نفسها التي بُنيت عليها نافذة أفراد.');
  }

  // ===== 13. Business: benefits =====
  {
    const s = pres.addSlide();
    chrome(s, 13, 'نافذة أعمال', 'ماذا تقدم نافذة أعمال لأصحاب الأعمال؟', GREEN);
    const items = [
      ['FaDoorOpen', 'نافذة واحدة', 'بدل التنقل بين قنوات متعددة'],
      ['FaStopwatch', 'وقت أقل', 'وصول أسرع إلى الخدمات'],
      ['FaEye', 'رؤية أوضح', 'اطلاع على الفرص في المدن الأربع'],
      ['FaLink', 'تجربة متصلة', 'هوية وتجربة موحدة مع نافذة أفراد'],
    ];
    const cw = 5.85, ch = 1.85, gx = 0.43, gy = 0.35, y0 = 1.95;
    for (let i = 0; i < items.length; i++) {
      const [ic, t, d] = items[i];
      const col = i % 2, row = Math.floor(i / 2);
      const x = W - M - cw - col * (cw + gx), y = y0 + row * (ch + gy);
      S(s, pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.18, fill: { color: TINT_G }, line: { type: 'none' }, step: i + 1 });
      circleIcon(s, x + cw - 0.35 - 1.0, y + (ch - 1.0) / 2, 1.0, await icon(ic, WHITE), GREEN, i + 1, 0.27);
      T(s, t, { x: x + 0.35, y: y + 0.42, w: cw - 1.85, h: 0.5, fontSize: 21, bold: true, color: GREEN, step: i + 1 });
      T(s, d, { x: x + 0.35, y: y + 0.95, w: cw - 1.85, h: 0.45, fontSize: 15, color: INK, step: i + 1 });
    }
    s.addNotes('لأصحاب الأعمال، الفائدة الأولى هي نافذة واحدة بدل التنقل بين قنوات متعددة. وهذا يعني وقتًا أقل، ورؤية أوضح للفرص، وتجربة متصلة مع نافذة أفراد.');
  }

  // ===== 14. Closing =====
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addImage({ path: MARK_NAVY, x: 9.3, y: -2.3, w: 6.0, h: 6.0 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    s.addImage({ path: MARK_NAVY, x: -2.2, y: 4.3, w: 5.4, h: 5.4 * MR, transparency: 94, altText: 'زخرفة شعار نافذة' });
    const x = 5.0, w = W - M - 0.4 - x;
    I(s, { path: MARK, x: W - M - 0.4 - 1.3, y: 1.1, w: 1.3, h: 1.3 * MR, step: 1 });
    T(s, 'نافذة أفراد: نافذتكم إلى المدينة', { x, y: 2.75, w, h: 0.9, fontSize: 36, bold: true, color: NAVY, step: 2 });
    T(s, 'نعلن اليوم تدشين التطبيق للسكان والزوار.', { x, y: 3.7, w, h: 0.9, fontSize: 20, color: INK, step: 3 });
    T(s, 'شكرًا لكم', { x, y: 4.75, w, h: 0.6, fontSize: 22, bold: true, color: MUTED, step: 3 });
    // QR slot (replace with the store QR code)
    const qx = 1.2, qy = 1.55, qs = 3.0;
    S(s, pres.shapes.ROUNDED_RECTANGLE, { x: qx, y: qy, w: qs, h: qs, rectRadius: 0.2, fill: { color: WHITE }, line: { color: NAVY, width: 1.5 }, objectName: 'QR_CODE', step: 2 });
    I(s, { data: await icon('FaQrcode', NAVY), x: qx + 0.55, y: qy + 0.55, w: qs - 1.1, h: qs - 1.1, step: 2 });
    T(s, 'امسح الرمز لتحميل التطبيق', { x: qx - 0.3, y: qy + qs + 0.2, w: qs + 0.6, h: 0.45, fontSize: 15, bold: true, color: NAVY, align: 'center', step: 2 });
    T(s, 'تطوير شركة رقيم للحلول الذكية', { x: 2, y: 6.55, w: W - 4, h: 0.45, fontSize: 14, color: MUTED, align: 'center', step: 4 });
    s.addNotes('بهذا نعلن تدشين تطبيق نافذة أفراد. يمكنكم تحميله الآن بمسح الرمز على الشاشة. نشكر الهيئة الملكية للجبيل وينبع وشركة جبين على دعمهما، والتطبيق من تطوير شركة رقيم للحلول الذكية.');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
