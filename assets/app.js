/* ============================================================
   بيت هايل — لوحة التوطين
   منطق التطبيق: قراءة البيانات + الحسابات + الرسوم
   ============================================================ */

const DATA_URL = 'data/nitaqat.xlsx';

const COMPANIES = {
  materials: {
    key: 'مواد_البناء',
    name: 'بيت هايل لمواد البناء',
    id: '32-1942229',
    page: 'materials.html',
    bands: [
      { name: 'أحمر',        min: 0,  max: 22,  color: 'var(--band-red)'  },
      { name: 'أخضر منخفض',  min: 22, max: 29,  color: 'var(--band-low)'  },
      { name: 'أخضر متوسط',  min: 29, max: 32,  color: 'var(--band-mid)'  },
      { name: 'أخضر مرتفع',  min: 32, max: 38,  color: 'var(--band-high)' },
      { name: 'بلاتيني',     min: 38, max: 100, color: 'var(--band-plat)' },
    ],
    scaleCap: 46 // لعرض المسطرة (البلاتيني مفتوح النهاية)
  },
  transport: {
    key: 'النقليات',
    name: 'بيت هائل للنقليات',
    id: '32-1942435',
    page: 'transport.html',
    bands: [
      { name: 'أحمر',        min: 0,  max: 18,  color: 'var(--band-red)'  },
      { name: 'أخضر منخفض',  min: 18, max: 22,  color: 'var(--band-low)'  },
      { name: 'أخضر متوسط',  min: 22, max: 25,  color: 'var(--band-mid)'  },
      { name: 'أخضر مرتفع',  min: 25, max: 36,  color: 'var(--band-high)' },
      { name: 'بلاتيني',     min: 36, max: 100, color: 'var(--band-plat)' },
    ],
    scaleCap: 44
  },
  maintenance: {
    key: 'الصيانة',
    name: 'هائل لصيانة السيارات',
    id: '32-1942463',
    page: 'maintenance.html',
    bands: [
      { name: 'أحمر (فئة أ)', min: 0,  max: 50,  color: 'var(--band-red)'  },
      { name: 'أخضر (فئة أ)', min: 50, max: 100, color: 'var(--band-high)' },
    ],
    scaleCap: 100
  }
};

/* ============================================================
   قرارات توطين المهن (وزارة الموارد البشرية والتنمية الاجتماعية)
   مستخرجة من الأدلة الإجرائية الرسمية المرفقة
   ============================================================ */
const PROFESSION_DECISIONS = [
  {
    id: 'admin_support',
    title: 'المهن الإدارية المساندة',
    decree: 'قرار وزاري رقم 132249 وتاريخ 1447/9/17هـ (2026/4/5م)',
    pct: 100,
    minEmployees: 1,
    effective: '2026/4/5م (وبعض المهن حتى 2026/10/4م)',
    note: 'قصر العمل بالكامل على السعوديين — أي عامل غير سعودي في هذه المهن يُعد مخالفة فورية.',
    professions: [
      'مدير علاقات عامة','مدير موارد بشرية أمن سيبراني','مدير تحقيق إداري','مدير تنظيم إداري',
      'مدير تطوير مؤسسي','مدير مكتب خدمة مدنية','مدير شؤون عمل','مدير مكتب عمل','مدير مكتب استقدام',
      'مدير علاقات الأفراد','مدير تطوير موارد بشرية','مدير مواهب','مدير تعويضات','مدير توظيف',
      'مدير تصنيف مهن','مدير تخطيط قوى عاملة','مدير عمليات موارد بشرية','أخصائي علاقات عامة',
      'خبير موارد بشرية','مستشار موارد بشرية','أخصائي استقدام','أخصائي مراقبة موارد بشرية',
      'أخصائي تخطيط قوى عاملة','عضو هيئة عمالية','محقق عمالي','أخصائي توجيه مهني','أخصائي تعويضات',
      'أخصائي توظيف','أخصائي تصنيف وظيفي ومهني','أخصائي عمليات موارد بشرية','أخصائي ارتباط وظيفي',
      'أخصائي تنظيم مؤتمرات وفعاليات','مدير مراسم','مدير جمع التبرعات','أخصائي توعية جمهور',
      'أخصائي تواصل داخلي','أخصائي مراسم','مستشار علاقات عامة','خبير علاقات عامة','كاتب سجل',
      'كاتب علاقات حكومية','سكرتير تنفيذي','مخلص جمركي','وكيل جمركي','وكيل شحن','أخصائي لغوي',
      'مترجم لغة إشارة','مصحح لغوي','مترجم فوري','مترجم','أمين صندوق','سكرتير','كاتب استقبال مرضى',
      'كاتب شكاوى','موظف استقبال فندق','مدخل بيانات','طباع','كاتب اختزال','حارس شخصي','حارس أمن',
      'مراقب كاميرات أمنية','كاتب موارد بشرية','ناسخ','كاتب حركة مخزون','كاتب شحن','أمين مخزن',
      'موظف استقبال','كاتب استعلامات','مساعد إداري',
    ]
  },
  {
    id: 'sales',
    title: 'مهن المبيعات',
    decree: 'قرار وزاري رقم 101278 وتاريخ 2026/01/19م',
    pct: 60,
    minEmployees: 3,
    effective: '2026/04/19م',
    note: 'يُحتسب على مستوى الكيان لا الفرع؛ يشمل جميع مسميات المبيعات.',
    professions: [
      'مدير مبيعات','مدير مبيعات تجزئة','مدير مبيعات جملة','مندوب مبيعات','وسيط سلع مستقبلية',
      'أخصائي مبيعات أجهزة تقنية المعلومات والاتصالات','أخصائي مبيعات','أخصائي تجاري','وسيط سلع',
    ]
  },
  {
    id: 'purchasing',
    title: 'مهن المشتريات',
    decree: 'قرار وزاري رقم 77050 وتاريخ 2025/11/30م',
    pct: 70,
    minEmployees: 3,
    effective: '2026/05/31م',
    note: '',
    professions: [
      'مدير مشتريات','مندوب مشتريات','مدير عقود','مدير خدمات لوجستية','مدير مستودع','أمين مستودع',
      'أخصائي مناقصات','أخصائي مشتريات','أخصائي توريد للعلامات التجارية الخاصة',
      'أخصائي تجارة إلكترونية','أخصائي أبحاث أسواق','أخصائي مستودعات',
    ]
  },
  {
    id: 'marketing',
    title: 'مهن التسويق',
    decree: 'قرار وزاري رقم 101319 وتاريخ 2026/01/19م',
    pct: 60,
    minEmployees: 3,
    effective: '2026/04/19م',
    note: 'لا يُحتسب السعودي إن قل أجره الخاضع للتأمينات عن 5,500 ريال.',
    professions: [
      'مدير تسويق','وكيل دعاية وإعلان','مدير دعاية وإعلان','أخصائي علاقات عامة',
      'أخصائي دعاية وإعلان','أخصائي تسويق','مدير علاقات عامة','مصمم جرافيك','مصمم إعلان',
      'مصور فوتوغرافي','أخصائي إعلامي',
    ]
  },
  {
    id: 'accounting',
    title: 'المهن المحاسبية',
    decree: 'قرار وزاري رقم 103108 وتاريخ 2025/01/26م',
    pct: 40,
    minEmployees: 5,
    effective: 'تدريجي من 2025/10/27م حتى 70% على مدى 5 سنوات',
    note: 'يشترط اعتماد مهني من الهيئة السعودية للمحاسبين، وحد أدنى للأجر (6,000 ريال بكالوريوس / 4,500 دبلوم).',
    professions: [
      'مدير مالي','مدير حسابات','مدير حسابات تعرفة','مدير خزينة','مدير ميزانية','مدير مراجعة',
      'مدير مراجعة داخلية','مدير تحصيل','مدير خزانة','محاسب تكاليف','مراقب مالي','مراجع داخلي',
      'كاتب حسابات','محاسب','أخصائي موازنة مالية','أخصائي حسابات ضرائب','محاسب قانوني',
      'مسؤول ضرائب','مساعد حسابات','أخصائي مراقبة مخزون','مراقب مخزون','كاتب مالي',
    ]
  },
];

function normalizeAr(s){
  return (s || '')
    .trim()
    .replace(/[\u064B-\u065F\u0670]/g, '')      // تشكيل
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/\s+/g, ' ');
}

function computeDecisionCompliance(stats, decision) {
  const targetSet = new Set(decision.professions.map(normalizeAr));
  const matchedJobs = [];
  let saudiMatched = 0, nonSaudiMatched = 0;

  Object.keys(stats.jobLocalization ? {} : {}); // no-op guard
  stats.all.forEach(p => {
    if (targetSet.has(normalizeAr(p.job))) {
      matchedJobs.push(p);
      if (p.group === 'سعودي') saudiMatched++; else nonSaudiMatched++;
    }
  });

  const totalMatched = saudiMatched + nonSaudiMatched;
  const applicable = totalMatched >= decision.minEmployees;
  const requiredCount = Math.round(totalMatched * (decision.pct / 100));
  const gap = Math.max(0, requiredCount - saudiMatched);
  const jobsFound = [...new Set(matchedJobs.map(p => p.job))];

  return { decision, totalMatched, saudiMatched, nonSaudiMatched, applicable, requiredCount, gap, jobsFound, compliant: gap === 0 && totalMatched > 0 };
}

let WORKBOOK = null;
let OFFICIAL_OVERRIDES = null; // { 'اسم الشركة|القرار': {total, saudi, nonSaudi, note} }

function loadOfficialOverrides(wb) {
  if (OFFICIAL_OVERRIDES) return OFFICIAL_OVERRIDES;
  OFFICIAL_OVERRIDES = {};
  const rows = sheetToRawRows(wb, 'ارقام_قوى_الرسمية');
  rows.forEach(r => {
    const company = String(r[0] ?? '').trim();
    const decision = String(r[1] ?? '').trim();
    const total = r[2], saudi = r[3], nonSaudi = r[4];
    if (!company || !decision || total === '' || total === undefined) return;
    OFFICIAL_OVERRIDES[company + '|' + decision] = {
      total: Number(total) || 0,
      saudi: Number(saudi) || 0,
      nonSaudi: Number(nonSaudi) || 0,
      note: String(r[5] ?? '').trim()
    };
  });
  return OFFICIAL_OVERRIDES;
}

function sheetToRawRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const [header, ...body] = rows;
  return body.filter(r => r.some(c => String(c).trim() !== ''));
}

// يرجع نتيجة القرار: يفضّل الرقم الرسمي من قوى إذا موجود، وإلا يرجع للتقدير التلقائي من المسميات الوظيفية
function getDecisionDisplay(stats, decision, wb) {
  const overrides = loadOfficialOverrides(wb);
  const key = stats.cfg.name + '|' + decision.title;
  const override = overrides[key];

  if (override) {
    const requiredCount = Math.round(override.total * (decision.pct / 100));
    const gap = Math.max(0, requiredCount - override.saudi);
    const applicable = override.total >= decision.minEmployees;
    return {
      decision, totalMatched: override.total, saudiMatched: override.saudi, nonSaudiMatched: override.nonSaudi,
      applicable, requiredCount, gap, jobsFound: [], compliant: gap === 0 && override.total > 0,
      isOfficial: true, officialNote: override.note
    };
  }
  const r = computeDecisionCompliance(stats, decision);
  return { ...r, isOfficial: false };
}

async function loadWorkbook() {
  if (WORKBOOK) return WORKBOOK;
  const res = await fetch(DATA_URL, { cache: 'no-store' });
  if (!res.ok) throw new Error('تعذر تحميل ملف البيانات');
  const buf = await res.arrayBuffer();
  WORKBOOK = XLSX.read(buf, { type: 'array' });
  return WORKBOOK;
}

function sheetToRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const [header, ...body] = rows;
  return body
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => ({
      id: String(r[0] ?? '').trim(),
      name: String(r[1] ?? '').trim(),
      gender: String(r[2] ?? '').trim(),
      nationality: String(r[3] ?? '').trim(),
      job: String(r[4] ?? '').trim(),
    }));
}

function getBand(pct, bands) {
  for (const b of bands) {
    if (pct >= b.min && pct < b.max) return b;
  }
  return bands[bands.length - 1];
}

function nextBand(pct, bands) {
  const cur = getBand(pct, bands);
  const idx = bands.indexOf(cur);
  return bands[idx + 1] || null;
}

// عدد السعوديين الإضافيين المطلوبين للوصول لعتبة نسبة معيّنة (بدون تغيير عدد غير السعوديين)
function neededToReach(saudi, total, thresholdPct) {
  const t = thresholdPct / 100;
  const nonSaudi = total - saudi;
  // (saudi + n) / (total + n) >= t  =>  n >= (t*total - saudi) / (1 - t)
  const n = (t * (nonSaudi + saudi) - saudi) / (1 - t);
  return Math.max(0, Math.ceil(n));
}

async function getCompanyStats(companyKey) {
  const cfg = COMPANIES[companyKey];
  const wb = await loadWorkbook();
  const saudis = sheetToRows(wb, `${cfg.key}_سعوديين`);
  const nonSaudis = sheetToRows(wb, `${cfg.key}_غير سعوديين`);
  const total = saudis.length + nonSaudis.length;
  const pct = total > 0 ? (saudis.length / total) * 100 : 0;
  const band = getBand(pct, cfg.bands);
  const nb = nextBand(pct, cfg.bands);

  const natCounts = {};
  nonSaudis.forEach(p => {
    const n = p.nationality || 'غير محدد';
    natCounts[n] = (natCounts[n] || 0) + 1;
  });

  const jobCounts = {};
  [...saudis, ...nonSaudis].forEach(p => {
    const j = p.job || 'غير محدد';
    jobCounts[j] = (jobCounts[j] || 0) + 1;
  });

  // نسبة توطين كل مهنة على حدة (سعودي ÷ إجمالي المهنة)
  const jobBreakdown = {};
  saudis.forEach(p => {
    const j = p.job || 'غير محدد';
    if (!jobBreakdown[j]) jobBreakdown[j] = { saudi: 0, nonSaudi: 0 };
    jobBreakdown[j].saudi++;
  });
  nonSaudis.forEach(p => {
    const j = p.job || 'غير محدد';
    if (!jobBreakdown[j]) jobBreakdown[j] = { saudi: 0, nonSaudi: 0 };
    jobBreakdown[j].nonSaudi++;
  });
  const jobLocalization = Object.entries(jobBreakdown).map(([job, c]) => {
    const t = c.saudi + c.nonSaudi;
    return { job, saudi: c.saudi, nonSaudi: c.nonSaudi, total: t, pct: t > 0 ? (c.saudi / t) * 100 : 0 };
  }).sort((a, b) => b.total - a.total);

  const genderCounts = { 'ذكر': 0, 'أنثى': 0 };
  [...saudis, ...nonSaudis].forEach(p => {
    const g = p.gender.replace('انثى', 'أنثى').trim();
    if (g === 'ذكر' || g === 'أنثى') genderCounts[g]++;
  });

  return {
    cfg, saudis, nonSaudis, total,
    saudiCount: saudis.length,
    nonSaudiCount: nonSaudis.length,
    pct, band, nextBand: nb,
    neededForNext: nb ? neededToReach(saudis.length, total, nb.min) : 0,
    natCounts, jobCounts, genderCounts, jobLocalization,
    all: [...saudis.map(p => ({ ...p, group: 'سعودي' })), ...nonSaudis.map(p => ({ ...p, group: 'غير سعودي' }))]
  };
}

async function getAllStats() {
  const keys = Object.keys(COMPANIES);
  const results = await Promise.all(keys.map(getCompanyStats));
  const map = {};
  keys.forEach((k, i) => map[k] = results[i]);
  return map;
}

/* ---------------- الرسم: مسطرة النطاق ---------------- */
function renderRuler(container, bands, pct, scaleCap) {
  container.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.className = 'ruler-wrap';

  const title = document.createElement('div');
  title.className = 'ruler-title';
  title.textContent = 'مسطرة نطاقات التوطين';
  wrap.appendChild(title);

  const track = document.createElement('div');
  track.className = 'ruler-track';

  const segWrap = document.createElement('div');
  segWrap.className = 'ruler-segments';
  bands.forEach(b => {
    const top = Math.min(b.max, scaleCap);
    const w = ((top - b.min) / scaleCap) * 100;
    const seg = document.createElement('div');
    seg.className = 'seg';
    seg.style.width = w + '%';
    seg.style.background = b.color;
    seg.title = `${b.name} (${b.min}%–${b.max >= 100 ? b.max + '%+' : b.max + '%'})`;
    track.dataset.ready = '1';
    segWrap.appendChild(seg);
  });

  const pointer = document.createElement('div');
  pointer.className = 'ruler-pointer';
  const clamped = Math.min(pct, scaleCap);
  pointer.style.right = `calc(${(clamped / scaleCap) * 100}% )`;
  pointer.innerHTML = `<div class="flag">${pct.toFixed(1)}%</div><div class="stem"></div>`;

  const labels = document.createElement('div');
  labels.className = 'ruler-labels';
  bands.forEach(b => {
    const top = Math.min(b.max, scaleCap);
    const w = ((top - b.min) / scaleCap) * 100;
    const lab = document.createElement('div');
    lab.className = 'lab';
    lab.style.width = w + '%';
    lab.textContent = b.min + '%';
    labels.appendChild(lab);
  });

  track.appendChild(pointer);
  track.appendChild(segWrap);
  wrap.appendChild(track);
  wrap.appendChild(labels);
  container.appendChild(wrap);
}

/* ---------------- الرسم: عداد نصف دائري (Gauge) ---------------- */
function renderGauge(svgEl, pct, band, scaleCap) {
  const size = 260, cx = size / 2, cy = size / 2 + 10, r = 100, sw = 22;
  const clamped = Math.min(Math.max(pct, 0), scaleCap);
  const angle = (clamped / scaleCap) * 180;

  function polar(cxx, cyy, rr, deg) {
    const rad = (Math.PI / 180) * (180 - deg);
    return { x: cxx - rr * Math.cos(rad), y: cyy - rr * Math.sin(rad) };
  }
  function arcPath(a0, a1) {
    const p0 = polar(cx, cy, r, a0);
    const p1 = polar(cx, cy, r, a1);
    const large = (a1 - a0) > 180 ? 1 : 0;
    return `M ${p0.x} ${p0.y} A ${r} ${r} 0 ${large} 1 ${p1.x} ${p1.y}`;
  }

  let s = `<svg viewBox="0 0 ${size} ${size*0.62}" xmlns="http://www.w3.org/2000/svg">`;
  s += `<path d="${arcPath(0,180)}" fill="none" stroke="var(--line)" stroke-width="${sw}" stroke-linecap="round"/>`;
  s += `<path d="${arcPath(0, angle)}" fill="none" stroke="${band.color}" stroke-width="${sw}" stroke-linecap="round"/>`;
  const needle = polar(cx, cy, r + 16, angle);
  s += `<line x1="${cx}" y1="${cy}" x2="${needle.x}" y2="${needle.y}" stroke="var(--ink)" stroke-width="2.5" stroke-linecap="round"/>`;
  s += `<circle cx="${cx}" cy="${cy}" r="5" fill="var(--ink)"/>`;
  s += `</svg>`;
  svgEl.innerHTML = s;
}

/* ---------------- أدوات مساعدة عامة ---------------- */
function cssColor(varExpr) {
  // يحوّل "var(--x)" إلى قيمة hex فعلية لاستخدامه في Chart.js
  const name = varExpr.replace('var(', '').replace(')', '').trim();
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}

function topEntries(obj, n) {
  return Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, n);
}

function initNav() {
  const btn = document.querySelector('.menu-btn');
  const nav = document.querySelector('.mainnav');
  if (btn && nav) {
    btn.addEventListener('click', () => nav.classList.toggle('open'));
  }
  const path = location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.mainnav a').forEach(a => {
    if (a.getAttribute('href') === path) a.classList.add('active');
  });
}

document.addEventListener('DOMContentLoaded', initNav);
