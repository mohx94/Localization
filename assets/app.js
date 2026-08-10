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
    id: 'admin_support', title: 'المهن الإدارية المساندة',
    decree: 'قرار وزاري رقم 132249 وتاريخ 2026/4/5م', pct: 100, minEmployees: 1,
    effective: 'من 2026/4/5م للمهن الفورية، وباقي المهن من 2026/10/4م', effectiveISO: '2026-04-05', delayedEffectiveISO: '2026-10-04',
    delayedProfessions: ['مدير تنظيم إداري','مدير تطوير مؤسسي','مدير مكتب خدمة مدنية','مدير مكتب عمل','مدير مكتب استقدام','مدير تطوير موارد بشرية','مدير مواهب','مدير تعويضات','مدير توظيف','مدير تصنيف مهن','مدير تخطيط قوى عاملة','مدير عمليات موارد بشرية','مدير موارد بشرية أمن سيبراني','مدير تحقيق إداري','محقق عمالي','أخصائي توجيه مهني','أخصائي تعويضات','أخصائي توظيف','أخصائي تصنيف وظيفي ومهني','أخصائي عمليات موارد بشرية','أخصائي ارتباط وظيفي','مدير مراسم','مدير جمع التبرعات','مدير علاقات عامة','خبير علاقات عامة','مستشار علاقات عامة','أخصائي مراسم','أخصائي تنظيم مؤتمرات وفعاليات','أخصائي توعية جمهور','أخصائي تواصل داخلي','أخصائي علاقات عامة','خبير موارد بشرية','مستشار موارد بشرية','أخصائي استقدام','أخصائي مراقبة موارد بشرية','أخصائي تخطيط قوى عاملة','عضو هيئة عمالية','مترجم لغة إشارة','مصحح لغوي','مراقب كاميرات أمنية','كاتب حركة مخزون','كاتب شحن','موظف استقبال','كاتب استعلامات','طباع','مساعد إداري','كاتب سجل','كاتب علاقات حكومية','وكيل جمركي','وكيل شحن'],
    note: 'قصر العمل بالكامل على السعوديين. تنبيه: بعض المهن لها تاريخ تطبيق لاحق حسب الدليل.',
    professions: ['مدير علاقات عامة','مدير موارد بشرية أمن سيبراني','مدير تحقيق إداري','مدير تنظيم إداري','مدير تطوير مؤسسي','مدير مكتب خدمة مدنية','مدير شؤون عمل','مدير مكتب عمل','مدير مكتب استقدام','مدير علاقات الأفراد','مدير تطوير موارد بشرية','مدير مواهب','مدير تعويضات','مدير توظيف','مدير تصنيف مهن','مدير تخطيط قوى عاملة','مدير عمليات موارد بشرية','أخصائي علاقات عامة','خبير موارد بشرية','مستشار موارد بشرية','أخصائي استقدام','أخصائي مراقبة موارد بشرية','أخصائي تخطيط قوى عاملة','عضو هيئة عمالية','محقق عمالي','أخصائي توجيه مهني','أخصائي تعويضات','أخصائي توظيف','أخصائي تصنيف وظيفي ومهني','أخصائي عمليات موارد بشرية','أخصائي ارتباط وظيفي','أخصائي تنظيم مؤتمرات وفعاليات','مدير مراسم','مدير جمع التبرعات','أخصائي توعية جمهور','أخصائي تواصل داخلي','أخصائي مراسم','مستشار علاقات عامة','خبير علاقات عامة','كاتب سجل','كاتب علاقات حكومية','سكرتير تنفيذي','مخلص جمركي','وكيل جمركي','وكيل شحن','أخصائي لغوي','مترجم لغة إشارة','مصحح لغوي','مترجم فوري','مترجم','أمين صندوق','سكرتير','كاتب استقبال مرضى','كاتب شكاوى','موظف استقبال فندق','مدخل بيانات','طباع','كاتب اختزال','حارس شخصي','حارس أمن','مراقب كاميرات أمنية','كاتب موارد بشرية','ناسخ','كاتب حركة مخزون','كاتب شحن','أمين مخزن','موظف استقبال','كاتب استعلامات','مساعد إداري']
  },
  {
    id: 'customer_service', title: 'مهن خدمة العملاء',
    decree: 'قرار وزاري رقم 208892 وتاريخ 1443/11/23هـ', pct: 100, minEmployees: 1,
    effective: '1444/05/23هـ', activeBySource: true,
    note: '100% لخدمة العملاء المقدمة بالعربية والإنجليزية، ويشمل مدراء العمليات وقادة/مشرفي فرق مراكز وإدارات خدمة العملاء. القرار يرتبط أيضاً بالعمل الفعلي؛ المطابقة النصية هنا أداة إنذار فقط.',
    professions: ['كاتب مركز اتصالات','كاتب استعلامات مركز خدمة عملاء','كاتب استعلامات','كاتب بيانات عملاء','أخصائي خدمة عملاء','ممثل خدمة عملاء','موظف خدمة عملاء','مدير خدمة عملاء','مدير عمليات مركز خدمة عملاء','مشرف خدمة عملاء']
  },
  {
    id: 'sales', title: 'مهن المبيعات', decree: 'قرار وزاري رقم 101278 وتاريخ 2026/01/19م',
    pct: 60, minEmployees: 3, effective: '2026/04/19م', effectiveISO: '2026-04-19', note: 'يُحتسب على مستوى الكيان عند وجود 3 عاملين فأكثر في المهن المستهدفة.',
    professions: ['مدير مبيعات','مدير مبيعات تجزئة','مدير مبيعات جملة','مندوب مبيعات','وسيط سلع مستقبلية','أخصائي مبيعات أجهزة تقنية المعلومات والاتصالات','أخصائي مبيعات','أخصائي تجاري','وسيط سلع']
  },
  {
    id: 'purchasing', title: 'مهن المشتريات', decree: 'قرار وزاري رقم 77050 وتاريخ 2025/11/30م',
    pct: 70, minEmployees: 3, effective: '2026/05/31م', effectiveISO: '2026-05-31', note: '',
    professions: ['مدير مشتريات','مندوب مشتريات','مدير عقود','مدير خدمات لوجستية','مدير مستودع','أمين مستودع','أخصائي مناقصات','أخصائي مشتريات','أخصائي توريد للعلامات التجارية الخاصة','أخصائي تجارة إلكترونية','أخصائي أبحاث أسواق','أخصائي مستودعات']
  },
  {
    id: 'marketing', title: 'مهن التسويق', decree: 'قرار وزاري رقم 101319 وتاريخ 2026/01/19م',
    pct: 60, minEmployees: 3, effective: '2026/04/19م', effectiveISO: '2026-04-19', note: 'لا يُحتسب السعودي إن قل أجره الخاضع للتأمينات عن 5,500 ريال؛ ملف الموظفين الحالي لا يحتوي الأجر لذلك الرقم الرسمي من قوى هو الأدق.',
    professions: ['مدير تسويق','وكيل دعاية وإعلان','مدير دعاية وإعلان','أخصائي علاقات عامة','أخصائي دعاية وإعلان','أخصائي تسويق','مدير علاقات عامة','مصمم جرافيك','مصمم إعلان','مصور فوتوغرافي','أخصائي إعلامي']
  },
  {
    id: 'accounting', title: 'المهن المحاسبية', decree: 'قرار وزاري رقم 103108 وتاريخ 2025/01/26م',
    pct: 40, minEmployees: 5, effective: 'تدريجي: 40% من 2025/10/27م، 50% من 2026/10/27م، 60% من 2027/10/27م، 70% من 2028/10/27م، وفي 2029/10/27م يبدأ 30% للمنشآت التي لديها 3–4 محاسبين',
    phases: [{date:'2025-10-27',pct:40,minEmployees:5},{date:'2026-10-27',pct:50,minEmployees:5},{date:'2027-10-27',pct:60,minEmployees:5},{date:'2028-10-27',pct:70,minEmployees:5},{date:'2029-10-27',pct:30,minEmployees:3,smallOnly:true}], note: 'يشترط الاعتماد المهني وحد أدنى للأجر؛ الحساب النصي تقديري ما لم تُدخل أرقام قوى الرسمية.',
    professions: ['مدير مالي','مدير حسابات','مدير حسابات تعرفة','مدير خزينة','مدير ميزانية','مدير مراجعة','مدير مراجعة داخلية','مدير تحصيل','مدير خزانة','محاسب تكاليف','مراقب مالي','مراجع داخلي','كاتب حسابات','محاسب','أخصائي موازنة مالية','أخصائي حسابات ضرائب','محاسب قانوني','مسؤول ضرائب','مساعد حسابات','أخصائي مراقبة مخزون','مراقب مخزون','كاتب مالي']
  },
  {
    id: 'engineering', title: 'المهن الهندسية', decree: 'قرار وزاري رقم 93483 وتاريخ 2025/12/31م',
    pct: 30, minEmployees: 5, effective: '2026/06/30م', effectiveISO: '2026-06-30', note: 'يُحتسب على مستوى الكيان عند وجود 5 مهندسين فأكثر، والحد الأدنى للأجر المحتسب 8,000 ريال مع متطلبات الاعتماد المهني.',
    professions: ['مهندس مواد','مهندس مدني','مهندس تعدين','مهندس معماري','مهندس ميكانيكي آليات ثقيلة','مهندس إنتاج تقني','مهندس مدني تقني','مهندس جيوتقني','مهندس إنشائي','مهندس معالجة مياه','مهندس صحي','مهندس ميكانيكي تقني','مهندس ميكانيكي قطارات','مهندس توربينات','مهندس ميكاترونكس','مهندس كيميائي تقني','مهندس تحلية مياه','مهندس نفط وغاز','مهندس تكاليف','مهندس متفجرات','مهندس ضبط وتحكم في الجودة','مهندس بحث وتطوير','مهندس كهربائي تقني','مهندس كهروميكانيكي','مهندس أتمتة','مهندس توليد طاقة','مهندس معماري داخلي','مهندس طرق','مهندس جسور','مهندس طيران','مهندس كهربائي تمديدات','مهندس معادن','مهندس ميكانيكي لحام','مهندس إنتاج','مهندس كهربائي','مهندس ميكانيكي','مهندس تصميم مواقع','مهندس صناعي','مهندس إلكترونيات','مهندس بيئي','مهندس تدفئة وتهوية وتكييف','مهندس مركبات','مهندس كيميائي','مهندس نووي','مهندس بحري','مهندس نقل وتوزيع طاقة']
  },
  {
    id: 'legal', title: 'المهن القانونية', decree: 'قرار وزاري رقم 212607 وتاريخ 1442/11/24هـ',
    pct: 70, minEmployees: 1, effective: 'نافذ حسب الدليل المرفق', activeBySource: true, note: '70% على مستوى الكيان، والحد الأدنى للأجر المحتسب 5,500 ريال.',
    professions: ['مدير شؤون قانونية','أخصائي قانوني','أخصائي عقود','سكرتير قانوني']
  },
  {
    id: 'project_management', title: 'مهن إدارة المشاريع', decree: 'قرار وزاري رقم 141749 وتاريخ 1444/09/11هـ',
    pct: 40, minEmployees: 3, effective: 'المرحلة الثانية 40% من 1446/06/11هـ', activeBySource: true, note: 'يُحتسب على مستوى الكيان عند وجود 3 عاملين فأكثر، والحد الأدنى للأجر المحتسب 6,000 ريال.',
    professions: ['مدير اتصالات','مدير هندسة اتصالات','مدير إدارة مشاريع','أخصائي إدارة مشاريع','مهندس إدارة مشاريع']
  }
];

// قرارات منافذ البيع مستبعدة عمداً من لوحة المخالفات الإلكترونية؛ تتطلب تحققاً ميدانياً.
const OUTLET_DECISION = null;

function normalizeAr(s){
  return (s || '')
    .trim()
    .replace(/[\u064B-\u065F\u0670]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/\s+/g, ' ');
}

function officialRound(value) {
  return Math.floor(Number(value) + 0.5);
}
function requiredSaudiCount(total, pct) {
  return officialRound(Number(total || 0) * (Number(pct || 0) / 100));
}
function dateOnly(value) {
  const d = value instanceof Date ? new Date(value) : new Date(value + 'T00:00:00');
  d.setHours(0,0,0,0); return d;
}
function todayLocal() { const d = new Date(); d.setHours(0,0,0,0); return d; }
function daysUntil(iso, now=todayLocal()) {
  if(!iso) return null;
  return Math.ceil((dateOnly(iso) - now) / 86400000);
}
function fmtDate(iso){
  if(!iso) return '';
  return dateOnly(iso).toLocaleDateString('ar-SA-u-ca-gregory', {year:'numeric',month:'long',day:'numeric'});
}

function getDecisionRule(decision, totalMatched=0, now=todayLocal()) {
  // المحاسبة قرار متدرج ويجب أن تتغير النسبة آلياً مع التاريخ.
  if (decision.id === 'accounting' && decision.phases) {
    const regular = decision.phases.filter(p => !p.smallOnly);
    let current = null, next = null;
    for (const p of regular) {
      if (dateOnly(p.date) <= now) current = p;
      else { next = p; break; }
    }
    if (!current) return {active:false,pct:40,minEmployees:5,nextChange:regular[0]};
    // من 2029/10/27: المنشآت ذات 3-4 محاسبين تدخل بنسبة 30%، و5 فأكثر تستمر 70%.
    const smallPhase = decision.phases.find(p => p.smallOnly);
    if (smallPhase && dateOnly(smallPhase.date) <= now && totalMatched >= 3 && totalMatched <= 4) {
      return {active:true,pct:30,minEmployees:3,nextChange:null,phaseLabel:'مرحلة 3–4 محاسبين'};
    }
    if (smallPhase && dateOnly(smallPhase.date) <= now && totalMatched >= 5) {
      return {active:true,pct:70,minEmployees:5,nextChange:null,phaseLabel:'استمرار المرحلة الرابعة'};
    }
    return {active:true,pct:current.pct,minEmployees:current.minEmployees,nextChange:next,phaseLabel:`المرحلة الحالية ${current.pct}%`};
  }
  let active = true;
  if (decision.effectiveISO) active = dateOnly(decision.effectiveISO) <= now;
  else if (decision.activeBySource) active = true;
  return {active,pct:decision.pct,minEmployees:decision.minEmployees,nextChange:null};
}

function isProfessionActive(decision, job, now=todayLocal()) {
  const normalized = normalizeAr(job);
  if (decision.effectiveISO && dateOnly(decision.effectiveISO) > now) return false;
  if (decision.id === 'admin_support' && decision.delayedProfessions && decision.delayedEffectiveISO) {
    const delayed = new Set(decision.delayedProfessions.map(normalizeAr));
    if (delayed.has(normalized) && dateOnly(decision.delayedEffectiveISO) > now) return false;
  }
  return true;
}

function computeDecisionCompliance(stats, decision, now=todayLocal()) {
  const targetSet = new Set(decision.professions.map(normalizeAr));
  const activeMatched = [], futureMatched = [];
  stats.all.forEach(p => {
    if (!targetSet.has(normalizeAr(p.job))) return;
    (isProfessionActive(decision, p.job, now) ? activeMatched : futureMatched).push(p);
  });

  const saudiMatched = activeMatched.filter(p=>p.group==='سعودي').length;
  const nonSaudiMatched = activeMatched.length - saudiMatched;
  const totalMatched = activeMatched.length;
  const rule = getDecisionRule(decision, totalMatched, now);
  const applicable = rule.active && totalMatched >= rule.minEmployees;
  const requiredCount = applicable ? requiredSaudiCount(totalMatched, rule.pct) : 0;
  const gap = applicable ? Math.max(0, requiredCount - saudiMatched) : 0;
  const jobsFound = [...new Set(activeMatched.map(p => p.job))];
  const futureJobsFound = [...new Set(futureMatched.map(p => p.job))];

  return {
    decision, rule, totalMatched, saudiMatched, nonSaudiMatched, applicable, requiredCount, gap,
    jobsFound, futureJobsFound, futureMatchedCount:futureMatched.length,
    nonSaudiEmployees: activeMatched.filter(p=>p.group!=='سعودي'),
    futureNonSaudiEmployees: futureMatched.filter(p=>p.group!=='سعودي'),
    compliant: applicable && gap === 0,
    isOfficial:false
  };
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
    const totalN = Number(total) || 0;
    const saudiN = Number(saudi) || 0;
    const nonSaudiN = Number(nonSaudi) || 0;
    const mismatch = (saudi !== '' && nonSaudi !== '' && totalN !== saudiN + nonSaudiN);
    OFFICIAL_OVERRIDES[company + '|' + decision] = {
      total: totalN, saudi: saudiN, nonSaudi: nonSaudiN,
      note: String(r[5] ?? '').trim(),
      mismatch
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
function getDecisionDisplay(stats, decision, wb, now=todayLocal()) {
  const overrides = loadOfficialOverrides(wb);
  const key = stats.cfg.name + '|' + decision.title;
  const override = overrides[key];

  if (override) {
    const rule = getDecisionRule(decision, override.total, now);
    const applicable = rule.active && override.total >= rule.minEmployees;
    const requiredCount = applicable ? requiredSaudiCount(override.total, rule.pct) : 0;
    const gap = applicable ? Math.max(0, requiredCount - override.saudi) : 0;
    return {
      decision, rule, totalMatched: override.total, saudiMatched: override.saudi, nonSaudiMatched: override.nonSaudi,
      applicable, requiredCount, gap, jobsFound: [], futureJobsFound:[], futureMatchedCount:0,
      nonSaudiEmployees:[], futureNonSaudiEmployees:[], compliant: applicable && gap === 0,
      isOfficial: true, officialNote: override.note, officialMismatch: override.mismatch
    };
  }
  return computeDecisionCompliance(stats, decision, now);
}

function getDecisionStatus(result) {
  if (result.officialMismatch) return {key:'review',label:'تدقيق مطلوب',icon:'⚠',rank:4};
  if (!result.rule.active) return {key:'upcoming',label:'لم يبدأ بعد',icon:'⏳',rank:2};
  if (!result.applicable) {
    if (result.futureMatchedCount > 0) return {key:'upcoming',label:'التزام قادم',icon:'⏳',rank:2};
    return {key:'na',label:'لا ينطبق حالياً',icon:'—',rank:0};
  }
  if (result.gap > 0) return {key:'critical',label:'مخالفة محتملة',icon:'🔴',rank:5};
  return {key:'ok',label:'ملتزم',icon:'🟢',rank:1};
}

function getCompanyDecisionResults(stats, now=todayLocal()) {
  return PROFESSION_DECISIONS.map(d => getDecisionDisplay(stats, d, WORKBOOK, now));
}

function getCompanyRiskSummary(stats, now=todayLocal()) {
  const results = getCompanyDecisionResults(stats, now);
  const critical = results.filter(r=>getDecisionStatus(r).key==='critical');
  const review = results.filter(r=>getDecisionStatus(r).key==='review');
  const upcoming = getUpcomingComplianceEvents(stats, now, 180);
  return {
    results, critical, review, upcoming,
    totalGap: critical.reduce((s,r)=>s+r.gap,0),
    riskCount: critical.length + review.length
  };
}

function getUpcomingComplianceEvents(stats, now=todayLocal(), horizonDays=180) {
  const events = [];
  const results = getCompanyDecisionResults(stats, now);
  results.forEach(r=>{
    const d = r.decision;
    // المهن الإدارية المؤجلة حتى 4 أكتوبر 2026.
    if (d.id==='admin_support' && d.delayedEffectiveISO && r.futureMatchedCount>0) {
      const days = daysUntil(d.delayedEffectiveISO, now);
      if (days >= 0 && days <= horizonDays) {
        const futureNonSaudi = r.futureNonSaudiEmployees.length;
        events.push({decision:d,date:d.delayedEffectiveISO,days,type:'activation',
          title:'بدء تطبيق مجموعة المهن الإدارية المساندة المؤجلة',
          detail:`${r.futureMatchedCount} موظف مطابق حالياً، منهم ${futureNonSaudi} غير سعودي.`});
      }
    }
    // المرحلة التالية للمحاسبة.
    if (r.rule.nextChange && r.totalMatched >= r.rule.minEmployees) {
      const days = daysUntil(r.rule.nextChange.date, now);
      if (days >= 0 && days <= horizonDays) {
        const nextRequired = requiredSaudiCount(r.totalMatched, r.rule.nextChange.pct);
        const futureGap = Math.max(0, nextRequired - r.saudiMatched);
        events.push({decision:d,date:r.rule.nextChange.date,days,type:'phase',
          title:`ارتفاع توطين المحاسبة إلى ${r.rule.nextChange.pct}%`,
          detail:`بالعدد الحالي سيصبح المطلوب ${nextRequired} سعودي، والفجوة المتوقعة ${futureGap}.`});
      }
    }
    // أي قرار كامل لم يبدأ بعد.
    if (!r.rule.active && d.effectiveISO) {
      const days = daysUntil(d.effectiveISO, now);
      if (days >= 0 && days <= horizonDays) events.push({decision:d,date:d.effectiveISO,days,type:'activation',title:`بدء تطبيق ${d.title}`,detail:d.note||''});
    }
  });
  return events.sort((a,b)=>dateOnly(a.date)-dateOnly(b.date));
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


function injectComplianceStyles(){
  if(document.getElementById('electronic-risk-styles')) return;
  const s=document.createElement('style'); s.id='electronic-risk-styles';
  s.textContent=`
  .risk-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:14px 0 18px}.risk-kpi{border:1px solid var(--line);border-radius:12px;padding:14px;background:white}.risk-kpi b{display:block;font-size:25px}.risk-kpi small{color:var(--ink-soft)}
  .risk-board{display:grid;gap:10px}.risk-row{border:1px solid var(--line);border-radius:12px;padding:13px 14px;background:white;display:grid;grid-template-columns:1.3fr .7fr .8fr 1fr;gap:10px;align-items:center}.risk-row.critical{border-inline-start:5px solid #b83232}.risk-row.review{border-inline-start:5px solid #b77a16}.risk-row.upcoming{border-inline-start:5px solid #607d8b}.risk-row.ok{border-inline-start:5px solid #4e7e52}
  .risk-badge{display:inline-block;border-radius:999px;padding:4px 9px;font-size:11px;font-weight:700}.risk-badge.critical{background:#fbe9e8;color:#a62222}.risk-badge.review{background:#fff3dc;color:#855600}.risk-badge.upcoming{background:#eef3f5;color:#455a64}.risk-badge.ok{background:#e8f3e8;color:#326737}.risk-badge.na{background:#f4f4f4;color:#777}
  .risk-people{margin-top:10px;border-top:1px dashed var(--line);padding-top:9px}.risk-person{display:flex;gap:8px;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--line-soft);font-size:12px}.risk-person:last-child{border-bottom:0}.countdown{font-weight:800}.source-note{font-size:11px;color:var(--ink-soft)}
  @media(max-width:800px){.risk-kpis{grid-template-columns:repeat(2,1fr)}.risk-row{grid-template-columns:1fr}.risk-person{display:block}}
  `; document.head.appendChild(s);
}

function renderElectronicRiskPanel(stats){
  injectComplianceStyles();
  let section=document.getElementById('electronic-risk-panel');
  if(!section){
    section=document.createElement('section'); section.id='electronic-risk-panel'; section.style.paddingTop='0';
    const decisionSection=document.getElementById('decisions-section');
    if(decisionSection) decisionSection.insertAdjacentElement('afterend',section); else return;
  }
  const summary=getCompanyRiskSummary(stats);
  const estimatedRisks=summary.critical.filter(r=>!r.isOfficial);
  const peopleCount=estimatedRisks.reduce((n,r)=>n+r.nonSaudiEmployees.length,0);
  const upcoming=summary.upcoming;
  const rows=summary.results.filter(r=>['critical','review'].includes(getDecisionStatus(r).key));

  const riskHtml = rows.length ? rows.map(r=>{
    const st=getDecisionStatus(r);
    const pct=r.rule.pct;
    let people='';
    if(st.key==='critical'){
      if(r.isOfficial){ people=`<div class="risk-people source-note">النتيجة من أرقام قوى الرسمية؛ ملف قوى المجمّع لا يحدد أسماء الموظفين المسببين للفجوة.</div>`; }
      else if(r.nonSaudiEmployees.length){
        people=`<div class="risk-people"><div class="source-note">غير السعوديين داخل المهن المستهدفة — يلزم معالجة ${r.gap} من الفجوة الحالية:</div>${r.nonSaudiEmployees.map(p=>`<div class="risk-person"><span><b>${p.name||'—'}</b></span><span>${p.job||'—'}</span></div>`).join('')}</div>`;
      }
    }
    return `<div class="risk-row ${st.key}"><div><b>${r.decision.title}</b><div class="source-note">${r.decision.decree}</div>${people}</div><div><span class="risk-badge ${st.key}">${st.icon} ${st.label}</span></div><div><b>${pct}%</b><div class="source-note">المطلوب ${r.requiredCount} / الفعلي ${r.saudiMatched}</div></div><div><b>${r.gap ? 'فجوة '+r.gap : (r.officialMismatch?'راجع الأرقام':'—')}</b><div class="source-note">${r.isOfficial?'معتمد من قوى':'تقدير من المسميات'}</div></div></div>`;
  }).join('') : `<div class="alert ok"><span class="ico">✓</span><div class="body"><strong>لا توجد مخالفات إلكترونية محتملة ظاهرة من البيانات الحالية</strong><span>يبقى التحقق من قوى هو المرجع النهائي، خصوصاً للقرارات المرتبطة بالأجر أو الاعتماد المهني.</span></div></div>`;

  const upcomingHtml=upcoming.length ? `<div style="margin-top:18px"><h3 style="font-size:15px;margin-bottom:8px">📅 التزامات قادمة</h3>${upcoming.map(e=>`<div class="risk-row upcoming"><div><b>${e.title}</b><div class="source-note">${e.detail}</div></div><div><span class="risk-badge upcoming">⏳ قادم</span></div><div><span class="countdown">${e.days===0?'اليوم':e.days+' يوم'}</span></div><div>${fmtDate(e.date)}</div></div>`).join('')}</div>`:'';

  section.innerHTML=`<div class="wrap"><div class="section-head"><h2>🚨 الإنذار المبكر للمخالفات الإلكترونية</h2><span class="hint">منافذ البيع مستبعدة لأنها تتطلب تحققاً ميدانياً</span></div><div class="risk-kpis"><div class="risk-kpi"><b>${summary.critical.length}</b><small>مخالفات محتملة</small></div><div class="risk-kpi"><b>${summary.totalGap}</b><small>إجمالي فجوة السعوديين</small></div><div class="risk-kpi"><b>${peopleCount}</b><small>غير سعوديين ضمن مهن الخطر (تقديري)</small></div><div class="risk-kpi"><b>${upcoming.length}</b><small>تغييرات خلال 180 يوم</small></div></div><div class="risk-board">${riskHtml}</div>${upcomingHtml}</div>`;
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

document.addEventListener('DOMContentLoaded', ()=>{ initNav(); injectComplianceStyles(); });
