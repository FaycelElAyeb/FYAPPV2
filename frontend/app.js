/* =====================================================
   نظام التنبؤ الأكاديمي - يعمل 100% في المتصفح
   بدون Python - بدون Server - بدون تثبيت
   ===================================================== */

let allStudents = [];
let charts = {};
let reportData = null;

// ---- DOM ----
const gradebookInput = document.getElementById('gradebookInput');
const analyticsInput = document.getElementById('analyticsInput');
const analyzeBtn     = document.getElementById('analyzeBtn');
const loadingOverlay = document.getElementById('loadingOverlay');
const uploadSection  = document.getElementById('uploadSection');
const reportSection  = document.getElementById('reportSection');
const studentsBody   = document.getElementById('studentsBody');
const kpiGrid        = document.getElementById('kpiGrid');
const searchInput    = document.getElementById('searchInput');

// ---- File selection ----
gradebookInput.addEventListener('change', () => {
  if (gradebookInput.files[0]) {
    document.getElementById('gradebookName').textContent = gradebookInput.files[0].name;
    document.getElementById('gradebookCheck').textContent = '✅';
    document.getElementById('gradebookBox').classList.add('has-file');
  }
  checkReady();
});

analyticsInput.addEventListener('change', () => {
  if (analyticsInput.files[0]) {
    document.getElementById('analyticsName').textContent = analyticsInput.files[0].name;
    document.getElementById('analyticsCheck').textContent = '✅';
    document.getElementById('analyticsBox').classList.add('has-file');
  }
  checkReady();
});

function checkReady() {
  analyzeBtn.disabled = !(gradebookInput.files[0] && analyticsInput.files[0]);
}

// ---- Analyze ----
analyzeBtn.addEventListener('click', async () => {
  loadingOverlay.style.display = 'flex';
  try {
    await new Promise(r => setTimeout(r, 300));

    // إرسال الملفات إلى Flask backend
const formData = new FormData();
formData.append("gradebook", gradebookInput.files[0]);
formData.append("analytics", analyticsInput.files[0]);

const res = await fetch("/api/analyze", {
  method: "POST",
  body: formData
});

if (!res.ok) {
  throw new Error("فشل الاتصال بالخادم");
}

reportData = await res.json();
    
   
    allStudents  = reportData.students;
    renderReport(reportData);
    uploadSection.style.display = 'none';
    reportSection.style.display = 'block';
    window.scrollTo({ top: 0, behavior: 'smooth' });

  } catch (e) {
    alert('❌ ' + e.message);
  } finally {
    loadingOverlay.style.display = 'none';
  }
});

// ---- قراءة الملف (CSV أو XLS أو XLSX) ----
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split('.').pop().toLowerCase();
    const reader = new FileReader();

    if (ext === 'csv') {
      reader.onload = e => {
        try {
          const text = e.target.result;
          const data = parseCSV(text);
          if (!data.length) throw new Error('الملف فارغ أو لا يحتوي على بيانات');
          console.log(`✅ تم قراءة CSV: ${file.name} | ${data.length} صف`);
          resolve(data);
        } catch (err) { reject(err); }
      };
      reader.onerror = () => reject(new Error('فشل قراءة ملف CSV'));
      reader.readAsText(file, 'UTF-8');

    } else if (ext === 'xls' || ext === 'xlsx') {
      reader.onload = e => {
        try {
          const wb  = XLSX.read(e.target.result, { type: 'array', codepage: 65001 });
          const ws  = wb.Sheets[wb.SheetNames[0]];

          // قراءة كـ array خام للحصول على القيم بالموضع الدقيق
          const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (arr.length < 2) throw new Error('الملف فارغ');

          // اكتشاف صف العناوين الحقيقي
          // صف العناوين = الصف الذي تكون فيه أغلب القيم نصوصاً (وليس أرقاماً)
          let headerRowIdx = 0;
          let bestTextCount = 0;
          for (let i = 0; i < Math.min(arr.length, 5); i++) {
            const textCount = arr[i].filter(c => {
              const s = String(c||'').trim();
              return s.length > 0 && isNaN(Number(s)) && !s.match(/^\d{4}-\d{2}-\d{2}/);
            }).length;
            if (textCount > bestTextCount) {
              bestTextCount = textCount;
              headerRowIdx = i;
            }
          }

          const headerRow = arr[headerRowIdx];
          // لا نفلتر الصفوف - نأخذ كل صف يحتوي على أي قيمة في أي خلية
          const dataRows  = arr.slice(headerRowIdx + 1).filter(r => 
            Array.isArray(r) && r.length > 0 && r.some(v => v !== '' && v !== null && v !== undefined)
          );

          console.log(`صف العناوين [${headerRowIdx}] | إجمالي صفوف البيانات: ${dataRows.length}`);

          // اكتشاف موضع عمودي الاسم
          // نبحث في صف العناوين عن الأعمدة التي تحتوي نصاً عربياً
          // ثم نتحقق أن قيم تلك الأعمدة في صفوف البيانات هي أسماء عربية
          let lastNameIdx  = -1;
          let firstNameIdx = -1;

          for (let ci = 0; ci < headerRow.length; ci++) {
            // فحص قيم هذا العمود في أول 5 صفوف من البيانات
            const colVals = dataRows.slice(0, 5).map(r => String(r[ci]||'').trim());
            const arabicNameCount = colVals.filter(v =>
              /[\u0600-\u06FF]/.test(v) &&   // يحتوي عربي
              v.length >= 2 &&               // طول معقول
              v.length <= 25 &&              // ليس طويلاً جداً
              !/\d{4}/.test(v) &&            // لا يحتوي سنة
              !/^\d/.test(v)                 // لا يبدأ برقم
            ).length;

            if (arabicNameCount >= 3) {
              if (lastNameIdx === -1) lastNameIdx = ci;
              else if (firstNameIdx === -1) { firstNameIdx = ci; break; }
            }
          }

          // إذا لم يُجد، استخدم العمودين 0 و 1 كـ fallback
          if (lastNameIdx  === -1) lastNameIdx  = 0;
          if (firstNameIdx === -1) firstNameIdx = 1;

          console.log('✅ عمودا الاسم:', {
            lastNameIdx, firstNameIdx,
            sampleLast:  String(dataRows[0]?.[lastNameIdx]  || ''),
            sampleFirst: String(dataRows[0]?.[firstNameIdx] || ''),
          });

          // بناء JSON مع مفاتيح ثابتة للاسم والمعرف
          const data = dataRows.map(row => {
            const obj = { __firstName: '', __lastName: '', __studentId: '' };
            headerRow.forEach((h, i) => {
              const key = String(h || '').trim();
              if (key) obj[key] = row[i] !== undefined ? row[i] : '';
              obj[`__c${i}`] = row[i] !== undefined ? row[i] : '';
            });
            obj.__firstName = String(row[firstNameIdx] || '').trim();
            obj.__lastName  = String(row[lastNameIdx]  || '').trim();
            // عمود C (index 2) = اسم المستخدم / معرف الطالب
            obj.__studentId = String(row[2] || '').trim();
            return obj;
          });

          // بناء map بالمعرف للبحث السريع
          const gbMap = {};
          data.forEach(r => {
            if (r.__studentId) gbMap[r.__studentId] = r;
          });
          data.__gbMap = gbMap;

          console.log(`✅ XLS: ${file.name} | ${data.length} صف`);
          console.log('🔑 معرفات Gradebook:', Object.keys(gbMap).slice(0,5));
          if (data[0]) {
            console.log('🔑 كل مفاتيح Gradebook:', JSON.stringify(Object.keys(data[0])));
            console.log('📋 أول صف كامل:', JSON.stringify(data[0]));
          }
          resolve(data);
        } catch (err) { reject(err); }
      };
      reader.onerror = () => reject(new Error('فشل قراءة ملف Excel'));
      reader.readAsArrayBuffer(file);

    } else {
      reject(new Error(`صيغة الملف غير مدعومة: .${ext}\nالصيغ المدعومة: CSV, XLS, XLSX`));
    }
  });
}

// ---- تحليل CSV إلى JSON مع اكتشاف صف العناوين الحقيقي ----
function parseCSV(text) {
  // إزالة BOM
  const cleaned = text.replace(/^\uFEFF/, '').trim();
  if (!cleaned) return [];

  // تحديد الفاصل تلقائياً
  const firstLine = cleaned.split('\n')[0];
  let sep = ',';
  if ((firstLine.match(/;/g)||[]).length > (firstLine.match(/,/g)||[]).length) sep = ';';
  else if ((firstLine.match(/\t/g)||[]).length > (firstLine.match(/,/g)||[]).length) sep = '\t';

  const lines = cleaned.split('\n').map(l => l.replace(/\r$/, ''));

  // ---- اكتشاف صف العناوين الحقيقي ----
  // نبحث عن أول صف يحتوي على كلمات مفتاحية معروفة
  const headerKeywords = [
    'اسم', 'الاسم', 'name', 'student', 'طالب',
    'grade', 'درجة', 'total', 'overall',
    'ساعات', 'hours', 'أيام', 'days', 'وصول', 'access',
    'username', 'مستخدم', 'activity', 'نشاط'
  ];

  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(lines.length, 10); i++) {
    const lineLower = lines[i].toLowerCase();
    const matchCount = headerKeywords.filter(kw => lineLower.includes(kw)).length;
    if (matchCount >= 2) {
      headerRowIdx = i;
      break;
    }
  }

  console.log(`=== CSV: صف العناوين المكتشف: السطر ${headerRowIdx} ===`, lines[headerRowIdx]);

  const headers = splitCSVLine(lines[headerRowIdx], sep).map(h => h.trim().replace(/^"|"$/g, ''));

  const rows = [];
  for (let i = headerRowIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const vals = splitCSVLine(line, sep);
    // تخطي الصفوف الفارغة أو التي تحتوي على metadata
    if (vals.every(v => !v.trim())) continue;
    const row = {};
    headers.forEach((h, idx) => {
      if (h) row[h] = (vals[idx] || '').trim().replace(/^"|"$/g, '');
    });
    rows.push(row);
  }
  return rows;
}

// ---- تقسيم سطر CSV مع دعم القيم بين علامات اقتباس ----
function splitCSVLine(line, sep) {
  const result = [];
  let current  = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i+1] === '"') { current += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === sep && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}

// ---- دالة استخراج اسم الطالب من صف Gradebook ----
function getStudentName(row) {
  const first = String(row.__firstName || '').trim();
  const last  = String(row.__lastName  || '').trim();
  const name  = (first + ' ' + last).trim();
  // إزالة أي شيء ليس حرفاً عربياً أو إنجليزياً أو مسافة
  const cleaned = name.replace(/[^\u0600-\u06FFa-zA-Z\s]/g, '').trim();
  return cleaned || 'غير متوفر';
}
function findCol(row, keywords) {
  const keys = Object.keys(row);
  // أولاً: بحث دقيق
  for (const key of keys) {
    const k = key.toLowerCase().trim();
    for (const kw of keywords) {
      if (k === kw.toLowerCase().trim()) return key;
    }
  }
  // ثانياً: بحث جزئي
  for (const key of keys) {
    const k = key.toLowerCase().trim();
    for (const kw of keywords) {
      if (k.includes(kw.toLowerCase().trim()) || kw.toLowerCase().trim().includes(k)) return key;
    }
  }
  return null;
}

// ---- Main Analysis ----
function analyze(gbRows, anRows) {
  if (!gbRows.length || !anRows.length) throw new Error('الملفات فارغة');

  // DEBUG: طباعة أسماء الأعمدة الفعلية في console
  console.log('=== أعمدة Gradebook ===', Object.keys(gbRows[0]));
  console.log('=== أعمدة Analytics ===', Object.keys(anRows[0]));
  console.log('=== أول صف Analytics ===', anRows[0]);
  console.log('=== أول صف Gradebook ===', gbRows[0]);

  const gbSample = gbRows[0];
  const anSample = anRows[0];
  // استخراج gbMap إذا كان موجوداً
  const gbMap = gbRows.__gbMap || {};

  // Gradebook columns - بالأسماء الحقيقية من CSV
  const gbKeys = Object.keys(gbSample);
  const cleanKey = k => k.replace(/[^\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFFa-zA-Z0-9\s\[\]:]/g, '').trim();

  // عمود الدرجة الكلية - يحتوي "Overall Grade"
  const totalColGB = gbKeys.find(k => k.includes('Overall Grade') || k.includes('إجمالي النقاط: حتى')) || null;

  // عمود اسم المستخدم للربط
  const nameColGB = gbKeys.find(k => cleanKey(k) === 'اسم المستخدم' || k.toLowerCase().includes('username')) || null;

  console.log('=== عمود الدرجة في Gradebook ===', totalColGB);
  console.log('=== عمود المستخدم في Gradebook ===', nameColGB);

  // ---- اكتشاف أعمدة الاسم باستخدام __idx_ ----
  // عمود 0 = الاسم الأخير، عمود 1 = الاسم الأول (كما في ملف Excel)
  // نتحقق من القيم الفعلية: إذا كانت نصية → عمود اسم
  const looksNumeric = v => {
    const s = String(v).replace(/\s/g,'');
    if (!s) return true;
    return /^[\d.+\-eE]+$/.test(s) || /^\d{6,}$/.test(s);
  };

  // نفحص أول صفين من البيانات للتحقق
  const sampleRows = gbRows.slice(0, Math.min(5, gbRows.length));

  // نبحث عن أعمدة __idx_ التي تحتوي قيم نصية (أسماء)
  const idxKeys = gbKeys.filter(k => k.startsWith('__idx_'));
  const nameIdxCols = idxKeys.filter(k => {
    const vals = sampleRows.map(r => String(r[k] || '').trim());
    const textCount = vals.filter(v => v && !looksNumeric(v) && v.length > 1).length;
    return textCount >= Math.min(2, sampleRows.length);
  }).sort((a, b) => {
    const ai = parseInt(a.replace('__idx_', ''));
    const bi = parseInt(b.replace('__idx_', ''));
    return ai - bi;
  });

  console.log('=== أعمدة الاسم بالـ index ===', nameIdxCols.map(k => ({
    key: k,
    sample: String(gbRows[0][k] || '')
  })));

  // عمود الاسم الأخير = أول عمود نصي، عمود الاسم الأول = ثاني عمود نصي
  const lastNameColGBFinal  = nameIdxCols[0] || null;
  const firstNameColGBFinal = nameIdxCols[1] || null;

  // عمود الدرجة الكلية - نبحث بالاسم
  // (تم تعريفه مسبقاً أعلاه)

  console.log('=== أعمدة Gradebook النهائية ===', {
    lastNameSample:  lastNameColGBFinal  ? String(gbRows[0][lastNameColGBFinal]  || '') : 'null',
    firstNameSample: firstNameColGBFinal ? String(gbRows[0][firstNameColGBFinal] || '') : 'null',
  });

  // ---- استخراج الدرجة القصوى من اسم عمود Overall Grade ----
  // اسم العمود يحتوي "حتى 40" لكن مع Unicode خاص
  // نستخرج الرقم من اسم العمود مباشرة
  let maxGrade = 100;
  if (totalColGB) {
    const colStr = String(totalColGB);
    // نستخرج كل الأرقام من اسم العمود ونأخذ أول رقم بين 20 و 200
    const nums = colStr.match(/\d+(?:\.\d+)?/g);
    if (nums) {
      const candidates = nums.map(Number).filter(n => n >= 20 && n <= 200 && n !== 1956436);
      if (candidates.length > 0) maxGrade = candidates[0];
    }
  }
  console.log('=== الدرجة القصوى من اسم العمود ===', maxGrade);

  // دالة استخراج الدرجة النهائية من خلية قد تحتوي نصاً طويلاً
  function extractFinalScore(cellValue) {
    if (cellValue === null || cellValue === undefined || cellValue === '') return NaN;
    const s = String(cellValue);
    // إذا كان رقماً مباشراً
    const direct = parseFloat(s);
    if (!isNaN(direct) && s.trim().length < 10) return direct;
    // استخراج كل الأرقام من النص وأخذ آخر رقم معقول (بين 0 و maxGrade*1.1)
    const nums = s.match(/\d+(?:\.\d+)?/g);
    if (!nums) return NaN;
    const validNums = nums.map(Number).filter(n => n >= 0 && n <= maxGrade * 1.1);
    if (validNums.length === 0) return NaN;
    // آخر رقم في النص هو الدرجة الكلية
    return validNums[validNums.length - 1];
  }

  // Analytics columns - بالأسماء الحقيقية من الملف
  const firstNameColAN = findCol(anSample, ['الاسم الأول']);
  const lastNameColAN  = findCol(anSample, ['الاسم الأخير','اسم العائلة']);
  const nameColAN      = findCol(anSample, ['اسم المستخدم','username','user']) || Object.keys(anSample)[0];
  const totalColAN     = findCol(anSample, ['التقدير الكلي','التقدير الكلي المعدل','Overall Grade','total','final','grade','الكلي','الإجمالي','overall']);
  const missedCol      = findCol(anSample, ['تنبيه بخصوص توزيع الاستحقاق الفائتة','الاستحقاق الفائتة','missed','due','فائتة','متأخر','missing','overdue']);
  const hoursCol       = findCol(anSample, ['عدد الساعات في المقرر الدراسي','عدد الساعات في المقرر','ساعات في المقرر الدراسي','عدد الساعات','hours']);
  const daysCol        = findCol(anSample, ['عدد الأيام منذ آخر وصول','أيام منذ آخر وصول','عدد الأيام','days since','days']);
  const lastAccessCol  = findCol(anSample, ['تاريخ آخر وصول','آخر وصول','last access','last login']);

  console.log('=== الأعمدة المكتشفة في Analytics ===', {
    firstNameColAN, lastNameColAN, nameColAN,
    totalColAN, missedCol, hoursCol, daysCol, lastAccessCol
  });

  // Build analytics map - الربط بالترتيب (index) لأن المعرفات قد تختلف بين الملفين
  // + احتياطياً بالمعرف إذا تطابق
  const anMap      = {};   // مفتاح: username/id
  const anByIndex  = [];   // مفتاح: رقم الصف

  for (let i = 0; i < anRows.length; i++) {
    const row = anRows[i];
    const username = String(row[nameColAN] || '').trim().toLowerCase();
    // استخراج الاسم من Analytics مباشرة (أعمدته واضحة وصحيحة)
    const anFirstName = String(row[firstNameColAN] || row['الاسم الأول'] || '').trim();
    const anLastName  = String(row[lastNameColAN]  || row['الاسم الأخير'] || '').trim();
    const entry = {
      total_an:         toNum(totalColAN    ? row[totalColAN]    : ''),
      missed_deadlines: toNumOrZero(missedCol     ? row[missedCol]     : 0),
      hours_spent:      toNumOrZero(hoursCol      ? row[hoursCol]      : 0),
      days_since_access:toNumOrZero(daysCol       ? row[daysCol]       : 0),
      last_access:      lastAccessCol ? String(row[lastAccessCol] || '').trim() : '',
      // الاسم من Analytics - واضح وصحيح
      firstName: anFirstName,
      lastName:  anLastName,
      fullName:  (anFirstName + ' ' + anLastName).trim(),
      // معرف الطالب للبريد الإلكتروني
      username:   String(row[nameColAN] || row['اسم المستخدم'] || row['معرف الطالب'] || '').trim(),
      studentId:  String(row['معرف الطالب'] || row['اسم المستخدم'] || row[nameColAN] || '').trim(),
    };
    anByIndex.push(entry);
    if (username) anMap[username] = entry;
    const userBase = username.split('@')[0];
    if (userBase) anMap[userBase] = entry;
  }

  console.log('=== anByIndex length ===', anByIndex.length);
  console.log('=== عينة اسم من Analytics ===', anByIndex[0]?.fullName);

  // Detect exam columns - أعمدة الاختبارات (تحتوي درجات 0-100)
  // نستخدم فقط أعمدة __c بالموضع لتجنب مشاكل Unicode
  const examCols = gbKeys.filter(k => {
    if (!k.startsWith('__c')) return false;
    const idx = parseInt(k.replace('__c',''));
    // تخطي أول 5 أعمدة (أسماء ومعرفات)
    if (idx < 5) return false;
    // تخطي عمود الدرجة الكلية
    if (k === totalColGB) return false;
    const vals = gbRows.map(r => parseFloat(r[k])).filter(v => !isNaN(v) && v >= 0);
    if (vals.length < 3) return false;
    const maxVal = Math.max(...vals);
    return maxVal <= 100 && maxVal > 0;
  });
  console.log('=== أعمدة الاختبارات المكتشفة ===', examCols);

  // بناء gbMap من Gradebook CSV بمعرف الطالب
  const gbMapById = {};
  for (const row of gbRows) {
    const uid = String(row['اسم المستخدم'] || row['معرف الطالب'] || '').trim();
    if (uid) gbMapById[uid] = row;
  }
  console.log('=== gbMap معرفات (أول 5) ===', Object.keys(gbMapById).slice(0,5));

  // Build student list - نبني من Analytics (17 طالب) وليس Gradebook (10 فقط)
  const students = [];

  for (let anIdx = 0; anIdx < anByIndex.length; anIdx++) {
    const an  = anByIndex[anIdx];

  // ---- اسم الطالب من Analytics ----
    const name = an.fullName || an.firstName || an.lastName || 'غير متوفر';

    // ---- الدرجة: من Gradebook CSV بالمعرف ----
    const studentId = an.studentId || an.username || '';
    const gbRow = gbMapById[studentId] || {};
    const gbRawGrade = toNum(totalColGB ? gbRow[totalColGB] : '');

    let totalGrade;
    if (!isNaN(gbRawGrade) && gbRawGrade >= 0) {
      // درجة خام من Gradebook → نسبة مئوية
      totalGrade = round(gbRawGrade / maxGrade * 100);
      console.log(`✅ ${name}: ${gbRawGrade}/${maxGrade} = ${totalGrade}%`);
    } else {
      // fallback: Analytics
      const anGrade = toNum(an.total_an);
      totalGrade = (!isNaN(anGrade) && anGrade >= 0) ? round(anGrade) : NaN;
      console.log(`⚠️ ${name}: Gradebook missing → Analytics=${totalGrade}%`);
    }

    // الصف المقابل في Gradebook للاختبارات
    const row = gbRow;

    // ---- أعمدة الاختبارات ----
    const examScores = examCols.map(c => {
      const v = toNum(row[c]);
      if (isNaN(v) || v < 0) return NaN;
      const m = c.match(/من\s*(\d+(?:\.\d+)?)/);
      const colMax = m ? parseFloat(m[1]) : 100;
      return round(v / colMax * 100);
    }).filter(v => !isNaN(v) && v >= 0 && v <= 100);

    const examAvg = examScores.length ? round(avg(examScores)) : (isNaN(totalGrade) ? NaN : totalGrade);
    const examStd = examScores.length > 1 ? std(examScores) : 0;
    const examMin = examScores.length ? Math.min(...examScores) : (isNaN(totalGrade) ? NaN : totalGrade);
    const examMax = examScores.length ? Math.max(...examScores) : (isNaN(totalGrade) ? NaN : totalGrade);
    const below50 = examScores.filter(v => v < 50).length;

    // ---- بيانات Analytics بالترتيب (index) ----
    const missed     = an.missed_deadlines  || 0;
    const hours      = an.hours_spent       || 0;
    const days       = an.days_since_access || 0;
    const lastAccess = an.last_access       || '';

    const riskScore   = calcRisk(totalGrade, missed, hours, days, below50);
    const riskLevel   = getRiskLevel(totalGrade);
    const engagement  = getEngagement(hours, days);
    const trend       = getTrend(examStd, examAvg, examMin, examMax);
    const recs        = getRecommendations(totalGrade, missed, hours, days, below50);

    students.push({
      name, totalGrade, examAvg, examStd, examMin, examMax,
      missed, hours, days, lastAccess, below50,
      riskScore, riskLevel, engagement, trend, recs,
      atRisk: riskScore >= 30 || (!isNaN(totalGrade) && totalGrade < 60),
      // معرف الطالب من Analytics (واضح وصحيح)
      username: String(an.studentId || an.username || '').trim(),
    });
  }

  students.sort((a, b) => b.riskScore - a.riskScore);

  // Summary - تجاهل الطلاب الذين ليس لديهم درجات
  const total   = students.length;
  const atRisk  = students.filter(s => s.atRisk).length;
  const grades  = students.map(s => s.totalGrade).filter(v => !isNaN(v) && v > 0);
  const avgGrade= grades.length ? avg(grades) : 0;
  const passRate= grades.length ? (grades.filter(v => v >= 60).length / grades.length * 100) : 0;

  console.log('=== عينة درجات ===', grades.slice(0,5));
  console.log('=== نسبة النجاح ===', passRate, 'من', grades.length, 'طالب');

  // Distributions
  const riskDist = countBy(students, s => s.riskLevel);
  const engDist  = countBy(students, s => s.engagement);
  const trendDist= countBy(students, s => s.trend);
  const gradeDist = {
    'ممتاز (90-100)':   grades.filter(v => v >= 90).length,
    'جيد جداً (80-89)': grades.filter(v => v >= 80 && v < 90).length,
    'جيد (70-79)':      grades.filter(v => v >= 70 && v < 80).length,
    'مقبول (60-69)':    grades.filter(v => v >= 60 && v < 70).length,
    'راسب (<60)':       grades.filter(v => v < 60).length,
  };

  return {
    summary: {
      total, atRisk, safe: total - atRisk,
      avgGrade: round(avgGrade), passRate: round(passRate),
      avgHours: round(avg(students.map(s => s.hours))),
      avgMissed: round(avg(students.map(s => s.missed))),
      avgDays: round(avg(students.map(s => s.days))),
    },
    riskDist, engDist, trendDist, gradeDist, students,
  };
}

// ---- Helpers ----
// تحويل القيمة لرقم - يدعم النسب المئوية مثل "39.2%"
function toNum(v) {
  if (v === null || v === undefined || v === '') return NaN;
  const s = String(v).trim();
  if (s === '' || s === '-') return NaN;
  // إزالة كل شيء عدا الأرقام والنقطة والسالب
  const n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? NaN : n;
}

function toNumOrZero(v) {
  const n = toNum(v);
  return isNaN(n) ? 0 : n;
}
function avg(arr) { return arr.length ? arr.reduce((a,b) => a+b, 0) / arr.length : 0; }
function std(arr) {
  const m = avg(arr);
  return Math.sqrt(arr.reduce((s,v) => s + (v-m)**2, 0) / arr.length);
}
function round(v, d=1) { return Math.round(v * 10**d) / 10**d; }
function countBy(arr, fn) {
  return arr.reduce((acc, v) => { const k = fn(v); acc[k] = (acc[k]||0)+1; return acc; }, {});
}

function calcRisk(grade, missed, hours, days, below50) {
  let s = 0;
  const g = isNaN(grade) ? 0 : grade;

  // الدرجة الكلية - المؤشر الأساسي
  if (g < 50)      s += 40;
  else if (g < 60) s += 30;
  else if (g < 70) s += 20;
  else if (g < 80) s += 10;
  else if (g < 90) s += 5;

  // المهام والاختبارات الفائتة - مؤشر قوي جداً
  if (missed >= 10)     s += 35;
  else if (missed >= 7) s += 25;
  else if (missed >= 5) s += 20;
  else if (missed >= 3) s += 12;
  else if (missed >= 1) s += 5;

  // أيام الغياب
  if (days > 21)      s += 20;
  else if (days > 14) s += 15;
  else if (days > 7)  s += 10;

  // ساعات الدراسة
  if (hours < 1)      s += 15;
  else if (hours < 3) s += 7;
  else if (hours < 5) s += 3;

  // اختبارات أقل من 50%
  if (below50 > 3)      s += 10;
  else if (below50 > 1) s += 5;

  return Math.min(s, 100);
}

function getRiskLevel(grade) {
  if (isNaN(grade) || grade <= 0) return 'غير محدد';
  if (grade >= 80) return 'منخفض';
  if (grade >= 60) return 'متوسط';
  if (grade >= 50) return 'مرتفع';
  return 'حرج';
}

function getRiskColor(level) {
  return { 'منخفض':'#27ae60','متوسط':'#f39c12','مرتفع':'#e67e22','حرج':'#e74c3c','غير محدد':'#95a5a6' }[level] || '#95a5a6';
}

function getEngagement(hours, days) {
  if (hours > 10 && days < 3)  return 'ممتاز';
  if (hours > 5  && days < 7)  return 'جيد';
  if (hours > 2  && days < 14) return 'متوسط';
  return 'ضعيف';
}

function getTrend(sd, av, mn, mx) {
  if (sd === 0) return 'مستقر';
  const cv = av > 0 ? sd / av : 0;
  if (cv < 0.15) return 'مستقر';
  if (mx - mn > 30) return 'متذبذب';
  if (av >= 70) return 'تحسن';
  return 'تراجع';
}

function getRecommendations(grade, missed, hours, days, below50) {
  const r = [];
  if (isNaN(grade) || grade < 50) r.push('⚠️ تدخل عاجل: جلسة دعم فردية');
  if (missed > 3)  r.push('📅 متابعة الواجبات الفائتة');
  if (days > 14)   r.push('📧 إرسال تنبيه للطالب');
  if (hours < 2)   r.push('⏱️ زيادة وقت الدراسة');
  if (below50 > 2) r.push('📚 مراجعة المفاهيم الأساسية');
  if (!r.length)   r.push('✅ الطالب يسير بشكل جيد');
  return r.join(' | ');
}

// ---- Render ----
function renderReport(data) {
  renderKPIs(data.summary);
  renderCharts(data);
  renderTable(data.students);
}

function renderKPIs(s) {
  const cards = [
    { icon:'👥', value: s.total,    label:'إجمالي الطلاب',        cls:'' },
    { icon:'⚠️', value: s.atRisk,   label:'طلاب في خطر',          cls:'danger' },
    { icon:'✅', value: s.safe,     label:'طلاب بأمان',            cls:'success' },
    { icon:'📊', value: s.avgGrade+'%', label:'متوسط الدرجات',    cls:'' },
    { icon:'🎯', value: s.passRate+'%', label:'نسبة النجاح',       cls: s.passRate>=70?'success':'warning' },
    { icon:'⏱️', value: s.avgHours+'h', label:'متوسط ساعات المقرر', cls:'info' },
    { icon:'📅', value: s.avgMissed,    label:'متوسط المهام الفائتة', cls: s.avgMissed>3?'danger':'warning' },
    { icon:'📆', value: s.avgDays+'d',  label:'متوسط أيام منذ آخر وصول',   cls: s.avgDays>14?'danger':'' },
  ];
  kpiGrid.innerHTML = cards.map(c => `
    <div class="kpi-card ${c.cls}">
      <div class="kpi-icon">${c.icon}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-label">${c.label}</div>
    </div>`).join('');
}

function destroyCharts() {
  Object.values(charts).forEach(c => c && c.destroy());
  charts = {};
}

function renderCharts(data) {
  destroyCharts();

  const riskLabels = Object.keys(data.riskDist);
  const riskColors = riskLabels.map(l =>
    l==='حرج'?'#c62828':l==='مرتفع'?'#e65100':l==='متوسط'?'#f57f17':l==='منخفض'?'#2e7d32':'#9e9e9e');
  charts.risk = new Chart(document.getElementById('riskChart'), {
    type:'doughnut',
    data:{ labels:riskLabels, datasets:[{ data:Object.values(data.riskDist), backgroundColor:riskColors, borderWidth:2 }] },
    options:{ plugins:{ legend:{ position:'bottom' } }, cutout:'60%' },
  });

  const gradeLabels = Object.keys(data.gradeDist);
  charts.grade = new Chart(document.getElementById('gradeChart'), {
    type:'bar',
    data:{ labels:gradeLabels, datasets:[{ label:'عدد الطلاب', data:Object.values(data.gradeDist),
      backgroundColor:['#1b5e20','#2e7d32','#f57f17','#e65100','#c62828'], borderRadius:6 }] },
    options:{ plugins:{ legend:{ display:false } }, scales:{ y:{ beginAtZero:true } } },
  });

  const engLabels = Object.keys(data.engDist);
  const engColors = engLabels.map(l =>
    l==='ممتاز'?'#1b5e20':l==='جيد'?'#2e7d32':l==='متوسط'?'#f57f17':'#c62828');
  charts.eng = new Chart(document.getElementById('engagementChart'), {
    type:'pie',
    data:{ labels:engLabels, datasets:[{ data:Object.values(data.engDist), backgroundColor:engColors, borderWidth:2 }] },
    options:{ plugins:{ legend:{ position:'bottom' } } },
  });

  const trendLabels = Object.keys(data.trendDist);
  const trendColors = trendLabels.map(l =>
    l==='تحسن'?'#1b5e20':l==='مستقر'?'#0288d1':l==='تراجع'?'#c62828':'#f57f17');
  charts.trend = new Chart(document.getElementById('trendChart'), {
    type:'bar',
    data:{ labels:trendLabels, datasets:[{ label:'عدد الطلاب', data:Object.values(data.trendDist),
      backgroundColor:trendColors, borderRadius:6 }] },
    options:{ plugins:{ legend:{ display:false } }, scales:{ y:{ beginAtZero:true } } },
  });
}

function renderTable(students) {
  studentsBody.innerHTML = students.map((s, i) => {
    const grade = isNaN(s.totalGrade) ? 'غير متاح' : s.totalGrade.toFixed(1) + '%';
    const examA = isNaN(s.examAvg)    ? '-'         : s.examAvg.toFixed(1) + '%';
    const displayName = s.name || 'غير متوفر';
    const color = getRiskColor(s.riskLevel);
    const trendIcon = s.trend==='تحسن'?'📈':s.trend==='تراجع'?'📉':s.trend==='مستقر'?'➡️':'〰️';
    const trendCls  = s.trend==='تحسن'?'trend-up':s.trend==='تراجع'?'trend-down':
                      s.trend==='مستقر'?'trend-stable':'trend-wave';
    return `
      <tr data-risk="${s.riskLevel}" data-name="${s.name.toLowerCase()}">
        <td>${i+1}</td>
        <td><strong>${displayName}</strong></td>
        <td>${grade}</td>
        <td>${examA}</td>
        <td>${s.hours > 0 ? s.hours.toFixed(2) + ' h' : '-'}</td>
        <td>${s.lastAccess || '-'}</td>
        <td>${s.days > 0 ? s.days + ' يوم' : '-'}</td>
        <td><span class="risk-badge risk-${s.riskLevel}">${s.riskLevel}</span></td>
        <td>
          <div class="risk-bar-wrap">
            <span>${s.riskScore}</span>
            <div class="risk-bar">
              <div class="risk-bar-fill" style="width:${s.riskScore}%;background:${color}"></div>
            </div>
          </div>
        </td>
        <td>${s.engagement}</td>
        <td class="${trendCls}">${trendIcon} ${s.trend}</td>
        <td class="rec-cell">${s.recs}</td>
        <td>
          <button class="btn-email" onclick="openEmailModal(${i})">
            <i class="fas fa-envelope"></i> إرسال بريد
          </button>
        </td>
      </tr>`;
  }).join('');
}

// ---- Filters ----
document.querySelectorAll('.filter-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    applyFilters();
  });
});
searchInput.addEventListener('input', applyFilters);

function applyFilters() {
  const filter = document.querySelector('.filter-btn.active')?.dataset.filter || 'all';
  const search = searchInput.value.trim().toLowerCase();
  let list = allStudents;
  if (filter !== 'all') list = list.filter(s => s.riskLevel === filter);
  if (search) list = list.filter(s => s.name.toLowerCase().includes(search));
  renderTable(list);
}

// ---- New Analysis ----
document.getElementById('newAnalysisBtn').addEventListener('click', () => {
  reportSection.style.display = 'none';
  uploadSection.style.display = 'block';
  gradebookInput.value = '';
  analyticsInput.value = '';
  document.getElementById('gradebookName').textContent = 'لم يتم اختيار ملف';
  document.getElementById('analyticsName').textContent = 'لم يتم اختيار ملف';
  document.getElementById('gradebookCheck').textContent = '';
  document.getElementById('analyticsCheck').textContent = '';
  document.getElementById('gradebookBox').classList.remove('has-file');
  document.getElementById('analyticsBox').classList.remove('has-file');
  analyzeBtn.disabled = true;
  destroyCharts();
  window.scrollTo({ top:0, behavior:'smooth' });
});

// ---- Download Excel ----
document.getElementById('downloadBtn').addEventListener('click', () => {
  if (!reportData) return;
  const wb = XLSX.utils.book_new();

  // Sheet 1: Summary
  const s = reportData.summary;
  const summaryData = [
    ['نظام التنبؤ المبكر بالتعثر الأكاديمي - تقرير شامل'],
    [],
    ['المؤشر', 'القيمة'],
    ['إجمالي الطلاب', s.total],
    ['الطلاب في خطر', s.atRisk],
    ['الطلاب بأمان', s.safe],
    ['متوسط الدرجات', s.avgGrade + '%'],
    ['نسبة النجاح', s.passRate + '%'],
    ['متوسط ساعات الدراسة', s.avgHours],
    ['متوسط المهام الفائتة', s.avgMissed],
    ['متوسط أيام الغياب', s.avgDays],
    [],
    ['توزيع الدرجات', ''],
    ...Object.entries(reportData.gradeDist),
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
  ws1['!cols'] = [{ wch: 30 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, ws1, 'ملخص تنفيذي');

  // Sheet 2: All Students
  const headers = ['اسم الطالب','الدرجة الكلية','متوسط الاختبارات','أدنى درجة','أعلى درجة',
    'ساعات في المقرر','تاريخ آخر وصول','أيام منذ آخر وصول','مستوى الخطر','درجة الخطر',
    'مستوى التفاعل','اتجاه الأداء','التوصيات'];
  const rows = reportData.students.map(s => [
    s.name, s.totalGrade, s.examAvg, s.examMin, s.examMax,
    s.hours, s.lastAccess, s.days, s.riskLevel, s.riskScore,
    s.engagement, s.trend, s.recs,
  ]);
  const ws2 = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws2['!cols'] = headers.map((h,i) => ({ wch: i===12?50:i===0?25:15 }));
  XLSX.utils.book_append_sheet(wb, ws2, 'تفاصيل الطلاب');

  // Sheet 3: At-Risk
  const atRiskRows = reportData.students.filter(s => s.atRisk).map(s => [
    s.name, s.totalGrade, s.examAvg, s.examMin, s.examMax,
    s.hours, s.lastAccess, s.days, s.riskLevel, s.riskScore,
    s.engagement, s.trend, s.recs,
  ]);
  const ws3 = XLSX.utils.aoa_to_sheet([headers, ...atRiskRows]);
  ws3['!cols'] = headers.map((h,i) => ({ wch: i===12?50:i===0?25:15 }));
  XLSX.utils.book_append_sheet(wb, ws3, 'الطلاب في خطر');

  XLSX.writeFile(wb, 'academic_analytics_report.xlsx');
});

// ---- Email Modal ----
const emailModal   = document.getElementById('emailModal');
const emailTo      = document.getElementById('emailTo');
const emailSubject = document.getElementById('emailSubject');
const emailBody    = document.getElementById('emailBody');
const modalSend    = document.getElementById('modalSend');

// فتح الـ Modal لطالب معين
function openEmailModal(studentIndex) {
  // البحث عن الطالب في القائمة الحالية المعروضة
  const rows = document.querySelectorAll('#studentsBody tr');
  const row  = rows[studentIndex];
  if (!row) return;

  const name       = row.querySelector('td:nth-child(2)')?.textContent?.trim() || '';
  const grade      = row.querySelector('td:nth-child(3)')?.textContent?.trim() || '-';
  const examAvg    = row.querySelector('td:nth-child(4)')?.textContent?.trim() || '-';
  const hours      = row.querySelector('td:nth-child(5)')?.textContent?.trim() || '-';
  const lastAccess = row.querySelector('td:nth-child(6)')?.textContent?.trim() || '-';
  const days       = row.querySelector('td:nth-child(7)')?.textContent?.trim() || '-';
  const riskLevel  = row.querySelector('td:nth-child(8) .risk-badge')?.textContent?.trim() || '-';
  const engagement = row.querySelector('td:nth-child(10)')?.textContent?.trim() || '-';

  // توليد البريد الإلكتروني من معرف الطالب (اسم المستخدم)
  const student = allStudents.find(s => s.name === name) || {};
  const studentId = student.username || '';
  const toEmail = studentId ? `${studentId}@qu.edu.sa` : `${name.replace(/\s+/g,'.')}@qu.edu.sa`;

  // تحديد الحالة
  const statusMap = { 'منخفض': 'جيد ✅', 'متوسط': 'متوسط ⚠️', 'مرتفع': 'يحتاج متابعة 🔶', 'حرج': 'في خطر 🔴', 'غير محدد': 'غير محدد' };
  const status = statusMap[riskLevel] || riskLevel;

  // صياغة الرسالة
  const message = `عزيزي الطالب ${name}،

السلام عليكم ورحمة الله وبركاته،

نود إبلاغك بتقريرك الأكاديمي الخاص بالمقرر الدراسي:

📊 الدرجة الكلية: ${grade}
📝 متوسط الاختبارات: ${examAvg}
⏱️ ساعات المقرر: ${hours}
📅 تاريخ آخر وصول: ${lastAccess}
📆 أيام منذ آخر تفاعل: ${days}
🎯 مستوى التفاعل: ${engagement}
🔔 الحالة الأكاديمية: ${status}

${riskLevel === 'حرج' || riskLevel === 'مرتفع'
  ? '⚠️ نرجو منك الاهتمام بالمقرر وزيادة التفاعل والمشاركة، والتواصل مع أستاذ المقرر في أقرب وقت.'
  : riskLevel === 'متوسط'
  ? '📌 أداؤك في المستوى المتوسط، ننصحك بزيادة وقت الدراسة والمراجعة المنتظمة.'
  : '✅ أداؤك جيد، استمر في التميز والمثابرة.'}

يرجى تحسين أدائك في حال الحاجة، ونحن هنا لدعمك.

مع تحيات أستاذ المقرر 🎓`;

  // تعبئة الـ Modal
  emailTo.textContent      = toEmail;
  emailSubject.textContent = `تقرير أكاديمي - ${name}`;
  emailBody.textContent    = message;

  // إعادة زر الإرسال لحالته الأصلية
  modalSend.classList.remove('sent');
  modalSend.innerHTML = '<i class="fas fa-paper-plane"></i> إرسال';
  modalSend.disabled  = false;

  emailModal.style.display = 'flex';
}

// إغلاق الـ Modal
document.getElementById('modalClose').addEventListener('click',  () => emailModal.style.display = 'none');
document.getElementById('modalCancel').addEventListener('click', () => emailModal.style.display = 'none');
emailModal.addEventListener('click', e => { if (e.target === emailModal) emailModal.style.display = 'none'; });

// زر الإرسال
modalSend.addEventListener('click', () => {
  modalSend.classList.add('sent');
  modalSend.innerHTML = '<i class="fas fa-check"></i> تم الإرسال ✓';
  modalSend.disabled  = true;
  setTimeout(() => { emailModal.style.display = 'none'; }, 1500);
});
