/* =====================================================
   Academic Early Risk Prediction Dashboard
   Frontend = UI only
   Backend Flask = analysis, report generation, email
   ===================================================== */

let allStudents = [];
let displayedStudents = [];
let charts = {};
let reportData = null;
let selectedStudent = null;

// ---- DOM ----
const gradebookInput = document.getElementById('gradebookInput');
const analyticsInput = document.getElementById('analyticsInput');
const analyzeBtn = document.getElementById('analyzeBtn');
const loadingOverlay = document.getElementById('loadingOverlay');
const uploadSection = document.getElementById('uploadSection');
const reportSection = document.getElementById('reportSection');
const studentsBody = document.getElementById('studentsBody');
const kpiGrid = document.getElementById('kpiGrid');
const searchInput = document.getElementById('searchInput');

const newAnalysisBtn = document.getElementById('newAnalysisBtn');
const downloadBtn = document.getElementById('downloadBtn');

const emailModal = document.getElementById('emailModal');
const emailFrom = document.getElementById('emailFrom');
const emailTo = document.getElementById('emailTo');
const emailSubject = document.getElementById('emailSubject');
const emailBody = document.getElementById('emailBody');
const modalClose = document.getElementById('modalClose');
const modalCancel = document.getElementById('modalCancel');
const modalSend = document.getElementById('modalSend');

// ---- Init ----
document.addEventListener('DOMContentLoaded', () => {
  setupEvents();
  checkReady();
});

// ---- Events ----
function setupEvents() {
  gradebookInput.addEventListener('change', () => {
    handleFileSelection(
      gradebookInput,
      'gradebookName',
      'gradebookCheck',
      'gradebookBox'
    );
  });

  analyticsInput.addEventListener('change', () => {
    handleFileSelection(
      analyticsInput,
      'analyticsName',
      'analyticsCheck',
      'analyticsBox'
    );
  });

  analyzeBtn.addEventListener('click', analyzeFiles);

  newAnalysisBtn.addEventListener('click', resetAnalysis);

  downloadBtn.addEventListener('click', downloadBackendReport);

  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document
        .querySelectorAll('.filter-btn')
        .forEach(b => b.classList.remove('active'));

      btn.classList.add('active');
      applyFilters();
    });
  });

  searchInput.addEventListener('input', applyFilters);

  modalClose.addEventListener('click', closeEmailModal);
  modalCancel.addEventListener('click', closeEmailModal);
  modalSend.addEventListener('click', sendEmailToStudent);
}

// ---- File UI ----
function handleFileSelection(input, nameId, checkId, boxId) {
  if (input.files && input.files[0]) {
    document.getElementById(nameId).textContent = input.files[0].name;
    document.getElementById(checkId).textContent = '✅';
    document.getElementById(boxId).classList.add('has-file');
  }

  checkReady();
}

function checkReady() {
  analyzeBtn.disabled = !(
    gradebookInput.files[0] &&
    analyticsInput.files[0]
  );
}

// ---- Analyze using Flask backend ----
async function analyzeFiles() {
  if (!gradebookInput.files[0] || !analyticsInput.files[0]) {
    alert('يرجى اختيار ملفي Gradebook و Analytics أولاً');
    return;
  }

  const formData = new FormData();
  formData.append('gradebook', gradebookInput.files[0]);
  formData.append('analytics', analyticsInput.files[0]);

  try {
    showLoading(true);

    const response = await fetch('/api/analyze', {
      method: 'POST',
      body: formData
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'حدث خطأ أثناء التحليل');
    }

    // Compatible with both:
    // return jsonify(report)
    // return jsonify({ success: true, report })
    reportData = data.report || data;

    if (!reportData.students || !Array.isArray(reportData.students)) {
      throw new Error('استجابة الخادم لا تحتوي على بيانات الطلاب');
    }

    allStudents = normalizeStudents(reportData.students);
    reportData.students = allStudents;

    renderReport(reportData);

    uploadSection.style.display = 'none';
    reportSection.style.display = 'block';

    window.scrollTo({ top: 0, behavior: 'smooth' });
  } catch (error) {
    console.error(error);
    alert('❌ فشل التحليل:\n\n' + error.message);
  } finally {
    showLoading(false);
  }
}

function showLoading(show) {
  loadingOverlay.style.display = show ? 'flex' : 'none';
}

// ---- Normalize backend student fields ----
function normalizeStudents(students) {
  return students.map(st => ({
    name: st.name || st.student_name || 'غير متوفر',
    student_id: st.student_id || st.username || '',
    total_grade: toNumber(st.total_grade),
    exam_avg: toNumber(st.exam_avg),
    exam_std: toNumber(st.exam_std),
    exam_min: toNumber(st.exam_min),
    exam_max: toNumber(st.exam_max),
    missed_deadlines: toNumber(st.missed_deadlines),
    hours_spent: toNumber(st.hours_spent),
    days_since_access: toNumber(st.days_since_access),
    exams_below_50: toNumber(st.exams_below_50),
    risk_score: toNumber(st.risk_score),
    risk_level: st.risk_level || 'غير محدد',
    risk_color: st.risk_color || getRiskColor(st.risk_level),
    at_risk: Boolean(st.at_risk),
    engagement: st.engagement || 'غير محدد',
    trend: st.trend || 'غير محدد',
    recommendations: st.recommendations || '',
    last_access: st.last_access || '-'
  }));
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

// ---- Render ----
function renderReport(data) {
  renderKPIs(data.summary);
  renderCharts(data);
  renderTable(allStudents);
}

function renderKPIs(summary) {
  const s = summary || {};

  const cards = [
    {
      icon: '👥',
      value: s.total ?? 0,
      label: 'إجمالي الطلاب',
      cls: ''
    },
    {
      icon: '⚠️',
      value: s.atRisk ?? 0,
      label: 'طلاب في خطر',
      cls: 'danger'
    },
    {
      icon: '✅',
      value: s.safe ?? 0,
      label: 'طلاب بأمان',
      cls: 'success'
    },
    {
      icon: '📊',
      value: `${s.avgGrade ?? 0}%`,
      label: 'متوسط الدرجات',
      cls: ''
    },
    {
      icon: '🎯',
      value: `${s.passRate ?? 0}%`,
      label: 'نسبة النجاح',
      cls: (s.passRate ?? 0) >= 70 ? 'success' : 'warning'
    },
    {
      icon: '⏱️',
      value: `${s.avgHours ?? 0}h`,
      label: 'متوسط ساعات المقرر',
      cls: 'info'
    },
    {
      icon: '📅',
      value: s.avgMissed ?? 0,
      label: 'متوسط المهام الفائتة',
      cls: (s.avgMissed ?? 0) > 3 ? 'danger' : 'warning'
    },
    {
      icon: '📆',
      value: `${s.avgDays ?? 0}d`,
      label: 'متوسط أيام منذ آخر وصول',
      cls: (s.avgDays ?? 0) > 14 ? 'danger' : ''
    }
  ];

  kpiGrid.innerHTML = cards.map(card => `
    <div class="kpi-card ${card.cls}">
      <div class="kpi-icon">${card.icon}</div>
      <div class="kpi-value">${card.value}</div>
      <div class="kpi-label">${card.label}</div>
    </div>
  `).join('');
}

function destroyCharts() {
  Object.values(charts).forEach(chart => {
    if (chart) chart.destroy();
  });

  charts = {};
}

function renderCharts(data) {
  destroyCharts();

  const riskDist = data.riskDist || {};
  const gradeDist = data.gradeDist || {};
  const engDist = data.engDist || {};
  const trendDist = data.trendDist || {};

  charts.risk = new Chart(document.getElementById('riskChart'), {
    type: 'doughnut',
    data: {
      labels: Object.keys(riskDist),
      datasets: [{
        data: Object.values(riskDist),
        backgroundColor: Object.keys(riskDist).map(getRiskColor),
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom' }
      },
      cutout: '60%'
    }
  });

  charts.grade = new Chart(document.getElementById('gradeChart'), {
    type: 'bar',
    data: {
      labels: Object.keys(gradeDist),
      datasets: [{
        label: 'عدد الطلاب',
        data: Object.values(gradeDist),
        backgroundColor: [
          '#1b5e20',
          '#2e7d32',
          '#f57f17',
          '#e65100',
          '#c62828'
        ],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: { beginAtZero: true }
      }
    }
  });

  charts.engagement = new Chart(document.getElementById('engagementChart'), {
    type: 'pie',
    data: {
      labels: Object.keys(engDist),
      datasets: [{
        data: Object.values(engDist),
        backgroundColor: [
          '#1b5e20',
          '#2e7d32',
          '#f57f17',
          '#c62828'
        ],
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom' }
      }
    }
  });

  charts.trend = new Chart(document.getElementById('trendChart'), {
    type: 'bar',
    data: {
      labels: Object.keys(trendDist),
      datasets: [{
        label: 'عدد الطلاب',
        data: Object.values(trendDist),
        backgroundColor: [
          '#1565c0',
          '#2e7d32',
          '#f57f17',
          '#c62828'
        ],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      indexAxis: 'y',
      plugins: {
        legend: { display: false }
      },
      scales: {
        x: { beginAtZero: true }
      }
    }
  });
}

function renderTable(students) {
  displayedStudents = students;

  if (!students.length) {
    studentsBody.innerHTML = `
      <tr>
        <td colspan="13">لا توجد بيانات مطابقة</td>
      </tr>
    `;
    return;
  }

  studentsBody.innerHTML = students.map((student, index) => `
    <tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(student.name)}</td>
      <td>${formatPercent(student.total_grade)}</td>
      <td>${formatPercent(student.exam_avg)}</td>
      <td>${formatNumber(student.hours_spent)}</td>
      <td>${escapeHtml(student.last_access || '-')}</td>
      <td>${formatNumber(student.days_since_access)}</td>
      <td>
        <span class="risk-badge" style="background:${student.risk_color}">
          ${escapeHtml(student.risk_level)}
        </span>
      </td>
      <td>${formatNumber(student.risk_score)}</td>
      <td>${escapeHtml(student.engagement)}</td>
      <td>${escapeHtml(student.trend)}</td>
      <td class="recommendations-cell">
        ${escapeHtml(student.recommendations)}
      </td>
      <td>
        <button class="btn-email" onclick="openEmailModal(${index})">
          <i class="fas fa-envelope"></i> إرسال بريد
        </button>
      </td>
    </tr>
  `).join('');
}

// ---- Filters ----
function applyFilters() {
  const activeFilter =
    document.querySelector('.filter-btn.active')?.dataset.filter || 'all';

  const searchValue = searchInput.value.trim().toLowerCase();

  let filtered = [...allStudents];

  if (activeFilter !== 'all') {
    filtered = filtered.filter(
      student => student.risk_level === activeFilter
    );
  }

  if (searchValue) {
    filtered = filtered.filter(student =>
      student.name.toLowerCase().includes(searchValue) ||
      String(student.student_id).toLowerCase().includes(searchValue)
    );
  }

  renderTable(filtered);
}

// ---- Reset ----
function resetAnalysis() {
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

  allStudents = [];
  displayedStudents = [];
  reportData = null;
  selectedStudent = null;

  destroyCharts();

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ---- Download report from backend ----
async function downloadBackendReport() {
  try {
    if (!reportData) {
      alert('لا يوجد تقرير للتنزيل');
      return;
    }

    downloadBtn.disabled = true;
    downloadBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> جارٍ التنزيل...';

    const response = await fetch('/api/download-report', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ download: true })
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error || 'فشل تنزيل التقرير');
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'academic_analytics_report.xlsx';
    document.body.appendChild(a);
    a.click();

    a.remove();
    window.URL.revokeObjectURL(url);
  } catch (error) {
    console.error(error);
    alert('❌ ' + error.message);
  } finally {
    downloadBtn.disabled = false;
    downloadBtn.innerHTML = '<i class="fas fa-file-excel"></i> تنزيل تقرير Excel';
  }
}

// ---- Email Modal ----
function openEmailModal(index) {
  selectedStudent = displayedStudents[index];

  if (!selectedStudent) return;

  const toEmail = selectedStudent.student_id
    ? `${selectedStudent.student_id}@qu.edu.sa`
    : 'غير متوفر';

  emailFrom.textContent = 'من إعدادات الخادم backend/.env';
  emailTo.textContent = toEmail;
  emailSubject.textContent = 'تنبيه أكاديمي - حالة أدائك في المقرر';

  emailBody.innerHTML = `
    <p>عزيزي الطالب/ـة <strong>${escapeHtml(selectedStudent.name)}</strong>,</p>
    <p>نحن نتابع أداءك في المقرر من خلال نظام التنبؤ المبكر بالتعثر الأكاديمي.</p>
    <p><strong>مستوى الخطر الحالي:</strong> ${escapeHtml(selectedStudent.risk_level)}</p>
    <p><strong>درجة الخطر:</strong> ${formatNumber(selectedStudent.risk_score)}</p>
    <p><strong>التوصيات:</strong></p>
    <p>${escapeHtml(selectedStudent.recommendations)}</p>
    <p>يرجى التواصل مع المدرس أو المشرف الأكاديمي للحصول على الدعم اللازم.</p>
  `;

  emailModal.style.display = 'flex';
}

function closeEmailModal() {
  emailModal.style.display = 'none';
  selectedStudent = null;
}

async function sendEmailToStudent() {
  if (!selectedStudent) return;

  try {
    modalSend.disabled = true;
    modalSend.innerHTML = '<i class="fas fa-spinner fa-spin"></i> جارٍ الإرسال...';

    const response = await fetch('/api/send-email', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        student_id: selectedStudent.student_id,
        student_name: selectedStudent.name,
        risk_level: selectedStudent.risk_level,
        recommendations: selectedStudent.recommendations
      })
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'فشل إرسال البريد');
    }

    alert('✅ ' + (data.message || 'تم إرسال البريد بنجاح'));
    closeEmailModal();
  } catch (error) {
    console.error(error);
    alert('❌ ' + error.message);
  } finally {
    modalSend.disabled = false;
    modalSend.innerHTML = '<i class="fas fa-paper-plane"></i> إرسال';
  }
}

// ---- Helpers ----
function getRiskColor(level) {
  const colors = {
    'منخفض': '#27ae60',
    'متوسط': '#f39c12',
    'مرتفع': '#e67e22',
    'حرج': '#e74c3c',
    'غير محدد': '#95a5a6'
  };

  return colors[level] || '#95a5a6';
}

function formatPercent(value) {
  return `${formatNumber(value)}%`;
}

function formatNumber(value) {
  const n = Number(value);

  if (!Number.isFinite(n)) return '0';

  return Math.round(n * 10) / 10;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}