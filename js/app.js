/**
 * مؤشرات الأداء - كلية الشريعة والأنظمة
 * يقرأ من ملف data/data.csv
 */

let programsData = [];
let currentData = null;
let compareData = { item1: null, item2: null, title1: '', title2: '' };

// ========================================
// المؤشرات الـ 13
// ========================================
const INDICATORS = [
    { id: 1, name: "تقويم الطالب لجودة خبرات التعلم", unit: "درجة", key: "experience_eval", numeric: true },
    { id: 2, name: "تقييم الطالب لجودة المقررات", unit: "درجة", key: "course_eval", numeric: true },
    { id: 3, name: "معدّل التخرج بالوقت المحدد", unit: "%", key: "graduation_rate", numeric: true },
    { id: 4, name: "معدّل استبقاء طلاب السنة الأولى", unit: "%", key: "retention_rate", numeric: true },
    { id: 5, name: "مستوى أداء الطالب", unit: "%", key: "student_performance", numeric: true },
    { id: 6, name: "توظيف الخريجين", unit: "%", key: "employment_rate", numeric: true },
    { id: 7, name: "تقويم جهات التوظيف", unit: "درجة", key: "employer_eval", numeric: true },
    { id: 8, name: "نسبة الطلاب/هيئة التدريس", unit: "نسبة", key: "student_faculty_ratio" },
    { id: 9, name: "نسبة النشر العلمي", unit: "%", key: "publication_pct", numeric: true },
    { id: 10, name: "البحوث/عضو هيئة تدريس", unit: "بحث", key: "research_per_faculty", numeric: true },
    { id: 11, name: "الاقتباسات/عضو هيئة تدريس", unit: "اقتباس", key: "citations_per_faculty", numeric: true },
    { id: 12, name: "نسبة نشر طلاب الدراسات العليا", unit: "%", key: "student_publication", numeric: true, gradOnly: true },
    { id: 13, name: "براءات الاختراع", unit: "براءة", key: "patents", numeric: true, gradOnly: true },
];

const RAW_DATA_LABELS = {
    students: "عدد الطلاب المنتظمين",
    graduates: "عدد الخريجين",
    faculty_total: "إجمالي أعضاء هيئة التدريس",
    faculty_published: "الأعضاء الذين نشروا بحثاً",
    research_count: "عدد الأبحاث المنشورة",
    citations: "إجمالي الاقتباسات",
    course_eval: "تقييم جودة المقررات",
    experience_eval: "تقييم خبرة البرنامج",
};

// ========================================
// تحويل السنة: 46 → 1446
// ========================================
function formatYear(y) {
    if (!y) return '';
    let n = Math.floor(parseFloat(y));
    return n < 100 ? '14' + String(n).padStart(2, '0') : String(n);
}

function shortYear(y) {
    if (!y) return '';
    let n = Math.floor(parseFloat(y));
    return n >= 1400 ? String(n).slice(-2) : String(n);
}

// ========================================
// تحميل البيانات من CSV
// ========================================
async function loadData() {
    const statusDot = document.getElementById('status-dot');
    const statusText = document.getElementById('data-source-text');
    
    try {
        statusText.textContent = 'جاري التحميل...';
        
        const response = await fetch('data/data.csv');
        const csvText = await response.text();
        
        console.log('📄 تم تحميل CSV، الحجم:', csvText.length);
        
        programsData = parseCSV(csvText);
        
        statusText.textContent = '✓ تم تحميل ' + programsData.length + ' برنامج';
        statusDot.classList.add('connected');
        
        console.log('✅ البرامج:', programsData.map(p => p.name + ' (' + p.degree + ')'));
        
        initSelects();
        
    } catch (error) {
        console.error('❌ خطأ:', error);
        statusText.textContent = '✗ خطأ في التحميل';
        statusDot.classList.add('error');
    }
}

// ========================================
// تحليل CSV
// ========================================
function parseCSV(csvText) {
    const lines = csvText.trim().split('\n');
    if (lines.length < 2) return [];
    
    // الرؤوس
    const headers = lines[0].split(',').map(h => h.trim().replace(/^\uFEFF/, ''));
    console.log('📋 الأعمدة:', headers);
    
    // فهرس الأعمدة
    const col = {};
    headers.forEach((h, i) => {
        const key = h.toLowerCase();
        if (key === 'dept_aname') col.dept = i;
        if (key === 'major_aname') col.program = i;
        if (key === 'degree_aname') col.degree = i;
        if (key === 'semester') col.semester = i;
        if (key === 'gender_aname') col.gender = i;
        if (key === 'join_semester') col.join = i;
        if (key === 'students' || key === 'عدد المنتظمين') col.students = i;
        if (key === 'graduates' || key === 'عدد الخريجين') col.graduates = i;
        if (key === 'course_eval' || key.includes('جودة المقررات')) col.courseEval = i;
        if (key === 'experience_eval' || key.includes('خبرة')) col.expEval = i;
        if (key === 'faculty_total' || key.includes('إجمالي') && key.includes('أعضاء')) col.facultyTotal = i;
        if (key === 'faculty_published' || key.includes('نشروا')) col.facultyPub = i;
        if (key === 'research_count' || key.includes('أبحاث')) col.research = i;
        if (key === 'citations' || key.includes('اقتباس')) col.citations = i;
    });
    
    console.log('🔢 فهرس الأعمدة:', col);
    
    // معالجة البيانات
    const programs = {};
    
    for (let i = 1; i < lines.length; i++) {
        const vals = lines[i].split(',').map(v => v.trim());
        
        const prog = vals[col.program] || '';
        const degree = vals[col.degree] || '';
        const semester = vals[col.semester] || '';
        
        if (!prog || !semester) continue;
        
        const key = prog + '|' + degree;
        if (!programs[key]) {
            programs[key] = {
                name: prog,
                degree: degree,
                dept: vals[col.dept] || '',
                years: {}
            };
        }
        
        const yr = shortYear(semester);
        programs[key].years[yr] = {
            students: vals[col.students] || '',
            graduates: vals[col.graduates] || '',
            course_eval: vals[col.courseEval] || '',
            experience_eval: vals[col.expEval] || '',
            faculty_total: vals[col.facultyTotal] || '',
            faculty_published: vals[col.facultyPub] || '',
            research_count: vals[col.research] || '',
            citations: vals[col.citations] || ''
        };
    }
    
    return Object.values(programs);
}

// ========================================
// حساب المؤشرات من البيانات الخام
// ========================================
function calculateIndicators(raw) {
    if (!raw?.data) return {};
    
    const d = raw.data;
    const ind = {};
    
    // البيانات الأساسية
    const students = parseFloat(d.students) || 0;
    const graduates = parseFloat(d.graduates) || 0;
    const faculty = parseFloat(d.faculty_total) || 0;
    const published = parseFloat(d.faculty_published) || 0;
    const research = parseFloat(d.research_count) || 0;
    const citations = parseFloat(d.citations) || 0;
    
    // المؤشرات المباشرة (من الاستبيانات)
    ind.experience_eval = d.experience_eval ? parseFloat(d.experience_eval).toFixed(2) : null;
    ind.course_eval = d.course_eval ? parseFloat(d.course_eval).toFixed(2) : null;
    
    // المؤشرات غير المتوفرة حالياً
    ind.graduation_rate = null;
    ind.retention_rate = null;
    ind.student_performance = null;
    ind.employment_rate = null;
    ind.employer_eval = null;
    ind.student_publication = null;
    ind.patents = null;
    
    // ═══════════════════════════════════════
    // المؤشرات المحسوبة
    // ═══════════════════════════════════════
    
    // نسبة الطلاب : هيئة التدريس
    if (faculty > 0 && students > 0) {
        ind.student_faculty_ratio = '1:' + Math.round(students / faculty);
    } else {
        ind.student_faculty_ratio = null;
    }
    
    // نسبة النشر العلمي = (الأعضاء الناشرين ÷ إجمالي الأعضاء) × 100
    if (faculty > 0 && published > 0) {
        ind.publication_pct = ((published / faculty) * 100).toFixed(1);
    } else {
        ind.publication_pct = null;
    }
    
    // البحوث لكل عضو = عدد الأبحاث ÷ عدد الأعضاء
    if (faculty > 0 && research > 0) {
        ind.research_per_faculty = (research / faculty).toFixed(2);
    } else {
        ind.research_per_faculty = null;
    }
    
    // الاقتباسات لكل عضو = إجمالي الاقتباسات ÷ عدد الأعضاء
    if (faculty > 0 && citations > 0) {
        ind.citations_per_faculty = (citations / faculty).toFixed(1);
    } else {
        ind.citations_per_faculty = null;
    }
    
    return ind;
}

// ========================================
// تهيئة القوائم
// ========================================
function initSelects() {
    fillProgramSelect('program-select');
    fillProgramSelect('compare-program');
    fillProgramSelect('compare-prog1');
    fillProgramSelect('compare-prog2');
    fillAllYears('compare-common-year');
}

function fillProgramSelect(id) {
    const sel = document.getElementById(id);
    if (!sel) return;
    sel.innerHTML = '<option value="">-- اختر البرنامج --</option>';
    programsData.forEach((p, i) => {
        sel.innerHTML += `<option value="${i}">${p.name} (${p.degree})</option>`;
    });
}

function fillYearSelect(id, progIdx) {
    const sel = document.getElementById(id);
    const prog = programsData[progIdx];
    sel.innerHTML = '<option value="">-- اختر --</option>';
    if (prog?.years) {
        Object.keys(prog.years).sort().forEach(y => {
            sel.innerHTML += `<option value="${y}">${formatYear(y)}</option>`;
        });
        sel.disabled = false;
    }
}

function fillAllYears(id) {
    const sel = document.getElementById(id);
    if (!sel) return;
    const years = new Set();
    programsData.forEach(p => Object.keys(p.years || {}).forEach(y => years.add(y)));
    sel.innerHTML = '<option value="">-- اختر السنة --</option>';
    [...years].sort().forEach(y => {
        sel.innerHTML += `<option value="${y}">${formatYear(y)}</option>`;
    });
}

function fillProgramsForYear(id, year) {
    const sel = document.getElementById(id);
    sel.innerHTML = '<option value="">-- اختر --</option>';
    programsData.forEach((p, i) => {
        if (p.years?.[year]) {
            sel.innerHTML += `<option value="${i}">${p.name} (${p.degree})</option>`;
        }
    });
    sel.disabled = false;
}

// ========================================
// عرض النتائج
// ========================================
function getData(progIdx, year) {
    const prog = programsData[progIdx];
    if (!prog?.years?.[year]) return null;
    return { program: prog, year: year, data: prog.years[year] };
}

function showSingleResults() {
    const progIdx = document.getElementById('program-select').value;
    const year = document.getElementById('year-select').value;
    
    if (!progIdx || !year) return alert('اختر البرنامج والسنة');
    
    const raw = getData(parseInt(progIdx), year);
    if (!raw) return alert('لا توجد بيانات');
    
    currentData = raw;
    const ind = calculateIndicators(raw);
    
    // عرض المعلومات
    document.getElementById('degree-badge').textContent = raw.program.degree;
    document.getElementById('program-name').textContent = raw.program.name;
    document.getElementById('program-year').textContent = 'السنة: ' + formatYear(year);
    
    // عرض المؤشرات
    const grid = document.getElementById('indicators-grid');
    grid.innerHTML = '';
    
    INDICATORS.forEach(indicator => {
        if (indicator.gradOnly && !['الماجستير', 'دكتوراه'].includes(raw.program.degree)) return;
        
        const val = ind[indicator.key];
        const hasVal = val !== null && val !== '';
        
        grid.innerHTML += `
            <div class="indicator-card ${hasVal ? '' : 'no-data'}" data-index="${indicator.id}">
                <span class="card-number">${indicator.id}</span>
                <h3 class="card-title">${indicator.name}</h3>
                <div class="card-value">${hasVal ? val : 'غير متوفر'}</div>
                <div class="card-unit">${indicator.unit}</div>
            </div>
        `;
    });
    
    // عرض البيانات الخام
    const dataGrid = document.getElementById('data-grid');
    dataGrid.innerHTML = '';
    
    Object.entries(RAW_DATA_LABELS).forEach(([key, label]) => {
        const val = raw.data[key];
        const hasVal = val !== null && val !== undefined && val !== '';
        dataGrid.innerHTML += `
            <div class="data-item">
                <div class="data-label">${label}</div>
                <div class="data-value ${hasVal ? '' : 'empty'}">${hasVal ? val : '—'}</div>
            </div>
        `;
    });
    
    // إظهار الأقسام
    document.getElementById('program-info').classList.remove('hidden');
    document.getElementById('indicators-section').classList.remove('hidden');
    document.getElementById('raw-data-section').classList.remove('hidden');
    document.getElementById('program-info').scrollIntoView({ behavior: 'smooth' });
}

// ========================================
// المقارنة
// ========================================
function showCompareResults() {
    const type = document.querySelector('.compare-type-btn.active').dataset.type;
    let data1, data2, title1, title2;
    
    if (type === 'years') {
        const prog = document.getElementById('compare-program').value;
        const y1 = document.getElementById('compare-year1').value;
        const y2 = document.getElementById('compare-year2').value;
        if (!prog || !y1 || !y2) return alert('اختر جميع الحقول');
        
        data1 = getData(parseInt(prog), y1);
        data2 = getData(parseInt(prog), y2);
        title1 = formatYear(y1);
        title2 = formatYear(y2);
    } else {
        const year = document.getElementById('compare-common-year').value;
        const p1 = document.getElementById('compare-prog1').value;
        const p2 = document.getElementById('compare-prog2').value;
        if (!year || !p1 || !p2) return alert('اختر جميع الحقول');
        
        data1 = getData(parseInt(p1), year);
        data2 = getData(parseInt(p2), year);
        title1 = programsData[p1].name;
        title2 = programsData[p2].name;
    }
    
    if (!data1 || !data2) return alert('لا توجد بيانات كافية');
    
    compareData = { item1: data1, item2: data2, title1, title2 };
    
    const ind1 = calculateIndicators(data1);
    const ind2 = calculateIndicators(data2);
    
    document.getElementById('th-col1').textContent = title1;
    document.getElementById('th-col2').textContent = title2;
    
    const tbody = document.getElementById('compare-tbody');
    tbody.innerHTML = '';
    
    INDICATORS.forEach(indicator => {
        const v1 = ind1[indicator.key];
        const v2 = ind2[indicator.key];
        
        let diffHtml = '<span class="diff-neutral">—</span>';
        if (indicator.numeric && v1 && v2) {
            const diff = parseFloat(v1) - parseFloat(v2);
            const sign = diff > 0 ? '+' : '';
            const cls = diff > 0 ? 'diff-positive' : diff < 0 ? 'diff-negative' : 'diff-neutral';
            diffHtml = `<span class="${cls}">${sign}${diff.toFixed(2)}</span>`;
        }
        
        tbody.innerHTML += `
            <tr>
                <td>${indicator.name}</td>
                <td class="value-cell">${v1 || '—'}</td>
                <td class="value-cell">${v2 || '—'}</td>
                <td>${diffHtml}</td>
            </tr>
        `;
    });
    
    document.getElementById('compare-results').classList.remove('hidden');
    document.getElementById('compare-results').scrollIntoView({ behavior: 'smooth' });
}

// ========================================
// التصدير
// ========================================
function exportToPDF() {
    if (!currentData) return;
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    
    doc.setFontSize(16);
    doc.text('KPI Report - ' + currentData.program.name, 105, 20, { align: 'center' });
    doc.setFontSize(12);
    doc.text('Year: ' + formatYear(currentData.year), 105, 30, { align: 'center' });
    
    const ind = calculateIndicators(currentData);
    const tableData = INDICATORS.map(i => [i.name, ind[i.key] || '-', i.unit]);
    
    doc.autoTable({
        startY: 40,
        head: [['Indicator', 'Value', 'Unit']],
        body: tableData,
        styles: { halign: 'center', fontSize: 9 },
        headStyles: { fillColor: [13, 79, 79] }
    });
    
    doc.save(`KPI-${currentData.program.name}-${formatYear(currentData.year)}.pdf`);
}

function exportCompareToPDF() {
    if (!compareData.item1) return;
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    
    doc.setFontSize(14);
    doc.text('KPI Comparison', 105, 20, { align: 'center' });
    
    const ind1 = calculateIndicators(compareData.item1);
    const ind2 = calculateIndicators(compareData.item2);
    const tableData = INDICATORS.map(i => [i.name, ind1[i.key] || '-', ind2[i.key] || '-']);
    
    doc.autoTable({
        startY: 35,
        head: [['Indicator', compareData.title1, compareData.title2]],
        body: tableData,
        styles: { halign: 'center', fontSize: 9 },
        headStyles: { fillColor: [13, 79, 79] }
    });
    
    doc.save('KPI-Comparison.pdf');
}

function exportToExcel() {
    if (!currentData) return;
    const ind = calculateIndicators(currentData);
    
    const data = [
        ['تقرير مؤشرات الأداء'],
        ['البرنامج: ' + currentData.program.name],
        ['السنة: ' + formatYear(currentData.year)],
        [],
        ['المؤشر', 'القيمة', 'الوحدة'],
        ...INDICATORS.map(i => [i.name, ind[i.key] || '—', i.unit])
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المؤشرات');
    XLSX.writeFile(wb, `مؤشرات-${formatYear(currentData.year)}.xlsx`);
}

function exportCompareToExcel() {
    if (!compareData.item1) return;
    const ind1 = calculateIndicators(compareData.item1);
    const ind2 = calculateIndicators(compareData.item2);
    
    const data = [
        ['مقارنة المؤشرات'],
        [],
        ['المؤشر', compareData.title1, compareData.title2],
        ...INDICATORS.map(i => [i.name, ind1[i.key] || '—', ind2[i.key] || '—'])
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المقارنة');
    XLSX.writeFile(wb, 'مقارنة-المؤشرات.xlsx');
}

// ========================================
// مستمعات الأحداث
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    
    // التنقل
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.view-section').forEach(v => v.classList.add('hidden'));
            btn.classList.add('active');
            document.getElementById(btn.dataset.view + '-view').classList.remove('hidden');
        });
    });
    
    // العرض الفردي
    document.getElementById('program-select').addEventListener('change', e => {
        if (e.target.value) {
            fillYearSelect('year-select', parseInt(e.target.value));
        } else {
            document.getElementById('year-select').disabled = true;
        }
        document.getElementById('show-btn').disabled = true;
        ['program-info', 'indicators-section', 'raw-data-section'].forEach(id => 
            document.getElementById(id).classList.add('hidden'));
    });
    
    document.getElementById('year-select').addEventListener('change', e => {
        document.getElementById('show-btn').disabled = !e.target.value;
    });
    
    document.getElementById('show-btn').addEventListener('click', showSingleResults);
    
    // نوع المقارنة
    document.querySelectorAll('.compare-type-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.compare-type-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById('compare-years-form').classList.toggle('hidden', btn.dataset.type !== 'years');
            document.getElementById('compare-programs-form').classList.toggle('hidden', btn.dataset.type !== 'programs');
            document.getElementById('compare-results').classList.add('hidden');
            document.getElementById('compare-btn').disabled = true;
        });
    });
    
    // مقارنة السنوات
    document.getElementById('compare-program').addEventListener('change', e => {
        if (e.target.value) {
            fillYearSelect('compare-year1', parseInt(e.target.value));
            fillYearSelect('compare-year2', parseInt(e.target.value));
        }
        updateCompareBtn();
    });
    
    ['compare-year1', 'compare-year2'].forEach(id => {
        document.getElementById(id).addEventListener('change', updateCompareBtn);
    });
    
    // مقارنة البرامج
    document.getElementById('compare-common-year').addEventListener('change', e => {
        if (e.target.value) {
            fillProgramsForYear('compare-prog1', e.target.value);
            fillProgramsForYear('compare-prog2', e.target.value);
        }
        updateCompareBtn();
    });
    
    ['compare-prog1', 'compare-prog2'].forEach(id => {
        document.getElementById(id).addEventListener('change', updateCompareBtn);
    });
    
    document.getElementById('compare-btn').addEventListener('click', showCompareResults);
});

function updateCompareBtn() {
    const type = document.querySelector('.compare-type-btn.active').dataset.type;
    let valid = false;
    
    if (type === 'years') {
        const prog = document.getElementById('compare-program').value;
        const y1 = document.getElementById('compare-year1').value;
        const y2 = document.getElementById('compare-year2').value;
        valid = prog && y1 && y2 && y1 !== y2;
    } else {
        const year = document.getElementById('compare-common-year').value;
        const p1 = document.getElementById('compare-prog1').value;
        const p2 = document.getElementById('compare-prog2').value;
        valid = year && p1 && p2 && p1 !== p2;
    }
    
    document.getElementById('compare-btn').disabled = !valid;
}
