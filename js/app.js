/**
 * مؤشرات الأداء - كلية الشريعة والأنظمة
 * نسخة مصححة مع دعم Google Sheets
 */

// ========================================
// المتغيرات العامة
// ========================================
let programsData = [];
let currentData = null;
let compareData = { item1: null, item2: null, title1: '', title2: '' };

// تعريف المؤشرات
const INDICATORS = [
    { id: 1, name: "تقويم الطالب لجودة خبرات التعلم", unit: "درجة", key: "experience_eval", numeric: true },
    { id: 2, name: "تقييم الطالب لجودة المقررات", unit: "درجة", key: "course_eval", numeric: true },
    { id: 3, name: "معدّل التخرج بالوقت المحدد", unit: "%", key: "graduation_rate" },
    { id: 4, name: "معدّل استبقاء طلاب السنة الأولى", unit: "%", key: "retention_rate" },
    { id: 5, name: "مستوى أداء الطالب", unit: "%", key: "student_performance" },
    { id: 6, name: "توظيف الخريجين", unit: "%", key: "employment_rate" },
    { id: 7, name: "تقويم جهات التوظيف", unit: "درجة", key: "employer_eval" },
    { id: 8, name: "نسبة الطلاب/هيئة التدريس", unit: "نسبة", key: "student_faculty_ratio" },
    { id: 9, name: "نسبة النشر العلمي", unit: "%", key: "publication_pct", numeric: true },
    { id: 10, name: "البحوث/عضو هيئة تدريس", unit: "بحث", key: "research_per_faculty", numeric: true },
    { id: 11, name: "الاقتباسات/عضو هيئة تدريس", unit: "اقتباس", key: "citations_per_faculty", numeric: true },
    { id: 12, name: "نسبة نشر طلاب الدراسات العليا", unit: "%", key: "student_publication", gradOnly: true },
    { id: 13, name: "براءات الاختراع", unit: "براءة", key: "patents", gradOnly: true },
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
// دالة تحويل السنة: 46 → 1446
// ========================================
function formatYear(semester) {
    if (!semester) return '';
    let num = parseFloat(semester);
    if (isNaN(num)) return semester;
    num = Math.floor(num);
    if (num < 100) {
        return '14' + String(num).padStart(2, '0');
    }
    return String(num);
}

function getShortYear(semester) {
    if (!semester) return '';
    let num = parseFloat(semester);
    if (isNaN(num)) return semester;
    num = Math.floor(num);
    if (num >= 1400) {
        return String(num).slice(-2);
    }
    return String(num);
}

// ========================================
// تحميل البيانات
// ========================================
async function loadData() {
    const statusDot = document.getElementById('status-dot');
    const statusText = document.getElementById('data-source-text');
    
    try {
        if (typeof GOOGLE_SHEET_URL !== 'undefined' && GOOGLE_SHEET_URL && GOOGLE_SHEET_URL.trim() !== '') {
            statusText.textContent = 'جاري الاتصال بـ Google Sheets...';
            console.log('🔗 محاولة الاتصال بـ:', GOOGLE_SHEET_URL);
            
            const response = await fetch(GOOGLE_SHEET_URL);
            if (!response.ok) {
                throw new Error('فشل الاتصال: ' + response.status);
            }
            
            const csvText = await response.text();
            console.log('📄 تم استلام CSV، الحجم:', csvText.length, 'حرف');
            console.log('📄 أول 500 حرف:', csvText.substring(0, 500));
            
            programsData = parseCSVToPrograms(csvText);
            
            if (programsData.length === 0) {
                throw new Error('لم يتم العثور على بيانات في الملف');
            }
            
            statusText.textContent = '✓ متصل بـ Google Sheets (' + programsData.length + ' برنامج)';
            statusDot.classList.add('connected');
            
        } else {
            console.log('📂 تحميل من الملف المحلي...');
            const response = await fetch('data/programs.json');
            programsData = await response.json();
            statusText.textContent = '✓ البيانات من ملف محلي (' + programsData.length + ' برنامج)';
            statusDot.classList.add('connected');
        }
        
        console.log('✅ تم تحميل', programsData.length, 'برنامج');
        console.log('📊 البرامج:', programsData.map(p => p.name));
        
        initializeAllSelects();
        
    } catch (error) {
        console.error('❌ خطأ:', error);
        statusText.textContent = '✗ خطأ: ' + error.message;
        statusDot.classList.add('error');
        
        try {
            console.log('🔄 محاولة تحميل الملف المحلي كاحتياطي...');
            const response = await fetch('data/programs.json');
            programsData = await response.json();
            statusText.textContent = '⚠️ تم استخدام الملف المحلي (فشل Google Sheets)';
            initializeAllSelects();
        } catch (e) {
            console.error('❌ فشل تحميل الملف المحلي أيضاً:', e);
        }
    }
}

// ========================================
// تحويل CSV إلى JSON
// ========================================
function parseCSVToPrograms(csvText) {
    console.log('🔄 بدء تحليل CSV...');
    
    csvText = csvText.trim();
    
    const firstLine = csvText.split('\n')[0];
    let delimiter = ',';
    if (firstLine.includes(';') && !firstLine.includes(',')) {
        delimiter = ';';
    } else if (firstLine.split(';').length > firstLine.split(',').length) {
        delimiter = ';';
    }
    console.log('📌 الفاصل المستخدم:', delimiter === ',' ? 'فاصلة (,)' : 'فاصلة منقوطة (;)');
    
    const lines = csvText.split('\n').filter(l => l.trim());
    console.log('📝 عدد الأسطر:', lines.length);
    
    if (lines.length < 2) {
        console.error('❌ لا توجد بيانات كافية');
        return [];
    }
    
    const headers = parseCSVLine(lines[0], delimiter);
    console.log('📋 الرؤوس:', headers);
    
    const colMap = findColumns(headers);
    console.log('🗺️ خريطة الأعمدة:', colMap);
    
    if (colMap.program === -1) {
        console.error('❌ لم يتم العثور على عمود البرنامج');
        return [];
    }
    
    const programs = {};
    
    for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i], delimiter);
        if (values.length < 3) continue;
        
        const prog = values[colMap.program]?.trim();
        const degree = colMap.degree !== -1 ? values[colMap.degree]?.trim() : '';
        const semester = colMap.semester !== -1 ? values[colMap.semester]?.trim() : '';
        
        if (!prog || !semester) continue;
        
        const key = `${prog}|${degree}`;
        if (!programs[key]) {
            programs[key] = {
                name: prog,
                degree: degree,
                dept: colMap.dept !== -1 ? values[colMap.dept]?.trim() : '',
                years: {}
            };
        }
        
        const shortYear = getShortYear(semester);
        if (!programs[key].years[shortYear]) {
            programs[key].years[shortYear] = {};
        }
        
        const gender = colMap.gender !== -1 ? (values[colMap.gender]?.trim() || 'All') : 'All';
        const joinSem = colMap.joinSemester !== -1 ? (values[colMap.joinSemester]?.trim() || 'All') : 'All';
        const filterKey = `${gender}_${joinSem}`;
        
        programs[key].years[shortYear][filterKey] = {
            students: colMap.students !== -1 ? cleanNumber(values[colMap.students]) : '',
            graduates: colMap.graduates !== -1 ? cleanNumber(values[colMap.graduates]) : '',
            course_eval: colMap.courseEval !== -1 ? cleanNumber(values[colMap.courseEval]) : '',
            experience_eval: colMap.expEval !== -1 ? cleanNumber(values[colMap.expEval]) : '',
            faculty_total: colMap.facultyTotal !== -1 ? cleanNumber(values[colMap.facultyTotal]) : '',
            faculty_published: colMap.facultyPublished !== -1 ? cleanNumber(values[colMap.facultyPublished]) : '',
            research_count: colMap.researchCount !== -1 ? cleanNumber(values[colMap.researchCount]) : '',
            citations: colMap.citations !== -1 ? cleanNumber(values[colMap.citations]) : '',
        };
    }
    
    const result = Object.values(programs);
    console.log('✅ تم تحليل', result.length, 'برنامج');
    return result;
}

function parseCSVLine(line, delimiter) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === delimiter && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current.trim());
    
    return result.map(v => v.replace(/^"|"$/g, '').replace(/^\uFEFF/, ''));
}

function findColumns(headers) {
    const map = {
        dept: -1,
        program: -1,
        degree: -1,
        semester: -1,
        gender: -1,
        joinSemester: -1,
        students: -1,
        graduates: -1,
        courseEval: -1,
        expEval: -1,
        facultyTotal: -1,
        facultyPublished: -1,
        researchCount: -1,
        citations: -1
    };
    
    headers.forEach((h, i) => {
        const header = h.toLowerCase().trim();
        
        if (header.includes('dept') || header.includes('قسم')) map.dept = i;
        if (header.includes('major') || header.includes('برنامج')) map.program = i;
        if (header.includes('degree') || header.includes('درجة')) map.degree = i;
        if (header.includes('semester') || header.includes('فصل') || header.includes('سنة')) map.semester = i;
        if (header.includes('gender') || header.includes('جنس')) map.gender = i;
        if (header.includes('join') || header.includes('التحاق')) map.joinSemester = i;
        if (header.includes('منتظم') || header.includes('طلاب')) map.students = i;
        if (header.includes('خريج')) map.graduates = i;
        if (header.includes('مقرر')) map.courseEval = i;
        if (header.includes('خبرة')) map.expEval = i;
        if (header.includes('إجمالي') && header.includes('أعضاء')) map.facultyTotal = i;
        if (header.includes('نشر') && header.includes('أعضاء')) map.facultyPublished = i;
        if (header.includes('أبحاث') || header.includes('بحث')) map.researchCount = i;
        if (header.includes('اقتباس')) map.citations = i;
    });
    
    return map;
}

function cleanNumber(value) {
    if (!value) return '';
    const str = String(value).trim();
    if (str === '' || str === 'NaN' || str === 'undefined') return '';
    return str;
}

// ========================================
// تهيئة القوائم
// ========================================
function initializeAllSelects() {
    console.log('🔧 تهيئة القوائم...');
    populateProgramSelect('program-select');
    populateProgramSelect('compare-program');
    populateProgramSelect('compare-prog1');
    populateProgramSelect('compare-prog2');
    populateAllYears('compare-common-year');
}

function populateProgramSelect(selectId) {
    const select = document.getElementById(selectId);
    if (!select) return;
    
    select.innerHTML = '<option value="">-- اختر البرنامج --</option>';
    
    programsData.forEach((p, i) => {
        const opt = document.createElement('option');
        opt.value = i;
        opt.textContent = `${p.name} (${p.degree})`;
        select.appendChild(opt);
    });
    
    console.log(`📝 تم ملء ${selectId} بـ ${programsData.length} برنامج`);
}

function populateYearSelect(selectId, programIndex) {
    const select = document.getElementById(selectId);
    const program = programsData[programIndex];
    
    select.innerHTML = '<option value="">-- اختر --</option>';
    
    if (program?.years) {
        const years = Object.keys(program.years).sort();
        years.forEach(y => {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = formatYear(y);
            select.appendChild(opt);
        });
        select.disabled = false;
        console.log(`📅 السنوات المتاحة لـ ${program.name}:`, years.map(y => formatYear(y)));
    } else {
        select.disabled = true;
    }
}

function populateAllYears(selectId) {
    const select = document.getElementById(selectId);
    if (!select) return;
    
    const allYears = new Set();
    programsData.forEach(p => Object.keys(p.years || {}).forEach(y => allYears.add(y)));
    
    select.innerHTML = '<option value="">-- اختر السنة --</option>';
    [...allYears].sort().forEach(y => {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = formatYear(y);
        select.appendChild(opt);
    });
}

function populateProgramsWithYear(selectId, year) {
    const select = document.getElementById(selectId);
    select.innerHTML = '<option value="">-- اختر --</option>';
    
    programsData.forEach((p, i) => {
        if (p.years && p.years[year]) {
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = `${p.name} (${p.degree})`;
            select.appendChild(opt);
        }
    });
    select.disabled = false;
}

// ========================================
// الحصول على البيانات وحساب المؤشرات
// ========================================
function getData(programIndex, year) {
    const program = programsData[programIndex];
    if (!program?.years?.[year]) {
        console.log('❌ لا توجد بيانات للبرنامج', programIndex, 'السنة', year);
        return null;
    }
    
    const yearData = program.years[year];
    let data = yearData["All_All"] || yearData["All_all"];
    
    if (!data) {
        for (const k of Object.keys(yearData)) {
            if (k.startsWith("All_")) { 
                data = yearData[k]; 
                break; 
            }
        }
    }
    
    if (!data && Object.keys(yearData).length > 0) {
        data = yearData[Object.keys(yearData)[0]];
    }
    
    console.log('📊 بيانات', program.name, 'سنة', year, ':', data);
    return { program, year, data };
}

function calculateIndicators(rawData) {
    if (!rawData?.data) return {};
    const d = rawData.data;
    const ind = {};
    
    ind.experience_eval = parseFloat(d.experience_eval) || null;
    ind.course_eval = parseFloat(d.course_eval) || null;
    ind.graduation_rate = null;
    ind.retention_rate = null;
    ind.student_performance = null;
    ind.employment_rate = null;
    ind.employer_eval = null;
    
    const students = parseFloat(d.students) || 0;
    const faculty = parseFloat(d.faculty_total) || 0;
    const published = parseFloat(d.faculty_published) || 0;
    const research = parseFloat(d.research_count) || 0;
    const citations = parseFloat(d.citations) || 0;
    
    ind.student_faculty_ratio = (faculty > 0 && students > 0) ? `1:${Math.round(students/faculty)}` : null;
    ind.publication_pct = (faculty > 0 && published > 0) ? ((published/faculty)*100).toFixed(1) : null;
    ind.research_per_faculty = (faculty > 0 && research > 0) ? (research/faculty).toFixed(2) : null;
    ind.citations_per_faculty = (faculty > 0 && citations > 0) ? (citations/faculty).toFixed(1) : null;
    ind.student_publication = null;
    ind.patents = null;
    
    return ind;
}

// ========================================
// العرض الفردي
// ========================================
function showSingleResults() {
    const progIdx = document.getElementById('program-select').value;
    const year = document.getElementById('year-select').value;
    
    console.log('🔍 عرض البرنامج:', progIdx, 'السنة:', year);
    
    if (!progIdx || !year) {
        alert('يرجى اختيار البرنامج والسنة');
        return;
    }
    
    const rawData = getData(parseInt(progIdx), year);
    if (!rawData) { 
        alert('لا توجد بيانات لهذا الاختيار'); 
        return; 
    }
    
    currentData = rawData;
    const indicators = calculateIndicators(rawData);
    
    document.getElementById('degree-badge').textContent = rawData.program.degree;
    document.getElementById('program-name').textContent = rawData.program.name;
    document.getElementById('program-year').textContent = `السنة: ${formatYear(year)}`;
    
    const grid = document.getElementById('indicators-grid');
    grid.innerHTML = '';
    
    INDICATORS.forEach(ind => {
        if (ind.gradOnly && !['الماجستير','دكتوراه'].includes(rawData.program.degree)) return;
        
        const val = indicators[ind.key];
        const hasVal = val !== null && val !== '';
        
        const card = document.createElement('div');
        card.className = `indicator-card ${hasVal ? '' : 'no-data'}`;
        card.setAttribute('data-index', ind.id);
        card.innerHTML = `
            <span class="card-number">${ind.id}</span>
            <h3 class="card-title">${ind.name}</h3>
            <div class="card-value">${hasVal ? val : 'غير متوفر'}</div>
            <div class="card-unit">${ind.unit}</div>
        `;
        grid.appendChild(card);
    });
    
    const dataGrid = document.getElementById('data-grid');
    dataGrid.innerHTML = '';
    
    Object.entries(RAW_DATA_LABELS).forEach(([key, label]) => {
        const val = rawData.data?.[key];
        const hasVal = val !== null && val !== undefined && val !== '';
        
        const item = document.createElement('div');
        item.className = 'data-item';
        item.innerHTML = `
            <div class="data-label">${label}</div>
            <div class="data-value ${hasVal ? '' : 'empty'}">${hasVal ? val : '—'}</div>
        `;
        dataGrid.appendChild(item);
    });
    
    document.getElementById('program-info').classList.remove('hidden');
    document.getElementById('indicators-section').classList.remove('hidden');
    document.getElementById('raw-data-section').classList.remove('hidden');
    document.getElementById('program-info').scrollIntoView({ behavior: 'smooth' });
}

// ========================================
// المقارنة
// ========================================
function showCompareResults() {
    const compareType = document.querySelector('.compare-type-btn.active').dataset.type;
    let data1, data2, title1, title2;
    
    if (compareType === 'years') {
        const progIdx = document.getElementById('compare-program').value;
        const y1 = document.getElementById('compare-year1').value;
        const y2 = document.getElementById('compare-year2').value;
        
        if (!progIdx || !y1 || !y2) { 
            alert('اختر جميع الحقول'); 
            return; 
        }
        
        data1 = getData(parseInt(progIdx), y1);
        data2 = getData(parseInt(progIdx), y2);
        title1 = formatYear(y1);
        title2 = formatYear(y2);
        
    } else {
        const year = document.getElementById('compare-common-year').value;
        const p1 = document.getElementById('compare-prog1').value;
        const p2 = document.getElementById('compare-prog2').value;
        
        if (!year || !p1 || !p2) { 
            alert('اختر جميع الحقول'); 
            return; 
        }
        
        data1 = getData(parseInt(p1), year);
        data2 = getData(parseInt(p2), year);
        title1 = programsData[p1].name;
        title2 = programsData[p2].name;
    }
    
    if (!data1 || !data2) { 
        alert('لا توجد بيانات كافية للمقارنة'); 
        return; 
    }
    
    compareData = { item1: data1, item2: data2, title1, title2 };
    
    const ind1 = calculateIndicators(data1);
    const ind2 = calculateIndicators(data2);
    
    document.getElementById('th-col1').textContent = title1;
    document.getElementById('th-col2').textContent = title2;
    
    const tbody = document.getElementById('compare-tbody');
    tbody.innerHTML = '';
    
    INDICATORS.forEach(ind => {
        const v1 = ind1[ind.key];
        const v2 = ind2[ind.key];
        
        let diffHtml = '<span class="diff-neutral">—</span>';
        if (ind.numeric && v1 && v2) {
            const diff = parseFloat(v1) - parseFloat(v2);
            const sign = diff > 0 ? '+' : '';
            const cls = diff > 0 ? 'diff-positive' : diff < 0 ? 'diff-negative' : 'diff-neutral';
            diffHtml = `<span class="${cls}">${sign}${diff.toFixed(2)}</span>`;
        }
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${ind.name}</td>
            <td class="value-cell">${v1 || '—'}</td>
            <td class="value-cell">${v2 || '—'}</td>
            <td>${diffHtml}</td>
        `;
        tbody.appendChild(row);
    });
    
    document.getElementById('compare-results').classList.remove('hidden');
    document.getElementById('compare-results').scrollIntoView({ behavior: 'smooth' });
}

// ========================================
// التصدير - PDF
// ========================================
function exportToPDF() {
    if (!currentData) return;
    
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    
    doc.setFont('helvetica');
    doc.setFontSize(16);
    doc.text('KPI Report - ' + currentData.program.name, 105, 20, { align: 'center' });
    doc.setFontSize(12);
    doc.text('Year: ' + formatYear(currentData.year), 105, 30, { align: 'center' });
    
    const indicators = calculateIndicators(currentData);
    const tableData = INDICATORS.map(ind => [ind.name, indicators[ind.key] || '-', ind.unit]);
    
    doc.autoTable({
        startY: 40,
        head: [['Indicator', 'Value', 'Unit']],
        body: tableData,
        styles: { font: 'helvetica', halign: 'center', fontSize: 9 },
        headStyles: { fillColor: [13, 79, 79] }
    });
    
    doc.save(`KPI-${currentData.program.name}-${formatYear(currentData.year)}.pdf`);
}

function exportCompareToPDF() {
    if (!compareData.item1) return;
    
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    
    doc.setFont('helvetica');
    doc.setFontSize(14);
    doc.text('KPI Comparison Report', 105, 20, { align: 'center' });
    
    const ind1 = calculateIndicators(compareData.item1);
    const ind2 = calculateIndicators(compareData.item2);
    const tableData = INDICATORS.map(ind => [ind.name, ind1[ind.key] || '-', ind2[ind.key] || '-']);
    
    doc.autoTable({
        startY: 35,
        head: [['Indicator', compareData.title1, compareData.title2]],
        body: tableData,
        styles: { font: 'helvetica', halign: 'center', fontSize: 9 },
        headStyles: { fillColor: [13, 79, 79] }
    });
    
    doc.save('KPI-Comparison.pdf');
}

// ========================================
// التصدير - Excel
// ========================================
function exportToExcel() {
    if (!currentData) return;
    
    const indicators = calculateIndicators(currentData);
    
    const data = [
        ['تقرير مؤشرات الأداء'],
        [`البرنامج: ${currentData.program.name}`],
        [`السنة: ${formatYear(currentData.year)}`],
        [],
        ['المؤشر', 'القيمة', 'الوحدة'],
        ...INDICATORS.map(ind => [ind.name, indicators[ind.key] || '—', ind.unit])
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المؤشرات');
    XLSX.writeFile(wb, `مؤشرات-${currentData.program.name}-${formatYear(currentData.year)}.xlsx`);
}

function exportCompareToExcel() {
    if (!compareData.item1) return;
    
    const ind1 = calculateIndicators(compareData.item1);
    const ind2 = calculateIndicators(compareData.item2);
    
    const data = [
        ['تقرير مقارنة المؤشرات'],
        [],
        ['المؤشر', compareData.title1, compareData.title2],
        ...INDICATORS.map(ind => [ind.name, ind1[ind.key] || '—', ind2[ind.key] || '—'])
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
    console.log('🚀 بدء التطبيق...');
    loadData();
    
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.view-section').forEach(v => v.classList.add('hidden'));
            btn.classList.add('active');
            document.getElementById(`${btn.dataset.view}-view`).classList.remove('hidden');
        });
    });
    
    document.getElementById('program-select').addEventListener('change', e => {
        if (e.target.value) {
            populateYearSelect('year-select', parseInt(e.target.value));
        } else {
            document.getElementById('year-select').disabled = true;
            document.getElementById('year-select').innerHTML = '<option value="">-- اختر السنة --</option>';
        }
        document.getElementById('show-btn').disabled = true;
        ['program-info', 'indicators-section', 'raw-data-section'].forEach(id => 
            document.getElementById(id).classList.add('hidden'));
    });
    
    document.getElementById('year-select').addEventListener('change', e => {
        const prog = document.getElementById('program-select').value;
        document.getElementById('show-btn').disabled = !e.target.value || !prog;
    });
    
    document.getElementById('show-btn').addEventListener('click', showSingleResults);
    
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
    
    document.getElementById('compare-program').addEventListener('change', e => {
        if (e.target.value) {
            populateYearSelect('compare-year1', parseInt(e.target.value));
            populateYearSelect('compare-year2', parseInt(e.target.value));
        } else {
            document.getElementById('compare-year1').disabled = true;
            document.getElementById('compare-year2').disabled = true;
        }
        updateCompareBtn();
    });
    
    ['compare-year1', 'compare-year2'].forEach(id => {
        document.getElementById(id).addEventListener('change', updateCompareBtn);
    });
    
    document.getElementById('compare-common-year').addEventListener('change', e => {
        if (e.target.value) {
            populateProgramsWithYear('compare-prog1', e.target.value);
            populateProgramsWithYear('compare-prog2', e.target.value);
        } else {
            document.getElementById('compare-prog1').disabled = true;
            document.getElementById('compare-prog2').disabled = true;
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

// للتصحيح
window.debugData = () => {
    console.log('📊 البيانات:', programsData);
    console.log('📊 الحالية:', currentData);
};
