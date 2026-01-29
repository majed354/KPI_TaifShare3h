/**
 * مؤشرات الأداء - كلية الشريعة
 * نسخة كاملة مع المقارنة والتصدير وربط Google Sheets
 *
 * ✅ إصلاحات مهمة:
 * - قراءة CSV تلقائياً سواء كان delimiter (,) أو (;)
 * - دعم CSV مع علامات الاقتباس "..."
 * - التعامل مع سطر Excel الشهير: sep=,
 * - محاولة تحويل روابط Google Sheets الشائعة إلى رابط CSV مناسب
 * - رسائل خطأ أوضح لو الرابط ليس CSV (مثلاً صفحة HTML أو صلاحيات)
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
// أدوات مساعدة
// ========================================
function normalizeKey(s) {
    // توحيد العناوين (مهم لاختلاف التشكيل/المسافات)
    return String(s || '')
        .trim()
        .toLowerCase()
        .replace(/[\u064B-\u0652]/g, '')       // إزالة التشكيل
        .replace(/\s+/g, '')                   // إزالة المسافات
        .replace(/[ـ_]/g, '')                  // إزالة التطويل/الشرطة السفلية
        .replace(/[^\w\u0600-\u06FF]/g, '');   // إبقاء الحروف/الأرقام العربية واللاتينية
}

function toNumber(v) {
    if (v === null || v === undefined) return NaN;
    let s = String(v).trim();
    if (!s) return NaN;

    // تحويل الأرقام العربية/الفارسية إلى لاتينية
    const map = {
        '٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9',
        '۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9'
    };
    s = s.replace(/[٠-٩۰-۹]/g, ch => map[ch] ?? ch);

    // فواصل عربية/آلاف
    s = s.replace(/٬/g, '');   // آلاف عربي
    s = s.replace(/،/g, ',');  // فاصلة عربية إلى فاصلة عادية
    s = s.replace(/٫/g, '.');  // فاصلة عشرية عربية إلى نقطة

    // إذا لا يوجد نقطة وفيه فاصلة واحدة غالباً فاصلة عشرية
    if (!s.includes('.') && (s.match(/,/g) || []).length === 1) {
        s = s.replace(',', '.');
    } else {
        // غير ذلك اعتبر الفواصل آلاف
        s = s.replace(/,/g, '');
    }

    return parseFloat(s);
}

function normalizeGoogleSheetUrl(url) {
    if (!url) return url;
    const u = url.trim();

    // إذا هو فعلاً CSV
    if (u.includes('output=csv') || u.includes('format=csv') || u.endsWith('.csv')) return u;

    // روابط publish html
    if (u.includes('/pubhtml')) {
        // غالباً نشر على الويب: pubhtml -> pub?output=csv
        const base = u.replace('/pubhtml', '/pub');
        return base.includes('?') ? `${base}&output=csv` : `${base}?output=csv`;
    }

    // روابط /pub بدون output
    if (u.includes('/pub') && !u.includes('output=')) {
        return u.includes('?') ? `${u}&output=csv` : `${u}?output=csv`;
    }

    // روابط spreadsheets/d/<ID>/edit... -> export?format=csv&gid=...
    const m = u.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (m && m[1]) {
        const id = m[1];
        let gid = '0';
        const gidMatch = u.match(/[?&]gid=(\d+)/);
        if (gidMatch && gidMatch[1]) gid = gidMatch[1];
        return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
    }

    return u;
}

async function fetchCsvText(url) {
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) {
        throw new Error(`فشل الاتصال (${res.status})`);
    }
    const text = await res.text();

    // لو رجعت HTML (صلاحيات/صفحة تسجيل دخول/عرض)
    const t = text.trim().toLowerCase();
    if (t.startsWith('<!doctype html') || t.startsWith('<html')) {
        throw new Error('الرابط لا يُرجع CSV (يبدو صفحة HTML/صلاحيات). تأكد من "نشر على الويب" بصيغة CSV.');
    }
    // أحياناً تظهر صفحة Google Accounts داخل النص
    if (t.includes('accounts.google.com') || t.includes('sign in')) {
        throw new Error('الرابط يتطلب تسجيل دخول. تأكد أن الملف منشور للعامة (Publish to web) بصيغة CSV.');
    }

    return text;
}

// ========================================
// تحميل البيانات
// ========================================
async function loadData() {
    const statusDot = document.getElementById('status-dot');
    const statusText = document.getElementById('data-source-text');

    statusDot.classList.remove('connected', 'error');

    try {
        // التحقق من وجود رابط Google Sheets
        if (typeof GOOGLE_SHEET_URL !== 'undefined' && GOOGLE_SHEET_URL && GOOGLE_SHEET_URL.trim() !== '') {
            const finalUrl = normalizeGoogleSheetUrl(GOOGLE_SHEET_URL);

            statusText.textContent = 'جاري الاتصال بـ Google Sheets...';
            const csvText = await fetchCsvText(finalUrl);

            programsData = parseCSVToPrograms(csvText);

            if (!Array.isArray(programsData) || programsData.length === 0) {
                throw new Error('تم تحميل CSV لكن لم يتم العثور على أي برامج. غالباً مشكلة فاصل CSV أو أسماء الأعمدة.');
            }

            statusText.textContent = '✓ متصل بـ Google Sheets';
            statusDot.classList.add('connected');
        } else {
            // استخدام الملف المحلي
            statusText.textContent = 'جاري تحميل البيانات المحلية...';
            const response = await fetch('data/programs.json', { cache: 'no-store' });
            if (!response.ok) throw new Error(`فشل تحميل الملف المحلي (${response.status})`);
            programsData = await response.json();

            statusText.textContent = '✓ البيانات من ملف محلي';
            statusDot.classList.add('connected');
        }

        console.log('تم تحميل', programsData.length, 'برنامج');
        initializeAllSelects();

    } catch (error) {
        console.error('خطأ:', error);
        statusText.textContent = `✗ خطأ: ${error.message || 'تعذر التحميل'}`;
        statusDot.classList.add('error');
    }
}

// ========================================
// CSV Parser (يدعم ", ; \t" + علامات الاقتباس)
// ========================================
function detectDelimiter(firstLine) {
    const candidates = [',', ';', '\t'];
    let best = ',';
    let bestCount = -1;

    for (const d of candidates) {
        const re = d === '\t' ? /\t/g : new RegExp(`\\${d}`, 'g');
        const count = (firstLine.match(re) || []).length;
        if (count > bestCount) {
            bestCount = count;
            best = d;
        }
    }
    return best;
}

function parseCsvRows(text, delimiter) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];

        if (inQuotes) {
            if (ch === '"') {
                if (text[i + 1] === '"') {
                    field += '"';
                    i++;
                } else {
                    inQuotes = false;
                }
            } else {
                field += ch;
            }
        } else {
            if (ch === '"') {
                inQuotes = true;
            } else if (ch === delimiter) {
                row.push(field);
                field = '';
            } else if (ch === '\n') {
                row.push(field);
                field = '';
                if (row.some(v => String(v).trim() !== '')) rows.push(row);
                row = [];
            } else {
                field += ch;
            }
        }
    }

    // آخر سطر
    row.push(field);
    if (row.some(v => String(v).trim() !== '')) rows.push(row);

    return rows;
}

// تحويل CSV إلى JSON (هيكل البرامج)
function parseCSVToPrograms(csvText) {
    if (!csvText) return [];

    // توحيد نهاية الأسطر + إزالة BOM
    let text = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/^\uFEFF/, '');

    // معالجة "sep=," أو "sep=;" (Excel)
    const nonEmptyLines = text.split('\n').filter(l => l.trim() !== '');
    if (nonEmptyLines.length < 2) return [];

    let delimiter = null;
    const firstLine = nonEmptyLines[0].trim();
    if (firstLine.toLowerCase().startsWith('sep=')) {
        delimiter = firstLine.substring(4, 5) || ',';
        // احذف سطر sep= من النص
        text = text.split('\n').slice(1).join('\n');
    }

    const previewFirstLine = text.split('\n').find(l => l.trim() !== '') || '';
    if (!delimiter) delimiter = detectDelimiter(previewFirstLine);

    const rows = parseCsvRows(text, delimiter);
    if (rows.length < 2) return [];

    const headersRaw = rows[0].map(h => String(h ?? '').trim().replace(/^\uFEFF/, ''));
    const headerIndex = {};
    headersRaw.forEach((h, idx) => {
        const nk = normalizeKey(h);
        if (nk && headerIndex[nk] === undefined) headerIndex[nk] = idx;
    });

    function getVal(values, key) {
        const idx = headerIndex[normalizeKey(key)];
        return idx === undefined ? '' : String(values[idx] ?? '').trim();
    }

    function getValContains(values, containsText) {
        const target = normalizeKey(containsText);
        for (const [nk, idx] of Object.entries(headerIndex)) {
            if (nk.includes(target)) return String(values[idx] ?? '').trim();
        }
        return '';
    }

    // مفاتيح مهمة
    const programs = {};

    for (let i = 1; i < rows.length; i++) {
        const values = rows[i];

        const prog = getVal(values, 'Major_aName');
        const degree = getVal(values, 'Degree_aName');
        const semester = getVal(values, 'Semester');
        if (!prog || !semester) continue;

        const key = `${prog}|${degree}`;
        if (!programs[key]) {
            programs[key] = {
                name: prog,
                degree: degree,
                dept: getVal(values, 'Dept_aName') || '',
                years: {}
            };
        }

        const year = String(semester).trim();
        if (!programs[key].years[year]) programs[key].years[year] = {};

        const gender = getVal(values, 'Gender_aName') || 'All';
        const joinSem = getVal(values, 'Join_Semester') || 'All';
        const filterKey = `${gender}_${joinSem}`;

        // عناوين عربية قد تختلف تشكيل/صياغة، لذلك نستخدم getVal + contains fallback
        const students = getVal(values, 'عدد المنتظمين') || getValContains(values, 'عددالمنتظمين');
        const graduates = getVal(values, 'عدد الخريجين') || getValContains(values, 'عددالخريجين');

        const courseEval =
            getVal(values, 'المتوسط العام لتقييم جودة المقررات') ||
            getValContains(values, 'تقييمجودةالمقررات') ||
            getVal(values, 'تقييم جودة المقررات');

        const expEval =
            getVal(values, 'المتوسط العام لتقييم خبرة البرنامج') ||
            getValContains(values, 'تقييمخبرةالبرنامج') ||
            getVal(values, 'تقييم خبرة البرنامج');

        const facultyTotal =
            getVal(values, 'العدد الإجمالي للأعضاء') ||
            getValContains(values, 'العدداجماليالاعضاء');

        const facultyPublished =
            getVal(values, 'عدد الأعضاء الذين نشروا بحثًا') ||
            getVal(values, 'عدد الأعضاء الذين نشروا بحثاً') ||
            getVal(values, 'عدد الأعضاء الذين نشروا بحثا') ||
            getValContains(values, 'عددالاعضاءالذيننشروابحث');

        const researchCount =
            getVal(values, 'عدد الأبحاث المنشورة') ||
            getValContains(values, 'عددالابحاثالمنشورة');

        const citations =
            getVal(values, 'إجمالي الاقتباسات') ||
            getValContains(values, 'اجماليالاقتباسات');

        programs[key].years[year][filterKey] = {
            students: students || '',
            graduates: graduates || '',
            course_eval: courseEval || '',
            experience_eval: expEval || '',
            faculty_total: facultyTotal || '',
            faculty_published: facultyPublished || '',
            research_count: researchCount || '',
            citations: citations || '',
        };
    }

    return Object.values(programs);
}

// ========================================
// تهيئة القوائم
// ========================================
function initializeAllSelects() {
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
}

function populateYearSelect(selectId, programIndex) {
    const select = document.getElementById(selectId);
    const program = programsData[programIndex];
    select.innerHTML = '<option value="">-- اختر --</option>';
    if (program?.years) {
        Object.keys(program.years).sort().forEach(y => {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = `سنة ${y}`;
            select.appendChild(opt);
        });
        select.disabled = false;
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
        opt.textContent = `سنة ${y}`;
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
    if (!program?.years?.[year]) return null;

    const yearData = program.years[year];
    let data = yearData["All_All"] || yearData["All_all"];
    if (!data) {
        for (const k of Object.keys(yearData)) {
            if (k.startsWith("All_")) { data = yearData[k]; break; }
        }
    }
    if (!data && Object.keys(yearData).length > 0) {
        data = yearData[Object.keys(yearData)[0]];
    }
    return { program, year, data };
}

function calculateIndicators(rawData) {
    if (!rawData?.data) return {};
    const d = rawData.data;
    const ind = {};

    const expEval = toNumber(d.experience_eval);
    const courseEval = toNumber(d.course_eval);

    ind.experience_eval = Number.isFinite(expEval) ? expEval : null;
    ind.course_eval = Number.isFinite(courseEval) ? courseEval : null;

    ind.graduation_rate = null;
    ind.retention_rate = null;
    ind.student_performance = null;
    ind.employment_rate = null;
    ind.employer_eval = null;

    const students = toNumber(d.students) || 0;
    const faculty = toNumber(d.faculty_total) || 0;
    const published = toNumber(d.faculty_published) || 0;
    const research = toNumber(d.research_count) || 0;
    const citations = toNumber(d.citations) || 0;

    ind.student_faculty_ratio = (faculty > 0 && students > 0) ? `1:${Math.round(students / faculty)}` : null;
    ind.publication_pct = (faculty > 0 && published > 0) ? ((published / faculty) * 100).toFixed(1) : null;
    ind.research_per_faculty = (faculty > 0 && research > 0) ? (research / faculty).toFixed(2) : null;
    ind.citations_per_faculty = (faculty > 0 && citations > 0) ? (citations / faculty).toFixed(1) : null;

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
    if (!progIdx || !year) return;

    const rawData = getData(parseInt(progIdx), year);
    if (!rawData) { alert('لا توجد بيانات'); return; }

    currentData = rawData;
    const indicators = calculateIndicators(rawData);

    // عرض المعلومات
    document.getElementById('degree-badge').textContent = rawData.program.degree;
    document.getElementById('program-name').textContent = rawData.program.name;
    document.getElementById('program-year').textContent = `السنة: ${year}`;

    // عرض المؤشرات
    const grid = document.getElementById('indicators-grid');
    grid.innerHTML = '';
    INDICATORS.forEach(ind => {
        if (ind.gradOnly && !['الماجستير', 'دكتوراه'].includes(rawData.program.degree)) return;
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

    // عرض البيانات الخام
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
    const compareType = document.querySelector('.compare-type-btn.active').dataset.type;
    let data1, data2, title1, title2;

    if (compareType === 'years') {
        const progIdx = document.getElementById('compare-program').value;
        const y1 = document.getElementById('compare-year1').value;
        const y2 = document.getElementById('compare-year2').value;
        if (!progIdx || !y1 || !y2) { alert('اختر جميع الحقول'); return; }

        data1 = getData(parseInt(progIdx), y1);
        data2 = getData(parseInt(progIdx), y2);
        title1 = `سنة ${y1}`;
        title2 = `سنة ${y2}`;
    } else {
        const year = document.getElementById('compare-common-year').value;
        const p1 = document.getElementById('compare-prog1').value;
        const p2 = document.getElementById('compare-prog2').value;
        if (!year || !p1 || !p2) { alert('اختر جميع الحقول'); return; }

        data1 = getData(parseInt(p1), year);
        data2 = getData(parseInt(p2), year);
        title1 = programsData[p1].name;
        title2 = programsData[p2].name;
    }

    if (!data1 || !data2) { alert('لا توجد بيانات كافية'); return; }

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
        if (ind.numeric && v1 !== null && v2 !== null && v1 !== '' && v2 !== '') {
            const diff = parseFloat(v1) - parseFloat(v2);
            if (Number.isFinite(diff)) {
                const sign = diff > 0 ? '+' : '';
                const cls = diff > 0 ? 'diff-positive' : diff < 0 ? 'diff-negative' : 'diff-neutral';
                diffHtml = `<span class="${cls}">${sign}${diff.toFixed(2)}</span>`;
            }
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
    doc.text('Year: ' + currentData.year, 105, 30, { align: 'center' });

    const indicators = calculateIndicators(currentData);
    const tableData = INDICATORS.map(ind => [ind.name, indicators[ind.key] || '-', ind.unit]);

    doc.autoTable({
        startY: 40,
        head: [['Indicator', 'Value', 'Unit']],
        body: tableData,
        styles: { font: 'helvetica', halign: 'center', fontSize: 9 },
        headStyles: { fillColor: [13, 79, 79] }
    });

    doc.save(`KPI-${currentData.program.name}-${currentData.year}.pdf`);
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
        [`السنة: ${currentData.year}`],
        [],
        ['المؤشر', 'القيمة', 'الوحدة'],
        ...INDICATORS.map(ind => [ind.name, indicators[ind.key] || '—', ind.unit])
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المؤشرات');
    XLSX.writeFile(wb, `مؤشرات-${currentData.program.name}-${currentData.year}.xlsx`);
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
    loadData();

    // التنقل بين العروض
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.view-section').forEach(v => v.classList.add('hidden'));
            btn.classList.add('active');
            document.getElementById(`${btn.dataset.view}-view`).classList.remove('hidden');
        });
    });

    // العرض الفردي
    document.getElementById('program-select').addEventListener('change', e => {
        if (e.target.value) {
            populateYearSelect('year-select', parseInt(e.target.value));
        } else {
            document.getElementById('year-select').disabled = true;
        }
        document.getElementById('show-btn').disabled = true;
        ['program-info', 'indicators-section', 'raw-data-section'].forEach(id =>
            document.getElementById(id).classList.add('hidden'));
    });

    document.getElementById('year-select').addEventListener('change', e => {
        document.getElementById('show-btn').disabled = !e.target.value || !document.getElementById('program-select').value;
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
            populateYearSelect('compare-year1', parseInt(e.target.value));
            populateYearSelect('compare-year2', parseInt(e.target.value));
        }
        updateCompareBtn();
    });

    ['compare-year1', 'compare-year2'].forEach(id => {
        document.getElementById(id).addEventListener('change', updateCompareBtn);
    });

    // مقارنة البرامج
    document.getElementById('compare-common-year').addEventListener('change', e => {
        if (e.target.value) {
            populateProgramsWithYear('compare-prog1', e.target.value);
            populateProgramsWithYear('compare-prog2', e.target.value);
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
