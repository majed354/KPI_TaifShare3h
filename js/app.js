/**
 * مؤشرات الأداء - كلية الشريعة والأنظمة
 * الإصدار 2.0 - محدث
 * 
 * المميزات:
 * - يدمج البيانات من جميع الصفوف (ذكور+إناث، سعوديين+غير سعوديين)
 * - يحفظ البيانات التفصيلية للعرض
 * - تحديث تلقائي كل 30 ثانية
 * - عرض الإحصائيات التفصيلية قبل المؤشرات
 */

let programsData = [];
let allRawData = []; // جميع البيانات الخام للعرض التفصيلي
let currentData = null;
let compareData = { item1: null, item2: null, title1: '', title2: '' };
let autoRefreshInterval = null;

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
// تحميل البيانات مع تحديث تلقائي
// ========================================
async function loadData(silent = false) {
    const statusDot = document.getElementById('status-dot');
    const statusText = document.getElementById('data-source-text');
    
    try {
        if (!silent) statusText.textContent = 'جاري التحميل...';
        
        // إضافة timestamp لتجنب التخزين المؤقت
        const timestamp = new Date().getTime();
        const response = await fetch(`data/data.csv?t=${timestamp}`);
        const csvText = await response.text();
        
        console.log('📄 تم تحميل الملف، الحجم:', csvText.length, 'حرف');
        
        const result = parseCSV(csvText);
        programsData = result.programs;
        allRawData = result.rawRows;
        
        const now = new Date().toLocaleTimeString('ar-SA');
        statusText.textContent = `✓ تم تحميل ${programsData.length} برنامج (${now})`;
        statusDot.classList.remove('error');
        statusDot.classList.add('connected');
        
        console.log('✅ البرامج:', programsData.map(p => p.name + ' (' + p.degree + ')'));
        
        if (!silent) initSelects();
        
    } catch (error) {
        console.error('❌ خطأ:', error);
        statusText.textContent = '✗ خطأ في التحميل';
        statusDot.classList.remove('connected');
        statusDot.classList.add('error');
    }
}

// تفعيل التحديث التلقائي
function startAutoRefresh(intervalSeconds = 30) {
    if (autoRefreshInterval) clearInterval(autoRefreshInterval);
    autoRefreshInterval = setInterval(() => {
        console.log('🔄 تحديث تلقائي...');
        loadData(true);
    }, intervalSeconds * 1000);
}

// ========================================
// تحليل CSV الشامل
// ========================================
function parseCSV(csvText) {
    console.log('🔄 بدء تحليل CSV...');
    
    // تنظيف وتقسيم
    const lines = csvText.trim().split('\n').map(l => l.replace(/\r/g, ''));
    if (lines.length < 2) return { programs: [], rawRows: [] };
    
    // تحديد الفاصل
    const firstLine = lines[0];
    const delimiter = firstLine.includes(';') ? ';' : ',';
    console.log('📌 الفاصل:', delimiter === ';' ? 'فاصلة منقوطة' : 'فاصلة');
    
    // استخراج الرؤوس
    const headers = lines[0].split(delimiter).map(h => h.trim().replace(/^\uFEFF/, ''));
    console.log('📋 الأعمدة:', headers);
    
    // إيجاد فهرس كل عمود (بحث مرن)
    const col = {};
    headers.forEach((h, i) => {
        const header = h.toLowerCase();
        const headerAr = h;
        
        // الأعمدة الأساسية
        if (header === 'dept_aname' || header.includes('dept')) col.dept = i;
        if (header === 'major_aname' || header.includes('major')) col.prog = i;
        if (header === 'degree_aname' || header.includes('degree')) col.degree = i;
        if (header === 'semester') col.semester = i;
        if (header === 'gender_aname' || header.includes('gender')) col.gender = i;
        if (header.includes('considered') || header.includes('nat')) col.nationality = i;
        if (header === 'join_semester' || header.includes('join')) col.join = i;
        
        // الأعمدة العربية
        if (headerAr.includes('المنتظمين')) col.students = i;
        if (headerAr.includes('الخريجين')) col.graduates = i;
        if (headerAr.includes('تقييم جودة المقررات') || headerAr.includes('جودة المقررات')) col.courseEval = i;
        if (headerAr.includes('خبرة البرنامج')) col.expEval = i;
        if (headerAr.includes('الإجمالي للأعضاء') || headerAr === 'العدد الإجمالي للأعضاء') col.facultyTotal = i;
        if (headerAr.includes('الأعضاء الذكور')) col.facultyMale = i;
        if (headerAr.includes('نشروا بحث')) col.facultyPub = i;
        if (headerAr.includes('الأعضاء الدكاترة') || headerAr.includes('الدكاترة')) col.facultyPhd = i;
        if (headerAr.includes('الأبحاث المنشورة')) col.research = i;
        if (headerAr.includes('الاقتباسات')) col.citations = i;
        if (headerAr.includes('الشعب الإجمالي') || headerAr === 'عدد الشعب الإجمالي') col.sectionsTotal = i;
        if (headerAr.includes('شعب الذكور')) col.sectionsMale = i;
    });
    
    console.log('🔢 فهرس الأعمدة:', col);
    
    // ═══════════════════════════════════════════════════════════
    // جمع البيانات الخام وتصنيفها
    // ═══════════════════════════════════════════════════════════
    const rawRows = [];
    const aggregated = {}; // للبيانات المجمعة
    
    for (let i = 1; i < lines.length; i++) {
        const vals = lines[i].split(delimiter).map(v => v.trim());
        if (vals.length < 5) continue;
        
        const prog = vals[col.prog] || '';
        const degree = vals[col.degree] || '';
        const semester = vals[col.semester] || '';
        const gender = vals[col.gender] || '';
        const nationality = vals[col.nationality] || '';
        const joinSem = vals[col.join] || '';
        
        // تجاهل الصفوف بدون برنامج
        if (!prog) continue;
        
        // استخراج القيم
        const rowData = {
            dept: vals[col.dept] || '',
            prog: prog,
            degree: degree,
            semester: shortYear(semester),
            gender: gender,
            nationality: nationality,
            joinSemester: joinSem,
            students: parseFloat(vals[col.students]) || 0,
            graduates: parseFloat(vals[col.graduates]) || 0,
            courseEval: parseFloat(vals[col.courseEval]) || 0,
            expEval: parseFloat(vals[col.expEval]) || 0,
            facultyTotal: parseFloat(vals[col.facultyTotal]) || 0,
            facultyMale: parseFloat(vals[col.facultyMale]) || 0,
            facultyPub: parseFloat(vals[col.facultyPub]) || 0,
            facultyPhd: parseFloat(vals[col.facultyPhd]) || 0,
            research: parseFloat(vals[col.research]) || 0,
            citations: parseFloat(vals[col.citations]) || 0,
            sectionsTotal: parseFloat(vals[col.sectionsTotal]) || 0,
            sectionsMale: parseFloat(vals[col.sectionsMale]) || 0,
        };
        
        rawRows.push(rowData);
        
        // ═══════════════════════════════════════════════════════════
        // تجميع البيانات لكل برنامج/سنة
        // ═══════════════════════════════════════════════════════════
        if (!rowData.semester) continue; // تجاهل صفوف بدون سنة
        
        const key = `${prog}|${degree}|${rowData.semester}`;
        
        if (!aggregated[key]) {
            aggregated[key] = {
                dept: rowData.dept,
                prog: prog,
                degree: degree,
                semester: rowData.semester,
                
                // البيانات الإجمالية
                students_total: 0,
                graduates_total: 0,
                courseEval: 0,
                expEval: 0,
                facultyTotal: 0,
                facultyMale: 0,
                facultyPub: 0,
                facultyPhd: 0,
                research: 0,
                citations: 0,
                sectionsTotal: 0,
                sectionsMale: 0,
                
                // البيانات التفصيلية
                details: {
                    students_male: 0,
                    students_female: 0,
                    students_saudi: 0,
                    students_nonSaudi: 0,
                    graduates_male: 0,
                    graduates_female: 0,
                    graduates_saudi: 0,
                    graduates_nonSaudi: 0,
                    
                    // للتتبع
                    has_gender_breakdown: false,
                    has_nationality_breakdown: false,
                }
            };
        }
        
        const agg = aggregated[key];
        
        // ═══════════════════════════════════════════════════════════
        // منطق الدمج الذكي
        // ═══════════════════════════════════════════════════════════
        
        // 1. البيانات الإجمالية (Gender=All, Nationality=All, Join_Semester=All)
        if (gender === 'All' && (nationality === 'All' || nationality === '') && joinSem === 'All') {
            // هذا صف إجمالي - نأخذ منه البيانات المباشرة
            if (rowData.students > 0) agg.students_total = rowData.students;
            if (rowData.graduates > 0) agg.graduates_total = rowData.graduates;
            if (rowData.courseEval > 0) agg.courseEval = rowData.courseEval;
            if (rowData.expEval > 0) agg.expEval = rowData.expEval;
            if (rowData.facultyTotal > 0) agg.facultyTotal = rowData.facultyTotal;
            if (rowData.facultyMale > 0) agg.facultyMale = rowData.facultyMale;
            if (rowData.facultyPub > 0) agg.facultyPub = rowData.facultyPub;
            if (rowData.facultyPhd > 0) agg.facultyPhd = rowData.facultyPhd;
            if (rowData.research > 0) agg.research = rowData.research;
            if (rowData.citations > 0) agg.citations = rowData.citations;
            if (rowData.sectionsTotal > 0) agg.sectionsTotal = rowData.sectionsTotal;
            if (rowData.sectionsMale > 0) agg.sectionsMale = rowData.sectionsMale;
        }
        
        // 2. بيانات حسب الجنس (Gender=ذكر أو أنثى)
        if (gender === 'ذكر' && (nationality === 'All' || nationality === '') && joinSem === 'All') {
            agg.details.students_male = rowData.students;
            agg.details.graduates_male = rowData.graduates;
            agg.details.has_gender_breakdown = true;
        }
        if (gender === 'أنثى' && (nationality === 'All' || nationality === '') && joinSem === 'All') {
            agg.details.students_female = rowData.students;
            agg.details.graduates_female = rowData.graduates;
            agg.details.has_gender_breakdown = true;
        }
        
        // 3. بيانات حسب الجنسية (Nationality=سعودي أو غير سعودي)
        if (gender === 'All' && nationality === 'سعودي' && joinSem === 'All') {
            agg.details.students_saudi = rowData.students;
            agg.details.graduates_saudi = rowData.graduates;
            agg.details.has_nationality_breakdown = true;
        }
        if (gender === 'All' && (nationality === 'غير سعودي' || nationality === 'Non-Saudi') && joinSem === 'All') {
            agg.details.students_nonSaudi = rowData.students;
            agg.details.graduates_nonSaudi = rowData.graduates;
            agg.details.has_nationality_breakdown = true;
        }
    }
    
    console.log('📊 السجلات المجمعة:', Object.keys(aggregated).length);
    
    // ═══════════════════════════════════════════════════════════
    // حساب الإجماليات من التفاصيل إذا لم تكن موجودة
    // ═══════════════════════════════════════════════════════════
    for (const key in aggregated) {
        const agg = aggregated[key];
        
        // إذا لم يكن هناك إجمالي للطلاب، نحسبه من الذكور والإناث
        if (agg.students_total === 0 && agg.details.has_gender_breakdown) {
            agg.students_total = agg.details.students_male + agg.details.students_female;
        }
        
        // إذا لم يكن هناك إجمالي للخريجين، نحسبه من الذكور والإناث
        if (agg.graduates_total === 0 && agg.details.has_gender_breakdown) {
            agg.graduates_total = agg.details.graduates_male + agg.details.graduates_female;
        }
        
        // حساب غير السعوديين إذا كان لدينا الإجمالي والسعوديين
        if (agg.details.students_saudi > 0 && agg.students_total > 0 && agg.details.students_nonSaudi === 0) {
            agg.details.students_nonSaudi = agg.students_total - agg.details.students_saudi;
            agg.details.has_nationality_breakdown = true;
        }
        
        // حساب الإناث إذا كان لدينا الإجمالي والذكور
        if (agg.details.students_male > 0 && agg.students_total > 0 && agg.details.students_female === 0) {
            agg.details.students_female = agg.students_total - agg.details.students_male;
            agg.details.has_gender_breakdown = true;
        }
    }
    
    // ═══════════════════════════════════════════════════════════
    // تحويل إلى هيكل البرامج
    // ═══════════════════════════════════════════════════════════
    const programs = {};
    
    for (const key in aggregated) {
        const d = aggregated[key];
        const progKey = `${d.prog}|${d.degree}`;
        
        if (!programs[progKey]) {
            programs[progKey] = {
                name: d.prog,
                degree: d.degree,
                dept: d.dept,
                years: {}
            };
        }
        
        programs[progKey].years[d.semester] = {
            // البيانات الأساسية
            students: d.students_total || '',
            graduates: d.graduates_total || '',
            course_eval: d.courseEval || '',
            experience_eval: d.expEval || '',
            faculty_total: d.facultyTotal || '',
            faculty_male: d.facultyMale || '',
            faculty_published: d.facultyPub || '',
            faculty_phd: d.facultyPhd || '',
            research_count: d.research || '',
            citations: d.citations || '',
            sections_total: d.sectionsTotal || '',
            sections_male: d.sectionsMale || '',
            
            // التفاصيل
            details: d.details
        };
    }
    
    const result = Object.values(programs);
    console.log('✅ البرامج النهائية:', result.length);
    
    return { programs: result, rawRows: rawRows };
}

// ========================================
// حساب المؤشرات من البيانات
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
    
    // المؤشرات المباشرة
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
    
    // نسبة النشر العلمي
    if (faculty > 0 && published > 0) {
        ind.publication_pct = ((published / faculty) * 100).toFixed(1);
    } else {
        ind.publication_pct = null;
    }
    
    // البحوث لكل عضو
    if (faculty > 0 && research > 0) {
        ind.research_per_faculty = (research / faculty).toFixed(2);
    } else {
        ind.research_per_faculty = null;
    }
    
    // الاقتباسات لكل عضو
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
    
    // ═══════════════════════════════════════
    // عرض الإحصائيات التفصيلية
    // ═══════════════════════════════════════
    const detailsSection = document.getElementById('details-section');
    const detailsGrid = document.getElementById('details-grid');
    
    if (detailsSection && detailsGrid) {
        detailsGrid.innerHTML = '';
        
        const data = raw.data;
        const details = data.details || {};
        
        // عدد الطلاب
        if (data.students) {
            detailsGrid.innerHTML += createDetailCard('👥', 'إجمالي الطلاب المنتظمين', data.students, 'طالب', 'primary');
        }
        
        // تفصيل الطلاب حسب الجنس
        if (details.has_gender_breakdown) {
            if (details.students_male > 0) {
                detailsGrid.innerHTML += createDetailCard('👨', 'الطلاب الذكور', details.students_male, 'طالب', 'blue');
            }
            if (details.students_female > 0) {
                detailsGrid.innerHTML += createDetailCard('👩', 'الطالبات الإناث', details.students_female, 'طالبة', 'pink');
            }
        }
        
        // تفصيل الطلاب حسب الجنسية
        if (details.has_nationality_breakdown) {
            if (details.students_saudi > 0) {
                detailsGrid.innerHTML += createDetailCard('🇸🇦', 'الطلاب السعوديون', details.students_saudi, 'طالب', 'green');
            }
            if (details.students_nonSaudi > 0) {
                detailsGrid.innerHTML += createDetailCard('🌍', 'الطلاب غير السعوديين', details.students_nonSaudi, 'طالب', 'orange');
            }
        }
        
        // عدد الخريجين
        if (data.graduates) {
            detailsGrid.innerHTML += createDetailCard('🎓', 'إجمالي الخريجين', data.graduates, 'خريج', 'success');
        }
        
        // تفصيل الخريجين حسب الجنس
        if (details.graduates_male > 0) {
            detailsGrid.innerHTML += createDetailCard('👨‍🎓', 'الخريجون الذكور', details.graduates_male, 'خريج', 'blue');
        }
        if (details.graduates_female > 0) {
            detailsGrid.innerHTML += createDetailCard('👩‍🎓', 'الخريجات الإناث', details.graduates_female, 'خريجة', 'pink');
        }
        
        // هيئة التدريس
        if (data.faculty_total) {
            detailsGrid.innerHTML += createDetailCard('👨‍🏫', 'إجمالي أعضاء هيئة التدريس', data.faculty_total, 'عضو', 'purple');
        }
        if (data.faculty_male) {
            detailsGrid.innerHTML += createDetailCard('👔', 'الأعضاء الذكور', data.faculty_male, 'عضو', 'blue');
        }
        if (data.faculty_phd) {
            detailsGrid.innerHTML += createDetailCard('📜', 'الأعضاء الدكاترة', data.faculty_phd, 'دكتور', 'indigo');
        }
        if (data.faculty_published) {
            detailsGrid.innerHTML += createDetailCard('📝', 'الأعضاء الناشرون', data.faculty_published, 'عضو', 'teal');
        }
        
        // الأبحاث
        if (data.research_count) {
            detailsGrid.innerHTML += createDetailCard('📚', 'الأبحاث المنشورة', data.research_count, 'بحث', 'amber');
        }
        if (data.citations) {
            detailsGrid.innerHTML += createDetailCard('📎', 'الاقتباسات', data.citations, 'اقتباس', 'cyan');
        }
        
        // الشعب
        if (data.sections_total) {
            detailsGrid.innerHTML += createDetailCard('🏛️', 'إجمالي الشعب', data.sections_total, 'شعبة', 'gray');
        }
        
        // التقييمات
        if (data.course_eval) {
            detailsGrid.innerHTML += createDetailCard('⭐', 'تقييم جودة المقررات', parseFloat(data.course_eval).toFixed(2), 'من 5', 'yellow');
        }
        if (data.experience_eval) {
            detailsGrid.innerHTML += createDetailCard('✨', 'تقييم خبرة البرنامج', parseFloat(data.experience_eval).toFixed(2), 'من 5', 'yellow');
        }
        
        detailsSection.classList.remove('hidden');
    }
    
    // ═══════════════════════════════════════
    // عرض المؤشرات
    // ═══════════════════════════════════════
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
    
    // ═══════════════════════════════════════
    // إخفاء قسم البيانات الخام القديم (تم استبداله بالتفاصيل)
    // ═══════════════════════════════════════
    const rawDataSection = document.getElementById('raw-data-section');
    if (rawDataSection) {
        rawDataSection.classList.add('hidden');
    }
    
    // إظهار الأقسام
    document.getElementById('program-info').classList.remove('hidden');
    document.getElementById('indicators-section').classList.remove('hidden');
    document.getElementById('program-info').scrollIntoView({ behavior: 'smooth' });
}

// إنشاء بطاقة تفصيلية
function createDetailCard(icon, label, value, unit, color) {
    return `
        <div class="detail-card detail-${color}">
            <div class="detail-icon">${icon}</div>
            <div class="detail-content">
                <div class="detail-label">${label}</div>
                <div class="detail-value">${value}</div>
                <div class="detail-unit">${unit}</div>
            </div>
        </div>
    `;
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
    const data = currentData.data;
    const details = data.details || {};
    
    const sheetData = [
        ['تقرير مؤشرات الأداء'],
        ['البرنامج: ' + currentData.program.name],
        ['السنة: ' + formatYear(currentData.year)],
        [],
        ['═══════════════════════════════════════'],
        ['الإحصائيات التفصيلية'],
        ['═══════════════════════════════════════'],
        [],
        ['البيان', 'القيمة'],
        ['إجمالي الطلاب', data.students || '—'],
        ['الطلاب الذكور', details.students_male || '—'],
        ['الطالبات الإناث', details.students_female || '—'],
        ['الطلاب السعوديون', details.students_saudi || '—'],
        ['الطلاب غير السعوديين', details.students_nonSaudi || '—'],
        [],
        ['إجمالي الخريجين', data.graduates || '—'],
        ['الخريجون الذكور', details.graduates_male || '—'],
        ['الخريجات الإناث', details.graduates_female || '—'],
        [],
        ['إجمالي أعضاء هيئة التدريس', data.faculty_total || '—'],
        ['الأعضاء الذكور', data.faculty_male || '—'],
        ['الأعضاء الدكاترة', data.faculty_phd || '—'],
        ['الأعضاء الناشرون', data.faculty_published || '—'],
        [],
        ['الأبحاث المنشورة', data.research_count || '—'],
        ['إجمالي الاقتباسات', data.citations || '—'],
        [],
        ['═══════════════════════════════════════'],
        ['المؤشرات المحسوبة'],
        ['═══════════════════════════════════════'],
        [],
        ['المؤشر', 'القيمة', 'الوحدة'],
        ...INDICATORS.map(i => [i.name, ind[i.key] || '—', i.unit])
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'التقرير الكامل');
    XLSX.writeFile(wb, `مؤشرات-${currentData.program.name}-${formatYear(currentData.year)}.xlsx`);
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
    startAutoRefresh(30); // تحديث كل 30 ثانية
    
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
        ['program-info', 'details-section', 'indicators-section', 'raw-data-section'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.classList.add('hidden');
        });
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

// ========================================
// للتصحيح
// ========================================
window.debugData = () => {
    console.log('📊 البرامج:', programsData);
    console.log('📊 البيانات الخام:', allRawData);
    console.log('📊 الحالي:', currentData);
};

window.refreshData = () => {
    loadData(false);
};
