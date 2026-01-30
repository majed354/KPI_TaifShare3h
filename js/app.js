/**
 * مؤشرات الأداء - كلية الشريعة والأنظمة
 * الإصدار 2.1 - يدعم الهيكل الجديد والقديم
 */

let programsData = [];
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
// تحويل السنة
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
// تحميل البيانات
// ========================================
async function loadData(silent = false) {
    const statusDot = document.getElementById('status-dot');
    const statusText = document.getElementById('data-source-text');
    
    try {
        if (!silent) statusText.textContent = 'جاري التحميل...';
        
        const timestamp = new Date().getTime();
        const response = await fetch(`data/data.csv?t=${timestamp}`);
        const csvText = await response.text();
        
        console.log('📄 تم تحميل الملف، الحجم:', csvText.length, 'حرف');
        
        programsData = parseCSV(csvText);
        
        const now = new Date().toLocaleTimeString('ar-SA');
        statusText.textContent = `✓ تم تحميل ${programsData.length} برنامج (${now})`;
        statusDot.classList.remove('error');
        statusDot.classList.add('connected');
        
        if (!silent) initSelects();
        
    } catch (error) {
        console.error('❌ خطأ:', error);
        statusText.textContent = '✗ خطأ في التحميل';
        statusDot.classList.remove('connected');
        statusDot.classList.add('error');
    }
}

function startAutoRefresh(intervalSeconds = 30) {
    if (autoRefreshInterval) clearInterval(autoRefreshInterval);
    autoRefreshInterval = setInterval(() => {
        console.log('🔄 تحديث تلقائي...');
        loadData(true);
    }, intervalSeconds * 1000);
}

// ========================================
// تحليل CSV (يدعم الهيكل الجديد والقديم)
// ========================================
function parseCSV(csvText) {
    console.log('🔄 بدء تحليل CSV...');
    
    const lines = csvText.trim().split('\n').map(l => l.replace(/\r/g, ''));
    if (lines.length < 2) return [];
    
    const firstLine = lines[0];
    const delimiter = firstLine.includes(';') ? ';' : ',';
    
    const headers = lines[0].split(delimiter).map(h => h.trim().replace(/^\uFEFF/, ''));
    console.log('📋 الأعمدة:', headers);
    
    // تحديد نوع الملف (جديد أو قديم)
    const isNewFormat = headers.includes('students_total') || headers.includes('students_male');
    console.log('📌 نوع الملف:', isNewFormat ? 'الهيكل الجديد' : 'الهيكل القديم');
    
    // إيجاد فهرس الأعمدة
    const col = {};
    headers.forEach((h, i) => {
        const header = h.toLowerCase();
        const headerAr = h;
        
        // الأعمدة الأساسية
        if (header.includes('dept')) col.dept = i;
        if (header.includes('major')) col.prog = i;
        if (header.includes('degree')) col.degree = i;
        if (header === 'semester') col.semester = i;
        if (header.includes('gender')) col.gender = i;
        if (header.includes('nat') || header.includes('considered')) col.nationality = i;
        if (header.includes('join')) col.join = i;
        
        // ═══════════════════════════════════════
        // أعمدة الهيكل الجديد (إنجليزي)
        // ═══════════════════════════════════════
        if (header === 'students_total') col.students_total = i;
        if (header === 'students_male') col.students_male = i;
        if (header === 'students_female') col.students_female = i;
        if (header === 'students_saudi') col.students_saudi = i;
        if (header === 'students_international') col.students_intl = i;
        if (header === 'students_new') col.students_new = i;
        if (header === 'students_retained') col.students_retained = i;
        if (header === 'graduates_total') col.graduates_total = i;
        if (header === 'graduates_ontime') col.graduates_ontime = i;
        if (header === 'sections_total') col.sections_total = i;
        if (header === 'sections_male') col.sections_male = i;
        if (header === 'sections_female') col.sections_female = i;
        if (header === 'faculty_total') col.faculty_total = i;
        if (header === 'faculty_phd') col.faculty_phd = i;
        if (header === 'faculty_male') col.faculty_male = i;
        if (header === 'faculty_female') col.faculty_female = i;
        if (header === 'faculty_published') col.faculty_published = i;
        if (header === 'research_count') col.research_count = i;
        if (header === 'citations') col.citations = i;
        if (header === 'eval_courses') col.eval_courses = i;
        if (header === 'eval_experience') col.eval_experience = i;
        if (header === 'eval_employers') col.eval_employers = i;
        if (header === 'performance_rate') col.performance_rate = i;
        if (header === 'employment_rate') col.employment_rate = i;
        
        // ═══════════════════════════════════════
        // أعمدة الهيكل القديم (عربي)
        // ═══════════════════════════════════════
        if (headerAr.includes('المنتظمين')) col.students_old = i;
        if (headerAr.includes('الخريجين')) col.graduates_old = i;
        if (headerAr.includes('تقييم جودة المقررات')) col.courseEval_old = i;
        if (headerAr.includes('خبرة البرنامج')) col.expEval_old = i;
        if (headerAr.includes('الإجمالي للأعضاء')) col.facultyTotal_old = i;
        if (headerAr.includes('الأعضاء الذكور')) col.facultyMale_old = i;
        if (headerAr.includes('نشروا بحث')) col.facultyPub_old = i;
        if (headerAr.includes('الدكاترة')) col.facultyPhd_old = i;
        if (headerAr.includes('الأبحاث المنشورة')) col.research_old = i;
        if (headerAr.includes('الاقتباسات')) col.citations_old = i;
        if (headerAr.includes('الشعب الإجمالي')) col.sectionsTotal_old = i;
        if (headerAr.includes('شعب الذكور')) col.sectionsMale_old = i;
    });
    
    console.log('🔢 فهرس الأعمدة:', col);
    
    // ═══════════════════════════════════════════════════════════
    // معالجة الصفوف
    // ═══════════════════════════════════════════════════════════
    const aggregated = {};
    
    for (let i = 1; i < lines.length; i++) {
        const vals = lines[i].split(delimiter).map(v => v.trim());
        if (vals.length < 5) continue;
        
        const prog = vals[col.prog] || '';
        const degree = vals[col.degree] || '';
        const semester = vals[col.semester] || '';
        
        if (!prog || !semester) continue;
        
        const key = `${prog}|${degree}|${shortYear(semester)}`;
        
        // للهيكل الجديد: صف واحد لكل برنامج/سنة
        if (isNewFormat) {
            aggregated[key] = {
                dept: vals[col.dept] || '',
                prog: prog,
                degree: degree,
                semester: shortYear(semester),
                
                // بيانات الطلاب
                students_total: parseFloat(vals[col.students_total]) || 0,
                students_male: parseFloat(vals[col.students_male]) || 0,
                students_female: parseFloat(vals[col.students_female]) || 0,
                students_saudi: parseFloat(vals[col.students_saudi]) || 0,
                students_intl: parseFloat(vals[col.students_intl]) || 0,
                students_new: parseFloat(vals[col.students_new]) || 0,
                students_retained: parseFloat(vals[col.students_retained]) || 0,
                
                // بيانات الخريجين
                graduates_total: parseFloat(vals[col.graduates_total]) || 0,
                graduates_ontime: parseFloat(vals[col.graduates_ontime]) || 0,
                
                // بيانات الشعب
                sections_total: parseFloat(vals[col.sections_total]) || 0,
                sections_male: parseFloat(vals[col.sections_male]) || 0,
                sections_female: parseFloat(vals[col.sections_female]) || 0,
                
                // بيانات هيئة التدريس
                faculty_total: parseFloat(vals[col.faculty_total]) || 0,
                faculty_phd: parseFloat(vals[col.faculty_phd]) || 0,
                faculty_male: parseFloat(vals[col.faculty_male]) || 0,
                faculty_female: parseFloat(vals[col.faculty_female]) || 0,
                faculty_published: parseFloat(vals[col.faculty_published]) || 0,
                
                // بيانات البحث العلمي
                research_count: parseFloat(vals[col.research_count]) || 0,
                citations: parseFloat(vals[col.citations]) || 0,
                
                // بيانات التقييم
                eval_courses: parseFloat(vals[col.eval_courses]) || 0,
                eval_experience: parseFloat(vals[col.eval_experience]) || 0,
                eval_employers: parseFloat(vals[col.eval_employers]) || 0,
                performance_rate: parseFloat(vals[col.performance_rate]) || 0,
                employment_rate: parseFloat(vals[col.employment_rate]) || 0,
            };
        } else {
            // للهيكل القديم: دمج الصفوف المتعددة
            const gender = vals[col.gender] || '';
            const nationality = vals[col.nationality] || '';
            const joinSem = vals[col.join] || '';
            
            if (!aggregated[key]) {
                aggregated[key] = {
                    dept: vals[col.dept] || '',
                    prog: prog,
                    degree: degree,
                    semester: shortYear(semester),
                    students_total: 0, students_male: 0, students_female: 0,
                    students_saudi: 0, students_intl: 0, students_new: 0, students_retained: 0,
                    graduates_total: 0, graduates_ontime: 0,
                    sections_total: 0, sections_male: 0, sections_female: 0,
                    faculty_total: 0, faculty_phd: 0, faculty_male: 0, faculty_female: 0, faculty_published: 0,
                    research_count: 0, citations: 0,
                    eval_courses: 0, eval_experience: 0, eval_employers: 0,
                    performance_rate: 0, employment_rate: 0,
                };
            }
            
            const agg = aggregated[key];
            
            // دمج البيانات من الصفوف المتعددة
            if (gender === 'All' && (nationality === 'All' || nationality === '') && joinSem === 'All') {
                if (col.students_old !== undefined && vals[col.students_old]) 
                    agg.students_total = parseFloat(vals[col.students_old]) || 0;
                if (col.graduates_old !== undefined && vals[col.graduates_old]) 
                    agg.graduates_total = parseFloat(vals[col.graduates_old]) || 0;
                if (col.courseEval_old !== undefined && vals[col.courseEval_old]) 
                    agg.eval_courses = parseFloat(vals[col.courseEval_old]) || 0;
                if (col.expEval_old !== undefined && vals[col.expEval_old]) 
                    agg.eval_experience = parseFloat(vals[col.expEval_old]) || 0;
                if (col.facultyTotal_old !== undefined && vals[col.facultyTotal_old]) 
                    agg.faculty_total = parseFloat(vals[col.facultyTotal_old]) || 0;
                if (col.facultyMale_old !== undefined && vals[col.facultyMale_old]) 
                    agg.faculty_male = parseFloat(vals[col.facultyMale_old]) || 0;
                if (col.facultyPhd_old !== undefined && vals[col.facultyPhd_old]) 
                    agg.faculty_phd = parseFloat(vals[col.facultyPhd_old]) || 0;
                if (col.facultyPub_old !== undefined && vals[col.facultyPub_old]) 
                    agg.faculty_published = parseFloat(vals[col.facultyPub_old]) || 0;
                if (col.research_old !== undefined && vals[col.research_old]) 
                    agg.research_count = parseFloat(vals[col.research_old]) || 0;
                if (col.citations_old !== undefined && vals[col.citations_old]) 
                    agg.citations = parseFloat(vals[col.citations_old]) || 0;
                if (col.sectionsTotal_old !== undefined && vals[col.sectionsTotal_old]) 
                    agg.sections_total = parseFloat(vals[col.sectionsTotal_old]) || 0;
                if (col.sectionsMale_old !== undefined && vals[col.sectionsMale_old]) 
                    agg.sections_male = parseFloat(vals[col.sectionsMale_old]) || 0;
            }
            
            // بيانات الذكور
            if (gender === 'ذكر' && (nationality === 'All' || nationality === '') && joinSem === 'All') {
                if (col.students_old !== undefined && vals[col.students_old])
                    agg.students_male = parseFloat(vals[col.students_old]) || 0;
            }
            
            // بيانات السعوديين
            if (gender === 'All' && nationality === 'سعودي' && joinSem === 'All') {
                if (col.students_old !== undefined && vals[col.students_old])
                    agg.students_saudi = parseFloat(vals[col.students_old]) || 0;
            }
        }
    }
    
    // ═══════════════════════════════════════════════════════════
    // حساب القيم المشتقة
    // ═══════════════════════════════════════════════════════════
    for (const key in aggregated) {
        const agg = aggregated[key];
        
        // حساب الإناث
        if (agg.students_female === 0 && agg.students_total > 0 && agg.students_male > 0) {
            agg.students_female = agg.students_total - agg.students_male;
        }
        
        // حساب الدوليين
        if (agg.students_intl === 0 && agg.students_total > 0 && agg.students_saudi > 0) {
            agg.students_intl = agg.students_total - agg.students_saudi;
        }
        
        // حساب شعب الإناث
        if (agg.sections_female === 0 && agg.sections_total > 0 && agg.sections_male > 0) {
            agg.sections_female = agg.sections_total - agg.sections_male;
        }
        
        // حساب أعضاء هيئة التدريس الإناث
        if (agg.faculty_female === 0 && agg.faculty_total > 0 && agg.faculty_male > 0) {
            agg.faculty_female = agg.faculty_total - agg.faculty_male;
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
        
        programs[progKey].years[d.semester] = d;
    }
    
    return Object.values(programs);
}

// ========================================
// حساب المؤشرات
// ========================================
function calculateIndicators(raw) {
    if (!raw?.data) return {};
    
    const d = raw.data;
    const ind = {};
    
    // المؤشرات المباشرة
    ind.experience_eval = d.eval_experience ? d.eval_experience.toFixed(2) : null;
    ind.course_eval = d.eval_courses ? d.eval_courses.toFixed(2) : null;
    ind.employer_eval = d.eval_employers ? d.eval_employers.toFixed(2) : null;
    ind.student_performance = d.performance_rate ? d.performance_rate.toFixed(1) : null;
    ind.employment_rate = d.employment_rate ? d.employment_rate.toFixed(1) : null;
    
    // معدل التخرج بالوقت المحدد
    if (d.graduates_total > 0 && d.graduates_ontime > 0) {
        ind.graduation_rate = ((d.graduates_ontime / d.graduates_total) * 100).toFixed(1);
    } else {
        ind.graduation_rate = null;
    }
    
    // معدل الاستبقاء (يحتاج بيانات السنة السابقة - غير متوفر حالياً)
    ind.retention_rate = null;
    
    // نسبة الطلاب/هيئة التدريس
    if (d.faculty_total > 0 && d.students_total > 0) {
        const ratio = Math.round(d.students_total / d.faculty_total);
        ind.student_faculty_ratio = `1:${ratio}`;
    } else {
        ind.student_faculty_ratio = null;
    }
    
    // نسبة النشر العلمي
    if (d.faculty_total > 0 && d.faculty_published > 0) {
        ind.publication_pct = ((d.faculty_published / d.faculty_total) * 100).toFixed(1);
    } else {
        ind.publication_pct = null;
    }
    
    // البحوث لكل عضو
    if (d.faculty_total > 0 && d.research_count > 0) {
        ind.research_per_faculty = (d.research_count / d.faculty_total).toFixed(2);
    } else {
        ind.research_per_faculty = null;
    }
    
    // الاقتباسات لكل عضو
    if (d.faculty_total > 0 && d.citations > 0) {
        ind.citations_per_faculty = (d.citations / d.faculty_total).toFixed(1);
    } else {
        ind.citations_per_faculty = null;
    }
    
    // المؤشرات غير المتوفرة
    ind.student_publication = null;
    ind.patents = null;
    
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
    const d = raw.data;
    
    // عرض المعلومات
    document.getElementById('degree-badge').textContent = raw.program.degree;
    document.getElementById('program-name').textContent = raw.program.name;
    document.getElementById('program-year').textContent = 'السنة: ' + formatYear(year);
    
    // ═══════════════════════════════════════
    // عرض الإحصائيات التفصيلية
    // ═══════════════════════════════════════
    const detailsGrid = document.getElementById('details-grid');
    detailsGrid.innerHTML = '';
    
    // === بيانات الطلاب ===
    if (d.students_total > 0) {
        detailsGrid.innerHTML += createDetailCard('👥', 'إجمالي الطلاب المنتظمين', d.students_total, 'طالب', 'primary');
    }
    if (d.students_male > 0) {
        detailsGrid.innerHTML += createDetailCard('👨', 'الطلاب الذكور', d.students_male, 'طالب', 'blue');
    }
    if (d.students_female > 0) {
        detailsGrid.innerHTML += createDetailCard('👩', 'الطالبات الإناث', d.students_female, 'طالبة', 'pink');
    }
    if (d.students_saudi > 0) {
        detailsGrid.innerHTML += createDetailCard('🇸🇦', 'الطلاب السعوديون', d.students_saudi, 'طالب', 'green');
    }
    if (d.students_intl > 0) {
        detailsGrid.innerHTML += createDetailCard('🌍', 'الطلاب الدوليون', d.students_intl, 'طالب', 'orange');
    }
    if (d.students_new > 0) {
        detailsGrid.innerHTML += createDetailCard('🆕', 'الطلاب المستجدون', d.students_new, 'طالب', 'cyan');
    }
    
    // === بيانات الخريجين ===
    if (d.graduates_total > 0) {
        detailsGrid.innerHTML += createDetailCard('🎓', 'إجمالي الخريجين', d.graduates_total, 'خريج', 'success');
    }
    if (d.graduates_ontime > 0) {
        detailsGrid.innerHTML += createDetailCard('⏱️', 'الخريجين بالوقت المحدد', d.graduates_ontime, 'خريج', 'teal');
    }
    
    // === بيانات الشعب ===
    if (d.sections_total > 0) {
        detailsGrid.innerHTML += createDetailCard('🏛️', 'إجمالي الشعب', d.sections_total, 'شعبة', 'gray');
        
        // متوسط الطلاب/شعبة
        const avgPerSection = (d.students_total / d.sections_total).toFixed(1);
        detailsGrid.innerHTML += createDetailCard('📊', 'متوسط الطلاب/شعبة', avgPerSection, 'طالب', 'indigo');
    }
    if (d.sections_male > 0) {
        detailsGrid.innerHTML += createDetailCard('🏛️', 'شعب الذكور', d.sections_male, 'شعبة', 'blue');
    }
    if (d.sections_female > 0) {
        detailsGrid.innerHTML += createDetailCard('🏛️', 'شعب الإناث', d.sections_female, 'شعبة', 'pink');
    }
    
    // === بيانات هيئة التدريس ===
    if (d.faculty_total > 0) {
        detailsGrid.innerHTML += createDetailCard('👨‍🏫', 'إجمالي أعضاء هيئة التدريس', d.faculty_total, 'عضو', 'purple');
    }
    if (d.faculty_phd > 0) {
        detailsGrid.innerHTML += createDetailCard('📜', 'الأعضاء الدكاترة', d.faculty_phd, 'دكتور', 'indigo');
    }
    if (d.faculty_male > 0) {
        detailsGrid.innerHTML += createDetailCard('👔', 'الأعضاء الذكور', d.faculty_male, 'عضو', 'blue');
    }
    if (d.faculty_female > 0) {
        detailsGrid.innerHTML += createDetailCard('👩‍🏫', 'الأعضاء الإناث', d.faculty_female, 'عضو', 'pink');
    }
    if (d.faculty_published > 0) {
        detailsGrid.innerHTML += createDetailCard('📝', 'الأعضاء الناشرون', d.faculty_published, 'عضو', 'teal');
    }
    
    // === بيانات البحث العلمي ===
    if (d.research_count > 0) {
        detailsGrid.innerHTML += createDetailCard('📚', 'الأبحاث المنشورة', d.research_count, 'بحث', 'amber');
    }
    if (d.citations > 0) {
        detailsGrid.innerHTML += createDetailCard('📎', 'الاقتباسات', d.citations, 'اقتباس', 'cyan');
    }
    
    // === بيانات التقييم ===
    if (d.eval_courses > 0) {
        detailsGrid.innerHTML += createDetailCard('⭐', 'تقييم جودة المقررات', d.eval_courses.toFixed(2), 'من 5', 'yellow');
    }
    if (d.eval_experience > 0) {
        detailsGrid.innerHTML += createDetailCard('✨', 'تقييم خبرة البرنامج', d.eval_experience.toFixed(2), 'من 5', 'yellow');
    }
    if (d.eval_employers > 0) {
        detailsGrid.innerHTML += createDetailCard('🏢', 'تقويم جهات التوظيف', d.eval_employers.toFixed(2), 'من 5', 'yellow');
    }
    if (d.performance_rate > 0) {
        detailsGrid.innerHTML += createDetailCard('📈', 'مستوى أداء الطالب', d.performance_rate + '%', '', 'success');
    }
    if (d.employment_rate > 0) {
        detailsGrid.innerHTML += createDetailCard('💼', 'نسبة توظيف الخريجين', d.employment_rate + '%', '', 'success');
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
    
    // إظهار الأقسام
    document.getElementById('program-info').classList.remove('hidden');
    document.getElementById('details-section').classList.remove('hidden');
    document.getElementById('indicators-section').classList.remove('hidden');
    document.getElementById('program-info').scrollIntoView({ behavior: 'smooth' });
}

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
    const d = currentData.data;
    
    const sheetData = [
        ['تقرير مؤشرات الأداء'],
        ['البرنامج: ' + currentData.program.name],
        ['السنة: ' + formatYear(currentData.year)],
        [],
        ['═══════════════════════════════════════════════════════════'],
        ['الإحصائيات التفصيلية'],
        ['═══════════════════════════════════════════════════════════'],
        [],
        ['البيان', 'القيمة'],
        ['--- بيانات الطلاب ---', ''],
        ['إجمالي الطلاب المنتظمين', d.students_total || '—'],
        ['الطلاب الذكور', d.students_male || '—'],
        ['الطالبات الإناث', d.students_female || '—'],
        ['الطلاب السعوديون', d.students_saudi || '—'],
        ['الطلاب الدوليون', d.students_intl || '—'],
        ['الطلاب المستجدون', d.students_new || '—'],
        [],
        ['--- بيانات الخريجين ---', ''],
        ['إجمالي الخريجين', d.graduates_total || '—'],
        ['الخريجين بالوقت المحدد', d.graduates_ontime || '—'],
        [],
        ['--- بيانات الشعب ---', ''],
        ['إجمالي الشعب', d.sections_total || '—'],
        ['شعب الذكور', d.sections_male || '—'],
        ['شعب الإناث', d.sections_female || '—'],
        [],
        ['--- بيانات هيئة التدريس ---', ''],
        ['إجمالي أعضاء هيئة التدريس', d.faculty_total || '—'],
        ['الأعضاء الدكاترة', d.faculty_phd || '—'],
        ['الأعضاء الذكور', d.faculty_male || '—'],
        ['الأعضاء الإناث', d.faculty_female || '—'],
        ['الأعضاء الناشرون', d.faculty_published || '—'],
        [],
        ['--- بيانات البحث العلمي ---', ''],
        ['الأبحاث المنشورة', d.research_count || '—'],
        ['إجمالي الاقتباسات', d.citations || '—'],
        [],
        ['--- بيانات التقييم ---', ''],
        ['تقييم جودة المقررات', d.eval_courses || '—'],
        ['تقييم خبرة البرنامج', d.eval_experience || '—'],
        ['تقويم جهات التوظيف', d.eval_employers || '—'],
        ['مستوى أداء الطالب', d.performance_rate || '—'],
        ['نسبة توظيف الخريجين', d.employment_rate || '—'],
        [],
        ['═══════════════════════════════════════════════════════════'],
        ['المؤشرات المحسوبة'],
        ['═══════════════════════════════════════════════════════════'],
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
    startAutoRefresh(30);
    
    // التنقل
    document.querySelectorAll('.nav-btn[data-view]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.nav-btn[data-view]').forEach(b => b.classList.remove('active'));
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
        ['program-info', 'details-section', 'indicators-section'].forEach(id => {
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

// للتصحيح
window.debugData = () => console.log('📊 البرامج:', programsData, '📊 الحالي:', currentData);
window.refreshData = () => loadData(false);
