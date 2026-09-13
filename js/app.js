/**
 * مؤشرات الأداء - كلية الشريعة والأنظمة
 * الإصدار 3.0 - لوحة معلومات احترافية
 */

// ========================================
// الثوابت
// ========================================
const DISPLAY_YEARS = [39, 40, 41, 42, 44, 45, 46, 47];
const CHART_COLORS = [
    '#0d8e8e','#c9a227','#3b82f6','#ef4444','#10b981',
    '#8b5cf6','#ec4899','#f97316','#06b6d4','#6366f1',
    '#22c55e','#eab308','#14b8a6','#e11d48','#7c3aed','#0ea5e9'
];
const KPI_CONFIG = (typeof window !== 'undefined' && window.KPI_CONFIG) ? window.KPI_CONFIG : {};
const GRADUATES_SURVEY_SHEET_URL = String(KPI_CONFIG.graduatesSurveySheetUrl || '').trim();
const GRADUATES_SURVEY_BACHELOR_SHEET_URL = String(KPI_CONFIG.graduatesSurveyBachelorSheetUrl || '').trim();
const GRADUATES_SURVEY_POSTGRAD_SHEET_URL = String(KPI_CONFIG.graduatesSurveyPostgradSheetUrl || '').trim();
const SHARI3AH_SURVEYS_DATA_URL = String(
    KPI_CONFIG.shari3ahSurveysDataUrl || 'https://raw.githubusercontent.com/majed354/Shari3ahSurveys/main/js/surveys-data.js'
).trim();
const ACTIVITIES_RAW_BASE = 'https://raw.githubusercontent.com/majed354/faculty-activities/main/data';
const ALL_BRANCH_FILTER_VALUE = '__all__';
const ALL_BRANCH_FILTER_LABEL = 'الكل';
const RESEARCH_KPI_EXCLUDED_RANKS = new Set(['معيد', 'محاضر', 'متعاون', 'مدرس']);
const GRADUATE_PROGRAM_ALIASES = {
    // توحيد الاختلافات الإملائية فقط دون دمج برامج مختلفة.
    'القران وعلومه': 'القرآن وعلومه',
    'الدراسات القرانيه': 'الدراسات القرآنية',
    'الانظمة': 'الأنظمة',
};
const SHARI3AH_SURVEYS_PROGRAM_ID_MAP = {
    'الأنظمة|بكالوريوس': 'p01',
    'الدراسات الإسلامية|بكالوريوس': 'p02',
    'الشريعة|بكالوريوس': 'p03',
    'القرآن وعلومه|بكالوريوس': 'p04',
    'القراءات|بكالوريوس': 'p05',
    'القانون|الماجستير': 'p06',
    'العقيدة|الماجستير': 'p07',
    'أصول الفقه|الماجستير': 'p08',
    'الفقه|الماجستير': 'p09',
    'الدراسات القرآنية المعاصرة|الماجستير': 'p10',
    'القراءات|الماجستير': 'p11',
    'أصول الفقه|دكتوراه': 'p12',
    'الفقه|دكتوراه': 'p13',
    'الدراسات القرآنية|دكتوراه': 'p14',
    'القراءات|دكتوراه': 'p15',
};
const SHARI3AH_PROGRAM_EVAL_SURVEY_TITLE = 'استبانة تقويم برنامج';
const INDICATORS_UG = [
    { id:1,  code:'KPI-1',  name:"تقويم الطلاب لجودة خبرات التعلم في البرنامج", unit:"درجة", key:"experience_eval", numeric:true },
    { id:2,  code:'KPI-2',  name:"تقييم الطلاب لجودة المقررات", unit:"درجة", key:"course_eval", numeric:true },
    { id:3,  code:'KPI-3',  name:"معدّل التخرج بالوقت المحدد", unit:"%", key:"graduation_rate", numeric:true },
    { id:4,  code:'KPI-4',  name:"معدّل استبقاء طلاب السنة الأولى", unit:"%", key:"retention_rate", numeric:true },
    { id:5,  code:'KPI-5',  name:"مستوى أداء الطالب", unit:"%", key:"student_performance", numeric:true },
    { id:6,  code:'KPI-6',  name:"توظيف الخريجين أو التحاقهم بالدراسات العليا", unit:"%", key:"employment_rate", numeric:true },
    { id:7,  code:'KPI-7',  name:"تقويم جهات التوظيف لكفاءة خريجي البرنامج", unit:"درجة", key:"employer_eval", numeric:true },
    { id:8,  code:'KPI-8',  name:"نسبة الطلاب إلى أعضاء هيئة التدريس", unit:"نسبة", key:"student_faculty_ratio" },
    { id:9,  code:'KPI-9',  name:"النسبة المئوية للنشر العلمي لأعضاء هيئة التدريس", unit:"%", key:"publication_pct", numeric:true },
    { id:10, code:'KPI-10', name:"معدل البحوث المنشورة لكل عضو هيئة تدريس", unit:"بحث", key:"research_per_faculty", numeric:true },
    { id:11, code:'KPI-11', name:"معدل الاقتباسات في المجلات المحكمة لكل عضو هيئة تدريس", unit:"اقتباس", key:"citations_per_faculty", numeric:true },
    { id:12, code:'KPI-12', name:"نسبة النشر العلمي للطلاب", unit:"%", key:"student_publication", numeric:true, gradOnly:true },
    { id:13, code:'KPI-13', name:"عدد براءات الاختراع والابتكار وجوائز التميز", unit:"براءة", key:"patents", numeric:true, gradOnly:true },
];
const INDICATORS_PG = [
    { id:1,  code:'KPI-PG-1',  name:"تقويم الطلاب لجودة خبرات التعلم في البرنامج", unit:"درجة", key:"experience_eval", numeric:true },
    { id:2,  code:'KPI-PG-2',  name:"تقييم الطلاب لجودة المقررات", unit:"درجة", key:"course_eval", numeric:true },
    { id:3,  code:'KPI-PG-3',  name:"تقييم الطلاب لجودة الإشراف العلمي", unit:"درجة", key:"supervision_eval", numeric:true },
    { id:4,  code:'KPI-PG-4',  name:"متوسط المدة التي يتخرج فيها الطالب", unit:"سنة", key:"avg_time_to_graduate", numeric:true },
    { id:5,  code:'KPI-PG-5',  name:"معدل تسرب الطلاب من البرنامج", unit:"%", key:"dropout_rate", numeric:true },
    { id:6,  code:'KPI-PG-6',  name:"تقويم جهات التوظيف لكفاءة خريجي البرنامج", unit:"درجة", key:"employer_eval", numeric:true },
    { id:7,  code:'KPI-PG-7',  name:"رضا الطلاب عن الخدمات المقدمة", unit:"درجة", key:"services_satisfaction", numeric:true },
    { id:8,  code:'KPI-PG-8',  name:"نسبة الطلاب إلى أعضاء هيئة التدريس", unit:"نسبة", key:"student_faculty_ratio" },
    { id:9,  code:'KPI-PG-9',  name:"النسبة المئوية للنشر العلمي لأعضاء هيئة التدريس", unit:"%", key:"publication_pct", numeric:true },
    { id:10, code:'KPI-PG-10', name:"معدل البحوث المنشورة لكل عضو هيئة تدريس", unit:"بحث", key:"research_per_faculty", numeric:true },
    { id:11, code:'KPI-PG-11', name:"معدل الاقتباسات في المجلات المحكمة لكل عضو هيئة تدريس", unit:"اقتباس", key:"citations_per_faculty", numeric:true },
    { id:12, code:'KPI-PG-12', name:"نسبة النشر العلمي للطلاب", unit:"%", key:"student_publication", numeric:true },
    { id:13, code:'KPI-PG-13', name:"عدد براءات الاختراع والابتكار وجوائز التميز", unit:"براءة", key:"patents", numeric:true },
];
const BACHELOR_DEGREES = new Set(['بكالوريوس']);
const POSTGRAD_DEGREES = new Set(['الماجستير', 'دكتوراه']);
const SURVEY_DEPT_FALLBACK_RULES = {
    // قواعد الأعمال المعتمدة:
    // - الشريعة والقراءات: تعدد برامج داخل نفس الدرجة => منع التعويض على مستوى القسم.
    // - الأنظمة والدراسات الإسلامية (الثقافة الإسلامية): برنامج واحد لكل درجة => يسمح بالتعويض عند غياب اسم البرنامج.
    'الشريعة': 'disabled',
    'القراءات': 'disabled',
    'الأنظمة': 'single_program_per_degree',
    'الدراسات الإسلامية': 'single_program_per_degree',
};
const GRADUATE_SAMPLE_METRICS_BY_DEGREE = {
    // استبانة الخريجين مصدر تكميلي لمؤشرات لا تغطيها الاستطلاعات الأصلية الحالية.
    'بكالوريوس': new Set(['performance_rate', 'employment_rate', 'eval_employers']),
    'الماجستير': new Set(['eval_supervision', 'eval_services', 'eval_employers']),
    'دكتوراه': new Set(['eval_supervision', 'eval_services', 'eval_employers']),
};
const STUDENT_CATEGORY_DEFS = [
    {
        id: 'excellent',
        label: 'المتفوقون',
        shortLabel: 'متفوقون',
        rangeLabel: 'ممتاز: 3.50 إلى 4.00',
        accent: '#0d8e8e',
        matches: gpa => gpa >= 3.5
    },
    {
        id: 'diligent',
        label: 'المجتهدون',
        shortLabel: 'مجتهدون',
        rangeLabel: 'جيد جدًا مرتفع: 3.25 إلى أقل من 3.50',
        accent: '#c9a227',
        matches: gpa => gpa >= 3.25 && gpa < 3.5
    },
    {
        id: 'weak',
        label: 'الضعفاء',
        shortLabel: 'ضعفاء',
        rangeLabel: 'أكبر من 1.00 وأقل من 2.00',
        accent: '#f97316',
        matches: gpa => gpa > 1 && gpa < 2
    },
    {
        id: 'struggling',
        label: 'المتعثرون',
        shortLabel: 'متعثرون',
        rangeLabel: '1.00 فأقل',
        accent: '#dc2626',
        matches: gpa => gpa <= 1
    },
];
const STUDENT_CATEGORY_DEFAULT_IDS = STUDENT_CATEGORY_DEFS.map(category => category.id);

// ========================================
// البيانات
// ========================================
let allRows = [];      // raw rows from CSV
let programs = [];     // organized {name, degree, dept, years:{y: data}}
let charts = {};       // Chart instances
let currentProg = null;// for export
let compareThirdEnabled = false;
let gradData = [];     // graduate records
let facultyRosterData = []; // faculty roster for 1445 through 1448
let entrantData = [];  // accepted new-entrant records for Islamic Studies branches
let islamicBranchData = null;
let ncData = [];       // non-completer records
let studentDetailData = [];
let appBootstrapped = false;
let appBootPromise = null;
let loginMembersById = null;
let analyticsChart = null;
let currentAnalyticsReport = null;
let studentCategoryChart = null;
let currentStudentCategoryReport = null;
let teachingProgramSupportByKey = {};
let programFacultyFteByBranchKey = {};
let researchActivitySupportByYearDept = {};
let availableFacultyBranches = [];

const SUPPORTED_KPI_DEGREES = new Set(['بكالوريوس','الماجستير','دكتوراه']);
const RANK_ALLOWED_DEGREES = {
    'معيد': new Set(['بكالوريوس']),
    'محاضر': new Set(['بكالوريوس']),
    'مدرس': new Set(['بكالوريوس']),
    'أستاذ مساعد': new Set(['بكالوريوس','الماجستير']),
    'أستاذ مشارك': new Set(['بكالوريوس','الماجستير','دكتوراه']),
    'أستاذ': new Set(['بكالوريوس','الماجستير','دكتوراه']),
    // المتعاون ليس ضمن القاعدة الرسمية في السؤال، نسمح له بكل الدرجات كحل عملي
    'متعاون': new Set(['بكالوريوس','الماجستير','دكتوراه']),
};
const RANK_BASE_FTE = {
    'متعاون': 0.5,
};

// ========================================
// المساعدات
// ========================================
function fmtYear(y) {
    if (!y) return '';
    const n = parseInt(y);
    return n < 100 ? '14' + String(n).padStart(2,'0') : String(n);
}
function shortYear(y) {
    const n = parseInt(y);
    return n >= 1400 ? n % 100 : n;
}
function fmtNum(n) {
    return n != null ? n.toLocaleString('ar-SA') : '—';
}
function numberOrNull(value) {
    if (value == null || value === '') return null;
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
}
function escapeHTML(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}
function fmtNumFlex(n, frac = 2) {
    if (n == null) return '—';
    const rounded = Math.round(n);
    if (Math.abs(n - rounded) < 0.001) return rounded.toLocaleString('ar-SA');
    return Number(n).toLocaleString('ar-SA', {
        minimumFractionDigits: frac,
        maximumFractionDigits: frac
    });
}
function pct(num, den) {
    if (!den || den === 0) return null;
    return Math.round((num / den) * 1000) / 10;
}
function rateBadge(val) {
    if (val == null) return '<span class="rate-na">—</span>';
    const cls = val >= 70 ? 'rate-high' : val >= 40 ? 'rate-mid' : 'rate-low';
    return `<span class="rate-badge ${cls}">${val.toFixed(1)}%</span>`;
}
function destroyChart(id) {
    if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

function absYearFromSemester(semester) {
    const n = parseInt(semester) || 0;
    return n < 100 ? 1400 + n : n;
}

function normalizeDepartment(dept) {
    const d = String(dept || '').trim();
    if (!d) return '';
    if (d === 'الثقافة الإسلامية') return 'الدراسات الإسلامية';
    return d;
}

function normalizeDegree(degree) {
    const d = String(degree || '').trim();
    if (!d) return '';
    if (d === 'البكالوريوس' || d === 'بكالوريوس' || d === 'بكالوريوس انتساب') return 'بكالوريوس';
    if (d === 'الماجستير' || d === 'ماجستير') return 'الماجستير';
    if (d === 'الدكتوراه' || d === 'دكتوراه') return 'دكتوراه';
    return d;
}

function isIslamicStudiesBachelor(programLike) {
    const major = normalizeSurveyProgramName(programLike?.name || programLike?.Major_aName || '');
    const degree = normalizeDegree(programLike?.degree || programLike?.Degree_aName || '');
    return major === 'الدراسات الإسلامية' && degree === 'بكالوريوس';
}

function getEffectiveStudentBranch(program, degree, explicitBranch = '') {
    const normalizedExplicitBranch = normalizeBranchName(explicitBranch);
    const programLike = { Major_aName: program, Degree_aName: degree };
    if (isIslamicStudiesBachelor(programLike)) return normalizedExplicitBranch;

    const normalizedProgram = normalizeSurveyProgramName(program || '');
    const normalizedDegree = normalizeDegree(degree || '');
    if (normalizedProgram && normalizedDegree) return 'الحوية';
    return normalizedExplicitBranch;
}

function getIslamicBranchNames() {
    const configured = Array.isArray(islamicBranchData?.branch_order)
        ? islamicBranchData.branch_order.map(normalizeBranchName).filter(Boolean)
        : [];
    return configured.length ? configured : ['الحوية', 'تربة', 'رنية', 'الخرمة'];
}

function getOrderedBranchNames(extraValues = []) {
    const configured = getIslamicBranchNames();
    const configuredSet = new Set(configured);
    const extras = [...new Set((extraValues || []).map(normalizeBranchName).filter(Boolean))]
        .filter(branch => !configuredSet.has(branch))
        .sort((a, b) => a.localeCompare(b, 'ar'));
    return [...configured, ...extras];
}

function getBranchesForProgram(programLike) {
    if (isIslamicStudiesBachelor(programLike)) return getIslamicBranchNames();
    return ['الحوية'];
}

function getIslamicBranchMetric(programLike, year, branch) {
    if (!isIslamicStudiesBachelor(programLike)) return null;
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch || normalizedBranch === ALL_BRANCH_FILTER_VALUE) return null;
    return islamicBranchData?.metrics?.[String(year)]?.[normalizedBranch] || null;
}

function getIslamicBranchCoverage(branch, year = null) {
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch) return null;
    if (year != null) {
        const yearKey = String(shortYear(year));
        return islamicBranchData?.coverage?.[yearKey]?.[normalizedBranch] || null;
    }
    const yearKeys = Object.keys(islamicBranchData?.coverage || {})
        .sort((a, b) => Number(b) - Number(a));
    for (const yearKey of yearKeys) {
        const coverage = islamicBranchData?.coverage?.[yearKey]?.[normalizedBranch];
        if (coverage) return coverage;
    }
    return null;
}

function isDeptSurveyFallbackAllowed(dept) {
    const normalizedDept = normalizeDepartment(dept);
    const rule = SURVEY_DEPT_FALLBACK_RULES[normalizedDept] || 'disabled';
    return rule === 'single_program_per_degree';
}

function normalizeRank(rank) {
    const r = String(rank || '').trim().replace('استاذ', 'أستاذ');
    if (!r) return '';
    if (r.includes('أستاذ مساعد')) return 'أستاذ مساعد';
    if (r.includes('أستاذ مشارك')) return 'أستاذ مشارك';
    if (r.includes('أستاذ')) return 'أستاذ';
    if (r.includes('محاضر')) return 'محاضر';
    if (r.includes('معيد')) return 'معيد';
    if (r.includes('مدرس')) return 'مدرس';
    if (r.includes('متعاون')) return 'متعاون';
    return r;
}

function getFacultyBaseForRatio(d) {
    if ((d.faculty_ratio_base || 0) > 0) return d.faculty_ratio_base;
    if ((d.faculty_total || 0) > 0) return d.faculty_total;
    return 0;
}

function parseFlatCSV(text, forcedSep = null) {
    if (!text || !text.trim()) return [];
    const lines = text.trim().split('\n').map(l => l.replace(/\r/g,''));
    if (lines.length < 2) return [];
    const sep = forcedSep || (lines[0].includes(';') ? ';' : ',');
    const headers = lines[0].split(sep).map(h => h.trim().replace(/^\uFEFF/, ''));
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const vals = lines[i].split(sep);
        const row = {};
        headers.forEach((h, idx) => { row[h] = (vals[idx] || '').trim(); });
        rows.push(row);
    }
    return rows;
}

function pickCell(row, keys) {
    for (const key of keys) {
        if (Object.prototype.hasOwnProperty.call(row, key) && row[key] != null && String(row[key]).trim() !== '') {
            return String(row[key]).trim();
        }
    }
    return '';
}

async function fetchTextIfExists(url) {
    try {
        const res = await fetch(url);
        if (!res.ok) return null;
        return await res.text();
    } catch {
        return null;
    }
}

async function fetchJSONIfExists(url) {
    try {
        const res = await fetch(url);
        if (!res.ok) return null;
        return await res.json();
    } catch {
        return null;
    }
}

function parseWindowAssignedJSON(text, windowVarName) {
    const raw = String(text || '').trim().replace(/^\uFEFF/, '');
    if (!raw) return null;
    const prefix = `window.${windowVarName}`;
    if (!raw.startsWith(prefix)) return null;

    const equalsIndex = raw.indexOf('=');
    if (equalsIndex === -1) return null;
    const jsonText = raw.slice(equalsIndex + 1).trim().replace(/;$/, '').trim();
    if (!jsonText) return null;

    try {
        return JSON.parse(jsonText);
    } catch {
        return null;
    }
}

function setBodyBlocked(blocked) {
    document.body.classList.toggle('app-blocked', Boolean(blocked));
}

function updateLoadingMessage(message) {
    const label = document.getElementById('loadingMessage');
    if (label) label.textContent = message;
}

function showLoadingOverlay(message = 'جاري تحميل مؤشرات الأداء...') {
    const overlay = document.getElementById('initialLoadingOverlay');
    updateLoadingMessage(message);
    if (overlay) overlay.classList.add('active');
    setBodyBlocked(true);
}

function hideLoadingOverlay() {
    const overlay = document.getElementById('initialLoadingOverlay');
    if (overlay) overlay.classList.remove('active');
    const loginVisible = !document.getElementById('loginOverlay')?.classList.contains('hidden');
    setBodyBlocked(loginVisible);
}

function setLoginError(message = '') {
    const error = document.getElementById('loginErrorMsg');
    if (error) error.textContent = message;
}

function showMainApp(memberName) {
    document.getElementById('loginOverlay')?.classList.add('hidden');
    document.getElementById('mainApp')?.classList.remove('hidden');

    const welcomeBadge = document.getElementById('welcomeBadge');
    const welcomeText = document.getElementById('welcomeText');
    if (welcomeBadge && welcomeText) {
        welcomeText.textContent = `مرحباً، ${memberName || 'مستخدم'}`;
        welcomeBadge.classList.remove('hidden');
    }
    setBodyBlocked(false);
}

function showLoginOverlay() {
    document.getElementById('mainApp')?.classList.add('hidden');
    document.getElementById('loginOverlay')?.classList.remove('hidden');
    document.getElementById('welcomeBadge')?.classList.add('hidden');
    hideLoadingOverlay();
    setBodyBlocked(true);
}

async function loadLoginMembersById() {
    if (loginMembersById) return loginMembersById;

    const csvText = await fetchTextIfExists(`data/faculty.csv?t=${Date.now()}`);
    if (!csvText) throw new Error('missing-login-members');

    const facultyRows = parseFlatCSV(csvText, ',');
    const byId = {};
    facultyRows.forEach(row => {
        const id = normalizeLoginDigits(pickCell(row, ['id', 'ID']));
        const name = pickCell(row, ['name', 'Name']);
        const year = parseInt(pickCell(row, ['year', 'Year']), 10) || 0;
        const active = pickCell(row, ['active', 'Active']) === 'نعم';
        if (!id || !name) return;

        const score = (active ? 100000 : 0) + year;
        if (!byId[id] || score > byId[id].score) {
            byId[id] = { name, score };
        }
    });

    loginMembersById = byId;
    return byId;
}

async function startApp() {
    if (appBootstrapped) return true;
    if (appBootPromise) return appBootPromise;

    appBootPromise = (async () => {
        showLoadingOverlay('جاري تحميل مؤشرات الأداء...');

        const ok = await loadData();
        if (!ok) return false;

        initDashboard();
        initProgramView();
        initCompare();

        updateLoadingMessage('جاري تحميل السجلات التفصيلية...');
        await Promise.all([loadGraduates(), loadNonCompleters(), loadStudentDetails(), loadFacultyRoster()]);
        initFacultyView();
        initEntrantsView();
        initGraduatesView();
        initNonCompleteView();
        initStudentCategoriesView();
        initAnalyticsView();

        appBootstrapped = true;
        return true;
    })();

    try {
        const ok = await appBootPromise;
        if (!ok) alert('تعذر تحميل بيانات الموقع.');
        return ok;
    } finally {
        appBootPromise = null;
        hideLoadingOverlay();
    }
}

async function handleLogin() {
    const employeeInput = document.getElementById('loginEmployeeId');
    const passwordInput = document.getElementById('loginPassword');
    const loginBtn = document.getElementById('loginBtn');
    const employeeId = normalizeLoginDigits(employeeInput?.value || '');
    const password = normalizeLoginDigits(passwordInput?.value || '');

    setLoginError('');
    if (!employeeId) {
        setLoginError('يرجى إدخال رقم المنسوب');
        employeeInput?.focus();
        return;
    }
    if (password !== '1429') {
        setLoginError('كلمة المرور غير صحيحة');
        passwordInput?.focus();
        passwordInput?.select();
        return;
    }

    if (loginBtn) loginBtn.disabled = true;
    try {
        const membersById = await loadLoginMembersById();
        const member = membersById[employeeId];
        if (!member) {
            setLoginError('رقم المنسوب غير موجود في السجلات');
            employeeInput?.focus();
            employeeInput?.select();
            return;
        }

        sessionStorage.setItem('loggedIn', 'true');
        sessionStorage.setItem('employeeId', employeeId);
        sessionStorage.setItem('employeeName', member.name);

        showMainApp(member.name);
        await startApp();
    } catch (error) {
        console.error(error);
        setLoginError('تعذر التحقق من بيانات الدخول');
    } finally {
        if (loginBtn) loginBtn.disabled = false;
    }
}

function handleLogout() {
    sessionStorage.removeItem('loggedIn');
    sessionStorage.removeItem('employeeId');
    sessionStorage.removeItem('employeeName');

    const employeeInput = document.getElementById('loginEmployeeId');
    const passwordInput = document.getElementById('loginPassword');
    if (employeeInput) employeeInput.value = '';
    if (passwordInput) passwordInput.value = '';

    setLoginError('');
    showLoginOverlay();
}

function buildWeights(programKeys, studentsByProgramKey) {
    const uniqueKeys = [...new Set(programKeys)];
    if (uniqueKeys.length === 0) return [];
    if (uniqueKeys.length === 1) return [{ key: uniqueKeys[0], weight: 1 }];

    const demands = uniqueKeys.map(k => Math.max(0, Number(studentsByProgramKey[k]) || 0));
    const demandTotal = demands.reduce((s, x) => s + x, 0);
    if (demandTotal > 0) {
        return uniqueKeys.map((key, idx) => ({ key, weight: demands[idx] / demandTotal }));
    }
    const equal = 1 / uniqueKeys.length;
    return uniqueKeys.map(key => ({ key, weight: equal }));
}

function normalizeArabicDigits(str) {
    if (str == null) return '';
    return String(str).replace(/[٠-٩]/g, d => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
}

function normalizeArabicText(str) {
    if (str == null) return '';
    return normalizeArabicDigits(String(str))
        .replace(/[\u200E\u200F]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

function normalizeLoginDigits(value) {
    if (value == null) return '';
    const map = {
        '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
        '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9',
        '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
        '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
    };
    return String(value).replace(/[٠-٩۰-۹]/g, ch => map[ch] || ch).trim();
}

function normalizeSurveyProgramName(name) {
    const base = normalizeArabicText(name)
        .replace(/^برنامج\s+/,'')
        .replace(/^(?:ال)?(?:بكالوريوس|ماجستير|الماجستير|دكتوراه)\s+/, '');
    if (!base) return '';
    return GRADUATE_PROGRAM_ALIASES[base] || base;
}

function detectSurveyDegree(programName, allowedDegrees = null) {
    const raw = normalizeArabicText(programName);
    if (/دكتوراه/.test(raw)) return 'دكتوراه';
    if (/ماجستير/.test(raw)) return 'الماجستير';
    if (/بكالوريوس/.test(raw)) return 'بكالوريوس';
    if (allowedDegrees && allowedDegrees.size === 1) return [...allowedDegrees][0];
    return '';
}

function buildProgramMajorDegreeKey(programName, degreeName) {
    const major = normalizeSurveyProgramName(programName);
    const degree = normalizeDegree(degreeName);
    if (!major || !degree) return '';
    return `${major}|${degree}`;
}

function buildProgramDataKey(year, dept, programName, degreeName) {
    const normalizedYear = parseInt(year, 10);
    const normalizedDept = normalizeDepartment(dept);
    const normalizedMajor = normalizeSurveyProgramName(programName);
    const normalizedDegree = normalizeDegree(degreeName);
    if (!normalizedYear || !normalizedDept || !normalizedMajor || !normalizedDegree) return '';
    return `${normalizedYear}|${normalizedDept}|${normalizedMajor}|${normalizedDegree}`;
}

function normalizeBranchName(branch) {
    return String(branch || '').trim();
}

function getFacultyBranchNames(branch) {
    return [...new Set(
        String(branch || '')
            .split(/[|,،]/)
            .map(normalizeBranchName)
            .filter(Boolean)
    )];
}

function getBranchFromTeachingLocation(location, allowedBranches = []) {
    const value = String(location || '').trim();
    const candidates = allowedBranches.length ? allowedBranches : getIslamicBranchNames();
    return candidates.find(branch => {
        const normalized = normalizeBranchName(branch);
        const withoutArticle = normalized.replace(/^ال/, '');
        return [normalized, withoutArticle]
            .filter(Boolean)
            .some(alias => value.includes(alias));
    }) || '';
}

function buildProgramBranchDataKey(programKey, branch) {
    const normalizedBranch = normalizeBranchName(branch);
    if (!programKey || !normalizedBranch) return '';
    return `${programKey}|${normalizedBranch}`;
}

function buildYearDeptKey(year, dept) {
    const normalizedYear = parseInt(year, 10);
    const normalizedDept = normalizeDepartment(dept);
    if (!normalizedYear || !normalizedDept) return '';
    return `${normalizedYear}|${normalizedDept}`;
}

function buildYearDeptBranchKey(year, dept, branch) {
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch) return '';
    const yearDeptKey = buildYearDeptKey(year, dept);
    return yearDeptKey ? `${yearDeptKey}|${normalizedBranch}` : '';
}

function detectSectionGender(course) {
    const location = String(course?.l || '').trim();
    if (location.includes('طالبات') || location.includes('الطالبات')) return 'female';
    return 'male';
}

function classifyFacultyNationality(value) {
    const normalized = normalizeArabicText(value);
    if (!normalized) return 'unknown';
    return normalized.includes('سعود') ? 'saudi' : 'nonSaudi';
}

function classifyFacultyGender(value) {
    const normalized = normalizeArabicText(value);
    if (!normalized) return 'unknown';
    if (normalized.includes('ذكر')) return 'male';
    if (normalized.includes('انث')) return 'female';
    return 'unknown';
}

function classifyStudentNationality(value) {
    const normalized = normalizeArabicText(value);
    if (!normalized) return 'unknown';
    if (
        normalized === 'سعودي' ||
        normalized === 'سعودية' ||
        normalized === 'السعودي' ||
        normalized === 'السعودية'
    ) {
        return 'saudi';
    }
    if (normalized.includes('غير سعودي')) {
        return 'nonSaudi';
    }
    return 'nonSaudi';
}

function classifyStudentGender(value) {
    const normalized = normalizeArabicText(value);
    if (!normalized) return 'unknown';
    if (normalized.includes('ذكر')) return 'male';
    if (normalized.includes('أنث') || normalized.includes('انث')) return 'female';
    return 'unknown';
}

function createFacultyRankAggregate() {
    return {
        count: 0,
        loadSum: 0,
        maleLoadSum: 0,
        maleLoadCount: 0,
        femaleLoadSum: 0,
        femaleLoadCount: 0,
        saudiMaleCount: 0,
        saudiFemaleCount: 0,
        saudiCount: 0,
        nonSaudiMaleCount: 0,
        nonSaudiFemaleCount: 0,
        nonSaudiCount: 0,
        unknownNationalityCount: 0,
        unknownGenderCount: 0
    };
}

function mergeFacultyRankAggregate(target, source) {
    if (!source) return target;
    target.count += Number(source.count) || 0;
    target.loadSum += Number(source.loadSum) || 0;
    target.maleLoadSum += Number(source.maleLoadSum) || 0;
    target.maleLoadCount += Number(source.maleLoadCount) || 0;
    target.femaleLoadSum += Number(source.femaleLoadSum) || 0;
    target.femaleLoadCount += Number(source.femaleLoadCount) || 0;
    target.saudiMaleCount += Number(source.saudiMaleCount) || 0;
    target.saudiFemaleCount += Number(source.saudiFemaleCount) || 0;
    target.saudiCount += Number(source.saudiCount) || 0;
    target.nonSaudiMaleCount += Number(source.nonSaudiMaleCount) || 0;
    target.nonSaudiFemaleCount += Number(source.nonSaudiFemaleCount) || 0;
    target.nonSaudiCount += Number(source.nonSaudiCount) || 0;
    target.unknownNationalityCount += Number(source.unknownNationalityCount) || 0;
    target.unknownGenderCount += Number(source.unknownGenderCount) || 0;
    return target;
}

function createResearchActivityAggregate() {
    return {
        faculty_total: 0,
        faculty_published: 0,
        research_count: 0,
        citations: 0,
        citations_per_publication: 0,
        research_source: 'faculty_activities_live'
    };
}

function createRawTeachingSupportBucket(includeBranches = true) {
    const bucket = {
        totalSections: 0,
        totalStudents: 0,
        maleSections: 0,
        maleStudents: 0,
        femaleSections: 0,
        femaleStudents: 0,
        exclusiveSections: 0,
        exclusiveStudents: 0,
        exclusiveMaleSections: 0,
        exclusiveMaleStudents: 0,
        exclusiveFemaleSections: 0,
        exclusiveFemaleStudents: 0,
        facultyById: {}
    };
    if (includeBranches) bucket.branchBuckets = {};
    return bucket;
}

function getRawTeachingSupportBranchBucket(bucket, branch) {
    const normalizedBranch = normalizeBranchName(branch);
    if (!bucket || !normalizedBranch) return null;
    if (!bucket.branchBuckets) bucket.branchBuckets = {};
    if (!bucket.branchBuckets[normalizedBranch]) {
        bucket.branchBuckets[normalizedBranch] = createRawTeachingSupportBucket(false);
    }
    return bucket.branchBuckets[normalizedBranch];
}

function finalizeTeachingSupportBucket(bucket) {
    if (!bucket) return null;
    const useExclusive = (bucket.exclusiveSections || 0) > 0;
    const sectionsTotal = useExclusive ? (bucket.exclusiveSections || 0) : (bucket.totalSections || 0);
    const studentsTotal = useExclusive ? (bucket.exclusiveStudents || 0) : (bucket.totalStudents || 0);
    const maleSections = useExclusive ? (bucket.exclusiveMaleSections || 0) : (bucket.maleSections || 0);
    const maleStudents = useExclusive ? (bucket.exclusiveMaleStudents || 0) : (bucket.maleStudents || 0);
    const femaleSections = useExclusive ? (bucket.exclusiveFemaleSections || 0) : (bucket.femaleSections || 0);
    const femaleStudents = useExclusive ? (bucket.exclusiveFemaleStudents || 0) : (bucket.femaleStudents || 0);

    const facultyEntries = Object.values(bucket.facultyById || {});
    const facultyIncluded = facultyEntries.filter(entry =>
        (entry.nonSharedHours || 0) > 0 || (entry.exclusiveHours || 0) > 0
    );
    const facultyFallback = facultyIncluded.length ? facultyIncluded : facultyEntries.filter(entry => (entry.totalHours || 0) > 0);
    const facultyRanks = {};
    let facultyLoadSum = 0;

    facultyFallback.forEach(entry => {
        const rank = normalizeRank(entry.rank || '') || 'غير مصنف';
        const effectiveHours = entry.nonSharedHours > 0
            ? entry.nonSharedHours
            : (entry.exclusiveHours > 0 ? entry.exclusiveHours : entry.totalHours);
        if (!Number.isFinite(effectiveHours) || effectiveHours <= 0) return;

        if (!facultyRanks[rank]) facultyRanks[rank] = createFacultyRankAggregate();
        facultyRanks[rank].count++;
        facultyRanks[rank].loadSum += effectiveHours;
        const genderGroup = classifyFacultyGender(entry.gender);
        if (genderGroup === 'male') {
            facultyRanks[rank].maleLoadSum += effectiveHours;
            facultyRanks[rank].maleLoadCount++;
        } else if (genderGroup === 'female') {
            facultyRanks[rank].femaleLoadSum += effectiveHours;
            facultyRanks[rank].femaleLoadCount++;
        } else {
            facultyRanks[rank].unknownGenderCount++;
        }

        const nationalityGroup = classifyFacultyNationality(entry.nationality);
        if (nationalityGroup === 'saudi') {
            facultyRanks[rank].saudiCount++;
            if (genderGroup === 'male') facultyRanks[rank].saudiMaleCount++;
            else if (genderGroup === 'female') facultyRanks[rank].saudiFemaleCount++;
        } else if (nationalityGroup === 'nonSaudi') {
            facultyRanks[rank].nonSaudiCount++;
            if (genderGroup === 'male') facultyRanks[rank].nonSaudiMaleCount++;
            else if (genderGroup === 'female') facultyRanks[rank].nonSaudiFemaleCount++;
        } else {
            facultyRanks[rank].unknownNationalityCount++;
        }
        facultyLoadSum += effectiveHours;
    });

    return {
        sectionSource: useExclusive ? 'exclusive' : (sectionsTotal > 0 ? 'mapped' : 'none'),
        totalSections: sectionsTotal,
        totalStudentsInSections: studentsTotal,
        avgStudentsPerSection: sectionsTotal > 0 ? Math.round((studentsTotal / sectionsTotal) * 10) / 10 : null,
        maleSections,
        maleStudents,
        avgMaleStudentsPerSection: maleSections > 0 ? Math.round((maleStudents / maleSections) * 10) / 10 : null,
        femaleSections,
        femaleStudents,
        avgFemaleStudentsPerSection: femaleSections > 0 ? Math.round((femaleStudents / femaleSections) * 10) / 10 : null,
        facultyCount: facultyFallback.length,
        facultyAvgLoad: facultyFallback.length > 0
            ? Math.round((facultyLoadSum / facultyFallback.length) * 10) / 10 : null,
        facultyRanks
    };
}

function roundToNearestFive(value) {
    if (!Number.isFinite(value)) return null;
    return Math.max(0, Math.round(value / 5) * 5);
}

function estimateNextValue(points) {
    const usable = (points || [])
        .filter(point => Number.isFinite(point?.value))
        .sort((a, b) => a.year - b.year);

    if (!usable.length) return null;
    if (usable.length === 1) return usable[0].value;

    const last = usable[usable.length - 1];
    const prev = usable[usable.length - 2];
    const gap = Math.max(1, (last.year || 0) - (prev.year || 0));
    const deltaPerYear = (last.value - prev.value) / gap;
    return last.value + deltaPerYear;
}

function estimateNextRoundedCount(points) {
    const estimated = estimateNextValue(points);
    return Number.isFinite(estimated) ? roundToNearestFive(estimated) : null;
}

function getStudentCategoryDefinition(categoryId) {
    return STUDENT_CATEGORY_DEFS.find(category => category.id === categoryId) || null;
}

function classifyStudentCategoryId(gpa) {
    if (!Number.isFinite(gpa)) return '';
    const match = STUDENT_CATEGORY_DEFS.find(category => category.matches(gpa));
    return match ? match.id : '';
}

function getOfficialGpaEstimate(gpa) {
    if (!Number.isFinite(gpa)) return 'غير محدد';
    if (gpa >= 3.5) return 'ممتاز';
    if (gpa >= 2.75) return 'جيد جدًا';
    if (gpa >= 1.75) return 'جيد';
    if (gpa >= 1) return 'مقبول';
    return 'أقل من مقبول';
}

function buildShari3ahSurveysProgramKey(programName, degreeName) {
    const normalizedProgram = normalizeSurveyProgramName(programName);
    const normalizedDegree = normalizeDegree(degreeName);
    if (!normalizedProgram || !normalizedDegree) return '';
    return `${normalizedProgram}|${normalizedDegree}`;
}

function getShari3ahSurveysProgramId(programName, degreeName) {
    const key = buildShari3ahSurveysProgramKey(programName, degreeName);
    return SHARI3AH_SURVEYS_PROGRAM_ID_MAP[key] || '';
}

function parseCSVQuotedRows(text, separator = ',') {
    if (!text || !text.trim()) return [];
    const rows = [];
    let row = [];
    let cell = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '"') {
            if (inQuotes && text[i + 1] === '"') {
                cell += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
            continue;
        }
        if (!inQuotes && ch === separator) {
            row.push(cell);
            cell = '';
            continue;
        }
        if (!inQuotes && (ch === '\n' || ch === '\r')) {
            row.push(cell);
            cell = '';
            if (row.some(v => String(v || '').trim() !== '')) rows.push(row);
            row = [];
            if (ch === '\r' && text[i + 1] === '\n') i++;
            continue;
        }
        cell += ch;
    }

    if (cell.length || row.length) {
        row.push(cell);
        if (row.some(v => String(v || '').trim() !== '')) rows.push(row);
    }
    return rows;
}

function parseCSVQuotedObjects(text, separator = ',') {
    const matrix = parseCSVQuotedRows(text, separator);
    if (matrix.length < 2) return [];
    const headers = matrix[0].map(h => String(h || '').trim().replace(/^\uFEFF/, ''));
    const rows = [];
    for (let i = 1; i < matrix.length; i++) {
        const vals = matrix[i];
        const row = {};
        headers.forEach((h, idx) => { row[h] = String(vals[idx] || '').trim(); });
        rows.push(row);
    }
    return rows;
}

function parseAcademicYear(value) {
    const raw = normalizeArabicText(value);
    if (!raw) return null;
    const nums = raw.match(/\d+/g);
    if (!nums || !nums.length) return null;
    const n = parseInt(nums[nums.length - 1], 10);
    if (!Number.isFinite(n)) return null;
    return n >= 1400 ? (n % 100) : n;
}

function parseDaySerial(value) {
    const raw = normalizeArabicText(value);
    if (!raw) return null;

    // YYYY-MM-DD
    const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) {
        const y = parseInt(iso[1], 10);
        const m = parseInt(iso[2], 10) - 1;
        const d = parseInt(iso[3], 10);
        const ms = Date.UTC(y, m, d);
        return Number.isFinite(ms) ? (ms / 86400000) : null;
    }

    // Excel serial date
    if (/^\d+(?:\.\d+)?$/.test(raw)) {
        const serial = parseFloat(raw);
        if (Number.isFinite(serial) && serial >= 20000 && serial <= 70000) {
            return Math.floor(serial) - 25569;
        }
    }

    const parsed = new Date(raw);
    if (Number.isNaN(parsed.getTime())) return null;
    return Math.floor(parsed.getTime() / 86400000);
}

async function applyAverageGraduationDurationFromDetails(rows) {
    const csvText = await fetchTextIfExists(`data/graduates_detail.csv?t=${Date.now()}`);
    if (!csvText) return { applied: false, reason: 'missing-graduates-detail' };

    const gradRows = parseCSVQuotedObjects(csvText, ';');
    if (!gradRows.length) return { applied: false, reason: 'empty-graduates-detail' };

    const aggregates = {}; // year|program|degree -> {sumYears, count}
    let validRecords = 0;

    gradRows.forEach(g => {
        const degree = normalizeDegree(g['الدرجة'] || g['Degree_aName'] || g['degree']);
        if (!isGradDegree(degree)) return;

        const year = parseAcademicYear(g['السنة'] || g['Semester'] || g['year']);
        const program = normalizeSurveyProgramName(g['التخصص'] || g['Major_aName'] || g['program']);
        if (!year || !program) return;

        const admissionDay = parseDaySerial(g['تاريخ_القبول'] || g['admission_date'] || g['AdmissionDate']);
        const graduateDay = parseDaySerial(g['تاريخ_التخرج'] || g['grad_date'] || g['GraduationDate']);
        if (admissionDay == null || graduateDay == null || graduateDay < admissionDay) return;

        const years = (graduateDay - admissionDay) / 365.25;
        if (!Number.isFinite(years) || years < 0.5 || years > 10) return;

        const key = `${year}|${program}|${degree}`;
        if (!aggregates[key]) aggregates[key] = { sumYears: 0, count: 0 };
        aggregates[key].sumYears += years;
        aggregates[key].count++;
        validRecords++;
    });

    if (!validRecords) return { applied: false, reason: 'no-valid-duration-records' };

    let appliedRows = 0;
    const matchedGroups = new Set();
    rows.forEach(r => {
        const degree = normalizeDegree(r.Degree_aName);
        if (!isGradDegree(degree)) return;

        const year = parseInt(r.Semester, 10);
        const program = normalizeSurveyProgramName(r.Major_aName);
        const key = `${year}|${program}|${degree}`;
        const agg = aggregates[key];
        if (!agg || agg.count <= 0) return;

        r.avg_time_to_graduate = Math.round((agg.sumYears / agg.count) * 100) / 100;
        r.avg_time_to_graduate_count = agg.count;
        r.avg_time_source = 'graduates_detail';
        matchedGroups.add(key);
        appliedRows++;
    });

    return {
        applied: appliedRows > 0,
        appliedRows,
        matchedGroups: matchedGroups.size,
        validRecords
    };
}

function extractNumericValues(value) {
    if (value == null) return [];
    const normalized = normalizeArabicText(value)
        .replace(/[٫،]/g, '.')
        .replace(/[–—]/g, '-');
    const matches = normalized.match(/\d+(?:\.\d+)?/g);
    if (!matches) return [];
    return matches.map(n => parseFloat(n)).filter(Number.isFinite);
}

function clampValue(value, min, max) {
    if (!Number.isFinite(value)) return null;
    return Math.min(max, Math.max(min, value));
}

function parseRangeOrSingleValue(value, min, max) {
    const raw = normalizeArabicText(value);
    if (!raw) return null;
    const nums = extractNumericValues(raw);
    if (!nums.length) return null;

    let parsed = nums[0];
    const hasRange = /-\s*\d/.test(raw);
    if (hasRange && nums.length >= 2) parsed = (nums[0] + nums[1]) / 2;

    if (/^(?:أقل|اقل)\s*من/.test(raw)) parsed = nums[0] - 0.25;
    if (/^(?:أعلى|اعلى|أكثر)\s*من/.test(raw)) parsed = nums[0] + 0.25;

    return clampValue(parsed, min, max);
}

function parseSurveyYear(value) {
    const nums = extractNumericValues(value).map(n => Math.round(n));
    if (!nums.length) return null;
    const fullYear = nums.find(n => n >= 1400);
    if (fullYear) return fullYear % 100;
    const candidate = nums[nums.length - 1];
    if (!candidate) return null;
    return candidate >= 100 ? candidate % 100 : candidate;
}

function parseSurveyEmploymentStatus(value) {
    const raw = normalizeArabicText(value);
    if (!raw) return null;

    const positive = [
        'موظف', 'يعمل', 'أعمل', 'اعمل', 'عمل حر', 'رائد أعمال', 'صاحب عمل',
        'أكمل دراسات عليا', 'اكمل دراسات عليا', 'مكمل دراسات عليا', 'دراسات عليا'
    ];
    const negative = ['أبحث عن عمل', 'ابحث عن عمل', 'باحث عن عمل', 'عاطل', 'لا أعمل', 'غير موظف'];

    if (positive.some(x => raw.includes(x))) return true;
    if (negative.some(x => raw.includes(x))) return false;
    return null;
}

function normalizeHeaderKey(header) {
    return normalizeArabicText(header).toLowerCase();
}

function findHeaderByAllParts(headers, parts) {
    const normalizedParts = parts.map(p => normalizeHeaderKey(p));
    const match = headers.find(h => {
        const key = normalizeHeaderKey(h);
        return normalizedParts.every(part => key.includes(part));
    });
    return match || '';
}

function findHeaderByCandidates(headers, candidates) {
    for (const parts of candidates) {
        const hit = findHeaderByAllParts(headers, parts);
        if (hit) return hit;
    }
    return '';
}

function detectGraduateSurveyColumns(headers) {
    return {
        program: findHeaderByCandidates(headers, [
            ['اسم البرنامج'],
            ['البرنامج الأكاديمي'],
            ['البرنامج']
        ]),
        year: findHeaderByCandidates(headers, [
            ['سنة التخرج'],
            ['سنه التخرج'],
            ['التخرج من البرنامج']
        ]),
        courseEval: findHeaderByCandidates(headers, [
            ['جودة المقررات'],
            ['تقييم المقررات']
        ]),
        experience: findHeaderByCandidates(headers, [
            ['تقييمك العام', 'جودة التعلم'],
            ['جودة خبرات التعلم']
        ]),
        supervision: findHeaderByCandidates(headers, [
            ['جودة الإشراف'],
            ['الاشراف', 'الرسالة'],
            ['الإشراف العلمي']
        ]),
        services: findHeaderByCandidates(headers, [
            ['رضاك', 'الخدمات المقدمة'],
            ['رضا الطلاب', 'الخدمات'],
            ['مستوى الخدمات']
        ]),
        status: findHeaderByCandidates(headers, [
            ['وضعك الحالي بعد التخرج'],
            ['وضعك الحالي']
        ]),
        performance: findHeaderByCandidates(headers, [
            ['درجتك', 'الاختبارات الوطنية'],
            ['الاختبارات', 'مهنية']
        ]),
        employerEval: findHeaderByCandidates(headers, [
            ['تقييم رئيسك'],
            ['تقيّم نفسك'],
            ['التقييم من ٥'],
            ['التقييم من 5']
        ]),
    };
}

function resolveGraduateSurveyCsvUrl(url) {
    const raw = String(url || '').trim();
    if (!raw) return '';
    if (raw.includes('output=csv') || raw.includes('format=csv')) return raw;

    try {
        const parsed = new URL(raw);
        const match = parsed.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!match) return raw;
        const gidFromHash = parsed.hash.match(/gid=(\d+)/);
        const gid = parsed.searchParams.get('gid') || (gidFromHash ? gidFromHash[1] : '');
        return `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv${gid ? `&gid=${gid}` : ''}`;
    } catch {
        return raw;
    }
}

function aggregateGraduateSurveyRows(surveyRows, allowedDegrees = null) {
    if (!surveyRows.length) return { metricsByKey: {}, groups: 0, matchedRows: 0, reason: 'empty' };

    const headers = Object.keys(surveyRows[0]);
    const columns = detectGraduateSurveyColumns(headers);
    if (!columns.program || !columns.year) {
        return { metricsByKey: {}, groups: 0, matchedRows: 0, reason: 'missing-required-columns' };
    }

    const grouped = {};
    let matchedRows = 0;
    surveyRows.forEach(row => {
        const program = normalizeSurveyProgramName(row[columns.program]);
        const degree = detectSurveyDegree(row[columns.program], allowedDegrees);
        const year = parseSurveyYear(row[columns.year]);
        if (!program || !year) return;

        const key = `${year}|${program}|${degree || '*'}`;
        if (!grouped[key]) {
            grouped[key] = {
                responses: 0,
                courseSum: 0, courseCount: 0,
                experienceSum: 0, experienceCount: 0,
                supervisionSum: 0, supervisionCount: 0,
                servicesSum: 0, servicesCount: 0,
                perfSum: 0, perfCount: 0,
                employedCount: 0, employmentCount: 0,
                employerEvalSum: 0, employerEvalCount: 0
            };
        }
        const g = grouped[key];
        g.responses++;
        matchedRows++;

        const courseEval = parseRangeOrSingleValue(row[columns.courseEval], 1, 5);
        if (courseEval != null) {
            g.courseSum += courseEval;
            g.courseCount++;
        }

        const experience = parseRangeOrSingleValue(row[columns.experience], 1, 5);
        if (experience != null) {
            g.experienceSum += experience;
            g.experienceCount++;
        }

        const supervisionEval = parseRangeOrSingleValue(row[columns.supervision], 1, 5);
        if (supervisionEval != null) {
            g.supervisionSum += supervisionEval;
            g.supervisionCount++;
        }

        const servicesEval = parseRangeOrSingleValue(row[columns.services], 1, 5);
        if (servicesEval != null) {
            g.servicesSum += servicesEval;
            g.servicesCount++;
        }

        const performance = parseRangeOrSingleValue(row[columns.performance], 0, 100);
        if (performance != null) {
            g.perfSum += performance;
            g.perfCount++;
        }

        const employmentStatus = parseSurveyEmploymentStatus(row[columns.status]);
        if (employmentStatus != null) {
            g.employmentCount++;
            if (employmentStatus) g.employedCount++;
        }

        const employerEval = parseRangeOrSingleValue(row[columns.employerEval], 1, 5);
        if (employerEval != null) {
            g.employerEvalSum += employerEval;
            g.employerEvalCount++;
        }
    });

    const metricsByKey = {};
    Object.entries(grouped).forEach(([key, g]) => {
        metricsByKey[key] = {
            eval_courses: g.courseCount > 0
                ? Math.round((g.courseSum / g.courseCount) * 100) / 100 : null,
            eval_courses_sample: g.courseCount,
            eval_experience: g.experienceCount > 0
                ? Math.round((g.experienceSum / g.experienceCount) * 100) / 100 : null,
            eval_experience_sample: g.experienceCount,
            eval_supervision: g.supervisionCount > 0
                ? Math.round((g.supervisionSum / g.supervisionCount) * 100) / 100 : null,
            eval_supervision_sample: g.supervisionCount,
            eval_services: g.servicesCount > 0
                ? Math.round((g.servicesSum / g.servicesCount) * 100) / 100 : null,
            eval_services_sample: g.servicesCount,
            performance_rate: g.perfCount > 0
                ? Math.round((g.perfSum / g.perfCount) * 10) / 10 : null,
            performance_rate_sample: g.perfCount,
            employment_rate: g.employmentCount > 0
                ? pct(g.employedCount, g.employmentCount) : null,
            employment_employed_count: g.employedCount,
            employment_rate_sample: g.employmentCount,
            eval_employers: g.employerEvalCount > 0
                ? Math.round((g.employerEvalSum / g.employerEvalCount) * 100) / 100 : null,
            eval_employers_sample: g.employerEvalCount,
        };
    });

    return {
        metricsByKey,
        groups: Object.keys(metricsByKey).length,
        matchedRows,
        reason: Object.keys(metricsByKey).length ? '' : 'no-metrics'
    };
}

function isDegreeAllowed(row, allowedDegrees) {
    if (!allowedDegrees || !allowedDegrees.size) return true;
    const degree = normalizeDegree(row.Degree_aName);
    return allowedDegrees.has(degree);
}

function isMetricMissing(value) {
    return value == null || String(value).trim() === '';
}

function getGraduateSampleMetricKeys(degreeName) {
    const degree = normalizeDegree(degreeName);
    return GRADUATE_SAMPLE_METRICS_BY_DEGREE[degree] || new Set();
}

function applyGraduateSurveyMetrics(rows, metricsByKey, allowedDegrees = null, sourceLabel = '') {
    let appliedRows = 0;
    const matchedGroups = new Set();
    const deptDegreeProgramCounts = {};

    rows.forEach(row => {
        if (!isDegreeAllowed(row, allowedDegrees)) return;
        const year = parseInt(row.Semester, 10);
        const degree = normalizeDegree(row.Degree_aName);
        const dept = normalizeSurveyProgramName(normalizeDepartment(row.Dept_aName));
        if (!year || !degree || !dept) return;
        const key = `${year}|${dept}|${degree}`;
        deptDegreeProgramCounts[key] = (deptDegreeProgramCounts[key] || 0) + 1;
    });

    rows.forEach(row => {
        if (!isDegreeAllowed(row, allowedDegrees)) return;

        const year = parseInt(row.Semester, 10);
        const degree = normalizeDegree(row.Degree_aName);
        const normalizedMajor = normalizeSurveyProgramName(row.Major_aName);
        const majorKey = `${year}|${normalizedMajor}|${degree}`;
        const legacyMajorKey = `${year}|${normalizedMajor}|*`;
        const normalizedDept = normalizeSurveyProgramName(normalizeDepartment(row.Dept_aName));
        const deptKey = `${year}|${normalizedDept}|${degree}`;
        const legacyDeptKey = `${year}|${normalizedDept}|*`;
        const deptDegreeKey = `${year}|${normalizedDept}|${degree}`;

        const matchedProgramKey = metricsByKey[majorKey] ? majorKey : (metricsByKey[legacyMajorKey] ? legacyMajorKey : '');
        const matchedDeptKey = metricsByKey[deptKey] ? deptKey : (metricsByKey[legacyDeptKey] ? legacyDeptKey : '');
        const hasProgramMetrics = Boolean(matchedProgramKey);
        const deptFallbackEnabled = isDeptSurveyFallbackAllowed(normalizedDept);
        const canUseDeptFallback = !hasProgramMetrics
            && deptFallbackEnabled
            && Boolean(matchedDeptKey)
            && deptDegreeProgramCounts[deptDegreeKey] === 1;
        const metrics = hasProgramMetrics
            ? metricsByKey[matchedProgramKey]
            : (canUseDeptFallback ? metricsByKey[matchedDeptKey] : null);
        if (!metrics) return;

        const eligibleMetrics = getGraduateSampleMetricKeys(degree);
        const usedProgramMetrics = hasProgramMetrics;
        const scope = usedProgramMetrics ? 'program' : 'dept';
        const sampleSource = `graduates_sample_${scope}${sourceLabel ? `_${sourceLabel}` : ''}`;
        let touched = false;
        if (eligibleMetrics.has('eval_supervision') && isMetricMissing(row.eval_supervision) && metrics.eval_supervision != null) {
            row.eval_supervision = metrics.eval_supervision;
            row.eval_supervision_sample = metrics.eval_supervision_sample || 0;
            row.eval_supervision_source = sampleSource;
            touched = true;
        }
        if (eligibleMetrics.has('eval_services') && isMetricMissing(row.eval_services) && metrics.eval_services != null) {
            row.eval_services = metrics.eval_services;
            row.eval_services_sample = metrics.eval_services_sample || 0;
            row.eval_services_source = sampleSource;
            touched = true;
        }
        if (eligibleMetrics.has('performance_rate') && isMetricMissing(row.performance_rate) && metrics.performance_rate != null) {
            row.performance_rate = metrics.performance_rate;
            row.performance_rate_sample = metrics.performance_rate_sample || 0;
            row.performance_rate_source = sampleSource;
            touched = true;
        }
        if (eligibleMetrics.has('employment_rate') && isMetricMissing(row.employment_rate) && metrics.employment_rate != null) {
            row.employment_rate = metrics.employment_rate;
            row.employment_employed_count = metrics.employment_employed_count || 0;
            row.employment_rate_sample = metrics.employment_rate_sample || 0;
            row.employment_rate_source = sampleSource;
            touched = true;
        }
        if (eligibleMetrics.has('eval_employers') && isMetricMissing(row.eval_employers) && metrics.eval_employers != null) {
            row.eval_employers = metrics.eval_employers;
            row.eval_employers_sample = metrics.eval_employers_sample || 0;
            row.eval_employers_source = sampleSource;
            touched = true;
        }
        if (touched) {
            row.survey_source = sampleSource;
            matchedGroups.add(usedProgramMetrics ? matchedProgramKey : matchedDeptKey);
            appliedRows++;
        }
    });
    return { appliedRows, matchedGroups: matchedGroups.size };
}

function isShari3ahProgramEvaluationSurvey(title) {
    return normalizeArabicText(title) === normalizeArabicText(SHARI3AH_PROGRAM_EVAL_SURVEY_TITLE);
}

function aggregateProgramExperienceMetricFromSurvey(survey) {
    if (!survey || !Array.isArray(survey.topics)) return null;

    let scoreTotal = 0;
    let responseTotal = 0;
    const itemResponseMap = new Map();

    survey.topics.forEach(topic => {
        (topic.items || []).forEach(item => {
            const itemKey = [
                normalizeArabicText(topic.label),
                normalizeArabicText(item.number),
                normalizeArabicText(item.label)
            ].join('||');

            let itemResponses = 0;
            (item.genders || []).forEach(genderEntry => {
                const responses = Number(genderEntry.responses || 0);
                const weightedTotal = Number(genderEntry.scoreTotal || 0);
                if (!Number.isFinite(responses) || responses <= 0) return;

                itemResponses += responses;
                responseTotal += responses;
                scoreTotal += Number.isFinite(weightedTotal) ? weightedTotal : 0;
            });

            if (itemResponses > 0) {
                itemResponseMap.set(itemKey, (itemResponseMap.get(itemKey) || 0) + itemResponses);
            }
        });
    });

    if (!responseTotal || !itemResponseMap.size) return null;

    return {
        eval_experience: Math.round((scoreTotal / responseTotal) * 100) / 100,
        eval_experience_sample: Math.max(0, ...itemResponseMap.values())
    };
}

function extractProgramExperienceMetricsFromShari3ahSurveys(payload) {
    const extractedData = payload && payload.extractedData ? payload.extractedData : {};
    const metricsByDatasetKey = {};

    Object.entries(extractedData).forEach(([datasetKey, dataset]) => {
        const survey = (dataset.surveys || []).find(item => isShari3ahProgramEvaluationSurvey(item.title));
        if (!survey) return;

        const metric = aggregateProgramExperienceMetricFromSurvey(survey);
        if (!metric) return;
        metricsByDatasetKey[datasetKey] = metric;
    });

    return metricsByDatasetKey;
}

async function applyProgramExperienceFromShari3ahSurveys(rows) {
    if (!SHARI3AH_SURVEYS_DATA_URL) return { applied: false, reason: 'no-url' };

    const requestUrl = `${SHARI3AH_SURVEYS_DATA_URL}${SHARI3AH_SURVEYS_DATA_URL.includes('?') ? '&' : '?'}t=${Date.now()}`;
    const sourceText = await fetchTextIfExists(requestUrl);
    if (!sourceText) return { applied: false, reason: 'unreachable-source' };

    const payload = parseWindowAssignedJSON(sourceText, 'SURVEYS_DATA');
    if (!payload || !payload.extractedData) return { applied: false, reason: 'invalid-payload' };

    const metricsByDatasetKey = extractProgramExperienceMetricsFromShari3ahSurveys(payload);
    const datasetKeys = Object.keys(metricsByDatasetKey);
    if (!datasetKeys.length) return { applied: false, reason: 'no-program-evaluation-metrics' };

    let appliedRows = 0;
    const matchedDatasets = new Set();

    rows.forEach(row => {
        const programId = getShari3ahSurveysProgramId(row.Major_aName, row.Degree_aName);
        if (!programId) return;

        const datasetKey = `${programId}::${fmtYear(row.Semester)}`;
        const metric = metricsByDatasetKey[datasetKey];
        if (!metric) return;

        row.eval_experience = metric.eval_experience;
        row.eval_experience_sample = metric.eval_experience_sample || 0;
        row.eval_experience_source = 'shari3ah_surveys_program_eval';
        matchedDatasets.add(datasetKey);
        appliedRows++;
    });

    return {
        applied: appliedRows > 0,
        appliedRows,
        matchedDatasets: matchedDatasets.size,
        availableDatasets: datasetKeys.length,
        sourceFile: payload.sourceFile || ''
    };
}

async function applyGraduateSurveyIndicatorsFromSheet(rows, rawUrl, allowedDegrees = null, sourceLabel = '') {
    const csvUrl = resolveGraduateSurveyCsvUrl(rawUrl);
    if (!csvUrl) return { applied: false, reason: 'no-url' };

    const requestUrl = `${csvUrl}${csvUrl.includes('?') ? '&' : '?'}t=${Date.now()}`;
    const csvText = await fetchTextIfExists(requestUrl);
    if (!csvText) return { applied: false, reason: 'unreachable-sheet' };

    const surveyRows = parseCSVQuotedObjects(csvText, ',');
    if (!surveyRows.length) return { applied: false, reason: 'empty-sheet' };

    const surveyAgg = aggregateGraduateSurveyRows(surveyRows, allowedDegrees);
    if (!Object.keys(surveyAgg.metricsByKey).length) {
        return { applied: false, reason: surveyAgg.reason || 'no-metrics', surveyRows: surveyRows.length };
    }

    const applyInfo = applyGraduateSurveyMetrics(rows, surveyAgg.metricsByKey, allowedDegrees, sourceLabel);
    return {
        applied: applyInfo.appliedRows > 0,
        appliedRows: applyInfo.appliedRows,
        groups: surveyAgg.groups,
        matchedRows: surveyAgg.matchedRows,
        reason: applyInfo.appliedRows > 0 ? '' : 'no-target-rows'
    };
}

async function applyGraduateSurveyIndicators(rows) {
    const hasSplitConfig = Boolean(
        GRADUATES_SURVEY_BACHELOR_SHEET_URL || GRADUATES_SURVEY_POSTGRAD_SHEET_URL
    );
    const sources = hasSplitConfig ? [
        {
            url: GRADUATES_SURVEY_BACHELOR_SHEET_URL,
            degrees: BACHELOR_DEGREES,
            label: 'bachelor'
        },
        {
            url: GRADUATES_SURVEY_POSTGRAD_SHEET_URL,
            degrees: POSTGRAD_DEGREES,
            label: 'postgrad'
        }
    ] : [
        {
            url: GRADUATES_SURVEY_SHEET_URL,
            degrees: null,
            label: 'all'
        }
    ];

    const configuredSources = sources.filter(s => String(s.url || '').trim() !== '');
    if (!configuredSources.length) return { applied: false, reason: 'no-url' };

    let appliedRows = 0;
    let groups = 0;
    let matchedRows = 0;
    let sourcesUsed = 0;
    let failures = 0;

    for (const source of configuredSources) {
        const info = await applyGraduateSurveyIndicatorsFromSheet(rows, source.url, source.degrees, source.label);
        if (info.applied) {
            appliedRows += info.appliedRows || 0;
            groups += info.groups || 0;
            matchedRows += info.matchedRows || 0;
            sourcesUsed++;
        } else {
            failures++;
        }
    }

    if (!appliedRows) {
        return { applied: false, reason: failures ? 'all-sources-failed' : 'no-metrics' };
    }
    return { applied: true, appliedRows, groups, matchedRows, sourcesUsed };
}

function parseCitationValue(value, citationsMap = null) {
    if (value == null) return 0;
    const raw = String(value).trim();
    if (!raw) return 0;

    if (citationsMap && Object.prototype.hasOwnProperty.call(citationsMap, raw)) {
        return Number(citationsMap[raw]) || 0;
    }

    const normalized = normalizeArabicDigits(raw)
        .replace(/[()（）]/g, '')
        .replace(/[–—]/g, '-')
        .trim();

    const rangeMatch = normalized.match(/^(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)$/);
    if (rangeMatch) {
        const a = parseFloat(rangeMatch[1]);
        const b = parseFloat(rangeMatch[2]);
        if (Number.isFinite(a) && Number.isFinite(b)) return (a + b) / 2;
    }

    const allNumbers = normalized.match(/\d+(?:\.\d+)?/g);
    if (allNumbers && allNumbers.length === 1) return parseFloat(allNumbers[0]) || 0;
    if (allNumbers && allNumbers.length >= 2) {
        const a = parseFloat(allNumbers[0]);
        const b = parseFloat(allNumbers[1]);
        if (Number.isFinite(a) && Number.isFinite(b)) return (a + b) / 2;
    }
    return 0;
}

function splitIdsPipe(text) {
    return String(text || '')
        .split('|')
        .map(x => x.trim())
        .filter(Boolean);
}

function mergeActivityPublications(baseRows, extraRows) {
    const merged = [...baseRows];
    const seen = new Set(
        baseRows.map(p => `${String(p.title || '').trim()}|${String(p.authors_ids || p.participant_ids || '').trim()}`)
    );
    (extraRows || []).forEach(row => {
        const sig = `${String(row.title || '').trim()}|${String(row.authors_ids || row.participant_ids || '').trim()}`;
        if (!sig || seen.has(sig)) return;
        seen.add(sig);
        merged.push(row);
    });
    return merged;
}

async function fetchActivitySheetsData(apiUrl) {
    if (!apiUrl) return null;
    try {
        const response = await fetch(`${apiUrl}?action=read`, {
            headers: { 'Accept': 'application/json' }
        });
        if (!response.ok) return null;
        const payload = await response.json();
        if (payload && !payload.error) return payload;
        return null;
    } catch {
        return null;
    }
}

async function fetchTextFromSources(paths) {
    for (const p of paths) {
        const text = await fetchTextIfExists(p);
        if (text && text.trim()) return text;
    }
    return null;
}

async function fetchJSONFromSources(paths) {
    for (const p of paths) {
        const obj = await fetchJSONIfExists(p);
        if (obj && typeof obj === 'object') return obj;
    }
    return null;
}

async function applyResearchIndicatorsFromActivities(rows) {
    const years = [...new Set(rows.map(r => absYearFromSemester(r.Semester)))];
    researchActivitySupportByYearDept = {};
    if (!years.length) return { applied: false, reason: 'no-years' };

    const stamp = Date.now();
    const facultyPaths = [
        `data/faculty.csv?t=${stamp}`,
        `${ACTIVITIES_RAW_BASE}/faculty.csv?t=${stamp}`,
    ];
    const publicationsPaths = [
        `${ACTIVITIES_RAW_BASE}/publications.csv?t=${stamp}`,
    ];
    const configPaths = [
        `${ACTIVITIES_RAW_BASE}/config.json?t=${stamp}`,
    ];

    const [facultyText, publicationsText, configObj] = await Promise.all([
        fetchTextFromSources(facultyPaths),
        fetchTextFromSources(publicationsPaths),
        fetchJSONFromSources(configPaths),
    ]);

    if (!facultyText || !publicationsText) {
        return { applied: false, reason: 'missing-faculty-or-publications' };
    }

    const facultyRows = parseFlatCSV(facultyText, ',');
    let publicationsRows = parseFlatCSV(publicationsText, ',');
    if (!facultyRows.length || !publicationsRows.length) {
        return { applied: false, reason: 'empty-faculty-or-publications' };
    }

    availableFacultyBranches = [...new Set([
        ...availableFacultyBranches,
        ...facultyRows.flatMap(row => getFacultyBranchNames(pickCell(row, ['branch', 'Branch', 'الفرع'])))
    ])].sort((a, b) => a.localeCompare(b, 'ar'));

    // مطابقة موقع الأنشطة: دمج البيانات الحية من Google Sheets فوق CSV
    const sheetsPayload = await fetchActivitySheetsData(configObj?.google_sheets_api);
    if (sheetsPayload?.publications?.length) {
        publicationsRows = mergeActivityPublications(publicationsRows, sheetsPayload.publications);
    }

    const citationsMap = (configObj && configObj.citations_ranges) ? configObj.citations_ranges : null;

    // فهارس أعضاء هيئة التدريس لمطابقة منطق موقع الأنشطة
    const authorDeptsByYearId = {};   // year|id -> Set(depts)
    const authorDeptBranchesByYearId = {}; // year|id -> Set(dept|branch)
    const eligibleIdsByYearDept = {}; // year|dept -> Set(ids) (نشط + مؤهل للـ KPI)
    const eligibleIdsByYearDeptBranch = {}; // year|dept|branch -> Set(ids)
    facultyRows.forEach(row => {
        const id = pickCell(row, ['id', 'ID']);
        const year = parseInt(pickCell(row, ['year', 'Year']), 10);
        if (!id || !year) return;

        const active = pickCell(row, ['active', 'Active']) === 'نعم';
        const dept = normalizeDepartment(pickCell(row, ['department', 'Department']));
        const rank = normalizeRank(pickCell(row, ['rank', 'Rank']));
        const branches = getFacultyBranchNames(pickCell(row, ['branch', 'Branch', 'الفرع']));

        if (dept) {
            const authorYearKey = `${year}|${id}`;
            if (!authorDeptsByYearId[authorYearKey]) authorDeptsByYearId[authorYearKey] = new Set();
            authorDeptsByYearId[authorYearKey].add(dept);
            branches.forEach(branch => {
                const deptBranchKey = `${dept}|${branch}`;
                if (!authorDeptBranchesByYearId[authorYearKey]) authorDeptBranchesByYearId[authorYearKey] = new Set();
                authorDeptBranchesByYearId[authorYearKey].add(deptBranchKey);
            });
        }
        if (!active || !dept) return;

        if (!RESEARCH_KPI_EXCLUDED_RANKS.has(rank)) {
            const key = buildYearDeptKey(year, dept);
            if (!eligibleIdsByYearDept[key]) eligibleIdsByYearDept[key] = new Set();
            eligibleIdsByYearDept[key].add(id);
            branches.forEach(branch => {
                const branchKey = buildYearDeptBranchKey(year, dept, branch);
                if (!eligibleIdsByYearDeptBranch[branchKey]) eligibleIdsByYearDeptBranch[branchKey] = new Set();
                eligibleIdsByYearDeptBranch[branchKey].add(id);
            });
        }
    });

    // aggregates per year+dept
    const publishingMembersByYearDept = {};
    const publicationsCountByYearDept = {};
    const citationsTotalByYearDept = {};
    const publishingMembersByYearDeptBranch = {};
    const publicationsCountByYearDeptBranch = {};
    const citationsTotalByYearDeptBranch = {};

    publicationsRows.forEach(pub => {
        const year = parseInt(pickCell(pub, ['year', 'Year']), 10);
        if (!year) return;
        if (!years.includes(year)) return;

        const authorIds = splitIdsPipe(pickCell(pub, ['authors_ids', 'participant_ids']));
        if (!authorIds.length) return;

        const departmentsTouched = new Set();
        const deptBranchesTouched = new Set();
        authorIds.forEach(fid => {
            const authorYearKey = `${year}|${fid}`;
            const depts = authorDeptsByYearId[authorYearKey];
            if (!depts) return;
            depts.forEach(d => departmentsTouched.add(d));
            const deptBranches = authorDeptBranchesByYearId[authorYearKey];
            if (deptBranches) deptBranches.forEach(key => deptBranchesTouched.add(key));
        });

        if (!departmentsTouched.size && !deptBranchesTouched.size) return;
        const citations = parseCitationValue(pickCell(pub, ['citations_range', 'Citations', 'citations']), citationsMap);
        departmentsTouched.forEach(dept => {
            const ydKey = buildYearDeptKey(year, dept);
            publicationsCountByYearDept[ydKey] = (publicationsCountByYearDept[ydKey] || 0) + 1;
            citationsTotalByYearDept[ydKey] = (citationsTotalByYearDept[ydKey] || 0) + citations;

            const eligible = eligibleIdsByYearDept[ydKey];
            if (!eligible || !eligible.size) return;
            if (!publishingMembersByYearDept[ydKey]) publishingMembersByYearDept[ydKey] = new Set();
            authorIds.forEach(fid => {
                if (eligible.has(fid)) publishingMembersByYearDept[ydKey].add(fid);
            });
        });
        deptBranchesTouched.forEach(deptBranchKey => {
            const [dept, ...branchParts] = deptBranchKey.split('|');
            const branch = branchParts.join('|');
            const ydbKey = buildYearDeptBranchKey(year, dept, branch);
            publicationsCountByYearDeptBranch[ydbKey] = (publicationsCountByYearDeptBranch[ydbKey] || 0) + 1;
            citationsTotalByYearDeptBranch[ydbKey] = (citationsTotalByYearDeptBranch[ydbKey] || 0) + citations;

            const eligible = eligibleIdsByYearDeptBranch[ydbKey];
            if (!eligible || !eligible.size) return;
            if (!publishingMembersByYearDeptBranch[ydbKey]) publishingMembersByYearDeptBranch[ydbKey] = new Set();
            authorIds.forEach(fid => {
                if (eligible.has(fid)) publishingMembersByYearDeptBranch[ydbKey].add(fid);
            });
        });
    });

    let appliedRows = 0;
    rows.forEach(r => {
        const year = absYearFromSemester(r.Semester);
        const dept = normalizeDepartment(r.Dept_aName);
        const ydKey = buildYearDeptKey(year, dept);

        const eligibleSet = eligibleIdsByYearDept[ydKey];
        if (!eligibleSet || !eligibleSet.size) return;

        const facultyTotal = eligibleSet.size; // مطابق لـ totalEligibleMembers في موقع الأنشطة
        const facultyPublished = (publishingMembersByYearDept[ydKey] || new Set()).size;
        const publicationsCount = publicationsCountByYearDept[ydKey] || 0;
        const citationsTotal = citationsTotalByYearDept[ydKey] || 0;
        const citationsPerPublication = publicationsCount > 0 ? (citationsTotal / publicationsCount) : 0;

        r.faculty_total = facultyTotal;
        r.faculty_published = facultyPublished;
        r.research_count = publicationsCount;
        r.citations = Math.round(citationsTotal * 10) / 10;
        r.citations_per_publication = Math.round(citationsPerPublication * 10) / 10;
        r.research_source = 'faculty_activities_live';
        appliedRows++;
    });

    Object.entries(eligibleIdsByYearDept).forEach(([ydKey, eligibleSet]) => {
        if (!eligibleSet?.size) return;
        const aggregate = createResearchActivityAggregate();
        const publicationsCount = publicationsCountByYearDept[ydKey] || 0;
        const citationsTotal = citationsTotalByYearDept[ydKey] || 0;
        aggregate.faculty_total = eligibleSet.size;
        aggregate.faculty_published = (publishingMembersByYearDept[ydKey] || new Set()).size;
        aggregate.research_count = publicationsCount;
        aggregate.citations = Math.round(citationsTotal * 10) / 10;
        aggregate.citations_per_publication = publicationsCount > 0
            ? Math.round((citationsTotal / publicationsCount) * 10) / 10 : 0;
        researchActivitySupportByYearDept[ydKey] = { ...aggregate, byBranch: {} };
    });

    Object.entries(eligibleIdsByYearDeptBranch).forEach(([ydbKey, eligibleSet]) => {
        if (!eligibleSet?.size) return;
        const [yearStr, dept, ...branchParts] = ydbKey.split('|');
        const branch = branchParts.join('|');
        const ydKey = `${yearStr}|${dept}`;
        if (!researchActivitySupportByYearDept[ydKey]) {
            researchActivitySupportByYearDept[ydKey] = { ...createResearchActivityAggregate(), byBranch: {} };
        }
        const aggregate = createResearchActivityAggregate();
        const publicationsCount = publicationsCountByYearDeptBranch[ydbKey] || 0;
        const citationsTotal = citationsTotalByYearDeptBranch[ydbKey] || 0;
        aggregate.faculty_total = eligibleSet.size;
        aggregate.faculty_published = (publishingMembersByYearDeptBranch[ydbKey] || new Set()).size;
        aggregate.research_count = publicationsCount;
        aggregate.citations = Math.round(citationsTotal * 10) / 10;
        aggregate.citations_per_publication = publicationsCount > 0
            ? Math.round((citationsTotal / publicationsCount) * 10) / 10 : 0;
        researchActivitySupportByYearDept[ydKey].byBranch[branch] = aggregate;
    });

    return {
        applied: appliedRows > 0,
        appliedRows,
        branches: availableFacultyBranches.length
    };
}

function resolveFacultyProfile(facultyProfilesByYearId, facultyProfilesById, year, fid) {
    const direct = facultyProfilesByYearId[`${year}|${fid}`];
    if (direct) return direct;
    const list = facultyProfilesById[fid];
    if (!list || !list.length) return null;
    let best = null;
    list.forEach(p => {
        if (p.year <= year && (!best || p.year > best.year)) best = p;
    });
    return best || list[list.length - 1];
}

function getRankAllowedDegrees(rank) {
    return RANK_ALLOWED_DEGREES[normalizeRank(rank)] || null;
}

function getRankBaseFTE(rank) {
    const normalized = normalizeRank(rank);
    return RANK_BASE_FTE[normalized] || 1;
}

function getDeptProgramKeys(programsByDeptYear, year, dept) {
    return (programsByDeptYear[`${year}|${dept}`] || []).map(p => p.key);
}

function getDeptDegreeProgramKeys(programsByDeptYear, year, dept, degree) {
    return (programsByDeptYear[`${year}|${dept}`] || [])
        .filter(p => p.degree === degree)
        .map(p => p.key);
}

async function applyTeachingBasedFacultyFTE(rows) {
    const years = [...new Set(rows.map(r => absYearFromSemester(r.Semester)))].sort((a,b) => a - b);
    if (!years.length) return { applied: false, reason: 'no-years' };

    const stamp = Date.now();
    const [plansText, facultyText, teachingMeta] = await Promise.all([
        fetchTextIfExists(`data/new_all_plans.csv?t=${stamp}`),
        fetchTextIfExists(`data/faculty.csv?t=${stamp}`),
        fetchJSONIfExists(`data/teaching/meta.json?t=${stamp}`)
    ]);

    if (!plansText || !facultyText) {
        return { applied: false, reason: 'missing-plans-or-faculty' };
    }

    const planRows = parseFlatCSV(plansText, ';');
    const facultyRows = parseFlatCSV(facultyText, ',');
    const teachingFacultyIndex = teachingMeta?.faculty_index || {};
    if (!planRows.length || !facultyRows.length) {
        return { applied: false, reason: 'empty-plans-or-faculty' };
    }

    availableFacultyBranches = [...new Set([
        ...availableFacultyBranches,
        ...facultyRows.flatMap(row => getFacultyBranchNames(pickCell(row, ['branch', 'Branch', 'الفرع'])))
    ])].sort((a, b) => a.localeCompare(b, 'ar'));

    const yearPayloads = await Promise.all(
        years.map(y => fetchJSONIfExists(`data/teaching/years/${y}.json?t=${stamp}`))
    );

    const teachingByYear = {};
    years.forEach((y, idx) => {
        const payload = yearPayloads[idx];
        if (!payload) return;
        const records = Array.isArray(payload) ? payload : payload.records;
        if (Array.isArray(records) && records.length) teachingByYear[y] = records;
    });

    // برنامج/سنة من data.csv
    const studentsByProgramKey = {};
    const programMetaByKey = {};
    const programLookupByYearMajorDegree = {};
    const programsByDeptYear = {};
    rows.forEach(r => {
        const year = absYearFromSemester(r.Semester);
        const dept = normalizeDepartment(r.Dept_aName);
        const major = normalizeSurveyProgramName(r.Major_aName);
        const degree = normalizeDegree(r.Degree_aName);
        if (!major || !SUPPORTED_KPI_DEGREES.has(degree)) return;

        const pKey = buildProgramDataKey(year, dept, major, degree);
        studentsByProgramKey[pKey] = Number(r.students_total) || 0;
        programMetaByKey[pKey] = { year, dept, major, degree };

        const mdKey = `${year}|${buildProgramMajorDegreeKey(major, degree)}`;
        if (!programLookupByYearMajorDegree[mdKey]) programLookupByYearMajorDegree[mdKey] = [];
        if (!programLookupByYearMajorDegree[mdKey].includes(pKey)) {
            programLookupByYearMajorDegree[mdKey].push(pKey);
        }

        const dyKey = `${year}|${dept}`;
        if (!programsByDeptYear[dyKey]) programsByDeptYear[dyKey] = [];
        if (!programsByDeptYear[dyKey].some(x => x.key === pKey)) {
            programsByDeptYear[dyKey].push({
                key: pKey,
                degree,
                students: Number(r.students_total) || 0
            });
        }
    });

    // خريطة المقرر -> البرامج
    const courseToPrograms = {};
    const programExclusiveCodes = {};
    const programNonSharedCodes = {};
    planRows.forEach(row => {
        const code = pickCell(row, ['Code', '\uFEFFCode', 'رمز المقرر']);
        const major = normalizeSurveyProgramName(pickCell(row, ['Program', '\uFEFFProgram', 'البرنامج']));
        const degree = normalizeDegree(pickCell(row, ['Degree', 'الدرجة']));
        const courseType = pickCell(row, ['نوع المقرر']);
        if (!code || !major || !SUPPORTED_KPI_DEGREES.has(degree)) return;
        const planKey = buildProgramMajorDegreeKey(major, degree);
        if (!courseToPrograms[code]) courseToPrograms[code] = [];
        if (!courseToPrograms[code].some(x => x.planKey === planKey)) {
            courseToPrograms[code].push({ major, degree, planKey });
        }

        if (courseType.includes('فريد')) {
            if (!programExclusiveCodes[planKey]) programExclusiveCodes[planKey] = new Set();
            programExclusiveCodes[planKey].add(code);
        }
        if (!courseType.includes('مشترك')) {
            if (!programNonSharedCodes[planKey]) programNonSharedCodes[planKey] = new Set();
            programNonSharedCodes[planKey].add(code);
        }
    });

    Object.entries(courseToPrograms).forEach(([code, mappings]) => {
        if (mappings.length !== 1) return;
        const planKey = mappings[0].planKey;
        if (!programExclusiveCodes[planKey]) programExclusiveCodes[planKey] = new Set();
        if (!programNonSharedCodes[planKey]) programNonSharedCodes[planKey] = new Set();
        programExclusiveCodes[planKey].add(code);
        programNonSharedCodes[planKey].add(code);
    });

    // فهرس أعضاء هيئة التدريس (للـ fallback)
    const facultyProfilesByYearId = {};
    const facultyProfilesById = {};
    const activeFacultyByYear = {};
    facultyRows.forEach(row => {
        const id = pickCell(row, ['id', 'ID']);
        const yearRaw = pickCell(row, ['year', 'Year']);
        const year = parseInt(yearRaw, 10);
        if (!id || !year) return;
        const profile = {
            id,
            year,
            rank: normalizeRank(pickCell(row, ['rank', 'Rank'])),
            dept: normalizeDepartment(pickCell(row, ['department', 'Department'])),
            active: pickCell(row, ['active', 'Active']) === 'نعم',
            nationality: String(pickCell(row, ['nationality', 'Nationality', 'الجنسية']) || '').trim(),
            gender: String(pickCell(row, ['gender', 'Gender', 'الجنس']) || '').trim(),
            branch: String(pickCell(row, ['branch', 'Branch', 'الفرع']) || '').trim()
        };

        facultyProfilesByYearId[`${year}|${id}`] = profile;
        if (!facultyProfilesById[id]) facultyProfilesById[id] = [];
        facultyProfilesById[id].push(profile);

        if (profile.active) {
            if (!activeFacultyByYear[year]) activeFacultyByYear[year] = [];
            if (!activeFacultyByYear[year].some(x => x.id === id)) {
                activeFacultyByYear[year].push(profile);
            }
        }
    });
    Object.values(facultyProfilesById).forEach(list => list.sort((a,b) => a.year - b.year));

    const facultyProgramLoads = {}; // year|fid -> {programKey: weighted_load}
    const facultyProgramBranchLoads = {}; // year|fid -> {programKey: {branch: weighted_load}}
    const teachingRankByFaculty = {}; // year|fid -> rank from the same-year roster or teaching metadata
    const supportByProgramKey = {};
    programFacultyFteByBranchKey = {};

    function getSupportBucket(pKey) {
        if (!supportByProgramKey[pKey]) {
            supportByProgramKey[pKey] = createRawTeachingSupportBucket(true);
        }
        return supportByProgramKey[pKey];
    }

    // 1) تحميل تدريسي فعلي من ملفات teaching/years/*.json
    Object.entries(teachingByYear).forEach(([yearStr, records]) => {
        const year = parseInt(yearStr, 10);
        records.forEach(rec => {
            const fid = String(rec.fid || '').trim();
            if (!fid) return;
            const facKey = `${year}|${fid}`;
            const directProfile = facultyProfilesByYearId[facKey];
            const profile = resolveFacultyProfile(facultyProfilesByYearId, facultyProfilesById, year, fid);
            const teachingMetaProfile = teachingFacultyIndex[fid] || {};
            const teachingRank = normalizeRank(directProfile?.rank || teachingMetaProfile.r || profile?.rank || '');
            const deptHint = normalizeDepartment(directProfile?.dept || teachingMetaProfile.d || profile?.dept || '');
            teachingRankByFaculty[facKey] = teachingRank;
            if (!facultyProgramLoads[facKey]) facultyProgramLoads[facKey] = {};

            const courses = Array.isArray(rec.cs) ? rec.cs : [];
            courses.forEach(c => {
                const secDegree = normalizeDegree(c.dg);
                if (!SUPPORTED_KPI_DEGREES.has(secDegree)) return;

                const loadHours = Number(c.h);
                const load = Number.isFinite(loadHours) && loadHours > 0 ? loadHours : 1;
                const code = String(c.cc || '').trim();

                let candidateProgramKeys = [];
                let supportCandidates = [];

                // محاولة ربط مباشر من رمز المقرر
                const mapped = courseToPrograms[code] || [];
                mapped.forEach(mp => {
                    if (mp.degree !== secDegree) return;
                    const mdKey = `${year}|${buildProgramMajorDegreeKey(mp.major, mp.degree)}`;
                    const options = programLookupByYearMajorDegree[mdKey] || [];
                    if (!options.length) return;
                    if (options.length === 1) {
                        candidateProgramKeys.push(options[0]);
                        supportCandidates.push({ key: options[0], planKey: mp.planKey });
                        return;
                    }
                    const deptMatched = deptHint ? options.filter(k => k.split('|')[1] === deptHint) : [];
                    (deptMatched.length ? deptMatched : options).forEach(key => {
                        candidateProgramKeys.push(key);
                        supportCandidates.push({ key, planKey: mp.planKey });
                    });
                });

                // fallback: لو الرمز غير موجود في الخطط نوزع داخل القسم/الدرجة
                candidateProgramKeys = [...new Set(candidateProgramKeys)];
                if (!candidateProgramKeys.length && deptHint) {
                    candidateProgramKeys = getDeptDegreeProgramKeys(programsByDeptYear, year, deptHint, secDegree);
                }
                if (!candidateProgramKeys.length) return;

                const profileBranches = getFacultyBranchNames(profile?.branch || '');
                const sectionBranch = getBranchFromTeachingLocation(c.l);
                const allocationBranch = sectionBranch || (profileBranches.length === 1 ? profileBranches[0] : '');
                const weights = buildWeights(candidateProgramKeys, studentsByProgramKey);
                weights.forEach(w => {
                    const weightedLoad = load * w.weight;
                    facultyProgramLoads[facKey][w.key] = (facultyProgramLoads[facKey][w.key] || 0) + weightedLoad;
                    if (allocationBranch) {
                        if (!facultyProgramBranchLoads[facKey]) facultyProgramBranchLoads[facKey] = {};
                        if (!facultyProgramBranchLoads[facKey][w.key]) facultyProgramBranchLoads[facKey][w.key] = {};
                        const branchLoads = facultyProgramBranchLoads[facKey][w.key];
                        branchLoads[allocationBranch] = (branchLoads[allocationBranch] || 0) + weightedLoad;
                    }
                });

                const supportMap = new Map();
                supportCandidates.forEach(item => {
                    if (!item?.key || !item?.planKey) return;
                    if (!supportMap.has(item.key)) supportMap.set(item.key, item.planKey);
                });

                const sectionStudents = Number(c.e) || 0;
                const sectionGender = detectSectionGender(c);
                const supportBranchNames = allocationBranch ? [allocationBranch] : [];
                supportMap.forEach((planKey, pKey) => {
                    const bucket = getSupportBucket(pKey);
                    const exclusiveSet = programExclusiveCodes[planKey];
                    const nonSharedSet = programNonSharedCodes[planKey];
                    const isExclusive = Boolean(exclusiveSet && exclusiveSet.has(code));
                    const isNonShared = Boolean(nonSharedSet && nonSharedSet.has(code));
                    const targetBuckets = [bucket];
                    supportBranchNames.forEach(branchName => {
                        const branchBucket = getRawTeachingSupportBranchBucket(bucket, branchName);
                        if (branchBucket) targetBuckets.push(branchBucket);
                    });

                    targetBuckets.forEach(targetBucket => {
                        targetBucket.totalSections++;
                        targetBucket.totalStudents += sectionStudents;

                        if (sectionGender === 'female') {
                            targetBucket.femaleSections++;
                            targetBucket.femaleStudents += sectionStudents;
                        } else {
                            targetBucket.maleSections++;
                            targetBucket.maleStudents += sectionStudents;
                        }

                        if (isExclusive) {
                            targetBucket.exclusiveSections++;
                            targetBucket.exclusiveStudents += sectionStudents;
                            if (sectionGender === 'female') {
                                targetBucket.exclusiveFemaleSections++;
                                targetBucket.exclusiveFemaleStudents += sectionStudents;
                            } else {
                                targetBucket.exclusiveMaleSections++;
                                targetBucket.exclusiveMaleStudents += sectionStudents;
                            }
                        }

                        if (!targetBucket.facultyById[fid]) {
                            targetBucket.facultyById[fid] = {
                                rank: teachingRank,
                                nationality: String(profile?.nationality || '').trim(),
                                gender: String(profile?.gender || '').trim(),
                                branch: String(profile?.branch || '').trim(),
                                totalHours: 0,
                                nonSharedHours: 0,
                                exclusiveHours: 0,
                                maleHours: 0,
                                femaleHours: 0,
                                nonSharedMaleHours: 0,
                                nonSharedFemaleHours: 0,
                                exclusiveMaleHours: 0,
                                exclusiveFemaleHours: 0,
                                sections: 0
                            };
                        }
                        const facultyEntry = targetBucket.facultyById[fid];
                        facultyEntry.rank = facultyEntry.rank || teachingRank;
                        if (!facultyEntry.nationality && profile?.nationality) {
                            facultyEntry.nationality = String(profile.nationality).trim();
                        }
                        if (!facultyEntry.gender && profile?.gender) {
                            facultyEntry.gender = String(profile.gender).trim();
                        }
                        if (!facultyEntry.branch && profile?.branch) {
                            facultyEntry.branch = String(profile.branch).trim();
                        }
                        facultyEntry.totalHours += load;
                        facultyEntry.sections++;
                        if (sectionGender === 'female') {
                            facultyEntry.femaleHours += load;
                        } else {
                            facultyEntry.maleHours += load;
                        }
                        if (isNonShared) {
                            facultyEntry.nonSharedHours += load;
                            if (sectionGender === 'female') facultyEntry.nonSharedFemaleHours += load;
                            else facultyEntry.nonSharedMaleHours += load;
                        }
                        if (isExclusive) {
                            facultyEntry.exclusiveHours += load;
                            if (sectionGender === 'female') facultyEntry.exclusiveFemaleHours += load;
                            else facultyEntry.exclusiveMaleHours += load;
                        }
                    });
                });
            });
        });
    });

    // نحول أحمال كل عضو إلى FTE = 1 موزعة على البرامج حسب نسب الأحمال
    const fteByProgramKey = {};
    const facultyWithActualTeaching = new Set();
    Object.entries(facultyProgramLoads).forEach(([facKey, byProgram]) => {
        const totalLoad = Object.values(byProgram).reduce((s, x) => s + x, 0);
        if (totalLoad <= 0) return;
        facultyWithActualTeaching.add(facKey);
        const [yearStr, fid] = facKey.split('|');
        const year = parseInt(yearStr, 10);
        const profile = resolveFacultyProfile(facultyProfilesByYearId, facultyProfilesById, year, fid);
        const baseFTE = getRankBaseFTE(teachingRankByFaculty[facKey] || profile?.rank);
        const branchNames = getFacultyBranchNames(profile?.branch || '');
        Object.entries(byProgram).forEach(([pKey, load]) => {
            const allocated = (load / totalLoad) * baseFTE;
            fteByProgramKey[pKey] = (fteByProgramKey[pKey] || 0) + allocated;
            const observedBranchLoads = facultyProgramBranchLoads[facKey]?.[pKey] || {};
            const observedLoad = Object.values(observedBranchLoads).reduce((sum, value) => sum + value, 0);
            const branchWeights = observedLoad > 0
                ? Object.entries(observedBranchLoads).map(([branch, branchLoad]) => ({
                    branch,
                    weight: branchLoad / load
                }))
                : branchNames.map(branch => ({ branch, weight: 1 / branchNames.length }));
            branchWeights.forEach(({ branch: branchName, weight }) => {
                const branchKey = buildProgramBranchDataKey(pKey, branchName);
                if (branchKey) {
                    programFacultyFteByBranchKey[branchKey] = (programFacultyFteByBranchKey[branchKey] || 0) + (allocated * weight);
                }
            });
        });
    });

    // 2) fallback: الأعضاء النشطون الذين لا يوجد لهم تدريس فعلي
    years.forEach(year => {
        const activeFaculty = activeFacultyByYear[year] || [];
        activeFaculty.forEach(profile => {
            const facKey = `${year}|${profile.id}`;
            if (facultyWithActualTeaching.has(facKey)) return;
            if (!profile.dept) return;

            const allowedDegrees = getRankAllowedDegrees(profile.rank);
            let candidateKeys = getDeptProgramKeys(programsByDeptYear, year, profile.dept);
            if (allowedDegrees) {
                candidateKeys = candidateKeys.filter(k => allowedDegrees.has(programMetaByKey[k]?.degree));
            }
            if (!candidateKeys.length) return;

            const baseFTE = getRankBaseFTE(profile.rank);
            const branchNames = getFacultyBranchNames(profile.branch || '');
            const weights = buildWeights(candidateKeys, studentsByProgramKey);
            weights.forEach(w => {
                const allocated = w.weight * baseFTE;
                fteByProgramKey[w.key] = (fteByProgramKey[w.key] || 0) + allocated;
                const branchAllocated = branchNames.length ? allocated / branchNames.length : 0;
                branchNames.forEach(branchName => {
                    const branchKey = buildProgramBranchDataKey(w.key, branchName);
                    if (branchKey) {
                        programFacultyFteByBranchKey[branchKey] = (programFacultyFteByBranchKey[branchKey] || 0) + branchAllocated;
                    }
                });
            });
        });
    });

    const finalizedSupportByProgramKey = {};
    Object.entries(supportByProgramKey).forEach(([pKey, bucket]) => {
        const finalizedBucket = finalizeTeachingSupportBucket(bucket) || {};
        const byBranch = {};
        Object.entries(bucket.branchBuckets || {}).forEach(([branch, branchBucket]) => {
            const finalizedBranch = finalizeTeachingSupportBucket(branchBucket);
            if (finalizedBranch) byBranch[branch] = finalizedBranch;
        });
        finalizedSupportByProgramKey[pKey] = {
            ...finalizedBucket,
            byBranch
        };
    });

    teachingProgramSupportByKey = finalizedSupportByProgramKey;

    let programsWithComputedFTE = 0;
    rows.forEach(r => {
        const year = absYearFromSemester(r.Semester);
        const dept = normalizeDepartment(r.Dept_aName);
        const major = normalizeSurveyProgramName(r.Major_aName);
        const degree = normalizeDegree(r.Degree_aName);
        const pKey = buildProgramDataKey(year, dept, major, degree);
        const computed = Number(fteByProgramKey[pKey]) || 0;

        if (computed > 0) {
            r.faculty_ratio_base = Math.round(computed * 100) / 100;
            r.faculty_ratio_source = 'teaching_fte';
            programsWithComputedFTE++;
        } else if (r.faculty_total > 0) {
            r.faculty_ratio_base = r.faculty_total;
            r.faculty_ratio_source = 'csv';
        } else {
            r.faculty_ratio_base = 0;
            r.faculty_ratio_source = 'none';
        }
    });

    return {
        applied: programsWithComputedFTE > 0,
        programsWithComputedFTE,
        yearsWithTeaching: Object.keys(teachingByYear).length,
        programsWithSupport: Object.keys(teachingProgramSupportByKey).length
    };
}

// ========================================
// تحميل وتحليل البيانات
// ========================================
async function loadIslamicBranchData() {
    const payload = await fetchJSONIfExists(`data/islamic_branch_students.json?t=${Date.now()}`);
    if (!payload || typeof payload !== 'object') {
        islamicBranchData = null;
        entrantData = [];
        return { applied: false, reason: 'missing-islamic-branch-data' };
    }
    if (!payload.metrics || !Array.isArray(payload.entrants) || !Array.isArray(payload.graduates)) {
        islamicBranchData = null;
        entrantData = [];
        return { applied: false, reason: 'invalid-islamic-branch-data' };
    }
    islamicBranchData = payload;
    entrantData = payload.entrants
        .filter(row => ['official', 'very_high', 'high'].includes(String(row.confidence || '')))
        .map(row => ({ ...row }));
    availableFacultyBranches = [...new Set([
        ...availableFacultyBranches,
        ...getIslamicBranchNames()
    ])].sort((a, b) => a.localeCompare(b, 'ar'));
    return {
        applied: true,
        entrants: entrantData.length,
        graduates: payload.graduates.length,
        branches: getIslamicBranchNames().length
    };
}

async function loadData() {
    try {
        const res = await fetch('data/data.csv?t=' + Date.now());
        const csv = await res.text();
        allRows = parseCSV(csv);
        const branchInfo = await loadIslamicBranchData();
        const durationInfo = await applyAverageGraduationDurationFromDetails(allRows);
        const experienceInfo = await applyProgramExperienceFromShari3ahSurveys(allRows);
        const surveyInfo = await applyGraduateSurveyIndicators(allRows);
        const researchInfo = await applyResearchIndicatorsFromActivities(allRows);
        const fteInfo = await applyTeachingBasedFacultyFTE(allRows);
        programs = buildPrograms(allRows);
        console.info('KPI data loaded', {
            programs: programs.length,
            rows: allRows.length,
            durationInfo,
            surveyInfo,
            experienceInfo,
            researchInfo,
            fteInfo,
            branchInfo,
        });
        return true;
    } catch (e) {
        console.error(e);
        return false;
    }
}

function parseCSV(text) {
    const lines = text.trim().split('\n').map(l => l.replace(/\r/g,''));
    if (lines.length < 2) return [];
    const sep = lines[0].includes(';') ? ';' : ',';
    const headers = lines[0].split(sep).map(h => h.trim().replace(/^\uFEFF/,''));
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        const vals = lines[i].split(sep);
        if (vals.length < 5) continue;
        const row = {};
        headers.forEach((h, j) => { row[h] = (vals[j] || '').trim(); });
        // convert numbers
        const numFields = [
            'students_total','students_male','students_female','students_saudi','students_international',
            'students_new','students_retained','graduates_total','graduates_ontime',
            'prev_new_count','new_4_ago_count',
            'sections_total','sections_male','sections_female',
            'faculty_total','faculty_phd','faculty_male','faculty_female','faculty_published',
            'research_count','citations','citations_per_publication'
        ];
        numFields.forEach(f => { row[f] = parseFloat(row[f]) || 0; });

        const optionalMetricFields = [
            'eval_courses','eval_experience','eval_supervision','eval_services','eval_employers',
            'performance_rate','employment_rate','avg_time_to_graduate'
        ];
        optionalMetricFields.forEach(f => {
            const raw = String(row[f] || '').trim();
            if (raw === '') {
                row[f] = null;
                return;
            }
            const parsed = parseFloat(raw);
            row[f] = Number.isFinite(parsed) ? parsed : null;
        });

        row.Semester = parseInt(row.Semester) || 0;
        // skip year 38 (base year only)
        if (row.Semester === 38) continue;
        rows.push(row);
    }
    return rows;
}

function buildPrograms(rows) {
    const map = {};
    rows.forEach(r => {
        const key = `${r.Major_aName}|${r.Degree_aName}`;
        if (!map[key]) {
            map[key] = { name: r.Major_aName, degree: r.Degree_aName, dept: r.Dept_aName, years: {} };
        }
        map[key].years[r.Semester] = r;
    });
    return Object.values(map).sort((a,b) => {
        const degOrder = {'بكالوريوس':0,'الماجستير':1,'دكتوراه':2};
        const da = degOrder[a.degree] ?? 9, db = degOrder[b.degree] ?? 9;
        if (da !== db) return da - db;
        return a.name.localeCompare(b.name, 'ar');
    });
}

function getYearData(year) {
    return programs.map(p => ({prog: p, data: p.years[year]})).filter(x => x.data);
}

function getAvailableYears() {
    const s = new Set();
    programs.forEach(p => Object.keys(p.years).forEach(y => s.add(parseInt(y))));
    return [...s].filter(y => DISPLAY_YEARS.includes(y)).sort((a,b)=>a-b);
}

// ========================================
// حساب المؤشرات
// ========================================
function calcKPIs(d, degree) {
    const kpi = {};
    kpi.experience_eval = d.eval_experience ?? null;
    kpi.course_eval = d.eval_courses ?? null;
    kpi.supervision_eval = d.eval_supervision ?? null;
    kpi.services_satisfaction = d.eval_services ?? null;
    kpi.employer_eval = d.eval_employers ?? null;
    kpi.student_performance = d.performance_rate ?? null;
    kpi.employment_rate = d.employment_rate ?? null;

    // معدل التخرج بالوقت المحدد
    kpi.graduation_rate = pct(d.graduates_ontime, d.new_4_ago_count);
    kpi.graduation_detail = d.new_4_ago_count > 0 ? `${d.graduates_ontime} من ${d.new_4_ago_count}` : null;

    // معدل الاستبقاء
    kpi.retention_rate = pct(d.students_retained, d.prev_new_count);
    kpi.retention_detail = d.prev_new_count > 0 ? `${d.students_retained} من ${d.prev_new_count}` : null;

    // مؤشرات خاصة بالدراسات العليا
    kpi.avg_time_to_graduate = d.avg_time_to_graduate ?? null;
    kpi.dropout_rate = kpi.retention_rate != null
        ? Math.round((100 - kpi.retention_rate) * 10) / 10 : null;

    // نسبة الطلاب/هيئة التدريس
    const facultyBase = getFacultyBaseForRatio(d);
    if (facultyBase > 0 && d.students_total > 0) {
        kpi.student_faculty_ratio = `1:${(d.students_total / facultyBase).toFixed(1)}`;
    } else {
        kpi.student_faculty_ratio = null;
    }

    // نسبة النشر العلمي (مطابقة لموقع الأنشطة: عدد الأعضاء الناشرين ÷ الأعضاء المؤهلين)
    kpi.publication_pct = d.faculty_total > 0
        ? pct(d.faculty_published, d.faculty_total) : null;

    // البحوث/عضو (مطابقة لموقع الأنشطة)
    kpi.research_per_faculty = d.faculty_total > 0
        ? Math.round((d.research_count / d.faculty_total) * 100) / 100 : null;

    // متوسط الاقتباسات لكل بحث (مطابقة لموقع الأنشطة)
    if (String(d.research_source || '').startsWith('faculty_activities_live') && Number.isFinite(d.citations_per_publication)) {
        kpi.citations_per_faculty = Math.round(d.citations_per_publication * 10) / 10;
    } else {
        kpi.citations_per_faculty = d.faculty_total > 0
            ? Math.round((d.citations / d.faculty_total) * 10) / 10 : null;
    }

    kpi.student_publication = null;
    kpi.patents = null;

    return kpi;
}

function fmtKPI(val, unit) {
    if (val == null) return { text: 'غير متوفر', cls: 'na' };
    if (unit === '%') return { text: val.toFixed(1) + '%', cls: '' };
    if (unit === 'درجة') return { text: parseFloat(val).toFixed(2), cls: '' };
    if (unit === 'سنة') return { text: parseFloat(val).toFixed(2), cls: '' };
    return { text: String(val), cls: '' };
}

function getSurveyEvidence(d, indicatorKey) {
    const keyMap = {
        experience_eval: { count: 'eval_experience_sample', source: 'eval_experience_source' },
        course_eval: { count: 'eval_courses_sample', source: 'eval_courses_source' },
        supervision_eval: { count: 'eval_supervision_sample', source: 'eval_supervision_source' },
        services_satisfaction: { count: 'eval_services_sample', source: 'eval_services_source' },
        student_performance: { count: 'performance_rate_sample', source: 'performance_rate_source' },
        employment_rate: { count: 'employment_rate_sample', source: 'employment_rate_source' },
        employer_eval: { count: 'eval_employers_sample', source: 'eval_employers_source' },
    };
    const keys = keyMap[indicatorKey];
    if (!keys) return null;

    const countValue = Number(d[keys.count]);
    const count = Number.isFinite(countValue) && countValue > 0 ? Math.round(countValue) : 0;
    const source = String(d[keys.source] || '');

    if (source.startsWith('graduates_sample_')) {
        return { kind: 'sample', label: 'استطلاع عينة', count };
    }
    return null;
}

function formatSurveyEvidenceText(evidence) {
    if (!evidence) return '';
    return evidence.count > 0
        ? `${evidence.label} - عدد المشاركين: ${fmtNum(evidence.count)}`
        : evidence.label;
}

function getTeachingSupportForProgramYear(prog, year, branch = ALL_BRANCH_FILTER_VALUE) {
    const pKey = buildProgramDataKey(year, prog?.dept, prog?.name, prog?.degree);
    if (!pKey) return null;
    const support = teachingProgramSupportByKey[pKey] || null;
    if (!support) return null;
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch || normalizedBranch === ALL_BRANCH_FILTER_VALUE) return support;
    return support.byBranch?.[normalizedBranch] || null;
}

function getProgramFacultyBaseForBranch(prog, year, branch = ALL_BRANCH_FILTER_VALUE) {
    const dataRow = prog?.years?.[year] || null;
    if (!dataRow) return 0;
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch || normalizedBranch === ALL_BRANCH_FILTER_VALUE) {
        return getFacultyBaseForRatio(dataRow);
    }
    const pKey = buildProgramDataKey(year, prog?.dept, prog?.name, prog?.degree);
    const branchKey = buildProgramBranchDataKey(pKey, normalizedBranch);
    const value = Number(programFacultyFteByBranchKey[branchKey]);
    return Number.isFinite(value) && value > 0 ? (Math.round(value * 100) / 100) : 0;
}

function getResearchSupportForProgramYear(prog, year, branch = ALL_BRANCH_FILTER_VALUE) {
    const yearDeptKey = buildYearDeptKey(year, prog?.dept);
    if (!yearDeptKey) return null;
    const support = researchActivitySupportByYearDept[yearDeptKey] || null;
    if (!support) return null;
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch || normalizedBranch === ALL_BRANCH_FILTER_VALUE) return support;
    return support.byBranch?.[normalizedBranch] || null;
}

function applyIslamicBranchStudentMetrics(displayData, programLike, year, branch) {
    if (!displayData || !isIslamicStudiesBachelor(programLike)) return displayData;
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch || normalizedBranch === ALL_BRANCH_FILTER_VALUE) return displayData;

    const metric = getIslamicBranchMetric(programLike, year, normalizedBranch);
    const coverage = getIslamicBranchCoverage(normalizedBranch, year);
    const regularFields = ['students_total', 'students_male', 'students_female'];
    regularFields.forEach(field => {
        displayData[field] = coverage?.allow_regular && metric && Object.prototype.hasOwnProperty.call(metric, field)
            ? metric[field]
            : null;
    });
    displayData.students_new = coverage?.allow_entrant && metric && Object.prototype.hasOwnProperty.call(metric, 'students_new')
        ? metric.students_new
        : null;
    ['graduates_total', 'graduates_branch_matched', 'graduates_ontime', 'students_retained', 'prev_new_count', 'new_4_ago_count']
        .forEach(field => {
            displayData[field] = metric && Object.prototype.hasOwnProperty.call(metric, field)
                ? metric[field]
                : null;
        });
    displayData.students_saudi = null;
    displayData.students_international = null;
    displayData.branch_student_source = metric ? 'islamic_branch_results' : 'branch_data_unavailable';
    displayData.branch_coverage_status = metric?.coverage_status || coverage?.status || 'unavailable';
    return displayData;
}

function buildProgramDisplayData(prog, year, branch = ALL_BRANCH_FILTER_VALUE) {
    const baseData = prog?.years?.[year] || null;
    if (!baseData) return null;
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch || normalizedBranch === ALL_BRANCH_FILTER_VALUE) {
        return { ...baseData };
    }

    const displayData = applyIslamicBranchStudentMetrics(
        { ...baseData },
        prog,
        year,
        normalizedBranch
    );
    const teachingSupport = getTeachingSupportForProgramYear(prog, year, normalizedBranch);
    if (teachingSupport) {
        displayData.sections_total = Number(teachingSupport.totalSections) || 0;
        displayData.sections_male = Number(teachingSupport.maleSections) || 0;
        displayData.sections_female = Number(teachingSupport.femaleSections) || 0;
    } else {
        displayData.sections_total = 0;
        displayData.sections_male = 0;
        displayData.sections_female = 0;
    }
    const facultyBase = getProgramFacultyBaseForBranch(prog, year, normalizedBranch);
    displayData.faculty_ratio_base = facultyBase;
    displayData.faculty_ratio_source = facultyBase > 0 ? 'teaching_fte_branch' : 'branch_none';

    const researchSupport = getResearchSupportForProgramYear(prog, year, normalizedBranch);
    if (researchSupport) {
        displayData.faculty_total = Number(researchSupport.faculty_total) || 0;
        displayData.faculty_published = Number(researchSupport.faculty_published) || 0;
        displayData.research_count = Number(researchSupport.research_count) || 0;
        displayData.citations = Math.round((Number(researchSupport.citations) || 0) * 10) / 10;
        displayData.citations_per_publication = Math.round((Number(researchSupport.citations_per_publication) || 0) * 10) / 10;
        displayData.research_source = `${researchSupport.research_source || 'faculty_activities_live'}_branch`;
    } else {
        displayData.faculty_total = 0;
        displayData.faculty_published = 0;
        displayData.research_count = 0;
        displayData.citations = 0;
        displayData.citations_per_publication = 0;
        displayData.research_source = 'faculty_activities_live_branch';
    }

    return displayData;
}

function buildProgramDisplayDataFromRow(row, branch = ALL_BRANCH_FILTER_VALUE) {
    if (!row) return null;
    const normalizedBranch = normalizeBranchName(branch);
    if (!normalizedBranch || normalizedBranch === ALL_BRANCH_FILTER_VALUE) {
        return { ...row };
    }

    const year = parseInt(row.Semester, 10) || null;
    const dept = normalizeDepartment(row.Dept_aName);
    const major = normalizeSurveyProgramName(row.Major_aName);
    const degree = normalizeDegree(row.Degree_aName);
    const programLike = { dept, name: major, degree };
    const displayData = applyIslamicBranchStudentMetrics(
        { ...row },
        programLike,
        year,
        normalizedBranch
    );
    displayData._selectedBranch = normalizedBranch;
    const teachingSupport = getTeachingSupportForProgramYear(programLike, year, normalizedBranch);
    if (teachingSupport) {
        displayData.sections_total = Number(teachingSupport.totalSections) || 0;
        displayData.sections_male = Number(teachingSupport.maleSections) || 0;
        displayData.sections_female = Number(teachingSupport.femaleSections) || 0;
    } else {
        displayData.sections_total = 0;
        displayData.sections_male = 0;
        displayData.sections_female = 0;
    }
    const pKey = buildProgramDataKey(year, dept, major, degree);
    const branchKey = buildProgramBranchDataKey(pKey, normalizedBranch);
    const facultyBase = Number(programFacultyFteByBranchKey[branchKey]) || 0;
    displayData.faculty_ratio_base = facultyBase > 0 ? Math.round(facultyBase * 100) / 100 : 0;
    displayData.faculty_ratio_source = facultyBase > 0 ? 'teaching_fte_branch' : 'branch_none';

    const researchSupport = getResearchSupportForProgramYear(
        programLike,
        year,
        normalizedBranch
    );
    if (researchSupport) {
        displayData.faculty_total = Number(researchSupport.faculty_total) || 0;
        displayData.faculty_published = Number(researchSupport.faculty_published) || 0;
        displayData.research_count = Number(researchSupport.research_count) || 0;
        displayData.citations = Math.round((Number(researchSupport.citations) || 0) * 10) / 10;
        displayData.citations_per_publication = Math.round((Number(researchSupport.citations_per_publication) || 0) * 10) / 10;
        displayData.research_source = `${researchSupport.research_source || 'faculty_activities_live'}_branch`;
    } else {
        displayData.faculty_total = 0;
        displayData.faculty_published = 0;
        displayData.research_count = 0;
        displayData.citations = 0;
        displayData.citations_per_publication = 0;
        displayData.research_source = 'faculty_activities_live_branch';
    }

    return displayData;
}

function getStudentDetailRowsForProgramYear(prog, year) {
    const normalizedProgram = normalizeSurveyProgramName(prog?.name || '');
    const normalizedDegree = normalizeDegree(prog?.degree || '');
    const normalizedDept = normalizeDepartment(prog?.dept || '');
    return studentDetailData.filter(row =>
        row._year === year &&
        row['التخصص'] === normalizedProgram &&
        row['الدرجة'] === normalizedDegree &&
        row['القسم'] === normalizedDept
    );
}

function createStudySystemBucket() {
    return {
        saudiMale: null,
        saudiFemale: null,
        saudiTotal: null,
        nonSaudiMale: null,
        nonSaudiFemale: null,
        nonSaudiTotal: null,
        total: null
    };
}

function buildStudySystemBreakdown(detailRows, dataRow) {
    const regular = createStudySystemBucket();
    const remote = createStudySystemBucket();

    const applyCounts = (bucket, rows) => {
        const saudiRows = rows.filter(row => classifyStudentNationality(row['الجنسية']) === 'saudi');
        const nonSaudiRows = rows.filter(row => classifyStudentNationality(row['الجنسية']) === 'nonSaudi');
        const countGender = (subset, group) => subset.filter(row => classifyStudentGender(row['الجنس']) === group).length;

        bucket.saudiMale = countGender(saudiRows, 'male');
        bucket.saudiFemale = countGender(saudiRows, 'female');
        bucket.saudiTotal = saudiRows.length;
        bucket.nonSaudiMale = countGender(nonSaudiRows, 'male');
        bucket.nonSaudiFemale = countGender(nonSaudiRows, 'female');
        bucket.nonSaudiTotal = nonSaudiRows.length;
        bucket.total = rows.length;
    };

    if (detailRows.length) {
        const regularRows = detailRows.filter(row => {
            const studyType = normalizeArabicText(row['نوع_الدراسة']);
            return !studyType || studyType.includes('منتظم') || studyType.includes('انتظام');
        });
        const remoteRows = detailRows.filter(row => normalizeArabicText(row['نوع_الدراسة']).includes('عن بعد'));

        applyCounts(regular, regularRows);
        applyCounts(remote, remoteRows);
        return { regular, remote, source: 'detail' };
    }

    if (dataRow) {
        regular.saudiTotal = Number.isFinite(dataRow.students_saudi) ? dataRow.students_saudi : null;
        regular.nonSaudiTotal = Number.isFinite(dataRow.students_international) ? dataRow.students_international : null;
        regular.total = Number.isFinite(dataRow.students_total) ? dataRow.students_total : null;
        remote.total = 0;
        return { regular, remote, source: 'aggregate' };
    }

    return { regular, remote, source: 'none' };
}

function buildProgramYearSnapshot(prog, year, branch = ALL_BRANCH_FILTER_VALUE, dataOverride = null) {
    const dataRow = dataOverride || buildProgramDisplayData(prog, year, branch);
    const teachingSupport = getTeachingSupportForProgramYear(prog, year, branch);
    const normalizedBranch = normalizeBranchName(branch);
    const detailRows = normalizedBranch && normalizedBranch !== ALL_BRANCH_FILTER_VALUE
        ? []
        : getStudentDetailRowsForProgramYear(prog, year);

    const internationalMale = detailRows.length
        ? detailRows.filter(row =>
            classifyStudentNationality(row['الجنسية']) === 'nonSaudi' &&
            classifyStudentGender(row['الجنس']) === 'male'
        ).length
        : null;
    const internationalFemale = detailRows.length
        ? detailRows.filter(row =>
            classifyStudentNationality(row['الجنسية']) === 'nonSaudi' &&
            classifyStudentGender(row['الجنس']) === 'female'
        ).length
        : null;

    const ratioValue = dataRow && getFacultyBaseForRatio(dataRow) > 0 && dataRow.students_total > 0
        ? (dataRow.students_total / getFacultyBaseForRatio(dataRow))
        : null;

    return {
        year,
        dataRow,
        teachingSupport,
        detailRows,
        studySystem: buildStudySystemBreakdown(detailRows, dataRow),
        studentsNewTotal: dataRow ? numberOrNull(dataRow.students_new) : null,
        studentsTotal: dataRow ? numberOrNull(dataRow.students_total) : null,
        studentsMale: dataRow ? numberOrNull(dataRow.students_male) : null,
        studentsFemale: dataRow ? numberOrNull(dataRow.students_female) : null,
        internationalTotal: dataRow ? numberOrNull(dataRow.students_international) : null,
        internationalMale,
        internationalFemale,
        avgStudentsPerSection: teachingSupport?.avgStudentsPerSection ?? null,
        avgMaleStudentsPerSection: teachingSupport?.avgMaleStudentsPerSection ?? null,
        avgFemaleStudentsPerSection: teachingSupport?.avgFemaleStudentsPerSection ?? null,
        studentFacultyRatioValue: ratioValue,
        graduatesTotal: dataRow ? numberOrNull(dataRow.graduates_total) : null,
        employmentEmployedCount: dataRow && Number.isFinite(Number(dataRow.employment_employed_count))
            ? Number(dataRow.employment_employed_count) : null,
        employmentRate: dataRow?.employment_rate ?? null,
        employmentSample: dataRow ? (Number(dataRow.employment_rate_sample) || 0) : 0,
        facultyRanks: teachingSupport?.facultyRanks || {},
        facultyCount: teachingSupport?.facultyCount ?? null,
        facultyAvgLoad: teachingSupport?.facultyAvgLoad ?? null
    };
}

function buildSnapshotPoints(snapshots, selector) {
    return (snapshots || [])
        .map(snapshot => ({ year: snapshot.year, value: selector(snapshot) }))
        .filter(point => Number.isFinite(point.value));
}

function formatSelfStudyCountCell(value) {
    return Number.isFinite(value) ? fmtNum(Math.round(value)) : '—';
}

function formatSelfStudyDecimalCell(value, frac = 1) {
    return Number.isFinite(value) ? fmtNumFlex(value, frac) : '—';
}

function formatSelfStudyPercentCell(value) {
    return Number.isFinite(value) ? `${value.toFixed(1)}%` : '—';
}

function formatSelfStudyRatioCell(value, projected = false) {
    if (!Number.isFinite(value)) return '—';
    if (projected) {
        const rounded = roundToNearestFive(value);
        return Number.isFinite(rounded) ? `1:${fmtNum(rounded)}` : '—';
    }
    return `1:${fmtNumFlex(value, 1)}`;
}

function buildSelfStudyReport(prog, year, branch = ALL_BRANCH_FILTER_VALUE) {
    const normalizedBranch = normalizeBranchName(branch);
    const trendYears = [year - 2, year - 1, year];
    const trendSnapshots = trendYears.map(targetYear => buildProgramYearSnapshot(prog, targetYear, normalizedBranch));
    const trendMap = Object.fromEntries(trendSnapshots.map(snapshot => [snapshot.year, snapshot]));
    const currentSnapshot = trendMap[year];
    const projectedYear = year + 1;

    const projected = {
        studentsNewTotal: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.studentsNewTotal)),
        studentsTotal: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.studentsTotal)),
        studentsMale: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.studentsMale)),
        studentsFemale: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.studentsFemale)),
        internationalTotal: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.internationalTotal)),
        internationalMale: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.internationalMale)),
        internationalFemale: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.internationalFemale)),
        avgStudentsPerSection: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.avgStudentsPerSection)),
        avgMaleStudentsPerSection: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.avgMaleStudentsPerSection)),
        avgFemaleStudentsPerSection: estimateNextRoundedCount(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.avgFemaleStudentsPerSection)),
        studentFacultyRatioValue: estimateNextValue(buildSnapshotPoints(trendSnapshots, snapshot => snapshot.studentFacultyRatioValue))
    };

    const graduatesYears = [year - 3, year - 2, year - 1];
    const graduateSnapshots = graduatesYears.map(targetYear => buildProgramYearSnapshot(prog, targetYear, normalizedBranch));
    const facultyRanks = currentSnapshot?.facultyRanks || {};
    const facultyDataAvailable = Number.isFinite(currentSnapshot?.facultyCount) || Object.keys(facultyRanks).length > 0;
    const facultyGroupDefinitions = [
        { group: 'أعضاء هيئة التدريس', ranks: ['أستاذ', 'أستاذ مشارك', 'أستاذ مساعد', 'متعاون'] },
        { group: 'هيئة التدريس من غير حملة الدكتوراه', ranks: ['محاضر', 'معيد', 'مدرس'] }
    ];
    const allFacultyRanks = [...new Set([
        ...facultyGroupDefinitions.flatMap(item => item.ranks),
        ...Object.keys(facultyRanks)
    ])];
    const aggregateFacultyRanks = ranks => {
        const aggregate = createFacultyRankAggregate();
        (ranks || []).forEach(rank => mergeFacultyRankAggregate(aggregate, facultyRanks[rank]));
        return aggregate;
    };
    const formatFacultyLoadCell = (sum, count) => {
        if (!facultyDataAvailable || !Number.isFinite(sum) || !Number.isFinite(count) || count <= 0) return '—';
        return formatSelfStudyDecimalCell(sum / count);
    };
    const buildFacultyRow = (group, label, stats) => {
        const info = stats || null;
        return {
            group,
            label,
            values: [
                formatSelfStudyCountCell(facultyDataAvailable ? (info?.saudiMaleCount ?? 0) : null),
                formatSelfStudyCountCell(facultyDataAvailable ? (info?.saudiFemaleCount ?? 0) : null),
                formatSelfStudyCountCell(facultyDataAvailable ? (info?.saudiCount ?? 0) : null),
                formatSelfStudyCountCell(facultyDataAvailable ? (info?.nonSaudiMaleCount ?? 0) : null),
                formatSelfStudyCountCell(facultyDataAvailable ? (info?.nonSaudiFemaleCount ?? 0) : null),
                formatSelfStudyCountCell(facultyDataAvailable ? (info?.nonSaudiCount ?? 0) : null),
                formatFacultyLoadCell(info?.maleLoadSum, info?.maleLoadCount),
                formatFacultyLoadCell(info?.femaleLoadSum, info?.femaleLoadCount),
                formatFacultyLoadCell(info?.loadSum, info?.count)
            ]
        };
    };
    const facultyRows = [];
    facultyGroupDefinitions.forEach(groupInfo => {
        groupInfo.ranks.forEach((rank, index) => {
            facultyRows.push(buildFacultyRow(index === 0 ? groupInfo.group : '', rank, facultyRanks[rank] || null));
        });
        facultyRows.push(buildFacultyRow('', 'الإجمالي', aggregateFacultyRanks(groupInfo.ranks)));
    });
    const facultyOverall = aggregateFacultyRanks(allFacultyRanks);

    const notes = [
        'عمود المتوقع بعد عام تقديري، مبني على الاتجاه الفعلي ومقرب لأقرب 5.',
        'صف عدد الطلاب المخطط التحاقهم بالبرنامج يعتمد على عدد المستجدين الفعلي عند غياب ملف الهدف السنوي.',
        'بيانات نظام الدراسة والتفصيل بالجنسية والجنس، بما في ذلك توزيع غير السعوديين إلى ذكور وإناث، تعتمد على ملف الطلاب التفصيلي والمتاح حاليًا لعامي 1445 و1446.',
        'إحصاءات الشعب وعبء التدريس مبنية على سجلات النشاط التدريسي المرتبطة بمقررات البرنامج.',
        'جنسية أعضاء هيئة التدريس وجنسهم يعتمدان على ملف الأعضاء المحدث داخل مشروع الأنشطة.',
        'متوسط عبء التدريس للذكور والإناث في جدول هيئة التدريس يعتمد على متوسط العبء الفعلي لأعضاء هيئة التدريس المصنفين بهذا الجنس داخل البرنامج.'
    ];
    if (normalizedBranch && normalizedBranch !== ALL_BRANCH_FILTER_VALUE) {
        if (isIslamicStudiesBachelor(prog)) {
            const coverage = getIslamicBranchCoverage(normalizedBranch, year);
            const reasons = Array.isArray(coverage?.reasons) ? coverage.reasons.join(' ') : '';
            notes.unshift(`الفرع المختار: ${normalizedBranch}. أعداد الطلاب المعروضة رصدٌ محافظ للسجلات التي أمكن إسنادها للبرنامج والفرع بدليل كافٍ، وليست إجماليات رسمية مكتملة. يبقى مستوى الثقة محفوظًا داخليًا لأغراض التدقيق.${reasons ? ` ${reasons}` : ''}`);
        } else {
            notes.unshift(`المقر المعتمد لهذا البرنامج: الحوية. تتغير بيانات هيئة التدريس والنشر العلمي ونسبة الطلاب إلى أعضاء هيئة التدريس بحسب المقر.`);
        }
    }
    if (facultyDataAvailable && facultyOverall.unknownNationalityCount > 0) {
        notes.push(`يوجد ${fmtNum(facultyOverall.unknownNationalityCount)} عضو/أعضاء هيئة تدريس بلا جنسية مصنفة في الملف الحالي، ويظهرون ضمن الإجمالي فقط.`);
    }
    if (facultyDataAvailable && facultyOverall.unknownGenderCount > 0) {
        notes.push(`يوجد ${fmtNum(facultyOverall.unknownGenderCount)} عضو/أعضاء هيئة تدريس بلا جنس مصنف في الملف الحالي، ويظهرون ضمن الإجمالي فقط.`);
    }

    return {
        title: 'أرقام الدراسة الذاتية',
        notes,
        enrollmentTable: {
            title: '1.9.1 تطور أعداد الطلاب الملتحقين بالبرنامج',
            headers: [
                'البند',
                'الفئة',
                `قبل عامين (${fmtYear(year - 2)})`,
                `العام الماضي (${fmtYear(year - 1)})`,
                `العام الحالي (${fmtYear(year)})`,
                `المتوقع بعد عام (${fmtYear(projectedYear)})`
            ],
            rows: [
                {
                    group: 'عدد الطلاب المخطط التحاقهم بالبرنامج',
                    label: 'ذكور',
                    values: ['—', '—', '—', '—']
                },
                {
                    group: 'عدد الطلاب المخطط التحاقهم بالبرنامج',
                    label: 'إناث',
                    values: ['—', '—', '—', '—']
                },
                {
                    group: 'عدد الطلاب المخطط التحاقهم بالبرنامج',
                    label: 'الإجمالي',
                    values: [
                        formatSelfStudyCountCell(trendMap[year - 2]?.studentsNewTotal),
                        formatSelfStudyCountCell(trendMap[year - 1]?.studentsNewTotal),
                        formatSelfStudyCountCell(trendMap[year]?.studentsNewTotal),
                        formatSelfStudyCountCell(projected.studentsNewTotal)
                    ]
                },
                {
                    group: 'العدد الكلي للطلاب الملتحقين بالبرنامج',
                    label: 'ذكور',
                    values: [
                        formatSelfStudyCountCell(trendMap[year - 2]?.studentsMale),
                        formatSelfStudyCountCell(trendMap[year - 1]?.studentsMale),
                        formatSelfStudyCountCell(trendMap[year]?.studentsMale),
                        formatSelfStudyCountCell(projected.studentsMale)
                    ]
                },
                {
                    group: 'العدد الكلي للطلاب الملتحقين بالبرنامج',
                    label: 'إناث',
                    values: [
                        formatSelfStudyCountCell(trendMap[year - 2]?.studentsFemale),
                        formatSelfStudyCountCell(trendMap[year - 1]?.studentsFemale),
                        formatSelfStudyCountCell(trendMap[year]?.studentsFemale),
                        formatSelfStudyCountCell(projected.studentsFemale)
                    ]
                },
                {
                    group: 'العدد الكلي للطلاب الملتحقين بالبرنامج',
                    label: 'الإجمالي',
                    values: [
                        formatSelfStudyCountCell(trendMap[year - 2]?.studentsTotal),
                        formatSelfStudyCountCell(trendMap[year - 1]?.studentsTotal),
                        formatSelfStudyCountCell(trendMap[year]?.studentsTotal),
                        formatSelfStudyCountCell(projected.studentsTotal)
                    ]
                },
                {
                    group: 'عدد الطلاب الدوليين الملتحقين بالبرنامج',
                    label: 'ذكور',
                    values: [
                        formatSelfStudyCountCell(trendMap[year - 2]?.internationalMale),
                        formatSelfStudyCountCell(trendMap[year - 1]?.internationalMale),
                        formatSelfStudyCountCell(trendMap[year]?.internationalMale),
                        formatSelfStudyCountCell(projected.internationalMale)
                    ]
                },
                {
                    group: 'عدد الطلاب الدوليين الملتحقين بالبرنامج',
                    label: 'إناث',
                    values: [
                        formatSelfStudyCountCell(trendMap[year - 2]?.internationalFemale),
                        formatSelfStudyCountCell(trendMap[year - 1]?.internationalFemale),
                        formatSelfStudyCountCell(trendMap[year]?.internationalFemale),
                        formatSelfStudyCountCell(projected.internationalFemale)
                    ]
                },
                {
                    group: 'عدد الطلاب الدوليين الملتحقين بالبرنامج',
                    label: 'الإجمالي',
                    values: [
                        formatSelfStudyCountCell(trendMap[year - 2]?.internationalTotal),
                        formatSelfStudyCountCell(trendMap[year - 1]?.internationalTotal),
                        formatSelfStudyCountCell(trendMap[year]?.internationalTotal),
                        formatSelfStudyCountCell(projected.internationalTotal)
                    ]
                },
                {
                    group: 'متوسط عدد الطلاب في الشعب الدراسية',
                    label: 'ذكور',
                    values: [
                        formatSelfStudyDecimalCell(trendMap[year - 2]?.avgMaleStudentsPerSection),
                        formatSelfStudyDecimalCell(trendMap[year - 1]?.avgMaleStudentsPerSection),
                        formatSelfStudyDecimalCell(trendMap[year]?.avgMaleStudentsPerSection),
                        formatSelfStudyCountCell(projected.avgMaleStudentsPerSection)
                    ]
                },
                {
                    group: 'متوسط عدد الطلاب في الشعب الدراسية',
                    label: 'إناث',
                    values: [
                        formatSelfStudyDecimalCell(trendMap[year - 2]?.avgFemaleStudentsPerSection),
                        formatSelfStudyDecimalCell(trendMap[year - 1]?.avgFemaleStudentsPerSection),
                        formatSelfStudyDecimalCell(trendMap[year]?.avgFemaleStudentsPerSection),
                        formatSelfStudyCountCell(projected.avgFemaleStudentsPerSection)
                    ]
                },
                {
                    group: 'متوسط عدد الطلاب في الشعب الدراسية',
                    label: 'الإجمالي',
                    values: [
                        formatSelfStudyDecimalCell(trendMap[year - 2]?.avgStudentsPerSection),
                        formatSelfStudyDecimalCell(trendMap[year - 1]?.avgStudentsPerSection),
                        formatSelfStudyDecimalCell(trendMap[year]?.avgStudentsPerSection),
                        formatSelfStudyCountCell(projected.avgStudentsPerSection)
                    ]
                },
                {
                    group: 'نسبة عدد الطلاب إلى هيئة التدريس',
                    label: 'ذكور',
                    values: ['—', '—', '—', '—']
                },
                {
                    group: 'نسبة عدد الطلاب إلى هيئة التدريس',
                    label: 'إناث',
                    values: ['—', '—', '—', '—']
                },
                {
                    group: 'نسبة عدد الطلاب إلى هيئة التدريس',
                    label: 'الإجمالي',
                    values: [
                        formatSelfStudyRatioCell(trendMap[year - 2]?.studentFacultyRatioValue),
                        formatSelfStudyRatioCell(trendMap[year - 1]?.studentFacultyRatioValue),
                        formatSelfStudyRatioCell(trendMap[year]?.studentFacultyRatioValue),
                        formatSelfStudyRatioCell(projected.studentFacultyRatioValue, true)
                    ]
                }
            ]
        },
        studySystemTable: {
            title: '2.9.1 تصنيف الطلاب حسب نظام الدراسة والجنسية',
            headers: [
                'التصنيف',
                'سعودي (ذكور)',
                'سعودي (إناث)',
                'سعودي (الإجمالي)',
                'غير سعودي (ذكور)',
                'غير سعودي (إناث)',
                'غير سعودي (الإجمالي)',
                'الإجمالي'
            ],
            rows: [
                {
                    label: 'انتظام',
                    values: [
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.regular?.saudiMale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.regular?.saudiFemale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.regular?.saudiTotal),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.regular?.nonSaudiMale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.regular?.nonSaudiFemale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.regular?.nonSaudiTotal),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.regular?.total)
                    ]
                },
                {
                    label: 'تعليم عن بعد',
                    values: [
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.remote?.saudiMale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.remote?.saudiFemale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.remote?.saudiTotal),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.remote?.nonSaudiMale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.remote?.nonSaudiFemale),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.remote?.nonSaudiTotal),
                        formatSelfStudyCountCell(currentSnapshot?.studySystem?.remote?.total)
                    ]
                }
            ],
            source: currentSnapshot?.studySystem?.source || 'none'
        },
        graduatesTable: {
            title: '3.9.1 تطور أعداد خريجي البرنامج',
            headers: [
                'البند',
                `قبل ثلاثة أعوام (${fmtYear(year - 3)})`,
                `قبل عامين (${fmtYear(year - 2)})`,
                `العام الماضي (${fmtYear(year - 1)})`
            ],
            rows: [
                {
                    group: 'أعداد الخريجين',
                    label: prog.degree,
                    values: graduateSnapshots.map(snapshot => formatSelfStudyCountCell(snapshot.graduatesTotal))
                },
                {
                    group: 'أعداد الخريجين',
                    label: 'المجموع',
                    values: graduateSnapshots.map(snapshot => formatSelfStudyCountCell(snapshot.graduatesTotal))
                },
                {
                    group: 'توظيف الخريجين أو التحاقهم بالدراسات العليا',
                    label: 'عدد الموظفين أو الملتحقين',
                    values: graduateSnapshots.map(snapshot => formatSelfStudyCountCell(snapshot.employmentEmployedCount))
                },
                {
                    group: 'توظيف الخريجين أو التحاقهم بالدراسات العليا',
                    label: 'النسبة إلى إجمالي الخريجين',
                    values: graduateSnapshots.map(snapshot => formatSelfStudyPercentCell(snapshot.employmentRate))
                }
            ]
        },
        facultyTable: {
            title: '4.9.1 أعداد هيئة التدريس',
            caption: facultyDataAvailable
                ? 'يعرض الجدول أعداد أعضاء هيئة التدريس حسب الرتبة والجنسية والجنس، ويعرض متوسط العبء التدريسي للذكور والإناث والإجمالي.'
                : 'لا تتوفر لهذا البرنامج في هذه السنة سجلات ربط كافية لإظهار تفصيل أعضاء هيئة التدريس والعبء التدريسي.',
            headers: ['الفئة', 'الرتبة', 'سعودي (ذكور)', 'سعودي (إناث)', 'سعودي (الإجمالي)', 'غير سعودي (ذكور)', 'غير سعودي (إناث)', 'غير سعودي (الإجمالي)', 'متوسط عبء التدريس (ذكور)', 'متوسط عبء التدريس (إناث)', 'متوسط عبء التدريس (الإجمالي)'],
            rows: [
                ...facultyRows,
                buildFacultyRow('الإجمالي العام', 'الإجمالي الكلي', facultyOverall)
            ]
        }
    };
}

function renderSelfStudyTable(title, headers, rows, caption = '') {
    return `
        <div class="self-study-section">
            <h4>${title}</h4>
            ${caption ? `<p class="self-study-caption">${caption}</p>` : ''}
            <div class="table-wrap">
                <table class="data-table self-study-table">
                    <thead>
                        <tr>${headers.map(header => `<th>${header}</th>`).join('')}</tr>
                    </thead>
                    <tbody>
                        ${rows.map(row => `
                            <tr>
                                ${row.group != null ? `<td class="self-study-group">${row.group}</td>` : ''}
                                <td class="self-study-label">${row.label}</td>
                                ${row.values.map(value => `<td>${value}</td>`).join('')}
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

function renderSelfStudyReport(report) {
    if (!report) return '<p class="self-study-empty">لا توجد بيانات متاحة لهذا القسم.</p>';
    const notesHtml = report.notes.map(note => `<li>${note}</li>`).join('');
    const studySystemSource = report.studySystemTable.source === 'detail'
        ? 'يعتمد هذا الجدول على ملف الطلاب التفصيلي.'
        : (report.studySystemTable.source === 'aggregate'
            ? 'يظهر هذا الجدول الإجماليات المتاحة فقط لعدم توفر الملف التفصيلي لهذه السنة.'
            : 'لا توجد بيانات طلاب تفصيلية لهذه السنة.');

    return `
        <div class="self-study-notes">
            <ul>${notesHtml}</ul>
        </div>
        ${renderSelfStudyTable(report.enrollmentTable.title, report.enrollmentTable.headers, report.enrollmentTable.rows, report.enrollmentTable.caption)}
        <div class="self-study-section">
            <h4>${report.studySystemTable.title}</h4>
            <p class="self-study-caption">${studySystemSource}</p>
            <div class="table-wrap">
                <table class="data-table self-study-table">
                    <thead>
                        <tr>${report.studySystemTable.headers.map(header => `<th>${header}</th>`).join('')}</tr>
                    </thead>
                    <tbody>
                        ${report.studySystemTable.rows.map(row => `
                            <tr>
                                <td class="self-study-label">${row.label}</td>
                                ${row.values.map(value => `<td>${value}</td>`).join('')}
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
        ${renderSelfStudyTable(report.graduatesTable.title, report.graduatesTable.headers, report.graduatesTable.rows, report.graduatesTable.caption)}
        ${renderSelfStudyTable(report.facultyTable.title, report.facultyTable.headers, report.facultyTable.rows, report.facultyTable.caption)}
    `;
}

function getSelfStudyExportRows(report) {
    if (!report) return [];
    const rows = [
        [report.title],
        ['ملاحظات'],
        ...report.notes.map(note => [note]),
        []
    ];

    const pushTable = (table) => {
        rows.push([table.title]);
        if (table.caption) rows.push([table.caption]);
        rows.push(table.headers);
        table.rows.forEach(row => rows.push([row.group || '', row.label, ...row.values]));
        rows.push([]);
    };

    pushTable(report.enrollmentTable);
    rows.push([report.studySystemTable.title]);
    rows.push(report.studySystemTable.headers);
    report.studySystemTable.rows.forEach(row => rows.push([row.label, ...row.values]));
    rows.push([]);
    pushTable(report.graduatesTable);
    pushTable(report.facultyTable);
    return rows;
}

// ========================================
// لوحة المعلومات
// ========================================
function initDashboard() {
    const sel = document.getElementById('dash-year');
    const years = getAvailableYears();
    sel.innerHTML = years.map(y =>
        `<option value="${y}" ${y === years[years.length-1] ? 'selected' : ''}>${fmtYear(y)}</option>`
    ).join('');
    sel.addEventListener('change', () => renderDashboard(parseInt(sel.value)));
    renderDashboard(years[years.length - 1]);
}

function renderDashboard(year) {
    const items = getYearData(year);
    renderSummary(items, year);
    renderStudentsChart(items);
    renderRatesChart(items, year);
    renderGenderChart(items);
    renderFlowChart(items);
    renderProgramsTable(items, year);
}

function renderSummary(items, year) {
    const totalStudents = items.reduce((s,x) => s + x.data.students_total, 0);
    const totalGrads = items.reduce((s,x) => s + x.data.graduates_total, 0);
    const totalNew = items.reduce((s,x) => s + x.data.students_new, 0);
    const totalRetained = items.reduce((s,x) => s + x.data.students_retained, 0);
    const totalPrevNew = items.reduce((s,x) => s + x.data.prev_new_count, 0);
    const avgRetention = totalPrevNew > 0 ? pct(totalRetained, totalPrevNew) : null;

    document.getElementById('summary-row').innerHTML = `
        <div class="summary-card">
            <div class="sc-icon">📚</div>
            <div class="sc-value">${items.length}</div>
            <div class="sc-label">برنامج أكاديمي</div>
        </div>
        <div class="summary-card">
            <div class="sc-icon">👥</div>
            <div class="sc-value">${fmtNum(totalStudents)}</div>
            <div class="sc-label">طالب منتظم</div>
        </div>
        <div class="summary-card">
            <div class="sc-icon">🎓</div>
            <div class="sc-value">${fmtNum(totalGrads)}</div>
            <div class="sc-label">خريج</div>
        </div>
        <div class="summary-card">
            <div class="sc-icon">🆕</div>
            <div class="sc-value">${fmtNum(totalNew)}</div>
            <div class="sc-label">طالب مستجد</div>
        </div>
        <div class="summary-card">
            <div class="sc-icon">📊</div>
            <div class="sc-value">${avgRetention != null ? avgRetention.toFixed(1) + '%' : '—'}</div>
            <div class="sc-label">متوسط الاستبقاء</div>
        </div>
    `;
}

function renderStudentsChart(items) {
    destroyChart('students');
    const sorted = [...items].sort((a,b) => b.data.students_total - a.data.students_total);
    const labels = sorted.map(x => x.prog.name + ' (' + x.prog.degree + ')');
    const data = sorted.map(x => x.data.students_total);
    const ctx = document.getElementById('chart-students').getContext('2d');
    charts.students = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                data,
                backgroundColor: sorted.map((_,i) => CHART_COLORS[i % CHART_COLORS.length]),
                borderRadius: 4,
                barThickness: 18,
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                x: { grid: { color: '#f0f0f0' }, ticks: { font: { family: 'Tajawal' } } },
                y: { ticks: { font: { family: 'Tajawal', size: 11 } } }
            }
        }
    });
}

function renderRatesChart(items, year) {
    destroyChart('rates');
    const wrap = document.getElementById('chart-rates-wrap');
    wrap.innerHTML = '<canvas id="chart-rates"></canvas>';
    const withRates = items.filter(x => {
        const gr = pct(x.data.graduates_ontime, x.data.new_4_ago_count);
        const rr = pct(x.data.students_retained, x.data.prev_new_count);
        return gr != null || rr != null;
    });
    if (withRates.length === 0) {
        wrap.innerHTML = '<p style="text-align:center;color:#999;padding:40px;">لا توجد بيانات معدلات لهذه السنة</p>';
        return;
    }
    const labels = withRates.map(x => x.prog.name + ' (' + x.prog.degree + ')');
    const gradRates = withRates.map(x => pct(x.data.graduates_ontime, x.data.new_4_ago_count) || 0);
    const retRates = withRates.map(x => pct(x.data.students_retained, x.data.prev_new_count) || 0);
    const ctx = document.getElementById('chart-rates').getContext('2d');
    charts.rates = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'معدل التخرج بالوقت', data: gradRates, backgroundColor: '#0d8e8e', borderRadius: 4, barThickness: 14 },
                { label: 'معدل الاستبقاء', data: retRates, backgroundColor: '#c9a227', borderRadius: 4, barThickness: 14 },
            ]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top', labels: { font: { family: 'Tajawal' } } },
                tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + ctx.parsed.x.toFixed(1) + '%' } }
            },
            scales: {
                x: { max: 100, grid: { color: '#f0f0f0' }, ticks: { callback: v => v + '%', font: { family: 'Tajawal' } } },
                y: { ticks: { font: { family: 'Tajawal', size: 11 } } }
            }
        }
    });
}

function renderGenderChart(items) {
    destroyChart('gender');
    const totalMale = items.reduce((s,x) => s + x.data.students_male, 0);
    const totalFemale = items.reduce((s,x) => s + x.data.students_female, 0);
    const ctx = document.getElementById('chart-gender').getContext('2d');
    charts.gender = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['ذكور', 'إناث'],
            datasets: [{
                data: [totalMale, totalFemale],
                backgroundColor: ['#3b82f6', '#ec4899'],
                borderWidth: 0,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { font: { family: 'Tajawal', size: 13 } } },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const total = ctx.dataset.data.reduce((a,b) => a+b, 0);
                            const p = ((ctx.parsed / total) * 100).toFixed(1);
                            return ctx.label + ': ' + ctx.parsed.toLocaleString('ar-SA') + ' (' + p + '%)';
                        }
                    }
                }
            }
        }
    });
}

function renderFlowChart(items) {
    destroyChart('flow');
    // Top programs by students
    const sorted = [...items].sort((a,b) => b.data.students_total - a.data.students_total).slice(0, 10);
    const labels = sorted.map(x => x.prog.name.length > 15 ? x.prog.name.slice(0,15) + '..' : x.prog.name);
    const ctx = document.getElementById('chart-flow').getContext('2d');
    charts.flow = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'المستجدون', data: sorted.map(x => x.data.students_new), backgroundColor: '#3b82f6', borderRadius: 4 },
                { label: 'الخريجين', data: sorted.map(x => x.data.graduates_total), backgroundColor: '#10b981', borderRadius: 4 },
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'top', labels: { font: { family: 'Tajawal' } } } },
            scales: {
                x: { ticks: { font: { family: 'Tajawal', size: 10 } } },
                y: { grid: { color: '#f0f0f0' }, ticks: { font: { family: 'Tajawal' } } }
            }
        }
    });
}

function renderProgramsTable(items, year) {
    const tbody = document.getElementById('tbl-body');
    const tfoot = document.getElementById('tbl-foot');
    let totals = { students:0, male:0, female:0, newS:0, grads:0 };

    tbody.innerHTML = items.map((x, idx) => {
        const d = x.data;
        totals.students += d.students_total;
        totals.male += d.students_male;
        totals.female += d.students_female;
        totals.newS += d.students_new;
        totals.grads += d.graduates_total;
        const gr = pct(d.graduates_ontime, d.new_4_ago_count);
        const rr = pct(d.students_retained, d.prev_new_count);
        const retainedText = d.prev_new_count > 0 ? `${fmtNum(d.students_retained)} من ${fmtNum(d.prev_new_count)}` : '';
        const ontimeText = d.new_4_ago_count > 0 ? `${fmtNum(d.graduates_ontime)} من ${fmtNum(d.new_4_ago_count)}` : '';
        // find program index
        const pIdx = programs.indexOf(x.prog);
        return `<tr class="clickable-row" data-pidx="${pIdx}" data-year="${year}">
            <td>${x.prog.name}</td>
            <td>${x.prog.degree}</td>
            <td>${x.prog.dept}</td>
            <td>${fmtNum(d.students_total)}</td>
            <td>${fmtNum(d.students_male)}</td>
            <td>${fmtNum(d.students_female)}</td>
            <td>${fmtNum(d.students_new)}</td>
            <td>${fmtNum(d.graduates_total)}</td>
            <td title="${retainedText}">${rateBadge(rr)}</td>
            <td title="${ontimeText}">${rateBadge(gr)}</td>
        </tr>`;
    }).join('');

    tfoot.innerHTML = `<tr>
        <td colspan="3">الإجمالي</td>
        <td>${fmtNum(totals.students)}</td>
        <td>${fmtNum(totals.male)}</td>
        <td>${fmtNum(totals.female)}</td>
        <td>${fmtNum(totals.newS)}</td>
        <td>${fmtNum(totals.grads)}</td>
        <td colspan="2"></td>
    </tr>`;

    // Click to navigate to program detail
    tbody.querySelectorAll('.clickable-row').forEach(tr => {
        tr.addEventListener('click', () => {
            const pidx = tr.dataset.pidx;
            const yr = tr.dataset.year;
            switchView('program');
            document.getElementById('prog-select').value = pidx;
            populateProgramBranchFilter(programs[parseInt(pidx, 10)] || null);
            fillProgYears(parseInt(pidx));
            document.getElementById('prog-year').value = yr;
            document.getElementById('prog-show').disabled = false;
            showProgramDetail();
        });
    });
}

// ========================================
// تفاصيل البرنامج
// ========================================
function populateProgramBranchFilter(prog = null) {
    const branchSelect = document.getElementById('prog-branch');
    if (!branchSelect) return;
    const branches = prog ? getBranchesForProgram(prog) : [];
    branchSelect.innerHTML = buildBranchOptionsHtml(branches);
    if (!prog) {
        branchSelect.value = ALL_BRANCH_FILTER_VALUE;
        branchSelect.disabled = true;
        return;
    }
    if (isIslamicStudiesBachelor(prog)) {
        branchSelect.value = ALL_BRANCH_FILTER_VALUE;
        branchSelect.disabled = false;
    } else {
        branchSelect.value = 'الحوية';
        branchSelect.disabled = true;
    }
}

function getProgramBranchSelection() {
    const rawValue = document.getElementById('prog-branch')?.value || ALL_BRANCH_FILTER_VALUE;
    return normalizeBranchName(rawValue) || ALL_BRANCH_FILTER_VALUE;
}

function initProgramView() {
    const sel = document.getElementById('prog-select');
    const branchSel = document.getElementById('prog-branch');
    sel.innerHTML = '<option value="">-- اختر البرنامج --</option>' +
        programs.map((p,i) => `<option value="${i}">${p.name} (${p.degree})</option>`).join('');
    populateProgramBranchFilter();

    sel.addEventListener('change', () => {
        const v = sel.value;
        document.getElementById('prog-results').classList.add('hidden');
        if (v !== '') {
            const selectedProgram = programs[parseInt(v, 10)];
            populateProgramBranchFilter(selectedProgram);
            fillProgYears(parseInt(v));
        } else {
            populateProgramBranchFilter();
            document.getElementById('prog-year').disabled = true;
            document.getElementById('prog-show').disabled = true;
        }
    });

    document.getElementById('prog-year').addEventListener('change', e => {
        document.getElementById('prog-show').disabled = !e.target.value;
    });

    branchSel?.addEventListener('change', () => {
        const idx = parseInt(document.getElementById('prog-select').value, 10);
        const year = parseInt(document.getElementById('prog-year').value, 10);
        if (!Number.isNaN(idx) && !Number.isNaN(year)) {
            showProgramDetail();
        }
    });

    document.getElementById('prog-show').addEventListener('click', showProgramDetail);
}

function fillProgYears(idx) {
    const yearSel = document.getElementById('prog-year');
    const p = programs[idx];
    const years = Object.keys(p.years).map(Number).filter(y => DISPLAY_YEARS.includes(y)).sort();
    yearSel.innerHTML = '<option value="">-- اختر --</option>' +
        years.map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');
    yearSel.disabled = false;
    document.getElementById('prog-show').disabled = true;
}

function showProgramDetail() {
    const idx = parseInt(document.getElementById('prog-select').value);
    const year = parseInt(document.getElementById('prog-year').value);
    const branch = getProgramBranchSelection();
    if (isNaN(idx) || isNaN(year)) return;

    const prog = programs[idx];
    const d = buildProgramDisplayData(prog, year, branch);
    if (!d) return;

    const selfStudyReport = buildSelfStudyReport(prog, year, branch);
    currentProg = {
        prog,
        year,
        branch,
        branchLabel: branch === ALL_BRANCH_FILTER_VALUE ? ALL_BRANCH_FILTER_LABEL : branch,
        data: d,
        selfStudyReport
    };

    // Header
    document.getElementById('prog-badge').textContent = prog.degree;
    document.getElementById('prog-name').textContent = prog.name;
    document.getElementById('prog-year-label').textContent =
        'السنة: ' + fmtYear(year) + ' | القسم: ' + prog.dept + ' | الفرع: ' + currentProg.branchLabel;
    const scopeNote = document.getElementById('prog-scope-note');
    if (scopeNote) {
        if (branch === ALL_BRANCH_FILTER_VALUE) {
            scopeNote.textContent = 'يعرض التقرير الحالي أرقام البرنامج على مستوى جميع الفروع.';
        } else if (isIslamicStudiesBachelor(prog)) {
            const coverage = getIslamicBranchCoverage(branch, year);
            const coverageNote = Array.isArray(coverage?.reasons) && coverage.reasons.length
                ? ` ${coverage.reasons.join(' ')}`
                : '';
            scopeNote.textContent = `الفرع المختار: ${branch}. الأعداد المعروضة رصد محافظ للسجلات المقبولة التي أمكن إسنادها للبرنامج والفرع، وليست إجماليات رسمية مكتملة.${coverageNote}`;
        } else {
            scopeNote.textContent = 'المقر المعتمد لهذا البرنامج: الحوية.';
        }
        scopeNote.classList.remove('hidden');
    }

    // Stats
    const stats = [];
    const hasBranchStudentMetrics = d.branch_student_source === 'islamic_branch_results';
    const regularLabel = hasBranchStudentMetrics ? 'المنتظمون المرصودون' : 'إجمالي المنتظمين';
    const entrantsLabel = hasBranchStudentMetrics ? 'المستجدون المرصودون' : 'المستجدون';
    if (d.students_total > 0 || (hasBranchStudentMetrics && d.students_total === 0)) stats.push({icon:'👥', label:regularLabel, value: fmtNum(d.students_total), color:'var(--primary)'});
    if (d.students_male > 0) stats.push({icon:'👨', label:'الذكور', value: fmtNum(d.students_male), color:'#3b82f6'});
    if (d.students_female > 0) stats.push({icon:'👩', label:'الإناث', value: fmtNum(d.students_female), color:'#ec4899'});
    if (d.students_saudi > 0) stats.push({icon:'🇸🇦', label:'السعوديون', value: fmtNum(d.students_saudi), color:'#059669'});
    if (d.students_international > 0) stats.push({icon:'🌍', label:'الدوليون', value: fmtNum(d.students_international), color:'#f97316'});
    if (d.students_new > 0 || (hasBranchStudentMetrics && d.students_new === 0)) stats.push({icon:'🆕', label:entrantsLabel, value: fmtNum(d.students_new), color:'#06b6d4'});
    if (d.students_retained > 0 || d.prev_new_count > 0) stats.push({icon:'🔄', label:'استبقاء الدفعة السابقة', value: d.prev_new_count > 0 ? fmtNum(d.students_retained) + ' من ' + fmtNum(d.prev_new_count) : fmtNum(d.students_retained), color:'#8b5cf6'});
    if (d.graduates_total > 0 || (hasBranchStudentMetrics && d.graduates_total === 0)) stats.push({icon:'🎓', label:'إجمالي الخريجين', value: fmtNum(d.graduates_total), color:'#10b981'});
    if (hasBranchStudentMetrics && d.graduates_branch_matched > 0) stats.push({icon:'🎓', label:'خريجون أمكن إسنادهم للفرع', value: fmtNum(d.graduates_branch_matched), color:'#10b981'});
    if (d.graduates_ontime > 0 || d.new_4_ago_count > 0) stats.push({icon:'⏱️', label:'خريجو الدفعة بالوقت', value: d.new_4_ago_count > 0 ? fmtNum(d.graduates_ontime) + ' من ' + fmtNum(d.new_4_ago_count) : fmtNum(d.graduates_ontime), color:'#0d8e8e'});
    const branchTeachingSupport = getTeachingSupportForProgramYear(prog, year, branch);
    const sectionsCount = branch !== ALL_BRANCH_FILTER_VALUE && Number.isFinite(branchTeachingSupport?.totalSections)
        ? branchTeachingSupport.totalSections
        : d.sections_total;
    if (sectionsCount > 0) stats.push({icon:'🏛️', label:'الشعب', value: fmtNum(sectionsCount), color:'#6b7280'});
    const facultyBase = getFacultyBaseForRatio(d);
    if (facultyBase > 0) {
        const facultyLabel = String(d.faculty_ratio_source || '').startsWith('teaching_fte')
            ? 'هيئة التدريس (مكافئ FTE)'
            : 'هيئة التدريس';
        stats.push({icon:'👨‍🏫', label: facultyLabel, value: fmtNumFlex(facultyBase), color:'#7c3aed'});
    }
    if (d.research_count > 0) stats.push({icon:'📚', label:'الأبحاث', value: fmtNum(d.research_count), color:'#eab308'});
    if (d.citations > 0) stats.push({icon:'📎', label:'الاقتباسات', value: fmtNum(d.citations), color:'#14b8a6'});

    document.getElementById('stats-grid').innerHTML = stats.map(s =>
        `<div class="stat-card" style="border-right-color:${s.color}">
            <div class="stat-icon">${s.icon}</div>
            <div class="stat-info">
                <div class="stat-label">${s.label}</div>
                <div class="stat-value">${s.value}</div>
            </div>
        </div>`
    ).join('');

    // KPIs
    const kpi = calcKPIs(d, prog.degree);
    const programIndicators = getIndicatorsForDegree(prog.degree);
    document.getElementById('kpi-grid').innerHTML = programIndicators
        .map(ind => {
            const v = kpi[ind.key];
            const f = fmtKPI(v, ind.unit);
            // إضافة تفاصيل "X من Y" لمعدل التخرج والاستبقاء
            let detailHtml = '';
            if (ind.key === 'graduation_rate' && kpi.graduation_detail) {
                detailHtml = `<div class="kpi-detail">(${kpi.graduation_detail})</div>`;
            } else if (ind.key === 'retention_rate' && kpi.retention_detail) {
                detailHtml = `<div class="kpi-detail">(${kpi.retention_detail})</div>`;
            } else if (ind.key === 'avg_time_to_graduate' && (d.avg_time_to_graduate_count || 0) > 0) {
                detailHtml = `<div class="kpi-detail">(من ${fmtNum(d.avg_time_to_graduate_count)} خريج)</div>`;
            }
            const surveyEvidence = getSurveyEvidence(d, ind.key);
            if (surveyEvidence) {
                const countHtml = surveyEvidence.count > 0
                    ? `<span class="kpi-survey-count">عدد المشاركين: ${fmtNum(surveyEvidence.count)}</span>`
                    : '';
                detailHtml += `<div class="kpi-survey-meta">
                    <span class="kpi-survey-badge ${surveyEvidence.kind}">${surveyEvidence.label}</span>
                    ${countHtml}
                </div>`;
            }
            return `<div class="kpi-card">
                <div class="kpi-num">${ind.code || ind.id}</div>
                <div class="kpi-body">
                    <div class="kpi-name">${ind.name}</div>
                    <div class="kpi-val ${f.cls}">${f.text} <span class="kpi-unit">${v != null ? ind.unit : ''}</span></div>
                    ${detailHtml}
                </div>
            </div>`;
        }).join('');

    const selfStudyWrap = document.getElementById('self-study-wrap');
    if (selfStudyWrap) selfStudyWrap.innerHTML = renderSelfStudyReport(selfStudyReport);

    // Trend chart
    renderTrendChart(prog, branch);

    document.getElementById('prog-results').classList.remove('hidden');
    document.getElementById('prog-results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderTrendChart(prog, branch = ALL_BRANCH_FILTER_VALUE) {
    destroyChart('trend');
    const years = Object.keys(prog.years).map(Number).filter(y => DISPLAY_YEARS.includes(y)).sort();
    const labels = years.map(y => fmtYear(y));
    const scopedRows = years.map(y => buildProgramDisplayData(prog, y, branch));
    const branchScoped = branch !== ALL_BRANCH_FILTER_VALUE && isIslamicStudiesBachelor(prog);
    const students = scopedRows.map(row => row?.students_total ?? null);
    const grads = scopedRows.map(row => branchScoped
        ? (row?.graduates_branch_matched ?? null)
        : (row?.graduates_total ?? null));
    const newS = scopedRows.map(row => row?.students_new ?? null);

    const ctx = document.getElementById('chart-trend').getContext('2d');
    charts.trend = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: branchScoped ? 'المنتظمون المرصودون' : 'الطلاب المنتظمون', data: students, borderColor: '#0d8e8e', backgroundColor: 'rgba(13,142,142,0.1)', fill: true, tension: 0.3, pointRadius: 5 },
                { label: branchScoped ? 'خريجون أمكن إسنادهم للفرع' : 'الخريجين', data: grads, borderColor: '#10b981', backgroundColor: 'transparent', tension: 0.3, pointRadius: 5 },
                { label: branchScoped ? 'المستجدون المرصودون' : 'المستجدون', data: newS, borderColor: '#3b82f6', backgroundColor: 'transparent', tension: 0.3, pointRadius: 5 },
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'top', labels: { font: { family: 'Tajawal' } } } },
            scales: {
                x: { ticks: { font: { family: 'Tajawal' } } },
                y: { beginAtZero: true, grid: { color: '#f0f0f0' }, ticks: { font: { family: 'Tajawal' } } }
            }
        }
    });
}

// ========================================
// المقارنة
// ========================================
function populateCompareBranchFilter(slot) {
    const select = document.getElementById(`cmp-${slot}-branch`);
    const programSelect = document.getElementById(`cmp-${slot}-prog`);
    if (!select) return;
    const program = programSelect?.value !== '' ? programs[parseInt(programSelect.value, 10)] : null;
    const branches = program ? getBranchesForProgram(program) : [];
    select.innerHTML = buildBranchOptionsHtml(branches);
    if (!program) {
        select.value = ALL_BRANCH_FILTER_VALUE;
        select.disabled = true;
    } else if (isIslamicStudiesBachelor(program)) {
        select.value = ALL_BRANCH_FILTER_VALUE;
        select.disabled = false;
    } else {
        select.value = 'الحوية';
        select.disabled = true;
    }
}

function populateCompareBranchFilters() {
    ['a', 'b', 'c'].forEach(slot => {
        populateCompareBranchFilter(slot);
    });
}

function initCompare() {
    const progOpts = '<option value="">-- اختر البرنامج --</option>' +
        programs.map((p,i) => `<option value="${i}">${p.name} (${p.degree}) - ${p.dept}</option>`).join('');
    ['cmp-a-prog','cmp-b-prog','cmp-c-prog'].forEach(id => {
        document.getElementById(id).innerHTML = progOpts;
    });
    populateCompareBranchFilters();

    ['a','b','c'].forEach(slot => {
        document.getElementById(`cmp-${slot}-prog`).addEventListener('change', () => {
            populateCompareBranchFilter(slot);
            populateCompareYears(slot);
            document.getElementById('cmp-results').classList.add('hidden');
            updateCmpBtn();
        });
        document.getElementById(`cmp-${slot}-year`).addEventListener('change', () => {
            document.getElementById('cmp-results').classList.add('hidden');
            updateCmpBtn();
        });
        document.getElementById(`cmp-${slot}-branch`)?.addEventListener('change', () => {
            document.getElementById('cmp-results').classList.add('hidden');
            updateCmpBtn();
        });
    });

    document.getElementById('cmp-add-third').addEventListener('click', () => {
        compareThirdEnabled = true;
        document.getElementById('cmp-slot-c').classList.remove('hidden');
        document.getElementById('cmp-add-third').classList.add('hidden');
        document.getElementById('cmp-remove-third').classList.remove('hidden');
        updateCmpBtn();
    });

    document.getElementById('cmp-remove-third').addEventListener('click', () => {
        compareThirdEnabled = false;
        document.getElementById('cmp-slot-c').classList.add('hidden');
        document.getElementById('cmp-add-third').classList.remove('hidden');
        document.getElementById('cmp-remove-third').classList.add('hidden');
        document.getElementById('cmp-c-prog').value = '';
        const cYear = document.getElementById('cmp-c-year');
        cYear.innerHTML = '<option value="">-- اختر السنة --</option>';
        cYear.disabled = true;
        cYear.value = '';
        const cBranch = document.getElementById('cmp-c-branch');
        if (cBranch) cBranch.value = ALL_BRANCH_FILTER_VALUE;
        document.getElementById('cmp-results').classList.add('hidden');
        updateCmpBtn();
    });

    document.getElementById('cmp-btn').addEventListener('click', showComparison);
    updateCmpBtn();
}

function populateCompareYears(slot) {
    const progSel = document.getElementById(`cmp-${slot}-prog`);
    const yearSel = document.getElementById(`cmp-${slot}-year`);
    yearSel.value = '';
    if (!progSel.value) {
        yearSel.innerHTML = '<option value="">-- اختر السنة --</option>';
        yearSel.disabled = true;
        return;
    }
    const idx = parseInt(progSel.value, 10);
    const p = programs[idx];
    if (!p) {
        yearSel.innerHTML = '<option value="">-- اختر السنة --</option>';
        yearSel.disabled = true;
        return;
    }
    const years = Object.keys(p.years).map(Number).filter(y => DISPLAY_YEARS.includes(y)).sort((a,b) => a - b);
    yearSel.innerHTML = '<option value="">-- اختر السنة --</option>' +
        years.map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');
    yearSel.disabled = false;
}

function getCompareSelection(slot) {
    const progValue = document.getElementById(`cmp-${slot}-prog`).value;
    const yearValue = document.getElementById(`cmp-${slot}-year`).value;
    const branchValue = normalizeBranchName(document.getElementById(`cmp-${slot}-branch`)?.value || ALL_BRANCH_FILTER_VALUE) || ALL_BRANCH_FILTER_VALUE;
    const hasProgram = progValue !== '';
    const hasYear = yearValue !== '';
    if (!hasProgram || !hasYear) return null;
    const programIndex = parseInt(progValue, 10);
    const year = parseInt(yearValue, 10);
    const program = programs[programIndex];
    if (!program || !program.years[year]) return null;
    const data = buildProgramDisplayData(program, year, branchValue);
    return {
        slot,
        programIndex,
        program,
        year,
        branch: branchValue,
        branchLabel: branchValue === ALL_BRANCH_FILTER_VALUE ? ALL_BRANCH_FILTER_LABEL : branchValue,
        data
    };
}

function updateCmpBtn() {
    const first = getCompareSelection('a');
    const second = getCompareSelection('b');
    const third = compareThirdEnabled ? getCompareSelection('c') : { slot: 'c' };
    const valid = !!(first && second && (!compareThirdEnabled || third));
    document.getElementById('cmp-btn').disabled = !valid;
}

function isGradDegree(deg) {
    return ['الماجستير','دكتوراه'].includes(String(deg || '').trim());
}

function getIndicatorsForDegree(degree) {
    if (isGradDegree(degree)) return INDICATORS_PG;
    return INDICATORS_UG.filter(ind => !ind.gradOnly);
}

function getIndicatorLabel(indicator) {
    const code = String(indicator.code || '').trim();
    return code ? `${code} - ${indicator.name}` : indicator.name;
}

function compareShortLabel(entry) {
    const branchSuffix = entry.branch && entry.branch !== ALL_BRANCH_FILTER_VALUE
        ? ` - ${entry.branch}`
        : '';
    return `${entry.program.name} - ${fmtYear(entry.year)}${branchSuffix}`;
}

function buildComparisonModel(entries) {
    const allGrad = entries.length > 0 && entries.every(e => isGradDegree(e.program.degree));
    const indicators = allGrad ? INDICATORS_PG : INDICATORS_UG.filter(ind => !ind.gradOnly);
    const header = [
        'المؤشر',
        ...entries.map(e => e.label),
        ...entries.slice(1).map((e, idx) => `الفرق (${idx + 2} - 1)`)
    ];
    const rows = indicators.map(ind => {
        const rawValues = entries.map(e => e.kpi[ind.key]);
        const formattedValues = rawValues.map(v => fmtKPI(v, ind.unit).text);
        const surveyEvidence = entries.map(e => getSurveyEvidence(e.data, ind.key));
        const diffs = [];
        for (let i = 1; i < rawValues.length; i++) {
            const base = rawValues[0];
            const val = rawValues[i];
            if (ind.numeric && base != null && val != null) {
                diffs.push(Math.round((parseFloat(val) - parseFloat(base)) * 100) / 100);
            } else {
                diffs.push(null);
            }
        }
        return { indicator: ind, rawValues, formattedValues, surveyEvidence, diffs };
    });
    return { header, rows, indicators };
}

function formatDiffCell(diff) {
    if (diff == null) return '<span class="diff-same">—</span>';
    if (diff > 0) return `<span class="diff-up">+${fmtNumFlex(diff, 2)}</span>`;
    if (diff < 0) return `<span class="diff-down">${fmtNumFlex(diff, 2)}</span>`;
    return '<span class="diff-same">0</span>';
}

function buildBranchOptionsHtml(branches = availableFacultyBranches) {
    return [
        `<option value="${ALL_BRANCH_FILTER_VALUE}">${ALL_BRANCH_FILTER_LABEL}</option>`,
        ...branches.map(branch => `<option value="${escapeHTML(branch)}">${escapeHTML(branch)}</option>`)
    ].join('');
}

function showComparison() {
    const selections = [getCompareSelection('a'), getCompareSelection('b')];
    if (compareThirdEnabled) selections.push(getCompareSelection('c'));
    if (selections.some(s => !s)) return alert('أكمل اختيار البرنامج والسنة لكل مقارنة مطلوبة');

    const entries = selections.map(sel => ({
        ...sel,
        label: `${sel.program.name} (${sel.program.degree}) - ${fmtYear(sel.year)}${sel.branch !== ALL_BRANCH_FILTER_VALUE ? ` - ${sel.branchLabel}` : ''}`,
        shortLabel: compareShortLabel(sel),
        kpi: calcKPIs(sel.data, sel.program.degree)
    }));

    const model = buildComparisonModel(entries);

    document.getElementById('cmp-thead').innerHTML =
        `<tr>${model.header.map(h => `<th>${h}</th>`).join('')}</tr>`;

    document.getElementById('cmp-tbody').innerHTML = model.rows.map(row => {
        const valueCells = row.formattedValues.map((v, index) => {
            const evidence = row.surveyEvidence[index];
            if (!evidence) return `<td>${v}</td>`;
            const countHtml = evidence.count > 0
                ? `<span class="kpi-survey-count">عدد المشاركين: ${fmtNum(evidence.count)}</span>`
                : '';
            return `<td>
                <div>${v}</div>
                <div class="kpi-survey-meta">
                    <span class="kpi-survey-badge ${evidence.kind}">${evidence.label}</span>
                    ${countHtml}
                </div>
            </td>`;
        }).join('');
        const diffCells = row.diffs.map(d => `<td>${formatDiffCell(d)}</td>`).join('');
        return `<tr><td>${getIndicatorLabel(row.indicator)}</td>${valueCells}${diffCells}</tr>`;
    }).join('');

    const chartLabels = [];
    const chartData = entries.map(() => []);
    model.rows.forEach(row => {
        if (!row.indicator.numeric) return;
        const chartLabel = row.indicator.code || row.indicator.name;
        chartLabels.push(chartLabel.length > 26 ? `${chartLabel.slice(0,26)}...` : chartLabel);
        row.rawValues.forEach((val, idx) => {
            chartData[idx].push(val == null ? null : parseFloat(val));
        });
    });

    destroyChart('compare');
    if (chartLabels.length > 0) {
        const ctx = document.getElementById('chart-compare').getContext('2d');
        charts.compare = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: chartLabels,
                datasets: entries.map((entry, idx) => ({
                    label: entry.shortLabel,
                    data: chartData[idx],
                    backgroundColor: CHART_COLORS[idx % CHART_COLORS.length],
                    borderRadius: 4
                }))
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top', labels: { font: { family: 'Tajawal' } } } },
                scales: {
                    x: { ticks: { font: { family: 'Tajawal', size: 10 }, maxRotation: 45 } },
                    y: { grid: { color: '#f0f0f0' }, ticks: { font: { family: 'Tajawal' } } }
                }
            }
        });
    }

    currentProg = { cmp: true, entries, model };
    document.getElementById('cmp-results').classList.remove('hidden');
    document.getElementById('cmp-results').scrollIntoView({ behavior: 'smooth' });
}

// ========================================
// التصدير
// ========================================
function safeFileName(text) {
    return String(text || '')
        .replace(/[\\/:*?"<>|]/g, '-')
        .replace(/\s+/g, '-');
}

function formatDiffText(diff) {
    if (diff == null) return '—';
    if (diff > 0) return `+${fmtNumFlex(diff, 2)}`;
    if (diff < 0) return fmtNumFlex(diff, 2);
    return '0';
}

function csvEscape(value) {
    const text = value == null ? '' : String(value);
    if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
    return text;
}

function downloadCSV(filename, rows) {
    const content = rows.map(row => row.map(csvEscape).join(',')).join('\n');
    const blob = new Blob([`\uFEFF${content}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}

async function exportSectionAsPDF(elementId, filename) {
    if (!window.jspdf || !window.html2canvas) {
        alert('ميزة PDF تحتاج تحميل مكتبة html2canvas بشكل صحيح.');
        return;
    }
    const target = document.getElementById(elementId);
    if (!target) return;

    const canvas = await window.html2canvas(target, {
        scale: 2.2,
        useCORS: true,
        backgroundColor: '#ffffff',
    });
    const imgData = canvas.toDataURL('image/jpeg', 0.95);
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 8;
    const contentWidth = pageWidth - (margin * 2);
    const contentHeight = (canvas.height * contentWidth) / canvas.width;
    const pageContentHeight = pageHeight - (margin * 2);

    let remaining = contentHeight;
    let position = margin;
    doc.addImage(imgData, 'JPEG', margin, position, contentWidth, contentHeight, undefined, 'FAST');
    remaining -= pageContentHeight;
    while (remaining > 0) {
        doc.addPage();
        position = margin - (contentHeight - remaining);
        doc.addImage(imgData, 'JPEG', margin, position, contentWidth, contentHeight, undefined, 'FAST');
        remaining -= pageContentHeight;
    }

    doc.save(filename);
}

async function withSelfStudyAccordionOpen(task) {
    const accordion = document.getElementById('self-study-details');
    const wasOpen = accordion ? accordion.open : false;
    if (accordion) {
        accordion.open = true;
        await new Promise(resolve => requestAnimationFrame(resolve));
    }
    try {
        return await task();
    } finally {
        if (accordion) accordion.open = wasOpen;
    }
}

function getProgramIndicatorRows() {
    if (!currentProg || currentProg.cmp) return [];
    const kpi = calcKPIs(currentProg.data, currentProg.prog.degree);
    return getIndicatorsForDegree(currentProg.prog.degree)
        .map(ind => {
            const f = fmtKPI(kpi[ind.key], ind.unit);
            const evidence = getSurveyEvidence(currentProg.data, ind.key);
            return [getIndicatorLabel(ind), f.text, ind.unit, formatSurveyEvidenceText(evidence)];
        });
}

async function exportPDF() {
    if (!currentProg || currentProg.cmp) return;
    const branchSuffix = currentProg.branch && currentProg.branch !== ALL_BRANCH_FILTER_VALUE
        ? `-${safeFileName(currentProg.branch)}`
        : '';
    const filename = `مؤشرات-${safeFileName(currentProg.prog.name)}-${fmtYear(currentProg.year)}${branchSuffix}.pdf`;
    await withSelfStudyAccordionOpen(() => exportSectionAsPDF('prog-results', filename));
}

function exportExcel() {
    if (!currentProg || currentProg.cmp) return;
    const indicatorRows = [
        ['تقرير مؤشرات الأداء'],
        ['البرنامج', currentProg.prog.name],
        ['الدرجة', currentProg.prog.degree],
        ['السنة', fmtYear(currentProg.year)],
        ['الفرع', currentProg.branchLabel || ALL_BRANCH_FILTER_LABEL],
        [],
        ['المؤشر', 'القيمة', 'الوحدة', 'نوع الاستطلاع'],
        ...getProgramIndicatorRows()
    ];
    const ws = XLSX.utils.aoa_to_sheet(indicatorRows);
    ws['!cols'] = [{ wch: 42 }, { wch: 18 }, { wch: 12 }, { wch: 34 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المؤشرات');
    if (currentProg.selfStudyReport) {
        const selfStudyRows = getSelfStudyExportRows(currentProg.selfStudyReport);
        const selfStudySheet = XLSX.utils.aoa_to_sheet(selfStudyRows);
        selfStudySheet['!cols'] = [
            { wch: 28 },
            { wch: 24 },
            { wch: 16 },
            { wch: 16 },
            { wch: 16 },
            { wch: 18 },
            { wch: 18 },
            { wch: 18 },
            { wch: 18 },
            { wch: 18 },
            { wch: 18 }
        ];
        XLSX.utils.book_append_sheet(wb, selfStudySheet, 'الدراسة الذاتية');
    }
    const branchSuffix = currentProg.branch && currentProg.branch !== ALL_BRANCH_FILTER_VALUE
        ? `-${safeFileName(currentProg.branch)}`
        : '';
    const filename = `مؤشرات-${safeFileName(currentProg.prog.name)}-${fmtYear(currentProg.year)}${branchSuffix}.xlsx`;
    XLSX.writeFile(wb, filename);
}

function exportCSV() {
    if (!currentProg || currentProg.cmp) return;
    const rows = [
        ['تقرير مؤشرات الأداء'],
        ['البرنامج', currentProg.prog.name],
        ['الدرجة', currentProg.prog.degree],
        ['السنة', fmtYear(currentProg.year)],
        ['الفرع', currentProg.branchLabel || ALL_BRANCH_FILTER_LABEL],
        [],
        ['المؤشر', 'القيمة', 'الوحدة', 'نوع الاستطلاع'],
        ...getProgramIndicatorRows(),
        [],
        ...getSelfStudyExportRows(currentProg.selfStudyReport)
    ];
    const branchSuffix = currentProg.branch && currentProg.branch !== ALL_BRANCH_FILTER_VALUE
        ? `-${safeFileName(currentProg.branch)}`
        : '';
    const filename = `مؤشرات-${safeFileName(currentProg.prog.name)}-${fmtYear(currentProg.year)}${branchSuffix}.csv`;
    downloadCSV(filename, rows);
}

function getCompareExportRows() {
    if (!currentProg || !currentProg.cmp || !currentProg.model) return [];
    return currentProg.model.rows.map(row => [
        getIndicatorLabel(row.indicator),
        ...row.formattedValues.map((value, index) => {
            const evidenceText = formatSurveyEvidenceText(row.surveyEvidence[index]);
            return evidenceText ? `${value} (${evidenceText})` : value;
        }),
        ...row.diffs.map(formatDiffText)
    ]);
}

async function exportComparePDF() {
    if (!currentProg || !currentProg.cmp) return;
    const filename = `مقارنة-المؤشرات-${Date.now()}.pdf`;
    await exportSectionAsPDF('cmp-export-area', filename);
}

function exportCompareExcel() {
    if (!currentProg || !currentProg.cmp || !currentProg.model) return;
    const rows = [
        ['مقارنة المؤشرات'],
        ...currentProg.entries.map((e, idx) => [`مدخل ${idx + 1}`, e.label]),
        [],
        currentProg.model.header,
        ...getCompareExportRows()
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const colCount = currentProg.model.header.length;
    ws['!cols'] = Array.from({ length: colCount }, (_, idx) => ({ wch: idx === 0 ? 40 : 20 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المقارنة');
    XLSX.writeFile(wb, 'مقارنة-المؤشرات.xlsx');
}

function exportCompareCSV() {
    if (!currentProg || !currentProg.cmp || !currentProg.model) return;
    const rows = [currentProg.model.header, ...getCompareExportRows()];
    downloadCSV('مقارنة-المؤشرات.csv', rows);
}

// ========================================
// قائمة أعضاء هيئة التدريس للأعوام 1445-1448
// ========================================
async function loadFacultyRoster() {
    try {
        const csvText = await fetchTextIfExists(`data/faculty.csv?t=${Date.now()}`);
        if (!csvText) {
            facultyRosterData = [];
            return false;
        }
        facultyRosterData = parseFlatCSV(csvText, ',')
            .map(row => ({
                year: parseInt(pickCell(row, ['year', 'Year']), 10) || null,
                id: String(pickCell(row, ['id', 'ID']) || '').trim(),
                name: String(pickCell(row, ['name', 'Name']) || '').trim(),
                rank: normalizeRank(pickCell(row, ['rank', 'Rank'])),
                email: String(pickCell(row, ['email', 'Email']) || '').trim(),
                active: String(pickCell(row, ['active', 'Active']) || '').trim(),
                department: normalizeDepartment(pickCell(row, ['department', 'Department'])),
                nationality: String(pickCell(row, ['nationality', 'Nationality', 'الجنسية']) || '').trim(),
                gender: String(pickCell(row, ['gender', 'Gender', 'الجنس']) || '').trim(),
                branch: getFacultyBranchNames(pickCell(row, ['branch', 'Branch', 'الفرع'])).join('|')
            }))
            .filter(row => [1445, 1446, 1447, 1448].includes(row.year) && row.id && row.name);
        return true;
    } catch (error) {
        console.error('خطأ في تحميل قائمة أعضاء هيئة التدريس:', error);
        facultyRosterData = [];
        return false;
    }
}

function initFacultyView() {
    const yearSelect = document.getElementById('faculty-year');
    const deptSelect = document.getElementById('faculty-dept');
    const branchSelect = document.getElementById('faculty-branch');
    const activeSelect = document.getElementById('faculty-active');
    const searchInput = document.getElementById('faculty-search');
    if (!yearSelect || !deptSelect || !branchSelect || !activeSelect || !searchInput) return;

    const years = [...new Set(facultyRosterData.map(row => row.year).filter(Boolean))].sort((a, b) => a - b);
    const departments = [...new Set(facultyRosterData.map(row => row.department).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'ar'));
    const branches = getOrderedBranchNames(facultyRosterData.flatMap(row => getFacultyBranchNames(row.branch)));
    yearSelect.innerHTML = '<option value="">الكل</option>' +
        years.map(year => `<option value="${year}">${fmtYear(year)}</option>`).join('');
    deptSelect.innerHTML = '<option value="">الكل</option>' +
        departments.map(department => `<option value="${escapeHTML(department)}">${escapeHTML(department)}</option>`).join('');
    branchSelect.innerHTML = '<option value="">الكل</option>' +
        branches.map(branch => `<option value="${escapeHTML(branch)}">${escapeHTML(branch)}</option>`).join('');

    [yearSelect, deptSelect, branchSelect, activeSelect].forEach(element =>
        element.addEventListener('change', renderFacultyRoster)
    );
    searchInput.addEventListener('input', renderFacultyRoster);
    renderFacultyRoster();
}

function getFilteredFacultyRoster() {
    const year = parseInt(document.getElementById('faculty-year')?.value, 10) || null;
    const department = document.getElementById('faculty-dept')?.value || '';
    const branch = document.getElementById('faculty-branch')?.value || '';
    const active = document.getElementById('faculty-active')?.value || '';
    const search = (document.getElementById('faculty-search')?.value || '').trim().toLowerCase();
    return facultyRosterData.filter(row => {
        if (year && row.year !== year) return false;
        if (department && row.department !== department) return false;
        if (branch && !getFacultyBranchNames(row.branch).includes(branch)) return false;
        if (active && row.active !== active) return false;
        if (!search) return true;
        return row.name.toLowerCase().includes(search) || row.id.includes(search) || row.email.toLowerCase().includes(search);
    });
}

function renderFacultyRoster() {
    const filtered = getFilteredFacultyRoster();
    const count = document.getElementById('faculty-count');
    if (count) count.textContent = `${filtered.length.toLocaleString('ar-SA')} سجل من أصل ${facultyRosterData.length.toLocaleString('ar-SA')}`;
    const tbody = document.getElementById('faculty-tbody');
    if (!tbody) return;
    const maxShow = 500;
    tbody.innerHTML = filtered.slice(0, maxShow).map((row, index) => `<tr class="${index % 2 ? 'alt' : ''}">
        <td>${index + 1}</td>
        <td>${escapeHTML(fmtYear(row.year))}</td>
        <td>${escapeHTML(row.id)}</td>
        <td>${escapeHTML(row.name)}</td>
        <td>${escapeHTML(row.rank)}</td>
        <td>${escapeHTML(row.department)}</td>
        <td>${escapeHTML(getFacultyBranchNames(row.branch).join('، ') || '—')}</td>
        <td>${escapeHTML(row.gender)}</td>
        <td>${escapeHTML(row.nationality)}</td>
        <td>${escapeHTML(row.email || '—')}</td>
        <td>${row.active === 'نعم' ? 'نشط' : 'غير نشط'}</td>
    </tr>`).join('');
    if (filtered.length > maxShow) {
        tbody.innerHTML += `<tr><td colspan="11" style="text-align:center;color:var(--text-light);padding:16px">
            يتم عرض أول ${maxShow} سجل. استخدم الفلاتر أو صدّر Excel لرؤية الكل.
        </td></tr>`;
    }
}

function exportFacultyExcel() {
    const rows = [
        ['السنة','الرقم الوظيفي','الاسم','الرتبة','القسم','الفرع','الجنس','الجنسية','البريد الإلكتروني','الحالة'],
        ...getFilteredFacultyRoster().map(row => [
            fmtYear(row.year), row.id, row.name, row.rank, row.department, getFacultyBranchNames(row.branch).join('، '),
            row.gender, row.nationality, row.email, row.active === 'نعم' ? 'نشط' : 'غير نشط'
        ])
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'هيئة التدريس');
    XLSX.writeFile(workbook, 'قائمة-هيئة-التدريس-1445-1448.xlsx');
}

// ========================================
// سجل المستجدين في فروع الدراسات الإسلامية
// ========================================
function formatAcademicSemester(value) {
    const semester = parseInt(value, 10);
    if (semester === 1) return 'الفصل الأول';
    if (semester === 2) return 'الفصل الثاني';
    if (semester === 3) return 'الفصل الصيفي';
    return '—';
}

function initEntrantsView() {
    const yearSelect = document.getElementById('entrant-year');
    const branchSelect = document.getElementById('entrant-branch');
    const searchInput = document.getElementById('entrant-search');
    if (!yearSelect || !branchSelect || !searchInput) return;

    const years = [...new Set(entrantData.map(row => String(row.year || '')).filter(Boolean))]
        .sort((a, b) => Number(a) - Number(b));
    const branches = getOrderedBranchNames(entrantData.map(row => row.branch));
    yearSelect.innerHTML = '<option value="">الكل</option>' +
        years.map(year => `<option value="${escapeHTML(year)}">${fmtYear(year)}</option>`).join('');
    branchSelect.innerHTML = '<option value="">الكل</option>' +
        branches.map(branch => `<option value="${escapeHTML(branch)}">${escapeHTML(branch)}</option>`).join('');

    [yearSelect, branchSelect].forEach(element => element.addEventListener('change', renderEntrants));
    searchInput.addEventListener('input', renderEntrants);
    renderEntrants();
}

function getFilteredEntrants() {
    const yearFilter = document.getElementById('entrant-year')?.value || '';
    const branchFilter = document.getElementById('entrant-branch')?.value || '';
    const search = (document.getElementById('entrant-search')?.value || '').trim().toLowerCase();
    return entrantData.filter(row => {
        if (yearFilter && String(row.year) !== yearFilter) return false;
        if (branchFilter && normalizeBranchName(row.branch) !== branchFilter) return false;
        if (!search) return true;
        return String(row.name || '').toLowerCase().includes(search) ||
            String(row.student_id || '').includes(search);
    });
}

function renderEntrants() {
    const filtered = getFilteredEntrants();
    const count = document.getElementById('entrant-count');
    if (count) {
        count.textContent = `${filtered.length.toLocaleString('ar-SA')} سجل من أصل ${entrantData.length.toLocaleString('ar-SA')}`;
    }
    const tbody = document.getElementById('entrant-tbody');
    if (!tbody) return;
    const maxShow = 500;
    tbody.innerHTML = filtered.slice(0, maxShow).map((row, index) => `<tr class="${index % 2 ? 'alt' : ''}">
        <td>${index + 1}</td>
        <td>${fmtYear(row.year)}</td>
        <td>${formatAcademicSemester(row.semester)}</td>
        <td>${escapeHTML(row.student_id)}</td>
        <td>${escapeHTML(row.name)}</td>
        <td>${escapeHTML(row.program)}</td>
        <td>${escapeHTML(row.degree)}</td>
        <td>${escapeHTML(row.branch)}</td>
        <td>${escapeHTML(row.gender)}</td>
    </tr>`).join('');
    if (filtered.length > maxShow) {
        tbody.innerHTML += `<tr><td colspan="9" style="text-align:center;color:var(--text-light);padding:16px">
            يتم عرض أول ${maxShow} سجل. استخدم الفلاتر أو صدّر Excel لرؤية الكل.
        </td></tr>`;
    }
}

function exportEntrantsExcel() {
    const filtered = getFilteredEntrants();
    const rows = [
        ['السنة','الفصل','الرقم الجامعي','الاسم','التخصص','الدرجة','القسم','الفرع','الجنس'],
        ...filtered.map(row => [
            fmtYear(row.year), formatAcademicSemester(row.semester), row.student_id, row.name,
            row.program, row.degree, row.department, row.branch, row.gender
        ])
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'المستجدون');
    XLSX.writeFile(workbook, 'سجل-المستجدين-حسب-الفرع.xlsx');
}

// ========================================
// سجل الخريجين
// ========================================
const STATUS_MAP = {
    'منسحب': { cls: 'st-withdrawn', label: 'منسحب' },
    'مؤجل': { cls: 'st-postponed', label: 'مؤجل' },
    'مؤجل قبول': { cls: 'st-postponed', label: 'مؤجل قبول' },
    'معتذر': { cls: 'st-excused', label: 'معتذر' },
    'منقطع عن الدراسة': { cls: 'st-absent', label: 'منقطع' },
    'مفصول اكاديميا': { cls: 'st-dismissed', label: 'مفصول' },
    'مطوي قيده': { cls: 'st-folded', label: 'مطوي قيده' },
    'موقوف تأديبي / مف': { cls: 'st-suspended', label: 'موقوف' },
    'متوفى': { cls: 'st-deceased', label: 'متوفى' },
};

function parseDetailCSV(text) {
    const lines = text.trim().split('\n').map(l => l.replace(/\r/g,''));
    if (lines.length < 2) return [];
    const sep = lines[0].includes(';') ? ';' : ',';
    const headers = lines[0].split(sep).map(h => h.trim().replace(/^\uFEFF/,''));
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        const vals = lines[i].split(sep);
        if (vals.length < 3) continue;
        const row = {};
        headers.forEach((h, j) => { row[h] = (vals[j] || '').trim(); });
        rows.push(row);
    }
    return rows;
}

async function loadGraduates() {
    try {
        const res = await fetch('data/graduates_detail.csv?t=' + Date.now());
        const csv = await res.text();
        gradData = parseDetailCSV(csv).map(row => ({
            ...row,
            'الفرع': getEffectiveStudentBranch(row['التخصص'], row['الدرجة'], row['الفرع'])
        }));
        mergeIslamicBranchGraduates();
        return true;
    } catch (e) {
        console.error('خطأ في تحميل بيانات الخريجين:', e);
        return false;
    }
}

function mergeIslamicBranchGraduates() {
    const accepted = Array.isArray(islamicBranchData?.graduates)
        ? islamicBranchData.graduates.filter(row =>
            ['official', 'very_high', 'high'].includes(String(row.confidence || ''))
        )
        : [];
    if (!accepted.length) return;

    const keyFor = (year, studentId, program, degree) => [
        parseInt(year, 10) || 0,
        String(studentId || '').trim(),
        normalizeSurveyProgramName(program || ''),
        normalizeDegree(degree || '')
    ].join('|');
    const branchRowsByKey = new Map(accepted.map(row => [
        keyFor(row.year, row.student_id, row.program, row.degree),
        row
    ]));

    gradData = gradData.map(row => {
        if (normalizeSurveyProgramName(row['التخصص']) !== 'الدراسات الإسلامية') return row;
        if (normalizeDegree(row['الدرجة']) !== 'بكالوريوس') return row;
        const matched = branchRowsByKey.get(keyFor(
            row['السنة'], row['الرقم_الجامعي'], row['التخصص'], row['الدرجة']
        ));
        if (!matched) return row;
        return {
            ...row,
            'الفرع': String(matched.branch || '').trim(),
            _confidence: matched.confidence,
            _branchConfidence: matched.branch_confidence || '',
            _branchSource: matched.source || ''
        };
    });

    const existing = new Set(gradData.map(row => keyFor(
        row['السنة'], row['الرقم_الجامعي'], row['التخصص'], row['الدرجة']
    )));
    accepted.forEach(row => {
        const key = keyFor(row.year, row.student_id, row.program, row.degree);
        if (existing.has(key)) return;
        gradData.push({
            'السنة': String(row.year ?? ''),
            'الرقم_الجامعي': String(row.student_id || ''),
            'الاسم': String(row.name || ''),
            'التخصص': String(row.program || 'الدراسات الإسلامية'),
            'الدرجة': String(row.degree || 'بكالوريوس'),
            'القسم': String(row.department || 'الدراسات الإسلامية'),
            'الفرع': String(row.branch || ''),
            'الجنس': String(row.gender || ''),
            'الجنسية': String(row.nationality || ''),
            'تاريخ_القبول': String(row.official_admission_date || row.admission_date || ''),
            'تاريخ_التخرج': String(row.official_graduation_date || row.graduation_date || ''),
            'تاريخ_التخرج_المتوقع': '',
            'المعدل': String(row.official_gpa || row.gpa || ''),
            _confidence: row.confidence,
            _branchConfidence: row.branch_confidence || '',
            _branchSource: row.source || ''
        });
        existing.add(key);
    });
}

async function loadNonCompleters() {
    try {
        const res = await fetch('data/non_completers.csv?t=' + Date.now());
        const csv = await res.text();
        ncData = parseDetailCSV(csv).map(row => ({
            ...row,
            'الفرع': getEffectiveStudentBranch(row['التخصص'], row['الدرجة'], row['الفرع'])
        }));
        mergeIslamicBranchNonCompleters();
        return true;
    } catch (e) {
        console.error('خطأ في تحميل بيانات غير المكملين:', e);
        return false;
    }
}

function mergeIslamicBranchNonCompleters() {
    const accepted = Array.isArray(islamicBranchData?.noncompleters)
        ? islamicBranchData.noncompleters.filter(row =>
            ['official', 'very_high', 'high'].includes(String(row.event_confidence || row.confidence || '')) &&
            ['very_high', 'high'].includes(String(row.branch_confidence || ''))
        )
        : [];
    if (!accepted.length) return;
    const keyFor = (year, studentId, program, degree) => [
        parseInt(year, 10) || 0,
        String(studentId || '').trim(),
        normalizeSurveyProgramName(program || ''),
        normalizeDegree(degree || '')
    ].join('|');
    const mapped = new Map(accepted.map(row => [
        keyFor(row.year, row.student_id, row.program, row.degree),
        row
    ]));
    ncData = ncData.map(row => {
        const match = mapped.get(keyFor(
            row['آخر_سنة'], row['الرقم_الجامعي'], row['التخصص'], row['الدرجة']
        ));
        if (!match) return row;
        return {
            ...row,
            'الفرع': normalizeBranchName(match.branch),
            _branchConfidence: match.branch_confidence
        };
    });
}

async function loadStudentDetails() {
    try {
        const res = await fetch('data/students_detail.csv?t=' + Date.now());
        const csv = await res.text();
        studentDetailData = parseDetailCSV(csv).map(row => {
            const normalizedProgram = normalizeSurveyProgramName(row['التخصص'] || '');
            const normalizedDegree = normalizeDegree(row['الدرجة'] || '');
            const normalizedDept = normalizeDepartment(row['القسم'] || '');
            const gpa = analyticsParseGPA(row['المعدل']);
            const categoryId = classifyStudentCategoryId(gpa);
            const branch = getEffectiveStudentBranch(normalizedProgram, normalizedDegree, row['الفرع']);
            const searchText = normalizeArabicText([
                row['السنة'],
                row['الرقم_الجامعي'],
                row['الاسم'],
                normalizedProgram,
                normalizedDegree,
                normalizedDept,
                branch,
                row['الحالة'],
                row['الجنس'],
                row['الجنسية']
            ].join(' ')).toLowerCase();
            return {
                ...row,
                'التخصص': normalizedProgram,
                'الدرجة': normalizedDegree,
                'القسم': normalizedDept,
                'الفرع': branch,
                _year: parseInt(row['السنة'], 10) || 0,
                _gpa: gpa,
                _categoryId: categoryId,
                _officialEstimate: getOfficialGpaEstimate(gpa),
                _searchText: searchText
            };
        });
        mergeStudentDetailBranches();
        return true;
    } catch (e) {
        console.error('خطأ في تحميل بيانات الطلاب التفصيلية:', e);
        return false;
    }
}

function mergeStudentDetailBranches() {
    const accepted = Array.isArray(islamicBranchData?.student_detail_branches)
        ? islamicBranchData.student_detail_branches.filter(row =>
            ['very_high', 'high'].includes(String(row.branch_confidence || ''))
        )
        : [];
    if (!accepted.length) return;
    const keyFor = (year, studentId, program, degree) => [
        parseInt(year, 10) || 0,
        String(studentId || '').trim(),
        normalizeSurveyProgramName(program || ''),
        normalizeDegree(degree || '')
    ].join('|');
    const mapped = new Map(accepted.map(row => [
        keyFor(row.year, row.student_id, row.program, row.degree),
        row
    ]));
    studentDetailData = studentDetailData.map(row => {
        const match = mapped.get(keyFor(
            row['السنة'], row['الرقم_الجامعي'], row['التخصص'], row['الدرجة']
        ));
        if (!match) return row;
        const branch = normalizeBranchName(match.branch);
        return {
            ...row,
            'الفرع': branch,
            _branchConfidence: match.branch_confidence,
            _searchText: `${row._searchText} ${normalizeArabicText(branch).toLowerCase()}`
        };
    });
}

function initGraduatesView() {
    if (!gradData.length) return;

    // Populate filters
    const years = [...new Set(gradData.map(g => g['السنة']))].sort();
    const progs = [...new Set(gradData.map(g => g['التخصص']))].sort((a,b) => a.localeCompare(b,'ar'));
    const degs = [...new Set(gradData.map(g => g['الدرجة']))];
    const branches = getOrderedBranchNames(gradData.map(row => row['الفرع']));

    const ySel = document.getElementById('grad-year');
    ySel.innerHTML = '<option value="">الكل</option>' +
        years.map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');

    const pSel = document.getElementById('grad-prog');
    pSel.innerHTML = '<option value="">الكل</option>' +
        progs.map(p => `<option value="${p}">${p}</option>`).join('');

    const dSel = document.getElementById('grad-deg');
    dSel.innerHTML = '<option value="">الكل</option>' +
        degs.map(d => `<option value="${d}">${d}</option>`).join('');

    const bSel = document.getElementById('grad-branch');
    bSel.innerHTML = '<option value="">الكل</option>' +
        branches.map(branch => `<option value="${escapeHTML(branch)}">${escapeHTML(branch)}</option>`).join('');

    // Event listeners
    [ySel, pSel, dSel, bSel].forEach(el => el.addEventListener('change', renderGraduates));
    document.getElementById('grad-search').addEventListener('input', renderGraduates);

    renderGraduates();
}

function renderGraduates() {
    const yearFilter = document.getElementById('grad-year').value;
    const progFilter = document.getElementById('grad-prog').value;
    const degFilter = document.getElementById('grad-deg').value;
    const branchFilter = document.getElementById('grad-branch').value;
    const search = document.getElementById('grad-search').value.trim().toLowerCase();

    let filtered = gradData;
    if (yearFilter) filtered = filtered.filter(g => g['السنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (branchFilter) filtered = filtered.filter(g => normalizeBranchName(g['الفرع']) === branchFilter);
    if (search) filtered = filtered.filter(g =>
        String(g['الاسم'] || '').toLowerCase().includes(search) ||
        String(g['الرقم_الجامعي'] || '').includes(search)
    );

    document.getElementById('grad-count').textContent =
        `${filtered.length.toLocaleString('ar-SA')} سجل من أصل ${gradData.length.toLocaleString('ar-SA')}`;

    const tbody = document.getElementById('grad-tbody');
    const MAX_SHOW = 500;
    const showing = filtered.slice(0, MAX_SHOW);

    tbody.innerHTML = showing.map((g, i) => `<tr class="${i%2?'alt':''}">
        <td>${i+1}</td>
        <td>${escapeHTML(fmtYear(g['السنة']))}</td>
        <td>${escapeHTML(g['الرقم_الجامعي'])}</td>
        <td>${escapeHTML(g['الاسم'])}</td>
        <td>${escapeHTML(g['التخصص'])}</td>
        <td>${escapeHTML(g['الدرجة'])}</td>
        <td>${escapeHTML(g['الفرع'] || '—')}</td>
        <td>${escapeHTML(g['الجنس'])}</td>
        <td>${escapeHTML(g['الجنسية'])}</td>
        <td>${escapeHTML(g['تاريخ_القبول'])}</td>
        <td>${escapeHTML(g['تاريخ_التخرج'])}</td>
        <td>${escapeHTML(g['المعدل'])}</td>
    </tr>`).join('');

    if (filtered.length > MAX_SHOW) {
        tbody.innerHTML += `<tr><td colspan="12" style="text-align:center;color:var(--text-light);padding:16px">
            يتم عرض أول ${MAX_SHOW} سجل. استخدم الفلاتر لتصفية النتائج أو صدّر Excel لرؤية الكل.
        </td></tr>`;
    }
}

function exportGradsExcel() {
    const yearFilter = document.getElementById('grad-year').value;
    const progFilter = document.getElementById('grad-prog').value;
    const degFilter = document.getElementById('grad-deg').value;
    const branchFilter = document.getElementById('grad-branch').value;
    const search = document.getElementById('grad-search').value.trim().toLowerCase();

    let filtered = gradData;
    if (yearFilter) filtered = filtered.filter(g => g['السنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (branchFilter) filtered = filtered.filter(g => normalizeBranchName(g['الفرع']) === branchFilter);
    if (search) filtered = filtered.filter(g =>
        String(g['الاسم'] || '').toLowerCase().includes(search) || String(g['الرقم_الجامعي'] || '').includes(search)
    );

    const rows = [
        ['السنة','الرقم الجامعي','الاسم','التخصص','الدرجة','القسم','الفرع','الجنس','الجنسية','تاريخ القبول','تاريخ التخرج','المعدل'],
        ...filtered.map(g => [
            fmtYear(g['السنة']), g['الرقم_الجامعي'], g['الاسم'], g['التخصص'], g['الدرجة'],
            g['القسم'], g['الفرع'] || '', g['الجنس'], g['الجنسية'], g['تاريخ_القبول'], g['تاريخ_التخرج'], g['المعدل']
        ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'الخريجين');
    XLSX.writeFile(wb, `سجل-الخريجين${yearFilter ? '-'+fmtYear(yearFilter) : ''}.xlsx`);
}

// ========================================
// غير المكملين
// ========================================
function initNonCompleteView() {
    if (!ncData.length) return;

    const years = [...new Set(ncData.map(g => g['آخر_سنة']))].sort();
    const progs = [...new Set(ncData.map(g => g['التخصص']))].sort((a,b) => a.localeCompare(b,'ar'));
    const degs = [...new Set(ncData.map(g => g['الدرجة']))];
    const statuses = [...new Set(ncData.map(g => g['الحالة']))].sort((a,b) => a.localeCompare(b,'ar'));
    const branches = getOrderedBranchNames(ncData.map(row => row['الفرع']));

    document.getElementById('nc-year').innerHTML = '<option value="">الكل</option>' +
        years.map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');
    document.getElementById('nc-prog').innerHTML = '<option value="">الكل</option>' +
        progs.map(p => `<option value="${p}">${p}</option>`).join('');
    document.getElementById('nc-deg').innerHTML = '<option value="">الكل</option>' +
        degs.map(d => `<option value="${d}">${d}</option>`).join('');
    document.getElementById('nc-status').innerHTML = '<option value="">الكل</option>' +
        statuses.map(s => `<option value="${s}">${s}</option>`).join('');
    document.getElementById('nc-branch').innerHTML = '<option value="">الكل</option>' +
        branches.map(branch => `<option value="${escapeHTML(branch)}">${escapeHTML(branch)}</option>`).join('');

    ['nc-year','nc-prog','nc-deg','nc-status','nc-branch'].forEach(id =>
        document.getElementById(id).addEventListener('change', renderNonCompleters));
    document.getElementById('nc-search').addEventListener('input', renderNonCompleters);

    renderNonCompleters();
}

function renderNonCompleters() {
    const yearFilter = document.getElementById('nc-year').value;
    const progFilter = document.getElementById('nc-prog').value;
    const degFilter = document.getElementById('nc-deg').value;
    const statusFilter = document.getElementById('nc-status').value;
    const branchFilter = document.getElementById('nc-branch').value;
    const search = document.getElementById('nc-search').value.trim().toLowerCase();

    let filtered = ncData;
    if (yearFilter) filtered = filtered.filter(g => g['آخر_سنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (statusFilter) filtered = filtered.filter(g => g['الحالة'] === statusFilter);
    if (branchFilter) filtered = filtered.filter(g => normalizeBranchName(g['الفرع']) === branchFilter);
    if (search) filtered = filtered.filter(g =>
        String(g['الاسم'] || '').toLowerCase().includes(search) || String(g['الرقم_الجامعي'] || '').includes(search)
    );

    document.getElementById('nc-count').textContent =
        `${filtered.length.toLocaleString('ar-SA')} سجل من أصل ${ncData.length.toLocaleString('ar-SA')}`;

    // Status summary cards
    const statusCounts = {};
    filtered.forEach(nc => {
        statusCounts[nc['الحالة']] = (statusCounts[nc['الحالة']] || 0) + 1;
    });
    const cardsDiv = document.getElementById('nc-summary-cards');
    cardsDiv.innerHTML = Object.entries(statusCounts)
        .sort((a,b) => b[1] - a[1])
        .map(([st, cnt]) => {
            const info = STATUS_MAP[st] || { cls: '', label: st };
            return `<div class="status-card">
                <span class="sc-count">${cnt.toLocaleString('ar-SA')}</span>
                <span class="sc-label"><span class="status-badge ${info.cls}">${escapeHTML(info.label)}</span></span>
            </div>`;
        }).join('');

    // Table
    const tbody = document.getElementById('nc-tbody');
    const MAX_SHOW = 500;
    const showing = filtered.slice(0, MAX_SHOW);

    tbody.innerHTML = showing.map((nc, i) => {
        const info = STATUS_MAP[nc['الحالة']] || { cls: '', label: nc['الحالة'] };
        return `<tr class="${i%2?'alt':''}">
            <td>${i+1}</td>
            <td>${escapeHTML(fmtYear(nc['آخر_سنة']))}</td>
            <td>${escapeHTML(nc['الرقم_الجامعي'])}</td>
            <td>${escapeHTML(nc['الاسم'])}</td>
            <td>${escapeHTML(nc['التخصص'])}</td>
            <td>${escapeHTML(nc['الدرجة'])}</td>
            <td>${escapeHTML(nc['الفرع'] || '—')}</td>
            <td><span class="status-badge ${info.cls}">${escapeHTML(nc['الحالة'])}</span></td>
            <td>${escapeHTML(nc['الجنس'])}</td>
            <td>${escapeHTML(nc['الجنسية'])}</td>
            <td>${escapeHTML(nc['تاريخ_القبول'])}</td>
            <td>${escapeHTML(nc['المعدل'])}</td>
        </tr>`;
    }).join('');

    if (filtered.length > MAX_SHOW) {
        tbody.innerHTML += `<tr><td colspan="12" style="text-align:center;color:var(--text-light);padding:16px">
            يتم عرض أول ${MAX_SHOW} سجل. استخدم الفلاتر لتصفية النتائج أو صدّر Excel لرؤية الكل.
        </td></tr>`;
    }
}

function exportNCExcel() {
    const yearFilter = document.getElementById('nc-year').value;
    const progFilter = document.getElementById('nc-prog').value;
    const degFilter = document.getElementById('nc-deg').value;
    const statusFilter = document.getElementById('nc-status').value;
    const branchFilter = document.getElementById('nc-branch').value;
    const search = document.getElementById('nc-search').value.trim().toLowerCase();

    let filtered = ncData;
    if (yearFilter) filtered = filtered.filter(g => g['آخر_سنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (statusFilter) filtered = filtered.filter(g => g['الحالة'] === statusFilter);
    if (branchFilter) filtered = filtered.filter(g => normalizeBranchName(g['الفرع']) === branchFilter);
    if (search) filtered = filtered.filter(g =>
        String(g['الاسم'] || '').toLowerCase().includes(search) || String(g['الرقم_الجامعي'] || '').includes(search)
    );

    const rows = [
        ['آخر سنة','الرقم الجامعي','الاسم','التخصص','الدرجة','القسم','الفرع','الحالة','الجنس','الجنسية','تاريخ القبول','المعدل','نوع الدراسة'],
        ...filtered.map(nc => [
            fmtYear(nc['آخر_سنة']), nc['الرقم_الجامعي'], nc['الاسم'], nc['التخصص'], nc['الدرجة'],
            nc['القسم'], nc['الفرع'] || '', nc['الحالة'], nc['الجنس'], nc['الجنسية'], nc['تاريخ_القبول'], nc['المعدل'], nc['نوع_الدراسة']
        ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'غير المكملين');
    XLSX.writeFile(wb, `غير-المكملين${statusFilter ? '-'+statusFilter : ''}.xlsx`);
}

// ========================================
// فئات الطلاب
// ========================================
function setSelectOptions(selectElement, options, placeholder = 'الكل', currentValue = '') {
    if (!selectElement) return;
    const previous = currentValue || selectElement.value || '';
    selectElement.innerHTML = `<option value="">${placeholder}</option>` +
        options.map(option => `<option value="${option.value}">${option.label}</option>`).join('');

    const exists = options.some(option => option.value === previous);
    selectElement.value = exists ? previous : '';
}

function buildStudentCategoryLabel(row) {
    const category = getStudentCategoryDefinition(row._categoryId);
    return category ? category.label : 'خارج الفئات المحددة';
}

function getSelectedStudentCategoryIds() {
    return [...document.querySelectorAll('#student-category-picks input[type="checkbox"]:checked')]
        .map(input => input.value)
        .filter(Boolean);
}

function renderStudentCategoryChoices() {
    const wrap = document.getElementById('student-category-picks');
    if (!wrap) return;
    wrap.innerHTML = STUDENT_CATEGORY_DEFS.map(category => `
        <label class="student-category-chip is-checked" data-category-id="${category.id}">
            <input type="checkbox" value="${category.id}" checked>
            <span class="student-category-chip-body">
                <span class="student-category-chip-title">${category.label}</span>
                <span class="student-category-chip-meta">${category.rangeLabel}</span>
            </span>
        </label>
    `).join('');

    wrap.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', () => {
            input.closest('.student-category-chip')?.classList.toggle('is-checked', input.checked);
            document.getElementById('student-cat-results')?.classList.add('hidden');
        });
    });
}

function refreshStudentCategoryProgramOptions() {
    const year = document.getElementById('student-cat-year')?.value || '';
    const dept = document.getElementById('student-cat-dept')?.value || '';
    const degree = document.getElementById('student-cat-degree')?.value || '';
    const programSelect = document.getElementById('student-cat-prog');
    if (!programSelect) return;

    let rows = studentDetailData;
    if (year) rows = rows.filter(row => String(row['السنة']) === year);
    if (dept) rows = rows.filter(row => row['القسم'] === dept);
    if (degree) rows = rows.filter(row => row['الدرجة'] === degree);

    const options = [...new Set(rows.map(row => row['التخصص']).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'ar'))
        .map(value => ({ value, label: value }));

    setSelectOptions(programSelect, options, 'كل البرامج');
}

function refreshStudentCategoryBaseFilters() {
    const yearSelect = document.getElementById('student-cat-year');
    const deptSelect = document.getElementById('student-cat-dept');
    const degreeSelect = document.getElementById('student-cat-degree');
    const branchSelect = document.getElementById('student-cat-branch');
    const statusSelect = document.getElementById('student-cat-status');
    const genderSelect = document.getElementById('student-cat-gender');

    const years = [...new Set(studentDetailData.map(row => String(row['السنة'])).filter(Boolean))]
        .sort((a, b) => Number(a) - Number(b))
        .map(value => ({ value, label: fmtYear(value) }));
    const departments = [...new Set(studentDetailData.map(row => row['القسم']).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'ar'))
        .map(value => ({ value, label: value }));
    const degrees = [...new Set(studentDetailData.map(row => row['الدرجة']).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'ar'))
        .map(value => ({ value, label: value }));
    const statuses = [...new Set(studentDetailData.map(row => row['الحالة']).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'ar'))
        .map(value => ({ value, label: value }));
    const genders = [...new Set(studentDetailData.map(row => row['الجنس']).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'ar'))
        .map(value => ({ value, label: value }));
    const branches = getOrderedBranchNames(studentDetailData.map(row => row['الفرع']))
        .map(value => ({ value, label: value }));

    setSelectOptions(yearSelect, years, 'كل السنوات');
    setSelectOptions(deptSelect, departments, 'كل الأقسام');
    setSelectOptions(degreeSelect, degrees, 'كل الدرجات');
    setSelectOptions(branchSelect, branches, 'كل الفروع');
    setSelectOptions(statusSelect, statuses, 'كل الحالات');
    setSelectOptions(genderSelect, genders, 'الكل');
    refreshStudentCategoryProgramOptions();
}

function resetStudentCategoriesView() {
    const yearSelect = document.getElementById('student-cat-year');
    const deptSelect = document.getElementById('student-cat-dept');
    const degreeSelect = document.getElementById('student-cat-degree');
    const programSelect = document.getElementById('student-cat-prog');
    const branchSelect = document.getElementById('student-cat-branch');
    const statusSelect = document.getElementById('student-cat-status');
    const genderSelect = document.getElementById('student-cat-gender');
    const searchInput = document.getElementById('student-cat-search');

    const latestYear = Math.max(...studentDetailData.map(row => Number(row['السنة']) || 0));
    if (yearSelect) yearSelect.value = latestYear ? String(latestYear) : '';
    if (deptSelect) deptSelect.value = '';
    if (degreeSelect) degreeSelect.value = '';
    refreshStudentCategoryProgramOptions();
    if (programSelect) programSelect.value = '';
    if (branchSelect) branchSelect.value = '';
    if (statusSelect) statusSelect.value = '';
    if (genderSelect) genderSelect.value = '';
    if (searchInput) searchInput.value = '';

    document.querySelectorAll('#student-category-picks input[type="checkbox"]').forEach(input => {
        input.checked = true;
        input.closest('.student-category-chip')?.classList.add('is-checked');
    });
}

function getStudentCategoryFilteredRows() {
    const year = document.getElementById('student-cat-year')?.value || '';
    const dept = document.getElementById('student-cat-dept')?.value || '';
    const degree = document.getElementById('student-cat-degree')?.value || '';
    const program = document.getElementById('student-cat-prog')?.value || '';
    const branch = document.getElementById('student-cat-branch')?.value || '';
    const status = document.getElementById('student-cat-status')?.value || '';
    const gender = document.getElementById('student-cat-gender')?.value || '';
    const search = normalizeArabicText(document.getElementById('student-cat-search')?.value || '').toLowerCase();

    let rows = studentDetailData;
    if (year) rows = rows.filter(row => String(row['السنة']) === year);
    if (dept) rows = rows.filter(row => row['القسم'] === dept);
    if (degree) rows = rows.filter(row => row['الدرجة'] === degree);
    if (program) rows = rows.filter(row => row['التخصص'] === program);
    if (branch) rows = rows.filter(row => normalizeBranchName(row['الفرع']) === branch);
    if (status) rows = rows.filter(row => row['الحالة'] === status);
    if (gender) rows = rows.filter(row => row['الجنس'] === gender);
    if (search) rows = rows.filter(row => row._searchText.includes(search));
    return rows;
}

function buildStudentCategoryCaption(totalRows) {
    const labels = [];
    const year = document.getElementById('student-cat-year')?.value || '';
    const dept = document.getElementById('student-cat-dept')?.value || '';
    const degree = document.getElementById('student-cat-degree')?.value || '';
    const program = document.getElementById('student-cat-prog')?.value || '';
    const branch = document.getElementById('student-cat-branch')?.value || '';
    const status = document.getElementById('student-cat-status')?.value || '';
    const gender = document.getElementById('student-cat-gender')?.value || '';
    const search = String(document.getElementById('student-cat-search')?.value || '').trim();

    if (year) labels.push(`سنة ${fmtYear(year)}`);
    if (dept) labels.push(`القسم: ${dept}`);
    if (degree) labels.push(`الدرجة: ${degree}`);
    if (program) labels.push(`البرنامج: ${program}`);
    if (branch) labels.push(`الفرع: ${branch} (السجلات الممكن إسنادها فقط)`);
    if (status) labels.push(`الحالة: ${status}`);
    if (gender) labels.push(`الجنس: ${gender}`);
    if (search) labels.push(`البحث: ${search}`);

    const contextText = labels.length ? labels.join(' | ') : 'كل السجلات المتاحة';
    return `تحليل ${fmtNum(totalRows)} سجلًا بعد الفلاتر. النسبة = عدد الفئة ÷ إجمالي السجلات بعد الفلاتر. ${contextText}`;
}

function buildStudentCategoryReport() {
    const selectedCategoryIds = getSelectedStudentCategoryIds();
    if (!selectedCategoryIds.length) {
        alert('حدد فئة واحدة على الأقل لبناء التقرير.');
        return null;
    }

    const filteredRows = getStudentCategoryFilteredRows();
    const selectedDefs = STUDENT_CATEGORY_DEFS.filter(category => selectedCategoryIds.includes(category.id));
    const totalRows = filteredRows.length;
    const averageGpa = analyticsAverage(filteredRows, row => row._gpa);

    const summaryRows = selectedDefs.map(category => {
        const matches = filteredRows
            .filter(row => row._categoryId === category.id)
            .sort((a, b) => (b._gpa || -1) - (a._gpa || -1));
        const count = matches.length;
        const percentage = totalRows > 0 ? Math.round((count / totalRows) * 1000) / 10 : 0;
        const average = analyticsAverage(matches, row => row._gpa);
        return {
            id: category.id,
            label: category.label,
            rangeLabel: category.rangeLabel,
            accent: category.accent,
            count,
            percentage,
            average,
            highest: matches.length ? matches[0]._gpa : null,
            lowest: matches.length ? matches[matches.length - 1]._gpa : null,
            matches
        };
    });

    const detailRowsSource = summaryRows
        .flatMap(summary => summary.matches.map(row => ({ ...row, _reportCategory: summary.label })))
        .sort((a, b) => {
            const aIdx = selectedCategoryIds.indexOf(a._categoryId);
            const bIdx = selectedCategoryIds.indexOf(b._categoryId);
            if (aIdx !== bIdx) return aIdx - bIdx;
            if ((b._gpa || -1) !== (a._gpa || -1)) return (b._gpa || -1) - (a._gpa || -1);
            return String(a['الاسم'] || '').localeCompare(String(b['الاسم'] || ''), 'ar');
        });
    const matchedRowsCount = detailRowsSource.length;
    const outsideCount = Math.max(0, totalRows - matchedRowsCount);

    const summaryHeaders = ['الفئة', 'النطاق المعتمد', 'العدد', 'النسبة', 'متوسط المعدل', 'أعلى معدل', 'أقل معدل'];
    const summaryTableRows = summaryRows.map(summary => [
        summary.label,
        summary.rangeLabel,
        fmtNum(summary.count),
        `${fmtNumFlex(summary.percentage, 1)}%`,
        summary.average == null ? '—' : fmtNumFlex(summary.average, 2),
        summary.highest == null ? '—' : fmtNumFlex(summary.highest, 2),
        summary.lowest == null ? '—' : fmtNumFlex(summary.lowest, 2)
    ]);

    const detailHeaders = ['السنة', 'الرقم الجامعي', 'الاسم', 'البرنامج', 'الدرجة', 'القسم', 'الفرع', 'الحالة', 'الجنس', 'المعدل', 'التقدير الرسمي', 'الفئة', 'مصدر الملف'];
    const detailRows = detailRowsSource.map(row => [
        fmtYear(row['السنة']),
        row['الرقم_الجامعي'],
        row['الاسم'],
        row['التخصص'],
        row['الدرجة'],
        row['القسم'],
        row['الفرع'] || '',
        row['الحالة'],
        row['الجنس'],
        row['المعدل'],
        row._officialEstimate,
        row._reportCategory,
        row['مصدر_الملف']
    ]);

    const filenameParts = [
        'فئات-الطلاب',
        document.getElementById('student-cat-year')?.value ? fmtYear(document.getElementById('student-cat-year').value) : 'كل-السنوات',
        document.getElementById('student-cat-prog')?.value || document.getElementById('student-cat-dept')?.value || 'كل-البرامج'
    ];
    if (document.getElementById('student-cat-branch')?.value) {
        filenameParts.push(document.getElementById('student-cat-branch').value);
    }

    return {
        totalRows,
        averageGpa,
        matchedRowsCount,
        outsideCount,
        caption: `${buildStudentCategoryCaption(totalRows)} الفئات المختارة تغطي ${fmtNum(matchedRowsCount)} سجلًا وتترك ${fmtNum(outsideCount)} سجلًا خارج التصنيف التشغيلي.`,
        summaries: summaryRows,
        detailRowsSource,
        summaryHeaders,
        summaryTableRows,
        detailHeaders,
        detailRows,
        filenameBase: safeFileName(filenameParts.join('-'))
    };
}

function renderStudentCategorySummary(report) {
    const wrap = document.getElementById('student-cat-summary');
    if (!wrap) return;

    const baseCards = [
        {
            icon: '🧾',
            label: 'إجمالي السجلات بعد الفلاتر',
            value: fmtNum(report.totalRows),
            note: 'الأساس الذي تحسب منه النسب'
        },
        {
            icon: '📊',
            label: 'متوسط المعدل',
            value: report.averageGpa == null ? '—' : fmtNumFlex(report.averageGpa, 2),
            note: 'على السلم الرباعي الرسمي'
        },
        {
            icon: '🧭',
            label: 'خارج الفئات المختارة',
            value: fmtNum(report.outsideCount),
            note: `${report.totalRows > 0 ? fmtNumFlex((report.outsideCount / report.totalRows) * 100, 1) : '0.0'}% من السجلات`
        }
    ];

    const categoryCards = report.summaries.map(summary => ({
        icon: '🎯',
        label: summary.label,
        value: fmtNum(summary.count),
        note: `${fmtNumFlex(summary.percentage, 1)}% من السجلات`
    }));

    wrap.innerHTML = [...baseCards, ...categoryCards].map(card => `
        <div class="summary-card student-category-card">
            <div class="sc-icon">${card.icon}</div>
            <div class="sc-value">${card.value}</div>
            <div class="sc-label">${card.label}</div>
            <div class="student-category-card-note">${card.note}</div>
        </div>
    `).join('');
}

function renderStudentCategorySummaryTable(report) {
    const thead = document.getElementById('student-cat-summary-thead');
    const tbody = document.getElementById('student-cat-summary-tbody');
    if (!thead || !tbody) return;

    thead.innerHTML = `<tr>${report.summaryHeaders.map(header => `<th>${header}</th>`).join('')}</tr>`;
    tbody.innerHTML = report.summaryTableRows.map((row, index) => `
        <tr class="${index % 2 ? 'alt' : ''}">
            ${row.map(cell => `<td>${cell}</td>`).join('')}
        </tr>
    `).join('');
}

function renderStudentCategoryDetailTable(report) {
    const tbody = document.getElementById('student-cat-detail-tbody');
    if (!tbody) return;

    const MAX_SHOW = 500;
    const showing = report.detailRowsSource.slice(0, MAX_SHOW);

    tbody.innerHTML = showing.map((row, index) => `
        <tr class="${index % 2 ? 'alt' : ''}">
            <td>${index + 1}</td>
            <td>${escapeHTML(fmtYear(row['السنة']))}</td>
            <td>${escapeHTML(row['الرقم_الجامعي'])}</td>
            <td>${escapeHTML(row['الاسم'])}</td>
            <td>${escapeHTML(row['التخصص'])}</td>
            <td>${escapeHTML(row['الدرجة'])}</td>
            <td>${escapeHTML(row['القسم'])}</td>
            <td>${escapeHTML(row['الفرع'] || '—')}</td>
            <td>${escapeHTML(row['الحالة'])}</td>
            <td>${escapeHTML(row['الجنس'])}</td>
            <td>${escapeHTML(row['المعدل'])}</td>
            <td>${escapeHTML(row._officialEstimate)}</td>
            <td>${escapeHTML(row._reportCategory)}</td>
        </tr>
    `).join('');

    if (report.detailRowsSource.length > MAX_SHOW) {
        tbody.innerHTML += `<tr><td colspan="13" style="text-align:center;color:var(--text-light);padding:16px">
            يتم عرض أول ${MAX_SHOW} سجل فقط داخل الصفحة. التصدير يشمل جميع السجلات المطابقة.
        </td></tr>`;
    }
}

function renderStudentCategoryChart(report) {
    if (studentCategoryChart) {
        studentCategoryChart.destroy();
        studentCategoryChart = null;
    }

    const canvas = document.getElementById('student-cat-chart');
    const empty = document.getElementById('student-cat-chart-empty');
    if (!canvas || !empty) return;

    if (!report.summaries.length || report.summaries.every(summary => summary.count === 0)) {
        canvas.classList.add('hidden');
        empty.classList.remove('hidden');
        return;
    }

    canvas.classList.remove('hidden');
    empty.classList.add('hidden');

    studentCategoryChart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels: report.summaries.map(summary => summary.label),
            datasets: [{
                label: 'عدد الطلاب',
                data: report.summaries.map(summary => summary.count),
                backgroundColor: report.summaries.map(summary => summary.accent),
                borderRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: context => {
                            const summary = report.summaries[context.dataIndex];
                            return `${fmtNum(summary.count)} طالب (${fmtNumFlex(summary.percentage, 1)}%)`;
                        }
                    }
                }
            },
            scales: {
                x: { ticks: { font: { family: 'Tajawal' } } },
                y: {
                    beginAtZero: true,
                    ticks: { font: { family: 'Tajawal' }, precision: 0 },
                    grid: { color: '#eef2f7' }
                }
            }
        }
    });
}

function renderStudentCategoriesReport(scrollToResults = true) {
    const report = buildStudentCategoryReport();
    if (!report) return;

    currentStudentCategoryReport = report;
    document.getElementById('student-cat-caption').textContent = report.caption;

    renderStudentCategorySummary(report);
    renderStudentCategorySummaryTable(report);
    renderStudentCategoryDetailTable(report);
    renderStudentCategoryChart(report);

    document.getElementById('student-cat-results').classList.remove('hidden');
    if (scrollToResults) {
        document.getElementById('student-cat-results').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function initStudentCategoriesView() {
    if (!studentDetailData.length) return;

    renderStudentCategoryChoices();
    refreshStudentCategoryBaseFilters();
    resetStudentCategoriesView();

    ['student-cat-year', 'student-cat-dept', 'student-cat-degree'].forEach(id => {
        document.getElementById(id)?.addEventListener('change', () => {
            refreshStudentCategoryProgramOptions();
            document.getElementById('student-cat-results')?.classList.add('hidden');
        });
    });

    ['student-cat-prog', 'student-cat-branch', 'student-cat-status', 'student-cat-gender'].forEach(id => {
        document.getElementById(id)?.addEventListener('change', () => {
            document.getElementById('student-cat-results')?.classList.add('hidden');
        });
    });

    document.getElementById('student-cat-search')?.addEventListener('input', () => {
        document.getElementById('student-cat-results')?.classList.add('hidden');
    });

    document.getElementById('student-cat-select-all')?.addEventListener('click', () => {
        document.querySelectorAll('#student-category-picks input[type="checkbox"]').forEach(input => {
            input.checked = true;
            input.closest('.student-category-chip')?.classList.add('is-checked');
        });
        document.getElementById('student-cat-results')?.classList.add('hidden');
    });

    document.getElementById('student-cat-clear-all')?.addEventListener('click', () => {
        document.querySelectorAll('#student-category-picks input[type="checkbox"]').forEach(input => {
            input.checked = false;
            input.closest('.student-category-chip')?.classList.remove('is-checked');
        });
        document.getElementById('student-cat-results')?.classList.add('hidden');
    });

    document.getElementById('student-cat-run')?.addEventListener('click', () => renderStudentCategoriesReport(true));
    document.getElementById('student-cat-reset')?.addEventListener('click', () => {
        resetStudentCategoriesView();
        renderStudentCategoriesReport(true);
    });

    renderStudentCategoriesReport(false);
}

async function exportStudentCategoriesPDF() {
    if (!currentStudentCategoryReport) return;
    await exportSectionAsPDF('student-cat-export-area', `${currentStudentCategoryReport.filenameBase}.pdf`);
}

function exportStudentCategoriesExcel() {
    if (!currentStudentCategoryReport) return;

    const summarySheet = XLSX.utils.aoa_to_sheet([
        ['فئات الطلاب'],
        ['الوصف', currentStudentCategoryReport.caption],
        [],
        currentStudentCategoryReport.summaryHeaders,
        ...currentStudentCategoryReport.summaryTableRows
    ]);
    const detailSheet = XLSX.utils.aoa_to_sheet([
        currentStudentCategoryReport.detailHeaders,
        ...currentStudentCategoryReport.detailRows
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, summarySheet, 'ملخص الفئات');
    XLSX.utils.book_append_sheet(wb, detailSheet, 'التفاصيل');
    XLSX.writeFile(wb, `${currentStudentCategoryReport.filenameBase}.xlsx`);
}

function exportStudentCategoriesCSV() {
    if (!currentStudentCategoryReport) return;
    downloadCSV(`${currentStudentCategoryReport.filenameBase}.csv`, [
        ['فئات الطلاب'],
        ['الوصف', currentStudentCategoryReport.caption],
        [],
        currentStudentCategoryReport.summaryHeaders,
        ...currentStudentCategoryReport.summaryTableRows,
        [],
        currentStudentCategoryReport.detailHeaders,
        ...currentStudentCategoryReport.detailRows
    ]);
}

// ========================================
// الاستوديو الإحصائي
// ========================================
function analyticsSum(records, getter) {
    let sum = 0;
    let count = 0;
    records.forEach(row => {
        const rawValue = getter(row);
        if (rawValue == null || rawValue === '') return;
        const value = Number(rawValue);
        if (!Number.isFinite(value)) return;
        sum += value;
        count++;
    });
    return count ? sum : null;
}

function analyticsAverage(records, getter) {
    let sum = 0;
    let count = 0;
    records.forEach(row => {
        const rawValue = getter(row);
        if (rawValue == null || rawValue === '') return;
        const value = Number(rawValue);
        if (!Number.isFinite(value)) return;
        sum += value;
        count++;
    });
    return count > 0 ? (sum / count) : null;
}

function analyticsWeightedAverage(records, valueGetter, weightGetter, fallbackWeight = 1) {
    let weightedSum = 0;
    let totalWeight = 0;
    records.forEach(row => {
        const rawValue = valueGetter(row);
        if (rawValue == null || rawValue === '') return;
        const value = Number(rawValue);
        if (!Number.isFinite(value)) return;
        const weightValue = weightGetter ? weightGetter(row) : fallbackWeight;
        const rawWeight = weightValue == null || weightValue === '' ? NaN : Number(weightValue);
        const weight = Number.isFinite(rawWeight) && rawWeight > 0 ? rawWeight : fallbackWeight;
        weightedSum += value * weight;
        totalWeight += weight;
    });
    return totalWeight > 0 ? (weightedSum / totalWeight) : null;
}

function analyticsCountWhere(records, predicate) {
    return records.reduce((count, row) => count + (predicate(row) ? 1 : 0), 0);
}

function analyticsDistinctCount(records, getter) {
    const values = new Set();
    records.forEach(row => {
        const value = String(getter(row) ?? '').trim();
        if (value) values.add(value);
    });
    return values.size;
}

function analyticsUniqueRows(records, keyGetter) {
    const map = new Map();
    records.forEach(row => {
        const key = String(keyGetter(row) ?? '').trim();
        if (!key || map.has(key)) return;
        map.set(key, row);
    });
    return [...map.values()];
}

function analyticsProgramResearchRows(records) {
    return analyticsUniqueRows(records, row => {
        const year = parseInt(row.Semester, 10);
        const dept = normalizeDepartment(row.Dept_aName);
        return year && dept ? `${year}|${dept}` : '';
    });
}

function analyticsFormatCount(value) {
    return value == null ? '—' : fmtNum(Math.round(value));
}

function analyticsFormatDecimal(value, frac = 2) {
    return value == null ? '—' : fmtNumFlex(Number(value), frac);
}

function analyticsFormatPercent(value) {
    return value == null ? '—' : `${Number(value).toFixed(1)}%`;
}

function analyticsFormatRatio(value) {
    return value == null ? '—' : `1:${Number(value).toFixed(1)}`;
}

function analyticsMetricHeader(metric) {
    return metric.unit ? `${metric.label} (${metric.unit})` : metric.label;
}

function markAnalyticsSampleMetrics(metricDefs, records) {
    return metricDefs.map(metric => {
        if (!metric.sampleSourceKey) return metric;
        const usesGraduateSample = records.some(row =>
            String(row[metric.sampleSourceKey] || '').startsWith('graduates_sample_')
        );
        if (!usesGraduateSample) return metric;
        return { ...metric, label: `${metric.label} - استطلاع عينة` };
    });
}

function analyticsText(value, fallback = 'غير محدد') {
    const text = String(value ?? '').trim();
    return text || fallback;
}

function analyticsQueryMatch(text, query) {
    return normalizeArabicText(text).toLowerCase().includes(normalizeArabicText(query).toLowerCase());
}

function analyticsParseGPA(value) {
    const num = parseFloat(String(value ?? '').trim().replace(',', '.'));
    return Number.isFinite(num) ? num : null;
}

function analyticsDurationYears(startValue, endValue) {
    const start = parseDaySerial(startValue);
    const end = parseDaySerial(endValue);
    if (start == null || end == null || end < start) return null;
    const years = (end - start) / 365.25;
    if (!Number.isFinite(years) || years < 0.2 || years > 15) return null;
    return years;
}

function analyticsCountByStatus(records, keywords) {
    const normalizedKeywords = keywords.map(k => normalizeArabicText(k));
    return analyticsCountWhere(records, row => {
        const status = normalizeArabicText(row['الحالة']);
        return normalizedKeywords.some(keyword => status.includes(keyword));
    });
}

function getAnalyticsSelectedBranchValue() {
    const rawValue = document.getElementById('analytics-filter-branch')?.value || '';
    return normalizeBranchName(rawValue);
}

function getAnalyticsSourceDefinitions() {
    return {
        programs: {
            key: 'programs',
            label: 'بيانات المؤشرات والبرامج',
            rowLabel: 'سجل برنامج/سنة',
            getRows: () => allRows,
            getYear: row => parseInt(row.Semester, 10) || null,
            searchText: row => [
                row.Major_aName,
                row.Dept_aName,
                row.Degree_aName,
                fmtYear(row.Semester)
            ].join(' '),
            filters: [
                { id: 'dept', label: 'القسم', getValue: row => row.Dept_aName },
                { id: 'degree', label: 'الدرجة', getValue: row => row.Degree_aName },
                { id: 'program', label: 'البرنامج', getValue: row => row.Major_aName },
                { id: 'branch', label: 'الفرع', getValue: () => '', options: () => getOrderedBranchNames(availableFacultyBranches) },
            ],
            groups: [
                { id: 'year', label: 'السنة', getValue: row => parseInt(row.Semester, 10) || null, format: value => fmtYear(value), sort: 'numeric' },
                { id: 'dept', label: 'القسم', getValue: row => row.Dept_aName, format: value => analyticsText(value), sort: 'text' },
                { id: 'degree', label: 'الدرجة', getValue: row => row.Degree_aName, format: value => analyticsText(value), sort: 'text' },
                { id: 'program', label: 'البرنامج', getValue: row => row.Major_aName, format: value => analyticsText(value), sort: 'text' },
            ],
            defaultMetrics: ['record_count', 'students_total', 'graduates_total', 'graduation_rate', 'retention_rate'],
            metrics: [
                { id: 'record_count', label: 'عدد السجلات', unit: 'سجل', compute: records => records.length, format: analyticsFormatCount },
                { id: 'distinct_programs', label: 'عدد البرامج المختلفة', unit: 'برنامج', compute: records => analyticsDistinctCount(records, row => `${row.Major_aName}|${row.Degree_aName}`), format: analyticsFormatCount },
                { id: 'students_total', label: 'إجمالي الطلاب', unit: 'طالب', compute: records => analyticsSum(records, row => row.students_total), format: analyticsFormatCount },
                { id: 'students_male', label: 'الطلاب الذكور', unit: 'طالب', compute: records => analyticsSum(records, row => row.students_male), format: analyticsFormatCount },
                { id: 'students_female', label: 'الطالبات', unit: 'طالبة', compute: records => analyticsSum(records, row => row.students_female), format: analyticsFormatCount },
                { id: 'students_new', label: 'الطلاب المستجدون', unit: 'طالب', compute: records => analyticsSum(records, row => row.students_new), format: analyticsFormatCount },
                { id: 'graduates_total', label: 'إجمالي الخريجين', unit: 'خريج', compute: records => analyticsSum(records, row => row.graduates_total), format: analyticsFormatCount },
                { id: 'graduates_branch_matched', label: 'خريجون أمكن إسنادهم للفرع', unit: 'خريج', compute: records => analyticsSum(records, row => row.graduates_branch_matched), format: analyticsFormatCount },
                { id: 'graduates_ontime', label: 'الخريجون بالوقت المحدد', unit: 'خريج', compute: records => analyticsSum(records, row => row.graduates_ontime), format: analyticsFormatCount },
                { id: 'sections_total', label: 'إجمالي الشعب', unit: 'شعبة', compute: records => analyticsSum(records, row => row.sections_total), format: analyticsFormatCount },
                { id: 'faculty_fte_total', label: 'إجمالي هيئة التدريس (FTE)', unit: 'عضو', compute: records => analyticsSum(records, row => getFacultyBaseForRatio(row)), format: value => analyticsFormatDecimal(value, 2) },
                {
                    id: 'student_faculty_ratio',
                    label: 'نسبة الطلاب إلى هيئة التدريس',
                    unit: 'نسبة',
                    compute: records => {
                        const totalStudents = analyticsSum(records, row => row.students_total);
                        const totalFacultyBase = analyticsSum(records, row => getFacultyBaseForRatio(row));
                        return totalFacultyBase > 0 ? (totalStudents / totalFacultyBase) : null;
                    },
                    format: analyticsFormatRatio
                },
                {
                    id: 'course_eval',
                    label: 'متوسط تقييم جودة المقررات',
                    unit: 'درجة',
                    compute: records => analyticsWeightedAverage(records, row => row.eval_courses, row => row.eval_courses_sample, 1),
                    format: value => analyticsFormatDecimal(value, 2)
                },
                {
                    id: 'experience_eval',
                    label: 'متوسط جودة خبرات التعلم',
                    unit: 'درجة',
                    compute: records => analyticsWeightedAverage(records, row => row.eval_experience, row => row.eval_experience_sample, 1),
                    format: value => analyticsFormatDecimal(value, 2)
                },
                {
                    id: 'supervision_eval',
                    label: 'متوسط جودة الإشراف العلمي',
                    unit: 'درجة',
                    sampleSourceKey: 'eval_supervision_source',
                    compute: records => analyticsWeightedAverage(records, row => row.eval_supervision, row => row.eval_supervision_sample, 1),
                    format: value => analyticsFormatDecimal(value, 2)
                },
                {
                    id: 'services_eval',
                    label: 'متوسط رضا الطلاب عن الخدمات',
                    unit: 'درجة',
                    sampleSourceKey: 'eval_services_source',
                    compute: records => analyticsWeightedAverage(records, row => row.eval_services, row => row.eval_services_sample, 1),
                    format: value => analyticsFormatDecimal(value, 2)
                },
                {
                    id: 'student_performance',
                    label: 'متوسط مستوى أداء الطالب',
                    unit: '%',
                    sampleSourceKey: 'performance_rate_source',
                    compute: records => analyticsWeightedAverage(records, row => row.performance_rate, row => row.performance_rate_sample, 1),
                    format: analyticsFormatPercent
                },
                {
                    id: 'employment_rate',
                    label: 'توظيف الخريجين أو التحاقهم بالدراسات العليا',
                    unit: '%',
                    sampleSourceKey: 'employment_rate_source',
                    compute: records => analyticsWeightedAverage(records, row => row.employment_rate, row => row.employment_rate_sample, 1),
                    format: analyticsFormatPercent
                },
                {
                    id: 'employer_eval',
                    label: 'متوسط تقويم جهات التوظيف',
                    unit: 'درجة',
                    sampleSourceKey: 'eval_employers_source',
                    compute: records => analyticsWeightedAverage(records, row => row.eval_employers, row => row.eval_employers_sample, 1),
                    format: value => analyticsFormatDecimal(value, 2)
                },
                {
                    id: 'graduation_rate',
                    label: 'معدل التخرج بالوقت المحدد',
                    unit: '%',
                    compute: records => pct(
                        analyticsSum(records, row => row.graduates_ontime),
                        analyticsSum(records, row => row.new_4_ago_count)
                    ),
                    format: analyticsFormatPercent
                },
                {
                    id: 'retention_rate',
                    label: 'معدل الاستبقاء',
                    unit: '%',
                    compute: records => pct(
                        analyticsSum(records, row => row.students_retained),
                        analyticsSum(records, row => row.prev_new_count)
                    ),
                    format: analyticsFormatPercent
                },
                {
                    id: 'dropout_rate',
                    label: 'معدل التسرب',
                    unit: '%',
                    compute: records => {
                        const retention = pct(
                            analyticsSum(records, row => row.students_retained),
                            analyticsSum(records, row => row.prev_new_count)
                        );
                        return retention == null ? null : Math.round((100 - retention) * 10) / 10;
                    },
                    format: analyticsFormatPercent
                },
                {
                    id: 'publication_pct',
                    label: 'النسبة المئوية للنشر العلمي',
                    unit: '%',
                    compute: records => {
                        const baseRows = analyticsProgramResearchRows(records);
                        return pct(
                            analyticsSum(baseRows, row => row.faculty_published),
                            analyticsSum(baseRows, row => row.faculty_total)
                        );
                    },
                    format: analyticsFormatPercent
                },
                {
                    id: 'research_count',
                    label: 'إجمالي البحوث المنشورة',
                    unit: 'بحث',
                    compute: records => analyticsSum(analyticsProgramResearchRows(records), row => row.research_count),
                    format: analyticsFormatCount
                },
                {
                    id: 'research_per_faculty',
                    label: 'معدل البحوث لكل عضو هيئة تدريس',
                    unit: 'بحث',
                    compute: records => {
                        const baseRows = analyticsProgramResearchRows(records);
                        const facultyTotal = analyticsSum(baseRows, row => row.faculty_total);
                        const researchTotal = analyticsSum(baseRows, row => row.research_count);
                        return facultyTotal > 0 ? (researchTotal / facultyTotal) : null;
                    },
                    format: value => analyticsFormatDecimal(value, 2)
                },
                {
                    id: 'citations_total',
                    label: 'إجمالي الاقتباسات',
                    unit: 'اقتباس',
                    compute: records => analyticsSum(analyticsProgramResearchRows(records), row => row.citations),
                    format: value => analyticsFormatDecimal(value, 1)
                },
                {
                    id: 'citations_per_publication',
                    label: 'معدل الاقتباسات لكل بحث',
                    unit: 'اقتباس',
                    compute: records => {
                        const baseRows = analyticsProgramResearchRows(records);
                        const researchTotal = analyticsSum(baseRows, row => row.research_count);
                        const citationsTotal = analyticsSum(baseRows, row => row.citations);
                        return researchTotal > 0 ? (citationsTotal / researchTotal) : null;
                    },
                    format: value => analyticsFormatDecimal(value, 2)
                },
            ],
        },
        faculty: {
            key: 'faculty',
            label: 'قائمة أعضاء هيئة التدريس',
            rowLabel: 'سجل عضو هيئة تدريس',
            getRows: () => facultyRosterData,
            getYear: row => row.year,
            searchText: row => [
                row.name, row.id, row.rank, row.department, row.branch,
                row.gender, row.nationality, row.email
            ].join(' '),
            filters: [
                { id: 'dept', label: 'القسم', getValue: row => row.department },
                { id: 'rank', label: 'الرتبة', getValue: row => row.rank },
                {
                    id: 'branch',
                    label: 'الفرع',
                    getValue: row => getFacultyBranchNames(row.branch).join('، '),
                    getOptionsValues: row => getFacultyBranchNames(row.branch),
                    matches: (row, wanted) => getFacultyBranchNames(row.branch).includes(wanted)
                },
                { id: 'gender', label: 'الجنس', getValue: row => row.gender },
                { id: 'nationality', label: 'الجنسية', getValue: row => row.nationality },
                { id: 'active', label: 'الحالة', getValue: row => row.active === 'نعم' ? 'نشط' : 'غير نشط' },
            ],
            groups: [
                { id: 'year', label: 'السنة', getValue: row => row.year, format: value => fmtYear(value), sort: 'numeric' },
                { id: 'dept', label: 'القسم', getValue: row => row.department, format: value => analyticsText(value), sort: 'text' },
                { id: 'rank', label: 'الرتبة', getValue: row => row.rank, format: value => analyticsText(value), sort: 'text' },
                {
                    id: 'branch',
                    label: 'الفرع',
                    getValue: row => getFacultyBranchNames(row.branch).join('، '),
                    getValues: row => getFacultyBranchNames(row.branch),
                    format: value => analyticsText(value),
                    sort: 'text'
                },
                { id: 'gender', label: 'الجنس', getValue: row => row.gender, format: value => analyticsText(value), sort: 'text' },
                { id: 'nationality', label: 'الجنسية', getValue: row => row.nationality, format: value => analyticsText(value), sort: 'text' },
            ],
            defaultMetrics: ['record_count', 'active_count'],
            metrics: [
                { id: 'record_count', label: 'عدد السجلات', unit: 'عضو', compute: records => records.length, format: analyticsFormatCount },
                { id: 'active_count', label: 'الأعضاء النشطون', unit: 'عضو', compute: records => analyticsCountWhere(records, row => row.active === 'نعم'), format: analyticsFormatCount },
                { id: 'inactive_count', label: 'الأعضاء غير النشطين', unit: 'عضو', compute: records => analyticsCountWhere(records, row => row.active !== 'نعم'), format: analyticsFormatCount },
                { id: 'distinct_members', label: 'الأعضاء المختلفون', unit: 'عضو', compute: records => analyticsDistinctCount(records, row => row.id), format: analyticsFormatCount },
            ],
        },
        entrants: {
            key: 'entrants',
            label: 'سجل المستجدين في الفروع',
            rowLabel: 'سجل مستجد',
            getRows: () => entrantData,
            getYear: row => parseInt(row.year, 10) || null,
            searchText: row => [
                row.name,
                row.student_id,
                row.program,
                row.degree,
                row.department,
                row.branch,
                row.gender
            ].join(' '),
            filters: [
                { id: 'branch', label: 'الفرع', getValue: row => normalizeBranchName(row.branch), options: rows => getOrderedBranchNames(rows.map(row => row.branch)) },
                { id: 'gender', label: 'الجنس', getValue: row => row.gender },
            ],
            groups: [
                { id: 'year', label: 'السنة', getValue: row => parseInt(row.year, 10) || null, format: value => fmtYear(value), sort: 'numeric' },
                { id: 'semester', label: 'الفصل', getValue: row => parseInt(row.semester, 10) || null, format: value => formatAcademicSemester(value), sort: 'numeric' },
                { id: 'branch', label: 'الفرع', getValue: row => normalizeBranchName(row.branch), format: value => analyticsText(value), sort: 'text' },
                { id: 'gender', label: 'الجنس', getValue: row => row.gender, format: value => analyticsText(value), sort: 'text' },
            ],
            defaultMetrics: ['record_count', 'male_count', 'female_count'],
            metrics: [
                { id: 'record_count', label: 'عدد المستجدين', unit: 'طالب', compute: records => records.length, format: analyticsFormatCount },
                { id: 'male_count', label: 'عدد الذكور', unit: 'طالب', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row.gender).includes('ذكر')), format: analyticsFormatCount },
                { id: 'female_count', label: 'عدد الإناث', unit: 'طالبة', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row.gender).includes('أنث')), format: analyticsFormatCount },
            ],
        },
        graduates: {
            key: 'graduates',
            label: 'سجل الخريجين',
            rowLabel: 'سجل خريج',
            getRows: () => gradData,
            getYear: row => parseAcademicYear(row['السنة']),
            searchText: row => [
                row['الاسم'],
                row['الرقم_الجامعي'],
                row['التخصص'],
                row['الدرجة'],
                row['القسم'],
                row['الفرع']
            ].join(' '),
            filters: [
                { id: 'dept', label: 'القسم', getValue: row => row['القسم'] },
                { id: 'degree', label: 'الدرجة', getValue: row => row['الدرجة'] },
                { id: 'program', label: 'البرنامج', getValue: row => row['التخصص'] },
                { id: 'branch', label: 'الفرع', getValue: row => normalizeBranchName(row['الفرع']), options: rows => getOrderedBranchNames(rows.map(row => row['الفرع'])) },
                { id: 'gender', label: 'الجنس', getValue: row => row['الجنس'] },
                { id: 'nationality', label: 'الجنسية', getValue: row => row['الجنسية'] },
            ],
            groups: [
                { id: 'year', label: 'السنة', getValue: row => parseAcademicYear(row['السنة']), format: value => fmtYear(value), sort: 'numeric' },
                { id: 'dept', label: 'القسم', getValue: row => row['القسم'], format: value => analyticsText(value), sort: 'text' },
                { id: 'degree', label: 'الدرجة', getValue: row => row['الدرجة'], format: value => analyticsText(value), sort: 'text' },
                { id: 'program', label: 'البرنامج', getValue: row => row['التخصص'], format: value => analyticsText(value), sort: 'text' },
                { id: 'branch', label: 'الفرع', getValue: row => normalizeBranchName(row['الفرع']), format: value => analyticsText(value), sort: 'text' },
                { id: 'gender', label: 'الجنس', getValue: row => row['الجنس'], format: value => analyticsText(value), sort: 'text' },
                { id: 'nationality', label: 'الجنسية', getValue: row => row['الجنسية'], format: value => analyticsText(value), sort: 'text' },
            ],
            defaultMetrics: ['record_count', 'avg_gpa', 'avg_study_duration', 'ontime_expected_rate'],
            metrics: [
                { id: 'record_count', label: 'عدد السجلات', unit: 'خريج', compute: records => records.length, format: analyticsFormatCount },
                { id: 'distinct_programs', label: 'عدد البرامج المختلفة', unit: 'برنامج', compute: records => analyticsDistinctCount(records, row => `${row['التخصص']}|${row['الدرجة']}`), format: analyticsFormatCount },
                { id: 'male_count', label: 'عدد الذكور', unit: 'خريج', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row['الجنس']).includes('ذكر')), format: analyticsFormatCount },
                { id: 'female_count', label: 'عدد الإناث', unit: 'خريجة', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row['الجنس']).includes('أنث')), format: analyticsFormatCount },
                { id: 'saudi_count', label: 'عدد السعوديين', unit: 'خريج', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row['الجنسية']).includes('سعود')), format: analyticsFormatCount },
                { id: 'non_saudi_count', label: 'عدد غير السعوديين', unit: 'خريج', compute: records => analyticsCountWhere(records, row => !normalizeArabicText(row['الجنسية']).includes('سعود')), format: analyticsFormatCount },
                { id: 'avg_gpa', label: 'متوسط المعدل', unit: 'معدل', compute: records => analyticsAverage(records, row => analyticsParseGPA(row['المعدل'])), format: value => analyticsFormatDecimal(value, 2) },
                { id: 'min_gpa', label: 'أقل معدل', unit: 'معدل', compute: records => {
                    const values = records.map(row => analyticsParseGPA(row['المعدل'])).filter(v => v != null);
                    return values.length ? Math.min(...values) : null;
                }, format: value => analyticsFormatDecimal(value, 2) },
                { id: 'max_gpa', label: 'أعلى معدل', unit: 'معدل', compute: records => {
                    const values = records.map(row => analyticsParseGPA(row['المعدل'])).filter(v => v != null);
                    return values.length ? Math.max(...values) : null;
                }, format: value => analyticsFormatDecimal(value, 2) },
                { id: 'avg_study_duration', label: 'متوسط مدة الدراسة حتى التخرج', unit: 'سنة', compute: records => analyticsAverage(records, row =>
                    analyticsDurationYears(row['تاريخ_القبول'], row['تاريخ_التخرج'])
                ), format: value => analyticsFormatDecimal(value, 2) },
                { id: 'ontime_expected_rate', label: 'الالتزام بموعد التخرج المتوقع', unit: '%', compute: records => {
                    let matched = 0;
                    let total = 0;
                    records.forEach(row => {
                        const expected = parseDaySerial(row['تاريخ_التخرج_المتوقع']);
                        const actual = parseDaySerial(row['تاريخ_التخرج']);
                        if (expected == null || actual == null) return;
                        total++;
                        if (actual <= expected) matched++;
                    });
                    return pct(matched, total);
                }, format: analyticsFormatPercent },
            ],
        },
        noncomplete: {
            key: 'noncomplete',
            label: 'سجل غير المكملين',
            rowLabel: 'سجل طالب',
            getRows: () => ncData,
            getYear: row => parseAcademicYear(row['آخر_سنة']),
            searchText: row => [
                row['الاسم'],
                row['الرقم_الجامعي'],
                row['التخصص'],
                row['الدرجة'],
                row['القسم'],
                row['الفرع'],
                row['الحالة']
            ].join(' '),
            filters: [
                { id: 'dept', label: 'القسم', getValue: row => row['القسم'] },
                { id: 'degree', label: 'الدرجة', getValue: row => row['الدرجة'] },
                { id: 'program', label: 'البرنامج', getValue: row => row['التخصص'] },
                { id: 'branch', label: 'الفرع', getValue: row => normalizeBranchName(row['الفرع']), options: rows => getOrderedBranchNames(rows.map(row => row['الفرع'])) },
                { id: 'status', label: 'الحالة', getValue: row => row['الحالة'] },
                { id: 'gender', label: 'الجنس', getValue: row => row['الجنس'] },
                { id: 'nationality', label: 'الجنسية', getValue: row => row['الجنسية'] },
                { id: 'study_type', label: 'نوع الدراسة', getValue: row => row['نوع_الدراسة'] },
            ],
            groups: [
                { id: 'year', label: 'آخر سنة', getValue: row => parseAcademicYear(row['آخر_سنة']), format: value => fmtYear(value), sort: 'numeric' },
                { id: 'dept', label: 'القسم', getValue: row => row['القسم'], format: value => analyticsText(value), sort: 'text' },
                { id: 'degree', label: 'الدرجة', getValue: row => row['الدرجة'], format: value => analyticsText(value), sort: 'text' },
                { id: 'program', label: 'البرنامج', getValue: row => row['التخصص'], format: value => analyticsText(value), sort: 'text' },
                { id: 'branch', label: 'الفرع', getValue: row => normalizeBranchName(row['الفرع']), format: value => analyticsText(value), sort: 'text' },
                { id: 'status', label: 'الحالة', getValue: row => row['الحالة'], format: value => analyticsText(value), sort: 'text' },
                { id: 'gender', label: 'الجنس', getValue: row => row['الجنس'], format: value => analyticsText(value), sort: 'text' },
                { id: 'nationality', label: 'الجنسية', getValue: row => row['الجنسية'], format: value => analyticsText(value), sort: 'text' },
                { id: 'study_type', label: 'نوع الدراسة', getValue: row => row['نوع_الدراسة'], format: value => analyticsText(value), sort: 'text' },
            ],
            defaultMetrics: ['record_count', 'avg_gpa', 'withdrawn_count', 'folded_count'],
            metrics: [
                { id: 'record_count', label: 'عدد السجلات', unit: 'طالب', compute: records => records.length, format: analyticsFormatCount },
                { id: 'distinct_programs', label: 'عدد البرامج المختلفة', unit: 'برنامج', compute: records => analyticsDistinctCount(records, row => `${row['التخصص']}|${row['الدرجة']}`), format: analyticsFormatCount },
                { id: 'male_count', label: 'عدد الذكور', unit: 'طالب', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row['الجنس']).includes('ذكر')), format: analyticsFormatCount },
                { id: 'female_count', label: 'عدد الإناث', unit: 'طالبة', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row['الجنس']).includes('أنث')), format: analyticsFormatCount },
                { id: 'saudi_count', label: 'عدد السعوديين', unit: 'طالب', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row['الجنسية']).includes('سعود')), format: analyticsFormatCount },
                { id: 'non_saudi_count', label: 'عدد غير السعوديين', unit: 'طالب', compute: records => analyticsCountWhere(records, row => !normalizeArabicText(row['الجنسية']).includes('سعود')), format: analyticsFormatCount },
                { id: 'avg_gpa', label: 'متوسط المعدل', unit: 'معدل', compute: records => analyticsAverage(records, row => analyticsParseGPA(row['المعدل'])), format: value => analyticsFormatDecimal(value, 2) },
                { id: 'withdrawn_count', label: 'حالات الانسحاب', unit: 'حالة', compute: records => analyticsCountByStatus(records, ['منسحب']), format: analyticsFormatCount },
                { id: 'folded_count', label: 'مطوي القيد', unit: 'حالة', compute: records => analyticsCountByStatus(records, ['مطوي']), format: analyticsFormatCount },
                { id: 'absent_count', label: 'المنقطعون عن الدراسة', unit: 'حالة', compute: records => analyticsCountByStatus(records, ['منقطع']), format: analyticsFormatCount },
                { id: 'postponed_count', label: 'حالات التأجيل', unit: 'حالة', compute: records => analyticsCountByStatus(records, ['مؤجل']), format: analyticsFormatCount },
                { id: 'dismissed_count', label: 'الحالات الأكاديمية المفصولة', unit: 'حالة', compute: records => analyticsCountByStatus(records, ['مفصول']), format: analyticsFormatCount },
                { id: 'excused_count', label: 'حالات الاعتذار', unit: 'حالة', compute: records => analyticsCountByStatus(records, ['معتذر']), format: analyticsFormatCount },
                { id: 'regular_study_count', label: 'عدد المنتظمين', unit: 'طالب', compute: records => analyticsCountWhere(records, row => normalizeArabicText(row['نوع_الدراسة']).includes('منتظم')), format: analyticsFormatCount },
                { id: 'distinct_statuses', label: 'عدد الحالات المختلفة', unit: 'حالة', compute: records => analyticsDistinctCount(records, row => row['الحالة']), format: analyticsFormatCount },
            ],
        }
    };
}

function getAnalyticsCurrentSource() {
    const sources = getAnalyticsSourceDefinitions();
    const selectedKey = document.getElementById('analytics-source')?.value;
    return sources[selectedKey] || Object.values(sources)[0];
}

function initAnalyticsView() {
    const sourceSelect = document.getElementById('analytics-source');
    const modeSelect = document.getElementById('analytics-mode');
    if (!sourceSelect || !modeSelect) return;
    if (sourceSelect.dataset.ready === 'true') return;

    const sources = Object.values(getAnalyticsSourceDefinitions());
    sourceSelect.innerHTML = sources.map(source =>
        `<option value="${source.key}">${source.label}</option>`
    ).join('');

    sourceSelect.addEventListener('change', () => {
        renderAnalyticsBuilder(true);
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
    });
    modeSelect.addEventListener('change', () => {
        syncAnalyticsGroupingState();
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
    });
    document.getElementById('analytics-group-primary')?.addEventListener('change', () => {
        updateAnalyticsSecondaryGroups();
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
    });
    document.getElementById('analytics-group-secondary')?.addEventListener('change', () => {
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
    });
    document.getElementById('analytics-year-from')?.addEventListener('change', () => {
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
    });
    document.getElementById('analytics-year-to')?.addEventListener('change', () => {
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
    });
    document.getElementById('analytics-search')?.addEventListener('input', () => {
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
    });
    document.getElementById('analytics-run')?.addEventListener('click', runAnalyticsReport);
    document.getElementById('analytics-reset')?.addEventListener('click', resetAnalyticsView);

    sourceSelect.dataset.ready = 'true';
    resetAnalyticsView();
}

function resetAnalyticsView() {
    const sourceSelect = document.getElementById('analytics-source');
    const modeSelect = document.getElementById('analytics-mode');
    const searchInput = document.getElementById('analytics-search');
    if (modeSelect) modeSelect.value = 'summary';
    if (sourceSelect && !sourceSelect.value) {
        const firstOption = sourceSelect.querySelector('option');
        sourceSelect.value = firstOption ? firstOption.value : '';
    }
    if (searchInput) searchInput.value = '';
    renderAnalyticsBuilder(true);
    document.getElementById('analytics-results').classList.add('hidden');
    currentAnalyticsReport = null;
    destroyAnalyticsChart();
}

function renderAnalyticsBuilder(resetSelections = false) {
    const source = getAnalyticsCurrentSource();
    if (!source) return;

    renderAnalyticsYears(source, resetSelections);
    renderAnalyticsGroups(source, resetSelections);
    renderAnalyticsFilters(source);
    renderAnalyticsMetrics(source);
    syncAnalyticsGroupingState();

    const filterCaption = document.getElementById('analytics-filter-caption');
    if (filterCaption) {
        const labels = source.filters.map(field => field.label).join('، ');
        filterCaption.textContent = labels
            ? `متاحة لهذا المصدر: ${labels}.`
            : 'لا توجد فلاتر إضافية لهذا المصدر.';
    }
}

function renderAnalyticsYears(source, resetSelections = false) {
    const fromSelect = document.getElementById('analytics-year-from');
    const toSelect = document.getElementById('analytics-year-to');
    const years = [...new Set(source.getRows().map(row => source.getYear(row)).filter(Boolean))].sort((a, b) => a - b);

    if (!years.length) {
        fromSelect.innerHTML = '<option value="">—</option>';
        toSelect.innerHTML = '<option value="">—</option>';
        return;
    }

    const currentFrom = !resetSelections ? fromSelect.value : '';
    const currentTo = !resetSelections ? toSelect.value : '';
    const options = years.map(year => `<option value="${year}">${fmtYear(year)}</option>`).join('');
    fromSelect.innerHTML = options;
    toSelect.innerHTML = options;
    fromSelect.value = years.includes(parseInt(currentFrom, 10)) ? currentFrom : String(years[0]);
    toSelect.value = years.includes(parseInt(currentTo, 10)) ? currentTo : String(years[years.length - 1]);
}

function renderAnalyticsGroups(source, resetSelections = false) {
    const primarySelect = document.getElementById('analytics-group-primary');
    const currentPrimary = !resetSelections ? primarySelect.value : '';

    primarySelect.innerHTML = '<option value="">بدون تجميع</option>' +
        source.groups.map(group => `<option value="${group.id}">${group.label}</option>`).join('');

    if (currentPrimary && source.groups.some(group => group.id === currentPrimary)) {
        primarySelect.value = currentPrimary;
    } else {
        primarySelect.value = '';
    }
    updateAnalyticsSecondaryGroups(resetSelections);
}

function updateAnalyticsSecondaryGroups(resetSelections = false) {
    const source = getAnalyticsCurrentSource();
    const primarySelect = document.getElementById('analytics-group-primary');
    const secondarySelect = document.getElementById('analytics-group-secondary');
    const currentSecondary = !resetSelections ? secondarySelect.value : '';
    const primaryValue = primarySelect.value;

    const options = source.groups
        .filter(group => group.id !== primaryValue)
        .map(group => `<option value="${group.id}">${group.label}</option>`)
        .join('');
    secondarySelect.innerHTML = '<option value="">بدون تجميع ثانوي</option>' + options;

    if (currentSecondary && source.groups.some(group => group.id === currentSecondary && group.id !== primaryValue)) {
        secondarySelect.value = currentSecondary;
    } else {
        secondarySelect.value = '';
    }
}

function syncAnalyticsGroupingState() {
    const source = getAnalyticsCurrentSource();
    const mode = document.getElementById('analytics-mode')?.value || 'summary';
    const primarySelect = document.getElementById('analytics-group-primary');
    const secondarySelect = document.getElementById('analytics-group-secondary');
    const isDetail = mode === 'detail';

    if (!isDetail) {
        primarySelect.value = '';
        secondarySelect.value = '';
        primarySelect.disabled = true;
        secondarySelect.disabled = true;
        return;
    }

    primarySelect.disabled = false;
    if (!primarySelect.value && source.groups.length) primarySelect.value = source.groups[0].id;
    updateAnalyticsSecondaryGroups();
    secondarySelect.disabled = false;
}

function renderAnalyticsFilters(source) {
    const filtersWrap = document.getElementById('analytics-filters');
    const sourceRows = source.getRows();
    filtersWrap.innerHTML = source.filters.map(field => {
        const rawValues = typeof field.options === 'function'
            ? field.options(sourceRows)
            : typeof field.getOptionsValues === 'function'
                ? sourceRows.flatMap(row => field.getOptionsValues(row) || [])
                : sourceRows.map(row => analyticsText(field.getValue(row), '')).filter(Boolean);
        const values = [...new Set((rawValues || []).filter(Boolean))];
        if (typeof field.options !== 'function') {
            values.sort((a, b) => String(a).localeCompare(String(b), 'ar'));
        }
        return `<div class="form-group">
            <label>${field.label}</label>
            <select id="analytics-filter-${field.id}">
                <option value="">الكل</option>
                ${values.map(value => `<option value="${value}">${value}</option>`).join('')}
            </select>
        </div>`;
    }).join('');

    filtersWrap.querySelectorAll('select').forEach(select => {
        select.addEventListener('change', () => {
            document.getElementById('analytics-results').classList.add('hidden');
            currentAnalyticsReport = null;
        });
    });
}

function renderAnalyticsMetrics(source) {
    const metricsWrap = document.getElementById('analytics-metrics');
    metricsWrap.innerHTML = source.metrics.map(metric => {
        const checked = source.defaultMetrics.includes(metric.id) ? 'checked' : '';
        return `<label class="metric-chip ${checked ? 'is-checked' : ''}">
            <input type="checkbox" value="${metric.id}" ${checked}>
            <span class="metric-chip-body">
                <span class="metric-chip-title">${metric.label}</span>
                <span class="metric-chip-meta">${metric.unit ? `الوحدة: ${metric.unit}` : 'قيمة مباشرة'}</span>
            </span>
        </label>`;
    }).join('');

    metricsWrap.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', () => {
            input.closest('.metric-chip')?.classList.toggle('is-checked', input.checked);
            document.getElementById('analytics-results').classList.add('hidden');
            currentAnalyticsReport = null;
        });
    });
}

function getSelectedAnalyticsMetricDefs(source) {
    const selected = [...document.querySelectorAll('#analytics-metrics input[type="checkbox"]:checked')]
        .map(input => input.value);
    return source.metrics.filter(metric => selected.includes(metric.id));
}

function getAnalyticsFilteredRows(source) {
    const fromYear = parseInt(document.getElementById('analytics-year-from')?.value, 10) || null;
    const toYear = parseInt(document.getElementById('analytics-year-to')?.value, 10) || null;
    const yearMin = fromYear && toYear ? Math.min(fromYear, toYear) : (fromYear || toYear);
    const yearMax = fromYear && toYear ? Math.max(fromYear, toYear) : (fromYear || toYear);
    const search = document.getElementById('analytics-search')?.value || '';

    let rows = [...source.getRows()];
    rows = rows.filter(row => {
        const year = source.getYear(row);
        if (!year) return false;
        if (yearMin && year < yearMin) return false;
        if (yearMax && year > yearMax) return false;
        return true;
    });

    source.filters.forEach(field => {
        const select = document.getElementById(`analytics-filter-${field.id}`);
        const wanted = String(select?.value || '').trim();
        if (!wanted) return;
        if (source.key === 'programs' && field.id === 'branch') {
            rows = rows
                .filter(row => {
                    if (isIslamicStudiesBachelor(row)) {
                        const year = parseInt(row.Semester, 10) || null;
                        return Boolean(getIslamicBranchMetric(row, year, wanted));
                    }
                    return wanted === 'الحوية';
                })
                .map(row => buildProgramDisplayDataFromRow(row, wanted));
            return;
        }
        rows = rows.filter(row => typeof field.matches === 'function'
            ? field.matches(row, wanted)
            : analyticsText(field.getValue(row), '') === wanted
        );
    });

    if (search.trim()) {
        rows = rows.filter(row => analyticsQueryMatch(source.searchText(row), search));
    }

    return rows;
}

function computeAnalyticsMetricValues(records, metricDefs) {
    const values = {};
    metricDefs.forEach(metric => {
        values[metric.id] = metric.compute(records);
    });
    return values;
}

function buildAnalyticsYearLabel(filteredRows, source) {
    const years = [...new Set(filteredRows.map(row => source.getYear(row)).filter(Boolean))].sort((a, b) => a - b);
    if (!years.length) return '—';
    if (years.length === 1) return fmtYear(years[0]);
    return `${fmtYear(years[0])} - ${fmtYear(years[years.length - 1])}`;
}

function sortAnalyticsGroups(rows, primaryGroup, secondaryGroup) {
    const compareValue = (a, b, group) => {
        if (!group) return 0;
        if (group.sort === 'numeric') return (Number(a) || 0) - (Number(b) || 0);
        return analyticsText(group.format ? group.format(a) : a).localeCompare(
            analyticsText(group.format ? group.format(b) : b),
            'ar'
        );
    };

    rows.sort((a, b) => {
        const primaryCmp = compareValue(a.primaryValue, b.primaryValue, primaryGroup);
        if (primaryCmp !== 0) return primaryCmp;
        return compareValue(a.secondaryValue, b.secondaryValue, secondaryGroup);
    });
}

function buildAnalyticsCaption(report) {
    const parts = [
        report.source.label,
        report.mode === 'summary' ? 'تقرير إجمالي' : 'تقرير تفصيلي',
        `الفترة: ${report.yearLabel}`
    ];
    if (report.mode === 'detail' && report.primaryGroup) {
        const grouping = report.secondaryGroup
            ? `${report.primaryGroup.label} ثم ${report.secondaryGroup.label}`
            : report.primaryGroup.label;
        parts.push(`التجميع: ${grouping}`);
    }
    if (report.activeFilters.length) {
        parts.push(`الفلاتر: ${report.activeFilters.join('، ')}`);
    }
    if (report.searchText) {
        parts.push(`البحث: ${report.searchText}`);
    }
    if (report.source.key === 'programs' && report.branchFilterValue) {
        parts.push('قراءة الفرع: سجلات مرصودة متحفظة، وليست إجماليات رسمية شاملة');
    }
    return parts.join(' | ');
}

function runAnalyticsReport() {
    const source = getAnalyticsCurrentSource();
    const mode = document.getElementById('analytics-mode')?.value || 'summary';
    const branchFilterValue = source.key === 'programs' ? getAnalyticsSelectedBranchValue() : '';
    const selectedMetricDefs = getSelectedAnalyticsMetricDefs(source);
    const branchMetricLabels = {
        students_total: 'المنتظمون المرصودون',
        students_male: 'الطلاب الذكور المرصودون',
        students_female: 'الطالبات المرصودات',
        students_new: 'المستجدون المرصودون'
    };
    const branchMetricDefs = branchFilterValue
        ? selectedMetricDefs.map(metric => ({
            ...metric,
            label: branchMetricLabels[metric.id] || metric.label
        }))
        : selectedMetricDefs;

    if (!branchMetricDefs.length) {
        alert('اختر إحصائية واحدة على الأقل قبل بناء التقرير.');
        return;
    }

    const filteredRows = getAnalyticsFilteredRows(source);
    if (!filteredRows.length) {
        document.getElementById('analytics-results').classList.add('hidden');
        currentAnalyticsReport = null;
        destroyAnalyticsChart();
        alert('لا توجد بيانات مطابقة للفلاتر المختارة.');
        return;
    }
    const metricDefs = markAnalyticsSampleMetrics(branchMetricDefs, filteredRows);

    const activeFilters = source.filters
        .map(field => {
            const value = String(document.getElementById(`analytics-filter-${field.id}`)?.value || '').trim();
            return value ? `${field.label}: ${value}` : '';
        })
        .filter(Boolean);
    const yearLabel = buildAnalyticsYearLabel(filteredRows, source);
    const searchText = String(document.getElementById('analytics-search')?.value || '').trim();
    let reportRows = [];
    let primaryGroup = null;
    let secondaryGroup = null;

    if (mode === 'summary') {
        reportRows = [{
            primaryValue: 'الإجمالي',
            secondaryValue: '',
            records: filteredRows,
            metricValues: computeAnalyticsMetricValues(filteredRows, metricDefs)
        }];
    } else {
        const primaryId = document.getElementById('analytics-group-primary')?.value || '';
        const secondaryId = document.getElementById('analytics-group-secondary')?.value || '';
        primaryGroup = source.groups.find(group => group.id === primaryId) || null;
        secondaryGroup = source.groups.find(group => group.id === secondaryId) || null;

        if (!primaryGroup) {
            alert('اختر التجميع الأساسي للتقرير التفصيلي.');
            return;
        }

        const grouped = new Map();
        filteredRows.forEach(row => {
            const readGroupValues = group => {
                if (!group) return [''];
                const rawValues = typeof group.getValues === 'function'
                    ? group.getValues(row)
                    : [group.getValue(row)];
                const values = (Array.isArray(rawValues) ? rawValues : [rawValues])
                    .map(value => analyticsText(value, 'غير محدد'))
                    .filter(Boolean);
                return values.length ? [...new Set(values)] : ['غير محدد'];
            };
            const primaryValues = readGroupValues(primaryGroup);
            const secondaryValues = secondaryGroup ? readGroupValues(secondaryGroup) : [''];
            primaryValues.forEach(primaryValue => {
                secondaryValues.forEach(secondaryValue => {
                    const key = `${analyticsText(primaryValue)}||${analyticsText(secondaryValue, '')}`;
                    if (!grouped.has(key)) {
                        grouped.set(key, {
                            primaryValue,
                            secondaryValue,
                            records: []
                        });
                    }
                    grouped.get(key).records.push(row);
                });
            });
        });

        reportRows = [...grouped.values()].map(group => ({
            ...group,
            metricValues: computeAnalyticsMetricValues(group.records, metricDefs)
        }));
        sortAnalyticsGroups(reportRows, primaryGroup, secondaryGroup);
    }

    const tableHeaders = [];
    if (mode === 'detail' && primaryGroup) tableHeaders.push(primaryGroup.label);
    if (mode === 'detail' && secondaryGroup) tableHeaders.push(secondaryGroup.label);
    if (mode === 'summary') tableHeaders.push('النطاق');
    tableHeaders.push(...metricDefs.map(analyticsMetricHeader));

    const tableRows = reportRows.map(row => {
        const cells = [];
        if (mode === 'detail' && primaryGroup) {
            cells.push(primaryGroup.format ? primaryGroup.format(row.primaryValue) : analyticsText(row.primaryValue));
        }
        if (mode === 'detail' && secondaryGroup) {
            cells.push(secondaryGroup.format ? secondaryGroup.format(row.secondaryValue) : analyticsText(row.secondaryValue));
        }
        if (mode === 'summary') {
            cells.push('إجمالي البيانات المطابقة');
        }
        metricDefs.forEach(metric => {
            cells.push(metric.format(row.metricValues[metric.id]));
        });
        return cells;
    });

    currentAnalyticsReport = {
        source,
        mode,
        branchFilterValue,
        yearLabel,
        searchText,
        activeFilters,
        filteredRowsCount: filteredRows.length,
        outputRowsCount: reportRows.length,
        metricDefs,
        primaryGroup,
        secondaryGroup,
        rows: reportRows,
        tableHeaders,
        tableRows,
        caption: '',
        filenameBase: safeFileName(`الاستوديو-الإحصائي-${source.label}-${mode === 'summary' ? 'إجمالي' : 'تفصيلي'}${branchFilterValue ? `-${branchFilterValue}` : ''}`)
    };
    currentAnalyticsReport.caption = buildAnalyticsCaption(currentAnalyticsReport);

    renderAnalyticsResults();
}

function renderAnalyticsResults() {
    if (!currentAnalyticsReport) return;

    document.getElementById('analytics-caption').textContent = currentAnalyticsReport.caption;
    renderAnalyticsSummary();
    renderAnalyticsTable();
    renderAnalyticsChart();
    document.getElementById('analytics-results').classList.remove('hidden');
    document.getElementById('analytics-results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderAnalyticsSummary() {
    const container = document.getElementById('analytics-summary');
    if (!currentAnalyticsReport) {
        container.innerHTML = '';
        return;
    }

    let cards = [];
    if (currentAnalyticsReport.mode === 'summary') {
        const summaryRow = currentAnalyticsReport.rows[0];
        cards = [
            { value: fmtNum(currentAnalyticsReport.filteredRowsCount), label: currentAnalyticsReport.source.rowLabel },
            { value: currentAnalyticsReport.yearLabel, label: 'الفترة الزمنية' },
            ...currentAnalyticsReport.metricDefs.slice(0, 4).map(metric => ({
                value: metric.format(summaryRow.metricValues[metric.id]),
                label: metric.label
            }))
        ];
    } else {
        const groupingLabel = currentAnalyticsReport.secondaryGroup
            ? `${currentAnalyticsReport.primaryGroup.label} + ${currentAnalyticsReport.secondaryGroup.label}`
            : currentAnalyticsReport.primaryGroup?.label || '—';
        cards = [
            { value: currentAnalyticsReport.source.label, label: 'مصدر البيانات' },
            { value: fmtNum(currentAnalyticsReport.filteredRowsCount), label: currentAnalyticsReport.source.rowLabel },
            { value: fmtNum(currentAnalyticsReport.outputRowsCount), label: 'المجموعات الناتجة' },
            { value: currentAnalyticsReport.yearLabel, label: 'الفترة الزمنية' },
            { value: groupingLabel, label: 'مستوى التفصيل' },
            { value: fmtNum(currentAnalyticsReport.metricDefs.length), label: 'الإحصاءات المختارة' },
        ];
    }

    container.innerHTML = cards.map(card => `
        <div class="summary-card">
            <div class="sc-value">${card.value}</div>
            <div class="sc-label">${card.label}</div>
        </div>
    `).join('');
}

function renderAnalyticsTable() {
    const thead = document.getElementById('analytics-thead');
    const tbody = document.getElementById('analytics-tbody');
    if (!currentAnalyticsReport) {
        thead.innerHTML = '';
        tbody.innerHTML = '';
        return;
    }

    thead.innerHTML = `<tr>${currentAnalyticsReport.tableHeaders.map(header => `<th>${header}</th>`).join('')}</tr>`;
    tbody.innerHTML = currentAnalyticsReport.tableRows.map(row => `
        <tr>${row.map(cell => `<td>${cell}</td>`).join('')}</tr>
    `).join('');
}

function destroyAnalyticsChart() {
    if (analyticsChart) {
        analyticsChart.destroy();
        analyticsChart = null;
    }
}

function ensureAnalyticsChartCanvas() {
    const wrap = document.querySelector('#analytics-view .analytics-chart-card .chart-wrap');
    if (!wrap) return null;
    if (!document.getElementById('analytics-chart')) {
        wrap.innerHTML = '<canvas id="analytics-chart"></canvas>';
    }
    return wrap;
}

function renderAnalyticsChart() {
    const cardTitle = document.querySelector('.analytics-chart-card h3');
    const wrap = ensureAnalyticsChartCanvas();
    if (!currentAnalyticsReport || !wrap) return;

    destroyAnalyticsChart();
    const firstMetric = currentAnalyticsReport.metricDefs[0];
    if (!firstMetric) {
        wrap.innerHTML = '<div class="analytics-empty">لا توجد إحصاءات مختارة للرسم.</div>';
        return;
    }

    const ctx = document.getElementById('analytics-chart')?.getContext('2d');
    if (!ctx) return;

    if (currentAnalyticsReport.mode === 'summary') {
        const labels = currentAnalyticsReport.metricDefs.map(metric => metric.label);
        const data = currentAnalyticsReport.metricDefs.map(metric => {
            const value = currentAnalyticsReport.rows[0].metricValues[metric.id];
            return value == null ? null : Number(value);
        });
        if (cardTitle) cardTitle.textContent = 'التمثيل البياني للإحصاءات المختارة';
        analyticsChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'القيم الإجمالية',
                    data,
                    backgroundColor: labels.map((_, index) => CHART_COLORS[index % CHART_COLORS.length]),
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { callbacks: { label: item => `${item.dataset.label}: ${fmtNumFlex(item.parsed.y ?? item.parsed, 2)}` } }
                },
                scales: {
                    x: { ticks: { font: { family: 'Tajawal', size: 10 } } },
                    y: { beginAtZero: true, ticks: { font: { family: 'Tajawal' } }, grid: { color: '#f0f0f0' } }
                }
            }
        });
        return;
    }

    const maxGroups = 20;
    const chartRows = currentAnalyticsReport.rows.slice(0, maxGroups);
    if (currentAnalyticsReport.secondaryGroup) {
        const primaryGroup = currentAnalyticsReport.primaryGroup;
        const secondaryGroup = currentAnalyticsReport.secondaryGroup;
        const primaryLabels = [...new Set(chartRows.map(row => primaryGroup.format ? primaryGroup.format(row.primaryValue) : analyticsText(row.primaryValue)))];
        const secondaryLabels = [...new Set(chartRows.map(row => secondaryGroup.format ? secondaryGroup.format(row.secondaryValue) : analyticsText(row.secondaryValue)))];

        const datasets = secondaryLabels.map((secondaryLabel, index) => ({
            label: secondaryLabel,
            data: primaryLabels.map(primaryLabel => {
                const found = chartRows.find(row =>
                    (primaryGroup.format ? primaryGroup.format(row.primaryValue) : analyticsText(row.primaryValue)) === primaryLabel &&
                    (secondaryGroup.format ? secondaryGroup.format(row.secondaryValue) : analyticsText(row.secondaryValue)) === secondaryLabel
                );
                if (!found) return 0;
                const value = found.metricValues[firstMetric.id];
                return value == null ? null : Number(value);
            }),
            backgroundColor: CHART_COLORS[index % CHART_COLORS.length],
            borderRadius: 4
        }));

        if (cardTitle) cardTitle.textContent = `التمثيل البياني: ${firstMetric.label}`;
        analyticsChart = new Chart(ctx, {
            type: 'bar',
            data: { labels: primaryLabels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top', labels: { font: { family: 'Tajawal' } } } },
                scales: {
                    x: { ticks: { font: { family: 'Tajawal', size: 10 } } },
                    y: { beginAtZero: true, ticks: { font: { family: 'Tajawal' } }, grid: { color: '#f0f0f0' } }
                }
            }
        });
        return;
    }

    const labels = chartRows.map(row =>
        currentAnalyticsReport.primaryGroup?.format
            ? currentAnalyticsReport.primaryGroup.format(row.primaryValue)
            : analyticsText(row.primaryValue)
    );
    const data = chartRows.map(row => {
        const value = row.metricValues[firstMetric.id];
        return value == null ? null : Number(value);
    });

    if (cardTitle) {
        const suffix = currentAnalyticsReport.rows.length > maxGroups ? ' (أول 20 مجموعة)' : '';
        cardTitle.textContent = `التمثيل البياني: ${firstMetric.label}${suffix}`;
    }
    analyticsChart = new Chart(ctx, {
        type: currentAnalyticsReport.primaryGroup?.id === 'year' ? 'line' : 'bar',
        data: {
            labels,
            datasets: [{
                label: firstMetric.label,
                data,
                borderColor: '#0d8e8e',
                backgroundColor: currentAnalyticsReport.primaryGroup?.id === 'year'
                    ? 'rgba(13,142,142,0.15)'
                    : labels.map((_, index) => CHART_COLORS[index % CHART_COLORS.length]),
                borderRadius: 6,
                fill: currentAnalyticsReport.primaryGroup?.id === 'year',
                tension: 0.25,
                pointRadius: currentAnalyticsReport.primaryGroup?.id === 'year' ? 4 : 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: currentAnalyticsReport.primaryGroup?.id === 'year', labels: { font: { family: 'Tajawal' } } }
            },
            scales: {
                x: { ticks: { font: { family: 'Tajawal', size: 10 } } },
                y: { beginAtZero: true, ticks: { font: { family: 'Tajawal' } }, grid: { color: '#f0f0f0' } }
            }
        }
    });
}

async function exportAnalyticsPDF() {
    if (!currentAnalyticsReport) return;
    await exportSectionAsPDF('analytics-export-area', `${currentAnalyticsReport.filenameBase}.pdf`);
}

function exportAnalyticsExcel() {
    if (!currentAnalyticsReport) return;
    const rows = [
        ['الاستوديو الإحصائي'],
        ['المصدر', currentAnalyticsReport.source.label],
        ['نوع التقرير', currentAnalyticsReport.mode === 'summary' ? 'إجمالي' : 'تفصيلي'],
        ['الفترة', currentAnalyticsReport.yearLabel],
        ['الوصف', currentAnalyticsReport.caption],
        [],
        currentAnalyticsReport.tableHeaders,
        ...currentAnalyticsReport.tableRows
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = currentAnalyticsReport.tableHeaders.map((_, index) => ({
        wch: index === 0 ? 28 : 20
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'الاستوديو الإحصائي');
    XLSX.writeFile(wb, `${currentAnalyticsReport.filenameBase}.xlsx`);
}

function exportAnalyticsCSV() {
    if (!currentAnalyticsReport) return;
    downloadCSV(`${currentAnalyticsReport.filenameBase}.csv`, [
        currentAnalyticsReport.tableHeaders,
        ...currentAnalyticsReport.tableRows
    ]);
}

// ========================================
// التنقل بين العروض
// ========================================
function switchView(viewName) {
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(viewName + '-view').classList.add('active');
    document.querySelector(`.nav-btn[data-view="${viewName}"]`).classList.add('active');
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ========================================
// التهيئة
// ========================================
document.addEventListener('DOMContentLoaded', async () => {
    // Navigation
    document.querySelectorAll('.nav-btn[data-view]').forEach(btn => {
        btn.addEventListener('click', () => switchView(btn.dataset.view));
    });

    document.getElementById('loginBtn')?.addEventListener('click', handleLogin);
    document.getElementById('logoutBtn')?.addEventListener('click', handleLogout);
    document.getElementById('loginPassword')?.addEventListener('keypress', e => {
        if (e.key === 'Enter') handleLogin();
    });
    document.getElementById('loginEmployeeId')?.addEventListener('keypress', e => {
        if (e.key === 'Enter') document.getElementById('loginPassword')?.focus();
    });

    if (sessionStorage.getItem('loggedIn') === 'true') {
        showMainApp(sessionStorage.getItem('employeeName') || '');
        await startApp();
    } else {
        showLoginOverlay();
    }
});
