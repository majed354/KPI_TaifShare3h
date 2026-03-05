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
const ACTIVITIES_RAW_BASE = 'https://raw.githubusercontent.com/majed354/faculty-activities/main/data';
const RESEARCH_KPI_EXCLUDED_RANKS = new Set(['معيد', 'محاضر', 'متعاون', 'مدرس']);
const GRADUATE_PROGRAM_ALIASES = {
    'القران وعلومه': 'القراءات',
    'القرآن وعلومه': 'القراءات',
    'الدراسات القرانيه': 'الدراسات القرآنية',
    'الانظمة': 'الأنظمة',
};
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

// ========================================
// البيانات
// ========================================
let allRows = [];      // raw rows from CSV
let programs = [];     // organized {name, degree, dept, years:{y: data}}
let charts = {};       // Chart instances
let currentProg = null;// for export
let compareThirdEnabled = false;
let gradData = [];     // graduate records
let ncData = [];       // non-completer records

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

function normalizeSurveyProgramName(name) {
    const base = normalizeArabicText(name)
        .replace(/^برنامج\s+/,'')
        .replace(/^(?:ال)?(?:بكالوريوس|ماجستير|الماجستير|دكتوراه)\s+/, '');
    if (!base) return '';
    return GRADUATE_PROGRAM_ALIASES[base] || base;
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

function aggregateGraduateSurveyRows(surveyRows) {
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
        const year = parseSurveyYear(row[columns.year]);
        if (!program || !year) return;

        const key = `${year}|${program}`;
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
            eval_experience: g.experienceCount > 0
                ? Math.round((g.experienceSum / g.experienceCount) * 100) / 100 : null,
            eval_supervision: g.supervisionCount > 0
                ? Math.round((g.supervisionSum / g.supervisionCount) * 100) / 100 : null,
            eval_services: g.servicesCount > 0
                ? Math.round((g.servicesSum / g.servicesCount) * 100) / 100 : null,
            performance_rate: g.perfCount > 0
                ? Math.round((g.perfSum / g.perfCount) * 10) / 10 : null,
            employment_rate: g.employmentCount > 0
                ? pct(g.employedCount, g.employmentCount) : null,
            eval_employers: g.employerEvalCount > 0
                ? Math.round((g.employerEvalSum / g.employerEvalCount) * 100) / 100 : null,
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

function applyGraduateSurveyMetrics(rows, metricsByKey, allowedDegrees = null, sourceLabel = '') {
    let appliedRows = 0;
    const matchedGroups = new Set();
    rows.forEach(row => {
        if (!isDegreeAllowed(row, allowedDegrees)) return;

        const year = parseInt(row.Semester, 10);
        const majorKey = `${year}|${normalizeSurveyProgramName(row.Major_aName)}`;
        const deptKey = `${year}|${normalizeSurveyProgramName(normalizeDepartment(row.Dept_aName))}`;
        const metrics = metricsByKey[majorKey] || metricsByKey[deptKey];
        if (!metrics) return;

        let touched = false;
        if (metrics.eval_courses != null) {
            row.eval_courses = metrics.eval_courses;
            touched = true;
        }
        if (metrics.eval_experience != null) {
            row.eval_experience = metrics.eval_experience;
            touched = true;
        }
        if (metrics.eval_supervision != null) {
            row.eval_supervision = metrics.eval_supervision;
            touched = true;
        }
        if (metrics.eval_services != null) {
            row.eval_services = metrics.eval_services;
            touched = true;
        }
        if (metrics.performance_rate != null) {
            row.performance_rate = metrics.performance_rate;
            touched = true;
        }
        if (metrics.employment_rate != null) {
            row.employment_rate = metrics.employment_rate;
            touched = true;
        }
        if (metrics.eval_employers != null) {
            row.eval_employers = metrics.eval_employers;
            touched = true;
        }
        if (touched) {
            const baseSource = metricsByKey[majorKey] ? 'graduates_survey_program' : 'graduates_survey_dept';
            row.survey_source = sourceLabel ? `${baseSource}_${sourceLabel}` : baseSource;
            matchedGroups.add(metricsByKey[majorKey] ? majorKey : deptKey);
            appliedRows++;
        }
    });
    return { appliedRows, matchedGroups: matchedGroups.size };
}

async function applyGraduateSurveyIndicatorsFromSheet(rows, rawUrl, allowedDegrees = null, sourceLabel = '') {
    const csvUrl = resolveGraduateSurveyCsvUrl(rawUrl);
    if (!csvUrl) return { applied: false, reason: 'no-url' };

    const requestUrl = `${csvUrl}${csvUrl.includes('?') ? '&' : '?'}t=${Date.now()}`;
    const csvText = await fetchTextIfExists(requestUrl);
    if (!csvText) return { applied: false, reason: 'unreachable-sheet' };

    const surveyRows = parseCSVQuotedObjects(csvText, ',');
    if (!surveyRows.length) return { applied: false, reason: 'empty-sheet' };

    const surveyAgg = aggregateGraduateSurveyRows(surveyRows);
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
    if (!years.length) return { applied: false, reason: 'no-years' };

    const stamp = Date.now();
    const facultyPaths = [
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

    // مطابقة موقع الأنشطة: دمج البيانات الحية من Google Sheets فوق CSV
    const sheetsPayload = await fetchActivitySheetsData(configObj?.google_sheets_api);
    if (sheetsPayload?.publications?.length) {
        publicationsRows = mergeActivityPublications(publicationsRows, sheetsPayload.publications);
    }

    const citationsMap = (configObj && configObj.citations_ranges) ? configObj.citations_ranges : null;

    // فهارس أعضاء هيئة التدريس لمطابقة منطق موقع الأنشطة
    const deptIds = {};               // dept -> Set(all ids) عبر كل السنوات
    const eligibleIdsByYearDept = {}; // year|dept -> Set(ids) (نشط + مؤهل للـ KPI)
    facultyRows.forEach(row => {
        const id = pickCell(row, ['id', 'ID']);
        const year = parseInt(pickCell(row, ['year', 'Year']), 10);
        if (!id || !year) return;

        const active = pickCell(row, ['active', 'Active']) === 'نعم';
        const dept = normalizeDepartment(pickCell(row, ['department', 'Department']));
        const rank = normalizeRank(pickCell(row, ['rank', 'Rank']));

        if (dept) {
            if (!deptIds[dept]) deptIds[dept] = new Set();
            deptIds[dept].add(id);
        }
        if (!active || !dept) return;

        if (!RESEARCH_KPI_EXCLUDED_RANKS.has(rank)) {
            const key = `${year}|${dept}`;
            if (!eligibleIdsByYearDept[key]) eligibleIdsByYearDept[key] = new Set();
            eligibleIdsByYearDept[key].add(id);
        }
    });

    // author -> depts (من جميع سنوات faculty كما في موقع الأنشطة)
    const authorDeptMap = {};
    Object.entries(deptIds).forEach(([dept, idsSet]) => {
        idsSet.forEach(fid => {
            if (!authorDeptMap[fid]) authorDeptMap[fid] = new Set();
            authorDeptMap[fid].add(dept);
        });
    });

    // aggregates per year+dept
    const publishingMembersByYearDept = {};
    const publicationsCountByYearDept = {};
    const citationsTotalByYearDept = {};

    publicationsRows.forEach(pub => {
        const year = parseInt(pickCell(pub, ['year', 'Year']), 10);
        if (!year) return;
        if (!years.includes(year)) return;

        const authorIds = splitIdsPipe(pickCell(pub, ['authors_ids', 'participant_ids']));
        if (!authorIds.length) return;

        const departmentsTouched = new Set();
        authorIds.forEach(fid => {
            const depts = authorDeptMap[fid];
            if (!depts) return;
            depts.forEach(d => departmentsTouched.add(d));
        });

        if (!departmentsTouched.size) return;
        const citations = parseCitationValue(pickCell(pub, ['citations_range', 'Citations', 'citations']), citationsMap);
        departmentsTouched.forEach(dept => {
            const ydKey = `${year}|${dept}`;
            publicationsCountByYearDept[ydKey] = (publicationsCountByYearDept[ydKey] || 0) + 1;
            citationsTotalByYearDept[ydKey] = (citationsTotalByYearDept[ydKey] || 0) + citations;

            const eligible = eligibleIdsByYearDept[ydKey];
            if (!eligible || !eligible.size) return;
            if (!publishingMembersByYearDept[ydKey]) publishingMembersByYearDept[ydKey] = new Set();
            authorIds.forEach(fid => {
                if (eligible.has(fid)) publishingMembersByYearDept[ydKey].add(fid);
            });
        });
    });

    let appliedRows = 0;
    rows.forEach(r => {
        const year = absYearFromSemester(r.Semester);
        const dept = normalizeDepartment(r.Dept_aName);
        const ydKey = `${year}|${dept}`;

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

    return { applied: appliedRows > 0, appliedRows };
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
    const [plansText, facultyText] = await Promise.all([
        fetchTextIfExists(`data/new_all_plans.csv?t=${stamp}`),
        fetchTextIfExists(`data/faculty.csv?t=${stamp}`)
    ]);

    if (!plansText || !facultyText) {
        return { applied: false, reason: 'missing-plans-or-faculty' };
    }

    const planRows = parseFlatCSV(plansText, ';');
    const facultyRows = parseFlatCSV(facultyText, ',');
    if (!planRows.length || !facultyRows.length) {
        return { applied: false, reason: 'empty-plans-or-faculty' };
    }

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
        const major = String(r.Major_aName || '').trim();
        const degree = normalizeDegree(r.Degree_aName);
        if (!major || !SUPPORTED_KPI_DEGREES.has(degree)) return;

        const pKey = `${year}|${dept}|${major}|${degree}`;
        studentsByProgramKey[pKey] = Number(r.students_total) || 0;
        programMetaByKey[pKey] = { year, dept, major, degree };

        const mdKey = `${year}|${major}|${degree}`;
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
    planRows.forEach(row => {
        const code = pickCell(row, ['Code', '\uFEFFCode', 'رمز المقرر']);
        const major = pickCell(row, ['Program', '\uFEFFProgram', 'البرنامج']);
        const degree = normalizeDegree(pickCell(row, ['Degree', 'الدرجة']));
        if (!code || !major || !SUPPORTED_KPI_DEGREES.has(degree)) return;
        if (!courseToPrograms[code]) courseToPrograms[code] = [];
        if (!courseToPrograms[code].some(x => x.major === major && x.degree === degree)) {
            courseToPrograms[code].push({ major, degree });
        }
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
            active: pickCell(row, ['active', 'Active']) === 'نعم'
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

    // 1) تحميل تدريسي فعلي من ملفات teaching/years/*.json
    Object.entries(teachingByYear).forEach(([yearStr, records]) => {
        const year = parseInt(yearStr, 10);
        records.forEach(rec => {
            const fid = String(rec.fid || '').trim();
            if (!fid) return;
            const facKey = `${year}|${fid}`;
            const profile = resolveFacultyProfile(facultyProfilesByYearId, facultyProfilesById, year, fid);
            const deptHint = normalizeDepartment(profile?.dept || '');
            if (!facultyProgramLoads[facKey]) facultyProgramLoads[facKey] = {};

            const courses = Array.isArray(rec.cs) ? rec.cs : [];
            courses.forEach(c => {
                const secDegree = normalizeDegree(c.dg);
                if (!SUPPORTED_KPI_DEGREES.has(secDegree)) return;

                const loadHours = Number(c.h);
                const load = Number.isFinite(loadHours) && loadHours > 0 ? loadHours : 1;
                const code = String(c.cc || '').trim();

                let candidateProgramKeys = [];

                // محاولة ربط مباشر من رمز المقرر
                const mapped = courseToPrograms[code] || [];
                mapped.forEach(mp => {
                    if (mp.degree !== secDegree) return;
                    const mdKey = `${year}|${mp.major}|${mp.degree}`;
                    const options = programLookupByYearMajorDegree[mdKey] || [];
                    if (!options.length) return;
                    if (options.length === 1) {
                        candidateProgramKeys.push(options[0]);
                        return;
                    }
                    const deptMatched = deptHint ? options.filter(k => k.split('|')[1] === deptHint) : [];
                    candidateProgramKeys.push(...(deptMatched.length ? deptMatched : options));
                });

                // fallback: لو الرمز غير موجود في الخطط نوزع داخل القسم/الدرجة
                candidateProgramKeys = [...new Set(candidateProgramKeys)];
                if (!candidateProgramKeys.length && deptHint) {
                    candidateProgramKeys = getDeptDegreeProgramKeys(programsByDeptYear, year, deptHint, secDegree);
                }
                if (!candidateProgramKeys.length) return;

                const weights = buildWeights(candidateProgramKeys, studentsByProgramKey);
                weights.forEach(w => {
                    facultyProgramLoads[facKey][w.key] = (facultyProgramLoads[facKey][w.key] || 0) + (load * w.weight);
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
        const baseFTE = getRankBaseFTE(profile?.rank);
        Object.entries(byProgram).forEach(([pKey, load]) => {
            fteByProgramKey[pKey] = (fteByProgramKey[pKey] || 0) + ((load / totalLoad) * baseFTE);
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
            const weights = buildWeights(candidateKeys, studentsByProgramKey);
            weights.forEach(w => {
                fteByProgramKey[w.key] = (fteByProgramKey[w.key] || 0) + (w.weight * baseFTE);
            });
        });
    });

    let programsWithComputedFTE = 0;
    rows.forEach(r => {
        const year = absYearFromSemester(r.Semester);
        const dept = normalizeDepartment(r.Dept_aName);
        const major = String(r.Major_aName || '').trim();
        const degree = normalizeDegree(r.Degree_aName);
        const pKey = `${year}|${dept}|${major}|${degree}`;
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
    };
}

// ========================================
// تحميل وتحليل البيانات
// ========================================
async function loadData() {
    const dot = document.getElementById('status-dot');
    const txt = document.getElementById('status-text');
    try {
        const res = await fetch('data/data.csv?t=' + Date.now());
        const csv = await res.text();
        allRows = parseCSV(csv);
        const durationInfo = await applyAverageGraduationDurationFromDetails(allRows);
        const surveyInfo = await applyGraduateSurveyIndicators(allRows);
        const researchInfo = await applyResearchIndicatorsFromActivities(allRows);
        const fteInfo = await applyTeachingBasedFacultyFTE(allRows);
        programs = buildPrograms(allRows);
        dot.className = 'dot ok';
        const parts = [`تم تحميل ${programs.length} برنامج - ${allRows.length} سجل`];
        if (durationInfo.applied) {
            parts.push(`احتساب متوسط مدة التخرج من سجل الخريجين (${durationInfo.appliedRows} سجل)`);
        }
        if (surveyInfo.applied) {
            const srcCount = surveyInfo.sourcesUsed || 1;
            parts.push(`تحديث مؤشرات الخريجين من الشيت (${surveyInfo.appliedRows} سجل، ${srcCount} مصدر)`);
        } else if (
            GRADUATES_SURVEY_SHEET_URL ||
            GRADUATES_SURVEY_BACHELOR_SHEET_URL ||
            GRADUATES_SURVEY_POSTGRAD_SHEET_URL
        ) {
            parts.push('تعذر تحديث مؤشرات الخريجين من الشيت');
        }
        if (researchInfo.applied) parts.push(`تحديث مؤشرات البحث من faculty-activities (${researchInfo.appliedRows} سجل)`);
        if (fteInfo.applied) parts.push(`احتساب هيئة التدريس من التدريس الفعلي (${fteInfo.programsWithComputedFTE} سجل)`);
        txt.textContent = parts.join(' | ');
        return true;
    } catch (e) {
        console.error(e);
        dot.className = 'dot err';
        txt.textContent = 'خطأ في تحميل البيانات';
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
    if (d.research_source === 'faculty_activities_live' && Number.isFinite(d.citations_per_publication)) {
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
function initProgramView() {
    const sel = document.getElementById('prog-select');
    sel.innerHTML = '<option value="">-- اختر البرنامج --</option>' +
        programs.map((p,i) => `<option value="${i}">${p.name} (${p.degree})</option>`).join('');

    sel.addEventListener('change', () => {
        const v = sel.value;
        document.getElementById('prog-results').classList.add('hidden');
        if (v !== '') {
            fillProgYears(parseInt(v));
        } else {
            document.getElementById('prog-year').disabled = true;
            document.getElementById('prog-show').disabled = true;
        }
    });

    document.getElementById('prog-year').addEventListener('change', e => {
        document.getElementById('prog-show').disabled = !e.target.value;
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
    if (isNaN(idx) || isNaN(year)) return;

    const prog = programs[idx];
    const d = prog.years[year];
    if (!d) return;

    currentProg = { prog, year, data: d };

    // Header
    document.getElementById('prog-badge').textContent = prog.degree;
    document.getElementById('prog-name').textContent = prog.name;
    document.getElementById('prog-year-label').textContent = 'السنة: ' + fmtYear(year) + ' | القسم: ' + prog.dept;

    // Stats
    const stats = [];
    if (d.students_total > 0) stats.push({icon:'👥', label:'إجمالي الطلاب', value: fmtNum(d.students_total), color:'var(--primary)'});
    if (d.students_male > 0) stats.push({icon:'👨', label:'الذكور', value: fmtNum(d.students_male), color:'#3b82f6'});
    if (d.students_female > 0) stats.push({icon:'👩', label:'الإناث', value: fmtNum(d.students_female), color:'#ec4899'});
    if (d.students_saudi > 0) stats.push({icon:'🇸🇦', label:'السعوديون', value: fmtNum(d.students_saudi), color:'#059669'});
    if (d.students_international > 0) stats.push({icon:'🌍', label:'الدوليون', value: fmtNum(d.students_international), color:'#f97316'});
    if (d.students_new > 0) stats.push({icon:'🆕', label:'المستجدون', value: fmtNum(d.students_new), color:'#06b6d4'});
    if (d.students_retained > 0 || d.prev_new_count > 0) stats.push({icon:'🔄', label:'استبقاء الدفعة السابقة', value: d.prev_new_count > 0 ? fmtNum(d.students_retained) + ' من ' + fmtNum(d.prev_new_count) : fmtNum(d.students_retained), color:'#8b5cf6'});
    if (d.graduates_total > 0) stats.push({icon:'🎓', label:'إجمالي الخريجين', value: fmtNum(d.graduates_total), color:'#10b981'});
    if (d.graduates_ontime > 0 || d.new_4_ago_count > 0) stats.push({icon:'⏱️', label:'خريجو الدفعة بالوقت', value: d.new_4_ago_count > 0 ? fmtNum(d.graduates_ontime) + ' من ' + fmtNum(d.new_4_ago_count) : fmtNum(d.graduates_ontime), color:'#0d8e8e'});
    if (d.sections_total > 0) stats.push({icon:'🏛️', label:'الشعب', value: fmtNum(d.sections_total), color:'#6b7280'});
    const facultyBase = getFacultyBaseForRatio(d);
    if (facultyBase > 0) {
        const facultyLabel = d.faculty_ratio_source === 'teaching_fte'
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
            return `<div class="kpi-card">
                <div class="kpi-num">${ind.code || ind.id}</div>
                <div class="kpi-body">
                    <div class="kpi-name">${ind.name}</div>
                    <div class="kpi-val ${f.cls}">${f.text} <span class="kpi-unit">${v != null ? ind.unit : ''}</span></div>
                    ${detailHtml}
                </div>
            </div>`;
        }).join('');

    // Trend chart
    renderTrendChart(prog);

    document.getElementById('prog-results').classList.remove('hidden');
    document.getElementById('prog-results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderTrendChart(prog) {
    destroyChart('trend');
    const years = Object.keys(prog.years).map(Number).filter(y => DISPLAY_YEARS.includes(y)).sort();
    const labels = years.map(y => fmtYear(y));
    const students = years.map(y => prog.years[y].students_total);
    const grads = years.map(y => prog.years[y].graduates_total);
    const newS = years.map(y => prog.years[y].students_new);

    const ctx = document.getElementById('chart-trend').getContext('2d');
    charts.trend = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'الطلاب المنتظمون', data: students, borderColor: '#0d8e8e', backgroundColor: 'rgba(13,142,142,0.1)', fill: true, tension: 0.3, pointRadius: 5 },
                { label: 'الخريجين', data: grads, borderColor: '#10b981', backgroundColor: 'transparent', tension: 0.3, pointRadius: 5 },
                { label: 'المستجدون', data: newS, borderColor: '#3b82f6', backgroundColor: 'transparent', tension: 0.3, pointRadius: 5 },
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
function initCompare() {
    const progOpts = '<option value="">-- اختر البرنامج --</option>' +
        programs.map((p,i) => `<option value="${i}">${p.name} (${p.degree}) - ${p.dept}</option>`).join('');
    ['cmp-a-prog','cmp-b-prog','cmp-c-prog'].forEach(id => {
        document.getElementById(id).innerHTML = progOpts;
    });

    ['a','b','c'].forEach(slot => {
        document.getElementById(`cmp-${slot}-prog`).addEventListener('change', () => {
            populateCompareYears(slot);
            document.getElementById('cmp-results').classList.add('hidden');
            updateCmpBtn();
        });
        document.getElementById(`cmp-${slot}-year`).addEventListener('change', () => {
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
    const hasProgram = progValue !== '';
    const hasYear = yearValue !== '';
    if (!hasProgram || !hasYear) return null;
    const programIndex = parseInt(progValue, 10);
    const year = parseInt(yearValue, 10);
    const program = programs[programIndex];
    if (!program || !program.years[year]) return null;
    return { slot, programIndex, program, year, data: program.years[year] };
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
    return `${entry.program.name} - ${fmtYear(entry.year)}`;
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
        return { indicator: ind, rawValues, formattedValues, diffs };
    });
    return { header, rows, indicators };
}

function formatDiffCell(diff) {
    if (diff == null) return '<span class="diff-same">—</span>';
    if (diff > 0) return `<span class="diff-up">+${fmtNumFlex(diff, 2)}</span>`;
    if (diff < 0) return `<span class="diff-down">${fmtNumFlex(diff, 2)}</span>`;
    return '<span class="diff-same">0</span>';
}

function showComparison() {
    const selections = [getCompareSelection('a'), getCompareSelection('b')];
    if (compareThirdEnabled) selections.push(getCompareSelection('c'));
    if (selections.some(s => !s)) return alert('أكمل اختيار البرنامج والسنة لكل مقارنة مطلوبة');

    const entries = selections.map(sel => ({
        ...sel,
        label: `${sel.program.name} (${sel.program.degree}) - ${fmtYear(sel.year)}`,
        shortLabel: compareShortLabel(sel),
        kpi: calcKPIs(sel.data, sel.program.degree)
    }));

    const model = buildComparisonModel(entries);

    document.getElementById('cmp-thead').innerHTML =
        `<tr>${model.header.map(h => `<th>${h}</th>`).join('')}</tr>`;

    document.getElementById('cmp-tbody').innerHTML = model.rows.map(row => {
        const valueCells = row.formattedValues.map(v => `<td>${v}</td>`).join('');
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

function getProgramIndicatorRows() {
    if (!currentProg || currentProg.cmp) return [];
    const kpi = calcKPIs(currentProg.data, currentProg.prog.degree);
    return getIndicatorsForDegree(currentProg.prog.degree)
        .map(ind => {
            const f = fmtKPI(kpi[ind.key], ind.unit);
            return [getIndicatorLabel(ind), f.text, ind.unit];
        });
}

async function exportPDF() {
    if (!currentProg || currentProg.cmp) return;
    const filename = `مؤشرات-${safeFileName(currentProg.prog.name)}-${fmtYear(currentProg.year)}.pdf`;
    await exportSectionAsPDF('prog-results', filename);
}

function exportExcel() {
    if (!currentProg || currentProg.cmp) return;
    const rows = [
        ['تقرير مؤشرات الأداء'],
        ['البرنامج', currentProg.prog.name],
        ['الدرجة', currentProg.prog.degree],
        ['السنة', fmtYear(currentProg.year)],
        [],
        ['المؤشر', 'القيمة', 'الوحدة'],
        ...getProgramIndicatorRows()
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 42 }, { wch: 18 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المؤشرات');
    const filename = `مؤشرات-${safeFileName(currentProg.prog.name)}-${fmtYear(currentProg.year)}.xlsx`;
    XLSX.writeFile(wb, filename);
}

function exportCSV() {
    if (!currentProg || currentProg.cmp) return;
    const rows = [
        ['المؤشر', 'القيمة', 'الوحدة'],
        ...getProgramIndicatorRows()
    ];
    const filename = `مؤشرات-${safeFileName(currentProg.prog.name)}-${fmtYear(currentProg.year)}.csv`;
    downloadCSV(filename, rows);
}

function getCompareExportRows() {
    if (!currentProg || !currentProg.cmp || !currentProg.model) return [];
    return currentProg.model.rows.map(row => [
        getIndicatorLabel(row.indicator),
        ...row.formattedValues,
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
        gradData = parseDetailCSV(csv);
        return true;
    } catch (e) {
        console.error('خطأ في تحميل بيانات الخريجين:', e);
        return false;
    }
}

async function loadNonCompleters() {
    try {
        const res = await fetch('data/non_completers.csv?t=' + Date.now());
        const csv = await res.text();
        ncData = parseDetailCSV(csv);
        return true;
    } catch (e) {
        console.error('خطأ في تحميل بيانات غير المكملين:', e);
        return false;
    }
}

function initGraduatesView() {
    if (!gradData.length) return;

    // Populate filters
    const years = [...new Set(gradData.map(g => g['السنة']))].sort();
    const progs = [...new Set(gradData.map(g => g['التخصص']))].sort((a,b) => a.localeCompare(b,'ar'));
    const degs = [...new Set(gradData.map(g => g['الدرجة']))];

    const ySel = document.getElementById('grad-year');
    ySel.innerHTML = '<option value="">الكل</option>' +
        years.map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');

    const pSel = document.getElementById('grad-prog');
    pSel.innerHTML = '<option value="">الكل</option>' +
        progs.map(p => `<option value="${p}">${p}</option>`).join('');

    const dSel = document.getElementById('grad-deg');
    dSel.innerHTML = '<option value="">الكل</option>' +
        degs.map(d => `<option value="${d}">${d}</option>`).join('');

    // Event listeners
    [ySel, pSel, dSel].forEach(el => el.addEventListener('change', renderGraduates));
    document.getElementById('grad-search').addEventListener('input', renderGraduates);

    renderGraduates();
}

function renderGraduates() {
    const yearFilter = document.getElementById('grad-year').value;
    const progFilter = document.getElementById('grad-prog').value;
    const degFilter = document.getElementById('grad-deg').value;
    const search = document.getElementById('grad-search').value.trim().toLowerCase();

    let filtered = gradData;
    if (yearFilter) filtered = filtered.filter(g => g['السنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (search) filtered = filtered.filter(g =>
        g['الاسم'].toLowerCase().includes(search) ||
        g['الرقم_الجامعي'].includes(search)
    );

    document.getElementById('grad-count').textContent =
        `${filtered.length.toLocaleString('ar-SA')} سجل من أصل ${gradData.length.toLocaleString('ar-SA')}`;

    const tbody = document.getElementById('grad-tbody');
    const MAX_SHOW = 500;
    const showing = filtered.slice(0, MAX_SHOW);

    tbody.innerHTML = showing.map((g, i) => `<tr class="${i%2?'alt':''}">
        <td>${i+1}</td>
        <td>${fmtYear(g['السنة'])}</td>
        <td>${g['الرقم_الجامعي']}</td>
        <td>${g['الاسم']}</td>
        <td>${g['التخصص']}</td>
        <td>${g['الدرجة']}</td>
        <td>${g['الجنس']}</td>
        <td>${g['الجنسية']}</td>
        <td>${g['تاريخ_القبول']}</td>
        <td>${g['تاريخ_التخرج']}</td>
        <td>${g['المعدل']}</td>
    </tr>`).join('');

    if (filtered.length > MAX_SHOW) {
        tbody.innerHTML += `<tr><td colspan="11" style="text-align:center;color:var(--text-light);padding:16px">
            يتم عرض أول ${MAX_SHOW} سجل. استخدم الفلاتر لتصفية النتائج أو صدّر Excel لرؤية الكل.
        </td></tr>`;
    }
}

function exportGradsExcel() {
    const yearFilter = document.getElementById('grad-year').value;
    const progFilter = document.getElementById('grad-prog').value;
    const degFilter = document.getElementById('grad-deg').value;
    const search = document.getElementById('grad-search').value.trim().toLowerCase();

    let filtered = gradData;
    if (yearFilter) filtered = filtered.filter(g => g['السنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (search) filtered = filtered.filter(g =>
        g['الاسم'].toLowerCase().includes(search) || g['الرقم_الجامعي'].includes(search)
    );

    const rows = [
        ['السنة','الرقم الجامعي','الاسم','التخصص','الدرجة','القسم','الجنس','الجنسية','تاريخ القبول','تاريخ التخرج','المعدل'],
        ...filtered.map(g => [
            fmtYear(g['السنة']), g['الرقم_الجامعي'], g['الاسم'], g['التخصص'], g['الدرجة'],
            g['القسم'], g['الجنس'], g['الجنسية'], g['تاريخ_القبول'], g['تاريخ_التخرج'], g['المعدل']
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

    document.getElementById('nc-year').innerHTML = '<option value="">الكل</option>' +
        years.map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');
    document.getElementById('nc-prog').innerHTML = '<option value="">الكل</option>' +
        progs.map(p => `<option value="${p}">${p}</option>`).join('');
    document.getElementById('nc-deg').innerHTML = '<option value="">الكل</option>' +
        degs.map(d => `<option value="${d}">${d}</option>`).join('');
    document.getElementById('nc-status').innerHTML = '<option value="">الكل</option>' +
        statuses.map(s => `<option value="${s}">${s}</option>`).join('');

    ['nc-year','nc-prog','nc-deg','nc-status'].forEach(id =>
        document.getElementById(id).addEventListener('change', renderNonCompleters));
    document.getElementById('nc-search').addEventListener('input', renderNonCompleters);

    renderNonCompleters();
}

function renderNonCompleters() {
    const yearFilter = document.getElementById('nc-year').value;
    const progFilter = document.getElementById('nc-prog').value;
    const degFilter = document.getElementById('nc-deg').value;
    const statusFilter = document.getElementById('nc-status').value;
    const search = document.getElementById('nc-search').value.trim().toLowerCase();

    let filtered = ncData;
    if (yearFilter) filtered = filtered.filter(g => g['آخر_سنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (statusFilter) filtered = filtered.filter(g => g['الحالة'] === statusFilter);
    if (search) filtered = filtered.filter(g =>
        g['الاسم'].toLowerCase().includes(search) || g['الرقم_الجامعي'].includes(search)
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
                <span class="sc-label"><span class="status-badge ${info.cls}">${info.label}</span></span>
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
            <td>${fmtYear(nc['آخر_سنة'])}</td>
            <td>${nc['الرقم_الجامعي']}</td>
            <td>${nc['الاسم']}</td>
            <td>${nc['التخصص']}</td>
            <td>${nc['الدرجة']}</td>
            <td><span class="status-badge ${info.cls}">${nc['الحالة']}</span></td>
            <td>${nc['الجنس']}</td>
            <td>${nc['الجنسية']}</td>
            <td>${nc['تاريخ_القبول']}</td>
            <td>${nc['المعدل']}</td>
        </tr>`;
    }).join('');

    if (filtered.length > MAX_SHOW) {
        tbody.innerHTML += `<tr><td colspan="11" style="text-align:center;color:var(--text-light);padding:16px">
            يتم عرض أول ${MAX_SHOW} سجل. استخدم الفلاتر لتصفية النتائج أو صدّر Excel لرؤية الكل.
        </td></tr>`;
    }
}

function exportNCExcel() {
    const yearFilter = document.getElementById('nc-year').value;
    const progFilter = document.getElementById('nc-prog').value;
    const degFilter = document.getElementById('nc-deg').value;
    const statusFilter = document.getElementById('nc-status').value;
    const search = document.getElementById('nc-search').value.trim().toLowerCase();

    let filtered = ncData;
    if (yearFilter) filtered = filtered.filter(g => g['آخر_سنة'] === yearFilter);
    if (progFilter) filtered = filtered.filter(g => g['التخصص'] === progFilter);
    if (degFilter) filtered = filtered.filter(g => g['الدرجة'] === degFilter);
    if (statusFilter) filtered = filtered.filter(g => g['الحالة'] === statusFilter);
    if (search) filtered = filtered.filter(g =>
        g['الاسم'].toLowerCase().includes(search) || g['الرقم_الجامعي'].includes(search)
    );

    const rows = [
        ['آخر سنة','الرقم الجامعي','الاسم','التخصص','الدرجة','القسم','الحالة','الجنس','الجنسية','تاريخ القبول','المعدل','نوع الدراسة'],
        ...filtered.map(nc => [
            fmtYear(nc['آخر_سنة']), nc['الرقم_الجامعي'], nc['الاسم'], nc['التخصص'], nc['الدرجة'],
            nc['القسم'], nc['الحالة'], nc['الجنس'], nc['الجنسية'], nc['تاريخ_القبول'], nc['المعدل'], nc['نوع_الدراسة']
        ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'غير المكملين');
    XLSX.writeFile(wb, `غير-المكملين${statusFilter ? '-'+statusFilter : ''}.xlsx`);
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

    // Load data
    const ok = await loadData();
    if (!ok) return;

    // Init views
    initDashboard();
    initProgramView();
    initCompare();

    // Load detail records
    await Promise.all([loadGraduates(), loadNonCompleters()]);
    initGraduatesView();
    initNonCompleteView();
});
