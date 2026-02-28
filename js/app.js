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
const INDICATORS = [
    { id:1,  name:"تقويم الطالب لجودة خبرات التعلم", unit:"درجة", key:"experience_eval", numeric:true },
    { id:2,  name:"تقييم الطالب لجودة المقررات",     unit:"درجة", key:"course_eval",     numeric:true },
    { id:3,  name:"معدّل التخرج بالوقت المحدد",      unit:"%",    key:"graduation_rate", numeric:true },
    { id:4,  name:"معدّل استبقاء طلاب السنة الأولى",  unit:"%",    key:"retention_rate",  numeric:true },
    { id:5,  name:"مستوى أداء الطالب",                unit:"%",    key:"student_performance", numeric:true },
    { id:6,  name:"توظيف الخريجين",                   unit:"%",    key:"employment_rate", numeric:true },
    { id:7,  name:"تقويم جهات التوظيف",               unit:"درجة", key:"employer_eval",   numeric:true },
    { id:8,  name:"نسبة الطلاب/هيئة التدريس",         unit:"نسبة", key:"student_faculty_ratio" },
    { id:9,  name:"نسبة النشر العلمي",                unit:"%",    key:"publication_pct", numeric:true },
    { id:10, name:"البحوث/عضو هيئة تدريس",           unit:"بحث",  key:"research_per_faculty", numeric:true },
    { id:11, name:"الاقتباسات/عضو هيئة تدريس",       unit:"اقتباس", key:"citations_per_faculty", numeric:true },
    { id:12, name:"نسبة نشر طلاب الدراسات العليا",   unit:"%",    key:"student_publication", numeric:true, gradOnly:true },
    { id:13, name:"براءات الاختراع",                  unit:"براءة", key:"patents", numeric:true, gradOnly:true },
];

// ========================================
// البيانات
// ========================================
let allRows = [];      // raw rows from CSV
let programs = [];     // organized {name, degree, dept, years:{y: data}}
let charts = {};       // Chart instances
let currentProg = null;// for export
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
        const fteInfo = await applyTeachingBasedFacultyFTE(allRows);
        programs = buildPrograms(allRows);
        dot.className = 'dot ok';
        if (fteInfo.applied) {
            txt.textContent = `تم تحميل ${programs.length} برنامج - ${allRows.length} سجل | تم احتساب هيئة التدريس من التدريس الفعلي (${fteInfo.programsWithComputedFTE} سجل)`;
        } else {
            txt.textContent = `تم تحميل ${programs.length} برنامج - ${allRows.length} سجل`;
        }
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
            'research_count','citations',
            'eval_courses','eval_experience','eval_employers',
            'performance_rate','employment_rate'
        ];
        numFields.forEach(f => { row[f] = parseFloat(row[f]) || 0; });
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
    kpi.experience_eval = d.eval_experience || null;
    kpi.course_eval = d.eval_courses || null;
    kpi.employer_eval = d.eval_employers || null;
    kpi.student_performance = d.performance_rate || null;
    kpi.employment_rate = d.employment_rate || null;

    // معدل التخرج بالوقت المحدد
    kpi.graduation_rate = pct(d.graduates_ontime, d.new_4_ago_count);
    kpi.graduation_detail = d.new_4_ago_count > 0 ? `${d.graduates_ontime} من ${d.new_4_ago_count}` : null;

    // معدل الاستبقاء
    kpi.retention_rate = pct(d.students_retained, d.prev_new_count);
    kpi.retention_detail = d.prev_new_count > 0 ? `${d.students_retained} من ${d.prev_new_count}` : null;

    // نسبة الطلاب/هيئة التدريس
    const facultyBase = getFacultyBaseForRatio(d);
    if (facultyBase > 0 && d.students_total > 0) {
        kpi.student_faculty_ratio = `1:${(d.students_total / facultyBase).toFixed(1)}`;
    } else {
        kpi.student_faculty_ratio = null;
    }

    // نسبة النشر العلمي
    kpi.publication_pct = d.faculty_total > 0 && d.faculty_published > 0
        ? pct(d.faculty_published, d.faculty_total) : null;

    // البحوث/عضو
    kpi.research_per_faculty = d.faculty_total > 0 && d.research_count > 0
        ? Math.round((d.research_count / d.faculty_total) * 100) / 100 : null;

    // الاقتباسات/عضو
    kpi.citations_per_faculty = d.faculty_total > 0 && d.citations > 0
        ? Math.round((d.citations / d.faculty_total) * 10) / 10 : null;

    kpi.student_publication = null;
    kpi.patents = null;

    return kpi;
}

function fmtKPI(val, unit) {
    if (val == null) return { text: 'غير متوفر', cls: 'na' };
    if (unit === '%') return { text: val.toFixed(1) + '%', cls: '' };
    if (unit === 'درجة') return { text: parseFloat(val).toFixed(2), cls: '' };
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
    document.getElementById('kpi-grid').innerHTML = INDICATORS
        .filter(ind => !ind.gradOnly || ['الماجستير','دكتوراه'].includes(prog.degree))
        .map(ind => {
            const v = kpi[ind.key];
            const f = fmtKPI(v, ind.unit);
            // إضافة تفاصيل "X من Y" لمعدل التخرج والاستبقاء
            let detailHtml = '';
            if (ind.key === 'graduation_rate' && kpi.graduation_detail) {
                detailHtml = `<div class="kpi-detail">(${kpi.graduation_detail})</div>`;
            } else if (ind.key === 'retention_rate' && kpi.retention_detail) {
                detailHtml = `<div class="kpi-detail">(${kpi.retention_detail})</div>`;
            }
            return `<div class="kpi-card">
                <div class="kpi-num">${ind.id}</div>
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
    // fill selects
    const progOpts = '<option value="">-- اختر --</option>' +
        programs.map((p,i) => `<option value="${i}">${p.name} (${p.degree})</option>`).join('');
    document.getElementById('cmp-prog').innerHTML = progOpts;
    document.getElementById('cmp-p1').innerHTML = progOpts;
    document.getElementById('cmp-p2').innerHTML = progOpts;

    const yearOpts = '<option value="">-- اختر --</option>' +
        getAvailableYears().map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');
    document.getElementById('cmp-year').innerHTML = yearOpts;

    // Mode toggle
    document.querySelectorAll('.mode-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.mode-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById('cmp-years-form').classList.toggle('hidden', btn.dataset.mode !== 'years');
            document.getElementById('cmp-progs-form').classList.toggle('hidden', btn.dataset.mode !== 'progs');
            document.getElementById('cmp-results').classList.add('hidden');
            document.getElementById('cmp-btn').disabled = true;
        });
    });

    // Years mode
    document.getElementById('cmp-prog').addEventListener('change', e => {
        if (e.target.value !== '') {
            const idx = parseInt(e.target.value);
            const p = programs[idx];
            const years = Object.keys(p.years).map(Number).filter(y => DISPLAY_YEARS.includes(y)).sort();
            const opts = '<option value="">--</option>' + years.map(y => `<option value="${y}">${fmtYear(y)}</option>`).join('');
            document.getElementById('cmp-y1').innerHTML = opts;
            document.getElementById('cmp-y2').innerHTML = opts;
            document.getElementById('cmp-y1').disabled = false;
            document.getElementById('cmp-y2').disabled = false;
        }
        updateCmpBtn();
    });

    // Programs mode
    document.getElementById('cmp-year').addEventListener('change', e => {
        if (e.target.value) {
            const yr = parseInt(e.target.value);
            const avail = programs.filter(p => p.years[yr]);
            const opts = '<option value="">--</option>' + avail.map((p,i) => {
                const realIdx = programs.indexOf(p);
                return `<option value="${realIdx}">${p.name} (${p.degree})</option>`;
            }).join('');
            document.getElementById('cmp-p1').innerHTML = opts;
            document.getElementById('cmp-p2').innerHTML = opts;
            document.getElementById('cmp-p1').disabled = false;
            document.getElementById('cmp-p2').disabled = false;
        }
        updateCmpBtn();
    });

    ['cmp-y1','cmp-y2','cmp-p1','cmp-p2'].forEach(id => {
        document.getElementById(id).addEventListener('change', updateCmpBtn);
    });

    document.getElementById('cmp-btn').addEventListener('click', showComparison);
}

function updateCmpBtn() {
    const mode = document.querySelector('.mode-btn.active').dataset.mode;
    let valid = false;
    if (mode === 'years') {
        const p = document.getElementById('cmp-prog').value;
        const y1 = document.getElementById('cmp-y1').value;
        const y2 = document.getElementById('cmp-y2').value;
        valid = p && y1 && y2 && y1 !== y2;
    } else {
        const yr = document.getElementById('cmp-year').value;
        const p1 = document.getElementById('cmp-p1').value;
        const p2 = document.getElementById('cmp-p2').value;
        valid = yr && p1 && p2 && p1 !== p2;
    }
    document.getElementById('cmp-btn').disabled = !valid;
}

function showComparison() {
    const mode = document.querySelector('.mode-btn.active').dataset.mode;
    let d1, d2, t1, t2, deg;

    if (mode === 'years') {
        const pIdx = parseInt(document.getElementById('cmp-prog').value);
        const y1 = parseInt(document.getElementById('cmp-y1').value);
        const y2 = parseInt(document.getElementById('cmp-y2').value);
        const p = programs[pIdx];
        d1 = p.years[y1]; d2 = p.years[y2];
        t1 = fmtYear(y1); t2 = fmtYear(y2);
        deg = p.degree;
    } else {
        const yr = parseInt(document.getElementById('cmp-year').value);
        const p1 = programs[parseInt(document.getElementById('cmp-p1').value)];
        const p2 = programs[parseInt(document.getElementById('cmp-p2').value)];
        d1 = p1.years[yr]; d2 = p2.years[yr];
        t1 = p1.name; t2 = p2.name;
        deg = p1.degree;
    }

    if (!d1 || !d2) return alert('لا توجد بيانات كافية');

    const kpi1 = calcKPIs(d1, deg);
    const kpi2 = calcKPIs(d2, deg);

    document.getElementById('cmp-th1').textContent = t1;
    document.getElementById('cmp-th2').textContent = t2;

    const tbody = document.getElementById('cmp-tbody');
    const chartLabels = [];
    const chartData1 = [];
    const chartData2 = [];

    tbody.innerHTML = INDICATORS
        .filter(ind => !ind.gradOnly || ['الماجستير','دكتوراه'].includes(deg))
        .map(ind => {
            const v1 = kpi1[ind.key];
            const v2 = kpi2[ind.key];
            const f1 = fmtKPI(v1, ind.unit);
            const f2 = fmtKPI(v2, ind.unit);

            let diffHtml = '<span class="diff-same">—</span>';
            if (ind.numeric && v1 != null && v2 != null) {
                const diff = parseFloat(v1) - parseFloat(v2);
                if (diff > 0) diffHtml = `<span class="diff-up">+${diff.toFixed(2)}</span>`;
                else if (diff < 0) diffHtml = `<span class="diff-down">${diff.toFixed(2)}</span>`;
                else diffHtml = `<span class="diff-same">0</span>`;

                chartLabels.push(ind.name.length > 20 ? ind.name.slice(0,20) + '..' : ind.name);
                chartData1.push(parseFloat(v1));
                chartData2.push(parseFloat(v2));
            }

            return `<tr>
                <td>${ind.name}</td>
                <td>${f1.text}</td>
                <td>${f2.text}</td>
                <td>${diffHtml}</td>
            </tr>`;
        }).join('');

    // Compare chart
    destroyChart('compare');
    if (chartLabels.length > 0) {
        const ctx = document.getElementById('chart-compare').getContext('2d');
        charts.compare = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: chartLabels,
                datasets: [
                    { label: t1, data: chartData1, backgroundColor: '#0d8e8e', borderRadius: 4 },
                    { label: t2, data: chartData2, backgroundColor: '#c9a227', borderRadius: 4 },
                ]
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

    // Store for export
    currentProg = { cmp: true, d1, d2, t1, t2, deg, kpi1, kpi2 };
    document.getElementById('cmp-results').classList.remove('hidden');
    document.getElementById('cmp-results').scrollIntoView({ behavior: 'smooth' });
}

// ========================================
// التصدير
// ========================================
function exportPDF() {
    if (!currentProg || currentProg.cmp) return;
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p','mm','a4');
    const d = currentProg.data;
    const kpi = calcKPIs(d, currentProg.prog.degree);

    doc.setFontSize(16);
    doc.text('KPI - ' + currentProg.prog.name + ' (' + currentProg.prog.degree + ')', 105, 20, { align: 'center' });
    doc.setFontSize(12);
    doc.text(fmtYear(currentProg.year), 105, 28, { align: 'center' });

    const body = INDICATORS
        .filter(ind => !ind.gradOnly || ['الماجستير','دكتوراه'].includes(currentProg.prog.degree))
        .map(ind => {
            const f = fmtKPI(kpi[ind.key], ind.unit);
            return [ind.name, f.text, ind.unit];
        });

    doc.autoTable({
        startY: 35,
        head: [['Indicator','Value','Unit']],
        body,
        styles: { halign: 'center', fontSize: 9 },
        headStyles: { fillColor: [13,110,110] }
    });
    doc.save(`KPI-${currentProg.prog.name}-${fmtYear(currentProg.year)}.pdf`);
}

function exportExcel() {
    if (!currentProg || currentProg.cmp) return;
    const d = currentProg.data;
    const kpi = calcKPIs(d, currentProg.prog.degree);
    const rows = [
        ['تقرير مؤشرات الأداء'],
        ['البرنامج: ' + currentProg.prog.name],
        ['الدرجة: ' + currentProg.prog.degree],
        ['السنة: ' + fmtYear(currentProg.year)],
        [],
        ['المؤشر','القيمة','الوحدة'],
        ...INDICATORS
            .filter(ind => !ind.gradOnly || ['الماجستير','دكتوراه'].includes(currentProg.prog.degree))
            .map(ind => {
                const f = fmtKPI(kpi[ind.key], ind.unit);
                return [ind.name, f.text, ind.unit];
            })
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المؤشرات');
    XLSX.writeFile(wb, `مؤشرات-${currentProg.prog.name}-${fmtYear(currentProg.year)}.xlsx`);
}

function exportComparePDF() {
    if (!currentProg || !currentProg.cmp) return;
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p','mm','a4');
    doc.setFontSize(14);
    doc.text('KPI Comparison', 105, 20, { align: 'center' });

    const body = INDICATORS
        .filter(ind => !ind.gradOnly || ['الماجستير','دكتوراه'].includes(currentProg.deg))
        .map(ind => {
            const f1 = fmtKPI(currentProg.kpi1[ind.key], ind.unit);
            const f2 = fmtKPI(currentProg.kpi2[ind.key], ind.unit);
            return [ind.name, f1.text, f2.text];
        });

    doc.autoTable({
        startY: 30,
        head: [['Indicator', currentProg.t1, currentProg.t2]],
        body,
        styles: { halign: 'center', fontSize: 9 },
        headStyles: { fillColor: [13,110,110] }
    });
    doc.save('KPI-Comparison.pdf');
}

function exportCompareExcel() {
    if (!currentProg || !currentProg.cmp) return;
    const rows = [
        ['مقارنة المؤشرات'],
        [],
        ['المؤشر', currentProg.t1, currentProg.t2],
        ...INDICATORS
            .filter(ind => !ind.gradOnly || ['الماجستير','دكتوراه'].includes(currentProg.deg))
            .map(ind => {
                const f1 = fmtKPI(currentProg.kpi1[ind.key], ind.unit);
                const f2 = fmtKPI(currentProg.kpi2[ind.key], ind.unit);
                return [ind.name, f1.text, f2.text];
            })
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المقارنة');
    XLSX.writeFile(wb, 'مقارنة-المؤشرات.xlsx');
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
