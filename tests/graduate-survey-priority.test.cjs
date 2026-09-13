const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const appSource = fs.readFileSync(require('node:path').join(__dirname, '..', 'js', 'app.js'), 'utf8');
const context = {
    console,
    URL,
    window: { KPI_CONFIG: {} },
    document: { addEventListener() {} },
    sessionStorage: { getItem() { return null; } },
};

vm.runInNewContext(`${appSource}\n;globalThis.__surveyTestHooks = {
    applyGraduateSurveyMetrics,
    aggregateGraduateSurveyRows,
    getSurveyEvidence,
    buildComparisonModel,
    calcKPIs,
    markAnalyticsSampleMetrics,
};`, context);

const {
    applyGraduateSurveyMetrics,
    aggregateGraduateSurveyRows,
    getSurveyEvidence,
    buildComparisonModel,
    calcKPIs,
    markAnalyticsSampleMetrics,
} = context.__surveyTestHooks;

const bachelorMetrics = {
    '45|الأنظمة|بكالوريوس': {
        eval_experience: 3.2,
        eval_experience_sample: 7,
        performance_rate: 81.2,
        performance_rate_sample: 6,
        employment_rate: 42.9,
        employment_rate_sample: 7,
        employment_employed_count: 3,
        eval_employers: 4.75,
        eval_employers_sample: 3,
    },
};

const bachelorRow = {
    Dept_aName: 'الأنظمة',
    Major_aName: 'الأنظمة',
    Degree_aName: 'بكالوريوس',
    Semester: 45,
    eval_experience: 4.6,
    eval_experience_source: 'shari3ah_surveys_program_eval',
    eval_experience_sample: 300,
    performance_rate: null,
    employment_rate: null,
    eval_employers: null,
};

applyGraduateSurveyMetrics([bachelorRow], bachelorMetrics, new Set(['بكالوريوس']), 'bachelor');
assert.equal(bachelorRow.eval_experience, 4.6, 'graduate sample must not replace program experience');
assert.equal(bachelorRow.performance_rate, 81.2);
assert.equal(bachelorRow.employment_rate, 42.9);
assert.equal(bachelorRow.eval_employers, 4.75);
assert.equal(bachelorRow.performance_rate_source, 'graduates_sample_program_bachelor');
assert.equal(getSurveyEvidence(bachelorRow, 'experience_eval'), null);
assert.equal(getSurveyEvidence(bachelorRow, 'student_performance').kind, 'sample');

const preservedBachelorRow = {
    Dept_aName: 'الأنظمة',
    Major_aName: 'الأنظمة',
    Degree_aName: 'بكالوريوس',
    Semester: 45,
    performance_rate: 91,
};
applyGraduateSurveyMetrics([preservedBachelorRow], bachelorMetrics, new Set(['بكالوريوس']), 'bachelor');
assert.equal(preservedBachelorRow.performance_rate, 91, 'graduate sample must not replace an existing KPI value');
assert.equal(preservedBachelorRow.performance_rate_source, undefined);

const postgradMetrics = {
    '46|الفقه|الماجستير': {
        eval_experience: 4.25,
        eval_experience_sample: 8,
        eval_supervision: 4.25,
        eval_supervision_sample: 8,
        eval_services: 4.12,
        eval_services_sample: 8,
        employment_rate: 37.5,
        employment_rate_sample: 8,
        eval_employers: 4.75,
        eval_employers_sample: 2,
    },
};
const postgradRow = {
    Dept_aName: 'الشريعة',
    Major_aName: 'الفقه',
    Degree_aName: 'الماجستير',
    Semester: 46,
    eval_experience: null,
    eval_supervision: null,
    eval_services: null,
    employment_rate: null,
    eval_employers: null,
};
applyGraduateSurveyMetrics([postgradRow], postgradMetrics, new Set(['الماجستير']), 'postgrad');
assert.equal(postgradRow.eval_experience, null, 'graduate sample is not a source for KPI-PG-1');
assert.equal(postgradRow.employment_rate, null, 'employment is not a postgraduate KPI in the supplied definition');
assert.equal(postgradRow.eval_supervision, 4.25);
assert.equal(postgradRow.eval_services, 4.12);
assert.equal(postgradRow.eval_employers, 4.75);
assert.equal(getSurveyEvidence(postgradRow, 'supervision_eval').kind, 'sample');

const separatedPostgrad = aggregateGraduateSurveyRows([
    {
        'ما اسم البرنامج ': 'ماجستير الفقه',
        'سنة التخرج من البرنامج الأكاديمي': '١٤٤٦ هـ',
        'ما تقييمك العام لجودة الإشراف أثناء الرسالة أو المشروع البحثي': 4,
    },
    {
        'ما اسم البرنامج ': 'دكتوراه الفقه',
        'سنة التخرج من البرنامج الأكاديمي': '١٤٤٦ هـ',
        'ما تقييمك العام لجودة الإشراف أثناء الرسالة أو المشروع البحثي': 2,
    },
], new Set(['الماجستير', 'دكتوراه']));
assert.equal(separatedPostgrad.metricsByKey['46|الفقه|الماجستير'].eval_supervision, 4);
assert.equal(separatedPostgrad.metricsByKey['46|الفقه|دكتوراه'].eval_supervision, 2);

const comparison = buildComparisonModel([
    { label: 'الأنظمة - 1445', program: { degree: 'بكالوريوس' }, data: bachelorRow, kpi: calcKPIs(bachelorRow, 'بكالوريوس') },
    { label: 'الأنظمة - 1446', program: { degree: 'بكالوريوس' }, data: preservedBachelorRow, kpi: calcKPIs(preservedBachelorRow, 'بكالوريوس') },
]);
const performanceComparison = comparison.rows.find(row => row.indicator.key === 'student_performance');
assert.equal(performanceComparison.surveyEvidence[0].kind, 'sample');
assert.equal(performanceComparison.surveyEvidence[1], null);

const analyticsMetrics = markAnalyticsSampleMetrics([
    { id: 'performance', label: 'متوسط مستوى أداء الطالب', sampleSourceKey: 'performance_rate_source' },
    { id: 'experience', label: 'متوسط جودة خبرات التعلم', sampleSourceKey: 'eval_experience_source' },
], [bachelorRow]);
assert.equal(analyticsMetrics[0].label, 'متوسط مستوى أداء الطالب - استطلاع عينة');
assert.equal(analyticsMetrics[1].label, 'متوسط جودة خبرات التعلم');

console.log('graduate survey priority tests passed');
