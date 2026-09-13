const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const appSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');
const context = {
    console,
    URL,
    window: { KPI_CONFIG: {} },
    document: { addEventListener() {} },
    sessionStorage: { getItem() { return null; } },
};

vm.runInNewContext(`${appSource}\n;globalThis.__surveyTestHooks = {
    applyGraduateSamplesFromShari3ahSurveys,
    getSurveyEvidence,
    buildComparisonModel,
    calcKPIs,
    markAnalyticsSampleMetrics,
    extractCourseEvaluationMetricsFromShari3ahSurveys,
    applyStatisticalKpiEstimates,
    buildProgramDisplayData,
    buildProgramExportReport,
    getStatisticalEvidence,
    setProgramsForTest: value => { programs = value; },
};`, context);

const {
    applyGraduateSamplesFromShari3ahSurveys,
    getSurveyEvidence,
    buildComparisonModel,
    calcKPIs,
    markAnalyticsSampleMetrics,
    extractCourseEvaluationMetricsFromShari3ahSurveys,
    applyStatisticalKpiEstimates,
    buildProgramDisplayData,
    buildProgramExportReport,
    getStatisticalEvidence,
    setProgramsForTest,
} = context.__surveyTestHooks;

const trendProgram = {
    name: 'برنامج تجريبي',
    degree: 'بكالوريوس',
    years: {
        45: { eval_courses: 3.5 },
        46: { eval_courses: 3.7 },
        47: { eval_courses: null },
    },
};
const trendInfo = applyStatisticalKpiEstimates([trendProgram]);
assert.equal(trendInfo.targetYear, '1447');
assert.equal(trendProgram.years[47].statistical_estimates.course_eval.displayValue, 3.75);
assert.equal(getStatisticalEvidence(trendProgram.years[47], 'course_eval').label, 'إحصائي');

setProgramsForTest([trendProgram]);
const scopedTrendData = buildProgramDisplayData(trendProgram, 47, 'الحوية');
const scopedTrendKpis = calcKPIs(scopedTrendData, trendProgram.degree);
assert.equal(scopedTrendKpis.course_eval, 3.75);
assert.equal(scopedTrendKpis.student_faculty_ratio, null);
assert.equal(getStatisticalEvidence(scopedTrendData, 'student_faculty_ratio'), null);

const multiYearExport = buildProgramExportReport([trendProgram], [45, 47], [0]);
assert.deepEqual([...multiYearExport.years], [45, 47]);
assert.equal(multiYearExport.selectedPrograms.length, 1);
assert.equal(multiYearExport.detailRecords.length, 2);
assert.equal(multiYearExport.indicatorRecords.length, 22);
assert.equal(
    multiYearExport.indicatorRecords.find(record => record.detail.year === '1447' && record.code === 'KPI-2').status,
    'إحصائي'
);
assert.equal(
    multiYearExport.indicatorRecords.find(record => record.detail.year === '1445' && record.code === 'KPI-8').status,
    'غير متوفر'
);

const courseMetrics = extractCourseEvaluationMetricsFromShari3ahSurveys({ courseRecords: [
    { year: '1447', program: 'ماجستير القانون', degree: 'الماجستير', courseName: 'مشروع بحثي', score: 4.5, respondents: 8 },
    { year: '1447', program: 'ماجستير القانون', degree: 'الماجستير', courseName: 'الرسالة', score: 4.0, respondents: 2 },
    { year: '1447', program: 'ماجستير القانون', degree: 'الماجستير', courseName: 'بحث التخرج', score: 2.0, respondents: 5, researchCourse: true },
] });
assert.equal(courseMetrics['1447|القانون|الماجستير'].eval_supervision, 4.25);
assert.equal(courseMetrics['1447|القانون|الماجستير'].eval_supervision_sample, 10);
assert.equal(courseMetrics['1447|القانون|الماجستير'].supervision_course_count, 2);

const payload = {
    extractedData: {},
    graduateSampleKpis: {
        'p01::1445': {
            metrics: {
                performance_rate: { value: 81.2, sampleCount: 6 },
                employment_rate: { value: 42.9, sampleCount: 7, positiveCount: 3 },
                eval_employers: { value: 4.75, sampleCount: 3 },
            },
        },
        'p09::1446': {
            metrics: {
                eval_supervision: { value: 4.25, sampleCount: 8 },
                eval_services: { value: 4.12, sampleCount: 8 },
                eval_employers: { value: 4.75, sampleCount: 2 },
                employment_rate: { value: 37.5, sampleCount: 8 },
            },
        },
    },
};

const bachelorRow = {
    Dept_aName: 'الأنظمة',
    Major_aName: 'الأنظمة',
    Degree_aName: 'بكالوريوس',
    Semester: 45,
    eval_experience: 4.6,
    eval_experience_source: 'shari3ah_surveys_program_eval',
    performance_rate: null,
    employment_rate: null,
    eval_employers: null,
};

const info = applyGraduateSamplesFromShari3ahSurveys([bachelorRow], payload);
assert.equal(info.applied, true);
assert.equal(bachelorRow.eval_experience, 4.6, 'graduate sample must not replace program experience');
assert.equal(bachelorRow.performance_rate, 81.2);
assert.equal(bachelorRow.employment_rate, 42.9);
assert.equal(bachelorRow.employment_employed_count, 3);
assert.equal(bachelorRow.eval_employers, 4.75);
assert.equal(bachelorRow.performance_rate_source, 'graduates_sample_program_shari3ah_surveys');
assert.equal(getSurveyEvidence(bachelorRow, 'experience_eval'), null);
assert.equal(getSurveyEvidence(bachelorRow, 'student_performance').kind, 'sample');

const preservedBachelorRow = {
    Dept_aName: 'الأنظمة',
    Major_aName: 'الأنظمة',
    Degree_aName: 'بكالوريوس',
    Semester: 45,
    performance_rate: 91,
};
applyGraduateSamplesFromShari3ahSurveys([preservedBachelorRow], payload);
assert.equal(preservedBachelorRow.performance_rate, 91, 'graduate sample must not replace an existing KPI value');
assert.equal(preservedBachelorRow.performance_rate_source, undefined);

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
applyGraduateSamplesFromShari3ahSurveys([postgradRow], payload);
assert.equal(postgradRow.eval_experience, null, 'graduate sample is not a source for KPI-PG-1');
assert.equal(postgradRow.employment_rate, null, 'employment is not a postgraduate KPI in the supplied definition');
assert.equal(postgradRow.eval_supervision, 4.25);
assert.equal(postgradRow.eval_services, 4.12);
assert.equal(postgradRow.eval_employers, 4.75);
assert.equal(getSurveyEvidence(postgradRow, 'supervision_eval').kind, 'sample');

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
