const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const surveyUrl = 'https://raw.githubusercontent.com/majed354/Shari3ahSurveys/main/js/surveys-data.js';
const courseUrl = 'https://raw.githubusercontent.com/majed354/Shari3ahSurveys/main/js/course-evaluations-data.js';
const appSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');
const context = {
    console,
    fetch,
    URL,
    window: {
        KPI_CONFIG: {
            shari3ahSurveysDataUrl: surveyUrl,
            shari3ahCourseEvaluationsDataUrl: courseUrl,
        },
    },
    document: { addEventListener() {} },
    sessionStorage: { getItem() { return null; } },
};

vm.runInNewContext(`${appSource}\n;globalThis.__liveSurveyTestHooks = {
    parseCSV,
    fetchShari3ahSurveysPayload,
    applyProgramExperienceFromShari3ahSurveys,
    applyGraduateSamplesFromShari3ahSurveys,
    applyCourseEvaluationsFromShari3ahSurveys,
    getSurveyEvidence,
    buildPrograms,
    applyStatisticalKpiEstimates,
    calcKPIs,
    getIndicatorsForDegree,
};`, context);

async function main() {
    const hooks = context.__liveSurveyTestHooks;
    const csv = fs.readFileSync(path.join(__dirname, '..', 'data', 'data.csv'), 'utf8');
    const rows = hooks.parseCSV(csv);
    const payload = await hooks.fetchShari3ahSurveysPayload();

    assert.ok(payload?.graduateSampleKpis, 'unified survey payload should include graduate samples');
    const programSurveyInfo = hooks.applyProgramExperienceFromShari3ahSurveys(rows, payload);
    const sampleInfo = hooks.applyGraduateSamplesFromShari3ahSurveys(rows, payload);
    const courseInfo = await hooks.applyCourseEvaluationsFromShari3ahSurveys(rows);
    assert.equal(programSurveyInfo.applied, true, 'program survey should be available');
    assert.equal(sampleInfo.applied, true, 'graduate sample surveys should be available');
    assert.equal(courseInfo.applied, true, 'course evaluations should be available');

    const bachelorSampleRows = rows.filter(row => row.Degree_aName === 'بكالوريوس' && (
        String(row.performance_rate_source || '').startsWith('graduates_sample_') ||
        String(row.employment_rate_source || '').startsWith('graduates_sample_') ||
        String(row.eval_employers_source || '').startsWith('graduates_sample_')
    ));
    const postgradSampleRows = rows.filter(row => ['الماجستير', 'دكتوراه'].includes(row.Degree_aName) && (
        String(row.eval_supervision_source || '').startsWith('graduates_sample_') ||
        String(row.eval_services_source || '').startsWith('graduates_sample_') ||
        String(row.eval_employers_source || '').startsWith('graduates_sample_')
    ));
    const programSurveyRows = rows.filter(row => row.eval_experience_source === 'shari3ah_surveys_program_eval');
    const programSurveyRows1447 = programSurveyRows.filter(row => Number(row.Semester) === 47);
    const courseSurveyRows = rows.filter(row => row.eval_courses_source === 'shari3ah_surveys_course_evaluations');
    const supervisionCourseRows1447 = rows.filter(row =>
        Number(row.Semester) === 47 && row.eval_supervision_source === 'shari3ah_surveys_course_supervision'
    );

    assert.ok(bachelorSampleRows.length > 0, 'bachelor sample KPIs should be populated');
    assert.ok(postgradSampleRows.length > 0, 'postgraduate sample KPIs should be populated');
    assert.ok(programSurveyRows.length > 0, 'program experience should come from the unified source');
    assert.equal(programSurveyRows1447.length, 15, '1447 program experience should support the average-based source format');
    assert.ok(courseSurveyRows.length > 0, 'course evaluation should come from the unified source');
    assert.equal(supervisionCourseRows1447.length, 8, '1447 supervision should use only research project and thesis courses');
    assert.equal(rows.some(row => String(row.eval_experience_source || '').startsWith('graduates_sample_')), false);
    assert.equal(rows.some(row => String(row.eval_courses_source || '').startsWith('graduates_sample_')), false);
    assert.equal(rows.some(row => String(row.employment_rate_source || '').includes('postgrad')), false);
    assert.equal(hooks.getSurveyEvidence(programSurveyRows[0], 'experience_eval'), null);
    assert.equal(
        bachelorSampleRows.some(row => hooks.getSurveyEvidence(row, 'student_performance')?.kind === 'sample'),
        true
    );

    const programs = hooks.buildPrograms(rows);
    const estimateInfo = hooks.applyStatisticalKpiEstimates(programs);
    const remainingMissing1447 = programs.flatMap(program => {
        const data = program.years[47];
        if (!data) return [];
        const values = hooks.calcKPIs(data, program.degree);
        return hooks.getIndicatorsForDegree(program.degree)
            .filter(indicator => values[indicator.key] == null)
            .map(indicator => `${program.name}|${program.degree}|${indicator.code}`);
    });
    assert.equal(estimateInfo.targetYear, '1447');
    const estimatesWithoutHistory = programs.flatMap(program => {
        const estimates = program.years[47]?.statistical_estimates || {};
        return Object.entries(estimates)
            .filter(([, estimate]) => !Array.isArray(estimate.basisYears) || estimate.basisYears.length === 0)
            .map(([key]) => `${program.name}|${program.degree}|${key}`);
    });
    assert.equal(estimatesWithoutHistory.length, 0, 'statistical estimates must have actual historical basis');
    assert.ok(remainingMissing1447.length > 0, 'KPIs without actual history should remain unavailable');

    const graduateYears = [...new Set(Object.keys(payload.graduateSampleKpis).map(key => key.split('::')[1]))].sort();
    console.log(JSON.stringify({
        programSurveyRows: programSurveyRows.length,
        courseSurveyRows: courseSurveyRows.length,
        bachelorSampleRows: bachelorSampleRows.length,
        postgradSampleRows: postgradSampleRows.length,
        graduateSampleDatasets: Object.keys(payload.graduateSampleKpis).length,
        graduateYears,
        statisticalEstimates1447: estimateInfo.estimatedValues,
        unavailableWithoutHistory1447: remainingMissing1447.length,
    }));
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
