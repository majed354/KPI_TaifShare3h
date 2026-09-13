const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const bachelorUrl = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTFZxQeVsbWxxIdmHOugA-NYiGVZOmKL44KquMVUaJJaJRYo0O8XjrMbCETrw0yY9y49aucRSrnXUCy/pub?gid=1084828379&single=true&output=csv';
const postgradUrl = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vScNn4rvHcKi2EB6mSZ31GsqSb4VV7LFnKnUgz5qsu-U8823iHG6gIPmKZgSMofqvsZNc5QeBXWem9C/pub?gid=520745508&single=true&output=csv';
const programSurveyUrl = 'https://raw.githubusercontent.com/majed354/Shari3ahSurveys/main/js/surveys-data.js';

const appSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');
const context = {
    console,
    fetch,
    URL,
    window: {
        KPI_CONFIG: {
            graduatesSurveyBachelorSheetUrl: bachelorUrl,
            graduatesSurveyPostgradSheetUrl: postgradUrl,
            shari3ahSurveysDataUrl: programSurveyUrl,
        },
    },
    document: { addEventListener() {} },
    sessionStorage: { getItem() { return null; } },
};

vm.runInNewContext(`${appSource}\n;globalThis.__liveSurveyTestHooks = {
    parseCSV,
    applyProgramExperienceFromShari3ahSurveys,
    applyGraduateSurveyIndicators,
    getSurveyEvidence,
};`, context);

async function main() {
    const hooks = context.__liveSurveyTestHooks;
    const csv = fs.readFileSync(path.join(__dirname, '..', 'data', 'data.csv'), 'utf8');
    const rows = hooks.parseCSV(csv);

    const programSurveyInfo = await hooks.applyProgramExperienceFromShari3ahSurveys(rows);
    const sampleInfo = await hooks.applyGraduateSurveyIndicators(rows);
    assert.equal(programSurveyInfo.applied, true, 'program survey should be available');
    assert.equal(sampleInfo.applied, true, 'graduate sample surveys should be available');

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

    assert.ok(bachelorSampleRows.length > 0, 'bachelor sample KPIs should be populated');
    assert.ok(postgradSampleRows.length > 0, 'postgraduate sample KPIs should be populated');
    assert.ok(programSurveyRows.length > 0, 'program experience should come from the existing source');
    assert.equal(rows.some(row => String(row.eval_experience_source || '').startsWith('graduates_sample_')), false);
    assert.equal(rows.some(row => String(row.eval_courses_source || '').startsWith('graduates_sample_')), false);
    assert.equal(rows.some(row => String(row.employment_rate_source || '').includes('postgrad')), false);

    assert.equal(hooks.getSurveyEvidence(programSurveyRows[0], 'experience_eval'), null);
    assert.equal(
        bachelorSampleRows.some(row => hooks.getSurveyEvidence(row, 'student_performance')?.kind === 'sample'),
        true
    );

    console.log(JSON.stringify({
        programSurveyRows: programSurveyRows.length,
        bachelorSampleRows: bachelorSampleRows.length,
        postgradSampleRows: postgradSampleRows.length,
        liveGraduateResponses: sampleInfo.matchedRows,
    }));
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
