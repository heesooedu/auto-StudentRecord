const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
const plain = value => JSON.parse(JSON.stringify(value));

// Minimal Sheets double: keep submitted values separate from displayed literal text.
class Sheet {
  constructor(rows) { this.rows = rows.map(r => [...r]); this.notes = new Map(); this.formulas = new Map(); this.writes = []; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(1, ...this.rows.map(r => r.length)); }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
  deleteRow(row) { this.rows.splice(row - 1, 1); }
  getRange(row, col, height = 1, width = 1) {
    const sh = this;
    const key = (r, c) => `${r}:${c}`;
    const range = {
      getValues: () => Array.from({ length: height }, (_, i) => Array.from({ length: width }, (_, j) => sh.rows[row + i - 1]?.[col + j - 1] ?? '')),
      getValue: () => range.getValues()[0][0],
      getNotes: () => Array.from({ length: height }, (_, i) => Array.from({ length: width }, (_, j) => sh.notes.get(key(row + i, col + j)) || '')),
      getNote: () => range.getNotes()[0][0],
      setValues(values) {
        sh.writes.push({ row, col, values: plain(values) });
        values.forEach((r, i) => r.forEach((v, j) => {
          sh.rows[row + i - 1] ||= [];
          sh.rows[row + i - 1][col + j - 1] = typeof v === 'string' && /^'[=+\-@]/.test(v) ? v.slice(1) : v;
        }));
        return range;
      },
      setValue: v => range.setValues([[v]]),
      setNote(v) { sh.notes.set(key(row, col), v); return range; },
      clearNote() { sh.notes.delete(key(row, col)); return range; },
      clearContent() { return range.setValues(Array.from({ length: height }, () => Array(width).fill(''))); },
      setFormulas(values) { values.forEach((r, i) => r.forEach((v, j) => sh.formulas.set(key(row + i, col + j), v))); return range; },
      setFormula: v => range.setFormulas([[v]]),
      setNumberFormat() { return range; },
      setWrapStrategy() { return range; },
      setVerticalAlignment() { return range; },
    };
    return range;
  }
}

function server(sheets = {}) {
  const props = new Map();
  const store = { getProperty: k => props.get(k) || null, setProperty: (k, v) => props.set(k, v), deleteProperty: k => props.delete(k) };
  const lock = { tryLock: () => true, waitLock() {}, releaseLock() {} };
  const c = vm.createContext({
    console, Date,
    PropertiesService: { getUserProperties: () => store, getDocumentProperties: () => store },
    LockService: { getScriptLock: () => lock, getDocumentLock: () => lock },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => sheets[name] }), flush() {}, WrapStrategy: { WRAP: 'WRAP' } },
    ScriptApp: { getProjectTriggers: () => [] }, Utilities: { sleep() {} },
  });
  for (const file of ['Code.js', 'ModelCatalog.js', 'Ui.js', 'UiTools.js', 'UiOcr.js']) vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });
  c.log_ = () => {};
  c.withDocumentLock_ = fn => fn();
  return { c, props };
}

const manualConfig = { startRow: 2, inputCols: [2], outputCol: 3, byteCol: 4, prompt: 'test' };
function manualServer(sh) {
  const h = server({ '수동추가': sh });
  h.c.getConfigMap_ = () => ({});
  h.c.getManualRecordConfig_ = () => manualConfig;
  return h;
}

test('manual aggregates all 2,000 rows while displaying only the first 500', () => {
  const sh = new Sheet([['번호', '활동', '생기부', '바이트'], ...Array.from({ length: 2000 }, (_, i) => [i + 1, '활동 근거가 충분히 담긴 문장입니다.'.repeat(2), '완료', ''])]);
  sh.rows[501][2] = ''; // beyond the visible 500
  sh.rows[2000][2] = '';
  sh.notes.set('2001:3', '[수동추가 오류]\nAPI 오류');
  sh.rows[502] = [502, '', '', ''];
  sh.rows[503][2] = '';
  sh.notes.set('504:3', '[수동추가 근거부족]');
  const { c } = manualServer(sh);
  let data = c.uiGetManual();
  assert.equal(data.rows.length, 500);
  assert.equal(data.truncated, true);
  assert.equal(data.pendingCount, 1);
  assert.equal(data.errorCount, 1);
  assert.equal(c.countPendingManualRowsInSheet_(sh, manualConfig), 1);
  // Resetting and editing visible rows must refresh totals including hidden rows.
  c.assertManualRow_ = () => ({ sh, config: manualConfig });
  data = c.uiResetManualOutput(2, 'unused').data;
  assert.equal(data.pendingCount, 2);
  data = c.uiUpdateManualOutput(2, 'unused', '=교사가 수정한 문구').data;
  assert.equal(data.pendingCount, 1);
  assert.equal(data.errorCount, 1);
  assert.equal(sh.writes.at(-1).values[0][0], "'=교사가 수정한 문구");
});

test('manual threshold, handled notes, retry and generation share the same eligibility', () => {
  const sh = new Sheet([['번호', '활동', '생기부', '바이트'], [1, 'a'.repeat(14), '', ''], [2, 'short', '', ''], [3, 'a'.repeat(30), '', '']]);
  sh.notes.set('4:3', '[수동추가 오류]\n실패');
  const { c } = manualServer(sh);
  const input = c.buildManualInputFromRowValues_(sh.rows[1], sh.rows[0], [2]);
  assert.equal(input.length, 19);
  assert.equal(c.countPendingManualRowsInSheet_(sh, manualConfig), 0);
  sh.rows[1][1] += 'a';
  assert.equal(c.uiGetManual().pendingCount, 1);
  assert.equal(c.countPendingManualRowsInSheet_(sh, manualConfig), 1);
  c.requireCurrentAiApiKey_ = () => {};
  c.runWithCollectedUi_ = () => ({ messages: [] });
  c.uiGetStatus = () => ({});
  c.uiStartManual({ retryErrors: true });
  assert.equal(c.uiGetManual().pendingCount, 2);
  c.ensureManualOutputColumns_ = () => {};
  c.getAiProviderAndModel_ = () => ({ provider: 'gpt', model: 'test', reasoning: 'none' });
  c.getSystemGuide_ = () => 'guide';
  c.isStopRequestedDuringRun_ = () => false;
  c.createAiPrefetcher_ = () => ({ prefetch() {}, finish() {} });
  c.createRecordDraftWithLength_ = () => ({ parsed: { record_draft: '=활동 결과' } });
  const result = c.processManualAddedRecordsCore_();
  assert.equal(result.processed, 2);
  assert.equal(result.remaining, 0);
  assert.equal(c.uiGetManual().pendingCount, 0);
  assert.equal(sh.rows[1][2], '=활동 결과.');
  assert.equal(c.manualRowStatus_('a'.repeat(30), '', '[수동추가 근거부족]'), '근거부족');
});

test('behavior byte migration is repeatable, preserves current values and ambiguous CHARS', () => {
  const sh = new Sheet([['key', 'value', 'description'],
    ['BEHAVIOR_RECORD_BYTE_MIN', '800', 'old'], ['BEHAVIOR_RECORD_BYTE_MAX', '1,300', 'old'],
    ['BEHAVIOR_RECORD_MIN_CHARS', '250', 'ambiguous'], ['BEHAVIOR_RECORD_MAX_CHARS', '400', 'ambiguous']]);
  const { c } = server({ '설정': sh });
  c.migrateLegacyConfigKeys_(sh);
  const first = plain(sh.rows);
  c.migrateLegacyConfigKeys_(sh);
  assert.deepEqual(sh.rows, first);
  assert.equal(sh.rows.find(r => r[0] === 'BEHAVIOR_RECORD_MIN_BYTES')[1], '800');
  assert.equal(sh.rows.find(r => r[0] === 'BEHAVIOR_RECORD_MAX_BYTES')[1], '1,300');
  assert.equal(sh.rows.find(r => r[0] === 'BEHAVIOR_RECORD_MIN_CHARS')[1], '250');
  sh.rows.push(['BEHAVIOR_RECORD_BYTE_MIN', '999', 'conflict'], ['BEHAVIOR_RECORD_BYTE_MAX', '1300', 'duplicate']);
  c.migrateLegacyConfigKeys_(sh);
  assert.equal(sh.rows.find(r => r[0] === 'BEHAVIOR_RECORD_MIN_BYTES')[1], '800');
  assert.ok(sh.rows.find(r => r[0] === 'BEHAVIOR_RECORD_BYTE_MIN')); // preserve conflicting original
  assert.ok(!sh.rows.find(r => r[0] === 'BEHAVIOR_RECORD_BYTE_MAX'));
  c.uiGetModelCatalog = () => ({ models: [] });
  assert.match(c.uiGetSettings().behaviorMigrationNotice, /250.*400.*800~1300/);
  const charsOnly = new Sheet([['key', 'value'], ['BEHAVIOR_RECORD_MIN_CHARS', '200']]);
  c.migrateLegacyConfigKeys_(charsOnly);
  assert.deepEqual(charsOnly.rows, [['key', 'value'], ['BEHAVIOR_RECORD_MIN_CHARS', '200']]);
  const emptyCurrent = new Sheet([['key', 'value'], ['BEHAVIOR_RECORD_MIN_BYTES', ''], ['BEHAVIOR_RECORD_BYTE_MIN', '850']]);
  c.migrateLegacyConfigKeys_(emptyCurrent);
  assert.equal(emptyCurrent.rows[1][1], '850');
  assert.equal(emptyCurrent.rows.length, 2);
});

test('Classroom caches complete pagination and profiles only within one execution; failures retry', () => {
  let lists = 0, profiles = 0;
  const { c } = server();
  c.Classroom = { Courses: { list: args => { lists++; return args.pageToken ? { courses: [{ id: 2 }] } : { courses: [{ id: 1 }], nextPageToken: 'next' }; } }, UserProfiles: { get: id => { profiles++; return { id }; } } };
  assert.equal(c.isMyTeacherCourse_(1), true);
  assert.equal(c.isMyTeacherCourse_(2), true);
  assert.equal(c.isMyTeacherCourse_(3), false);
  assert.equal(lists, 2);
  c.getUserProfile_(10); c.getUserProfile_('10'); assert.equal(profiles, 1);
  c.Classroom.UserProfiles.get = () => { profiles++; throw Error('temporary'); };
  assert.throws(() => c.getUserProfile_(11)); assert.throws(() => c.getUserProfile_(11)); assert.equal(profiles, 3);
  const fresh = server().c;
  fresh.Classroom = c.Classroom;
  fresh.isMyTeacherCourse_(2); assert.equal(lists, 4);
  const failed = server().c;
  let attempts = 0;
  failed.Classroom = { Courses: { list: args => { attempts++; if (attempts === 2) throw Error('page failure'); return args.pageToken ? { courses: [{ id: 2 }] } : { courses: [{ id: 1 }], nextPageToken: 'next' }; } } };
  assert.throws(() => failed.isMyTeacherCourse_(1));
  assert.equal(failed.isMyTeacherCourse_(2), true);
  assert.equal(attempts, 4);
});

function configureJobs(c) {
  c.getUi_ = () => ({ alert() {} });
  c.requireCurrentAiApiKey_ = () => {};
  c.ensureBehaviorRecommendationSheet_ = () => ({ getLastRow: () => 2 });
  c.styleBehaviorRecommendationSheet_ = () => {};
  c.headerMap_ = () => ({});
  c.getConfigMap_ = () => ({});
  c.countPendingBehaviorRecommendationRows_ = () => 2;
  c.countPendingManualAddedRows_ = () => 3;
  c.beginAutoRun_ = () => {};
  c.endAutoRun_ = () => {};
  c.logBehaviorRecommendationResult_ = () => {};
  c.toastAutoRecordDraft_ = () => {};
  for (const kind of ['BehaviorRecommendation', 'ManualRecord']) {
    c[`scheduleNextAuto${kind}Trigger_`] = () => {};
    c[`deleteAuto${kind}Triggers_`] = () => {};
  }
}

test('behavior and manual first/subsequent batches persist blocked, progress, completion and exception results', () => {
  for (const kind of ['BEHAVIOR_RECOMMENDATION', 'MANUAL_RECORD']) {
    for (const auto of [false, true]) {
      const sh = new Sheet([['status'], ['대기']]);
      const { c, props } = server({ '행발문구추천': sh });
      configureJobs(c);
      const behavior = kind === 'BEHAVIOR_RECOMMENDATION';
      const process = behavior ? 'processBehaviorRecommendationRows_' : 'processManualAddedRecordsCore_';
      const run = behavior ? (auto ? 'autoProcessPendingBehaviorRecommendations' : 'generateBehaviorRecommendationsCore_') : (auto ? 'autoProcessPendingManualAddedRecords' : 'generateManualAddedRecords');
      for (const expected of ['BLOCKED', 'RUNNING', 'DONE', 'EXCEPTION']) {
        props.set(`AUTO_${kind}_TRIGGER`, 'ON');
        c[process] = () => {
          if (expected === 'EXCEPTION') throw Error('exception detail');
          if (!auto) assert.equal(c.getAutoToolLastResult_(kind).blocked, false);
          return { processed: 1, failed: 0, skipped: 0, remaining: expected === 'DONE' ? 0 : 2, blocked: expected === 'BLOCKED', blockingReason: 'quota', blockingMessage: 'quota detail' };
        };
        c[run]();
        const last = c.getAutoToolLastResult_(kind);
        assert.equal(last.status, expected === 'EXCEPTION' ? 'BLOCKED' : expected, `${run}: ${expected}`);
        assert.equal(last.blocked, expected === 'BLOCKED' || expected === 'EXCEPTION');
        if (expected === 'EXCEPTION') assert.equal(last.blockingMessage, 'exception detail');
      }
      for (const name of ['SubmissionCollections', 'RecordDrafts', 'StudentFinalRecords', 'CommonPhrases', 'Grades']) c[`countPending${name}_`] = () => 0;
      assert.equal(c.uiGetStatus()[behavior ? 'behavior' : 'manual'].last.blocked, true);
      props.set(`AUTO_${kind}_LAST_RESULT`, 'broken json');
      assert.equal(c.getAutoToolLastResult_(kind), null);
    }
  }
});

test('OCR extraction/edit and behavior input/output escape literal text while keeping byte formulas', () => {
  const sh = new Sheet([['쪽', '학번', '이름', '답안', '확인할 점', '모델', '추출 시각'], ['1', '', '', '', '', '', '']]);
  const { c } = server();
  c.getOcrSheet_ = () => sh;
  c.ocrSettings_ = () => ({ config: {}, provider: 'gpt', model: 'test', reasoning: 'none' });
  c.requireOcrModel_ = () => 'test'; c.ocrPrompt_ = () => 'guide'; c.ocrItemsFresh_ = () => true;
  c.ocrItems_ = () => [{ name: '답안' }]; c.ocrLayoutMatches_ = () => true;
  c.callAi_ = () => ({ text: JSON.stringify({ student_id: '001', name: '=이름', answers: [{ item: 1, text: '=1+1' }], unreadable: '+확인' }) });
  let row = c.uiOcrExtract({ pages: '1', data: 'AAAA', model: 'test' });
  assert.equal(row.answers['답안'], '=1+1');
  assert.equal(sh.writes.at(-1).values[0][3], "'=1+1");
  assert.equal(sh.writes.at(-1).values[0][1], '001');
  assert.match(sh.writes.at(-1).values[0][6], /^\d{4}-/); // Date stays a Date before transport
  row = c.uiOcrUpdateRow(2, '1', { studentId: '002', name: '@이름', answers: { '답안': '- 첫째' } });
  assert.equal(row.answers['답안'], '- 첫째');
  assert.equal(sh.writes.at(-1).values[0][2], "'- 첫째");
  const behavior = new Sheet([['studentName', 'customTraits', 'recommendedRecord', 'byteCount', 'status'], ['교사', '', '', '', '']]);
  const h = { studentName: 1, customTraits: 2, recommendedRecord: 3, byteCount: 4, status: 5 };
  c.getSheet_ = () => behavior; c.assertBehaviorRow_ = () => h; c.readBehaviorRow_ = () => ({});
  c.uiSaveBehaviorRow(2, '교사', { customTraits: '=관찰' });
  assert.equal(behavior.writes.at(-1).values[0][0], "'=관찰");
  c.uiUpdateBehaviorRecord(2, '교사', '+추천');
  assert.equal(behavior.writes.at(-1).values[0][0], "'+추천");
  c.writeBehaviorRecommendationResult_(behavior, h, 2, '=추천', '생성완료', '');
  assert.equal(behavior.writes.at(-1).values[0][0], "'=추천");
  assert.equal(behavior.formulas.get('2:4'), '=IF($C2="","",2*LENB($C2)-LEN($C2))');
});

test('OCR split server uses overlap as context and returns only new absolute pages', () => {
  const { c } = server();
  c.ocrSettings_ = () => ({ provider: 'gpt', model: 'test' });
  c.requireOcrModel_ = () => 'test';
  c.callAi_ = (_p, _m, guide) => {
    assert.match(guide, /참고 문맥/);
    return { text: JSON.stringify({ pages: [{ page: 1, new_student: true }, { page: 2, new_student: false }, { page: 3, new_student: true }] }) };
  };
  const pages = c.uiOcrDetectStarts({ firstPage: 4, count: 3, contextPages: 1, data: 'AAAA' });
  assert.deepEqual(plain(pages.map(p => [p.page, p.newStudent])), [[5, false], [6, true]]);
  assert.throws(() => c.uiOcrDetectStarts({ firstPage: 4, count: 1, contextPages: 1, data: 'AAAA' }));
});

test('manual import, append and single-row input preserve formula-like text through actual write paths', () => {
  const sh = new Sheet([]);
  for (const method of ['setFrozenRows', 'setColumnWidth']) sh[method] = () => {};
  sh.getMaxColumns = () => 26; sh.getMaxRows = () => 1000;
  const { c } = manualServer(sh);
  let exists = false;
  c.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: () => exists ? sh : null, getNumSheets: () => 0,
    insertSheet: () => { exists = true; return sh; },
  });
  c.hideNewWorkSheet_ = () => {};
  c.writeConfigValues_ = () => {};
  // These style-only methods are irrelevant to the storage contract under test.
  const getRange = sh.getRange.bind(sh);
  sh.getRange = (...args) => {
    const range = getRange(...args);
    range.setFontWeight = range.setBackground = () => range;
    range.getDisplayValues = range.getValues;
    return range;
  };
  c.uiImportManualCsv({ headers: ['이름', '활동'], rows: [['001', '=가져온 활동 근거가 충분한 문장입니다.']], inputCols: [2] });
  assert.equal(sh.rows[1][0], '001');
  assert.equal(sh.writes.find(w => w.row === 2).values[0][1], "'=가져온 활동 근거가 충분한 문장입니다.");
  c.uiAddManualRow({ 1: '002', 2: '+직접 추가한 활동 근거가 충분한 문장입니다.' });
  assert.equal(sh.writes.at(-1).values[0][1], "'+직접 추가한 활동 근거가 충분한 문장입니다.");
  c.uiAppendManualRows({ headers: ['이름', '활동'], rows: [['003', '@설문 응답의 활동 근거가 충분한 문장입니다.']] });
  assert.equal(sh.writes.at(-1).values[0][1], "'@설문 응답의 활동 근거가 충분한 문장입니다.");
  assert.equal(c.uiGetManual().pendingCount, 3);
});

test('record editing restores a byte formula, including empty text, without touching adjacent cells', () => {
  const sh = new Sheet([['key', 'finalRecord', 'charCount', 'teacherChecked'], ['student-1', '기존 문구', 999, true]]);
  const { c } = server();
  c.getSheet_ = () => sh;
  c.updateRecordRow_('학생별생기부', 2, 'student-1', { finalRecord: '=교사가 수정' }, r => r.key);
  assert.equal(sh.rows[1][1], '=교사가 수정');
  assert.equal(sh.writes.at(-1).values[0][0], "'=교사가 수정");
  assert.equal(sh.formulas.get('2:3'), '=IF($B2="","",2*LENB($B2)-LEN($B2))');
  assert.equal(sh.rows[1][3], true);
  c.updateRecordRow_('학생별생기부', 2, 'student-1', { finalRecord: '' }, r => r.key);
  assert.equal(sh.rows[1][1], '');
  assert.equal(sh.formulas.get('2:3'), '=IF($B2="","",2*LENB($B2)-LEN($B2))');
  c.applyNeisByteFormulas_(sh, 2, 3, 2, 1);
  assert.equal(sh.rows[1][1], '');
  assert.equal(sh.rows[1][3], true);
});

function client() {
  const c = vm.createContext({
    console, Date, setTimeout, clearTimeout, setInterval, clearInterval,
    document: { addEventListener() {}, querySelectorAll: () => [], querySelector: () => ({ addEventListener() {} }) },
    window: { addEventListener() {} },
  });
  for (const file of ['AppScript.html', 'AppTools.html']) {
    const src = fs.readFileSync(path.join(root, file), 'utf8').replace(/^<script>/, '').replace(/<\/script>\s*$/, '').replace(/^  startApp\(\);$/m, '');
    vm.runInContext(src, c, { filename: file });
  }
  c.render = () => {}; c.banner = () => {}; c.showError = () => {}; c.isEditing = () => false;
  return c;
}

test('OCR prefetch is reused, stale entry refreshes once and invalidated responses cannot overwrite new data', async () => {
  const c = client();
  const resolves = [];
  c.call = () => new Promise(resolve => resolves.push(resolve));
  const prefetch = c.loadTool('ocr');
  const enter = c.loadFreshTool('ocr');
  assert.equal(resolves.length, 1);
  resolves[0]({ version: 1 }); await Promise.all([prefetch, enter]);
  await c.loadFreshTool('ocr'); assert.equal(resolves.length, 1);
  vm.runInContext('toolLoadedAt.ocr = Date.now() - 6000', c);
  const stale = c.loadFreshTool('ocr'), another = c.loadFreshTool('ocr');
  assert.equal(resolves.length, 2);
  c.dropTool('ocr');
  const fresh = c.loadFreshTool('ocr');
  resolves[2]({ version: 3 }); await fresh;
  resolves[1]({ version: 2 }); await Promise.all([stale, another]);
  assert.equal(vm.runInContext('S.ocr.version', c), 3);
});

test('OCR client overlaps chunks without changing earlier decisions, progress or failed-page boundaries', async () => {
  for (const scenario of ['success', 'failure', 'stop']) {
    const c = client();
    vm.runInContext("S.ocrPdf={pageCount:9}; S.ocr={model:'test'}; S.view='ocr'; S.ocrStarts=[1,3]", c);
    c.confirmModal = async () => true;
    const pieces = [], requests = [];
    c.ocrPiece = async chunk => { pieces.push(plain(chunk)); return 'AAAA'; };
    c.call = async (_name, args) => {
      requests.push(plain(args));
      if (requests.length === 2 && scenario === 'failure') throw Error('failed');
      if (scenario === 'stop') vm.runInContext('S.ocrRun.stop = true', c);
      return Array.from({ length: args.count }, (_, i) => ({ page: args.firstPage + i, newStudent: args.firstPage + i === 6 || (args.contextPages === 1 && i === 0) }));
    };
    await vm.runInContext("TOOL_ACTIONS['ocr-split-ai']()", c);
    const state = vm.runInContext('S', c);
    if (scenario === 'stop') {
      assert.equal(requests.length, 1); assert.equal(state.ocrRun.stopped, true);
      assert.deepEqual(plain(state.ocrStarts), [1, 3]);
    } else {
      assert.deepEqual(pieces.map(p => [p.from, p.to]), [[1, 4], [4, 8], [8, 9]]);
      assert.deepEqual(requests.map(p => p.contextPages), [0, 1, 1]);
      assert.equal(state.ocrRun.done, 9);
      assert.deepEqual(plain(state.ocrStarts), scenario === 'success' ? [1, 6] : [1, 5, 6, 7, 8]);
      assert.equal(state.ocrDetected[4].newStudent, false);
    }
  }
});

test('manual start uses hidden totals for confirmation and enables retry; status panel displays stop reasons', async () => {
  const c = client();
  vm.runInContext("S.manual={pendingCount:0,errorCount:2,rows:[]}; S.boot={status:{manual:{state:'OFF',remaining:0,last:{blocked:true,blockingReason:'quota',blockingMessage:'try later'}}}}", c);
  let message, options;
  c.preparePrompt = async () => true;
  c.confirmModal = async (_title, text) => { message = text; return true; };
  c.withBusy = async (_text, fn) => fn();
  c.call = async (_name, args) => { options = args; return { data: {}, status: {} }; };
  c.afterToolStart = () => {};
  await vm.runInContext("TOOL_ACTIONS['start-manual']()", c);
  assert.match(message, /2개 행/); assert.equal(options.retryErrors, true);
  assert.match(c.statusPanel('manual', '행'), /quota[\s\S]*try later/);
});
