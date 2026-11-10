// 생기부 도우미 HTML 화면(App.html)용 서버 함수 모음
// 시트는 DB로 그대로 쓰고, 화면은 google.script.run으로 아래 ui* 함수만 호출한다.

const ASSISTANT_DIALOG_TITLE = '생기부 도우미';
const ASSISTANT_DIALOG_WIDTH = 1280;
const ASSISTANT_DIALOG_HEIGHT = 880;
const AUTO_OPEN_ASSISTANT_PROPERTY = 'AUTO_OPEN_ASSISTANT';

const API_KEY_PROPERTY_BY_PROVIDER = {
  claude: 'CLAUDE_API_KEY',
  gpt: 'OPENAI_API_KEY',
  gemini: 'GEMINI_API_KEY',
};

// 도우미 설정 화면에서 다루는 설정 키. 이 목록에 없는 키는 저장하지 않는다.
const ASSISTANT_SETTING_GROUPS = {
  basic: [
    'NEIS_MAX_BYTES',
    'NEIS_WARNING_BYTES',
  ],
  ai: ['AI_PROVIDER', 'CLAUDE_MODEL', 'GPT_MODEL', 'GEMINI_MODEL', 'REASONING'],
  prompts: ['PROMPT_1', 'PROMPT_2', 'SYSTEM_GUIDE', 'GRADING_GUIDE', 'BEHAVIOR_RECORD_PROMPT', 'MANUAL_PROMPT', 'OCR_PROMPT'],
  tools: ['COMMON_PHRASE_PREVIOUS_LIMIT'],
  // 분량 기준은 설정 화면이 아니라 각 화면의 프롬프트 카드 아래에서 바꾼다.
  lengths: [
    'RECORD_DRAFT_MIN_BYTES',
    'RECORD_DRAFT_MAX_BYTES',
    'FINAL_RECORD_MIN_BYTES',
    'FINAL_RECORD_MAX_BYTES',
    'MANUAL_RECORD_MIN_BYTES',
    'MANUAL_RECORD_MAX_BYTES',
    'BEHAVIOR_RECORD_MIN_BYTES',
    'BEHAVIOR_RECORD_MAX_BYTES',
  ],
  advanced: ['MAX_RECORD_CHARS', 'MAX_INPUT_CHARS', 'MAX_RUN_SECONDS', 'AUTO_NEXT_DELAY_MS', 'FAST_GENERATION'],
};

const ASSISTANT_SETTING_LABELS = {
  RECORD_DRAFT_MIN_BYTES: '초안 최소 바이트',
  RECORD_DRAFT_MAX_BYTES: '초안 최대 바이트',
  FINAL_RECORD_MIN_BYTES: '초안 합치기 최소 바이트',
  FINAL_RECORD_MAX_BYTES: '초안 합치기 최대 바이트',
  MANUAL_RECORD_MIN_BYTES: '수동추가 최소 바이트',
  MANUAL_RECORD_MAX_BYTES: '수동추가 최대 바이트',
  BEHAVIOR_RECORD_MIN_BYTES: '행발 최소 바이트',
  BEHAVIOR_RECORD_MAX_BYTES: '행발 최대 바이트',
  NEIS_MAX_BYTES: '나이스 최대 바이트',
  NEIS_WARNING_BYTES: '나이스 경고 바이트',
  AI_PROVIDER: 'AI 서비스',
  CLAUDE_MODEL: 'Claude 모델',
  GPT_MODEL: 'GPT 모델',
  GEMINI_MODEL: 'Gemini 모델',
  REASONING: '추론 수준',
  PROMPT_1: '과제별 초안 지시사항',
  PROMPT_2: '초안 합치기 지시사항',
  SYSTEM_GUIDE: '공통 작성 지침',
  GRADING_GUIDE: '채점 공통 지침',
  BEHAVIOR_RECORD_PROMPT: '행발 추천 문구 작성 지침',
  MANUAL_PROMPT: '수동추가 생기부 작성 지침',
  OCR_PROMPT: '손글씨 추출 지침',
  COMMON_PHRASE_PREVIOUS_LIMIT: '공통문구 중복 방지용 참고 문구 수',
  MAX_RECORD_CHARS: '초안 최대 글자수(구버전 설정)',
  MAX_INPUT_CHARS: 'AI에 보낼 과제 원문 최대 글자수',
  MAX_RUN_SECONDS: '1회 실행 안전 중단 시간(초)',
  AUTO_NEXT_DELAY_MS: '다음 자동 실행까지 대기(ms)',
  FAST_GENERATION: '빠른 생성(실험)',
};

// --- UI 어댑터 ---------------------------------------------------------
// 기존 메뉴 함수는 getUi_().alert()로 결과를 알린다. 도우미 화면에서 같은 함수를
// 실행할 때는 알림을 모아 화면에 돌려주고, 확인창은 화면에서 이미 확인했으므로 OK로 처리한다.

let UI_ADAPTER_ = null;

function getUi_() {
  return UI_ADAPTER_ || SpreadsheetApp.getUi();
}

function makeCollectingUi_() {
  const messages = [];
  const ButtonSet = { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' };
  const Button = { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE' };

  return {
    messages,
    ButtonSet,
    Button,
    alert(a, b, c) {
      // alert(prompt) / alert(prompt, buttons) / alert(title, prompt, buttons)
      const buttons = c !== undefined ? c : (ButtonSet[b] ? b : undefined);
      const isConfirm = buttons && buttons !== ButtonSet.OK;

      if (!isConfirm) {
        const text = c !== undefined || (b !== undefined && !ButtonSet[b]) ? `${a}\n${b}` : a;
        messages.push(String(text));
      }

      return buttons === ButtonSet.YES_NO || buttons === ButtonSet.YES_NO_CANCEL ? Button.YES : Button.OK;
    },
    prompt() {
      throw new Error('도우미 화면에서는 입력창을 사용할 수 없습니다.');
    },
  };
}

function runWithCollectedUi_(fn) {
  const adapter = makeCollectingUi_();
  UI_ADAPTER_ = adapter;

  try {
    fn();
  } finally {
    UI_ADAPTER_ = null;
  }

  return { messages: adapter.messages };
}

// --- 화면 열기 -----------------------------------------------------------

function openAssistant() {
  try {
    hideWorkSheetsOnce_();
  } catch (err) {
    log_('openAssistant', 'ERROR', `시트 숨기기 실패: ${err.message}`);
  }

  const html = HtmlService.createTemplateFromFile('App')
    .evaluate()
    .setWidth(ASSISTANT_DIALOG_WIDTH)
    .setHeight(ASSISTANT_DIALOG_HEIGHT);

  SpreadsheetApp.getUi().showModalDialog(html, ASSISTANT_DIALOG_TITLE);
}

function include_(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

// 설치형 onOpen 트리거에서 호출된다. (단순 트리거 onOpen에서는 대화상자를 띄울 수 없다)
function openAssistantOnOpen() {
  try {
    openAssistant();
  } catch (err) {
    log_('openAssistantOnOpen', 'ERROR', String(err.message || err));
  }
}

function showAssistantHintOnOpen_() {
  let autoOpen = false;

  try {
    autoOpen = PropertiesService.getDocumentProperties().getProperty(AUTO_OPEN_ASSISTANT_PROPERTY) === 'ON';
  } catch (err) {
    // 단순 트리거에서 속성 읽기가 막히면 안내만 표시한다.
  }

  if (autoOpen) return;

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '메뉴 [⭐ 생기부 자동화] > [▶ 생기부 도우미 열기]에서 시작하세요.',
      '생기부 도우미',
      8
    );
  } catch (err) {
    // toast 실패는 무시
  }
}

function hasAutoOpenTrigger_() {
  return ScriptApp.getProjectTriggers().some(trigger => {
    return trigger.getHandlerFunction() === 'openAssistantOnOpen';
  });
}

function uiSetAutoOpen(enabled) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'openAssistantOnOpen') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  if (enabled) {
    ScriptApp.newTrigger('openAssistantOnOpen')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
  }

  PropertiesService.getDocumentProperties().setProperty(AUTO_OPEN_ASSISTANT_PROPERTY, enabled ? 'ON' : 'OFF');
  log_('uiSetAutoOpen', 'OK', enabled ? '시트 열 때 도우미 자동 열기 켬' : '시트 열 때 도우미 자동 열기 끔');

  return { autoOpen: !!enabled };
}

// --- 공통 조회 -----------------------------------------------------------

function isAssistantSheetsReady_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return [
    SHEET_NAMES.config,
    SHEET_NAMES.assignments,
    SHEET_NAMES.submissions,
    SHEET_NAMES.records,
    SHEET_NAMES.studentFinalRecords,
  ].every(name => !!ss.getSheetByName(name));
}

// 지정한 열만 읽어 {row, ...} 객체 배열로 돌려준다. 빈 행은 keyField 기준으로 건너뛴다.
// 제출물 원문(extractedText)처럼 긴 열이 있는 시트가 많아서, 시트 전체가 아니라 필요한 열 묶음만 읽는다.
function readSheetRecords_(sheetName, fields, keyField) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const h = headerIndex_(sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim()));
  const columns = readColumnValues_(sh, 2, lastRow - 1, fields.concat(keyField ? [keyField] : []).map(field => h[field]));
  const valueAt = (field, i) => (h[field] === undefined ? '' : columns[h[field]][i]);
  const records = [];

  for (let i = 0; i < lastRow - 1; i++) {
    if (keyField && h[keyField] !== undefined && String(valueAt(keyField, i) || '').trim() === '') continue;

    const record = { row: i + 2 };

    fields.forEach(field => {
      record[field] = toClientValue_(valueAt(field, i));
    });

    records.push(record);
  }

  return records;
}

// 0부터 센 열 번호 목록을 받아, 붙어 있는 열은 한 번에 읽고 {열 번호: [값, ...]}로 돌려준다.
function readColumnValues_(sh, startRow, rowCount, indexes) {
  const sorted = Array.from(new Set(indexes.filter(i => i !== undefined))).sort((a, b) => a - b);
  const columns = {};
  let i = 0;

  while (i < sorted.length) {
    let end = i;
    while (end + 1 < sorted.length && sorted[end + 1] === sorted[end] + 1) end++;

    const values = sh.getRange(startRow, sorted[i] + 1, rowCount, end - i + 1).getValues();
    for (let k = i; k <= end; k++) {
      columns[sorted[k]] = values.map(r => r[k - i]);
    }
    i = end + 1;
  }

  return columns;
}

// google.script.run은 Date를 돌려줄 수 없으므로 문자열로 바꾼다.
function toClientValue_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }

  return value;
}

function isChecked_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

function countColumnValues_(sheetName, header, predicate) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;

  const h = headerMap_(sh);
  if (!h[header]) return 0;

  return sh.getRange(2, h[header], sh.getLastRow() - 1, 1)
    .getValues()
    .filter(r => predicate(r[0]))
    .length;
}

// 과제목록 시트를 한 번만 읽어 과제 수, 체크한 과제 수, 목록에 보이는 과제 키를 함께 센다.
function readAssignmentSummary_() {
  const summary = { assignments: 0, includedAssignments: 0, listed: new Set() };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.assignments);
  if (!sh) return summary;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return summary;

  const h = headerMap_(sh);
  const col = header => (h[header] ? h[header] - 1 : undefined);
  const columns = readColumnValues_(sh, 2, lastRow - 1, ['courseId', 'courseWorkId', 'includeInFinal', 'hidden'].map(col));
  const valuesOf = header => (h[header] ? columns[h[header] - 1] : null);
  const courseIds = valuesOf('courseId');
  const courseWorkIds = valuesOf('courseWorkId');
  const included = valuesOf('includeInFinal');
  const hidden = valuesOf('hidden');

  for (let i = 0; i < lastRow - 1; i++) {
    if (courseWorkIds && String(courseWorkIds[i] || '').trim() !== '') summary.assignments++;
    if (included && isChecked_(included[i])) summary.includedAssignments++;
    // 비활성화되거나 숨긴 수업·과제의 자료는 시트에 남아도 채점과 초안 확인에서 뺀다.
    if (courseIds && courseWorkIds && !(hidden && isChecked_(hidden[i]))) {
      summary.listed.add(`${courseIds[i]}|${courseWorkIds[i]}`);
    }
  }

  return summary;
}

// 과제목록에 보이는 과제의 키
function listedAssignmentKeys_() {
  return readAssignmentSummary_().listed;
}

// 초안 시트는 본문이 길어서 필요한 열만 읽는다.
function countListedDrafts_(listed) {
  const counts = { drafts: 0, draftsChecked: 0 };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.records);
  if (!sh || sh.getLastRow() < 2) return counts;

  const h = headerMap_(sh);
  if (!h.submissionKey) return counts;

  const rowCount = sh.getLastRow() - 1;
  const listedKeys = listed || listedAssignmentKeys_();
  const columns = readColumnValues_(sh, 2, rowCount, [h.submissionKey - 1, h.teacherChecked ? h.teacherChecked - 1 : undefined]);
  const checks = h.teacherChecked ? columns[h.teacherChecked - 1] : [];

  columns[h.submissionKey - 1].forEach((key, i) => {
    if (!isSubmissionFromIncludedAssignment_(key, listedKeys)) return;
    counts.drafts++;
    if (isChecked_(checks[i])) counts.draftsChecked++;
  });

  return counts;
}

// draftPending은 진행 상태(status.drafts.remaining)에서 이미 센 값을 받아 다시 세지 않는다.
function getAssistantCounts_(draftPending) {
  const notBlank = v => String(v || '').trim() !== '';
  const assignmentSummary = readAssignmentSummary_();
  const draftCounts = countListedDrafts_(assignmentSummary.listed);

  return {
    assignments: assignmentSummary.assignments,
    includedAssignments: assignmentSummary.includedAssignments,
    submissions: countColumnValues_(SHEET_NAMES.submissions, 'submissionKey', notBlank),
    draftPending: draftPending === undefined ? countPendingRecordDrafts_() : draftPending,
    drafts: draftCounts.drafts,
    draftsChecked: draftCounts.draftsChecked,
    finalRecords: countColumnValues_(SHEET_NAMES.studentFinalRecords, 'finalRecord', notBlank),
  };
}

// withSettings가 true면 설정 화면 값도 함께 돌려줘서, 화면을 처음 열 때 요청을 한 번 줄인다.
function uiGetBootstrap(withSettings) {
  const props = PropertiesService.getUserProperties();
  const isSetup = isAssistantSheetsReady_();
  const config = isSetup ? getConfigMap_() : {};
  const ai = getAiProviderAndModel_(config);
  const keys = {};
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const status = isSetup ? getStatus_(triggers) : null;

  Object.keys(API_KEY_PROPERTY_BY_PROVIDER).forEach(provider => {
    // 키 값은 절대 화면으로 보내지 않고 저장 여부만 보낸다.
    keys[provider] = !!props.getProperty(API_KEY_PROPERTY_BY_PROVIDER[provider]);
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const boot = {
    isSetup,
    keys,
    provider: ai.provider,
    model: ai.model,
    reasoning: ai.reasoning,
    counts: isSetup ? getAssistantCounts_(status.drafts.remaining) : null,
    status,
    autoOpen: triggers.includes('openAssistantOnOpen'),
    hiddenCourses: getHiddenCourses_(),
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
  };

  if (withSettings === true && isSetup) boot.settings = uiGetSettings();
  return boot;
}

// --- 시작하기 ------------------------------------------------------------

function uiRunSetup() {
  const result = runWithCollectedUi_(() => setupSheets());
  return Object.assign(result, { bootstrap: uiGetBootstrap() });
}

function uiSaveApiKey(provider, key) {
  const propertyName = API_KEY_PROPERTY_BY_PROVIDER[provider];
  if (!propertyName) throw new Error(`알 수 없는 AI 서비스입니다: ${provider}`);

  const trimmed = String(key || '').trim();
  if (!trimmed) throw new Error('API 키가 비어 있습니다.');

  PropertiesService.getUserProperties().setProperty(propertyName, trimmed);
  log_('uiSaveApiKey', 'OK', `현재 사용자 ${provider} API 키 저장 완료`);

  return uiGetBootstrap();
}

function uiClearApiKey(provider) {
  const propertyName = API_KEY_PROPERTY_BY_PROVIDER[provider];
  if (!propertyName) throw new Error(`알 수 없는 AI 서비스입니다: ${provider}`);

  PropertiesService.getUserProperties().deleteProperty(propertyName);
  log_('uiClearApiKey', 'OK', `현재 사용자 ${provider} API 키 삭제 완료`);

  return uiGetBootstrap();
}

// --- 설정 ---------------------------------------------------------------

function uiGetSettings() {
  const sh = getSheet_(SHEET_NAMES.config);
  const values = sh.getDataRange().getValues();
  const rowsByKey = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) rowsByKey[key] = values[i];
  }

  const groups = {};

  Object.keys(ASSISTANT_SETTING_GROUPS).forEach(group => {
    groups[group] = ASSISTANT_SETTING_GROUPS[group].map(key => {
      const row = rowsByKey[key] || [];

      return {
        key,
        label: ASSISTANT_SETTING_LABELS[key] || key,
        value: String(row[1] === undefined ? '' : row[1]),
        description: String(row[2] || ''),
        options: CONFIG_VALIDATION_OPTIONS[key] || null,
      };
    });
  });

  // 화면에서 비어 있는 프롬프트 대신 보여줄 기본값
  const defaults = {};
  ASSISTANT_SETTING_GROUPS.prompts.forEach(key => {
    try {
      defaults[key] = uiGetDefaultPrompt(key);
    } catch (err) {
      // 기본값이 없는 키는 건너뛴다.
    }
  });
  // 행발 기본 프롬프트는 진지한 버전(기본값) 외에 가벼운 버전도 고를 수 있다.
  defaults.BEHAVIOR_RECORD_PROMPT_LIGHT = defaultBehaviorRecordPrompt_('light');

  return { groups, defaults };
}

function uiSaveSettings(patch) {
  saveSettingsPatch_(patch);
  return uiGetSettings();
}

// 프롬프트·분량처럼 화면이 값을 이미 알고 있을 때 쓴다. 설정 전체를 다시 읽어 보내지 않아 빠르다.
function uiSaveSettingValues(patch) {
  return { saved: saveSettingsPatch_(patch) };
}

// 허용한 설정만 검사해서 저장하고, 저장한 키 목록을 돌려준다.
function saveSettingsPatch_(patch) {
  const allowed = {};
  Object.keys(ASSISTANT_SETTING_GROUPS).forEach(group => {
    ASSISTANT_SETTING_GROUPS[group].forEach(key => {
      allowed[key] = true;
    });
  });

  const entries = Object.keys(patch || {}).filter(key => allowed[key]);
  if (entries.length === 0) return [];

  entries.forEach(key => {
    const options = CONFIG_VALIDATION_OPTIONS[key];
    const value = String(patch[key] === undefined || patch[key] === null ? '' : patch[key]);

    if (options && !options.includes(value)) {
      throw new Error(`${ASSISTANT_SETTING_LABELS[key] || key}에 쓸 수 없는 값입니다: ${value}`);
    }

    if (ASSISTANT_SETTING_GROUPS.lengths.includes(key) && !/^[1-9][0-9]{0,4}$/.test(value.trim())) {
      throw new Error(`${ASSISTANT_SETTING_LABELS[key] || key}는 1 이상의 숫자로 적어 주세요: ${value}`);
    }
  });

  const values = {};
  entries.forEach(key => {
    values[key] = patch[key];
  });
  writeConfigValues_(values);

  log_('uiSaveSettings', 'OK', `도우미 화면에서 설정 저장: ${entries.join(', ')}`);
  return entries;
}

// 설정 시트에 key → value를 쓴다. 없는 키는 맨 아래에 추가한다.
function writeConfigValues_(patch) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    // 행을 찾을 key 열만 읽는다. 값 열에는 긴 프롬프트가 들어 있다.
    const sh = getSheet_(SHEET_NAMES.config);
    const values = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
    const rowByKey = {};

    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim();
      if (key) rowByKey[key] = i + 1;
    }

    Object.keys(patch || {}).forEach(key => {
      const value = String(patch[key] === undefined || patch[key] === null ? '' : patch[key]);

      if (rowByKey[key]) {
        sh.getRange(rowByKey[key], 2).setValue(value);
      } else {
        sh.appendRow([key, value, '']);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function uiGetDefaultPrompt(key) {
  if (key === 'PROMPT_1') return defaultRecordDraftPrompt_();
  if (key === 'PROMPT_2') return defaultStudentFinalRecordPrompt_();
  if (key === 'SYSTEM_GUIDE') return defaultSystemGuide_();
  if (key === 'GRADING_GUIDE') return defaultGradingGuide_();
  if (key === 'BEHAVIOR_RECORD_PROMPT') return defaultBehaviorRecordPrompt_();
  if (key === 'MANUAL_PROMPT') return defaultManualRecordPrompt_();
  if (key === 'OCR_PROMPT') return defaultOcrPrompt_();

  throw new Error(`기본값이 없는 설정입니다: ${key}`);
}

// --- 과제 ---------------------------------------------------------------

function uiImportAssignments() {
  const result = importClassroomAssignmentsCore_();

  return {
    messages: buildAssignmentImportMessages_(result),
    assignments: uiListAssignments(),
  };
}

function uiListAssignments() {
  return readSheetRecords_(SHEET_NAMES.assignments, [
    'courseId',
    'courseName',
    'section',
    'courseWorkId',
    'assignmentTitle',
    'workType',
    'state',
    'dueAt',
    'alternateLink',
    'includeInFinal',
    'collectStatus',
    'hidden',
  ], 'courseWorkId').map(a => {
    a.key = `${a.courseId}|${a.courseWorkId}`;
    a.includeInFinal = isChecked_(a.includeInFinal);
    a.hidden = isChecked_(a.hidden);
    return a;
  });
}

// changes: {'courseId|courseWorkId': true/false}. 화면은 체크를 이미 반영했으므로 목록을 돌려주지 않는다.
function uiSetAssignmentsIncluded(changes) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sh = getSheet_(SHEET_NAMES.assignments);
    const h = headerMap_(sh);
    const lastRow = sh.getLastRow();

    if (!h.includeInFinal || !h.courseId || !h.courseWorkId) {
      throw new Error('과제목록 시트의 열 구성이 올바르지 않습니다. 시작하기에서 시트 초기 설정을 실행하세요.');
    }

    if (lastRow < 2) return;

    const courseIds = sh.getRange(2, h.courseId, lastRow - 1, 1).getValues();
    const courseWorkIds = sh.getRange(2, h.courseWorkId, lastRow - 1, 1).getValues();
    const includeRange = sh.getRange(2, h.includeInFinal, lastRow - 1, 1);
    const included = includeRange.getValues();

    for (let i = 0; i < courseIds.length; i++) {
      const key = `${courseIds[i][0]}|${courseWorkIds[i][0]}`;
      if (Object.prototype.hasOwnProperty.call(changes || {}, key)) included[i][0] = !!changes[key];
    }

    includeRange.setValues(included);
  } finally {
    lock.releaseLock();
  }
}

// changes: {'courseId|courseWorkId': true/false}. 행을 지우지 않고 표시만 해서 바로 되돌릴 수 있다.
// 숨긴 과제가 수집과 합치기에 들어가지 않도록 체크를 함께 푼다.
function uiSetAssignmentsHidden(changes) {
  withDocumentLock_(() => {
    let sh = getSheet_(SHEET_NAMES.assignments);
    if (!headerMap_(sh).hidden) sh = ensureSheetHeaders_(SHEET_NAMES.assignments);

    const h = headerMap_(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const courseIds = sh.getRange(2, h.courseId, lastRow - 1, 1).getValues();
    const courseWorkIds = sh.getRange(2, h.courseWorkId, lastRow - 1, 1).getValues();
    const hiddenRange = sh.getRange(2, h.hidden, lastRow - 1, 1);
    const includeRange = sh.getRange(2, h.includeInFinal, lastRow - 1, 1);
    const hidden = hiddenRange.getValues();
    const included = includeRange.getValues();

    courseIds.forEach((r, i) => {
      const key = `${r[0]}|${courseWorkIds[i][0]}`;
      if (!Object.prototype.hasOwnProperty.call(changes, key)) return;
      hidden[i][0] = changes[key] ? true : '';
      if (changes[key]) included[i][0] = false;
    });

    hiddenRange.setValues(hidden);
    includeRange.setValues(included);
  });
}

// 과제 화면의 변경을 한 번의 요청으로 저장한다. 요청을 나누면 창을 닫을 때 뒤의 요청이 나가지 못한다.
function uiSaveAssignmentChanges(changes) {
  if (Object.keys(changes.includes).length) uiSetAssignmentsIncluded(changes.includes);
  if (changes.clears.length) uiClearCollectStatus(changes.clears);
  if (Object.keys(changes.assignmentHides).length) uiSetAssignmentsHidden(changes.assignmentHides);
  if (Object.keys(changes.courseHides).length) uiHideCourses(changes.courseHides);
}

// 과제의 수집 표시만 지운다. 모은 제출물은 그대로 둔다.
function uiClearCollectStatus(keys) {
  const keySet = new Set(keys || []);

  withDocumentLock_(() => {
    const sh = getSheet_(SHEET_NAMES.assignments);
    const h = headerMap_(sh);
    const lastRow = sh.getLastRow();

    if (!h.collectStatus || !h.courseId || !h.courseWorkId) {
      throw new Error('과제목록 시트의 열 구성이 올바르지 않습니다. 시작하기에서 시트 초기 설정을 실행하세요.');
    }

    if (lastRow < 2) return;

    const courseIds = sh.getRange(2, h.courseId, lastRow - 1, 1).getValues();
    const courseWorkIds = sh.getRange(2, h.courseWorkId, lastRow - 1, 1).getValues();
    const cells = [];

    courseIds.forEach((r, i) => {
      if (keySet.has(`${r[0]}|${courseWorkIds[i][0]}`)) cells.push(sh.getRange(i + 2, h.collectStatus).getA1Notation());
    });

    if (cells.length) sh.getRangeList(cells).clearContent().clearNote();
  });
}

// courses: {courseId: 이름}. 과제목록에서 빼고, 다음에 과제를 가져올 때도 건너뛴다. 모은 제출물은 그대로 둔다.
function uiHideCourses(courses) {
  // 수집은 시작할 때 읽은 행 번호로 상태를 쓰므로, 그동안 행을 지우면 다른 과제에 상태가 적힌다.
  if (PropertiesService.getUserProperties().getProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY) === 'ON') {
    throw new Error('수집이 끝난 뒤에 수업을 숨겨 주세요.');
  }

  withDocumentLock_(() => {
    saveHiddenCourses_(Object.assign(getHiddenCourses_(), courses));

    const sh = getSheet_(SHEET_NAMES.assignments);
    const h = headerMap_(sh);
    const lastRow = sh.getLastRow();
    if (!h.courseId || lastRow < 2) return;

    const rows = [];
    sh.getRange(2, h.courseId, lastRow - 1, 1).getValues().forEach((r, i) => {
      if (Object.prototype.hasOwnProperty.call(courses, String(r[0]))) rows.push(i + 2);
    });

    deleteSheetRows_(sh, rows, sh.getLastColumn());
  });
}

function uiShowCourse(courseId) {
  withDocumentLock_(() => {
    const hidden = getHiddenCourses_();
    delete hidden[courseId];
    saveHiddenCourses_(hidden);
  });
}

// --- 제출물 수집 ---------------------------------------------------------

function uiStartCollect() {
  const result = runWithCollectedUi_(() => collectSubmissionsForIncludedAssignments());
  return Object.assign(result, { status: uiGetStatus() });
}

function uiListSubmissions() {
  const draftKeys = new Set(
    readSheetRecords_(SHEET_NAMES.records, ['submissionKey'], 'submissionKey')
      .map(r => String(r.submissionKey))
  );

  return readSheetRecords_(SHEET_NAMES.submissions, [
    'submissionKey',
    'studentNo',
    'studentName',
    'courseId',
    'courseName',
    'courseWorkId',
    'assignmentTitle',
    'state',
    'late',
    'textChars',
    'fileTitles',
    'submissionLink',
    'aiStatus',
  ], 'submissionKey').map(s => {
    s.hasDraft = draftKeys.has(String(s.submissionKey));
    return s;
  });
}

// 제출물 하나의 원문만 본다. 키 열로 행을 찾은 뒤 그 행만 읽는다.
function uiGetSubmissionText(submissionKey) {
  const notFound = () => new Error('제출물을 찾을 수 없습니다. 목록을 새로고침해 주세요.');
  const sh = getSheet_(SHEET_NAMES.submissions);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) throw notFound();

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const keyCol = headers.indexOf('submissionKey') + 1;
  if (!keyCol) throw notFound();

  const index = sh.getRange(2, keyCol, lastRow - 1, 1).getValues()
    .findIndex(r => String(r[0]) === String(submissionKey));
  if (index < 0) throw notFound();

  const values = sh.getRange(index + 2, 1, 1, lastCol).getValues()[0];
  const found = { row: index + 2 };
  ['submissionKey', 'studentName', 'assignmentTitle', 'extractedText', 'fileTitles', 'fileLinks', 'submissionLink'].forEach(field => {
    const i = headers.indexOf(field);
    found[field] = i < 0 ? '' : toClientValue_(values[i]);
  });

  return found;
}

// --- 초안 생성 -----------------------------------------------------------

function uiQueueDrafts(submissionKeys) {
  const keySet = new Set((submissionKeys || []).map(String));
  if (keySet.size === 0) throw new Error('초안을 만들 제출물을 선택하세요.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  let queued = 0;

  try {
    const sh = getSheet_(SHEET_NAMES.submissions);
    const h = headerMap_(sh);
    const lastRow = sh.getLastRow();

    if (!h.aiStatus || !h.submissionKey) throw new Error('제출물 시트에서 aiStatus 열을 찾을 수 없습니다.');
    if (lastRow < 2) throw new Error('제출물이 없습니다.');

    const keys = sh.getRange(2, h.submissionKey, lastRow - 1, 1).getValues();
    const rows = [];

    keys.forEach((r, i) => {
      if (keySet.has(String(r[0]))) rows.push(i + 2);
    });

    if (rows.length === 0) throw new Error('선택한 제출물을 시트에서 찾지 못했습니다. 목록을 새로고침해 주세요.');

    queued = markSubmissionRowsPending_(sh, h, rows);
  } finally {
    lock.releaseLock();
  }

  log_('uiQueueDrafts', 'OK', `${queued}개 제출물을 대기 상태로 표시`);
  return { queued };
}

function uiStartDrafts() {
  // 실행이 끊겨 '생성중'으로 남은 제출물도 이어서 만든다.
  if (PropertiesService.getUserProperties().getProperty('AUTO_RECORD_DRAFT_TRIGGER') !== 'ON') {
    withDocumentLock_(() => resetStatusValues_(SHEET_NAMES.submissions, 'aiStatus', ['생성중'], '대기'));
  }
  const result = runWithCollectedUi_(() => startAutoRecordDraftTrigger());
  return Object.assign(result, { status: uiGetStatus() });
}

// 초안 합치기를 처음부터 다시 하지 않고, 멈추거나 오류 난 학생만 이어서 만든다.
function uiResumeFinalRecords() {
  if (PropertiesService.getUserProperties().getProperty('AUTO_STUDENT_FINAL_TRIGGER') !== 'ON') {
    withDocumentLock_(() => resetStatusValues_(SHEET_NAMES.studentFinalRecords, 'finalStatus', ['생성중', '오류'], '대기'));
  }
  const result = runWithCollectedUi_(() => startAutoStudentFinalRecordTrigger_());
  return Object.assign(result, { status: uiGetStatus() });
}

// header 열에서 from 값인 칸을 to로 바꾼다. 바꾼 개수를 돌려준다.
function resetStatusValues_(sheetName, header, from, to) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;
  const col = headerMap_(sh)[header];
  if (!col) return 0;

  const range = sh.getRange(2, col, sh.getLastRow() - 1, 1);
  const values = range.getValues();
  let changed = 0;
  values.forEach(r => {
    if (from.includes(String(r[0] || '').trim())) { r[0] = to; changed++; }
  });
  if (changed) range.setValues(values);
  return changed;
}

function uiStopAll() {
  stopAutoRecordDraftTrigger(false);
  return { messages: ['진행 중인 자동 작업을 모두 중지했습니다.'], status: uiGetStatus() };
}

// --- 진행 상태 -----------------------------------------------------------

function uiGetStatus() {
  return getStatus_();
}

// triggers: 이미 읽은 트리거 핸들러 이름 목록. 없으면 여기서 읽는다.
function getStatus_(triggerNames) {
  const props = PropertiesService.getUserProperties();
  const triggers = triggerNames || ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const triggerCount = name => triggers.filter(n => n === name).length;

  // 자동 작업은 실행 중에 스크립트 잠금을 잡고, 끝날 때 다음 트리거를 만든다.
  // 그래서 'ON'인데 트리거도 없고 잠금도 비어 있으면, 예전에 멈춘 작업이 남긴 표시이므로 끈다.
  let idle = null;
  const isIdle = () => {
    if (idle === null) {
      const lock = LockService.getScriptLock();
      idle = lock.tryLock(0);
      if (idle) lock.releaseLock();
    }
    return idle;
  };
  const stateOf = (prop, handler) => {
    const state = props.getProperty(prop) || 'OFF';
    if (state === 'ON' && triggerCount(handler) === 0 && isIdle()) {
      props.setProperty(prop, 'OFF');
      log_('uiGetStatus', 'RESET', `${prop}: 실행 중인 작업과 트리거가 없어 OFF로 바꿈`);
      return 'OFF';
    }
    return state;
  };

  return {
    collect: {
      state: stateOf(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'autoProcessPendingSubmissionCollections'),
      triggers: triggerCount('autoProcessPendingSubmissionCollections'),
      remaining: countPendingSubmissionCollections_(),
      last: getAutoSubmissionCollectionLastResult_(),
    },
    drafts: {
      state: stateOf('AUTO_RECORD_DRAFT_TRIGGER', 'autoProcessPendingRecordDrafts'),
      triggers: triggerCount('autoProcessPendingRecordDrafts'),
      remaining: countPendingRecordDrafts_(),
      last: getAutoRecordDraftLastResult_(),
    },
    final: {
      state: stateOf('AUTO_STUDENT_FINAL_TRIGGER', 'autoProcessPendingStudentFinalRecords'),
      triggers: triggerCount('autoProcessPendingStudentFinalRecords'),
      remaining: countPendingStudentFinalRecords_(),
      last: getAutoStudentFinalLastResult_(),
    },
    common: {
      state: stateOf('AUTO_COMMON_PHRASE_TRIGGER', 'autoProcessPendingCommonPhrases'),
      triggers: triggerCount('autoProcessPendingCommonPhrases'),
      // 공통문구는 [문구 만들기]를 다시 누르면 처음부터 새로 만든다. 오류로 멈춘 뒤 남은 빈 칸은 화면의 [남은 것 이어서 만들기]로 따로 보여 준다.
      remaining: props.getProperty('AUTO_COMMON_PHRASE_TRIGGER') === 'ON' ? countPendingCommonPhrases_() : 0,
      last: getAutoCommonPhraseLastResult_(),
    },
    behavior: {
      state: stateOf(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'autoProcessPendingBehaviorRecommendations'),
      triggers: triggerCount('autoProcessPendingBehaviorRecommendations'),
      remaining: countPendingBehaviorRecommendationRows_(),
      last: null,
    },
    manual: {
      state: stateOf(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'autoProcessPendingManualAddedRecords'),
      triggers: triggerCount('autoProcessPendingManualAddedRecords'),
      remaining: countPendingManualAddedRows_(),
      last: null,
    },
    grading: {
      state: stateOf(AUTO_GRADING_TRIGGER_PROPERTY, 'autoProcessPendingGrades'),
      triggers: triggerCount('autoProcessPendingGrades'),
      remaining: countPendingGrades_(),
      last: getAutoGradingLastResult_(),
    },
  };
}

// --- 초안 확인 / 초안 합치기 ----------------------------------------------

function uiListDrafts() {
  const listed = listedAssignmentKeys_();

  return readSheetRecords_(SHEET_NAMES.records, [
    'submissionKey',
    'studentNo',
    'studentName',
    'courseName',
    'assignmentTitle',
    'finalRecord',
    'charCount',
    'caution',
    'teacherChecked',
    'createdAt',
    'model',
  ], 'submissionKey').filter(r => isSubmissionFromIncludedAssignment_(r.submissionKey, listed)).map(r => {
    r.key = String(r.submissionKey);
    r.teacherChecked = isChecked_(r.teacherChecked);
    return r;
  });
}

function uiUpdateDraft(row, key, patch) {
  return updateRecordRow_(SHEET_NAMES.records, row, key, patch, r => String(r.submissionKey));
}

function uiStartFinalRecords() {
  const result = runWithCollectedUi_(() => generateStudentFinalRecordsForIncludedAssignments());
  return Object.assign(result, { status: uiGetStatus() });
}

function uiListFinalRecords() {
  return readSheetRecords_(SHEET_NAMES.studentFinalRecords, [
    'finalQueueKey',
    'studentNo',
    'studentName',
    'courseName',
    'assignmentCount',
    'sourceAssignments',
    'finalRecord',
    'charCount',
    'caution',
    'teacherChecked',
    'createdAt',
    'finalStatus',
    'finalError',
  ], 'studentName').map(r => {
    r.key = finalRecordRowKey_(r);
    r.teacherChecked = isChecked_(r.teacherChecked);
    return r;
  });
}

function uiUpdateFinalRecord(row, key, patch) {
  return updateRecordRow_(SHEET_NAMES.studentFinalRecords, row, key, patch, finalRecordRowKey_);
}

function finalRecordRowKey_(r) {
  return String(r.finalQueueKey || `${r.studentNo}|${r.studentName}`);
}

// 붙어 있는 행은 한 번에 지운다. 아래 묶음부터 지워야 위쪽 행 번호가 바뀌지 않는다.
// 시트는 고정 행 밖의 행을 모두 지울 수 없으므로 마지막 한 행은 내용만 비운다.
function deleteSheetRows_(sh, rows, lastCol) {
  const sorted = Array.from(new Set(rows)).sort((a, b) => a - b);
  const runs = [];

  sorted.forEach(row => {
    const run = runs[runs.length - 1];
    if (run && run.start + run.count === row) run.count++;
    else runs.push({ start: row, count: 1 });
  });

  let remaining = sh.getMaxRows() - sh.getFrozenRows();

  runs.reverse().forEach(run => {
    if (run.count < remaining) {
      sh.deleteRows(run.start, run.count);
      remaining -= run.count;
      return;
    }

    if (run.count > 1) sh.deleteRows(run.start + 1, run.count - 1);
    sh.getRange(run.start, 1, 1, lastCol).clearContent();
    remaining = 1;
  });
}

// 화면에서 고른 초안(kind='review') 또는 합친 생기부(kind='final') 행을 지운다.
// items: [{row, key}] — 행 번호가 바뀌었을 수 있으므로 같은 행의 키가 일치할 때만 지운다.
function uiDeleteRecords(kind, items) {
  const isDraft = kind === 'review';
  const sheetName = isDraft ? SHEET_NAMES.records : SHEET_NAMES.studentFinalRecords;
  const keyOf = isDraft ? r => String(r.submissionKey) : finalRecordRowKey_;
  const targets = (items || []).map(it => ({ row: Number(it.row), key: String(it.key) }));
  if (targets.length === 0) throw new Error('지울 항목을 선택하세요.');

  const deleted = withDocumentLock_(() => {
    const sh = getSheet_(sheetName);
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());

    // 본문이 긴 시트라 전체를 읽지 않고 행을 확인할 열만 읽는다.
    const keyColumns = {};
    (isDraft ? ['submissionKey'] : ['finalQueueKey', 'studentNo', 'studentName']).forEach(header => {
      const col = headers.indexOf(header) + 1;
      if (col) keyColumns[header] = sh.getRange(1, col, lastRow, 1).getValues();
    });

    const rows = targets.filter(t => {
      if (t.row < 2 || t.row > lastRow) return false;
      const current = {};
      Object.keys(keyColumns).forEach(header => {
        current[header] = keyColumns[header][t.row - 1][0];
      });
      return keyOf(current) === t.key;
    }).map(t => t.row);

    if (rows.length < targets.length) {
      throw new Error('시트 내용이 바뀌어 지우지 못했습니다. 목록을 새로고침한 뒤 다시 시도해 주세요.');
    }

    deleteSheetRows_(sh, rows, lastCol);

    if (isDraft) resetDeletedDraftStatus_(targets.map(t => t.key));
    SpreadsheetApp.flush();
    return rows.length;
  });

  log_('uiDeleteRecords', 'OK', `${sheetName}: ${deleted}개 삭제`);
  return { deleted };
}

// 초안을 지운 제출물은 제출물 시트의 aiStatus를 비워 '초안 없음'으로 되돌린다.
function resetDeletedDraftStatus_(keys) {
  const remaining = new Set(
    readSheetRecords_(SHEET_NAMES.records, ['submissionKey'], 'submissionKey').map(r => String(r.submissionKey))
  );
  const keySet = new Set(keys.filter(k => !remaining.has(k)));
  if (keySet.size === 0) return;

  const sh = getSheet_(SHEET_NAMES.submissions);
  const h = headerMap_(sh);
  const lastRow = sh.getLastRow();
  if (!h.submissionKey || !h.aiStatus || lastRow < 2) return;

  const keyValues = sh.getRange(2, h.submissionKey, lastRow - 1, 1).getValues();
  const statusRange = sh.getRange(2, h.aiStatus, lastRow - 1, 1);
  const statuses = statusRange.getValues();
  let changed = false;

  keyValues.forEach((r, i) => {
    if (keySet.has(String(r[0])) && ['생성완료', '근거부족'].includes(String(statuses[i][0]))) {
      statuses[i][0] = '';
      changed = true;
    }
  });

  if (changed) statusRange.setValues(statuses);
}

// 화면에서 수정한 finalRecord / teacherChecked를 저장한다.
// 화면을 연 뒤 행이 바뀌었을 수 있으므로 같은 행의 키가 일치할 때만 쓴다.
function updateRecordRow_(sheetName, row, key, patch, keyOf) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sh = getSheet_(sheetName);
    const h = headerMap_(sh);
    const lastCol = sh.getLastColumn();

    if (row < 2 || row > sh.getLastRow()) throw new Error('수정할 행을 찾을 수 없습니다. 목록을 새로고침해 주세요.');

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    const values = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const current = {};
    headers.forEach((header, i) => {
      if (header) current[header] = values[i];
    });

    if (keyOf(current) !== String(key)) {
      throw new Error('시트 내용이 바뀌어 저장하지 못했습니다. 목록을 새로고침한 뒤 다시 시도해 주세요.');
    }

    if (patch && patch.finalRecord !== undefined && h.finalRecord) {
      sh.getRange(row, h.finalRecord).setValue(String(patch.finalRecord));
    }

    if (patch && patch.teacherChecked !== undefined && h.teacherChecked) {
      sh.getRange(row, h.teacherChecked).setValue(!!patch.teacherChecked);
    }

    SpreadsheetApp.flush();

    const updated = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    return {
      row,
      key: String(key),
      finalRecord: h.finalRecord ? String(updated[h.finalRecord - 1] || '') : '',
      charCount: h.charCount ? toClientValue_(updated[h.charCount - 1]) : '',
      teacherChecked: h.teacherChecked ? isChecked_(updated[h.teacherChecked - 1]) : false,
    };
  } finally {
    lock.releaseLock();
  }
}

// --- 다운로드 -------------------------------------------------------------

// 화면의 [엑셀로 받기]용: 산출물 시트 전체를 { fileName, headers, rows } 표로 돌려준다.
// textCol이 있으면 화면에서 그 열의 나이스 바이트를 계산해 붙인다.
function uiExportData(kind) {
  if (kind === 'review' || kind === 'final') return exportRecordSheet_(kind);
  if (kind === 'common') return exportCommonPhrases_();
  if (kind === 'behavior') return exportBehaviorRows_();
  if (kind === 'manual') return exportManualSheet_();
  if (kind === 'grading') return exportGradingResults_();
  throw new Error('알 수 없는 다운로드 종류입니다.');
}

function exportRecordSheet_(kind) {
  const columns = kind === 'review'
    ? [['studentNo', '번호'], ['studentName', '이름'], ['courseName', '수업'], ['assignmentTitle', '과제'],
      ['finalRecord', '초안'], ['caution', '확인할 점'], ['teacherChecked', '확인함'], ['createdAt', '만든 시각']]
    : [['studentNo', '번호'], ['studentName', '이름'], ['courseName', '수업'], ['assignmentCount', '반영 과제 수'],
      ['sourceAssignments', '반영 과제'], ['finalRecord', '합친 생기부'], ['caution', '확인할 점'],
      ['teacherChecked', '확인함'], ['finalStatus', '상태'], ['createdAt', '만든 시각']];
  const sheetName = kind === 'review' ? SHEET_NAMES.records : SHEET_NAMES.studentFinalRecords;
  let records = readSheetRecords_(sheetName, columns.map(c => c[0]).concat('submissionKey'), kind === 'review' ? 'submissionKey' : 'studentName');

  if (kind === 'review') {
    const listed = listedAssignmentKeys_();
    records = records.filter(r => isSubmissionFromIncludedAssignment_(r.submissionKey, listed));
  }

  return {
    fileName: kind === 'review' ? '초안' : '합친 생기부',
    headers: columns.map(c => c[1]),
    rows: records.map(r => columns.map(c => c[0] === 'teacherChecked' ? (isChecked_(r[c[0]]) ? 'O' : '') : r[c[0]])),
    textCol: columns.findIndex(c => c[0] === 'finalRecord'),
  };
}

function exportCommonPhrases_() {
  const phrases = uiGetCommonPhrases().phrases.filter(p => p.text);
  return {
    fileName: '공통문구',
    headers: ['번호', '문구'],
    rows: phrases.map((p, i) => [i + 1, p.text]),
    textCol: 1,
  };
}

function exportBehaviorRows_() {
  const columns = [['studentName', '이름'], ['selectedTraits', '특성'], ['customTraits', '목록에 없는 특성'],
    ['positiveObservation', '잘한 점'], ['improvementNeeded', '보완할 점'], ['growthEvidence', '성장한 모습'],
    ['recommendedRecord', '행동특성 및 종합의견']];
  return {
    fileName: '행발',
    headers: columns.map(c => c[1]),
    rows: uiListBehaviorRows().rows.map(r => columns.map(c => r[c[0]])),
    textCol: columns.length - 1,
  };
}

// 수동추가는 올린 원본 열을 그대로 두고 받는다(화면 목록처럼 500행에서 자르지 않는다).
function exportManualSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANUAL_SHEET_NAME);
  if (!sh || sh.getLastRow() < 1) throw new Error('수동추가 시트가 없습니다.');

  const manualConfig = getManualRecordConfig_(getConfigMap_());
  const lastCol = Math.max(sh.getLastColumn(), manualConfig.outputCol);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map((v, i) => String(v || '').trim() || colToA1_(i + 1));
  const lastRow = sh.getLastRow();
  const values = lastRow >= manualConfig.startRow
    ? sh.getRange(manualConfig.startRow, 1, lastRow - manualConfig.startRow + 1, lastCol).getValues()
    : [];

  return {
    fileName: '수동추가',
    headers,
    rows: values
      .filter(r => r.some(v => String(v || '').trim() !== ''))
      .map(r => r.map(toClientValue_)),
  };
}
