const SHEET_NAMES = {
  config: '설정',
  assignments: '과제목록',
  students: '학생명단',
  submissions: '제출물',
  records: '생기부초안',
  studentFinalRecords: '학생별생기부',
  commonPhrases: '공통문구생성',
  behaviorRecommendations: '행발문구추천',
  logs: '로그',
};

const ASSIGNMENTS_TRAILING_BLANK_ROWS = 1000;
const MANUAL_SHEET_NAME = '수동추가';
const MANUAL_RESET_ROWS = 1000;
const AUTO_MANUAL_RECORD_TRIGGER_PROPERTY = 'AUTO_MANUAL_RECORD_TRIGGER';
const AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY = 'AUTO_SUBMISSION_COLLECTION_TRIGGER';
const AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY = 'AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER';
const AUTO_BUSY_RETRY_DELAY_MS = 60000;

// 생기부 분량 기준(나이스 바이트). [최소 키, 최대 키, 기본 최소, 기본 최대]
const RECORD_BYTE_RANGE_SETTINGS = {
  draft: ['RECORD_DRAFT_MIN_BYTES', 'RECORD_DRAFT_MAX_BYTES', 900, 1200],
  final: ['FINAL_RECORD_MIN_BYTES', 'FINAL_RECORD_MAX_BYTES', 900, 1200],
  manual: ['MANUAL_RECORD_MIN_BYTES', 'MANUAL_RECORD_MAX_BYTES', 900, 1200],
  behavior: ['BEHAVIOR_RECORD_MIN_BYTES', 'BEHAVIOR_RECORD_MAX_BYTES', 750, 1200],
};
const COMMON_PHRASE_DEFAULT_PREVIOUS_LIMIT = 10;
const COMMON_PHRASE_DUPLICATE_RETRY_LIMIT = 3;
const BEHAVIOR_RECORD_MIN_CHARS_DESCRIPTION = '행발 최소 글자수\n(나이스 바이트로는 약 3을 곱하시면 됩니다)\n(완벽하게 지키지는 못합니다)';
const BEHAVIOR_RECORD_MAX_CHARS_DESCRIPTION = '행발 최대 글자수\n(나이스 바이트로는 약 3을 곱하시면 됩니다)\n(완벽하게 지키지는 못합니다)';
const BEHAVIOR_RECOMMENDATION_HEADERS = [
  'studentName',
  'selectedTraits',
  'customTraits',
  'positiveObservation',
  'improvementNeeded',
  'growthEvidence',
  'recommendedRecord',
  'byteCount',
  'status',
];
// 행발문구추천 시트 selectedTraits 열의 드롭다운 값과 같아야 한다.
const BEHAVIOR_TRAIT_KEYWORDS = [
  '없음',
  '성실',
  '책임감',
  '자발성',
  '꾸준함',
  '근면성',
  '정직함',
  '침착함',
  '자기통제력',
  '규칙 준수',
  '예의 바름',
  '겸손',
  '자기주도성',
  '자기성찰',
  '목표 의식',
  '인내심',
  '성장의지',
  '개선의지',
  '배려',
  '존중',
  '협동심',
  '봉사정신',
  '공감능력',
  '경청',
  '나눔',
  '의사소통',
  '리더십',
  '솔선수범',
  '적극성',
  '수업태도',
  '탐구력',
  '호기심',
  '문제 해결력',
  '비판적 사고',
  '창의적 사고',
  '논리적 사고',
  '긍정적 태도',
  '친화력',
  '정서적 지지',
  '따뜻함',
  '학업 열의',
  '회복탄력성',
  '공동체 의식',
  '도전 정신',
];

// 보완이 필요한 특성. 도우미 화면에서 다른 색으로 보여 준다.
const BEHAVIOR_NEGATIVE_TRAIT_KEYWORDS = [
  '규칙 미준수',
  '지각 잦음',
  '수업 방해',
  '불성실',
  '충동적',
  '거짓말',
  '공격적',
  '자기중심적',
];

BEHAVIOR_TRAIT_KEYWORDS.push(...BEHAVIOR_NEGATIVE_TRAIT_KEYWORDS);

const DASHBOARD_SHEET_NAME = '현황판';

const HEADERS = {
  '설정': ['key', 'value', 'description'],
  '과제목록': [
    'courseId',
    'courseName',
    'section',
    'courseWorkId',
    'assignmentTitle',
    'workType',
    'state',
    'dueAt',
    'updateTime',
    'alternateLink',
    'assignmentDescription',
    'maxPoints',
    'teacherGrade',
    'includeInFinal',
    'weight',
    'category',
    'collectStatus',
    'memo'
  ],
  '학생명단': ['studentNo', 'name', 'email', 'classroomUserId', 'memo'],
  '제출물': [
    'submissionKey', 'studentNo', 'studentName', 'email',
    'courseId', 'courseName', 'courseWorkId', 'assignmentTitle',
    'submissionId', 'state', 'late', 'draftGrade', 'assignedGrade',
    'maxPoints', 'teacherGrade',
    'updateTime', 'submissionLink',
    'fileTitles', 'fileLinks',
    'extractedText', 'textChars',
    'collectedAt', 'aiStatus',
    'assignmentDescription',
  ],


  '생기부초안': [
    'submissionKey',
    'studentNo',
    'studentName',
    'email',
    'courseName',
    'assignmentTitle',
    'finalRecord',
    'charCount',
    'caution',
    'teacherChecked',
    'createdAt',
    'model',
    'evidenceSummary',
    'recordDraft',
    'gradeText',
    'usage'
  ],

  '학생별생기부': [
    'studentNo',
    'studentName',
    'email',
    'courseName',
    'assignmentCount',
    'sourceSubmissionKeys',
    'sourceAssignments',
    'evidencePack',
    'finalRecord',
    'charCount',
    'caution',
    'teacherChecked',
    'createdAt',
    'model',
    'usage',
    'finalQueueKey',
    'finalStatus',
    'finalError'
  ],
  '공통문구생성': ['prompt', 'classWork', 'byte', 'number'],
  '로그': ['timestamp', 'action', 'status', 'message'],
  
};

function onOpen() {
  const ui = getUi_();

  // 기존 메뉴는 '고급' 하위 메뉴로 유지하고, 기본 흐름은 도우미 화면에서 진행한다.
  const advancedMenu = ui
    .createMenu('고급 (기존 메뉴)')
    .addItem('1. 시트 초기 설정 (모델 목록 갱신)', 'setupSheets')
    .addItem('2. Claude API 키 저장', 'saveClaudeApiKey')
    .addItem('Claude API 키 삭제', 'clearClaudeApiKey')
    .addItem('GPT API 키 저장', 'saveOpenAiApiKey')
    .addItem('GPT API 키 삭제', 'clearOpenAiApiKey')
    .addItem('Gemini API 키 저장', 'saveGeminiApiKey')
    .addItem('Gemini API 키 삭제', 'clearGeminiApiKey')

    .addSeparator()
    .addItem('3. Classroom 과제목록 가져오기', 'importClassroomAssignments')
    .addItem('4. 체크한 과제 모두 수집', 'collectSubmissionsForIncludedAssignments')
    //.addItem('현재 선택한 과제 1개만 수집', 'collectSubmissionsForSelectedAssignment')
    //.addItem('선택 과제 생기부 반영 체크', 'markSelectedAssignmentsIncluded')
    //.addItem('선택 과제 생기부 반영 해제', 'unmarkSelectedAssignmentsIncluded')
    //.addItem('체크한 과제 제출물 일괄 수집', 'collectSubmissionsForIncludedAssignments')
    .addSeparator()
    
    .addItem('5-1. 선택 제출물 대기열에 넣기', 'markSelectedSubmissionsPending')
    //.addItem('대기열 1회 처리', 'processPendingRecordDrafts')
    .addItem('5-2. 자동 생성 트리거 시작', 'startAutoRecordDraftTrigger')
    //.addItem('5-3. 선택한 제출물로 생기부 초안 생성', 'generateRecordsForSelectedSubmissions')
    .addItem('자동 생성 트리거 중지', 'stopAutoRecordDraftTrigger')
    .addItem('자동 생성 상태 확인', 'showAutoRecordDraftStatus')
    .addSeparator()
    .addItem('6. 학생별 최종 생기부 생성', 'generateStudentFinalRecordsForIncludedAssignments')
    .addSeparator()
    .addItem('7. 공통문구 생성', 'generateCommonPhrases')
    //.addItem('7-1. 공통문구 중복검사', 'checkCommonPhraseDuplicates')
    .addSeparator()
    .addItem('8. 보고있는 시트의 학생 데이터 삭제', 'deleteActiveSheetStudentData')
    .addSeparator()
    .addItem('9. 수동추가 데이터 생기부 생성', 'generateManualAddedRecords')
    .addSeparator()
    .addItem('10. 행발 추천 문구 생성', 'generateBehaviorRecommendationsForSelectedRows')
    //.addItem('생기부초안 열 재배열', 'reorganizeRecordsSheetLayout_')
    //.addItem('화면 정리 / 서식 적용', 'applyFrontendLayout')
    //.addItem('시트 순서 / 숨김 정리', 'organizeSheetTabs_')
    //.addItem('현황판 새로고침', 'refreshDashboard')
    .addSeparator()
    .addItem('제출물 원문 열 보이기/숨기기', 'toggleExtractedTextColumn');
    //.addSeparator()
    //.addItem('AI 모델·추론 조합 점검', 'checkAllAiCombinations')
    //.addItem('행발 추천 모든 모델·추론으로 비교 (선택한 행)', 'compareBehaviorAcrossModels');

  ui.createMenu('⭐ 생기부 자동화')
    .addItem('▶ 생기부 도우미 열기', 'openAssistant')
    .addSeparator()
    .addSubMenu(advancedMenu)
    .addToUi();

  try {
    syncConfigModelsOnOpen_();
  } catch (err) {
    // onOpen에서 설정 시트 동기화 실패 시 메뉴 표시는 유지
  }

  showAssistantHintOnOpen_();
}

function saveOpenAiApiKey() {
  const ui = getUi_();
  const res = ui.prompt(
    'OpenAI API 키 저장',
    '본인 OpenAI API 키를 입력하세요. 이 키는 현재 사용자 계정의 User Properties에만 저장됩니다.',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const key = res.getResponseText().trim();
  if (!key) {
    ui.alert('API 키가 비어 있습니다.');
    return;
  }

  PropertiesService.getUserProperties().setProperty('OPENAI_API_KEY', key);

  log_('saveOpenAiApiKey', 'OK', '현재 사용자 OpenAI API 키 저장 완료');
  ui.alert('OpenAI API 키를 현재 사용자 계정에 저장했습니다.');
}

function saveGeminiApiKey() {
  const ui = getUi_();
  const res = ui.prompt(
    'Gemini API 키 저장',
    '본인 Gemini API 키를 입력하세요. 이 키는 현재 사용자 계정의 User Properties에만 저장됩니다.',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const key = res.getResponseText().trim();
  if (!key) {
    ui.alert('API 키가 비어 있습니다.');
    return;
  }

  PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', key);

  log_('saveGeminiApiKey', 'OK', '현재 사용자 Gemini API 키 저장 완료');
  ui.alert('Gemini API 키를 현재 사용자 계정에 저장했습니다.');
}

function clearAiApiKeys() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('CLAUDE_API_KEY');
  props.deleteProperty('OPENAI_API_KEY');
  props.deleteProperty('GEMINI_API_KEY');

  PropertiesService.getScriptProperties().deleteProperty('CLAUDE_API_KEY');
  PropertiesService.getScriptProperties().deleteProperty('OPENAI_API_KEY');
  PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');

  log_('clearAiApiKeys', 'OK', 'Claude/OpenAI/Gemini API 키 삭제 완료');
  getUi_().alert('Claude/OpenAI/Gemini API 키를 삭제했습니다.');
}

function generateCommonPhrases() {
  const ui = getUi_();
  const sh = getOrCreateSheet_(SHEET_NAMES.commonPhrases);

  ensureCommonPhraseSheet_();
  ensureCommonPhraseConfigDefaults_();
  styleCommonPhraseSheet_();

  try {
    requireCurrentAiApiKey_();
  } catch (err) {
    ui.alert(String(err.message || err));
    return;
  }

  const prompt = String(sh.getRange('A2').getValue() || '').trim();
  const sourceText = String(sh.getRange('B2').getValue() || '').trim();
  const count = Number(sh.getRange('D2').getValue() || 0);

  if (!prompt) {
    ui.alert('공통문구생성 시트 A2에 프롬프트를 입력하세요.');
    return;
  }

  if (!sourceText) {
    ui.alert('공통문구생성 시트 B2에 바탕이 되는 문구를 입력하세요.');
    return;
  }

  if (!Number.isInteger(count) || count < 1) {
    ui.alert('공통문구생성 시트 D2에 생성할 개수를 1 이상의 정수로 입력하세요.');
    return;
  }

  clearCommonPhraseOutputs_(count);
  saveCommonPhraseRun_({
    runId: makeCommonPhraseRunId_(),
    total: count,
  });
  PropertiesService.getUserProperties().setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'ON');

  log_('generateCommonPhrases', 'QUEUE', `${count}개 공통문구 대기열 생성`);
  startAutoCommonPhraseTrigger_();
}

function checkCommonPhraseDuplicates() {
  const sh = getOrCreateSheet_(SHEET_NAMES.commonPhrases);
  ensureCommonPhraseSheet_();
  styleCommonPhraseSheet_();

  const configuredCount = Number(sh.getRange('D2').getValue() || 0);
  const runTotal = getCommonPhraseTotal_();
  const rowCount = Math.max(sh.getLastRow() - 2, configuredCount, runTotal, 0);

  if (rowCount < 1) {
    getUi_().alert('검사할 공통문구가 없습니다.\nB3 아래에 생성된 문구가 있는지 확인하세요.');
    return;
  }

  const range = sh.getRange(3, 2, rowCount, 1);
  const values = range.getValues();
  const notes = range.getNotes();
  const groups = {};
  let nonEmptyCount = 0;

  values.forEach((row, index) => {
    const text = String(row[0] || '').trim();
    if (!text) return;

    nonEmptyCount++;
    const key = normalizeCommonPhraseDuplicateKey_(text);
    if (!key) return;

    if (!groups[key]) groups[key] = [];
    groups[key].push(index + 3);
  });

  const duplicateGroups = Object.values(groups).filter(rows => rows.length > 1);
  const duplicateRows = new Set();

  duplicateGroups.forEach(rows => {
    rows.forEach(row => duplicateRows.add(row));
  });

  if (nonEmptyCount === 0) {
    getUi_().alert('검사할 공통문구가 없습니다.\nB3 아래에 생성된 문구가 있는지 확인하세요.');
    return;
  }

  const backgrounds = values.map((row, index) => {
    const rowNumber = index + 3;
    const text = String(row[0] || '').trim();
    return [text && duplicateRows.has(rowNumber) ? '#f4cccc' : '#ffffff'];
  });

  const updatedNotes = notes.map((row, index) => {
    const rowNumber = index + 3;
    const currentNote = String(row[0] || '');

    if (!duplicateRows.has(rowNumber)) {
      return [currentNote.indexOf('[중복검사]') === 0 ? '' : currentNote];
    }

    const rows = duplicateGroups.find(groupRows => groupRows.includes(rowNumber)) || [];
    return [`[중복검사]\n같은 문구가 있는 행: ${rows.join(', ')}`];
  });

  range.setBackgrounds(backgrounds);
  range.setNotes(updatedNotes);

  log_(
    'checkCommonPhraseDuplicates',
    'OK',
    `검사 ${nonEmptyCount}개, 중복 그룹 ${duplicateGroups.length}개, 중복 행 ${duplicateRows.size}개`
  );

  if (duplicateGroups.length === 0) {
    getUi_().alert(`공통문구 중복검사 완료\n\n검사 문구: ${nonEmptyCount}개\n중복 없음`);
    return;
  }

  const detail = duplicateGroups
    .slice(0, 10)
    .map((rows, index) => `${index + 1}. ${rows.join(', ')}행`)
    .join('\n');
  const extra = duplicateGroups.length > 10 ? `\n외 ${duplicateGroups.length - 10}개 그룹` : '';

  getUi_().alert(
    `공통문구 중복검사 완료\n\n` +
    `검사 문구: ${nonEmptyCount}개\n` +
    `중복 그룹: ${duplicateGroups.length}개\n` +
    `중복 행: ${duplicateRows.size}개\n\n` +
    `${detail}${extra}`
  );
}

function normalizeCommonPhraseDuplicateKey_(text) {
  return normalizeAiText_(text)
    .trim()
    .replace(/\s+/g, ' ');
}

function ensureCommonPhraseSheet_() {
  const sh = getOrCreateSheet_(SHEET_NAMES.commonPhrases);
  const headers = HEADERS[SHEET_NAMES.commonPhrases];
  const current = sh.getRange(1, 1, 1, headers.length).getValues()[0]
    .map(value => String(value || '').trim());
  const needsHeader = headers.some((header, index) => current[index] !== header);

  if (needsHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sh;
}

function ensureCommonPhraseConfigDefaults_() {
  const config = getSheet_(SHEET_NAMES.config);
  const defaults = getCommonPhraseConfigDefaults_();
  const lastRow = config.getLastRow();
  const values = lastRow >= 2
    ? config.getRange(2, 1, lastRow - 1, 3).getValues()
    : [];
  const existingKeys = {};

  values.forEach((row, index) => {
    const key = String(row[0] || '').trim();
    if (!key) return;

    existingKeys[key] = index + 2;
  });

  const rowsToAppend = [];

  defaults.forEach(row => {
    const existingRow = existingKeys[row[0]];

    if (!existingRow) {
      rowsToAppend.push(row);
      return;
    }

    const currentDescription = String(values[existingRow - 2][2] || '');
    if (currentDescription !== row[2]) {
      config.getRange(existingRow, 3).setValue(row[2]);
    }
  });

  if (rowsToAppend.length > 0) {
    config.getRange(config.getLastRow() + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
  }
}

function getCommonPhraseConfigDefaults_() {
  return [[
    'COMMON_PHRASE_PREVIOUS_LIMIT',
    String(COMMON_PHRASE_DEFAULT_PREVIOUS_LIMIT),
    '공통문구 생성 시 AI 프롬프트에 포함할 최근 생성 문구 수\n기본 10개. 입력 토큰 절감을 위해 너무 크게 설정하지 않는 것을 권장합니다.',
  ]];
}

function clearCommonPhraseOutputs_(count) {
  const sh = getSheet_(SHEET_NAMES.commonPhrases);
  ensureCommonPhraseOutputRows_(count);
  const clearRows = Math.max(sh.getLastRow() - 2, count, 1);

  const outputRange = sh.getRange(3, 2, clearRows, 2);
  outputRange.clearContent();
  outputRange.clearNote();
}

function ensureCommonPhraseOutputRows_(count) {
  const sh = getSheet_(SHEET_NAMES.commonPhrases);
  const requiredRows = 2 + Math.max(Number(count || 0), 1);

  if (sh.getMaxRows() < requiredRows) {
    sh.insertRowsAfter(sh.getMaxRows(), requiredRows - sh.getMaxRows());
  }
}

function makeCommonPhraseRunId_() {
  return `CP-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')}-${Utilities.getUuid().slice(0, 8)}`;
}

function saveCommonPhraseRun_(run) {
  PropertiesService.getUserProperties().setProperty(
    'AUTO_COMMON_PHRASE_RUN',
    JSON.stringify(run || {})
  );
}

function getCommonPhraseRun_() {
  const raw = PropertiesService.getUserProperties().getProperty('AUTO_COMMON_PHRASE_RUN');
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function getCommonPhraseTotal_() {
  const run = getCommonPhraseRun_();
  const total = run && Number(run.total || 0);

  if (total > 0) return total;

  return 0;
}

function getGeneratedCommonPhrases_(total, outputValues) {
  total = Number(total || 0);
  if (total < 1) return [];

  const values = outputValues || getSheet_(SHEET_NAMES.commonPhrases).getRange(3, 2, total, 1).getValues();

  return values
    .map(row => String(row[0] || '').trim())
    .filter(text => !!text);
}

function getNextCommonPhraseIndex_(total, outputValues, startIndex) {
  total = Number(total || 0);
  if (total < 1) return 0;

  const values = outputValues || getSheet_(SHEET_NAMES.commonPhrases).getRange(3, 2, total, 1).getValues();
  const start = Math.max(Number(startIndex || 1), 1);

  for (let i = start - 1; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) {
      return i + 1;
    }
  }

  return 0;
}

function countPendingCommonPhrases_() {
  const total = getCommonPhraseTotal_();
  if (total < 1) return 0;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.commonPhrases);
  if (!sh) return 0;

  const values = sh.getRange(3, 2, total, 1).getValues();

  return values.filter(row => !String(row[0] || '').trim()).length;
}

function processPendingCommonPhrasesCore_() {
  const sh = getSheet_(SHEET_NAMES.commonPhrases);
  const config = getConfigMap_();
  const total = getCommonPhraseTotal_();
  const inputValues = sh.getRange(2, 1, 1, 2).getValues()[0];
  const prompt = String(inputValues[0] || '').trim();
  const sourceText = String(inputValues[1] || '').trim();
  const batchSize = RUN_ITEM_LIMIT_;
  const maxRunSeconds = Number(config.MAX_RUN_SECONDS || 270);
  const previousLimit = getCommonPhrasePreviousLimit_(config);
  const startedAt = Date.now();
  const ai = getAiProviderAndModel_(config);
  const provider = ai.provider;
  const model = ai.model;
  const reasoning = ai.reasoning;
  const systemGuide = commonPhraseSystemGuide_();

  let processed = 0;
  let skipped = 0;
  let failed = 0;
  let blocked = false;
  let blockingReason = '';
  let blockingMessage = '';

  if (!prompt || !sourceText || total < 1) {
    return {
      processed,
      skipped,
      failed,
      remaining: 0,
      blocked: true,
      blockingReason: '공통문구 입력값 오류',
      blockingMessage: '공통문구생성 시트의 A2, B2, D2 값을 확인하세요.',
    };
  }

  ensureCommonPhraseOutputRows_(total);
  const outputValues = sh.getRange(3, 2, total, 1).getValues();
  const generatedTexts = getGeneratedCommonPhrases_(total, outputValues);
  const generatedTextSet = buildExactTextSet_(generatedTexts);
  let nextIndex = getNextCommonPhraseIndex_(total, outputValues);

  // 빠른 생성: 다음 몇 개 문구를 같은 '이전 문구' 목록으로 한꺼번에 받는다.
  // 그 문구의 첫 시도는 미리 보낸 프롬프트(fastPrompts)를 그대로 써야 받아 둔 답과 맞는다.
  // 서로의 결과는 보지 못하므로, 똑같은 문구가 나오면 기존 중복 재시도가 최신 목록으로 다시 만든다.
  const prefetcher = createAiPrefetcher_(config, 'processPendingCommonPhrasesCore_');
  const fastPrompts = {};
  const upcomingPhraseRequests = fromIndex => {
    const requests = [];
    const limit = Math.min(FAST_GENERATION_CONCURRENCY, batchSize - processed);
    const previousTexts = getRecentCommonPhraseTexts_(generatedTexts, previousLimit);

    for (let idx = fromIndex; idx > 0 && requests.length < limit; idx = getNextCommonPhraseIndex_(total, outputValues, idx + 1)) {
      if (!fastPrompts[idx]) {
        fastPrompts[idx] = buildCommonPhrasePrompt_(prompt, sourceText, idx, total, previousTexts, {
          previousLimit,
          duplicatePhrase: '',
          duplicateAttempt: 0,
        });
      }
      requests.push({ provider, model, systemGuide, userPrompt: fastPrompts[idx], reasoning });
    }

    return requests;
  };

  while (processed < batchSize) {
    const elapsedSeconds = (Date.now() - startedAt) / 1000;

    if (elapsedSeconds > maxRunSeconds) {
      log_(
        'processPendingCommonPhrasesCore_',
        'STOP',
        `실행 시간 제한 접근으로 중단. 처리 ${processed}개, 실패 ${failed}개`
      );
      break;
    }

    const index = nextIndex;
    if (index === 0) break;

    const row = 2 + index;
    let phrase = '';
    let duplicatePhrase = '';

    prefetcher.prefetch(() => upcomingPhraseRequests(index));

    try {
      for (let attempt = 1; attempt <= COMMON_PHRASE_DUPLICATE_RETRY_LIMIT + 1; attempt++) {
        const previousTexts = getRecentCommonPhraseTexts_(generatedTexts, previousLimit);
        const userPrompt = attempt === 1 && fastPrompts[index] ? fastPrompts[index] : buildCommonPhrasePrompt_(
          prompt,
          sourceText,
          index,
          total,
          previousTexts,
          {
            previousLimit,
            duplicatePhrase,
            duplicateAttempt: attempt > 1 ? attempt - 1 : 0,
          }
        );
        const result = callAi_(provider, model, systemGuide, userPrompt, reasoning);
        const parsed = parseClaudeJson_(result.text);

        phrase = normalizeAiText_(parsed.record_draft || result.text).trim();

        if (phrase && !generatedTextSet[phrase]) {
          break;
        }

        duplicatePhrase = phrase;
        log_(
          'processPendingCommonPhrasesCore_',
          'DUPLICATE_RETRY',
          `${index}번째 문구 중복으로 재생성 시도 ${attempt}/${COMMON_PHRASE_DUPLICATE_RETRY_LIMIT + 1}`
        );
      }

      if (!phrase) {
        throw new Error('공통문구 생성 결과가 비어 있습니다.');
      }

      if (generatedTextSet[phrase]) {
        throw new Error(`공통문구 중복으로 저장하지 않음: ${phrase}`);
      }

      const phraseCell = sh.getRange(row, 2);
      phraseCell.setValue(phrase);
      phraseCell.clearNote();
      sh.getRange(row, 3).setFormula(`=IF(B${row}="","",2*LENB(B${row})-LEN(B${row}))`);
      sh.setRowHeight(row, 150);
      outputValues[index - 1][0] = phrase;
      generatedTexts.push(phrase);
      generatedTextSet[phrase] = true;
      nextIndex = getNextCommonPhraseIndex_(total, outputValues, index);

      processed++;
      Utilities.sleep(1200);

    } catch (err) {
      if (isAiTimeBudgetError_(err)) {
        log_('processPendingCommonPhrasesCore_', 'STOP', `${index}번째 문구: ${err.message}`);
        break;
      }

      const message = String(err.message || err);
      const errorInfo = classifyRecordDraftError_(message);

      failed++;
      blocked = true;
      blockingReason = errorInfo.reason;
      blockingMessage = errorInfo.userMessage;
      sh.getRange(row, 2).setNote(`${errorInfo.reason}: ${message}`);
      log_('processPendingCommonPhrasesCore_', 'STOP', `${errorInfo.reason}: ${message}`);
      break;
    }
  }

  prefetcher.finish();

  styleCommonPhraseSheet_();

  const remaining = countPendingCommonPhrases_();

  log_(
    'processPendingCommonPhrasesCore_',
    'OK',
    `처리 ${processed}개, 실패 ${failed}개, 스킵 ${skipped}개, 남은 대기 ${remaining}개`
  );

  return {
    processed,
    skipped,
    failed,
    remaining,
    blocked,
    blockingReason,
    blockingMessage,
  };
}

function startAutoCommonPhraseTrigger_() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/spreadsheets.currentonly',
    'https://www.googleapis.com/auth/script.external_request',
  ]);

  const ui = getUi_();
  const pendingCount = countPendingCommonPhrases_();

  if (pendingCount === 0) {
    ui.alert('공통문구 생성 대기열이 없습니다.');
    return;
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    ui.alert('이미 다른 자동 생성 작업이 실행 중입니다.\n현재 작업이 끝난 뒤 다시 시작해 주세요.');
    return;
  }

  try {
    deleteAutoCommonPhraseTriggers_();

    const props = PropertiesService.getUserProperties();
    props.setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'ON');

    const result = runFirstBatchQuickly_(() => processPendingCommonPhrasesCore_());

    if (result.blocked) {
      props.setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'OFF');
      deleteAutoCommonPhraseTriggers_();
      saveAutoCommonPhraseLastResult_('BLOCKED', result);
      ui.alert(buildRecordDraftResultMessage_('공통문구 생성을 중단했습니다.', result));
      return;
    }

    if (result.remaining === 0) {
      props.setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'OFF');
      deleteAutoCommonPhraseTriggers_();
      saveAutoCommonPhraseLastResult_('DONE', result);
      ui.alert(buildRecordDraftResultMessage_('공통문구 생성을 완료했습니다.', result));
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoCommonPhraseTrigger_(nextDelayMs);
    saveAutoCommonPhraseLastResult_('RUNNING', result);

    log_(
      'startAutoCommonPhraseTrigger_',
      'OK',
      `첫 배치 처리 완료. 처리 ${result.processed}개, 실패 ${result.failed}개, 남은 대기 ${result.remaining}개`
    );

  } catch (err) {
    const message = String(err.message || err);
    const result = {
      processed: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingCommonPhrases_(),
      blocked: true,
      blockingReason: '공통문구 생성 시작 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'OFF');
    deleteAutoCommonPhraseTriggers_();
    saveAutoCommonPhraseLastResult_('BLOCKED', result);
    log_('startAutoCommonPhraseTrigger_', 'ERROR', message);
    ui.alert(buildRecordDraftResultMessage_('공통문구 생성 시작 중 오류가 발생했습니다.', result));
  } finally {
    lock.releaseLock();
  }
}

function autoProcessPendingCommonPhrases() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    rescheduleAutoJobWhenBusy_('autoProcessPendingCommonPhrases', 'AUTO_COMMON_PHRASE_TRIGGER', scheduleNextAutoCommonPhraseTrigger_);
    return;
  }

  try {
    deleteAutoCommonPhraseTriggers_();

    const status = PropertiesService.getUserProperties().getProperty('AUTO_COMMON_PHRASE_TRIGGER') || 'OFF';

    if (status !== 'ON') {
      log_('autoProcessPendingCommonPhrases', 'STOP', '자동 생성 상태가 OFF라 실행하지 않음');
      return;
    }

    beginAutoRun_(scheduleNextAutoCommonPhraseTrigger_);

    const result = processPendingCommonPhrasesCore_();

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'OFF');
      saveAutoCommonPhraseLastResult_('BLOCKED', result);
      toastAutoRecordDraft_(result.blockingMessage || '공통문구 생성을 중단했습니다.', result.blockingReason || '공통문구 생성 중단');
      log_('autoProcessPendingCommonPhrases', 'BLOCKED', `${result.blockingReason}: ${result.blockingMessage}`);
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'OFF');
      saveAutoCommonPhraseLastResult_('DONE', result);
      toastAutoRecordDraft_(
        `생성 완료 ${result.processed}개, 실패 ${result.failed}개`,
        '공통문구 생성 완료'
      );
      log_('autoProcessPendingCommonPhrases', 'DONE', '남은 공통문구 대기열이 없어 종료했습니다.');
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoCommonPhraseTrigger_(nextDelayMs);
    saveAutoCommonPhraseLastResult_('RUNNING', result);

    log_(
      'autoProcessPendingCommonPhrases',
      'NEXT',
      `남은 대기 ${result.remaining}개. ${Math.round(nextDelayMs / 1000)}초 뒤 다음 실행 예약`
    );

  } catch (err) {
    const message = String(err.message || err);
    const result = {
      processed: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingCommonPhrases_(),
      blocked: true,
      blockingReason: '공통문구 자동 처리 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'OFF');
    saveAutoCommonPhraseLastResult_('BLOCKED', result);
    toastAutoRecordDraft_(message, '공통문구 자동 처리 오류');
    log_('autoProcessPendingCommonPhrases', 'ERROR', message);

  } finally {
    endAutoRun_('AUTO_COMMON_PHRASE_TRIGGER', deleteAutoCommonPhraseTriggers_);
    lock.releaseLock();
  }
}

function scheduleNextAutoCommonPhraseTrigger_(delayMs) {
  deleteAutoCommonPhraseTriggers_();

  ScriptApp.newTrigger('autoProcessPendingCommonPhrases')
    .timeBased()
    .after(delayMs)
    .create();

  log_(
    'scheduleNextAutoCommonPhraseTrigger_',
    'OK',
    `${Math.round(delayMs / 1000)}초 뒤 다음 공통문구 자동 생성 예약`
  );
}

function deleteAutoCommonPhraseTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoProcessPendingCommonPhrases') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function saveAutoCommonPhraseLastResult_(status, result) {
  const payload = {
    status,
    processed: result.processed || 0,
    failed: result.failed || 0,
    skipped: result.skipped || 0,
    remaining: result.remaining || 0,
    blocked: !!result.blocked,
    blockingReason: result.blockingReason || '',
    blockingMessage: result.blockingMessage || '',
    updatedAt: new Date().toISOString(),
  };

  PropertiesService.getUserProperties().setProperty(
    'AUTO_COMMON_PHRASE_LAST_RESULT',
    JSON.stringify(payload)
  );
}

function getAutoCommonPhraseLastResult_() {
  const raw = PropertiesService.getUserProperties().getProperty('AUTO_COMMON_PHRASE_LAST_RESULT');
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function getCommonPhrasePreviousLimit_(config) {
  const raw = String((config && config.COMMON_PHRASE_PREVIOUS_LIMIT) || '').trim();
  if (!raw) return COMMON_PHRASE_DEFAULT_PREVIOUS_LIMIT;

  const value = Number(raw);

  if (Number.isFinite(value) && value >= 0) {
    return Math.floor(value);
  }

  return COMMON_PHRASE_DEFAULT_PREVIOUS_LIMIT;
}

function buildExactTextSet_(texts) {
  const set = Object.create(null);
  (texts || []).forEach(text => {
    const value = String(text || '').trim();
    if (value) set[value] = true;
  });
  return set;
}

function getRecentCommonPhraseTexts_(texts, limit) {
  const max = Math.max(Number(limit || 0), 0);
  if (max < 1) return [];

  return (texts || [])
    .filter(text => String(text || '').trim())
    .slice(-max);
}

function buildCommonPhrasePrompt_(prompt, sourceText, index, total, previousTexts, options) {
  options = options || {};
  const rawPreviousLimit = options.previousLimit === undefined
    ? COMMON_PHRASE_DEFAULT_PREVIOUS_LIMIT
    : Number(options.previousLimit);
  const previousLimit = Number.isFinite(rawPreviousLimit) && rawPreviousLimit >= 0
    ? Math.floor(rawPreviousLimit)
    : COMMON_PHRASE_DEFAULT_PREVIOUS_LIMIT;
  const limitedPreviousTexts = getRecentCommonPhraseTexts_(previousTexts, previousLimit);
  const duplicatePhrase = String(options.duplicatePhrase || '').trim();
  const duplicateAttempt = Number(options.duplicateAttempt || 0);

  return [
    '[공통문구 생성 요청]',
    '아래 원문을 프롬프트에 따라 공통문구 1개로 작성한다.',
    '생활기록부나 수업 기록에 바로 옮겨 적을 문장은 record_draft에만 작성한다.',
    '반드시 JSON 형식으로만 답한다.',
    '',
    '[프롬프트]',
    prompt,
    '',
    '[원문]',
    sourceText,
    '',
    '[생성 조건]',
    `총 ${total}개 중 ${index}번째 문구를 작성한다.`,
    '이미 생성된 문구와 같은 표현을 반복하지 않는다.',
    `참고용 이전 문구는 최근 ${previousLimit}개까지만 제공된다. 제공되지 않은 이전 문구도 내부 중복 검사에서 완전히 같은 문구면 저장되지 않는다.`,
    'record_draft에는 문구 본문만 작성하고 번호, 따옴표, 설명을 붙이지 않는다.',
    'evidence_summary와 caution은 빈 문자열로 둔다.',
    limitedPreviousTexts.length > 0 ? '' : '',
    limitedPreviousTexts.length > 0 ? `[최근 생성된 문구 최대 ${previousLimit}개]` : '',
    limitedPreviousTexts.length > 0 ? limitedPreviousTexts.map((text, i) => `${i + 1}. ${text}`).join('\n') : '',
    duplicatePhrase ? '' : '',
    duplicatePhrase ? `[직전 중복 생성 결과 - 저장하지 않음${duplicateAttempt ? ` / 재시도 ${duplicateAttempt}` : ''}]` : '',
    duplicatePhrase ? duplicatePhrase : '',
  ].filter(line => line !== '').join('\n');
}

function commonPhraseSystemGuide_() {
  return [
    '너는 교사가 수업 기록과 생활기록부 작성에 활용할 공통문구를 생성하는 도우미다.',
    '사용자가 준 프롬프트와 원문만 근거로 삼고, 없는 사실을 꾸며내지 않는다.',
    '결과는 반드시 JSON 형식으로만 답한다.',
    '{"evidence_summary":"", "record_draft":"", "caution":""}',
    '실제로 사용할 문구는 record_draft에만 작성한다.',
  ].join('\n');
}

function ensureBehaviorRecommendationSheet_() {
  const sh = getOrCreateSheet_(SHEET_NAMES.behaviorRecommendations);
  sh.getRange(1, 1, 1, BEHAVIOR_RECOMMENDATION_HEADERS.length)
    .setValues([BEHAVIOR_RECOMMENDATION_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function styleBehaviorRecommendationSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.behaviorRecommendations);
  if (!sh) return;

  ensureBehaviorRecommendationSheet_();

  const h = headerMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 2);
  const headerCount = BEHAVIOR_RECOMMENDATION_HEADERS.length;
  const bodyRows = Math.max(lastRow - 1, 1);

  sh.setFrozenRows(1);
  setColumnWidthsByHeader_(sh, {
    studentName: 130,
    selectedTraits: 220,
    customTraits: 220,
    positiveObservation: 300,
    improvementNeeded: 300,
    growthEvidence: 300,
    recommendedRecord: 620,
    byteCount: 90,
    status: 120,
  });

  sh.getRange(1, 1, 1, headerCount)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.getRange(2, 1, bodyRows, 6)
    .setBackground('#eaf4ff')
    .setVerticalAlignment('top');
  sh.getRange(2, 7, bodyRows, 3)
    .setBackground('#fff2cc')
    .setVerticalAlignment('top');

  ['selectedTraits', 'customTraits', 'positiveObservation', 'improvementNeeded', 'growthEvidence', 'recommendedRecord'].forEach(header => {
    if (h[header]) {
      sh.getRange(2, h[header], bodyRows, 1)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setVerticalAlignment('top');
    }
  });

  ['byteCount', 'status'].forEach(header => {
    if (h[header]) {
      sh.getRange(2, h[header], bodyRows, 1)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }
  });

  if (h.studentName) {
    sh.getRange(1, h.studentName).setNote('교사가 학생을 구분하기 위한 이름입니다. AI 프롬프트에는 전송하지 않습니다.');
  }
  if (h.selectedTraits) {
    sh.getRange(1, h.selectedTraits).setNote(
      '쉼표로 구분하여 여러 특성을 입력하세요.\n예: ' + BEHAVIOR_TRAIT_KEYWORDS.join(', ')
    );
  }
  if (h.customTraits) {
    sh.getRange(1, h.customTraits).setNote('기본 키워드 목록에 없는 특성을 직접 입력하세요. 여러 개면 쉼표로 구분할 수 있습니다.');
  }
  if (h.recommendedRecord) {
    sh.getRange(1, h.recommendedRecord).setNote('AI가 생성한 행발 추천 문구입니다. caution이 있으면 이 셀 메모에 저장됩니다.');
  }
  if (h.byteCount) {
    sh.getRange(1, h.byteCount).setNote('recommendedRecord 기준 나이스 바이트 계산값입니다.');
  }

  try {
    const filter = sh.getFilter();
    if (filter) filter.remove();
    sh.getRange(1, 1, lastRow, headerCount).createFilter();
  } catch (err) {
    // 필터 생성 실패는 기능 자체를 막지 않는다.
  }
}

function generateBehaviorRecommendationsForSelectedRows() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!sh || sh.getName() !== SHEET_NAMES.behaviorRecommendations) {
    getUi_().alert('행발문구추천 시트에서 처리할 행을 선택한 뒤 실행하세요.');
    return;
  }

  generateBehaviorRecommendationsCore_();
}

// 행발문구추천 시트의 전체 데이터 행을 처리한다. (도우미 화면에서도 호출)
function generateBehaviorRecommendationsCore_() {
  const ui = getUi_();
  const lock = LockService.getScriptLock();
  const sh = ensureBehaviorRecommendationSheet_();

  styleBehaviorRecommendationSheet_();

  try {
    requireCurrentAiApiKey_();
  } catch (err) {
    ui.alert(String(err.message || err));
    return;
  }

  const h = headerMap_(sh);
  const target = resolveBehaviorRecommendationTargetRange_(sh);

  if (!target) {
    ui.alert('행발문구추천 시트에 처리할 데이터 행이 없습니다.');
    return;
  }

  if (!lock.tryLock(1000)) {
    ui.alert('이미 다른 자동 생성 작업이 실행 중입니다.\n현재 작업이 끝난 뒤 다시 시작해 주세요.');
    return;
  }

  try {
    deleteAutoBehaviorRecommendationTriggers_();

    const config = getConfigMap_();
    const result = runFirstBatchQuickly_(() => processBehaviorRecommendationRows_(sh, h, target.startRow, target.rowCount, config));
    result.scopeLabel = target.scopeLabel;

    styleBehaviorRecommendationSheet_();
    logBehaviorRecommendationResult_('generateBehaviorRecommendationsForSelectedRows', result);

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
      deleteAutoBehaviorRecommendationTriggers_();
      ui.alert(buildBehaviorRecommendationResultMessage_(result));
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
      deleteAutoBehaviorRecommendationTriggers_();
      ui.alert(buildBehaviorRecommendationResultMessage_(result));
      return;
    }

    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);
    PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'ON');
    scheduleNextAutoBehaviorRecommendationTrigger_(nextDelayMs);

    ui.alert(
      buildBehaviorRecommendationResultMessage_(result) +
      `\n\n남은 행은 ${Math.round(nextDelayMs / 1000)}초 뒤 자동으로 이어서 처리합니다.`
    );
  } catch (err) {
    const message = String(err.message || err);
    PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
    deleteAutoBehaviorRecommendationTriggers_();
    log_('generateBehaviorRecommendationsForSelectedRows', 'ERROR', message);
    ui.alert(`행발 추천 문구 생성 중 오류가 발생했습니다.\n\n${message}`);
  } finally {
    lock.releaseLock();
  }
}

function resolveBehaviorRecommendationTargetRange_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  return {
    startRow: 2,
    rowCount: lastRow - 1,
    scopeLabel: `전체 데이터행 2-${lastRow}`,
  };
}

function processBehaviorRecommendationRows_(sh, h, startRow, rowCount, config, options) {
  const missingHeaders = BEHAVIOR_RECOMMENDATION_HEADERS.filter(header => !h[header]);
  if (missingHeaders.length > 0) {
    throw new Error(`행발문구추천 시트에 필요한 헤더가 없습니다: ${missingHeaders.join(', ')}`);
  }

  const pendingOnly = !!(options && options.pendingOnly);
  const batchSize = RUN_ITEM_LIMIT_;
  const maxRunSeconds = Math.max(Number(config.MAX_RUN_SECONDS || 270), 30);
  const maxInputChars = Number(config.MAX_INPUT_CHARS || 30000);
  const behaviorPrompt = getBehaviorRecordPrompt_(config);
  const charRange = getBehaviorRecordCharRange_(config);
  const startedAt = Date.now();
  const ai = getAiProviderAndModel_(config);
  const systemGuide = behaviorRecommendationSystemGuide_();
  const lastCol = Math.max(sh.getLastColumn(), BEHAVIOR_RECOMMENDATION_HEADERS.length);

  if (rowCount <= 0) {
    return {
      processed: 0,
      skipped: 0,
      skippedAlreadyGenerated: 0,
      skippedNoTraits: 0,
      skippedBlank: 0,
      failed: 0,
      remaining: countPendingBehaviorRecommendationRows_(),
      blocked: false,
      blockingReason: '',
      blockingMessage: '',
    };
  }

  const rowValues = sh.getRange(startRow, 1, rowCount, lastCol).getValues();
  const statusUpdates = [];
  const candidates = [];
  const candidateRows = [];
  const skippedDetails = [];

  let processed = 0;
  let skipped = 0;
  let skippedAlreadyGenerated = 0;
  let skippedNoTraits = 0;
  let skippedBlank = 0;
  let failed = 0;
  let attempted = 0;
  let blocked = false;
  let blockingReason = '';
  let blockingMessage = '';

  rowValues.forEach((values, index) => {
    const row = startRow + index;
    const status = String(rowValueAt_(values, h.status) || '').trim();

    if (pendingOnly && status !== '대기' && status !== '재시도필요') {
      addBehaviorSkippedDetail_(skippedDetails, row, `상태 ${status || '빈값'}`);
      return;
    }

    const data = buildBehaviorRecommendationData_(values, h);

    if (!data.hasSelectedTraits) {
      if (data.hasAnyInput) {
        statusUpdates.push({ row, status: '특성없음' });
        skippedNoTraits++;
        addBehaviorSkippedDetail_(skippedDetails, row, '특성없음');
      } else {
        skippedBlank++;
        addBehaviorSkippedDetail_(skippedDetails, row, '빈 행');
      }
      skipped++;
      return;
    }

    const recommendedRecord = String(rowValueAt_(values, h.recommendedRecord) || '').trim();
    if (recommendedRecord) {
      skipped++;
      skippedAlreadyGenerated++;
      addBehaviorSkippedDetail_(skippedDetails, row, '이미 생성됨');
      return;
    }

    candidates.push({ row, data });
    candidateRows.push(row);
    statusUpdates.push({ row, status: '대기' });
  });

  writeBehaviorStatusRows_(sh, h.status, statusUpdates);

  // generateBehaviorRecommendationWithCharRetry_의 첫 요청과 같은 요청을 만든다.
  const prefetcher = createAiPrefetcher_(config, 'processBehaviorRecommendationRows_');
  const upcomingBehaviorRequests = fromIndex => candidates
    .slice(fromIndex, fromIndex + Math.min(FAST_GENERATION_CONCURRENCY, batchSize - processed))
    .map(candidate => ({
      provider: ai.provider,
      model: ai.model,
      systemGuide,
      userPrompt: buildBehaviorRecommendationPrompt_(candidate.data, maxInputChars, behaviorPrompt, charRange, null),
      reasoning: ai.reasoning,
      responseSchema: BEHAVIOR_RESPONSE_SCHEMA,
    }));

  for (let i = 0; i < candidates.length; i++) {
    const elapsedSeconds = (Date.now() - startedAt) / 1000;
    if (processed >= batchSize || elapsedSeconds > maxRunSeconds) break;

    const item = candidates[i];
    attempted++;

    prefetcher.prefetch(() => upcomingBehaviorRequests(i));

    try {
      sh.getRange(item.row, h.status).setValue('생성중');

      const generated = generateBehaviorRecommendationWithCharRetry_({
        ai,
        systemGuide,
        data: item.data,
        maxInputChars,
        behaviorPrompt,
        charRange,
      });
      const recommendedRecord = generated.recommendedRecord;
      const caution = generated.caution;

      if (!recommendedRecord) {
        throw makeBehaviorResponseFormatError_('recommended_record 값이 비어 있습니다.');
      }

      writeBehaviorRecommendationResult_(sh, h, item.row, recommendedRecord, '생성완료', caution, generated.reframes);
      processed++;
      Utilities.sleep(1200);
    } catch (err) {
      if (isAiTimeBudgetError_(err)) {
        sh.getRange(item.row, h.status).setValue('대기');
        log_('processBehaviorRecommendationRows_', 'STOP', `${item.row}행: ${err.message}`);
        break;
      }

      failed++;

      if (isBehaviorResponseFormatError_(err)) {
        writeBehaviorRecommendationResult_(sh, h, item.row, '', '응답형식오류', String(err.message || err));
        log_('processBehaviorRecommendationRows_', 'FORMAT_ERROR', `${item.row}행: ${err.message}`);
        continue;
      }

      const message = String(err.message || err);
      const errorInfo = classifyRecordDraftError_(message);
      writeBehaviorRecommendationResult_(sh, h, item.row, '', errorInfo.status, `${errorInfo.reason}: ${message}`);
      log_(
        'processBehaviorRecommendationRows_',
        errorInfo.shouldStop ? 'STOP' : 'ERROR',
        `${item.row}행: ${errorInfo.reason} / ${message}`
      );

      if (errorInfo.shouldStop) {
        blocked = true;
        blockingReason = errorInfo.reason;
        blockingMessage = errorInfo.userMessage;
        break;
      }
    }
  }

  prefetcher.finish();

  return {
    processed,
    skipped,
    skippedAlreadyGenerated,
    skippedNoTraits,
    skippedBlank,
    candidateRows,
    skippedDetails,
    failed,
    remaining: countPendingBehaviorRecommendationRows_(),
    blocked,
    blockingReason,
    blockingMessage,
  };
}

function buildBehaviorRecommendationData_(rowValues, h) {
  const studentName = String(rowValueAt_(rowValues, h.studentName) || '').trim();
  const sanitizeOptions = { studentName };
  const selectedTraits = sanitizePromptText_(rowValueAt_(rowValues, h.selectedTraits), sanitizeOptions).trim();
  const customTraits = sanitizePromptText_(rowValueAt_(rowValues, h.customTraits), sanitizeOptions).trim();
  const positiveObservation = sanitizePromptText_(rowValueAt_(rowValues, h.positiveObservation), sanitizeOptions).trim();
  const improvementNeeded = sanitizePromptText_(rowValueAt_(rowValues, h.improvementNeeded), sanitizeOptions).trim();
  const growthEvidence = sanitizePromptText_(rowValueAt_(rowValues, h.growthEvidence), sanitizeOptions).trim();

  const inputFields = [selectedTraits, customTraits, positiveObservation, improvementNeeded, growthEvidence];

  return {
    selectedTraits,
    customTraits,
    positiveObservation,
    improvementNeeded,
    growthEvidence,
    // 목록에서 고른 특성이 없어도 '목록에 없는 특성'을 적었으면 만든다.
    hasSelectedTraits: isMeaningfulBehaviorInput_(selectedTraits) || isMeaningfulBehaviorInput_(customTraits),
    hasAnyInput: inputFields.some(isMeaningfulBehaviorInput_),
  };
}

function addBehaviorSkippedDetail_(details, row, reason) {
  if (!details || details.length >= 12) return;
  details.push(`${row}행 ${reason}`);
}

function isMeaningfulBehaviorInput_(text) {
  return String(text || '')
    .replace(/\[(학생명|이메일|학번) 제거\]/g, '')
    .replace(/[\s,.;:|/\\()[\]{}'"`~!?\-_=+]+/g, '')
    .trim()
    .length > 0;
}

function behaviorRecommendationSystemGuide_() {
  return [
    '너는 대한민국 학교 담임교사가 행동특성 및 종합의견을 쓸 때 돕는 도구이다.',
    '교사가 고른 특성과 관찰 내용을 바탕으로 쓰고, 입력에 없는 구체적인 일화, 교우관계, 가정환경, 학업 성취를 지어내지 않는다.',
    '학생 이름과 개인정보를 결과에 포함하지 않는다.',
    '비난, 낙인, 모욕적 표현은 쓰지 않는다.',
    '보완점은 부정적인 사실을 그대로 쓰지 않고, 그 이면의 장점이나 성장 가능성으로 돌려 표현한다.',
    'caution에는 교사가 참고할 점(근거가 부족한 특성 등)이 있을 때만 짧게 쓰고, 없으면 빈 문자열로 둔다.',
    '반드시 JSON 형식으로만 답한다. reframes를 먼저 채운 뒤 recommended_record를 쓴다.',
    BEHAVIOR_RESPONSE_FORMAT,
  ].join('\n');
}

// reframes를 recommended_record보다 앞에 두어, 보완점을 돌려 쓴 표현 후보를 먼저 떠올린 뒤 본문을 쓰게 한다.
const BEHAVIOR_RESPONSE_FORMAT = '{"reframes":[{"trait":"", "candidates":["", "", ""], "chosen":""}], "recommended_record":"", "caution":""}';
const BEHAVIOR_RESPONSE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    reframes: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          trait: { type: 'string' },
          candidates: { type: 'array', items: { type: 'string' } },
          chosen: { type: 'string' },
        },
        required: ['trait', 'candidates', 'chosen'],
      },
    },
    recommended_record: { type: 'string' },
    caution: { type: 'string' },
  },
  required: ['reframes', 'recommended_record', 'caution'],
};

function buildBehaviorReframeInstruction_() {
  return [
    '[보완점 돌려 쓰기 절차]',
    '1. improvementNeeded 등 입력에 부정적인 내용이 있으면, 하나씩 reframes에 담는다. trait에는 원래 내용을 짧게 쓴다. 부정적인 내용이 없으면 reframes는 빈 배열로 둔다.',
    '2. 각 보완점마다 "행동 → 그 행동 뒤에 있는 성향 → 그 성향을 가치 있게 보는 관점" 순서로 생각하여, 서로 다른 관점의 돌려 쓴 표현을 candidates에 3개 쓴다.',
    "3. 처음 떠오르는 무난한 표현에 머물지 않는다. '성장 가능성이 기대됨', '노력한다면', '~할 필요가 있음', '~하면 좋겠음' 같은 상투적인 표현은 후보로 쓰지 않는다.",
    '4. 후보는 긍정적으로 읽히되, 담임교사가 읽으면 어떤 모습을 말하는지 알아챌 수 있어야 한다.',
    '5. 후보 중 학생의 다른 특성과 가장 자연스럽게 이어지는 것을 chosen에 쓰고, recommended_record에 그 표현을 녹여 쓴다.',
  ].join('\n');
}

function defaultBehaviorRecordPrompt_() {
  return [
    '[역할]',
    "너는 이 학생을 한 해 동안 지켜본 담임교사다. 교사 메모를 바탕으로 학교생활기록부 '행동특성 및 종합의견' 문단을 직접 쓴다.",
    '',
    '[입력 읽는 법]',
    '1. selectedTraits, customTraits는 교사가 이 학생을 보고 고른 특성이다. 교사의 판단이므로 학생의 성품이나 태도로 자연스럽게 서술해도 된다.',
    '2. 다만 특성에 맞추어 입력에 없는 구체적인 일화(무슨 일을 했는지, 언제 그랬는지)를 지어내지 않는다.',
    '3. positiveObservation, improvementNeeded, growthEvidence는 실제 관찰한 장면이다. 구체적인 행동은 이 안에서만 쓴다.',
    "4. '없음'은 고른 특성이 없다는 뜻이다.",
    '5. growthEvidence가 있을 때만 변화나 향상을 쓴다. 없으면 보완점은 앞으로의 기대로 쓴다.',
    '6. 고른 특성을 모두 넣을 필요는 없다. 문단 흐름에 어울리는 것만 고른다.',
    '',
    '[문체]',
    '1. 학생이 어떤 사람인지 그려지도록, 학생의 행동을 주어로 한 생생한 문장으로 쓴다.',
    "2. 관찰 보고서 말투를 쓰지 않는다. '~이 관찰됨', '~모습이 드러남', '~태도가 확인됨', '~하는 학생임' 같은 표현은 쓰지 않는다.",
    "3. 행동의 일반적인 의미를 설명하지 않는다. 예: '인사는 상대를 존중하는 뜻을 전하는 행동임' 같은 문장은 쓰지 않는다.",
    "4. 기록이나 근거에 대한 말을 문단에 쓰지 않는다. 예: '구체적인 근거와 함께 기록할 필요가 있음', '추가 관찰이 필요함'. 이런 내용은 caution에만 쓴다.",
    '5. 같은 내용을 말만 바꿔 되풀이하지 않는다. 분량을 채워야 하면 고른 특성을 엮어 서술한다.',
    "6. 문장 끝은 '~함', '~임', '~지님', '~기대됨' 같은 명사형 어미로 쓰되, 같은 어미를 연달아 쓰지 않는다.",
    '7. 장점을 먼저 쓴다. 보완점(improvementNeeded 등 부정적인 내용)은 아래 [보완점 돌려 쓰기]에 따라 쓴다.',
    "8. 한 번 있었던 일을 '항상', '꾸준히'처럼 부풀리지 않는다.",
    '9. 교우관계, 가정환경, 학업 성취처럼 입력에 없는 영역은 쓰지 않는다.',
    '10. 이름과 개인정보, 제목, 번호, 따옴표 없이 한 문단으로 쓴다.',
    '',
    '[보완점 돌려 쓰기]',
    '생활기록부에는 학생에 대한 부정적인 서술을 쓰지 않는다. 보완점은 다음 중 하나로 바꾸어 쓴다.',
    '- 그 행동의 이면에 있는 특성을 장점처럼 표현한다.',
    '- 앞으로의 성장 가능성이나 기대로 표현한다.',
    "'~하지 못함', '~이 부족함', '~이 잦음', '~하는 경우가 있으나' 처럼 부정적인 사실을 그대로 드러내는 표현은 쓰지 않는다.",
    '문장은 긍정적으로 읽히되, 담임교사가 읽으면 어떤 모습을 말하는지 알아챌 수 있게 쓴다.',
    '무난하게 순화하지 말고 과감하게 돌려 쓴다. 단점을 거창한 장점이나 철학처럼 부풀려, 읽는 교사가 피식 웃을 만큼 능청스러워도 좋다. 다만 문장은 생활기록부 문체(명사형 어미)를 지킨다.',
    '아래 예시는 바꾸는 방식과 말맛만 참고한다. 같은 보완점이 나와도 예시 문장을 그대로 쓰지 말고 새로 쓴다. 예시를 그대로 쓰면 한 반 학생들의 문장이 똑같아진다.',
    '- 수업 중 자주 잠 → 현실과 꿈의 경계를 자유롭게 넘나들며 수업 시간에도 상상력을 충전하는 낭만가임.',
    '- 교복 규정을 지키지 않음 → 획일화된 복장에 자신의 개성을 투영하려는 패션 철학을 가지고 있으며 학교 안에서 트렌드를 선도함.',
    '- 급식실에서 새치기함 → 배식 시간의 효율성을 누구보다 중시하며 행동력이 빠름.',
    '- 청소 시간에 도망감 → 학급의 청결보다 더 큰 대의를 위해 자리를 비우는 경우가 있으며 항상 바쁜 모습을 보임.',
    '- 말이 거침 → 돌려 말하지 않는 담백함이 있으며 언어의 본질적 의미 전달에 충실함.',
    '- 고집이 셈 → 천 번을 흔들려도 한 번도 꺾이지 않는 철학적 확신을 지님.',
    '- 눈치가 없음 → 사회적 시선이라는 보이지 않는 틀에서 자유로운 순수한 영혼임.',
    '- 학원 숙제만 함 → 자신의 우선순위가 분명하고 효율을 극대화하는 전략적 학습 설계 능력을 갖추고 있음.',
    '- 화를 잘 냄 → 불의를 보면 참지 못하는 뜨거운 가슴을 지니고 있으며 감정 에너지가 풍부함.',
    '',
    '[예시] (문체만 참고하고 문장을 그대로 가져다 쓰지 않는다)',
    '입력: 특성 - 책임감, 배려 / 긍정 관찰 - 1인 1역으로 칠판 정리를 맡음, 아픈 친구를 보건실에 데려다줌 / 보완 - 발표할 때 목소리가 작음',
    '좋은 결과: 맡은 일에 책임감을 갖고 임하여 1인 1역인 칠판 정리를 스스로 챙겨 해냄. 몸이 불편한 친구를 먼저 살피고 보건실까지 함께 가 주는 등 주변을 따뜻하게 챙기는 마음을 지님. 말하기에 앞서 생각을 차분히 다듬는 신중함을 지녔으며, 자신의 생각을 당당히 펼쳐 보이는 경험이 더해진다면 한층 크게 성장할 것으로 기대됨.',
    '나쁜 결과: 칠판 정리를 하는 모습이 관찰되어 책임감 있는 태도가 드러남. 칠판 정리는 학급 공동체를 위한 행동임. 발표할 때 목소리가 작아 의사 전달이 부족함. 배려는 구체적인 근거와 함께 기록할 필요가 있음.',
  ].join('\n');
}

function getBehaviorRecordPrompt_(config) {
  const prompt = String((config && config.BEHAVIOR_RECORD_PROMPT) || '').trim();
  return prompt || defaultBehaviorRecordPrompt_();
}

function getBehaviorRecordCharRange_(config) {
  return getRecordByteRange_(config, 'behavior');
}

function parsePositiveInteger_(value) {
  const parsed = Number(String(value || '').replace(/,/g, '').trim());
  if (!Number.isFinite(parsed) || parsed <= 0) return 0;
  return Math.floor(parsed);
}

function buildBehaviorRecommendationPrompt_(data, maxInputChars, behaviorPrompt, charRange, retryContext) {
  const inputLines = [
    data.selectedTraits ? `[selectedTraits]\n${data.selectedTraits}` : '',
    data.customTraits ? `[customTraits]\n${data.customTraits}` : '',
    data.positiveObservation ? `[positiveObservation]\n${data.positiveObservation}` : '',
    data.improvementNeeded ? `[improvementNeeded]\n${data.improvementNeeded}` : '',
    data.growthEvidence ? `[growthEvidence]\n${data.growthEvidence}` : '',
  ].filter(Boolean).join('\n\n');
  const charInstruction = buildBehaviorCharInstruction_(charRange);
  const traitInstruction = buildBehaviorTraitInstruction_();
  const retryInstruction = retryContext
    ? [
      '[재작성 요청]',
      `이전 응답은 나이스 바이트 기준 ${retryContext.charCount}바이트였고, 설정 범위(${formatBehaviorCharRange_(charRange)})를 벗어났다.`,
      '위 입력 사실은 유지하되, recommended_record를 설정 분량 범위 안에 들어오도록 다시 작성한다.',
    ].join('\n')
    : '';

  return truncate_([
    '[행발 추천 문구 생성 요청]',
    '아래 입력을 바탕으로 행동특성 및 종합의견에 사용할 추천 문구를 작성한다.',
    '',
    '[입력]',
    inputLines,
    '',
    behaviorPrompt || defaultBehaviorRecordPrompt_(),
    traitInstruction,
    buildBehaviorReframeInstruction_(),
    charInstruction,
    retryInstruction,
    '',
    '[응답 형식]',
    BEHAVIOR_RESPONSE_FORMAT,
  ].join('\n'), maxInputChars || 30000);
}

function buildBehaviorTraitInstruction_() {
  return [
    '[특성 키워드 사용 원칙]',
    'selectedTraits와 customTraits는 교사가 고른 특성이므로 학생의 성품이나 태도로 서술할 수 있다.',
    '모두 넣을 필요는 없고, 문단 흐름에 어울리는 것을 골라 쓴다.',
    '특성에 맞추어 입력에 없는 구체적인 일화를 지어내지 않는다.',
    '근거가 부족하다는 말은 recommended_record에 쓰지 않고, 필요하면 caution에만 쓴다.',
  ].join('\n');
}

function buildBehaviorCharInstruction_(charRange) {
  if (!charRange || !charRange.enabled) return '';
  return buildByteRangeInstruction_('recommended_record', charRange);
}

function formatBehaviorCharRange_(charRange) {
  if (!charRange || !charRange.enabled) return '제한 없음';
  return formatByteRange_(charRange);
}

function generateBehaviorRecommendationWithCharRetry_(options) {
  const call = retryContext => {
    const userPrompt = buildBehaviorRecommendationPrompt_(
      options.data,
      options.maxInputChars,
      options.behaviorPrompt,
      options.charRange,
      retryContext
    );
    const result = callAi_(
      options.ai.provider,
      options.ai.model,
      options.systemGuide,
      userPrompt,
      options.ai.reasoning,
      BEHAVIOR_RESPONSE_SCHEMA
    );
    return readBehaviorGenerated_(result.text);
  };

  const first = call(null);
  if (!needsBehaviorCharRetry_(first, options.charRange)) {
    if (isBehaviorCharOutOfRange_(first.charCount, options.charRange)) {
      first.caution = appendBehaviorCaution_(first.caution, nearByteRangeNote_(first.charCount, options.charRange.min, options.charRange.max));
    }
    return first;
  }

  Utilities.sleep(800);
  try {
    return settleBehaviorCharRetry_(first, call({ charCount: first.charCount }), '', options.charRange);
  } catch (err) {
    return settleBehaviorCharRetry_(first, null, String(err.message || err), options.charRange);
  }
}

function readBehaviorGenerated_(text) {
  const parsed = parseBehaviorRecommendationJson_(text);
  const recommendedRecord = normalizeAiText_(parsed.recommended_record);

  if (!recommendedRecord) {
    throw makeBehaviorResponseFormatError_('recommended_record 값이 비어 있습니다.');
  }

  return {
    recommendedRecord,
    caution: String(parsed.caution || '').trim(),
    charCount: neisByteCount_(recommendedRecord),
    reframes: parsed.reframes,
  };
}

function needsBehaviorCharRetry_(generated, charRange) {
  return !!(charRange && charRange.enabled &&
    isBehaviorCharOutOfRange_(generated.charCount, charRange) &&
    !isNearByteRange_(generated.charCount, charRange.min, charRange.max));
}

// 분량 재요청 결과를 정리한다. 범위를 벗어나도 결과는 버리지 않고, 벗어났다는 말을 caution에 붙인다.
// retried가 없으면(재요청 실패) 첫 답을 남긴다.
function settleBehaviorCharRetry_(first, retried, retryErrorMessage, charRange) {
  if (retried && !isBehaviorCharOutOfRange_(retried.charCount, charRange)) {
    retried.caution = appendBehaviorCaution_(retried.caution, `분량 범위 재요청 후 ${retried.charCount}바이트로 생성됨.`);
    return retried;
  }

  const fallback = retried || first;
  const rangeMessage = `설정 분량 범위(${formatBehaviorCharRange_(charRange)})를 벗어남: ${fallback.charCount}바이트.`;
  const retryMessage = retryErrorMessage
    ? `재요청 응답 처리 실패: ${retryErrorMessage}`
    : '재요청 후에도 범위에 들어오지 않음.';
  fallback.caution = appendBehaviorCaution_(fallback.caution, `${rangeMessage} ${retryMessage}`);
  return fallback;
}

function isBehaviorCharOutOfRange_(charCount, charRange) {
  if (!charRange || !charRange.enabled) return false;
  if (charRange.min && charCount < charRange.min) return true;
  if (charRange.max && charCount > charRange.max) return true;
  return false;
}

function appendBehaviorCaution_(caution, message) {
  const existing = String(caution || '').trim();
  const extra = String(message || '').trim();
  if (!existing) return extra;
  if (!extra) return existing;
  return `${existing}\n${extra}`;
}

function parseBehaviorRecommendationJson_(text) {
  const raw = String(text || '').trim();
  let parsed = null;

  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    const match = raw.match(/\{[\s\S]*\}/);
    if (!match) {
      throw makeBehaviorResponseFormatError_('JSON 객체를 찾을 수 없습니다.');
    }

    try {
      parsed = JSON.parse(match[0]);
    } catch (parseErr) {
      throw makeBehaviorResponseFormatError_('JSON 파싱에 실패했습니다.');
    }
  }

  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw makeBehaviorResponseFormatError_('JSON 객체 형식이 아닙니다.');
  }

  const recommendedRecord = parsed.recommended_record !== undefined
    ? parsed.recommended_record
    : parsed.record_draft;

  if (typeof recommendedRecord !== 'string') {
    throw makeBehaviorResponseFormatError_('recommended_record 키가 없거나 문자열이 아닙니다.');
  }

  return {
    recommended_record: recommendedRecord,
    caution: parsed.caution === undefined ? '' : String(parsed.caution || ''),
    reframes: normalizeBehaviorReframes_(parsed.reframes),
  };
}

// 모델마다 형식이 조금씩 달라도 [{trait, candidates[], chosen}] 모양으로 맞춘다. 형식이 틀리면 빈 배열로 두고 생성은 계속한다.
function normalizeBehaviorReframes_(reframes) {
  if (!Array.isArray(reframes)) return [];

  return reframes
    .filter(item => item && typeof item === 'object')
    .map(item => ({
      trait: String(item.trait || '').trim(),
      candidates: (Array.isArray(item.candidates) ? item.candidates : [])
        .map(candidate => String(candidate || '').trim())
        .filter(Boolean),
      chosen: String(item.chosen || '').trim(),
    }))
    .filter(item => item.trait || item.candidates.length);
}

function formatBehaviorReframes_(reframes) {
  if (!reframes || !reframes.length) return '';

  return reframes.map(item => [
    `- ${item.trait || '보완점'}`,
    ...item.candidates.map(candidate => `  · ${candidate}${candidate === item.chosen ? ' (사용)' : ''}`),
  ].join('\n')).join('\n');
}

function makeBehaviorResponseFormatError_(message) {
  const err = new Error(message);
  err.name = 'BehaviorResponseFormatError';
  return err;
}

function isBehaviorResponseFormatError_(err) {
  return err && err.name === 'BehaviorResponseFormatError';
}

function writeBehaviorRecommendationResult_(sh, h, row, recommendedRecord, status, note, reframes) {
  const record = String(recommendedRecord || '').trim();
  const byteCount = record ? neisByteCount_(record) : '';
  const cleanNote = String(note || '').trim();
  const reframeText = formatBehaviorReframes_(reframes);
  const noteText = [
    cleanNote
      ? status === '생성완료'
        ? `[caution]\n${cleanNote}`
        : `[${status}]\n${cleanNote}`
      : '',
    reframeText ? `[돌려 쓴 표현 후보]\n${reframeText}` : '',
  ].filter(Boolean).join('\n\n');

  sh.getRange(row, h.recommendedRecord, 1, 3).setValues([[
    record,
    byteCount,
    status,
  ]]);
  sh.getRange(row, h.recommendedRecord).setNote(noteText);
}

function writeBehaviorStatusRows_(sh, statusCol, items) {
  if (!statusCol || !items || items.length === 0) return;

  const sorted = items
    .filter(item => item && item.row && item.status)
    .sort((a, b) => a.row - b.row);

  let startRow = 0;
  let values = [];

  const flush = () => {
    if (!values.length) return;
    sh.getRange(startRow, statusCol, values.length, 1).setValues(values);
    startRow = 0;
    values = [];
  };

  sorted.forEach(item => {
    const expectedRow = startRow ? startRow + values.length : item.row;

    if (startRow && item.row !== expectedRow) {
      flush();
    }

    if (!startRow) startRow = item.row;
    values.push([item.status]);
  });

  flush();
}

function buildBehaviorRecommendationResultMessage_(result) {
  const lines = [
    '행발 추천 문구 생성 결과',
    '',
    `처리 범위: ${result.scopeLabel || '선택 행'}`,
    `생성 완료: ${result.processed || 0}행`,
    `실패: ${result.failed || 0}행`,
    `스킵: ${result.skipped || 0}행`,
    `대기: ${result.remaining || 0}행`,
    `생성 후보: ${(result.candidateRows || []).length}행`,
  ];

  if (result.skipped) {
    lines.push(
      '',
      `스킵 사유 - 이미 생성됨: ${result.skippedAlreadyGenerated || 0}행`,
      `스킵 사유 - 특성없음: ${result.skippedNoTraits || 0}행`,
      `스킵 사유 - 빈 행: ${result.skippedBlank || 0}행`
    );
  }

  if (result.candidateRows && result.candidateRows.length > 0) {
    lines.push('', `생성 후보 행: ${result.candidateRows.slice(0, 12).join(', ')}${result.candidateRows.length > 12 ? '...' : ''}`);
  }

  if (result.skippedDetails && result.skippedDetails.length > 0) {
    lines.push('', `스킵 상세: ${result.skippedDetails.join(', ')}`);
  }

  if (result.blocked) {
    lines.push('', `중단 사유: ${result.blockingReason || '확인 필요'}`);
    lines.push(result.blockingMessage || '더 이상 생성을 계속할 수 없어 중단했습니다.');
  }

  return lines.join('\n');
}

function logBehaviorRecommendationResult_(action, result) {
  log_(
    action,
    result.blocked ? 'BLOCKED' : 'OK',
    `범위 ${result.scopeLabel || '대기 행'}, 처리 ${result.processed}행, 실패 ${result.failed}행, 스킵 ${result.skipped}행` +
      ` (이미생성 ${result.skippedAlreadyGenerated || 0}, 특성없음 ${result.skippedNoTraits || 0}, 빈행 ${result.skippedBlank || 0}), ` +
      `후보 ${(result.candidateRows || []).join(',') || '없음'}, 스킵상세 ${(result.skippedDetails || []).join(', ') || '없음'}, 남은 대기 ${result.remaining}행`
  );
}

function autoProcessPendingBehaviorRecommendations() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    rescheduleAutoJobWhenBusy_('autoProcessPendingBehaviorRecommendations', AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, scheduleNextAutoBehaviorRecommendationTrigger_);
    return;
  }

  try {
    deleteAutoBehaviorRecommendationTriggers_();

    const status = PropertiesService.getUserProperties().getProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY) || 'OFF';
    if (status !== 'ON') {
      log_('autoProcessPendingBehaviorRecommendations', 'STOP', '자동 생성 상태가 OFF라 실행하지 않음');
      return;
    }

    beginAutoRun_(scheduleNextAutoBehaviorRecommendationTrigger_, [SHEET_NAMES.behaviorRecommendations, 'status']);

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.behaviorRecommendations);
    if (!sh) {
      PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
      log_('autoProcessPendingBehaviorRecommendations', 'STOP', '행발문구추천 시트를 찾을 수 없음');
      return;
    }

    const lastRow = sh.getLastRow();
    const h = headerMap_(sh);
    const config = getConfigMap_();
    const result = processBehaviorRecommendationRows_(
      sh,
      h,
      2,
      Math.max(lastRow - 1, 0),
      config,
      { pendingOnly: true }
    );
    result.scopeLabel = '대기/재시도필요 행';

    logBehaviorRecommendationResult_('autoProcessPendingBehaviorRecommendations', result);

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
      toastAutoRecordDraft_(result.blockingMessage || '행발 추천 문구 생성을 중단했습니다.', result.blockingReason || '행발 생성 중단');
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
      toastAutoRecordDraft_(
        `생성 완료 ${result.processed}행, 실패 ${result.failed}행, 스킵 ${result.skipped}행`,
        '행발 추천 문구 생성 완료'
      );
      return;
    }

    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);
    scheduleNextAutoBehaviorRecommendationTrigger_(nextDelayMs);
  } catch (err) {
    const message = String(err.message || err);
    PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
    deleteAutoBehaviorRecommendationTriggers_();
    toastAutoRecordDraft_(message, '행발 추천 문구 자동 처리 오류');
    log_('autoProcessPendingBehaviorRecommendations', 'ERROR', message);
  } finally {
    endAutoRun_(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, deleteAutoBehaviorRecommendationTriggers_);
    lock.releaseLock();
  }
}

function countPendingBehaviorRecommendationRows_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.behaviorRecommendations);
    if (!sh) return 0;

    const h = headerMap_(sh);
    if (!h.selectedTraits || !h.positiveObservation || !h.improvementNeeded || !h.growthEvidence || !h.recommendedRecord || !h.status) {
      return 0;
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return 0;

    const lastCol = Math.max(sh.getLastColumn(), BEHAVIOR_RECOMMENDATION_HEADERS.length);
    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let count = 0;

    values.forEach(row => {
      const status = String(rowValueAt_(row, h.status) || '').trim();
      if (status !== '대기' && status !== '재시도필요') return;

      const recommendedRecord = String(rowValueAt_(row, h.recommendedRecord) || '').trim();
      if (recommendedRecord) return;

      // 관찰 내용은 모두 선택 입력이라 특성(목록 또는 직접 적은 것)만 있으면 만든다.
      const traits = [rowValueAt_(row, h.selectedTraits), h.customTraits ? rowValueAt_(row, h.customTraits) : ''];
      if (traits.some(isMeaningfulBehaviorInput_)) count++;
    });

    return count;
  } catch (err) {
    return 0;
  }
}

function scheduleNextAutoBehaviorRecommendationTrigger_(delayMs) {
  deleteAutoBehaviorRecommendationTriggers_();

  ScriptApp.newTrigger('autoProcessPendingBehaviorRecommendations')
    .timeBased()
    .after(delayMs)
    .create();

  log_(
    'scheduleNextAutoBehaviorRecommendationTrigger_',
    'OK',
    `${Math.round(delayMs / 1000)}초 뒤 다음 행발 추천 문구 자동 생성 예약`
  );
}

function deleteAutoBehaviorRecommendationTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoProcessPendingBehaviorRecommendations') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function setupSheets() {
  Object.keys(HEADERS).forEach(sheetName => {
    const sh = ensureSheetHeaders_(sheetName);
    const headers = HEADERS[sheetName];

    // 생기부초안, 학생별생기부는 전용 서식이 있으므로 자동 너비 조정 제외
    if (
      sheetName !== SHEET_NAMES.records &&
      sheetName !== SHEET_NAMES.studentFinalRecords
    ) {
      sh.autoResizeColumns(1, headers.length);
    }
  });

  const config = getSheet_(SHEET_NAMES.config);
  migrateLegacyConfigKeys_(config);
  const existing = getConfigMap_();

  const defaults = [
    ['AI_PROVIDER', 'claude', 'claude 또는 gpt 또는 gemini'],
    ['CLAUDE_MODEL', 'claude-haiku-5-5', 'claude-haiku-5-5 / claude-sonnet-5-5\nclaude-haiku-5-5 - 기본값. 값이 싸고 대답이 빠름. 생기부 문장 쓰기에 충분함\nclaude-sonnet-5-5 - 상대적으로 비쌈. 조금 느리지만 더 똑똑한 편'],
    ['GPT_MODEL', 'gpt-6-luna', 'gpt-6-luna / gpt-6.1-terra\ngpt-6.1-terra / 상위 고성능 모델\ngpt-6-luna / 무난하게 쓸 수 있는 기본 모델'],
    ['GEMINI_MODEL', 'gemini-3.8-flash', 'gemini-3.8-flash / gemini-3.7-flash / gemini-3.6-flash / gemini-3.5-flash'],
    ['REASONING', 'none', '추론 수준입니다\nlow, medium을 사용하면 글을 조금 더 생각하고 써줍니다\n다만, none이 아니라 low를 사용하면 비용을 대략 2배정도 사용하는 것 같습니다'],
    ['MAX_RECORD_CHARS', '500', '생기부 초안 최대 글자수'],
    ['MAX_INPUT_CHARS', '30000', 'Claude에 보낼 과제 원문 최대 글자수'],
    ['RECORD_DRAFT_MIN_BYTES', '900', '초안 만들기 최소 바이트(나이스 기준)\n도우미의 [초안 만들기] 화면에서 바꿀 수 있습니다.'],
    ['RECORD_DRAFT_MAX_BYTES', '1200', '초안 만들기 최대 바이트(나이스 기준)'],
    ['FINAL_RECORD_MIN_BYTES', '900', '초안 합치기 최소 바이트(나이스 기준)\n도우미의 [초안 합치기] 화면에서 바꿀 수 있습니다.'],
    ['FINAL_RECORD_MAX_BYTES', '1200', '초안 합치기 최대 바이트(나이스 기준)'],
    ['MANUAL_RECORD_MIN_BYTES', '900', '수동추가 최소 바이트(나이스 기준)\n도우미의 [수동추가] 화면에서 바꿀 수 있습니다.'],
    ['MANUAL_RECORD_MAX_BYTES', '1200', '수동추가 최대 바이트(나이스 기준)'],
    ['NEIS_MAX_BYTES', '1500', ''],
    ['NEIS_WARNING_BYTES', '1400', ''],
    ['MAX_RUN_SECONDS', '270', 'Apps Script 6분 제한 전에 안전 중단할 시간'],
    ['AUTO_NEXT_DELAY_MS', '30000', ''],
    ['FAST_GENERATION', 'OFF', '빠른 생성(실험)\nON이면 학생 3명까지 AI에 한꺼번에 요청해 더 빨리 만듭니다.\n요청 한도(429)나 서버 오류가 나면 30분 동안 자동으로 한 명씩 보내는 방식으로 돌아갑니다.\nAPI 키를 처음 만들어 사용량 등급이 낮다면 켜지 않는 것을 권장합니다.'],
    ['COMMON_PHRASE_PREVIOUS_LIMIT', String(COMMON_PHRASE_DEFAULT_PREVIOUS_LIMIT), '공통문구 생성 시 AI 프롬프트에 포함할 최근 생성 문구 수\n기본 10개. 입력 토큰 절감을 위해 너무 크게 설정하지 않는 것을 권장합니다.'],
    ['PROMPT_1', defaultRecordDraftPrompt_(), '생기부초안 만들 때 지시 또는 강조 사항(없으면 SYSTEM_GUIDE를 따라감)'],
    ['PROMPT_2', defaultStudentFinalRecordPrompt_(), '학생별생기부 만들 때 지시 또는 강조 사항(없으면 SYSTEM_GUIDE를 따라감)'],
    ['MANUAL_START_ROW', '2', '수동추가를 위한 설정\n학생 데이터가 시작하는 행번호(ex. 2)'],
    ['MANUAL_INPUT_COLS', '', '수동추가를 위한 설정\n인공지능에게 줄 학생 데이터가 적힌 열(ex. B,C) /  콤마로 구분하여 입력 가능\n이름과 같은 개인정보를 포함하지 않게 주의해주세요'],
    ['MANUAL_OUTPUT_COL', '', '수동추가를 위한 설정\n인공지능에게 받아올 생기부를 적을 열(ex. C)'],
    ['MANUAL_PROMPT', defaultManualRecordPrompt_(), '수동추가를 위한 생기부 작성 지침\n\n\n인공지능에게 채점도 시켜보고 싶으시면 채점하라고 지시하시면 됩니다.\n그리고 "evidence_summary에 점수를 입력하여 보낸다" 라는 프롬프트를 \n작업에 넣어주시면 메모로 받아보실 수 있습니다.'],
    ['BEHAVIOR_RECORD_MIN_BYTES', '750', '행발 최소 바이트(나이스 기준)\n도우미의 [행발 추천] 화면에서 바꿀 수 있습니다.'],
    ['BEHAVIOR_RECORD_MAX_BYTES', '1200', '행발 최대 바이트(나이스 기준)'],
    ['BEHAVIOR_RECORD_PROMPT', defaultBehaviorRecordPrompt_(), '작성하시는 스타일에 맞게 수정하시면 됩니다.'],
    ['SYSTEM_GUIDE', defaultSystemGuide_(), '공통 작성 지침\n초안, 초안 합치기, 수동추가에 모두 적용되는 원칙입니다.\n과목에 상관없는 원칙(근거, 문체)만 두는 것을 권장합니다.\n\n과제마다 달라지는 내용(예: 첫 문장을 "인공지능과 미래사회에 대한 글쓰기 수행평가에서"로 시작)은\n여기가 아니라 각 프롬프트 맨 위의 "작성 순서/형식(선택)" 줄에 적어 주세요.']
  ];

  const rows = defaults.filter(r => !existing[r[0]]);
  if (rows.length > 0) {
    config.getRange(config.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  }

  applyConfigValidations_(config);
  syncConfigDefaults_(config, defaults);
  styleConfigSheet_();

  ensureBehaviorRecommendationSheet_();
  styleCommonPhraseSheet_();
  styleBehaviorRecommendationSheet_();
  styleRecordsSheet_();
  styleStudentFinalRecordsSheet_();
  organizeSheetTabs_();

  log_('setupSheets', 'OK', '시트 초기 설정 완료');
}

function saveClaudeApiKey() {
  const ui = getUi_();
  const res = ui.prompt(
    'Claude API 키 저장',
    '본인 Claude API 키를 입력하세요. 이 키는 현재 사용자 계정의 User Properties에만 저장됩니다.',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const key = res.getResponseText().trim();
  if (!key) {
    ui.alert('API 키가 비어 있습니다.');
    return;
  }

  PropertiesService.getUserProperties().setProperty('CLAUDE_API_KEY', key);

  log_('saveClaudeApiKey', 'OK', '현재 사용자 Claude API 키 저장 완료');
  ui.alert('Claude API 키를 현재 사용자 계정에 저장했습니다.');
}

function importClassroomAssignments() {
  const result = importClassroomAssignmentsCore_();

  getUi_().alert(`${result.assignmentCount}개 과제를 가져왔습니다.`);
}

function importClassroomAssignmentsCore_() {
  const sh = ensureAssignmentsSheetForImport_();

  const existingOptions = getAssignmentOptionMap_();
  const courses = listAllCourses_();
  const rows = [];

  courses.forEach(course => {
    try {
      const works = listAllCourseWork_(course.id);

      works.forEach(w => {
        const key = `${course.id}|${w.id}`;
        const option = existingOptions[key] || {};


      rows.push([
        course.id,
        course.name || '',
        course.section || '',
        w.id,
        w.title || '',
        w.workType || '',
        w.state || '',
        formatDue_(w.dueDate, w.dueTime),
        w.updateTime || '',
        w.alternateLink || '',
        w.description || '',
        w.maxPoints ?? '',
        '',
        option.includeInFinal === true ? true : false,
        option.weight || 1,
        option.category || '',
        option.collectStatus || '',
        option.memo || '',
      ]);

      });
    } catch (err) {
      log_('importClassroomAssignments', 'ERROR', `${course.name}: ${err.message}`);
    }
  });

  clearData_(sh);

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, HEADERS[SHEET_NAMES.assignments].length).setValues(rows);
  }

  styleAssignmentsSheet_();

  log_('importClassroomAssignments', 'OK', `${rows.length}개 과제 가져옴`);

  return {
    assignmentCount: rows.length,
  };
}

function ensureAssignmentsSheetForImport_() {
  return ensureSheetHeaders_(SHEET_NAMES.assignments);
}


function getAssignmentOptionMap_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.assignments);
  const map = {};

  if (!sh || sh.getLastRow() < 2) return map;

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const idx = headerIndex_(headers);

  if (idx.courseId === undefined || idx.courseWorkId === undefined) return map;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const courseId = String(row[idx.courseId] || '').trim();
    const courseWorkId = String(row[idx.courseWorkId] || '').trim();

    if (!courseId || !courseWorkId) continue;

    const key = `${courseId}|${courseWorkId}`;

    map[key] = {
      includeInFinal: idx.includeInFinal !== undefined ? row[idx.includeInFinal] === true || String(row[idx.includeInFinal]).toUpperCase() === 'TRUE' : false,
      weight: idx.weight !== undefined ? row[idx.weight] : 1,
      category: idx.category !== undefined ? row[idx.category] : '',
      collectStatus: idx.collectStatus !== undefined ? row[idx.collectStatus] : '',
      memo: idx.memo !== undefined ? row[idx.memo] : '',
    };
  }

  return map;
}

function collectSubmissionsForSelectedAssignment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  if (sh.getName() !== SHEET_NAMES.assignments) {
    getUi_().alert('과제목록 시트에서 과제 행 하나를 선택한 뒤 실행하세요.');
    return;
  }

  const row = sh.getActiveRange().getRow();
  if (row <= 1) {
    getUi_().alert('헤더가 아닌 과제 행을 선택하세요.');
    return;
  }

  const h = headerMap_(sh);
  const courseId = valueAt_(sh, row, h.courseId);
  const courseName = valueAt_(sh, row, h.courseName);
  const courseWorkId = valueAt_(sh, row, h.courseWorkId);
  const assignmentTitle = valueAt_(sh, row, h.assignmentTitle);
  const maxPoints = h.maxPoints ? valueAt_(sh, row, h.maxPoints) : '';
  const assignmentDescription = valueAt_(sh, row, h.assignmentDescription);

  if (!isMyTeacherCourse_(courseId)) {
    getUi_().alert(
      '안전 차단: 이 수업은 현재 계정이 교사로 등록된 수업이 아닙니다.\n' +
      '관리자 권한으로 보이는 다른 교사의 수업일 가능성이 있어 제출물 수집을 중단합니다.'
    );
    log_('collectSubmissionsForSelectedAssignment', 'BLOCKED', `내 교사 수업 아님: ${courseId} / ${assignmentTitle}`);
    return;
  }  

  if (!courseId || !courseWorkId) {
    getUi_().alert('courseId 또는 courseWorkId가 비어 있습니다.');
    return;
  }

  const submissions = listAllStudentSubmissions_(courseId, courseWorkId);
  const studentLookupIndex = buildStudentLookupIndex_();

  const rows = [];

  submissions.forEach(sub => {
    try {
      const profile = getUserProfile_(sub.userId);
      const email = profile.emailAddress || '';
      const name = profile.name && profile.name.fullName ? profile.name.fullName : '';
      const studentInfo = lookupStudent_(email, name, sub.userId, studentLookupIndex);
      const extracted = extractSubmissionText_(sub);

      const submissionKey = `${courseId}|${courseWorkId}|${sub.id}`;

      rows.push([
        submissionKey,
        studentInfo.studentNo || '',
        studentInfo.name || name,
        studentInfo.email || email,
        courseId,
        courseName,
        courseWorkId,
        assignmentTitle,
        sub.id,
        sub.state || '',
        sub.late === true ? 'TRUE' : 'FALSE',
        sub.draftGrade ?? '',
        sub.assignedGrade ?? '',
        maxPoints,
        '',
        sub.updateTime || '',
        sub.alternateLink || '',
        extracted.fileTitles.join('\n'),
        extracted.fileLinks.join('\n'),
        extracted.text,
        extracted.text.length,
        new Date(),
        '',
        assignmentDescription,
      ]);
    } catch (err) {
      log_('collectSubmissionsForSelectedAssignment', 'ERROR', `${sub.id}: ${err.message}`);
    }
  });

  ensureSheetHeaders_(SHEET_NAMES.submissions);
  upsertByKey_(SHEET_NAMES.submissions, 'submissionKey', rows);

  log_('collectSubmissionsForSelectedAssignment', 'OK', `${rows.length}개 제출물 수집`);
  getUi_().alert(`${rows.length}개 제출물을 수집했습니다.`);
}

function generateRecordsForSelectedSubmissions() {
  const ui = getUi_();
  const selection = getSelectedSubmissionRows_();

  if (selection.error) {
    ui.alert(selection.error);
    return;
  }

  const pendingCount = markSubmissionRowsPending_(selection.sh, selection.h, selection.rows);

  log_(
    'generateRecordsForSelectedSubmissions',
    'QUEUE',
    `${pendingCount}개 선택 제출물을 대기열에 넣고 자동 생성을 시작`
  );
  refreshDashboard();

  startAutoRecordDraftTrigger();
}


function extractSubmissionText_(sub) {
  const parts = [];
  const fileTitles = [];
  const fileLinks = [];

  if (sub.shortAnswerSubmission && sub.shortAnswerSubmission.answer) {
    parts.push(`[단답형 응답]\n${sub.shortAnswerSubmission.answer}`);
  }

  if (sub.multipleChoiceSubmission && sub.multipleChoiceSubmission.answer) {
    parts.push(`[선택형 응답]\n${sub.multipleChoiceSubmission.answer}`);
  }

  const attachments = sub.assignmentSubmission && sub.assignmentSubmission.attachments
    ? sub.assignmentSubmission.attachments
    : [];

  attachments.forEach(att => {
    if (att.driveFile) {
      const fileId = att.driveFile.id;
      const title = att.driveFile.title || fileId;
      const link = att.driveFile.alternateLink || '';

      fileTitles.push(title);
      fileLinks.push(link);

      try {
        const file = DriveApp.getFileById(fileId);
        const mime = file.getMimeType();
        let text = '';

        if (mime === MimeType.GOOGLE_DOCS || mime === 'application/vnd.google-apps.document') {
          text = DocumentApp.openById(fileId).getBody().getText();
        } else if (mime === MimeType.PLAIN_TEXT || mime === 'text/plain') {
          text = file.getBlob().getDataAsString('UTF-8');
        } else {
          text = `[텍스트 자동 추출 미지원 파일 형식: ${mime}]\n파일 링크를 직접 확인하세요: ${link}`;
        }

        parts.push(`[첨부파일: ${title}]\n${truncate_(text, 40000)}`);
      } catch (err) {
        parts.push(`[첨부파일: ${title}]\n파일을 열 수 없음: ${err.message}\n${link}`);
      }
    } else if (att.link) {
      const title = att.link.title || att.link.url || 'link';
      fileTitles.push(title);
      fileLinks.push(att.link.url || '');
      parts.push(`[링크 제출]\n${title}\n${att.link.url || ''}`);
    } else if (att.form) {
      const title = att.form.title || 'Google Form';
      fileTitles.push(title);
      fileLinks.push(att.form.formUrl || '');
      parts.push(`[Google Form 첨부]\n${title}\n폼 응답 원문은 Classroom 제출물에서 직접 추출되지 않을 수 있습니다.`);
    }
  });

  return {
    text: parts.join('\n\n---\n\n'),
    fileTitles,
    fileLinks,
  };
}

function buildClaudePayload_(model, systemGuide, userPrompt, reasoning) {
  // temperature는 보내지 않는다. Claude Haiku 5.5·Sonnet 5.5는 기본값이 아닌 temperature를 400으로 거절한다.
  const payload = {
    model,
    max_tokens: getClaudeMaxTokens_(reasoning),
    system: systemGuide,
    messages: [
      {
        role: 'user',
        content: userPrompt,
      },
    ],
  };

  applyClaudeReasoning_(payload, model, reasoning);
  return payload;
}

function callClaude_(model, systemGuide, userPrompt, reasoning) {
  const apiKey = PropertiesService.getUserProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    throw new Error('Claude API 키가 없습니다. 메뉴에서 본인 Claude API 키를 먼저 저장하세요.');
  }

  const payload = buildClaudePayload_(model, systemGuide, userPrompt, reasoning);

  const maxAttempts = 5;
  let outputTokenRetryUsed = false;
  let lastErrorMessage = '';

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let res;

    try {
      res = fetchAiWithinBudget_(AI_API_ENDPOINTS.claude.url, buildAiFetchOptions_('claude', apiKey, payload));
    } catch (err) {
      // 응답 없이 끊긴 경우(시간 초과, 연결 오류)도 잠시 뒤 다시 시도한다.
      if (isAiTimeBudgetError_(err) || attempt === maxAttempts) throw err;
      lastErrorMessage = `Claude API 네트워크 오류: ${String(err.message || err)}`;
      log_('callClaude_', 'RETRY', `${lastErrorMessage}. ${attempt}/${maxAttempts}회 시도 실패`);
      sleepWithinAiBudget_(Math.min(2000 * Math.pow(2, attempt - 1), 30000));
      continue;
    }

    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);

      // 출력 한도에 걸려 잘린 답은 JSON이 끊겨 있으므로, 한도를 늘려 한 번 더 받는다.
      if (json.stop_reason === 'max_tokens' && !outputTokenRetryUsed) {
        outputTokenRetryUsed = true;
        payload.max_tokens = Math.min(Number(payload.max_tokens) * 2, 16000);
        log_('callClaude_', 'RETRY', `Claude 응답이 max_tokens로 잘려 ${payload.max_tokens} 토큰으로 1회 재시도합니다.`);
        sleepWithinAiBudget_(1200);
        continue;
      }

      if (json.stop_reason === 'max_tokens') {
        throw new Error(`Claude 응답이 출력 한도(max_tokens ${payload.max_tokens})에서 잘렸습니다. 다시 만들어 보세요.`);
      }

      const text = extractClaudeOutputText_(json);

      return {
        text,
        usage: json.usage || {},
      };
    }

    lastErrorMessage = `Claude API 오류 ${code}: ${body}`;

    if (isRetryableClaudeError_(code)) {
      const waitMs = getClaudeRetryDelayMs_(res, attempt);

      log_(
        'callClaude_',
        'RETRY',
        `Claude API ${code} 오류. ${attempt}/${maxAttempts}회 시도 실패. ${Math.round(waitMs / 1000)}초 후 재시도`
      );

      sleepWithinAiBudget_(waitMs);
      continue;
    }

    throw new Error(lastErrorMessage);
  }

  throw new Error(`Claude API 재시도 실패: ${lastErrorMessage}`);
}

// Apps Script 한 번 실행은 6분까지다. 전역 코드는 실행마다 다시 평가되므로 이 값이 이번 실행의 시작 시각이다.
// 남은 시간이 AI 요청 한 번(최대 약 1분)에 모자라면 요청이나 재시도 대기를 시작하지 않고 멈춰서,
// 강제 종료로 '생성중'이 남거나 다음 실행 예약이 사라지지 않게 한다.
const AI_EXECUTION_STARTED_AT = Date.now();
const AI_EXECUTION_LIMIT_MS = 345 * 1000;
const AI_FETCH_RESERVE_MS = 70 * 1000;

// 생성 루프가 한 번 실행에서 처리하는 최대 수. 예전에는 BATCH_SIZE(기본 5)만큼 만들고 다음 실행을 기다려
// 쓸 수 있는 시간의 절반도 못 쓰고 쉬었다. 이제는 MAX_RUN_SECONDS와 위의 시간 예산까지 이어서 만들고,
// 이 값은 만일을 위한 상한으로만 둔다.
const MAX_ITEMS_PER_RUN = 100;

// [시작]을 눌렀을 때 바로 도는 첫 실행은 사용자가 화면에서 기다리므로 예전처럼 몇 명만 만들고 돌려준다.
// 나머지는 뒤에서 도는 자동 실행이 시간 한도까지 이어서 만든다.
const FIRST_RUN_MAX_ITEMS = 5;
let RUN_ITEM_LIMIT_ = MAX_ITEMS_PER_RUN;

function runFirstBatchQuickly_(run) {
  RUN_ITEM_LIMIT_ = FIRST_RUN_MAX_ITEMS;
  try {
    return run();
  } finally {
    RUN_ITEM_LIMIT_ = MAX_ITEMS_PER_RUN;
  }
}

function ensureAiTimeBudget_(extraMs) {
  const remainingMs = AI_EXECUTION_STARTED_AT + AI_EXECUTION_LIMIT_MS - Date.now();

  if (remainingMs - (extraMs || 0) < AI_FETCH_RESERVE_MS) {
    const err = new Error('실행 시간 제한(6분)에 가까워 이번 실행에서는 AI 요청을 보내지 않았습니다. 다음 실행에서 이어서 만듭니다.');
    err.aiTimeBudget = true;
    throw err;
  }
}

function isAiTimeBudgetError_(err) {
  return !!(err && err.aiTimeBudget);
}

function sleepWithinAiBudget_(ms) {
  ensureAiTimeBudget_(ms);
  Utilities.sleep(ms);
}

function fetchAiWithinBudget_(url, options) {
  ensureAiTimeBudget_(0);
  return UrlFetchApp.fetch(url, options);
}

const AI_API_ENDPOINTS = {
  claude: { url: 'https://api.anthropic.com/v1/messages', keyProperty: 'CLAUDE_API_KEY' },
  gpt: { url: 'https://api.openai.com/v1/responses', keyProperty: 'OPENAI_API_KEY' },
  gemini: { url: 'https://generativelanguage.googleapis.com/v1beta/interactions', keyProperty: 'GEMINI_API_KEY' },
};

function buildAiFetchOptions_(provider, apiKey, payload) {
  const headers = {
    claude: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
    gpt: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
    gemini: { 'x-goog-api-key': apiKey, 'Content-Type': 'application/json' },
  }[provider];

  return {
    method: 'post',
    muteHttpExceptions: true,
    headers,
    payload: JSON.stringify(payload),
  };
}

function buildAiPayload_(provider, model, systemGuide, userPrompt, reasoning, responseSchema) {
  if (provider === 'gpt') return buildOpenAiPayload_(model, systemGuide, userPrompt, reasoning, responseSchema);
  if (provider === 'gemini') return buildGeminiPayload_(model, systemGuide, userPrompt, reasoning, responseSchema);
  return buildClaudePayload_(model, systemGuide, userPrompt, reasoning);
}

// 성공 응답에서 글을 꺼낸다. 출력 한도로 잘린 답이면 null을 돌려준다(그 요청은 하나씩 보내는 경로로 다시 받는다).
function readAiSuccessResponse_(provider, json) {
  if (provider === 'gpt') {
    if (isOpenAiMaxOutputTokenIncomplete_(json)) return null;
    return { text: extractOpenAiOutputText_(json), usage: json.usage || {} };
  }

  if (provider === 'gemini') {
    if (isGeminiOutputTruncated_(json)) return null;
    return { text: extractGeminiOutputText_(json), usage: json.usage || json.usage_metadata || json.usageMetadata || {} };
  }

  if (json.stop_reason === 'max_tokens') return null;
  return { text: extractClaudeOutputText_(json), usage: json.usage || {} };
}

// 여러 AI 요청을 UrlFetchApp.fetchAll로 한 번에 보낸다. requests: [{provider, model, systemGuide, userPrompt, reasoning, responseSchema}]
// 재시도는 하지 않는다. outcomes[i]는 {result} 또는 {error}이고, 오류 응답·잘린 답·연결 실패면 null이다.
// codes[i]는 받은 HTTP 상태 코드(못 받았으면 0), fetchError는 fetchAll 자체가 실패했을 때의 오류다.
function fetchAiBatch_(requests) {
  const props = PropertiesService.getUserProperties();
  const outcomes = requests.map(() => null);
  const codes = requests.map(() => 0);
  const sendable = [];

  requests.forEach((req, index) => {
    const provider = normalizeAiProvider_(req.provider);
    const endpoint = AI_API_ENDPOINTS[provider];
    const apiKey = endpoint ? props.getProperty(endpoint.keyProperty) : '';

    if (!apiKey) {
      outcomes[index] = { error: new Error(`${provider} API 키가 저장되어 있지 않습니다.`) };
      return;
    }

    const payload = buildAiPayload_(provider, req.model, req.systemGuide, req.userPrompt, normalizeReasoning_(req.reasoning), req.responseSchema);
    sendable.push({ index, provider, fetch: Object.assign({ url: endpoint.url }, buildAiFetchOptions_(provider, apiKey, payload)) });
  });

  let responses = [];
  let fetchError = null;
  if (sendable.length) {
    try {
      ensureAiTimeBudget_(0);
      responses = UrlFetchApp.fetchAll(sendable.map(item => item.fetch));
    } catch (err) {
      if (isAiTimeBudgetError_(err)) throw err;
      // 한 요청이라도 연결이 끊기면 fetchAll 전체가 실패한다.
      fetchError = err;
      responses = [];
    }
  }

  sendable.forEach((item, i) => {
    const res = responses[i];
    if (!res) return;

    const code = res.getResponseCode();
    codes[item.index] = code;
    if (code < 200 || code >= 300) return;

    try {
      const read = readAiSuccessResponse_(item.provider, JSON.parse(res.getContentText()));
      if (read) outcomes[item.index] = { result: read };
    } catch (err) {
      outcomes[item.index] = { error: err };
    }
  });

  return { outcomes, codes, fetchError, sentIndexes: sendable.map(item => item.index) };
}

// fetchAiBatch_로 한꺼번에 보낸 뒤, 오류(429 등)이거나 잘렸거나 네트워크가 끊긴 요청만 callAi_로 하나씩 다시 보내
// 기존 재시도 규칙을 그대로 따른다. 결과는 요청과 같은 순서로 {result} 또는 {error}다.
function callAiParallel_(requests) {
  const batch = fetchAiBatch_(requests);
  const outcomes = batch.outcomes;

  if (batch.fetchError) {
    log_('callAiParallel_', 'RETRY', `동시 요청 실패, 하나씩 다시 보냅니다: ${String(batch.fetchError.message || batch.fetchError)}`);
  }

  batch.sentIndexes.forEach(index => {
    if (outcomes[index]) return;

    const req = requests[index];
    try {
      outcomes[index] = { result: callAi_(req.provider, req.model, req.systemGuide, req.userPrompt, req.reasoning, req.responseSchema) };
    } catch (err) {
      outcomes[index] = { error: err };
    }
  });

  return outcomes;
}

// --- 빠른 생성(실험) -----------------------------------------------------
// 설정 FAST_GENERATION이 ON이면 다음 학생 몇 명의 첫 요청을 fetchAll로 미리 한꺼번에 받아 둔다.
// 생성 루프는 지금처럼 한 명씩 돌고, callAi_가 똑같은 요청이 오면 받아 둔 답을 꺼내 쓴다.
// 분량 재요청, 오류 판정, 중단 규칙은 모두 기존 순차 경로를 그대로 따른다.
// 안전장치:
// - 한 번에 최대 FAST_GENERATION_CONCURRENCY개까지만 받는다.
// - 미리 받기에서 오류 응답이 하나라도 오면 이번 실행은 한 명씩으로 돌아가고, 오류 응답은 순차 경로가 다시 보낸다.
// - 429나 서버 오류(5xx)면 FAST_GENERATION_PAUSE_MS 동안 다음 실행들도 한 명씩 보낸다.
// - 남은 실행 시간이 요청 한 번과 FAST_GENERATION_TAIL_MS에 모자라면 미리 받지 않는다.
const FAST_GENERATION_CONCURRENCY = 3;
const FAST_GENERATION_TAIL_MS = 30 * 1000;
const FAST_GENERATION_PAUSE_MS = 30 * 60 * 1000;
const FAST_GENERATION_PAUSED_UNTIL_PROPERTY = 'FAST_GENERATION_PAUSED_UNTIL';
let AI_PREFETCHED_ = {};
let AI_PREFETCHED_COUNT_ = 0;

function aiRequestKey_(provider, model, systemGuide, userPrompt, reasoning, responseSchema) {
  const text = JSON.stringify([
    normalizeAiProvider_(provider),
    String(model || ''),
    String(systemGuide || ''),
    String(userPrompt || ''),
    normalizeReasoning_(reasoning),
    responseSchema || null,
  ]);
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8));
}

function takePrefetchedAi_(provider, model, systemGuide, userPrompt, reasoning, responseSchema) {
  if (AI_PREFETCHED_COUNT_ < 1) return null;

  const key = aiRequestKey_(provider, model, systemGuide, userPrompt, reasoning, responseSchema);
  const result = AI_PREFETCHED_[key];
  if (!result) return null;

  delete AI_PREFETCHED_[key];
  AI_PREFETCHED_COUNT_--;
  return result;
}

function clearPrefetchedAi_() {
  AI_PREFETCHED_ = {};
  AI_PREFETCHED_COUNT_ = 0;
}

function isFastGenerationOn_(config) {
  if (String((config && config.FAST_GENERATION) || '').trim().toUpperCase() !== 'ON') return false;

  const pausedUntil = Number(PropertiesService.getUserProperties().getProperty(FAST_GENERATION_PAUSED_UNTIL_PROPERTY) || 0);
  return Date.now() >= pausedUntil;
}

// 생성 루프마다 하나 만든다. prefetch(makeRequests)는 지금 학생부터 차례로 보낼 첫 요청 목록을 만드는 함수를 받는다.
// 빠른 생성이 꺼져 있으면 목록을 만들지도 않는다. 루프가 끝나면 finish()로 남은 답을 비운다.
function createAiPrefetcher_(config, label) {
  clearPrefetchedAi_();

  let on = isFastGenerationOn_(config);
  const attempted = {};

  return {
    prefetch(makeRequests) {
      if (!on) return;

      try {
        const requests = makeRequests() || [];
        const keys = requests.map(req => aiRequestKey_(req.provider, req.model, req.systemGuide, req.userPrompt, req.reasoning, req.responseSchema));

        // 지금 학생의 요청을 이미 미리 보냈으면(받았든 못 받았든) 기다렸다가 다음 묶음을 보낸다.
        if (!requests.length || attempted[keys[0]]) return;

        const batch = [];
        const batchKeys = [];
        requests.forEach((req, i) => {
          if (batch.length >= FAST_GENERATION_CONCURRENCY || attempted[keys[i]]) return;
          batch.push(req);
          batchKeys.push(keys[i]);
        });

        if (batch.length < 2) return;

        // 실행이 끝나 갈 때 미리 받으면 쓰지 못하고 버리기 쉬우므로, 넉넉히 남았을 때만 받는다.
        ensureAiTimeBudget_(FAST_GENERATION_TAIL_MS);
        batchKeys.forEach(key => { attempted[key] = true; });

        const sent = fetchAiBatch_(batch);
        sent.outcomes.forEach((outcome, i) => {
          if (!outcome || !outcome.result || AI_PREFETCHED_[batchKeys[i]]) return;
          AI_PREFETCHED_[batchKeys[i]] = outcome.result;
          AI_PREFETCHED_COUNT_++;
        });

        const badCodes = sent.codes.filter(code => code && (code < 200 || code >= 300));
        if (sent.fetchError || badCodes.length) {
          on = false;

          const overloaded = badCodes.some(code => code === 429 || code >= 500);
          if (overloaded) {
            PropertiesService.getUserProperties().setProperty(
              FAST_GENERATION_PAUSED_UNTIL_PROPERTY,
              String(Date.now() + FAST_GENERATION_PAUSE_MS)
            );
          }

          log_(label, 'FAST_OFF', [
            sent.fetchError ? `동시 요청 실패: ${String(sent.fetchError.message || sent.fetchError)}` : `오류 응답 ${badCodes.join(', ')}`,
            overloaded ? `요청 한도 또는 서버 과부하로 ${FAST_GENERATION_PAUSE_MS / 60000}분 동안 한 명씩 보냅니다.` : '이번 실행은 한 명씩 보냅니다.',
          ].join(' / '));
        }
      } catch (err) {
        // 미리 받기는 덤이다. 실패하면 조용히 한 명씩 보내는 경로로 넘긴다.
        on = false;
        if (!isAiTimeBudgetError_(err)) {
          log_(label, 'FAST_OFF', `미리 받기 실패, 한 명씩 보냅니다: ${String(err.message || err)}`);
        }
      }
    },

    finish() {
      if (AI_PREFETCHED_COUNT_ > 0) {
        log_(label, 'FAST', `미리 받았지만 이번 실행에서 쓰지 못한 답 ${AI_PREFETCHED_COUNT_}개`);
      }
      clearPrefetchedAi_();
    },
  };
}

// responseSchema를 주지 않으면 GPT·Gemini는 생기부 초안 형식(evidence_summary, record_draft, caution)으로 답하게 한다.
// Claude는 형식을 강제하지 않고 프롬프트의 JSON 예시를 따른다.
function callAi_(provider, model, systemGuide, userPrompt, reasoning, responseSchema) {
  provider = normalizeAiProvider_(provider);
  reasoning = normalizeReasoning_(reasoning);

  const prefetched = takePrefetchedAi_(provider, model, systemGuide, userPrompt, reasoning, responseSchema);
  if (prefetched) return prefetched;

  if (provider === 'gpt') {
    return callOpenAi_(model, systemGuide, userPrompt, reasoning, responseSchema);
  }

  if (provider === 'gemini') {
    return callGemini_(model, systemGuide, userPrompt, reasoning, responseSchema);
  }

  if (provider === 'claude') {
    return callClaude_(model, systemGuide, userPrompt, reasoning);
  }

  throw new Error(`지원하지 않는 AI_PROVIDER입니다. claude, gpt, gemini 중 하나로 설정해주세요: ${provider}`);
}

function normalizeAiProvider_(provider) {
  const value = String(provider || 'claude').trim().toLowerCase();

  if (value === 'openai') return 'gpt';
  if (value === 'anthropic') return 'claude';
  if (value === 'google') return 'gemini';

  return value;
}

function normalizeReasoning_(reasoning) {
  const value = String(reasoning || 'none').trim().toLowerCase();

  if (value === 'low' || value === 'medium') return value;

  return 'none';
}

function getClaudeMaxTokens_(reasoning) {
  reasoning = normalizeReasoning_(reasoning);

  // 한국어 500자 초안과 근거 요약이 들어갈 만큼 넉넉하게 둔다. 실제로 쓴 토큰만 과금된다.
  if (reasoning === 'medium') return 8000;
  if (reasoning === 'low') return 4000;

  return 2000;
}

function applyClaudeReasoning_(payload, model, reasoning) {
  reasoning = normalizeReasoning_(reasoning);

  if (reasoning === 'none') {
    // Haiku 5.5와 Sonnet 5.5는 thinking을 안 보내면 스스로 생각을 켠다. 그래서 '없음'일 때 명시적으로 끈다.
    // Sonnet 5.5는 disabled를 400으로 거절하고 between_tools만 받는다.
    const value = String(model || '').toLowerCase();
    if (/claude-sonnet-5-5/.test(value)) payload.thinking = { type: 'between_tools' };
    else if (/claude-haiku-5-5/.test(value)) payload.thinking = { type: 'disabled' };
    return;
  }

  if (usesLegacyClaudeThinking_(model)) {
    payload.thinking = {
      type: 'enabled',
      budget_tokens: reasoning === 'medium' ? 2048 : 1024,
    };
    return;
  }

  payload.thinking = {
    type: 'adaptive',
  };
  payload.output_config = {
    effort: reasoning,
  };
}

// Claude 3·4.0·4.1·4.5 모델만 budget_tokens 방식을 쓴다. 그 뒤 모델(4.6, 5.x 등)은 adaptive만 받는다.
function usesLegacyClaudeThinking_(model) {
  const value = String(model || '').toLowerCase();
  return /claude-3/.test(value) ||
    /claude-(opus|sonnet|haiku)-4-[015](\b|$)/.test(value) ||
    /claude-(opus|sonnet|haiku)-4-\d{8}/.test(value);
}

function extractClaudeOutputText_(json) {
  const content = json && json.content ? json.content : [];

  return content
    .filter(block => block && block.text && (!block.type || block.type === 'text'))
    .map(block => block.text)
    .join('\n');
}

function getGeminiThinkingLevel_(reasoning, model) {
  reasoning = normalizeReasoning_(reasoning);

  if (reasoning === 'medium') return 'medium';
  if (reasoning === 'low') return 'low';

  return isGeminiMinimalUnsupported_(model) ? 'low' : 'minimal';
}

// minimal을 400으로 거절한 모델을 기억해 둔다. 알려진 모델은 처음부터 low로 보내 헛 호출을 줄인다.
const GEMINI_MINIMAL_UNSUPPORTED_MODELS = ['gemini-3.8-flash', 'gemini-3.7-flash'];

function isGeminiMinimalUnsupported_(model) {
  const value = String(model || '').trim().toLowerCase();
  if (GEMINI_MINIMAL_UNSUPPORTED_MODELS.includes(value)) return true;

  try {
    return CacheService.getUserCache().get(`GEMINI_NO_MINIMAL_${value}`) === '1';
  } catch (err) {
    return false;
  }
}

function rememberGeminiMinimalUnsupported_(model) {
  try {
    CacheService.getUserCache().put(`GEMINI_NO_MINIMAL_${String(model || '').trim().toLowerCase()}`, '1', 21600);
  } catch (err) {
    // 기억하지 못해도 다음 호출에서 다시 low로 재요청하므로 괜찮다.
  }
}


const RECORD_DRAFT_RESPONSE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    evidence_summary: { type: 'string' },
    record_draft: { type: 'string' },
    caution: { type: 'string' },
  },
  required: ['evidence_summary', 'record_draft', 'caution'],
};

// Gemini 형식에는 additionalProperties를 넣지 않는다(기존 초안 형식도 넣지 않고 써 왔다).
function stripAdditionalProperties_(schema) {
  if (Array.isArray(schema)) return schema.map(stripAdditionalProperties_);
  if (!schema || typeof schema !== 'object') return schema;

  const result = {};
  Object.keys(schema).forEach(key => {
    if (key !== 'additionalProperties') result[key] = stripAdditionalProperties_(schema[key]);
  });
  return result;
}

function buildOpenAiPayload_(model, systemGuide, userPrompt, reasoning, responseSchema) {
  reasoning = normalizeReasoning_(reasoning);
  const baseMaxOutputTokens = getOpenAiMaxOutputTokens_(reasoning);

  return {
    model: model || 'gpt-6-luna',
    // 응답을 OpenAI 서버에 대화 기록으로 남기지 않는다(기본값은 30일 보관).
    store: false,
    reasoning: {
      effort: reasoning,
    },
    instructions: systemGuide,
    input: [
      {
        role: 'user',
        content: [
          {
            type: 'input_text',
            text: userPrompt,
          },
        ],
      },
    ],
    max_output_tokens: baseMaxOutputTokens,
    text: {
      format: {
        type: 'json_schema',
        name: 'student_record_draft',
        strict: true,
        schema: responseSchema || RECORD_DRAFT_RESPONSE_SCHEMA,
      },
    },
  };
}

function callOpenAi_(model, systemGuide, userPrompt, reasoning, responseSchema) {
  const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error('OpenAI API 키가 없습니다. 메뉴에서 본인 OpenAI API 키를 먼저 저장하세요.');
  }

  reasoning = normalizeReasoning_(reasoning);
  const baseMaxOutputTokens = getOpenAiMaxOutputTokens_(reasoning);
  const payload = buildOpenAiPayload_(model, systemGuide, userPrompt, reasoning, responseSchema);

  const maxAttempts = 5;
  let outputTokenRetryUsed = false;
  let lastErrorMessage = '';

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let res;

    try {
      res = fetchAiWithinBudget_(AI_API_ENDPOINTS.gpt.url, buildAiFetchOptions_('gpt', apiKey, payload));
    } catch (err) {
      // 응답 없이 끊긴 경우(시간 초과, 연결 오류)도 잠시 뒤 다시 시도한다.
      if (isAiTimeBudgetError_(err) || attempt === maxAttempts) throw err;
      lastErrorMessage = `OpenAI API 네트워크 오류: ${String(err.message || err)}`;
      log_('callOpenAi_', 'RETRY', `${lastErrorMessage}. ${attempt}/${maxAttempts}회 시도 실패`);
      sleepWithinAiBudget_(Math.min(2000 * Math.pow(2, attempt - 1), 30000));
      continue;
    }

    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);

      if (isOpenAiMaxOutputTokenIncomplete_(json) && !outputTokenRetryUsed) {
        outputTokenRetryUsed = true;
        payload.max_output_tokens = Math.min(Number(payload.max_output_tokens || baseMaxOutputTokens) * 2, 16000);
        log_(
          'callOpenAi_',
          'RETRY',
          `OpenAI 응답이 max_output_tokens로 중단되어 ${payload.max_output_tokens} 토큰으로 1회 재시도합니다. ${summarizeOpenAiUsage_(json)}`
        );
        sleepWithinAiBudget_(1200);
        continue;
      }

      if (isOpenAiMaxOutputTokenIncomplete_(json)) {
        throw new Error(`OpenAI 응답이 max_output_tokens 한도로 중단되어 output_text가 없습니다. ${summarizeOpenAiUsage_(json)}`);
      }

      const text = extractOpenAiOutputText_(json);

      return {
        text,
        usage: json.usage || {},
      };
    }

    lastErrorMessage = `OpenAI API 오류 ${code}: ${body}`;

    if ([408, 409, 429, 500, 502, 503, 504].includes(code)) {
      const waitMs = 2000 * Math.pow(2, attempt - 1) + Math.floor(Math.random() * 1000);
      sleepWithinAiBudget_(Math.min(waitMs, 30000));
      continue;
    }

    throw new Error(lastErrorMessage);
  }

  throw new Error(`OpenAI API 재시도 실패: ${lastErrorMessage}`);
}

function getOpenAiMaxOutputTokens_(reasoning) {
  reasoning = normalizeReasoning_(reasoning);

  if (reasoning === 'medium') return 6000;
  if (reasoning === 'low') return 3000;

  return 2000;
}

function isOpenAiMaxOutputTokenIncomplete_(json) {
  const reason = json && json.incomplete_details
    ? String(json.incomplete_details.reason || '').trim()
    : '';

  return json && json.status === 'incomplete' && reason === 'max_output_tokens';
}

function summarizeOpenAiUsage_(json) {
  const usage = json && json.usage ? json.usage : {};
  const outputDetails = usage.output_tokens_details || usage.outputTokensDetails || {};

  return [
    `status=${json && json.status ? json.status : ''}`,
    `output_tokens=${formatUsageNumber_(usage.output_tokens, usage.outputTokens)}`,
    `reasoning_tokens=${formatUsageNumber_(outputDetails.reasoning_tokens, outputDetails.reasoningTokens)}`,
  ].filter(part => !part.endsWith('=')).join(', ');
}

function formatUsageNumber_(snakeValue, camelValue) {
  if (snakeValue !== undefined && snakeValue !== null) return snakeValue;
  if (camelValue !== undefined && camelValue !== null) return camelValue;

  return '';
}

function buildGeminiPayload_(model, systemGuide, userPrompt, reasoning, responseSchema) {
  reasoning = normalizeReasoning_(reasoning);

  return {
    model: model || 'gemini-3.8-flash',
    // 요청을 Gemini 서버에 대화 기록으로 남기지 않는다(기본값은 보관).
    store: false,
    system_instruction: systemGuide,
    input: userPrompt,
    generation_config: {
      temperature: 0.2,
      // 생각에 쓴 토큰도 이 한도에 들어가므로 추론 수준에 맞춰 늘린다.
      max_output_tokens: getClaudeMaxTokens_(reasoning),
      thinking_level: getGeminiThinkingLevel_(reasoning, model || 'gemini-3.8-flash'),
    },
    response_format: {
      type: 'text',
      mime_type: 'application/json',
      schema: stripAdditionalProperties_(responseSchema || RECORD_DRAFT_RESPONSE_SCHEMA),
    },
  };
}

function callGemini_(model, systemGuide, userPrompt, reasoning, responseSchema) {
  const apiKey = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('Gemini API 키가 없습니다. 메뉴에서 본인 Gemini API 키를 먼저 저장하세요.');
  }

  reasoning = normalizeReasoning_(reasoning);
  const payload = buildGeminiPayload_(model, systemGuide, userPrompt, reasoning, responseSchema);

  const maxAttempts = 5;
  let outputTokenRetryUsed = false;
  let lastErrorMessage = '';

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let res;

    try {
      res = fetchAiWithinBudget_(AI_API_ENDPOINTS.gemini.url, buildAiFetchOptions_('gemini', apiKey, payload));
    } catch (err) {
      // 응답 없이 끊긴 경우(시간 초과, 연결 오류)도 잠시 뒤 다시 시도한다.
      if (isAiTimeBudgetError_(err) || attempt === maxAttempts) throw err;
      lastErrorMessage = `Gemini API 네트워크 오류: ${String(err.message || err)}`;
      log_('callGemini_', 'RETRY', `${lastErrorMessage}. ${attempt}/${maxAttempts}회 시도 실패`);
      sleepWithinAiBudget_(Math.min(2000 * Math.pow(2, attempt - 1), 30000));
      continue;
    }

    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);

      if (isGeminiOutputTruncated_(json) && !outputTokenRetryUsed) {
        outputTokenRetryUsed = true;
        payload.generation_config.max_output_tokens = Math.min(Number(payload.generation_config.max_output_tokens) * 2, 16000);
        log_('callGemini_', 'RETRY', `Gemini 응답이 출력 한도로 잘려 ${payload.generation_config.max_output_tokens} 토큰으로 1회 재시도합니다.`);
        sleepWithinAiBudget_(1200);
        continue;
      }

      if (isGeminiOutputTruncated_(json)) {
        throw new Error(`Gemini 응답이 출력 한도(max_output_tokens ${payload.generation_config.max_output_tokens})에서 잘렸습니다. 다시 만들어 보세요.`);
      }

      const text = extractGeminiOutputText_(json);

      return {
        text,
        usage: json.usage || json.usage_metadata || json.usageMetadata || {},
      };
    }

    lastErrorMessage = `Gemini API 오류 ${code}: ${body}`;

    // gemini-3.8-flash·3.7-flash처럼 minimal을 받지 않는 모델은 가장 낮은 low로 다시 보내고, 다음부터는 처음부터 low로 보낸다.
    if (code === 400 && payload.generation_config.thinking_level === 'minimal' && /THINKING_LEVEL_MINIMAL/.test(body)) {
      payload.generation_config.thinking_level = 'low';
      rememberGeminiMinimalUnsupported_(payload.model);
      log_('callGemini_', 'RETRY', `${payload.model}은 thinking_level minimal을 지원하지 않아 low로 다시 요청합니다.`);
      continue;
    }

    if ([408, 409, 429, 500, 502, 503, 504].includes(code)) {
      const waitMs = 2000 * Math.pow(2, attempt - 1) + Math.floor(Math.random() * 1000);
      sleepWithinAiBudget_(Math.min(waitMs, 30000));
      continue;
    }

    throw new Error(lastErrorMessage);
  }

  throw new Error(`Gemini API 재시도 실패: ${lastErrorMessage}`);
}

// interactions 응답의 status, 또는 어느 단계/후보든 finish reason이 MAX_TOKENS면 잘린 답이다.
function isGeminiOutputTruncated_(json) {
  if (!json) return false;
  if (String(json.status || '').toLowerCase() === 'incomplete') return true;

  return /"finish_?reason"\s*:\s*"max_tokens"/i.test(JSON.stringify(json));
}

function extractOpenAiOutputText_(json) {
  if (json.output_text) {
    return json.output_text;
  }

  if (json.output && Array.isArray(json.output)) {
    for (const item of json.output) {
      if (item.content && Array.isArray(item.content)) {
        for (const content of item.content) {
          if (content.type === 'output_text' && content.text) {
            return content.text;
          }
        }
      }
    }
  }

  throw new Error('OpenAI 응답에서 output_text를 찾을 수 없습니다.');
}

function extractGeminiOutputText_(json) {
  if (json.output_text) {
    return json.output_text;
  }

  if (json.outputText) {
    return json.outputText;
  }

  if (json.steps && Array.isArray(json.steps)) {
    const parts = [];

    json.steps.forEach(step => {
      const content = step.content || step.output || [];
      const blocks = Array.isArray(content) ? content : [content];

      blocks.forEach(block => {
        if (block && block.type === 'text' && block.text) {
          parts.push(block.text);
        }
      });
    });

    if (parts.length > 0) {
      return parts.join('');
    }
  }

  if (json.candidates && Array.isArray(json.candidates)) {
    const parts = [];

    json.candidates.forEach(candidate => {
      const candidateParts = candidate.content && Array.isArray(candidate.content.parts)
        ? candidate.content.parts
        : [];

      candidateParts.forEach(part => {
        if (part.text) parts.push(part.text);
      });
    });

    if (parts.length > 0) {
      return parts.join('');
    }
  }

  throw new Error('Gemini 응답에서 output_text를 찾을 수 없습니다.');
}

function isRetryableClaudeError_(code) {
  return [429, 500, 502, 503, 504, 529].includes(code);
}

function getClaudeRetryDelayMs_(res, attempt) {
  const headers = res.getAllHeaders ? res.getAllHeaders() : {};
  const retryAfter =
    headers['Retry-After'] ||
    headers['retry-after'] ||
    headers['Retry-after'];

  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (!isNaN(seconds) && seconds > 0) {
      return Math.min(seconds * 1000, 60000);
    }
  }

  const baseMs = 2000;
  const exponentialMs = baseMs * Math.pow(2, attempt - 1);
  const jitterMs = Math.floor(Math.random() * 1000);

  return Math.min(exponentialMs + jitterMs, 30000);
}
function parseClaudeJson_(text) {
  try {
    return JSON.parse(text);
  } catch (err) {
    const match = text.match(/\{[\s\S]*\}/);
    if (!match) {
      // JSON을 쓰다 끊긴 답이면 조각이 초안 본문으로 들어가지 않게 오류로 넘긴다.
      if (/^\s*(```(json)?\s*)?\{/i.test(text)) {
        throw new Error('AI 응답 JSON이 중간에 끊겨 읽을 수 없습니다. 다시 만들어 보세요.');
      }
      return {
        evidence_summary: '',
        record_draft: text.trim(),
        caution: 'JSON 파싱 실패. 원문을 직접 검토하세요.',
      };
    }
    return JSON.parse(match[0]);
  }
}

function getAllClassroomPages_(path, arrayKey, params) {
  const all = [];
  let pageToken = '';

  do {
    const merged = Object.assign({}, params || {});
    if (pageToken) merged.pageToken = pageToken;

    const json = classroomGet_(path, merged);
    const items = json[arrayKey] || [];
    all.push(...items);
    pageToken = json.nextPageToken || '';
  } while (pageToken);

  return all;
}

function classroomGet_(path, params) {
  const query = Object.keys(params || {})
    .filter(k => params[k] !== undefined && params[k] !== null && params[k] !== '')
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');

  const url = `https://classroom.googleapis.com/v1/${path}${query ? '?' + query : ''}`;

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Classroom API 오류 ${code}: ${body}`);
  }

  return body ? JSON.parse(body) : {};
}



function lookupStudent_(email, name, classroomUserId, studentLookupIndex) {
  const index = studentLookupIndex || buildStudentLookupIndex_();
  const lookupUserId = normalizeStudentLookupText_(classroomUserId);
  const lookupEmail = normalizeStudentLookupEmail_(email);
  const lookupName = normalizeStudentLookupText_(name);

  if (lookupUserId && index.byClassroomUserId[lookupUserId]) {
    return index.byClassroomUserId[lookupUserId];
  }

  if (lookupEmail && index.byEmail[lookupEmail]) {
    return index.byEmail[lookupEmail];
  }

  if (lookupName) {
    const nameMatches = index.byName[lookupName] || [];

    if (nameMatches.length === 1) {
      return nameMatches[0];
    }

    if (nameMatches.length > 1) {
      logAmbiguousStudentName_(index, lookupName, nameMatches.length);
    }
  }

  return {
    studentNo: '',
    name: lookupName,
    email: String(email || '').trim(),
  };
}

function buildStudentLookupIndex_() {
  const sh = getSheet_(SHEET_NAMES.students);
  const values = sh.getDataRange().getValues();
  return buildStudentLookupIndexFromValues_(values);
}

function buildStudentLookupIndexFromValues_(values) {
  const index = {
    byClassroomUserId: Object.create(null),
    byEmail: Object.create(null),
    byName: Object.create(null),
    ambiguousNameLogged: Object.create(null),
  };

  if (!values || values.length < 2) return index;

  const h = headerIndex_(values[0] || []);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const studentNo = studentLookupRowValue_(row, h.studentNo);
    const studentName = normalizeStudentLookupText_(studentLookupRowValue_(row, h.name));
    const studentEmail = String(studentLookupRowValue_(row, h.email) || '').trim();
    const classroomUserId = normalizeStudentLookupText_(studentLookupRowValue_(row, h.classroomUserId));
    const emailKey = normalizeStudentLookupEmail_(studentEmail);

    if (!studentNo && !studentName && !studentEmail && !classroomUserId) continue;

    const student = {
      studentNo: studentNo || '',
      name: studentName,
      email: studentEmail,
    };

    if (classroomUserId && !index.byClassroomUserId[classroomUserId]) {
      index.byClassroomUserId[classroomUserId] = student;
    }

    if (emailKey && !index.byEmail[emailKey]) {
      index.byEmail[emailKey] = student;
    }

    if (studentName) {
      if (!index.byName[studentName]) index.byName[studentName] = [];
      index.byName[studentName].push(student);
    }
  }

  return index;
}

function studentLookupRowValue_(row, index) {
  if (index === undefined || index === null) return '';
  return row[index];
}

function normalizeStudentLookupText_(value) {
  return String(value || '').trim();
}

function normalizeStudentLookupEmail_(value) {
  return normalizeStudentLookupText_(value).toLowerCase();
}

function logAmbiguousStudentName_(index, name, count) {
  if (index.ambiguousNameLogged && index.ambiguousNameLogged[name]) return;

  log_(
    'lookupStudent_',
    'AMBIGUOUS_NAME',
    `동명이인 자동 연결 안 함: ${name} (학생명단 ${count}명)`
  );

  if (index.ambiguousNameLogged) index.ambiguousNameLogged[name] = true;
}

function getConfigMap_() {
  const sh = getSheet_(SHEET_NAMES.config);
  const values = sh.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const value = String(values[i][1] || '');
    if (key) map[key] = value;
  }

  return map;
}

const CONFIG_VALIDATION_OPTIONS = {
  AI_PROVIDER: ['claude', 'gpt', 'gemini'],
  CLAUDE_MODEL: ['claude-haiku-5-5', 'claude-sonnet-5-5'],
  GPT_MODEL: ['gpt-6-luna', 'gpt-6.1-terra'],
  GEMINI_MODEL: ['gemini-3.8-flash', 'gemini-3.7-flash', 'gemini-3.6-flash', 'gemini-3.5-flash'],
  REASONING: ['none', 'low', 'medium'],
  FAST_GENERATION: ['OFF', 'ON'],
};

function syncConfigModelsOnOpen_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  const config = ss.getSheetByName(SHEET_NAMES.config);
  if (!config || config.getLastRow() < 2) return;

  migrateLegacyConfigKeys_(config);
  applyConfigValidations_(config);
  syncConfigDefaults_(config, [
    ['AI_PROVIDER', 'claude', 'claude 또는 gpt 또는 gemini'],
    ['CLAUDE_MODEL', 'claude-haiku-5-5', 'claude-haiku-5-5 / claude-sonnet-5-5\nclaude-haiku-5-5 - 기본값. 값이 싸고 대답이 빠름. 생기부 문장 쓰기에 충분함\nclaude-sonnet-5-5 - 상대적으로 비쌈. 조금 느리지만 더 똑똑한 편'],
    ['GPT_MODEL', 'gpt-6-luna', 'gpt-6-luna / gpt-6.1-terra\ngpt-6.1-terra / 상위 고성능 모델\ngpt-6-luna / 무난하게 쓸 수 있는 기본 모델'],
    ['GEMINI_MODEL', 'gemini-3.8-flash', 'gemini-3.8-flash / gemini-3.7-flash / gemini-3.6-flash / gemini-3.5-flash'],
    ['REASONING', 'none', '추론 수준입니다\nlow, medium을 사용하면 글을 조금 더 생각하고 써줍니다\n다만, none이 아니라 low를 사용하면 비용을 대략 2배정도 사용하는 것 같습니다'],
  ]);
}

function syncConfigDefaults_(config, defaults) {
  const lastRow = config.getLastRow();
  if (lastRow < 2) return;

  const defaultMap = {};
  defaults.forEach(row => {
    defaultMap[row[0]] = row;
  });

  const values = config.getRange(2, 1, lastRow - 1, 3).getValues();

  values.forEach((row, index) => {
    const sheetRow = index + 2;
    const key = String(row[0] || '').trim();
    if (!key || !defaultMap[key]) return;

    const description = defaultMap[key][2];
    if (String(row[2] || '') !== description) {
      config.getRange(sheetRow, 3).setValue(description);
    }

    if (key === 'AI_PROVIDER') {
      const provider = String(row[1] || '').trim().toLowerCase();

      if (provider === 'openai') {
        config.getRange(sheetRow, 2).setValue('gpt');
      } else if (provider === 'anthropic') {
        config.getRange(sheetRow, 2).setValue('claude');
      } else if (provider === 'google') {
        config.getRange(sheetRow, 2).setValue('gemini');
      }
    } else if (key === 'CLAUDE_MODEL') {
      const model = String(row[1] || '').trim();

      if (!CONFIG_VALIDATION_OPTIONS.CLAUDE_MODEL.includes(model)) {
        const nextModel = model.indexOf('sonnet') !== -1 ? 'claude-sonnet-5-5' : defaultMap[key][1];
        config.getRange(sheetRow, 2).setValue(nextModel);
      }
    } else if (key === 'GPT_MODEL') {
      const model = String(row[1] || '').trim();

      if (!CONFIG_VALIDATION_OPTIONS.GPT_MODEL.includes(model)) {
        config.getRange(sheetRow, 2).setValue(defaultMap[key][1]);
      }
    } else if (key === 'GEMINI_MODEL') {
      const model = String(row[1] || '').trim();

      if (!CONFIG_VALIDATION_OPTIONS.GEMINI_MODEL.includes(model)) {
        config.getRange(sheetRow, 2).setValue(defaultMap[key][1]);
      }
    } else if (key === 'REASONING') {
      const reasoning = normalizeReasoning_(row[1]);

      if (reasoning === 'none' && String(row[1] || '').trim().toLowerCase() !== 'none') {
        config.getRange(sheetRow, 2).setValue(defaultMap[key][1]);
      }
    }
  });
}

function migrateLegacyConfigKeys_(config) {
  const lastRow = config.getLastRow();
  if (lastRow < 2) return;

  const values = config.getRange(2, 1, lastRow - 1, 3).getValues();
  let gptModelRow = 0;
  const openAiModelRows = [];
  let firstOpenAiModelValue = '';
  let reasoningRow = 0;
  const legacyReasoningRows = [];
  let firstLegacyReasoningValue = '';
  let behaviorMinCharsRow = 0;
  let behaviorMaxCharsRow = 0;
  const legacyBehaviorMinRows = [];
  const legacyBehaviorMaxRows = [];
  let firstLegacyBehaviorMinValue = '';
  let firstLegacyBehaviorMaxValue = '';

  values.forEach((row, index) => {
    const sheetRow = index + 2;
    const key = String(row[0] || '').trim();
    const normalizedKey = key.toUpperCase();

    if (key === 'GPT_MODEL') {
      gptModelRow = sheetRow;
    } else if (key === 'OPENAI_MODEL') {
      openAiModelRows.push(sheetRow);
      if (!firstOpenAiModelValue) {
        firstOpenAiModelValue = row[1];
      }
    } else if (key === 'REASONING') {
      reasoningRow = sheetRow;
    } else if (normalizedKey === 'REASONING') {
      legacyReasoningRows.push(sheetRow);
      if (!firstLegacyReasoningValue) {
        firstLegacyReasoningValue = row[1];
      }
    } else if (key === 'BEHAVIOR_RECORD_MIN_CHARS') {
      behaviorMinCharsRow = sheetRow;
    } else if (key === 'BEHAVIOR_RECORD_MAX_CHARS') {
      behaviorMaxCharsRow = sheetRow;
    } else if (key === 'BEHAVIOR_RECORD_BYTE_MIN') {
      legacyBehaviorMinRows.push(sheetRow);
      if (!firstLegacyBehaviorMinValue) {
        firstLegacyBehaviorMinValue = row[1];
      }
    } else if (key === 'BEHAVIOR_RECORD_BYTE_MAX') {
      legacyBehaviorMaxRows.push(sheetRow);
      if (!firstLegacyBehaviorMaxValue) {
        firstLegacyBehaviorMaxValue = row[1];
      }
    }
  });

  if (openAiModelRows.length > 0 && !gptModelRow) {
    const migratedRow = openAiModelRows.shift();
    config.getRange(migratedRow, 1, 1, 3).setValues([[
      'GPT_MODEL',
      firstOpenAiModelValue || 'gpt-6-luna',
      'gpt-6-luna / gpt-6.1-terra\ngpt-6.1-terra / 상위 고성능 모델\ngpt-6-luna / 무난하게 쓸 수 있는 기본 모델',
    ]]);
  }

  if (legacyReasoningRows.length > 0) {
    if (!reasoningRow) {
      const migratedRow = legacyReasoningRows.shift();
      config.getRange(migratedRow, 1, 1, 3).setValues([[
        'REASONING',
        normalizeReasoning_(firstLegacyReasoningValue),
        '추론 수준입니다\nlow, medium을 사용하면 글을 조금 더 생각하고 써줍니다\n다만, none이 아니라 low를 사용하면 비용을 대략 2배정도 사용하는 것 같습니다',
      ]]);
      reasoningRow = migratedRow;
    } else {
      const currentValue = String(config.getRange(reasoningRow, 2).getValue() || '').trim();

      if (!currentValue) {
        config.getRange(reasoningRow, 2).setValue(normalizeReasoning_(firstLegacyReasoningValue));
      }
    }
  }

  if (legacyBehaviorMinRows.length > 0) {
    if (!behaviorMinCharsRow) {
      const migratedRow = legacyBehaviorMinRows.shift();
      config.getRange(migratedRow, 1, 1, 3).setValues([[
        'BEHAVIOR_RECORD_MIN_CHARS',
        firstLegacyBehaviorMinValue || '',
        BEHAVIOR_RECORD_MIN_CHARS_DESCRIPTION,
      ]]);
      behaviorMinCharsRow = migratedRow;
    } else {
      const currentValue = String(config.getRange(behaviorMinCharsRow, 2).getValue() || '').trim();

      if (!currentValue && firstLegacyBehaviorMinValue) {
        config.getRange(behaviorMinCharsRow, 2).setValue(firstLegacyBehaviorMinValue);
      }
    }
  }

  if (legacyBehaviorMaxRows.length > 0) {
    if (!behaviorMaxCharsRow) {
      const migratedRow = legacyBehaviorMaxRows.shift();
      config.getRange(migratedRow, 1, 1, 3).setValues([[
        'BEHAVIOR_RECORD_MAX_CHARS',
        firstLegacyBehaviorMaxValue || '',
        BEHAVIOR_RECORD_MAX_CHARS_DESCRIPTION,
      ]]);
      behaviorMaxCharsRow = migratedRow;
    } else {
      const currentValue = String(config.getRange(behaviorMaxCharsRow, 2).getValue() || '').trim();

      if (!currentValue && firstLegacyBehaviorMaxValue) {
        config.getRange(behaviorMaxCharsRow, 2).setValue(firstLegacyBehaviorMaxValue);
      }
    }
  }

  openAiModelRows
    .concat(legacyReasoningRows)
    .concat(legacyBehaviorMinRows)
    .concat(legacyBehaviorMaxRows)
    .sort((a, b) => b - a)
    .forEach(row => config.deleteRow(row));
}

function applyConfigValidations_(config) {
  const lastRow = config.getLastRow();
  if (lastRow < 2) return;

  const values = config.getRange(2, 1, lastRow - 1, 1).getValues();
  const validationOptions = CONFIG_VALIDATION_OPTIONS;

  values.forEach((row, index) => {
    const key = String(row[0] || '').trim();
    const options = validationOptions[key];
    if (!options) return;

    const range = config.getRange(index + 2, 2);

    // Replacing an equivalent dropdown through Apps Script drops the Sheets chip display style.
    if (hasSameValueInListValidation_(range.getDataValidation(), options)) {
      return;
    }

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(false)
      .build();

    range.setDataValidation(rule);
  });
}

function hasSameValueInListValidation_(rule, options) {
  if (!rule) return false;
  if (rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return false;

  const values = rule.getCriteriaValues();
  const currentOptions = Array.isArray(values[0])
    ? values[0].map(value => String(value || '').trim())
    : [];

  if (currentOptions.length !== options.length) return false;

  return currentOptions.every((value, index) => value === String(options[index] || '').trim());
}

function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error(`${name} 시트가 없습니다. 먼저 시트 초기 설정을 실행하세요.`);
  return sh;
}

function clearData_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}

function ensureSheetHeaders_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  const headers = HEADERS[sheetName];
  if (!headers || headers.length === 0) return sh;

  const currentWidth = Math.max(sh.getLastColumn(), 1);
  const currentHeaders = sh.getRange(1, 1, 1, currentWidth).getValues()[0]
    .map(value => String(value || '').trim());

  headers.forEach((header, index) => {
    const targetCol = index + 1;

    if (currentHeaders[index] === header) return;
    if (currentHeaders.indexOf(header) !== -1) return;

    if (targetCol <= sh.getMaxColumns()) {
      sh.insertColumnBefore(targetCol);
    } else {
      sh.insertColumnsAfter(sh.getMaxColumns(), 1);
    }

    currentHeaders.splice(index, 0, header);
  });

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function deleteActiveSheetStudentData() {
  const ui = getUi_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sh) return;

  const sheetName = sh.getName();

  if (sheetName === MANUAL_SHEET_NAME) {
    clearManualSheetData_(sh, ui);
    return;
  }

  const startRowsBySheetName = {
    [SHEET_NAMES.assignments]: 2,
    [SHEET_NAMES.submissions]: 2,
    [SHEET_NAMES.records]: 2,
    [SHEET_NAMES.studentFinalRecords]: 2,
    [SHEET_NAMES.commonPhrases]: 3,
    [SHEET_NAMES.behaviorRecommendations]: 2,
  };
  const startRow = startRowsBySheetName[sheetName];

  if (!startRow) {
    ui.alert(
      '활성화된 시트 학생 데이터 삭제',
      `${sheetName} 시트는 이 기능으로 삭제할 수 없습니다.\n과제목록, 제출물, 생기부초안, 학생별생기부, 공통문구생성, 수동추가, 행발문구추천 시트에서만 실행할 수 있습니다.`,
      ui.ButtonSet.OK
    );
    log_('deleteActiveSheetStudentData', 'BLOCKED', `${sheetName}: 삭제 허용 시트 아님`);
    return;
  }

  const lastRow = sh.getLastRow();
  const rowsToDelete = lastRow - startRow + 1;

  if (rowsToDelete < 1) {
    ui.alert(`${sheetName} 시트에 삭제할 데이터 행이 없습니다.`);
    return;
  }

  const confirm = ui.alert(
    '활성화된 시트 학생 데이터 삭제',
    `${sheetName} 시트의 ${startRow}행부터 ${lastRow}행까지 ${rowsToDelete}개 행을 삭제합니다.\n계속할까요?`,
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) return;

  if (sheetName === SHEET_NAMES.behaviorRecommendations) {
    deleteAutoBehaviorRecommendationTriggers_();
    PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');
  }

  sh.deleteRows(startRow, rowsToDelete);

  if (sh.getMaxRows() < startRow) {
    sh.insertRowsAfter(sh.getMaxRows(), startRow - sh.getMaxRows());
  }

  refreshDashboard();
  log_('deleteActiveSheetStudentData', 'OK', `${sheetName}: ${startRow}-${lastRow}행 삭제`);
  ui.alert(`${sheetName} 시트의 데이터 행 ${rowsToDelete}개를 삭제했습니다.`);
}

function clearManualSheetData_(sh, ui) {
  const confirm = ui.alert(
    '수동추가 시트 데이터 삭제',
    `수동추가 시트의 모든 내용, 메모, 서식, 데이터 검증을 삭제하고 빈 ${MANUAL_RESET_ROWS}행을 남깁니다.\n계속할까요?`,
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) return;

  deleteAutoManualRecordTriggers_();
  PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');

  try {
    const filter = sh.getFilter();
    if (filter) filter.remove();
  } catch (err) {
    // 필터 제거 실패는 삭제 자체를 막지 않는다.
  }

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearNote();
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearNote();
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);

  const maxRows = sh.getMaxRows();

  if (maxRows < MANUAL_RESET_ROWS) {
    sh.insertRowsAfter(maxRows, MANUAL_RESET_ROWS - maxRows);
  } else if (maxRows > MANUAL_RESET_ROWS) {
    sh.deleteRows(MANUAL_RESET_ROWS + 1, maxRows - MANUAL_RESET_ROWS);
  }

  log_('deleteActiveSheetStudentData', 'OK', `${MANUAL_SHEET_NAME}: 전체 데이터 삭제, ${MANUAL_RESET_ROWS}행 유지`);
  ui.alert(`수동추가 시트의 데이터를 삭제하고 빈 ${MANUAL_RESET_ROWS}행을 남겼습니다.`);
}

function upsertByKey_(sheetName, keyHeader, rows) {
  if (!rows || rows.length === 0) {
    return {
      appended: 0,
      updated: 0,
      targetRows: [],
    };
  }

  const sh = getSheet_(sheetName);
  const lastRow = Math.max(sh.getLastRow(), 1);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const sheetValues = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const actualHeaders = sheetValues[0] || [];
  const keyCol = actualHeaders.indexOf(keyHeader) + 1;

  if (keyCol < 1) {
    throw new Error(`${sheetName} 시트에서 ${keyHeader} 헤더를 찾을 수 없습니다.`);
  }

  const existingValues = sheetValues.slice(1);
  const existing = Object.create(null);
  let lastDataRow = 1;

  existingValues.forEach((row, i) => {
    const key = String(row[keyCol - 1] || '').trim();
    const rowNumber = i + 2;

    if (key) {
      existing[key] = {
        rowNumber,
        values: row,
      };
      lastDataRow = rowNumber;
    }
  });

  const maxIncomingWidth = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const writeWidth = Math.max(maxIncomingWidth, 1);

  if (writeWidth > sh.getMaxColumns()) {
    sh.insertColumnsAfter(sh.getMaxColumns(), writeWidth - sh.getMaxColumns());
  }

  let appended = 0;
  let updated = 0;
  const targetRows = [];
  const writesByRow = Object.create(null);

  let nextAppendRow = Math.max(lastDataRow + 1, 2);

  rows.forEach(row => {
    let outputRow = row.slice();
    const key = String(outputRow[keyCol - 1] || '').trim();

    if (!key) {
      throw new Error(`${sheetName} 저장 실패: ${keyHeader} 값이 비어 있습니다.`);
    }

    if (existing[key]) {
      const targetRow = existing[key].rowNumber;
      outputRow = preserveSubmissionEditableValues_(sh, actualHeaders, targetRow, outputRow, existing[key].values);
      updated++;
      targetRows.push(targetRow);
      existing[key].values = outputRow.slice();
      writesByRow[targetRow] = {
        rowNumber: targetRow,
        values: outputRow,
      };
    } else {
      const targetRow = nextAppendRow;
      appended++;
      targetRows.push(targetRow);
      existing[key] = {
        rowNumber: targetRow,
        values: outputRow.slice(),
      };
      writesByRow[targetRow] = {
        rowNumber: targetRow,
        values: outputRow,
      };
      nextAppendRow++;
    }
  });

  const writeItems = Object.keys(writesByRow).map(rowNumber => writesByRow[rowNumber]);
  const maxTargetRow = writeItems.reduce((max, item) => Math.max(max, item.rowNumber), 0);
  ensureSheetRows_(sh, maxTargetRow);
  writeRowsInContiguousBlocks_(sh, writeItems);

  log_(
    'upsertByKey_',
    'OK',
    `${sheetName}: 추가 ${appended}개, 업데이트 ${updated}개, 행 ${targetRows.join(', ')}`
  );

  return {
    appended,
    updated,
    targetRows,
  };
}

function preserveSubmissionEditableValues_(sh, headers, targetRow, row, existingRow) {
  if (sh.getName() !== SHEET_NAMES.submissions) return row;

  ['teacherGrade'].forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col < 1 || col > row.length) return;

    const incoming = String(row[col - 1] || '').trim();
    if (incoming) return;

    const existing = existingRow ? existingRow[col - 1] : sh.getRange(targetRow, col).getValue();
    if (String(existing || '').trim()) {
      row[col - 1] = existing;
    }
  });

  return row;
}

function ensureSheetRows_(sh, requiredLastRow) {
  if (!requiredLastRow || requiredLastRow <= sh.getMaxRows()) return;
  sh.insertRowsAfter(sh.getMaxRows(), requiredLastRow - sh.getMaxRows());
}

function writeRowsInContiguousBlocks_(sh, writeItems) {
  if (!writeItems || writeItems.length === 0) return;

  const items = writeItems
    .filter(item => item && item.rowNumber && item.values)
    .sort((a, b) => a.rowNumber - b.rowNumber);

  let blockStart = null;
  let blockWidth = 0;
  let blockValues = [];

  const flush = () => {
    if (!blockValues.length) return;
    sh.getRange(blockStart, 1, blockValues.length, blockWidth).setValues(blockValues);
    blockStart = null;
    blockWidth = 0;
    blockValues = [];
  };

  items.forEach(item => {
    const values = item.values.slice();
    const width = values.length;
    const expectedRow = blockStart === null ? item.rowNumber : blockStart + blockValues.length;

    if (blockStart !== null && item.rowNumber === expectedRow && width === blockWidth) {
      blockValues.push(values);
      return;
    }

    flush();
    blockStart = item.rowNumber;
    blockWidth = width;
    blockValues = [values];
  });

  flush();
}

function getLastNonEmptyRowInColumn_(sh, col) {
  const maxRows = sh.getMaxRows();

  if (maxRows < 2) return 1;

  const values = sh.getRange(2, col, maxRows - 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim() !== '') {
      return i + 2;
    }
  }

  return 1;
}

function getLastDataRowByHeaders_(sh, headers) {
  const h = headerMap_(sh);
  const targetCols = headers
    .map(header => h[header])
    .filter(col => !!col);

  if (targetCols.length === 0) {
    return Math.max(sh.getLastRow(), 1);
  }

  return targetCols.reduce((lastRow, col) => {
    return Math.max(lastRow, getLastNonEmptyRowInColumn_(sh, col));
  }, 1);
}

function clearRowsBelowLastData_(sh, lastDataRow) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const maxRows = sh.getMaxRows();
  const startRow = Math.max(lastDataRow + 1, 2);

  if (lastCol < 1 || maxRows < startRow) return;

  const endRow = Math.min(Math.max(lastRow, startRow), maxRows);

  sh.getRange(startRow, 1, endRow - startRow + 1, lastCol)
    .clearContent()
    .clearDataValidations();
}

function resizeBlankRowsAfterLastData_(sh, lastDataRow, blankRowsToKeep) {
  const safeLastDataRow = Math.max(Number(lastDataRow || 1), 1);
  const blankRows = Math.max(Number(blankRowsToKeep || 0), 0);
  const targetRows = Math.max(safeLastDataRow + blankRows, 2);
  const maxRows = sh.getMaxRows();

  if (maxRows < targetRows) {
    sh.insertRowsAfter(maxRows, targetRows - maxRows);
  } else if (maxRows > targetRows) {
    sh.deleteRows(targetRows + 1, maxRows - targetRows);
  }

  if (blankRows > 0) {
    const lastCol = sh.getLastColumn();

    if (lastCol > 0) {
      sh.getRange(safeLastDataRow + 1, 1, blankRows, lastCol)
        .clearContent()
        .clearDataValidations();
    }
  }
}


function headerMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[h] = i + 1;
  });
  return map;
}

function headerIndex_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[h] = i;
  });
  return map;
}

function valueAt_(sh, row, col) {
  if (!col) return '';
  return sh.getRange(row, col).getValue();
}

function rowValueAt_(rowValues, col) {
  if (!col) return '';
  return rowValues[col - 1];
}

function setCellValueAndNote_(sh, row, col, value, note) {
  const range = sh.getRange(row, col);
  range.setValue(value);

  if (note !== undefined) {
    range.setNote(note || '');
  }
}

function truncate_(text, maxChars) {
  text = String(text || '');
  if (text.length <= maxChars) return text;
  return text.slice(0, maxChars) + '\n\n[이하 글자수 제한으로 생략됨]';
}

function sanitizePromptText_(text, options) {
  let sanitized = String(text || '')
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, '[이메일 제거]');
  const opts = options || {};

  sanitized = replaceLiteralInPrompt_(sanitized, opts.email, '[이메일 제거]', 3);
  sanitized = replaceLiteralInPrompt_(sanitized, opts.studentName, '[학생명 제거]', 2);
  sanitized = replaceLiteralInPrompt_(sanitized, opts.studentNo, '[학번 제거]', 3);

  return sanitized;
}

function replaceLiteralInPrompt_(text, value, replacement, minLength) {
  const literal = String(value || '').trim();
  if (literal.length < minLength) return text;

  return String(text || '').replace(new RegExp(escapeRegExp_(literal), 'g'), replacement);
}

function escapeRegExp_(text) {
  return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function isPersonalInfoHeader_(header) {
  const normalized = String(header || '')
    .toLowerCase()
    .replace(/[\s_\-.()[\]{}]/g, '');

  if (!normalized) return false;

  const likelyNonStudentName = /(과제|파일|학교|수업|활동|탐구|주제|제목|문항|단원|모둠)/.test(normalized);
  const nameLike =
    (/^(이름|성명|학생명|학생이름|name|fullname|studentname)/.test(normalized) ||
      /(학생명|학생이름|성명|studentname)/.test(normalized)) &&
    !likelyNonStudentName;
  const studentNumberLike =
    /^(학번|출석번호|학생번호|번호|studentno|studentnumber|studentid|userid|classroomuserid)/.test(normalized) ||
    /(학번|출석번호|학생번호|studentno|studentnumber|studentid|userid|classroomuserid)/.test(normalized);

  return (
    nameLike ||
    /(이메일|메일|email|e-mail|mail)/.test(normalized) ||
    studentNumberLike
  );
}

function normalizeAiText_(text) {
  const normalized = normalizeAiLineBreaks_(text)
    .replace(/[‘’‚‛′＇ꞌ]/g, "'")
    .replace(/[“”„‟″＂]/g, "'")
    .replace(/[『』「」｢｣《》〈〉]/g, "'")
    .replace(/([\uac00-\ud7a3])\s*[\u00b7]\s*([\uac00-\ud7a3])/g, '$1, $2')
    .replace(/[。．]/g, '.')
    .replace(/[ \t\f\v\u00a0]+/g, ' ')
    .trim();

  if (!normalized) return '';

  return normalized.endsWith('.') ? normalized : `${normalized}.`;
}

function normalizeAiLineBreaks_(text) {
  return String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t\f\v\u00a0]*\n+[ \t\f\v\u00a0]*/g, ' ');
}

function formatDue_(dueDate, dueTime) {
  if (!dueDate) return '';

  const y = dueDate.year;
  const m = String(dueDate.month || 1).padStart(2, '0');
  const d = String(dueDate.day || 1).padStart(2, '0');

  if (!dueTime) return `${y}-${m}-${d}`;

  const hh = String(dueTime.hours || 0).padStart(2, '0');
  const mm = String(dueTime.minutes || 0).padStart(2, '0');
  return `${y}-${m}-${d} ${hh}:${mm}`;
}

function log_(action, status, message) {
  try {
    const sh = getSheet_(SHEET_NAMES.logs);
    sh.appendRow([new Date(), action, status, message]);
  } catch (err) {
    console.log(action, status, message);
  }
}

function defaultSystemGuide_() {
  return [
    '# 역할',
    '너는 대한민국 고등학교 교사가 학교생활기록부 문장(교과 세부능력 및 특기사항, 창의적 체험활동 특기사항 등)을 작성할 때 돕는 도구이다.',
    '과목, 작성 영역, 작성 순서는 요청 맨 위의 [교사 안내] 또는 [자료 안내]를 따른다.',
    '',
    '# 근거 원칙',
    '1. 학생이 직접 작성하거나 수행한 내용에 드러난 활동, 사고 과정, 탐구 태도, 문제 해결 과정만 근거로 삼는다.',
    '2. 원문에 없는 사실, 성격, 태도, 역량을 추정해서 쓰지 않는다.',
    '3. 제출하지 않은 과제나 원문이 부족한 과제는 긍정 근거로 사용하지 않고, 근거가 부족하면 caution에 적는다.',
    '4. 점수 정보가 있으면 교사 입력 점수/수준을 우선하고, 없으면 만점 대비 점수를 참고해 표현 수준만 조절한다. 점수나 등급을 본문에 쓰지 않는다.',
    '',
    '# 문체',
    '1. 문장은 주로 ~함, ~임, ~음 같은 명사형 종결어미로 쓰되, 문장이 자연스럽게 이어지게 한다.',
    '2. 영어 표현은 지양한다. 예: AI 대신 인공지능.',
    "3. '이 학생은'이라는 표현을 쓰지 않는다.",
    '4. 학생 이름, 학번, 이메일 등 개인정보는 근거로 쓰거나 본문에 넣지 않는다.',
    '',
    '반드시 JSON 형식으로만 답한다.',
    '{"evidence_summary":"", "record_draft":"", "caution":""}',
  ].join('\n');
}

function getSystemGuide_(config) {
  return ensureJsonResponseGuide_(config.SYSTEM_GUIDE || defaultSystemGuide_());
}

function ensureJsonResponseGuide_(systemGuide) {
  const guide = String(systemGuide || '').trim();
  if (!guide) return defaultSystemGuide_();

  const requiredLines = [];

  if (!/JSON\s*형식으로만\s*답/i.test(guide)) {
    requiredLines.push('반드시 JSON 형식으로만 답한다.');
  }

  if (
    !guide.includes('"evidence_summary"') ||
    !guide.includes('"record_draft"') ||
    !guide.includes('"caution"')
  ) {
    requiredLines.push('{"evidence_summary":"", "record_draft":"", "caution":""}');
  }

  if (!/(개인정보|이메일|학번)/.test(guide)) {
    requiredLines.push('학생 이름, 학번, 이메일 등 개인정보는 근거로 사용하거나 record_draft 본문에 쓰지 않는다.');
  }

  if (!/(점수 정보|draftGrade|만점 대비|교사 입력 점수)/.test(guide)) {
    requiredLines.push('점수 정보가 있으면 교사 입력 점수/수준을 우선하고, 없으면 만점 대비 Classroom draftGrade를 참고해 수행 수준에 맞게 작성한다. 단, 점수 자체를 record_draft에 직접 쓰거나 점수만으로 원문에 없는 내용을 꾸며내지 않는다.');
  }

  return [guide]
    .concat(requiredLines)
    .filter(line => String(line || '').trim())
    .join('\n');
}

function defaultRecordDraftPrompt_() {
  return [
    '[교사 안내] (선택)',
    '- 과목/작성 영역: ',
    '- 강조할 점: ',
    '- 작성 순서/형식: ',
    '',
    '',
    '[작성 원칙]',
    '1. [교사 안내]의 과목/작성 영역에 맞는 어휘와 관점으로 쓰고, 작성 순서/형식이 있으면 그대로 따른다.',
    '2. 강조할 점은 제출물에 근거가 있을 때만 반영한다.',
    '3. 과제 문항, 교사 지시문, 예시문은 학생 답변을 이해하는 맥락으로만 쓰고, 학생이 직접 쓴 답변만 근거로 삼는다.',
  ].join('\n');
}

function defaultStudentFinalRecordPrompt_() {
  return [
    '[교사 안내] (선택)',
    '- 과목/작성 영역: ',
    '- 강조할 점: ',
    '- 작성 순서/형식: ',
    '',
    '',
    '[작성 원칙]',
    '1. [교사 안내]의 과목/작성 영역에 맞는 어휘와 관점으로 쓰고, 작성 순서/형식이 있으면 그대로 따른다.',
    '2. 강조할 점은 과제별 근거가 있을 때만 반영한다.',
    '3. 여러 과제를 단순 나열하지 말고, 학생의 수행 특성이 자연스럽게 드러나도록 종합한다.',
  ].join('\n');
}

function defaultManualRecordPrompt_() {
  return [
    '[자료 안내] (선택)',
    '- 자료 종류: ',
    '- 작성할 영역: ',
    '- 작성 순서/형식: ',
    '',
    '',
    '[작성 원칙]',
    '1. [수동추가 데이터]에서 학생이 직접 한 활동, 결과물, 생각, 탐구 과정만 근거로 쓴다. 데이터에 없는 사실, 성격, 역량을 지어내지 않는다.',
    '2. 데이터 형태가 무엇이든(활동지 답변, 보고서, 교사 관찰 메모, 점수표 등) 먼저 학생이 실제로 한 것이 무엇인지 파악한 뒤 쓴다.',
    '3. 문항, 안내문, 평가 기준처럼 학생이 쓴 것이 아닌 내용은 근거로 삼지 않는다.',
    '4. [ ] 안의 열 이름은 데이터가 무엇인지 알려 주는 힌트로만 쓰고 본문에 옮겨 적지 않는다.',
    '5. 점수나 등급은 수행 수준을 가늠하는 데만 참고하고 본문에 숫자로 쓰지 않는다.',
    '6. [자료 안내]의 작성할 영역에 맞는 어휘와 관점으로 쓴다. 비어 있으면 데이터 내용에 가장 어울리는 영역으로 판단한다. 작성 순서/형식이 있으면 그대로 따른다.',
    "7. 활동을 나열하지 말고 '무엇을 했는지 → 어떻게 했는지 → 그 과정에서 드러난 역량' 흐름으로 엮는다.",
    "8. 데이터의 표현을 그대로 옮기지 말고 생활기록부 문어체로 다듬는다. '이 학생은', 학생 이름, 개인정보는 쓰지 않는다.",
    "9. 문장 끝은 '~함', '~임', '~보임' 같은 명사형 어미로 쓰되 같은 어미를 연달아 쓰지 않는다.",
    '10. record_draft는 나이스 바이트 기준 {{recordDraftMinBytes}}바이트 이상 {{recordDraftMaxBytes}}바이트 이하(공백 포함 약 {{recordDraftMinChars}}~{{recordDraftMaxChars}}자)로 쓴다.',
    '11. evidence_summary에는 근거로 쓴 데이터 내용을 짧게 정리하고, caution에는 교사가 확인할 점이 있을 때만 쓴다.',
  ].join('\n');
}

function buildRecordDraftNoEvidenceRule_() {
  return [
    '[자동 적용 최우선 제외 규칙]',
    '이 규칙은 설정 시트의 PROMPT_1 내용 및 record_draft 분량 조건보다 우선한다.',
    '학생 제출물 원문에 학생이 직접 작성한 답변, 활동 결과, 사고 과정, 탐구 내용이 확인되지 않으면 생활기록부를 작성하지 않는다.',
    '파일을 열 수 없음 또는 텍스트 자동 추출 미지원 문구가 있는 첨부파일은 해당 첨부파일만 근거에서 제외한다.',
    '같은 제출물 안의 다른 단답형 응답, 선택형 응답, Google Docs 본문, 텍스트 파일 등에서 학생이 직접 작성한 답변이 확인되면 그 확인 가능한 답변만 근거로 record_draft를 작성한다.',
    '제출하지 않은 과제, 빈 제출물, 모든 첨부파일을 열 수 없음, 모든 첨부파일이 텍스트 자동 추출 미지원, 교사 문항만 있고 학생 답변이 없는 경우는 근거 부족으로 본다.',
    '근거 부족 상황에서는 반드시 아래 JSON 구조처럼 답한다.',
    '{"evidence_summary":"", "record_draft":"", "caution":"학생 제출물 원문에서 생활기록부 작성에 사용할 수 있는 학생 수행 근거가 확인되지 않음"}',
    '근거 부족 상황에서는 record_draft 최소 분량 조건을 무시하고 record_draft를 반드시 빈 문자열로 둔다.',
    'record_draft에 "작성하지 않음", "근거 부족", "제출물 없음" 같은 설명 문장을 쓰지 않는다.',
  ].join('\n');
}

function buildStudentFinalNoEvidenceRule_() {
  return [
    '[자동 적용 최우선 제외 규칙]',
    '이 규칙은 설정 시트의 PROMPT_2 내용 및 record_draft 분량 조건보다 우선한다.',
    '과제별 근거 묶음에서 근거 부족, 제출물 없음, 미제출, 학생 답변 없음이 표시된 과제는 생활기록부 근거에서 제외한다.',
    '파일을 열 수 없음 또는 텍스트 자동 추출 미지원 표시는 해당 첨부파일만 제외하라는 뜻이며, 같은 과제의 다른 확인 가능한 근거 요약이 있으면 그 과제 전체를 제외하지 않는다.',
    '근거 요약이 비어 있거나 학생의 직접 수행이 확인되지 않는 과제도 제외한다.',
    '제외 후 사용할 수 있는 학생 수행 근거가 하나도 없으면 생활기록부를 작성하지 않는다.',
    '이 경우 반드시 아래 JSON 구조처럼 답한다.',
    '{"evidence_summary":"", "record_draft":"", "caution":"생활기록부 작성에 사용할 수 있는 학생 수행 근거가 확인되지 않음"}',
    '근거 부족 상황에서는 record_draft 최소 분량 조건을 무시하고 record_draft를 반드시 빈 문자열로 둔다.',
    'record_draft에 "작성하지 않음", "근거 부족", "제출물 없음" 같은 설명 문장을 쓰지 않는다.',
  ].join('\n');
}

function buildGradeContextRule_() {
  return [
    '[점수 정보 사용 규칙]',
    '점수 정보는 제출물의 수행 수준을 판단하기 위한 중요한 보조 정보로 참고한다.',
    '교사 입력 점수/수준이 있으면 Classroom draftGrade 또는 assignedGrade보다 우선 적용한다.',
    '교사 입력 점수/수준이 없으면 만점 대비 draftGrade를 우선 참고하고, draftGrade가 없으면 assignedGrade를 참고한다.',
    '점수가 높을수록 제출물에서 확인되는 성취와 수행 수준을 더 적극적으로 반영하고, 점수가 낮으면 과장된 표현을 피한다.',
    '단, 점수만으로 제출물 원문에 없는 역량, 태도, 활동을 새로 꾸며내지 않는다.',
    'record_draft에는 점수, 만점, 등급, 점수 비율을 직접 쓰지 않는다.',
    'evidence_summary에는 필요한 경우 점수 정보를 내부 참고용으로 간단히 남길 수 있다.',
    '점수 정보가 없으면 학생 제출물 원문에서 확인되는 내용만 근거로 작성한다.',
  ].join('\n');
}

function formatGradeContext_(draftGrade, assignedGrade, maxPoints, teacherGrade) {
  const draft = String(draftGrade ?? '').trim();
  const assigned = String(assignedGrade ?? '').trim();
  const max = String(maxPoints ?? '').trim();
  const teacher = String(teacherGrade ?? '').trim();

  if (teacher) {
    const teacherMax = max || (isPlainNumericGrade_(teacher) ? '100' : '');
    return `교사 입력 점수/수준: ${formatGradeValueWithMax_(teacher, teacherMax)} (Classroom 점수보다 우선 적용)`;
  }

  const gradeParts = [];

  if (draft) gradeParts.push(`Classroom 초안점수 ${formatGradeValueWithMax_(draft, max)}`);
  if (assigned) gradeParts.push(`Classroom 확정점수 ${formatGradeValueWithMax_(assigned, max)}`);

  if (gradeParts.length > 0) return gradeParts.join(', ');
  if (max) return `만점 ${max}, 학생 점수 없음`;
  return '점수 정보 없음';
}

function formatGradeValueWithMax_(grade, maxPoints) {
  const gradeText = String(grade ?? '').trim();
  const max = String(maxPoints ?? '').trim();

  if (!gradeText) return '점수 없음';
  if (!isPlainNumericGrade_(gradeText)) return gradeText;

  const ratioText = formatGradeRatio_(gradeText, max);
  if (max) return `${gradeText}/${max}${ratioText}`;

  return gradeText;
}

function formatGradeRatio_(grade, maxPoints) {
  const gradeNumber = parseGradeNumber_(grade);
  const maxNumber = parseGradeNumber_(maxPoints);

  if (gradeNumber === null || maxNumber === null || maxNumber <= 0) return '';

  const percent = Math.round((gradeNumber / maxNumber) * 100);
  return ` (${percent}%)`;
}

function parseGradeNumber_(value) {
  const text = String(value ?? '').trim().replace(/,/g, '');
  const match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) return null;

  const number = Number(match[0]);
  return Number.isFinite(number) ? number : null;
}

function isPlainNumericGrade_(value) {
  return /^-?\d+(?:\.\d+)?$/.test(String(value ?? '').trim());
}

function getConfigPromptTemplate_(config, key, fallback) {
  return String(config[key] || '').trim() || fallback;
}

function renderPromptTemplate_(template, data) {
  const values = {};

  Object.keys(data || {}).forEach(key => {
    values[key.toLowerCase()] = data[key];
  });

  return String(template || '')
    .replace(/\{\{\s*([A-Za-z0-9_]+)\s*\}\}/g, (match, key) => {
      const value = values[String(key || '').toLowerCase()];
      return value === undefined || value === null ? '' : String(value);
    })
    .trim();
}

function promptTemplateHasPlaceholder_(template, key) {
  const pattern = new RegExp('\\{\\{\\s*' + key + '\\s*\\}\\}', 'i');
  return pattern.test(String(template || ''));
}

function promptTemplateHasAnyPlaceholder_(template, keys) {
  return keys.some(key => promptTemplateHasPlaceholder_(template, key));
}

function buildMissingPromptLines_(template, placeholderLinePairs) {
  return placeholderLinePairs
    .filter(pair => !promptTemplateHasPlaceholder_(template, pair[0]))
    .map(pair => pair[1]);
}
function listAllCourses_() {
  const all = [];
  let pageToken = null;

  do {
    const optionalArgs = {
      pageSize: 100,
      courseStates: ['ACTIVE'],
      teacherId: 'me',   // 핵심: 내가 교사로 등록된 수업만 가져오기
    };

    if (pageToken) optionalArgs.pageToken = pageToken;

    const res = Classroom.Courses.list(optionalArgs);

    if (res.courses && res.courses.length > 0) {
      all.push(...res.courses);
    }

    pageToken = res.nextPageToken || null;
  } while (pageToken);

  return all;
}

function listAllCourseWork_(courseId) {
  const all = [];
  let pageToken = null;

  do {
    const optionalArgs = {
      pageSize: 100,
      orderBy: 'updateTime desc',
    };

    if (pageToken) optionalArgs.pageToken = pageToken;

    const res = Classroom.Courses.CourseWork.list(courseId, optionalArgs);

    if (res.courseWork && res.courseWork.length > 0) {
      all.push(...res.courseWork);
    }

    pageToken = res.nextPageToken || null;
  } while (pageToken);

  return all;
}

function listAllStudentSubmissions_(courseId, courseWorkId) {
  const all = [];
  let pageToken = null;

  do {
    const optionalArgs = {
      pageSize: 100,
    };

    if (pageToken) optionalArgs.pageToken = pageToken;

    const res = Classroom.Courses.CourseWork.StudentSubmissions.list(
      courseId,
      courseWorkId,
      optionalArgs
    );

    if (res.studentSubmissions && res.studentSubmissions.length > 0) {
      all.push(...res.studentSubmissions);
    }

    pageToken = res.nextPageToken || null;
  } while (pageToken);

  return all;
}

function getUserProfile_(userId) {
  return Classroom.UserProfiles.get(userId);
}
function testMyClassroomCoursesOnly() {
  const res = Classroom.Courses.list({
    pageSize: 20,
    courseStates: ['ACTIVE'],
    teacherId: 'me',
  });

  const courses = res.courses || [];

  Logger.log(`내가 교사인 수업 수: ${courses.length}`);

  courses.forEach(course => {
    Logger.log(`${course.name} / ${course.id}`);
  });

  getUi_().alert(`내가 교사인 수업만 조회 성공: ${courses.length}개`);
}

function isMyTeacherCourse_(courseId) {
  let pageToken = null;

  do {
    const optionalArgs = {
      pageSize: 100,
      courseStates: ['ACTIVE'],
      teacherId: 'me',
    };

    if (pageToken) optionalArgs.pageToken = pageToken;

    const res = Classroom.Courses.list(optionalArgs);
    const courses = res.courses || [];

    if (courses.some(course => String(course.id) === String(courseId))) {
      return true;
    }

    pageToken = res.nextPageToken || null;
  } while (pageToken);

  return false;
}


function createRecordDraftWithLength_(provider, model, systemGuide, userPrompt, minBytes, maxBytes, reasoning) {
  let result = callAi_(provider, model, systemGuide, userPrompt, reasoning);
  let parsed = normalizeNoEvidenceRecordResponse_(parseClaudeJson_(result.text));

  let draft = parsed.record_draft || '';
  let len = neisByteCount_(draft);

  if (shouldKeepEmptyRecordDraft_(parsed) || (len >= minBytes && len <= maxBytes)) {
    return {
      parsed,
      usage: result.usage || {},
      retry: false,
    };
  }

  if (isNearByteRange_(len, minBytes, maxBytes)) {
    const note = nearByteRangeNote_(len, minBytes, maxBytes);
    const caution = String(parsed.caution || '').trim();
    parsed.caution = caution ? `${caution}\n${note}` : note;

    return {
      parsed,
      usage: result.usage || {},
      retry: false,
    };
  }

  const retryPrompt = [
    userPrompt,
    '',
    '[재작성 요청]',
    `방금 생성한 record_draft는 나이스 바이트 기준 ${len}바이트(공백 포함 ${draft.length}자)이다.`,
    `record_draft 값만 ${formatByteRange_({ min: minBytes, max: maxBytes })}(공백 포함 약 ${approxCharsFromBytes_(minBytes)}~${approxCharsFromBytes_(maxBytes)}자)가 되도록 ${len < minBytes ? '근거 안에서 내용을 보태어 늘려' : '덜 중요한 내용을 덜어 줄여'} 다시 작성하라.`,
    '단, 근거 부족 또는 제출물 없음으로 record_draft를 빈 문자열로 둔 경우에는 분량 조건을 무시하고 빈 문자열을 유지하라.',
    'evidence_summary와 caution은 간단히 유지해도 된다.',
    '반드시 JSON 형식으로만 답하라.',
    '',
    '[이전 record_draft]',
    draft,
  ].join('\n');

  // 재요청이 실패하거나 빈 답을 주거나 범위를 벗어나도 결과는 남긴다.
  // 두 답 중 범위에 더 가까운 것을 쓰고, 범위를 벗어났으면 caution에 적는다.
  const first = { parsed, usage: result.usage || {}, len };
  let chosen = first;
  let retryNote = '재요청 후에도 범위에 들어오지 않음.';

  try {
    const retryResult = callAi_(provider, model, systemGuide, retryPrompt, reasoning);
    const retryParsed = normalizeNoEvidenceRecordResponse_(parseClaudeJson_(retryResult.text));
    const retryLen = neisByteCount_(retryParsed.record_draft || '');

    if (retryLen > 0 && charRangeDistance_(retryLen, minBytes, maxBytes) <= charRangeDistance_(len, minBytes, maxBytes)) {
      chosen = { parsed: retryParsed, usage: retryResult.usage || {}, len: retryLen };
    }
  } catch (err) {
    retryNote = `재요청 응답 처리 실패: ${String(err.message || err)}`;
    log_('createRecordDraftWithLength_', 'RETRY_ERROR', retryNote);
  }

  if (charRangeDistance_(chosen.len, minBytes, maxBytes) > 0) {
    const note = `설정 분량 범위(${formatByteRange_({ min: minBytes, max: maxBytes })})를 벗어남: ${chosen.len}바이트. ${retryNote}`;
    const caution = String(chosen.parsed.caution || '').trim();
    chosen.parsed.caution = caution ? `${caution}\n${note}` : note;
  }

  return {
    parsed: chosen.parsed,
    usage: chosen.usage,
    retry: true,
  };
}

// 길이가 [min, max] 범위에서 얼마나 벗어났는지. 범위 안이면 0.
// 설정 시트의 분량 기준(나이스 바이트). 비어 있거나 잘못된 값이면 기본값을 쓴다.
function getRecordByteRange_(config, kind) {
  const setting = RECORD_BYTE_RANGE_SETTINGS[kind];
  let min = parsePositiveInteger_(config && config[setting[0]]) || setting[2];
  let max = parsePositiveInteger_(config && config[setting[1]]) || setting[3];

  if (min > max) {
    const tmp = min;
    min = max;
    max = tmp;
  }

  return { min, max, enabled: true };
}

// 한글 생기부는 공백을 섞어 1자에 약 2.7바이트다. AI에 대략의 글자수를 함께 알려 줄 때 쓴다.
function approxCharsFromBytes_(bytes) {
  return Math.round(Number(bytes || 0) / 2.7);
}

function formatByteRange_(range) {
  return `${range.min}바이트 이상 ${range.max}바이트 이하`;
}

function buildByteRangeInstruction_(field, range) {
  return [
    '[분량 조건]',
    `${field}는 나이스 바이트 기준 ${formatByteRange_(range)}로 작성한다. 공백 포함 약 ${approxCharsFromBytes_(range.min)}~${approxCharsFromBytes_(range.max)}자에 해당한다.`,
    '나이스 바이트는 한글·한자·특수문자 1자를 3바이트, 영문·숫자·공백·일반 문장부호 1자를 1바이트로 센다.',
    '이 분량 조건이 위에 적힌 다른 글자수 안내보다 우선한다.',
  ].join('\n');
}

function charRangeDistance_(len, minChars, maxChars) {
  if (len < minChars) return minChars - len;
  if (len > maxChars) return len - maxChars;
  return 0;
}

// 설정 범위를 이 비율 이내로 벗어난 분량은 재요청하지 않는다. 재요청하면 그 학생은 시간과 비용이 두 배가 든다.
// 위쪽 여유는 나이스 최대 바이트(설정 NEIS_MAX_BYTES, 기본 1500)를 넘지 않게 한다.
const BYTE_RETRY_TOLERANCE = 0.1;

function isNearByteRange_(len, minBytes, maxBytes) {
  let neisMax = 1500;
  try {
    neisMax = parsePositiveInteger_(getConfigMap_().NEIS_MAX_BYTES) || neisMax;
  } catch (err) {
    // 설정을 못 읽으면 기본값을 쓴다.
  }

  const low = Math.floor(minBytes * (1 - BYTE_RETRY_TOLERANCE));
  const high = Math.max(maxBytes, Math.min(Math.ceil(maxBytes * (1 + BYTE_RETRY_TOLERANCE)), neisMax));
  return len >= low && len <= high;
}

function nearByteRangeNote_(len, minBytes, maxBytes) {
  return `설정 분량 범위(${formatByteRange_({ min: minBytes, max: maxBytes })})를 조금 벗어남: ${len}바이트. 10% 이내라 재요청하지 않음.`;
}

function normalizeNoEvidenceRecordResponse_(parsed) {
  parsed = parsed || {};

  if (shouldSuppressRecordDraft_(parsed)) {
    parsed.record_draft = '';
  }

  return parsed;
}

function shouldKeepEmptyRecordDraft_(parsed) {
  return !String((parsed && parsed.record_draft) || '').trim() &&
    !!String((parsed && parsed.caution) || '').trim();
}

// 근거가 없으면 프롬프트대로 AI가 record_draft를 비워 보낸다. 초안이 왔으면 지우지 않는다.
// (예전에는 caution·evidence_summary의 단어로 근거 부족을 짐작해 초안을 지웠는데,
// "생활기록부 작성 시 … 점수 정보가 없어"처럼 평범한 확인 문장에도 걸려 멀쩡한 초안이 사라졌다.)
function shouldSuppressRecordDraft_(parsed) {
  const draft = String((parsed && parsed.record_draft) || '').trim();
  const caution = String((parsed && parsed.caution) || '').trim();

  return !draft && !!caution;
}

function hasExtractedTextWarning_(text) {
  return /파일을\s*열\s*수\s*없음|텍스트\s*자동\s*추출\s*미지원/.test(String(text || ''));
}

function getExtractedTextForStatus_(text) {
  return String(text || '')
    .split(/\r?\n/)
    .filter(line => {
      const value = line.trim();
      return value &&
        !/^\[첨부파일:/.test(value) &&
        !/^파일을\s*열\s*수\s*없음/.test(value) &&
        !/^텍스트\s*자동\s*추출\s*미지원/.test(value) &&
        !/^\[텍스트\s*자동\s*추출\s*미지원/.test(value) &&
        !/^파일\s*링크를\s*직접\s*확인하세요/.test(value) &&
        !/^https?:\/\//i.test(value);
    })
    .join('\n')
    .trim();
}

function hasUsableExtractedText_(text) {
  const value = String(text || '').trim();
  if (!value) return false;
  if (!hasExtractedTextWarning_(value)) return value.length >= 20;
  return getExtractedTextForStatus_(value).length >= 20;
}

function clearClaudeApiKey() {
  PropertiesService.getUserProperties().deleteProperty('CLAUDE_API_KEY');
  PropertiesService.getScriptProperties().deleteProperty('CLAUDE_API_KEY');

  log_('clearClaudeApiKey', 'OK', 'Claude API 키 삭제 완료');
  getUi_().alert('현재 사용자 및 Script Properties의 Claude API 키를 삭제했습니다.');
}


function applyFrontendLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  styleSheetBase_(SHEET_NAMES.config);
  styleSheetBase_(SHEET_NAMES.assignments);
  styleSheetBase_(SHEET_NAMES.students);
  styleSheetBase_(SHEET_NAMES.submissions);
  styleSheetBase_(SHEET_NAMES.records);
  styleSheetBase_(SHEET_NAMES.commonPhrases);
  styleSheetBase_(SHEET_NAMES.behaviorRecommendations);
  styleSheetBase_(SHEET_NAMES.logs);

  styleConfigSheet_();
  styleCommonPhraseSheet_();
  styleBehaviorRecommendationSheet_();
  styleAssignmentsSheet_();
  styleSubmissionsSheet_();
  styleRecordsSheet_();
  styleStudentFinalRecordsSheet_();
  styleLogsSheet_();

  refreshDashboard();
  organizeSheetTabs_();
  getUi_().alert('화면 정리와 서식 적용을 완료했습니다.');
}

function styleSheetBase_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return;

  const lastCol = sh.getLastColumn();
  const lastRow = Math.max(sh.getLastRow(), 1);

  if (lastCol < 1) return;

  sh.setFrozenRows(1);

  const header = sh.getRange(1, 1, 1, lastCol);
  header
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.getRange(1, 1, lastRow, lastCol)
    .setVerticalAlignment('top')
    .setFontSize(10);

  try {
    const filter = sh.getFilter();
    if (filter) filter.remove();
    sh.getRange(1, 1, lastRow, lastCol).createFilter();
  } catch (err) {
    // 필터 생성 실패는 화면 기능에 치명적이지 않으므로 무시
  }

  sh.setRowHeight(1, 32);

  if (lastRow > 1) {
    sh.setRowHeights(2, lastRow - 1, 34);
  }
}

function styleConfigSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.config);
  if (!sh) return;

  const lastRow = Math.max(sh.getLastRow(), 1);

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 210);
  sh.setColumnWidth(2, 680);
  sh.setColumnWidth(3, 420);

  sh.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.getRange(1, 1, lastRow, 3)
    .setVerticalAlignment('middle')
    .setFontSize(10);

  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.getRange(2, 2, lastRow - 1, 2)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');
    sh.autoResizeRows(2, lastRow - 1);
  }

  applyConfigValidations_(sh);
}

function styleCommonPhraseSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.commonPhrases);
  if (!sh) return;

  ensureCommonPhraseSheet_();

  const lastRow = Math.max(sh.getLastRow(), 2);

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 360);
  sh.setColumnWidth(2, 720);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 90);

  sh.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.getRange(1, 1, lastRow, 4)
    .setFontSize(10)
    .setVerticalAlignment('top');

  sh.getRange(2, 1, Math.max(lastRow - 1, 1), 2)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sh.getRange(2, 3, Math.max(lastRow - 1, 1), 2)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.getRange('C2').setFormula('=IF(B2="","",2*LENB(B2)-LEN(B2))');

  if (lastRow > 1) {
    sh.setRowHeights(2, lastRow - 1, 150);
  }

  if (lastRow >= 3) {
    const formulaRows = [];

    for (let row = 3; row <= lastRow; row++) {
      formulaRows.push([`=IF(B${row}="","",2*LENB(B${row})-LEN(B${row}))`]);
    }

    sh.getRange(3, 3, formulaRows.length, 1).setFormulas(formulaRows);
  }

  try {
    const filter = sh.getFilter();
    if (filter) filter.remove();
    sh.getRange(1, 1, lastRow, 4).createFilter();
  } catch (err) {
    // 필터 생성 실패는 무시
  }
}

function styleAssignmentsSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.assignments);
  if (!sh) return;

  setColumnWidthsByHeader_(sh, {
    courseId: 120,
    courseName: 220,
    section: 100,
    courseWorkId: 140,
    assignmentTitle: 320,
    workType: 100,
    state: 90,
    dueAt: 140,
    updateTime: 160,
    alternateLink: 220,
    maxPoints: 80,
    teacherGrade: 120,
    includeInFinal: 115,
    weight: 70,
    category: 120,
    collectStatus: 130,
    memo: 220,
  });

  const h = headerMap_(sh);
  const lastRow = getLastDataRowByHeaders_(sh, ['courseId', 'courseWorkId', 'assignmentTitle']);
  clearRowsBelowLastData_(sh, lastRow);

  // 데이터 영역 기본 서식 초기화: 과제목록 본문은 항상 검정 글자
  const lastCol = sh.getLastColumn();

  if (lastRow > 1 && lastCol > 0) {
    sh.getRange(2, 1, lastRow - 1, lastCol)
      .setFontColor('#000000')
      .setBackground('#ffffff')
      .setFontWeight('normal');
  }

  if (h.includeInFinal) {
    sh.getRange(1, h.includeInFinal).setBackground('#274e13');

    if (lastRow > 1) {
      sh.getRange(2, h.includeInFinal, lastRow - 1, 1).insertCheckboxes();
    }
  }

  if (h.weight && lastRow > 1) {
    const range = sh.getRange(2, h.weight, lastRow - 1, 1);
    const values = range.getValues().map(r => [r[0] || 1]);
    range.setValues(values).setHorizontalAlignment('center');
  }

  ['includeInFinal', 'weight', 'maxPoints', 'teacherGrade', 'collectStatus'].forEach(header => {
    if (h[header] && lastRow > 1) {
      sh.getRange(2, h[header], lastRow - 1, 1)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }
  });

  if (h.assignmentTitle && lastRow > 1) {
    sh.getRange(2, h.assignmentTitle, lastRow - 1, 1).setWrap(true);
  }

  if (lastRow > 1) {
    sh.setRowHeights(2, lastRow - 1, 21);
  }

  applyAssignmentCourseConditionalFormats_(sh, lastRow);

  try {
    const filter = sh.getFilter();
    if (filter) filter.remove();
    sh.getRange(1, 1, Math.max(lastRow, 1), lastCol).createFilter();
  } catch (err) {
    // Ignore filter recreation failures; formatting can still be applied.
  }

  hideColumnsByHeaders_(sh, [
    'courseId',
    'courseWorkId',
    'workType',
    'state',
    'dueAt',
    'updateTime',
    'maxPoints',
    'teacherGrade',
    'assignmentDescription',
    'alternateLink'
  ]);

  resizeBlankRowsAfterLastData_(sh, lastRow, ASSIGNMENTS_TRAILING_BLANK_ROWS);

}

function applyAssignmentCourseConditionalFormats_(sh, lastRow) {
  const h = headerMap_(sh);
  if (!h.courseName) return;

  lastRow = lastRow || getLastDataRowByHeaders_(sh, ['courseName', 'assignmentTitle']);
  if (lastRow < 2) {
    sh.setConditionalFormatRules([]);
    return;
  }

  const lastCol = sh.getLastColumn();
  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
  const courseValues = sh.getRange(2, h.courseName, lastRow - 1, 1).getValues();
  const courseNames = [];
  const seen = {};

  courseValues.forEach(row => {
    const courseName = String(row[0] || '').trim();
    if (!courseName || seen[courseName]) return;

    seen[courseName] = true;
    courseNames.push(courseName);
  });

  const colors = [
    '#eaf4ff',
    '#eaf7ea',
    '#fff4d8',
    '#fcebea',
    '#eef0ff',
    '#e8f7f4',
    '#f7ecff',
    '#f2f2f2',
    '#fff0e0',
    '#e9f2ec',
  ];
  const courseNameCol = colToA1_(h.courseName);

  const rules = courseNames.map((courseName, index) => {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=TRIM($${courseNameCol}2)=${toSheetFormulaString_(courseName)}`)
      .setBackground(colors[index % colors.length])
      .setRanges([dataRange])
      .build();
  });

  sh.setConditionalFormatRules(rules);
}

function toSheetFormulaString_(value) {
  return `"${String(value || '').replace(/"/g, '""')}"`;
}


function styleSubmissionsSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.submissions);
  if (!sh) return;
  const h = headerMap_(sh);

  setColumnWidthsByHeader_(sh, {
    submissionKey: 80,
    studentNo: 80,
    studentName: 120,
    email: 180,
    courseId: 80,
    courseName: 180,
    courseWorkId: 80,
    assignmentTitle: 260,
    submissionId: 80,
    state: 100,
    late: 70,
    draftGrade: 90,
    assignedGrade: 80,
    maxPoints: 80,
    teacherGrade: 120,
    updateTime: 150,
    submissionLink: 220,
    fileTitles: 260,
    fileLinks: 220,
    extractedText: 500,
    textChars: 80,
    collectedAt: 150,
    aiStatus: 110,
  });

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).setWrap(false);

    if (h.fileTitles) sh.getRange(2, h.fileTitles, lastRow - 1, 1).setWrap(true);
    if (h.fileLinks) sh.getRange(2, h.fileLinks, lastRow - 1, 1).setWrap(true);
    if (h.extractedText) sh.getRange(2, h.extractedText, lastRow - 1, 1).setWrap(true);

    ['late', 'draftGrade', 'assignedGrade', 'maxPoints', 'teacherGrade', 'textChars', 'aiStatus'].forEach(header => {
      if (h[header]) {
        sh.getRange(2, h[header], lastRow - 1, 1)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    });

    sh.setRowHeights(2, lastRow - 1, 21);
  }

  if (h.teacherGrade) {
    sh.getRange(1, h.teacherGrade).setNote(
      '[선택입력] 클래스룸에 과제를 채점하셔서 점수를 입력하지 않으셨거나, 클래스룸 점수와 다르게 점수를 입력하고 싶으시다면 기입해주세요. 클래스룸에 특별한 입력이 없다면 100점 만점 기준입니다.'
    );
  }

  if (h.maxPoints) {
    sh.getRange(1, h.maxPoints).setNote(
      'Classroom 과제의 만점입니다. teacherGrade에 숫자만 입력하면 이 만점 대비 점수로 해석됩니다.'
    );
  }



  hideColumnsByHeaders_(sh, [
    'submissionKey',
    'courseId',
    'courseWorkId',
    'submissionId',

    'state',
    'late',
    
    'assignedGrade',
    'updateTime',
    'fileTitles',

    'fileLinks',
    'extractedText',
    'assignmentDescription'
  ]);


  applySubmissionConditionalFormats_(sh);
}

function styleRecordsSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.records);
  if (!sh) return;

  const filter = sh.getFilter();
  if (filter) filter.remove();

  const h = headerMap_(sh);
  const lastDataRow = getLastDataRowByHeaders_(sh, [
    'submissionKey',
    'studentName',
    'assignmentTitle',
    'finalRecord',
    'recordDraft',
    'createdAt'
  ]);
  clearRowsBelowLastData_(sh, lastDataRow);
  const lastCol = sh.getLastColumn();

  sh.setFrozenRows(1);
  sh.setFrozenColumns(6);

  const headerRange = sh.getRange(1, 1, 1, lastCol);
  headerRange
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.setRowHeight(1, 34);

  setColumnWidthsByHeader_(sh, {
    submissionKey: 80,
    studentNo: 70,
    studentName: 90,
    email: 160,
    courseName: 150,
    assignmentTitle: 220,
    finalRecord: 620,
    charCount: 85,
    caution: 260,
    teacherChecked: 85,
    createdAt: 140,
    evidenceSummary: 320,
    recordDraft: 420,
    gradeText: 260,
    model: 180,
    usage: 180,
  });

  if (lastDataRow > 1) {
    const dataRange = sh.getRange(2, 1, lastDataRow - 1, lastCol);

    dataRange
      .setVerticalAlignment('top')
      .setFontSize(10)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // 핵심 검토 열은 줄바꿈
    ['finalRecord', 'caution', 'gradeText'].forEach(header => {
      if (h[header]) {
        sh.getRange(2, h[header], lastDataRow - 1, 1)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top');
      }
    });

    // 과제명도 너무 길 수 있으므로 줄바꿈
    if (h.assignmentTitle) {
      sh.getRange(2, h.assignmentTitle, lastDataRow - 1, 1)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }

    // 학생명, 글자수, 체크박스 등은 가운데 정렬
    ['studentNo', 'studentName', 'charCount', 'teacherChecked'].forEach(header => {
      if (h[header]) {
        sh.getRange(2, h[header], lastDataRow - 1, 1)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    });

    // finalRecord가 보기 좋게 보이도록 행 높이 확보
    sh.setRowHeights(2, lastDataRow - 1, 135);
    applyEvidenceSummaryNotesToFinalRecords_(
      Array.from({ length: lastDataRow - 1 }, (_, index) => index + 2)
    );
  }

  if (h.teacherChecked && lastDataRow > 1) {
    sh.getRange(2, h.teacherChecked, lastDataRow - 1, 1).insertCheckboxes();
  }

  // 숨길 열
  hideColumnsByHeaders_(sh, [
    'submissionKey',
    'email',
    'evidenceSummary',
    'recordDraft',
    'usage'
  ]);

  // 핵심 열 강조
  if (h.finalRecord) {
    sh.getRange(1, h.finalRecord).setBackground('#274e13');
  }

  if (h.charCount) {
    sh.getRange(1, h.charCount).setBackground('#7f6000');
  }

  if (h.caution) {
    sh.getRange(1, h.caution).clearNote();
  }

  applyNeisByteFormulasToRecords_();
  applyRecordConditionalFormats_(sh);

  try {
    sh.getRange(1, 1, Math.max(lastDataRow, 1), lastCol).createFilter();
  } catch (err) {
    // 필터 생성 실패는 무시
  }
}

function cleanEmptyRecordRows_() {
  const sh = getSheet_(SHEET_NAMES.records);
  const h = headerMap_(sh);

  const keyCol = h.submissionKey || 1;
  const lastDataRow = Math.max(getLastNonEmptyRowInColumn_(sh, keyCol), 1);
  const maxRows = sh.getMaxRows();

  if (maxRows > lastDataRow + 20) {
    sh.deleteRows(lastDataRow + 21, maxRows - lastDataRow - 20);
  }

  log_('cleanEmptyRecordRows_', 'OK', `생기부초안 빈 하단 행 정리 완료. 실제 마지막 데이터 행: ${lastDataRow}`);
  getUi_().alert(`생기부초안 빈 하단 행을 정리했습니다.\n실제 마지막 데이터 행: ${lastDataRow}`);
}
function styleLogsSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.logs);
  if (!sh) return;

  setColumnWidthsByHeader_(sh, {
    timestamp: 150,
    action: 260,
    status: 100,
    message: 600,
  });

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).setWrap(true);
  }
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const activeSheet = ss.getActiveSheet();
  const activeRange = activeSheet ? activeSheet.getActiveRange() : null;

  const sh = getOrCreateSheet_(DASHBOARD_SHEET_NAME);

  sh.clear();

  const submissions = getSheetObjects_(SHEET_NAMES.submissions);
  const records = getSheetObjects_(SHEET_NAMES.records);
  const logs = getSheetObjects_(SHEET_NAMES.logs);
  const config = getConfigMap_();

  const neisMaxBytes = Number(config.NEIS_MAX_BYTES || 1500);
  const neisWarningBytes = Number(config.NEIS_WARNING_BYTES || 1400);

  const submissionCount = submissions.length;
  const extractionSuccessCount = submissions.filter(r => hasUsableExtractedText_(r.extractedText)).length;

  const extractionFailCount = submissionCount - extractionSuccessCount;

  const recordCount = records.length;
  const checkedCount = records.filter(r => r.teacherChecked === true || String(r.teacherChecked).toUpperCase() === 'TRUE').length;
  const retryCount = submissions.filter(r => String(r.aiStatus || '') === '재시도필요').length;
  const errorCount = submissions.filter(r => String(r.aiStatus || '') === '오류').length;


  const overneisByteCount = records.filter(r => {
    const n = Number(r.charCount || 0);
    return n > neisMaxBytes;
  }).length;

  const dashboardRows = [
    ['생기부 자동화 현황판', '', '', ''],
    ['항목', '값', '기준/설명', ''],
    ['수집된 제출물', submissionCount, '제출물 시트 기준', ''],
    ['원문 추출 성공', extractionSuccessCount, 'Google Docs 텍스트 추출 성공 추정', ''],
    ['원문 추출 실패/확인 필요', extractionFailCount, '권한/파일형식/빈 문서 확인', ''],
    ['생성된 생기부 초안', recordCount, '생기부초안 시트 기준', ''],
    ['교사 검토 완료', checkedCount, 'teacherChecked 체크 기준', ''],
    ['재시도 필요', retryCount, 'Claude 529/일시 오류 등', ''],
    ['오류', errorCount, '처리 실패', ''],
    ['나이스 바이트 초과', overneisByteCount, `${neisMaxBytes}바이트 초과 금지`, ''],
    ['', '', '', ''],
    ['최근 로그', '', '', ''],
    ['시간', '작업', '상태', '메시지'],
  ];

  sh.getRange(1, 1, dashboardRows.length, 4).setValues(dashboardRows);

  const recentLogs = logs.slice(-10).reverse().map(r => [
    r.timestamp || '',
    r.action || '',
    r.status || '',
    r.message || '',
  ]);

  if (recentLogs.length > 0) {
    sh.getRange(dashboardRows.length + 1, 1, recentLogs.length, 4).setValues(recentLogs);
  }

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 260);
  sh.setColumnWidth(4, 650);

  sh.getRange('A1:D1')
    .merge()
    .setFontWeight('bold')
    .setFontSize(16)
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center');

  sh.getRange('A2:D2')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#5b9bd5')
    .setHorizontalAlignment('center');

  sh.getRange('A12:D12')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79');

  sh.getRange('A13:D13')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#5b9bd5')
    .setHorizontalAlignment('center');

  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 4)
    .setVerticalAlignment('top')
    .setWrap(true);

  const warningRanges = [
    ['A5:D5', extractionFailCount > 0],
    ['A8:D8', retryCount > 0],
    ['A9:D9', errorCount > 0],
    ['A10:D10', overneisByteCount > 0],
  ];

  warningRanges.forEach(item => {
    if (item[1]) {
      sh.getRange(item[0]).setBackground('#fff2cc');
    }
  });

  try {
    if (activeSheet) {
      activeSheet.activate();

      if (activeRange) {
        activeSheet.setActiveRange(activeRange);
      }
    }
  } catch (err) {
    // 원래 시트/셀 복귀 실패는 무시
  }


}

function toggleExtractedTextColumn() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.submissions);
  if (!sh) return;

  const h = headerMap_(sh);
  const col = h.extractedText;

  if (!col) {
    getUi_().alert('extractedText 열을 찾을 수 없습니다.');
    return;
  }

  if (sh.isColumnHiddenByUser(col)) {
    sh.showColumns(col);
    getUi_().alert('extractedText 원문 열을 표시했습니다.');
  } else {
    sh.hideColumns(col);
    getUi_().alert('extractedText 원문 열을 숨겼습니다.');
  }
}

function applySubmissionConditionalFormats_(sh) {
  const h = headerMap_(sh);
  const maxRows = sh.getMaxRows();
  const rules = [];

  if (h.studentName) {
    const studentNameCol = colToA1_(h.studentName);
    const dataRange = sh.getRange(2, 1, maxRows - 1, sh.getLastColumn());

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($${studentNameCol}2<>"",ISEVEN(COUNTUNIQUE($${studentNameCol}$2:$${studentNameCol}2)))`)
        .setBackground('#e2f0ff')
        .setRanges([dataRange])
        .build()
    );
  }

  if (h.aiStatus) {
    const range = sh.getRange(2, h.aiStatus, maxRows - 1, 1);

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('생성완료')
        .setBackground('#d9ead3')
        .setRanges([range])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('재시도필요')
        .setBackground('#fff2cc')
        .setRanges([range])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('오류')
        .setBackground('#f4cccc')
        .setRanges([range])
        .build()
    );
  }

  if (h.extractedText) {
    const colLetter = colToA1_(h.extractedText);
    const warningRange = sh.getRange(2, h.extractedText, maxRows - 1, 1);

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=REGEXMATCH($${colLetter}2,"파일을 열 수 없음|텍스트 자동 추출 미지원")`)
        .setBackground('#fce4d6')
        .setRanges([warningRange])
        .build()
    );
  }

  sh.setConditionalFormatRules(rules);
}

function applyRecordConditionalFormats_(sh) {
  const h = headerMap_(sh);
  const maxRows = sh.getMaxRows();
  const config = getConfigMap_();

  const neisMaxBytes = Number(config.NEIS_MAX_BYTES || 1500);
  const neisWarningBytes = Number(config.NEIS_WARNING_BYTES || 1400);

  const rules = [];

  if (h.charCount) {
    const colLetter = colToA1_(h.charCount);
    const range = sh.getRange(2, h.charCount, maxRows - 1, 1);

    // 1500바이트 초과: 빨간색
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($${colLetter}2<>"",$${colLetter}2>${neisMaxBytes})`)
        .setBackground('#f4cccc')
        .setRanges([range])
        .build()
    );

    // 1400~1500바이트: 노란색
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($${colLetter}2<>"",$${colLetter}2>${neisWarningBytes},$${colLetter}2<=${neisMaxBytes})`)
        .setBackground('#fff2cc')
        .setRanges([range])
        .build()
    );
  }

  if (h.teacherChecked) {
    const colLetter = colToA1_(h.teacherChecked);
    const dataRange = sh.getRange(2, 1, maxRows - 1, sh.getLastColumn());

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$${colLetter}2=TRUE`)
        .setBackground('#d9ead3')
        .setRanges([dataRange])
        .build()
    );
  }

  sh.setConditionalFormatRules(rules);
}
function setColumnWidthsByHeader_(sh, widthMap) {
  const h = headerMap_(sh);

  Object.keys(widthMap).forEach(header => {
    if (h[header]) {
      sh.setColumnWidth(h[header], widthMap[header]);
    }
  });
}

function hideColumnsByHeaders_(sh, headers) {
  const h = headerMap_(sh);

  headers.forEach(header => {
    if (h[header]) {
      sh.hideColumns(h[header]);
    }
  });
}

function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName, 0);
  }

  return sh;
}

function getSheetObjects_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return [];

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h || '').trim());

  return values.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        if (header) obj[header] = row[i];
      });
      return obj;
    });
}

function colToA1_(col) {
  let temp = '';
  let letter = '';

  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }

  return letter;
}

function neisByteCount_(text) {
  text = String(text || '');

  let bytes = 0;

  for (const ch of text) {
    const code = ch.codePointAt(0);

    // NEIS 계산식 = 2*LENB - LEN 기준 근사
    // 영문/숫자/공백/일반 ASCII = 1바이트
    // 한글/한자/일본어 등 비ASCII = 3바이트
    if (code <= 0x7F) {
      bytes += 1;
    } else {
      bytes += 3;
    }
  }

  return bytes;
}

function applyNeisByteFormulasToRecords_(targetRows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.records);
  if (!sh) return;

  const h = headerMap_(sh);

  if (!h.charCount || !h.finalRecord) {
    getUi_().alert('charCount 또는 finalRecord 열을 찾을 수 없습니다.');
    return;
  }

  const finalRecordCol = colToA1_(h.finalRecord);

  if (targetRows && targetRows.length > 0) {
    targetRows.forEach(row => {
      sh.getRange(row, h.charCount).setFormula(
        `=IF($${finalRecordCol}${row}="","",2*LENB($${finalRecordCol}${row})-LEN($${finalRecordCol}${row}))`
      );
    });
  } else {
    const keyCol = h.submissionKey || 1;
    const lastDataRow = getLastNonEmptyRowInColumn_(sh, keyCol);
    if (lastDataRow < 2) return;

    const formulas = [];
    for (let row = 2; row <= lastDataRow; row++) {
      formulas.push([
        `=IF($${finalRecordCol}${row}="","",2*LENB($${finalRecordCol}${row})-LEN($${finalRecordCol}${row}))`
      ]);
    }

    sh.getRange(2, h.charCount, lastDataRow - 1, 1).setFormulas(formulas);
  }

  sh.getRange(1, h.charCount).setNote(
    'NEIS 바이트 계산값입니다. finalRecord 기준으로 =2*LENB(finalRecord)-LEN(finalRecord) 공식이 적용됩니다.'
  );

  if (h.caution) {
    sh.getRange(1, h.caution).clearNote();
  }
}

function applyEvidenceSummaryNotesToFinalRecords_(targetRows) {
  if (!targetRows || targetRows.length === 0) return;

  const sh = getSheet_(SHEET_NAMES.records);
  const h = headerMap_(sh);

  if (!h.finalRecord || !h.evidenceSummary) return;

  targetRows.forEach(row => {
    if (row < 2) return;

    const evidenceSummary = String(sh.getRange(row, h.evidenceSummary).getValue() || '').trim();
    sh.getRange(row, h.finalRecord).setNote(
      evidenceSummary ? `[evidence_summary]\n${evidenceSummary}` : ''
    );
  });
}

function showRecordsSheetAfterSave_(targetRows) {
  const sh = getSheet_(SHEET_NAMES.records);

  const filter = sh.getFilter();
  if (filter) filter.remove();

  const h = headerMap_(sh);
  applyEvidenceSummaryNotesToFinalRecords_(targetRows);

  if (targetRows && targetRows.length > 0) {
    targetRows.forEach(row => {
      sh.setRowHeight(row, 135);

      if (h.finalRecord) {
        sh.getRange(row, h.finalRecord)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top');
      }

      if (h.caution) {
        sh.getRange(row, h.caution)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top');
      }

      if (h.assignmentTitle) {
        sh.getRange(row, h.assignmentTitle)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      }

      ['studentNo', 'studentName', 'charCount', 'teacherChecked'].forEach(header => {
        if (h[header]) {
          sh.getRange(row, h[header])
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
        }
      });
    });
  }

  sh.activate();

  if (targetRows && targetRows.length > 0) {
    const row = targetRows[0];
    sh.setActiveRange(sh.getRange(row, h.finalRecord || 1));
  }

  try {
    const keyCol = h.submissionKey || 1;
    const lastDataRow = Math.max(getLastNonEmptyRowInColumn_(sh, keyCol), 2);
    sh.getRange(1, 1, lastDataRow, sh.getLastColumn()).createFilter();
  } catch (err) {
    // 필터 재생성 실패는 무시
  }
}

function appendRecordsOnly_(rows) {
  if (!rows || rows.length === 0) {
    return {
      appended: 0,
      updated: 0,
      targetRows: [],
    };
  }

  const sh = ensureSheetHeaders_(SHEET_NAMES.records);
  const targetRows = [];

  let nextRow = getFirstBlankRecordRow_(sh, 2);

  rows.forEach(row => {
    if (nextRow > sh.getMaxRows()) {
      sh.insertRowsAfter(sh.getMaxRows(), 20);
    }

    sh.getRange(nextRow, 1, 1, row.length).setValues([row]);
    targetRows.push(nextRow);

    nextRow = getFirstBlankRecordRow_(sh, nextRow + 1);
  });

  log_(
    'appendRecordsOnly_',
    'OK',
    `${SHEET_NAMES.records}: 새 초안 ${rows.length}개 추가, 행 ${targetRows.join(', ')}`
  );

  return {
    appended: rows.length,
    updated: 0,
    targetRows,
  };
}

function getFirstBlankRecordRow_(sh, startRow) {
  const h = headerMap_(sh);

  const coreHeaders = [
    'submissionKey',
    'studentName',
    'recordDraft',
    'finalRecord',
    'createdAt'
  ];

  const coreCols = coreHeaders
    .map(header => h[header])
    .filter(col => !!col);

  if (coreCols.length === 0) {
    throw new Error('생기부초안 시트의 핵심 헤더를 찾을 수 없습니다.');
  }

  const maxRows = sh.getMaxRows();
  const lastCol = sh.getLastColumn();
  const rowStart = Math.max(startRow || 2, 2);

  if (rowStart > maxRows) {
    return rowStart;
  }

  const values = sh.getRange(rowStart, 1, maxRows - rowStart + 1, lastCol).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const hasCoreData = coreCols.some(col => {
      return String(row[col - 1] || '').trim() !== '';
    });

    if (!hasCoreData) {
      return rowStart + i;
    }
  }

  return maxRows + 1;
}

function reorganizeRecordsSheetLayout_() {
  const sh = getSheet_(SHEET_NAMES.records);

  const filter = sh.getFilter();
  if (filter) filter.remove();

  const oldLastRow = sh.getLastRow();
  const oldLastCol = sh.getLastColumn();

  if (oldLastRow < 1 || oldLastCol < 1) return;

  const oldValues = sh.getRange(1, 1, oldLastRow, oldLastCol).getValues();
  const oldHeaders = oldValues[0].map(v => String(v || '').trim());

  const newHeaders = HEADERS[SHEET_NAMES.records];

  const oldIndex = {};
  oldHeaders.forEach((header, i) => {
    if (header) oldIndex[header] = i;
  });

  const newRows = [newHeaders];

  for (let r = 1; r < oldValues.length; r++) {
    const oldRow = oldValues[r];

    const hasAnyData = oldRow.some(cell => String(cell || '').trim() !== '');
    if (!hasAnyData) continue;

    const newRow = newHeaders.map(header => {
      if (oldIndex[header] !== undefined) {
        return oldRow[oldIndex[header]];
      }
      return '';
    });

    newRows.push(newRow);
  }

  sh.clear();

  sh.getRange(1, 1, newRows.length, newHeaders.length).setValues(newRows);

  // 남는 열이 있으면 정리
  if (sh.getMaxColumns() > newHeaders.length) {
    sh.deleteColumns(newHeaders.length + 1, sh.getMaxColumns() - newHeaders.length);
  }

  applyNeisByteFormulasToRecords_();
  styleRecordsSheet_();

  log_('reorganizeRecordsSheetLayout_', 'OK', `생기부초안 열 재배열 완료: ${newRows.length - 1}행`);
  getUi_().alert(`생기부초안 열 재배열을 완료했습니다.\n데이터 ${newRows.length - 1}행`);
}

function markSelectedSubmissionsPending() {
  const selection = getSelectedSubmissionRows_();

  if (selection.error) {
    getUi_().alert(selection.error);
    return;
  }

  const rows = selection.rows;
  markSubmissionRowsPending_(selection.sh, selection.h, rows);

  log_('markSelectedSubmissionsPending', 'OK', `${rows.length}개 제출물을 대기 상태로 표시`);
  refreshDashboard();

  getUi_().alert(`${rows.length}개 제출물을 대기 상태로 표시했습니다.`);
}

function getSelectedSubmissionRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  if (sh.getName() !== SHEET_NAMES.submissions) {
    return {
      error: '제출물 시트에서 학생 제출물 행을 선택한 뒤 실행하세요.',
    };
  }

  const h = headerMap_(sh);
  if (!h.aiStatus) {
    return {
      error: 'aiStatus 열을 찾을 수 없습니다.',
    };
  }

  const activeRangeList = ss.getActiveRangeList();
  const selectedRanges = activeRangeList ? activeRangeList.getRanges() : [sh.getActiveRange()];
  const selectedRows = new Set();

  selectedRanges.forEach(range => {
    if (!range || range.getSheet().getName() !== SHEET_NAMES.submissions) return;

    const startRow = Math.max(range.getRow(), 2);
    const endRow = range.getLastRow();

    for (let row = startRow; row <= endRow; row++) {
      selectedRows.add(row);
    }
  });

  const rows = Array.from(selectedRows).sort((a, b) => a - b);

  if (rows.length === 0) {
    return {
      error: '헤더가 아닌 제출물 행을 선택하세요.',
    };
  }

  return {
    sh,
    h,
    rows,
  };
}

function markSubmissionRowsPending_(sh, h, rows) {
  let groupStart = rows[0];
  let previousRow = rows[0];
  const rowGroups = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    if (row === previousRow + 1) {
      previousRow = row;
      continue;
    }

    rowGroups.push([groupStart, previousRow]);
    groupStart = row;
    previousRow = row;
  }

  rowGroups.push([groupStart, previousRow]);

  rowGroups.forEach(([startRow, endRow]) => {
    const numRows = endRow - startRow + 1;
    const values = Array.from({ length: numRows }, () => ['대기']);
    sh.getRange(startRow, h.aiStatus, numRows, 1).setValues(values);
  });

  return rows.length;
}

const RECORD_DRAFT_ERROR_KINDS = {
  auth: {
    status: '오류',
    shouldStop: true,
    reason: 'API 키 인증 오류',
    userMessage: 'API 키 인증에 실패했습니다. 저장된 API 키가 올바른지 확인하세요.',
  },
  permission: {
    status: '오류',
    shouldStop: true,
    reason: 'API 권한 오류',
    userMessage: 'API 권한 문제로 생성을 계속할 수 없습니다. API 키 권한 또는 프로젝트 설정을 확인하세요.',
  },
  billing: {
    status: '오류',
    shouldStop: true,
    reason: 'API 결제/사용량 한도',
    userMessage: 'API 결제 또는 사용량 한도 문제로 생성을 계속할 수 없습니다. 결제 상태와 사용량 한도를 확인하세요.',
  },
  rateLimit: {
    status: '재시도필요',
    shouldStop: true,
    reason: 'API 호출 한도 또는 속도 제한',
    userMessage: 'API 호출 한도 또는 속도 제한에 걸린 것으로 보입니다. 잠시 후 또는 할당량 갱신 후 다시 시작하세요.',
  },
  request: {
    status: '오류',
    shouldStop: true,
    reason: '모델명 또는 요청 형식 오류',
    userMessage: '모델명 또는 API 요청 형식 문제로 생성을 계속할 수 없습니다. 설정 시트의 모델명과 reasoning 값을 확인하세요.',
  },
  temporary: {
    status: '재시도필요',
    shouldStop: true,
    reason: 'API 일시 오류 또는 서버 과부하',
    userMessage: 'API 서버 과부하 또는 일시 오류로 보입니다. 잠시 후 다시 시작하세요.',
  },
  other: {
    status: '오류',
    shouldStop: false,
    reason: '기타 오류',
    userMessage: '개별 제출물 처리 중 오류가 발생했습니다. 로그 시트의 오류 메시지를 확인하세요.',
  },
};

// 오류 종류는 callClaude_/callOpenAi_/callGemini_가 붙이는 'API 오류 ###'의 HTTP 상태 코드로 판정한다.
// 메시지 전체에서 숫자를 찾으면 'JSON ... at position 1401'이나 요청 ID 속 숫자를 401 등으로 잘못 읽는다.
function classifyRecordDraftError_(message) {
  const raw = String(message || '');
  const lower = raw.toLowerCase();
  const codeMatch = raw.match(/API 오류 (\d{3})\b/);
  const code = codeMatch ? Number(codeMatch[1]) : 0;
  const pick = kind => Object.assign({}, RECORD_DRAFT_ERROR_KINDS[kind]);
  const isBilling = () => /insufficient_quota|billing|credit balance|payment|결제/.test(lower);

  if (code) {
    if (code === 401) return pick('auth');
    if (code === 402) return pick('billing');
    if (code === 403) return pick('permission');
    if (code === 429) return pick(isBilling() ? 'billing' : 'rateLimit');
    if (code === 400) {
      if (/api key not valid|api_key_invalid|invalid api key|invalid x-api-key/.test(lower)) return pick('auth');
      if (isBilling()) return pick('billing');
      return pick('request');
    }
    if (code === 408 || code === 409 || code >= 500) return pick('temporary');
    if (code >= 400) return pick('request');
    return pick('other');
  }

  // HTTP 응답 없이 난 오류: API 키 없음, Apps Script 호출 한도, 네트워크 시간 초과 등.
  if (lower.includes('api 키')) return pick('auth');
  if (/too many times|quota|할당량/.test(lower)) return pick('rateLimit');
  if (/timeout|timed out|address unavailable|dns error/.test(lower)) return pick('temporary');

  return pick('other');
}

function buildRecordDraftResultMessage_(title, result, unit) {
  unit = unit || '개';

  const lines = [
    title,
    '',
    `생성 완료: ${result.processed || 0}${unit}`,
    `실패: ${result.failed || 0}${unit}`,
    `스킵: ${result.skipped || 0}${unit}`,
    `남은 대기/재시도: ${result.remaining || 0}${unit}`,
  ];

  if (result.blocked) {
    lines.push('', `중단 사유: ${result.blockingReason || '확인 필요'}`);
    lines.push(result.blockingMessage || '더 이상 자동 생성을 계속할 수 없어 중단했습니다.');
  }

  return lines.join('\n');
}

function saveAutoRecordDraftLastResult_(status, result) {
  const payload = {
    status,
    processed: result.processed || 0,
    failed: result.failed || 0,
    skipped: result.skipped || 0,
    remaining: result.remaining || 0,
    blocked: !!result.blocked,
    blockingReason: result.blockingReason || '',
    blockingMessage: result.blockingMessage || '',
    updatedAt: new Date().toISOString(),
  };

  PropertiesService.getUserProperties().setProperty(
    'AUTO_RECORD_DRAFT_LAST_RESULT',
    JSON.stringify(payload)
  );
}

function getAutoRecordDraftLastResult_() {
  const raw = PropertiesService.getUserProperties().getProperty('AUTO_RECORD_DRAFT_LAST_RESULT');
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function toastAutoRecordDraft_(message, title) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title || '생기부 자동 생성', 10);
  } catch (err) {
    // 시간 기반 트리거에서 toast가 실패해도 자동 처리 자체에는 영향이 없다.
  }
}

function processPendingRecordDrafts() {
  const result = runFirstBatchQuickly_(() => processPendingRecordDraftsCore_(true));
  saveAutoRecordDraftLastResult_(result.blocked ? 'BLOCKED' : 'MANUAL_RUN', result);

  getUi_().alert(buildRecordDraftResultMessage_('이번 실행 결과', result));
}

// 실행이 6분 제한으로 강제 종료되면 끝에서 만드는 다음 예약이 사라져 자동 작업이 그대로 멈춘다.
// 그래서 시작할 때 6분보다 긴 시간 뒤의 예비 예약을 먼저 만들어 둔다. 정상으로 끝나면 이어서 할 때는 다음 예약으로 바뀌고,
// 끝났거나 멈췄을 때는 endAutoRun_이 지운다. 강제 종료됐을 때만 이 예약이 남아 이어서 만든다.
const AUTO_BACKUP_DELAY_MS = 7 * 60 * 1000;

// 자동 작업은 모두 스크립트 잠금을 잡고 돈다. 잠금을 잡은 지금 '생성중'인 행은 앞 실행이 강제 종료되며 남긴 것이라 '대기'로 돌린다.
function beginAutoRun_(scheduleNext, staleStatus) {
  scheduleNext(AUTO_BACKUP_DELAY_MS);

  if (staleStatus) {
    const changed = resetStatusValues_(staleStatus[0], staleStatus[1], ['생성중'], '대기');
    if (changed) log_('beginAutoRun_', 'RECOVER', `${staleStatus[0]}: 앞 실행이 끊기며 남긴 생성중 ${changed}행을 대기로 돌림`);
  }
}

function endAutoRun_(statusProperty, deleteTriggers) {
  if ((PropertiesService.getUserProperties().getProperty(statusProperty) || 'OFF') !== 'ON') deleteTriggers();
}

// 자동 작업들은 스크립트 잠금 하나를 같이 쓴다. 다른 작업이 돌고 있어 잠금을 못 잡으면
// 그냥 끝내지 말고 잠시 뒤로 다시 예약한다. 그러지 않으면 이 작업은 ON인 채로 멈춘다.
function rescheduleAutoJobWhenBusy_(handlerName, statusProperty, scheduleNext) {
  const status = PropertiesService.getUserProperties().getProperty(statusProperty) || 'OFF';

  if (status !== 'ON') {
    log_(handlerName, 'SKIP', '다른 자동 작업이 실행 중이고, 이 작업은 이미 꺼져 있어 다시 예약하지 않음');
    return;
  }

  scheduleNext(AUTO_BUSY_RETRY_DELAY_MS);
  log_(handlerName, 'WAIT', `다른 자동 작업이 실행 중이라 ${Math.round(AUTO_BUSY_RETRY_DELAY_MS / 1000)}초 뒤 다시 시도`);
}

function autoProcessPendingRecordDrafts() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    rescheduleAutoJobWhenBusy_('autoProcessPendingRecordDrafts', 'AUTO_RECORD_DRAFT_TRIGGER', scheduleNextAutoRecordDraftTrigger_);
    return;
  }

  try {
    deleteAutoRecordDraftTriggers_();

    const status = PropertiesService.getUserProperties().getProperty('AUTO_RECORD_DRAFT_TRIGGER') || 'OFF';

    if (status !== 'ON') {
      log_('autoProcessPendingRecordDrafts', 'STOP', '자동 생성 상태가 OFF라 실행하지 않음');
      return;
    }

    beginAutoRun_(scheduleNextAutoRecordDraftTrigger_, [SHEET_NAMES.submissions, 'aiStatus']);

    const result = processPendingRecordDraftsCore_(false);

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'OFF');
      saveAutoRecordDraftLastResult_('BLOCKED', result);
      toastAutoRecordDraft_(result.blockingMessage || '자동 생성을 중단했습니다.', result.blockingReason || '자동 생성 중단');
      log_(
        'autoProcessPendingRecordDrafts',
        'BLOCKED',
        `${result.blockingReason}: ${result.blockingMessage}`
      );
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'OFF');
      saveAutoRecordDraftLastResult_('DONE', result);
      toastAutoRecordDraft_(
        `생성 완료 ${result.processed}개, 실패 ${result.failed}개, 스킵 ${result.skipped}개`,
        '자동 생성 완료'
      );
      log_('autoProcessPendingRecordDrafts', 'DONE', '남은 대기열이 없어 자동 생성을 종료했습니다.');
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoRecordDraftTrigger_(nextDelayMs);
    saveAutoRecordDraftLastResult_('RUNNING', result);

    log_(
      'autoProcessPendingRecordDrafts',
      'NEXT',
      `남은 대기 ${result.remaining}개. ${Math.round(nextDelayMs / 1000)}초 뒤 다음 실행 예약`
    );

  } catch (err) {
    const message = String(err.message || err);
    const result = {
      processed: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingRecordDrafts_(),
      blocked: true,
      blockingReason: '자동 처리 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'OFF');
    saveAutoRecordDraftLastResult_('BLOCKED', result);
    toastAutoRecordDraft_(message, '자동 처리 오류');
    log_('autoProcessPendingRecordDrafts', 'ERROR', message);

  } finally {
    endAutoRun_('AUTO_RECORD_DRAFT_TRIGGER', deleteAutoRecordDraftTriggers_);
    lock.releaseLock();
  }
}

function processPendingRecordDraftsCore_(showUi) {
  const sh = getSheet_(SHEET_NAMES.submissions);
  const h = headerMap_(sh);
  const config = getConfigMap_();

  const batchSize = RUN_ITEM_LIMIT_;
  const maxRunSeconds = Number(config.MAX_RUN_SECONDS || 270);
  const startedAt = Date.now();

  const ai = getAiProviderAndModel_(config);
  const provider = ai.provider;
  const model = ai.model;
  const reasoning = ai.reasoning;
  const maxInputChars = Number(config.MAX_INPUT_CHARS || 30000);

  const byteRange = getRecordByteRange_(config, 'draft');
  const recordDraftMinChars = approxCharsFromBytes_(byteRange.min);
  const recordDraftMaxChars = approxCharsFromBytes_(byteRange.max);

  const systemGuide = getSystemGuide_(config);
  const recordPromptTemplate = getConfigPromptTemplate_(config, 'PROMPT_1', defaultRecordDraftPrompt_());

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  let processed = 0;
  let skipped = 0;
  let failed = 0;
  let blocked = false;
  let blockingReason = '';
  let blockingMessage = '';
  const savedRows = [];

  if (lastRow < 2) {
    return {
      processed,
      skipped,
      failed,
      remaining: 0,
      blocked,
      blockingReason,
      blockingMessage,
    };
  }

  const submissionValues = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const isDraftRowReady = values => {
    const status = String(rowValueAt_(values, h.aiStatus) || '').trim();
    const text = rowValueAt_(values, h.extractedText);
    return (status === '대기' || status === '재시도필요') &&
      !!rowValueAt_(values, h.submissionKey) && !!text && String(text).length >= 20;
  };

  const buildDraftUserPrompt = values => {
    const promptSanitizeOptions = {
      studentName: rowValueAt_(values, h.studentName),
      studentNo: rowValueAt_(values, h.studentNo),
      email: rowValueAt_(values, h.email),
    };
    const assignmentDescription = h.assignmentDescription ? rowValueAt_(values, h.assignmentDescription) : '';

    return buildRecordDraftPrompt_({
      promptTemplate: recordPromptTemplate,
      courseName: rowValueAt_(values, h.courseName),
      assignmentTitle: rowValueAt_(values, h.assignmentTitle),
      state: rowValueAt_(values, h.state),
      late: rowValueAt_(values, h.late),
      gradeText: formatGradeContext_(
        h.draftGrade ? rowValueAt_(values, h.draftGrade) : '',
        h.assignedGrade ? rowValueAt_(values, h.assignedGrade) : '',
        h.maxPoints ? rowValueAt_(values, h.maxPoints) : '',
        h.teacherGrade ? rowValueAt_(values, h.teacherGrade) : ''
      ),
      assignmentDescription: sanitizePromptText_(
        assignmentDescription || '(과제 설명란에 별도 지시사항 없음)',
        promptSanitizeOptions
      ),
      recordDraftMinChars,
      recordDraftMaxChars,
      recordDraftMinBytes: byteRange.min,
      recordDraftMaxBytes: byteRange.max,
      extractedText: truncate_(sanitizePromptText_(rowValueAt_(values, h.extractedText), promptSanitizeOptions), maxInputChars),
    });
  };

  const prefetcher = createAiPrefetcher_(config, 'processPendingRecordDraftsCore_');
  const upcomingDraftRequests = fromRow => {
    const requests = [];
    const limit = Math.min(FAST_GENERATION_CONCURRENCY, batchSize - processed);

    for (let r = fromRow; r <= lastRow && requests.length < limit; r++) {
      const values = submissionValues[r - 2] || [];
      if (!isDraftRowReady(values)) continue;
      requests.push({ provider, model, systemGuide, userPrompt: buildDraftUserPrompt(values), reasoning });
    }

    return requests;
  };

  for (let row = 2; row <= lastRow; row++) {
    const elapsedSeconds = (Date.now() - startedAt) / 1000;

    if (elapsedSeconds > maxRunSeconds) {
      log_(
        'processPendingRecordDraftsCore_',
        'STOP',
        `실행 시간 제한 접근으로 중단. 처리 ${processed}개, 실패 ${failed}개`
      );
      break;
    }

    if (processed >= batchSize) {
      break;
    }

    const rowValues = submissionValues[row - 2] || [];
    const aiStatus = String(rowValueAt_(rowValues, h.aiStatus) || '').trim();

    if (aiStatus !== '대기' && aiStatus !== '재시도필요') {
      continue;
    }

    const submissionKey = rowValueAt_(rowValues, h.submissionKey);
    const studentNo = rowValueAt_(rowValues, h.studentNo);
    const studentName = rowValueAt_(rowValues, h.studentName);
    const email = rowValueAt_(rowValues, h.email);
    const courseName = rowValueAt_(rowValues, h.courseName);
    const assignmentTitle = rowValueAt_(rowValues, h.assignmentTitle);
    const assignmentDescription = h.assignmentDescription
      ? rowValueAt_(rowValues, h.assignmentDescription)
      : '';
    const state = rowValueAt_(rowValues, h.state);
    const late = rowValueAt_(rowValues, h.late);
    const draftGrade = h.draftGrade ? rowValueAt_(rowValues, h.draftGrade) : '';
    const assignedGrade = h.assignedGrade ? rowValueAt_(rowValues, h.assignedGrade) : '';
    const maxPoints = h.maxPoints ? rowValueAt_(rowValues, h.maxPoints) : '';
    const teacherGrade = h.teacherGrade ? rowValueAt_(rowValues, h.teacherGrade) : '';
    const extractedText = rowValueAt_(rowValues, h.extractedText);

    if (!submissionKey || !extractedText || String(extractedText).length < 20) {
      sh.getRange(row, h.aiStatus).setValue('오류');
      log_('processPendingRecordDraftsCore_', 'SKIP', `${studentName}: 추출 텍스트 부족`);
      skipped++;
      continue;
    }

    const gradeText = formatGradeContext_(draftGrade, assignedGrade, maxPoints, teacherGrade);
    const userPrompt = buildDraftUserPrompt(rowValues);

    prefetcher.prefetch(() => upcomingDraftRequests(row));

    try {
      sh.getRange(row, h.aiStatus).setValue('생성중');
      const generated = createRecordDraftWithLength_(
        provider,
        model,
        systemGuide,
        userPrompt,
        byteRange.min,
        byteRange.max,
        reasoning
      );


      const parsed = generated.parsed;
      const recordDraft = normalizeAiText_(parsed.record_draft);
      const caution = normalizeAiText_(parsed.caution);
      const evidenceSummary = normalizeAiText_(parsed.evidence_summary);

      const outputRow = [
        submissionKey,
        studentNo,
        studentName,
        email,
        courseName,
        assignmentTitle,
        recordDraft,
        neisByteCount_(recordDraft),
        caution,
        '',
        new Date(),
        `${provider}:${model}`,
        evidenceSummary,
        recordDraft,
        gradeText,
        JSON.stringify(generated.usage || {}),
      ];
      const saveResult = appendRecordsOnly_([outputRow]);

      savedRows.push(...saveResult.targetRows);

      sh.getRange(row, h.aiStatus).setValue(recordDraft ? '생성완료' : '근거부족');

      processed++;

      Utilities.sleep(1200);

    } catch (err) {
      if (isAiTimeBudgetError_(err)) {
        sh.getRange(row, h.aiStatus).setValue('대기');
        log_('processPendingRecordDraftsCore_', 'STOP', `${studentName}: ${err.message}`);
        break;
      }

      const message = String(err.message || '');
      const errorInfo = classifyRecordDraftError_(message);

      sh.getRange(row, h.aiStatus).setValue(errorInfo.status);

      if (errorInfo.shouldStop) {
        blocked = true;
        blockingReason = errorInfo.reason;
        blockingMessage = errorInfo.userMessage;
      }

      failed++;
      log_(
        'processPendingRecordDraftsCore_',
        blocked ? 'STOP' : 'ERROR',
        `${studentName}: ${errorInfo.reason} / ${message}`
      );

      if (blocked) break;
    }
  }

  prefetcher.finish();

  if (savedRows.length > 0) {
    applyNeisByteFormulasToRecords_(savedRows);
    formatRecordRows_(savedRows);
  }

  if (savedRows.length > 0 && showUi) {
    showRecordsSheetAfterSave_(savedRows);
  }

  refreshDashboard();




  const remaining = countPendingRecordDrafts_();

  log_(
    'processPendingRecordDraftsCore_',
    'OK',
    `처리 ${processed}개, 실패 ${failed}개, 스킵 ${skipped}개, 남은 대기 ${remaining}개`
  );

  return {
    processed,
    skipped,
    failed,
    remaining,
    blocked,
    blockingReason,
    blockingMessage,
  };
}

function countPendingRecordDrafts_() {
  const sh = getSheet_(SHEET_NAMES.submissions);
  const h = headerMap_(sh);

  if (!h.aiStatus) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const values = sh.getRange(2, h.aiStatus, lastRow - 1, 1).getValues();

  return values.filter(r => {
    const status = String(r[0] || '').trim();
    return status === '대기' || status === '재시도필요';
  }).length;
}

function startAutoRecordDraftTrigger() {

  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/spreadsheets.currentonly',
    'https://www.googleapis.com/auth/script.external_request',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/classroom.courses.readonly',
    'https://www.googleapis.com/auth/classroom.coursework.students.readonly',
    'https://www.googleapis.com/auth/classroom.rosters.readonly',
    'https://www.googleapis.com/auth/classroom.profile.emails'
  ]);  
  const ui = getUi_();


  try {
    requireCurrentAiApiKey_();
  } catch (err) {
    ui.alert(String(err.message || err));
    return;
  }


  const pendingCount = countPendingRecordDrafts_();

  if (pendingCount === 0) {
    ui.alert('대기 또는 재시도필요 상태인 제출물이 없습니다.\n먼저 제출물 시트에서 학생을 선택한 뒤 대기열에 넣으세요.');
    return;
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    ui.alert('이미 다른 자동 생성 작업이 실행 중입니다.\n현재 작업이 끝난 뒤 다시 시작해 주세요.');
    return;
  }

  try {
    deleteAutoRecordDraftTriggers_();

    const props = PropertiesService.getUserProperties();
    props.setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'ON');

    const result = runFirstBatchQuickly_(() => processPendingRecordDraftsCore_(false));

    if (result.blocked) {
      props.setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'OFF');
      deleteAutoRecordDraftTriggers_();
      saveAutoRecordDraftLastResult_('BLOCKED', result);

      log_(
        'startAutoRecordDraftTrigger',
        'BLOCKED',
        `${result.blockingReason}: ${result.blockingMessage}`
      );

      ui.alert(buildRecordDraftResultMessage_('자동 생성을 중단했습니다.', result));
      return;
    }

    if (result.remaining === 0) {
      props.setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'OFF');
      saveAutoRecordDraftLastResult_('DONE', result);

      log_(
        'startAutoRecordDraftTrigger',
        'DONE',
        `첫 배치 즉시 처리 후 남은 대기열이 없어 자동 생성을 종료했습니다. 처리 ${result.processed}개, 실패 ${result.failed}개, 스킵 ${result.skipped}개`
      );

      ui.alert(buildRecordDraftResultMessage_('자동 생성을 완료했습니다.', result));
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoRecordDraftTrigger_(nextDelayMs);
    saveAutoRecordDraftLastResult_('RUNNING', result);

    log_(
      'startAutoRecordDraftTrigger',
      'OK',
      `첫 배치 즉시 처리 완료. 처리 ${result.processed}개, 실패 ${result.failed}개, 스킵 ${result.skipped}개, 남은 대기열 ${result.remaining}개. ${Math.round(nextDelayMs / 1000)}초 뒤 다음 실행 예약`
    );

  } catch (err) {
    const message = String(err.message || err);
    const remaining = countPendingRecordDrafts_();
    const result = {
      processed: 0,
      failed: 1,
      skipped: 0,
      remaining,
      blocked: true,
      blockingReason: '자동 생성 시작 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'OFF');
    deleteAutoRecordDraftTriggers_();
    saveAutoRecordDraftLastResult_('BLOCKED', result);
    log_('startAutoRecordDraftTrigger', 'ERROR', message);

    ui.alert(buildRecordDraftResultMessage_('자동 생성 시작 중 오류가 발생했습니다.', result));
  } finally {
    lock.releaseLock();
  }
}

function stopAutoRecordDraftTrigger(showUi) {
  deleteAutoRecordDraftTriggers_();
  deleteAutoStudentFinalRecordTriggers_();
  deleteAutoCommonPhraseTriggers_();
  deleteAutoManualRecordTriggers_();
  deleteAutoSubmissionCollectionTriggers_();
  deleteAutoBehaviorRecommendationTriggers_();

  PropertiesService.getUserProperties().setProperty('AUTO_RECORD_DRAFT_TRIGGER', 'OFF');
  PropertiesService.getUserProperties().setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'OFF');
  PropertiesService.getUserProperties().setProperty('AUTO_COMMON_PHRASE_TRIGGER', 'OFF');
  PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');
  PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'OFF');
  PropertiesService.getUserProperties().setProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY, 'OFF');

  log_('stopAutoRecordDraftTrigger', 'OK', '자동 생성 트리거 중지');

  if (showUi !== false) {
    getUi_().alert('제출물 수집/생기부초안/학생별 최종 생기부/공통문구/수동추가/행발 추천 자동 트리거를 중지했습니다.');
  }
}


function showAutoRecordDraftStatus() {
  const submissionCollectionPendingCount = countPendingSubmissionCollections_();
  const draftPendingCount = countPendingRecordDrafts_();
  const studentFinalPendingCount = countPendingStudentFinalRecords_();
  const commonPhrasePendingCount = countPendingCommonPhrases_();
  const behaviorRecommendationPendingCount = countPendingBehaviorRecommendationRows_();

  const triggers = ScriptApp.getProjectTriggers();
  const submissionCollectionTriggers = triggers.filter(trigger => {
    return trigger.getHandlerFunction() === 'autoProcessPendingSubmissionCollections';
  });
  const draftTriggers = triggers.filter(trigger => {
    return trigger.getHandlerFunction() === 'autoProcessPendingRecordDrafts';
  });
  const studentFinalTriggers = triggers.filter(trigger => {
    return trigger.getHandlerFunction() === 'autoProcessPendingStudentFinalRecords';
  });
  const commonPhraseTriggers = triggers.filter(trigger => {
    return trigger.getHandlerFunction() === 'autoProcessPendingCommonPhrases';
  });
  const behaviorRecommendationTriggers = triggers.filter(trigger => {
    return trigger.getHandlerFunction() === 'autoProcessPendingBehaviorRecommendations';
  });

  const submissionCollectionStatus = PropertiesService.getUserProperties().getProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY) || 'OFF';
  const draftStatus = PropertiesService.getUserProperties().getProperty('AUTO_RECORD_DRAFT_TRIGGER') || 'OFF';
  const studentFinalStatus = PropertiesService.getUserProperties().getProperty('AUTO_STUDENT_FINAL_TRIGGER') || 'OFF';
  const commonPhraseStatus = PropertiesService.getUserProperties().getProperty('AUTO_COMMON_PHRASE_TRIGGER') || 'OFF';
  const behaviorRecommendationStatus = PropertiesService.getUserProperties().getProperty(AUTO_BEHAVIOR_RECOMMENDATION_TRIGGER_PROPERTY) || 'OFF';
  const submissionCollectionLastResult = formatAutoLastResultForStatus_('마지막 제출물 자동 수집 결과', getAutoSubmissionCollectionLastResult_(), '개', '수집 완료');
  const draftLastResult = formatAutoLastResultForStatus_('마지막 생기부초안 자동 생성 결과', getAutoRecordDraftLastResult_(), '개');
  const studentFinalLastResult = formatAutoLastResultForStatus_('마지막 학생별 최종 자동 생성 결과', getAutoStudentFinalLastResult_(), '명');
  const commonPhraseLastResult = formatAutoLastResultForStatus_('마지막 공통문구 자동 생성 결과', getAutoCommonPhraseLastResult_(), '개');

  getUi_().alert(
    `자동 생성 상태\n\n` +
    `[제출물 수집]\n` +
    `상태값: ${submissionCollectionStatus}\n` +
    `실제 트리거 수: ${submissionCollectionTriggers.length}개\n` +
    `남은 수집대기 과제: ${submissionCollectionPendingCount}개\n` +
    submissionCollectionLastResult +
    `\n\n` +
    `[생기부초안]\n` +
    `상태값: ${draftStatus}\n` +
    `실제 트리거 수: ${draftTriggers.length}개\n` +
    `남은 대기/재시도 제출물: ${draftPendingCount}개\n` +
    draftLastResult +
    `\n\n[학생별 최종 생기부]\n` +
    `상태값: ${studentFinalStatus}\n` +
    `실제 트리거 수: ${studentFinalTriggers.length}개\n` +
    `남은 대기/재시도 학생: ${studentFinalPendingCount}명` +
    studentFinalLastResult +
    `\n\n[공통문구]\n` +
    `상태값: ${commonPhraseStatus}\n` +
    `실제 트리거 수: ${commonPhraseTriggers.length}개\n` +
    `남은 대기/재시도 문구: ${commonPhrasePendingCount}개` +
    commonPhraseLastResult +
    `\n\n[행발문구추천]\n` +
    `상태값: ${behaviorRecommendationStatus}\n` +
    `실제 트리거 수: ${behaviorRecommendationTriggers.length}개\n` +
    `남은 대기/재시도 행: ${behaviorRecommendationPendingCount}행`
  );
}

function formatAutoLastResultForStatus_(title, lastResult, unit, processedLabel) {
  if (!lastResult) return '';
  processedLabel = processedLabel || '생성 완료';

  return [
    '',
    `[${title}]`,
    `상태: ${lastResult.status || ''}`,
    `시각: ${lastResult.updatedAt || ''}`,
    `${processedLabel}: ${lastResult.processed || 0}${unit}`,
    `실패: ${lastResult.failed || 0}${unit}`,
    `스킵: ${lastResult.skipped || 0}${unit}`,
    `남은 대기/재시도: ${lastResult.remaining || 0}${unit}`,
    lastResult.blocked ? `중단 사유: ${lastResult.blockingReason || '확인 필요'}` : '',
    lastResult.blocked ? String(lastResult.blockingMessage || '') : '',
  ].filter(Boolean).join('\n');
}

function scheduleNextAutoRecordDraftTrigger_(delayMs) {
  deleteAutoRecordDraftTriggers_();

  ScriptApp.newTrigger('autoProcessPendingRecordDrafts')
    .timeBased()
    .after(delayMs)
    .create();

  log_(
    'scheduleNextAutoRecordDraftTrigger_',
    'OK',
    `${Math.round(delayMs / 1000)}초 뒤 다음 자동 생성 예약`
  );
}

function deleteAutoRecordDraftTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoProcessPendingRecordDrafts') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function generateManualAddedRecords() {
  const ui = getUi_();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    ui.alert('이미 다른 자동 생성 작업이 실행 중입니다.\n현재 작업이 끝난 뒤 다시 시작해 주세요.');
    return;
  }

  try {
    deleteAutoManualRecordTriggers_();
    PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'ON');

    const result = runFirstBatchQuickly_(() => processManualAddedRecordsCore_());

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');
      deleteAutoManualRecordTriggers_();
      ui.alert(buildRecordDraftResultMessage_('수동추가 생기부 생성을 중단했습니다.', result, '행'));
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');
      deleteAutoManualRecordTriggers_();
      ui.alert(buildRecordDraftResultMessage_('수동추가 생기부 생성을 완료했습니다.', result, '행'));
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);
    scheduleNextAutoManualRecordTrigger_(nextDelayMs);

    ui.alert(
      buildRecordDraftResultMessage_('수동추가 첫 배치 처리 결과', result, '행') +
      `\n\n남은 행은 ${Math.round(nextDelayMs / 1000)}초 뒤 자동으로 이어서 처리합니다.`
    );
  } catch (err) {
    const message = String(err.message || err);
    const result = {
      processed: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingManualAddedRows_(),
      blocked: true,
      blockingReason: '수동추가 처리 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');
    deleteAutoManualRecordTriggers_();
    log_('generateManualAddedRecords', 'ERROR', message);
    ui.alert(buildRecordDraftResultMessage_('수동추가 생기부 생성 중 오류가 발생했습니다.', result, '행'));
  } finally {
    lock.releaseLock();
  }
}

function autoProcessPendingManualAddedRecords() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    rescheduleAutoJobWhenBusy_('autoProcessPendingManualAddedRecords', AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, scheduleNextAutoManualRecordTrigger_);
    return;
  }

  try {
    deleteAutoManualRecordTriggers_();

    const status = PropertiesService.getUserProperties().getProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY) || 'OFF';

    if (status !== 'ON') {
      log_('autoProcessPendingManualAddedRecords', 'STOP', '자동 생성 상태가 OFF라 실행하지 않음');
      return;
    }

    beginAutoRun_(scheduleNextAutoManualRecordTrigger_);

    const result = processManualAddedRecordsCore_();

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');
      toastAutoRecordDraft_(result.blockingMessage || '수동추가 생성을 중단했습니다.', result.blockingReason || '수동추가 생성 중단');
      log_('autoProcessPendingManualAddedRecords', 'BLOCKED', `${result.blockingReason}: ${result.blockingMessage}`);
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');
      toastAutoRecordDraft_(
        `생성 완료 ${result.processed}행, 실패 ${result.failed}행, 스킵 ${result.skipped}행`,
        '수동추가 생성 완료'
      );
      log_('autoProcessPendingManualAddedRecords', 'DONE', '남은 수동추가 행이 없어 종료했습니다.');
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoManualRecordTrigger_(nextDelayMs);
    log_(
      'autoProcessPendingManualAddedRecords',
      'NEXT',
      `남은 수동추가 ${result.remaining}행. ${Math.round(nextDelayMs / 1000)}초 뒤 다음 실행 예약`
    );
  } catch (err) {
    const message = String(err.message || err);
    PropertiesService.getUserProperties().setProperty(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, 'OFF');
    toastAutoRecordDraft_(message, '수동추가 자동 처리 오류');
    log_('autoProcessPendingManualAddedRecords', 'ERROR', message);
  } finally {
    endAutoRun_(AUTO_MANUAL_RECORD_TRIGGER_PROPERTY, deleteAutoManualRecordTriggers_);
    lock.releaseLock();
  }
}

function processManualAddedRecordsCore_() {
  const config = getConfigMap_();
  const manualConfig = getManualRecordConfig_(config);
  const sh = getSheet_(MANUAL_SHEET_NAME);
  ensureManualOutputColumns_(sh, manualConfig);

  const batchSize = RUN_ITEM_LIMIT_;
  const maxRunSeconds = Number(config.MAX_RUN_SECONDS || 270);
  const startedAt = Date.now();
  const ai = getAiProviderAndModel_(config);
  const provider = ai.provider;
  const model = ai.model;
  const reasoning = ai.reasoning;
  const maxInputChars = Number(config.MAX_INPUT_CHARS || 30000);
  const byteRange = getRecordByteRange_(config, 'manual');
  const recordDraftMinChars = approxCharsFromBytes_(byteRange.min);
  const recordDraftMaxChars = approxCharsFromBytes_(byteRange.max);
  const systemGuide = getSystemGuide_(config);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), manualConfig.byteCol);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowCount = Math.max(lastRow - manualConfig.startRow + 1, 0);
  const manualValues = rowCount > 0
    ? sh.getRange(manualConfig.startRow, 1, rowCount, lastCol).getValues()
    : [];
  const outputNotes = rowCount > 0
    ? sh.getRange(manualConfig.startRow, manualConfig.outputCol, rowCount, 1).getNotes()
    : [];

  let processed = 0;
  let skipped = 0;
  let failed = 0;
  let blocked = false;
  let blockingReason = '';
  let blockingMessage = '';

  if (lastRow < manualConfig.startRow) {
    return {
      processed,
      skipped,
      failed,
      remaining: 0,
      blocked,
      blockingReason,
      blockingMessage,
    };
  }

  const buildManualUserPrompt = (row, manualInput) => buildManualRecordPrompt_({
    promptTemplate: manualConfig.prompt,
    rowNumber: row,
    manualInput: truncate_(manualInput, maxInputChars),
    inputColumns: manualConfig.inputCols.map(colToA1_).join(', '),
    recordDraftMinChars,
    recordDraftMaxChars,
    recordDraftMinBytes: byteRange.min,
    recordDraftMaxBytes: byteRange.max,
  });

  const prefetcher = createAiPrefetcher_(config, 'processManualAddedRecordsCore_');
  const upcomingManualRequests = fromRow => {
    const requests = [];
    const limit = Math.min(FAST_GENERATION_CONCURRENCY, batchSize - processed);

    for (let r = fromRow; r <= lastRow && requests.length < limit; r++) {
      const index = r - manualConfig.startRow;
      const values = manualValues[index] || [];
      const note = String((outputNotes[index] && outputNotes[index][0]) || '').trim();
      if (String(rowValueAt_(values, manualConfig.outputCol) || '').trim() || isManualOutputHandledNote_(note)) continue;

      const manualInput = buildManualInputFromRowValues_(values, headers, manualConfig.inputCols);
      if (manualInput.length < 20) continue;

      requests.push({ provider, model, systemGuide, userPrompt: buildManualUserPrompt(r, manualInput), reasoning });
    }

    return requests;
  };

  for (let row = manualConfig.startRow; row <= lastRow; row++) {
    const elapsedSeconds = (Date.now() - startedAt) / 1000;

    if (elapsedSeconds > maxRunSeconds) {
      log_(
        'processManualAddedRecordsCore_',
        'STOP',
        `실행 시간 제한 접근으로 중단. 처리 ${processed}행, 실패 ${failed}행`
      );
      break;
    }

    if (processed >= batchSize) break;

    const rowIndex = row - manualConfig.startRow;
    const rowValues = manualValues[rowIndex] || [];
    const outputValue = String(rowValueAt_(rowValues, manualConfig.outputCol) || '').trim();
    const outputNote = String((outputNotes[rowIndex] && outputNotes[rowIndex][0]) || '').trim();

    if (outputValue || isManualOutputHandledNote_(outputNote)) {
      continue;
    }

    const manualInput = buildManualInputFromRowValues_(rowValues, headers, manualConfig.inputCols);

    if (manualInput.length < 20) {
      skipped++;
      continue;
    }

    const userPrompt = buildManualUserPrompt(row, manualInput);

    prefetcher.prefetch(() => upcomingManualRequests(row));

    const outputCell = sh.getRange(row, manualConfig.outputCol);

    try {
      const generated = createRecordDraftWithLength_(
        provider,
        model,
        systemGuide,
        userPrompt,
        byteRange.min,
        byteRange.max,
        reasoning
      );

      const parsed = generated.parsed;
      const recordDraft = normalizeAiText_(parsed.record_draft);
      const caution = normalizeAiText_(parsed.caution);
      const evidenceSummary = normalizeAiText_(parsed.evidence_summary);

      outputCell
        .setValue(recordDraft)
        .setNote(buildManualOutputNote_(recordDraft ? '처리완료' : '근거부족', caution, evidenceSummary, `${provider}:${model}`));
      setManualByteFormula_(sh, row, manualConfig.outputCol);
      formatManualProcessedRow_(sh, row, manualConfig);

      processed++;
      Utilities.sleep(1200);
    } catch (err) {
      if (isAiTimeBudgetError_(err)) {
        log_('processManualAddedRecordsCore_', 'STOP', `${row}행: ${err.message}`);
        break;
      }

      const message = String(err.message || '');
      const errorInfo = classifyRecordDraftError_(message);

      failed++;
      outputCell.setNote(buildManualOutputErrorNote_(errorInfo.reason, message));
      formatManualProcessedRow_(sh, row, manualConfig);
      log_(
        'processManualAddedRecordsCore_',
        errorInfo.shouldStop ? 'STOP' : 'ERROR',
        `${row}행: ${errorInfo.reason} / ${message}`
      );

      if (errorInfo.shouldStop) {
        blocked = true;
        blockingReason = errorInfo.reason;
        blockingMessage = errorInfo.userMessage;
        break;
      }
    }
  }

  prefetcher.finish();

  const remaining = countPendingManualAddedRows_();

  log_(
    'processManualAddedRecordsCore_',
    'OK',
    `처리 ${processed}행, 실패 ${failed}행, 스킵 ${skipped}행, 남은 대기 ${remaining}행`
  );

  return {
    processed,
    skipped,
    failed,
    remaining,
    blocked,
    blockingReason,
    blockingMessage,
  };
}

function countPendingManualAddedRows_() {
  try {
    const config = getConfigMap_();
    const manualConfig = getManualRecordConfig_(config);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANUAL_SHEET_NAME);

    if (!sh) return 0;

    return countPendingManualRowsInSheet_(sh, manualConfig);
  } catch (err) {
    return 0;
  }
}

function countPendingManualRowsInSheet_(sh, manualConfig) {
  const lastRow = sh.getLastRow();
  if (lastRow < manualConfig.startRow) return 0;

  const rowCount = lastRow - manualConfig.startRow + 1;
  const lastCol = Math.max(sh.getLastColumn(), manualConfig.outputCol);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sh.getRange(manualConfig.startRow, 1, rowCount, lastCol).getValues();
  const outputNotes = sh.getRange(manualConfig.startRow, manualConfig.outputCol, rowCount, 1).getNotes();
  let count = 0;

  values.forEach((row, index) => {
    const outputValue = String(row[manualConfig.outputCol - 1] || '').trim();
    const outputNote = String(outputNotes[index][0] || '').trim();

    if (outputValue || isManualOutputHandledNote_(outputNote)) return;

    const input = manualConfig.inputCols
      .map(col => {
        const label = String(headers[col - 1] || '').trim();
        if (isPersonalInfoHeader_(label)) return '';
        return sanitizePromptText_(row[col - 1]).trim();
      })
      .filter(Boolean)
      .join('\n');

    if (input.length >= 20) count++;
  });

  return count;
}

function getManualRecordConfig_(config) {
  const inputCols = parseManualColumnList_(config.MANUAL_INPUT_COLS);
  const outputCol = parseManualColumn_(config.MANUAL_OUTPUT_COL, 'MANUAL_OUTPUT_COL');
  const byteCol = outputCol + 1;

  if (inputCols.length === 0) {
    throw new Error('설정 시트의 MANUAL_INPUT_COLS에 API로 보낼 열을 입력하세요. 예: E,F');
  }

  if (inputCols.includes(outputCol)) {
    throw new Error('MANUAL_OUTPUT_COL은 MANUAL_INPUT_COLS에 포함될 수 없습니다.');
  }

  if (inputCols.includes(byteCol)) {
    throw new Error('MANUAL_OUTPUT_COL 바로 오른쪽 바이트 계산 열은 MANUAL_INPUT_COLS에 포함될 수 없습니다.');
  }

  return {
    startRow: Math.max(Number(config.MANUAL_START_ROW || 2), 1),
    inputCols,
    outputCol,
    byteCol,
    prompt: getConfigPromptTemplate_(config, 'MANUAL_PROMPT', defaultManualRecordPrompt_()),
  };
}

function parseManualColumnList_(value) {
  return String(value || '')
    .split(',')
    .map(part => part.trim())
    .filter(Boolean)
    .map(part => parseManualColumn_(part, 'MANUAL_INPUT_COLS'));
}

function parseManualColumn_(value, key) {
  const text = String(value || '').trim().toUpperCase();

  if (!text) {
    throw new Error(`설정 시트의 ${key} 값을 입력하세요.`);
  }

  if (/^\d+$/.test(text)) {
    const col = Number(text);
    if (col > 0) return col;
  }

  if (/^[A-Z]+$/.test(text)) {
    return text.split('').reduce((col, ch) => col * 26 + ch.charCodeAt(0) - 64, 0);
  }

  throw new Error(`${key} 값은 열 문자 또는 숫자로 입력하세요: ${value}`);
}

function ensureManualOutputColumns_(sh, manualConfig) {
  const maxInputCol = Math.max.apply(null, manualConfig.inputCols);

  if (sh.getMaxColumns() < maxInputCol) {
    throw new Error(`수동추가 시트에 MANUAL_INPUT_COLS 열(${manualConfig.inputCols.map(colToA1_).join(', ')})이 없습니다.`);
  }

  if (sh.getMaxColumns() < manualConfig.byteCol) {
    sh.insertColumnsAfter(sh.getMaxColumns(), manualConfig.byteCol - sh.getMaxColumns());
  }
}

function buildManualInputForRow_(sh, row, headers, inputCols) {
  const lastCol = Math.max(sh.getLastColumn(), Math.max.apply(null, inputCols));
  const rowValues = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  return buildManualInputFromRowValues_(rowValues, headers, inputCols);
}

function buildManualInputFromRowValues_(rowValues, headers, inputCols) {
  return inputCols.map(col => {
    const label = String(headers[col - 1] || '').trim() || `열 ${colToA1_(col)}`;
    if (isPersonalInfoHeader_(label)) return '';

    const value = sanitizePromptText_(rowValueAt_(rowValues, col)).trim();

    if (!value) return '';

    return `[${label}]\n${value}`;
  }).filter(Boolean).join('\n\n');
}

function buildManualRecordPrompt_(data) {
  const template = data.promptTemplate || defaultManualRecordPrompt_();
  const renderedPrompt = renderPromptTemplate_(template, data);
  const parts = [renderedPrompt];
  const missingInputLines = buildMissingPromptLines_(template, [
    ['rowNumber', `수동추가 시트 행: ${data.rowNumber}`],
    ['inputColumns', `입력 열: ${data.inputColumns}`],
  ]);

  if (missingInputLines.length > 0) {
    parts.push(['[입력 정보]'].concat(missingInputLines).join('\n'));
  }

  parts.push(buildByteRangeInstruction_('record_draft', {
    min: data.recordDraftMinBytes,
    max: data.recordDraftMaxBytes,
  }));
  parts.push([
    '[기본 출력 조건]',
    '생활기록부에 실제로 옮겨 적을 문장은 record_draft에만 작성한다.',
    '반드시 JSON 형식으로만 답한다.',
  ].join('\n'));

  if (!promptTemplateHasPlaceholder_(template, 'manualInput')) {
    parts.push([
      '[수동추가 데이터]',
      data.manualInput,
    ].join('\n'));
  }

  if (!/JSON\s*형식으로만\s*답/i.test(template)) {
    parts.push([
      '반드시 JSON 형식으로만 답한다.',
      '{"evidence_summary":"", "record_draft":"", "caution":""}',
    ].join('\n'));
  }

  parts.push(buildRecordDraftNoEvidenceRule_());

  return parts.filter(part => String(part || '').trim()).join('\n\n');
}

function buildManualOutputNote_(status, caution, evidenceSummary, model) {
  return [
    `[수동추가 ${status}]`,
    model ? `model: ${model}` : '',
    caution ? `[caution]\n${caution}` : '',
    evidenceSummary ? `[evidence_summary]\n${evidenceSummary}` : '',
  ].filter(Boolean).join('\n\n');
}

function buildManualOutputErrorNote_(reason, message) {
  return [
    '[수동추가 오류]',
    `${reason}: ${message}`,
  ].filter(Boolean).join('\n');
}

function isManualOutputHandledNote_(note) {
  return /^\[수동추가 (처리완료|근거부족|오류)\]/.test(String(note || '').trim());
}

function setManualByteFormula_(sh, row, outputCol) {
  const outputColA1 = colToA1_(outputCol);

  sh.getRange(row, outputCol + 1)
    .setFormula(`=IF($${outputColA1}${row}="","",2*LENB($${outputColA1}${row})-LEN($${outputColA1}${row}))`);
}

function formatManualProcessedRow_(sh, row, manualConfig) {
  const lastCol = Math.max(sh.getLastColumn(), manualConfig.byteCol);

  sh.getRange(row, 1, 1, lastCol)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('top');
}

function scheduleNextAutoManualRecordTrigger_(delayMs) {
  deleteAutoManualRecordTriggers_();

  ScriptApp.newTrigger('autoProcessPendingManualAddedRecords')
    .timeBased()
    .after(delayMs)
    .create();

  log_(
    'scheduleNextAutoManualRecordTrigger_',
    'OK',
    `${Math.round(delayMs / 1000)}초 뒤 다음 수동추가 자동 생성 예약`
  );
}

function deleteAutoManualRecordTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoProcessPendingManualAddedRecords') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function getAiProviderAndModel_(config) {
  const provider = normalizeAiProvider_(config.AI_PROVIDER);
  const reasoning = normalizeReasoning_(config.REASONING || config.reasoning);

  let model = config.CLAUDE_MODEL || 'claude-haiku-5-5';

  if (provider === 'gpt') {
    model = config.GPT_MODEL || config.OPENAI_MODEL || 'gpt-6-luna';
  } else if (provider === 'gemini') {
    model = config.GEMINI_MODEL || 'gemini-3.8-flash';
  }

  return {
    provider,
    model,
    reasoning,
  };
}


// 저장된 API 키로 회사별 선택 가능한 모델 × 추론 수준을 모두 짧게 호출해 보고 'AI점검' 시트에 결과를 적는다.
// 실제 생성과 같은 callAi_ → parseClaudeJson_ 경로를 쓰므로, 여기서 통과하면 같은 설정으로 생성할 때도 요청 형식 오류가 나지 않는다.
function checkAllAiCombinations() {
  const props = PropertiesService.getUserProperties();
  const modelOptions = {
    claude: CONFIG_VALIDATION_OPTIONS.CLAUDE_MODEL,
    gpt: CONFIG_VALIDATION_OPTIONS.GPT_MODEL,
    gemini: CONFIG_VALIDATION_OPTIONS.GEMINI_MODEL,
  };
  const systemGuide = '너는 JSON으로만 답한다.';
  const userPrompt = [
    '고등학생이 과학 탐구 보고서를 성실히 제출했다는 내용으로 생활기록부 문장을 한 문장(50자 내외)으로 써라.',
    '반드시 {"evidence_summary": "...", "record_draft": "...", "caution": ""} 형식의 JSON으로만 답하라.',
  ].join('\n');
  const startedAt = Date.now();
  const rows = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('AI점검') || ss.insertSheet('AI점검');

  // 한 번에 다 못 끝내면(6분 제한) 다시 실행했을 때 이미 '정상'인 조합은 그대로 두고 나머지만 시험한다.
  const passed = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach(r => {
      if (r[3] === '정상') passed[`${r[0]}|${r[1]}|${r[2]}`] = r;
    });
  }

  Object.keys(modelOptions).forEach(provider => {
    const hasKey = !!props.getProperty(AI_API_ENDPOINTS[provider].keyProperty);

    modelOptions[provider].forEach(model => {
      CONFIG_VALIDATION_OPTIONS.REASONING.forEach(reasoning => {
        const previous = passed[`${provider}|${model}|${reasoning}`];
        if (previous) {
          rows.push(previous);
          return;
        }

        const row = [provider, model, reasoning, '', '', '', ''];
        rows.push(row);

        if (!hasKey) {
          row[3] = '건너뜀';
          row[6] = 'API 키가 저장되어 있지 않음';
          return;
        }

        if ((Date.now() - startedAt) / 1000 > 300) {
          row[3] = '건너뜀';
          row[6] = '실행 시간 제한에 가까워 이번에는 시험하지 않음. 메뉴를 한 번 더 실행하면 이어서 시험합니다.';
          return;
        }

        const callStartedAt = Date.now();

        try {
          const result = callAi_(provider, model, systemGuide, userPrompt, reasoning);
          const parsed = parseClaudeJson_(result.text);
          const draft = String(parsed.record_draft || '').trim();
          row[3] = draft ? '정상' : '응답 비어 있음';
          row[5] = draft;
        } catch (err) {
          row[3] = '오류';
          row[6] = String(err.message || err).slice(0, 1000);
        }

        row[4] = Math.round((Date.now() - callStartedAt) / 100) / 10;
      });
    });
  });

  sh.clear();
  sh.getRange(1, 1, 1, 7).setValues([['회사', '모델', '추론', '결과', '걸린 시간(초)', '받은 문장', '오류 내용']]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
  sh.getRange(1, 8).setValue(`점검 시각: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`);
  sh.setFrozenRows(1);
  sh.setColumnWidth(6, 320);
  sh.setColumnWidth(7, 420);
  sh.getRange(2, 6, rows.length, 2).setWrap(true);
  ss.setActiveSheet(sh);

  const count = status => rows.filter(r => r[3] === status).length;
  const failed = rows.filter(r => r[3] === '오류' || r[3] === '응답 비어 있음');
  const lines = [
    `AI 조합 점검 결과 ('AI점검' 시트)`,
    '',
    `정상 ${count('정상')}개, 오류 ${failed.length}개, 건너뜀 ${count('건너뜀')}개`,
  ];

  if (failed.length) {
    lines.push('', '[오류 난 조합]');
    failed.forEach(r => lines.push(`- ${r[1]} / ${r[2]}: ${String(r[6] || r[3]).slice(0, 150)}`));
  }

  log_('checkAllAiCombinations', failed.length ? 'ERROR' : 'OK', lines.join('\n'));
  getUi_().alert(lines.join('\n'));
}

// 행발문구추천 시트에서 선택한 학생 한 명의 입력으로, 키가 저장된 모든 모델 × 추론 수준(none·low·medium)으로 행발 추천을 만들어 '행발비교' 시트에 나란히 적는다.
// 실제 생성과 같은 프롬프트·분량 범위·재요청 경로를 쓰고, 행발문구추천 시트에는 쓰지 않는다.
function compareBehaviorAcrossModels() {
  const ui = getUi_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getActiveSheet();

  if (!source || source.getName() !== SHEET_NAMES.behaviorRecommendations) {
    ui.alert('행발문구추천 시트에서 비교할 학생 행을 하나 선택한 뒤 실행하세요.');
    return;
  }

  const sourceRow = source.getActiveRange().getRow();
  if (sourceRow < 2) {
    ui.alert('제목 줄이 아닌 학생 행을 선택하세요.');
    return;
  }

  const h = headerMap_(source);
  const lastCol = Math.max(source.getLastColumn(), BEHAVIOR_RECOMMENDATION_HEADERS.length);
  const data = buildBehaviorRecommendationData_(source.getRange(sourceRow, 1, 1, lastCol).getValues()[0], h);

  if (!data.hasSelectedTraits) {
    ui.alert(`${sourceRow}행에 고른 특성이 없습니다. 특성을 고른 행을 선택하세요.`);
    return;
  }

  const config = getConfigMap_();
  const behaviorPrompt = getBehaviorRecordPrompt_(config);
  const charRange = getBehaviorRecordCharRange_(config);
  const maxInputChars = Number(config.MAX_INPUT_CHARS || 30000);
  const systemGuide = behaviorRecommendationSystemGuide_();
  const props = PropertiesService.getUserProperties();
  const modelOptions = {
    claude: CONFIG_VALIDATION_OPTIONS.CLAUDE_MODEL,
    gpt: CONFIG_VALIDATION_OPTIONS.GPT_MODEL,
    gemini: CONFIG_VALIDATION_OPTIONS.GEMINI_MODEL,
  };
  const headers = ['회사', '모델', '추론', '결과', '걸린 시간(초)', '바이트', '추천 문구', '돌려 쓴 표현 후보', '확인할 점 / 오류'];
  const sh = ss.getSheetByName('행발비교') || ss.insertSheet('행발비교');

  // 거절된 요청을 하나씩 다시 보내다 6분 제한에 걸리면 일부가 남는다. 같은 입력·프롬프트로 다시 실행하면 이미 '정상'인 조합은 그대로 두고 나머지만 만든다.
  const runKey = Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    JSON.stringify([sourceRow, data, behaviorPrompt, charRange.min, charRange.max]),
    Utilities.Charset.UTF_8
  ));
  const passed = {};
  if (sh.getRange(1, headers.length + 2).getValue() === runKey && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues().forEach(r => {
      if (r[3] === '정상') passed[`${r[0]}|${r[1]}|${r[2]}`] = r;
    });
  }

  const rows = [];
  const todo = [];

  Object.keys(modelOptions).forEach(provider => {
    const hasKey = !!props.getProperty(AI_API_ENDPOINTS[provider].keyProperty);

    modelOptions[provider].forEach(model => CONFIG_VALIDATION_OPTIONS.REASONING.forEach(reasoning => {
      const previous = passed[`${provider}|${model}|${reasoning}`];
      if (previous) {
        rows.push(previous);
        return;
      }

      const row = [provider, model, reasoning, '', '', '', '', '', ''];
      rows.push(row);

      if (!hasKey) {
        row[3] = '건너뜀';
        row[8] = 'API 키가 저장되어 있지 않음';
        return;
      }

      todo.push({ row, ai: { provider, model, reasoning } });
    }));
  });

  // 모든 조합의 첫 요청을 한꺼번에 보내고, 분량 범위를 벗어난 것만 모아 재요청도 한꺼번에 보낸다.
  // 한 번에 보내다 거절(429 등)된 요청은 callAiParallel_이 하나씩 다시 보낸다.
  const sendAll = (items, retryContextOf) => {
    const startedAt = Date.now();
    const outcomes = callAiParallel_(items.map(item => ({
      provider: item.ai.provider,
      model: item.ai.model,
      reasoning: item.ai.reasoning,
      systemGuide,
      userPrompt: buildBehaviorRecommendationPrompt_(data, maxInputChars, behaviorPrompt, charRange, retryContextOf(item)),
      responseSchema: BEHAVIOR_RESPONSE_SCHEMA,
    })));
    const seconds = Math.round((Date.now() - startedAt) / 100) / 10;

    return outcomes.map(outcome => {
      if (outcome.error) return { error: outcome.error, seconds };
      try {
        return { generated: readBehaviorGenerated_(outcome.result.text), seconds };
      } catch (err) {
        return { error: err, seconds };
      }
    });
  };

  sendAll(todo, () => null).forEach((outcome, i) => {
    todo[i].first = outcome;
    todo[i].seconds = outcome.seconds;
  });

  const retries = todo.filter(item => item.first.generated && needsBehaviorCharRetry_(item.first.generated, charRange));
  if (retries.length) {
    try {
      sendAll(retries, item => ({ charCount: item.first.generated.charCount })).forEach((outcome, i) => {
        retries[i].retry = outcome;
        retries[i].seconds = Math.round((retries[i].seconds + outcome.seconds) * 10) / 10;
      });
    } catch (err) {
      // 시간 제한에 걸리면 재요청 없이 첫 답을 남긴다.
      retries.forEach(item => { item.retry = { error: err }; });
    }
  }

  todo.forEach(item => {
    const row = item.row;
    row[4] = item.seconds;

    if (item.first.error) {
      row[3] = isAiTimeBudgetError_(item.first.error) ? '건너뜀' : '오류';
      row[8] = String(item.first.error.message || item.first.error).slice(0, 1000);
      return;
    }

    const generated = item.retry
      ? settleBehaviorCharRetry_(
        item.first.generated,
        item.retry.generated || null,
        item.retry.error ? String(item.retry.error.message || item.retry.error) : '',
        charRange
      )
      : item.first.generated;

    row[3] = '정상';
    row[5] = generated.charCount;
    row[6] = generated.recommendedRecord;
    row[7] = formatBehaviorReframes_(generated.reframes);
    row[8] = generated.caution;
  });

  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(1, headers.length + 1).setValue(
    `${sourceRow}행 비교 · ${formatByteRange_(charRange)} · ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`
  );
  sh.getRange(1, headers.length + 2).setValue(runKey).setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(7, 420);
  sh.setColumnWidth(8, 320);
  sh.setColumnWidth(9, 260);
  sh.getRange(2, 1, rows.length, headers.length).setWrap(true).setVerticalAlignment('top');
  ss.setActiveSheet(sh);

  const count = status => rows.filter(r => r[3] === status).length;
  const lines = [
    `행발 추천 모델·추론 비교 ('행발비교' 시트, ${sourceRow}행 입력)`,
    '',
    `정상 ${count('정상')}개, 오류 ${count('오류')}개, 건너뜀 ${count('건너뜀')}개`,
  ];
  if (rows.some(r => r[3] === '건너뜀' && /실행 시간/.test(r[8]))) {
    lines.push('', '시간이 모자라 일부 모델을 건너뛰었습니다. 같은 행을 선택하고 메뉴를 한 번 더 실행하면 이어서 만듭니다.');
  }

  log_('compareBehaviorAcrossModels', count('오류') ? 'ERROR' : 'OK', lines.join('\n'));
  ui.alert(lines.join('\n'));
}

function clearOpenAiApiKey() {
  PropertiesService.getUserProperties().deleteProperty('OPENAI_API_KEY');

  // 혹시 예전에 Script Properties에 저장한 흔적이 있으면 같이 삭제
  PropertiesService.getScriptProperties().deleteProperty('OPENAI_API_KEY');

  log_('clearOpenAiApiKey', 'OK', 'OpenAI API 키 삭제 완료');
  getUi_().alert('현재 사용자 계정의 OpenAI API 키를 삭제했습니다.');
}

function clearGeminiApiKey() {
  PropertiesService.getUserProperties().deleteProperty('GEMINI_API_KEY');
  PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');

  log_('clearGeminiApiKey', 'OK', 'Gemini API 키 삭제 완료');
  getUi_().alert('현재 사용자 계정의 Gemini API 키를 삭제했습니다.');
}

function markSelectedAssignmentsIncluded() {
  setSelectedAssignmentsIncluded_(true);
}

function unmarkSelectedAssignmentsIncluded() {
  setSelectedAssignmentsIncluded_(false);
}

function setSelectedAssignmentsIncluded_(value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sh.getName() !== SHEET_NAMES.assignments) {
    getUi_().alert('과제목록 시트에서 과제 행을 선택한 뒤 실행하세요.');
    return;
  }

  const h = headerMap_(sh);
  if (!h.includeInFinal) {
    getUi_().alert('includeInFinal 열을 찾을 수 없습니다. 먼저 시트 초기 설정을 실행하세요.');
    return;
  }

  const range = sh.getActiveRange();
  const startRow = range.getRow();
  const numRows = range.getNumRows();

  if (startRow <= 1) {
    getUi_().alert('헤더가 아닌 과제 행을 선택하세요.');
    return;
  }

  const values = [];
  for (let i = 0; i < numRows; i++) {
    values.push([value]);
  }

  sh.getRange(startRow, h.includeInFinal, numRows, 1).setValues(values);

  getUi_().alert(
    value
      ? `${numRows}개 과제를 생기부 반영 대상으로 체크했습니다.`
      : `${numRows}개 과제의 생기부 반영 체크를 해제했습니다.`
  );
}

function collectSubmissionsForIncludedAssignments() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/spreadsheets.currentonly',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/classroom.courses.readonly',
    'https://www.googleapis.com/auth/classroom.coursework.students.readonly',
    'https://www.googleapis.com/auth/classroom.rosters.readonly',
    'https://www.googleapis.com/auth/classroom.profile.emails'
  ]);

  const ui = getUi_();
  const sh = ensureSheetHeaders_(SHEET_NAMES.assignments);
  const h = headerMap_(sh);
  const lastRow = sh.getLastRow();

  if (!h.includeInFinal) {
    ui.alert('includeInFinal 열을 찾을 수 없습니다. 먼저 시트 초기 설정을 실행하세요.');
    return;
  }

  if (lastRow < 2) {
    ui.alert('과제목록에 과제가 없습니다.');
    return;
  }

  const includedCount = countIncludedAssignments_();

  if (includedCount === 0) {
    ui.alert('includeInFinal에 체크된 과제가 없습니다.');
    return;
  }

  const confirm = ui.alert(
    '체크한 과제 제출물 일괄 수집',
    `현재 includeInFinal에 체크된 과제 ${includedCount}개의 제출물을 수집합니다.\n과제가 많으면 자동으로 나누어 이어서 처리합니다.\n계속할까요?`,
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) return;

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    ui.alert('이미 다른 자동 작업이 실행 중입니다.\n현재 작업이 끝난 뒤 다시 시작해 주세요.');
    return;
  }

  try {
    queueIncludedAssignmentsForSubmissionCollection_(sh, h, lastRow);
    deleteAutoSubmissionCollectionTriggers_();
    PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'ON');

    const result = processPendingSubmissionCollectionsCore_();
    saveAutoSubmissionCollectionLastResult_(result.blocked ? 'BLOCKED' : result.remaining === 0 ? 'DONE' : 'RUNNING', result);

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'OFF');
      deleteAutoSubmissionCollectionTriggers_();
      ui.alert(buildSubmissionCollectionResultMessage_('제출물 수집을 중단했습니다.', result));
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'OFF');
      deleteAutoSubmissionCollectionTriggers_();
      ui.alert(buildSubmissionCollectionResultMessage_('제출물 수집을 완료했습니다.', result));
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);
    scheduleNextAutoSubmissionCollectionTrigger_(nextDelayMs);

    ui.alert(
      buildSubmissionCollectionResultMessage_('제출물 첫 배치 수집 결과', result) +
      `\n\n남은 과제는 ${Math.round(nextDelayMs / 1000)}초 뒤 자동으로 이어서 수집합니다.`
    );
  } catch (err) {
    const message = String(err.message || err);
    const result = {
      processed: 0,
      submissionCount: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingSubmissionCollections_(),
      blocked: true,
      blockingReason: '제출물 수집 시작 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'OFF');
    deleteAutoSubmissionCollectionTriggers_();
    saveAutoSubmissionCollectionLastResult_('BLOCKED', result);
    log_('collectSubmissionsForIncludedAssignments', 'ERROR', message);
    ui.alert(buildSubmissionCollectionResultMessage_('제출물 수집 중 오류가 발생했습니다.', result));
  } finally {
    lock.releaseLock();
  }
}

function autoProcessPendingSubmissionCollections() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    rescheduleAutoJobWhenBusy_('autoProcessPendingSubmissionCollections', AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, scheduleNextAutoSubmissionCollectionTrigger_);
    return;
  }

  try {
    deleteAutoSubmissionCollectionTriggers_();

    const status = PropertiesService.getUserProperties().getProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY) || 'OFF';

    if (status !== 'ON') {
      log_('autoProcessPendingSubmissionCollections', 'STOP', '자동 수집 상태가 OFF라 실행하지 않음');
      return;
    }

    beginAutoRun_(scheduleNextAutoSubmissionCollectionTrigger_);

    const result = processPendingSubmissionCollectionsCore_();
    saveAutoSubmissionCollectionLastResult_(result.blocked ? 'BLOCKED' : result.remaining === 0 ? 'DONE' : 'RUNNING', result);

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'OFF');
      toastAutoRecordDraft_(result.blockingMessage || '제출물 수집을 중단했습니다.', result.blockingReason || '제출물 수집 중단');
      log_('autoProcessPendingSubmissionCollections', 'BLOCKED', `${result.blockingReason}: ${result.blockingMessage}`);
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'OFF');
      toastAutoRecordDraft_(
        `수집 완료 ${result.processed}개 과제, 제출물 ${result.submissionCount || 0}개, 실패 ${result.failed}개`,
        '제출물 수집 완료'
      );
      log_('autoProcessPendingSubmissionCollections', 'DONE', '남은 수집대기 과제가 없어 종료했습니다.');
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoSubmissionCollectionTrigger_(nextDelayMs);
    log_(
      'autoProcessPendingSubmissionCollections',
      'NEXT',
      `남은 제출물 수집 과제 ${result.remaining}개. ${Math.round(nextDelayMs / 1000)}초 뒤 다음 실행 예약`
    );
  } catch (err) {
    const message = String(err.message || err);
    PropertiesService.getUserProperties().setProperty(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, 'OFF');
    saveAutoSubmissionCollectionLastResult_('BLOCKED', {
      processed: 0,
      submissionCount: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingSubmissionCollections_(),
      blocked: true,
      blockingReason: '제출물 자동 수집 오류',
      blockingMessage: message,
    });
    toastAutoRecordDraft_(message, '제출물 자동 수집 오류');
    log_('autoProcessPendingSubmissionCollections', 'ERROR', message);
  } finally {
    endAutoRun_(AUTO_SUBMISSION_COLLECTION_TRIGGER_PROPERTY, deleteAutoSubmissionCollectionTriggers_);
    lock.releaseLock();
  }
}

function processPendingSubmissionCollectionsCore_() {
  const config = getConfigMap_();
  const maxRunSeconds = Number(config.MAX_RUN_SECONDS || 270);
  const startedAt = Date.now();
  const sh = ensureSheetHeaders_(SHEET_NAMES.assignments);
  const h = headerMap_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  let processed = 0;
  let submissionCount = 0;
  let skipped = 0;
  let failed = 0;
  let blocked = false;
  let blockingReason = '';
  let blockingMessage = '';
  let touchedSubmissions = false;

  if (!h.includeInFinal || !h.collectStatus || lastRow < 2) {
    return {
      processed,
      submissionCount,
      skipped,
      failed,
      remaining: 0,
      blocked,
      blockingReason,
      blockingMessage,
    };
  }

  const assignmentValues = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const studentLookupIndex = buildStudentLookupIndex_();
  let submissionsSheetEnsured = false;

  for (let row = 2; row <= lastRow; row++) {
    const elapsedSeconds = (Date.now() - startedAt) / 1000;

    if (elapsedSeconds > maxRunSeconds) {
      log_(
        'processPendingSubmissionCollectionsCore_',
        'STOP',
        `실행 시간 제한 접근으로 중단. 수집 과제 ${processed}개, 실패 ${failed}개`
      );
      break;
    }

    const rowValues = assignmentValues[row - 2] || [];
    const include = rowValueAt_(rowValues, h.includeInFinal) === true;
    if (!include) continue;

    const status = String(rowValueAt_(rowValues, h.collectStatus) || '').trim();
    if (!isPendingSubmissionCollectionStatus_(status)) continue;

    const assignment = {
      row,
      courseId: rowValueAt_(rowValues, h.courseId),
      courseName: rowValueAt_(rowValues, h.courseName),
      courseWorkId: rowValueAt_(rowValues, h.courseWorkId),
      assignmentTitle: rowValueAt_(rowValues, h.assignmentTitle),
      maxPoints: h.maxPoints ? rowValueAt_(rowValues, h.maxPoints) : '',
      assignmentDescription: h.assignmentDescription
        ? rowValueAt_(rowValues, h.assignmentDescription)
        : '',
    };

    if (!assignment.courseId || !assignment.courseWorkId) {
      setCellValueAndNote_(sh, row, h.collectStatus, '수집오류', 'courseId 또는 courseWorkId가 비어 있습니다.');
      failed++;
      continue;
    }

    try {
      setCellValueAndNote_(sh, row, h.collectStatus, '수집중', '');

      const rows = buildSubmissionRowsForAssignment_(assignment, studentLookupIndex);
      if (!submissionsSheetEnsured) {
        ensureSheetHeaders_(SHEET_NAMES.submissions);
        submissionsSheetEnsured = true;
      }

      if (rows.length > 0) {
        upsertByKey_(SHEET_NAMES.submissions, 'submissionKey', rows);
        touchedSubmissions = true;
      }

      processed++;
      submissionCount += rows.length;
      setCellValueAndNote_(sh, row, h.collectStatus, `수집완료 ${rows.length}명`, '');

      Utilities.sleep(500);
    } catch (err) {
      const message = String(err.message || err);
      failed++;
      setCellValueAndNote_(sh, row, h.collectStatus, '수집오류', message);
      log_(
        'processPendingSubmissionCollectionsCore_',
        'ERROR',
        `${assignment.assignmentTitle}: ${message}`
      );
    }
  }

  if (touchedSubmissions) {
    sortSubmissionsSheet_();
    styleSubmissionsSheet_();
  }

  const remaining = countPendingSubmissionCollections_();

  log_(
    'processPendingSubmissionCollectionsCore_',
    'OK',
    `수집 과제 ${processed}개, 제출물 ${submissionCount}개, 실패 ${failed}개, 스킵 ${skipped}개, 남은 대기 ${remaining}개`
  );

  return {
    processed,
    submissionCount,
    skipped,
    failed,
    remaining,
    blocked,
    blockingReason,
    blockingMessage,
  };
}

function queueIncludedAssignmentsForSubmissionCollection_(sh, h, lastRow) {
  if (!h.includeInFinal || !h.collectStatus || lastRow < 2) return 0;

  const rowCount = lastRow - 1;
  const includeValues = sh.getRange(2, h.includeInFinal, rowCount, 1).getValues();
  const statusRange = sh.getRange(2, h.collectStatus, rowCount, 1);
  const statusValues = statusRange.getValues();
  const statusNotes = statusRange.getNotes();
  let queued = 0;

  for (let i = 0; i < rowCount; i++) {
    const include = includeValues[i][0] === true;

    if (!include) continue;

    statusValues[i][0] = '수집대기';
    statusNotes[i][0] = '';
    queued++;
  }

  if (queued > 0) {
    statusRange.setValues(statusValues);
    statusRange.setNotes(statusNotes);
  }

  return queued;
}

function countPendingSubmissionCollections_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.assignments);
    if (!sh) return 0;

    const h = headerMap_(sh);
    if (!h.includeInFinal || !h.collectStatus) return 0;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return 0;

    const includeValues = sh.getRange(2, h.includeInFinal, lastRow - 1, 1).getValues();
    const statusValues = sh.getRange(2, h.collectStatus, lastRow - 1, 1).getValues();
    let count = 0;

    for (let i = 0; i < includeValues.length; i++) {
      const include = includeValues[i][0] === true;
      const status = String(statusValues[i][0] || '').trim();

      if (include && isPendingSubmissionCollectionStatus_(status)) count++;
    }

    return count;
  } catch (err) {
    return 0;
  }
}

function isPendingSubmissionCollectionStatus_(status) {
  const value = String(status || '').trim();
  return value === '수집대기' || value === '수집중';
}

function buildSubmissionCollectionResultMessage_(title, result) {
  const lines = [
    title,
    '',
    `수집 완료 과제: ${result.processed || 0}개`,
    `수집 제출물: ${result.submissionCount || 0}개`,
    `실패 과제: ${result.failed || 0}개`,
    `스킵: ${result.skipped || 0}개`,
    `남은 수집대기 과제: ${result.remaining || 0}개`,
  ];

  if (result.blocked) {
    lines.push('', `중단 사유: ${result.blockingReason || '확인 필요'}`);
    lines.push(result.blockingMessage || '제출물 수집을 계속할 수 없어 중단했습니다.');
  }

  return lines.join('\n');
}

function scheduleNextAutoSubmissionCollectionTrigger_(delayMs) {
  deleteAutoSubmissionCollectionTriggers_();

  ScriptApp.newTrigger('autoProcessPendingSubmissionCollections')
    .timeBased()
    .after(delayMs)
    .create();

  log_(
    'scheduleNextAutoSubmissionCollectionTrigger_',
    'OK',
    `${Math.round(delayMs / 1000)}초 뒤 다음 제출물 수집 예약`
  );
}

function deleteAutoSubmissionCollectionTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoProcessPendingSubmissionCollections') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function saveAutoSubmissionCollectionLastResult_(status, result) {
  const payload = {
    status,
    processed: result.processed || 0,
    submissionCount: result.submissionCount || 0,
    failed: result.failed || 0,
    skipped: result.skipped || 0,
    remaining: result.remaining || 0,
    blocked: !!result.blocked,
    blockingReason: result.blockingReason || '',
    blockingMessage: result.blockingMessage || '',
    updatedAt: new Date().toISOString(),
  };

  PropertiesService.getUserProperties().setProperty(
    'AUTO_SUBMISSION_COLLECTION_LAST_RESULT',
    JSON.stringify(payload)
  );
}

function getAutoSubmissionCollectionLastResult_() {
  const raw = PropertiesService.getUserProperties().getProperty('AUTO_SUBMISSION_COLLECTION_LAST_RESULT');
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}


function countIncludedAssignments_() {
  const sh = getSheet_(SHEET_NAMES.assignments);
  const h = headerMap_(sh);
  const lastRow = sh.getLastRow();

  if (!h.includeInFinal || lastRow < 2) return 0;

  return sh.getRange(2, h.includeInFinal, lastRow - 1, 1)
    .getValues()
    .filter(row => row[0] === true)
    .length;
}

function buildSubmissionRowsForAssignment_(assignment, studentLookupIndex) {
  const courseId = assignment.courseId;
  const courseName = assignment.courseName;
  const courseWorkId = assignment.courseWorkId;
  const assignmentTitle = assignment.assignmentTitle;
  const maxPoints = assignment.maxPoints || '';
  const assignmentDescription = assignment.assignmentDescription || '';


  if (!isMyTeacherCourse_(courseId)) {
    throw new Error(`현재 계정이 교사로 등록된 수업이 아닙니다: ${courseName}`);
  }

  const submissions = listAllStudentSubmissions_(courseId, courseWorkId);
  const rows = [];

  submissions.forEach(sub => {
    try {
      const profile = getUserProfile_(sub.userId);
      const email = profile.emailAddress || '';
      const name = profile.name && profile.name.fullName ? profile.name.fullName : '';
      const studentInfo = lookupStudent_(email, name, sub.userId, studentLookupIndex);
      const extracted = extractSubmissionText_(sub);

      const submissionKey = `${courseId}|${courseWorkId}|${sub.id}`;

      rows.push([
        submissionKey,
        studentInfo.studentNo || '',
        studentInfo.name || name,
        studentInfo.email || email,
        courseId,
        courseName,
        courseWorkId,
        assignmentTitle,
        sub.id,
        sub.state || '',
        sub.late === true ? 'TRUE' : 'FALSE',
        sub.draftGrade ?? '',
        sub.assignedGrade ?? '',
        maxPoints,
        '',
        sub.updateTime || '',
        sub.alternateLink || '',
        extracted.fileTitles.join('\n'),
        extracted.fileLinks.join('\n'),
        extracted.text,
        extracted.text.length,
        new Date(),
        '',
        assignmentDescription,
      ]);
    } catch (err) {
      log_('buildSubmissionRowsForAssignment_', 'ERROR', `${assignmentTitle} / ${sub.id}: ${err.message}`);
    }
  });

  return rows;
}

function sortSubmissionsSheet_() {
  const sh = getSheet_(SHEET_NAMES.submissions);
  const h = headerMap_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 3) return;

  const sortSpecs = [];

  if (h.courseName) {
    sortSpecs.push({ column: h.courseName, ascending: true });
  }

  if (h.studentNo) {
    sortSpecs.push({ column: h.studentNo, ascending: true });
  }

  if (h.studentName) {
    sortSpecs.push({ column: h.studentName, ascending: true });
  }

  if (h.assignmentTitle) {
    sortSpecs.push({ column: h.assignmentTitle, ascending: true });
  }

  if (sortSpecs.length === 0) return;

  sh.getRange(2, 1, lastRow - 1, lastCol).sort(sortSpecs);
}

function generateStudentFinalRecordsForIncludedAssignments() {
  const ui = getUi_();
  const checkedAssignments = getIncludedAssignmentKeySet_();

  if (checkedAssignments.size === 0) {
    ui.alert('생기부 반영 대상으로 체크된 과제가 없습니다.');
    return;
  }

  try {
    requireCurrentAiApiKey_();
  } catch (err) {
    ui.alert(String(err.message || err));
    return;
  }

  const groups = buildStudentFinalRecordGroups_(checkedAssignments);

  if (groups.length === 0) {
    ui.alert('학생별로 합칠 생기부초안이 없습니다.\n과제목록에서 생기부 반영 대상을 체크하고 생기부초안을 먼저 생성하세요.');
    return;
  }

  const runId = makeStudentFinalRunId_();
  const queuedCount = queueStudentFinalRecordGroups_(groups, runId);

  log_(
    'generateStudentFinalRecordsForIncludedAssignments',
    'QUEUE',
    `${queuedCount}명 학생별 최종 생기부 대기열 생성`
  );

  refreshDashboard();
  startAutoStudentFinalRecordTrigger_();
}

function makeStudentFinalRunId_() {
  return `SFR-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')}-${Utilities.getUuid().slice(0, 8)}`;
}

function ensureStudentFinalRecordQueueHeaders_() {
  const sh = getSheet_(SHEET_NAMES.studentFinalRecords);
  const requiredHeaders = HEADERS[SHEET_NAMES.studentFinalRecords];
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const currentHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(v => String(v || '').trim());
  const hasAnyHeader = currentHeaders.some(header => !!header);

  if (!hasAnyHeader) {
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return headerMap_(sh);
  }

  const existing = {};

  currentHeaders.forEach(header => {
    if (header) existing[header] = true;
  });

  const missing = requiredHeaders.filter(header => !existing[header]);

  if (missing.length > 0) {
    sh.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing]);
  }

  return headerMap_(sh);
}

function buildStudentFinalRecordGroups_(checkedAssignments) {
  checkedAssignments = checkedAssignments || getIncludedAssignmentKeySet_();
  if (checkedAssignments.size === 0) return [];

  const records = getSheetObjects_(SHEET_NAMES.records);
  const latestBySubmissionKey = {};

  records.forEach(r => {
    const submissionKey = String(r.submissionKey || '').trim();
    if (!submissionKey) return;

    if (!isSubmissionFromIncludedAssignment_(submissionKey, checkedAssignments)) return;

    const existing = latestBySubmissionKey[submissionKey];
    const currentTime = new Date(r.createdAt || 0).getTime();
    const existingTime = existing ? new Date(existing.createdAt || 0).getTime() : 0;

    if (!existing || currentTime >= existingTime) {
      latestBySubmissionKey[submissionKey] = r;
    }
  });

  const grouped = {};

  Object.values(latestBySubmissionKey).forEach(r => {
    const studentNo = String(r.studentNo || '').trim();
    const email = String(r.email || '').trim();
    const studentName = String(r.studentName || '').trim();
    const courseName = String(r.courseName || '').trim();

    const key = [
      studentNo || 'NO_STUDENT_NO',
      email || 'NO_EMAIL',
      courseName || 'NO_COURSE'
    ].join('|');

    if (!grouped[key]) {
      grouped[key] = {
        key,
        studentNo,
        studentName,
        email,
        courseName,
        items: [],
      };
    }

    grouped[key].items.push(r);
  });

  return Object.keys(grouped)
    .sort()
    .map(key => {
      const group = grouped[key];
      const sourceSubmissionKeys = group.items.map(r => r.submissionKey).join('\n');
      const sourceAssignments = group.items.map(r => r.assignmentTitle).join('\n');
      const evidencePack = group.items.map((r, index) => {
        const assignmentTitle = r.assignmentTitle || `과제${index + 1}`;
        const evidence = r.evidenceSummary || r.finalRecord || r.recordDraft || '';
        const caution = r.caution || '';
        const gradeText = r.gradeText || '';

        return [
          `# ${index + 1}. ${assignmentTitle}`,
          gradeText ? `[점수 정보]\n${gradeText}` : '',
          '[근거 요약]',
          evidence,
          caution ? `[주의] ${caution}` : '',
        ].filter(Boolean).join('\n');
      }).join('\n\n---\n\n');

      return {
        key: group.key,
        studentNo: group.studentNo,
        studentName: group.studentName,
        email: group.email,
        courseName: group.courseName,
        assignmentCount: group.items.length,
        sourceSubmissionKeys,
        sourceAssignments,
        evidencePack,
      };
    });
}

function queueStudentFinalRecordGroups_(groups, runId) {
  if (!groups || groups.length === 0) return 0;

  const sh = getSheet_(SHEET_NAMES.studentFinalRecords);
  ensureStudentFinalRecordQueueHeaders_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(v => String(v || '').trim());
  const rows = groups.map(group => {
    const data = {
      studentNo: group.studentNo,
      studentName: group.studentName,
      email: group.email,
      courseName: group.courseName,
      assignmentCount: group.assignmentCount,
      sourceSubmissionKeys: group.sourceSubmissionKeys,
      sourceAssignments: group.sourceAssignments,
      evidencePack: group.evidencePack,
      finalRecord: '',
      charCount: '',
      caution: '',
      teacherChecked: '',
      createdAt: '',
      model: '',
      usage: '',
      finalQueueKey: `${runId}|${group.key}`,
      finalStatus: '대기',
      finalError: '',
    };

    return headers.map(header => data[header] !== undefined ? data[header] : '');
  });

  const startRow = Math.max(sh.getLastRow() + 1, 2);
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  formatStudentFinalRecordRows_(Array.from({ length: rows.length }, (_, index) => startRow + index));

  return rows.length;
}

function processPendingStudentFinalRecordsCore_(showUi) {
  const sh = getSheet_(SHEET_NAMES.studentFinalRecords);
  const h = ensureStudentFinalRecordQueueHeaders_();
  const config = getConfigMap_();
  const batchSize = RUN_ITEM_LIMIT_;
  const maxRunSeconds = Number(config.MAX_RUN_SECONDS || 270);
  const startedAt = Date.now();
  const ai = getAiProviderAndModel_(config);
  const provider = ai.provider;
  const model = ai.model;
  const reasoning = ai.reasoning;
  const byteRange = getRecordByteRange_(config, 'final');
  const recordDraftMinChars = approxCharsFromBytes_(byteRange.min);
  const recordDraftMaxChars = approxCharsFromBytes_(byteRange.max);
  const systemGuide = getSystemGuide_(config);
  const studentFinalPromptTemplate = getConfigPromptTemplate_(config, 'PROMPT_2', defaultStudentFinalRecordPrompt_());
  const lastRow = sh.getLastRow();

  let processed = 0;
  let skipped = 0;
  let failed = 0;
  let blocked = false;
  let blockingReason = '';
  let blockingMessage = '';
  const touchedRows = [];

  if (!h.finalStatus || lastRow < 2) {
    return {
      processed,
      skipped,
      failed,
      remaining: 0,
      blocked,
      blockingReason,
      blockingMessage,
    };
  }

  // get(col)은 그 행의 칸 값을 돌려준다. 루프는 칸을 바로 읽고, 미리 받기는 한 번에 읽어 둔 값을 쓴다.
  const buildFinalUserPrompt = get => buildStudentFinalRecordPrompt_({
    promptTemplate: studentFinalPromptTemplate,
    courseName: get(h.courseName),
    assignmentCount: get(h.assignmentCount),
    evidencePack: sanitizePromptText_(get(h.evidencePack), {
      studentName: get(h.studentName),
      studentNo: get(h.studentNo),
      email: get(h.email),
    }),
    recordDraftMinChars,
    recordDraftMaxChars,
    recordDraftMinBytes: byteRange.min,
    recordDraftMaxBytes: byteRange.max,
  });

  const prefetcher = createAiPrefetcher_(config, 'processPendingStudentFinalRecordsCore_');
  let finalValues = null;
  const upcomingFinalRequests = fromRow => {
    if (!finalValues) finalValues = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

    const requests = [];
    const limit = Math.min(FAST_GENERATION_CONCURRENCY, batchSize - processed);

    for (let r = fromRow; r <= lastRow && requests.length < limit; r++) {
      const values = finalValues[r - 2] || [];
      const get = col => rowValueAt_(values, col);
      const status = String(get(h.finalStatus) || '').trim();
      if (status !== '대기' && status !== '재시도필요') continue;
      if (!get(h.studentName) || !get(h.evidencePack)) continue;

      requests.push({ provider, model, systemGuide, userPrompt: buildFinalUserPrompt(get), reasoning });
    }

    return requests;
  };

  for (let row = 2; row <= lastRow; row++) {
    const elapsedSeconds = (Date.now() - startedAt) / 1000;

    if (elapsedSeconds > maxRunSeconds) {
      log_(
        'processPendingStudentFinalRecordsCore_',
        'STOP',
        `실행 시간 제한 접근으로 중단. 처리 ${processed}명, 실패 ${failed}명`
      );
      break;
    }

    if (processed >= batchSize) break;

    const status = String(valueAt_(sh, row, h.finalStatus) || '').trim();

    if (status !== '대기' && status !== '재시도필요') continue;

    const studentNo = valueAt_(sh, row, h.studentNo);
    const studentName = valueAt_(sh, row, h.studentName);
    const email = valueAt_(sh, row, h.email);
    const courseName = valueAt_(sh, row, h.courseName);
    const assignmentCount = valueAt_(sh, row, h.assignmentCount);
    const evidencePack = valueAt_(sh, row, h.evidencePack);

    if (!studentName || !evidencePack) {
      if (h.finalStatus) sh.getRange(row, h.finalStatus).setValue('오류');
      if (h.finalError) sh.getRange(row, h.finalError).setValue('학생 정보 또는 근거 묶음이 비어 있습니다.');
      skipped++;
      touchedRows.push(row);
      continue;
    }

    const userPrompt = buildFinalUserPrompt(col => valueAt_(sh, row, col));

    prefetcher.prefetch(() => upcomingFinalRequests(row));

    try {
      sh.getRange(row, h.finalStatus).setValue('생성중');
      if (h.finalError) sh.getRange(row, h.finalError).setValue('');

      const generated = createRecordDraftWithLength_(
        provider,
        model,
        systemGuide,
        userPrompt,
        byteRange.min,
        byteRange.max,
        reasoning
      );

      const parsed = generated.parsed;
      const finalRecord = normalizeAiText_(parsed.record_draft);
      const caution = normalizeAiText_(parsed.caution);
      const evidenceSummary = normalizeAiText_(parsed.evidence_summary);

      sh.getRange(row, h.finalRecord)
        .setValue(finalRecord)
        .setNote(evidenceSummary ? `[evidence_summary]\n${evidenceSummary}` : '');
      sh.getRange(row, h.charCount).setValue(neisByteCount_(finalRecord));
      sh.getRange(row, h.caution).setValue(caution);
      sh.getRange(row, h.teacherChecked).setValue(false);
      sh.getRange(row, h.createdAt).setValue(new Date());
      sh.getRange(row, h.model).setValue(`${provider}:${model}`);
      sh.getRange(row, h.usage).setValue(JSON.stringify(generated.usage || {}));
      sh.getRange(row, h.finalStatus).setValue(finalRecord ? '생성완료' : '근거부족');

      processed++;
      touchedRows.push(row);
      Utilities.sleep(1200);

    } catch (err) {
      if (isAiTimeBudgetError_(err)) {
        sh.getRange(row, h.finalStatus).setValue('대기');
        log_('processPendingStudentFinalRecordsCore_', 'STOP', `${studentNo}${studentName}: ${err.message}`);
        break;
      }

      const message = String(err.message || '');
      const errorInfo = classifyRecordDraftError_(message);

      sh.getRange(row, h.finalStatus).setValue(errorInfo.status);
      if (h.finalError) {
        sh.getRange(row, h.finalError).setValue(`${errorInfo.reason}: ${message}`);
      }

      if (errorInfo.shouldStop) {
        blocked = true;
        blockingReason = errorInfo.reason;
        blockingMessage = errorInfo.userMessage;
      }

      failed++;
      touchedRows.push(row);
      log_(
        'processPendingStudentFinalRecordsCore_',
        blocked ? 'STOP' : 'ERROR',
        `${studentNo}${studentName}: ${errorInfo.reason} / ${message}`
      );

      if (blocked) break;
    }
  }

  prefetcher.finish();

  if (touchedRows.length > 0) {
    formatStudentFinalRecordRows_(touchedRows);
  }

  refreshDashboard();

  const remaining = countPendingStudentFinalRecords_();

  log_(
    'processPendingStudentFinalRecordsCore_',
    'OK',
    `처리 ${processed}명, 실패 ${failed}명, 스킵 ${skipped}명, 남은 대기 ${remaining}명`
  );

  return {
    processed,
    skipped,
    failed,
    remaining,
    blocked,
    blockingReason,
    blockingMessage,
  };
}

function buildRecordDraftPrompt_(data) {
  const template = data.promptTemplate || defaultRecordDraftPrompt_();
  const renderedPrompt = renderPromptTemplate_(template, data);
  const parts = [renderedPrompt];
  const missingInputLines = buildMissingPromptLines_(template, [
    ['courseName', `수업명: ${data.courseName}`],
    ['assignmentTitle', `과제명: ${data.assignmentTitle}`],
    ['state', `제출상태: ${data.state}`],
    ['late', `지각제출: ${data.late}`],
    ['gradeText', `점수 정보: ${data.gradeText}`],
  ]);

  if (missingInputLines.length > 0) {
    parts.push(['[입력 정보]'].concat(missingInputLines).join('\n'));
  }

  if (!promptTemplateHasPlaceholder_(template, 'assignmentDescription')) {
    parts.push([
      '[과제 문항/교사 지시사항]',
      data.assignmentDescription,
    ].join('\n'));
  }

  parts.push(buildByteRangeInstruction_('record_draft', {
    min: data.recordDraftMinBytes,
    max: data.recordDraftMaxBytes,
  }));
  parts.push([
    '[기본 출력 조건]',
    '생활기록부에 실제로 옮겨 적을 문장은 record_draft에만 작성한다.',
    '반드시 JSON 형식으로만 답한다.',
  ].join('\n'));

  if (!promptTemplateHasPlaceholder_(template, 'extractedText')) {
    parts.push([
      '[학생 제출물 원문]',
      data.extractedText,
    ].join('\n'));
  }

  parts.push(buildRecordDraftNoEvidenceRule_());
  parts.push(buildGradeContextRule_());

  return parts.filter(part => String(part || '').trim()).join('\n\n');
}

function buildStudentFinalRecordPrompt_(data) {
  const template = data.promptTemplate || defaultStudentFinalRecordPrompt_();
  const renderedPrompt = renderPromptTemplate_(template, data);
  const parts = [renderedPrompt];
  const missingInputLines = buildMissingPromptLines_(template, [
    ['courseName', `수업명: ${data.courseName}`],
    ['assignmentCount', `반영 과제 수: ${data.assignmentCount}`],
  ]);

  if (missingInputLines.length > 0) {
    parts.push(['[입력 정보]'].concat(missingInputLines).join('\n'));
  }

  parts.push(buildByteRangeInstruction_('record_draft', {
    min: data.recordDraftMinBytes,
    max: data.recordDraftMaxBytes,
  }));
  parts.push([
    '[기본 출력 조건]',
    '생활기록부에 실제로 옮겨 적을 문장은 record_draft에만 작성한다.',
    '반드시 JSON 형식으로만 답한다.',
  ].join('\n'));

  if (!promptTemplateHasPlaceholder_(template, 'evidencePack')) {
    parts.push([
      '[과제별 근거 묶음]',
      data.evidencePack,
    ].join('\n'));
  }

  parts.push(buildStudentFinalNoEvidenceRule_());
  parts.push(buildGradeContextRule_());

  return parts.filter(part => String(part || '').trim()).join('\n\n');
}

function countPendingStudentFinalRecords_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.studentFinalRecords);
  if (!sh) return 0;

  const h = headerMap_(sh);
  if (!h.finalStatus) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const values = sh.getRange(2, h.finalStatus, lastRow - 1, 1).getValues();

  return values.filter(row => {
    const status = String(row[0] || '').trim();
    return status === '대기' || status === '재시도필요';
  }).length;
}

function startAutoStudentFinalRecordTrigger_() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/spreadsheets.currentonly',
    'https://www.googleapis.com/auth/script.external_request',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/classroom.courses.readonly',
    'https://www.googleapis.com/auth/classroom.coursework.students.readonly',
    'https://www.googleapis.com/auth/classroom.rosters.readonly',
    'https://www.googleapis.com/auth/classroom.profile.emails'
  ]);

  const ui = getUi_();
  const pendingCount = countPendingStudentFinalRecords_();

  if (pendingCount === 0) {
    ui.alert('학생별 최종 생기부 대기열이 없습니다.');
    return;
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    ui.alert('이미 다른 자동 생성 작업이 실행 중입니다.\n현재 작업이 끝난 뒤 다시 시작해 주세요.');
    return;
  }

  try {
    deleteAutoStudentFinalRecordTriggers_();

    const props = PropertiesService.getUserProperties();
    props.setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'ON');

    const result = runFirstBatchQuickly_(() => processPendingStudentFinalRecordsCore_(false));

    if (result.blocked) {
      props.setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'OFF');
      deleteAutoStudentFinalRecordTriggers_();
      saveAutoStudentFinalLastResult_('BLOCKED', result);

      log_(
        'startAutoStudentFinalRecordTrigger_',
        'BLOCKED',
        `${result.blockingReason}: ${result.blockingMessage}`
      );

      ui.alert(buildRecordDraftResultMessage_('학생별 최종 생기부 생성을 중단했습니다.', result, '명'));
      return;
    }

    if (result.remaining === 0) {
      props.setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'OFF');
      saveAutoStudentFinalLastResult_('DONE', result);
      log_(
        'startAutoStudentFinalRecordTrigger_',
        'DONE',
        `첫 배치 즉시 처리 후 남은 대기열이 없어 종료. 처리 ${result.processed}명, 실패 ${result.failed}명, 스킵 ${result.skipped}명`
      );

      ui.alert(buildRecordDraftResultMessage_('학생별 최종 생기부 생성을 완료했습니다.', result, '명'));
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoStudentFinalRecordTrigger_(nextDelayMs);
    saveAutoStudentFinalLastResult_('RUNNING', result);

    log_(
      'startAutoStudentFinalRecordTrigger_',
      'OK',
      `첫 배치 처리 완료. 처리 ${result.processed}명, 실패 ${result.failed}명, 스킵 ${result.skipped}명, 남은 대기 ${result.remaining}명`
    );

  } catch (err) {
    const message = String(err.message || err);
    const result = {
      processed: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingStudentFinalRecords_(),
      blocked: true,
      blockingReason: '학생별 최종 생성 시작 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'OFF');
    deleteAutoStudentFinalRecordTriggers_();
    saveAutoStudentFinalLastResult_('BLOCKED', result);
    log_('startAutoStudentFinalRecordTrigger_', 'ERROR', message);
    ui.alert(buildRecordDraftResultMessage_('학생별 최종 생기부 생성 시작 중 오류가 발생했습니다.', result, '명'));
  } finally {
    lock.releaseLock();
  }
}

function autoProcessPendingStudentFinalRecords() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    rescheduleAutoJobWhenBusy_('autoProcessPendingStudentFinalRecords', 'AUTO_STUDENT_FINAL_TRIGGER', scheduleNextAutoStudentFinalRecordTrigger_);
    return;
  }

  try {
    deleteAutoStudentFinalRecordTriggers_();

    const status = PropertiesService.getUserProperties().getProperty('AUTO_STUDENT_FINAL_TRIGGER') || 'OFF';

    if (status !== 'ON') {
      log_('autoProcessPendingStudentFinalRecords', 'STOP', '자동 생성 상태가 OFF라 실행하지 않음');
      return;
    }

    beginAutoRun_(scheduleNextAutoStudentFinalRecordTrigger_, [SHEET_NAMES.studentFinalRecords, 'finalStatus']);

    const result = processPendingStudentFinalRecordsCore_(false);

    if (result.blocked) {
      PropertiesService.getUserProperties().setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'OFF');
      saveAutoStudentFinalLastResult_('BLOCKED', result);
      toastAutoRecordDraft_(result.blockingMessage || '학생별 최종 생성을 중단했습니다.', result.blockingReason || '학생별 최종 생성 중단');
      log_(
        'autoProcessPendingStudentFinalRecords',
        'BLOCKED',
        `${result.blockingReason}: ${result.blockingMessage}`
      );
      return;
    }

    if (result.remaining === 0) {
      PropertiesService.getUserProperties().setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'OFF');
      saveAutoStudentFinalLastResult_('DONE', result);
      toastAutoRecordDraft_(
        `생성 완료 ${result.processed}명, 실패 ${result.failed}명, 스킵 ${result.skipped}명`,
        '학생별 최종 생성 완료'
      );
      log_('autoProcessPendingStudentFinalRecords', 'DONE', '남은 학생별 최종 대기열이 없어 종료했습니다.');
      return;
    }

    const config = getConfigMap_();
    const nextDelayMs = Number(config.AUTO_NEXT_DELAY_MS || 60000);

    scheduleNextAutoStudentFinalRecordTrigger_(nextDelayMs);
    saveAutoStudentFinalLastResult_('RUNNING', result);

    log_(
      'autoProcessPendingStudentFinalRecords',
      'NEXT',
      `남은 대기 ${result.remaining}명. ${Math.round(nextDelayMs / 1000)}초 뒤 다음 실행 예약`
    );

  } catch (err) {
    const message = String(err.message || err);
    const result = {
      processed: 0,
      failed: 1,
      skipped: 0,
      remaining: countPendingStudentFinalRecords_(),
      blocked: true,
      blockingReason: '학생별 최종 자동 처리 오류',
      blockingMessage: message,
    };

    PropertiesService.getUserProperties().setProperty('AUTO_STUDENT_FINAL_TRIGGER', 'OFF');
    saveAutoStudentFinalLastResult_('BLOCKED', result);
    toastAutoRecordDraft_(message, '학생별 최종 자동 처리 오류');
    log_('autoProcessPendingStudentFinalRecords', 'ERROR', message);

  } finally {
    endAutoRun_('AUTO_STUDENT_FINAL_TRIGGER', deleteAutoStudentFinalRecordTriggers_);
    lock.releaseLock();
  }
}

function scheduleNextAutoStudentFinalRecordTrigger_(delayMs) {
  deleteAutoStudentFinalRecordTriggers_();

  ScriptApp.newTrigger('autoProcessPendingStudentFinalRecords')
    .timeBased()
    .after(delayMs)
    .create();

  log_(
    'scheduleNextAutoStudentFinalRecordTrigger_',
    'OK',
    `${Math.round(delayMs / 1000)}초 뒤 다음 학생별 최종 자동 생성 예약`
  );
}

function deleteAutoStudentFinalRecordTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoProcessPendingStudentFinalRecords') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function saveAutoStudentFinalLastResult_(status, result) {
  const payload = {
    status,
    processed: result.processed || 0,
    failed: result.failed || 0,
    skipped: result.skipped || 0,
    remaining: result.remaining || 0,
    blocked: !!result.blocked,
    blockingReason: result.blockingReason || '',
    blockingMessage: result.blockingMessage || '',
    updatedAt: new Date().toISOString(),
  };

  PropertiesService.getUserProperties().setProperty(
    'AUTO_STUDENT_FINAL_LAST_RESULT',
    JSON.stringify(payload)
  );
}

function getAutoStudentFinalLastResult_() {
  const raw = PropertiesService.getUserProperties().getProperty('AUTO_STUDENT_FINAL_LAST_RESULT');
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function getIncludedAssignmentKeySet_() {
  const sh = getSheet_(SHEET_NAMES.assignments);
  const h = headerMap_(sh);
  const set = new Set();

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return set;

  for (let row = 2; row <= lastRow; row++) {
    const include = h.includeInFinal && sh.getRange(row, h.includeInFinal).getValue() === true;
    if (!include) continue;

    const courseId = valueAt_(sh, row, h.courseId);
    const courseWorkId = valueAt_(sh, row, h.courseWorkId);

    if (courseId && courseWorkId) {
      set.add(`${courseId}|${courseWorkId}`);
    }
  }

  return set;
}

function isSubmissionFromIncludedAssignment_(submissionKey, assignmentKeySet) {
  const parts = String(submissionKey || '').split('|');
  if (parts.length < 2) return false;

  const assignmentKey = `${parts[0]}|${parts[1]}`;
  return assignmentKeySet.has(assignmentKey);
}

function appendStudentFinalRecords_(rows) {
  if (!rows || rows.length === 0) return;

  const sh = getSheet_(SHEET_NAMES.studentFinalRecords);
  const startRow = Math.max(sh.getLastRow() + 1, 2);

  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function styleStudentFinalRecordsSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.studentFinalRecords);
  if (!sh) return;

  const filter = sh.getFilter();
  if (filter) filter.remove();

  const h = headerMap_(sh);
  const lastRow = getLastDataRowByHeaders_(sh, [
    'studentNo',
    'studentName',
    'courseName',
    'assignmentCount',
    'finalRecord',
    'createdAt'
  ]);
  clearRowsBelowLastData_(sh, lastRow);
  const lastCol = sh.getLastColumn();

  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);

  sh.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  setColumnWidthsByHeader_(sh, {
    studentNo: 70,
    studentName: 90,
    email: 160,
    courseName: 160,
    assignmentCount: 90,
    sourceSubmissionKeys: 180,
    sourceAssignments: 260,
    evidencePack: 420,
    finalRecord: 650,
    charCount: 85,
    caution: 260,
    teacherChecked: 90,
    createdAt: 140,
    model: 180,
    usage: 180,
    finalQueueKey: 180,
    finalStatus: 110,
    finalError: 320,
  });

  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol)
      .setVerticalAlignment('top')
      .setFontSize(10)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    ['sourceAssignments', 'evidencePack', 'finalRecord', 'caution', 'finalError'].forEach(header => {
      if (h[header]) {
        sh.getRange(2, h[header], lastRow - 1, 1)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top');
      }
    });

    ['studentNo', 'studentName', 'assignmentCount', 'charCount', 'teacherChecked', 'finalStatus'].forEach(header => {
      if (h[header]) {
        sh.getRange(2, h[header], lastRow - 1, 1)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    });

    sh.setRowHeights(2, lastRow - 1, 150);

    if (h.teacherChecked) {
      sh.getRange(2, h.teacherChecked, lastRow - 1, 1).insertCheckboxes();
    }
  }

  hideColumnsByHeaders_(sh, [
    'finalQueueKey',
    'sourceSubmissionKeys',
    'evidencePack',
    'usage'
  ]);

  if (h.finalRecord) {
    sh.getRange(1, h.finalRecord).setBackground('#274e13');
  }

  if (h.charCount) {
    sh.getRange(1, h.charCount).setBackground('#7f6000');
  }

  applyRecordConditionalFormats_(sh);

  try {
    sh.getRange(1, 1, Math.max(lastRow, 1), lastCol).createFilter();
  } catch (err) {
    // 무시
  }
}

function formatStudentFinalRecordRows_(targetRows) {
  if (!targetRows || targetRows.length === 0) return;

  const sh = getSheet_(SHEET_NAMES.studentFinalRecords);
  const h = headerMap_(sh);

  targetRows.forEach(row => {
    sh.setRowHeight(row, 150);

    sh.getRange(row, 1, 1, sh.getLastColumn())
      .setVerticalAlignment('top')
      .setFontSize(10)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    ['sourceAssignments', 'evidencePack', 'finalRecord', 'caution', 'finalError'].forEach(header => {
      if (h[header]) {
        sh.getRange(row, h[header])
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top');
      }
    });

    ['studentNo', 'studentName', 'assignmentCount', 'charCount', 'teacherChecked', 'finalStatus'].forEach(header => {
      if (h[header]) {
        sh.getRange(row, h[header])
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    });

    if (h.teacherChecked) {
      sh.getRange(row, h.teacherChecked).insertCheckboxes();
    }
  });

  hideColumnsByHeaders_(sh, [
    'finalQueueKey',
    'sourceSubmissionKeys',
    'evidencePack',
    'usage'
  ]);

  SpreadsheetApp.flush();
}

function requireCurrentAiApiKey_() {
  const config = getConfigMap_();
  const ai = getAiProviderAndModel_(config);
  const provider = normalizeAiProvider_(ai.provider);

  if (provider === 'gpt') {
    const key = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');

    if (!key) {
      throw new Error(
        '현재 AI_PROVIDER가 gpt로 설정되어 있지만 OpenAI API 키가 저장되어 있지 않습니다.\n' +
        '메뉴에서 OpenAI API 키를 저장하세요.'
      );
    }

    return {
      provider: 'gpt',
      model: ai.model,
    };
  }

  if (provider === 'gemini') {
    const key = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');

    if (!key) {
      throw new Error(
        '현재 AI_PROVIDER가 gemini로 설정되어 있지만 Gemini API 키가 저장되어 있지 않습니다.\n' +
        '메뉴에서 Gemini API 키를 저장하세요.'
      );
    }

    return {
      provider: 'gemini',
      model: ai.model,
    };
  }

  if (provider === 'claude') {
    const key = PropertiesService.getUserProperties().getProperty('CLAUDE_API_KEY');

    if (!key) {
      throw new Error(
        '현재 AI_PROVIDER가 claude로 설정되어 있지만 Claude API 키가 저장되어 있지 않습니다.\n' +
        '메뉴에서 Claude API 키를 저장하세요.'
      );
    }

    return {
      provider: 'claude',
      model: ai.model,
    };
  }

  throw new Error(`지원하지 않는 AI_PROVIDER입니다. claude, gpt, gemini 중 하나로 설정해주세요: ${provider}`);
}


function organizeSheetTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const activeSheet = ss.getActiveSheet();
  const activeRange = activeSheet ? activeSheet.getActiveRange() : null;
  const activeSheetName = activeSheet ? activeSheet.getName() : '';

  const sheetOrder = [
    DASHBOARD_SHEET_NAME,
    SHEET_NAMES.config,
    SHEET_NAMES.assignments,
    SHEET_NAMES.submissions,
    SHEET_NAMES.records,
    SHEET_NAMES.studentFinalRecords,
    SHEET_NAMES.commonPhrases,
    SHEET_NAMES.behaviorRecommendations,
    SHEET_NAMES.logs,
  ];

  sheetOrder.forEach((sheetName, index) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    sh.showSheet();
    sh.activate();
    ss.moveActiveSheet(index + 1);
  });

  const studentsSheet = ss.getSheetByName(SHEET_NAMES.students);
  if (studentsSheet) {
    studentsSheet.hideSheet();
  }

  try {
    if (activeSheetName && activeSheetName !== SHEET_NAMES.students) {
      const restoreSheet = ss.getSheetByName(activeSheetName);
      if (restoreSheet && !restoreSheet.isSheetHidden()) {
        restoreSheet.activate();

        if (activeRange) {
          restoreSheet.setActiveRange(activeRange);
        }
      }
    } else {
      const dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
      if (dashboard) dashboard.activate();
    }
  } catch (err) {
    // 시트 복귀 실패는 무시
  }
}

function formatRecordRows_(targetRows) {
  if (!targetRows || targetRows.length === 0) return;

  const sh = getSheet_(SHEET_NAMES.records);
  const h = headerMap_(sh);

  targetRows.forEach(row => {
    sh.setRowHeight(row, 145);

    // 일단 해당 행 전체는 위쪽 정렬 + 기본 글자 크기
    sh.getRange(row, 1, 1, sh.getLastColumn())
      .setVerticalAlignment('top')
      .setFontSize(10)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // 핵심 확인 열은 줄바꿈
    ['finalRecord', 'caution', 'assignmentTitle'].forEach(header => {
      if (h[header]) {
        sh.getRange(row, h[header])
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top');
      }
    });

    // finalRecord는 실제 검토용이므로 조금 더 보기 좋게
    if (h.finalRecord) {
      sh.getRange(row, h.finalRecord)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setVerticalAlignment('top')
        .setHorizontalAlignment('left');
    }

    // 숫자/체크/이름 계열은 가운데 정렬
    ['studentNo', 'studentName', 'charCount', 'teacherChecked', 'createdAt', 'model'].forEach(header => {
      if (h[header]) {
        sh.getRange(row, h[header])
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    });

    // 체크박스 다시 보장
    if (h.teacherChecked) {
      sh.getRange(row, h.teacherChecked).insertCheckboxes();
    }
  });

  SpreadsheetApp.flush();
}
