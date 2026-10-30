// 생기부 도우미의 '도구' 화면(공통문구 · 행발 추천 · 수동추가)용 서버 함수
// 생성 로직은 Code.js의 기존 함수를 그대로 쓰고, 여기서는 화면 입출력만 다룬다.

const MANUAL_DEFAULT_HEADERS = ['번호', '이름', '활동 내용', '생기부', '바이트'];
const MANUAL_ROW_LIMIT = 500;

// --- 공통문구 -------------------------------------------------------------

function uiGetCommonPhrases() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.commonPhrases);
  if (!sh) return { prompt: '', sourceText: '', count: 10, phrases: [] };

  const input = sh.getRange(2, 1, 1, 4).getValues()[0];
  const count = Number(input[3] || 0);
  const rowCount = Math.max(sh.getLastRow() - 2, getCommonPhraseTotal_(), count, 0);
  const phrases = [];

  if (rowCount > 0) {
    const range = sh.getRange(3, 2, rowCount, 1);
    const values = range.getValues();
    const notes = range.getNotes();

    values.forEach((r, i) => {
      const text = String(r[0] || '').trim();
      const note = String(notes[i][0] || '').trim();
      if (!text && !note) return;
      // 중복검사 메모는 화면에서 따로 표시하므로 오류 메모만 넘긴다.
      phrases.push({ row: i + 3, text, error: text ? '' : note });
    });
  }

  return {
    prompt: String(input[0] || ''),
    sourceText: String(input[1] || ''),
    count: count || 10,
    phrases,
  };
}

function uiStartCommonPhrases(input) {
  const prompt = String((input && input.prompt) || '').trim();
  const sourceText = String((input && input.sourceText) || '').trim();
  const count = Number(input && input.count);

  if (!prompt) throw new Error('어떤 문구를 만들지 지시사항을 입력하세요.');
  if (!sourceText) throw new Error('바탕이 되는 문구를 입력하세요.');
  if (!Number.isInteger(count) || count < 1 || count > 200) throw new Error('만들 개수는 1~200 사이의 숫자로 입력하세요.');

  const sh = ensureCommonPhraseSheet_();
  sh.getRange(2, 1, 1, 2).setValues([[prompt, sourceText]]);
  sh.getRange(2, 4).setValue(count);
  SpreadsheetApp.flush();

  const result = runWithCollectedUi_(() => generateCommonPhrases());
  return Object.assign(result, { status: uiGetStatus(), data: uiGetCommonPhrases() });
}

function uiUpdateCommonPhrase(row, expected, text) {
  return withDocumentLock_(() => {
    const sh = getSheet_(SHEET_NAMES.commonPhrases);
    row = Number(row);
    if (row < 3 || row > sh.getLastRow()) throw new Error('수정할 문구를 찾을 수 없습니다. 새로고침해 주세요.');

    const cell = sh.getRange(row, 2);
    if (String(cell.getValue() || '').trim() !== String(expected || '').trim()) {
      throw new Error('시트 내용이 바뀌어 저장하지 못했습니다. 새로고침한 뒤 다시 시도해 주세요.');
    }

    const value = String(text || '').trim();
    cell.setValue(value);
    sh.getRange(row, 3).setFormula(`=IF(B${row}="","",2*LENB(B${row})-LEN(B${row}))`);

    return { row, text: value };
  });
}

// --- 행발 추천 ------------------------------------------------------------

const BEHAVIOR_INPUT_FIELDS = ['studentName', 'selectedTraits', 'customTraits', 'positiveObservation', 'improvementNeeded', 'growthEvidence'];

function uiListBehaviorRows() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.behaviorRecommendations);
  const result = { traits: BEHAVIOR_TRAIT_KEYWORDS, rows: [] };
  if (!sh) return result;

  const h = headerMap_(sh);
  result.traits = behaviorTraitOptions_(sh, h);
  if (sh.getLastRow() < 2) return result;

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), BEHAVIOR_RECOMMENDATION_HEADERS.length);
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const notes = h.recommendedRecord ? sh.getRange(2, h.recommendedRecord, lastRow - 1, 1).getNotes() : [];

  values.forEach((r, i) => {
    const item = { row: i + 2 };
    BEHAVIOR_RECOMMENDATION_HEADERS.forEach(field => {
      item[field] = h[field] ? toClientValue_(r[h[field] - 1]) : '';
    });
    item.note = notes[i] ? String(notes[i][0] || '') : '';

    const hasAny = BEHAVIOR_INPUT_FIELDS.some(field => String(item[field] || '').trim() !== '') || String(item.recommendedRecord || '').trim() !== '';
    if (hasAny) result.rows.push(item);
  });

  return result;
}

// 시트의 드롭다운 목록을 그대로 쓴다. 드롭다운이 없으면 기본 목록을 쓴다.
function behaviorTraitOptions_(sh, h) {
  if (!h.selectedTraits) return BEHAVIOR_TRAIT_KEYWORDS;
  const rule = sh.getRange(2, h.selectedTraits).getDataValidation();
  if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return BEHAVIOR_TRAIT_KEYWORDS;
  const options = (rule.getCriteriaValues()[0] || []).map(v => String(v || '').trim()).filter(v => v && v !== '없음');
  return options.length ? options : BEHAVIOR_TRAIT_KEYWORDS;
}

function uiAddBehaviorStudents(names) {
  const list = (names || []).map(n => String(n || '').trim()).filter(Boolean);
  if (list.length === 0) throw new Error('추가할 학생 이름을 입력하세요.');

  withDocumentLock_(() => {
    const sh = ensureBehaviorRecommendationSheet_();
    const h = headerMap_(sh);
    const startRow = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(startRow, h.studentName, list.length, 1).setValues(list.map(n => [n]));
  });

  styleBehaviorRecommendationSheet_();
  log_('uiAddBehaviorStudents', 'OK', `행발문구추천 시트에 ${list.length}명 추가`);
  return uiListBehaviorRows();
}

function uiSaveBehaviorRow(row, expectedName, data) {
  return withDocumentLock_(() => {
    const sh = getSheet_(SHEET_NAMES.behaviorRecommendations);
    const h = assertBehaviorRow_(sh, row, expectedName);
    const values = Object.assign({}, data);

    // 드롭다운에 없는 특성은 '목록에 없는 특성' 칸으로 옮긴다.
    if (values.selectedTraits !== undefined) {
      const allowed = behaviorTraitOptions_(sh, h);
      const picked = String(values.selectedTraits || '').split(',').map(t => t.trim()).filter(Boolean);
      const extra = picked.filter(t => !allowed.includes(t));
      values.selectedTraits = picked.filter(t => allowed.includes(t)).join(', ');
      if (extra.length) {
        const custom = String(values.customTraits !== undefined ? values.customTraits : sh.getRange(row, h.customTraits).getValue() || '')
          .split(',').map(t => t.trim()).filter(Boolean);
        values.customTraits = custom.concat(extra.filter(t => !custom.includes(t))).join(', ');
      }
    }

    BEHAVIOR_INPUT_FIELDS.forEach(field => {
      if (values[field] !== undefined && h[field]) {
        try {
          sh.getRange(row, h[field]).setValue(String(values[field]).trim());
        } catch (err) {
          if (field === 'selectedTraits' && String(values.selectedTraits).includes(',')) {
            throw new Error('특성 칸 드롭다운이 여러 개 선택을 허용하지 않아요. 시트의 selectedTraits 열 → 데이터 확인에서 "여러 항목 선택 허용"을 켜거나, 특성을 하나만 골라 주세요.');
          }
          throw err;
        }
      }
    });

    return readBehaviorRow_(sh, h, row);
  });
}

function uiUpdateBehaviorRecord(row, expectedName, text) {
  return withDocumentLock_(() => {
    const sh = getSheet_(SHEET_NAMES.behaviorRecommendations);
    const h = assertBehaviorRow_(sh, row, expectedName);
    const value = String(text || '').trim();

    sh.getRange(row, h.recommendedRecord).setValue(value);
    sh.getRange(row, h.byteCount).setValue(value ? neisByteCount_(value) : '');

    return readBehaviorRow_(sh, h, row);
  });
}

// 결과를 지우면 다음 [추천 문구 만들기] 때 다시 만든다.
function uiResetBehaviorResult(row, expectedName) {
  return withDocumentLock_(() => {
    const sh = getSheet_(SHEET_NAMES.behaviorRecommendations);
    const h = assertBehaviorRow_(sh, row, expectedName);

    sh.getRange(row, h.recommendedRecord, 1, 3).clearContent();
    sh.getRange(row, h.recommendedRecord).clearNote();

    return readBehaviorRow_(sh, h, row);
  });
}

function uiDeleteBehaviorRow(row, expectedName) {
  withDocumentLock_(() => {
    const sh = getSheet_(SHEET_NAMES.behaviorRecommendations);
    assertBehaviorRow_(sh, row, expectedName);
    sh.deleteRow(row);
  });

  return uiListBehaviorRows();
}

function uiStartBehavior() {
  const result = runWithCollectedUi_(() => generateBehaviorRecommendationsCore_());
  return Object.assign(result, { status: uiGetStatus(), data: uiListBehaviorRows() });
}

function assertBehaviorRow_(sh, row, expectedName) {
  const h = headerMap_(sh);
  const missing = BEHAVIOR_RECOMMENDATION_HEADERS.filter(header => !h[header]);
  if (missing.length) throw new Error(`행발문구추천 시트에 필요한 열이 없습니다: ${missing.join(', ')}`);

  row = Number(row);
  if (row < 2 || row > sh.getLastRow()) throw new Error('학생 행을 찾을 수 없습니다. 새로고침해 주세요.');

  const name = String(sh.getRange(row, h.studentName).getValue() || '').trim();
  if (name !== String(expectedName || '').trim()) {
    throw new Error('시트 내용이 바뀌어 저장하지 못했습니다. 새로고침한 뒤 다시 시도해 주세요.');
  }

  return h;
}

function readBehaviorRow_(sh, h, row) {
  SpreadsheetApp.flush();
  const lastCol = Math.max(sh.getLastColumn(), BEHAVIOR_RECOMMENDATION_HEADERS.length);
  const values = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const item = { row };

  BEHAVIOR_RECOMMENDATION_HEADERS.forEach(field => {
    item[field] = toClientValue_(values[h[field] - 1]);
  });
  item.note = String(sh.getRange(row, h.recommendedRecord).getNote() || '');

  return item;
}

// --- 수동추가 -------------------------------------------------------------

function uiGetManual() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANUAL_SHEET_NAME);
  if (!sh) return { exists: false };

  const config = getConfigMap_();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headerValues = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const headers = headerValues.map((name, i) => ({
    col: i + 1,
    letter: colToA1_(i + 1),
    name: String(name || '').trim(),
  }));

  let manualConfig = null;
  let configError = '';
  try {
    manualConfig = getManualRecordConfig_(config);
  } catch (err) {
    configError = String(err.message || err);
  }

  const result = {
    exists: true,
    headers,
    configured: !!manualConfig,
    configError,
    startRow: manualConfig ? manualConfig.startRow : Math.max(Number(config.MANUAL_START_ROW || 2), 1),
    inputCols: manualConfig ? manualConfig.inputCols : [],
    outputCol: manualConfig ? manualConfig.outputCol : 0,
    rows: [],
    truncated: false,
  };

  if (!manualConfig) return result;

  const lastRow = sh.getLastRow();
  if (lastRow < manualConfig.startRow) return result;

  const rowCount = Math.min(lastRow - manualConfig.startRow + 1, MANUAL_ROW_LIMIT);
  const readCols = Math.max(sh.getLastColumn(), manualConfig.byteCol);
  const fullHeaders = sh.getRange(1, 1, 1, readCols).getValues()[0];
  const values = sh.getRange(manualConfig.startRow, 1, rowCount, readCols).getValues();
  const notes = sh.getRange(manualConfig.startRow, manualConfig.outputCol, rowCount, 1).getNotes();
  // 입력·출력 열이 아닌 앞쪽 열(번호, 이름 등)을 행 구분용으로 보여준다.
  const labelCols = [];
  for (let col = 1; col <= readCols && labelCols.length < 2; col++) {
    if (!manualConfig.inputCols.includes(col) && col !== manualConfig.outputCol && col !== manualConfig.byteCol) labelCols.push(col);
  }

  values.forEach((r, i) => {
    const input = buildManualInputFromRowValues_(r, fullHeaders, manualConfig.inputCols);
    const output = String(rowValueAt_(r, manualConfig.outputCol) || '').trim();
    if (!input && !output) return;

    const note = String(notes[i][0] || '').trim();
    result.rows.push({
      row: manualConfig.startRow + i,
      label: labelCols.map(col => String(rowValueAt_(r, col) || '').trim()).filter(Boolean).join(' '),
      input,
      output,
      note,
      status: manualRowStatus_(input, output, note),
    });
  });

  result.truncated = lastRow - manualConfig.startRow + 1 > MANUAL_ROW_LIMIT;
  return result;
}

function manualRowStatus_(input, output, note) {
  if (output) return '완료';
  if (/^\[수동추가 오류\]/.test(note)) return '오류';
  if (/^\[수동추가 근거부족\]/.test(note)) return '근거부족';
  if (input.length < 20) return '내용 부족';
  return '대기';
}

function uiCreateManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(MANUAL_SHEET_NAME)) throw new Error(`'${MANUAL_SHEET_NAME}' 시트가 이미 있습니다.`);

  const sh = ss.insertSheet(MANUAL_SHEET_NAME);
  sh.getRange(1, 1, 1, MANUAL_DEFAULT_HEADERS.length)
    .setValues([MANUAL_DEFAULT_HEADERS])
    .setFontWeight('bold')
    .setBackground('#eaf4ff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(3, 420);
  sh.setColumnWidth(4, 420);

  writeConfigValues_({ MANUAL_START_ROW: '2', MANUAL_INPUT_COLS: 'C', MANUAL_OUTPUT_COL: 'D' });
  log_('uiCreateManualSheet', 'OK', '수동추가 시트를 기본 열 구성으로 만듦');

  return uiGetManual();
}

function uiSaveManualSetup(setup) {
  const startRow = Number(setup && setup.startRow);
  const inputCols = ((setup && setup.inputCols) || []).map(Number).filter(n => n > 0);
  const outputCol = Number(setup && setup.outputCol);

  if (!Number.isInteger(startRow) || startRow < 2) throw new Error('시작 행은 2 이상의 숫자로 입력하세요.');
  if (inputCols.length === 0) throw new Error('AI에게 보낼 열을 하나 이상 고르세요.');
  if (!outputCol) throw new Error('생기부를 적을 열을 고르세요.');

  const patch = {
    MANUAL_START_ROW: String(startRow),
    MANUAL_INPUT_COLS: inputCols.map(colToA1_).join(','),
    MANUAL_OUTPUT_COL: colToA1_(outputCol),
  };

  // 기존 검사 로직으로 열 구성이 올바른지 먼저 확인한다.
  getManualRecordConfig_(patch);

  writeConfigValues_(patch);
  log_('uiSaveManualSetup', 'OK', `수동추가 열 설정: 입력 ${patch.MANUAL_INPUT_COLS}, 출력 ${patch.MANUAL_OUTPUT_COL}, 시작 ${startRow}행`);

  return uiGetManual();
}

function uiAddManualRow(valuesByCol) {
  withDocumentLock_(() => {
    const sh = getSheet_(MANUAL_SHEET_NAME);
    const cols = Object.keys(valuesByCol || {}).map(Number).filter(n => n > 0);
    if (cols.length === 0 || cols.every(col => !String(valuesByCol[col] || '').trim())) {
      throw new Error('추가할 내용을 입력하세요.');
    }

    const width = Math.max.apply(null, cols);
    const rowValues = [];
    for (let col = 1; col <= width; col++) rowValues.push(String(valuesByCol[col] || '').trim());

    const config = getManualRecordConfig_(getConfigMap_());
    const row = Math.max(sh.getLastRow() + 1, config.startRow);
    sh.getRange(row, 1, 1, width).setValues([rowValues]);
  });

  return uiGetManual();
}

function uiUpdateManualOutput(row, expectedInput, text) {
  return withDocumentLock_(() => {
    const ctx = assertManualRow_(row, expectedInput);
    const value = String(text || '').trim();

    ctx.sh.getRange(row, ctx.config.outputCol).setValue(value);
    if (value) setManualByteFormula_(ctx.sh, row, ctx.config.outputCol);

    return { row: Number(row), output: value };
  });
}

// 결과와 메모를 지우면 다음 [생기부 만들기] 때 다시 만든다.
function uiResetManualOutput(row, expectedInput) {
  return withDocumentLock_(() => {
    const ctx = assertManualRow_(row, expectedInput);
    const cell = ctx.sh.getRange(row, ctx.config.outputCol);
    cell.clearContent();
    cell.clearNote();
    ctx.sh.getRange(row, ctx.config.byteCol).clearContent();

    return { row: Number(row), output: '', note: '' };
  });
}

function uiStartManual() {
  const result = runWithCollectedUi_(() => generateManualAddedRecords());
  return Object.assign(result, { status: uiGetStatus(), data: uiGetManual() });
}

// 화면을 연 뒤 행이 바뀌었을 수 있으므로, AI 입력 내용이 같을 때만 쓴다.
function assertManualRow_(row, expectedInput) {
  const sh = getSheet_(MANUAL_SHEET_NAME);
  const config = getManualRecordConfig_(getConfigMap_());
  row = Number(row);

  if (row < config.startRow || row > sh.getLastRow()) throw new Error('행을 찾을 수 없습니다. 새로고침해 주세요.');

  const readCols = Math.max(sh.getLastColumn(), config.byteCol);
  const headers = sh.getRange(1, 1, 1, readCols).getValues()[0];
  const values = sh.getRange(row, 1, 1, readCols).getValues()[0];

  if (buildManualInputFromRowValues_(values, headers, config.inputCols) !== String(expectedInput || '')) {
    throw new Error('시트 내용이 바뀌어 저장하지 못했습니다. 새로고침한 뒤 다시 시도해 주세요.');
  }

  return { sh, config };
}

// --- 공용 -----------------------------------------------------------------

function withDocumentLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
