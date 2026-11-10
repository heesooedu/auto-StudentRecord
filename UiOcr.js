// 손글씨 추출: 종이로 받은 과제를 스캔한 PDF에서 학생별 손글씨를 추출해 '수동추가'로 넘긴다.
// PDF는 화면(브라우저)에서 학생별 조각으로 나눠 한 명씩 보낸다. 서버는 조각 하나를 AI에 보내고 결과를 시트에 적는다.

const OCR_SHEET_NAME = '손글씨추출';
// 시트 열: 쪽, 학번, 이름, [추출 지침의 항목…], 확인할 점, 모델, 추출 시각
const OCR_FIXED_HEAD = ['쪽', '학번', '이름'];
const OCR_FIXED_TAIL = ['확인할 점', '모델', '추출 시각'];
// 표의 열은 추출 지침에서 AI가 읽어 정한다. 지침이 바뀌었는지는 지침의 해시로 안다.
const OCR_ITEMS_PROPERTY = 'OCR_ITEMS';
const OCR_ITEMS_HASH_PROPERTY = 'OCR_ITEMS_PROMPT_HASH';
const OCR_MAX_ITEMS = 20;
// 기본 지침의 이 표시가 남아 있으면 아직 과제에 맞게 고치지 않은 것이다.
const OCR_PROMPT_PLACEHOLDER = '(여기에';
// 학번·이름은 늘 따로 뽑으므로 열로 두지 않는다. 수동추가 시트에서 쓰는 열 이름과도 겹치면 안 된다.
const OCR_RESERVED_NAMES = OCR_FIXED_HEAD.concat(OCR_FIXED_TAIL, ['생기부', '바이트', '번호', '성명', '학생 이름', '학생이름']);
const OCR_MODEL_PROPERTY_PREFIX = 'OCR_MODEL_';
const OCR_REASONING_PROPERTY = 'OCR_REASONING';
// 학생 한 명 조각의 최대 크기. 이보다 크면 스캔 해상도를 낮추라고 안내한다.
const OCR_MAX_PDF_BYTES = 10 * 1024 * 1024;
const OCR_FAILED_PREFIX = '[읽기 실패]';

// 회사별로 싼 모델부터 적는다. 화면은 이 순서대로 보여 주고, 처음에는 첫 모델을 고른다.
const OCR_MODEL_CHOICES = {
  claude: [
    { id: 'claude-sonnet-5-5', note: '손글씨를 잘 읽어요. 이걸로 쓰세요.' },
    { id: 'claude-haiku-5-5', note: '싸지만 손글씨를 잘 못 읽어요.' },
  ],
  gpt: [
    { id: 'gpt-6.1-terra', note: '손글씨를 잘 읽어요. 이걸로 쓰세요.' },
    { id: 'gpt-6-luna', note: '싸지만 손글씨를 잘 못 읽어요.' },
  ],
  gemini: [
    { id: 'gemini-3.8-flash', note: '가장 새 모델이에요.' },
    { id: 'gemini-3.7-flash', note: '' },
    { id: 'gemini-3.6-flash', note: '' },
    { id: 'gemini-3.5-flash', note: '' },
  ],
};

const OCR_RESPONSE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    student_id: { type: 'string' },
    name: { type: 'string' },
    answers: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          item: { type: 'number' },
          text: { type: 'string' },
        },
        required: ['item', 'text'],
      },
    },
    unreadable: { type: 'string' },
  },
  required: ['student_id', 'name', 'answers', 'unreadable'],
};

function defaultOcrPrompt_() {
  return [
    '[추출 내용 및 방법]',
    '"학번", "이름", (여기에 이번 과제에서 추출할 항목을 하나씩 큰따옴표로 감싸고 쉼표로 나눠 적으세요. 문항은 줄이지 말고 통째로 적어야 생기부를 쓸 때 AI가 무엇을 물었는지 알 수 있어요. 예: "악곡명", "Q1 이 곡을 듣고 떠오른 장면, 느낌을 쓰시오")을 추출한다.',
    OCR_TABLE_RULE,
    '',
    '[옮겨 적기 원칙]',
    '1. 학생이 손으로 쓴 글만 옮긴다.',
    '2. 보이는 그대로 옮긴다. 맞춤법을 고치거나, 요약하거나, 없는 내용을 보태지 않는다.',
    '3. 읽을 수 없는 글자나 낱말은 [?]로 적는다. 짐작이 가더라도 확실하지 않으면 [?]로 적는다.',
    '4. 줄을 바꾼 곳과 문단은 줄바꿈으로 살린다. 줄을 그어 지운 글은 옮기지 않는다.',
  ].join('\n');
}

function ocrPromptNeedsEdit_(prompt) {
  return String(prompt || '').indexOf(OCR_PROMPT_PLACEHOLDER) >= 0;
}

function ocrPromptHash_(prompt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(prompt || ''), Utilities.Charset.UTF_8));
}

function ocrPrompt_(config) {
  return getConfigPromptTemplate_(config || getConfigMap_(), 'OCR_PROMPT', defaultOcrPrompt_());
}

const OCR_TABLE_RULE = '이때 스프레드시트로 저장할 수 있는 표 형태로 추출한다.';

// 선생님이 지침에서 표 형태 문장을 지웠어도 AI에게는 늘 붙여 보낸다.
function ocrPromptWithTableRule_(prompt) {
  const text = String(prompt || '');
  if (text.replace(/\s+/g, '').indexOf('스프레드시트로저장할수있는표') >= 0) return text;
  return `${text.trim()}\n\n${OCR_TABLE_RULE}`;
}

function ocrSystemGuide_(prompt, items) {
  const list = items.map((it, i) => `${i + 1}. ${it.name}${it.hint ? ` — ${it.hint}` : ''}`);
  const one = items.length === 1;
  return [
    '# 역할',
    '너는 학생이 손으로 쓴 과제 스캔본을 글자로 옮겨 적는 사람이다. 첨부한 PDF는 학생 한 명의 과제다.',
    '',
    ocrPromptWithTableRule_(prompt),
    '',
    '# 표의 열 (추출 지침에서 읽은 항목. 학번·이름은 따로 적는다)',
    ...list,
    one
      ? '- 열이 하나뿐이므로 학생이 쓴 글 전체를 1번에 담는다.'
      : '- 학생의 글을 열마다 나눠 담는다. 그 항목에 해당하는 글만 넣고, 답이 없으면 빈 문자열로 둔다.',
    ...(one ? [] : ['- 위 열에 들어가지 않는 글은 옮기지 않는다. 어느 항목의 글인지 애매하면 가장 가까운 열에 넣고 unreadable에 적는다.']),
    '',
    '# 출력',
    '- student_id: 종이에 적힌 학번(또는 번호). 없으면 빈 문자열.',
    '- name: 종이에 적힌 학생 이름. 없으면 빈 문자열.',
    `- answers: 열마다 하나씩 {"item": 열 번호(1~${items.length}), "text": 옮겨 적은 글}.`,
    '- unreadable: 읽기 어려웠던 부분, 흐리거나 잘린 곳, 학생이 둘 이상 섞여 보이는 경우 등 교사가 원본과 비교해 봐야 할 점을 짧게. 없으면 빈 문자열.',
    '- 다른 말 없이 JSON 객체 하나만 출력한다: {"student_id":"","name":"","answers":[{"item":1,"text":""}],"unreadable":""}',
  ].join('\n');
}

// 추출 지침에서 읽어 둔 열 목록. 이름이 곧 손글씨추출·수동추가 시트의 열 이름이 된다.
function ocrItems_() {
  try {
    const saved = JSON.parse(PropertiesService.getDocumentProperties().getProperty(OCR_ITEMS_PROPERTY) || 'null');
    if (Array.isArray(saved)) return saved.filter(it => it && it.name).map(it => ({ name: String(it.name) }));
  } catch (err) {
    // 저장값이 망가졌으면 다시 읽게 한다.
  }
  return [];
}

function ocrItemsFresh_(prompt) {
  return PropertiesService.getDocumentProperties().getProperty(OCR_ITEMS_HASH_PROPERTY) === ocrPromptHash_(prompt) && ocrItems_().length > 0;
}

// AI가 돌려준 열 이름을 정리한다. 학번·이름 같은 고정 열과 겹치는 이름, 중복은 뺀다.
// 문항 전문을 열 이름으로 쓸 수 있게 넉넉히 둔다(생기부 AI에게 열 이름이 문항 맥락으로 함께 간다).
const OCR_MAX_COLUMN_NAME = 200;

function cleanOcrColumnNames_(names) {
  const seen = {};
  const out = [];
  (Array.isArray(names) ? names : []).forEach(n => {
    const name = String(n || '').replace(/\s+/g, ' ').trim()
      .replace(/^["'“”‘’「」『』]+|["'“”‘’「」『』]+$/g, '').trim().slice(0, OCR_MAX_COLUMN_NAME);
    if (!name || OCR_RESERVED_NAMES.includes(name) || seen[name]) return;
    seen[name] = true;
    out.push({ name });
  });
  return out.slice(0, OCR_MAX_ITEMS);
}

const OCR_COLUMNS_RESPONSE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: { columns: { type: 'array', items: { type: 'string' } } },
  required: ['columns'],
};

function ocrColumnsGuide_() {
  return [
    '# 역할',
    '너는 교사가 쓴 "손글씨 추출 지침"을 읽고, 학생 과제를 표로 옮길 때 쓸 열 이름을 정하는 사람이다.',
    '',
    '# 규칙',
    '- 지침에서 추출하라고 적은 항목을 적힌 순서대로 열 이름으로 쓴다.',
    '- 열 이름은 지침에 적힌 낱말을 그대로 쓴다. 고치거나, 줄이거나, 합치거나, 지침에 없는 열을 새로 만들지 않는다.',
    '- 항목이 문항 전체(질문 문장)로 적혀 있으면 문항 전체를 한 글자도 빼지 않고 열 이름 하나로 쓴다.',
    '- 따옴표("", “”, \'\')로 감싼 부분은 안에 쉼표가 있어도 열 하나다. 열 이름에는 바깥 따옴표를 빼고 쓴다.',
    '- 따옴표가 없어도 쉼표 앞뒤가 한 문항의 이어진 문장이면 나누지 않는다. 쉼표는 항목과 항목 사이를 나눌 때만 열을 나눈다.',
    '- 학번·이름(번호·성명 등)은 따로 적으므로 열에 넣지 않는다.',
    '- 옮겨 적기 원칙, 표 형태 같은 방법 설명은 열이 아니다.',
    '',
    '# 출력',
    '- 다른 말 없이 JSON 객체 하나만 출력한다: {"columns":["항목1","항목2"]}',
  ].join('\n');
}

// [추출 내용 및 방법]의 "…을/를 추출한다" 앞부분이 모두 따옴표로 감싼 항목이면 그 목록을, 아니면 null을 돌려준다.
function quotedOcrColumns_(prompt) {
  const text = String(prompt || '');
  const at = text.indexOf('[추출 내용 및 방법]');
  let section = at >= 0 ? text.slice(at + '[추출 내용 및 방법]'.length) : text;
  const next = section.search(/\n\s*\[/);
  if (next >= 0) section = section.slice(0, next);
  const end = section.indexOf('추출한다');
  if (end < 0) return null;
  const head = section.slice(0, end);
  const names = [];
  const rest = head.replace(/"([^"]*)"|“([^”]*)”/g, (m, a, b) => { names.push(a !== undefined ? a : b); return ','; });
  if (!names.length || !/^[\s,、·및과와을를]*$/.test(rest)) return null;
  return names;
}

// 추출 전에 부른다. 지침이 그대로면 읽어 둔 열을 쓰고, 바뀌었으면 AI에게 지침만 보내 열을 다시 읽는다.
function uiOcrPrepareColumns(input) {
  const s = ocrSettings_();
  const prompt = ocrPrompt_(s.config);
  if (ocrPromptNeedsEdit_(prompt)) {
    throw new Error('추출 지침의 [추출 내용 및 방법]을 이번 과제에 맞게 먼저 고쳐 주세요. 적은 항목이 표의 열이 돼요.');
  }
  if (ocrItemsFresh_(prompt)) return { items: ocrItems_(), prompt, changed: false };

  const model = s.models.some(c => c.id === (input && input.model)) ? input.model : s.model;
  // 항목을 모두 큰따옴표로 감쌌으면 AI 없이 그대로 열로 쓴다. 하나라도 빠졌으면 AI에게 맡긴다.
  const quoted = quotedOcrColumns_(prompt);
  let items, source;
  if (quoted) {
    items = cleanOcrColumnNames_(quoted);
    source = '따옴표';
  } else {
    const result = callAi_(s.provider, model, ocrColumnsGuide_(), `다음 지침에서 표의 열 이름을 뽑아 JSON으로 답하세요.\n\n<지침>\n${prompt}\n</지침>`, 'none', OCR_COLUMNS_RESPONSE_SCHEMA, { maxTokens: 1000 });
    items = cleanOcrColumnNames_(parseOcrJson_(result.text).columns);
    source = `${s.provider}:${model}`;
  }
  if (!items.length) throw new Error('추출 지침에서 추출할 항목을 찾지 못했어요. [추출 내용 및 방법]에 "학번, 이름, 악곡명, 느낀 점을 추출한다"처럼 항목을 쉼표로 나눠 적어 주세요.');

  return withDocumentLock_(() => {
    const sh = getOcrSheet_(false);
    const before = ocrItems_().map(it => it.name).join('\u0001');
    const after = items.map(it => it.name).join('\u0001');
    if (before !== after && readOcrRows_(sh).length) {
      throw new Error(`추출 지침의 항목(${items.map(it => it.name).join(' · ')})이 이미 추출한 결과의 열과 달라요. [결과 모두 지우기]를 누른 뒤 다시 추출하거나, 지침의 항목을 예전처럼 되돌려 주세요.`);
    }
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(OCR_ITEMS_PROPERTY, JSON.stringify(items));
    props.setProperty(OCR_ITEMS_HASH_PROPERTY, ocrPromptHash_(prompt));
    log_('uiOcrPrepareColumns', 'OK', `추출 지침에서 열 ${items.length}개를 읽음: ${items.map(it => it.name).join(', ')} (${source})`);
    return { items, prompt, changed: before !== after };
  });
}

function ocrHeaders_(items) {
  return OCR_FIXED_HEAD.concat(items.map(it => it.name), OCR_FIXED_TAIL);
}

function readOcrHeader_(sh) {
  if (!sh || sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
}

// 시트 머리글에서 항목 열 이름(이름 열과 확인할 점 열 사이)을 읽는다. 옛 시트(학번·이름·내용)도 그대로 읽힌다.
function ocrSheetItemNames_(header) {
  const from = header.indexOf('이름');
  const to = header.indexOf('확인할 점');
  if (from < 0 || to <= from) return null;
  return header.slice(from + 1, to);
}

function ocrLayoutMatches_(sh, items) {
  const names = ocrSheetItemNames_(readOcrHeader_(sh));
  return !!names && names.join('\u0001') === items.map(it => it.name).join('\u0001');
}

function ocrMaxTokens_(reasoning) {
  reasoning = normalizeReasoning_(reasoning);
  if (reasoning === 'medium') return 12000;
  if (reasoning === 'low') return 8000;
  return 4000;
}

function ocrSettings_() {
  const config = getConfigMap_();
  const provider = getAiProviderAndModel_(config).provider;
  const choices = OCR_MODEL_CHOICES[provider] || [];
  const props = PropertiesService.getUserProperties();
  const savedModel = props.getProperty(OCR_MODEL_PROPERTY_PREFIX + provider) || '';
  const savedReasoning = props.getProperty(OCR_REASONING_PROPERTY) || 'none';

  return {
    config,
    provider,
    models: choices,
    model: choices.some(c => c.id === savedModel) ? savedModel : (choices[0] ? choices[0].id : ''),
    reasoning: CONFIG_VALIDATION_OPTIONS.REASONING.includes(savedReasoning) ? savedReasoning : 'none',
  };
}

function getOcrSheet_(create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(OCR_SHEET_NAME);
  if (!create) return sh;

  const items = ocrItems_();
  if (sh) {
    if (ocrLayoutMatches_(sh, items)) return sh;
    // 결과가 남은 시트는 열을 바꾸지 않는다. 화면에서 결과를 지우게 한다.
    if (sh.getLastRow() >= 2) throw new Error('손글씨추출 시트의 열이 추출 지침의 항목과 달라요. [결과 모두 지우기]를 누른 뒤 다시 추출해 주세요.');
    ss.deleteSheet(sh);
  }

  const headers = ocrHeaders_(items);
  sh = ss.insertSheet(OCR_SHEET_NAME);
  hideNewWorkSheet_(sh);
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#eaf4ff');
  // '3-4' 같은 쪽 번호와 학번 앞자리 0이 날짜·숫자로 바뀌지 않게 텍스트로 둔다.
  sh.getRange(1, 1, sh.getMaxRows(), 3).setNumberFormat('@');
  sh.setFrozenRows(1);
  items.forEach((_, i) => sh.setColumnWidth(OCR_FIXED_HEAD.length + i + 1, 360));
  sh.setColumnWidth(OCR_FIXED_HEAD.length + items.length + 1, 220);
  return sh;
}

// 행마다 answers({문항 이름: 글})를 담아 돌려준다. 열 위치는 시트 머리글에서 읽는다.
function readOcrRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const header = readOcrHeader_(sh);
  const names = ocrSheetItemNames_(header) || [];
  const col = name => header.indexOf(name);
  const at = { pages: col('쪽'), studentId: col('학번'), name: col('이름'), note: col('확인할 점'), model: col('모델') };
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, header.length).getValues();
  const text = (r, i) => (i < 0 ? '' : String(r[i] == null ? '' : r[i]));
  const rows = [];
  values.forEach((r, i) => {
    const pages = text(r, at.pages).trim();
    if (!pages) return;
    const answers = {};
    names.forEach(n => { answers[n] = text(r, col(n)); });
    const note = text(r, at.note);
    rows.push({
      row: i + 2,
      pages,
      studentId: text(r, at.studentId),
      name: text(r, at.name),
      answers,
      filled: names.some(n => answers[n].trim()),
      note,
      failed: note.indexOf(OCR_FAILED_PREFIX) === 0,
      model: text(r, at.model),
    });
  });
  return rows;
}

function uiOcrGet() {
  const s = ocrSettings_();
  const items = ocrItems_();
  const sh = getOcrSheet_(false);
  const rows = readOcrRows_(sh);
  const prompt = ocrPrompt_(s.config);
  return {
    items,
    // 화면이 지금 고치는 지침과 비교해, 열을 다시 읽어야 하는지 알 수 있게 한다.
    columnsPrompt: ocrItemsFresh_(prompt) ? prompt : '',
    // 결과가 남은 시트의 열이 지금 문항과 다르면 추출·저장·보내기를 막는다.
    layoutMismatch: rows.length > 0 && !ocrLayoutMatches_(sh, items),
    provider: s.provider,
    models: s.models,
    model: s.model,
    reasoning: s.reasoning,
    reasoningOptions: CONFIG_VALIDATION_OPTIONS.REASONING,
    hasKey: !!PropertiesService.getUserProperties().getProperty(AI_API_ENDPOINTS[s.provider].keyProperty),
    rows,
    manualExists: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANUAL_SHEET_NAME),
  };
}

function uiOcrSaveSettings(input) {
  const s = ocrSettings_();
  const model = String((input && input.model) || '');
  const reasoning = String((input && input.reasoning) || '');
  const props = PropertiesService.getUserProperties();

  if (model && s.models.some(c => c.id === model)) props.setProperty(OCR_MODEL_PROPERTY_PREFIX + s.provider, model);
  if (CONFIG_VALIDATION_OPTIONS.REASONING.includes(reasoning)) props.setProperty(OCR_REASONING_PROPERTY, reasoning);
  return { model: model || s.model, reasoning: reasoning || s.reasoning };
}

// 학생 한 명 조각(input.data: PDF base64)을 AI에 보내 글씨를 추출하고, 같은 쪽의 행이 있으면 바꾸고 없으면 아래에 붙인다.
function uiOcrExtract(input) {
  const pages = String((input && input.pages) || '').trim();
  const data = String((input && input.data) || '');
  if (!/^\d+(-\d+)?$/.test(pages)) throw new Error('쪽 번호가 올바르지 않습니다.');
  if (!data) throw new Error('보낼 PDF가 없습니다.');
  if (Math.floor(data.length * 3 / 4) > OCR_MAX_PDF_BYTES) {
    throw new Error(`${pages}쪽 파일이 너무 커요(학생 한 명에 10MB까지). 스캔 해상도를 낮추거나(예: 200dpi) 흑백으로 다시 스캔해 주세요.`);
  }

  const s = ocrSettings_();
  const model = s.models.some(c => c.id === input.model) ? input.model : s.model;
  const reasoning = CONFIG_VALIDATION_OPTIONS.REASONING.includes(input.reasoning) ? input.reasoning : s.reasoning;
  const prompt = ocrPrompt_(s.config);
  if (!ocrItemsFresh_(prompt)) throw new Error('추출 지침이 바뀌어 표의 열을 다시 읽어야 해요. 추출 버튼을 다시 눌러 주세요.');
  const items = ocrItems_();
  // 열이 어긋난 시트라면 AI를 부르기 전에 멈춘다.
  const current = getOcrSheet_(false);
  if (current && readOcrRows_(current).length && !ocrLayoutMatches_(current, items)) {
    throw new Error('손글씨추출 시트의 열이 추출 지침의 항목과 달라요. [결과 모두 지우기]를 누른 뒤 다시 추출해 주세요.');
  }
  const label = `${s.provider}:${model}${reasoning === 'none' ? '' : ` (${reasoning})`}`;

  let record;
  try {
    const result = callAi_(s.provider, model, ocrSystemGuide_(prompt, items), '첨부한 PDF에서 학생이 손으로 쓴 글을 원칙에 따라 옮겨 적어 JSON으로 답하세요.', reasoning, OCR_RESPONSE_SCHEMA, {
      attachments: [{ mimeType: 'application/pdf', data, name: `pages-${pages}.pdf` }],
      maxTokens: ocrMaxTokens_(reasoning),
    });
    const json = parseOcrJson_(result.text);
    const texts = items.map(() => []);
    (Array.isArray(json.answers) ? json.answers : []).forEach(a => {
      const i = Math.round(Number(a && a.item)) - 1;
      const t = String((a && a.text) || '').trim();
      if (i >= 0 && i < items.length && t) texts[i].push(t);
    });
    record = {
      studentId: String(json.student_id || '').trim(),
      name: String(json.name || '').trim(),
      answers: texts.map(list => list.join('\n\n')),
      note: String(json.unreadable || '').trim(),
    };
    if (!record.answers.some(Boolean)) record.note = `${OCR_FAILED_PREFIX} 학생이 쓴 글을 찾지 못했어요. ${record.note}`.trim();
    log_('uiOcrExtract', 'OK', `${pages}쪽 추출 (${label})`);
  } catch (err) {
    const message = String(err.message || err).replace(/^Exception:\s*/, '');
    record = { studentId: '', name: '', answers: items.map(() => ''), note: `${OCR_FAILED_PREFIX} ${message.slice(0, 300)}` };
    log_('uiOcrExtract', 'ERROR', `${pages}쪽: ${message.slice(0, 500)}`);
  }

  return withDocumentLock_(() => {
    const sh = getOcrSheet_(true);
    const existing = readOcrRows_(sh).find(r => r.pages === pages);
    const row = existing ? existing.row : Math.max(sh.getLastRow() + 1, 2);
    const line = [pages, record.studentId, record.name].concat(record.answers, [record.note, label, new Date()]);
    sh.getRange(row, 1, 1, line.length).setValues([line]);
    return readOcrRows_(sh).find(r => r.row === row);
  });
}

const OCR_SPLIT_RESPONSE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    pages: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          page: { type: 'number' },
          new_student: { type: 'boolean' },
          student_id: { type: 'string' },
          name: { type: 'string' },
        },
        required: ['page', 'new_student', 'student_id', 'name'],
      },
    },
  },
  required: ['pages'],
};

function ocrSplitGuide_(count) {
  return [
    '# 역할',
    `너는 반 학생들의 손글씨 과제를 한꺼번에 스캔한 PDF를 학생별로 나누는 사람이다. 첨부한 PDF는 그중 연속된 ${count}쪽이다.`,
    '',
    '# 판단 기준',
    '- 쪽마다 새 학생의 과제가 시작되는 쪽인지(new_student) 판단한다.',
    '- 이름이나 학번이 새로 적혀 있거나, 과제 양식의 첫 쪽(제목, 1번 문항부터 시작)이면 true다.',
    '- 앞 쪽의 글이 이어지는 쪽, 같은 학생의 뒷면·덧붙인 종이면 false다.',
    '- 그 쪽에 적힌 학번(번호)과 이름을 student_id, name에 적는다. 없거나 읽을 수 없으면 빈 문자열.',
    '',
    '# 출력',
    `- pages에 1쪽부터 ${count}쪽까지 한 쪽에 하나씩, 첨부한 PDF 안의 쪽 번호(page: 1~${count}) 순서대로 적는다.`,
    '- 다른 말 없이 JSON 객체 하나만 출력한다: {"pages":[{"page":1,"new_student":true,"student_id":"","name":""}]}',
  ].join('\n');
}

// 연속된 몇 쪽(input.data: PDF base64, input.firstPage: 원래 PDF에서 첫 쪽 번호)을 보내 학생이 시작하는 쪽을 찾는다.
// 쪽이 바뀌는 것만 보면 되므로 추론 없이 보낸다.
function uiOcrDetectStarts(input) {
  const firstPage = Number(input && input.firstPage);
  const count = Number(input && input.count);
  const data = String((input && input.data) || '');
  if (!Number.isInteger(firstPage) || firstPage < 1 || !Number.isInteger(count) || count < 1 || count > 10) throw new Error('쪽 번호가 올바르지 않습니다.');
  if (!data) throw new Error('보낼 PDF가 없습니다.');
  if (Math.floor(data.length * 3 / 4) > OCR_MAX_PDF_BYTES) {
    throw new Error(`${firstPage}쪽부터의 파일이 너무 커요. 스캔 해상도를 낮추거나(예: 200dpi) 흑백으로 다시 스캔해 주세요.`);
  }

  const s = ocrSettings_();
  const model = s.models.some(c => c.id === input.model) ? input.model : s.model;
  const result = callAi_(s.provider, model, ocrSplitGuide_(count), `첨부한 ${count}쪽을 쪽마다 판단해 JSON으로 답하세요.`, 'none', OCR_SPLIT_RESPONSE_SCHEMA, {
    attachments: [{ mimeType: 'application/pdf', data, name: `pages-${firstPage}.pdf` }],
    maxTokens: 2000,
  });
  const json = parseOcrJson_(result.text);
  const byPage = {};
  (Array.isArray(json.pages) ? json.pages : []).forEach(p => {
    const page = Math.round(Number(p && p.page));
    if (page >= 1 && page <= count) byPage[page] = p;
  });

  const pages = [];
  for (let i = 1; i <= count; i++) {
    const p = byPage[i] || {};
    pages.push({
      page: firstPage + i - 1,
      newStudent: p.new_student === true,
      studentId: String(p.student_id || '').trim(),
      name: String(p.name || '').trim(),
    });
  }
  log_('uiOcrDetectStarts', 'OK', `${firstPage}~${firstPage + count - 1}쪽 나누기 (${s.provider}:${model})`);
  return pages;
}

function parseOcrJson_(text) {
  const raw = String(text || '').trim();
  try {
    return JSON.parse(raw);
  } catch (err) {
    const match = raw.match(/\{[\s\S]*\}/);
    if (match) return JSON.parse(match[0]);
    throw new Error('AI 답을 읽지 못했어요. 다시 읽어 보세요.');
  }
}

function assertOcrRow_(sh, row, expectedPages) {
  row = Number(row);
  if (!sh || !Number.isInteger(row) || row < 2 || row > sh.getLastRow()) throw new Error('행을 찾을 수 없습니다. 새로고침해 주세요.');
  const pages = String(sh.getRange(row, 1).getValue() || '').trim();
  if (pages !== String(expectedPages || '')) throw new Error('시트 내용이 바뀌었습니다. 새로고침한 뒤 다시 시도하세요.');
  return row;
}

function uiOcrUpdateRow(row, expectedPages, patch) {
  return withDocumentLock_(() => {
    const sh = getOcrSheet_(false);
    row = assertOcrRow_(sh, row, expectedPages);
    const items = ocrItems_();
    if (!ocrLayoutMatches_(sh, items)) throw new Error('손글씨추출 시트의 열이 추출 지침의 항목과 달라요. 새로고침해 주세요.');
    const answers = (patch && patch.answers) || {};
    const texts = items.map(it => String(answers[it.name] == null ? '' : answers[it.name]).trim());
    const values = [String((patch && patch.studentId) || '').trim(), String((patch && patch.name) || '').trim()].concat(texts);
    sh.getRange(row, 2, 1, values.length).setValues([values]);
    // 교사가 직접 고친 행은 읽기 실패 표시를 지운다.
    const noteCol = OCR_FIXED_HEAD.length + items.length + 1;
    const note = String(sh.getRange(row, noteCol).getValue() || '');
    if (texts.some(Boolean) && note.indexOf(OCR_FAILED_PREFIX) === 0) sh.getRange(row, noteCol).setValue('');
    return readOcrRows_(sh).find(r => r.row === row);
  });
}

function uiOcrDeleteRow(row, expectedPages) {
  withDocumentLock_(() => {
    const sh = getOcrSheet_(false);
    row = assertOcrRow_(sh, row, expectedPages);
    sh.deleteRow(row);
  });
  return uiOcrGet();
}

// 손글씨 추출을 처음 상태로 되돌린다. 결과 시트와 지침에서 읽어 둔 열을 지운다. 추출 지침과 모델 설정은 그대로 둔다.
function uiOcrResetAll() {
  withDocumentLock_(() => {
    const sh = getOcrSheet_(false);
    if (sh) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sh);
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(OCR_ITEMS_PROPERTY);
    props.deleteProperty(OCR_ITEMS_HASH_PROPERTY);
  });
  log_('uiOcrResetAll', 'OK', '손글씨 추출을 초기화함');
  return uiOcrGet();
}

function uiOcrClear() {
  withDocumentLock_(() => {
    const sh = getOcrSheet_(false);
    if (sh) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sh);
  });
  log_('uiOcrClear', 'OK', '손글씨 추출 결과를 모두 지움');
  return uiOcrGet();
}

// 추출 결과를 '수동추가'로 넘긴다. 열은 학번, 이름, 지침의 항목들이고, AI에게 보낼 열은 항목 열 모두다(학번·이름은 보내지 않는다).
// mode: backup(지금 시트 백업 후 새로) / overwrite / append(열 구성이 같은 시트 뒤에 붙이기)
function uiOcrSendToManual(mode) {
  const items = ocrItems_();
  const sh = getOcrSheet_(false);
  const rows = readOcrRows_(sh).filter(r => r.filled);
  if (rows.length === 0) throw new Error('옮길 내용이 없습니다. 먼저 글씨를 추출해 주세요.');
  if (!items.length || !ocrLayoutMatches_(sh, items)) throw new Error('손글씨추출 시트의 열이 추출 지침의 항목과 달라요. 결과를 지우고 다시 추출해 주세요.');
  const headers = ['학번', '이름'].concat(items.map(it => it.name));
  const values = rows.map(r => [r.studentId, r.name].concat(items.map(it => r.answers[it.name] || '')));

  if (mode === 'append') {
    withDocumentLock_(() => {
      const manual = getSheet_(MANUAL_SHEET_NAME);
      const existing = manual.getRange(1, 1, 1, Math.max(manual.getLastColumn(), 1)).getValues()[0].map(h => String(h || '').trim());
      const cols = headers.map(name => existing.indexOf(name) + 1);
      const missing = headers.filter((_, i) => cols[i] < 1);
      if (missing.length) {
        throw new Error(`지금 '수동추가' 시트에 ${missing.join('·')} 열이 없어 이어 붙일 수 없어요. 새로 만들기를 골라 주세요.`);
      }
      const config = getManualRecordConfig_(getConfigMap_());
      const want = items.map((_, i) => cols[i + 2]).sort((x, y) => x - y);
      const have = config.inputCols.slice().sort((x, y) => x - y);
      if (want.join(',') !== have.join(',')) {
        const haveNames = have.map(c => existing[c - 1] || colToA1_(c)).join('·');
        throw new Error(`지금 '수동추가' 시트는 AI에게 ${haveNames} 열을 보내도록 되어 있어, 추출한 항목(${items.map(it => it.name).join('·')})과 달라요. 새로 만들기를 골라 주세요.`);
      }
      const width = Math.max.apply(null, cols);
      const start = Math.max(manual.getLastRow() + 1, config.startRow);
      const block = values.map(v => {
        const line = Array(width).fill('');
        cols.forEach((col, i) => { line[col - 1] = v[i]; });
        return line;
      });
      manual.getRange(start, 1, block.length, width).setNumberFormat('@').setValues(block);
    });
    log_('uiOcrSendToManual', 'OK', `손글씨 추출 ${rows.length}행을 수동추가 시트 뒤에 붙임`);
    return Object.assign(uiGetManual(), { imported: rows.length, appended: true });
  }

  const inputCols = items.map((_, i) => i + 3);
  const result = uiImportManualCsv({ headers, rows: values, inputCols, overwrite: mode === 'overwrite' });
  log_('uiOcrSendToManual', 'OK', `손글씨 추출 ${rows.length}행으로 수동추가 시트 생성 (열 ${items.map(it => it.name).join('·')})`);
  return result;
}
