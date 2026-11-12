// API 키와 접근권한은 사용자별이다. 공유 설정 시트에 모델 목록을 저장하지 않는다.
const MODEL_SETTING_PROVIDERS = { CLAUDE_MODEL: 'claude', GPT_MODEL: 'gpt', GEMINI_MODEL: 'gemini' };
const MODEL_CATALOG_TTL = 21600;
const MODEL_SELECTION_HELP = 'API 키를 저장한 뒤 생기부 도우미에서 최신 모델을 선택하고 저장하세요.';
const MODEL_CATALOG_KEY_PROPERTIES = { claude: 'CLAUDE_API_KEY', gpt: 'OPENAI_API_KEY', gemini: 'GEMINI_API_KEY' };
const MODEL_CATALOG_URLS = {
  claude: 'https://api.anthropic.com/v1/models',
  gpt: 'https://api.openai.com/v1/models',
  gemini: 'https://generativelanguage.googleapis.com/v1beta/models',
};
let modelCatalogMemo_ = {};

function modelCatalogCacheKey_(provider, key) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, key, Utilities.Charset.UTF_8);
  return 'AI_MODELS_V1_' + provider + '_' + Utilities.base64EncodeWebSafe(digest);
}

function invalidateModelCatalog_(provider) {
  const key = PropertiesService.getUserProperties().getProperty(MODEL_CATALOG_KEY_PROPERTIES[provider]);
  if (key) {
    const cacheKey = modelCatalogCacheKey_(provider, key);
    CacheService.getUserCache().remove(cacheKey);
    delete modelCatalogMemo_[cacheKey];
  }
}

function modelCatalogFailure_(provider, status, message) {
  return { provider, status, message, models: [], fetchedAt: null };
}

// 숫자로 버전을 비교한다. 3.10은 3.9보다 새 버전이며 스냅샷 날짜는 버전에서 뺀다.
function compareModelVersions_(a, b) {
  const av = a.version.split(/[.-]/).map(Number);
  const bv = b.version.split(/[.-]/).map(Number);
  for (let i = 0; i < Math.max(av.length, bv.length); i++) {
    const diff = (bv[i] || 0) - (av[i] || 0);
    if (diff) return diff;
  }
  return (b.created || 0) - (a.created || 0) || (b.snapshot || 0) - (a.snapshot || 0) || a.id.localeCompare(b.id);
}

function selectCatalogModels_(provider, rows) {
  const versions = {};
  rows.forEach(row => {
    const id = String(row.id || row.name || '').replace(/^models\//, '');
    if (/(preview|experimental|\bexp\b|beta|alpha)/i.test(id + ' ' + (row.display_name || row.displayName || ''))) return;
    if (row.lifecycle && row.lifecycle !== 'active') return;
    if (row.shutdown_date && Date.parse(row.shutdown_date) <= Date.now()) return;
    let match, family, version;
    if (provider === 'claude') {
      match = id.match(/^claude-(haiku|sonnet)-(\d+(?:-\d{1,2})*)(?:-(\d{8})|-(latest))?$/);
      if (!match || (Object.prototype.hasOwnProperty.call(row, 'line') && row.line !== match[1])) return;
      family = match[1]; version = match[2];
    } else if (provider === 'gpt') {
      match = id.match(/^gpt-(\d+(?:\.\d+)*?)-(luna|terra)(?:-(\d{4}-\d{2}-\d{2}))?$/);
      if (!match) return;
      family = match[2]; version = match[1];
    } else {
      match = id.match(/^gemini-(\d+(?:\.\d+)*)-flash(?:-(\d{3}))?$/);
      if (!match || !(row.supportedGenerationMethods || []).includes('generateContent')) return;
      family = 'flash'; version = match[1];
    }
    const item = {
      id, family, version,
      created: Number(row.created) || Date.parse(row.created_at || '') || 0,
      snapshot: Number((id.match(/-(\d{8}|\d{4}-\d{2}-\d{2}|\d{3})$/) || [])[1]?.replace(/-/g, '')) || 0,
      capabilities: row.capabilities || null,
    };
    const group = family + ':' + version;
    const old = versions[group];
    // 같은 버전은 기본 ID를 우선하며, 기본 ID가 없으면 최신 스냅샷을 고른다.
    const base = provider === 'claude' ? `claude-${family}-${version}` : provider === 'gpt' ? `gpt-${version}-${family}` : `gemini-${version}-flash`;
    if (!old || id === base || (old.id !== base && compareModelVersions_(item, old) < 0)) versions[group] = item;
  });
  const sorted = Object.keys(versions).map(k => versions[k]).sort(compareModelVersions_);
  if (provider === 'gemini') return sorted.slice(0, 2);
  return (provider === 'claude' ? ['haiku', 'sonnet'] : ['luna', 'terra'])
    .map(family => sorted.find(m => m.family === family)).filter(Boolean);
}

function fetchModelCatalogRows_(provider, key) {
  const rows = [];
  const seen = {};
  let cursor = '';
  do {
    if (seen[cursor]) throw new Error('pagination');
    seen[cursor] = true;
    let url = MODEL_CATALOG_URLS[provider];
    const headers = provider === 'claude'
      ? { 'x-api-key': key, 'anthropic-version': '2026-10-21' }
      : provider === 'gpt' ? { Authorization: 'Bearer ' + key } : { 'x-goog-api-key': key };
    if (provider === 'claude') url += '?limit=1000' + (cursor ? '&after_id=' + encodeURIComponent(cursor) : '');
    if (provider === 'gemini') url += '?pageSize=1000' + (cursor ? '&pageToken=' + encodeURIComponent(cursor) : '');
    if (provider === 'gpt' && cursor) url += '?after=' + encodeURIComponent(cursor);
    const response = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    const code = response.getResponseCode();
    if (code < 200 || code >= 300) {
      const err = new Error('catalog HTTP ' + code);
      let invalidKey = false;
      try {
        const error = JSON.parse(response.getContentText()).error || {};
        invalidKey = error.status === 'UNAUTHENTICATED' || (error.details || []).some(d => d.reason === 'API_KEY_INVALID');
      } catch (parseError) { /* 오류 원문은 화면이나 로그로 보내지 않는다. */ }
      err.catalogStatus = code === 401 || invalidKey ? 'auth_error' : code === 403 ? 'permission_error' : 'unavailable';
      throw err;
    }
    const json = JSON.parse(response.getContentText());
    const page = provider === 'gemini' ? json.models : json.data;
    if (!Array.isArray(page)) throw new Error('invalid catalog');
    rows.push.apply(rows, page);
    cursor = provider === 'gemini' ? (json.nextPageToken || '') : (json.has_more ? json.last_id || (page.length && page[page.length - 1].id) : '');
    if (json.has_more && !cursor) throw new Error('missing cursor');
  } while (cursor);
  return rows;
}

function getModelCatalog_(provider, forceRefresh) {
  if (!MODEL_CATALOG_KEY_PROPERTIES[provider]) throw new Error('알 수 없는 AI 서비스입니다.');
  const key = PropertiesService.getUserProperties().getProperty(MODEL_CATALOG_KEY_PROPERTIES[provider]);
  if (!key) return modelCatalogFailure_(provider, 'no_key', '시작하기에서 API 키를 먼저 저장해 주세요.');
  const cacheKey = modelCatalogCacheKey_(provider, key);
  const cache = CacheService.getUserCache();
  if (!forceRefresh) {
    let saved = modelCatalogMemo_[cacheKey];
    if (!saved) {
      try { saved = JSON.parse(cache.get(cacheKey) || 'null'); } catch (err) { /* 손상된 캐시는 다시 조회 */ }
    }
    if (saved && Date.now() - Date.parse(saved.fetchedAt) < MODEL_CATALOG_TTL * 1000) return saved;
  }
  delete modelCatalogMemo_[cacheKey];
  cache.remove(cacheKey);
  try {
    const models = selectCatalogModels_(provider, fetchModelCatalogRows_(provider, key));
    const result = { provider, status: models.length ? 'ok' : 'empty', models, fetchedAt: new Date().toISOString(),
      message: models.length ? '정식 모델 목록을 불러왔어요.' : '이 키로 조회되는 모델 중 조건에 맞는 정식 모델이 없어요.' };
    modelCatalogMemo_[cacheKey] = result;
    cache.put(cacheKey, JSON.stringify(result), MODEL_CATALOG_TTL);
    return result;
  } catch (err) {
    const status = err.catalogStatus || 'unavailable';
    return modelCatalogFailure_(provider, status, status === 'auth_error'
      ? 'API 키 인증에 실패했어요. 키를 확인하고 다시 저장해 주세요.'
      : status === 'permission_error' ? '모델 목록 조회 권한이 없어요. 키의 조회 권한을 확인해 주세요.'
      : '모델 목록을 불러오지 못했어요. 잠시 뒤 새로고침해 주세요.');
  }
}

function uiGetModelCatalog(provider, forceRefresh) {
  return publicModelCatalog_(getModelCatalog_(provider, forceRefresh === true));
}

function publicModelCatalog_(result) {
  return Object.assign({}, result, { models: result.models.map(m => ({ id: m.id, family: m.family, version: m.version })) });
}

function requireCatalogModel_(provider, model) {
  const catalog = getModelCatalog_(provider);
  if (catalog.status !== 'ok') throw new Error(catalog.message);
  const selected = catalog.models.find(m => m.id === model);
  if (!selected) throw new Error('저장된 모델이 최신 선택 목록에 없어요. 시작하기에서 새 모델을 선택하고 저장해 주세요.');
  return selected;
}

function modelCatalogChoices_(provider, ocr) {
  const catalog = getModelCatalog_(provider);
  const models = catalog.models.slice();
  if (ocr) models.sort((a, b) => Number(['sonnet', 'terra'].includes(b.family)) - Number(['sonnet', 'terra'].includes(a.family)));
  return { catalog, models };
}

function catalogClaudeModel_(model) {
  return getModelCatalog_('claude').models.find(m => m.id === model);
}

// 특정 추론 값 미지원과 지원하는 값을 명시한 오류만 재시도한다.
function supportedReasoningRetry_(body, requested, provider) {
  let error;
  try { error = JSON.parse(body).error; } catch (err) { return ''; }
  if (!error) return '';
  const message = String(error.message || '');
  const param = String(error.param || '');
  if (!/reasoning|thinking_level|effort/i.test(param + ' ' + message) || !/not supported|unsupported|not.*allowed|invalid.*value/i.test(message)) return '';
  const supported = message.match(/supported (?:values|levels)(?: are| include)?\s*[:=]?\s*(.+)/i);
  if (!supported) return '';
  const allowed = (supported[1].match(/\b(none|minimal|low|medium|high|xhigh|max)\b/g) || []);
  const order = provider === 'gemini' ? ['minimal', 'low', 'medium'] : ['none', 'low', 'medium'];
  const current = order.indexOf(requested);
  return order.filter(v => allowed.includes(v) && v !== requested)
    .sort((a, b) => Math.abs(order.indexOf(a) - current) - Math.abs(order.indexOf(b) - current) || order.indexOf(a) - order.indexOf(b))[0] || '';
}

function reportReasoningAdjustment_(provider, model, requested, actual) {
  const message = `${provider} ${model}: 추론 ${requested} 미지원으로 ${actual} 적용 (최대 medium)`;
  log_('AI 추론', 'INFO', message);
  try { SpreadsheetApp.getActiveSpreadsheet().toast(message, '추론 수준 조정', 8); } catch (err) { /* 알림 실패로 생성을 중단하지 않는다. */ }
}
