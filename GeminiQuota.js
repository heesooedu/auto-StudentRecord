// Gemini 무료 키 호출 한도 지키기.
// 무료 등급은 키(프로젝트)마다, 그리고 모델마다 따로 분당 요청 5회, 분당 토큰 25만 개, 하루 요청 20회로 막혀 있다.
// 예: gemini-4-hydrogen 20회와 gemini-3.8-flash 20회는 서로 다른 몫이다.
// 같은 사용자의 자동 작업·도우미 화면 호출을 모델별로 합쳐 User Properties에 기록한다.
// 하루는 구글 기준인 태평양 시간 자정에 바뀐다. 유료 키로 표시하면 이 장치는 아무것도 하지 않는다.
const GEMINI_FREE_LIMITS = { rpm: 5, tpm: 250000, rpd: 20 };
const GEMINI_KEY_TIER_PROPERTY = 'GEMINI_KEY_TIER';
const GEMINI_FREE_USAGE_PROPERTY = 'GEMINI_FREE_USAGE_V2';
const GEMINI_QUOTA_TIME_ZONE = 'America/Los_Angeles';
const GEMINI_QUOTA_WINDOW_MS = 60 * 1000;

// 저장하지 않았으면 무료로 본다. 유료 키를 무료로 보면 느릴 뿐이지만, 반대는 한도 오류로 멈춘다.
function getGeminiKeyTier_() {
  return PropertiesService.getUserProperties().getProperty(GEMINI_KEY_TIER_PROPERTY) === 'paid' ? 'paid' : 'free';
}

function setGeminiKeyTier_(tier) {
  if (tier !== 'free' && tier !== 'paid') throw new Error('Gemini 키 종류는 free 또는 paid여야 합니다.');
  PropertiesService.getUserProperties().setProperty(GEMINI_KEY_TIER_PROPERTY, tier);
}

function isGeminiFreeTier_() {
  return getGeminiKeyTier_() === 'free';
}

function geminiQuotaDay_(now) {
  return Utilities.formatDate(new Date(now), GEMINI_QUOTA_TIME_ZONE, 'yyyy-MM-dd');
}

// 태평양 시간 자정은 한국 시간으로 서머타임 때 오후 4시, 아닐 때 오후 5시다.
function geminiQuotaResetText_(now) {
  const offset = Utilities.formatDate(new Date(now), GEMINI_QUOTA_TIME_ZONE, 'Z');
  return offset === '-0700' ? '한국 시간 오후 4시' : '한국 시간 오후 5시';
}

function geminiQuotaModelKey_(model) {
  return String(model || '').trim().toLowerCase().replace(/^models\//, '') || 'unknown';
}

// 기록 형식: { day, models: { 'gemini-4-hydrogen': { count, recent: [[보낸 시각, 토큰, id], ...] } } }
function readGeminiFreeUsage_(now) {
  let usage = null;
  try {
    usage = JSON.parse(PropertiesService.getUserProperties().getProperty(GEMINI_FREE_USAGE_PROPERTY) || 'null');
  } catch (err) { /* 손상된 기록은 새로 시작한다. */ }

  const day = geminiQuotaDay_(now);
  if (!usage || usage.day !== day || !usage.models || typeof usage.models !== 'object') usage = { day, models: {} };
  Object.keys(usage.models).forEach(model => {
    const entry = usage.models[model] || {};
    entry.count = Number(entry.count) || 0;
    entry.recent = (Array.isArray(entry.recent) ? entry.recent : [])
      .filter(item => Array.isArray(item) && now - Number(item[0]) < GEMINI_QUOTA_WINDOW_MS);
    usage.models[model] = entry;
  });
  return usage;
}

function geminiModelUsage_(usage, model) {
  const key = geminiQuotaModelKey_(model);
  if (!usage.models[key]) usage.models[key] = { count: 0, recent: [] };
  return usage.models[key];
}

function writeGeminiFreeUsage_(usage) {
  PropertiesService.getUserProperties().setProperty(GEMINI_FREE_USAGE_PROPERTY, JSON.stringify(usage));
}

function withGeminiQuotaLock_(fn) {
  const lock = LockService.getUserLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// 보내기 전에 쓸 토큰을 넉넉히 어림한다. 한국어는 글자 하나가 토큰 하나쯤이고, 출력 한도는 다 쓴다고 본다.
// PDF는 쪽마다 258토큰이 들며, 스캔 한 쪽을 약 200KB로 보고 바이트 크기로 어림한다.
function estimateGeminiRequestTokens_(payload) {
  const text = [payload.system_instruction, payload.input].map(part => {
    if (typeof part === 'string') return part.length;
    if (!Array.isArray(part)) return 0;
    return part.reduce((sum, item) => {
      if (item && item.type === 'text') return sum + String(item.text || '').length;
      if (item && item.data) return sum + Math.ceil(String(item.data).length * 0.75 / 800);
      return sum;
    }, 0);
  }).reduce((a, b) => a + b, 0);
  const output = Number(payload.generation_config && payload.generation_config.max_output_tokens) || 0;
  return text + output;
}

function geminiDailyLimitError_(model, now) {
  const err = new Error(
    `Gemini 무료 키로 ${geminiQuotaModelKey_(model)} 모델의 하루 요청 한도(${GEMINI_FREE_LIMITS.rpd}회)를 다 썼어요. ` +
    `한도는 모델마다 따로라서, 시작하기에서 다른 Gemini 모델로 바꾸면 이어서 만들 수 있어요. ` +
    `이 모델은 ${geminiQuotaResetText_(now)} 이후 다시 쓸 수 있어요. 결제를 연결한 유료 키라면 키 종류를 '유료'로 바꾸세요. (무료 할당량 소진)`
  );
  err.geminiDailyLimit = true;
  return err;
}

// 한도 안에서 요청 한 번을 예약한다. 분당 한도에 걸리면 자리가 날 때까지 기다리고,
// 실행 시간(6분)이 모자라면 시간 제한 오류로 넘겨 다음 실행에서 이어 가게 한다.
// 하루 한도를 다 썼으면 기다리지 않고 바로 멈춘다. 돌려주는 값은 recordGeminiFreeUsage_에 넘긴다.
function reserveGeminiFreeRequest_(model, estimatedTokens) {
  const tokens = Math.min(Math.max(0, Math.ceil(Number(estimatedTokens) || 0)), GEMINI_FREE_LIMITS.tpm);

  for (;;) {
    const decision = withGeminiQuotaLock_(() => {
      const now = Date.now();
      const usage = readGeminiFreeUsage_(now);
      const entry = geminiModelUsage_(usage, model);
      if (entry.count >= GEMINI_FREE_LIMITS.rpd) return { error: geminiDailyLimitError_(model, now) };

      let waitUntil = 0;
      if (entry.recent.length >= GEMINI_FREE_LIMITS.rpm) {
        waitUntil = Number(entry.recent[entry.recent.length - GEMINI_FREE_LIMITS.rpm][0]) + GEMINI_QUOTA_WINDOW_MS;
      }
      // 토큰 한도: 오래된 요청부터 빠져야 이번 요청이 들어갈 자리가 생긴다.
      let used = entry.recent.reduce((sum, item) => sum + Number(item[1] || 0), 0);
      for (let i = 0; used + tokens > GEMINI_FREE_LIMITS.tpm && i < entry.recent.length; i++) {
        used -= Number(entry.recent[i][1] || 0);
        waitUntil = Math.max(waitUntil, Number(entry.recent[i][0]) + GEMINI_QUOTA_WINDOW_MS);
      }
      if (waitUntil > now) return { waitMs: waitUntil - now };

      const id = now + Math.random();
      entry.recent.push([now, tokens, id]);
      entry.count++;
      writeGeminiFreeUsage_(usage);
      return { id, model: geminiQuotaModelKey_(model), remaining: GEMINI_FREE_LIMITS.rpd - entry.count };
    });

    if (decision.error) throw decision.error;
    if (!decision.waitMs) return decision;

    log_('callGemini_', 'WAIT', `Gemini 무료 키 분당 한도를 지키려고 ${Math.ceil(decision.waitMs / 1000)}초 기다립니다. (${geminiQuotaModelKey_(model)})`);
    sleepWithinAiBudget_(decision.waitMs + 300);
  }
}

// 응답에 실제 토큰 수가 있으면 어림값을 바꿔 둔다(분당 토큰 계산을 덜 보수적으로).
function recordGeminiFreeUsage_(reservation, usage) {
  if (!reservation || !usage) return;
  const total = Number(usage.total_tokens || usage.totalTokenCount || usage.total_token_count || 0);
  if (!total) return;

  withGeminiQuotaLock_(() => {
    const record = readGeminiFreeUsage_(Date.now());
    const item = geminiModelUsage_(record, reservation.model).recent.find(entry => entry[2] === reservation.id);
    if (!item) return;
    item[1] = total;
    writeGeminiFreeUsage_(record);
  });
}

function refundGeminiFreeRequest_(reservation) {
  if (!reservation) return;
  withGeminiQuotaLock_(() => {
    const usage = readGeminiFreeUsage_(Date.now());
    const entry = geminiModelUsage_(usage, reservation.model);
    if (entry.count > 0) entry.count--;
    writeGeminiFreeUsage_(usage);
  });
}

// 구글이 하루 한도를 다 썼다고 답하면(다른 기기·AI Studio에서 쓴 것 포함) 오늘은 이 모델로 더 보내지 않는다.
function markGeminiDailyQuotaExhausted_(model, body) {
  if (!/PerDay|per day|daily/i.test(String(body || ''))) return false;
  withGeminiQuotaLock_(() => {
    const usage = readGeminiFreeUsage_(Date.now());
    const entry = geminiModelUsage_(usage, model);
    entry.count = Math.max(entry.count, GEMINI_FREE_LIMITS.rpd);
    writeGeminiFreeUsage_(usage);
  });
  return true;
}

// 화면에 보여 줄 오늘 사용량(모델별).
function getGeminiFreeUsageSummary_() {
  const now = Date.now();
  const usage = readGeminiFreeUsage_(now);
  const models = {};
  Object.keys(usage.models).forEach(model => { models[model] = usage.models[model].count; });
  return {
    tier: getGeminiKeyTier_(),
    usedByModel: models,
    limits: GEMINI_FREE_LIMITS,
    resetText: geminiQuotaResetText_(now),
  };
}
