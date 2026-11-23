// Gemini 무료 키 호출 한도 지키기.
// 무료 등급은 키(프로젝트)마다 분당 요청 5회, 분당 토큰 25만 개, 하루 요청 20회로 막혀 있다.
// 한도는 키 기준이라 같은 사용자의 자동 작업·도우미 화면 호출을 모두 합쳐 User Properties에 기록한다.
// 하루는 구글 기준인 태평양 시간 자정에 바뀐다. 유료 키로 표시하면 이 장치는 아무것도 하지 않는다.
const GEMINI_FREE_LIMITS = { rpm: 5, tpm: 250000, rpd: 20 };
const GEMINI_KEY_TIER_PROPERTY = 'GEMINI_KEY_TIER';
const GEMINI_FREE_USAGE_PROPERTY = 'GEMINI_FREE_USAGE';
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

function readGeminiFreeUsage_(now) {
  let usage = null;
  try {
    usage = JSON.parse(PropertiesService.getUserProperties().getProperty(GEMINI_FREE_USAGE_PROPERTY) || 'null');
  } catch (err) { /* 손상된 기록은 새로 시작한다. */ }

  const day = geminiQuotaDay_(now);
  if (!usage || usage.day !== day) usage = { day, count: 0, recent: [] };
  usage.recent = (Array.isArray(usage.recent) ? usage.recent : [])
    .filter(item => Array.isArray(item) && now - Number(item[0]) < GEMINI_QUOTA_WINDOW_MS);
  return usage;
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

function geminiDailyLimitError_(now) {
  const err = new Error(
    `Gemini 무료 키의 하루 요청 한도(${GEMINI_FREE_LIMITS.rpd}회)를 다 썼어요. ` +
    `${geminiQuotaResetText_(now)} 이후 다시 시작하세요. 결제를 연결한 유료 키라면 시작하기에서 키 종류를 '유료'로 바꾸세요. (무료 할당량 소진)`
  );
  err.geminiDailyLimit = true;
  return err;
}

// 한도 안에서 요청 한 번을 예약한다. 분당 한도에 걸리면 자리가 날 때까지 기다리고,
// 실행 시간(6분)이 모자라면 시간 제한 오류로 넘겨 다음 실행에서 이어 가게 한다.
// 하루 한도를 다 썼으면 기다리지 않고 바로 멈춘다. 돌려주는 값은 recordGeminiFreeUsage_에 넘긴다.
function reserveGeminiFreeRequest_(estimatedTokens) {
  const tokens = Math.min(Math.max(0, Math.ceil(Number(estimatedTokens) || 0)), GEMINI_FREE_LIMITS.tpm);

  for (;;) {
    const decision = withGeminiQuotaLock_(() => {
      const now = Date.now();
      const usage = readGeminiFreeUsage_(now);
      if (usage.count >= GEMINI_FREE_LIMITS.rpd) return { error: geminiDailyLimitError_(now) };

      let waitUntil = 0;
      if (usage.recent.length >= GEMINI_FREE_LIMITS.rpm) {
        waitUntil = Number(usage.recent[usage.recent.length - GEMINI_FREE_LIMITS.rpm][0]) + GEMINI_QUOTA_WINDOW_MS;
      }
      // 토큰 한도: 오래된 요청부터 빠져야 이번 요청이 들어갈 자리가 생긴다.
      let used = usage.recent.reduce((sum, item) => sum + Number(item[1] || 0), 0);
      for (let i = 0; used + tokens > GEMINI_FREE_LIMITS.tpm && i < usage.recent.length; i++) {
        used -= Number(usage.recent[i][1] || 0);
        waitUntil = Math.max(waitUntil, Number(usage.recent[i][0]) + GEMINI_QUOTA_WINDOW_MS);
      }
      if (waitUntil > now) return { waitMs: waitUntil - now };

      const id = now + Math.random();
      usage.recent.push([now, tokens, id]);
      usage.count++;
      writeGeminiFreeUsage_(usage);
      return { id, remaining: GEMINI_FREE_LIMITS.rpd - usage.count };
    });

    if (decision.error) throw decision.error;
    if (!decision.waitMs) return decision;

    log_('callGemini_', 'WAIT', `Gemini 무료 키 분당 한도를 지키려고 ${Math.ceil(decision.waitMs / 1000)}초 기다립니다.`);
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
    const item = record.recent.find(entry => entry[2] === reservation.id);
    if (!item) return;
    item[1] = total;
    writeGeminiFreeUsage_(record);
  });
}

function refundGeminiFreeRequest_(reservation) {
  if (!reservation) return;
  withGeminiQuotaLock_(() => {
    const usage = readGeminiFreeUsage_(Date.now());
    if (usage.count > 0) usage.count--;
    writeGeminiFreeUsage_(usage);
  });
}

// 구글이 하루 한도를 다 썼다고 답하면(다른 기기·AI Studio에서 쓴 것 포함) 오늘은 더 보내지 않는다.
function markGeminiDailyQuotaExhausted_(body) {
  if (!/PerDay|per day|daily/i.test(String(body || ''))) return false;
  withGeminiQuotaLock_(() => {
    const usage = readGeminiFreeUsage_(Date.now());
    usage.count = Math.max(usage.count, GEMINI_FREE_LIMITS.rpd);
    writeGeminiFreeUsage_(usage);
  });
  return true;
}

// 화면에 보여 줄 오늘 사용량.
function getGeminiFreeUsageSummary_() {
  const now = Date.now();
  const usage = readGeminiFreeUsage_(now);
  return {
    tier: getGeminiKeyTier_(),
    usedToday: usage.count,
    limits: GEMINI_FREE_LIMITS,
    resetText: geminiQuotaResetText_(now),
  };
}
