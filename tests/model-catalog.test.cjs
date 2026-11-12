const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');
const path = require('node:path');
const root = path.join(__dirname, '..');

function harness() {
  const props = new Map(), cache = new Map(), requests = [], notices = [];
  const propertyStore = {
    getProperty: key => props.get(key) || null,
    setProperty: (key, value) => props.set(key, value),
    deleteProperty: key => props.delete(key),
  };
  const c = vm.createContext({
    console, Date,
    PropertiesService: { getUserProperties: () => propertyStore },
    CacheService: { getUserCache: () => ({ get: k => cache.get(k), put: (k, v) => cache.set(k, v), remove: k => cache.delete(k) }) },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, value) => [...crypto.createHash(alg).update(value).digest()],
      base64EncodeWebSafe: bytes => Buffer.from(bytes).toString('base64url'),
    },
    UrlFetchApp: { fetch: (url, options) => { requests.push({ url, options }); return c.nextResponse(url, options); } },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ toast: (...args) => notices.push(args) }) },
  });
  for (const f of ['Code.js', 'ModelCatalog.js', 'Ui.js', 'UiOcr.js']) vm.runInContext(fs.readFileSync(path.join(root, f), 'utf8'), c, { filename: f });
  c.log_ = () => {};
  c.nextResponse = () => response(200, { data: [] });
  return { c, props, cache, requests, notices, propertyStore };
}
const response = (status, body) => ({ getResponseCode: () => status, getContentText: () => JSON.stringify(body) });
const plain = x => JSON.parse(JSON.stringify(x));
const gemini = id => ({ name: 'models/' + id, supportedGenerationMethods: ['generateContent'] });

test('families, stable releases, numeric versions and aliases are filtered before limiting', () => {
  const { c } = harness();
  const ids = (provider, rows) => plain(c.selectCatalogModels_(provider, rows)).map(m => m.id);
  assert.deepEqual(ids('claude', [
    'claude-haiku-5-9', 'claude-haiku-5-10', 'claude-haiku-5-10-20261101',
    'claude-sonnet-6-1', 'claude-opus-7-0', 'claude-sonnet-7-0-preview',
  ].map(id => ({ id }))), ['claude-haiku-5-10', 'claude-sonnet-6-1']);
  assert.deepEqual(ids('gpt', [
    'gpt-6-luna', 'gpt-6.9-terra', 'gpt-6.10-terra', 'gpt-6.10-terra-2026-11-01',
    'gpt-7-sol', 'gpt-9-terra-codex', 'gpt-9-luna-preview', 'gpt-9-luna-audio', 'ft:gpt-9-luna:custom',
  ].map(id => ({ id }))), ['gpt-6-luna', 'gpt-6.10-terra']);
  assert.deepEqual(ids('gemini', [
    'gemini-3.9-flash', 'gemini-3.10-flash', 'gemini-3.10-flash-001', 'gemini-3.8-flash',
    'gemini-4-flash-lite', 'gemini-4-flash-image', 'gemini-4-flash-preview', 'gemini-4-pro',
  ].map(gemini)), ['gemini-3.10-flash', 'gemini-3.9-flash']);
  assert.deepEqual(ids('gemini', [{ name: 'models/gemini-5-flash', supportedGenerationMethods: ['embedContent'] }]), []);
  assert.deepEqual(ids('gpt', [{ id: 'gpt-6-luna', shutdown_date: '2000-01-01' }]), []);
  assert.deepEqual(ids('claude', [{ id: 'claude-haiku-5-5', lifecycle: 'retired' }]), []);
  assert.deepEqual(ids('claude', [{ id: 'claude-haiku-5-5', line: 'opus' }]), []);
  assert.deepEqual(ids('gemini', ['gemini-3.8-flash-001', 'gemini-3.8-flash-002'].map(gemini)), ['gemini-3.8-flash-002']);
});

test('pagination, headers, cache expiry, refresh and key isolation', () => {
  const { c, props, cache, requests } = harness();
  props.set('GEMINI_API_KEY', 'test-secret');
  c.nextResponse = url => url.includes('pageToken=')
    ? response(200, { models: [gemini('gemini-3.9-flash')] })
    : response(200, { models: [gemini('gemini-3.8-flash')], nextPageToken: 'page /2' });
  let catalog = c.uiGetModelCatalog('gemini');
  assert.equal(requests.length, 2);
  assert.match(requests[1].url, /pageToken=page%20%2F2/);
  assert.equal(requests[0].options.headers['x-goog-api-key'], 'test-secret');
  assert.equal(catalog.models.length, 2);
  assert.ok(!JSON.stringify(catalog).includes('test-secret'));
  assert.ok([...cache.keys()].every(k => !k.includes('test-secret')));
  c.uiGetModelCatalog('gemini'); assert.equal(requests.length, 2);
  c.uiGetModelCatalog('gemini', true); assert.equal(requests.length, 4);
  const key = [...cache.keys()][0], old = JSON.parse(cache.get(key));
  old.fetchedAt = '2000-01-01T00:00:00Z'; cache.set(key, JSON.stringify(old));
  vm.runInContext('modelCatalogMemo_ = {}', c);
  c.uiGetModelCatalog('gemini'); assert.equal(requests.length, 6);
  c.invalidateModelCatalog_('gemini'); assert.equal(cache.size, 0);
  props.set('GEMINI_API_KEY', 'replacement');
  c.uiGetModelCatalog('gemini'); assert.equal(requests.length, 8);
  c.invalidateModelCatalog_('gemini'); props.delete('GEMINI_API_KEY');
  assert.equal(c.uiGetModelCatalog('gemini').status, 'no_key');
  assert.equal(cache.size, 0);
});

test('Claude and OpenAI cursor pagination collect all candidates', () => {
  for (const provider of ['claude', 'gpt']) {
    const { c, props, requests } = harness();
    props.set(provider === 'claude' ? 'CLAUDE_API_KEY' : 'OPENAI_API_KEY', 'key');
    const first = provider === 'claude' ? 'claude-haiku-5-5' : 'gpt-6-luna';
    const second = provider === 'claude' ? 'claude-sonnet-5-5' : 'gpt-6.1-terra';
    c.nextResponse = () => requests.length === 1
      ? response(200, { data: [{ id: first }], has_more: true, last_id: first })
      : response(200, { data: [{ id: second }], has_more: false });
    assert.equal(c.uiGetModelCatalog(provider).models.length, 2);
    assert.equal(requests.length, 2);
    assert.ok(requests[1].url.includes(encodeURIComponent(first)));
  }
});

test('errors do not erase settings, expose server bodies, or reuse stale catalogs', () => {
  for (const [http, status] of [[401, 'auth_error'], [403, 'permission_error'], [429, 'unavailable'], [500, 'unavailable']]) {
    const { c, props } = harness(); props.set('OPENAI_API_KEY', 'secret');
    c.nextResponse = () => response(http, { error: { message: 'secret server response' } });
    const result = c.uiGetModelCatalog('gpt');
    assert.equal(result.status, status);
    assert.equal(result.models.length, 0);
    assert.ok(!JSON.stringify(result).includes('secret'));
  }
  const { c, props } = harness(); props.set('OPENAI_API_KEY', 'key');
  c.nextResponse = () => response(200, { data: [{ id: 'gpt-6-luna' }] });
  c.uiGetModelCatalog('gpt');
  c.nextResponse = () => { throw new Error('network secret'); };
  assert.equal(c.uiGetModelCatalog('gpt', true).status, 'unavailable');
  assert.equal(c.uiGetModelCatalog('gpt').status, 'unavailable');
  assert.throws(() => c.requireCatalogModel_('gpt', 'gpt-6-luna'));
});

test('save validation blocks old and arbitrary models; onOpen preserves saved model values', () => {
  const { c, props } = harness(); props.set('OPENAI_API_KEY', 'key');
  c.nextResponse = () => response(200, { data: [{ id: 'gpt-7-luna' }] });
  c.writeConfigValues_ = patch => { c.written = plain(patch); };
  c.saveSettingsPatch_({ GPT_MODEL: 'gpt-7-luna' });
  assert.deepEqual(c.written, { GPT_MODEL: 'gpt-7-luna' });
  assert.throws(() => c.saveSettingsPatch_({ GPT_MODEL: 'gpt-6-luna' }));
  assert.throws(() => c.saveSettingsPatch_({ REASONING: 'high' }));
  const writes = [];
  c.syncConfigDefaults_({ getLastRow: () => 2, getRange: (...args) => ({
    getValues: () => [['GPT_MODEL', 'gpt-7-luna', 'old description']], setValue: value => writes.push({ args, value }),
  }) }, [['GPT_MODEL', '', 'dynamic description']]);
  assert.ok(writes.every(write => write.args[1] !== 2));
  c.getConfigMap_ = () => ({ AI_PROVIDER: 'gpt', GPT_MODEL: 'gpt-6-luna' });
  assert.throws(() => c.requireCurrentAiApiKey_());
});

test('OCR shares the catalog, recommends upper family, and preserves invalid saved choices', () => {
  const { c, props } = harness(); props.set('OPENAI_API_KEY', 'key');
  props.set('OCR_MODEL_gpt', 'gpt-6-luna');
  c.getConfigMap_ = () => ({ AI_PROVIDER: 'gpt' });
  c.nextResponse = () => response(200, { data: [{ id: 'gpt-7-luna' }, { id: 'gpt-7-terra' }] });
  let settings = c.ocrSettings_();
  assert.equal(settings.models[0].id, 'gpt-7-terra');
  // Use the actual property prefix defined by the OCR module.
  vm.runInContext("PropertiesService.getUserProperties().setProperty(OCR_MODEL_PROPERTY_PREFIX + 'gpt', 'gpt-6-luna')", c);
  settings = c.ocrSettings_();
  assert.equal(settings.model, 'gpt-6-luna');
  assert.throws(() => c.requireOcrModel_(settings, 'gpt-7-terra'));
  assert.throws(() => c.uiOcrSaveSettings({ model: 'gpt-6-luna' }));
  assert.equal(c.uiOcrSaveSettings({ model: 'gpt-7-terra', reasoning: 'medium' }).model, 'gpt-7-terra');
});

test('Gemini reports minimal fallback once and never retries a generic error or high-only configuration', () => {
  const { c, props, notices } = harness(); props.set('GEMINI_API_KEY', 'key');
  const sent = [];
  c.fetchAiWithinBudget_ = (url, options) => {
    sent.push(JSON.parse(options.payload));
    return sent.length === 1
      ? response(400, { error: { message: 'THINKING_LEVEL_MINIMAL is not supported' } })
      : response(200, { output_text: 'ok' });
  };
  c.callGemini_('gemini-4-flash', 'test', 'test', 'none');
  assert.deepEqual(sent.map(p => p.generation_config.thinking_level), ['minimal', 'low']);
  assert.equal(notices.length, 1);
  assert.equal(c.getGeminiThinkingLevel_('none', 'gemini-4-flash'), 'low');
  let count = 0;
  c.fetchAiWithinBudget_ = () => { count++; return response(400, { error: { message: 'invalid temperature' } }); };
  assert.throws(() => c.callGemini_('gemini-4-flash', 'test', 'test', 'medium'));
  assert.equal(count, 1);
  c.fetchAiWithinBudget_ = () => response(400, { error: { param: 'thinking_level', message: 'Unsupported value medium. Supported levels: high' } });
  assert.throws(() => c.callGemini_('gemini-4-flash', 'test', 'test', 'medium'));
});

test('key lifecycle endpoints preserve bootstrap response and discard former key cache', () => {
  const { c, props, cache } = harness();
  c.uiGetBootstrap = () => ({ keys: { gpt: !!props.get('OPENAI_API_KEY') }, model: 'saved-model' });
  props.set('OPENAI_API_KEY', 'old-key');
  c.nextResponse = () => response(200, { data: [{ id: 'gpt-7-luna' }] });
  c.uiGetModelCatalog('gpt'); assert.equal(cache.size, 1);
  const saved = c.uiSaveApiKey('gpt', 'replacement');
  assert.equal(saved.model, 'saved-model'); assert.equal(cache.size, 0);
  c.uiGetModelCatalog('gpt'); assert.equal(cache.size, 1);
  assert.equal(c.uiClearApiKey('gpt').keys.gpt, false);
  assert.equal(cache.size, 0);
});

test('client keeps provider results separate when responses arrive in reverse order', async () => {
  const c = vm.createContext({
    document: { addEventListener() {}, querySelectorAll: () => [], querySelector: () => ({ addEventListener() {} }) },
    window: { addEventListener() {} }, console, setTimeout, clearTimeout, setInterval, clearInterval,
  });
  const source = fs.readFileSync(path.join(root, 'AppScript.html'), 'utf8').replace(/^<script>/, '').replace(/<\/script>\s*$/, '');
  vm.runInContext(source, c);
  vm.runInContext("S.boot={provider:'gpt'};S.settings={groups:{ai:[{key:'CLAUDE_MODEL',value:''},{key:'GPT_MODEL',value:''}]},catalogs:{}}", c);
  const pending = {};
  c.call = (name, provider) => new Promise(resolve => { pending[provider] = resolve; });
  const earlier = c.reloadModelCatalog('claude', true), later = c.reloadModelCatalog('gpt', true);
  pending.gpt({ provider: 'gpt', status: 'ok', models: [{ id: 'gpt-7-luna' }] }); await later;
  pending.claude({ provider: 'claude', status: 'ok', models: [{ id: 'claude-haiku-6-0' }] }); await earlier;
  assert.equal(vm.runInContext('catalogFor(S.boot.provider).models[0].id', c), 'gpt-7-luna');
  assert.equal(vm.runInContext("S.settings.groups.ai.find(s=>s.key==='GPT_MODEL').options[0]", c), 'gpt-7-luna');
});

test('reasoning uses only explicit supported values up to medium', () => {
  const { c } = harness();
  const error = message => JSON.stringify({ error: { param: 'reasoning.effort', message } });
  assert.equal(c.supportedReasoningRetry_(error("Unsupported value 'none'. Supported values are: 'low', 'medium', 'high'."), 'none', 'gpt'), 'low');
  assert.equal(c.supportedReasoningRetry_(error("Unsupported value 'medium'. Supported values are: 'high', 'xhigh'."), 'medium', 'gpt'), '');
  assert.equal(c.supportedReasoningRetry_(error('Invalid API key'), 'none', 'gpt'), '');
  assert.equal(c.supportedReasoningRetry_('not json', 'none', 'gpt'), '');
  assert.equal(c.normalizeReasoning_('max'), 'none');
  c.catalogClaudeModel_ = () => ({ capabilities: {
    thinking: { types: { disabled: { supported: true }, adaptive: { supported: true } } },
    effort: { low: { supported: true }, medium: { supported: true } },
  } });
  const payload = {};
  c.applyClaudeReasoning_(payload, 'claude-haiku-6-0', 'none');
  assert.equal(payload.thinking.type, 'disabled');
  c.applyClaudeReasoning_(payload, 'claude-haiku-6-0', 'medium');
  assert.equal(payload.output_config.effort, 'medium');
  c.catalogClaudeModel_ = () => ({ capabilities: { thinking: { types: { adaptive: { supported: true } } }, effort: { high: { supported: true } } } });
  assert.throws(() => c.applyClaudeReasoning_({}, 'claude-haiku-6-0', 'medium'));
});

test('single and parallel payloads normalize high effort; explicit GPT retry occurs only once', () => {
  const { c, props } = harness(); props.set('OPENAI_API_KEY', 'key');
  c.ensureAiTimeBudget_ = () => {};
  c.UrlFetchApp.fetchAll = requests => {
    assert.equal(JSON.parse(requests[0].payload).reasoning.effort, 'none');
    return [response(200, { output: [{ type: 'message', content: [{ type: 'output_text', text: 'ok' }] }] })];
  };
  c.fetchAiBatch_([{ provider: 'gpt', model: 'gpt-7-luna', reasoning: 'high', systemGuide: 'test', userPrompt: 'test' }]);
  const sent = [];
  c.fetchAiWithinBudget_ = (url, options) => {
    sent.push(JSON.parse(options.payload));
    return sent.length === 1 ? response(400, { error: { param: 'reasoning.effort', message: "Unsupported value 'none'. Supported values are: 'low', 'medium', 'high'." } })
      : response(200, { output: [{ type: 'message', content: [{ type: 'output_text', text: 'ok' }] }] });
  };
  assert.equal(c.callOpenAi_('gpt-7-luna', 'test', 'test', 'none').text, 'ok');
  assert.equal(sent.length, 2);
  assert.equal(sent[1].reasoning.effort, 'low');
  c.fetchAiWithinBudget_ = () => response(400, { error: { param: 'reasoning.effort', message: "Unsupported value. Supported values are: 'high', 'max'." } });
  assert.throws(() => c.callOpenAi_('gpt-7-luna', 'test', 'test', 'medium'));
});
