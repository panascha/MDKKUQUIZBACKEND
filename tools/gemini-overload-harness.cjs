// Offline harness for callGeminiAI overload/emergency fallback. Run: node tools/gemini-overload-harness.cjs (from repo root)
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');
const root = path.join(__dirname, '..');
const src = fs.readFileSync(path.join(root, 'ai-gemini.gs'), 'utf8');
new vm.Script(src, { filename: 'ai-gemini.gs' }); // syntax check

// behavior: { model: 'ok' | 'overload' | 'quota' } — anything else = ok
function makeCtx(behavior, preCache) {
  const fetched = [], warns = [], cache = Object.assign({}, preCache || {});
  const ctx = {
    console: { log() {}, warn: (...a) => warns.push(a.join(' ')), error() {} },
    fetched, warns,
    EXEC_START_MS: Date.now(),
    execBudgetExhausted_: () => false,
    CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
    SpreadsheetApp: { openById: () => { throw new Error('no sheet in harness'); } },
    UrlFetchApp: {
      fetch: (url) => {
        const model = url.match(/models\/([^:]+):/)[1];
        fetched.push(model);
        const b = behavior[model] || 'ok';
        if (b === 'overload') return { getResponseCode: () => 503, getContentText: () => JSON.stringify({ error: { message: 'This model is currently experiencing high demand.' } }), getAllHeaders: () => ({}) };
        if (b === 'gone') return { getResponseCode: () => 404, getContentText: () => JSON.stringify({ error: { message: 'This model models/' + model + ' is no longer available to new users.' } }), getAllHeaders: () => ({}) };
        if (b === 'quota') return { getResponseCode: () => 429, getContentText: () => JSON.stringify({ error: { message: 'quota' } }), getAllHeaders: () => ({}) };
        return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: 'answer from ' + model }] }, finishReason: 'STOP' }] }), getAllHeaders: () => ({}) };
      }
    },
  };
  vm.createContext(ctx);
  vm.runInContext(src, ctx);
  ctx.handleGemini429_ = () => {};
  return ctx;
}
function run(ctx, keyInfo) {
  try { return { text: ctx.callGeminiAI('q', keyInfo, null) }; } catch (e) { return { err: e.message }; }
}
const LITE = ['gemini-3.5-flash-lite', 'gemini-3.1-flash-lite'];
// key picked lite (askAIExpert) — fallbackModels = every model the key still has quota on, priority order
const keyAll = { key: 'K', model: 'gemini-3.5-flash-lite', usage: 0, limit: 500,
  fallbackModels: ['gemini-3.8-flash', 'gemini-3.7-flash', 'gemini-3.6-flash', 'gemini-3.5-flash', 'gemini-2.5-flash'].concat(LITE) };
const keyNoEmergency = { key: 'K', model: 'gemini-3.5-flash-lite', usage: 0, limit: 500,
  fallbackModels: ['gemini-3.8-flash', 'gemini-3.7-flash'].concat(LITE) }; // no 3.6/2.5/3.5-flash quota on this key
const bothOverloaded = { 'gemini-3.5-flash-lite': 'overload', 'gemini-3.1-flash-lite': 'overload' };
let pass = 0, fail = 0;
function check(name, fn) { try { fn(); pass++; console.log('PASS ' + name); } catch (e) { fail++; console.log('FAIL ' + name + ' — ' + e.message); } }

// (a) both lite 503 → served by emergency model outside converter chain
check('a: both lite overloaded -> emergency', () => {
  const c = makeCtx(bothOverloaded); const r = run(c, keyAll);
  assert.ok(r.text, 'threw: ' + r.err);
  assert.ok(/gemini-3\.6-flash|gemini-2\.5-flash/.test(r.text), r.text);
  assert.ok(!c.fetched.some(m => /3\.8-flash$|3\.7-flash$|^gemini-3\.5-flash$/.test(m)), 'touched converter model: ' + c.fetched);
});
// (b) both lite circuit-open → pass 0 skips straight to emergency, no lite fetch
check('b: both lite circuit-open -> emergency, no lite fetch', () => {
  const c = makeCtx({}, { 'model_down:gemini-3.5-flash-lite': '1', 'model_down:gemini-3.1-flash-lite': '1' });
  const r = run(c, keyAll);
  assert.ok(r.text, 'threw: ' + r.err);
  assert.ok(!c.fetched.some(m => /lite/.test(m)), 'fetched lite: ' + c.fetched);
});
// (c) no emergency model has quota on this key → throws cleanly, trail logged
check('c: no emergency quota -> clean throw + trail warn', () => {
  const c = makeCtx(bothOverloaded); const r = run(c, keyNoEmergency);
  assert.ok(r.err && /^AI Assistant Error: gemini-3\.1-flash-lite: .*high demand/.test(r.err), String(r.err));
  assert.ok(!c.fetched.some(m => !/lite/.test(m)), 'fetched non-lite: ' + c.fetched);
  assert.ok(c.warns.some(w => /trail/i.test(w)), 'no trail warn');
});
// (d) lite ok → no emergency fetch
check('d: lite ok -> single fetch', () => {
  const c = makeCtx({}); const r = run(c, keyAll);
  assert.strictEqual(r.text, 'answer from gemini-3.5-flash-lite');
  assert.deepStrictEqual(c.fetched, ['gemini-3.5-flash-lite']);
});
// (e) everything overloaded → fetch cap respected
check('e: all overloaded -> bounded fetches', () => {
  const all = {}; keyAll.fallbackModels.forEach(m => { all[m] = 'overload'; });
  const c = makeCtx(all); const r = run(c, keyAll);
  assert.ok(r.err, 'should throw');
  assert.ok(c.fetched.length <= 6, 'fetches=' + c.fetched.length);
  assert.ok(!c.fetched.some(m => /3\.8-flash$|3\.7-flash$/.test(m)), 'touched non-last-resort converter model: ' + c.fetched);
});
// (f) non-lite pick (e.g. verify judge 3.5-flash) → chain untouched, no emergency logic
check('f: full-flash pick unchanged', () => {
  const c = makeCtx({ 'gemini-3.5-flash': 'overload' });
  const r = run(c, { key: 'K', model: 'gemini-3.5-flash', usage: 0, limit: 20, fallbackModels: ['gemini-3.5-flash', 'gemini-2.5-flash'] });
  assert.strictEqual(r.text, 'answer from gemini-2.5-flash');
});
// (g) 2026-10-13 live: lite overloaded, 3.6 overloaded, 2.5-flash retired (404) → user must see the real cause, not the 404
const liteThen404 = Object.assign({ 'gemini-3.6-flash': 'overload', 'gemini-2.5-flash': 'gone' }, bothOverloaded);
check('g: retired model in chain, last resort also overloaded -> error shows overload, not 404', () => {
  const c = makeCtx(Object.assign({ 'gemini-3.5-flash': 'overload' }, liteThen404)); const r = run(c, keyAll);
  assert.ok(r.err, 'should throw');
  assert.ok(!/no longer available/.test(r.err), r.err);
  assert.ok(/high demand/.test(r.err), r.err);
});
// (h) retired model opens a long circuit (all keys) — next call skips it without fetching
check('h: retired model circuit is long-lived', () => {
  const viaFetch = makeCtx({ 'gemini-3.5-flash-lite': 'gone' }); const ttls = [];
  const rc = viaFetch.CacheService.getScriptCache;
  viaFetch.CacheService.getScriptCache = () => { const o = rc(); const p = o.put; o.put = (k, v, t) => { if (/^model_down:/.test(k)) ttls.push(t); return p(k, v, t); }; return o; };
  viaFetch.geminiFetchOnce_('gemini-3.5-flash-lite', 'K', {}, false);
  assert.ok(ttls.length === 1 && ttls[0] >= 3600, 'model_down ttl=' + ttls);
});
// (i) lite + 3.6 overloaded, 2.5 retired → one last-resort fetch on 3.5-flash serves the answer
check('i: last resort 3.5-flash serves after lite + emergency fail', () => {
  const c = makeCtx(liteThen404); const r = run(c, keyAll);
  assert.strictEqual(r.text, 'answer from gemini-3.5-flash', String(r.err));
  assert.strictEqual(c.fetched.filter(m => m === 'gemini-3.5-flash').length, 1);
  assert.strictEqual(c.fetched[c.fetched.length - 1], 'gemini-3.5-flash', 'last resort not last: ' + c.fetched);
});
// (j) 2.5 already circuit-open from an earlier 404 → skipped without fetch, 3.5-flash still reached
check('j: retired model circuit-open -> skipped, last resort reached', () => {
  const c = makeCtx(Object.assign({ 'gemini-3.6-flash': 'overload' }, bothOverloaded), { 'model_down:gemini-2.5-flash': '1' });
  const r = run(c, keyAll);
  assert.strictEqual(r.text, 'answer from gemini-3.5-flash', String(r.err));
  assert.ok(!c.fetched.includes('gemini-2.5-flash'), 'fetched retired: ' + c.fetched);
});
console.log(pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
