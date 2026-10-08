// Offline harness for Phase 3 step 3 mirror helpers. Run: node tools/sb-phase3-harness.js (from repo root)
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');
const root = path.join(__dirname, '..');
for (const f of ['router-doPost.gs', 'supabase-mirror.gs', 'discussion.gs']) {
  new vm.Script(fs.readFileSync(path.join(root, f), 'utf8'), { filename: f }); // syntax check
}

function makeCtx(withKeys, fetchImpl) {
  const calls = [], logs = [], cache = {};
  const ctx = {
    console, calls, logs,
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => withKeys ? { SUPABASE_URL: 'https://x.supabase.co', SUPABASE_SERVICE_KEY: 'k' }[k] : null }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
    Utilities: {
      computeDigest: () => [1, 2, 3], base64EncodeWebSafe: () => 'ABC',
      DigestAlgorithm: { MD5: 'MD5' }
    },
    UrlFetchApp: { fetch: (url, opts) => { calls.push({ url, opts }); return fetchImpl(url, opts); } },
    writeAdminLog: (...a) => logs.push(a),
    execBudgetExhausted_: () => false,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(root, 'supabase-mirror.gs'), 'utf8'), ctx);
  return ctx;
}
const ok = body => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify(body) });
const bad = code => ({ getResponseCode: () => code, getContentText: () => 'boom' });

// 1. no keys -> no network, null
let c = makeCtx(false, () => { throw new Error('should not fetch'); });
assert.strictEqual(c.sbMirrorProgress_('a@b', 's', 5, {}), null);
assert.strictEqual(c.sbMirrorFeedback_('Bug', 'd', 'a', 'c', 'x', []), null);
c.sbBackfillProgress_('a@b', 's', 5, {});
assert.strictEqual(c.calls.length, 0);

// 2. progress named args
c = makeCtx(true, () => ok({ result: 'success', timestamp: 5 }));
c.sbMirrorProgress_('a@b', 's', 5, { timestamp: 5 });
let b = JSON.parse(c.calls[0].opts.payload);
assert.ok(c.calls[0].url.endsWith('/rest/v1/rpc/save_progress'));
assert.deepStrictEqual(b, { p_email: 'a@b', p_subject: 's', p_ts: 5, p_state: { timestamp: 5 } });

// 3. stale is not logged as failure
c = makeCtx(true, () => ok({ result: 'stale', cloudTimestamp: 9 }));
c.sbMirrorProgress_('a@b', 's', 5, {});
assert.strictEqual(c.logs.length, 0);

// 4. HTTP failure logged, no throw
c = makeCtx(true, () => bad(500));
assert.doesNotThrow(() => c.sbMirrorFeedback_('Bug', 'd', 'a', 'c', 'x', ['u']));
assert.strictEqual(c.logs.length, 1);
// network throw swallowed
c = makeCtx(true, () => { throw new Error('net'); });
assert.doesNotThrow(() => c.sbMirrorDiscussionPost_('q', 'e', 'n', 't', 'txt'));
assert.strictEqual(c.logs.length, 1);

// 5. backfill: once per flag
c = makeCtx(true, () => ok({ result: 'success' }));
c.sbBackfillProgress_('a@b', 's', 5, {});
c.sbBackfillProgress_('a@b', 's', 5, {});
assert.strictEqual(c.calls.length, 1);
// backfill failure -> flag not set -> retried
c = makeCtx(true, () => bad(500));
c.sbBackfillProgress_('a@b', 's', 5, {});
c.sbBackfillProgress_('a@b', 's', 5, {});
assert.strictEqual(c.calls.length, 2);

// 6. discussion post args
c = makeCtx(true, () => ok({ ok: true }));
c.sbMirrorDiscussionPost_('q1', 'e', 'nick', 'ab12', 'hi');
b = JSON.parse(c.calls[0].opts.payload);
assert.deepStrictEqual(b, { p_question_id: 'q1', p_email: 'e', p_nickname: 'nick', p_tag: 'ab12', p_content: 'hi', p_tags: [], p_references: null });

// 7. delete/status: lookup then rpc
const pg = { qid: 'q1', email: 'e', text: 'hi', ts: '2026-10-08T10:00:00.000Z' };
c = makeCtx(true, (url) => url.includes('/rest/v1/discussions')
  ? ok([{ id: 'near', created_at: '2026-10-08T10:00:01.000Z' }])
  : ok({ ok: true }));
c.sbMirrorDiscussionDelete_(pg, 'e', false);
assert.strictEqual(c.calls.length, 2);
assert.strictEqual(c.calls[0].opts.method, 'get');
assert.deepStrictEqual(JSON.parse(c.calls[1].opts.payload), { p_id: 'near', p_requestor_email: 'e', p_is_admin: false });
c.sbMirrorDiscussionStatus_(pg, 'pinned');
assert.deepStrictEqual(JSON.parse(c.calls[3].opts.payload), { p_id: 'near', p_status: 'pinned', p_is_admin: true });
// lookup url filters undeleted rows
assert.ok(c.calls[0].url.includes('deleted_at=is.null'));
// ambiguous (2 rows) -> no rpc, logged
c = makeCtx(true, () => ok([{ id: 'a', created_at: '2026-10-08T10:00:01.000Z' }, { id: 'b', created_at: '2026-10-08T10:01:30.000Z' }]));
c.sbMirrorDiscussionDelete_(pg, 'e', true);
c.sbMirrorDiscussionStatus_(pg, 'pinned');
assert.strictEqual(c.calls.length, 2);
assert.strictEqual(c.logs.length, 2);
// lookup miss -> no rpc
c = makeCtx(true, () => ok([]));
c.sbMirrorDiscussionDelete_(pg, 'e', true);
assert.strictEqual(c.calls.length, 1);
// missing pg -> nothing
c.sbMirrorDiscussionStatus_(undefined, 'pinned');
assert.strictEqual(c.calls.length, 1);

console.log('sb-phase3-harness: all assertions passed');
