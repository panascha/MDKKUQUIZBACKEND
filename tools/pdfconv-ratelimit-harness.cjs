// Stub harness: convertPdfBatch rate-limit gate in router-doPost.gs (real source text, stubbed GAS services).
// Run: node tools/pdfconv-ratelimit-harness.cjs   (from MDKKUQUIZBACKEND/)
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const router = fs.readFileSync(path.join(root, 'router-doPost.gs'), 'utf8');
const rlSrc = fs.readFileSync(path.join(root, 'intelsphere.gs'), 'utf8').match(/function checkActionRateLimit[\s\S]*?\n}/)[0];
const gate = router.slice(router.indexOf("if (action === 'convertPdfBatch') {") + "if (action === 'convertPdfBatch') {".length, router.indexOf('var pcUser = null;'));

function makeGate(roles) {
  const store = {};
  const CacheService = { getScriptCache: () => ({ get: k => store[k] || null, put: (k, v) => { store[k] = v; } }) };
  const ContentService = { createTextOutput: s => ({ setMimeType: () => JSON.parse(s) }), MimeType: { JSON: 'json' } };
  const verifySessionToken = t => (roles[t] ? { role: roles[t] } : null);
  const requireRole_ = (u, r) => !!u && String(u.role).trim() === r;
  // eslint-disable-next-line no-new-func
  return new Function('CacheService', 'ContentService', 'verifySessionToken', 'requireRole_',
    rlSrc + '\nreturn function(data){' + gate + '\nreturn "PASS";};')(CacheService, ContentService, verifySessionToken, requireRole_);
}
const run = (g, data, n) => { let pass = 0; for (let i = 0; i < n; i++) if (g(data) === 'PASS') pass++; return pass; };

const g = makeGate({ DEV: 'DEVELOPER', ADM: 'Admin' });
const cases = [
  ['garbage token capped at 80', run(g, { sessionToken: 'junk' }, 100), 80],
  ['Admin (non-DEVELOPER) capped at 80', run(g, { sessionToken: 'ADM' }, 100), 80],
  ['DEVELOPER gets 80 + 600', run(g, { sessionToken: 'DEV' }, 1000), 680],
  ['username-only (no token) capped at 80', run(g, { username: 'u' }, 100), 80],
];
let fail = 0;
for (const [name, got, want] of cases) { const ok = got === want; if (!ok) fail++; console.log((ok ? 'PASS ' : 'FAIL ') + name + ` (got ${got}, want ${want})`); }
process.exitCode = fail ? 1 : 0;
