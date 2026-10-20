const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = process.argv[2] || path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dashboard.html'), 'utf8');
const names = ['usd', 'signedUsd', 'num', 'validateDashboardPayload'];
const context = {};
vm.createContext(context);
for (const name of names) {
  const match = html.match(new RegExp(`    function ${name}\\([^]*?\\n    }`));
  assert.ok(match, `Missing ${name}`);
  vm.runInContext(match[0], context);
}
for (const name of ['usd', 'signedUsd', 'num']) {
  for (const value of [null, undefined, '', NaN, Infinity, -Infinity]) assert.equal(context[name](value), '-');
}
assert.equal(context.usd(0), '$0');
assert.equal(context.signedUsd(0), '$0');
assert.equal(context.num(0), '0');
assert.equal(context.signedUsd(-100), '-$100');
const zero = {qty20:0,qty40:0,teu:0,expected:0,actual:0,gap:0};
const payload = row => ({rows:[row],meta:{targetRows:1,sourceModifiedAt:'2026-10-20T09:44:54+09:00'}});
context.validateDashboardPayload(payload(zero));
context.validateDashboardPayload({rows:[],meta:{targetRows:0,sourceModifiedAt:'2026-10-20T09:44:54+09:00'}});
for (const field of Object.keys(zero)) {
  for (const value of [null, undefined, '', NaN, Infinity]) {
    assert.throws(() => context.validateDashboardPayload(payload({...zero,[field]:value})), /missing or not numeric/);
  }
}
assert.throws(() => context.validateDashboardPayload({rows:[],meta:{targetRows:1}}), /row count/);
assert.throws(() => context.validateDashboardPayload({rows:[zero],meta:{targetRows:1}}), /source timestamp/);
assert.throws(() => context.validateDashboardPayload(payload({...zero,actual:10,gap:0})), /does not reconcile/);
assert.throws(() => context.validateDashboardPayload({meta:{targetRows:0}}), /rows and metadata/);
const liveFile = process.argv[3];
if (liveFile) {
  const live = JSON.parse(fs.readFileSync(liveFile, 'utf8'));
  context.validateDashboardPayload(live);
  console.log(`PASS: missing versus zero, malformed payloads, row count, timestamp, gap reconciliation; ${live.rows.length} actual published rows accepted`);
} else console.log('PASS: missing versus zero, malformed payloads, row count, timestamp and gap reconciliation');
