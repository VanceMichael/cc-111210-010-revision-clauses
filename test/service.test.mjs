import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { ReceiptStore } from '../src/store.mjs';
import { startService } from '../src/service.mjs';
import { startFixture } from '../fixtures/source.mjs';

test('接收记录支持同内容重复领取并拒绝覆盖', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'receipt-'));
  try {
    const store = new ReceiptStore(directory);
    assert.deepEqual(await store.save('one', { value: 3 }), await store.save('one', { value: 3 }));
    await assert.rejects(store.save('one', { value: 4 }), /冲突/);
    await assert.rejects(store.save('../escape', {}), /无效/);
    const receipt = JSON.parse(await readFile(join(directory, 'one.json'), 'utf8'));
    receipt.value.value = 8;
    await writeFile(join(directory, 'one.json'), JSON.stringify(receipt));
    await assert.rejects(store.read('one'), /校验失败/);
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test('服务健康检查和单次查看走真实来源', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'service-'));
  const fixture = await startFixture();
  const service = await startService({ source: fixture.url, directory });
  try {
    assert.equal((await (await fetch(service.url + '/health')).json()).state, 'ready');
    const response = await fetch(service.url + '/inspect');
    assert.equal(response.status, 200, await response.text());
    assert.equal((await fetch(service.url + '/missing')).status, 404);
  } finally { await service.close(); await fixture.close(); await rm(directory, { recursive: true, force: true }); }
});

test('条款采集视图必填且不同视图留存互不覆盖', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'service-'));
  const fixture = await startFixture();
  const service = await startService({ source: fixture.url, directory });
  try {
    assert.equal((await fetch(service.url + '/collect')).status, 400);
    assert.equal((await fetch(service.url + '/collect?view=final')).status, 400);
    const accept = await (await fetch(service.url + '/collect?view=accept')).json();
    assert.equal(accept.view, 'accept');
    assert.ok(accept.package.sha256);
    const fee = accept.clauses.find(clause => clause.paraId === '0000000A');
    assert.equal(fee.text, '服务费120元');
    assert.equal(fee.number.text, '5.');
    const saved = await fetch(service.url + '/receive?key=effective&view=accept');
    assert.equal(saved.status, 200);
    assert.equal((await fetch(service.url + '/receive?key=effective&view=accept')).status, 200);
    const conflict = await fetch(service.url + '/receive?key=effective&view=reject');
    assert.equal(conflict.status, 400);
    assert.match((await conflict.json()).error, /冲突/);
  } finally { await service.close(); await fixture.close(); await rm(directory, { recursive: true, force: true }); }
});
