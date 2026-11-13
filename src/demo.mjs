import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startFixture } from '../fixtures/source.mjs';
import { startService } from './service.mjs';
import { compareCollections } from './clauses.mjs';

const directory = await mkdtemp(join(tmpdir(), 'capture-demo-'));
const source = await startFixture();
let service;
try {
  service = await startService({ source: source.url, directory });
  const get = async path => {
    const response = await fetch(service.url + path);
    const body = await response.json();
    if (!response.ok) throw new Error(`${path} → ${response.status}: ${JSON.stringify(body)}`);
    return body;
  };
  const brief = clause => ({ key: clause.key, number: clause.number?.text ?? null, text: clause.text, inTable: clause.inTable, footnotes: clause.footnotes });

  const accept = await get('/collect?view=accept');
  console.log('=== 接受视图（拟生效条款，交登记系统） ===');
  console.log(JSON.stringify({ excludedParts: accept.excludedParts, clauses: accept.clauses.map(brief) }, null, 2));

  const [reject, pending] = [await get('/collect?view=reject'), await get('/collect?view=pending')];
  const key = 'word/document.xml#0000000A';
  const pick = collection => collection.clauses.find(clause => clause.key === key);
  console.log('=== 同一条款在三种修订视图下的内容与身份（身份不变，内容不混） ===');
  console.log(JSON.stringify({
    key,
    paraId: pick(accept).paraId,
    accept: pick(accept).text,
    reject: pick(reject).text,
    pending: pick(pending).segments,
  }, null, 2));

  const moved = pending.clauses.find(clause => clause.paraId === '0000000B');
  console.log('=== 移动条款在待审视图中的配对身份（来源段落 ↔ 目标单元格） ===');
  console.log(JSON.stringify({ from: moved.key, links: moved.links }, null, 2));

  const renumbered = await get('/collect?view=accept&kind=renumbered');
  const diff = compareCollections(accept, renumbered);
  console.log('=== 重新编号前后比对（内容未变，仅编号 5. → 9.） ===');
  console.log(JSON.stringify({ changed: diff.clauses.filter(entry => entry.numberChanged || entry.textChanged), summary: diff.summary }, null, 2));

  const unpaired = await get('/collect?view=pending&kind=unpaired');
  console.log('=== 移动配对不完整 → 不猜测归属，交人工复核 ===');
  console.log(JSON.stringify(unpaired.review, null, 2));

  const receipt = await get('/receive?key=effective&view=accept');
  console.log('=== 拟生效条款接收留存（同一 key 不能以其他视图覆盖） ===');
  console.log(JSON.stringify({ key: receipt.key, digest: receipt.digest, view: receipt.value.view }, null, 2));
} finally {
  await service?.close();
  await source.close();
  await rm(directory, { recursive: true, force: true });
}
