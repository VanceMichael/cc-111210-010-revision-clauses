import test from 'node:test';
import assert from 'node:assert/strict';
import { parts } from '../src/connector.mjs';
import { archive } from '../fixtures/source.mjs';
import { collectClauses, compareCollections } from '../src/clauses.mjs';

const files = kind => parts(archive(kind));
const clause = (collection, paraId) => collection.clauses.find(item => item.paraId === paraId);

test('接受视图只保留拟生效内容，删除与移出不进入文本', () => {
  const collection = collectClauses(files(), { view: 'accept' });
  const fee = clause(collection, '0000000A');
  assert.equal(fee.text, '服务费120元');
  assert.equal(fee.number.text, '5.');
  assert.ok(!collection.clauses.some(item => item.text.includes('100元')));
  assert.equal(clause(collection, '0000000B').text, '');
  const cell = clause(collection, '0000000C');
  assert.equal(cell.text, '进场时间');
  assert.equal(cell.inTable, true);
  assert.deepEqual(cell.footnotes, [{ id: '2', text: '另见进场规则' }]);
  assert.deepEqual(collection.review, []);
});

test('拒绝视图保留修订前原文，插入与移入不进入文本', () => {
  const collection = collectClauses(files(), { view: 'reject' });
  assert.equal(clause(collection, '0000000A').text, '服务费100元');
  assert.ok(!collection.clauses.some(item => item.text.includes('120元')));
  assert.equal(clause(collection, '0000000B').text, '进场时间');
  assert.equal(clause(collection, '0000000C').text, '');
});

test('待审视图并录双方内容且携带修订身份', () => {
  const collection = collectClauses(files(), { view: 'pending' });
  const fee = clause(collection, '0000000A');
  assert.deepEqual(fee.segments.map(segment => [segment.text, segment.status]), [
    ['服务费', 'unchanged'],
    ['100元', 'deleted'],
    ['120元', 'inserted'],
  ]);
  assert.equal(fee.segments[1].rev.id, '11');
  assert.equal(fee.segments[2].rev.id, '12');
  assert.equal(fee.segments[2].rev.author, '法务');
  const from = clause(collection, '0000000B').segments[0];
  const to = clause(collection, '0000000C').segments.find(segment => segment.status === 'moved-to');
  assert.equal(from.move.key, to.move.key);
  assert.equal(from.move.paired, true);
  assert.deepEqual(clause(collection, '0000000B').links[0].peers, [clause(collection, '0000000C').key]);
});

test('条款编号依据抽象编号与覆盖规则而非正文数字', () => {
  const collection = collectClauses(files(), { view: 'accept' });
  const fee = clause(collection, '0000000A');
  assert.equal(fee.number.value, 5);
  assert.deepEqual(fee.number.basis, {
    numId: 4,
    abstractNumId: 1,
    ilvl: 0,
    start: 1,
    startOverride: 5,
    numFmt: 'decimal',
    lvlText: '%1.',
  });
  assert.equal(clause(collection, '0000000B').number, null);
});

test('页眉重复文本不作为正文条款采集', () => {
  const collection = collectClauses(files(), { view: 'accept' });
  assert.ok(!collection.clauses.some(item => item.text.includes('服务费摘要')));
  assert.ok(collection.excludedParts.some(item => item.part === 'word/header1.xml'));
});

test('移动配对不完整时交人工复核而不是猜测归属', () => {
  const collection = collectClauses(files('unpaired'), { view: 'pending' });
  const entries = collection.review.filter(item => item.type === 'move-unpaired');
  assert.equal(entries.length, 2);
  const involved = entries.flatMap(entry => entry.clauses);
  assert.ok(involved.includes(clause(collection, '0000000B').key));
  assert.ok(involved.includes(clause(collection, '0000000C').key));
  assert.equal(clause(collection, '0000000B').segments[0].move.paired, false);
  assert.equal(clause(collection, '0000000B').links[0].peers.length, 0);
});

test('重新编号后条款身份与内容不变，仅编号变化', () => {
  const before = collectClauses(files(), { view: 'accept' });
  const after = collectClauses(files('renumbered'), { view: 'accept' });
  assert.equal(clause(after, '0000000A').number.text, '9.');
  assert.equal(clause(after, '0000000A').text, clause(before, '0000000A').text);
  const diff = compareCollections(before, after);
  assert.equal(diff.summary.textChanged, 0);
  assert.equal(diff.summary.renumberedOnly, 1);
  const fee = diff.clauses.find(entry => entry.key === clause(before, '0000000A').key);
  assert.equal(fee.textChanged, false);
  assert.equal(fee.numberChanged, true);
});

test('三种视图身份一致但内容不混，非法视图被拒绝', () => {
  const accept = collectClauses(files(), { view: 'accept' });
  const reject = collectClauses(files(), { view: 'reject' });
  assert.equal(clause(accept, '0000000A').key, clause(reject, '0000000A').key);
  assert.notEqual(clause(accept, '0000000A').text, clause(reject, '0000000A').text);
  assert.throws(() => collectClauses(files(), { view: 'final' }), /视图/);
  assert.throws(() => collectClauses(files(), {}), /视图/);
});

test('编号部件或脚注目标缺失时交人工复核', () => {
  const noNumbering = files();
  delete noNumbering['word/numbering.xml'];
  const collection = collectClauses(noNumbering, { view: 'accept' });
  assert.equal(clause(collection, '0000000A').number, null);
  assert.ok(collection.review.some(item => item.type === 'relationship-missing'));
  assert.ok(collection.review.some(item => item.type === 'numbering-missing'));
  const noFootnotes = files();
  delete noFootnotes['word/footnotes.xml'];
  const orphaned = collectClauses(noFootnotes, { view: 'accept' });
  assert.ok(orphaned.review.some(item => item.type === 'footnote-unresolved'));
});

test('跨段落移动按名称配对，范围未闭合交人工复核', () => {
  const document = moves => `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${moves}</w:body></w:document>`;
  const synthetic = moves => ({
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml': document(moves),
  });
  const paired = collectClauses(synthetic(
    '<w:p><w:moveFromRangeStart w:name="m1"/><w:moveFrom w:id="1" w:name="m1"><w:r><w:t>甲</w:t></w:r></w:moveFrom></w:p>'
    + '<w:p><w:moveFrom w:id="2" w:name="m1"><w:r><w:t>乙</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:name="m1"/></w:p>'
    + '<w:p><w:moveToRangeStart w:name="m1"/><w:moveTo w:id="3" w:name="m1"><w:r><w:t>甲</w:t></w:r></w:moveTo></w:p>'
    + '<w:p><w:moveTo w:id="4" w:name="m1"><w:r><w:t>乙</w:t></w:r></w:moveTo><w:moveToRangeEnd w:name="m1"/></w:p>',
  ), { view: 'accept' });
  assert.deepEqual(paired.review, []);
  assert.deepEqual(paired.clauses.map(item => item.text), ['', '', '甲', '乙']);
  const unclosed = collectClauses(synthetic(
    '<w:p><w:moveToRangeStart w:name="m2"/><w:moveTo w:id="5" w:name="m2"><w:r><w:t>丙</w:t></w:r></w:moveTo></w:p>',
  ), { view: 'pending' });
  assert.ok(unclosed.review.some(item => item.type === 'move-range-unclosed'));
  assert.ok(unclosed.review.some(item => item.type === 'move-unpaired'));
});
