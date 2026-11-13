import {
  NS_MAIN, NS_W14, parseXml, localName, qn, elements, textOf, attr,
  rootOf, prefixFor, relationships, relsPartName, resolveTarget,
} from './ooxml.mjs';

// 三种修订视图互不相同、不可混用：accept=拟生效（接受全部修订），reject=修订前原文（拒绝全部修订），pending=保留待审（双方并录）
export const VIEWS = ['accept', 'reject', 'pending'];

const KEEP = {
  accept: new Set(['unchanged', 'inserted', 'moved-to']),
  reject: new Set(['unchanged', 'deleted', 'moved-from']),
  pending: new Set(['unchanged', 'inserted', 'deleted', 'moved-from', 'moved-to']),
};

const REVISION_STATUS = { ins: 'inserted', del: 'deleted', moveFrom: 'moved-from', moveTo: 'moved-to' };
const RANGE_MARKS = {
  moveFromRangeStart: { side: 'from', mark: 'start' },
  moveFromRangeEnd: { side: 'from', mark: 'end' },
  moveToRangeStart: { side: 'to', mark: 'start' },
  moveToRangeEnd: { side: 'to', mark: 'end' },
};

export function collectClauses(files, options = {}) {
  const { view, packageHash = null } = options;
  if (!VIEWS.includes(view)) throw new Error(`修订视图须为 ${VIEWS.join(' / ')} 之一，收到：${view ?? '（缺失）'}`);
  const review = [];
  const partNames = Object.keys(files).sort();

  // 包关系定位正文部件；关系缺失时回退默认路径并交人工复核
  let documentPart = 'word/document.xml';
  if (files['_rels/.rels']) {
    const rel = relationships(files['_rels/.rels']).find(item => item.type?.endsWith('/officeDocument'));
    if (rel) documentPart = resolveTarget('', rel.target);
    else review.push({ type: 'relationship-missing', detail: '包关系缺少 officeDocument，回退默认正文部件 word/document.xml', clauses: [] });
  } else {
    review.push({ type: 'relationship-missing', detail: '缺少 _rels/.rels，回退默认正文部件 word/document.xml', clauses: [] });
  }
  if (!files[documentPart]) throw new Error('正文部件缺失');

  const relsPart = relsPartName(documentPart);
  const documentRels = files[relsPart] ? relationships(files[relsPart]) : [];
  const relTarget = suffix => {
    const rel = documentRels.find(item => item.type?.endsWith(suffix));
    return rel ? resolveTarget(documentPart, rel.target) : null;
  };

  // 编号部件：条款编号只依据抽象编号与覆盖规则，不从正文数字推断
  let numbering = null;
  const numberingPart = relTarget('/numbering');
  if (numberingPart) {
    if (files[numberingPart]) numbering = parseNumbering(files[numberingPart]);
    else review.push({ type: 'relationship-missing', detail: `编号关系指向缺失部件 ${numberingPart}`, clauses: [] });
  }

  const footnotesPart = relTarget('/footnotes');
  let footnotes = null;
  if (footnotesPart) {
    if (files[footnotesPart]) footnotes = parseFootnotes(files[footnotesPart], view, review);
    else review.push({ type: 'relationship-missing', detail: `脚注关系指向缺失部件 ${footnotesPart}`, clauses: [] });
  }

  // 页眉页脚是版式重复文本，不属于正文条款
  const excludedParts = partNames
    .filter(name => /(^|\/)header\d*\.xml$/.test(name) || /(^|\/)footer\d*\.xml$/.test(name))
    .map(part => ({ part, reason: '页眉/页脚为版式重复文本，不属于正文条款' }));

  const { clauses, ranges } = parseDocument(files[documentPart], documentPart, review);

  markDuplicateIdentities(clauses, review);
  pairMoves(clauses, ranges, review);
  const viewClauses = applyView(clauses, view, review);
  applyNumbering(viewClauses, numbering, review);
  applyFootnotes(viewClauses, footnotes, review);

  return {
    view,
    package: { sha256: packageHash, parts: partNames },
    documentPart,
    excludedParts,
    clauses: viewClauses.map(finalizeClause),
    review,
  };
}

// 跨版本比对：按条款身份（key）对齐，区分内容变化与仅重新编号
export function compareCollections(a, b) {
  const mapA = new Map(a.clauses.map(clause => [clause.key, clause]));
  const mapB = new Map(b.clauses.map(clause => [clause.key, clause]));
  const keys = [...a.clauses.map(clause => clause.key), ...b.clauses.filter(clause => !mapA.has(clause.key)).map(clause => clause.key)];
  const clauses = keys.map(key => {
    const ca = mapA.get(key);
    const cb = mapB.get(key);
    const entry = { key, inA: Boolean(ca), inB: Boolean(cb) };
    if (ca && cb) {
      entry.textChanged = ca.text !== cb.text;
      entry.numberChanged = (ca.number?.text ?? null) !== (cb.number?.text ?? null);
      entry.a = { text: ca.text, number: ca.number?.text ?? null };
      entry.b = { text: cb.text, number: cb.number?.text ?? null };
    }
    return entry;
  });
  return {
    views: [a.view, b.view],
    clauses,
    summary: {
      onlyInA: clauses.filter(entry => entry.inA && !entry.inB).length,
      onlyInB: clauses.filter(entry => !entry.inA && entry.inB).length,
      textChanged: clauses.filter(entry => entry.textChanged).length,
      renumberedOnly: clauses.filter(entry => entry.numberChanged && !entry.textChanged).length,
      unchanged: clauses.filter(entry => entry.inA && entry.inB && !entry.textChanged && !entry.numberChanged).length,
    },
  };
}

function parseDocument(xml, part, review) {
  const { root, namespaces } = rootOf(parseXml(xml), 'document');
  const w = prefixFor(namespaces, NS_MAIN);
  if (!w) throw new Error('正文部件未声明 OOXML 主命名空间');
  const ctx = { w, w14: prefixFor(namespaces, NS_W14), part, review, ranges: [] };
  const body = elements(root.children).find(element => element.tag === qn(ctx.w, 'body'));
  if (!body) throw new Error('正文部件缺少 w:body');
  const clauses = [];
  walkBlocks(body.children, 'w:body', ctx, clauses, false);
  return { clauses, ranges: ctx.ranges };
}

function walkBlocks(nodes, pathPrefix, ctx, clauses, inTable) {
  const counters = new Map();
  for (const element of elements(nodes)) {
    const local = localName(element.tag);
    const index = (counters.get(local) ?? -1) + 1;
    counters.set(local, index);
    const path = `${pathPrefix}/${element.tag}[${index}]`;
    if (element.tag === qn(ctx.w, 'p')) {
      clauses.push(parseParagraph(element, path, ctx, inTable));
    } else if (element.tag === qn(ctx.w, 'tbl')) {
      walkTable(element, path, ctx, clauses);
    } else if (RANGE_MARKS[local] && element.tag === qn(ctx.w, local)) {
      ctx.ranges.push({ name: attr(element, qn(ctx.w, 'name')), ...RANGE_MARKS[local], path });
    } else if (element.tag === qn(ctx.w, 'sectPr') || element.tag === qn(ctx.w, 'bookmarkStart') || element.tag === qn(ctx.w, 'bookmarkEnd')) {
      continue;
    } else {
      // 非标准容器（如 sdt）：深入取段落，但归属交人工确认
      const before = clauses.length;
      walkBlocks(element.children, path, ctx, clauses, inTable);
      if (clauses.length > before) {
        ctx.review.push({
          type: 'nonstandard-container',
          detail: `正文出现非标准容器 ${element.tag}（${path}），其中段落的条款归属需人工确认`,
          clauses: clauses.slice(before).map(clause => clause.key),
        });
      }
    }
  }
}

function walkTable(table, path, ctx, clauses) {
  const rows = elements(table.children).filter(element => element.tag === qn(ctx.w, 'tr'));
  rows.forEach((row, rowIndex) => {
    const cells = elements(row.children).filter(element => element.tag === qn(ctx.w, 'tc'));
    cells.forEach((cell, cellIndex) => {
      walkBlocks(cell.children, `${path}/${row.tag}[${rowIndex}]/${cell.tag}[${cellIndex}]`, ctx, clauses, true);
    });
  });
}

function parseParagraph(element, path, ctx, inTable) {
  const paraId = ctx.w14 ? attr(element, qn(ctx.w14, 'paraId')) : null;
  const clause = {
    key: `${ctx.part}#${paraId ?? path}`,
    part: ctx.part,
    paraId,
    identityBasis: paraId ? 'paraId' : 'path',
    path,
    inTable,
    numPr: null,
    mark: null,
    segments: [],
    links: [],
  };
  for (const child of elements(element.children)) {
    if (child.tag === qn(ctx.w, 'pPr')) {
      const props = elements(child.children);
      const numPrElement = props.find(item => item.tag === qn(ctx.w, 'numPr'));
      if (numPrElement) clause.numPr = parseNumPr(numPrElement, ctx);
      const rPr = props.find(item => item.tag === qn(ctx.w, 'rPr'));
      if (rPr) clause.mark = parseParagraphMark(rPr, ctx, path) ?? clause.mark;
    } else {
      collectRunSegments([child], null, ctx, clause.segments, path);
    }
  }
  return clause;
}

function parseNumPr(numPrElement, ctx) {
  const children = elements(numPrElement.children);
  const ilvl = children.find(item => item.tag === qn(ctx.w, 'ilvl'));
  const numId = children.find(item => item.tag === qn(ctx.w, 'numId'));
  const ilvlValue = ilvl ? attr(ilvl, qn(ctx.w, 'val')) : null;
  return {
    ilvl: ilvlValue != null ? Number(ilvlValue) : 0,
    numId: numId ? attr(numId, qn(ctx.w, 'val')) : null,
  };
}

function parseParagraphMark(rPr, ctx, path) {
  for (const element of elements(rPr.children)) {
    if (element.tag === qn(ctx.w, 'ins')) return { status: 'inserted', rev: revisionInfo(element, ctx) };
    if (element.tag === qn(ctx.w, 'del')) return { status: 'deleted', rev: revisionInfo(element, ctx) };
    if (element.tag === qn(ctx.w, 'moveFrom') || element.tag === qn(ctx.w, 'moveTo')) {
      ctx.review.push({ type: 'paragraph-mark-move', detail: `段落 ${path} 的段落标记参与移动，合并与拆分行为需人工确认`, clauses: [] });
    }
  }
  return null;
}

function collectRunSegments(elementList, revision, ctx, out, path) {
  for (const element of elementList) {
    const local = localName(element.tag);
    if (element.tag === qn(ctx.w, 'r')) {
      out.push(runSegment(element, revision, ctx));
    } else if (REVISION_STATUS[local] && element.tag === qn(ctx.w, local)) {
      const rev = revisionInfo(element, ctx);
      const move = local === 'moveFrom' || local === 'moveTo'
        ? { side: local === 'moveFrom' ? 'from' : 'to', key: moveKey(element, ctx), id: rev.id ?? null, name: attr(element, qn(ctx.w, 'name')) }
        : null;
      collectRunSegments(elements(element.children), { status: REVISION_STATUS[local], rev, move }, ctx, out, path);
    } else if (RANGE_MARKS[local] && element.tag === qn(ctx.w, local)) {
      ctx.ranges.push({ name: attr(element, qn(ctx.w, 'name')), ...RANGE_MARKS[local], path });
    } else if (element.tag === qn(ctx.w, 'txbxContent')) {
      ctx.review.push({ type: 'textbox-content', detail: `文本框内容（${path}）不参与正文顺序，未自动采集，需人工处理`, clauses: [] });
    } else {
      collectRunSegments(elements(element.children), revision, ctx, out, path); // 透明包装：hyperlink、sdt、smartTag 等
    }
  }
}

function runSegment(run, revision, ctx) {
  let text = '';
  const footnoteRefs = [];
  for (const element of elements(run.children)) {
    if (element.tag === qn(ctx.w, 't') || element.tag === qn(ctx.w, 'delText')) {
      text += textOf(element.children);
    } else if (element.tag === qn(ctx.w, 'tab')) {
      text += '\t';
    } else if (element.tag === qn(ctx.w, 'br') || element.tag === qn(ctx.w, 'cr')) {
      text += '\n';
    } else if (element.tag === qn(ctx.w, 'footnoteReference')) {
      footnoteRefs.push(attr(element, qn(ctx.w, 'id')));
    } else if (element.tag === qn(ctx.w, 'drawing') || element.tag === qn(ctx.w, 'pict')) {
      if (containsTag(element.children, ctx.w, 'txbxContent')) {
        ctx.review.push({ type: 'textbox-content', detail: '绘图文本框内容未自动采集，需人工处理', clauses: [] });
      }
    }
  }
  const segment = { text, status: revision?.status ?? 'unchanged', footnoteRefs };
  if (revision?.rev) segment.rev = revision.rev;
  if (revision?.move) segment.move = revision.move;
  return segment;
}

function containsTag(nodes, w, local) {
  for (const element of elements(nodes)) {
    if (element.tag === qn(w, local)) return true;
    if (containsTag(element.children, w, local)) return true;
  }
  return false;
}

function revisionInfo(element, ctx) {
  const info = {};
  const id = attr(element, qn(ctx.w, 'id'));
  const author = attr(element, qn(ctx.w, 'author'));
  const date = attr(element, qn(ctx.w, 'date'));
  if (id != null) info.id = id;
  if (author != null) info.author = author;
  if (date != null) info.date = date;
  return info;
}

// 移动配对键：真实文档按 w:name 关联跨段落的同一移动，样本按修订 w:id 关联
function moveKey(element, ctx) {
  const name = attr(element, qn(ctx.w, 'name'));
  const id = attr(element, qn(ctx.w, 'id'));
  return name != null ? `name:${name}` : `id:${id}`;
}

function markDuplicateIdentities(clauses, review) {
  const seen = new Map();
  for (const clause of clauses) {
    if (clause.paraId == null) continue;
    const first = seen.get(clause.paraId);
    if (first) {
      clause.identityBasis = 'paraId-duplicate';
      clause.key = `${clause.key}@${clause.path}`;
      review.push({
        type: 'paraId-duplicate',
        detail: `段落标识 ${clause.paraId} 出现多次，跨版本身份需人工确认`,
        clauses: [first.key, clause.key],
      });
    } else {
      seen.set(clause.paraId, clause);
    }
  }
}

function pairMoves(clauses, ranges, review) {
  const groups = new Map();
  for (const clause of clauses) {
    for (const segment of clause.segments) {
      if (!segment.move) continue;
      const group = groups.get(segment.move.key) ?? { from: [], to: [] };
      group[segment.move.side].push({ clause, segment });
      groups.set(segment.move.key, group);
    }
  }
  for (const [key, group] of groups) {
    const paired = group.from.length > 0 && group.to.length > 0;
    for (const side of ['from', 'to']) {
      for (const { clause, segment } of group[side]) {
        segment.move.paired = paired;
        const peers = [...new Set(group[side === 'from' ? 'to' : 'from'].map(item => item.clause.key))];
        clause.links.push({ type: 'move', key, role: side, paired, peers });
      }
    }
    if (!paired) {
      review.push({
        type: 'move-unpaired',
        detail: `移动配对不完整（${key}）：${group.from.length ? '存在移动来源但缺少移动目标' : '存在移动目标但缺少移动来源'}，条款归属需人工复核`,
        clauses: [...new Set([...group.from, ...group.to].map(item => item.clause.key))],
        revisions: [...new Set([...group.from, ...group.to].map(item => item.segment.move.id).filter(Boolean))],
      });
    }
  }
  const rangeGroups = new Map();
  for (const range of ranges) {
    const key = `${range.side}:${range.name}`;
    const group = rangeGroups.get(key) ?? { start: 0, end: 0 };
    group[range.mark] += 1;
    rangeGroups.set(key, group);
  }
  for (const [key, group] of rangeGroups) {
    if (group.start !== group.end) {
      review.push({ type: 'move-range-unclosed', detail: `移动范围标记未闭合（${key}）：起点 ${group.start} 处、终点 ${group.end} 处`, clauses: [] });
    }
  }
}

function applyView(clauses, view, review) {
  const keep = KEEP[view];
  const output = [];
  for (const clause of clauses) {
    if (clause.mark?.status === 'inserted' && view === 'reject') continue; // 整段为插入，拒绝后不存在
    if (clause.mark?.status === 'deleted' && view === 'accept') {
      review.push({
        type: 'paragraph-mark-deleted',
        detail: `段落 ${clause.key} 的段落标记被删除，接受修订后并入相邻段落，需人工确认`,
        clauses: [clause.key],
      });
      continue;
    }
    output.push({ ...clause, segments: clause.segments.filter(segment => keep.has(segment.status)) });
  }
  return output;
}

function applyNumbering(clauses, numbering, review) {
  const counters = new Map();
  for (const clause of clauses) {
    clause.number = null;
    if (!clause.numPr || clause.numPr.numId == null) continue;
    if (!numbering) {
      review.push({ type: 'numbering-missing', detail: `条款 ${clause.key} 带有编号引用（numId=${clause.numPr.numId}）但编号部件缺失`, clauses: [clause.key] });
      continue;
    }
    clause.number = resolveNumber(clause, numbering, counters, review);
  }
}

function resolveNumber(clause, numbering, counters, review) {
  const { numId, ilvl } = clause.numPr;
  const num = numbering.nums[numId];
  if (!num) {
    review.push({ type: 'numbering-unknown-num', detail: `条款 ${clause.key} 引用的 numId=${numId} 在编号部件中不存在`, clauses: [clause.key] });
    return null;
  }
  const abstract = numbering.abstract[num.abstractNumId];
  if (!abstract) {
    review.push({ type: 'numbering-unknown-abstract', detail: `numId=${numId} 指向的 abstractNumId=${num.abstractNumId} 不存在`, clauses: [clause.key] });
    return null;
  }
  const level = abstract.levels[String(ilvl)];
  if (!level) {
    review.push({ type: 'numbering-unknown-level', detail: `条款 ${clause.key} 的 ilvl=${ilvl} 在抽象编号 ${num.abstractNumId} 中未定义`, clauses: [clause.key] });
    return null;
  }
  const override = num.overrides[String(ilvl)];
  const start = override?.startOverride ?? level.start;
  if (start == null) {
    review.push({ type: 'numbering-start-missing', detail: `条款 ${clause.key} 的编号级别缺少起始值，不臆测编号`, clauses: [clause.key] });
    return null;
  }
  const basis = {
    numId: Number(numId),
    abstractNumId: Number(num.abstractNumId),
    ilvl,
    start: level.start,
    startOverride: override?.startOverride ?? null,
    numFmt: level.numFmt ?? 'decimal',
    lvlText: level.lvlText,
  };
  const counter = counters.get(numId) ?? new Map();
  counters.set(numId, counter);
  const value = (counter.get(ilvl) ?? start - 1) + 1;
  counter.set(ilvl, value);
  for (const key of [...counter.keys()]) if (key > ilvl) counter.delete(key);
  if (basis.numFmt !== 'decimal') {
    review.push({ type: 'numbering-format-unsupported', detail: `条款 ${clause.key} 的编号格式 ${basis.numFmt} 暂未支持，编号文本需人工确认`, clauses: [clause.key] });
    return { text: null, value, basis };
  }
  if (level.lvlText == null) {
    review.push({ type: 'numbering-lvltext-missing', detail: `条款 ${clause.key} 的编号级别缺少 lvlText，不臆测编号文本`, clauses: [clause.key] });
    return { text: null, value, basis };
  }
  const text = level.lvlText.replace(/%([1-9])/g, (match, digit) => {
    const levelIndex = Number(digit) - 1;
    if (counter.has(levelIndex)) return String(counter.get(levelIndex));
    const fallbackLevel = abstract.levels[String(levelIndex)];
    const fallbackStart = num.overrides[String(levelIndex)]?.startOverride ?? fallbackLevel?.start;
    return fallbackStart != null ? String(fallbackStart) : match;
  });
  return { text, value, basis };
}

function applyFootnotes(clauses, footnotes, review) {
  for (const clause of clauses) {
    const refs = clause.segments.flatMap(segment => segment.footnoteRefs);
    clause.footnotes = refs.map(id => {
      const text = footnotes?.[id];
      if (text == null) {
        review.push({ type: 'footnote-unresolved', detail: `条款 ${clause.key} 的脚注引用 id=${id} 无法对应脚注部件`, clauses: [clause.key] });
        return { id, text: null };
      }
      return { id, text };
    });
  }
}

function parseNumbering(xml) {
  const { root, namespaces } = rootOf(parseXml(xml), 'numbering');
  const w = prefixFor(namespaces, NS_MAIN);
  if (!w) throw new Error('编号部件未声明 OOXML 主命名空间');
  const toNumber = value => (value == null ? null : Number(value));
  const abstract = {};
  const nums = {};
  for (const element of elements(root.children)) {
    if (element.tag === qn(w, 'abstractNum')) {
      const levels = {};
      for (const lvl of elements(element.children)) {
        if (lvl.tag !== qn(w, 'lvl')) continue;
        const get = local => {
          const child = elements(lvl.children).find(item => item.tag === qn(w, local));
          return child ? attr(child, qn(w, 'val')) : null;
        };
        levels[attr(lvl, qn(w, 'ilvl')) ?? '0'] = { start: toNumber(get('start')), numFmt: get('numFmt'), lvlText: get('lvlText') };
      }
      abstract[attr(element, qn(w, 'abstractNumId'))] = { levels };
    } else if (element.tag === qn(w, 'num')) {
      const overrides = {};
      let abstractNumId = null;
      for (const child of elements(element.children)) {
        if (child.tag === qn(w, 'abstractNumId')) abstractNumId = attr(child, qn(w, 'val'));
        if (child.tag === qn(w, 'lvlOverride')) {
          const startOverride = elements(child.children).find(item => item.tag === qn(w, 'startOverride'));
          overrides[attr(child, qn(w, 'ilvl')) ?? '0'] = {
            startOverride: startOverride ? toNumber(attr(startOverride, qn(w, 'val'))) : null,
          };
        }
      }
      nums[attr(element, qn(w, 'numId'))] = { abstractNumId, overrides };
    }
  }
  return { abstract, nums };
}

function parseFootnotes(xml, view, review) {
  const { root, namespaces } = rootOf(parseXml(xml), 'footnotes');
  const w = prefixFor(namespaces, NS_MAIN);
  if (!w) throw new Error('脚注部件未声明 OOXML 主命名空间');
  const ctx = { w, w14: null, part: 'footnotes', review, ranges: [] };
  const map = {};
  for (const element of elements(root.children)) {
    if (element.tag !== qn(w, 'footnote')) continue;
    const type = attr(element, qn(w, 'type'));
    if (type != null && type !== 'normal') continue; // 分隔线等非正文脚注
    const paragraphs = [];
    for (const paragraph of elements(element.children)) {
      if (paragraph.tag !== qn(w, 'p')) continue;
      const segments = [];
      collectRunSegments(elements(paragraph.children), null, ctx, segments, 'footnote');
      paragraphs.push(segments.filter(segment => KEEP[view].has(segment.status)).map(segment => segment.text).join(''));
    }
    map[attr(element, qn(w, 'id'))] = paragraphs.join('\n');
  }
  return map;
}

function finalizeClause(clause) {
  return {
    key: clause.key,
    part: clause.part,
    paraId: clause.paraId,
    identityBasis: clause.identityBasis,
    path: clause.path,
    inTable: clause.inTable,
    number: clause.number ?? null,
    segments: clause.segments.map(segment => {
      const out = { text: segment.text, status: segment.status };
      if (segment.rev) out.rev = segment.rev;
      if (segment.move) out.move = segment.move;
      if (segment.footnoteRefs.length) out.footnoteRefs = segment.footnoteRefs;
      return out;
    }),
    text: clause.segments.map(segment => segment.text).join(''),
    footnotes: clause.footnotes ?? [],
    links: clause.links,
    ...(clause.mark ? { paragraphMark: clause.mark } : {}),
  };
}
