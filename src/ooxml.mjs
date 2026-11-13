import { XMLParser } from 'fast-xml-parser';

export const NS_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const PARSER_OPTIONS = {
  ignoreAttributes: false,
  preserveOrder: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
};

export function parseXml(xml) {
  return new XMLParser(PARSER_OPTIONS).parse(xml);
}

export function localName(tag) {
  const index = tag.indexOf(':');
  return index < 0 ? tag : tag.slice(index + 1);
}

export function qn(prefix, local) {
  return `${prefix}:${local}`;
}

// preserveOrder 节点 → 元素数组 [{tag, children, attrs}]，文本节点被跳过
export function elements(nodes) {
  const output = [];
  for (const node of nodes ?? []) {
    for (const tag of Object.keys(node)) {
      if (tag === ':@' || tag === '#text') continue;
      output.push({ tag, children: node[tag], attrs: node[':@'] ?? {} });
    }
  }
  return output;
}

export function textOf(nodes) {
  let text = '';
  for (const node of nodes ?? []) {
    if (typeof node['#text'] === 'string') text += node['#text'];
  }
  return text;
}

export function attr(element, qname) {
  const value = element.attrs['@_' + qname];
  return value === undefined ? null : value;
}

// 根元素及其命名空间声明（前缀 → URI），采集只认绑定到 OOXML 主命名空间的前缀
export function rootOf(parsed, local) {
  const root = elements(parsed).find(element => localName(element.tag) === local);
  if (!root) throw new Error(`XML 缺少根元素 ${local}`);
  const namespaces = {};
  for (const [key, value] of Object.entries(root.attrs)) {
    if (key === '@_xmlns') namespaces[''] = value;
    else if (key.startsWith('@_xmlns:')) namespaces[key.slice('@_xmlns:'.length)] = value;
  }
  return { root, namespaces };
}

export function prefixFor(namespaces, uri) {
  for (const [prefix, candidate] of Object.entries(namespaces)) {
    if (candidate === uri) return prefix;
  }
  return null;
}

export function relationships(xml) {
  const { root } = rootOf(parseXml(xml), 'Relationships');
  return elements(root.children)
    .filter(element => localName(element.tag) === 'Relationship')
    .map(element => ({ id: attr(element, 'Id'), type: attr(element, 'Type'), target: attr(element, 'Target') }));
}

export function relsPartName(part) {
  const slash = part.lastIndexOf('/');
  return `${slash < 0 ? '' : part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

export function resolveTarget(basePart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const segments = basePart.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return segments.join('/');
}
