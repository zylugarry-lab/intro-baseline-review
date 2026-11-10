export const NS = Object.freeze({
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  xml: 'http://www.w3.org/XML/1998/namespace',
  ss: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
});

const isElement = node => node?.nodeType === 1;

export function childList(node) {
  const out = [];
  for (let child = node?.firstChild; child; child = child.nextSibling) out.push(child);
  return out;
}

export function nodeList(collection) {
  const out = [];
  if (!collection) return out;
  for (let index = 0; index < collection.length; index += 1) out.push(collection[index]);
  return out;
}

export function parseXml(text) {
  const Parser = globalThis.DOMParser;
  if (!Parser) throw new Error('当前环境不能读取 XML。');
  let doc;
  try {
    doc = new Parser().parseFromString(text, 'application/xml');
  } catch (error) {
    throw new Error(`XML 读取失败：${error.message}`);
  }
  if (!doc?.documentElement || doc.getElementsByTagName('parsererror').length) throw new Error('XML 读取失败。');
  return doc;
}

export function serializeXml(doc) {
  const Serializer = globalThis.XMLSerializer;
  const body = new Serializer().serializeToString(doc);
  return body.startsWith('<?xml') ? body : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

export function children(node, ns, local) {
  const out = [];
  for (let child = node?.firstChild; child; child = child.nextSibling) {
    if (isElement(child) && child.namespaceURI === ns && child.localName === local) out.push(child);
  }
  return out;
}

export function elements(node, ns, local) {
  const out = [];
  const visit = parent => {
    for (let child = parent?.firstChild; child; child = child.nextSibling) {
      if (!isElement(child)) continue;
      if (child.namespaceURI === ns && child.localName === local) out.push(child);
      visit(child);
    }
  };
  visit(node);
  return out;
}

export function ancestor(node, ns, local) {
  for (let current = node?.parentNode; current; current = current.parentNode) {
    if (current.namespaceURI === ns && current.localName === local) return current;
  }
  return null;
}

export function isVisible(node) {
  for (let current = node; current; current = current.parentNode) {
    if (!isElement(current)) continue;
    if (current.namespaceURI === NS.w && (current.localName === 'del' || current.localName === 'moveFrom')) return false;
    if (current.namespaceURI === NS.mc && current.localName === 'Fallback') {
      const choice = childList(current.parentNode).some(item => isElement(item) && item.namespaceURI === NS.mc && item.localName === 'Choice');
      if (choice) return false;
    }
  }
  return true;
}

export function visibleText(paragraph) {
  let out = '';
  const visit = node => {
    if (!isElement(node)) return;
    if (node !== paragraph && node.namespaceURI === NS.w && node.localName === 'p') return;
    if (!isVisible(node)) return;
    if (node.namespaceURI === NS.w) {
      if (node.localName === 't') { out += node.textContent || ''; return; }
      if (node.localName === 'tab') { out += '\t'; return; }
      if (node.localName === 'br' || node.localName === 'cr') { out += '\n'; return; }
    }
    for (let child = node.firstChild; child; child = child.nextSibling) visit(child);
  };
  visit(paragraph);
  return out;
}

/** Body order, including text boxes. Fallback copies and deleted branches are skipped. */
export function collectParagraphs(doc, part) {
  const paragraphs = [];
  let tableSeq = 0;
  let boxSeq = 0;
  const walk = (node, ctx) => {
    if (!isElement(node)) return;
    if (node.namespaceURI === NS.w && (node.localName === 'del' || node.localName === 'moveFrom')) return;
    if (node.namespaceURI === NS.mc && node.localName === 'Fallback') return;
    if (node.namespaceURI === NS.w && node.localName === 'txbxContent') {
      const boxKey = `${part}#x${boxSeq += 1}`;
      for (let child = node.firstChild; child; child = child.nextSibling) walk(child, { ...ctx, boxKey });
      return;
    }
    if (node.namespaceURI === NS.w && node.localName === 'tbl') {
      const tableId = `${part}#t${tableSeq += 1}`;
      let rowIndex = -1;
      for (let row = node.firstChild; row; row = row.nextSibling) {
        if (!isElement(row) || row.namespaceURI !== NS.w || row.localName !== 'tr') continue;
        rowIndex += 1;
        let cellIndex = -1;
        for (let cell = row.firstChild; cell; cell = cell.nextSibling) {
          if (!isElement(cell) || cell.namespaceURI !== NS.w || cell.localName !== 'tc') continue;
          cellIndex += 1;
          const next = {
            tableId, rowIndex, cellIndex,
            cellKey: `${tableId}:${rowIndex}:${cellIndex}`,
            boxKey: ctx.boxKey || '',
          };
          for (let child = cell.firstChild; child; child = child.nextSibling) walk(child, next);
        }
      }
      return;
    }
    if (node.namespaceURI === NS.w && node.localName === 'p') {
      paragraphs.push({
        node, part,
        text: visibleText(node),
        tableId: ctx.tableId || '',
        rowIndex: ctx.rowIndex ?? -1,
        cellIndex: ctx.cellIndex ?? -1,
        cellKey: ctx.cellKey || '',
        boxKey: ctx.boxKey || '',
      });
      return;
    }
    for (let child = node.firstChild; child; child = child.nextSibling) walk(child, ctx);
  };
  walk(doc.documentElement || doc, {});
  paragraphs.forEach((paragraph, index) => {
    paragraph.index = index;
    paragraph.id = `${part}:${index}`;
  });
  return paragraphs;
}
