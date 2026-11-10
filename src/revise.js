import { NS, ancestor, childList, isVisible, nodeList, visibleText } from './xml.js';

const isElement = node => node?.nodeType === 1;

export function acceptRevisions(node) {
  const kids = childList(node);
  for (const child of kids) {
    if (!isElement(child)) continue;
    const revision = child.namespaceURI === NS.w && ['del', 'moveFrom', 'ins', 'moveTo'].includes(child.localName);
    if (child.namespaceURI === NS.w && (child.localName === 'del' || child.localName === 'moveFrom')) {
      node.removeChild(child);
      continue;
    }
    if (child.namespaceURI === NS.w && (child.localName === 'ins' || child.localName === 'moveTo')) {
      acceptRevisions(child);
      while (child.firstChild) node.insertBefore(child.firstChild, child);
      node.removeChild(child);
      continue;
    }
    if (!revision) acceptRevisions(child);
  }
}

function element(doc, name) {
  return doc.createElementNS(NS.w, `w:${name}`);
}

function stamp(node, meta) {
  node.setAttributeNS(NS.w, 'w:id', String(meta.next()));
  node.setAttributeNS(NS.w, 'w:author', meta.author);
  node.setAttributeNS(NS.w, 'w:date', meta.date);
  return node;
}

function preserveSpace(node, text) {
  if (/^\s|\s$/.test(text) || text === '') node.setAttributeNS(NS.xml, 'xml:space', 'preserve');
}

function runProperties(run) {
  return childList(run).find(node => isElement(node) && node.namespaceURI === NS.w && node.localName === 'rPr') || null;
}

function pieces(paragraph) {
  const items = [];
  const visit = node => {
    if (!isElement(node)) return;
    if (node !== paragraph && node.namespaceURI === NS.w && node.localName === 'p') return;
    if (!isVisible(node)) return;
    if (node.namespaceURI === NS.w) {
      if (node.localName === 't') { items.push({ type: 'text', node, text: node.textContent || '' }); return; }
      if (node.localName === 'br' || node.localName === 'cr') { items.push({ type: 'break', node }); return; }
      if (node.localName === 'tab') { items.push({ type: 'tab', node }); return; }
    }
    for (let child = node.firstChild; child; child = child.nextSibling) visit(child);
  };
  visit(paragraph);
  return items;
}

function pieceText(items) {
  return items.map(item => (item.type === 'text' ? item.text : item.type === 'tab' ? '\t' : '\n')).join('');
}

function splitText(node, offset) {
  const text = node.textContent || '';
  if (offset <= 0 || offset >= text.length) return offset <= 0 ? node : null;
  const right = node.ownerDocument.createElementNS(NS.w, 'w:t');
  const space = node.getAttributeNS?.(NS.xml, 'space') || node.getAttribute('xml:space');
  if (space) right.setAttributeNS(NS.xml, 'xml:space', space);
  right.textContent = text.slice(offset);
  node.textContent = text.slice(0, offset);
  node.parentNode.insertBefore(right, node.nextSibling);
  return right;
}

function isolateText(textNode) {
  const run = ancestor(textNode, NS.w, 'r');
  if (!run) return null;
  const properties = runProperties(run);
  const preceding = [];
  for (let child = run.firstChild; child && child !== textNode; child = child.nextSibling) preceding.push(child);
  if (preceding.some(node => node !== properties)) {
    const left = run.cloneNode(false);
    if (properties) left.appendChild(properties.cloneNode(true));
    run.parentNode.insertBefore(left, run);
    for (const node of preceding) if (node !== properties) left.appendChild(node);
  }
  const following = [];
  for (let child = textNode.nextSibling; child; child = child.nextSibling) following.push(child);
  if (following.length) {
    const right = run.cloneNode(false);
    if (properties) right.appendChild(properties.cloneNode(true));
    run.parentNode.insertBefore(right, run.nextSibling);
    for (const node of following) right.appendChild(node);
  }
  return run;
}

function toDeletedText(node) {
  if (node.localName !== 't') return;
  const deleted = node.ownerDocument.createElementNS(NS.w, 'w:delText');
  const space = node.getAttributeNS?.(NS.xml, 'space') || node.getAttribute('xml:space');
  if (space) deleted.setAttributeNS(NS.xml, 'xml:space', space);
  else preserveSpace(deleted, node.textContent || '');
  deleted.textContent = node.textContent || '';
  node.parentNode.replaceChild(deleted, node);
}

function appendLiteral(run, text) {
  const doc = run.ownerDocument;
  const lines = String(text).split('\n');
  lines.forEach((line, index) => {
    if (index) run.appendChild(element(doc, 'br'));
    if (!line) return;
    const textNode = element(doc, 't');
    preserveSpace(textNode, line);
    textNode.textContent = line;
    run.appendChild(textNode);
  });
}

function makeInsertedRun(doc, sampleRun, text, meta) {
  const inserted = stamp(element(doc, 'ins'), meta);
  const run = element(doc, 'r');
  const properties = sampleRun ? runProperties(sampleRun) : null;
  if (properties) run.appendChild(properties.cloneNode(true));
  appendLiteral(run, text);
  inserted.appendChild(run);
  return inserted;
}

function replaceWhole(paragraph, after, meta) {
  const doc = paragraph.ownerDocument;
  const runs = [];
  const visit = node => {
    if (!isElement(node)) return;
    if (node !== paragraph && node.namespaceURI === NS.w && node.localName === 'p') return;
    if (node.namespaceURI === NS.w && node.localName === 'r') { runs.push(node); return; }
    for (let child = node.firstChild; child; child = child.nextSibling) visit(child);
  };
  visit(paragraph);
  const sample = runs[0] || null;
  if (runs.length) {
    const deleted = stamp(element(doc, 'del'), meta);
    runs[0].parentNode.insertBefore(deleted, runs[0]);
    for (const run of runs) {
      for (const node of nodeList(run.getElementsByTagNameNS(NS.w, 't'))) toDeletedText(node);
      deleted.appendChild(run);
    }
  }
  if (after) paragraph.appendChild(makeInsertedRun(doc, sample, after, meta));
  else markParagraphDeleted(paragraph, meta);
}

function markParagraphDeleted(paragraph, meta) {
  const doc = paragraph.ownerDocument;
  let properties = childList(paragraph).find(node => isElement(node) && node.localName === 'pPr');
  if (!properties) {
    properties = element(doc, 'pPr');
    paragraph.insertBefore(properties, paragraph.firstChild);
  }
  let runPropertiesNode = childList(properties).find(node => isElement(node) && node.localName === 'rPr');
  if (!runPropertiesNode) {
    runPropertiesNode = element(doc, 'rPr');
    properties.appendChild(runPropertiesNode);
  }
  runPropertiesNode.appendChild(stamp(element(doc, 'del'), meta));
}

function textSpans(paragraph) {
  let cursor = 0;
  return pieces(paragraph).filter(item => item.type === 'text').map(item => {
    const span = { node: item.node, start: cursor, end: cursor + item.text.length, text: item.text };
    cursor = span.end;
    return span;
  });
}

function replaceRange(paragraph, edit, meta) {
  const items = pieces(paragraph);
  const text = pieceText(items);
  const start = edit.start;
  const end = start + edit.before.length;
  if (text.slice(start, end) !== edit.before) throw new Error('修订位置和原文不一致，已停止写入。');
  if (edit.before.includes('\n') || edit.after.includes('\n') || items.some(item => item.type !== 'text')) {
    replaceWhole(paragraph, text.slice(0, start) + edit.after + text.slice(end), meta);
    return;
  }
  const doc = paragraph.ownerDocument;
  const sample = textSpans(paragraph)[0] ? ancestor(textSpans(paragraph)[0].node, NS.w, 'r') : null;
  if (!edit.before) {
    const inserted = makeInsertedRun(doc, sample, edit.after, meta);
    const spans = textSpans(paragraph);
    const atStart = spans.find(span => span.start === start);
    const around = spans.find(span => span.start < start && start < span.end);
    const atEnd = spans.find(span => span.end === start);
    if (around) {
      const right = splitText(around.node, start - around.start);
      const run = isolateText(right);
      run.parentNode.insertBefore(inserted, run);
    } else if (atStart) {
      ancestor(atStart.node, NS.w, 'r').parentNode.insertBefore(inserted, ancestor(atStart.node, NS.w, 'r'));
    } else if (atEnd) {
      const run = ancestor(atEnd.node, NS.w, 'r');
      run.parentNode.insertBefore(inserted, run.nextSibling);
    } else paragraph.appendChild(inserted);
    return;
  }
  let spans = textSpans(paragraph);
  const first = spans.find(span => span.end > start && span.start < end);
  const last = [...spans].reverse().find(span => span.end > start && span.start < end);
  if (!first || !last) throw new Error('修订位置超出段落。');
  if (end < last.end) splitText(last.node, end - last.start);
  if (start > first.start) first.node = splitText(first.node, start - first.start);
  spans = textSpans(paragraph);
  const selected = [];
  for (const span of spans) {
    if (span.end <= start || span.start >= end) continue;
    selected.push(span.node);
  }
  if (selected.map(node => node.textContent || '').join('') !== edit.before) throw new Error('修订切分和原文不一致，已停止写入。');
  const runs = [];
  for (const node of selected) {
    const run = isolateText(node);
    if (run && !runs.includes(run)) runs.push(run);
  }
  const deleted = stamp(element(doc, 'del'), meta);
  runs[0].parentNode.insertBefore(deleted, runs[0]);
  for (const run of runs) {
    for (const node of childList(run)) if (isElement(node) && node.localName === 't') toDeletedText(node);
    deleted.appendChild(run);
  }
  if (edit.after) deleted.parentNode.insertBefore(makeInsertedRun(doc, runs[0], edit.after, meta), deleted.nextSibling);
}

export function applyParagraphEdits(paragraph, edits, meta) {
  if (!edits.length) return;
  acceptRevisions(paragraph);
  const original = visibleText(paragraph);
  const ordered = [...edits].sort((a, b) => b.start - a.start);
  let text = original;
  for (const edit of ordered) {
    if (text.slice(edit.start, edit.start + edit.before.length) !== edit.before) throw new Error('修订位置和原文不一致，已停止写入。');
    text = text.slice(0, edit.start) + edit.after + text.slice(edit.start + edit.before.length);
  }
  // A line break makes character offsets unsafe after the first rewrite, so the
  // whole paragraph gets one deletion and one insertion.
  const structural = pieces(paragraph).some(item => item.type !== 'text') || /[\n\t]/.test(text);
  if (structural) {
    replaceWhole(paragraph, text, meta);
    return;
  }
  for (const edit of ordered) replaceRange(paragraph, edit, meta);
}

export function insertParagraphAfter(paragraph, text, meta) {
  const doc = paragraph.ownerDocument;
  const created = element(doc, 'p');
  const properties = element(doc, 'pPr');
  const runPropertiesNode = element(doc, 'rPr');
  runPropertiesNode.appendChild(stamp(element(doc, 'ins'), meta));
  properties.appendChild(runPropertiesNode);
  created.appendChild(properties);
  const sample = nodeList(paragraph.getElementsByTagNameNS(NS.w, 'r'))[0] || null;
  created.appendChild(makeInsertedRun(doc, sample, text, meta));
  paragraph.parentNode.insertBefore(created, paragraph.nextSibling);
  return created;
}

export function deleteParagraph(paragraph, meta) {
  acceptRevisions(paragraph);
  replaceWhole(paragraph, '', meta);
}

export function maxRevisionId(docs) {
  let max = 0;
  for (const doc of docs) {
    for (const node of nodeList(doc.getElementsByTagName('*'))) {
      const id = node.getAttributeNS?.(NS.w, 'id') || node.getAttribute('w:id');
      if (id && /^\d+$/.test(id)) max = Math.max(max, Number(id));
    }
  }
  return max;
}
