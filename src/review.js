import { planCompanyEdits } from './company.js';
import { planManagerEdits } from './managers.js';
import { applyParagraphEdits, deleteParagraph, insertParagraphAfter, maxRevisionId } from './revise.js';
import { collectParagraphs, parseXml, serializeXml, NS } from './xml.js';

const PART = /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

function rangesOverlap(a, b) {
  return a.start < b.start + b.before.length && b.start < a.start + a.before.length;
}

export async function readDocxParagraphs(bytes) {
  const zip = await globalThis.JSZip.loadAsync(bytes);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('这份 Word 不是有效的 docx。');
  const doc = parseXml(await entry.async('string'));
  return collectParagraphs(doc, 'word/document.xml');
}

export async function reviewOnePager({ bytes, fileName, companyBytes, variant, records, asOf, author = '介绍底稿', companyTextById = null }) {
  if (!/\.docx$/i.test(fileName || '')) throw new Error('一页通请使用 docx 文件。');
  const sourceParagraphs = await readDocxParagraphs(companyBytes);
  if (companyTextById) {
    for (const paragraph of sourceParagraphs) {
      if (Object.prototype.hasOwnProperty.call(companyTextById, paragraph.id)) paragraph.text = companyTextById[paragraph.id];
    }
  }
  const zip = await globalThis.JSZip.loadAsync(bytes);
  const partNames = Object.keys(zip.files).filter(path => PART.test(path));
  if (!partNames.includes('word/document.xml')) throw new Error('一页通缺少正文，无法核对。');
  const docs = new Map();
  const paragraphs = [];
  for (const part of partNames) {
    const doc = parseXml(await zip.file(part).async('string'));
    docs.set(part, doc);
    paragraphs.push(...collectParagraphs(doc, part));
  }
  const company = planCompanyEdits(sourceParagraphs, paragraphs, variant);
  const managers = planManagerEdits(paragraphs, records, asOf);
  const edits = [...company.edits, ...managers.edits];
  const warnings = [...company.warnings, ...managers.warnings];
  const byParagraph = new Map();
  const blocked = new Set();
  for (const edit of edits) {
    if (!edit.paragraphId) continue;
    const list = byParagraph.get(edit.paragraphId) || [];
    if (list.some(item => rangesOverlap(item, edit) && item.source !== edit.source)) {
      warnings.push(`“${edit.label}”和已有修改重叠，这条保留原文。`);
      blocked.add(edit);
      continue;
    }
    if (list.some(item => rangesOverlap(item, edit))) continue;
    list.push(edit);
    byParagraph.set(edit.paragraphId, list);
  }
  const deletions = edits.filter(edit => edit.deleteParagraphId && !blocked.has(edit));
  const insertions = edits.filter(edit => edit.insertAfterId && !blocked.has(edit));
  const changed = byParagraph.size + deletions.length + insertions.length > 0;
  if (!changed) {
    return { fileName, changed: false, bytes, warnings, managers: managers.reviews, companyLabels: [], changeCount: 0 };
  }
  let revision = maxRevisionId([...docs.values()]);
  const meta = {
    author,
    date: new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'),
    next: () => { revision += 1; return revision; },
  };
  const paragraphById = new Map(paragraphs.map(paragraph => [paragraph.id, paragraph]));
  for (const [paragraphId, list] of byParagraph) {
    const paragraph = paragraphById.get(paragraphId);
    if (!paragraph) continue;
    applyParagraphEdits(paragraph.node, list.map(edit => ({ start: edit.start, before: edit.before, after: edit.after })), meta);
  }
  for (const edit of deletions) {
    const paragraph = paragraphById.get(edit.deleteParagraphId);
    if (paragraph && !byParagraph.has(paragraph.id)) deleteParagraph(paragraph.node, meta);
  }
  const created = new Map();
  for (const edit of insertions) {
    const anchor = created.get(edit.insertAfterId) || paragraphById.get(edit.insertAfterId)?.node;
    if (!anchor) continue;
    const node = insertParagraphAfter(anchor, edit.text, meta);
    if (edit.newId) created.set(edit.newId, node);
  }
  for (const [part, doc] of docs) zip.file(part, serializeXml(doc));
  await ensureTracking(zip);
  const output = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const companyLabels = [...new Set(edits.filter(edit => edit.source === 'company').map(edit => edit.label))];
  return {
    fileName,
    changed: true,
    bytes: output,
    warnings,
    managers: managers.reviews,
    companyLabels,
    changeCount: [...byParagraph.values()].reduce((total, list) => total + list.length, 0) + deletions.length + insertions.length,
  };
}

async function ensureTracking(zip) {
  const file = zip.file('word/settings.xml');
  if (!file) return;
  const doc = parseXml(await file.async('string'));
  if (doc.getElementsByTagNameNS(NS.w, 'trackRevisions').length) return;
  const marker = doc.createElementNS(NS.w, 'w:trackRevisions');
  doc.documentElement.insertBefore(marker, doc.documentElement.firstChild);
  zip.file('word/settings.xml', serializeXml(doc));
}

export function downloadName(fileName) {
  const stem = String(fileName || '一页通.docx').replace(/\.docx$/i, '');
  return `${stem}（介绍修订）.docx`;
}
