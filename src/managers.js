import { alignLines, similarity, textDiff } from './diff.js';
import { refreshBiography } from './years.js';

const BULLET = /^[\s•·●◆\uF0D8]*/;
const BOUNDARY = /^(?:[一二三四五六七八九十\d]+[、.．]\s*)?(?:产品要素|产品介绍|产品特点|产品信息|风险提示|风险揭示|重要提示|免责声明|特别提示|管理人近年|荣誉奖项|公司实力|雄厚公司|投资策略|认购要素|申购|赎回|费用|业绩表现|历史业绩|数据来源|销售安排|开放安排|开放期|存续期|投资范围|投资比例|认申购)/;
const IDENTITY = /学士|硕士|博士|学历|毕业|大学|学院|CFA|FRM|从业|经验|现任|历任|曾任|曾就职|任职/;

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function isBoundary(text) {
  const value = String(text || '').trim().replace(/^[\s•·●◆\uF0D8]+/, '');
  if (!value) return false;
  if (/^(?:投资经理|基金经理)(?:介绍|简介)?$/.test(value)) return true;
  return BOUNDARY.test(value) && value.length <= 40;
}

function matchName(text, names) {
  for (const name of names) {
    const pattern = new RegExp(`^[\\s•·●◆\\uF0D8]*(?:(?:拟任|拟聘)?(?:投资|基金)经理(?:简介|介绍)?[：:\\s]*)?${escapeRegExp(name)}(?=$|[\\s：:，,（(])`);
    const matched = text.match(pattern);
    if (!matched) continue;
    let end = matched[0].length;
    const suffix = text.slice(end).match(/^\s*(?:[（(](?:配|附)?(?:照片|图片)[）)])?(?:先生|女士)?[\s：:，,]*/);
    if (suffix) end += suffix[0].length;
    const remainder = text.slice(end).trim();
    if (remainder && !IDENTITY.test(remainder)) continue;
    return { name, start: end, prefix: text.slice(0, end) };
  }
  return null;
}

function regionOwner(paragraph) {
  return paragraph.boxKey || paragraph.cellKey || 'body';
}

export function biographyRegions(paragraphs, names) {
  const ordered = [...new Set(names.filter(Boolean))].sort((a, b) => b.length - a.length);
  const body = paragraphs.filter(paragraph => paragraph.part === 'word/document.xml');
  const anchors = body.map(paragraph => {
    const hit = matchName(paragraph.text, ordered);
    return hit ? { paragraph, ...hit } : null;
  }).filter(Boolean);
  const anchorIds = new Set(anchors.map(anchor => anchor.paragraph.id));
  return anchors.map(anchor => {
    const entries = [];
    const initial = anchor.paragraph.text.slice(anchor.start);
    if (initial.trim()) {
      const bullet = initial.match(BULLET)?.[0] || '';
      entries.push({ paragraph: anchor.paragraph, start: anchor.start + bullet.length, text: initial.slice(bullet.length), prefix: anchor.prefix + bullet });
    }
    for (const paragraph of body) {
      if (paragraph.index <= anchor.paragraph.index) continue;
      if (anchorIds.has(paragraph.id)) break;
      if (regionOwner(paragraph) !== regionOwner(anchor.paragraph)) break;
      if (!paragraph.text.trim()) continue;
      if (isBoundary(paragraph.text)) break;
      const bullet = paragraph.text.match(BULLET)?.[0] || '';
      entries.push({ paragraph, start: bullet.length, text: paragraph.text.slice(bullet.length), prefix: bullet });
    }
    return { ...anchor, entries };
  });
}

function linesOf(text, entryCount) {
  const lines = String(text || '').split('\n').map(line => line.trim()).filter(Boolean);
  if (entryCount <= 1) return [lines.join('\n')];
  if (lines.length === 1) return lines[0].match(/[^。；;]+[。；;]?/g)?.map(line => line.trim()).filter(Boolean) || lines;
  return lines;
}

function sameCopy(before, after) {
  const norm = value => String(value || '').replace(/\r\n?/g, '\n').split('\n').map(line => line.trim()).filter(Boolean).join('\n');
  return norm(before) === norm(after);
}

function managerSection(paragraphs) {
  const body = paragraphs.filter(paragraph => paragraph.part === 'word/document.xml').sort((left, right) => left.index - right.index);
  const heading = body.find(paragraph => /^(?:投资经理|基金经理)(?:介绍|简介)?$/.test(paragraph.text.trim().replace(/^[\s•·●◆\uF0D8]+/, '')));
  if (!heading) return null;
  const section = [];
  let sawTable = Boolean(heading.tableId);
  for (const paragraph of body) {
    if (paragraph.index <= heading.index) continue;
    if (paragraph.text.trim() && isBoundary(paragraph.text)) break;
    if (!heading.tableId && paragraph.tableId) sawTable = true;
    if (!heading.tableId && sawTable && !paragraph.tableId && paragraph.text.trim()) break;
    if (heading.tableId && paragraph.tableId && paragraph.tableId !== heading.tableId) break;
    section.push(paragraph);
  }
  return { heading, paragraphs: section };
}

function isNameParagraph(text) {
  const value = String(text || '').trim().replace(/^[\s•·●◆\uF0D8]+/, '');
  if (/^[\u4e00-\u9fff]{2,4}[：:]?$/.test(value)) return true;
  return /^[\u4e00-\u9fff]{2,4}[：:]/.test(value);
}

function nameOf(text) {
  return String(text || '').trim().replace(/^[\s•·●◆\uF0D8]+/, '').match(/^([\u4e00-\u9fff]{2,4})/)?.[1] || '';
}

function managerBlocks(sectionParagraphs) {
  const blocks = [];
  let current = null;
  for (const paragraph of sectionParagraphs) {
    if (!paragraph.text.trim()) continue;
    if (isNameParagraph(paragraph.text)) {
      current = { name: nameOf(paragraph.text), paragraphs: [paragraph] };
      blocks.push(current);
    } else if (current) current.paragraphs.push(paragraph);
  }
  return blocks;
}

export function planManagerEdits(paragraphs, records, asOf, contractNames = null) {
  const roster = Array.isArray(contractNames) && contractNames.length ? [...new Set(contractNames)] : null;
  const prepared = records.map(record => ({ record, ...refreshBiography(record, asOf) })).filter(item => item.record.name && item.text.trim());
  const regions = biographyRegions(paragraphs, prepared.map(item => item.record.name));
  const edits = [];
  const reviews = [];
  const warnings = [];
  for (const region of regions) {
    if (roster && !roster.includes(region.name)) continue;
    const before = region.entries.map(entry => entry.text).join('\n');
    const candidates = prepared.filter(item => item.record.name === region.name)
      .map(item => ({ ...item, score: similarity(before, item.text) }))
      .sort((a, b) => b.score - a.score);
    if (!candidates.length) continue;
    const chosen = candidates[0];
    const after = chosen.text;
    if (!region.entries.length) {
      reviews.push({ name: region.name, status: 'inserted', before: '', after, calculations: chosen.calculations, sheet: chosen.record.sheet, row: chosen.record.row });
      let afterId = region.paragraph.id;
      for (const line of linesOf(after, 2)) {
        const newId = `new-${region.paragraph.id}-${edits.length}`;
        edits.push({ insertAfterId: afterId, newId, text: line, label: `${region.name}简介`, source: 'manager' });
        afterId = newId;
      }
      warnings.push(`${region.name}在一页通里只有姓名，已按底稿补上简介。`);
      continue;
    }
    if (sameCopy(before, after)) {
      reviews.push({ name: region.name, status: 'unchanged', before, after, calculations: chosen.calculations, sheet: chosen.record.sheet, row: chosen.record.row });
      continue;
    }
    reviews.push({ name: region.name, status: 'updated', before, after, calculations: chosen.calculations, sheet: chosen.record.sheet, row: chosen.record.row });
    const pairs = alignLines(region.entries.map(entry => entry.text), linesOf(after, region.entries.length));
    let previous = region.paragraph;
    for (const pair of pairs) {
      if (pair.before && pair.after != null) {
        const entry = pair.beforeEntry || region.entries.find(item => item.text === pair.before && !item.used);
        const target = region.entries.find(item => item.text === pair.before && !item.used) || null;
        if (target) target.used = true;
        const entryRef = target;
        if (!entryRef) continue;
        previous = entryRef.paragraph;
        for (const diff of textDiff(entryRef.text, pair.after)) {
          edits.push({ paragraphId: entryRef.paragraph.id, start: entryRef.start + diff.start, before: diff.before, after: diff.after, label: `${region.name}简介`, source: 'manager' });
        }
      } else if (pair.before && pair.after == null) {
        const target = region.entries.find(item => item.text === pair.before && !item.used);
        if (!target) continue;
        target.used = true;
        previous = target.paragraph;
        if (target.start === 0 && target.text === target.paragraph.text) edits.push({ deleteParagraphId: target.paragraph.id, label: `${region.name}简介`, source: 'manager' });
        else edits.push({ paragraphId: target.paragraph.id, start: target.start, before: target.text, after: '', label: `${region.name}简介`, source: 'manager' });
      } else if (pair.after) {
        const newId = `new-${region.paragraph.id}-${edits.length}`;
        edits.push({ insertAfterId: previous.id, newId, text: pair.after, label: `${region.name}简介`, source: 'manager' });
        previous = { id: newId };
      }
    }
  }
  if (roster) syncRoster(paragraphs, roster, prepared, regions, edits, reviews, warnings);
  else {
    const matched = new Set(reviews.map(item => item.name));
    const seen = new Set();
    for (const region of regions) {
      if (matched.has(region.name) || seen.has(region.name)) continue;
      seen.add(region.name);
      warnings.push(`一页通里的“${region.name}”不在经理底稿中，介绍保持原样。`);
    }
  }
  return { edits, reviews, warnings };
}

function syncRoster(paragraphs, roster, prepared, regions, edits, reviews, warnings) {
  const section = managerSection(paragraphs);
  if (!section) {
    warnings.push('一页通里没有找到投资经理栏目，没有按合同增减经理。');
    return;
  }
  const blocks = managerBlocks(section.paragraphs);
  const present = new Set([...blocks.map(block => block.name), ...regions.map(region => region.name)]);
  for (const block of blocks) {
    if (roster.includes(block.name)) continue;
    for (const paragraph of block.paragraphs) {
      edits.push({ deleteParagraphId: paragraph.id, label: `${block.name}简介`, source: 'manager' });
    }
    reviews.push({
      name: block.name,
      status: 'removed',
      before: block.paragraphs.map(paragraph => paragraph.text.trim()).filter(Boolean).join('\n'),
      after: '',
      calculations: [],
    });
    warnings.push(`一页通里的「${block.name}」不在合同投资经理名单中，已删除。`);
  }
  let anchor = [...blocks].reverse().find(block => roster.includes(block.name));
  let afterId = anchor ? anchor.paragraphs[anchor.paragraphs.length - 1].id : section.heading.id;
  const styleColon = blocks.some(block => /[：:]/.test(block.paragraphs[0].text));
  for (const name of roster) {
    if (present.has(name) || regions.some(region => region.name === name)) continue;
    const chosen = prepared.find(item => item.record.name === name);
    if (!chosen) {
      warnings.push(`合同里的「${name}」不在经理底稿中，一页通没有补上简介。`);
      continue;
    }
    const nameLine = styleColon || !blocks.length ? `${name}：` : name;
    const bio = chosen.text.split('\n').map(line => line.trim()).filter(Boolean).map(line => line.replace(new RegExp(`^${name}[：:，,]\\s*`), '')).filter(Boolean);
    reviews.push({
      name,
      status: 'inserted',
      before: '',
      after: chosen.text,
      calculations: chosen.calculations,
      sheet: chosen.record.sheet,
      row: chosen.record.row,
    });
    for (const line of [nameLine, ...bio]) {
      const newId = `new-manager-${edits.length}`;
      edits.push({ insertAfterId: afterId, newId, text: line, label: `${name}简介`, source: 'manager' });
      afterId = newId;
    }
  }
}
