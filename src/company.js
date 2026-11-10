const compact = value => String(value || '').normalize('NFKC').replace(/\s/g, '');
const NOTE = /数据来源|金牛奖颁奖机构|统计时间|近两年管理人类获奖情况/;
const VERSIONS = new Map([['通用版本', 'general'], ['通用版', 'general'], ['建行版本', 'ccb'], ['建行版', 'ccb']]);

export const COMPANY_LABELS = {
  capital: '注册资本',
  qualification: '公司资质',
  institution: '公司定位',
  experience: '资管经验',
  team: '投研团队',
  product_system: '产品体系',
  awards: '公司奖项',
  aum: '资管规模',
  aum_date: '规模数据日期',
  award_issuer: '奖项来源',
  award_period: '奖项统计时间',
  award_recent: '近期获奖情况',
  shareholder_description: '股东介绍',
  shareholder_network: '股东覆盖',
  shareholder_isolation: '股东隔离说明',
  aum_label: '规模名称',
  awards_label: '奖项名称',
  experience_label: '经验名称',
};

function canonical(value) {
  return compact(value)
    .replace(/^(?:股东|投资顾问|管理人|受托人)[：:]?/, '')
    .replace(/(?:股份有限公司|有限责任公司|有限公司)$/, '')
    .replace(/证券资产管理$/, '资产管理')
    .replace(/资管$/, '资产管理');
}

export function sameSubject(a, b) {
  const left = canonical(a);
  const right = canonical(b);
  return Boolean(left && right && left === right);
}

function explicitName(value) {
  const name = compact(value).replace(/^(?:股东|投资顾问|管理人|受托人)[：:]?/, '');
  return name.length >= 4 && name.length <= 60 && !/[，。；、：]/.test(name) && /(?:资产管理|资管|证券|信托|银行|有限公司)$/.test(name) ? name : '';
}

function atom(paragraph, key, start, text) {
  return { key, text, start, paragraphId: paragraph.id };
}

function companyAtoms(paragraph, role = 'company') {
  if (NOTE.test(paragraph.text) || /更名为|变更为|发行与管理|兑付责任/.test(paragraph.text)) return [];
  if (role === 'shareholder') {
    const text = paragraph.text.trim();
    const start = paragraph.text.indexOf(text);
    const key = /业务隔离|股东不直接/.test(text) ? 'shareholder_isolation'
      : /分支机构|服务范围覆盖|资本实力位居/.test(text) ? 'shareholder_network'
        : /(?:有限公司|金融服务商)/.test(text) ? 'shareholder_description' : '';
    return key && text ? [atom(paragraph, key, start, text)] : [];
  }
  const found = [];
  const aum = paragraph.text.match(/截至[^。；;\n]*?(?:资产管理|资管)规模[^。；;\n]*/);
  if (aum) found.push(atom(paragraph, 'aum', aum.index, aum[0]));
  for (const match of paragraph.text.matchAll(/[^，,。；;\n]+/g)) {
    if (found.some(item => match.index >= item.start && match.index < item.start + item.text.length)) continue;
    const text = match[0].replace(/^[\s•·●◆]+/, '').trim();
    if (!text) continue;
    const start = match.index + match[0].indexOf(text);
    const key = /注册资本/.test(text) ? 'capital'
      : /首批.*资产管理子公司/.test(text) ? 'qualification'
        : /全能型资产管理机构/.test(text) ? 'institution'
          : /\d+多年?资产管理经验|\d+年资产管理经验/.test(text) ? 'experience'
            : /投研团队/.test(text) ? 'team'
              : /产品体系/.test(text) ? 'product_system'
                : /荣获.*金牛奖/.test(text) ? 'awards' : '';
    if (key) found.push(atom(paragraph, key, start, text));
  }
  return found.sort((a, b) => a.start - b.start);
}

function noteAtoms(paragraph) {
  const rules = [
    ['aum_date', /(?:资管规模数据时间|规模数据|数据时间)?截至\d{4}[.年/-]\d{1,2}[.月/-]\d{1,2}日?/g],
    ['award_issuer', /金牛奖颁奖机构为[^，,。；;\n]+/g],
    ['award_period', /统计时间为[^，,。；;\n]+/g],
    ['award_recent', /近两年管理人类获奖情况为[^。；;\n]+/g],
  ];
  const found = [];
  for (const [key, pattern] of rules) {
    for (const match of paragraph.text.matchAll(pattern)) found.push(atom(paragraph, key, match.index, match[0]));
  }
  const dates = found.filter(item => item.key === 'aum_date');
  const labeled = dates.filter(item => /^(?:资管规模数据时间|规模数据)/.test(item.text));
  return labeled.length ? found.filter(item => item.key !== 'aum_date' || labeled.includes(item)) : found;
}

function tablesOf(paragraphs) {
  const groups = new Map();
  for (const paragraph of paragraphs) {
    if (!paragraph.tableId || paragraph.part !== 'word/document.xml') continue;
    if (!groups.has(paragraph.tableId)) groups.set(paragraph.tableId, []);
    groups.get(paragraph.tableId).push(paragraph);
  }
  return [...groups.values()];
}

function narrative(table) {
  const notes = table.filter(paragraph => NOTE.test(paragraph.text));
  const cores = table.flatMap(paragraph => companyAtoms(paragraph));
  const keys = [...new Set(cores.map(item => item.key))];
  if (keys.length < 3) return null;
  const firstRow = Math.min(...cores.map(item => table.find(paragraph => paragraph.id === item.paragraphId).rowIndex));
  const headers = table.filter(paragraph => paragraph.rowIndex < firstRow && paragraph.text.trim());
  const columns = [];
  for (const cell of [...new Set(table.map(paragraph => paragraph.cellIndex))]) {
    const head = headers.filter(paragraph => paragraph.cellIndex === cell);
    const headerText = head.map(paragraph => paragraph.text).join('');
    const role = /股东/.test(headerText) ? 'shareholder' : 'company';
    const named = head.map(paragraph => explicitName(paragraph.text)).filter(Boolean);
    const body = table.filter(paragraph => paragraph.cellIndex === cell && paragraph.rowIndex >= firstRow && !notes.includes(paragraph));
    let subject = named[0] || '';
    if (!subject && role === 'shareholder') {
      const first = body.find(paragraph => /股份有限公司|有限责任公司|有限公司/.test(paragraph.text));
      subject = first?.text.match(/([^，。；\n]{4,55}?(?:股份有限公司|有限责任公司|有限公司))/)?.[1] || '国泰海通证券';
    }
    if (!subject && role === 'company') subject = '国泰海通资产管理';
    const entries = body.flatMap(paragraph => companyAtoms(paragraph, role));
    if (entries.length) columns.push({ role, subject, cellIndex: cell, entries, paragraphIds: body.map(paragraph => paragraph.id) });
  }
  if (!columns.some(column => column.role === 'company')) return null;
  return { layout: 'narrative', columns, notes: notes.map(paragraph => ({ paragraphId: paragraph.id, text: paragraph.text, entries: noteAtoms(paragraph) })) };
}

function kpi(table) {
  const labelOf = text => ({ 资管规模: 'aum', 金牛奖: 'awards', 荣获金牛奖: 'awards', 资管经验: 'experience' }[compact(text)] || '');
  const labels = table.filter(paragraph => labelOf(paragraph.text));
  const keys = [...new Set(labels.map(paragraph => labelOf(paragraph.text)))];
  if (keys.length !== 3) return null;
  const entries = [];
  for (const label of labels) {
    const values = table.filter(paragraph => paragraph.cellIndex === label.cellIndex && paragraph.rowIndex < label.rowIndex && paragraph.text.trim());
    if (values.length !== 1) return null;
    const value = values[0];
    entries.push(atom(value, labelOf(label.text), value.text.indexOf(value.text.trim()), value.text.trim()));
  }
  return { layout: 'kpi', columns: [{ role: 'company', subject: '国泰海通资产管理', entries, paragraphIds: entries.map(item => item.paragraphId) }], notes: [] };
}

export function companySections(paragraphs) {
  return tablesOf(paragraphs).map(table => narrative(table) || kpi(table)).filter(Boolean);
}

export function selectVariant(paragraphs, variant) {
  if (!['general', 'ccb'].includes(variant)) throw new Error('请选择通用版或建行版公司介绍。');
  const markers = paragraphs.filter(paragraph => paragraph.part === 'word/document.xml' && !paragraph.tableId && VERSIONS.has(compact(paragraph.text)));
  if (!markers.length) return paragraphs;
  const ranges = markers.flatMap((paragraph, index) => (
    VERSIONS.get(compact(paragraph.text)) === variant ? [[paragraph.index, markers[index + 1]?.index ?? Infinity]] : []
  ));
  if (!ranges.length) throw new Error(variant === 'ccb' ? '公司介绍文件里没有建行版。' : '公司介绍文件里没有通用版。');
  return paragraphs.filter(paragraph => paragraph.part !== 'word/document.xml' || ranges.some(([start, end]) => paragraph.index >= start && paragraph.index < end));
}

function dedupedNotes(section) {
  const seen = new Set();
  return section.notes.filter(note => {
    const key = note.entries.map(entry => `${entry.key}:${entry.text}`).join('|');
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function planCompanyEdits(sourceParagraphs, targetParagraphs, variant) {
  const source = companySections(selectVariant(sourceParagraphs, variant));
  const target = companySections(targetParagraphs);
  const edits = [];
  const warnings = [];
  if (!source.length) warnings.push('公司介绍底稿没有识别出双栏或三项指标，请检查 Word 底稿。');
  if (!target.length) warnings.push('这份一页通里没有找到可核对的公司介绍区域，公司介绍保持原样。');
  for (const section of target) {
    const candidates = source.filter(item => item.layout === section.layout && item.columns.some(column => section.columns.some(targetColumn => targetColumn.role === column.role && sameSubject(targetColumn.subject, column.subject))));
    if (candidates.length !== 1) {
      warnings.push(section.layout === 'kpi' ? '指标区不能唯一对应到底稿，已保留原文。' : '公司介绍双栏不能唯一对应到底稿，已保留原文。');
      continue;
    }
    const basis = candidates[0];
    for (const column of section.columns) {
      const sources = basis.columns.filter(item => item.role === column.role && sameSubject(item.subject, column.subject));
      if (sources.length !== 1) continue;
      const basisColumn = sources[0];
      for (const entry of column.entries) {
        const matches = basisColumn.entries.filter(item => item.key === entry.key);
        if (matches.length !== 1) {
          warnings.push(`“${COMPANY_LABELS[entry.key] || entry.key}”在底稿里没有唯一对应，已保留原文。`);
          continue;
        }
        if (entry.text === matches[0].text) continue;
        edits.push({ paragraphId: entry.paragraphId, start: entry.start, before: entry.text, after: matches[0].text, label: COMPANY_LABELS[entry.key] || entry.key, source: 'company' });
      }
      let insertAfterId = [...column.paragraphIds].reverse().find(id => targetParagraphs.find(paragraph => paragraph.id === id)?.text.trim()) || '';
      for (const entry of basisColumn.entries) {
        if (column.entries.some(item => item.key === entry.key)) continue;
        if (!insertAfterId) continue;
        const newId = `company-new-${edits.length}`;
        edits.push({ insertAfterId, newId, text: entry.text, label: COMPANY_LABELS[entry.key] || entry.key, source: 'company' });
        insertAfterId = newId;
        warnings.push(`一页通缺少“${COMPANY_LABELS[entry.key] || entry.key}”，已按底稿插入。`);
      }
    }
    const sourceNotes = dedupedNotes(basis).flatMap(note => note.entries);
    for (const note of section.notes) {
      for (const entry of note.entries) {
        const matches = sourceNotes.filter(item => item.key === entry.key);
        const unique = [...new Map(matches.map(item => [item.text, item])).values()];
        if (unique.length !== 1) {
          if (unique.length > 1) warnings.push(`“${COMPANY_LABELS[entry.key] || entry.key}”有多个底稿来源，已保留原文。`);
          continue;
        }
        if (entry.text === unique[0].text) continue;
        edits.push({ paragraphId: entry.paragraphId, start: entry.start, before: entry.text, after: unique[0].text, label: COMPANY_LABELS[entry.key] || entry.key, source: 'company' });
      }
    }
  }
  return { edits, warnings };
}
