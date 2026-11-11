const LABELS = [
  ['产品名称', /^产品名称$/],
  ['管理人', /^(?:集合|资产)?管理人$/],
  ['托管人', /^(?:集合|资产)?托管人$/],
  ['产品类型', /^(?:产品|计划)(?:类型|类别)$/],
  ['投资范围', /^投资范围$/],
  ['投资比例', /^投资比例$/],
  ['投资策略', /^投资策略$/],
  ['存续期', /^存续期(?:限)?$/],
  ['开放期', /^开放(?:期|安排)$/],
  ['认购/参与费率，退出费率', /^(?:认购[/／]参与费率[，,]退出费率|参与[/／、]?退出(?:费率|费)|认购(?:[/／]参与)?(?:费率|费)|申购(?:费率|费))$/],
  ['首参金额', /^(?:首参金额|首次参与(?:最低)?金额|参与起点|认购起点)$/],
  ['产品费率', /^(?:产品费率|固定管理费率|管理费率)$/],
  ['业绩报酬', /^业绩报酬(?:计提基准|计提)?$/],
  ['产品风险等级', /^(?:产品)?风险(?:等级|评级)$/],
];

export function planElementEdits(paragraphs, rows) {
  const incoming = (rows || []).map(row => ({
    key: canonicalLabel(row?.label),
    value: String(row?.value || '').trim(),
  })).filter(row => row.key);
  if (!incoming.length) return { edits: [], warnings: [], labels: [] };
  const table = findElementTable(paragraphs);
  if (!table) return { edits: [], warnings: ['这份一页通里没有找到产品要素表，要素没有写入正文。'], labels: [] };
  const byKey = new Map();
  for (const row of table.rows) {
    const key = canonicalLabel(row.label);
    if (key && !byKey.has(key)) byKey.set(key, row);
  }
  const edits = [];
  const warnings = [];
  const labels = [];
  for (const item of incoming) {
    const current = byKey.get(item.key);
    if (!item.value) {
      if (current) warnings.push(`要素表里「${item.key}」是空的，一页通这一格没有改。`);
      continue;
    }
    if (!current) {
      warnings.push(`一页通的产品要素里没有「${item.key}」，没有补进表格。`);
      continue;
    }
    const valueParagraphs = current.values.filter(paragraph => paragraph.text.trim());
    if (!valueParagraphs.length) {
      warnings.push(`一页通的「${item.key}」没有可改的正文。`);
      continue;
    }
    const joined = valueParagraphs.map(paragraph => paragraph.text.trim()).join('\n');
    if (compact(joined) === compact(item.value)) continue;
    const [first, ...rest] = valueParagraphs;
    const before = first.text.trim();
    edits.push({
      paragraphId: first.id,
      start: first.text.indexOf(before),
      before,
      after: item.value,
      label: item.key,
      source: 'element',
    });
    for (const paragraph of rest) {
      edits.push({ deleteParagraphId: paragraph.id, label: item.key, source: 'element' });
    }
    labels.push(item.key);
  }
  return { edits, warnings, labels };
}

function findElementTable(paragraphs) {
  const tables = new Map();
  for (const paragraph of paragraphs) {
    if (!paragraph.tableId || paragraph.rowIndex < 0) continue;
    const table = tables.get(paragraph.tableId) || { part: paragraph.part, rows: new Map() };
    const row = table.rows.get(paragraph.rowIndex) || [];
    row.push(paragraph);
    table.rows.set(paragraph.rowIndex, row);
    tables.set(paragraph.tableId, table);
  }
  let best = null;
  for (const table of tables.values()) {
    const rows = [...table.rows.entries()]
      .sort((left, right) => left[0] - right[0])
      .map(([, items]) => ({
        label: items.filter(item => item.cellIndex === 0).map(item => item.text).join('').trim(),
        values: items.filter(item => item.cellIndex > 0),
      }));
    const keys = rows.map(row => canonicalLabel(row.label)).filter(Boolean);
    if (!keys.includes('产品名称') || keys.length < 4) continue;
    const score = keys.length + (table.part === 'word/document.xml' ? 10 : 0);
    if (!best || score > best.score) best = { rows, score };
  }
  return best;
}

function canonicalLabel(text) {
  const value = String(text || '')
    .replace(/\s+/g, '')
    .replace(/（[^）]*）|\([^)]*\)/g, '')
    .replace(/[：:]/g, '');
  for (const [key, pattern] of LABELS) {
    if (pattern.test(value)) return key;
  }
  return '';
}

function compact(value) {
  return String(value || '').replace(/\s+/g, '').replace(/\(/g, '（').replace(/\)/g, '）');
}
