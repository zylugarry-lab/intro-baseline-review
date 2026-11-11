const WORD = '中低风险|中高风险|中风险|低风险|高风险';
const PAIR_CODE_FIRST = `R[1-5]\\s*[（(]\\s*(?:${WORD})\\s*[）)]`;
const PAIR_WORD_FIRST = `(?:${WORD})\\s*[（(]\\s*R[1-5]\\s*[）)]`;
const RATING = `(?:${PAIR_CODE_FIRST}|${PAIR_WORD_FIRST}|R[1-5]|${WORD})`;
const RISK_WORD = { R1: '低风险', R2: '中低风险', R3: '中风险', R4: '中高风险', R5: '高风险' };
const CCB_TYPE = { C1: '谨慎型', C2: '稳健型', C3: '平衡型', C4: '进取型', C5: '激进型' };
const CODE_OF = { 低风险: 'R1', 中低风险: 'R2', 中风险: 'R3', 中高风险: 'R4', 高风险: 'R5' };

export function parseRiskStatement(text) {
  const source = String(text || '');
  const sales = source.match(new RegExp(`销售机构风险评级为\\s*(R[1-5])\\s*[（(]?\\s*(${WORD})?`));
  const managed = source.match(new RegExp(`风险等级评定为\\s*(R[1-5])\\s*[（(]\\s*(${WORD})`))
    || source.match(new RegExp(`风险等级评定为\\s*(${WORD})\\s*[（(]\\s*(R[1-5])`))
    || source.match(new RegExp(`(?:^|[，,。\\s])(R[1-5])\\s*[（(]\\s*(${WORD})`))
    || source.match(new RegExp(`(?:^|[，,。\\s])(${WORD})\\s*[（(]\\s*(R[1-5])`));
  let code = '';
  let word = '';
  if (managed) {
    if (/^R[1-5]$/.test(managed[1])) {
      code = managed[1];
      word = managed[2] || '';
    } else {
      word = managed[1];
      code = managed[2] || '';
    }
  }
  if (code && !word) word = RISK_WORD[code] || '';
  if (word && !code) code = CODE_OF[word] || '';
  const salesCode = sales?.[1] || '';
  return {
    code,
    word,
    salesCode,
    salesWord: salesCode ? (sales?.[2] || RISK_WORD[salesCode] || '') : '',
  };
}

export function planRiskEdits(paragraphs, riskText, skipIds = null) {
  const canonical = String(riskText || '').trim();
  if (!canonical) return { edits: [], warnings: [] };
  const risk = parseRiskStatement(canonical);
  if (!risk.code && !risk.salesCode) {
    return { edits: [], warnings: ['要素表里的风险等级没能识别出 R 等级，一页通的风险揭示没有改。'] };
  }
  const rows = paragraphs.filter(paragraph => paragraph.part === 'word/document.xml');
  const edits = [];
  let seen = false;
  let skipped = false;
  for (let index = 0; index < rows.length; index += 1) {
    const paragraph = rows[index];
    if (skipIds?.has(paragraph.id)) {
      skipped = true;
      continue;
    }
    const trimmed = paragraph.text.trim();
    if (!isRiskStatement(trimmed)) continue;
    seen = true;
    const tail = rows[index + 1] && isRiskTail(rows[index + 1].text) ? rows[index + 1] : null;
    const salesInMain = /销售机构风险评级为/.test(trimmed);
    pushReplacement(edits, paragraph, rewriteRisk(trimmed, risk, { includeSales: !tail || salesInMain }));
    if (tail && !salesInMain) pushReplacement(edits, tail, rewriteTail(tail.text.trim(), risk));
  }
  const warnings = seen || skipped ? [] : ['这份一页通里没有找到风险等级表述，底部风险揭示没有改。'];
  return { edits, warnings };
}

function pushReplacement(edits, paragraph, after) {
  const before = paragraph.text.trim();
  if (compact(before) === compact(after)) return;
  edits.push({
    paragraphId: paragraph.id,
    start: paragraph.text.indexOf(before),
    before,
    after,
    label: '风险揭示',
    source: 'risk',
  });
}

function rewriteRisk(text, risk, { includeSales }) {
  let next = replaceManaged(text, risk);
  next = replaceLead(next, risk);
  next = replaceSuitability(next, risk);
  if (includeSales || /销售机构风险评级为/.test(next)) next = upsertSales(next, risk);
  return next;
}

function rewriteTail(text, risk) {
  return upsertSales(text, risk);
}

function replaceManaged(text, risk) {
  if (!risk.code && !risk.word) return text;
  return text.replace(new RegExp(`风险等级评定为\\s*${RATING}`), match => {
    const sample = match.replace(/^风险等级评定为\s*/, '');
    return `风险等级评定为${formatPair(sample, risk)}`;
  });
}

function replaceLead(text, risk) {
  if (!risk.code && !risk.word) return text;
  return text.replace(new RegExp(`^(\\d+[、.．]\\s*)?(?:${PAIR_CODE_FIRST}|${PAIR_WORD_FIRST})`), (full, prefix = '') => (
    `${prefix}${formatPair(full.slice(prefix.length), risk)}`
  ));
}

function replaceSuitability(text, risk) {
  if (!risk.code) return text;
  const level = Number(risk.code.slice(1));
  if (!level) return text;
  let next = text.replace(/C[1-5]\s*[-–~到至]\s*C5/g, `C${level}-C5`);
  next = next.replace(new RegExp(`C[1-5]\\s*[（(][^）)]+[）)](?:\\s*[、,，]\\s*C[1-5]\\s*[（(][^）)]+[）)])+`, 'g'), list => {
    const open = list.includes('(') && !list.includes('（') ? '(' : '（';
    const close = open === '(' ? ')' : '）';
    const separator = list.includes('、') ? '、' : '，';
    const items = [];
    for (let index = level; index <= 5; index += 1) items.push(`C${index}${open}${CCB_TYPE[`C${index}`]}${close}`);
    return items.join(separator);
  });
  return next;
}

function upsertSales(text, risk) {
  const current = text.match(new RegExp(`该产品销售机构风险评级为\\s*(R[1-5])\\s*[（(]\\s*(${WORD})`));
  if (!risk.salesCode) {
    if (!current) return text;
    return text
      .replace(/[，,]?\s*该产品销售机构风险评级为\s*R[1-5]\s*[（(][^）)]+[）)]\s*。?/g, '')
      .replace(/[，,]\s*。/g, '。')
      .replace(/。{2,}/g, '。');
  }
  if (current && current[1] === risk.salesCode && current[2] === risk.salesWord) return text;
  const open = current?.[0].includes('(') && !current[0].includes('（') ? '(' : (text.includes('(') && !text.includes('（') ? '(' : '（');
  const close = open === '(' ? ')' : '）';
  const clause = `该产品销售机构风险评级为${risk.salesCode}${open}${risk.salesWord}${close}`;
  const pattern = /[，,]?\s*该产品销售机构风险评级为\s*R[1-5]\s*[（(][^）)]+[）)]\s*。?/g;
  if (current) {
    return text.replace(pattern, (match, offset, source) => {
      const needsComma = offset > 0 && !/[。；;，,]/.test(source[offset - 1]);
      return `${needsComma ? '，' : ''}${clause}${/。\s*$/.test(match) ? '。' : ''}`;
    });
  }
  const trimmed = text.replace(/\s+$/, '');
  if (trimmed.endsWith('。')) return `${trimmed.slice(0, -1)}，${clause}。`;
  return `${trimmed}，${clause}。`;
}

export function applySalesRating(text, salesCode) {
  const risk = parseRiskStatement(text);
  return upsertSales(text, {
    ...risk,
    salesCode: salesCode || '',
    salesWord: salesCode ? (RISK_WORD[salesCode] || '') : '',
  });
}

function formatPair(sample, risk) {
  const codeFirst = /^\s*R[1-5]/.test(sample);
  const open = sample.includes('(') && !sample.includes('（') ? '(' : '（';
  const close = open === '(' ? ')' : '）';
  if (codeFirst) return `${risk.code}${open}${risk.word}${close}`;
  return `${risk.word}${open}${risk.code}${close}`;
}

function isRiskStatement(text) {
  if (!text || text.length > 500) return false;
  if (/本资料仅为宣传|不得转载|入市须谨慎/.test(text) && !/风险等级评定为/.test(text)) return false;
  if (/风险等级评定为/.test(text) && new RegExp(`R[1-5]|${WORD}`).test(text)) return true;
  if (/销售机构风险评级为\s*R[1-5]/.test(text) && text.length < 240) return true;
  return new RegExp(`^(?:\\d+[、.．]\\s*)?(?:${PAIR_CODE_FIRST}|${PAIR_WORD_FIRST})`).test(text)
    && /销售机构|风险/.test(text)
    && text.length < 240;
}

function isRiskTail(text) {
  const value = text.trim();
  if (!value || value.length > 180 || /风险等级评定为/.test(value)) return false;
  return /^投资者/.test(value) || /风险等级评定以(?:销售|代销)机构/.test(value);
}

function compact(value) {
  return String(value).replace(/\s+/g, '');
}
