import { CCB_TYPE, RISK_WORD } from './elements.js';
import { childList, children, nodeList, parseXml, serializeXml } from './xml.js';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

const DISCLAIMER_HEAD = [
  '1、本资料仅为宣传材料，不作为任何法律文件，仅供合格投资者参考，不得通过任何渠道进行公开宣传，不得印刷宣传，包含的内容仅供参考，具体以本集合计划合同、说明书及风险揭示书等法律文本为准，文件未经允许不得转载或给第三方传阅。',
  '2、本产品为非保本型产品，存在净值波动，投资风险由投资者自行承担，如需了解产品详情，请阅读产品资产管理合同、资产管理计划说明书等法律文件。投资有风险，入市需谨慎。',
  '3、管理人承诺以诚实信用、勤勉尽责的原则管理产品。产品的过往业绩和净值高低并不预示其未来业绩表现，管理人管理的其他产品的业绩也不构成对本产品业绩表现的保证。',
].join('');

function flat(text) {
  return String(text || '').replace(/\s*\n+\s*/g, '').replace(/[ \t]{2,}/g, ' ').trim();
}

function field(elements, label) {
  return elements?.find(row => row.label === label)?.value || '';
}

function productName(value) {
  return flat(value).split('\n')[0].replace(/[（(]\s*产品代码[：:][^）)]*[）)]/g, '').trim();
}

export function readPosterRisk(riskValue) {
  const text = String(riskValue || '');
  const code = text.match(/R[1-5]/)?.[0] || '';
  const word = text.match(/(低风险|中低风险|中风险|中高风险|高风险)/)?.[0] || RISK_WORD[code] || '';
  return { code, word };
}

function suitability(code, channel) {
  const level = Number(String(code || '').slice(1));
  if (!level) return '';
  if (channel === 'ccb') {
    const list = [];
    for (let index = level; index <= 5; index += 1) list.push(`C${index}（${CCB_TYPE[`C${index}`]}）`);
    return `适合合格投资者中的专业投资者以及风险承受能力为${list.join('、')}的普通投资者。`;
  }
  return `适合合格投资者中的专业投资者和风险承受能力C${level}-C5的普通投资者。`;
}

function feeLine(participate, productFee) {
  const left = flat(participate);
  const right = flat(productFee).replace(/固定管理费率/g, '管理费率');
  const parts = [];
  if (left) parts.push(/费率/.test(left) ? left : `参与/退出费率：${left}`);
  if (right) parts.push(right);
  return parts.join('    ');
}

export function posterCopy({ elements, channel = 'house', code = '', date = '', salesCode = '' } = {}) {
  const name = productName(field(elements, '产品名称'));
  const risk = readPosterRisk(field(elements, '产品风险等级'));
  const riskLine = risk.code
    ? `风险等级：${risk.code}（${risk.word}）`
    : (risk.word ? `风险等级：${risk.word}` : '');
  const fit = suitability(risk.code, channel);
  const rows = [];
  if (String(code || '').trim()) rows.push({ label: '产品代码：', value: String(code).trim() });
  if (String(date || '').trim()) rows.push({ label: '开放/销售日期：', value: String(date).trim() });
  rows.push(
    { label: '起购金额：', value: flat(field(elements, '首参金额')) },
    { label: '业绩报酬计提：', value: flat(field(elements, '业绩报酬')) },
    { label: '开放期：', value: flat(field(elements, '开放期')) },
    { label: '产品费率：', value: feeLine(field(elements, '参与/退出费率'), field(elements, '产品费率')) },
  );
  if (riskLine) rows.push({ label: riskLine, value: '' });
  if (fit) rows.push({ label: '', value: fit });
  const manager = flat(field(elements, '管理人')) || '上海国泰海通证券资产管理有限公司';
  const distributor = channel === 'ccb'
    ? '建设银行作为代销机构不承担产品的投资、兑付和风险管理责任。'
    : '销售机构不承担产品的投资和兑付责任。';
  let pointFive = '';
  if (fit) {
    pointFive = `5、${risk.code ? `管理人对于本集合计划的风险等级评定为${risk.code}，` : ''}${fit}`;
    if (salesCode && RISK_WORD[salesCode]) pointFive += `该产品销售机构风险评级为${salesCode}（${RISK_WORD[salesCode]}）。`;
  }
  const disclaimer = `${DISCLAIMER_HEAD}4、本产品由${manager}发行与管理，${distributor}${pointFive}`;
  return { name, riskLine, suitability: fit, rows, disclaimer };
}

export function posterFileName(productNameValue, asOf) {
  const stamp = String(asOf || '').match(/^(\d{4})-(\d{2})/);
  const short = productName(productNameValue).replace(/集合资产管理计划/g, '').replace(/[/／\\:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();
  if (stamp && short) return `【海报  ${stamp[1]}${stamp[2]}】${short}.pptx`;
  return `${short || '海报'}.pptx`;
}

function textOf(paragraph) {
  return nodeList(paragraph.getElementsByTagNameNS(A, 't')).map(node => node.textContent || '').join('');
}

function clearRuns(paragraph) {
  for (const child of childList(paragraph)) {
    if (child.localName === 'r' || child.localName === 'br') paragraph.removeChild(child);
  }
}

function makeRun(doc, sampleRun, text, options = {}) {
  const run = doc.createElementNS(A, 'a:r');
  const source = sampleRun?.getElementsByTagNameNS(A, 'rPr')?.[0];
  const props = source ? source.cloneNode(true) : doc.createElementNS(A, 'a:rPr');
  if (options.sz) props.setAttribute('sz', options.sz);
  if (options.bold === true) props.setAttribute('b', '1');
  if (options.bold === false) props.removeAttribute('b');
  run.appendChild(props);
  const node = doc.createElementNS(A, 'a:t');
  node.textContent = text;
  run.appendChild(node);
  return { run, props };
}

function writeParagraph(paragraph, pieces) {
  const sample = childList(paragraph).find(node => node.localName === 'r');
  clearRuns(paragraph);
  const end = childList(paragraph).find(node => node.localName === 'endParaRPr');
  pieces.forEach((piece, index) => {
    if (piece.breakBefore) {
      const made = makeRun(paragraph.ownerDocument, sample, '', piece);
      const br = paragraph.ownerDocument.createElementNS(A, 'a:br');
      br.appendChild(made.props);
      paragraph.insertBefore(br, end || null);
    }
    if (!piece.text) return;
    const made = makeRun(paragraph.ownerDocument, piece.sample || sample, piece.text, piece);
    if (index === 0 && piece.breakBefore) {
      // props already used by the break; rebuild the run's props
    }
    paragraph.insertBefore(made.run, end || null);
  });
}

function shapeByName(doc, name) {
  for (const shape of nodeList(doc.getElementsByTagNameNS(P, 'sp'))) {
    const mark = shape.getElementsByTagNameNS(P, 'cNvPr')[0];
    if (mark?.getAttribute('name') === name) return shape;
  }
  throw new Error(`海报模板里没有「${name}」。`);
}

function bodyParagraphs(shape) {
  const body = shape.getElementsByTagNameNS(P, 'txBody')[0];
  return { body, paragraphs: children(body, A, 'p') };
}

function pairAfter(paragraphs, prefix) {
  const index = paragraphs.findIndex(paragraph => textOf(paragraph).trim().startsWith(prefix));
  if (index < 0) return null;
  const value = paragraphs.slice(index + 1).find(paragraph => textOf(paragraph).trim());
  return { label: paragraphs[index], value };
}

function prependPair(body, labelSample, valueSample, spacerSample, label, value) {
  const labelParagraph = labelSample.cloneNode(true);
  const valueParagraph = valueSample.cloneNode(true);
  const spacer = spacerSample.cloneNode(true);
  writeParagraph(labelParagraph, [{ text: label }]);
  writeParagraph(valueParagraph, [{ text: value }]);
  const first = children(body, A, 'p')[0];
  body.insertBefore(labelParagraph, first);
  body.insertBefore(valueParagraph, first);
  body.insertBefore(spacer, first);
}

function writeTitle(shape, name, riskLine) {
  const paragraph = children(shape.getElementsByTagNameNS(P, 'txBody')[0], A, 'p')[0];
  const main = name.replace(/集合资产管理计划/g, '').trim();
  const pieces = [{ text: main || name, sz: '3200', bold: true }];
  if (name.includes('集合资产管理计划')) pieces.push({ text: '集合资产管理计划', sz: '2800', bold: true, breakBefore: true });
  if (riskLine) pieces.push({ text: riskLine, sz: '2800', bold: true, breakBefore: true });
  writeParagraph(paragraph, pieces);
}

export async function buildPoster(templateBytes, spec) {
  const copy = posterCopy(spec);
  const zip = await globalThis.JSZip.loadAsync(templateBytes);
  const doc = parseXml(await zip.file('ppt/slides/slide1.xml').async('string'));
  writeTitle(shapeByName(doc, '标题 1'), copy.name, copy.riskLine);
  const disclaimer = bodyParagraphs(shapeByName(doc, '副标题 2')).paragraphs.find(paragraph => textOf(paragraph).trim())
    || bodyParagraphs(shapeByName(doc, '副标题 2')).paragraphs[0];
  writeParagraph(disclaimer, [{ text: copy.disclaimer }]);
  const { body, paragraphs } = bodyParagraphs(shapeByName(doc, '文本框 20'));
  const slots = [
    ['起购金额', copy.rows.find(row => row.label === '起购金额：')],
    ['业绩报酬', copy.rows.find(row => row.label === '业绩报酬计提：')],
    ['开放期', copy.rows.find(row => row.label === '开放期：')],
    ['产品费率', copy.rows.find(row => row.label === '产品费率：')],
    ['风险等级', { label: copy.riskLine, value: '' }],
    ['适合合格', { label: '', value: copy.suitability }],
  ];
  for (const [prefix, row] of slots) {
    const found = pairAfter(paragraphs, prefix);
    if (!found || !row) continue;
    if (prefix === '风险等级') {
      const runs = childList(found.label).filter(node => node.localName === 'r');
      writeParagraph(found.label, [
        { text: '风险等级：', sample: runs[0], bold: true },
        { text: copy.riskLine.replace(/^风险等级：/, ''), sample: runs[1] || runs[0], bold: false },
      ]);
    } else if (prefix === '适合合格') {
      writeParagraph(found.label, [{ text: row.value }]);
    } else {
      writeParagraph(found.label, [{ text: row.label }]);
      if (found.value) writeParagraph(found.value, [{ text: row.value }]);
    }
  }
  const labelSample = pairAfter(paragraphs, '起购金额')?.label;
  const valueSample = pairAfter(paragraphs, '起购金额')?.value;
  const spacerSample = paragraphs.find(paragraph => !textOf(paragraph).trim()) || labelSample;
  const extras = copy.rows.filter(row => row.label === '产品代码：' || row.label === '开放/销售日期：');
  for (const row of [...extras].reverse()) prependPair(body, labelSample, valueSample, spacerSample, row.label, row.value);
  zip.file('ppt/slides/slide1.xml', serializeXml(doc));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
