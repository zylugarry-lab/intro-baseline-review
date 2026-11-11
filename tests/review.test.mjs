import { readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { parseManagerWorkbook } from '../src/excel.js';
import { companySections, selectVariant } from '../src/company.js';
import { refreshBiography, tenureYears } from '../src/years.js';
import { readDocxParagraphs, reviewOnePager } from '../src/review.js';
import { textDiff } from '../src/diff.js';
import { applyParagraphEdits } from '../src/revise.js';
import { collectParagraphs, parseXml } from '../src/xml.js';
import { planRiskEdits, parseRiskStatement } from '../src/risk.js';
import { composeElements } from '../src/elements.js';

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
globalThis.JSZip = JSZip;

const root = new URL('../', import.meta.url);
const read = name => new Uint8Array(readFileSync(new URL(name, root)));
const AS_OF = '2026-11-10';
let failed = 0;

function check(condition, message) {
  if (condition) console.log(`  ok  ${message}`);
  else { failed += 1; console.error(`  FAIL ${message}`); }
}

function visibleParagraphs(bytes) {
  return readDocxParagraphs(bytes);
}

const excel = read('assets/manager-introductions.xlsx');
const managers = await parseManagerWorkbook(excel, '经理介绍.xlsx');
check(managers.records.length > 40, `读到 ${managers.records.length} 位经理`);
const byName = name => managers.records.find(record => record.name === name);

const zhao = refreshBiography(byName('赵博阳'), AS_OF);
check(zhao.calculations.find(item => item.kind === 'work')?.years === tenureYears('2012-08-31', AS_OF), '赵博阳从业年限按日期计算');
check(zhao.text.startsWith(`${tenureYears('2012-08-31', AS_OF)}年证券期货从业经验`), '赵博阳简介开头年限已重算');
check(zhao.text.includes(`其中${tenureYears('2016-01-15', AS_OF)}年固定收益投资经验`), '赵博阳投资年限已重算');

const jia = refreshBiography(byName('贾聪'), AS_OF);
check(jia.text.includes('10年以上信用债投研经验'), '贾聪“10年以上投研”不被从业日期覆盖');

const zhu = refreshBiography(byName('朱咏熙'), AS_OF);
check(zhu.text.includes(`${tenureYears('2025-03-10', AS_OF, 1)}年投资经历`), '朱咏熙的小数投资年限按一位小数更新');

const liu = refreshBiography(byName('刘博洋'), AS_OF);
check(liu.calculations.some(item => item.kind === 'investment' && item.years < 1), '刘博洋未满一年的投资年限保留小数');

const general = await visibleParagraphs(read('assets/company-introduction-general.docx'));
const ccb = await visibleParagraphs(read('assets/company-introduction-ccb.docx'));
const generalText = companySections(selectVariant(general, 'general')).flatMap(section => section.columns.flatMap(column => column.entries.map(entry => entry.text))).join('\n');
const ccbText = companySections(selectVariant(ccb, 'ccb')).flatMap(section => section.columns.flatMap(column => column.entries.map(entry => entry.text))).join('\n');
check(generalText.includes('全面领先') && generalText.includes('资本实力位居行业最前列'), '通用版保留全面领先和资本实力');
check(!ccbText.includes('全面领先') && !ccbText.includes('资本实力位居行业最前列'), '建行版去掉全面领先和资本实力');
check(textDiff('长期、持续、全面领先的综合金融服务商', '长期、持续的综合金融服务商')[0].before === '、全面领先', '股东介绍只标出不同的词');

const tiny = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>8年证券从业经验，1年投资经验。</w:t></w:r></w:p></w:body></w:document>`;
const tinyDoc = parseXml(tiny);
const tinyParagraph = collectParagraphs(tinyDoc, 'word/document.xml')[0];
applyParagraphEdits(tinyParagraph.node, [{ start: 0, before: '8', after: '9' }], { author: '介绍底稿', date: '2026-11-10T00:00:00Z', next: (() => { let n = 0; return () => { n += 1; return n; }; })() });
const tinyXml = new XMLSerializer().serializeToString(tinyDoc);
check(tinyXml.includes('w:del') && tinyXml.includes('w:delText') && tinyXml.includes('>8<'), '年限修改写成删除');
check(tinyXml.includes('w:ins') && tinyXml.includes('>9<'), '年限修改写成插入');

mkdirSync(new URL('tests/output/', root), { recursive: true });
const onePagers = [
  '【一页通202608】国泰海通君盈瑞合增利FOF1号.docx',
  '【一页通202606】国泰海通君享恒利QDII2 3号.docx',
  '【一页通 202611】汇享全周期QDII1号.docx',
];
for (const variant of ['general', 'ccb']) {
  for (const fileName of onePagers) {
    const result = await reviewOnePager({
      bytes: read(fileName), fileName, variant, records: managers.records, asOf: AS_OF,
      companyBytes: read(variant === 'ccb' ? 'assets/company-introduction-ccb.docx' : 'assets/company-introduction-general.docx'),
    });
    check(result.managers.some(item => item.status !== 'unchanged' || result.changed), `${fileName} / ${variant} 完成核对，经理 ${result.managers.map(item => item.name).join('、') || '无'}`);
    if (!result.changed) continue;
    const outputName = `tests/output/${variant}-${fileName.replace(/[【】\s]/g, '')}`;
    writeFileSync(new URL(outputName, root), result.bytes);
    const xml = await JSZip.loadAsync(result.bytes).then(zip => zip.file('word/document.xml').async('string'));
    check(xml.includes('w:ins') && xml.includes('w:del'), `${variant} ${fileName} 含有修订标记`);
    const again = await visibleParagraphs(result.bytes);
    check(again.some(paragraph => paragraph.text.includes('国泰海通')), `${variant} ${fileName} 修订后仍能读出正文`);
  }
}

const fofCcb = await reviewOnePager({
  bytes: read(onePagers[0]), fileName: onePagers[0], variant: 'ccb', records: managers.records, asOf: AS_OF,
  companyBytes: read('assets/company-introduction-ccb.docx'),
});
const fofXml = await JSZip.loadAsync(fofCcb.bytes).then(zip => zip.file('word/document.xml').async('string'));
check(fofXml.includes('全面领先') && fofXml.includes('delText'), '建行版比对把“全面领先”标成删除');
const wang = refreshBiography(byName('王宜峰'), AS_OF);
const fofParagraphs = await visibleParagraphs(fofCcb.bytes);
check(fofParagraphs.some(paragraph => paragraph.text.includes(wang.text.split('\n')[0].slice(0, 12))), '王宜峰简介已按底稿改写');
check(fofCcb.companyLabels.includes('股东介绍'), '未改底稿时，建行版会改一页通里的股东介绍');

const generalShare = general.find(paragraph => paragraph.text.includes('全面领先') && paragraph.text.includes('综合金融服务商'));
const ccbShare = ccb.find(paragraph => paragraph.text.includes('综合金融服务商') && !paragraph.text.includes('全面领先'));
const editedCompany = await reviewOnePager({
  bytes: read(onePagers[0]), fileName: onePagers[0], variant: 'ccb', records: managers.records, asOf: AS_OF,
  companyBytes: read('assets/company-introduction-ccb.docx'),
  companyTextById: { [ccbShare.id]: generalShare.text },
});
check(!editedCompany.companyLabels.includes('股东介绍'), '网页上改过的公司介绍会用于比对');

const riskCell = composeElements(
  { risk: { code: 'R2', word: '中低风险' }, name: '测试产品', performance: { mode: 'none' } },
  { channel: 'house', salesCode: 'R3', code: '' },
).find(row => row[0] === '产品风险等级')[1];
check(riskCell.includes('评定为R2（中低风险）') && riskCell.includes('该产品销售机构风险评级为R3（中风险）'), '要素表风险话术带上销售机构评级');
const parsedRisk = parseRiskStatement(riskCell);
check(parsedRisk.code === 'R2' && parsedRisk.salesCode === 'R3', '从要素话术里分出管理人等级和销售机构等级');

const bottom = '5、管理人对于本集合计划的风险等级评定为R3（中风险），适合合格投资者中的专业投资者以及风险承受能力为C3-C5的普通投资者，产品风险等级以销售机构为准。';
const top = 'R2 (中低风险)，销售机构对于本集合计划的风险等级评定以销售机构发布为准。该产品销售机构风险评级为 R4(中高风险)。';
const noise = '1、本资料仅为宣传材料，具体以风险揭示书为准。投资有风险，入市须谨慎。';
const riskEdits = planRiskEdits([
  { id: 'word/document.xml:0', part: 'word/document.xml', text: top },
  { id: 'word/document.xml:1', part: 'word/document.xml', text: bottom },
  { id: 'word/document.xml:2', part: 'word/document.xml', text: noise },
], riskCell).edits;
check(riskEdits.length === 2, '只改两处风险等级表述');
const topAfter = riskEdits.find(edit => edit.before.startsWith('R2')).after;
const bottomAfter = riskEdits.find(edit => edit.before.startsWith('5、')).after;
check(topAfter.includes('R2(中低风险)') && topAfter.includes('R3(中风险)') && !topAfter.includes('R4'), '要素格里的销售机构评级改成 R3');
check(bottomAfter.includes('评定为R2（中低风险）') && bottomAfter.includes('C2-C5') && bottomAfter.includes('该产品销售机构风险评级为R3（中风险）'), '底部风险揭示跟着要素改等级和销售机构评级');
check(planRiskEdits([{ id: 'word/document.xml:1', part: 'word/document.xml', text: bottomAfter }], riskCell).edits.length === 0, '已经一致时不再改');

const ccbRisk = composeElements(
  { risk: { code: 'R2', word: '中低风险' }, performance: { mode: 'none' } },
  { channel: 'ccb', salesCode: '', code: '' },
).find(row => row[0] === '产品风险等级')[1];
const ccbSentence = '管理人对于本集合计划的风险等级评定为R3（中风险），适合合格投资者中的专业投资者以及风险承受能力为C3（平衡型）、C4（进取型）、C5（激进型）的普通投资者。代销机构对本集合计划的风险等级评定以代销机构发布为准。';
const ccbAfter = planRiskEdits([{ id: 'word/document.xml:0', part: 'word/document.xml', text: ccbSentence }], ccbRisk).edits[0].after;
check(ccbAfter.includes('评定为R2（中低风险）'), '建行句保留代号在前的写法');
check(ccbAfter.includes('C2（稳健型）、C3（平衡型）、C4（进取型）、C5（激进型）'), '适合的投资者从 C2 起');

const split = planRiskEdits([
  { id: 'word/document.xml:0', part: 'word/document.xml', text: '管理人对于本集合计划的风险等级评定为R2（中低风险），适合向合格投资者中的专业投资者和风险承受能力评估结果为C2-C5的普通' },
  { id: 'word/document.xml:1', part: 'word/document.xml', text: '投资者销售，销售机构对本集合计划的风险等级评定以销售机构发布为准。' },
], riskCell).edits;
check(split.some(edit => edit.before.startsWith('投资者') && edit.after.includes('该产品销售机构风险评级为R3')), '断行的后半句补上销售机构评级');

const riskRun = await reviewOnePager({
  bytes: read(onePagers[2]), fileName: onePagers[2], variant: 'general', records: managers.records, asOf: AS_OF,
  companyBytes: read('assets/company-introduction-general.docx'), elementRisk: riskCell,
});
check(riskRun.riskUpdated, '一页通风险揭示被改写');
const riskParagraphs = await visibleParagraphs(riskRun.bytes);
check(riskParagraphs.some(paragraph => paragraph.text.includes('评定为R2（中低风险）') && paragraph.text.includes('该产品销售机构风险评级为R3（中风险）')), '修订稿里的风险揭示已换成要素等级');
const riskXml = await JSZip.loadAsync(riskRun.bytes).then(zip => zip.file('word/document.xml').async('string'));
check(riskXml.includes('w:del') && riskXml.includes('w:ins'), '风险修改保留修订标记');

if (failed) {
  console.error(`\n${failed} checks failed`);
  process.exit(1);
}
console.log('\nall checks passed');
