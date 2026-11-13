import { existsSync, readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { buildManagerWorkbook, parseManagerWorkbook, buildElementWorkbook } from '../src/excel.js';
import { companySections, selectVariant } from '../src/company.js';
import { refreshBiography, tenureYears } from '../src/years.js';
import { downloadName, readDocxParagraphs, reviewOnePager } from '../src/review.js';
import { textDiff } from '../src/diff.js';
import { applyParagraphEdits } from '../src/revise.js';
import { collectParagraphs, parseXml } from '../src/xml.js';
import { planRiskEdits, parseRiskStatement } from '../src/risk.js';
import { composeElements, contractManagerNames, extractCore } from '../src/elements.js';
import { planCoverNameEdits, planElementEdits } from '../src/element-review.js';
import { planManagerEdits } from '../src/managers.js';
import { planLockEdits } from '../src/lock.js';
import { buildPoster, posterCopy, posterFileName, posterLayout, wrapTitle } from '../src/poster.js';
import { binaryDocText, wordKind } from '../src/word.js';

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

const elementBook = await buildElementWorkbook([['产品名称', '君享测试'], ['产品风险等级', 'R2（中低风险）\n该产品销售机构风险评级为R3（中风险）。']]);
const elementZip = await JSZip.loadAsync(elementBook);
const elementSheet = await elementZip.file('xl/worksheets/sheet1.xml').async('string');
const elementWorkbook = await elementZip.file('xl/workbook.xml').async('string');
check(elementWorkbook.includes('产品要素表'), '要素表 Excel 的工作表名是产品要素表');
check(elementSheet.includes('君享测试') && elementSheet.includes('销售机构风险评级为R3'), '要素表 Excel 写入当前格子里的内容');
const elementStyles = await elementZip.file('xl/styles.xml').async('string');
check(elementStyles.includes('style="thin"'), '要素表格子之间有分割线');
check(/ht="18"/.test(elementSheet), '短内容的行高收紧');
const multiBook = await buildElementWorkbook([
  { name: '君享年华定开债8号', rows: [['产品名称', '年华8号']] },
  { name: '君享年华定开债8号', rows: [['产品名称', '年华9号']] },
]);
const multiZip = await JSZip.loadAsync(multiBook);
const multiWorkbook = await multiZip.file('xl/workbook.xml').async('string');
const secondSheet = await multiZip.file('xl/worksheets/sheet2.xml').async('string');
check(multiWorkbook.includes('name="君享年华定开债8号"') && multiWorkbook.includes('name="君享年华定开债8号2"'), '多份合同分成不同工作表');
check(secondSheet.includes('年华9号'), '第二张工作表写入第二份合同的要素');

const docBytes = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, ...Buffer.from('集合计划的基本情况\n集合资产管理计划\n', 'utf16le')]);
check(wordKind(docBytes) === 'doc', '旧版 DOC 按文件头识别');
check(binaryDocText(docBytes).includes('基本情况'), '旧版 DOC 能读出正文');
const docxBytes = read(onePagers[2]);
check(wordKind(docxBytes) === 'docx', 'DOCX 按文件头识别');
const renamed = await reviewOnePager({
  bytes: docxBytes, fileName: '汇享全周期.doc', variant: 'general', records: managers.records, asOf: AS_OF,
  companyBytes: read('assets/company-introduction-general.docx'),
});
check(renamed.changed, '扩展名是 .doc、内容是 docx 的一页通可以核对');
check(downloadName('汇享全周期.doc') === '汇享全周期（介绍修订）.docx', '旧扩展名下载时改成 docx');
check(downloadName('旧名.docx', { productName: '国泰海通君盈瑞合增利FOF1号集合资产管理计划（产品代码：S12345）', asOf: '2026-11-10' }) === '【一页通  202611】国泰海通君盈瑞合增利FOF1号.docx', '下载文件名用制作年月和合同简称');
let binaryRejected = false;
try {
  await reviewOnePager({
    bytes: docBytes, fileName: '旧版.doc', variant: 'general', records: managers.records, asOf: AS_OF,
    companyBytes: read('assets/company-introduction-general.docx'),
  });
} catch (error) {
  binaryRejected = /旧版 \.doc/.test(error.message);
}
check(binaryRejected, '真正的旧版 DOC 一页通会说明写不了修订');

const contractPath = '/Users/luzhiyuan/Desktop/国泰海通/asset management/产品宣传/自动生成要素/国泰海通君盈瑞合增利FOF1号集合资产管理计划管理合同0831.doc';
if (existsSync(contractPath)) {
  const contract = binaryDocText(new Uint8Array(readFileSync(contractPath)));
  const core = extractCore(contract);
  check(core.name.includes('君盈瑞合增利') && core.risk.code === 'R3', '真实 DOC 合同能抽出名称和风险等级');
}

function cell(id, text, tableId, rowIndex, cellIndex) {
  return { id, part: 'word/document.xml', text, tableId, rowIndex, cellIndex, cellKey: `${tableId}:${rowIndex}:${cellIndex}` };
}
const elementRows = [
  cell('a', '产品名称', 'series', 0, 0), cell('b', '成立日', 'series', 0, 1),
  cell('c', '产品名称', 'facts', 0, 0), cell('d', '旧产品', 'facts', 0, 1),
  cell('e', '集合管理人', 'facts', 1, 0), cell('f', '上海国泰海通证券资产管理有限公司', 'facts', 1, 1),
  cell('g', '投资范围', 'facts', 2, 0), cell('h', '范围前半', 'facts', 2, 1), cell('i', '范围后半', 'facts', 2, 1),
  cell('j', '产品类型', 'facts', 3, 0), cell('k', '固定收益类集合资产管理计划', 'facts', 3, 1),
  cell('l', '参与 / 退出费率', 'facts', 4, 0), cell('m', '0%', 'facts', 4, 1),
];
const elementPlan = planElementEdits(elementRows, [
  { label: '产品名称', value: '新产品（产品代码：S1）' },
  { label: '管理人', value: '上海国泰海通证券资产管理有限公司' },
  { label: '投资范围', value: '新的投资范围' },
  { label: '产品类型', value: '固定收益类集合资产管理计划' },
  { label: '参与/退出费率', value: '0%' },
  { label: '投资策略', value: '合同里的策略' },
]);
check(elementPlan.labels.join('、') === '产品名称、投资范围', '只改产品要素表里对不上的格子');
check(elementPlan.edits.some(edit => edit.before === '旧产品' && edit.after.includes('S1')), '产品名称写成要素表的文字');
check(elementPlan.edits.some(edit => edit.deleteParagraphId === 'i'), '同一格里多出来的段落删掉');
check(elementPlan.warnings.some(warning => warning.includes('投资策略')), '一页通没有的要素行会提示');
const feePlan = planElementEdits(elementRows, [
  { label: '产品名称', value: '旧产品' },
  { label: '管理人', value: '上海国泰海通证券资产管理有限公司' },
  { label: '投资范围', value: '范围前半\n范围后半' },
  { label: '产品类型', value: '固定收益类集合资产管理计划' },
  { label: '参与/退出费率', value: '认购费率：0.2%。' },
]);
check(feePlan.edits.some(edit => edit.before === '0%' && edit.after === '认购费率：0.2%。'), '参与/退出费率仍对上一页通里的同名格子');
const cover = planCoverNameEdits([
  { id: 'a', part: 'word/document.xml', text: '国泰君安君享年华定开债8号' },
  { id: 'b', part: 'word/document.xml', text: '集合资产管理计划' },
  { id: 'c', part: 'word/document.xml', text: '产品有风险，投资须谨慎。' },
  { id: 'd', part: 'word/document.xml', text: '国泰君安君享年华定开债8号集合资产管理计划适合稳健投资者。' },
], '国泰海通建盈君享年华定开债8号集合资产管理计划（产品代码：S1）');
check(cover.edits[0]?.before === '国泰君安君享年华定开债8号' && cover.edits[0].after === '国泰海通建盈君享年华定开债8号集合资产管理计划', '一页通开头的产品名称改成合同名称');
check(cover.edits.some(edit => edit.deleteParagraphId === 'b') && !cover.edits.some(edit => edit.paragraphId === 'd'), '封面名称换行后的“集合资产管理计划”删掉，正文里的产品名不动');

const factRun = await reviewOnePager({
  bytes: read(onePagers[0]), fileName: onePagers[0], variant: 'general', records: managers.records, asOf: AS_OF,
  companyBytes: read('assets/company-introduction-general.docx'),
  elements: [
    { label: '产品名称', value: '国泰海通君盈瑞合增利FOF1号集合资产管理计划（产品代码：S12345）' },
    { label: '管理人', value: '上海国泰海通证券资产管理有限公司' },
    { label: '托管人', value: '招商银行股份有限公司南京分行' },
    { label: '产品类型', value: '固定收益类集合资产管理计划' },
    { label: '存续期', value: '自成立之日起10年（可展期）' },
  ],
});
check(factRun.elementLabels.includes('产品名称') && !factRun.elementLabels.includes('管理人') && !factRun.elementLabels.includes('存续期'), '真实一页通只改不一致的产品要素');
const factParagraphs = await visibleParagraphs(factRun.bytes);
check(factParagraphs.some(paragraph => paragraph.text.includes('产品代码：S12345')), '修订后的产品要素表写入了要素表里的产品名称');
const coverRun = await reviewOnePager({
  bytes: read(onePagers[0]), fileName: onePagers[0], variant: 'general', records: [], asOf: AS_OF,
  companyBytes: read('assets/company-introduction-general.docx'),
  elements: [{ label: '产品名称', value: '国泰海通君盈瑞合增利FOF2号集合资产管理计划' }],
});
const coverParagraphs = await visibleParagraphs(coverRun.bytes);
check(coverParagraphs.some(paragraph => paragraph.text.includes('FOF2号集合资产管理计划') && paragraph.text.length < 40), '一页通封面产品名称改成合同里的名称');

const lockSentence = '7、本产品存在锁定期，委托人每笔份额需持有满180天，可能存在流动性风险。';
const openingSentence = '特别的，投资者每笔份额持有时间不得少于180天。合同另有约定的，从其约定。';
const unrelated = '保存期限自基金账户销户之日起不得少于20年。';
const lockPlan = planLockEdits([
  { id: 'word/document.xml:0', part: 'word/document.xml', text: lockSentence },
  { id: 'word/document.xml:1', part: 'word/document.xml', text: openingSentence },
  { id: 'word/document.xml:2', part: 'word/document.xml', text: unrelated },
], [{ label: '开放期', value: '每月开放。特别的，每笔份额需持有满90个自然日。' }]);
check(lockPlan.edits.length === 2 && lockPlan.edits.every(edit => edit.before === '180' && edit.after === '90'), '锁定期和开放期里的天数都改成要素表的天数');
check(planLockEdits([
  { id: 'word/document.xml:0', part: 'word/document.xml', text: lockSentence },
], [{ label: '开放期', value: '每笔份额需持有满180天。' }]).edits.length === 0, '天数已经一致时不改锁定期');
const withLock = composeElements(
  { opening: '每月开放。', lock: { days: '180', unit: '个自然日' }, performance: { mode: 'none' }, risk: {} },
  { channel: 'house' },
).find(row => row[0] === '开放期')[1];
check(withLock.includes('持有满180个自然日'), '合同里的锁定期天数写进开放期');
const scopeByRange = extractCore([
  '一、名称：国泰海通君盈荟萃精选大类资产全周期QDII9号1期集合资产管理计划。',
  '七、投资目标、主要投资方向、投资比例及产品风险等级',
  '1、投资目标',
  '本集合计划主要投资于境内外资产。',
  '2、投资范围',
  '（1）资产管理计划可根据资金管理的需求，投资于现金、银行存款、货币市场基金；',
  '（2）资产管理计划可投资于公募证券投资基金。',
  '特别揭示：本计划可投资于债券回购。',
  '3、投资比例',
  '（1）固定收益类资产：占集合计划资产总值的20%～100%（不含）。',
].join('\n'));
check(scopeByRange.scope.includes('投资于现金、银行存款') && scopeByRange.scope.includes('公募证券投资基金') && scopeByRange.scope.includes('债券回购') && !scopeByRange.scope.includes('\n'), '标题写成投资范围时也能抽出投资范围，并收成一段');
check(!scopeByRange.scope.includes('2、投资范围') && !scopeByRange.scope.includes('固定收益类资产'), '投资范围不含标题，也不把投资比例带进来');
const noFee = composeElements(extractCore('6、业绩报酬\n本集合计划不收取业绩报酬。\n7、其他费用'), { channel: 'house' }).find(row => row[0] === '业绩报酬')[1];
check(noFee === '本集合计划不收取业绩报酬。', '不收取业绩报酬时直接抄合同这句话');
const tierFees = composeElements(extractCore([
  '1、认购费率：参与金额<1000万，0.2%；参与金额≥1000万，0%。',
  '2、参与费率：参与金额<1000万，0.2%；参与金额≥1000万，0%。',
  '3、退出费率：0%。',
  '4、管理费率：',
  '本集合计划的年固定管理费率为1.0%。',
  '5、托管费率',
  '本集合计划的年托管费率为0.05%。',
].join('\n')), { channel: 'house' });
const tierRate = tierFees.find(row => row[0] === '参与/退出费率')[1];
check(tierRate.includes('认购费率：参与金额<1000万，0.2%；参与金额≥1000万，0%。') && tierRate.includes('参与费率：参与金额<1000万，0.2%；参与金额≥1000万，0%。') && tierRate.includes('退出费率0%'), '认购费和参与费的分档写进参与/退出费率');
check(tierFees.find(row => row[0] === '产品费率')[1] === '固定管理费率：1%/年；托管费率：0.05%/年', '分档参与费不改管理费和托管费');
check(composeElements(extractCore('1、认购费率：0%\n2、存续期参与费率：0%\n3、退出费率：0%\n'), { channel: 'house' }).find(row => row[0] === '参与/退出费率')[1] === '0%', '没有分档时参与费和退出费都是 0 仍写成 0%');
const named = contractManagerNames('投资经理的指定\n投资经理：\n王宜峰，8年证券从业经验。\n刘博洋，6年证券从业经验。\n投资经理的变更\n管理人可根据业务需要变更投资经理。');
check(named.join('、') === '王宜峰、刘博洋', '合同投资经理名单按指定段落提取');
const rosterParagraphs = [
  { id: 'h', part: 'word/document.xml', index: 1, text: '投资经理介绍' },
  { id: 'n1', part: 'word/document.xml', index: 2, text: '王宜峰：' },
  { id: 'b1', part: 'word/document.xml', index: 3, text: '旧简介。' },
  { id: 'n2', part: 'word/document.xml', index: 4, text: '侯少蒙：' },
  { id: 'b2', part: 'word/document.xml', index: 5, text: '侯少蒙的旧介绍。' },
  { id: 'end', part: 'word/document.xml', index: 6, text: '风险揭示' },
];
const rosterEdits = planManagerEdits(rosterParagraphs, managers.records, AS_OF, ['王宜峰', '赵博阳']);
check(rosterEdits.edits.some(edit => edit.deleteParagraphId === 'n2') && rosterEdits.edits.some(edit => edit.deleteParagraphId === 'b2'), '合同里没有的投资经理从一页通删除');
check(rosterEdits.reviews.some(item => item.name === '赵博阳' && item.status === 'inserted'), '合同里有、一页通没有的投资经理按底稿补上');
check(!rosterEdits.edits.some(edit => edit.deleteParagraphId === 'n1'), '合同里有的投资经理保留');

const posterRows = [
  { label: '产品名称', value: '国泰海通福星多元稳健19号集合资产管理计划' },
  { label: '管理人', value: '上海国泰海通证券资产管理有限公司' },
  { label: '首参金额', value: '单个客户首次参与金额不低于30万元，追加参与的最低金额为人民币1万元。' },
  { label: '业绩报酬', value: '本集合计划管理人不提取业绩报酬。' },
  { label: '开放期', value: '每个开放期开放5个工作日。' },
  { label: '参与/退出费率', value: '0%' },
  { label: '产品费率', value: '固定管理费率：0.6%/年；托管费率：0.02%/年' },
  { label: '产品风险等级', value: '管理人对于本集合计划的风险等级评定为R3（中风险），适合向合格投资者中的专业投资者和风险承受能力评估结果为C3-C5的普通投资者销售。' },
];
const poster = posterCopy({ elements: posterRows, channel: 'house', code: '953794', date: '9月28日 正式开放' });
check(poster.riskLine === '风险等级：R3（中风险）', '海报产品名称下用合同里的风险等级');
check(poster.suitability === '适合合格投资者中的专业投资者和风险承受能力C3-C5的普通投资者。', '通用版海报把适合投资者收成 C3-C5');
check(poster.rows.some(row => row.label === '产品代码：' && row.value === '953794') && poster.rows.some(row => row.value === '9月28日 正式开放'), '海报带上产品代码和开放日期');
const posterCcb = posterCopy({ elements: posterRows, channel: 'ccb' });
check(posterCcb.suitability.includes('C3（平衡型）') && posterCcb.suitability.includes('C5（激进型）') && !posterCcb.suitability.includes('C1'), '建行海报从对应等级写到 C5');
check(posterFileName(posterRows[0].value, AS_OF) === '【海报  202611】国泰海通福星多元稳健19号.pptx', '海报文件名用制作年月和合同简称');
const posterBytes = await buildPoster(read('assets/poster-template.pptx'), { elements: posterRows, channel: 'house', code: '953794', date: '9月28日 正式开放' });
const posterXml = await JSZip.loadAsync(posterBytes).then(zip => zip.file('ppt/slides/slide1.xml').async('string'));
check(posterXml.includes('国泰海通福星') && posterXml.includes('多元稳健19号') && !posterXml.includes('星聚'), '海报模板换成合同产品名称');
check(posterXml.includes('风险等级：') && posterXml.includes('R3（中风险）') && posterXml.includes('C3-C5'), '海报写入风险等级大字和简写');
check(posterXml.includes('953794') && posterXml.includes('9月28日 正式开放') && posterXml.includes('参与费：0%') && posterXml.includes('退出费：0%') && posterXml.includes('管理费率：0.6%/年'), '海报产品费率写上参与费、退出费、管理费和托管费');
const tierPoster = posterCopy({
  elements: [
    { label: '产品名称', value: '测试集合资产管理计划' },
    { label: '参与/退出费率', value: '认购费率：参与金额<1000万，0.2%；参与金额≥1000万，0%。\n参与费率：参与金额<1000万，0.2%；参与金额≥1000万，0%。\n退出费率0%' },
    { label: '产品费率', value: '固定管理费率：1%/年；托管费率：0.05%/年' },
  ],
});
const tierFee = tierPoster.rows.find(row => row.label === '产品费率：').value;
check(tierFee.includes('认购费：参与金额<1000万，0.2%') && tierFee.includes('参与费：参与金额<1000万，0.2%') && tierFee.includes('退出费：0%') && tierFee.includes('托管费率：0.05%/年'), '分档时产品费率同时写出认购费、参与费和退出费');
check(!posterXml.includes('40万元'), '海报不再保留模板里的起购金额');
const frame = { slideH: 11878945, bodyWidth: 6000000, bodyTop: 7060000, titleWidth: 4200000, titleTop: 4540000, disclaimerWidth: 6400000 };
const shortLayout = posterLayout(poster, frame);
const longCopy = posterCopy({
  elements: posterRows.map(row => row.label === '业绩报酬' ? { ...row, value: '当年化收益率超过计提基准时提取业绩报酬。'.repeat(30) } : row),
  channel: 'house',
});
const longName = posterCopy({
  elements: posterRows.map(row => row.label === '产品名称' ? { ...row, value: '国泰海通君盈荟萃精选大类资产全周期QDII9号1期集合资产管理计划' } : row),
  channel: 'house',
});
const longLayout = posterLayout(longCopy, frame);
const nameLayout = posterLayout(longName, frame);
check(longLayout.body.sz < shortLayout.body.sz, '正文变长时字号缩小');
check(longLayout.body.y + longLayout.body.h <= longLayout.disclaimer.y, '正文不会压住底部风险揭示');
check(longLayout.disclaimer.y + longLayout.disclaimer.h <= frame.slideH, '风险揭示落在页面里面');
const nameLines = wrapTitle('国泰海通君盈大类资产全方略QDII1号', frame.titleWidth, 3200);
check(nameLines.some(line => line.includes('大类资产全方略')) && !nameLines.some(line => line.endsWith('大类资产')), '标题不把大类资产和全方略拆开');
check(wrapTitle('国泰海通君享恒利QDII2/QDII3号', frame.titleWidth, 3200).join('').includes('QDII2/QDII3'), '标题保留名称里的斜杠');
check(nameLayout.title.lines.every(line => !line.endsWith('大类资')), '长名称不会把大类资产拆成两行');
const starLayout = posterLayout(posterCopy({
  elements: [
    { label: '产品名称', value: '国泰海通星聚大类资产配置1号集合资产管理计划' },
    { label: '产品风险等级', value: 'R3（中风险）' },
  ],
}), {
  slideH: 11878945,
  bodyWidth: 6000000,
  bodyTop: 7066800,
  titleWidth: 4320000,
  titleNameWidth: 3420000,
  titleTop: 4896000,
  disclaimerWidth: 6400000,
});
check(starLayout.title.nameSz === 3200 && starLayout.title.subSz === 2800, '产品名称用32磅，副标题用28磅');
check(starLayout.title.lines.join('|') === '国泰海通星聚|大类资产配置1号', '星聚标题按词分成两行');
check(starLayout.title.y === 4896000 && starLayout.title.y + starLayout.title.h < 4896000 + 4.6 * 360000, '标题放在灯塔右侧、黄条上方');

const cycleFile = '/Users/luzhiyuan/Desktop/【一页通202604】君享大类资产全周期1号.docx';
if (existsSync(cycleFile)) {
  const cycle = await reviewOnePager({
    bytes: new Uint8Array(readFileSync(cycleFile)),
    fileName: '【一页通202604】君享大类资产全周期1号.docx',
    variant: 'general', records: [], asOf: AS_OF,
    companyBytes: read('assets/company-introduction-general.docx'),
    elements: [{ label: '开放期', value: '本集合计划每月开放。特别的，每笔份额需持有满90个自然日。' }],
  });
  check(cycle.lockUpdated, '全周期一页通识别了锁定期这句话');
  const cycleParagraphs = await visibleParagraphs(cycle.bytes);
  check(cycleParagraphs.some(paragraph => paragraph.text.includes('持有满90天')), '全周期一页通的锁定期改成90天');
  check(!cycleParagraphs.some(paragraph => paragraph.text.includes('持有满180天')), '全周期一页通不再保留持有满180天');
}

if (failed) {
  console.error(`\n${failed} checks failed`);
  process.exit(1);
}
console.log('\nall checks passed');
