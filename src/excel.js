import { NS, children, elements, parseXml } from './xml.js';

const MAX_BYTES = 30 * 1024 * 1024;

function columnIndex(address) {
  let result = 0;
  for (const char of String(address).match(/^[A-Z]+/i)?.[0] || '') {
    result = result * 26 + char.toUpperCase().charCodeAt(0) - 64;
  }
  return result - 1;
}

function targetPath(base, target) {
  if (!target || /^[a-z]+:/i.test(target)) return '';
  const out = [];
  for (const part of `${base.slice(0, base.lastIndexOf('/') + 1)}${target}`.replace(/^\//, '').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') out.pop();
    else out.push(part);
  }
  return out.join('/');
}

function headerKey(value) {
  const text = String(value || '').normalize('NFKC').replace(/[\s:：()（）【】[\]]/g, '');
  if (/^(?:(?:基金|投资)?经理)?姓名$|^(?:基金经理|投资经理)$/.test(text)) return 'name';
  if (/^(?:(?:基金|投资)?经理|个人|人员)?(?:简介|介绍|简历)$/.test(text)) return 'biography';
  if (/其他|备注|标签/.test(text)) return 'note';
  if (/投资/.test(text) && /日期|时间/.test(text)) return 'investmentStart';
  if (/从业|工作/.test(text) && /日期|时间|起始|开始/.test(text)) return 'workStart';
  return '';
}

function textDate(value) {
  const match = String(value || '').trim().match(/^(\d{4})[年/.\-](\d{1,2})[月/.\-](\d{1,2})/);
  if (!match) return '';
  const date = `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  const [year, month, day] = date.split('-').map(Number);
  const parsed = new Date(Date.UTC(year, month - 1, day));
  return parsed.getUTCFullYear() === year && parsed.getUTCMonth() === month - 1 && parsed.getUTCDate() === day ? date : '';
}

function excelDate(value, date1904) {
  if (!Number.isFinite(value) || value < 0) return '';
  const day = Math.floor(value);
  if (!date1904 && day === 60) return '';
  const offset = date1904 ? day : day - (day >= 60 ? 1 : 0);
  const date = new Date((date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)) + offset * 86400000);
  if (date.getUTCFullYear() < 1900 || date.getUTCFullYear() > 2200) return '';
  return date.toISOString().slice(0, 10);
}

function serialFromDate(iso) {
  const [year, month, day] = iso.split('-').map(Number);
  let days = Math.round((Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 31)) / 86400000);
  if (days >= 60) days += 1;
  return days;
}

function cellText(cell, shared) {
  const type = cell.getAttribute('t');
  const raw = children(cell, NS.ss, 'v')[0]?.textContent ?? '';
  if (type === 's') return shared[Number(raw)] ?? '';
  if (type === 'inlineStr') return elements(cell, NS.ss, 't').map(node => node.textContent || '').join('');
  if (raw !== '' && Number.isFinite(Number(raw))) return { text: raw, number: Number(raw) };
  return raw;
}

export async function parseManagerWorkbook(bytes, name = '经理介绍.xlsx') {
  if (!(bytes instanceof Uint8Array) && !(bytes instanceof ArrayBuffer)) throw new Error('经理介绍 Excel 内容无效。');
  const source = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (!source.byteLength || source.byteLength > MAX_BYTES) throw new Error('经理介绍 Excel 为空或超过 30 MB。');
  const zip = await globalThis.JSZip.loadAsync(source);
  const read = async path => {
    const file = zip.file(path);
    if (!file) return null;
    return parseXml(await file.async('string'));
  };
  const workbook = await read('xl/workbook.xml');
  const rels = await read('xl/_rels/workbook.xml.rels');
  if (!workbook || !rels) throw new Error('这份 Excel 不是有效的 xlsx。');
  const date1904 = ['1', 'true'].includes(children(workbook, NS.ss, 'workbookPr')[0]?.getAttribute('date1904'));
  const paths = new Map();
  for (const rel of elements(rels, NS.pr, 'Relationship')) {
    if (rel.getAttribute('TargetMode') === 'External') continue;
    paths.set(rel.getAttribute('Id'), targetPath('xl/workbook.xml', rel.getAttribute('Target')));
  }
  const strings = await read('xl/sharedStrings.xml');
  const shared = strings ? elements(strings, NS.ss, 'si').map(item => elements(item, NS.ss, 't').map(node => node.textContent || '').join('')) : [];
  const warnings = [];
  const records = [];
  for (const [sheetIndex, sheet] of elements(workbook, NS.ss, 'sheet').entries()) {
    const sheetName = sheet.getAttribute('name') || `工作表${sheetIndex + 1}`;
    const path = paths.get(sheet.getAttributeNS(NS.rel, 'id') || sheet.getAttribute('r:id'));
    const doc = path ? await read(path) : null;
    if (!doc) continue;
    const rows = [];
    for (const row of elements(doc, NS.ss, 'row')) {
      const rowIndex = Number(row.getAttribute('r')) - 1;
      if (rowIndex < 0 || rowIndex > 20000) continue;
      const values = rows[rowIndex] ||= [];
      for (const cell of children(row, NS.ss, 'c')) {
        const index = columnIndex(cell.getAttribute('r'));
        if (index < 0 || index > 40) continue;
        const value = cellText(cell, shared);
        values[index] = typeof value === 'object' ? value : { text: String(value || ''), number: null };
      }
    }
    let headerRow = -1;
    let header = null;
    for (let index = 0; index < Math.min(rows.length, 30); index += 1) {
      const map = new Map();
      for (let column = 0; column < 12; column += 1) {
        const key = headerKey(rows[index]?.[column]?.text);
        if (key && !map.has(key)) map.set(key, column);
      }
      if (map.has('name') && map.has('biography')) { headerRow = index; header = map; break; }
    }
    if (!header) continue;
    for (let index = headerRow + 1; index < rows.length; index += 1) {
      if (!rows[index]) continue;
      const readCell = key => rows[index][header.get(key)];
      const person = String(readCell('name')?.text || '').trim();
      const biography = String(readCell('biography')?.text || '').replace(/\r\n?/g, '\n');
      if (!person || !biography.trim()) continue;
      const dateOf = key => {
        const cell = readCell(key);
        if (!cell?.text?.trim()) return '';
        const value = typeof cell.number === 'number' ? excelDate(cell.number, date1904) : textDate(cell.text);
        if (!value) warnings.push(`${sheetName}第${index + 1}行的${key === 'workStart' ? '从业日期' : '投资起始日期'}无法识别，这条年限保留原文。`);
        return value;
      };
      records.push({
        id: `${sheetIndex + 1}-${index + 1}`,
        name: person,
        biography,
        note: String(readCell('note')?.text || '').trim(),
        workStart: header.has('workStart') ? dateOf('workStart') : '',
        investmentStart: header.has('investmentStart') ? dateOf('investmentStart') : '',
        sheet: sheetName,
        row: index + 1,
      });
    }
  }
  if (!records.length) throw new Error('Excel 里没有读到“姓名”和“简介”。请使用基金经理简介这张表。');
  return { name, records, warnings };
}

function escapeXml(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** A maintenance workbook: biography text is refreshed, and the year columns keep TODAY() formulas. */
export function buildManagerWorkbook(records) {
  const header = ['姓名', '简介', '其他', '从业日期', '从业年限（年）', '投资起始日期', '投资年限（年）'];
  const rows = [header.map(value => `<c t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`).map((cell, index) => cell.replace('<c ', `<c r="${columnName(index)}1" `)).join('')];
  records.forEach((record, index) => {
    const row = index + 2;
    const cells = [
      inline(0, row, record.name),
      inline(1, row, record.biography),
      inline(2, row, record.note || ''),
      record.workStart ? dateCell(3, row, record.workStart) : '',
      record.workStart ? `<c r="E${row}"><f>ROUNDDOWN((TODAY()-D${row})/365,0)</f></c>` : '',
      record.investmentStart ? dateCell(5, row, record.investmentStart) : '',
      record.investmentStart ? `<c r="G${row}"><f>ROUNDDOWN((TODAY()-F${row})/365,${record.investmentDigits || 0})</f></c>` : '',
    ].filter(Boolean);
    rows.push(`<row r="${row}">${cells.join('')}</row>`);
  });
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS.ss}"><sheetData><row r="1">${rows[0]}</row>${rows.slice(1).join('')}</sheetData></worksheet>`;
  const files = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS.ss}" xmlns:r="${NS.rel}"><sheets><sheet name="基金经理简介" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS.ss}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`,
    'xl/worksheets/sheet1.xml': sheet,
  };
  const zip = new globalThis.JSZip();
  for (const [path, content] of Object.entries(files)) zip.file(path, content);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

function columnName(index) {
  return 'ABCDEFG'[index];
}

function inline(column, row, value) {
  const space = /^\s|\s$|\n/.test(value) ? ' xml:space="preserve"' : '';
  return `<c r="${columnName(column)}${row}" t="inlineStr"><is><t${space}>${escapeXml(value)}</t></is></c>`;
}

function dateCell(column, row, iso) {
  return `<c r="${columnName(column)}${row}" s="1"><v>${serialFromDate(iso)}</v></c>`;
}
