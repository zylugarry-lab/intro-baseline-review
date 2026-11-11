import { companySections, selectVariant, COMPANY_LABELS } from './company.js';
import { buildManagerWorkbook, parseManagerWorkbook } from './excel.js';
import { downloadName, readDocxParagraphs, reviewOnePager } from './review.js';
import { readStored, removeStored, writeStored } from './store.js';
import { refreshBiography, shanghaiToday } from './years.js';

const defaults = {
  managers: { url: './assets/manager-introductions.xlsx', name: '投资经理介绍.xlsx' },
  general: { url: './assets/company-introduction-general.docx', name: '公司介绍-通用版.docx' },
  ccb: { url: './assets/company-introduction-ccb.docx', name: '公司介绍-建行版.docx' },
};
const DRAFT_KEY = 'draft';
const RECORD_FIELDS = ['id', 'name', 'biography', 'note', 'workStart', 'investmentStart'];

const state = {
  asOf: localStorage.getItem('intro-as-of') || shanghaiToday(),
  variant: localStorage.getItem('intro-variant') === 'ccb' ? 'ccb' : 'general',
  files: {},
  parsedRecords: [],
  records: [],
  warnings: [],
  company: { general: [], ccb: [] },
  companyParagraphs: { general: [], ccb: [] },
  companyBase: { general: {}, ccb: {} },
  openCard: null,
  onePagers: [],
};

const $ = id => document.getElementById(id);
let saveTimer = 0;

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
}

function saveBlob(bytes, name, type) {
  const blob = new Blob([bytes], { type });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 2000);
}

function cloneRecord(record) {
  return {
    id: record.id,
    name: record.name,
    biography: record.biography,
    note: record.note || '',
    workStart: record.workStart || '',
    investmentStart: record.investmentStart || '',
    sheet: record.sheet || '',
    row: record.row || 0,
  };
}

function normalizeRecord(record, index) {
  const date = value => (/^\d{4}-\d{2}-\d{2}$/.test(value || '') ? value : '');
  return {
    id: String(record?.id || `custom-${index}`),
    name: String(record?.name || ''),
    biography: String(record?.biography || '').replace(/\r\n?/g, '\n'),
    note: String(record?.note || ''),
    workStart: date(record?.workStart),
    investmentStart: date(record?.investmentStart),
    sheet: record?.sheet || '网页',
    row: Number(record?.row) || index + 2,
  };
}

function recordKey(record) {
  return RECORD_FIELDS.map(field => record[field] || '').join('\u0001');
}

function recordsDiffer() {
  if (state.records.length !== state.parsedRecords.length) return true;
  return state.records.some((record, index) => recordKey(record) !== recordKey(state.parsedRecords[index]));
}

function companyDiff(variant) {
  const diff = {};
  for (const paragraph of state.companyParagraphs[variant] || []) {
    if (paragraph.text !== state.companyBase[variant][paragraph.id]) diff[paragraph.id] = paragraph.text;
  }
  return diff;
}

function emptyDraft() {
  return { company: { general: {}, ccb: {} }, records: null };
}

async function loadDraft() {
  const stored = await readStored(DRAFT_KEY).catch(() => null);
  return {
    company: {
      general: stored?.company?.general || {},
      ccb: stored?.company?.ccb || {},
    },
    records: Array.isArray(stored?.records) ? stored.records.map(normalizeRecord) : null,
  };
}

async function loadBytes(key) {
  const response = await fetch(defaults[key].url);
  if (!response.ok) throw new Error(`读取内置${defaults[key].name}失败。`);
  return { name: defaults[key].name, bytes: new Uint8Array(await response.arrayBuffer()) };
}

function safeRefresh(record) {
  try {
    return refreshBiography(record, state.asOf);
  } catch {
    return { text: record.biography || '', calculations: [] };
  }
}

function meta() {
  $('managers-meta').textContent = `${recordsDiffer() ? '网页上修改' : '内置底稿'} · ${state.records.length} 人`;
  for (const variant of ['general', 'ccb']) {
    const count = Object.keys(companyDiff(variant)).length;
    $(`${variant}-meta`).textContent = count ? `网页上修改 · ${count} 段` : '内置底稿';
  }
}

function saveStatus(text) {
  $('save-state').textContent = text;
}

async function persistDraft() {
  const draft = {
    company: { general: companyDiff('general'), ccb: companyDiff('ccb') },
    records: recordsDiffer() ? state.records.map(cloneRecord) : null,
  };
  const untouched = !draft.records && !Object.keys(draft.company.general).length && !Object.keys(draft.company.ccb).length;
  if (untouched) await removeStored(DRAFT_KEY);
  else await writeStored(DRAFT_KEY, draft);
  meta();
  saveStatus(untouched ? '和内置底稿一致。' : '已保存在这台浏览器。');
}

function scheduleSave() {
  saveStatus('正在保存…');
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => { persistDraft().catch(error => { saveStatus(error.message || '保存失败。'); }); }, 250);
}

function tagCompany(variant) {
  let count = 0;
  for (const section of state.company[variant] || []) {
    for (const column of section.columns) {
      for (const entry of column.entries) entry.uid = `c${count += 1}`;
    }
    for (const note of section.notes) note.uid = `n${count += 1}`;
  }
}

function findTagged(variant, uid) {
  for (const section of state.company[variant] || []) {
    for (const column of section.columns) {
      const entry = column.entries.find(item => item.uid === uid);
      if (entry) return { entry };
    }
    const note = section.notes.find(item => item.uid === uid);
    if (note) return { note };
  }
  return null;
}

function atomsOf(variant) {
  const atoms = [];
  for (const section of state.company[variant] || []) {
    for (const column of section.columns) atoms.push(...column.entries);
    for (const note of section.notes) atoms.push(...note.entries);
  }
  return atoms;
}

function spliceClause(variant, entry, nextText) {
  const paragraph = state.companyParagraphs[variant].find(item => item.id === entry.paragraphId);
  if (!paragraph) return;
  const delta = nextText.length - entry.text.length;
  paragraph.text = paragraph.text.slice(0, entry.start) + nextText + paragraph.text.slice(entry.start + entry.text.length);
  for (const atom of atomsOf(variant)) {
    if (atom.paragraphId === entry.paragraphId && atom.start > entry.start) atom.start += delta;
  }
  entry.text = nextText;
}

async function reload() {
  $('baseline-error').hidden = true;
  try {
    const draft = await loadDraft();
    state.files.managers = await loadBytes('managers');
    state.files.general = await loadBytes('general');
    state.files.ccb = await loadBytes('ccb');
    const parsed = await parseManagerWorkbook(state.files.managers.bytes, state.files.managers.name);
    state.parsedRecords = parsed.records.map(cloneRecord);
    state.records = draft.records ? draft.records.map(cloneRecord) : state.parsedRecords.map(cloneRecord);
    state.warnings = parsed.warnings;
    for (const variant of ['general', 'ccb']) {
      const paragraphs = await readDocxParagraphs(state.files[variant].bytes);
      state.companyBase[variant] = Object.fromEntries(paragraphs.map(paragraph => [paragraph.id, paragraph.text]));
      for (const paragraph of paragraphs) {
        if (Object.prototype.hasOwnProperty.call(draft.company[variant], paragraph.id)) paragraph.text = draft.company[variant][paragraph.id];
      }
      state.companyParagraphs[variant] = paragraphs;
      state.company[variant] = companySections(selectVariant(paragraphs, variant));
    }
    meta();
    const edited = recordsDiffer() || Object.keys(companyDiff('general')).length || Object.keys(companyDiff('ccb')).length;
    saveStatus(edited ? '已载入这台浏览器里保存的修改。' : '和内置底稿一致。点一张卡片展开。');
    renderCompany('general');
    renderCompany('ccb');
    renderManagers();
    applyOpenCard();
  } catch (error) {
    $('baseline-error').hidden = false;
    $('baseline-error').textContent = error.message || '底稿读取失败。';
    saveStatus('底稿没有读出来。');
  }
}

function applyOpenCard() {
  for (const key of ['managers', 'general', 'ccb']) {
    const open = state.openCard === key;
    const detail = $(`detail-${key}`);
    detail.hidden = !open;
    $(`card-${key}`).classList.toggle('open', open);
    const chevron = $(`card-${key}`).querySelector('.chevron');
    if (chevron) chevron.textContent = open ? '收起' : '展开';
    if (open) detail.scrollIntoView({ block: 'nearest' });
  }
}

function renderCompany(variant) {
  const sections = state.company[variant] || [];
  tagCompany(variant);
  const hint = variant === 'ccb'
    ? '建行版股东介绍不含“全面领先”，也不写“资本实力位居行业最前列”。'
    : '通用版股东介绍保留“全面领先”和“资本实力位居行业最前列”。';
  const columns = sections.flatMap(section => section.columns).filter(column => column.role === 'company' || column.role === 'shareholder');
  const notes = sections.flatMap(section => section.notes);
  $(`detail-${variant}`).innerHTML = `
    <p class="lede">${escapeHtml(hint)}</p>
    <div class="company-columns">
      ${columns.map(column => `<div>${column.entries.map(entry => `
        <label class="clause"><b>${escapeHtml(COMPANY_LABELS[entry.key] || entry.key)}</b>
          <textarea data-uid="${entry.uid}" rows="${Math.min(6, Math.max(2, entry.text.split('\n').length + 1))}">${escapeHtml(entry.text)}</textarea>
        </label>`).join('')}</div>`).join('')}
    </div>
    ${notes.map(note => `<label class="note"><textarea data-uid="${note.uid}" rows="2" aria-label="数据说明">${escapeHtml(note.text)}</textarea></label>`).join('')}
  `;
}

function preparedRecords() {
  return state.records.map(record => ({ record, ...safeRefresh(record) }));
}

function tenureLabel(record) {
  const item = safeRefresh(record);
  const work = item.calculations.find(calc => calc.kind === 'work');
  const investment = item.calculations.find(calc => calc.kind === 'investment');
  return [work ? `从业 ${work.years}` : '', investment ? `投资 ${investment.years}` : ''].filter(Boolean).join(' · ');
}

function renderManagers() {
  const query = $('manager-search').value.trim();
  const rows = preparedRecords().filter(item => !query || item.record.name.includes(query));
  const warning = state.warnings.length ? `<p class="warn">${escapeHtml(state.warnings.slice(0, 4).join(' '))}</p>` : '';
  $('manager-preview').innerHTML = `${warning}<div class="manager-list">
    ${rows.map(item => {
      const record = item.record;
      const work = item.calculations.find(calc => calc.kind === 'work');
      const investment = item.calculations.find(calc => calc.kind === 'investment');
      return `<article class="manager-card" data-id="${escapeHtml(record.id)}">
        <div class="person-side">
          <div class="person-head">
            <div>
              <strong class="who">${escapeHtml(record.name || '未命名')}</strong>
              <span class="tenure">${escapeHtml(tenureLabel(record))}</span>
            </div>
            <button type="button" class="text-button remove-manager" data-remove-manager="${escapeHtml(record.id)}">删除</button>
          </div>
          <label>姓名<input class="cell-input name-input" data-field="name" value="${escapeHtml(record.name)}" aria-label="姓名"></label>
          <div class="date-pair">
            <label>从业日期<input class="cell-input" type="date" data-field="workStart" value="${escapeHtml(record.workStart || '')}" aria-label="从业日期"></label>
            <label>从业年限<strong data-year="work">${work ? escapeHtml(String(work.years)) : '—'}</strong></label>
          </div>
          <div class="date-pair">
            <label>投资起始<input class="cell-input" type="date" data-field="investmentStart" value="${escapeHtml(record.investmentStart || '')}" aria-label="投资起始日期"></label>
            <label>投资年限<strong data-year="investment">${investment ? escapeHtml(String(investment.years)) : '—'}</strong></label>
          </div>
        </div>
        <label class="wide">简介<textarea class="bio-input" data-field="biography" rows="6" aria-label="${escapeHtml(record.name || '经理')}简介">${escapeHtml(item.text)}</textarea></label>
      </article>`;
    }).join('') || '<p class="lede">没有匹配的经理。</p>'}
  </div>`;
}

function paintYears(card, record) {
  const item = safeRefresh(record);
  const work = item.calculations.find(calc => calc.kind === 'work');
  const investment = item.calculations.find(calc => calc.kind === 'investment');
  card.querySelector('[data-year="work"]').textContent = work ? String(work.years) : '—';
  card.querySelector('[data-year="investment"]').textContent = investment ? String(investment.years) : '—';
  card.querySelector('.who').textContent = record.name || '未命名';
  card.querySelector('.tenure').textContent = tenureLabel(record);
  return item;
}

async function buildPlainDocx(paragraphs) {
  const body = paragraphs.map(text => {
    const runs = String(text).split('\n').map((line, index) => `${index ? '<w:br/>' : ''}<w:t xml:space="preserve">${escapeHtml(line)}</w:t>`).join('');
    return `<w:p><w:r>${runs}</w:r></w:p>`;
  }).join('');
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/document.xml', document);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function downloadCompany(variant) {
  const title = variant === 'ccb' ? '建行版' : '通用版';
  const lines = [title];
  for (const section of state.company[variant] || []) {
    for (const column of section.columns) {
      if (column.role === 'shareholder') lines.push('股东');
      for (const entry of column.entries) lines.push(entry.text);
    }
    for (const note of section.notes) if (note.text.trim()) lines.push(note.text);
  }
  saveBlob(await buildPlainDocx(lines), `公司介绍-${title}.docx`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

function bindDrop(element, onFiles) {
  element.addEventListener('dragover', event => { event.preventDefault(); element.classList.add('dragging'); });
  element.addEventListener('dragleave', () => element.classList.remove('dragging'));
  element.addEventListener('drop', event => {
    event.preventDefault();
    element.classList.remove('dragging');
    onFiles([...event.dataTransfer.files]);
  });
}

function addOnePagers(files) {
  const accepted = [...files].filter(file => /\.docx$/i.test(file.name));
  state.onePagers.push(...accepted);
  renderOnePagers();
}

function renderOnePagers() {
  $('onepager-list').innerHTML = state.onePagers.map((file, index) => `
    <div class="file-entry"><span>${escapeHtml(file.name)}</span><button type="button" data-remove="${index}" class="text-button">移除</button></div>
  `).join('');
  $('run').disabled = state.onePagers.length === 0;
  $('compare-hint').textContent = state.onePagers.length
    ? `将用${state.variant === 'ccb' ? '建行版' : '通用版'}公司介绍，按 ${state.asOf} 的经理年限核对 ${state.onePagers.length} 份一页通。`
    : '选择文件后开始核对。';
}

function renderResults(results) {
  const changed = results.filter(item => item.changed).length;
  const managers = results.reduce((total, item) => total + item.managers.filter(manager => manager.status === 'updated').length, 0);
  $('results').innerHTML = `
    <div class="summary">
      <div><span>一页通</span><strong>${results.length}</strong></div>
      <div><span>有修改</span><strong>${changed}</strong></div>
      <div><span>经理简介已改</span><strong>${managers}</strong></div>
      <div><span>公司介绍版本</span><strong style="font-size:18px">${state.variant === 'ccb' ? '建行版' : '通用版'}</strong></div>
    </div>
    ${changed > 1 ? '<button id="download-zip" class="secondary" type="button">打包下载全部修订稿</button>' : ''}
    ${results.map((item, index) => `
      <article class="result">
        <h3>${escapeHtml(item.fileName)}</h3>
        <p>${item.changed ? `写了 ${item.changeCount} 处修订。` : '公司和经理介绍已经和底稿一致，原文件没有改。'}
          ${item.companyLabels.length ? `公司介绍改动：${escapeHtml(item.companyLabels.join('、'))}。` : ''}</p>
        ${item.managers.map(manager => `<p><span class="badge ${manager.status}">${manager.status === 'updated' ? '已修改' : manager.status === 'inserted' ? '已补入' : '一致'}</span>${escapeHtml(manager.name)} ${manager.calculations.map(calc => `${calc.label} ${calc.years}`).join(' · ')}</p>${manager.status === 'unchanged' ? '' : `<div class="diff"><del>${escapeHtml(manager.before || '（空）')}</del>\n<ins>${escapeHtml(manager.after)}</ins></div>`}`).join('')}
        ${item.warnings.map(warning => `<p class="warn">${escapeHtml(warning)}</p>`).join('')}
        ${item.changed ? `<button class="primary" type="button" data-download="${index}">下载修订稿</button>` : ''}
      </article>
    `).join('')}
  `;
  $('results').querySelectorAll('[data-download]').forEach(button => {
    button.addEventListener('click', () => {
      const item = results[Number(button.dataset.download)];
      saveBlob(item.bytes, downloadName(item.fileName), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    });
  });
  $('download-zip')?.addEventListener('click', async () => {
    const zip = new JSZip();
    results.filter(item => item.changed).forEach(item => zip.file(downloadName(item.fileName), item.bytes));
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    saveBlob(bytes, `介绍修订-${state.asOf}.zip`, 'application/zip');
  });
}

async function runReview() {
  $('run-error').hidden = true;
  $('run').disabled = true;
  $('compare-hint').textContent = '正在核对…';
  try {
    const paragraphs = state.companyParagraphs[state.variant] || [];
    const companyTextById = Object.fromEntries(paragraphs.map(paragraph => [paragraph.id, paragraph.text]));
    const results = [];
    for (const file of state.onePagers) {
      const bytes = new Uint8Array(await file.arrayBuffer());
      results.push(await reviewOnePager({
        bytes, fileName: file.name, companyBytes: state.files[state.variant].bytes, variant: state.variant,
        records: state.records, asOf: state.asOf, companyTextById,
      }));
    }
    renderResults(results);
    $('compare-hint').textContent = '核对完成。下载的 docx 用 Word 打开即可看到修订。';
  } catch (error) {
    $('run-error').hidden = false;
    $('run-error').textContent = error.message || '核对失败。';
    $('compare-hint').textContent = '核对没有完成。';
  } finally {
    $('run').disabled = state.onePagers.length === 0;
  }
}

function refreshVisibleBios() {
  for (const card of $('manager-preview').querySelectorAll('.manager-card[data-id]')) {
    const record = state.records.find(item => item.id === card.dataset.id);
    const area = card.querySelector('[data-field="biography"]');
    if (!record || !area || document.activeElement === area) continue;
    area.value = safeRefresh(record).text;
    paintYears(card, record);
  }
}

function onManagerEdit(target) {
  const card = target.closest('.manager-card');
  if (!card || !target.dataset.field) return;
  const record = state.records.find(item => item.id === card.dataset.id);
  if (!record) return;
  record[target.dataset.field] = target.value;
  paintYears(card, record);
  if (target.dataset.field !== 'biography') {
    const area = card.querySelector('[data-field="biography"]');
    if (area && document.activeElement !== area) area.value = safeRefresh(record).text;
  }
  scheduleSave();
}

async function resetCompany(variant) {
  const draft = await loadDraft();
  draft.company[variant] = {};
  await writeStored(DRAFT_KEY, { ...draft, records: recordsDiffer() ? state.records.map(cloneRecord) : null });
  for (const paragraph of state.companyParagraphs[variant]) paragraph.text = state.companyBase[variant][paragraph.id];
  state.company[variant] = companySections(selectVariant(state.companyParagraphs[variant], variant));
  renderCompany(variant);
  await persistDraft();
}

async function resetManagers() {
  state.records = state.parsedRecords.map(cloneRecord);
  renderManagers();
  await persistDraft();
}

function bind() {
  $('as-of').value = state.asOf;
  $('variant').value = state.variant;
  $('as-of').addEventListener('change', () => {
    if (!$('as-of').value) return;
    state.asOf = $('as-of').value;
    localStorage.setItem('intro-as-of', state.asOf);
    refreshVisibleBios();
    renderOnePagers();
  });
  $('use-today').addEventListener('click', () => {
    state.asOf = shanghaiToday();
    $('as-of').value = state.asOf;
    localStorage.setItem('intro-as-of', state.asOf);
    refreshVisibleBios();
    renderOnePagers();
  });
  $('variant').addEventListener('change', () => {
    state.variant = $('variant').value;
    localStorage.setItem('intro-variant', state.variant);
    renderOnePagers();
  });
  $('manager-search').addEventListener('input', renderManagers);
  document.querySelector('.cards').addEventListener('click', event => {
    const opener = event.target.closest('[data-open-panel]');
    if (!opener) return;
    const panel = opener.dataset.openPanel;
    state.openCard = state.openCard === panel ? null : panel;
    if (state.openCard === 'general' || state.openCard === 'ccb') {
      state.variant = state.openCard;
      $('variant').value = state.variant;
      localStorage.setItem('intro-variant', state.variant);
      renderOnePagers();
    }
    applyOpenCard();
  });
  $('add-manager').addEventListener('click', () => {
    const record = normalizeRecord({ id: `custom-${Date.now()}`, name: '', biography: '', sheet: '网页' }, 0);
    state.records.unshift(record);
    $('manager-search').value = '';
    renderManagers();
    $('manager-preview').querySelector(`[data-id="${CSS.escape(record.id)}"] [data-field="name"]`)?.focus();
    scheduleSave();
  });
  $('baseline').addEventListener('input', event => {
    const panel = event.target.closest('[data-company-variant]');
    const uid = event.target?.dataset?.uid;
    if (!panel || !uid) return;
    const found = findTagged(panel.dataset.companyVariant, uid);
    if (!found) return;
    const variant = panel.dataset.companyVariant;
    if (found.entry) spliceClause(variant, found.entry, event.target.value);
    else {
      const paragraph = state.companyParagraphs[variant].find(item => item.id === found.note.paragraphId);
      if (paragraph) paragraph.text = event.target.value;
      found.note.text = event.target.value;
    }
    scheduleSave();
  });
  $('manager-preview').addEventListener('input', event => onManagerEdit(event.target));
  $('manager-preview').addEventListener('change', event => {
    if (event.target?.dataset?.field === 'workStart' || event.target?.dataset?.field === 'investmentStart') onManagerEdit(event.target);
  });
  $('manager-preview').addEventListener('click', event => {
    const remove = event.target.closest('[data-remove-manager]');
    if (!remove) return;
    state.records = state.records.filter(record => record.id !== remove.dataset.removeManager);
    renderManagers();
    scheduleSave();
  });
  $('reset-managers').addEventListener('click', resetManagers);
  $('reset-general').addEventListener('click', () => resetCompany('general'));
  $('reset-ccb').addEventListener('click', () => resetCompany('ccb'));
  $('download-managers').addEventListener('click', async () => {
    const rows = preparedRecords().map(item => ({
      ...item.record,
      biography: item.text,
      investmentDigits: item.calculations.some(calc => calc.kind === 'investment' && !Number.isInteger(calc.years)) ? 1 : 0,
    }));
    saveBlob(await buildManagerWorkbook(rows), `经理介绍-年限截至${state.asOf}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  });
  $('download-general').addEventListener('click', () => downloadCompany('general'));
  $('download-ccb').addEventListener('click', () => downloadCompany('ccb'));
  $('onepager-file').addEventListener('change', event => { addOnePagers(event.target.files); event.target.value = ''; });
  $('onepager-list').addEventListener('click', event => {
    const index = event.target?.dataset?.remove;
    if (index == null) return;
    state.onePagers.splice(Number(index), 1);
    renderOnePagers();
  });
  $('run').addEventListener('click', runReview);
  bindDrop($('onepager-drop'), addOnePagers);
}

bind();
reload();
