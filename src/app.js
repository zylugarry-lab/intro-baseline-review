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

const state = {
  asOf: localStorage.getItem('intro-as-of') || shanghaiToday(),
  variant: localStorage.getItem('intro-variant') === 'ccb' ? 'ccb' : 'general',
  files: {},
  records: [],
  warnings: [],
  company: { general: [], ccb: [] },
  onePagers: [],
};

const $ = id => document.getElementById(id);

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

async function loadBytes(key) {
  const stored = await readStored(key).catch(() => null);
  if (stored?.bytes) return { name: stored.name, bytes: new Uint8Array(stored.bytes), source: '本机保存' };
  const response = await fetch(defaults[key].url);
  if (!response.ok) throw new Error(`读取内置${defaults[key].name}失败。`);
  return { name: defaults[key].name, bytes: new Uint8Array(await response.arrayBuffer()), source: '内置底稿' };
}

function meta(key) {
  const file = state.files[key];
  $(`${key}-meta`).textContent = file ? `${file.source} · ${file.name}` : '尚未读取';
}

async function reload() {
  $('baseline-error').hidden = true;
  try {
    state.files.managers = await loadBytes('managers');
    state.files.general = await loadBytes('general');
    state.files.ccb = await loadBytes('ccb');
    const parsed = await parseManagerWorkbook(state.files.managers.bytes, state.files.managers.name);
    state.records = parsed.records;
    state.warnings = parsed.warnings;
    for (const variant of ['general', 'ccb']) {
      const paragraphs = await readDocxParagraphs(state.files[variant].bytes);
      state.company[variant] = companySections(selectVariant(paragraphs, variant));
    }
    meta('managers');
    meta('general');
    meta('ccb');
    renderCompany();
    renderManagers();
  } catch (error) {
    $('baseline-error').hidden = false;
    $('baseline-error').textContent = error.message || '底稿读取失败。';
  }
}

function renderCompany() {
  const sections = state.company[state.variant] || [];
  const other = state.variant === 'ccb' ? 'general' : 'ccb';
  const textOf = variant => (state.company[variant] || []).flatMap(section => section.columns.flatMap(column => column.entries.map(entry => entry.text))).join('\n');
  const current = textOf(state.variant);
  const alternate = textOf(other);
  const hint = state.variant === 'ccb'
    ? '建行版用于比对。股东介绍不含“全面领先”，也不写“资本实力位居行业最前列”。'
    : '通用版用于比对。股东介绍保留“全面领先”和“资本实力位居行业最前列”。';
  const columns = sections.flatMap(section => section.columns).filter(column => column.role === 'company' || column.role === 'shareholder');
  const notes = [...new Set(sections.flatMap(section => section.notes.map(note => note.text)).filter(Boolean))];
  $('company-preview').innerHTML = `
    <p class="lede">${escapeHtml(hint)}${current === alternate ? '' : ' 两个版本的正文不一致，切换后下面的条目会跟着变。'}</p>
    <div class="company-columns">
      ${columns.map(column => `<div>${column.entries.map(entry => `<p class="clause"><b>${escapeHtml(COMPANY_LABELS[entry.key] || entry.key)}</b>${escapeHtml(entry.text)}</p>`).join('')}</div>`).join('')}
    </div>
    ${notes.map(note => `<p class="note">${escapeHtml(note)}</p>`).join('')}
  `;
}

function preparedRecords() {
  return state.records.map(record => ({ record, ...refreshBiography(record, state.asOf) }));
}

function renderManagers() {
  const query = $('manager-search').value.trim();
  const rows = preparedRecords().filter(item => !query || item.record.name.includes(query));
  const warning = state.warnings.length ? `<p class="warn">${escapeHtml(state.warnings.slice(0, 4).join(' '))}</p>` : '';
  $('manager-preview').innerHTML = `${warning}<table><thead><tr><th>姓名</th><th>从业日期</th><th>从业年限</th><th>投资起始</th><th>投资年限</th><th>按日期刷新后的开头</th></tr></thead><tbody>
    ${rows.map(item => {
      const work = item.calculations.find(calc => calc.kind === 'work');
      const investment = item.calculations.find(calc => calc.kind === 'investment');
      const yearCell = calc => calc ? `<span class="years">${escapeHtml(String(calc.years))}</span>` : '—';
      return `<tr><td>${escapeHtml(item.record.name)}</td><td>${escapeHtml(item.record.workStart || '—')}</td><td>${yearCell(work)}</td><td>${escapeHtml(item.record.investmentStart || '—')}</td><td>${yearCell(investment)}</td><td class="bio">${escapeHtml(item.text.split('\n')[0])}</td></tr>`;
    }).join('')}
  </tbody></table>`;
}

async function replaceBaseline(key, file) {
  if (!file) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  await writeStored(key, { name: file.name, bytes: Array.from(bytes), savedAt: new Date().toISOString() });
  await reload();
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
    const company = state.files[state.variant];
    const results = [];
    for (const file of state.onePagers) {
      const bytes = new Uint8Array(await file.arrayBuffer());
      results.push(await reviewOnePager({
        bytes, fileName: file.name, companyBytes: company.bytes, variant: state.variant,
        records: state.records, asOf: state.asOf,
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

function bind() {
  $('as-of').value = state.asOf;
  $('variant').value = state.variant;
  $('as-of').addEventListener('change', () => {
    if (!$('as-of').value) return;
    state.asOf = $('as-of').value;
    localStorage.setItem('intro-as-of', state.asOf);
    renderManagers();
    renderOnePagers();
  });
  $('use-today').addEventListener('click', () => {
    state.asOf = shanghaiToday();
    $('as-of').value = state.asOf;
    localStorage.setItem('intro-as-of', state.asOf);
    renderManagers();
    renderOnePagers();
  });
  $('variant').addEventListener('change', () => {
    state.variant = $('variant').value;
    localStorage.setItem('intro-variant', state.variant);
    renderCompany();
    renderOnePagers();
  });
  $('manager-search').addEventListener('input', renderManagers);
  $('managers-file').addEventListener('change', event => replaceBaseline('managers', event.target.files[0]));
  $('general-file').addEventListener('change', event => replaceBaseline('general', event.target.files[0]));
  $('ccb-file').addEventListener('change', event => replaceBaseline('ccb', event.target.files[0]));
  for (const key of ['managers', 'general', 'ccb']) {
    $(`reset-${key}`).addEventListener('click', async () => { await removeStored(key); await reload(); });
    $(`download-${key}`).addEventListener('click', async () => {
      if (key === 'managers') {
        const rows = preparedRecords().map(item => ({
          ...item.record,
          biography: item.text,
          investmentDigits: item.calculations.some(calc => calc.kind === 'investment' && !Number.isInteger(calc.years)) ? 1 : 0,
        }));
        saveBlob(await buildManagerWorkbook(rows), `经理介绍-年限截至${state.asOf}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        return;
      }
      const file = state.files[key];
      saveBlob(file.bytes, file.name, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    });
  }
  $('onepager-file').addEventListener('change', event => { addOnePagers(event.target.files); event.target.value = ''; });
  $('onepager-list').addEventListener('click', event => {
    const index = event.target?.dataset?.remove;
    if (index == null) return;
    state.onePagers.splice(Number(index), 1);
    renderOnePagers();
  });
  $('run').addEventListener('click', runReview);
  bindDrop($('onepager-drop'), addOnePagers);
  bindDrop($('card-managers'), files => replaceBaseline('managers', files[0]));
  bindDrop($('card-general'), files => replaceBaseline('general', files[0]));
  bindDrop($('card-ccb'), files => replaceBaseline('ccb', files[0]));
}

bind();
reload();
