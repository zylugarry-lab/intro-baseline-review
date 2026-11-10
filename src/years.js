const DATE = /^\d{4}-\d{2}-\d{2}$/;

export function assertDate(value) {
  if (typeof value !== 'string' || !DATE.test(value)) throw new Error('日期需要是完整的 YYYY-MM-DD。');
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error('日期无效，请检查年、月、日。');
  }
  return value;
}

/** Calendar day in China, so the tenure does not follow the laptop time zone. */
export function shanghaiToday(now = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(now);
  const pick = type => parts.find(part => part.type === type).value;
  return assertDate(`${pick('year')}-${pick('month')}-${pick('day')}`);
}

export function daySpan(start, end) {
  assertDate(start);
  assertDate(end);
  const [ys, ms, ds] = start.split('-').map(Number);
  const [ye, me, de] = end.split('-').map(Number);
  return Math.round((Date.UTC(ye, me - 1, de) - Date.UTC(ys, ms - 1, ds)) / 86400000);
}

/**
 * Same rule as the workbook: ROUNDDOWN((制作日期 - 起始日期) / 365, digits).
 * A negative span cannot be a tenure.
 */
export function tenureYears(start, end, digits = 0) {
  const days = daySpan(start, end);
  if (days < 0) return null;
  const factor = 10 ** digits;
  return Math.floor((days / 365) * factor) / factor;
}

export function formatYears(years) {
  if (!Number.isFinite(years)) return '';
  if (Number.isInteger(years)) return String(years);
  return String(Math.round(years * 10) / 10);
}

const SPAN = /(?:超过|超|近|约)?\d+(?:\.\d+)?\s*(?:余|多)?年(?:以上)?\s*[^\d\n，,。；;：:、]{0,26}?经[验历]/g;

function kindOf(description) {
  if (/从业|投研|研究投资/.test(description)) return 'work';
  if (/投资/.test(description) && /经[验历]$/.test(description)) return 'investment';
  if (/经验$/.test(description)) return 'work';
  return null;
}

/**
 * Replace the first career span and the first investment span.
 * “其中” marks a narrower subset. Approximate claims such as “10年以上投研”
 * stay untouched when a more specific span of the same kind comes first.
 */
export function refreshBiography(record, asOf) {
  const text = String(record?.biography ?? '').replace(/\r\n?/g, '\n');
  const date = assertDate(asOf);
  const spans = [];
  for (const match of text.matchAll(SPAN)) {
    const before = match[0];
    const number = before.match(/\d+(?:\.\d+)?/)?.[0] || '';
    if (!number || Number(number) >= 100) continue;
    const description = before.replace(/^(?:超过|超|近|约)?\d+(?:\.\d+)?\s*(?:余|多)?年(?:以上)?\s*/, '');
    const kind = kindOf(description);
    if (!kind) continue;
    const prefix = text.slice(Math.max(0, match.index - 8), match.index);
    spans.push({
      kind, start: match.index, before, description,
      subset: /其中\s*$/.test(prefix), decimal: number.includes('.'),
    });
  }
  const calculations = [];
  const edits = [];
  for (const kind of ['work', 'investment']) {
    const startDate = record?.[kind === 'work' ? 'workStart' : 'investmentStart'];
    if (!startDate) continue;
    const candidates = spans.filter(span => span.kind === kind);
    const span = candidates.find(item => !item.subset) || candidates[0];
    if (!span) continue;
    const days = daySpan(startDate, date);
    if (days < 0) continue;
    const digits = span.decimal || days < 365 ? 1 : 0;
    const years = tenureYears(startDate, date, digits);
    const shown = formatYears(years);
    const after = `${shown}年${span.description}`;
    calculations.push({
      kind, label: kind === 'work' ? '从业年限' : '投资年限',
      startDate, asOf: date, days, years, before: span.before, after,
    });
    if (after !== span.before) edits.push({ ...span, after });
  }
  let result = text;
  for (const edit of edits.sort((a, b) => b.start - a.start)) {
    result = result.slice(0, edit.start) + edit.after + result.slice(edit.start + edit.before.length);
  }
  return { text: result, calculations };
}
