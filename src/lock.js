const LOCK_PATTERN = /(?:每笔份额需持有满|份额持有时间不得少于|持有满)\s*(\d+)\s*(?:个自然日|个工作日|天)/;

export function parseLockDays(text) {
  const match = String(text || '').match(LOCK_PATTERN);
  return match ? match[1] : '';
}

export function lockDaysFromElements(rows) {
  const opening = (rows || []).find(row => row.label === '开放期')?.value || '';
  const fromOpening = parseLockDays(opening);
  if (fromOpening) return fromOpening;
  for (const row of rows || []) {
    const found = parseLockDays(row.value || '');
    if (found) return found;
  }
  return '';
}

export function planLockEdits(paragraphs, rows, skipIds = null) {
  const days = lockDaysFromElements(rows);
  const relevant = (paragraphs || []).filter(paragraph => paragraph.part === 'word/document.xml' && mentionsLock(paragraph.text));
  if (!relevant.length) return { edits: [], warnings: [] };
  if (!days) {
    return { edits: [], warnings: ['一页通写了锁定期天数，要素表的开放期里没有对应天数，这句没有改。'] };
  }
  const edits = [];
  for (const paragraph of relevant) {
    if (skipIds?.has(paragraph.id)) continue;
    const pattern = new RegExp(LOCK_PATTERN.source, 'g');
    let match;
    while ((match = pattern.exec(paragraph.text))) {
      if (match[1] === days) continue;
      const start = match.index + match[0].lastIndexOf(match[1]);
      edits.push({
        paragraphId: paragraph.id,
        start,
        before: match[1],
        after: days,
        label: '锁定期',
        source: 'lock',
      });
    }
  }
  return { edits, warnings: [] };
}

function mentionsLock(text) {
  return /锁定期|每笔份额/.test(text) && LOCK_PATTERN.test(text);
}
