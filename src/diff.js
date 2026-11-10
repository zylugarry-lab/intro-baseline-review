/** Character edits that keep the unchanged words outside the revision marks. */
export function textDiff(before, after) {
  if (before === after) return [];
  const n = before.length;
  const m = after.length;
  if (n * m > 250000) return edgeDiff(before, after);
  const grid = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i -= 1) {
    for (let j = m - 1; j >= 0; j -= 1) {
      grid[i][j] = before[i] === after[j] ? grid[i + 1][j + 1] + 1 : Math.max(grid[i + 1][j], grid[i][j + 1]);
    }
  }
  const edits = [];
  let i = 0;
  let j = 0;
  let pending = null;
  const flush = () => {
    if (!pending) return;
    edits.push(pending);
    pending = null;
  };
  while (i < n || j < m) {
    if (i < n && j < m && before[i] === after[j]) {
      flush();
      i += 1;
      j += 1;
      continue;
    }
    pending ||= { start: i, before: '', after: '' };
    if (j < m && (i === n || grid[i][j + 1] >= grid[i + 1][j])) pending.after += after[j++];
    else pending.before += before[i++];
  }
  flush();
  return edits.filter(edit => edit.before || edit.after);
}

function edgeDiff(before, after) {
  let start = 0;
  while (start < before.length && start < after.length && before[start] === after[start]) start += 1;
  let end = before.length;
  let tail = after.length;
  while (end > start && tail > start && before[end - 1] === after[tail - 1]) {
    end -= 1;
    tail -= 1;
  }
  if (end === start && tail === start) return [];
  return [{ start, before: before.slice(start, end), after: after.slice(start, tail) }];
}

export function similarity(a, b) {
  const norm = value => String(value || '').normalize('NFKC').replace(/[\s•·●◆\uF0D8]/g, '');
  const left = norm(a);
  const right = norm(b);
  if (left === right) return 1;
  if (!left || !right) return 0;
  const grams = value => {
    const out = new Map();
    for (let i = 0; i < value.length - 1; i += 1) {
      const gram = value.slice(i, i + 2);
      out.set(gram, (out.get(gram) || 0) + 1);
    }
    return out;
  };
  const x = grams(left);
  const y = grams(right);
  let shared = 0;
  for (const [gram, count] of x) shared += Math.min(count, y.get(gram) || 0);
  return (2 * shared) / Math.max(1, left.length + right.length - 2);
}

/** Align biography lines. 0 keep pair, 1 drop a template line, 2 insert a baseline line. */
export function alignLines(before, after) {
  const n = before.length;
  const m = after.length;
  const cost = Array.from({ length: n + 1 }, () => new Float64Array(m + 1));
  const step = Array.from({ length: n + 1 }, () => new Uint8Array(m + 1));
  for (let i = 1; i <= n; i += 1) {
    cost[i][0] = i * 0.7;
    step[i][0] = 1;
  }
  for (let j = 1; j <= m; j += 1) {
    cost[0][j] = j * 0.7;
    step[0][j] = 2;
  }
  for (let i = 1; i <= n; i += 1) {
    for (let j = 1; j <= m; j += 1) {
      const options = [
        cost[i - 1][j - 1] + 1.3 * (1 - similarity(before[i - 1], after[j - 1])),
        cost[i - 1][j] + 0.7,
        cost[i][j - 1] + 0.7,
      ];
      const best = Math.min(...options);
      cost[i][j] = best;
      step[i][j] = options.indexOf(best);
    }
  }
  const pairs = [];
  let i = n;
  let j = m;
  while (i || j) {
    const choice = step[i][j];
    if (choice === 0) pairs.push({ before: before[i - 1], after: after[j - 1] });
    else if (choice === 1) pairs.push({ before: before[i - 1], after: null });
    else pairs.push({ before: null, after: after[j - 1] });
    if (choice === 0) { i -= 1; j -= 1; }
    else if (choice === 1) i -= 1;
    else j -= 1;
  }
  return pairs.reverse();
}
