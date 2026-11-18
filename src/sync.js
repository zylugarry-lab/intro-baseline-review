export const LIBRARY_BIN = 'https://extendsclass.com/api/json-storage/bin/feeceba';
const LIBRARY_KEYS = ['managers', 'general', 'ccb'];

export function parseSharedBody(text) {
  const raw = String(text || '').trim();
  const cut = raw.search(/<!doctype|<\/?html/i);
  const json = cut > 0 ? raw.slice(0, cut) : raw;
  const data = JSON.parse(json || '{}');
  const libraries = data?.libraries && typeof data.libraries === 'object' ? data.libraries : {};
  return { kind: 'intro-baseline-libraries', libraries };
}

export function withLibrary(shared, key, library) {
  if (!LIBRARY_KEYS.includes(key)) throw new Error('没有这一份底稿。');
  return {
    kind: 'intro-baseline-libraries',
    libraries: { ...(shared?.libraries || {}), [key]: library },
  };
}

export async function readShared() {
  const response = await fetch(`${LIBRARY_BIN}?t=${Date.now()}`, { cache: 'no-store' });
  if (!response.ok) throw new Error('读不到已同步的底稿。');
  return parseSharedBody(await response.text());
}

export async function writeShared(data) {
  const response = await fetch(LIBRARY_BIN, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  if (!response.ok) throw new Error('同步没有写上去。');
}
