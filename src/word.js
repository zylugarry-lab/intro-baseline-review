export function wordKind(bytes) {
  if (bytes.length >= 4 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) return 'doc';
  if (bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b) return 'docx';
  return '';
}

export function binaryDocText(bytes) {
  const text = new TextDecoder('utf-16le').decode(bytes).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '\n');
  if (!/基本情况|集合资产管理计划|资产管理计划/.test(text)) {
    throw new Error('这份旧版 DOC 没能读出正文。请在 Word 中另存为 DOCX 后再上传。');
  }
  return text;
}
