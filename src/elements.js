const MANAGER = "上海国泰海通证券资产管理有限公司";
const RISK_WORD = { R1: "低风险", R2: "中低风险", R3: "中风险", R4: "中高风险", R5: "高风险" };
const CCB_TYPE = { C1: "谨慎型", C2: "稳健型", C3: "平衡型", C4: "进取型", C5: "激进型" };

export function extractCore(raw) {
  const text = clean(raw);
  const basic = sliceBasic(text) || text;
  const name = fieldName(basic, text);
  const type = fieldType(basic);
  const manager = fieldManager(text);
  const custodian = fieldCustodian(text);
  const scope = investmentBlock(basic, /主要投资方向|投资方向[：:]/, /投资比例/)
    || investmentBlock(basic, /(?:^|\n)\s*(?:[（(]?[0-9一二三四五六七八九十]+[）).、．]\s*)?投资范围[：:]?/, /投资比例/);
  const ratio = compact(investmentBlock(basic, /投资比例/, /产品风险等级|风险评级评价为/));
  const strategy = fieldStrategy(text);
  const duration = fieldDuration(basic);
  const opening = fieldOpening(basic);
  const lock = fieldLock(text);
  const fees = fieldFees(basic);
  const amounts = fieldAmounts(text);
  const performance = fieldPerformance(basic);
  const risk = fieldRisk(basic);
  const notes = [];
  if (!name) notes.push("未识别产品名称，请在要素表中手改。");
  if (!scope) notes.push("未识别投资范围，请对照合同「投资范围」或「主要投资方向」手补。");
  if (performance.mode === "announce") notes.push("合同只写业绩报酬计提基准、未给具体数字，需在下方手填。");
  if (risk.code && !basic.includes(risk.code) && /中低风险|低风险|中风险|中高风险|高风险/.test(basic)) notes.push("合同未写 R 代号，已按风险用语对应。");
  return { name, type, manager, custodian, scope, ratio, strategy, duration, opening, lock, ...fees, ...amounts, performance, risk, notes };
}

export function composeElements(core, manual) {
  const code = (manual.code || "").trim();
  const name = core.name ? (code ? `${core.name}（产品代码：${code}）` : core.name) : "";
  let opening = core.opening || "";
  if (core.lock?.days && !new RegExp(`(?:持有满|不得少于)\\s*${core.lock.days}\\s*(?:天|个自然日|个工作日)`).test(opening)) {
    opening = [opening, `特别的，每笔份额需持有满${core.lock.days}${core.lock.unit}。`].filter(Boolean).join("\n");
  }
  if (manual.ongoing && (manual.openDate || "").trim()) {
    opening = [opening, `本次开放日\n${manual.openDate.trim()}（若出现合同变更，以管理人公告为准）`].filter(Boolean).join("\n");
  }
  const participate = ratePair(core.participateRate, core.exitRate);
  const productFee = feeLine(core.managementFee, core.custodyFee);
  const first = firstLine(core.firstWan, core.extraWan, core.firstFeeNote);
  return [
    ["产品名称", name],
    ["管理人", core.manager || ""],
    ["托管人", core.custodian || ""],
    ["产品类型", core.type || ""],
    ["投资范围", core.scope || ""],
    ["投资比例", core.ratio || ""],
    ["投资策略", core.strategy || ""],
    ["存续期", core.duration || ""],
    ["开放期", opening],
    ["认购/参与费率，退出费率", participate],
    ["首参金额", first],
    ["产品费率", productFee],
    ["业绩报酬", performanceText(core, manual)],
    ["产品风险等级", riskText(core, manual.channel, manual.salesCode)],
  ];
}

function performanceText(core, manual) {
  const mode = manual.perfMode || "auto";
  const p = core.performance || { mode: "announce" };
  if (mode === "none" || (mode === "auto" && p.mode === "none")) return p.sentence || "本集合计划管理人不提取业绩报酬。";
  if (mode === "tier" || (mode === "auto" && p.mode === "tier")) {
    const b1 = num(manual.b1 || p.b1);
    const b2 = num(manual.b2 || p.b2);
    const r1 = num(manual.r1 || p.r1);
    const r2 = num(manual.r2 || p.r2);
    if (!b1 || !b2 || !r1 || !r2) return "业绩报酬为阶梯计提。合同未写明两档基准，请在下方填写两个计提基准和两个提取比例。业绩报酬计提基准以管理人公告为准。";
    return `当年化收益率≤${b1}%时，不提取；当${b1}%＜年化收益率≤${b2}%时，对年化收益率大于${b1}%且不高于${b2}%的收益部分提取${r1}%比例的业绩报酬；当年化收益率＞${b2}%时，对年化收益率高于${b2}%的收益部分提取${r2}%的业绩报酬，及对年化收益率大于${b1}%且不高于${b2}%的收益部分提取${r1}%的业绩报酬。业绩报酬计提日为计划分红登记日、投资者退出日和计划终止日。（业绩报酬计提标准以公告为准）\n业绩报酬计提基准不构成收益承诺，实际收益可能低于该基准。`;
  }
  const benchmark = num(manual.b1 || p.benchmark);
  const ratio = num(manual.r1 || p.ratio);
  if (benchmark && ratio) {
    return `本集合计划对年化收益超过${benchmark}%的部分提取${ratio}%的业绩报酬。业绩报酬计提日为计划分红权益登记日、投资者退出日和计划终止日。（业绩报酬计提基准以管理人公告为准）\n业绩报酬计提基准不构成收益承诺，实际收益可能低于该基准。`;
  }
  if (ratio) {
    return `本集合计划对投资者每笔份额的持有期年化收益率超过业绩报酬计提基准的部分提取${ratio}%的业绩报酬。\n业绩报酬计提基准以管理人公告为准。\n业绩报酬计提基准不构成收益承诺，实际收益可能低于该基准。`;
  }
  return "业绩报酬计提基准以管理人公告为准。";
}

function riskText(core, channel, salesCode) {
  const code = core.risk?.code || "";
  const word = core.risk?.word || (code ? RISK_WORD[code] : "");
  if (!code && !word) return "";
  const n = code ? Number(code.slice(1)) : 0;
  let sentence = "";
  if (channel === "ccb") {
    const from = n || 1;
    const list = [];
    for (let i = from; i <= 5; i++) list.push(`C${i}（${CCB_TYPE["C" + i]}）`);
    const head = code ? `${word}（${code}）` : word;
    sentence = `管理人对于本集合计划的风险等级评定为${head}，适合合格投资者中的专业投资者以及风险承受能力为${list.join("、")}的普通投资者（代销机构对于本集合计划的风险等级评定以代销机构发布为准）。`;
  } else if (!code) {
    sentence = `管理人对于本集合计划的风险等级评定为${word}。销售机构对本集合计划的风险等级评定以销售机构发布为准。`;
  } else {
    sentence = `管理人对于本集合计划的风险等级评定为${code}（${word}），适合向合格投资者中的专业投资者和风险承受能力评估结果为C${n}-C5的普通投资者销售，销售机构对本集合计划的风险等级评定以销售机构发布为准。`;
  }
  if (!salesCode || !RISK_WORD[salesCode]) return sentence;
  return `${sentence.replace(/。\s*$/, "")}该产品销售机构风险评级为${salesCode}（${RISK_WORD[salesCode]}）。`;
}

function clean(raw) {
  return String(raw || "")
    .replace(/\u0000/g, "")
    .replace(/\r\n?/g, "\n")
    .replace(/[\u0001-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n");
}

function sliceBasic(text) {
  const marks = ["集合计划的基本情况", "资产管理计划的基本情况"];
  let best = "";
  for (const mark of marks) {
    let from = 0;
    while (from < text.length) {
      const at = text.indexOf(mark, from);
      if (at < 0) break;
      const chunk = text.slice(at, at + 12000);
      if (/名称|类别|类型/.test(chunk.slice(0, 500)) && /存续|开放期|管理费率|风险/.test(chunk)) {
        const end = chunk.search(/\n(?:六、资产管理计划的募集|资产管理计划的募集|第\s*6\s*部分)/);
        const body = (end > 200 ? chunk.slice(0, end) : chunk).trim();
        if (body.length > best.length) best = body;
      }
      from = at + mark.length;
    }
  }
  return best;
}

function fieldName(basic, text) {
  const inline = basic.match(/名称[：:]\s*([^\n。]{4,80})/);
  if (inline) return inline[1].replace(/集合资产管理计划.*/, "集合资产管理计划").trim();
  const block = basic.match(/资产管理计划的名称\s*\n+\s*([^\n。]{4,80})/);
  if (block) return block[1].replace(/集合资产管理计划.*/, "集合资产管理计划").trim();
  const cover = text.match(/国泰[^\n]{0,40}集合资产管理计划/);
  return cover ? cover[0].replace(/\s+/g, "") : "";
}

function fieldType(basic) {
  const inline = basic.match(/类型[：:]\s*([^\n。]{4,40})/);
  if (inline) return inline[1].trim();
  const block = basic.match(/(?:类别|产品类型)\s*\n+\s*([^\n。]{4,40})/);
  return block ? block[1].trim() : "";
}

function fieldManager(text) {
  const found = text.match(/管理人[：:]\s*([^\n]{4,40})/);
  let name = found ? found[1].trim() : "";
  name = name.replace(/上海国泰君安证券资产管理有限公司/g, MANAGER);
  if (/国泰君安证券资产管理/.test(name) && !/国泰海通/.test(name)) name = MANAGER;
  return name || MANAGER;
}

function fieldCustodian(text) {
  const found = text.match(/(?:集合)?托管人[：:]\s*([^\n]{4,40})/);
  return found ? found[1].replace(/[。；].*$/, "").trim() : "";
}

const LABEL_LINE = /^(?:及比例|(?:[（(]?[0-9一二三四五六七八九十]+[）).、．]\s*)?(?:主要)?投资(?:方向|比例|范围)(?:及比例)?(?:（占[^）]*）)?[：:]?|[（(]?[0-9一二三四五六七八九十]+[）).、．]?)$/;

function investmentBlock(basic, anchorRe, stopRe) {
  const flags = anchorRe.flags.includes("g") ? anchorRe.flags : anchorRe.flags + "g";
  const finder = new RegExp(anchorRe.source, flags);
  for (const found of basic.matchAll(finder)) {
    const around = basic.slice(Math.max(0, found.index - 24), found.index + found[0].length + 8);
    if (/投资目标、主要投资方向、投资比例/.test(around)) continue;
    const from = Math.max(0, basic.lastIndexOf("\n", found.index) + 1);
    const stop = basic.slice(found.index + found[0].length).search(stopRe);
    let end = stop > 20 ? found.index + found[0].length + stop : found.index + 5000;
    const lineBreak = basic.lastIndexOf("\n", end);
    if (lineBreak > from) end = lineBreak;
    const body = basic.slice(from, end).split("\n").map((line) => line.trim()).filter((line) => line && !LABEL_LINE.test(line)).join("\n").trim();
    if (/投资于|固定收益|银行存款|现金|本计划/.test(body.slice(0, 180))) return body;
  }
  return "";
}

function fieldDuration(basic) {
  const m = basic.match(/存续期限(?:为)?\s*(?:自成立之日起)?\s*([0-9０-９一二三四五六七八九十]+\s*年(?:（可展期）|\(可展期\))?)/);
  if (!m) return "";
  const years = m[1].replace(/\s+/g, "");
  return years.startsWith("自") ? years : `自成立之日起${years.replace(/^自成立之日起/, "")}`;
}

function fieldLock(text) {
  const match = String(text || "").match(/(?:每笔份额需持有满|份额持有时间不得少于|持有满)\s*(\d+)\s*(个自然日|个工作日|天)/);
  if (!match) return null;
  return { days: match[1], unit: match[2] };
}

function fieldOpening(basic) {
  const numbered = basic.search(/2、开放期[：:]/);
  if (numbered >= 0) {
    let rest = basic.slice(numbered).replace(/^2、开放期[：:]/, "").trim();
    const end = rest.search(/\n(?:[三四五六七八九十]{1,3}、|•\s*资产管理计划的目标规模|资产管理计划的目标规模|资产管理计划的存续期限|•\s*资产管理计划的存续期限)/);
    if (end > 0) rest = rest.slice(0, end);
    return joinParagraphs(rest);
  }
  const mode = basic.search(/运作方式/);
  if (mode < 0) return "";
  let rest = basic.slice(mode).replace(/^[^\n]*运作方式[：:]?\s*/, "").trim();
  const end = rest.search(/\n[（(][五六七八九十][）)]|\n(?:资产管理计划的投资目标|投资目标、主要投资方向)/);
  if (end > 0) rest = rest.slice(0, end);
  return joinParagraphs(rest);
}

function fieldStrategy(text) {
  const method = text.search(/投资管理的方法和标准/);
  if (method >= 0) {
    let rest = text.slice(method).replace(/^投资管理的方法和标准\s*/, "");
    const end = rest.search(/\n\s*(?:2[、.）)]|[（(][四五六七八九十][）)]|[四五六七八九十]{1,3}、|业绩比较基准|投资限制|投资禁止)/);
    if (end > 20) rest = rest.slice(0, end);
    const lines = rest.split(/\n/).map((line) => line.trim()).filter(Boolean);
    if (lines.length > 1 && /^\d+[、.）)]\s*\S{1,16}$/.test(lines[0])) lines.shift();
    return compact(lines.join("\n"));
  }
  const basicAt = text.search(/投资策略[：:]/);
  if (basicAt < 0) return "";
  const line = text.slice(basicAt).match(/投资策略[：:]\s*([^\n]{12,})/);
  return line ? compact(line[1]) : "";
}

function joinParagraphs(rest) {
  const parts = rest.split(/\n+/).map((s) => s.trim()).filter((line) => line && !/^开放期[：:]?$/.test(line));
  const unique = [];
  for (const part of parts) {
    if (unique.at(-1) === part) continue;
    if (unique.length && (unique.at(-1) === part || unique.at(-1).endsWith(part))) continue;
    if (unique.length && /[，、：:]$/.test(unique.at(-1))) unique[unique.length - 1] += part;
    else unique.push(part);
  }
  return unique.join("\n");
}

function compact(text) {
  return String(text || "").replace(/\s*\n+\s*/g, "").replace(/[ \t]{2,}/g, " ").trim();
}

function fieldFees(basic) {
  const participate = participateTiers(basic) || bracketTiers(basic) || percentNear(basic, /(?:存续期)?参与费率|认购\/参与费率/);
  const exit = percentNear(basic, /退出费率/);
  const management = percentNear(basic, /年固定管理费率|管理费率/);
  const custody = percentNear(basic, /年托管费率|托管费率/);
  return { participateRate: participate, exitRate: exit, managementFee: management, custodyFee: custody };
}

function participateTiers(basic) {
  const at = basic.search(/存续期参与费率/);
  if (at < 0) return "";
  const window = basic.slice(at, at + 700);
  const parts = [];
  for (const line of window.split(/\n/).slice(0, 8)) {
    if (/退出费率|管理费率|托管费率|业绩报酬/.test(line)) break;
    const found = line.match(/参与金额(?:（含参与费）)?\s*(.+?)[，,]\s*参与费\s*(\d+(?:\.\d+)?)\s*%/);
    if (found) parts.push(`${found[1].trim()}${found[2]}%`);
  }
  return parts.length >= 2 ? parts.join("，") : "";
}

function bracketTiers(basic) {
  const lines = [];
  for (const found of basic.matchAll(/(认购费率|参与费率)[：:]\s*([^\n。]*参与金额[^\n。]*%)/g)) {
    const body = found[2].replace(/[；;]\s*$/, "").trim();
    const line = `${found[1]}：${body}。`;
    if (!lines.includes(line)) lines.push(line);
  }
  return lines.join("\n");
}

function percentNear(text, label) {
  const found = text.match(new RegExp("(?:" + label.source + ")[^\\n%\\d]{0,40}(\\d+(?:\\.\\d+)?)\\s*%"));
  return found ? stripZero(found[1]) : "";
}

function fieldAmounts(text) {
  const at = text.search(/最低认购金额和支付方式|最低认购金额|首次(?:参与|认购)(?:金额|最低金额)/);
  const slice = at >= 0 ? text.slice(at, at + 700) : text;
  const firstPatterns = [
    /(?:认购|参与)金额[^。\n]{0,24}?不得低于\s*\[?(\d+(?:\.\d+)?)\]?\s*万元/,
    /首次(?:参与|认购)(?:金额|最低金额)[^。\n]{0,40}?不低于\s*\[?(\d+(?:\.\d+)?)\]?\s*万元/,
  ];
  const extraPatterns = [
    /追加(?:参与|认购)的?最低金额为人民币\s*\[?(\d+(?:\.\d+)?)\]?\s*万元/,
    /追加(?:参与|认购)[^。\n]{0,24}?不低于\s*\[?(\d+(?:\.\d+)?)\]?\s*万元/,
  ];
  const firstHit = pickAmount(slice, text, firstPatterns);
  const first = firstHit.value;
  const extra = pickAmount(slice, text, extraPatterns).value;
  const around = firstHit.around;
  const firstFeeNote = first && /(?<!不)含(?:认购|参与)费/.test(around) ? "（含认购费）" : "";
  return { firstWan: first, extraWan: extra, firstFeeNote };
}

function pickAmount(slice, text, patterns) {
  const local = amountHit(slice, patterns);
  return local.value ? local : amountHit(text, patterns);
}

function amountHit(text, patterns) {
  for (const pattern of patterns) {
    const found = text.match(pattern);
    if (!found) continue;
    const around = text.slice(found.index, found.index + 90);
    return { value: stripZero(found[1]), around };
  }
  return { value: "", around: "" };
}

function fieldPerformance(basic) {
  const waived = basic.match(/(本集合计划|本计划)不收取业绩报酬[^。\n]*/);
  if (waived) return { mode: "none", sentence: `${waived[0].replace(/[。.]?$/, "")}。` };
  if (/不提取业绩报酬/.test(basic)) return { mode: "none" };
  if (/R1/.test(basic) && /R2/.test(basic) && /5%/.test(basic) && /50%/.test(basic)) {
    return { mode: "tier", r1: "5", r2: "50" };
  }
  const stated = basic.match(/计提基准\s*(\d+(?:\.\d+)?)\s*%[^。\n]{0,24}?提取\s*(\d+(?:\.\d+)?)\s*%/);
  if (stated) return { mode: "single", benchmark: stripZero(stated[1]), ratio: stripZero(stated[2]) };
  const ratioOnly = basic.match(/超过业绩报酬计提基准的部分提取\s*(\d+(?:\.\d+)?)\s*%/);
  if (ratioOnly) return { mode: "announce", ratio: stripZero(ratioOnly[1]) };
  const loose = basic.match(/提取\s*(\d+(?:\.\d+)?)\s*%[^。\n]{0,12}业绩报酬/);
  if (loose) return { mode: "announce", ratio: stripZero(loose[1]) };
  return { mode: "announce" };
}

function fieldRisk(basic) {
  let code = "";
  let word = "";
  const pair = basic.match(/(中低风险|中高风险|中风险|低风险|高风险)\s*[（(]\s*(R[1-5])\s*[）)]|(R[1-5])\s*[（(]\s*(中低风险|中高风险|中风险|低风险|高风险)\s*[）)]/);
  if (pair) {
    code = pair[2] || pair[3];
    word = pair[1] || pair[4];
  } else {
    const onlyCode = basic.match(/风险评级评价为[^。\n]{0,12}(R[1-5])/);
    const onlyWord = basic.match(/风险评级评价为(中低风险|中高风险|中风险|低风险|高风险)/);
    word = onlyWord ? onlyWord[1] : "";
    code = onlyCode ? onlyCode[1] : ({ 低风险: "R1", 中低风险: "R2", 中风险: "R3", 中高风险: "R4", 高风险: "R5" }[word] || "");
    if (code && !word) word = RISK_WORD[code];
  }
  return { code, word };
}

function ratePair(participate, exit) {
  if (!participate && !exit) return "";
  if (/(?:认购|参与)费率：/.test(participate)) {
    const exitText = !exit ? "" : exit === "0" ? "退出费率0%" : `退出费率${exit}%`;
    return [participate, exitText].filter(Boolean).join("\n");
  }
  if (participate === "0" && (!exit || exit === "0")) return "0%";
  if (!participate) return exit === "0" ? "退出费率0%" : `退出费率${exit}%`;
  const part = /%/.test(participate) ? participate : `${participate}%`;
  return `参与费率${part}；退出费率${exit || "0"}%`;
}

function feeLine(management, custody) {
  if (!management && !custody) return "";
  return `固定管理费率：${management || "—"}%/年；托管费率：${custody || "—"}%/年`;
}

function firstLine(first, extra, feeNote) {
  if (!first && !extra) return "";
  const head = first ? `不低于${first}万元${feeNote || ""}。` : "";
  const tail = extra ? `追加参与的最低金额为人民币${extra}万元。` : "";
  return `${head}${tail}`;
}

function num(value) {
  const text = String(value || "").trim().replace(/%/g, "");
  return /^\d+(?:\.\d+)?$/.test(text) ? text.replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "") : "";
}

function stripZero(value) {
  return String(value).replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "");
}

export { RISK_WORD, CCB_TYPE };
