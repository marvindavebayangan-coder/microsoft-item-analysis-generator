// Read values from ordinary .xlsx results exports in the browser. No workbook
// content is executed, uploaded, or inserted as another worksheet.
const decoder = new TextDecoder("utf-8", { fatal: true });
const FILE_LIMIT = 10 * 1024 * 1024, PART_LIMIT = 16 * 1024 * 1024;

function crc32(bytes) {
  let crc = -1;
  for (const byte of bytes) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) crc = crc >>> 1 ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ -1) >>> 0;
}

export async function unzipXlsx(buffer) {
  const bytes = new Uint8Array(buffer), view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 22 || bytes.length > FILE_LIMIT) throw new Error("Choose an .xlsx results file up to 10 MB.");
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === bytes.length) { end = i; break; }
  }
  if (end < 0 || view.getUint16(end + 4, true) || view.getUint16(end + 6, true)) throw new Error("Choose a normal, unencrypted Excel .xlsx results file.");
  const count = view.getUint16(end + 10, true), start = view.getUint32(end + 16, true), size = view.getUint32(end + 12, true);
  if (count !== view.getUint16(end + 8, true) || count > 2048 || start + size > end) throw new Error("The Excel archive has an unsupported directory.");
  const entries = new Map(); let offset = start, total = 0;
  for (let i = 0; i < count; i++) {
    if (offset + 46 > start + size || view.getUint32(offset, true) !== 0x02014b50) throw new Error("The Excel archive directory is damaged.");
    const flags = view.getUint16(offset + 8, true), method = view.getUint16(offset + 10, true);
    const packed = view.getUint32(offset + 20, true), unpacked = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true), extra = view.getUint16(offset + 30, true), comment = view.getUint16(offset + 32, true);
    if (offset + 46 + nameLength + extra + comment > start + size || flags & 1 || ![0, 8].includes(method) || unpacked > PART_LIMIT) throw new Error("The Excel file uses unsupported compression or encryption.");
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (entries.has(name) || name.startsWith("/") || name.split("/").includes("..")) throw new Error("The Excel archive has invalid part names.");
    entries.set(name, { flags, method, packed, unpacked, crc: view.getUint32(offset + 16, true), local: view.getUint32(offset + 42, true) });
    total += unpacked; offset += 46 + nameLength + extra + comment;
  }
  if (offset !== start + size || total > 50 * 1024 * 1024) throw new Error("The expanded Excel file is too large for this import.");
  return { has: name => entries.has(name), async read(name) {
    const entry = entries.get(name);
    if (!entry) throw new Error("Missing Excel part: " + name);
    const p = entry.local;
    if (p + 30 > start || view.getUint32(p, true) !== 0x04034b50) throw new Error("The Excel file has a damaged part.");
    const n = view.getUint16(p + 26, true), e = view.getUint16(p + 28, true), begin = p + 30 + n + e;
    if (begin + entry.packed > start || decoder.decode(bytes.subarray(p + 30, p + 30 + n)) !== name || view.getUint16(p + 8, true) !== entry.method) throw new Error("The Excel part does not match its directory.");
    const compressed = bytes.subarray(begin, begin + entry.packed);
    let content;
    if (entry.method === 0) content = compressed;
    else {
      let stream;
      try { stream = new DecompressionStream("deflate-raw"); }
      catch (_) { throw new Error("This browser cannot read compressed Excel files. Open Excel Online in a current Chrome or Edge browser."); }
      const reader = new Blob([compressed]).stream().pipeThrough(stream).getReader(), chunks = []; let length = 0;
      try {
        while (true) { const { done, value } = await reader.read(); if (done) break;
          length += value.length;
          if (length > entry.unpacked || length > PART_LIMIT) { await reader.cancel(); throw new Error("An Excel part exceeds its declared size."); }
          chunks.push(value);
        }
      } finally { reader.releaseLock(); }
      content = new Uint8Array(length); let at = 0;
      for (const chunk of chunks) { content.set(chunk, at); at += chunk.length; }
    }
    if (content.length !== entry.unpacked || crc32(content) !== entry.crc) throw new Error("The Excel file failed its integrity check. Download a fresh copy from Forms.");
    return decoder.decode(content).replace(/^\uFEFF/, "");
  } };
}

function entities(value) {
  return value.replace(/&([^;\s]+);/g, (_, code) => {
    const standard = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
    if (Object.hasOwn(standard, code)) return standard[code];
    if (!/^#(?:x[0-9a-f]+|\d+)$/i.test(code)) throw new Error("Unsupported XML entity in the results file.");
    const n = code[1]?.toLowerCase() === "x" ? parseInt(code.slice(2), 16) : Number(code.slice(1));
    if (!Number.isInteger(n) || n < 1 || n > 0x10ffff || n >= 0xd800 && n <= 0xdfff) throw new Error("Invalid character in the results file.");
    return String.fromCodePoint(n);
  });
}
const excelText = value => value.replace(/_x([0-9a-f]{4})_/gi, (_, code) => String.fromCharCode(parseInt(code, 16)));

// A small non-validating XML reader for Open XML parts. DTDs/external entities are
// rejected. Tag nesting, attributes, and unparsed text are checked explicitly.
export function xmlEvents(xml, visit) {
  const tokens = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<\/[A-Za-z_][\w.:-]*\s*>|<[A-Za-z_][\w.:-]*(?:\s+[A-Za-z_][\w.:-]*\s*=\s*(?:"[^"]*"|'[^']*'))*\s*\/?>|[^<]+/g;
  const stack = []; let last = 0, roots = 0;
  for (const match of xml.matchAll(tokens)) {
    if (match.index !== last) throw new Error("The Excel XML is malformed or contains unsupported declarations.");
    const token = match[0]; last += token.length;
    if (token.startsWith("<!--") || token.startsWith("<?")) continue;
    if (token.startsWith("<![CDATA[")) { if (!stack.length) throw new Error("Invalid XML text."); visit.text?.(token.slice(9, -3)); continue; }
    if (token[0] !== "<") { if (!stack.length && token.trim()) throw new Error("Invalid text outside the XML root."); visit.text?.(entities(token)); continue; }
    const closing = token.startsWith("</"), name = /^<\/?([^\s/>]+)/.exec(token)[1], local = name.split(":").at(-1);
    if (closing) { if (stack.pop() !== name) throw new Error("The Excel XML tags are not balanced."); visit.close?.(local); continue; }
    if (!stack.length && ++roots > 1) throw new Error("The Excel XML has more than one root.");
    const attributes = Object.create(null);
    for (const a of token.matchAll(/\s+([A-Za-z_][\w.:-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
      if (Object.hasOwn(attributes, a[1])) throw new Error("Duplicate XML attribute.");
      attributes[a[1]] = entities(a[2] ?? a[3]);
    }
    visit.open?.(local, attributes);
    if (token.endsWith("/>")) visit.close?.(local); else stack.push(name);
  }
  if (last !== xml.length || stack.length || roots !== 1) throw new Error("The Excel XML is incomplete.");
}

function strings(xml) {
  const result = []; let current = null, text = false, phonetic = 0;
  xmlEvents(xml, { open(name) { if (name === "si") current = ""; if (name === "rPh") phonetic++; if (name === "t" && !phonetic) text = true; },
    text(value) { if (current != null && text) current += value; },
    close(name) { if (name === "t") text = false; if (name === "rPh") phonetic--; if (name === "si") { result.push(excelText(current)); current = null; } } });
  return result;
}

export function worksheetValues(xml, shared = []) {
  const values = [], seen = new Set(); let cell = null, tag = "", maxRow = 0, maxCol = 0;
  xmlEvents(xml, { open(name, attrs) {
    if (name === "c") {
      const match = /^([A-Z]+)([1-9]\d*)$/i.exec(attrs.r || "");
      if (!match || seen.has(attrs.r)) throw new Error("Invalid or duplicate cell address in the results file.");
      const row = Number(match[2]) - 1, col = [...match[1].toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      if (row > 1000 || col > 399) throw new Error("Use a results export with up to 1,000 responses and 400 columns.");
      seen.add(attrs.r); cell = { row, col, type: attrs.t, value: "", inline: "" };
      maxRow = Math.max(maxRow, row + 1); maxCol = Math.max(maxCol, col + 1);
    }
    if (cell && ["v", "t"].includes(name)) tag = name;
  }, text(value) { if (cell && tag === "v") cell.value += value; else if (cell && tag === "t") cell.inline += value; },
  close(name) {
    if (name === tag) tag = "";
    if (name !== "c" || !cell) return;
    let value = cell.value;
    if (cell.type === "s") { if (!/^\d+$/.test(value) || !Object.hasOwn(shared, Number(value))) throw new Error("Invalid Excel shared string."); value = shared[Number(value)]; }
    else if (cell.type === "inlineStr") value = excelText(cell.inline);
    else if (cell.type === "str") value = excelText(value);
    else if (cell.type === "b") value = value === "1";
    else if (!cell.type || cell.type === "n") { if (value !== "") { value = Number(value); if (!Number.isFinite(value)) throw new Error("Invalid number in the results file."); } }
    (values[cell.row] ||= [])[cell.col] = value; cell = null;
  } });
  return Array.from({ length: maxRow }, (_, r) => Array.from({ length: maxCol }, (_, c) => values[r]?.[c] ?? ""));
}

export async function readFormsXlsx(buffer) {
  const archive = await unzipXlsx(buffer), relationships = new Map(), sheets = [];
  xmlEvents(await archive.read("xl/_rels/workbook.xml.rels"), { open(name, a) {
    if (name === "Relationship" && a.TargetMode !== "External" && a.Type?.endsWith("/worksheet")) {
      const url = new URL(a.Target, "https://xlsx.invalid/xl/workbook.xml");
      if (url.origin === "https://xlsx.invalid") relationships.set(a.Id, decodeURIComponent(url.pathname.slice(1)));
    }
  } });
  xmlEvents(await archive.read("xl/workbook.xml"), { open(name, a) { if (name === "sheet") sheets.push({ name: a.name, path: relationships.get(a["r:id"]) }); } });
  if (sheets.length > 20) throw new Error("Choose the Forms results export rather than a workbook with many subject tabs.");
  const shared = archive.has("xl/sharedStrings.xml") ? strings(await archive.read("xl/sharedStrings.xml")) : [];
  const candidates = [];
  for (const sheet of sheets) {
    if (!sheet.path) continue;
    const values = worksheetValues(await archive.read(sheet.path), shared);
    if (values[0]?.includes("Id") && values[0]?.includes("Total points")) candidates.push({ sheet: sheet.name, values });
  }
  if (candidates.length !== 1) throw new Error("Choose an Excel export containing one Microsoft Forms quiz results sheet.");
  return candidates[0];
}
