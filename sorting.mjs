import { itemOrder, key, text } from "./domain.mjs?v=0.3.1";

export const identityOrder = order => order.every((source, destination) => source === destination);
export const reorderRows = (rows, order) => order.map(source => rows[source].slice());
export const reorderColumns = (rows, order) => rows.map(row => order.map(source => row[source]));

export function participantOrder(names, ids, ranks) {
  if (names.length !== ids.length || names.length !== ranks.length) throw new Error("The participant table is incomplete.");
  const count = names.filter(name => text(name)).length;
  const seen = new Set();
  for (const name of names.filter(value => text(value))) {
    if (seen.has(key(name))) throw new Error("Keep participant names unique before sorting.");
    seen.add(key(name));
  }
  if (ids.some(id => typeof id !== "number" || !Number.isSafeInteger(id) || id < 1) || new Set(ids).size !== ids.length) {
    throw new Error("Keep a unique original Student # for every participant row before sorting.");
  }
  const populated = ranks.filter((_, i) => text(names[i]));
  if (new Set(populated).size !== count || populated.some(rank => typeof rank !== "number" || !Number.isSafeInteger(rank) || rank < 1 || rank > count)) {
    throw new Error("Each named participant needs one valid Analysis Rank. Keep the score, rank and group formulas in place.");
  }
  if (ranks.some((rank, i) => !text(names[i]) && rank !== "" && rank != null)) {
    throw new Error("Blank participant rows must have blank ranks before sorting.");
  }
  return names.map((_, i) => i).sort((a, b) =>
    (text(names[a]) ? 0 : 1) - (text(names[b]) ? 0 : 1) ||
    (text(names[a]) && text(names[b]) ? ranks[a] - ranks[b] : 0) || ids[a] - ids[b]);
}

export function questionOrder(headings, percentages, byPercent) {
  const items = itemOrder(headings, headings.length);
  if (byPercent && (percentages.length !== items.length || percentages.some(value =>
    typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 1))) {
    throw new Error("The % correct row needs a number between 0% and 100% for every item before sorting.");
  }
  return items.map((_, i) => i).sort((a, b) => (byPercent ? percentages[b] - percentages[a] : 0) || items[a] - items[b]);
}

// Range.formulas returns literal values too, so startsWith('=') alone cannot
// distinguish a formula from a name such as '=Student'. Use SpecialCells instead.
export function formulaMask(addresses, top, left, height, width) {
  const mask = Array.from({ length: height }, () => Array(width).fill(false));
  if (!addresses) return mask;
  const areas = [];
  let quoted = false, first = 0;
  for (let i = 0; i < addresses.length; i++) {
    if (addresses[i] === "'") {
      if (quoted && addresses[i + 1] === "'") i++;
      else quoted = !quoted;
    } else if (addresses[i] === "," && !quoted) { areas.push(addresses.slice(first, i)); first = i + 1; }
  }
  areas.push(addresses.slice(first));
  const parse = value => {
    const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(value.trim());
    if (!match) throw new Error("Could not read the formula cells in the selected table.");
    const col = [...match[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    return [Number(match[2]) - 1, col - 1];
  };
  for (const area of areas) {
    const suffix = area.slice(area.lastIndexOf("!") + 1);
    const [a, b = a] = suffix.split(":"), [r1, c1] = parse(a), [r2, c2] = parse(b);
    for (let r = Math.max(top, r1); r <= Math.min(top + height - 1, r2); r++) {
      for (let c = Math.max(left, c1); c <= Math.min(left + width - 1, c2); c++) mask[r - top][c - left] = true;
    }
  }
  return mask;
}

export function formulaRuns(matrix, mask) {
  const runs = [];
  for (let r = 0; r < matrix.length; r++) {
    for (let c = 0; c < matrix[r].length;) {
      if (!mask[r][c]) { c++; continue; }
      const first = c++;
      while (c < matrix[r].length && mask[r][c]) c++;
      runs.push({ row: r, col: first, formulas: [matrix[r].slice(first, c)] });
    }
  }
  return runs;
}

export function verifyLiteralCells(actual, expected, mask) {
  if (actual.length !== expected.length || actual.some((row, r) => row.length !== expected[r].length)) {
    throw new Error("The sorted table has an unexpected size.");
  }
  for (let r = 0; r < actual.length; r++) for (let c = 0; c < actual[r].length; c++) {
    if (!mask[r][c] && !Object.is(actual[r][c] ?? "", expected[r][c] ?? "")) {
      throw new Error("The sort could not verify that the names, original IDs and recorded answers stayed matched.");
    }
  }
}

// Excel may normalize function/reference case or redundant zero offsets. String
// literals keep their original contents, including spaces and letter case.
export function normalizedFormula(value) {
  return String(value ?? "").split(/("(?:[^"]|"")*")/g).map((part, i) => i % 2 ? part :
    part.replace(/\s/g, "").toUpperCase().replace(/([RC])\[[-+]?0\]/g, "$1")).join("");
}

export function verifyFormulaCells(actual, actualMask, expected, expectedMask) {
  for (let r = 0; r < expectedMask.length; r++) for (let c = 0; c < expectedMask[r].length; c++) {
    if (actualMask[r][c] !== expectedMask[r][c]) {
      throw new Error("A calculated cell changed between a formula and plain text. The sort was not accepted.");
    }
    if (expectedMask[r][c] && normalizedFormula(actual[r][c]) !== normalizedFormula(expected[r][c])) {
      throw new Error("A sorted formula did not keep its expected references. The sort was not accepted.");
    }
  }
}

// Recover only the generator's known label expressions from 0.3.0. An arbitrary
// literal beginning with '=' or a teacher's intentional label override is kept.
export function labelRepairMask(values, mask, a1, relative) {
  return values.map((row, r) => row.map((value, c) => {
    if (mask[r][c] || typeof value !== "string") return false;
    const candidate = normalizedFormula(value.replace(/^'(?==)/, ""));
    return candidate === normalizedFormula(a1[r][c]) || candidate === normalizedFormula(relative[r][c]);
  }));
}
