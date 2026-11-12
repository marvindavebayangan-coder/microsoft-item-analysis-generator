export const MAX_QUESTIONS = 100;
export const MAX_STUDENTS = 100;
export const CHECK = "✓";
export const PAPER_RED = "#ff0000";

export const text = value => value == null ? "" : String(value).trim();
export const key = value => text(value).normalize("NFKC").replace(/\s+/g, " ").toLowerCase();
export const literal = value => /^[=+\-@]/.test(String(value)) ? "'" + value : value;
export const isCheck = value => value === 1 || /^(p|1)$/i.test(text(value));

export function column(number) {
  let result = "";
  for (let n = number; n > 0; n = Math.floor((n - 1) / 26)) {
    result = String.fromCharCode(65 + (n - 1) % 26) + result;
  }
  return result;
}

export function readInput(raw, entries) {
  const names = [[0, "School"], [1, "District"], [2, "Assessment"],
    [3, "Period"], [5, "Teacher"], [6, "Subject / Grade(s)"]];
  for (const [index, label] of names) {
    if (!text(raw[index])) throw new Error(`Enter ${label} in Generate!B${index + 5}.`);
  }
  const questions = Number(raw[7]);
  if (typeof raw[7] === "boolean" || !Number.isSafeInteger(questions) || questions < 1 || questions > MAX_QUESTIONS) {
    throw new Error("Enter a whole number of questions from 1 to 100 in Generate!B12.");
  }
  if (typeof raw[4] !== "number" || !Number.isFinite(raw[4]) || raw[4] <= 0 || raw[4] > 2958465) {
    throw new Error("Enter a valid Excel date in Generate!B9.");
  }
  const roster = entries.filter(entry => text(entry.name)).map((entry, index) => ({
    name: text(entry.name), order: index + 1,
    paper: [PAPER_RED, "red"].includes(String(entry.color || "").toLowerCase())
  }));
  if (!roster.length || roster.length > MAX_STUDENTS) {
    throw new Error("Enter between 1 and 100 student names below Student Names here.");
  }
  const seen = new Set();
  for (const student of roster) {
    if (seen.has(key(student.name))) throw new Error(`Duplicate student name: ${student.name}. Use unique roster names.`);
    if (student.paper && key(student.name) === "show all") throw new Error("Show all is reserved for the paper student selector. Use another display name for that student.");
    seen.add(key(student.name));
  }
  const subject = text(raw[6]);
  if (subject.length > 100) throw new Error("Use a subject title of 100 characters or fewer.");
  return { school: text(raw[0]), district: text(raw[1]), assessment: text(raw[2]), period: text(raw[3]),
    date: raw[4], teacher: text(raw[5]), subject, questions, roster };
}

export function uniqueTabName(title, used, suffix = "") {
  const safe = text(title).replace(/[\\/\[\]*?:]/g, " ").replace(/\s+/g, " ").replace(/^'+|'+$/g, "") || "Subject";
  const existing = new Set(used.map(key));
  for (let index = 0; index < 10000; index++) {
    const tail = (index ? ` (${index + 1})` : "") + suffix;
    const candidate = safe.slice(0, 31 - tail.length).trimEnd() + tail;
    if (!existing.has(key(candidate)) && key(candidate) !== "history") return candidate;
  }
  throw new Error("Could not create a unique worksheet name. Shorten the subject title.");
}

export function itemOrder(values, count) {
  const order = values.map(value => Number(text(value).replace(/^item\s+/i, "")));
  if (order.length !== count || new Set(order).size !== count ||
      order.some(n => !Number.isSafeInteger(n) || n < 1 || n > count)) {
    throw new Error(`Keep every item number from 1 through ${count} exactly once.`);
  }
  return order;
}

export function paperGrades(matrix, destinationItems) {
  const count = destinationItems.length;
  const header = matrix[0];
  const rows = matrix.slice(1);
  const numbers = itemOrder(rows.map(row => row[0]), count);
  const destination = itemOrder(destinationItems, count);
  const letters = rows.map((row, index) => {
    const answer = text(row[1]).toUpperCase();
    if (!/^[ABCD]$/.test(answer)) throw new Error(`Enter A, B, C, or D as the key for Item ${numbers[index]}.`);
    return answer;
  });
  const seen = new Set();
  const results = [];
  for (let c = 2; c < header.length; c++) {
    const name = text(header[c]);
    const normalized = key(name);
    if (name && seen.has(normalized)) throw new Error(`Duplicate paper student: ${name}.`);
    if (name) seen.add(normalized);
    const answers = rows.map(row => text(row[c]).toUpperCase());
    if (!answers.some(Boolean)) continue;
    if (!name) throw new Error("Enter a name above every paper answer column.");
    const correct = new Set();
    answers.forEach((answer, r) => {
      if (answer && !/^[ABCD]$/.test(answer)) throw new Error(`Use A, B, C, or D for ${name}, Item ${numbers[r]}.`);
      if (answer === letters[r]) correct.add(numbers[r]);
    });
    results.push({ name, key: normalized, marks: destination.map(n => correct.has(n) ? CHECK : "") });
  }
  return results;
}

export function todaySerial() {
  const parts = new Intl.DateTimeFormat("en-US", { timeZone: "Asia/Manila", year: "numeric", month: "numeric", day: "numeric" })
    .formatToParts(new Date());
  const value = type => Number(parts.find(part => part.type === type).value);
  return Math.round((Date.UTC(value("year"), value("month") - 1, value("day")) - Date.UTC(1899, 11, 30)) / 86400000);
}
