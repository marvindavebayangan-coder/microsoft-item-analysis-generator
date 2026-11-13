import { CHECK, key, text, itemOrder } from "./domain.mjs?v=0.4.0";

export function quizUrl(value) {
  let url;
  try { url = new URL(text(value)); } catch (_) { throw new Error("Paste the Microsoft Forms quiz link from Collect responses."); }
  if (url.protocol !== "https:" || url.username || url.password || url.port ||
      !["forms.office.com", "forms.microsoft.com", "forms.cloud.microsoft"].includes(url.hostname) || url.href.length > 2048 ||
      !(url.searchParams.get("id") || /\/r\/[^/]+/i.test(url.pathname))) {
    throw new Error("Use the HTTPS Microsoft Forms quiz link from Collect responses.");
  }
  url.hash = ""; return url.href;
}

export function questionBlocks(headers) {
  const blocks = [];
  for (let c = 0; c < headers.length - 2; c++) {
    if (headers[c] && headers[c + 1] === "Points - " + headers[c] && headers[c + 2] === "Feedback - " + headers[c]) {
      blocks.push({ answer: c, points: c + 1, title: headers[c] }); c += 2;
    }
  }
  return blocks;
}

export async function formsDataset(source) {
  const values = source?.values;
  if (!Array.isArray(values) || !values.length || values.length > 1001) throw new Error("Choose a Microsoft Forms quiz results export with up to 1,000 responses.");
  const headers = values[0].map(value => text(value));
  if (!headers.length || headers.length > 400 || new Set(headers).size !== headers.length || headers.some(value => !value)) throw new Error("The results sheet needs unique, nonblank column headings.");
  const id = headers.indexOf("Id"), total = headers.indexOf("Total points"), blocks = questionBlocks(headers);
  if (id < 0 || total < 0 || !blocks.length) throw new Error("The file needs Forms response IDs, Total points, and per-question Points/Feedback columns.");
  const preferred = blocks.find(block => /^(?:name\d*|student['’]?s?\s*name|learner['’]?s?\s*name|pupil['’]?s?\s*name)$/i.test(block.title));
  const nameColumn = preferred?.answer ?? headers.indexOf("Name");
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(JSON.stringify(headers)));
  return { headers, rows: values.slice(1).filter(row => row.some(value => value !== "" && value != null)),
    blocks, id, total, nameColumn, sheet: source.sheet || "Sheet1", hash: [...new Uint8Array(hash)].map(n => n.toString(16).padStart(2, "0")).join("") };
}

export function formsMapping(data, nameColumn, questions) {
  if (!Number.isInteger(nameColumn) || nameColumn < 0 || nameColumn >= data.headers.length || [data.id, data.total].includes(nameColumn)) throw new Error("Choose the column containing the student's roster name.");
  const items = data.blocks.filter(block => block.answer !== nameColumn);
  if (items.length !== questions || questions < 1 || questions > 100) throw new Error(`The selected subject has ${questions} items, but this name-column selection leaves ${items.length} quiz items. Use a matching subject and exclude the name question.`);
  return { hash: data.hash, nameColumn, id: data.id, total: data.total, items: items.map((block, i) => ({ id: i + 1, answer: block.answer, points: block.points, title: block.title.slice(0, 100) })) };
}

export function planFormsImport(data, binding, names, headings, paperNames = []) {
  if (data.hash !== binding.mapping.hash) throw new Error("This file's column layout differs from the connected quiz. Check that it is the latest export of the same quiz; question changes require reviewing the connection.");
  const mapping = formsMapping(data, binding.mapping.nameColumn, headings.length);
  const destination = itemOrder(headings, headings.length), roster = new Map(), paper = new Set(paperNames.map(key));
  names.forEach((name, row) => { const k = key(name); if (!k) return; if (roster.has(k)) throw new Error("Keep roster names unique before importing."); roster.set(k, row); });
  const seenIds = new Map(), seenStudents = new Map(), updates = [], unmatched = [], paperSkipped = [], pending = [];
  const known = Object.assign(Object.create(null), binding.responses || {});
  for (const row of data.rows) {
    const id = text(row[mapping.id]), name = text(row[mapping.nameColumn]), k = key(name);
    if (!/^[A-Za-z0-9_-]{1,128}$/.test(id) || !name) throw new Error("Every response needs a response ID and a student name. Check the selected name column.");
    const points = mapping.items.map(item => row[item.points]), signature = JSON.stringify([k, points, row[mapping.total]]);
    if (seenIds.has(id)) { if (seenIds.get(id) !== signature) throw new Error(`Response ID ${id} appears with conflicting results.`); continue; }
    seenIds.set(id, signature);
    if (!roster.has(k)) { unmatched.push(name); continue; }
    if (paper.has(k)) { paperSkipped.push(name); continue; }
    if (seenStudents.has(k) && seenStudents.get(k) !== id || Object.entries(known).some(([old, student]) => student === k && old !== id)) throw new Error(`More than one attempt was found for ${name}. Review the repeated attempts before importing; no marks were changed.`);
    if (Object.hasOwn(known, id) && known[id] !== k) throw new Error(`Response ID ${id} now names a different student. Review the name mapping before importing.`);
    seenStudents.set(k, id);
    if (points.some(value => value === "" || value == null)) { pending.push(name); continue; }
    if (points.some(value => value !== 0 && value !== 1)) throw new Error(`Use one-point, automatically graded items. ${name} has an item score other than 0 or 1.`);
    const sum = points.reduce((a, b) => a + b, 0);
    if (typeof row[mapping.total] !== "number" || row[mapping.total] !== sum) throw new Error(`The per-item points do not match Total points for ${name}. Check that the name question has zero points and the export is complete.`);
    const byItem = new Map(mapping.items.map((item, i) => [item.id, points[i] ? CHECK : ""]));
    updates.push({ row: roster.get(k), id, key: k, name, correct: sum, marks: destination.map(id => byItem.get(id)) });
    known[id] = k;
  }
  return { updates, responses: known, unmatched, paperSkipped, pending };
}
