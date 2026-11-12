import { CHECK, column, key, literal, text, isCheck, readInput, uniqueTabName, itemOrder, paperGrades } from "./domain.mjs?v=0.2.1";
import { detectTemplate, subjectLayout, headerPlan, analysisFormulas, MIN_ITEM_WIDTH, GUIDE_LINES } from "./layout.mjs?v=0.2.1";

const REGISTRY = "TIA.Generator.Registry.v1";
const MAX_COLUMNS = 16384;
const GREEN = "#C6EFCE";
const BLUE = "#D9EAF7";
const PALE = "#FFF2CC";
let writing = false;
let eventsRegistered = false;

async function registry(context) {
  const setting = context.workbook.settings.getItemOrNullObject(REGISTRY);
  setting.load("value");
  await context.sync();
  if (setting.isNullObject) return { version: 1, subjects: [] };
  const state = JSON.parse(setting.value);
  if (!state || !Array.isArray(state.subjects)) throw new Error("The workbook's add-in settings are invalid. Keep this file and report the error details.");
  return state;
}

function saveRegistry(context, state) {
  context.workbook.settings.add(REGISTRY, JSON.stringify(state));
}

function baseFormat(range) {
  range.format.font.name = "Arial";
  range.format.font.size = 10;
  range.format.verticalAlignment = "Center";
  range.format.rowHeight = 16.5;
}

function coloredRule(range, formula, color, owned = false) {
  const rule = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
  rule.custom.rule.formula = owned ? '=AND(' + formula.replace(/^=/, "") + ',N("TIA_ANALYSIS")=0)' : formula;
  rule.custom.format.fill.color = color;
}

function hideTail(sheet, width) {
  sheet.getRangeByIndexes(0, 0, 1, width).getEntireColumn().columnHidden = false;
  if (width < MAX_COLUMNS) {
    sheet.getRangeByIndexes(0, width, 1, MAX_COLUMNS - width).getEntireColumn().columnHidden = true;
  }
}

async function readTemplate(context, sheet) {
  const used = sheet.getUsedRangeOrNullObject(true).load("rowIndex,rowCount,columnIndex,columnCount");
  await context.sync();
  if (used.isNullObject) throw new Error("Original is empty. Keep your formatted subject template in that tab.");
  const bottom = used.rowIndex + used.rowCount;
  if (bottom > 5000) throw new Error("Keep Original's template and notes within the first 5000 rows.");
  const labels = sheet.getRangeByIndexes(0, 0, bottom, 1).load("text");
  const headers = sheet.getRange("A6:DB9").load("text");
  await context.sync();
  const layout = detectTemplate(labels.text, headers.text);
  const head = sheet.getRangeByIndexes(0, 0, 9, Math.max(44, layout.totalCol + 1));
  const headerText = head.load("text");
  const styles = head.getCellProperties({ format: {
    font: { name: true, size: true, bold: true, italic: true, color: true },
    fill: { color: true }, horizontalAlignment: true, verticalAlignment: true,
    wrapText: true, textOrientation: true
  } });
  const widths = Array.from({ length: layout.totalCol + 1 }, (_, i) =>
    sheet.getCell(8, i).format.load("columnWidth"));
  const heights = Array.from({ length: 9 }, (_, i) => sheet.getCell(i, 0).format.load("rowHeight"));
  const studentRow = sheet.getRange("A10").format.load("rowHeight");
  await context.sync();
  return { ...layout, headerText: headerText.text, styles: styles.value,
    widths: widths.map(x => x.columnWidth), heights: heights.map(x => x.rowHeight),
    studentHeight: studentRow.rowHeight || 14.25, usedWidth: used.columnIndex + used.columnCount };
}

export async function prepareWorkbook() {
  return guarded(async () => Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/id");
    const state = await registry(context);
    const input = sheets.items.find(sheet => sheet.name === "Generate");
    const source = sheets.items.find(sheet => sheet.name === "Original");
    if (!input || !source) throw new Error("Open your template workbook with tabs named exactly Original and Generate. This version uses the tabs you created.");
    const template = await readTemplate(context, source);
    const data = await readGenerate(context, input);
    const guideUpdated = await upgradeLegacyGuide(context, sheets);
    state.setup = "template-2";
    saveRegistry(context, state);
    await context.sync();
    return "Template checked: " + template.capacity + " item slots and " + template.students +
      " student rows. Generate contains " + data.questions + " questions and " + data.roster.length +
      " students. Ready to generate " + data.subject + "." +
      (guideUpdated ? " Updated the supplied GUIDE's worked example and decision rules." : "");
  }));
}

async function upgradeLegacyGuide(context, sheets) {
  const guide = sheets.items.find(sheet => sheet.name === "GUIDE");
  if (!guide) return false;
  const example = guide.getRange("A1:M43").load("values");
  await context.sync();
  const v = example.values;
  // Only migrate the supplied legacy example. Leave other teachers' guides alone.
  if (key(v[1][0]).replace(/[^a-z]/g, "") !== "itemno" || key(v[1][7]) !== "decision" ||
      !/poor item.*reject/i.test(String(v[7][11])) || v.slice(1).some(row => text(row[8]))) return false;
  const counts = v.slice(3).flatMap(row => [Number(row[1]), Number(row[3])]);
  if (counts.some(x => !Number.isInteger(x) || x < 0 || x > 100)) return false;
  const n = Math.max(1, ...counts);
  guide.getRange("K1").values = [["Example N"]];
  guide.getRange("L1").values = [[n]];
  guide.getRange("L1").format.fill.color = PALE;
  guide.getRange("L1").dataValidation.rule = { wholeNumber: { operator: "Between", formula1: "=MAX(1,$B$4:$B$43,$D$4:$D$43)", formula2: 100 } };
  guide.getRange("B3").formulas = [['="F (N="&$L$1&")"']];
  guide.getRange("D3").formulas = [['="F (N="&$L$1&")"']];
  guide.getRange("H2").values = [["Description"]];
  guide.getRange("I2:I3").merge();
  guide.getRange("I2").values = [["Decision"]];
  guide.getRange("I:I").format.columnWidth = 75;
  const proportions = [], lower = [], results = [];
  for (let r = 4; r <= 43; r++) {
    proportions.push(['=IF($L$1=0,"",B' + r + '/$L$1)']);
    lower.push(['=IF($L$1=0,"",D' + r + '/$L$1)']);
    results.push(['=IF($L$1=0,"",(C' + r + '+E' + r + ')/2)',
      '=IF($L$1=0,"",ROUND(C' + r + '-E' + r + ',10))',
      '=IF(ISNUMBER(G' + r + '),IF(G' + r + '<0.2,"PI",IF(G' + r + '<0.3,"NI",IF(G' + r + '<0.4,"GI","VG"))),"")',
      '=IF(ISNUMBER(G' + r + '),IF(G' + r + '<0.2,"Review",IF(G' + r + '<0.3,"Revise","Retain")),"Review")']);
  }
  guide.getRange("C4:C43").formulas = proportions;
  guide.getRange("E4:E43").formulas = lower;
  guide.getRange("F4:I43").formulas = results;
  guide.getRange("F4:G43").numberFormat = Array.from({ length: 40 }, () => ["0.00", "0.00"]);
  guide.getRange("L8").values = [["PI = Poor Item: Review"]];
  guide.getRange("M13").values = [["Review: check the answer key, wording and learning objective."]];
  guide.getRange("M17").values = [["Upper correct responses / example group size in L1."]];
  guide.getRange("M19").values = [["Lower correct responses / example group size in L1."]];
  guide.getRange("M22").values = [["Description uses D. The next column gives the decision."]];
  guide.getRange("L23").values = [["Decision rules"]];
  guide.getRange("M23").values = [["Retain: D >= 0.30. Revise: 0.20 <= D < 0.30. Review: D < 0.20."]];
  guide.getRange("L24").values = [["Worked example"]];
  guide.getRange("M24").values = [["Example N is independent of the class roster. Each subject calculates its own groups."]];
  guide.getRange("I2:I43").format.verticalAlignment = "Center";
  guide.getRange("I2:I3").format.wrapText = true;
  guide.getRange("H2:H3").format.wrapText = true;
  await context.sync();
  return true;
}

async function readGenerate(context, sheet) {
  const raw = sheet.getRange("B5:B12").load("values");
  const used = sheet.getUsedRange(true).load("rowIndex,rowCount");
  await context.sync();
  const bottom = used.rowIndex + used.rowCount;
  if (bottom > 5000) throw new Error("Keep the input roster near the Student Names here heading; this prototype reads up to row 5000.");
  const names = sheet.getRangeByIndexes(0, 0, bottom, 1).load("text");
  await context.sync();
  const heading = names.text.findIndex(row => /^(student|pupil)names(?:here)?$/.test(key(row[0]).replace(/[^a-z0-9]/g, "")));
  if (heading < 0) throw new Error("Keep Student Names here in column A above the roster.");
  const entries = [];
  for (let r = heading + 1; r < names.text.length; r++) {
    if (text(names.text[r][0])) {
      const fill = sheet.getCell(r, 0).format.fill.load("color");
      entries.push({ name: names.text[r][0], fill });
    }
  }
  if (entries.length > 100) throw new Error("The roster contains more than 100 names.");
  await context.sync();
  return readInput(raw.values.map(row => row[0]), entries.map(entry => ({ name: entry.name, color: entry.fill.color })));
}

function sourceStyle(template, address) {
  const [, letters, row] = /^([A-Z]+)(\d+)$/.exec(address);
  const col = [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  return template.styles[Number(row) - 1]?.[col - 1]?.format || {};
}

function headerStyle(range, style, { field = false, title = false } = {}) {
  for (const [name, value] of Object.entries(style.font || {})) {
    if (value != null) range.format.font[name] = value;
  }
  if (!style.font?.name) range.format.font.name = "Calibri";
  if (!style.font?.size) range.format.font.size = 11;
  range.format.fill.color = style.fill?.color || (title ? GREEN : "#FFFFFF");
  range.format.textOrientation = 0;
  range.format.wrapText = true;
  range.format.verticalAlignment = "Center";
  range.format.horizontalAlignment = title ? "Center" : style.horizontalAlignment || "Left";
  for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]) {
    const border = range.format.borders.getItem(edge);
    border.style = "Continuous"; border.weight = "Thin"; border.color = "#000000";
  }
}

function renderHeader(sheet, data, template) {
  const plan = headerPlan(data.questions);
  const q = data.questions, itemWidth = Math.max(MIN_ITEM_WIDTH, template.widths[2]);
  const widthAt = col => col <= 2 ? template.widths[col - 1] :
    col <= q + 2 ? itemWidth : template.widths[template.capacity + col - q - 1];
  const height = template.heights.slice(0, 8);
  const top = sheet.getRangeByIndexes(0, 0, 8, plan.score);
  top.format.fill.color = "#FFFFFF";
  top.format.textOrientation = 0;
  for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]) {
    top.format.borders.getItem(edge).style = "None";
  }
  for (const block of plan.blocks) {
    const range = sheet.getRangeByIndexes(block.row - 1, block.col - 1, 1, block.width);
    if (block.width > 1) range.merge();
    const title = block.row === 2 && !block.field && q >= 6;
    const value = block.field ? block.field === "students" ? data.roster.length : data[block.field] :
      title ? template.headerText[1]?.[0] || block.value : block.value;
    sheet.getCell(block.row - 1, block.col - 1).values = [[literal(value)]];
    const style = sourceStyle(template, block.style);
    headerStyle(range, style, { field: !!block.field, title });
    if (block.field === "date") sheet.getCell(block.row - 1, block.col - 1).numberFormat = [["m/d/yyyy"]];
    const width = Array.from({ length: block.width }, (_, i) => widthAt(block.col + i)).reduce((a, b) => a + b, 0);
    const display = block.field === "date" ? "11/12/2026" : String(value);
    const font = style.font?.size || 11;
    const lines = Math.max(1, Math.ceil(display.length * font * 0.52 / Math.max(12, width - 5)));
    if (lines > 1) height[block.row - 1] = Math.max(height[block.row - 1], lines * font * 1.25 + 3);
  }
  if (q >= 6) {
    sheet.getRangeByIndexes(4, 0, 1, plan.score).format.fill.color = GREEN;
    sheet.getRange("A1:A2").getEntireRow().rowHidden = false;
    sheet.getRange("A5:A9").getEntireRow().rowHidden = false;
  } else {
    sheet.getRange("A1:A9").getEntireRow().rowHidden = false;
    for (let r = 0; r < height.length; r++) height[r] = Math.max(height[r], 18.75);
  }
  height.forEach((value, r) => { sheet.getCell(r, 0).format.rowHeight = Math.min(409, value || 15); });
  sheet.getRange("A9").values = [[template.headerText[8]?.[0] || "Student name"]];
  sheet.getRange("B9").values = [[template.headerText[8]?.[1] || "Student #"]];
  sheet.getRange("B9").format.textOrientation = 90;
  const items = sheet.getRangeByIndexes(8, 2, 1, q);
  items.values = [Array.from({ length: q }, (_, i) => i + 1)];
  items.format.textOrientation = 90;
  items.format.horizontalAlignment = "Center";
  items.format.verticalAlignment = "Bottom";
  sheet.getRangeByIndexes(8, 0, 1, plan.score).format.rowHeight =
    Math.max(template.heights[8], q < 6 ? 105 : 0);
  [q + 3, q + 4].forEach((col, i) => {
    const range = sheet.getRangeByIndexes(plan.scoreHeaderRow - 1, col - 1, 10 - plan.scoreHeaderRow, 1);
    if (plan.scoreHeaderRow < 9) range.merge();
    const oldAddress = column(template.totalCol + i) + (template.headerText[5]?.[template.totalCol - 1 + i] ? "6" : "9");
    headerStyle(range, sourceStyle(template, oldAddress));
    sheet.getCell(plan.scoreHeaderRow - 1, col - 1).values = [[i ? "Student Score" : "# Of Correct Answers"]];
    range.format.textOrientation = 90;
    range.format.horizontalAlignment = "Center";
  });
  sheet.getRange(plan.cells.students).formulas = [[ '=COUNTIF($A$10:$A$' + (9 + data.roster.length) + ',"<>")' ]];
  return plan.cells;
}

async function clearCopiedRules(context, sheet) {
  const rules = sheet.getUsedRange().conditionalFormats.load("items/type");
  await context.sync();
  const custom = rules.items.filter(rule => rule.type === Excel.ConditionalFormatType.custom);
  custom.forEach(rule => rule.custom.rule.load("formula"));
  await context.sync();
  custom.forEach(rule => {
    if (String(rule.custom.rule.formula || "").includes("TIA_ANALYSIS")) rule.delete();
  });
  await context.sync();
}

function blankRows(range) {
  return range.values.every((row, r) => row.every((value, c) =>
    (value == null || value === "") && !String(range.formulas[r][c] ?? "").startsWith("=")));
}

async function resizeSubject(context, sheet, source, template, data, progress = () => {}) {
  const layout = subjectLayout(data.questions, data.roster.length);
  const width = template.totalCol + 1;
  // Work only on the new copy. Remove old sample answers and cyclic legacy formulas first.
  progress("clearing the copied sample answers");
  sheet.getRangeByIndexes(0, 0, 9, width).unmerge();
  sheet.getRangeByIndexes(0, 0, 9, width).clear("Contents");
  sheet.getRangeByIndexes(9, 0, template.students + template.analysisRows, width).clear("Contents");
  await context.sync();
  progress("reading the copied color rules");
  await clearCopiedRules(context, sheet);
  progress("resizing the question columns");
  const extra = data.questions - template.capacity;
  if (extra > 0) {
    sheet.getRangeByIndexes(0, template.totalCol - 1, 1, extra).getEntireColumn().insert("Right");
    await context.sync();
    sheet.getRangeByIndexes(8, template.totalCol - 1, template.footer + template.analysisRows - 9, extra)
      .copyFrom(source.getRangeByIndexes(8, template.totalCol - 2, template.footer + template.analysisRows - 9, 1), "Formats");
  } else if (extra < 0) {
    sheet.getRangeByIndexes(0, data.questions + 2, 1, -extra).getEntireColumn().delete("Left");
  }
  await context.sync();
  progress("resizing the student rows");
  const difference = data.roster.length - template.students;
  if (difference > 0) {
    sheet.getRangeByIndexes(template.footer - 1, 0, difference, 1).getEntireRow().insert("Down");
  } else if (difference < 0) {
    sheet.getRangeByIndexes(data.roster.length + 9, 0, -difference, 1).getEntireRow().delete("Up");
  }
  await context.sync();
  progress("formatting the student rows");
  const body = sheet.getRangeByIndexes(9, 0, data.roster.length, layout.score);
  body.copyFrom(sheet.getRangeByIndexes(9, 0, 1, layout.score), "Formats");
  body.clear("Contents");
  body.getEntireRow().rowHidden = false;
  body.format.rowHeight = template.studentHeight;
  body.format.verticalAlignment = "Center";
  sheet.getRangeByIndexes(0, 2, 1, data.questions).getEntireColumn().format.columnWidth = Math.max(MIN_ITEM_WIDTH, template.widths[2]);
  // Keep the template's narrow, rotated score columns, rather than making four wide columns.
  sheet.getRangeByIndexes(0, layout.total - 1, 1, 1).getEntireColumn().format.columnWidth = Math.max(24, template.widths[template.totalCol - 1]);
  sheet.getRangeByIndexes(0, layout.score - 1, 1, 1).getEntireColumn().format.columnWidth = Math.max(33, template.widths[template.totalCol]);
  await context.sync();
  progress("making space for the analysis rows");
  if (template.analysisRows === 4) {
    const next = sheet.getRangeByIndexes(layout.footer + 3, 0, 2, Math.max(layout.score, template.usedWidth + extra)).load("values,formulas");
    await context.sync();
    if (!blankRows(next)) sheet.getRangeByIndexes(layout.footer + 3, 0, 2, 1).getEntireRow().insert("Down");
    await context.sync();
    sheet.getRangeByIndexes(layout.footer + 3, 0, 2, layout.score)
      .copyFrom(sheet.getRangeByIndexes(layout.footer + 2, 0, 1, layout.score), "Formats");
    sheet.getRangeByIndexes(layout.footer + 3, 0, 2, layout.score).format.rowHeight = 19.5;
  }
  const guideFirst = layout.footer + 6;
  progress("preserving the interpretation notes");
  const guides = sheet.getRangeByIndexes(guideFirst - 1, 0, 3, Math.max(layout.score, template.usedWidth + extra)).load("values,formulas");
  await context.sync();
  const existing = GUIDE_LINES.every((line, i) => String(guides.values[i][0] || "").startsWith(line.split(":")[0] + ":"));
  if (!existing && !blankRows(guides)) {
    sheet.getRangeByIndexes(guideFirst - 1, 0, 3, 1).getEntireRow().insert("Down");
    await context.sync();
  }
  return layout;
}

function renderAnalysis(sheet, data, cells) {
  const a = analysisFormulas(data.questions, data.roster.length, cells);
  const { questions: q, students: n, footer, total, score, rank, group } = a;
  const totalL = column(total), scoreL = column(score), groupL = column(group);
  const helper = sheet.getRangeByIndexes(5, rank - 1, n + 4, 2);
  helper.format.font.name = "Arial";
  helper.format.font.size = 10;
  helper.format.verticalAlignment = "Center";
  sheet.getRangeByIndexes(5, rank - 1, 3, 2).values = [["Group %", 0.33], ["Each group", ""], ["Basis", "All listed pupils"]];
  sheet.getRangeByIndexes(8, rank - 1, 1, 2).values = [["Analysis Rank", "Analysis Group"]];
  sheet.getRangeByIndexes(8, rank - 1, 1, 2).format.fill.color = GREEN;
  sheet.getRangeByIndexes(8, rank - 1, 1, 2).format.font.bold = true;
  sheet.getRangeByIndexes(5, rank - 1, 4, 2).format.wrapText = true;
  sheet.getRangeByIndexes(0, rank - 1, 1, 1).getEntireColumn().format.columnWidth = 75;
  sheet.getRangeByIndexes(0, group - 1, 1, 1).getEntireColumn().format.columnWidth = 108.75;
  sheet.getRange(groupL + "6").numberFormat = [["0%"]];
  sheet.getRange(groupL + "6").format.fill.color = PALE;
  sheet.getRange(groupL + "6").dataValidation.rule = { decimal: { operator: "Between", formula1: 0.000001, formula2: 0.5 } };
  sheet.getRange(groupL + "7").formulas = [[a.groupSize]];
  sheet.getRangeByIndexes(9, total - 1, n, 2).formulas = a.scores;
  sheet.getRangeByIndexes(9, total - 1, n, 2).format.font.name = "Arial";
  sheet.getRangeByIndexes(9, total - 1, n, 2).format.font.size = 10;
  sheet.getRangeByIndexes(9, total - 1, n, 2).format.horizontalAlignment = "Center";
  sheet.getRangeByIndexes(9, total - 1, n, 1).numberFormat = Array.from({ length: n }, () => ["0"]);
  sheet.getRangeByIndexes(9, score - 1, n, 1).numberFormat = Array.from({ length: n }, () => ["0%"]);
  sheet.getRangeByIndexes(9, rank - 1, n, 2).formulas = a.ranks;
  const labels = ["# Of Items Answered Correctly", "% Of Students Who Answered Correctly", "Item Difficulty (Df)", "Discrimination Index (D)", "Description", "Decision"];
  const analysis = sheet.getRangeByIndexes(footer - 1, 0, 6, score);
  analysis.unmerge(); analysis.clear("Contents");
  labels.forEach((value, i) => {
    sheet.getRange("A" + (footer + i) + ":B" + (footer + i)).merge();
    sheet.getRange("A" + (footer + i)).values = [[value]];
  });
  sheet.getRangeByIndexes(footer - 1, 0, 6, 2).format.wrapText = true;
  const results = sheet.getRangeByIndexes(footer - 1, 2, 6, q);
  results.formulas = a.itemRows;
  results.format.font.name = "Arial";
  results.format.horizontalAlignment = "Center";
  results.format.verticalAlignment = "Center";
  sheet.getRangeByIndexes(footer - 1, 2, 1, q).numberFormat = [Array(q).fill("0")];
  sheet.getRangeByIndexes(footer, 2, 1, q).numberFormat = [Array(q).fill("0%")];
  sheet.getRangeByIndexes(footer + 1, 2, 2, q).numberFormat = Array.from({ length: 2 }, () => Array(q).fill("0.00"));
  sheet.getRangeByIndexes(footer + 1, 2, 2, q).format.font.size = 10;
  sheet.getRangeByIndexes(footer + 1, 2, 4, q).format.textOrientation = 0;
  sheet.getRangeByIndexes(footer + 3, 2, 2, q).format.font.size = 9;
  sheet.getRangeByIndexes(footer + 3, 2, 2, q).numberFormat = [Array(q).fill("@"), Array(q).fill("@")];
  sheet.getRange(totalL + footer + ":" + scoreL + footer).merge();
  sheet.getRange(totalL + footer).values = [["Mean Percentage Score"]];
  sheet.getRange(totalL + footer + ":" + scoreL + footer).format.wrapText = true;
  sheet.getRange(totalL + footer + ":" + scoreL + footer).format.font.size = 8;
  sheet.getRange(totalL + (footer + 1) + ":" + scoreL + (footer + 1)).merge();
  sheet.getRange(totalL + (footer + 1)).formulas = [[a.mps]];
  sheet.getRange(totalL + (footer + 1)).numberFormat = [["0%"]];
  const grid = sheet.getRangeByIndexes(9, 2, n, q);
  grid.format.font.name = "Arial";
  grid.format.fill.color = "#FFFFFF";
  coloredRule(grid, '=$' + groupL + '10="Upper"', PALE, true);
  coloredRule(grid, '=$' + groupL + '10="Lower"', "#D9EAD3", true);
  const decisions = sheet.getRangeByIndexes(footer + 4, 2, 1, q);
  for (const [value, color] of [["Review", "#FCE4D6"], ["Revise", PALE], ["Retain", "#E2F0D9"]]) {
    coloredRule(decisions, '=C' + (footer + 5) + '="' + value + '"', color, true);
  }
  const guideFirst = footer + 6;
  sheet.getRangeByIndexes(guideFirst - 1, 0, 3, score).unmerge();
  const reportWidth = sheet.getRangeByIndexes(guideFirst - 1, 0, 3, score);
  reportWidth.format.font.name = "Arial"; reportWidth.format.font.size = 10;
  reportWidth.format.font.bold = false; reportWidth.format.font.color = "#333333";
  reportWidth.format.fill.color = "#FFFFFF"; reportWidth.format.wrapText = true;
  reportWidth.format.textOrientation = 0; reportWidth.format.verticalAlignment = "Center";
  GUIDE_LINES.forEach((value, i) => {
    const range = sheet.getRangeByIndexes(guideFirst + i - 1, 0, 1, score);
    range.merge(); sheet.getCell(guideFirst + i - 1, 0).values = [[value]];
    range.format.rowHeight = q < 6 ? 42 : 24;
  });
  sheet.freezePanes.freezeAt("A1:B9");
  hideTail(sheet, group);
}

function renderPaper(sheet, data) {
  const names = data.roster.filter(student => student.paper).map(student => student.name);
  const width = names.length + 2, first = 4;
  baseFormat(sheet.getRangeByIndexes(0, 0, Math.max(data.questions + 3, names.length + 1), width + 1));
  sheet.getRange("A1").values = [["Select Student"]];
  sheet.getRange("B1").values = [["Show all"]];
  sheet.getRange("A1:B1").format.fill.color = BLUE;
  sheet.getRange("A3").values = [["Item"]];
  sheet.getRange("B3").values = [["Key Answers"]];
  sheet.getRangeByIndexes(2, 2, 1, names.length).values = [names.map(literal)];
  sheet.getRangeByIndexes(3, 0, data.questions, 1).values = Array.from({ length: data.questions }, (_, i) => [`Item ${i + 1}`]);
  sheet.getRangeByIndexes(2, 0, 1, width).format.fill.color = BLUE;
  sheet.getRangeByIndexes(2, 0, 1, width).format.font.bold = true;
  sheet.getRangeByIndexes(2, 0, 1, width).format.wrapText = true;
  sheet.getRangeByIndexes(2, 0, 1, width).format.rowHeight = 33;
  sheet.getRange("A:B").format.columnWidth = 82.5;
  sheet.getRangeByIndexes(0, 2, 1, names.length).getEntireColumn().format.columnWidth = 108.75;
  sheet.getRangeByIndexes(3, 1, data.questions, 1).format.fill.color = PALE;
  const options = ["Show all", ...[...names].sort((a, b) => a.localeCompare(b, "en", { sensitivity: "base", numeric: true }))];
  const list = sheet.getRangeByIndexes(0, width, options.length, 1);
  list.values = options.map(value => [literal(value)]);
  sheet.getRange("B1").dataValidation.rule = { list: { inCellDropDown: true, source: list } };
  const answers = sheet.getRangeByIndexes(first - 1, 2, data.questions, names.length);
  const base = `=AND(LEN(TRIM($B${first}))>0,LEN(TRIM(C${first}))>0,UPPER(TRIM(C${first}))`;
  coloredRule(answers, `${base}=UPPER(TRIM($B${first})))`, "#D9EAD3");
  coloredRule(answers, `${base}<>UPPER(TRIM($B${first})))`, "#F4CCCC");
  sheet.freezePanes.freezeAt("A1:B3");
  hideTail(sheet, width);
  return names;
}

export async function generateSubject() {
  return guarded(async () => Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/id,items/position");
    const state = await registry(context);
    const input = sheets.items.find(sheet => sheet.name === "Generate");
    const source = sheets.items.find(sheet => sheet.name === "Original");
    if (!input || !source) throw new Error("Open the workbook containing your Original and Generate tabs.");
    const template = await readTemplate(context, source);
    const data = await readGenerate(context, input);
    state.subjects = state.subjects.filter(subject => sheets.items.some(sheet => sheet.id === subject.id));
    if (state.subjects.some(subject => key(subject.title) === key(data.subject)) ||
        sheets.items.some(sheet => key(sheet.name) === key(data.subject))) {
      throw new Error("The subject " + data.subject + " already exists. Use another subject title for a new exam.");
    }
    const tabName = uniqueTabName(data.subject, sheets.items.map(sheet => sheet.name));
    const paperName = uniqueTabName(data.subject, [...sheets.items.map(sheet => sheet.name), tabName], " - Manual Checking");
    const created = [];
    let stage = "copying Original";
    try {
      const subject = source.copy(Excel.WorksheetPositionType.after, input);
      subject.load("id,position");
      await context.sync();
      created.push(subject.id);
      subject.name = tabName;
      subject.visibility = "Visible";
      await resizeSubject(context, subject, source, template, data, value => { stage = value; });
      stage = "writing the student roster and subject header";
      subject.getRangeByIndexes(9, 0, data.roster.length, 2).values =
        data.roster.map(student => [literal(student.name), student.order]);
      const cells = renderHeader(subject, data, template);
      await context.sync();
      stage = "building the item analysis";
      renderAnalysis(subject, data, cells);
      await context.sync();
      let paper = null, paperNames = [];
      if (data.roster.some(student => student.paper)) {
        stage = "creating the paper checking sheet";
        paper = sheets.add(paperName);
        paper.load("id");
        await context.sync();
        created.push(paper.id);
        paper.position = subject.position + 1;
        paperNames = renderPaper(paper, data);
      }
      stage = "finishing the generated sheets";
      subject.activate();
      await context.sync();
      state.setup = "template-2";
      state.subjects.push({ id: subject.id, title: data.subject, tabName, questions: data.questions,
        students: data.roster.length, paperId: paper?.id || "", paperNames, cells,
        layoutVersion: 2, onlineStudentKeys: [] });
      stage = "saving the subject links";
      saveRegistry(context, state);
      await context.sync();
      return "Created " + tabName + ": " + data.questions + " questions and " + data.roster.length +
        " students from Original." + (paper ? " Created " + paperName + " for " + paperNames.length + " paper students." : "");
    } catch (error) {
      let cleanupError;
      try { await cleanup(created); } catch (failure) { cleanupError = failure; }
      const failure = new Error("Generation stopped while " + stage + ". " + (error.message || String(error)) +
        (cleanupError ? " Cleanup also failed: " + cleanupError.message : " The unfinished subject copy was removed."));
      failure.stage = stage;
      failure.code = error.code;
      failure.debugInfo = error.debugInfo;
      throw failure;
    }
  }));
}

async function cleanup(ids) {
  if (!ids.length) return;
  try {
    await Excel.run(async context => {
      for (const id of [...ids].reverse()) {
        const sheet = context.workbook.worksheets.getItemOrNullObject(id);
        sheet.load("name");
        await context.sync();
        if (!sheet.isNullObject) { sheet.visibility = "Visible"; sheet.delete(); await context.sync(); }
      }
    });
  } catch (error) {
    throw new Error(`The operation stopped and unfinished tabs could not be removed. Check the workbook. ${error.message || error}`);
  }
}

function selectedSubject(context, state, activeId) {
  const found = state.subjects.find(subject => subject.id === activeId || subject.paperId === activeId);
  if (!found) throw new Error("Open a generated subject tab or its Manual Checking tab first.");
  return found;
}

async function applyView(context, config) {
  if (!config.paperId) throw new Error("This subject has no Manual Checking tab.");
  const paper = context.workbook.worksheets.getItem(config.paperId);
  const selector = paper.getRange("B1").load("text");
  const names = paper.getRangeByIndexes(2, 2, 1, config.paperNames.length).load("text");
  await context.sync();
  const selected = key(selector.text[0][0]);
  const index = names.text[0].findIndex(name => key(name) === selected);
  if (selected !== "show all" && index < 0) throw new Error("Choose Show all or a student from the B1 dropdown.");
  paper.getRangeByIndexes(0, 0, 1, config.paperNames.length + 2).getEntireColumn().columnHidden = false;
  if (selected !== "show all") {
    if (index > 0) paper.getRangeByIndexes(0, 2, 1, index).getEntireColumn().columnHidden = true;
    if (index + 1 < config.paperNames.length) paper.getRangeByIndexes(0, index + 3, 1, config.paperNames.length - index - 1).getEntireColumn().columnHidden = true;
  }
  await context.sync();
}

export async function applyStudentView() {
  return guarded(async () => Excel.run(async context => {
    const active = context.workbook.worksheets.getActiveWorksheet().load("id");
    const state = await registry(context);
    const config = selectedSubject(context, state, active.id);
    await applyView(context, config);
    return "Applied the student selection from Manual Checking!B1.";
  }));
}

export async function gradePaper() {
  return guarded(async () => Excel.run(async context => {
    const active = context.workbook.worksheets.getActiveWorksheet().load("id");
    const state = await registry(context);
    const config = selectedSubject(context, state, active.id);
    if (!config.paperId) throw new Error("No Manual Checking tab exists. Generate a subject with paper students highlighted in red.");
    const subject = context.workbook.worksheets.getItem(config.id);
    const paper = context.workbook.worksheets.getItem(config.paperId);
    const matrix = paper.getRangeByIndexes(2, 0, config.questions + 1, config.paperNames.length + 2).load("text");
    const items = subject.getRangeByIndexes(8, 2, 1, config.questions).load("values");
    const names = subject.getRangeByIndexes(9, 0, config.students, 1).load("text");
    await context.sync();
    const order = itemOrder(items.values[0], config.questions);
    const entries = paperGrades(matrix.text, order);
    const target = new Map();
    names.text.forEach((row, index) => {
      const name = key(row[0]);
      if (!name) return;
      if (target.has(name)) throw new Error(`Duplicate roster name: ${row[0]}.`);
      target.set(name, index + 9);
    });
    const protectedNames = new Set(config.onlineStudentKeys || []);
    for (const entry of entries) if (!target.has(entry.key)) throw new Error(`Paper student ${entry.name} does not match the subject roster.`);
    const updates = entries.filter(entry => !protectedNames.has(entry.key));
    for (const entry of updates) subject.getRangeByIndexes(target.get(entry.key), 2, 1, config.questions).values = [entry.marks];
    await context.sync();
    return updates.length ? `Graded ${updates.length} paper student(s). Scores and analysis recalculated.` : "No paper answers were available to transfer. Enter the key and student answers first.";
  }));
}

async function guarded(action) {
  if (writing) throw new Error("Another operation is running. Wait for it to finish.");
  writing = true;
  try { return await action(); } finally { writing = false; }
}

export async function registerEvents(report) {
  if (eventsRegistered) return;
  await Excel.run(async context => {
    context.workbook.worksheets.onChanged.add(async event => {
      if (writing) return;
      try {
        await guarded(async () => Excel.run(async inner => {
          const state = await registry(inner);
          const config = state.subjects.find(subject => subject.id === event.worksheetId || subject.paperId === event.worksheetId);
          if (!config) return;
          const sheet = inner.workbook.worksheets.getItem(event.worksheetId);
          const changed = sheet.getRange(event.address).load("rowIndex,columnIndex,rowCount,columnCount");
          await inner.sync();
          if (config.paperId === event.worksheetId) {
            if (changed.rowIndex <= 0 && changed.rowIndex + changed.rowCount > 0 && changed.columnIndex <= 1 && changed.columnIndex + changed.columnCount > 1) await applyView(inner, config);
            return;
          }
          const top = Math.max(changed.rowIndex, 9), bottom = Math.min(changed.rowIndex + changed.rowCount, config.students + 9);
          const left = Math.max(changed.columnIndex, 2), right = Math.min(changed.columnIndex + changed.columnCount, config.questions + 2);
          if (bottom <= top || right <= left) return;
          const range = sheet.getRangeByIndexes(top, left, bottom - top, right - left).load("values,formulas");
          await inner.sync();
          for (let r = 0; r < range.values.length; r++) for (let c = 0; c < range.values[r].length; c++) {
            if (!String(range.formulas[r][c] ?? "").startsWith("=") && isCheck(range.values[r][c])) {
              sheet.getCell(top + r, left + c).values = [[CHECK]];
            }
          }
          await inner.sync();
        }));
      } catch (error) { report(error.message || String(error)); }
    });
    await context.sync();
  });
  eventsRegistered = true;
}
