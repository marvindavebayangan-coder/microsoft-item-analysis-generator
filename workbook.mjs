import { CHECK, column, key, literal, text, isCheck, readInput, uniqueTabName, itemOrder, paperGrades, todaySerial } from "./domain.mjs?v=0.1.0";

const REGISTRY = "TIA.Generator.Registry.v1";
const MAX_COLUMNS = 16384;
const GREEN = "#C6EFCE";
const BLUE = "#D9EAF7";
const PALE = "#FFF2CC";
const SAMPLE_NAMES = ["Student 01", "Student 02", "Student 03", "Student 04", "Student 05", "Student 06"];
let writing = false;
let eventsRegistered = false;

async function registry(context) {
  const setting = context.workbook.settings.getItemOrNullObject(REGISTRY);
  setting.load("value");
  await context.sync();
  return setting.isNullObject ? { version: 1, subjects: [] } : JSON.parse(setting.value);
}

function saveRegistry(context, state) {
  context.workbook.settings.add(REGISTRY, JSON.stringify(state));
}

function baseFormat(range) {
  range.format.font.name = "Arial";
  range.format.font.size = 10;
  range.format.verticalAlignment = "Center";
  range.format.rowHeight = 22;
}

function coloredRule(range, formula, color) {
  const rule = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
  rule.custom.rule.formula = formula;
  rule.custom.format.fill.color = color;
}

function hideTail(sheet, width) {
  if (width < MAX_COLUMNS) {
    sheet.getRangeByIndexes(0, width, 1, MAX_COLUMNS - width).getEntireColumn().columnHidden = true;
  }
}

function template(sheet) {
  baseFormat(sheet.getRange("A1:DB125"));
  sheet.getRange("A:A").format.columnWidth = 220;
  sheet.getRange("B:B").format.columnWidth = 52;
  sheet.getRange("C:CX").format.columnWidth = 38;
  sheet.getRange("A9:DB9").format.fill.color = GREEN;
  sheet.getRange("A9:DB9").format.font.bold = true;
  sheet.getRange("A9:DB9").format.rowHeight = 56;
  sheet.getRange("A9:DB9").format.wrapText = true;
  sheet.getRange("A9:B9").values = [["Student name", "Student #"]];
  sheet.getRange("C9:CX9").values = [Array.from({ length: 100 }, (_, i) => i + 1)];
  sheet.getRange("CY9:DB9").values = [["Correct answers", "Student score", "Analysis Rank", "Analysis Group"]];
  sheet.getRange("A110:A115").values = [["Items answered correctly"], ["% correct"], ["Item Difficulty (Df)"],
    ["Discrimination Index (D)"], ["Description"], ["Decision"]];
}

function generateInputs(sheet) {
  baseFormat(sheet.getRange("A1:H114"));
  sheet.getRange("A:A").format.columnWidth = 235;
  sheet.getRange("B:B").format.columnWidth = 230;
  sheet.getRange("D:H").format.columnWidth = 90;
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [["Microsoft Item Analysis Generator"]];
  sheet.getRange("A1:H1").format.fill.color = "#176443";
  sheet.getRange("A1:H1").format.font.color = "#FFFFFF";
  sheet.getRange("A1:H1").format.font.bold = true;
  sheet.getRange("A1:H1").format.rowHeight = 32;
  sheet.getRange("A3:H3").merge();
  sheet.getRange("A3").values = [["Prototype sample data. Replace these details and student names for another test."]];
  sheet.getRange("A3:H3").format.wrapText = true;
  sheet.getRange("A3:H3").format.rowHeight = 34;
  sheet.getRange("A5:A12").values = [["School"], ["District"], ["Assessment"], ["Period"],
    ["Date of Assessment"], ["Teacher"], ["Subject / Grade(s)"], ["# of Questions"]];
  sheet.getRange("B5:B12").values = [["Sample School"], ["Sample District"], ["Practice Test"],
    ["1st Quarter"], [todaySerial()], ["Sample Teacher"], ["Sample Math"], [5]];
  sheet.getRange("A5:A12").format.font.bold = true;
  sheet.getRange("B5:B12").format.fill.color = GREEN;
  sheet.getRange("B9").numberFormat = [["m/d/yyyy"]];
  sheet.getRange("B12").dataValidation.rule = { wholeNumber: { operator: "Between", formula1: 1, formula2: 100 } };
  sheet.getRange("A14").values = [["Student Names here"]];
  sheet.getRange("A14").format.font.bold = true;
  sheet.getRange("A14").format.fill.color = BLUE;
  sheet.getRange("A15:A20").values = SAMPLE_NAMES.map(name => [name]);
  sheet.getRange("A17").format.fill.color = "#FF0000";
  sheet.getRange("A20").format.fill.color = "#FF0000";
  const instructions = [
    "1. Enter the class details in the green cells.",
    "2. Enter up to 100 student names below the heading in column A.",
    "3. Fill a name cell with plain red (#FF0000) for a paper student.",
    "4. In the add-in, choose Generate Subject.",
    "Blank name cells are ignored; the student count is automatic.",
    "All red = all paper. No red = all online or direct score entry.",
    "The sample marks Student 03 and Student 06 as paper students."
  ];
  instructions.forEach((value, index) => {
    const row = index + 5;
    sheet.getRange(`D${row}:H${row}`).merge();
    sheet.getRange(`D${row}`).values = [[value]];
    sheet.getRange(`D${row}:H${row}`).format.wrapText = true;
    sheet.getRange(`D${row}:H${row}`).format.rowHeight = 36;
  });
  sheet.freezePanes.freezeAt("A1:A1");
}

export async function prepareWorkbook() {
  return guarded(async () => Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/id");
    const state = await registry(context);
    const input = sheets.items.find(sheet => key(sheet.name) === "generate");
    const original = sheets.items.find(sheet => key(sheet.name) === "original");
    if (input || original) {
      if (input && original && state.setup === "prototype-1") {
        input.activate();
        await context.sync();
        return "This workbook is already prepared. Existing inputs and results were kept.";
      }
      throw new Error("A Generate or Original tab already exists. Use a new blank workbook for this prototype.");
    }
    const created = [];
    try {
      const generate = sheets.add("Generate");
      generate.load("id");
      const source = sheets.add("Original");
      source.load("id");
      await context.sync();
      created.push(source.id, generate.id);
      template(source);
      generateInputs(generate);
      generate.activate();
      source.visibility = "Hidden";
      await context.sync();
      state.setup = "prototype-1";
      saveRegistry(context, state);
      await context.sync();
      return "Prepared Generate and hidden Original. Sample: 5 questions, 6 students, 2 red paper students. Next choose Generate Subject.";
    } catch (error) {
      await cleanup(created);
      throw error;
    }
  }));
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

function renderHeader(sheet, data) {
  const lastItem = data.questions + 2;
  const score = data.questions + 4;
  const cells = {};
  const label = (row, col, width, value) => {
    const range = sheet.getRangeByIndexes(row - 1, col - 1, 1, width);
    if (width > 1) range.merge();
    sheet.getCell(row - 1, col - 1).values = [[value]];
    range.format.font.bold = true;
    range.format.wrapText = true;
  };
  const field = (row, col, width, name) => {
    const range = sheet.getRangeByIndexes(row - 1, col - 1, 1, width);
    if (width > 1) range.merge();
    sheet.getCell(row - 1, col - 1).values = [[literal(name === "students" ? data.roster.length : data[name])]];
    range.format.fill.color = GREEN;
    range.format.wrapText = true;
    cells[name] = column(col) + row;
    if (name === "date") sheet.getCell(row - 1, col - 1).numberFormat = [["m/d/yyyy"]];
  };
  if (lastItem >= 8) {
    const left = Math.floor(lastItem / 2);
    const right = left + 1;
    const value = right + 2;
    [[1, "SCHOOL", "school", "DISTRICT", "district"],
      [6, "Assessment", "assessment", "Teacher", "teacher"],
      [7, "Period", "period", "Subject / Grade(s)", "subject"]].forEach(([row, a, fieldA, b, fieldB]) => {
      label(row, 1, 1, a); field(row, 2, left - 1, fieldA);
      label(row, right, 2, b); field(row, value, lastItem - value + 1, fieldB);
    });
    label(2, 1, score, "Test Item Analysis Calculator");
    sheet.getRangeByIndexes(1, 0, 1, score).format.horizontalAlignment = "Center";
    sheet.getRangeByIndexes(1, 0, 1, score).format.fill.color = GREEN;
    label(8, 1, 1, "Date of Assessment"); field(8, 2, left - 1, "date");
    const qValue = Math.floor((right + lastItem) / 2);
    label(8, right, qValue - right, "# of Questions"); field(8, qValue, 1, "questions");
    label(8, qValue + 1, lastItem - qValue - 1, "# of Students"); field(8, lastItem, 1, "students");
    sheet.getRangeByIndexes(4, 0, 1, score).format.fill.color = GREEN;
  } else {
    [["School", "school"], ["District", "district"], ["Assessment", "assessment"], ["Period", "period"],
      ["Date of Assessment", "date"], ["Teacher", "teacher"], ["Subject / Grade(s)", "subject"]].forEach(([name, property], i) => {
      label(i + 1, 1, 1, name); field(i + 1, 2, score - 1, property);
    });
    label(8, 1, 1, "# of Questions"); field(8, 2, 1, "questions");
    label(8, 3, score - 3, "# of Students"); field(8, score, 1, "students");
  }
  sheet.getRangeByIndexes(0, 0, 8, score).format.rowHeight = 36;
  sheet.getRange(cells.students).formulas = [[`=COUNTIF($A$10:$A$${9 + data.roster.length},"<>")`]];
  return cells;
}

function renderAnalysis(sheet, data, cells) {
  const q = data.questions, n = data.roster.length, last = n + 9, footer = last + 1;
  const total = q + 3, score = q + 4, rank = q + 5, group = q + 6;
  const totalL = column(total), rankL = column(rank), groupL = column(group), lastQ = column(q + 2);
  const absolute = address => address.replace(/([A-Z]+)(\d+)/, "$$$1$$$2");
  const students = absolute(cells.students), questions = absolute(cells.questions), size = `$${groupL}$7`;
  const names = `$A$10:$A$${last}`, totals = `$${totalL}$10:$${totalL}$${last}`, groups = `$${groupL}$10:$${groupL}$${last}`;
  sheet.getRangeByIndexes(8, total - 1, 1, 4).values = [["Correct answers", "Student score", "Analysis Rank", "Analysis Group"]];
  sheet.getRangeByIndexes(5, rank - 1, 3, 2).values = [["Group %", 0.33], ["Each group", ""], ["Basis", "All listed pupils"]];
  sheet.getRange(`${groupL}6`).numberFormat = [["0%"]];
  sheet.getRange(`${groupL}6`).format.fill.color = PALE;
  sheet.getRange(`${groupL}6`).dataValidation.rule = { decimal: { operator: "Between", formula1: 0.000001, formula2: 0.5 } };
  sheet.getRange(`${groupL}7`).formulas = [[`=IF(${students}<2,0,MAX(1,MIN(INT(${students}/2),ROUND(${students}*$${groupL}$6,0))))`]];
  const rowFormulas = [], rankFormulas = [];
  for (let r = 10; r <= last; r++) {
    rowFormulas.push([`=IF($A${r}="","",COUNTIF(C${r}:${lastQ}${r},"${CHECK}"))`,
      `=IF($A${r}="","",${totalL}${r}/${questions})`]);
    rankFormulas.push([`=IF($A${r}="","",COUNTIFS(${totals},">"&${totalL}${r},${names},"<>")+COUNTIFS(${totals},${totalL}${r},$B$10:$B$${last},"<="&B${r},${names},"<>"))`,
      `=IF(OR($A${r}="",${size}=0),"",IF(${rankL}${r}<=${size},"Upper",IF(${rankL}${r}>${students}-${size},"Lower","Middle")))`]);
  }
  sheet.getRangeByIndexes(9, total - 1, n, 2).formulas = rowFormulas;
  sheet.getRangeByIndexes(9, score - 1, n, 1).numberFormat = Array.from({ length: n }, () => ["0%"]);
  sheet.getRangeByIndexes(9, rank - 1, n, 2).formulas = rankFormulas;
  const labels = ["Items answered correctly", "% correct", "Item Difficulty (Df)", "Discrimination Index (D)", "Description", "Decision"];
  const formulas = labels.map(() => []);
  for (let i = 0; i < q; i++) {
    const c = column(i + 3), answers = `${c}$10:${c}$${last}`, d = `${c}${footer + 3}`;
    const upper = `COUNTIFS(${groups},"Upper",${answers},"${CHECK}")`;
    const lower = `COUNTIFS(${groups},"Lower",${answers},"${CHECK}")`;
    formulas[0].push(`=COUNTIFS(${names},"<>",${answers},"${CHECK}")`);
    formulas[1].push(`=${c}${footer}/${students}`);
    formulas[2].push(`=IF(${size}=0,"",(${upper}+${lower})/(2*${size}))`);
    formulas[3].push(`=IF(${size}=0,"",ROUND((${upper}-${lower})/${size},10))`);
    formulas[4].push(`=IF(ISNUMBER(${d}),IF(${d}<0.2,"PI",IF(${d}<0.3,"NI",IF(${d}<0.4,"GI","VG"))),"")`);
    formulas[5].push(`=IF(ISNUMBER(${d}),IF(${d}<0.2,"Review",IF(${d}<0.3,"Revise","Retain")),"Review")`);
  }
  labels.forEach((value, i) => {
    sheet.getRange(`A${footer + i}:B${footer + i}`).merge();
    sheet.getRange(`A${footer + i}`).values = [[value]];
  });
  sheet.getRangeByIndexes(footer - 1, 0, 6, 2).format.wrapText = true;
  sheet.getRangeByIndexes(footer - 1, 2, 6, q).formulas = formulas;
  sheet.getRangeByIndexes(footer, 2, 1, q).numberFormat = [Array(q).fill("0%")];
  sheet.getRangeByIndexes(footer + 1, 2, 2, q).numberFormat = Array.from({ length: 2 }, () => Array(q).fill("0.00"));
  sheet.getRangeByIndexes(footer - 1, total - 1, 1, 2).values = [["MPS", ""]];
  sheet.getRangeByIndexes(footer - 1, score - 1, 1, 1).formulas = [[`=AVERAGE(C${footer + 1}:${lastQ}${footer + 1})`]];
  sheet.getRangeByIndexes(footer - 1, score - 1, 1, 1).numberFormat = [["0%"]];
  const grid = sheet.getRangeByIndexes(9, 2, n, q);
  grid.format.fill.color = "#FFFFFF";
  coloredRule(grid, `=$${groupL}10="Upper"`, PALE);
  coloredRule(grid, `=$${groupL}10="Lower"`, "#D9EAD3");
  const decisions = sheet.getRangeByIndexes(footer + 4, 2, 1, q);
  for (const [value, color] of [["Review", "#FCE4D6"], ["Revise", PALE], ["Retain", "#E2F0D9"]]) {
    coloredRule(decisions, `=C${footer + 5}="${value}"`, color);
  }
  const guides = [
    "Decision rules: Retain if D >= 0.30. Revise if 0.20 <= D < 0.30. Review if D < 0.20 or group data are unavailable.",
    "Review guidance: check the answer key, wording and learning objective. All-correct/all-wrong items have D = 0. Teachers make the final decision.",
    "Analysis basis: all listed pupils. Finish marking/submissions before interpreting decisions. Difficulty uses upper/lower groups. Ties use roster order."
  ];
  guides.forEach((value, i) => {
    const row = footer + 6 + i;
    sheet.getRangeByIndexes(row - 1, 0, 1, score).merge();
    sheet.getRange(`A${row}`).values = [[value]];
    sheet.getRangeByIndexes(row - 1, 0, 1, score).format.wrapText = true;
    sheet.getRangeByIndexes(row - 1, 0, 1, score).format.rowHeight = q < 6 ? 56 : 32;
  });
  sheet.getRangeByIndexes(0, total - 1, 1, 4).getEntireColumn().format.columnWidth = 90;
  sheet.getRangeByIndexes(8, 0, 1, group).format.wrapText = true;
  sheet.getRangeByIndexes(8, 0, 1, group).format.rowHeight = 56;
  sheet.getRangeByIndexes(8, 0, 1, group).format.fill.color = GREEN;
  sheet.getRangeByIndexes(8, 0, 1, group).format.font.bold = true;
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
  sheet.getRangeByIndexes(2, 0, 1, width).format.rowHeight = 44;
  sheet.getRange("A:B").format.columnWidth = 110;
  sheet.getRangeByIndexes(0, 2, 1, names.length).getEntireColumn().format.columnWidth = 145;
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
    if (state.setup !== "prototype-1") throw new Error("Use Prepare Sample Workbook first in a blank workbook.");
    const input = sheets.items.find(sheet => sheet.name === "Generate");
    const source = sheets.items.find(sheet => sheet.name === "Original");
    if (!input || !source) throw new Error("Keep Generate and Original in this workbook.");
    const data = await readGenerate(context, input);
    state.subjects = state.subjects.filter(subject => sheets.items.some(sheet => sheet.id === subject.id));
    if (state.subjects.some(subject => key(subject.title) === key(data.subject))) {
      throw new Error(`The subject ${data.subject} already exists. Use a new subject title for another exam.`);
    }
    const tabName = uniqueTabName(data.subject, sheets.items.map(sheet => sheet.name));
    const paperName = uniqueTabName(data.subject, [...sheets.items.map(sheet => sheet.name), tabName], " - Manual Checking");
    const created = [];
    try {
      const subject = source.copy(Excel.WorksheetPositionType.after, input);
      subject.name = tabName;
      subject.visibility = "Visible";
      subject.load("id,position");
      await context.sync();
      created.push(subject.id);
      if (data.questions < 100) subject.getRangeByIndexes(0, data.questions + 2, 1, 100 - data.questions).getEntireColumn().delete(Excel.DeleteShiftDirection.left);
      if (data.roster.length < 100) subject.getRangeByIndexes(data.roster.length + 9, 0, 100 - data.roster.length, 1).getEntireRow().delete(Excel.DeleteShiftDirection.up);
      await context.sync();
      subject.getRangeByIndexes(9, 0, data.roster.length, 2).values = data.roster.map(student => [literal(student.name), student.order]);
      const cells = renderHeader(subject, data);
      renderAnalysis(subject, data, cells);
      let paper = null, paperNames = [];
      if (data.roster.some(student => student.paper)) {
        paper = sheets.add(paperName);
        paper.position = subject.position + 1;
        paper.load("id");
        await context.sync();
        created.push(paper.id);
        paperNames = renderPaper(paper, data);
      }
      subject.activate();
      await context.sync();
      state.subjects.push({ id: subject.id, title: data.subject, tabName, questions: data.questions,
        students: data.roster.length, paperId: paper?.id || "", paperNames, cells, onlineStudentKeys: [] });
      saveRegistry(context, state);
      await context.sync();
      return `Created ${tabName}: ${data.questions} questions and ${data.roster.length} students.` +
        (paper ? ` Created ${paperName} for ${paperNames.length} paper students.` : "");
    } catch (error) {
      await cleanup(created);
      throw error;
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
