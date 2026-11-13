import { CHECK, column, key, itemOrder } from "./domain.mjs?v=0.3.0";

const compact = value => key(value).replace(/[^a-z0-9]/g, "");
export const FIRST_STUDENT = 10;
export const MIN_ITEM_WIDTH = 33; // 44 screen pixels, matching the V14 analysis minimum.
export const GUIDE_LINES = [
  "Decision rules: Retain if D >= 0.30. Revise if 0.20 <= D < 0.30. Review if D < 0.20 or group data are unavailable.",
  "Review guidance: check the answer key, wording and learning objective. All-correct/all-wrong items have D = 0. Teachers make the final decision.",
  "Analysis basis: all listed pupils. Finish marking/submissions before interpreting decisions. Difficulty uses upper/lower groups. Ties use roster order."
];

export function detectTemplate(labels, headers) {
  const footer = labels.findIndex(row => compact(row[0]).includes("itemsansweredcorrectly")) + 1;
  const matches = new Set();
  for (const row of headers) row.forEach((value, index) => {
    if (compact(value).includes("correctanswers")) matches.add(index + 1);
  });
  if (footer < 11 || footer > 110 || matches.size !== 1) {
    throw new Error("Original needs item headings on row 9, student rows starting at row 10, an Items Answered Correctly row, and one Correct Answers column.");
  }
  const totalCol = [...matches][0], capacity = totalCol - 3;
  if (capacity < 1 || capacity > 100 ||
      !headers.some(row => compact(row[totalCol]).includes("studentscore"))) {
    throw new Error("Keep 1–100 numbered item columns followed by Correct Answers and Student Score in Original.");
  }
  itemOrder(headers[3].slice(2, totalCol - 1), capacity);
  const description = compact(labels[footer + 3]?.[0]) === "description";
  const decision = compact(labels[footer + 4]?.[0]) === "decision";
  if (description !== decision) throw new Error("Keep both Description and Decision rows in Original, or leave both for the add-in to add.");
  return { footer, capacity, totalCol, students: footer - FIRST_STUDENT, analysisRows: description ? 6 : 4 };
}

export function subjectLayout(questions, students) {
  if (!Number.isInteger(questions) || questions < 1 || questions > 100 ||
      !Number.isInteger(students) || students < 1 || students > 100) throw new Error("Use 1–100 questions and students.");
  return { questions, students, first: FIRST_STUDENT, last: students + 9, footer: students + 10,
    total: questions + 3, score: questions + 4, rank: questions + 5, group: questions + 6 };
}

export function headerPlan(questions) {
  const lastItem = questions + 2, score = questions + 4;
  const blocks = [], cells = {};
  const add = (row, col, width, value, field, style) => {
    if (width < 1) throw new Error("The question count cannot fit this header.");
    blocks.push({ row, col, width, value, field, style });
    if (field) cells[field] = column(col) + row;
  };
  if (questions >= 6) {
    const left = Math.floor(lastItem / 2), right = left + 1, start = right + 2;
    for (const [row, a, fieldA, b, fieldB, styleA, styleB] of [
      [1, "SCHOOL:", "school", "DISTRICT:", "district", "B1", "T1"],
      [6, "Assessment:", "assessment", "Teacher:", "teacher", "B6", "Q6"],
      [7, "Period:", "period", "Subject / Grade(s):", "subject", "B7", "Q7"]
    ]) {
      add(row, 1, 1, a, null, "A" + row); add(row, 2, left - 1, null, fieldA, styleA);
      add(row, right, 2, b, null, row === 1 ? "Q1" : row === 6 ? "N6" : "L7");
      add(row, start, lastItem - start + 1, null, fieldB, styleB);
    }
    add(2, 1, score, "<<< Test Item Analysis Calculator >>>", null, "A2");
    add(8, 1, 1, "Date of Assessment:", null, "A8"); add(8, 2, left - 1, null, "date", "B8");
    const count = Math.floor((right + lastItem) / 2);
    add(8, right, count - right, "# of Questions:", null, "N8"); add(8, count, 1, null, "questions", "T8");
    add(8, count + 1, lastItem - count - 1, "# of Students:", null, "U8");
    add(8, lastItem, 1, null, "students", "Y8");
  } else {
    for (const [i, [label, field, style]] of [
      ["School:", "school", "B1"], ["District:", "district", "T1"], ["Assessment:", "assessment", "B6"],
      ["Period:", "period", "B7"], ["Date of Assessment:", "date", "B8"], ["Teacher:", "teacher", "Q6"],
      ["Subject / Grade(s):", "subject", "Q7"]
    ].entries()) {
      add(i + 1, 1, 1, label, null, "A6"); add(i + 1, 2, score - 1, null, field, style);
    }
    add(8, 1, 1, "# of Questions:", null, "A8"); add(8, 2, 1, null, "questions", "T8");
    add(8, 3, score - 3, "# of Students:", null, "U8"); add(8, score, 1, null, "students", "Y8");
  }
  return { blocks, cells, score, scoreHeaderRow: questions >= 6 ? 6 : 9 };
}

export function analysisFormulas(questions, students, cells) {
  const layout = subjectLayout(questions, students);
  const { last, footer, total, rank, group } = layout;
  const totalL = column(total), rankL = column(rank), groupL = column(group), lastQ = column(questions + 2);
  const absolute = address => address.replace(/^([A-Z]+)(\d+)$/, (_, c, r) => `$${c}$${r}`);
  const count = absolute(cells.students), items = absolute(cells.questions), size = `$${groupL}$7`;
  const names = `$A$10:$A$${last}`, totals = `$${totalL}$10:$${totalL}$${last}`, groups = `$${groupL}$10:$${groupL}$${last}`;
  const scores = [], ranks = [], itemRows = Array.from({ length: 6 }, () => []);
  for (let r = 10; r <= last; r++) {
    scores.push([`=IF($A${r}="","",COUNTIF(C${r}:${lastQ}${r},"${CHECK}"))`,
      `=IF(OR($A${r}="",${items}=0),"",${totalL}${r}/${items})`]);
    ranks.push([`=IF($A${r}="","",COUNTIFS(${totals},">"&${totalL}${r},${names},"<>")+COUNTIFS(${totals},${totalL}${r},$B$10:$B$${last},"<="&B${r},${names},"<>"))`,
      `=IF(OR($A${r}="",${size}=0),"",IF(${rankL}${r}<=${size},"Upper",IF(${rankL}${r}>${count}-${size},"Lower","Middle")))`]);
  }
  for (let i = 0; i < questions; i++) {
    const c = column(i + 3), answers = `${c}$10:${c}$${last}`, d = `${c}${footer + 3}`;
    const upper = `COUNTIFS(${groups},"Upper",${answers},"${CHECK}")`;
    const lower = `COUNTIFS(${groups},"Lower",${answers},"${CHECK}")`;
    itemRows[0].push(`=COUNTIFS(${names},"<>",${answers},"${CHECK}")`);
    itemRows[1].push(`=IF(${count}=0,"",${c}${footer}/${count})`);
    itemRows[2].push(`=IF(${size}=0,"",(${upper}+${lower})/(2*${size}))`);
    itemRows[3].push(`=IF(${size}=0,"",ROUND((${upper}-${lower})/${size},10))`);
    itemRows[4].push(`=IF(ISNUMBER(${d}),IF(${d}<0.2,"PI",IF(${d}<0.3,"NI",IF(${d}<0.4,"GI","VG"))),"")`);
    itemRows[5].push(`=IF(ISNUMBER(${d}),IF(${d}<0.2,"Review",IF(${d}<0.3,"Revise","Retain")),"Review")`);
  }
  return { ...layout, scores, ranks, itemRows,
    groupSize: `=IF(${count}<2,0,MAX(1,MIN(INT(${count}/2),ROUND(${count}*$${groupL}$6,0))))`,
    mps: `=IF(${count}=0,"",AVERAGE(C${footer + 1}:${lastQ}${footer + 1}))` };
}
