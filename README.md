# Microsoft Item Analysis Generator — prototype 0.2

An Excel Online Office.js add-in for the A1 and A3 accounts used in this project. Program files are hosted on GitHub Pages. Class records stay in the teacher's workbook; this version has no backend or Microsoft Forms connection.

## This update

Open the workbook you prepared with **Original** and **Generate**. The add-in uses those sheets directly; it no longer creates a replacement template or requires a blank workbook.

- Detects Original's question columns and student area from its headings. The supplied template has 40 questions and 11 student rows.
- Copies Original, clears inherited sample answers and old analysis formulas in that copy, and expands or removes item columns and student rows for the actual exam.
- Preserves the template's name column, student row heights, item heading height, fonts, borders, and rotated headings. Metadata merges adjust to the requested item count. Very short quizzes use one field per row.
- Uses a minimum item width of 44 screen pixels, as the V14 analysis does, so decimal results and decision words fit. Correct Answers and Student Score retain narrow columns with vertical headings.
- Adds Description and Decision without overwriting interpretation notes. Existing notes and the group legend move with the resized student area.
- Reads Generate!B5:B12 and all names below Student Names here in column A. Blank names are ignored; the limits remain 100 questions and 100 students.
- Pure red name fills (#FF0000) create a linked Manual Checking sheet. A–D grading, answer colors, the student selector, and frozen panes remain available.
- Rebuilds score, percentage, upper/lower group, difficulty, discrimination, description, rank, and decision formulas using the named roster and the V14 rules. Ties retain the original student order.
- On **Check Template**, refreshes only the identifiable legacy GUIDE example: keeps its response counts, provides an editable Example N large enough for those counts, and calculates the current descriptions and decisions. Example N is separate from a subject's actual groups.
- Leaves Original, Generate, and existing subject results intact. A failed generation removes only the new sheets it created.

The green worksheet drawing labeled Click here to Generate has no assigned add-in action. Use Generate Subject in the task pane.

## Update GitHub Pages

1. Extract Microsoft_TIA_Template_Update_02.zip.
2. Open its web folder.
3. In the existing repository, select Add file > Upload files.
4. Upload **all files inside web** directly to the repository root and replace matching files. Include the new layout.mjs.
5. Commit to main and wait for the Pages deployment to complete.

Pages is already configured for main, /(root). This ZIP contains the add-in files; keep using your own workbook in OneDrive.

## Open the updated add-in

1. Open https://marvindavebayangan-coder.github.io/microsoft-item-analysis-generator/?v=0.2.0
2. Use the XML included at the root of the ZIP, or expand Install or update the add-in on the page to download it.
3. Open your TIA Prototype workbook in Excel Online.
4. Close the old task pane and upload the new XML through Home > Add-ins > Advanced > Upload My Add-in.
5. Confirm the pane says Prototype 0.2, then select Check Template.

The add-in ID is unchanged. The manifest version is 1.2.0.0, and every application module uses the 0.2.0 cache key.

## First Excel test

Use the existing sample: 20 questions in B12 and ten names in A15:A24.

1. Check Template should report 40 item slots, 11 template student rows, 20 requested questions, and ten students.
2. Select Generate Subject. The new Math sheet should have students on rows 10–19, items in C:V, Correct Answers in W, and Student Score in X.
3. The six analysis rows should start at row 20. Interpretation and the color legend should remain below the decision notes.
4. Check that student rows are visible, old sample P answers are absent, and item and score headings are vertical.
5. Enter P or 1 in a few answer cells while the pane is open. They should become ✓ and recalculate the scores.
6. To test mixed mode, change the subject title to Math Mixed and fill two roster names with plain red. Generate again, enter an A–D key and some answers in Manual Checking, then run Grade Paper Answers.
7. Repeat in the A1 account. After the small example passes, test 100 questions and 100 names.

No red names produce a subject sheet only. All-red names produce a paper sheet for the entire roster. Blank paper answer columns are skipped when grading.

## Validation and limits

Run portable regression checks with:

    node --test tests/*.test.mjs

The tests cover template detection, shrinking and expansion, merged headings, preservation of the master and notes, clearing hidden sample data, mixed paper grading, item order, selection, formula dependency cycles, duplicate protection, and failed-generation cleanup.

The Office adapter used by tests models structural operations; it does not replace Excel Online. Local calculation checks also compare generated formulas with independent expected results. Visual appearance and native behavior in the tenant still need the Excel test above on A3 and A1.

Sorting commands, Microsoft Forms response collection, quiz connection management, and migration of older generated subject layouts are subsequent milestones. Existing registered paper sheets remain usable; this update applies the template layout to newly generated subjects.

## Architecture

- index.html, styles.css, app.mjs: task pane and command handling.
- domain.mjs: input validation, identifiers, paper grading, and student/item mapping.
- layout.mjs: template detection, header positions, and analysis formula construction.
- workbook.mjs: Office.js worksheet copying, resizing, formatting, registry, grading, and edit events.
- manifest-builder.mjs: manifest with stable app identity and HTTPS source.
- tests: portable domain and structural regression checks.

Workbook settings use the existing TIA.Generator.Registry.v1 key to retain subject and paper sheet IDs. This code uses ExcelApi 1.9 features and does not depend on Office Scripts.

The uploaded workbook is the visual source. Google Apps Script V14 remains the functional source for the scoring and decision rules. The original workbook's obsolete calculations are replaced in each generated copy rather than trusted as the scoring engine.

API references used for this update:

- https://learn.microsoft.com/en-us/javascript/api/excel/excel.range
- https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeformat
- https://learn.microsoft.com/en-us/javascript/api/excel/excel.cellpropertiesformatloadoptions
- https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-conditional-formatting
