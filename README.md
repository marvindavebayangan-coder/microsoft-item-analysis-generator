# Microsoft Item Analysis Generator — prototype 0.4.0

An Excel Online Office.js add-in targeting the A1 and A3 accounts used in this project. Program files are hosted on GitHub Pages. This version reads downloaded Microsoft Forms quiz results in the task pane and applies item marks to the teacher's workbook. It has no backend, Forms sign-in integration, or automatic response collection. The hosted connection was tested by the user on both licenses; the new file-import workflow needs the native A3 check below, followed by A1.

This update adds **Connect Quiz**, **Import Responses**, and **Disconnect Quiz**, preserving generation, paper grading, sorting and the 0.3.1 label repair. Existing template-generated subjects remain registered. Use the 20-item Forms Test subject for the first import.

## Microsoft Forms file-import workflow

1. In the Microsoft Forms quiz section, choose the target TIA subject.
2. Paste the quiz link from Collect responses. This is an association and a link to open the quiz; it does not authorize access to Forms responses.
3. Choose the downloaded .xlsx results file from that same quiz. Files are read locally in the task pane, not uploaded to GitHub.
4. Check the student-name column and Review item mapping. In the supplied export, Name1 is the roster dropdown; Name is the account name. The Name1/Points/Feedback block is excluded from the 20 exam items.
5. Choose Connect Quiz. This stores the quiz association and column mapping without changing marks.
6. Choose Import Responses. The importer maps names to current roster rows and permanent item IDs to current item columns, including after sorting.
7. To refresh later, download a fresh results file, choose it in the pane, and run Import Responses again. Connecting alone does not fetch new submissions.
8. Disconnect Quiz stops further imports and keeps existing marks and response history. It does not delete or close the Microsoft Forms quiz. The button is shown only for a connected subject. Reconnect the same quiz/layout to resume; use a fresh subject for a different exam once results have been imported.

The initial supported export has English Forms headings, one-point automatically graded questions, and a zero-point name question. Every included per-item score must be numeric 0 or 1 and their sum must equal Total points. Additional ungraded metadata questions, partial credit and weighted scoring need an expanded mapping design. Item order in the source export establishes permanent item IDs; the file does not expose a durable quiz identity for verifying that a manually selected file belongs to the pasted URL. Select the correct export and review its mapping. Changed column layouts stop imports rather than silently remapping existing results.

Response IDs update existing records rather than adding roster rows. Unmatched names, paper students and rows with incomplete grading are reported/skipped. Multiple different response IDs for the same matched student stop the import for review; this version does not silently pick an attempt. An updated grade for the same response ID is accepted after validation.

Only matched online students' item marks are written. A temporary range backup protects an import until both its cell values and history are saved. Failures attempt to restore marks and the previous history; if recovery fails, keep the visible TIA Import Backup tab and report the error details. Imported students remain protected from paper-grading writes after disconnection.

The browser reader supports ordinary stored/deflated .xlsx packages, shared and inline strings, and cached scalar values. It never executes source workbook formulas. Limits are 10 MB compressed, 1,000 response rows, and 400 columns, covering the 100-question/100-student target. It requires a current browser with native deflate-raw decompression; no extra Excel API requirement or Office Scripts access is introduced.

## Description and Decision repair

Version 0.3.0 left the calculated Description and Decision rows formatted as Text. Re-entering their formulas during item sorting could therefore store formula text instead of calculated labels. Version 0.3.1 sets an appropriate formula-compatible format before those writes, retains numeric formats for the other analysis rows, and verifies that formula cells remain formulas with the expected references.

For the affected existing subject, update the add-in, open the subject and choose **Restore Original Item Order** once. Before sorting, the add-in recognizes the generator's known Description/Decision expressions stored as A1 or R1C1 text and restores those calculated labels. This also works if the item columns are already in original order. Responses, paper keys, student ranks, and the numerical difficulty/discrimination calculations are not rebuilt by this repair. Intentional literal label overrides and arbitrary text beginning with = are kept.

The current classification rules are unchanged:

| Discrimination D | Description | Decision |
| --- | --- | --- |
| Below 0.20, including zero and negative values | PI | Review |
| 0.20 to below 0.30 | NI | Revise |
| 0.30 to below 0.40 | GI | Retain |
| 0.40 and above | VG | Retain |
| Group data unavailable | Blank | Review |

The screenshot's values such as 0.29, 0.43, zero, and negative discrimination can occur under these rules. Verifying every numerical D in a particular exam requires its full response grid and upper/lower groups; a cropped footer alone does not establish those counts.

## Sorting

- Participants: rank 1 first, followed by the remaining ranks. Original Student # values move with their names and answers. Score and rank formulas are restored for their new row positions. Blank-name rows move to the bottom.
- Items: highest to lowest % correct; ties use the original item number. Each item heading, student answer column, and its six analysis cells move together. Restore Original Item Order returns the columns to ascending item number.
- Cell formatting and annotations move with the sorted cells. Headers above the item table, score columns during item sorting, frozen panes, and the linked Manual Checking matrix retain their positions.
- Paper grading still matches the current subject rows by student name and columns by permanent item number after either sort. Sorting does not change the key or student selector in Manual Checking.
- Each change uses a temporary hidden range backup, verifies the recorded values after sorting, and removes that backup after success. If sorting fails, it attempts to restore the original table. If restoration fails, the backup is retained and made visible for recovery; keep it and report the error details. An already sorted table needs no backup.

Grade the recorded answers before sorting. These commands work from a registered subject tab or its linked Manual Checking tab and open the subject results when a sort completes. As answers change, ranks and percentages recalculate, but the displayed order changes only when you run a sorting command again.

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

1. Extract Microsoft_TIA_Forms_Update_04.zip.
2. Open its web folder.
3. In the existing repository, select Add file > Upload files.
4. Upload **all files inside web** directly to the repository root and replace matching files. Include the new **forms.mjs** and **xlsx-reader.mjs**.
5. Commit to main. Open Actions and wait for the latest Pages build and deployment to show **Success** with a green check.

Pages is already configured for main, /(root). This ZIP contains the add-in files; keep using your own workbook in OneDrive.

## Open the updated add-in

1. Open https://marvindavebayangan-coder.github.io/microsoft-item-analysis-generator/?v=0.4.0
2. Use the XML included at the root of the ZIP, or expand Install or update the add-in on the page to download it.
3. Open your TIA Prototype workbook in Excel Online.
4. Close the old task pane and upload the new XML through Home > Add-ins > Advanced > Upload My Add-in.
5. Confirm the pane says **Prototype 0.4.0** and has the **Microsoft Forms quiz** section.

The add-in ID is unchanged. The manifest version is 1.4.0.0, and every application module uses the 0.4.0 cache key.

## First native Forms import test

1. Choose Forms Test, with 20 items and Student 2 as an online student.
2. Paste the quiz link and select the supplied Untitled quiz.xlsx results file.
3. Confirm the name selector shows Name1, one response, and 20 exam items in the mapping.
4. Choose Connect Quiz, then Import Responses. Student 2 should have 12 checks, a score of 12/20, and 60%. The account name is not the roster match.
5. Import the same file again. It should report that the result is already current, with no duplicate roster row.
6. Disconnect Quiz. The marks should remain and Import Responses should become unavailable until reconnection.
7. Reconnect the same quiz and try a new export after another student's submission. Test A1 separately after this passes on A3.

## Next native Excel check

Use the existing Advanced Algebra subject with its 30-item, 20-student paper results. Saving a workbook copy first gives you a convenient comparison.

1. Open Advanced Algebra and choose Sort Participants by Rank. Analysis Rank should read 1–20 down the rows, with the highest score first. Student # values keep their original IDs rather than being renumbered.
2. Choose Sort Items by % Correct. Percentages should descend from left to right; equal percentages use ascending item number. A student's score should stay the same, and each item's results should stay beneath its numbered heading.
3. Choose Restore Original Item Order. The headings should return to 1–30. Participant rank order remains as sorted.
4. In Manual Checking, change a known answer and choose Grade Paper Answers. Verify that it updates the same named student and numbered item in the subject, including after another item sort.
5. Check that the upper/lower colors, rotated headings, and frozen panes still display correctly. Repeat the native check with A1 at the planned next session.

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

The tests cover template detection, shrinking and expansion, merged headings, preservation of the master and notes, clearing hidden sample data, mixed paper grading, selection, formula dependency cycles, duplicate protection, and failed-generation cleanup. Sorting coverage includes combined row and column sorts, stable names and item IDs, ties, extra remarks, formulas, notes, formatting, literal names beginning with =, paper regrading, 100 × 100 marks, rejected invalid data, verified rollback, and retained backups if restoration fails. The Text-format regression is modeled explicitly, including existing-label recovery, preservation of overrides, and detection of formulas changing to text or acquiring shifted references.

58 portable tests pass. New checks cover reading compressed .xlsx values, XML/CRC validation, Name1 mapping, duplicate imports, name/item sorting, paper protection, pending grading, repeated attempts, changed layouts, score reconciliation, disconnect/reconnect, failed writes/history commits, 100 × 100 imports and task-pane event/state handling. The Office adapter and DOM model do not replace Excel Online or a visual browser check. Native behavior of the new workflow still needs A3 and A1 testing. The 0.2.2 fix for copied conditional-format IDs is retained.

Automatic Microsoft Forms response collection and migration of older pre-template generated subject layouts are subsequent milestones. Existing registered template subjects and paper sheets remain usable.

## Architecture

- index.html, styles.css, app.mjs: task pane and command handling.
- domain.mjs: input validation, identifiers, paper grading, and student/item mapping.
- layout.mjs: template detection, header positions, and analysis formula construction.
- sorting.mjs: validated sort order, stable identifiers, formula masks, and result verification.
- forms.mjs: quiz associations, export mapping, score validation and response import planning.
- xlsx-reader.mjs: local browser ZIP/Open XML values reader for Forms exports.
- workbook.mjs: Office.js worksheet copying, resizing, formatting, registry, grading, and edit events.
- manifest-builder.mjs: manifest with stable app identity and HTTPS source.
- tests: portable domain and structural regression checks.

Workbook settings use the existing TIA.Generator.Registry.v1 key to retain subject and paper sheet IDs. This code uses ExcelApi 1.9 features and does not depend on Office Scripts.

The uploaded workbook is the visual source. Google Apps Script V14 remains the functional source for the scoring and decision rules. The original workbook's obsolete calculations are replaced in each generated copy rather than trusted as the scoring engine.

API references used for this update:

- https://learn.microsoft.com/en-us/javascript/api/excel/excel.range
- https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangesort
- https://learn.microsoft.com/en-us/javascript/api/excel/excel.sortfield
- https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeareas
- https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeformat
- https://learn.microsoft.com/en-us/javascript/api/excel/excel.cellpropertiesformatloadoptions
- https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-conditional-formatting
- https://support.microsoft.com/en-us/excel/how-to-avoid-broken-formulas-in-excel
