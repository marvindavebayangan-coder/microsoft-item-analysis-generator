# Microsoft Item Analysis Generator — prototype 0.1

An Excel Online Office.js add-in for the A1 and A3 accounts tested in this project. Program files are hosted on GitHub Pages; workbook contents remain in the teacher's workbook. This version does not connect to Microsoft Forms or send workbook records to a backend.

## Included in this prototype

- Prepare Generate and a hidden, 100-question / 100-student Original template in a new workbook.
- Read the familiar Generate!B5:B12 inputs and the names below Student Names here in column A.
- Ignore blank names, reject duplicates, and enforce 1–100 questions and 1–100 named students.
- Copy Original for each subject. Full titles are retained in the workbook header; Excel tab names are shortened to 31 characters and made unique.
- Create a Manual Checking tab only for names with plain red fill (#FF0000).
- Compare A–D paper answers with native conditional formatting: matching answers light green, different answers light red, blanks neutral. Grading ignores case and surrounding spaces.
- Alphabetical student selector with Show all; freeze Manual Checking rows 1–3 and columns A–B.
- Transfer graded paper answers by student name and permanent item number.
- Generate native Excel score, percentage, difficulty, discrimination, description, rank, group, and decision formulas following the attached Google Version 14.
- Delete surplus template question slots in the generated copy and hide columns beyond the report. Excel retains its fixed 16,384-column grid.
- Convert typed P/p/1 to ✓ while the add-in is open.
- Preserve the template and existing subject results when preparing an already prepared workbook.

Sorting commands, Microsoft Forms response collection, quiz connection management, and update/migration commands are later milestones.

## Update the hosted files

1. Extract Microsoft_TIA_Generator_Prototype_01.zip.
2. Open its web folder.
3. In the GitHub repository, choose Add file > Upload files.
4. Upload all files inside web directly to the repository root, replacing matching existing files.
5. Commit changes to main and wait for the GitHub Pages deployment to finish.

GitHub Pages is already configured to publish main, /(root).

## Install this version in Excel

1. Open the published page with ?v=0.1.0 at the end of its address.
2. Under Install or update the add-in, download Microsoft_Item_Analysis_Generator.xml.
3. Open a blank Excel Online workbook with the A3 account.
4. Use Home > Add-ins > Advanced > Upload My Add-in to upload the new XML.
5. Open Item Analysis Generator. It uses the same add-in ID as the hosting test, with manifest version 1.1.0.0.
6. Select Prepare Sample Workbook.
7. Select Generate Subject.

The existing hosting-test installation points to the same page, but the updated XML changes its displayed name and pins the current prototype page.

## First generation check

The sample has 5 questions, 6 student names, and plain red fills on Student 03 and Student 06.

Expected results:

- Generate remains available.
- Original exists and is hidden.
- Sample Math is immediately after Generate.
- Sample Math - Manual Checking follows the subject.
- The subject contains exactly 5 item columns, followed by totals, score, rank and group.
- The paper tab contains Student 03 and Student 06.
- Paper tab B1 lists Show all and those two names.
- Paper rows 1–3 and columns A–B are frozen.

Use the B1 dropdown while the add-in is open. Apply Student View can also refresh the selected view.

## Optional paper check after generation succeeds

Enter this small fixture on Manual Checking:

| Item | Key | Student 03 | Student 06 |
|---|---|---|---|
| 1 | A | A | B |
| 2 | B | B | B |
| 3 | C | D | C |
| 4 | D | D | A |
| 5 | A | leave blank | A |

Click Grade Paper Answers from the paper or subject tab.

Each paper student should have 3 correct answers and a score of 60%. The remaining four listed students are unmarked, so the whole-roster MPS is provisionally 20%. Finish all marking/submissions before interpreting the item decisions.

## Analysis rules retained from V14

- Overall % correct uses all named roster students.
- Difficulty uses upper/lower groups.
- Default group proportion is 33%, rounded and capped at half the class.
- Ties use the original student number.
- Description: D < 0.20 = PI; D < 0.30 = NI; D < 0.40 = GI; otherwise VG.
- Decision: D < 0.20 = Review; 0.20 ≤ D < 0.30 = Revise; D ≥ 0.30 = Retain.
- Missing group data yields Review; one student has no upper/lower group statistics.
- All-correct and all-wrong items have discrimination 0 and therefore Review.
- Upper answer rows are yellow; Lower answer rows are green; Middle rows retain white.

## Validation performed

Local checks cover input limits, plain-red detection, duplicate normalization, Excel tab names, item-number mapping, invalid A–D data, a 100-by-100 paper matrix, literal-text handling, JavaScript syntax, and XML resource references.

These checks do not establish native Excel runtime behavior. Generation, Excel formula recalculation, pane handling, and conditional formatting still require the A3 and A1 workbook tests. The earlier hosting test established only basic hosted add-in read/write access.

## Reference

The supplied Google Apps Script Version 14 is the functional reference. The older worksheet guide's Reject wording is not used for the prototype's decisions.
