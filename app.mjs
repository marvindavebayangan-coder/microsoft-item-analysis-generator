import { buildManifest } from "./manifest-builder.mjs?v=0.4.0";
import { prepareWorkbook, generateSubject, gradePaper, applyStudentView, sortParticipantsByRank,
  sortItemsByPercentCorrect, restoreOriginalItemOrder, registerEvents,
  listQuizSubjects, connectQuiz, importQuizResponses, disconnectQuiz } from "./workbook.mjs?v=0.4.0";
import { readFormsXlsx } from "./xlsx-reader.mjs?v=0.4.0";
import { formsDataset, quizUrl } from "./forms.mjs?v=0.4.0";

const status = document.getElementById("status");
const connection = document.getElementById("connection");
const errorBox = document.getElementById("error-details");
const errorText = document.getElementById("error");
const buttons = [...document.querySelectorAll("[data-action]")];
let ready = false, busy = false;
let quizSubjects = [], quizData = null;
const subjectChoice = document.getElementById("quiz-subject"), nameChoice = document.getElementById("quiz-name");
const urlInput = document.getElementById("quiz-url"), fileInput = document.getElementById("quiz-file");
const quizSelected = () => quizSubjects.find(subject => subject.id === subjectChoice.value);
const base = new URL("./", window.location.href).href;
document.getElementById("page-address").textContent = base;

function message(value, state = "") {
  status.textContent = value;
  status.dataset.state = state;
}

function enabled() {
  for (const button of buttons) button.disabled = !ready || busy;
  const selected = quizSelected(), connected = !!selected?.connected;
  subjectChoice.disabled = !ready || busy || !quizSubjects.length;
  urlInput.disabled = !ready || busy || connected;
  fileInput.disabled = !ready || busy;
  nameChoice.disabled = !ready || busy || connected || !quizData;
  document.getElementById("quiz-connect").disabled = !ready || busy || connected || !selected || !quizData;
  document.getElementById("quiz-import").disabled = !ready || busy || !connected || !quizData;
  document.getElementById("quiz-disconnect").hidden = !connected;
  document.getElementById("quiz-disconnect").disabled = !ready || busy || !connected;
}

function quizPreview() {
  document.getElementById("quiz-preview").hidden = !quizData;
  if (!quizData) return;
  const items = quizData.blocks.filter(block => block.answer !== Number(nameChoice.value));
  document.getElementById("quiz-summary").textContent = `${quizData.rows.length} response(s), ${items.length} exam items. The selected name question is excluded.`;
  const list = document.getElementById("quiz-items"); list.replaceChildren();
  for (const item of items) { const li = document.createElement("li"); li.textContent = item.title; list.append(li); }
}

function quizState() {
  const subject = quizSelected(), link = document.getElementById("quiz-link");
  document.getElementById("quiz-state").textContent = !subject ? "Choose a generated subject." : subject.connected ?
    `Connected for file imports. ${subject.questions} items; names from ${subject.nameHeader}. ${subject.imported} response ID(s) imported.` :
    `${subject.questions} items. No active quiz connection.${subject.imported ? " Previous imported marks are kept." : ""}`;
  link.hidden = true;
  if (subject?.connected) {
    urlInput.value = subject.url;
    try { link.href = quizUrl(subject.url); link.hidden = false; } catch (_) { /* Keep invalid stored links inactive. */ }
  }
  if (quizData && subject?.connected) nameChoice.value = String(subject.nameColumn);
  quizPreview(); enabled();
}

async function refreshQuizSubjects() {
  if (!ready) return;
  const previous = subjectChoice.value, catalog = await listQuizSubjects();
  quizSubjects = catalog.subjects;
  subjectChoice.replaceChildren();
  const empty = document.createElement("option"); empty.value = ""; empty.textContent = "Choose a generated subject"; subjectChoice.append(empty);
  for (const subject of quizSubjects) { const option = document.createElement("option"); option.value = subject.id; option.textContent = subject.name; subjectChoice.append(option); }
  subjectChoice.value = quizSubjects.some(subject => subject.id === previous) ? previous : catalog.active;
  quizState();
}

function failed(error) {
  message(error.message || String(error), "failed");
  errorText.textContent = [error.stage && "Step: " + error.stage, error.code,
    error.debugInfo?.errorLocation, error.message || String(error)].filter(Boolean).join("\n");
  errorBox.hidden = false;
  errorBox.open = true;
  errorBox.scrollIntoView({ block: "nearest" });
}

async function perform(label, action) {
  if (!ready || busy) return;
  busy = true; enabled(); errorBox.hidden = true;
  message(label);
  try { message(await action(), "passed"); } catch (error) { failed(error); }
  finally {
    try { await refreshQuizSubjects(); } catch (_) { /* Keep the operation result if this read fails. */ }
    busy = false; enabled();
  }
}

document.getElementById("prepare").addEventListener("click", () => perform("Checking Original and Generate…", prepareWorkbook));
document.getElementById("generate").addEventListener("click", () => perform("Reading the roster and generating the subject…", generateSubject));
document.getElementById("grade").addEventListener("click", () => perform("Checking paper keys and answers…", gradePaper));
document.getElementById("view").addEventListener("click", () => perform("Applying the student selection…", applyStudentView));
document.getElementById("sort-participants").addEventListener("click", () => perform("Sorting participants by rank…", sortParticipantsByRank));
document.getElementById("sort-items").addEventListener("click", () => perform("Sorting items by % correct…", sortItemsByPercentCorrect));
document.getElementById("restore-items").addEventListener("click", () => perform("Restoring original item order…", restoreOriginalItemOrder));
subjectChoice.addEventListener("change", () => { urlInput.value = quizSelected()?.url || ""; quizState(); });
nameChoice.addEventListener("change", quizPreview);
fileInput.addEventListener("change", () => perform("Reading the Forms Excel results…", async () => {
  const file = fileInput.files[0]; quizData = null; quizPreview();
  if (!file || !/\.xlsx$/i.test(file.name) || file.size > 10 * 1024 * 1024) throw new Error("Choose the Forms .xlsx results export, up to 10 MB.");
  quizData = await formsDataset(await readFormsXlsx(await file.arrayBuffer()));
  nameChoice.replaceChildren();
  quizData.headers.forEach((header, index) => { const option = document.createElement("option"); option.value = String(index); option.textContent = `${index + 1}. ${header}`; nameChoice.append(option); });
  nameChoice.value = String(quizSelected()?.connected ? quizSelected().nameColumn : quizData.nameColumn);
  quizPreview();
  return "Results loaded. Check the selected subject, student-name column and item mapping. Connect Quiz saves the mapping; Import Responses applies the results.";
}));
document.getElementById("quiz-connect").addEventListener("click", () => perform("Connecting the selected quiz and subject…", () =>
  connectQuiz(subjectChoice.value, { url: urlInput.value, data: quizData, nameColumn: Number(nameChoice.value) })));
document.getElementById("quiz-import").addEventListener("click", () => perform("Matching Forms responses to the subject…", () => importQuizResponses(subjectChoice.value, quizData)));
document.getElementById("quiz-disconnect").addEventListener("click", () => perform("Disconnecting the quiz and keeping its marks…", () => disconnectQuiz(subjectChoice.value)));

try {
  const xml = buildManifest(base);
  const download = document.getElementById("download");
  download.disabled = false;
  download.addEventListener("click", () => {
    const url = URL.createObjectURL(new Blob([xml], { type: "application/xml;charset=utf-8" }));
    const link = document.createElement("a");
    link.href = url; link.download = "Microsoft_Item_Analysis_Generator.xml";
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
  document.getElementById("download-help").textContent = "This XML opens prototype 0.4.0 at this HTTPS address.";
} catch (error) {
  document.getElementById("download-help").textContent = "Publish these files at an HTTPS address to download the XML.";
}

const timer = setTimeout(() => {
  if (!ready) {
    document.getElementById("install").open = true;
    connection.textContent = "Open this add-in in Excel";
    message("Download its XML, then open the add-in in your workbook containing Original and Generate.");
  }
}, 7000);

if (typeof Office === "undefined") {
  clearTimeout(timer); connection.textContent = "Office library unavailable";
  message("The Microsoft Office library did not load. Reload the page and check the connection.", "failed");
} else {
  Office.onReady().then(async info => {
    clearTimeout(timer);
    if (info.host !== Office.HostType.Excel) {
      document.getElementById("install").open = true;
      connection.textContent = "Open this add-in in Excel";
      message("Download the XML and open this add-in in your template workbook in Excel Online.");
      return;
    }
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      connection.textContent = "Excel client needs an update";
      message("This prototype requires ExcelApi 1.9. Open the workbook in current Excel Online.", "failed");
      return;
    }
    ready = true; enabled(); connection.textContent = "Connected to Excel"; connection.dataset.ready = "true";
    message("Ready. Open your workbook with Original and Generate, then select Check Template.");
    await refreshQuizSubjects();
    try {
      await registerEvents(value => { if (!busy) message(value, "failed"); });
    } catch (error) {
      message("Ready. Automatic edit handling could not start; Apply Student View is still available. " + (error.message || error), "failed");
    }
  }).catch(error => { clearTimeout(timer); failed(error); });
}
