import { buildManifest } from "./manifest-builder.mjs?v=0.2.0";
import { prepareWorkbook, generateSubject, gradePaper, applyStudentView, registerEvents } from "./workbook.mjs?v=0.2.0";

const status = document.getElementById("status");
const connection = document.getElementById("connection");
const errorBox = document.getElementById("error-details");
const errorText = document.getElementById("error");
const buttons = [...document.querySelectorAll("[data-action]")];
let ready = false, busy = false;
const base = new URL("./", window.location.href).href;
document.getElementById("page-address").textContent = base;

function message(value, state = "") {
  status.textContent = value;
  status.dataset.state = state;
}

function enabled() { for (const button of buttons) button.disabled = !ready || busy; }

function failed(error) {
  message(error.message || String(error), "failed");
  errorText.textContent = [error.code, error.debugInfo?.errorLocation, error.message || String(error)].filter(Boolean).join("\n");
  errorBox.hidden = false;
}

async function perform(label, action) {
  if (!ready || busy) return;
  busy = true; enabled(); errorBox.hidden = true;
  message(label);
  try { message(await action(), "passed"); } catch (error) { failed(error); }
  finally { busy = false; enabled(); }
}

document.getElementById("prepare").addEventListener("click", () => perform("Checking Original and Generate…", prepareWorkbook));
document.getElementById("generate").addEventListener("click", () => perform("Reading the roster and generating the subject…", generateSubject));
document.getElementById("grade").addEventListener("click", () => perform("Checking paper keys and answers…", gradePaper));
document.getElementById("view").addEventListener("click", () => perform("Applying the student selection…", applyStudentView));

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
  document.getElementById("download-help").textContent = "This XML opens prototype 0.2 at this HTTPS address.";
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
    try {
      await registerEvents(value => { if (!busy) message(value, "failed"); });
    } catch (error) {
      message("Ready. Automatic edit handling could not start; Apply Student View is still available. " + (error.message || error), "failed");
    }
  }).catch(error => { clearTimeout(timer); failed(error); });
}
