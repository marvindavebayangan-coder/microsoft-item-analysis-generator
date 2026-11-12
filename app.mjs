import { buildManifest } from "./manifest-builder.mjs";

const download = document.getElementById("download");
const run = document.getElementById("run");
const status = document.getElementById("status");
const connection = document.getElementById("connection");
const errorDetails = document.getElementById("error");
const baseAddress = new URL("./", window.location.href).href;
document.getElementById("page-address").textContent = baseAddress;

function setStatus(message, state = "") {
  status.textContent = message;
  status.dataset.state = state;
}

try {
  const manifest = buildManifest(baseAddress);
  download.disabled = false;
  document.getElementById("download-help").textContent = "The XML file will use this page’s actual HTTPS address.";
  download.addEventListener("click", () => {
    const url = URL.createObjectURL(new Blob([manifest], { type: "application/xml;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "TIA_Hosting_Test.xml";
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
} catch (error) {
  document.getElementById("download-help").textContent = "Host these files at an HTTPS address before downloading the add-in XML.";
}

let excelReady = false;
const connectionTimer = setTimeout(() => {
  if (!excelReady) {
    connection.textContent = "Open the add-in from Excel to connect";
    setStatus("This page has loaded. To finish the workbook test, upload its XML file in Excel and open TIA Hosting Test.");
  }
}, 7000);

if (typeof Office === "undefined") {
  clearTimeout(connectionTimer);
  connection.textContent = "Office library unavailable";
  setStatus("The Microsoft Office library did not load. Check access to officeapis.public.onecdn.static.microsoft, then reload.", "failed");
} else {
  Office.onReady().then(info => {
    clearTimeout(connectionTimer);
    if (info.host !== Office.HostType.Excel) {
      connection.textContent = "Open the add-in from Excel to connect";
      setStatus("The page is ready. Download its XML, upload it in Excel, and open TIA Hosting Test.");
      return;
    }
    excelReady = true;
    connection.textContent = "Connected to Excel";
    run.disabled = false;
    setStatus("Connected. Use a blank workbook, then select Run Excel test.");
  }).catch(error => {
    clearTimeout(connectionTimer);
    connection.textContent = "Connection failed";
    setStatus("Excel initialization failed.", "failed");
    errorDetails.hidden = false;
    errorDetails.textContent = error.message || String(error);
  });
}

run.addEventListener("click", async () => {
  if (!excelReady) return;
  run.disabled = true;
  errorDetails.hidden = true;
  errorDetails.textContent = "";
  setStatus("Writing a test message and reading it back…");
  try {
    const actual = await Excel.run(async context => {
      const cell = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      cell.values = [["Hosted add-in works!"]];
      cell.format.fill.color = "#E2F0D9";
      cell.format.font.color = "#176443";
      cell.format.font.bold = true;
      await context.sync();
      cell.load("values");
      await context.sync();
      return cell.values[0][0];
    });
    if (actual !== "Hosted add-in works!") throw new Error("The value read from Excel did not match the test message.");
    setStatus("Passed: Excel saved the message in cell A1 and returned the same value.", "passed");
  } catch (error) {
    setStatus("The workbook test failed. Please report the message below.", "failed");
    errorDetails.hidden = false;
    errorDetails.textContent = [error.code, error.message || String(error)].filter(Boolean).join("\n");
  } finally {
    run.disabled = false;
  }
});
