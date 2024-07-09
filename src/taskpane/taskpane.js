import { Excel } from "@microsoft/office-js";

// Ensure Office.js is fully loaded and initialize the task pane add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Office.js is ready. Host: " + info.host);

    // Example: Call your function to insert text into Excel
    document.addEventListener("DOMContentLoaded", () => {
      insertText("Hello, Excel!");
    });
  } else {
    console.log("Office.js is ready, but this add-in is not supported in the current host: " + info.host);
  }
});

// Function to insert text into the Excel sheet
export async function insertText(text) {
  // Write text to the top left cell.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
  }
}
