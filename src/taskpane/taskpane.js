/* global Excel console */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    console.log("Office.js is ready. Host: " + info.host);
    // Example: Call your function to insert text into Excel
    insertText("Hello, Excel!");
  } else {
    console.log("Office.js is ready, but this add-in is not supported in the current host: " + info.host);
  }
});

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
