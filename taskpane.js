/* global Office, Excel */

// Choose the cell that should trigger the pane.
// You can change this to a named range, e.g. { name: "OpenPaneCell" }.
const TARGET = { sheet: "Overview", address: "B2" };

let selectionHookAdded = false;

// Map the manifest's FunctionName to this function (event-based activation).
Office.actions.associate("onDocumentOpen", onDocumentOpen);

Office.onReady(() => {
  const s = document.getElementById("status");
  if (s) s.textContent = `Watching for ${TARGET.sheet}!${TARGET.address}...`;
});

// Runs when the workbook opens (Excel on the web).
async function onDocumentOpen(event) {
  try {
    await ensureSelectionWatcher();
  } catch (err) {
    console.error(err);
  } finally {
    // REQUIRED for event-based activation.
    event.completed();
  }
}

async function ensureSelectionWatcher() {
  if (selectionHookAdded) return;
  await Excel.run(async (context) => {
    const context.workbook.worksheets.getItem(TARGET.sheet);
    sheet.onSelectionChanged.add(handleSelectionChanged);
    await context.sync();
  });
  selectionHookAdded = true;
}

// Fired whenever the selection changes anywhere in the workbook.
async function handleSelectionChanged(args) {
  try {
    await Excel.run(async (context) => {
      const wb = context.workbook;
      const s = document.getElementById("status");
      if (s) s.textContent = "Handling Selection Changed";
      // Get current selection and the target
      const sel = wb.getSelectedRange();
      const targetSheet = wb.worksheets.getItem(TARGET.sheet);
      const target = targetSheet.getRange(TARGET.address);
      const activeSheet = wb.worksheets.getActiveWorksheet();

      target.load("address");
      activeSheet.load("name");
      targetSheet.load("name");
      await context.sync();

      // Only compare ranges when you're on the right sheet
      if (activeSheet.name !== targetSheet.name) return;

      // Intersect to see if the selected cell overlaps the target
      const hit = target.getIntersectionOrNullObject(sel);
      hit.load("address"); // will be null object if no hit
      await context.sync();

      if (!hit.isNullObject) {
        // Open the task pane (idempotent).
        await Office.addin.showAsTaskpane();
      }
    });
  } catch (e) {
    console.error("Selection handler failed", e);
  }
}
