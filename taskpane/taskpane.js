/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
/* global Office, Excel */

// Choose the cell that should trigger the pane.
// You can change this to a named range, e.g. { name: "OpenPaneCell" }.
const TARGET = { sheet: "Overview", address: "B2", rowIndex: 1, columnIndex: 1 };

let selectionHookAdded = false;

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
    const sheet = context.workbook.worksheets.getItem(TARGET.sheet);
    sheet.onSelectionChanged.add(handleSelectionChanged);
    await context.sync();
  });
  selectionHookAdded = true;
}

// Fired whenever the selection changes anywhere in the workbook.
async function handleSelectionChanged(args) {
  try {
    if ((args.columnCount == 1) && (args.rowCount == 1)) {
      if ((rgs.startRow == TABLE.rowIndex) && (args.startColumn == TABLE.columnIndex)) {
        await Excel.run(async (context) => {
          const wb = context.workbook;
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
      }
    }  
  } catch (e) {
    console.error("Selection handler failed", e);
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign event handlers and other initialization logic.
    document.getElementById("open-dialog").onclick = (() => tryCatch(openDialog));
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    Office.context.document.bindings.getByIdAsync("Overview", function (result) {
          result.value.addHandlerAsync("bindingSelectionChanged", handleSelectionChanged);
    });
    await ensureSelectionWatcher();
  }
});

let dialog = null;
function openDialog() {
  Office.context.ui.displayDialogAsync(
    'https://meganreinard11.github.io/Test-AddIn/dialogs/popup.html',
    { height: 45, width: 55, displayInIframe: true },

    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  );
}

/**
 * This function writes the string provided by the dialog to the "user-name" element in the task pane.
 * @param arg The value returned from the dialog.
 */
function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }
}
