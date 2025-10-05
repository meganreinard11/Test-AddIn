/* global console, document, Excel, Office */
/* global Office, Excel */

// Choose the cell that should trigger the pane.
// You can change this to a named range, e.g. { name: "OpenPaneCell" }.
const TARGET = { sheet: "Overview", address: "B2", rowIndex: 1, columnIndex: 1 };

// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    try {
      let sheet = context.workbook.worksheets.getItem(TARGET.sheet);
      sheet.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();
      console.log("A handler has been registered for the onChanged event.");
    } catch() {
      console.log("Addition of the SelectionChanged handler failed");
    }
  });
};

// Fired whenever the selection changes anywhere in the workbook.
async function handleSelectionChanged(event) {
  try {
    await Excel.run(async (context) => {    
      await context.sync();
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  } catch () {
    console.log("Selection handler failed");
  }
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign event handlers and other initialization logic.
    document.getElementById("open-dialog").onclick = (() => tryCatch(openDialog));
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
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
  } catch (err) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(err);
  }
}
