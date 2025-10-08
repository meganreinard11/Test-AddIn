// Choose the cell that should trigger the pane.
const TARGET = { sheet: "Overview", address: "B2" };

Office.actions.associate("showTaskpane", showTaskpane);
Office.actions.associate("hideTaskpane", hideTaskpane);

// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  Excel.run(async context => {
      let sheet = context.workbook.worksheets.getItem(TARGET.sheet);
      sheet.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();
      console.log("A handler has been registered for the onSelectionChanged event.");
  }).catch(window.ValidationManager.handleError);
};

let isTaskpaneOpen = false;
// Fired whenever the selection changes anywhere in the workbook.
async function handleSelectionChanged(event) {
  await Excel.run(async (context) => {
    if (event.address !== TARGET.address) return;
    await context.sync();
  }).catch(window.ValidationManager.handleError);
};

function showPane() {
  if (isTaskpaneOpen) return;
  Office.addin.showAsTaskpane()
    .then(function() {
      isTaskpaneOpen = true;
    });
}

function hidePane() {
  if (!isTaskpaneOpen) return;
  Office.addin.hide()
    .then(function() {
      isTaskpaneOpen = false;
    });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
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
