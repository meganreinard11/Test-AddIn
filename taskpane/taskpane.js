/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign event handlers and other initialization logic.
    document.getElementById("open-dialog").onclick = (() => tryCatch(openDialog));
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

let dialog = null;

/**
 * This function opens a dialog that uses popup.html.
 */
function openDialog() {
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html',
    { height: 45, width: 55 },

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
