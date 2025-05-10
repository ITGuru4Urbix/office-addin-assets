/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // Office.js is ready to be called if needed.
});

/**
 * Opens the Urbix IT Support Portal in a new browser window.
 * @param event {Office.AddinCommands.Event}
 */
function openSupportPortal(event) {
  Office.context.ui.openBrowserWindow("https://urbixinc.sysaidit.com/servicePortal?openedInIframe=1");
  event.completed();
}

// Register the function with Office.
Office.actions.associate("openSupportPortal", openSupportPortal);
