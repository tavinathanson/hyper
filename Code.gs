/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
        .addToUi();
  replaceText();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
}

function replaceText() {
    var body = DocumentApp.getActiveDocument().getBody();
    body.replaceText("{{{replace_me}}}", "Replaced!");
}
