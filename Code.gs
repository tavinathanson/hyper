/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen(e) {
  replaceText();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
}

function replaceText() {
    var elements = getAllTextElementsWithUrls();
    for (var i = 0; i < elements.length; i++) {
        element = elements[i];
        url = element.getLinkUrl(0);
        if (url.indexOf("http://hammerlab.org/linker") == 0) {
            var key = url.split("http://hammerlab.org/linker/")[1].split("/http")[0]
            var realUrl = "http" + url.split("http://hammerlab.org/linker/")[1].split("/http")[1];
            var response = UrlFetchApp.fetch(realUrl).getContentText();
            var responseKey = "{{{" + key + ":";
            var responseKeyIndex = response.indexOf(responseKey);
            if (responseKeyIndex != -1) {
                responsePartial = response.substr(responseKeyIndex);
                responsePartialKeyEndIndex = responseKey.length;
                responsePartialValueEndIndex = responsePartial.indexOf("}}}");
                responseValue = responsePartial.substr(
                    responsePartialKeyEndIndex,
                    responsePartialValueEndIndex - responsePartialKeyEndIndex);

                element.replaceText(".*", responseValue);
            }
        }
    }
}

function getAllTextElementsWithUrls(node) {
    var elements = [];
    node = node || DocumentApp.getActiveDocument().getBody();
    if (node.getType() === DocumentApp.ElementType.TEXT) {
        var url = node.getLinkUrl(0);
        if (url != null) {
            elements.push(node);
        }
    }
    else {
        var numChildren = node.getNumChildren();
        for (var i = 0; i < numChildren; i++) {
            elements = elements.concat(getAllTextElementsWithUrls(node.getChild(i)));
        }
    }
    return elements;
}
