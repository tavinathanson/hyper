var HYPERIZED_COLOR = "#cc0099";

// Added using the project key in https://github.com/simula-innovation/gas-underscore
var _ = Underscore.load();

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem("Hyperize Links", "hyperize")
    .addItem("Hide Hyper Links", "hideHyperElements")
    .addItem("Show Hyper Links", "showHyperElements")
    .addItem("Remove GitHub Authorization", "reset")
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Color hyper elements with the hyper color.
 */
function showHyperElements() {
  colorHyperElements(HYPERIZED_COLOR);
}

/**
 * Set hyper elements back to black.
 */
function hideHyperElements() {
  colorHyperElements("#000000");
}

/**
 * Set hyper elements to a particular color.
 */
function colorHyperElements(color) {
  var links = getAllHyperLinks();
  _.each(links, function(link) {
    link.element.setForegroundColor(link.startOffset, link.endOffsetInclusive, color);
  });
}

/**
 * Replace hyper link elements with the text that they point to.
 */
function hyperize() {
  askForGitHubAccess();
  waitForGitHubAccess();

  // Split all links into their own text elements
  var textElements = getAllTextElements();
  _.each(textElements, function(textElement) {
    var links = getLinksFromText(textElement);
    var newTextElements = splitTextByLinks(textElement, links);
    var parentElement = textElement.getParent();
    var textElementIndex = parentElement.getChildIndex(textElement);
    parentElement.removeChild(textElement);
    var i = 0;
    _.each(newTextElements, function(newTextElement) {
      if (newTextElement.getText().length > 0) {
        parentElement.insertText(textElementIndex + i, newTextElement);
        i += 1;
      }
    });
  });

  var links = getAllHyperLinks();
  _.each(links, function(link) {
    var url = link.url;
    var realUrl = url.split("?hyper")[0]
    var response = null;
    if (realUrl.indexOf("https://github.com") === 0) {
      response = fetchGitHubUrl(realUrl);
    }
    else {
      response = UrlFetchApp.fetch(realUrl).getContentText();
    }

    var parentElement = link.element.getParent();

    if (url.indexOf("?hyper=") !== -1) {
      var key = url.split("?hyper=")[1]
      var responseKey = "{{{" + key + ":";
      var responseKeyIndex = response.indexOf(responseKey);
      if (responseKeyIndex != -1) {
        var responsePartial = response.substr(responseKeyIndex);
        var responsePartialKeyEndIndex = responseKey.length;
        var responsePartialValueEndIndex = responsePartial.indexOf("}}}");
        var responseValue = responsePartial.substr(
          responsePartialKeyEndIndex,
          responsePartialValueEndIndex - responsePartialKeyEndIndex);
        if (link.isText) {
          if (link.element.getText() !== responseValue) {
            var elementIndex = parentElement.getChildIndex(link.element);
            parentElement.removeChild(link.element);
            parentElement.insertText(elementIndex, responseValue);
            var newElement = parentElement.getChild(elementIndex);

            newElement.setLinkUrl(link.url);
            newElement.setUnderline(false);
            newElement.setForegroundColor(HYPERIZED_COLOR);
          }
        }
        else {
          // Weird case: image turns into text?
        }
      }
      // No hyper text label found? Look for images!
      else {
        var key = url.split("?hyper=")[1]
        var label = "{{{" + key + "}}}";
        var responseKeyIndex = response.indexOf(label);
        if (responseKeyIndex != -1) {
          var responseJSON = JSON.parse(response);
          var labelToImage = getAllResponseImages(response);
          var image = labelToImage[label];
          var elementIndex = parentElement.getChildIndex(link.element);
          parentElement.removeChild(link.element);
          parentElement.insertInlineImage(elementIndex, image);
          var newElement = parentElement.getChild(elementIndex);

          newElement.setLinkUrl(link.url);
        }
      }
    }
});
}

function splitTextByLinks(textElement, links) {
  var newTextElements = [];
  // Example starting points
  // |         |         |
  // abc[link1]cde[link2]hij
  var startingPoint = 0;
  var textLength = textElement.getText().length;
  _.each(links, function(link) {
    // abc, cde
    var newTextElement = textElement.copy();
    if (startingPoint > 0) {
      newTextElement.deleteText(0, startingPoint - 1);
    }
    newTextElement.deleteText(link.startOffset - startingPoint, textLength - 1 - startingPoint);
    newTextElements.push(newTextElement);
    startingPoint += newTextElement.getText().length;

    // link1, link2
    newTextElement = textElement.copy();
    if (link.endOffsetInclusive + 1 - link.startOffset <= textLength - 1 - link.startOffset) {
      if (link.startOffset > 0) {
        newTextElement.deleteText(0, link.startOffset - 1);
      }
      newTextElement.deleteText(link.endOffsetInclusive + 1 - link.startOffset, textLength - 1 - link.startOffset);
      newTextElements.push(newTextElement);
      startingPoint += newTextElement.getText().length;
    }
  });

  // hij
  newTextElement = textElement.copy();
  if (startingPoint > 0 && startingPoint < textLength) {
    newTextElement.deleteText(0, startingPoint - 1);
  }
  newTextElements.push(newTextElement);

  return newTextElements;
}

function getAllResponseImages(response) {
  var responseJSON = JSON.parse(response);
  var cells = responseJSON.cells;
  var labelsToImages = {};
  var labels = [];
  var lastLabel = null;
  var lastImage = null;
  _.each(cells, function(cell) {
    var outputs = cell.outputs;
    _.each(outputs, function(output) {
      // TODO: Use better contains.
      if ("data" in output) {
        if ("image/png" in output["data"]) {
          var decoded = Utilities.base64Decode(output["data"]["image/png"]);
          var blob = Utilities.newBlob(decoded);
          if (lastLabel !== null) {
            labelsToImages[lastLabel] = blob;
            lastLabel = null;
          }
        }
      }
      if ("text" in output) {
        var text = output["text"];
        _.each(text, function(line) {
          var line = String(line);
          if (line.indexOf("{{{") !== -1 && line.indexOf("}}}") !== -1) {
            var startLabelIndex = line.indexOf("{{{");
            var endLabelIndex = line.indexOf("}}}");
            var label = line.slice(startLabelIndex, endLabelIndex + "}}}".length);
            // Don't include non-image hyper labels
            if (label.indexOf(":") === -1) {
              labels.push(label);
              lastLabel = label;
            }
          }
        });
      }
    });
  });

  if (_.size(labels) !== _.size(labelsToImages)) {
    throw new Error("Not every label corresponds to an image!");
  }

  return labelsToImages;
}

/**
 * Returns a list of all hyper elements.
 */
function getAllHyperLinks() {
  var links = getAllLinks();
  var hyperLinks = [];
  for (var i = 0; i < links.length; i++) {
    var link = links[i];
    var url = link.url;
    if (url.indexOf("?hyper") != -1) {
      hyperLinks.push(link);
    }
  }
  return hyperLinks;
}

/**
 * Use GitHub authorization to fetch the contents of a GitHub URL.
 */
function fetchGitHubUrl(gitHubUrl) {
  var service = getGitHubService();
  if (service.hasAccess()) {
    gitHubUrlParts = gitHubUrl.split("/");
    gitHubIndex = gitHubUrlParts.indexOf("github.com");
    orgName = gitHubUrlParts[gitHubIndex + 1];
    repoName = gitHubUrlParts[gitHubIndex + 2];
    blobIndex = gitHubUrlParts.indexOf("blob");
    branch = gitHubUrlParts[blobIndex + 1];
    path = gitHubUrlParts.slice(blobIndex + 2).join("/");
    var url = "https://api.github.com/repos/" + orgName + "/" + repoName + "/contents/" +
        path + "?ref=" + branch;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: "Bearer " + service.getAccessToken(),
        Accept: "application/vnd.github.v3.raw"
      }
    });
    return response.getContentText();
  }
}

function getAllTextElements(element) {
  var textElements = [];
  var element = element || DocumentApp.getActiveDocument().getBody();
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    textElements.push(element);
  }
  else {
    var numChildren = 0;
    try {
      numChildren = element.getNumChildren();
    }
    catch (e) {}
    for (var i = 0; i < numChildren; i++) {
      textElements = textElements.concat(getAllTextElements(element.getChild(i)));
    }
  }

  return textElements;
}

function getLinksFromText(element) {
  var links = [];
  var textObj = element.editAsText();
  var text = element.getText();
  var inUrl = false;
  var url = null;
  var curUrl = null;
  for (var ch = 0; ch < text.length; ch++) {
    url = textObj.getLinkUrl(ch);
    if (url != null) {
      if (!inUrl) {
        inUrl = true;
        curUrl = {};
        curUrl.isText = true;
        curUrl.element = element;
        curUrl.url = String(url);
        curUrl.startOffset = ch;
      }
      else {
        curUrl.endOffsetInclusive = ch;
      }
    }
    else {
      if (inUrl) {
        inUrl = false;
        if (curUrl.endOffsetInclusive === undefined) {
          curUrl.endOffsetInclusive = curUrl.startOffset;
        }
        links.push(curUrl);
        curUrl = {};
      }
    }
  }
  // Needed for the case when no non-link character comes after a link.
  if (url != null && curUrl != null) {
    inUrl = false;
    if (curUrl.endOffsetInclusive === undefined) {
      curUrl.endOffsetInclusive = curUrl.startOffset;
    }
    links.push(curUrl);
    curUrl = {};
  }

  return links;
}

/**
 * Get all links from the document, which is tricky because links might be only
 * a portion of an element.
 *
 * Mostly taken from https://gist.github.com/mogsdad/6518632#file-getalllinks-js.
 * Added a bug fix to the case where no non-link character follows a link.
 */
function getAllLinks(element) {
  var links = [];
  element = element || DocumentApp.getActiveDocument().getBody();
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    links = links.concat(getLinksFromText(element));
  }
  else if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
    url = element.getLinkUrl();
    if (url != null) {
      curUrl = {};
      curUrl.isText = false;
      curUrl.element = element;
      curUrl.url = String(url);
      links.push(curUrl);
    }
  }
  else {
    var numChildren = 0;
    try {
      numChildren = element.getNumChildren();
    }
    catch (e) {}
    for (var i = 0; i < numChildren; i++) {
      links = links.concat(getAllLinks(element.getChild(i)));
    }
  }

  return links;
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getGitHubService();
  service.reset();
}

/**
 * Configures the GitHub authorization service.
 */
function getGitHubService() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var gitHubClientId = scriptProperties.getProperty("GITHUB_CLIENT_ID");
  var gitHubClientSecret = scriptProperties.getProperty("GITHUB_CLIENT_SECRET");
  return OAuth2.createService("GitHub")
    .setAuthorizationBaseUrl("https://github.com/login/oauth/authorize")
    .setTokenUrl("https://github.com/login/oauth/access_token")
    .setClientId(gitHubClientId)
    .setClientSecret(gitHubClientSecret)
    .setCallbackFunction("authCallback")
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope("repo") // Need access to code
    .setParam("allow_signup", true);
}

/**
 * Prompts the user for a modal dialog to gain GitHub access.
 */
function askForGitHubAccess() {
  var gitHubService = getGitHubService();
  if (!gitHubService.hasAccess()) {
    var authorizationUrl = gitHubService.getAuthorizationUrl();
    // Close the modal after authorization is attempted.
    var template = HtmlService.createTemplateFromFile("github_auth");
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate().setHeight(50);
    DocumentApp.getUi().showModalDialog(page, "Hyper: Authorize GitHub");
  }
}

/**
 * Blocks until GitHub access is established.
 */
function waitForGitHubAccess() {
  var gitHubService = getGitHubService();
  while (!gitHubService.hasAccess()) {
    Utilities.sleep(1000);
  }
}

/**
 * Handles the GitHub OAuth callback.
 */
function authCallback(request) {
  var service = getGitHubService();
  var authorized = service.handleCallback(request);
  var template = HtmlService.createTemplateFromFile("github_oauth_callback");
  var callbackMessage = template.callbackMessage = authorized ?
      "GitHub authorization complete. You may return to the document." :
      "GitHub authorization was not successful.";
  template.callbackMessage = callbackMessage;
  var page = template.evaluate();
  return page;
}
