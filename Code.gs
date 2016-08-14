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
  var links = getAllHyperLinks();
  var offsetIncrement = 0;
  var parentElement = null;
  _.each(links, function(link) {
    // When we see a new parent element, reset the offset increment.
    if (parentElement === null ||
        !isSameElement(parentElement, link.element.getParent())) {
      parentElement = link.element.getParent();
      offsetIncrement = 0;
    }
    var url = link.url;
    var realUrl = url.split("?hyper")[0]
    var response = null;
    if (realUrl.indexOf("https://github.com") === 0) {
      response = fetchGitHubUrl(realUrl);
    }
    else {
      response = UrlFetchApp.fetch(realUrl).getContentText();
    }

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
          if (link.element.getText().slice(offsetIncrement + link.startOffset, offsetIncrement + link.endOffsetInclusive + 1) !== responseValue) {
            link.element.deleteText(offsetIncrement + link.startOffset,
                                    offsetIncrement + link.endOffsetInclusive);
            link.element.insertText(offsetIncrement + link.startOffset, responseValue);
            var newEndOffsetInclusive = offsetIncrement + link.startOffset +
                responseValue.length - 1;
            // Set this new offset before we equate the two.
            // Note that this does not need to be added to the old offset, but is
            // the entire new offset (since offsetIncrement is already factored in).
            var newOffsetIncrement = (newEndOffsetInclusive - link.endOffsetInclusive);
            link.endOffsetInclusive = newEndOffsetInclusive;
            link.element.setLinkUrl(offsetIncrement + link.startOffset, link.endOffsetInclusive, link.url);
            link.element.setUnderline(offsetIncrement + link.startOffset, link.endOffsetInclusive, false);
            link.element.setForegroundColor(offsetIncrement + link.startOffset, link.endOffsetInclusive,
                                            HYPERIZED_COLOR);
            // Other links in this element are now at new locations. Use the
            // increment to adjust for that.
            offsetIncrement = newOffsetIncrement;
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

            // We can't insert an image into a Text element, so split the Text element into
            // multiple elements and insert the image into the parent Paragraph element.
            if (link.isText) {
              var beforeImageText = link.element.getText().slice(0, link.startOffset);
              var afterImageText = link.element.getText().slice(link.endOffsetInclusive + 1, link.element.getText().length);
            }
            var childIndex = parentElement.getChildIndex(link.element);
            // TODO: Something like...
            // if (link.element.getText().slice(link.startOffset, link.endOffsetInclusive + 1) !== responseValue) {
            parentElement.removeChild(link.element);
            var indexIncrement = 0;
            if (link.isText) {
              parentElement.insertText(childIndex + indexIncrement, beforeImageText);
              indexIncrement += 1;
            }
            var imageIndex = childIndex + indexIncrement;
            parentElement.insertInlineImage(imageIndex, image);
            indexIncrement += 1;
            if (link.isText) {
              parentElement.insertText(childIndex + indexIncrement, afterImageText);
              indexIncrement += 1;
            }

            // Set the link of the image.
            var imageElement = parentElement.getChild(imageIndex);
            imageElement.setLinkUrl(link.url);
          }
        }
      }
    }
  });
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

/**
 * From http://stackoverflow.com/a/28821735
 */
function bodyPath(el, path) {
  path = path? path: [];
  var parent = el.getParent();
  var index = parent.getChildIndex(el);
  path.push(index);
  var parentType = parent.getType();
  if (parentType !== DocumentApp.ElementType.BODY_SECTION) {
    path = bodyPath(parent, path);
  }
  else {
    return path;
  };
  return path;
};

/**
 * Check element equality based on path to body rather than contents.
 *
 * From http://stackoverflow.com/a/28821735
 */
function isSameElement(element1, element2) {
  var path1 = bodyPath(element1);
  var path2 = bodyPath(element2);
  if (path1.length == path2.length) {
    for (var i = 0 ; i < path1.length; i++) {
      if (path1[i] !== path2[i]) {
        return false;
      };
    };
  }
  else {
    return false;
  };
  return true;
};
