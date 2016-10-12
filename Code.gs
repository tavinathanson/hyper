var HYPERIZED_COLOR = "#cc0099";
var FIGORDER_COLOR = "#0099cc";

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
    .addItem("Grab Hyper URLs", "grabUrls")
    .addItem("Peek at Text Changes", "checkForTextChanges")
    .addItem("Find and Replace URLs", "replaceUrls")
    .addItem("Order Figures", "orderFigures")
    .addItem("Hide Figure Links", "hideFigureOrderElements")
    .addItem("Show Figure Links", "showFigureOrderElements")
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
    if (link.element.getType() == DocumentApp.ElementType.TEXT) {
      link.element.setForegroundColor(link.startOffset,
                                      link.endOffsetInclusive, color);
      link.element.setUnderline(false);
    }
  });
}

/**
 * Color figure order elements with the hyper color.
 */
function showFigureOrderElements() {
  colorFigureOrderElements(FIGORDER_COLOR);
}

/**
 * Set figure order elements back to black.
 */
function hideFigureOrderElements() {
  colorFigureOrderElements("#000000");
}

/**
 * Set figure order elements to a particular color.
 */
function colorFigureOrderElements(color) {
  var links = getAllFigureOrderLinks();
  _.each(links, function(link) {
    if (link.element.getType() == DocumentApp.ElementType.TEXT) {
      link.element.setForegroundColor(link.startOffset,
                                      link.endOffsetInclusive, color);
      link.element.setUnderline(false);
    }
  });
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

function grabUrls() {
  var urls = hyperize(true);
  DocumentApp.getUi().alert("URLs: \n" + _.uniq(urls).join("\n"));
}

function checkForTextChanges() {
  var changes = hyperize(false, true);
  DocumentApp.getUi().alert("Changes: \n" + changes.join("\n"));
}

function replaceUrls() {
  // From http://stackoverflow.com/a/17606289
  String.prototype.replaceAll = function(search, replacement) {
    var target = this;
    return target.split(search).join(replacement);
  };

  var ui = DocumentApp.getUi();
  var textToReplaceResponse = ui.prompt(
    "Find and Replace Hyper Link URLs",
    "URL Text to Replace",
    ui.ButtonSet.OK_CANCEL);
  var newTextResponse = ui.prompt(
    "Find and Replace Hyper Link URLs",
    "New URL Text",
    ui.ButtonSet.OK_CANCEL);
  if (textToReplaceResponse.getSelectedButton() === ui.Button.OK &&
      newTextResponse.getSelectedButton() === ui.Button.OK) {
    var textToReplace = textToReplaceResponse.getResponseText();
    var newText = newTextResponse.getResponseText();
    var links = getAllHyperLinks();
    _.each(links, function(link) {
      if (link.url.indexOf(textToReplace) !== -1) {
        var updatedUrl = link.url.replaceAll(textToReplace, newText);
        link.element.setLinkUrl(link.startOffset, link.endOffsetInclusive, updatedUrl);
      }
    });
  }
}

function orderFigures() {
  var changingHyperObjects = {};
  var errors = [];
  var labelToOrderedLabel = {};

  function unpackKey(link) {
    var url = link.url;
    var fullKey = url.split("?figorder=")[1];
    var group = "default";
    var isNew = false;
    var sectionNum = null;
    if (fullKey.indexOf("&group") !== -1) {
      group = fullKey.split("&group=")[1].split("&")[0];
    }
    if (fullKey.indexOf("&section") !== -1) {
      sectionNum = fullKey.split("&section=")[1].split("&")[0];
    }
    if (fullKey.indexOf("&new") !== -1) {
      isNew = true;
    }
    var key = fullKey.split("&")[0]

    return [key, group, sectionNum, isNew];
  }

  // First, split elements.
  var figOrderLinks = getAllFigureOrderLinks();

  if (figOrderLinks.length === 0) {
    return;
  }

  var textElements = getAllTextElements();
  _.each(textElements, function(textElement) {
    var figOrderLinksFromText = getAllFigureOrderLinks(textElement);
    var newTextElements = splitTextByLinks(textElement, figOrderLinksFromText);
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

  // Re-run after splitting.
  figOrderLinks = getAllFigureOrderLinks();
  var groups = {};
  _.each(figOrderLinks, function(link) {
    var unpackedKey = unpackKey(link);
    var group = unpackedKey[1];
    groups[group] = true;
  });
  groups = _.keys(groups);

  var sectionCount;
  var sectionFigureCount = {};
  _.each(groups, function(curGroup) {
    sectionCount = 1;
    _.each(figOrderLinks, function(link) {
      var unpackedKey = unpackKey(link);
      var key = unpackedKey[0];
      var group = unpackedKey[1];
      var sectionNum = unpackedKey[2];
      var isNew = unpackedKey[3];

      if (group === curGroup) {
        if (!labelToOrderedLabel.hasOwnProperty(key)) {
          if (sectionNum !== null) {
            if (!sectionFigureCount.hasOwnProperty(sectionNum)) {
              sectionFigureCount[sectionNum] = -1;
            }
            sectionFigureCount[sectionNum] += 1;
            labelToOrderedLabel[key] = sectionNum + String.fromCharCode(65 + sectionFigureCount[sectionNum]);
          }
          else {
            if (isNew) {
              sectionCount += 1;
            }
            if (!sectionFigureCount.hasOwnProperty(sectionCount)) {
              sectionFigureCount[sectionCount] = 0;
            }
            else {
              sectionFigureCount[sectionCount] += 1;
            }
            labelToOrderedLabel[key] = sectionCount + String.fromCharCode(65 + sectionFigureCount[sectionCount]);
          }
        }
      }
    });
  });

  _.each(figOrderLinks, function(link) {
    var linkElement = link.element;
    var parentElement = linkElement.getParent();
    var elementIndex = parentElement.getChildIndex(linkElement);
    var unpackedKey = unpackKey(link);
    var key = unpackedKey[0];
    var figValue = labelToOrderedLabel[key];

    parentElement.removeChild(linkElement);
    parentElement.insertText(elementIndex, figValue);
    var newElement = parentElement.getChild(elementIndex);

    newElement.setLinkUrl(link.url);
    newElement.setUnderline(false);
    newElement.setForegroundColor(FIGORDER_COLOR);
  });
}

/**
 * Replace hyper link elements with the text that they point to.
 */
function hyperize(onlyGrabUrls, onlyCheckForTextChanges) {
  var errors = [];

  var userProperties = PropertiesService.getUserProperties();
  var scriptProperties = PropertiesService.getScriptProperties();

  // Cached results after fetching URLs and looking for images in responses.
  var fetchedUrls = {};
  var responseImages = {};

  askForGitHubAccess();
  var gitHubService = waitForGitHubAccess();

  // Be sure not to grab this on every URL fetch.
  var accessToken = gitHubService.getAccessToken();

  /**
   * Configures the GitHub authorization service.
   */
  function getGitHubService() {
    return getGitHubServiceWithProperties(scriptProperties, userProperties);
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
    var numTries = 0;
    while (!gitHubService.hasAccess()) {
      if (numTries > 10) {
        throw new Error("Cannot reach GitHub.");
      }

      Utilities.sleep(1000);
      numTries += 1;
    }

    return gitHubService;
  }

  /**
   * Use GitHub authorization to fetch the contents of a GitHub URL.
   */
  function fetchGitHubUrl(gitHubUrl) {
    if (fetchedUrls.hasOwnProperty(gitHubUrl)) {
      return fetchedUrls[gitHubUrl];
    }

    gitHubUrlParts = gitHubUrl.split("/");
    gitHubIndex = gitHubUrlParts.indexOf("github.com");
    orgName = gitHubUrlParts[gitHubIndex + 1];
    repoName = gitHubUrlParts[gitHubIndex + 2];
    blobIndex = gitHubUrlParts.indexOf("blob");
    branch = gitHubUrlParts[blobIndex + 1];
    path = gitHubUrlParts.slice(blobIndex + 2).join("/");
    var url = "https://api.github.com/repos/" + orgName + "/" + repoName + "/contents/" +
        path + "?ref=" + branch;
    try {
      var response = UrlFetchApp.fetch(url, {
        headers: {
          Authorization: "Bearer " + accessToken,
          Accept: "application/vnd.github.v3.raw"
        }
      });
    }
    catch (e) {
      errors.push(e.message);
      return null;
    }

    var contentText = response.getContentText();
    fetchedUrls[gitHubUrl] = contentText;
    return contentText;
  }

  function getAllChangingHyperObjects() {
    var links = getAllHyperLinks();
    var changingHyperObjects = {};
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

      function pushError(link, key, label) {
        errors.push("Not found in response: " + label + " (" + link.url + ")");
        if (!changingHyperObjects.hasOwnProperty(link.url)) {
          changingHyperObjects[link.url] = [];
        }
        changingHyperObjects[link.url].push({"value": "[Not Found]", "to_text": true, "link": link, "image_size": null, "key": key});
      }

      if (response !== null) {
        if (url.indexOf("?hyper=") !== -1) {
          var key = url.split("?hyper=")[1]
          var responseKey = "{{{" + key + ":";
          var responseKeyIndex = response.indexOf(responseKey);
          if (responseKeyIndex !== -1) {
            var responsePartial = response.substr(responseKeyIndex);
            var responsePartialKeyEndIndex = responseKey.length;
            var responsePartialValueEndIndex = responsePartial.indexOf("}}}");
            var responseValue = responsePartial.substr(
              responsePartialKeyEndIndex,
              responsePartialValueEndIndex - responsePartialKeyEndIndex);
            if (link.isText) {
              // This is called before the links are chopped up. (And after, though the links will have different start/end offsets then.)
              if (link.element.getText().slice(link.startOffset, link.endOffsetInclusive + 1) !== responseValue) {
                if (!changingHyperObjects.hasOwnProperty(link.url)) {
                  changingHyperObjects[link.url] = [];
                }
                changingHyperObjects[link.url].push({"value": responseValue, "to_text": true, "link": link, "image_size": null, "key": key});
              }
            }
          }
          else {
            var keyAndSize = url.split("?hyper=")[1]
            var keyAndSizeSplit = keyAndSize.split("&size=")
            var key = keyAndSizeSplit[0]
            var imageSize = null;
            if (keyAndSizeSplit.length > 1) {
              imageSize = keyAndSizeSplit[1]
            }
            var label = "{{{" + key + "}}}";
            var responseKeyIndex = response.indexOf(label);
            if (responseKeyIndex != -1) {
              var responseJSON = JSON.parse(response);
              var labelToImage = getAllResponseImages(response, url);
              if (labelToImage !== null && labelToImage.hasOwnProperty(label)) {
                var responseValue = labelToImage[label];
                if (!changingHyperObjects.hasOwnProperty(link.url)) {
                  changingHyperObjects[link.url] = [];
                }
                changingHyperObjects[link.url].push({"value": responseValue, "to_text": false, "link": link, "image_size": imageSize, "key": key});
              }
              else {
                pushError(link, key, label);
              }
            }
            else {
              pushError(link, key, label);
            }
          }
        }
      }
    });
    return changingHyperObjects;
  }

  function getAllResponseImages(response, url) {
    if (responseImages.hasOwnProperty(response)) {
      return responseImages[response];
    }

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

    responseImages[response] = labelsToImages;

    if (_.size(labels) !== _.size(labelsToImages)) {
      var labelsNotInMap = _.difference(labels, _.keys(labelsToImages));
      errors.push("Not every label (" + labelsNotInMap +
                  ") corresponds to an image in: " + url);
      return null;
    }

    return labelsToImages;
  }

  if (onlyGrabUrls === true) {
    var realUrls = {};
    var links = getAllHyperLinks();
    _.each(links, function(link) {
      var url = link.url;
      var realUrl = url.split("?hyper")[0];
      realUrls[realUrl] = true;
    });

    return _.keys(realUrls);
  }

  // Split all links into their own text elements
  var changingHyperObjects = getAllChangingHyperObjects();

  if (onlyCheckForTextChanges === true) {
    var changingTexts = [];
    _.each(changingHyperObjects, function(hyperObjects) {
      _.each(hyperObjects, function(hyperObject) {
        var responseValue = hyperObject.value;
        var toText = hyperObject.to_text;
        if (toText == true) {
          var link = hyperObject.link;
          var linkElement = link.element;
          var oldText = linkElement.getText().slice(link.startOffset, link.endOffsetInclusive + 1)
          if (oldText !== responseValue) {
            changingTexts.push("Text [" + oldText + "] (" + hyperObject.key +
                               ") will change to [" + responseValue + "] for " +
                               link.url);
          }
        }
      });
    });

    return _.uniq(changingTexts);
  }

  var textElements = getAllTextElements();
  _.each(textElements, function(textElement) {
    var links = getLinksFromText(textElement);
    var changingLinksFromText = [];
    _.each(links, function(link) {
      if (changingHyperObjects.hasOwnProperty(link.url)) {
        changingLinksFromText.push(link);
      }
    });
    if (changingLinksFromText.length > 0) {
      var newTextElements = splitTextByLinks(textElement, changingLinksFromText);
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
    }
  });

  // We must re-run to get the split elements.
  changingHyperObjects = getAllChangingHyperObjects();
  _.each(_.allKeys(changingHyperObjects), function(url) {
    var objectsForUrl = changingHyperObjects[url];
    _.each(objectsForUrl, function(hyperObject) {
      var responseValue = hyperObject.value;
      var toText = hyperObject.to_text;
      var link = hyperObject.link;
      var linkElement = link.element;
      var parentElement = linkElement.getParent();
      var elementIndex = parentElement.getChildIndex(linkElement);
      if (toText === true) {
        parentElement.removeChild(linkElement);
        parentElement.insertText(elementIndex, responseValue);
        var newElement = parentElement.getChild(elementIndex);

        newElement.setLinkUrl(link.url);
        newElement.setUnderline(false);
        newElement.setForegroundColor(HYPERIZED_COLOR);
      }
      // No hyper text label found? Look for images!
      else {
        var imageSize = hyperObject.image_size;
        parentElement.removeChild(linkElement);
        parentElement.insertInlineImage(elementIndex, responseValue);
        var newElement = parentElement.getChild(elementIndex);

        newElement.setLinkUrl(link.url);

        // Resize if a ratio is specified.
        if (imageSize !== null) {
          var curWidth = newElement.getWidth();
          var curHeight = newElement.getHeight();
          newElement.setWidth(curWidth * imageSize);
          newElement.setHeight(curHeight * imageSize);
        }
      }
    });
  });

  // Some links didn't change, but color them anyway.
  showHyperElements();

  if (errors.length > 0) {
    DocumentApp.getUi().alert("Errors encountered: \n" + _.uniq(errors).join("\n"));
  }
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

function getAllFigureOrderLinks(element) {
  var element = element || DocumentApp.getActiveDocument().getBody();
  var links = getAllLinks(element);
  var figOrderLinks = [];
  for (var i = 0; i < links.length; i++) {
    var link = links[i];
    var url = link.url;
    if (url.indexOf("?figorder") != -1) {
      figOrderLinks.push(link);
    }
  }
  return figOrderLinks;
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
 * Configures the GitHub authorization service.
 */
function getGitHubServiceWithProperties(scriptProperties, userProperties) {
  var gitHubClientId = scriptProperties.getProperty("GITHUB_CLIENT_ID");
  var gitHubClientSecret = scriptProperties.getProperty("GITHUB_CLIENT_SECRET");
  return OAuth2.createService("GitHub")
    .setAuthorizationBaseUrl("https://github.com/login/oauth/authorize")
    .setTokenUrl("https://github.com/login/oauth/access_token")
    .setClientId(gitHubClientId)
    .setClientSecret(gitHubClientSecret)
    .setCallbackFunction("authCallback")
    .setPropertyStore(userProperties)
    .setScope("repo") // Need access to code
    .setParam("allow_signup", true);
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var userProperties = PropertiesService.getUserProperties();
  var scriptProperties = PropertiesService.getScriptProperties();
  var service = getGitHubServiceWithProperties(scriptProperties, userProperties);
  service.reset();
}

/**
 * Handles the GitHub OAuth callback.
 */
function authCallback(request) {
  var userProperties = PropertiesService.getUserProperties();
  var scriptProperties = PropertiesService.getScriptProperties();
  var service = getGitHubServiceWithProperties(scriptProperties, userProperties);
  var authorized = service.handleCallback(request);
  var template = HtmlService.createTemplateFromFile("github_oauth_callback");
  var callbackMessage = template.callbackMessage = authorized ?
      "GitHub authorization complete. You may return to the document." :
      "GitHub authorization was not successful.";
  template.callbackMessage = callbackMessage;
  var page = template.evaluate();
  return page;
}
