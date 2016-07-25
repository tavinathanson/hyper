var scriptProperties = PropertiesService.getScriptProperties();
var GITHUB_CLIENT_ID = scriptProperties.getProperty("GITHUB_CLIENT_ID");
var GITHUB_CLIENT_SECRET = scriptProperties.getProperty("GITHUB_CLIENT_SECRET");

var HYPERIZED_COLOR = "#cc0099";

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem("Hyperize Links", "hyperize")
    .addItem("Hide Hyper Links", "hideHyperElements")
    .addItem("Show Hyper Links", "showHyperElements")
    .addToUi();
  askForGitHubAccess();
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
  var elements = getAllHyperLinks();
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    element.setForegroundColor(color);
  }
}

/**
 * Replace hyper link elements with the text that they point to.
 */
function hyperize() {
  waitForGitHubAccess();
  var links = getAllHyperLinks();
  for (var i = 0; i < links.length; i++) {
    link = links[i];
    url = link.url;
    var key = url.split("?hyper=")[1]
    var realUrl = url.split("?hyper=")[0]
    var response = null;
    if (realUrl.indexOf("https://github.com") == 0) {
      response = fetchGitHubUrl(realUrl);
    }
    else {
      response = UrlFetchApp.fetch(realUrl).getContentText();
    }
    var responseKey = "{{{" + key + ":";
    var responseKeyIndex = response.indexOf(responseKey);
    if (responseKeyIndex != -1) {
      responsePartial = response.substr(responseKeyIndex);
      responsePartialKeyEndIndex = responseKey.length;
      responsePartialValueEndIndex = responsePartial.indexOf("}}}");
      responseValue = responsePartial.substr(
        responsePartialKeyEndIndex,
        responsePartialValueEndIndex - responsePartialKeyEndIndex);
      link.element.deleteText(link.startOffset, link.endOffsetInclusive);
      link.element.insertText(link.startOffset, responseValue);
      var newEndOffsetInclusive = link.startOffset + responseValue.length - 1;
      link.element.setLinkUrl(link.startOffset, newEndOffsetInclusive, link.url);
      link.element.setUnderline(link.startOffset, newEndOffsetInclusive, false);
      link.element.setForegroundColor(link.startOffset, newEndOffsetInclusive,
                                      HYPERIZED_COLOR);
    }
  }
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
    if (url.indexOf("?hyper=") != -1) {
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
  else {
    var numChildren = element.getNumChildren();
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
  return OAuth2.createService("GitHub")
    .setAuthorizationBaseUrl("https://github.com/login/oauth/authorize")
    .setTokenUrl("https://github.com/login/oauth/access_token")
    .setClientId(GITHUB_CLIENT_ID)
    .setClientSecret(GITHUB_CLIENT_SECRET)
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
