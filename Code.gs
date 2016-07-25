var scriptProperties = PropertiesService.getScriptProperties();
var GITHUB_CLIENT_ID = scriptProperties.getProperty("GITHUB_CLIENT_ID");
var GITHUB_CLIENT_SECRET = scriptProperties.getProperty("GITHUB_CLIENT_SECRET");

var HYPERIZED_COLOR = "#cc0099";

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen(e) {
  askForAccess();
  waitForAccess();
  DocumentApp.getUi().createAddonMenu()
    .addItem("Hyperize Links", "replaceText")
    .addItem("Hide Hyper Links", "hideHyperElements")
    .addItem("Show Hyper Links", "showHyperElements")
    .addToUi();
  var app = UiApp.getActiveApplication();
  app.close();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
}

function showHyperElements() {
  colorHyperElements(HYPERIZED_COLOR);
}

function hideHyperElements() {
  colorHyperElements("#000000");
}

function colorHyperElements(color) {
  var elements = getAllHyperElements();
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    element.setForegroundColor(color);
  }
}

function replaceText() {
  var elements = getAllHyperElements();
  for (var i = 0; i < elements.length; i++) {
    element = elements[i];
    url = element.getLinkUrl(0);
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
      element.replaceText(".*", responseValue);
      element.setUnderline(false);
      element.setForegroundColor(HYPERIZED_COLOR);
    }
  }
}

function getAllHyperElements() {
  var elements = getAllTextElementsWithUrls();
  var hyperElements = [];
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    var url = element.getLinkUrl(0);
    if (url.indexOf("?hyper=") != -1) {
      hyperElements.push(element);
    }
  }
  return hyperElements;
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
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getGitHubService();
  service.reset();
}

/**
 * Configures the service.
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

function askForAccess() {
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

function waitForAccess() {
  var gitHubService = getGitHubService();
  while (!gitHubService.hasAccess()) {
    Utilities.sleep(1000);
  }
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getGitHubService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput("GitHub authorization complete. You may return to the document.");
  } else {
    return HtmlService.createHtmlOutput("GitHub authorization was not successful.");
  }
}
