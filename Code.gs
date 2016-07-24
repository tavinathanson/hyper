var scriptProperties = PropertiesService.getScriptProperties();
var GITHUB_CLIENT_ID = scriptProperties.getProperty("GITHUB_CLIENT_ID");
var GITHUB_CLIENT_SECRET = scriptProperties.getProperty("GITHUB_CLIENT_SECRET");

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen(e) {
  showSidebar();
  DocumentApp.getUi().createAddonMenu()
    .addItem('Start', 'replaceText')
    .addToUi();
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
      var response = null;
      if (realUrl.indexOf("https://github.com") == 0) {
        response = run(realUrl);
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

function run(gitHubUrl) {
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
  } else {
    showSidebar();
    if (service.hasAccess()) {
      return run();
    }
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
  return OAuth2.createService('GitHub')
    .setAuthorizationBaseUrl('https://github.com/login/oauth/authorize')
    .setTokenUrl("https://github.com/login/oauth/access_token")
    .setClientId(GITHUB_CLIENT_ID)
    .setClientSecret(GITHUB_CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('repo') // Need access to code
    .setParam('allow_signup', true);
}

function showSidebar() {
  var service = getGitHubService();
  if (!service.hasAccess()) {
    var authorizationUrl = service.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
      '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    DocumentApp.getUi().showSidebar(page);
  }
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getGitHubService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied');
  }
}
