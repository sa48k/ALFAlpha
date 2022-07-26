/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService().reset();
}

/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('Google')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/v2/auth')
      .setTokenUrl('https://oauth2.googleapis.com/token')

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function that should be invoked to
      // complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scope and additional Google-specific parameters.
      .setScope('https://www.googleapis.com/auth/drive')
      .setScope('https://www.googleapis.com/auth/userinfo.profile')
      .setScope('https://www.googleapis.com/auth/userinfo.email')
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    var token = service.getAccessToken();
    var neotoken = alfLogin(token);     // sets the neotoken in the tokens sheet, too
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register in the Google Developers Console.
 */
function logRedirectUri() {
  Logger.log(OAuth2.getRedirectUri());
}

function loginFailed(e){
  Logger.log('Error: Not logged in');
  var timestamp = e.namedValues['Timestamp']; // TODO: Fix possible collisions
  var ss = SpreadsheetApp.openById("1RY-YwxcK3se4lGfZ9owpiTD3IaSRRoCYSWLW6ySPfzA"); // response spreadsheet
  var sh = ss.getSheetByName(e.range.getSheet().getName()); 
  var textFinder = sh.createTextFinder(timestamp);
  var searchRow = textFinder.findNext().getRow()
  var changeRange = sh.getRange(searchRow, 1, 1, sh.getLastColumn());
  changeRange.setBackgroundRGB(255, 255, 102);
}