// alfalpha v1.0
// Andy Little 2020
// andy@amesbury.school.nz

var CLIENT_ID = 'REDACTED';
var CLIENT_SECRET = 'REDACTED';

function formPosted(e) {                                  // run on form submit // TODO: Lock to prevent this being fired twice? known Google issue
  var responseSheet = e.range.getSheet().getName();       // string: used to differentiate between writing and maths responses
  var neotoken = getNeotoken(e.namedValues['Email address']);                      
  var students = e.namedValues['Student username (e.g. finn.l) or multiple usernames separated by commas'];
  var usersList = students.toString().toLowerCase().split(',').filter(Boolean);    // TODO: Fail gracefully if the vlookup returns no userids
  Logger.log('posting to %d user(s)', usersList.length);
  var timestamp = e.namedValues['Timestamp'];
  for (var i=0; i<usersList.length; i++) {                // iterate over all users in form response. for each, build json and post it
    var student = usersList[i].trim();
    Logger.log('Posting to %s from %s response sheet', student, responseSheet); // TODO: Counter
    var payload = buildJson(e, student);
    Logger.log(payload);
    var resp = uploadEvidence(student, payload, neotoken, timestamp, responseSheet);                        // WARNING: Comment this out when testing
    Logger.log(resp);
  }
}

function buildJson(e, student) {
  var vl = SpreadsheetApp.openById("17rmooC_1BoCM_5Es4WoAu90nmpZKr19kguGvpgJwTY4");  // vlookups spreadsheet
  var vlookupgoals = vl.getSheetByName(e.range.getSheet().getName());                // should be 'writing' or 'maths'
  var vlookupstudents = vl.getSheetByName("students");
  var vlookupteacher = vl.getSheetByName("teachers");
  
  var namedvals = e.namedValues;
  var email = namedvals['Email address'];
  var comment = namedvals['Comment'].toString().substring(0,950); // limit to 950 chars
  Logger.log(comment);
  
  var allgoals = e.values.slice(4).filter(Boolean);
  var goals = allgoals.toString().split(";");      // split to an array         
  goals.pop();
  
  for (var i=0; i<goals.length; i++) {             
    if (goals[i].substring(0, 2) == ', ') {
      goals[i] = goals[i].substring(2);            // remove any leading comma and space from form checkbox response
    } else if (goals[i].substring(0, 1) == ',') {  // probably a better way to do this with regex
      goals[i] = goals[i].substring(1);
    } 
  }
  // get id for student, teacher, and goals
  var studentid = vlookup(vlookupstudents, 2, 1, student);
  var teacherid = vlookup(vlookupteacher, 1, 1, email);
  var ids = []
  for (var i=0; i<goals.length; i++) {
    goalid = vlookup(vlookupgoals, 3, 1, goals[i]);        //    Logger.log('goalid for %s: %s', goals[i], goalid);
    ids.push(goalid);
  } 
  // build the JSON to send with the request to the API
  var payload = {
    "student": studentid,
    "goals": ids,
    "extraGoals": [],
    "comment": comment,
    "teacher": teacherid,
    "approved": true,
    "unit": ""
  }
  return(payload);
}

function uploadEvidence(student, payload, neotoken, timestamp, responseSheet) {                   // make the API call
  var url = "https://alf.amesbury.school.nz/api/evidence/multipleupload";  
  var headers = getHeaders();
  headers.neotoken = neotoken;
  var options = getOptions(headers);
  options.payload = JSON.stringify(payload);

  var response = UrlFetchApp.fetch(url, options);   // WARNING: This line posts to ALF when uncommented // WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING
  var content = response.getContentText();
  var parsed = JSON.parse(content);
  
  var ss = SpreadsheetApp.openById("1RY-YwxcK3se4lGfZ9owpiTD3IaSRRoCYSWLW6ySPfzA"); // response spreadsheet
  var sh = ss.getSheetByName(responseSheet);                                        // get the writing or maths sheet so we can highlight the correct line
  var textFinder = sh.createTextFinder(timestamp);
  var searchRow = textFinder.findNext().getRow()
  var changeRange = sh.getRange(searchRow, 1, 1, sh.getLastColumn());
  if (parsed.status == 'ok') {
    changeRange.setBackgroundRGB(200, 255, 150); // ok
  } else if (parsed.status == 'login') {
    changeRange.setBackgroundRGB(255, 255, 102); // not logged in / token expired
  } else {
    changeRange.setBackgroundRGB(255, 202, 150); // something bad happened
  }
  return(parsed.status);
}

function getNeotoken(email) {
  var ss = SpreadsheetApp.openById("1TkOjWEPJPffwp40U9VMnTX9rTvg2Dl7gFz7XlBhgffg"); // token spreadsheet
  var sh = ss.getSheetByName('tokens'); 
  var textFinder = sh.createTextFinder(email);
  var searchRow = textFinder.findNext().getRow()
  var neotoken = sh.getRange(searchRow, 2).getValue();
  return(neotoken);
}

function login(){              // should only be launched from the response sheet
  var service = getService();
  if (service.hasAccess()) {     
    var neotoken = alfLogin(service.getAccessToken());
    var template = HtmlService.createTemplate('Already logged in');
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    var template = HtmlService.createTemplate('<a href="<?= authorizationUrl ?>" target="_blank">Authorise</a><br>');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page); // This will cause the script to fail if launched from the code editor or a trigger // only login from response sheet
  }
}

function alfLogin(token) {
  // use the google access token to get a neotoken and store it on the tokens sheet 
  var currentUser = Session.getActiveUser().getEmail();
  Logger.log('current user: %s', currentUser);
  var url = "https://alf.amesbury.school.nz/api/neo/login";  
  var data = {
    "token":  "https://alf.amesbury.school.nz/proto/googleapi/oauth/#access_token=" + token + "&token_type=Bearer&expires_in=3599&scope=email%20profile%20https://www.googleapis.com/auth/drive%20https://www.googleapis.com/auth/userinfo.email%20openid%20https://www.googleapis.com/auth/userinfo.profile&authuser=1&hd=amesbury.school.nz&prompt=none"    
  }
  var headers = getHeaders();
  var options = getOptions(headers);
  options.payload = JSON.stringify(data);
  
  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText();
  var data = JSON.parse(content);
  var neotoken = data.result.neoToken;
    
  var ss = SpreadsheetApp.openById("1TkOjWEPJPffwp40U9VMnTX9rTvg2Dl7gFz7XlBhgffg"); // tokens spreadsheet
  var sh = ss.getSheetByName('tokens'); 
  var textFinder = sh.createTextFinder(currentUser);
  var searchRow = textFinder.findNext().getRow()
  var cell = sh.getRange(searchRow, 2);
  cell.setValue(neotoken);
  return(neotoken);
};

function getHeaders() {
   var headers = {
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://alf.amesbury.school.nz",
    "Referer": "https://alf.amesbury.school.nz"
  };
  return(headers);
}

function getOptions(headers) {
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'headers': headers
  };
  return(options);
}

function vlookup(sheet, column, index, value) {
  var lastRow=sheet.getLastRow();
  var data=sheet.getRange(1,column,lastRow,column+index).getValues();
  for(i=0;i<data.length;++i){
    if (data[i][0]==value){
      return data[i][index];
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('alfalpha')
    .addItem('Login', 'login')
    //.addItem('ALF latest row: maths sheet', 'alfLastLine')
    .addToUi();
}

function getCurrentUser() {
  Logger.log('current user: %s', Session.getActiveUser().getEmail());
}