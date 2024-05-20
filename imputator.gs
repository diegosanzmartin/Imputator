var jiraApiUrl = 'https://makingscience.atlassian.net';
var jiraApiToken = '';
var jiraUsername = '';

function getHeaders() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Imputator");
  var headers = {};

  for (var i = 1; i < sheet.getLastColumn(); i++) {
    h = sheet.getRange(1, i).getValue();
    headers[h] = i;
  }

  return headers;
}

function onEdit(e) {
  var sheet = e.source.getSheetByName('Imputator');
  var range = e.range;
  var headers = getHeaders();

  if (range.getColumn() == headers["Day"] && range.getRow() > 1) {
    var targetRow = range.getRow();
    var value = sheet.getRange(targetRow, headers["Day"]).getValue();

    if (value == "") {
      var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
      sheet.getRange(targetRow, headers["Day"]).setValue(currentDate);
    }
  }

  if (range.getColumn() == headers["Issue Ticket"] && range.getRow() > 1) {
    var targetRow = range.getRow();
    var issueCell = sheet.getRange(targetRow, headers["Issue Ticket"]).getValue();

    if (issueCell != "") {
      var hyperlink = "";

      if(issueCell.startsWith("http")) {
        var hyperlink = '=HYPERLINK("' + issueCell + '"; "' + issueCell + '")';
      }
      else {
        var issueKey = sheet.getRange(targetRow, headers["Issue Ticket"]).getValue().split(":")[0];
        var hyperlink = '=HYPERLINK("' + jiraApiUrl + '/browse/' + issueKey + '"; "' + issueCell + '")';
      }
      
      sheet.getRange(targetRow, headers["Issue Ticket"]).setValue(hyperlink);
    }
  }
}

function imputeWorklog() {
  Logger.log('EDIT');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Imputator");
  var headers = getHeaders();

  var lastRow = true;
  var r = 530;

  while(lastRow) {
    var ticket = sheet.getRange(r, headers["Issue Ticket"]).getValue();
    var comment = sheet.getRange(r, headers["Comment"]).getValue();
    var isImputed = sheet.getRange(r, headers["Imputed"]).getValue();
    var status = sheet.getRange(r, headers["Status"]).getValue();

    var day = sheet.getRange(r, headers["Day"]).getValue();
    var start = sheet.getRange(r, headers["Start"]).getValue();
    var end = sheet.getRange(r, headers["End"]).getValue();

    if (isImputed == true && ticket != '' && status == '') {
      if(ticket.startsWith("http")) {
        var issueKey = ticket.split("/")[4];
      }
      else {
        var issueKey = ticket.split(":")[0];
      }

      var startedHour = "09:00:00.000+0100";
      var started = Utilities.formatDate(day, "Europe/Madrid", "yyyy-MM-dd'T'") + startedHour;
      var timeSpent = String((end - start)/(1000 * 60)) + "m";

      response = logWorkInJira(issueKey, timeSpent, comment, started);

      sheet.getRange(r, headers["Status"]).setValue(response);
    }

    if (isImputed == false 
      && ticket == ''
      && status == ''
      && day == ''
      && start == ''
      && end == ''
    ) {
      lastRow = false;
    }

    r++;
  }
}

function getTickets() {
  Logger.log('OPEN');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Jira");

  if (sheet) {
    responseData = queryJiraAPI()

    if (responseData.issues) {
      for (var i = 0; i < 500; i++) {
        cell = String(2+i);
        sheet.getRange("A" + cell).setValue("");
      }

      for (var i = 0; i < responseData.issues.length; i++) {
        var issueData = responseData.issues[i];
        var issueKey = issueData.key;
        var issueSummary = issueData.fields.summary;

        var issue = issueKey + ":" + issueSummary;
        var hyperlink = '=HYPERLINK("' + jiraApiUrl + '/browse/' + issueKey + '"; "' + issue + '")';

        cell = String(2+i);
        sheet.getRange("A" + cell).setValue(hyperlink);
      }
    } else {
      Logger.log('No issues found');
    }

  } else {
    Logger.log("Sheet 'TEST' not found.");
  }
}

function queryJiraAPI() {
  var jqlQuery = 'assignee = currentUser() AND (status not in (Resolved, Archived, Done, Closed, Declined, Completed, Cancel, Canceled)) order by created DESC';

  // Basic authentication header
  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(jiraUsername + ':' + jiraApiToken),
    'Accept': 'application/json'
  };

  var params = {
    'method': 'get',
    'headers': headers,
    'muteHttpExceptions': true,
  };

  var apiUrlWithQuery = jiraApiUrl + '/rest/api/2/search?jql=' + encodeURIComponent(jqlQuery);

  // Make the request to the Jira API
  var response = UrlFetchApp.fetch(apiUrlWithQuery, params);

  // Parse the JSON response
  var responseData = JSON.parse(response.getContentText());

  return responseData;
}

function logWorkInJira(issueKey, timeSpent, comment, started) {
  // Set your Jira API URL
  var apiUrl = jiraApiUrl + '/rest/api/2/issue/' + issueKey + '/worklog';

  // Create the authentication header
  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(jiraUsername + ':' + jiraApiToken),
    'Content-Type': 'application/json',
  };

  // Set the worklog data
  var worklogData = {
    'timeSpent': timeSpent,
    'started': started,
    'comment': comment
  };

  // Set the Jira API parameters
  var params = {
    'method': 'post',
    'headers': headers,
    'muteHttpExceptions': true,
    'payload': JSON.stringify(worklogData),
  };

  // Make the request to log work in the Jira issue
  var response = UrlFetchApp.fetch(apiUrl, params);

  // Check if the worklog was successfully added
  if (response.getResponseCode() == 201) {
    Logger.log('Worklog added successfully for issue ' + issueKey);
    return 'Worklog added successfully for issue ' + issueKey;
  } else {
    Logger.log('Failed to add worklog. Response code: ' + response.getResponseCode());
    Logger.log('Response content: ' + response.getContentText());
    var responseError = 'Response code: ' + response.getResponseCode() + ' ' + response.getContentText();
    return responseError;
  }
}

////// Testing
function testSpreed() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Imputator");

  if (sheet) {
    var day = sheet.getRange("A2").getValue();
    var start = sheet.getRange("B2").getValue();
    var end = sheet.getRange("C2").getValue();

    var timeSpent = String((end - start)/(1000 * 60)) + "m";

    var issueKey = sheet.getRange("G2").getValue().split(":")[0];
    var comment = sheet.getRange("H2").getValue();

    var startedHour = Utilities.formatDate(start, "Europe/Madrid", "HH:mm:ss.SSS'+0100'")
    var started = Utilities.formatDate(day, "Europe/Madrid", "yyyy-MM-dd'T'") + startedHour;

    Logger.log("datetime: " + started + " endHour: " + timeSpent);
    //Logger.log("issueKey: " + issueKey + " comment: " + comment);

  } else {
    Logger.log("Sheet 'TEST' not found.");
  }
}
