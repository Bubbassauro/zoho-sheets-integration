/**
 * Script to import a Zoho Creator report to a Google Sheet
 * 
 * Required script properties:
 *  access_token
 *  owner_name
 */

/**
* Clears the current spreadsheet and imports data from a Zoho Creator report. 
* Note that all content is overwritten, including headers, which are built based on the object returned by the Zoho API.
*
* @param {string} name of the Zoho Creator project and view name, for example: 'event-staff/view/Event_Staffing_Report'
*/
function importZohoCreatorReport(viewName) {

  var scriptProperties = PropertiesService.getScriptProperties();
  var accessToken = scriptProperties.getProperty('access_token');
  var ownerName = scriptProperties.getProperty('owner_name');
  var ui = SpreadsheetApp.getUi();
  
  var url = "https://creator.zoho.com/api/json/" + viewName;
  var params = {'SingleLine': 'SingleLineValue',
                               "authtoken": accessToken,
                               "scope": "creatorapi",
                               "zc_ownername": ownerName,
                               "raw" : true
               };
  
  var headers = {"Content-type": "application/x-www-form-urlencoded","Accept": "text/plain"};

  var options =
      {
        "method"  : "POST",
        "payload" : params,   
        "followRedirects" : true,
        "muteHttpExceptions": true
      };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var text = response.getContentText();
  
  if (!text) {
    ui.alert('Zoho API returned no data.\nMake sure the access_token and owner_name are correctly specified in Project Properties.');
    return;
  }
  
  if (responseCode != 200) {
    displayZohoApiReturn(text);
    return;
  }

  var dataAll;
  try {
    dataAll = JSON.parse(text); 
  }
  catch(e) {
    var html = '<div>Error parsing Zoho API Json: ' + e + '</div>' + text;
    displayZohoApiReturn(html);
    return;
  }
  
  if (Object.keys(dataAll).length <= 0) {
    ui.alert('Unexpected data format returned by Zoho API.\n' + text);
    return;
  }
  
  var tableName = Object.keys(dataAll)[0];
  var table = dataAll[tableName];
  var rows = [];
  
  // Build header
  var header = [];
  var first = table[0];
  Object.keys(first).forEach(function(key,index) {
    header.push(key);
  });
  header.push("Updated");
  rows.push(header);
  
  // Read rows
  table.forEach(function(obj) {
    var row = [];
    Object.keys(obj).forEach(function(key,index) {
      row.push(obj[key]);
    });
    row.push(new Date());
    rows.push(row);
  });
  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  sheet.clearContents();
  
  var range = sheet.getRange(1, 1, table.length+1, header.length);
  range.setValues(rows);
}

function displayZohoApiReturn(html) {
  var html = HtmlService.createHtmlOutput(html)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Error');
}

function importEventStaff() {
  importZohoCreatorReport('event-staff/view/Event_Staffing_Report');
}

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menubuttons = [ {name: "Import Event Staff", functionName: "importEventStaff"}];
    ss.addMenu("Zoho Integration", menubuttons);
} 