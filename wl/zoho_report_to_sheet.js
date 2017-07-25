/**
 * Import a report from Zoho Reports to a Google Sheet
 * 
 * Required Script properties:
 *  reports_authtoken : See https://www.zoho.com/reports/api/#auth-token
 * 
 * owner_login : user name (e-mail)
 * database_id : number that identifies the Zoho Reports database
 * object_id   : number that identifies the Zoho Report
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
      {
            name: 'Refresh Report',
            functionName: 'getPipelineReport'
        }
    ];
    spreadsheet.addMenu('Wanderlust', menuItems);
  getPipelineReport();
}

/**
 * Get data from a Zoho Report
 * You must configure the script property reports_authtoken in File > Project properties
 */
function getZohoReport(login_email, database_id, object_id) {
  var projectProperties = PropertiesService.getScriptProperties();
  var authtoken = projectProperties.getProperty('reports_authtoken');
  var url = 'https://reportsapi.zoho.com/api/' + login_email + '?DBID=' + database_id + '&OBJID=' + object_id + '&ZOHO_ACTION=EXPORT&ZOHO_OUTPUT_FORMAT=JSON&ZOHO_API_VERSION=1.0&authtoken=' + authtoken;
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText().replace(/\\/gi, '');
  var data = JSON.parse(json);
  return data;
}

/**
 * 
 */
function getPipelineReport() {
  var projectProperties = PropertiesService.getScriptProperties();
  var owner_login = projectProperties.getProperty('owner_login');
  var database_id = projectProperties.getProperty('database_id');
  var object_id = projectProperties.getProperty('object_id');

  var data = getZohoReport(owner_login, database_id, object_id);
  var number_of_rows = data.response.result.rows.length;
  var number_of_columns = data.response.result.column_order.length;
  var active_sheet = SpreadsheetApp.getActiveSheet();
  active_sheet.clearContents();
  // Populate headers
  var header_range = active_sheet.getRange(1, 1, 1, number_of_columns);
  var headers = [];
  headers.push(data.response.result.column_order);
  header_range.setValues(headers);
  // Populate report rows
  var range = active_sheet.getRange(2, 1, number_of_rows, number_of_columns);
  range.setValues(data.response.result.rows);
}

