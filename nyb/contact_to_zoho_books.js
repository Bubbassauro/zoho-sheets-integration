/**
 * Scripts to send contact info from Google Sheet to Zoho Books
 * 
 * Required Script properties:
 *  Zoho Organization ID
 *  Zoho Token
 */

function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
      {
            name: 'How to Use',
            functionName: 'howToUse'
        },
      {
            name: 'Copy New Visitors',
            functionName: 'copyToIntroClass'
        },
        {
            name: 'Export Contact to Zoho Books',
            functionName: 'createContact'
        }
    ];
    spreadsheet.addMenu('New York Budokai', menuItems);
}

function howToUse() {
  
  var html = '<style>.nyb { font-family: Arial, Helvetica, sans-serif; font-size: small; }\n' +
    '.nyb h2 { font-size: large; }</style>' +
    '<div class="nyb">' +
    '<h2>Copy New Visitors</h2>' +
    '<p>To copy new visitors, place the cursor on a blank target row where you ' +
    'want to start including new records and click "Copy New Visitors"</p>' +
    '<p><strong>Notes:</strong></p>' +
    '<p>The script knows which records are new by ' +
    'looking at the class date (will copy records for classes in the future) and ' +
    'by comparing the e-mail address (if the e-mail is already in the Intro ' +
    'Class sheet, it will not be copied over).</p>' +
    '<p>Formatting and moving records around is still up to you.<br/>' +
    'You can manually modify the values after they are copied, as long as you ' +
    'keep the e-mail address, which is used as the identification for the ' +
    'record.</p>' +
    '<h2>Export Contact to Zoho Books</h2>' + 
    '<p>To export a contact to Zoho Books, select the row on "NYB Sign Up Information" ' +
    'and click "Export Contact to Zoho Books."' +
    '</div>';
    var htmlService = HtmlService.createHtmlOutput(html)
      .setTitle('How to Use')
      .setWidth(300);
  SpreadsheetApp.getUi() 
      .showSidebar(htmlService);
      
}
/**
 * arrayToObjectList - Convert a bi-dimentional DataRange array into an object list
 *  so you can access by column names like a human
 * Assumes that the first row in the values set has the original column names
 *
 * @param  {array}  values   two-dimentional array that you get with Range.getValues()
 * @param  {object} mappings to substitute the column names, for example:
 *  map = {"a hideous column-name!", "usableName", "anOthER BLErgh" : "another"};
 * @return {array}           object array with column names to use like this:
 *  var name = myObject[row].usableName;
 */
function arrayToObjectList(values, mappings) {

    var list = [];
    if (values.length <= 1) {
        return list;
    }
    for (var row = 1; row < values.length; row++) {
        var info = {};
        for (var col = 0; col < values[0].length; col++) {
            var colName = (mappings[values[0][col]] || values[0][col]);
            info[colName] = values[row][col];
        }
        list.push(info);
    }
    return list;
}

function addMonths(value, months) {
    var dt = new Date(value);
    return dt.setMonth(dt.getMonth() + months);
}

function formatMDY(value) {
    var dt = new Date(value);
    return dt.getMonth() + 1 + "/" + dt.getDate() + "/" + dt.getYear();
}

/**
 * getClassDate - Get the closest class date going back
 * class dates are 2 (Tuesday) and 4 (Thursday)
 *
 * @param  {object} value date to check if it's a class date
 * @return {Date}         class date closest to the provided date
 */
function getClassDate(value) {
    // try for 4 days at most
    var dt = new Date(value);
    var classDate;
    var ct = 1;
    while (dt.getDay() != 2 && dt.getDay != 4 && ct < 5) {
        classDate = dt.setDate(dt.getDate() - 1);
        ct++;
    }
    return (classDate || value);
}

/**
 * getPrice - read prices from another sheet
 *
 * @param  {string} promo       "Introduction Class", "One Month", etc
 * @param  {string} paymentType "Paypal", "Cash", "Groupon"
 * @return {float}              price
 */
function getPrice(promo, paymentType) {
    var prices = SpreadsheetApp.getActive().getRange("Prices").getValues();
    var paymentTypeIndex;
    for (var col = 1; col < prices[0].length; col++) {
        if (prices[0][col] == paymentType) {
            paymentTypeIndex = col;
            break;
        }
    }
    if (!paymentTypeIndex) {
        Logger.log("I can't find the payment type " + paymentType);
        return "";
    }
    for (var row = 1; row < prices.length; row++) {
        if (prices[row][0] == promo) {
            return prices[row][paymentTypeIndex];
        }
    }
    return "";
}

/**
 * appendInfoToContact - Append custom information to contact object
 *
 * @param  {type} contact original contact info
 * @return {type}         contact with additional fields
 */
function appendInfoToContact(contact) {
    // Add more information to contact
    contact.notes = contact.promotion;
    if (contact.classDay) {
        switch (contact.promotion) {
            case "One Month":
                contact.months = 1;
                break;
            case "Three Months":
                contact.months = 3;
                break;
            case "Observing":
                contact.months = "Observing";
                break;
            default:
                contact.months = "Intro";
                break;
        }

        contact.starts = formatMDY(contact.classDay);
        if (!isNaN(contact.months)) {
            contact.expires = formatMDY(getClassDate(addMonths(contact.classDay, contact.months)));
            contact.notes += " starts " + contact.starts + " expires " + contact.expires;
        } else {
            contact.expires = "";
            contact.notes += " " + contact.starts;
        }
    }
    return contact;
}


/**
 * getContacts - Read contacts from raw signup sheet
 *
 * @return {array}  array of contact objects
 */
function getContacts() {
    // This is where you map a column name from the raw spredsheet to a property
    var signup_map = {
        "Submitted": "submitted",
        "first-name": "firstName",
        "last-name": "lastName",
        "address": "address",
        "phone": "phone",
        "email": "email",
        "occupation": "occupation",
        "age": "age",
        "referer": "referer",
        "emergency-contact-name": "emergencyName",
        "emergency-contact-info": "emergencyPhone",
        "interest": "interest",
        "emergency-contact-relation": "designation",
        "expectations": "expectations",
        "emergency-contact-method": "contactMethod",
        "emergency-contact-method2": "contactMethod2",
        "experience": "experience",
        "promotion-type": "promotion",
        "class-day": "notUsed",
        "legal-waiver-agreement": "waiver",
        "sign-up-verification": "verification",
        "id:promotion-type": "promotion",
        "id:class-day": "classDay",
        "groupon-code": "groupon",
        "payment-option": "payment",
        "Submitted": "submitted",
        "Login": "login",
        "Submitted From": "ip"
    };

    var sourceSheet = "NYB Sign Up Information";
    var signup = SpreadsheetApp.getActive().getSheetByName(sourceSheet);
    var rawContacts = signup.getDataRange().getValues();
    var contacts = arrayToObjectList(rawContacts, signup_map);
    return contacts;
}

/**
 * createContact - Export a contact to Zoho Books
 *
 * @return {void} displays an alert at the end with response message from
 * the Zoho API
 */
function createContact() {
    const COMPANY = "New York Budokai";
    const DEFAULT_PAYMENT = 30;
    var sourceSheet = "NYB Sign Up Information";

    var ui = SpreadsheetApp.getUi();
    if (SpreadsheetApp.getActive().getActiveSheet().getSheetName() != sourceSheet) {
      ui.alert('Please click on a row to export in the "NYB Sign Up Information" sheet.');
      return;
    }

    var contacts = getContacts();
    // Use the selected row by selected row position in the raw signup information sheet
    var selectedRow = SpreadsheetApp.getActiveRange().getRow();
    var contact = appendInfoToContact(contacts[selectedRow - 2]); // -2 because array is zero based and header is removed from set

    // Format request with fields that are relevant for us
    // Api documentation:
    // https://www.zoho.com/books/api/v3/contacts/#create-a-contact
    var data = {
        "contact_name": contact.firstName + " " + contact.lastName,
        "company_name": COMPANY,
        "payment_terms": DEFAULT_PAYMENT,
        "billing_address": {
            "address": contact.address
        },
        "contact_persons": [{
                "first_name": contact.firstName,
                "last_name": contact.lastName,
                "email": contact.email,
                "phone": contact.phone,
                "is_primary_contact": true
            },
            {
                "first_name": contact.emergencyName,
                "phone": contact.emergencyPhone,
                "designation": contact.designation
            }
        ],
        "notes": contact.notes
    };

    var scriptProperties = PropertiesService.getScriptProperties();
    var token = scriptProperties.getProperty("Zoho Token");
    var organizationId = scriptProperties.getProperty("Zoho Organization ID");
    var url = "https://books.zoho.com/api/v3/contacts?authtoken=" + token + "&organization_id=" + organizationId;

    url += "&JSONString=" + escape(JSON.stringify(data));
    // Logger.log(url);

    var options = {
        'method': 'post',
        'muteHttpExceptions': true
    };

    try {
        var response = UrlFetchApp.fetch(url, options);
        var result = JSON.parse(response);
        ui.alert(result.message);
    } catch (e) {
        ui.alert("There was an unexpected error, check log for details");
        Logger.log(e);
    }
}

/**
 * isIn - Check if the e-mail is in the list
 *
 * @param  {array} values     list of e-mails (assumes that the e-mail is in col 0)
 * @param  {object} value     value to find
 * @return {bool}             true if found
 */
function isIn(values, email) {
    for (var i = 0; i < values.length; i++) {
        if (values[i][0] == email) {
            return true;
        }
    }
    return false;
}

var compareClassDate = function(info1, info2) {
  var date1 = new Date(info1.classDay);
  var date2 = new Date(info2.classDay);
  if (date1 < date2) {
    return -1;
  }
  if (date1 > date2) {
    return 1;
  }
  return 0;
};


/**
 * copyNewToIntroClass - Copy new contacts from signup sheet to intro class
 * starts at the active cell and adds cells down
 *
 * @return {void}
 */
function copyToIntroClass() {
    const EMAIL = 11;

    var ui = SpreadsheetApp.getUi();
    var intro = SpreadsheetApp.getActive().getActiveSheet();

    if (intro.getActiveCell().getValue()) {
      ui.alert('Please select a blank row to place new records.');
      return;
    }

    var targetRow = intro.getActiveCell().getRow();
    var emails = intro.getRange(2, EMAIL, intro.getLastRow() - 2).getValues();
    var contacts = getContacts().sort(compareClassDate);

    var today = new Date();
    var visitorCount = 0;
    for (var row = 0; row < contacts.length; row++) {
      if (!isIn(emails, contacts[row].email) && contacts[row].classDay >= today) {
        var contact = appendInfoToContact(contacts[row]);
        var record = [formatIntroClassRow(contact)];
        intro.getRange(targetRow, 1, 1, 12).setValues(record);
        intro.insertRowAfter(targetRow);
        targetRow++;
        visitorCount++;
      }
    }

    if (visitorCount) {
      ui.alert(visitorCount + ' new visitor(s) copied.');
    }
    else {
      ui.alert('There are no new visitors to copy.');
    }
}

function formatIntroClassRow(contact) {
    const NAME = 0;
    const CASH_CHECK = 1;
    const PAYPAL = 2;
    const GROUPON = 3;
    const MONTHS = 4;
    const ATTENDANCE = 5;
    const STARTED = 6;
    const EXPIRES = 7;
    const NOTES = 8;
    const PHONE = 9;
    const EMAIL = 10;
    const SUBMITTED = 11;

    // initialize with empty values
    var record = Array(12);
    var size = 12;
    while(size--) record[size] = "";

    record[NAME] = contact.firstName + " " + contact.lastName;
    record[MONTHS] = contact.months;
    record[STARTED] = contact.starts;
    record[EXPIRES] = contact.expires;
    record[PHONE] = contact.phone.toString();
    record[EMAIL] = contact.email;
    record[SUBMITTED] = contact.submitted;

    if (contact.payment == "Checkout now with") {
        record[PAYPAL] = getPrice(contact.promotion, "Paypal");
    } else if (contact.groupon) {
        record[GROUPON] = getPrice(contact.promotion, "Groupon");
    } else {
        record[CASH_CHECK] = getPrice(contact.promotion, "Cash");
    }
    return record;
}
