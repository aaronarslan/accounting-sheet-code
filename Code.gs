function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Plaid / Admin')
      .addItem('Create Admin Dashboards for Users', 'createAdminDashboardsForUsers')
      .addItem('Send Plaid Link to Checked Users',  'sendPlaidLinkToCheckedUsers')
      .addToUi();
}

function createAdminDashboardsForUsers() {

  const ui            = SpreadsheetApp.getUi();
  const ctrlSs        = SpreadsheetApp.getActiveSpreadsheet();   // the Control‑Panel file
  const sheet         = ctrlSs.getActiveSheet();                 // assumes you’re on the Users sheet
  const data          = sheet.getDataRange().getValues();        // 0‑based array
  const HEADER_ROW    = 3;                                       // zero‑based → row 4 in Sheet

  // column indexes in Control‑Panel
  const COL_NAME      = 0;     // “Name”            (A)
  const COL_EMAIL     = 1;     // “Email”           (B)
  const COL_SS        = 2;     // “Spreadsheet:”    (C)

  if (data.length <= HEADER_ROW) {
    ui.alert('No data rows found under the headers (row 4).');
    return;
  }

  const templateId = '1m3Q1mcicj69cMfi6ij5zbZgMkGiYI0xcoOFOhSj-KwY';   // ★ your Admin template ID
  let   created    = 0;

  for (let r = HEADER_ROW + 1; r < data.length; r++) {
    const row      = data[r];
    const name     = (row[COL_NAME]  || '').toString().trim();
    const email    = (row[COL_EMAIL] || '').toString().trim();
    const ssLink   = (row[COL_SS]    || '').toString().trim();

    // we only act on rows that **have** a name + email and **lack** a sheet link
    if (!name || !email || ssLink) continue;

    /*───────────────────
      1.  Copy the template
    ───────────────────*/
    const title   = `[Admin] - Accounting Sheet for ${name} (${email})`;
    const newFile = DriveApp.getFileById(templateId).makeCopy(title);
    const newSs   = SpreadsheetApp.open(newFile);

    /*───────────────────
      2.  Fill‑in the new dashboard’s “Users” sheet
    ───────────────────*/
    const usersSheet = newSs.getSheetByName('Users');
    if (usersSheet) {

      // find the “User” and “Primary Email” columns in ROW 2
      const hdr        = usersSheet.getRange(2,1,1,usersSheet.getLastColumn()).getValues()[0];
      const colUser    = hdr.indexOf('User');           // zero‑based
      const colEmail   = hdr.indexOf('Primary Email');  // zero‑based

      if (colUser >= 0)  usersSheet.getRange(3, colUser  + 1).setValue(name);
      if (colEmail >= 0) usersSheet.getRange(3, colEmail + 1).setValue(email);

    } else {
      Logger.log(`⚠️  “Users” sheet missing in new dashboard "${title}" – skipped owner row.`);
    }

    /*───────────────────
      3.  Write the dashboard’s URL back to Control‑Panel (column C)
    ───────────────────*/
    sheet.getRange(r + 1, COL_SS + 1).setValue(newSs.getUrl());
    created++;
  }

  ui.alert(
    created
      ? `✅  Created ${created} new Admin dashboard${created > 1 ? 's' : ''}.`
      : 'Everyone already has a dashboard – nothing to create.'
  );
}

function sendPlaidLinkToCheckedUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet(); // Ensure you're on the correct sheet
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Assuming headers are in row 4
  var headerRow = 3; // Zero-based index (Row 4)
  var emailCol = 1; // Column B (Email), zero-based index
  var spreadsheetCol = 2; // Column C (Spreadsheet)
  var sendLinkCol = 3; // Column D (Send Link?)

  var webAppUrl = 'https://script.google.com/macros/s/AKfycbxl6UvRhR4hffc9jJfg1tqluIZK4R3BfnlyxWNdYC-zjAgvKgX92xmUFCK9Ndg04tH9/exec'; // Replace with your web app URL

  for (var i = headerRow + 1; i < data.length; i++) {
    var row = data[i];
    var sendLink = row[sendLinkCol];
    if (sendLink === true) { // Checkbox is checked
      var userEmail = row[emailCol];
      var spreadsheetUrl = row[spreadsheetCol];
      var spreadsheetId = extractIdFromUrl(spreadsheetUrl);

      // Generate unique link with parameters
      var link = webAppUrl + '?userEmail=' + encodeURIComponent(userEmail) + '&spreadsheetId=' + encodeURIComponent(spreadsheetId);

      // Send email to user
      var subject = 'Connect Your Bank Account via Plaid';
      var body = 'Hello,\n\nPlease click the following link to connect your bank account:\n\n' + link + '\n\nThank you!';
      MailApp.sendEmail(userEmail, subject, body);

      // Uncheck the checkbox
      sheet.getRange(i + 1, sendLinkCol + 1).setValue(false); // Adjust for 1-based index
    }
  }
}

function extractIdFromUrl(url) {
  var id = '';
  var regex = /\/d\/([a-zA-Z0-9-_]+)/;
  var match = regex.exec(url);
  if (match && match[1]) {
    id = match[1];
  }
  return id;
}


// Serve the HTML file
function doGet(e) {
  var userEmail = e.parameter.userEmail;
  var spreadsheetId = e.parameter.spreadsheetId;

  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = userEmail;
  template.spreadsheetId = spreadsheetId;

  return template.evaluate();
}

/**
 * Function to create a Plaid Link token, requiring only Auth & Transactions
 * but consenting to Assets, Balance, Enrich, Investments, Liabilities,
 * Recurring Transactions, and Transactions Refresh.
 */
function createLinkToken(userEmail) {
  const props          = PropertiesService.getScriptProperties();
  const PLAID_CLIENT_ID  = props.getProperty('PLAID_CLIENT_ID');
  const PLAID_SECRET     = props.getProperty('PLAID_SECRET');
  const PLAID_ENV        = 'production'; // or 'sandbox' / 'development'

  // Create a non-sensitive unique ID for this user
  const clientUserId = hashEmailAddress(userEmail);

  const url = `https://${PLAID_ENV}.plaid.com/link/token/create`;
  const payload = {
    client_id: PLAID_CLIENT_ID,
    secret: PLAID_SECRET,
    user: {
      client_user_id: clientUserId
    },
    client_name: 'Your App Name',
    // only these two are *required*…
    products: ['transactions'],
    additional_consented_products: [
      'auth',
      'liabilities',
      'investments'
    ],
    country_codes: ['US'],
    language: 'en'
    // no account_filters here so all accounts show up
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    const err = JSON.parse(response.getContentText());
    throw new Error('Error creating link token: ' + err.error_message);
  }
  return JSON.parse(response.getContentText()).link_token;
}

// Function to hash the email address
function hashEmailAddress(email) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email.trim());
  var hash = rawHash.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
  return hash;
}

// Function to exchange public_token for access_token and save to the user's admin spreadsheet
function getAccessToken(public_token, userEmail, spreadsheetId) {
  Logger.log('getAccessToken called with public_token: ' + public_token);
  Logger.log('Spreadsheet ID: ' + spreadsheetId);

  var properties = PropertiesService.getScriptProperties();
  var PLAID_CLIENT_ID = properties.getProperty('PLAID_CLIENT_ID');
  var PLAID_SECRET = properties.getProperty('PLAID_SECRET');
  var PLAID_ENV = 'production'; // or 'sandbox'/'development'

  var url = 'https://' + PLAID_ENV + '.plaid.com/item/public_token/exchange';
  var payload = {
    client_id: PLAID_CLIENT_ID,
    secret: PLAID_SECRET,
    public_token: public_token
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());

    // Log the data we got from Plaid
    Logger.log('Received from Plaid: ' + JSON.stringify(data, null, 2));

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheetName = 'Access & Item';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('Sheet "' + sheetName + '" not found, creating it now...');
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['Item ID', 'Access Token']);
    }

    // Log the item_id and access_token before writing
    Logger.log('Writing item_id: ' + data.item_id + ' and access_token: ' + data.access_token);

    sheet.appendRow([data.item_id, data.access_token]);
    Logger.log('Appended row to "' + sheetName + '" successfully.');

    return data;
  } catch (err) {
    Logger.log('Error in getAccessToken: ' + err.message);
    throw err;
  }
}