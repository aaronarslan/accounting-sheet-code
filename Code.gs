//INITIALIZING CODE:

function initializeDashboard() {

    // 1) store our Admin ID in script properties
  storeAdminSheetId();

  // 0) For each user in our Admin “Users” sheet, write a Self‐owned entity
  const adminSs    = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (usersSheet) {
    const headerRow = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
    const nameCol   = 1; // column A
    const dashCol   = headerRow.indexOf("Dashboard") + 1;
    const lastUser  = usersSheet.getLastRow();
    for (let r = 3; r <= lastUser; r++) {
      const userName     = usersSheet.getRange(r, nameCol).getValue();
      const dashUrl      = usersSheet.getRange(r, dashCol).getValue();
      const dashboardId  = extractIdFromUrl(dashUrl);
      if (!dashboardId || !userName) continue;
      try {
        const userSs       = SpreadsheetApp.openById(dashboardId);
        const entitiesSheet= userSs.getSheetByName("Entities");
        if (!entitiesSheet) continue;

        // find our columns in row 2
        const entHdrs = entitiesSheet
          .getRange(2, 1, 1, entitiesSheet.getLastColumn())
          .getValues()[0];
        const entNameCol = entHdrs.indexOf("Entity Name:") + 1;
        const entTypeCol = entHdrs.indexOf("Entity Type:") + 1;
        const ownPctCol  = entHdrs.indexOf("Ownership Percentage:") + 1;
        if (entNameCol < 1 || entTypeCol < 1 || ownPctCol < 1) continue;

        // append a new row at the bottom
        const writeRow = entitiesSheet.getLastRow() + 1;
        entitiesSheet.getRange(writeRow, entNameCol).setValue(userName);
        entitiesSheet.getRange(writeRow, entTypeCol).setValue("Self");
        entitiesSheet.getRange(writeRow, ownPctCol ).setValue("100%");
      } catch (err) {
        Logger.log(`initializeDashboard → Entities seed failed for row ${r}: ${err}`);
      }
    }
  }

  // 2) create each user’s dashboard + form folder & files
  createUserDocuments();

  // 3) install all the onEdit triggers into each user’s dashboard
  installOnEditTriggersForAllUsers();

  // 4) install all the onFormSubmit triggers into each user’s form
  installOnFormSubmitTriggersForAllUsers();

  // 5) install any triggers you use for Access & Item sheet
  installAccessItemTriggers();

  // 6) create+link legal‐docs folders for each user’s Entities
  createAndLinkLegalDocsFolders();

  fetchAndStoreAllData();
}

function storeAdminSheetId() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty('ADMIN_SHEET_ID', id);
}

function getAdminSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('ADMIN_SHEET_ID');
  return SpreadsheetApp.openById(id);
}

function createUserDocuments() {
  const ss = getAdminSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const usersRange = usersSheet.getRange('A3:F' + usersSheet.getLastRow());
  const usersData = usersRange.getValues();

  const parentFolderId = '1_ybybfE09h3n8w9dOE2xQcHxwq7dWgU8';
  const dashboardTemplateId = '1_sWg3AXMKwtcbjlQWaejFWGAm8LrlwQjgtPpNl9267c';
  const transactionFormTemplateId = '1CNGY8gWHfwxMlqYm3d-p6l7TOuhs9HyGPE-NgceUszM';

  usersData.forEach((row, index) => {
    const userName = row[0];
    const dashboardUrl = row[1];
    const transactionFormUrl = row[2];
    // Check if the user name is not empty and documents haven't been created yet (URLs are empty)
    if (userName && !dashboardUrl && !transactionFormUrl) {
      const userFolder = DriveApp.getFolderById(parentFolderId).createFolder(`${userName} - Transaction Form Docs`);

      // Create Dashboard
      const dashboardCopy = DriveApp.getFileById(dashboardTemplateId).makeCopy(`${userName} - Accounting Sheet by Aaron Arslan`, userFolder);
      const newDashboardUrl = dashboardCopy.getUrl();
      usersSheet.getRange(index + 3, 2).setValue(newDashboardUrl); // B column for Dashboard URL

      // Create Transaction Form
      const formCopy = DriveApp.getFileById(transactionFormTemplateId).makeCopy(`${userName} - Transaction Form`, userFolder);
      const form = FormApp.openById(formCopy.getId());
      form.setTitle(`${userName} - Transaction Form`);
      const formEditUrl = form.getEditUrl();

      // Ensure the form has a response destination set
      let formResponseSheet;
      try {
        const formResponsesId = form.getDestinationId();
        if (formResponsesId) {
          formResponseSheet = SpreadsheetApp.openById(formResponsesId);
        } else {
          // If no destination is set, create a new one
          const newFormResponseSheet = SpreadsheetApp.create(`${userName} - Transaction Form Response`);
          form.setDestination(FormApp.DestinationType.SPREADSHEET, newFormResponseSheet.getId());
          formResponseSheet = newFormResponseSheet;
        }
      } catch (e) {
        // In case of any error, create a new spreadsheet for responses
        const newFormResponseSheet = SpreadsheetApp.create(`${userName} - Transaction Form Response`);
        form.setDestination(FormApp.DestinationType.SPREADSHEET, newFormResponseSheet.getId());
        formResponseSheet = newFormResponseSheet;
      }

      formResponseSheet.rename(`${userName} - Transaction Form Response`);
      const formResponseSheetUrl = formResponseSheet.getUrl();

      usersSheet.getRange(index + 3, 3).setValue(formEditUrl); // C column for Transaction Form Edit URL
      usersSheet.getRange(index + 3, 4).setValue(formResponseSheetUrl); // D column for Form Response Sheet URL
    }
  });
}

/**
 * Loop through "Users" sheet in this Admin Dashboard,
 * read each user's spreadsheet URL or ID,
 * and install an onEdit trigger that calls handleUserEdit(e).
 */
function installOnEditTriggersForAllUsers() {
  // 1) Open "Users" sheet in the Admin Dashboard (the same spreadsheet that has this script)
  const adminSs = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found => aborting.');
    return;
  }

  // 2) Identify the column that has the user’s spreadsheet URL/ID
  //    For example, row 2 might have "Dashboard" or "UserSheet" or something.
  const headerRow = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol = headerRow.indexOf("Dashboard"); // or "UserSheet"
  if (dashCol < 0) {
    Logger.log('No "Dashboard" column found => aborting.');
    return;
  }

  // 3) Loop rows in "Users" from row 3 downward
  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('No user rows => nothing to do.');
    return;
  }
  const userData = usersSheet.getRange(3, dashCol + 1, lastRow - 2, 1).getValues();

  userData.forEach((rowVal, i) => {
    const dashUrlOrId = rowVal[0];
    if (!dashUrlOrId) {
      Logger.log(`Row ${i+3}: blank => skipping`);
      return;
    }

    // 4) If it's a full URL, parse out the ID
    let userSheetId = extractIdFromUrl(dashUrlOrId);
    if (!userSheetId) {
      // maybe it's already just an ID
      userSheetId = dashUrlOrId;
    }

    Logger.log(`Row ${i+3}: Installing onEdit trigger for userSheetId=${userSheetId}`);
    installUserOnEditTrigger(userSheetId);
  });

  Logger.log('Done installing onEdit triggers for all users.');
}

/**
 * Installs or re-installs an onEdit installable trigger on a given user's sheet,
 * pointing back to the function handleUserEdit(e) in this Admin script.
 * 
 * @param {string} userSheetId ID (or possibly URL) of the user’s spreadsheet
 */
function installUserOnEditTrigger(userSheetId) {
  // 1) Delete any existing onEdit triggers for that same sheet & function
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (
      trigger.getHandlerFunction() === 'handleUserEdit' &&
      trigger.getTriggerSourceId() === userSheetId &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT
    ) {
      Logger.log(`Removing old onEdit trigger for sheet ID: ${userSheetId}`);
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 2) Create a new onEdit trigger
  ScriptApp.newTrigger('handleUserEdit')
    .forSpreadsheet(userSheetId)
    .onEdit()
    .create();

  Logger.log(`Installed a new onEdit trigger for user sheet ID=${userSheetId}`);
}

/**
 * Installs an onFormSubmit trigger for each user’s 
 * Form Response sheet – as listed in row 2 of "Users" 
 * under the "Form Response" column.
 */
function installOnFormSubmitTriggersForAllUsers() {
  const adminSs = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found in Admin Dashboard => Aborting.');
    return;
  }

  // Identify the "Form Response" column in row 2
  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("No user rows found => nothing to do.");
    return;
  }
  const headerRow = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const formRespCol = headerRow.indexOf("Form Response");
  if (formRespCol < 0) {
    Logger.log('No "Form Response" column found in row 2 => Aborting.');
    return;
  }

  // Loop each user row in "Users"
  const userData = usersSheet.getRange(3, formRespCol + 1, lastRow - 2, 1).getValues();
  userData.forEach((rowVal, i) => {
    const formRespUrlOrId = rowVal[0];
    if (!formRespUrlOrId) {
      Logger.log(`Row ${i+3}: blank => skipping.`);
      return;
    }

    // Attempt to parse the ID from a full URL, or just use it if it's already an ID
    let formRespId = extractIdFromUrl(formRespUrlOrId);
    if (!formRespId) {
      formRespId = formRespUrlOrId; // assume it's directly an ID
    }

    Logger.log(`Row ${i+3}: Installing onFormSubmit trigger for ID="${formRespId}"`);
    installUserOnFormSubmitTrigger(formRespId);
  });

  Logger.log("Done installing onFormSubmit triggers for all users.");
}
/**
 * Installs or re-installs an onFormSubmit trigger for the specified 
 * "Form Response" sheet, removing any old ones for the same function name.
 *
 * @param {string} formRespSheetId The ID of the form response spreadsheet
 */
function installUserOnFormSubmitTrigger(formRespSheetId) {
  // 1) Delete any existing onFormSubmit triggers for this sheet & function
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (
      trigger.getHandlerFunction() === 'handleUserFormSubmit' &&
      trigger.getTriggerSourceId() === formRespSheetId &&
      trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT
    ) {
      Logger.log(`Removing old onFormSubmit trigger for sheet ID=${formRespSheetId}`);
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 2) Create a new onFormSubmit trigger
  ScriptApp.newTrigger('handleUserFormSubmit')
    .forSpreadsheet(formRespSheetId)
    .onFormSubmit()
    .create();

  Logger.log(`Installed a new onFormSubmit trigger for ID="${formRespSheetId}"`);
}

/**
 * Run *once* from the Script‑editor ►▶ Run menu.  
 * Creates/re‑creates **both** an onEdit and an onChange trigger that
 * call handleAccessItemEvent(e) whenever *anything* changes in this
 * Admin spreadsheet.
 */
function installAccessItemTriggers() {
  const adminId = getAdminSpreadsheet().getId();

  // ── remove any stale copies ─────────────────────────────────────
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn   = t.getHandlerFunction();
    const src  = t.getTriggerSourceId();
    const type = t.getEventType();
    if (fn === 'handleAccessItemEvent' &&
        src === adminId &&
        (type === ScriptApp.EventType.ON_EDIT ||
         type === ScriptApp.EventType.ON_CHANGE)) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // ── install fresh ones ──────────────────────────────────────────
  ScriptApp.newTrigger('handleAccessItemEvent')
           .forSpreadsheet(adminId)
           .onEdit()                 // manual UI edits
           .create();

  ScriptApp.newTrigger('handleAccessItemEvent')
           .forSpreadsheet(adminId)
           .onChange()               // script‑driven writes, row inserts, etc.
           .create();

  Logger.log('✔ onEdit + onChange triggers installed for “Access & Item”.');
}

/**
 * Loop through each user’s Dashboard ID in the Admin “Users” sheet
 * and install a fetchAndStoreAllData onOpen trigger on it.
 */
function installOnOpenTriggersForAllUsers() {
  const adminSs    = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) return Logger.log('No "Users" sheet found.');

  // find the “Dashboard” column in row 2
  const hdrs    = usersSheet.getRange(2,1,1,usersSheet.getLastColumn()).getValues()[0];
  const dashCol = hdrs.indexOf("Dashboard") + 1;
  if (!dashCol) return Logger.log('No "Dashboard" column in row 2.');

  // for each user row…
  const lastRow = usersSheet.getLastRow();
  for (let r = 3; r <= lastRow; r++) {
    const url = usersSheet.getRange(r, dashCol).getValue();
    if (!url) continue;
    const sheetId = extractIdFromUrl(url) || url;
    Logger.log(`installOnOpen ▶ adding onOpen trigger for ${sheetId}`);
    installUserOnOpenTrigger(sheetId);
  }
}

/**
 * Deletes any old onOpen → fetchAndStoreAllData triggers
 * for that sheet, then creates a fresh one.
 */
function installUserOnOpenTrigger(userSheetId) {
  // 1) remove any existing onOpen/fetchAndStoreAllData for this sheet
  ScriptApp.getProjectTriggers().forEach(t => {
    if (
      t.getHandlerFunction() === "fetchAndStoreAllData" &&
      t.getEventType() === ScriptApp.EventType.ON_OPEN &&
      t.getTriggerSourceId() === userSheetId
    ) {
      ScriptApp.deleteTrigger(t);
      Logger.log(`  ↪ removed stale onOpen trigger for ${userSheetId}`);
    }
  });

  // 2) create the new onOpen trigger
  ScriptApp.newTrigger("fetchAndStoreAllData")
           .forSpreadsheet(userSheetId)
           .onOpen()
           .create();
  Logger.log(`  ✔ installed onOpen trigger for ${userSheetId}`);
}


/**
 * Utility to parse out the doc/spreadsheet ID from a typical Google URL
 * (Sheets, Forms, Drive). If not found, returns null.
 */
function extractIdFromUrl(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (let r of patterns) {
    const match = r.exec(url);
    if (match && match[1]) return match[1];
  }
  return null;
}

// END INITIALIZING CODE

/**
 * Sets up a custom menu when the Admin spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Admin Tools')
    .addItem('Share Dashboards & Forms', 'shareDashboardAndFormWithUsers')
    .addToUi();
}


/**
 * Share each user’s dashboard (as editor) and form (as viewer),
 * then email them the links + mobile-savings tips.
 */
function shareDashboardAndFormWithUsers() {
  const ss         = getAdminSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    throw new Error('Users sheet not found.');
  }

  // header row is row 2
  const hdr        = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol    = hdr.indexOf('Dashboard') + 1;
  const formCol    = hdr.indexOf('Transaction Form') + 1;
  const emailCol   = hdr.indexOf('Primary Email') + 1;
  if (dashCol < 1 || formCol < 1 || emailCol < 1) {
    throw new Error('Make sure “Dashboard”, “Transaction Form” and “Primary Email” headers exist in row 2.');
  }

  const lastRow = usersSheet.getLastRow();
  for (let r = 3; r <= lastRow; r++) {
    const dashUrl = usersSheet.getRange(r, dashCol   ).getValue();
    const formUrl = usersSheet.getRange(r, formCol   ).getValue();
    const email   = usersSheet.getRange(r, emailCol  ).getValue();
    if (!dashUrl || !formUrl || !email) continue;

    const dashId = extractIdFromUrl(dashUrl);
    const formId = extractIdFromUrl(formUrl);
    if (!dashId || !formId) {
      Logger.log(`Row ${r}: could not parse IDs — skipping.`);
      continue;
    }

    // 1) Share the dashboard as Editor
    try {
      DriveApp.getFileById(dashId).addEditor(email);
    } catch (e) {
      Logger.log(`Row ${r}: failed to share dashboard → ${e}`);
    }

    // 2) Share the form as Viewer
    let viewUrl = formUrl;
    try {
      // give them view-only on the Form file
      DriveApp.getFileById(formId).addViewer(email);
      // get the “fill-in” URL (ends in /viewform)
      const form = FormApp.openById(formId);
      viewUrl = form.getPublishedUrl();
    } catch (e) {
      Logger.log(`Row ${r}: failed to share form → ${e}`);
    }

    // 3) Email them the links + mobile tips
    const subject = 'Your new Dashboard & Transaction Form';
    const body = `
Hi there,

You’ve been granted access to your personal accounting dashboard and form.

• Dashboard (edit): ${dashUrl}
• Form     (fill): ${viewUrl}

For best experience, open the form on your phone and “Add to Home Screen”:

• iPhone (Safari): Tap the Share icon (⬆️), then “Add to Home Screen.”  
• Android (Chrome): Tap the ⋮ menu, then “Add to Home screen.”

Let me know if you have any trouble!

— Aaron
`.trim();

    try {
      MailApp.sendEmail(email, subject, body);
    } catch (e) {
      Logger.log(`Row ${r}: failed to send email to ${email} → ${e}`);
    }
  }

  Logger.log('Dashboard & Form sharing complete.');
}

/**
 * onFormSubmit → runs every time a user submits their Transaction Form.
 *     1)  Import the row into the user’s “Financial Journal”.
 *     2)  Run the Who‑Owes‑Who scenarios (your existing logic).
 *     3)  Sync Plaid (all items).                ← NEW
 *     4)  If that sync produced *any* new Plaid transactions,
 *         merge Manual ↔ Plaid across all users. ← NEW
 */
function handleUserFormSubmit(e) {
  // grab a global lock so we never race two submissions
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // up to 30s

  try {
    Logger.log("handleUserFormSubmit ➜ new form submission");
    Logger.log(JSON.stringify(e.values));

    // 1) bring that one new form row into the financial journal
    importToFinancialJournal();

    // 2) run all your Who‑Owes‑Who scenarios
    whoOwsWhoAllScenarios();

    // 3) pull in every user's Plaid data now:
    //    fetchAndStoreAllData() should return { newTransactionRows: N, … }
    const syncResult = fetchAndStoreAllData();
    const newTxns    = syncResult?.newTransactionRows || 0;
    Logger.log(`Plaid sync complete – ${newTxns} new transaction(s)`);

    // 4) merge only if we actually got new Plaid transactions
    if (newTxns > 0) {
      Logger.log("↳ Detected new Plaid transactions – merging manual vs. Plaid rows");
      bulkMergeManualAndPlaidAllUsers();
    }

  } catch (err) {
    Logger.log(`handleUserFormSubmit ➜ ERROR: ${err}`);
  } finally {
    // always release the lock
    lock.releaseLock();
  }
}


/**
 * Master event handler for onEdit triggers in the user's sheet.
 * This code lives in your Admin script. 
 * It's called via an installable onEdit trigger that you set up 
 * with ScriptApp.newTrigger(...).forSpreadsheet(userSheetId).onEdit().
 *
 * NOTE: This function does NOT write changes back to the user's sheet 
 * (based on your instructions). We just detect the edit and call 
 * the relevant function in the Admin script.
 */
function handleUserEdit(e) {
  // Acquire the global script lock
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait up to 30s for the lock

  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const newValue = e.value;

    Logger.log(`handleUserEdit => sheet="${sheetName}", R${row}C${col}, newVal="${newValue}"`);

    // === [ Your existing conditions below ] ===

    // 1) "Entities" sheet => new "Entity Name:" => createAndLinkLegalDocsFolders()
    if (sheetName === "Entities" && row >= 3) {
      const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      const entityNameCol = row2.indexOf("Entity Name:");
      if (entityNameCol >= 0 && (col === (entityNameCol + 1))) {
        Logger.log("=> Detected new/edited entity => createAndLinkLegalDocsFolders()");
        createAndLinkLegalDocsFolders();
      }
    }

    // 2) "Accounts" sheet => nickname changed => addNicknamesToTransactionsData()
    if (sheetName === "Accounts" && row >= 3) {
      const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      const nicknameCol = row2.indexOf("Nickname:");
      if (nicknameCol >= 0 && col === (nicknameCol + 1)) {
        Logger.log("=> Nickname changed => addNicknamesToTransactionsData()");
        addNicknamesToTransactionsData();
        transferPlaidTransactionsToUserDashboards_Optimized();
      }
    }

    // 3) "Recurring Transactions"
    if (sheetName === "Recurring Transactions") {
      const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      // (a) "✔" => clearRecurringMergedEntries()
      const checkColIndex = row2.indexOf("✔");
      if (checkColIndex >= 0 && col === (checkColIndex + 1)) {
        const isChecked = sheet.getRange(row, col).getValue() === true;
        if (isChecked) {
          Logger.log("=> Recurring '✔' => clearRecurringMergedEntries()");
          clearRecurringMergedEntries();
        }
      }

      // (b) if row >= 3 and freq/start/from/to/amt are filled => mergeRecurring
      if (row >= 3) {
        const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        const freqCol  = row2.indexOf("Frequency");
        const startCol = row2.indexOf("Start Date");
        const fromCol  = row2.indexOf("From");
        const toCol    = row2.indexOf("To");
        const amtCol   = row2.indexOf("Amount");

        if (freqCol>=0 && startCol>=0 && fromCol>=0 && toCol>=0 && amtCol>=0) {
          const freqVal  = rowValues[freqCol];
          const startVal = rowValues[startCol];
          const fromVal  = rowValues[fromCol];
          const toVal    = rowValues[toCol];
          const amtVal   = rowValues[amtCol];
          if (freqVal && startVal && fromVal && toVal && amtVal) {
            Logger.log("=> Recurring row fully filled => mergeRecurringTransactionsToFinancialJournal()");
            mergeRecurringTransactionsToFinancialJournal();
          }
        }
      }
    }

    // 4) "Financial Journal" => merges or sync
    if (sheetName === "Financial Journal") {
      const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      const mergeColIndex = row2.indexOf("Merge");
      if (mergeColIndex >= 0 && col === (mergeColIndex + 1)) {
        Logger.log("=> 'Merge' changed => merges + sync");
        bulkMergeManualAndPlaidAllUsers();
        syncMergeColBasedOnAccountsForAllUsers();
      }

      // If user edits "From" or "To", sync
      const fromIdx = row2.indexOf("From");
      const toIdx   = row2.indexOf("To");
      if ((fromIdx >= 0 && col === (fromIdx + 1)) || 
          (toIdx >= 0   && col === (toIdx + 1))) {
        Logger.log("=> 'From'/'To' changed => syncMergeColBasedOnAccountsForAllUsers()");
        syncMergeColBasedOnAccountsForAllUsers();
      }
    }

    // 5) "Merged" => if user checks "Un-Merge?" => bulkUnMergeAllCheckedRows()
    if (sheetName === "Merged") {
      const row3 = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
      const unmergeColIndex = row3.indexOf("Un-Merge?");
      if (row >= 4 && unmergeColIndex >= 0 && col === (unmergeColIndex + 1)) {
        const isChecked = sheet.getRange(row, col).getValue() === true;
        if (isChecked) {
          Logger.log("=> 'Un-Merge?' checkbox => bulkUnMergeAllCheckedRows()");
          bulkUnMergeAllCheckedRows();
        }
      }
    }

    // 6) If the edit overlaps named ranges => updateUsersTransactionForms()
    const namedRangesToWatch = [
      "AccountNicknames",
      "Categories",
      "Subcategories",
      "EntitiesList",
      "Tags"
    ];
    let runUpdateForms = false;
    for (let nrName of namedRangesToWatch) {
      const nr = e.source.getRangeByName(nrName);
      if (!nr) continue;
      if (e.range.getSheet().getName() === nr.getSheet().getName()) {
        if (rangesOverlap(e.range, nr)) {
          runUpdateForms = true;
          break;
        }
      }
    }
    if (runUpdateForms) {
      Logger.log("=> Named range changed => updateUsersTransactionForms()");
      updateUsersTransactionForms();
    }

    // 7) "Who Owes Who?" => Payment/Date/Action => processWhoOwesWhoPaymentsAllUsers()
    if (sheetName === "Who Owes Who?") {
      const row3 = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
      const paymentIdx = row3.indexOf("Payment");
      const dateIdx    = row3.indexOf("Date");
      const actionIdx  = row3.indexOf("Action");
      if (row >= 4) {
        if ((paymentIdx >= 0 && col === (paymentIdx + 1)) ||
            (dateIdx >= 0    && col === (dateIdx + 1))    ||
            (actionIdx >= 0  && col === (actionIdx + 1))) {
          Logger.log("=> 'Who Owes Who?' Payment/Date/Action => processWhoOwesWhoPaymentsAllUsers()");
          processWhoOwesWhoPaymentsAllUsers();
        }
      }
    }

  } catch (err) {
    Logger.log(`handleUserEdit => Error: ${err}`);
  } finally {
    // Release the lock so subsequent triggers can proceed
    lock.releaseLock();
  }
}

/**
 * Unified event handler for *either* trigger.
 *   • Scans "Access & Item" for rows that have Item ID + Access Token
 *     but no “Auth Done (ISO)” timestamp yet.
 *   • If at least one found, runs fetchAndStoreAllData().
 *   • Stamps the “Auth Done (ISO)” cell **immediately after** Auth succeeds
 *     (that timestamping happens inside fetchAndStore… already).
 *
 * Works no matter *who* wrote the row: a human, a Form, or another script.
 */
function handleAccessItemEvent(e) {
  const ss    = getAdminSpreadsheet();
  const sheet = ss.getSheetByName('Access & Item');
  if (!sheet) return;

  /* -------- locate columns once -------------------------------- */
  const hdr      = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colItem  = hdr.indexOf('Item ID')           + 1;
  const colToken = hdr.indexOf('Access Token')      + 1;
  const colAuth  = hdr.indexOf('Auth Done (ISO)')   + 1;  // added earlier
  if (colItem === 0 || colToken === 0 || colAuth === 0) {
    Logger.log('handleAccessItemEvent ▶ required columns missing'); 
    return;
  }

  /* -------- is there any *new* Item waiting to be processed? ---- */
  const rng   = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const rows  = rng.getValues();
  const freshRows = rows.filter(r =>
      r[colItem  - 1] &&           // has Item ID
      r[colToken - 1] &&           // has Access Token
      !r[colAuth - 1]);            // but no Auth stamp yet

  if (freshRows.length === 0) return;  // nothing to do

  /* -------- run the big sync, protected by a script lock -------- */
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) { Logger.log('Sync already running – skipping'); return; }

  try {
    Logger.log(`Detected ${freshRows.length} new Plaid item(s) ➜ syncing`);
    fetchAndStoreAllData();    // **your main function**
  } finally {
    lock.releaseLock();
  }
}


/**
 * Helper to see if two ranges overlap.
 * e.g. if e.range is within a named range. 
 */
function rangesOverlap(r1, r2) {
  const r1RowStart = r1.getRow();
  const r1RowEnd   = r1RowStart + r1.getNumRows() - 1;
  const r1ColStart = r1.getColumn();
  const r1ColEnd   = r1ColStart + r1.getNumColumns() - 1;

  const r2RowStart = r2.getRow();
  const r2RowEnd   = r2RowStart + r2.getNumRows() - 1;
  const r2ColStart = r2.getColumn();
  const r2ColEnd   = r2ColStart + r2.getNumColumns() - 1;

  // Overlap if row ranges intersect AND column ranges intersect
  const rowsOverlap = !(r1RowEnd < r2RowStart || r1RowStart > r2RowEnd);
  const colsOverlap = !(r1ColEnd < r2ColStart || r1ColStart > r2ColEnd);

  return rowsOverlap && colsOverlap;
}



































/*

  addNicknamesToTransactionsData();
  updateUserDashboardsAccountDropdowns();
  transferPlaidTransactionsToUserDashboards_Optimized();
  discoverRecurringTransactions();

  */



// Fetch and store Plaid data into Admin Dashboard

/**
 * 1) Fetches /auth, /balance, /transactions for every token row
 * 2) Writes balances out to “Balance Data”
 * 3) Appends only brand-new, de-duplicated transactions to “Transactions Data”
 * 4) Calls downstream routines
 */
function fetchAndStoreAllData() {
  const ss    = getAdminSpreadsheet();
  const sheet = ss.getSheetByName('Access & Item');
  if (!sheet) {
    Logger.log('ERROR ▶ "Access & Item" sheet not found.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('No tokens to process.');
    return;
  }

  const headers     = data[0];
  const tokCol      = headers.indexOf('Access Token');
  const authFlagCol = headers.indexOf('Auth Pulled?');
  const balFlagCol  = headers.indexOf('Balance Pulled?');
  if (tokCol < 0 || authFlagCol < 0 || balFlagCol < 0) {
    throw new Error('Missing one of: "Access Token", "Auth Pulled?", "Balance Pulled?" headers.');
  }

  let allAch          = [];
  let allNewAccounts  = [];
  let allTransactions = [];
  let successCount    = 0;
  let failCount       = 0;

  //  ── gather everything ─────────────────────────────────────────
  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const token = (row[tokCol] || '').toString().trim();
    if (!token) continue;
    Logger.log(`Processing token row ${i+1}`);

    if (!row[authFlagCol]) {
      try {
        const auth = fetchAuthData(token);
        if (auth.numbers?.ach) allAch.push(...auth.numbers.ach);
        sheet.getRange(i+1, authFlagCol+1).setValue(true);
      } catch (e) {
        if (!e.message.includes('NO_AUTH_ACCOUNTS')) {
          Logger.log(`  ✖ Auth error at row ${i+1}: ${e.message}`);
        }
      }
    }

    if (!row[balFlagCol]) {
      try {
        const bal = fetchBalanceData(token);
        if (bal.accounts) allNewAccounts.push(...bal.accounts);
        sheet.getRange(i+1, balFlagCol+1).setValue(true);
      } catch (e) {
        Logger.log(`  ✖ Balance error at row ${i+1}: ${e.message}`);
        failCount++;
        continue;
      }
    }

    try {
      const tx = fetchAllTransactions(token);
      if (tx.transactions) allTransactions.push(...tx.transactions);
      successCount++;
    } catch (e) {
      Logger.log(`  ✖ Transactions error at row ${i+1}: ${e.message}`);
      failCount++;
    }
  }

  if (successCount + failCount === 0) {
    Logger.log('No tokens processed.');
    return;
  }

  //  ── write balances ────────────────────────────────────────────
  if (allNewAccounts.length) {
    storeBalanceDataWithAuth(
      { accounts: allNewAccounts },
      { numbers: { ach: allAch } }
    );
  }

  //  ── de-dupe step 1: drop pendings that have a posted sibling ───
  const toDrop = new Set(
    allTransactions
      .filter(tx => tx.pending === false && tx.pending_transaction_id)
      .map(tx => tx.pending_transaction_id)
  );
  let batch = allTransactions.filter(tx =>
    !(tx.pending === true && toDrop.has(tx.transaction_id))
  );

  //  ── de-dupe step 2: unique by transaction_id ─────────────────
  batch = Array.from(
    new Map(batch.map(tx => [tx.transaction_id, tx])).values()
  );

  //  ── de-dupe step 3: fuzzy key (date|amount|merchant-prefix) ────
  const seen = new Set();
  batch = batch.filter(tx => {
    const merchant = (tx.merchant_name || tx.name || '').substring(0,10).toLowerCase();
    const key = `${tx.date}|${tx.amount}|${merchant}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  //  ── write transactions ────────────────────────────────────────
  Logger.log(`About to call storeTransactionsData with ${batch.length} items (filtered from ${allTransactions.length})`);
  storeTransactionsData(batch);

  Logger.log(`✅ Done: ${successCount} OK, ${failCount} failed`);
  addNicknamesToTransactionsData();
  updateUserDashboardsAccountDropdowns();
  transferPlaidTransactionsToUserDashboards_Optimized();
  discoverRecurringTransactions();
}

/**
 * Incremental balance writer: only appends accounts not already in “Balance Data”
 *
 * @param {{accounts:Array}} balanceData   — the { accounts } object
 * @param {{numbers:{ach:Array}}} authData — the { numbers: { ach } } object
 */
function storeBalanceDataWithAuth(balanceData, authData) {
  const ss     = getAdminSpreadsheet();
  let sheet    = ss.getSheetByName('Balance Data');
  const header = [
    'Account ID','Account Name','Official Name','Type','Subtype',
    'Available Balance','Current Balance','Currency','Routing Number','Account Number'
  ];

  // 1) create + header if missing
  if (!sheet) {
    sheet = ss.insertSheet('Balance Data');
    sheet.appendRow(header);
  }

  // 2) read existing account IDs
  const lastRow = sheet.getLastRow();
  const existingIds = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String)
    : [];

  // 3) build lookup of ACH by account_id
  const achMap = {};
  (Array.isArray(authData?.numbers?.ach) ? authData.numbers.ach : [])
    .forEach(a => achMap[a.account_id] = a);

  // 4) filter & build only brand-new rows
  const rowsToAppend = (Array.isArray(balanceData.accounts) ? balanceData.accounts : [])
    .filter(ac => existingIds.indexOf(ac.account_id) < 0)
    .map(ac => {
      const ach = achMap[ac.account_id] || {};
      return [
        ac.account_id,
        ac.name,
        ac.official_name,
        ac.type,
        ac.subtype,
        ac.balances.available,
        ac.balances.current,
        ac.balances.iso_currency_code || ac.balances.unofficial_currency_code || '',
        ach.routing || '',
        ach.account || ''
      ];
    });

  // 5) append them (if any)
  if (rowsToAppend.length) {
    sheet
      .getRange(lastRow + 1, 1, rowsToAppend.length, header.length)
      .setValues(rowsToAppend);
    Logger.log(`storeBalanceDataWithAuth ▶ Appended ${rowsToAppend.length} new account(s).`);
  } else {
    Logger.log(`storeBalanceDataWithAuth ▶ No new accounts to append.`);
  }
}

/**
 * Fetches all transactions via Plaid paginated endpoint.
 */
function fetchAllTransactions(accessToken) {
  let allTx = [];
  let offset = 0;
  let total = 0;
  do {
    const page = fetchTransactionsDataWithOffset(accessToken, offset);
    if (page.transactions) allTx.push(...page.transactions);
    total = page.total_transactions;
    offset = allTx.length;
  } while (offset < total);
  return { transactions: allTx };
}

/**
 * Helpers for Plaid API calls
 */
function fetchTransactionsDataWithOffset(accessToken, offset) {
  const props = PropertiesService.getScriptProperties();
  const env = props.getProperty('PLAID_ENV');
  const payload = {
    client_id: props.getProperty('PLAID_CLIENT_ID'),
    secret: props.getProperty('PLAID_SECRET'),
    access_token: accessToken,
    start_date: Utilities.formatDate(new Date(Date.now() - 90*24*3600*1000), 'GMT', 'yyyy-MM-dd'),
    end_date: Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd'),
    options: { count: 500, offset }
  };
  const res = UrlFetchApp.fetch(`https://${env}.plaid.com/transactions/get`, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload)
  });
  const data = JSON.parse(res.getContentText());
  if (data.error) throw new Error(data.error.error_message);
  return data;
}

function fetchBalanceData(accessToken) {
  const props = PropertiesService.getScriptProperties();
  const env = props.getProperty('PLAID_ENV');
  const payload = {
    client_id: props.getProperty('PLAID_CLIENT_ID'),
    secret: props.getProperty('PLAID_SECRET'),
    access_token: accessToken
  };
  const res = UrlFetchApp.fetch(`https://${env}.plaid.com/accounts/balance/get`, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload)
  });
  const data = JSON.parse(res.getContentText());
  if (data.error) throw new Error(data.error.error_message);
  return data;
}

function fetchAuthData(accessToken) {
  const props = PropertiesService.getScriptProperties();
  const env = props.getProperty('PLAID_ENV');
  const payload = {
    client_id: props.getProperty('PLAID_CLIENT_ID'),
    secret: props.getProperty('PLAID_SECRET'),
    access_token: accessToken
  };
  const res = UrlFetchApp.fetch(`https://${env}.plaid.com/auth/get`, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload)
  });
  const data = JSON.parse(res.getContentText());
  if (data.error) throw new Error(data.error.error_message);
  return data;
}

/**
 * Writes balances (with routing/account numbers) to 'Balance Data' sheet
 */
function storeBalanceData(accounts, achRecords) {
  const ss = getAdminSpreadsheet();
  let sheet = ss.getSheetByName('Balance Data');
  if (!sheet) {
    sheet = ss.insertSheet('Balance Data');
    sheet.appendRow(['Account ID','Account Name','Official Name','Type','Subtype','Available Balance','Current Balance','Currency','Routing Number','Account Number']);
  }
  // remove old rows but keep header
  if (sheet.getLastRow()>1) sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).clearContent();

  // map ACH by account_id
  const achMap = {};
  achRecords.forEach(a=> achMap[a.account_id] = a);

  const rows = accounts.map(ac => {
    const ach = achMap[ac.account_id] || {};
    return [
      ac.account_id,
      ac.name,
      ac.official_name,
      ac.type,
      ac.subtype,
      ac.balances.available,
      ac.balances.current,
      ac.balances.iso_currency_code||ac.balances.unofficial_currency_code,
      ach.routing||'',
      ach.account||''
    ];
  });
  if (rows.length) sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  Logger.log(`storeBalanceData ▶ Wrote ${rows.length} rows`);
}

/**
 * Appends only brand‑new transactions into “Transactions Data”.
 * Accepts either:
 *   • An array:       storeTransactionsData([ {…}, {…} ])
 *   • Or an object:   storeTransactionsData({ transactions: [ … ] })
 */
function storeTransactionsData(input) {
  // figure out the array
  let txArray;
  if (Array.isArray(input)) {
    txArray = input;
  } else if (input?.transactions && Array.isArray(input.transactions)) {
    txArray = input.transactions;
  } else {
    Logger.log('storeTransactionsData ▶ ERROR: no transactions array found in input', input);
    return;
  }

  Logger.log(`storeTransactionsData ▶ Received ${txArray.length} transaction(s)`);

  const ss        = getAdminSpreadsheet();
  const sheetName = 'Transactions Data';
  let sheet       = ss.getSheetByName(sheetName);

  // 1) Create + header if missing
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`storeTransactionsData ▶ Created sheet "${sheetName}"`);
    sheet.appendRow([
      'Transaction ID',
      'Account ID',
      'Date',
      'Name',
      'Amount',
      'Currency',
      'Category',
      'Category ID',
      'Pending',
      'Merchant Name',
      'Payment Channel',
      'Transaction Type',
      'Nickname',
      'Transferred?',
      'Recurring Imported?'
    ]);
  }

  // 2) Load existing IDs
  const lastRow   = sheet.getLastRow();
  const existing  = new Set(
    lastRow > 1
      ? sheet.getRange(2,1,lastRow-1,1).getValues().flat().filter(String)
      : []
  );
  Logger.log(`storeTransactionsData ▶ ${existing.size} existing transaction IDs`);

  // 3) Build new rows
  const newRows = txArray
    .filter(tx => !existing.has(tx.transaction_id))
    .map(tx => [
      tx.transaction_id,
      tx.account_id,
      tx.date,
      tx.name,
      tx.amount,
      tx.iso_currency_code || tx.unofficial_currency_code,
      tx.category ? tx.category.join(' > ') : '',
      tx.category_id,
      tx.pending ? 'Yes' : 'No',
      tx.merchant_name || '',
      tx.payment_channel,
      tx.transaction_type,
      '',  // Nickname
      '',  // Transferred?
      ''   // Recurring Imported?
    ]);

  // 4) Append
  if (newRows.length) {
    Logger.log(`storeTransactionsData ▶ Appending ${newRows.length} new transaction(s)`);
    sheet
      .getRange(lastRow + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
    Logger.log(`storeTransactionsData ▶ Append complete`);
  } else {
    Logger.log(`storeTransactionsData ▶ No new transactions to append`);
  }
}

























function updateUserDashboardsAccountDropdowns() {
  const adminSs = getAdminSpreadsheet();
  
  // 1) Get display names from "Balance Data"
  const balanceSheet = adminSs.getSheetByName("Balance Data");
  if (!balanceSheet) {
    Logger.log('No "Balance Data" sheet found.');
    return;
  }
  
  const lastRow = balanceSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows in "Balance Data".');
    return;
  }
  
  // Identify columns for "Account Name" & "Official Name"
  const headers = balanceSheet.getRange(1, 1, 1, balanceSheet.getLastColumn()).getValues()[0];
  const accountNameColIndex = headers.indexOf("Account Name");
  const officialNameColIndex = headers.indexOf("Official Name");
  if (accountNameColIndex < 0 || officialNameColIndex < 0) {
    Logger.log('Could not find "Account Name" or "Official Name" columns in row 1. Check your headers.');
    return;
  }
  
  // Read data from row 2 downward
  const rangeData = balanceSheet.getRange(2, 1, lastRow - 1, balanceSheet.getLastColumn()).getValues();
  
  // Build the displayNames array
  const displayNamesSet = new Set();
  rangeData.forEach(row => {
    const accountName = row[accountNameColIndex] || "";
    const officialName = row[officialNameColIndex] || "";
    
    let displayName;
    if (accountName && officialName) {
      if (accountName === officialName) {
        displayName = accountName;
      } else {
        displayName = accountName + " - " + officialName;
      }
    } else if (accountName && !officialName) {
      displayName = accountName;
    } else if (!accountName && officialName) {
      displayName = officialName;
    } else {
      // both blank => skip
      return;
    }
    displayNamesSet.add(displayName);
  });
  
  const displayNames = Array.from(displayNamesSet);
  if (displayNames.length === 0) {
    Logger.log('No valid display names found in Balance Data.');
    return;
  }
  
  // 2) Find "Dashboard" column in "Users" sheet
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found.');
    return;
  }
  
  const usersLastCol = usersSheet.getLastColumn();
  const row2values = usersSheet.getRange(2, 1, 1, usersLastCol).getValues()[0];
  const dashboardColIndex = row2values.indexOf("Dashboard");
  
  if (dashboardColIndex < 0) {
    Logger.log('Could not find "Dashboard" label in row 2 of "Users" sheet.');
    return;
  }
  
  // 3) For each user (row >= 3), update their Dashboard
  const usersLastRow = usersSheet.getLastRow();
  if (usersLastRow < 3) {
    Logger.log('No user rows found below row 2.');
    return;
  }
  
  const dashboardUrlsRange = usersSheet.getRange(3, dashboardColIndex + 1, usersLastRow - 2, 1);
  const dashboardUrls = dashboardUrlsRange.getValues(); // 2D array
  
  // Convert the displayNames array to a comma-separated literal: {"Name1","Name2","Name3"}
  // This is used in the FILTER() formula
  const arrayLiteral = `{"${displayNames.join('","')}"}`;
  
  dashboardUrls.forEach((rowVal, i) => {
    const url = rowVal[0];
    if (!url) return; // skip empty
    
    const userDashboardId = extractIdFromUrl(url);
    if (!userDashboardId) {
      Logger.log(`Row ${i+3}: Invalid Dashboard URL: ${url}`);
      return;
    }
    
    // Open user's dashboard
    let userSs;
    try {
      userSs = SpreadsheetApp.openById(userDashboardId);
    } catch (err) {
      Logger.log(`Row ${i+3}: Could not open user dashboard: ${err.message}`);
      return;
    }
    
    // (A) Update data validation in "AssignedAccount"
    updateAssignedAccountValidation(userSs, displayNames);

    // (B) Update the "Unassigned Accounts" formula in the "Accounts" sheet
    updateUnassignedAccountsFormula(userSs, arrayLiteral);
    
    Logger.log(`Row ${i+3}: Done updating user dashboard: ${userSs.getName()}.`);
  });
  
  Logger.log('All user dashboards updated successfully.');
}

/**
 * PART A: Sets data validation in the named range "AssignedAccount".
 */
function updateAssignedAccountValidation(userSs, displayNames) {
  const namedRange = userSs.getRangeByName("AssignedAccount");
  if (!namedRange) {
    Logger.log(`Named range "AssignedAccount" not found in ${userSs.getName()}.`);
    return;
  }
  
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(displayNames, true) 
    .setAllowInvalid(false)
    .build();
    
  namedRange.setDataValidation(rule);
}

/**
 * PART B: Sets the formula under "Unassigned Accounts:" in the user’s "Accounts" sheet.
 *   - "Unassigned Accounts:" label is in row 2
 *   - The formula goes in row 3 of the same column
 *   - The formula is: =FILTER(ARRAYFORMULA({...}), ISNA(MATCH(ARRAYFORMULA({...}), AssignedAccount, 0)))
 */
function updateUnassignedAccountsFormula(userSs, arrayLiteral) {
  const accountsSheet = userSs.getSheetByName("Accounts");
  if (!accountsSheet) {
    Logger.log(`No "Accounts" sheet found in ${userSs.getName()}.`);
    return;
  }

  // Find "Unassigned Accounts:" in row 2
  const lastCol = accountsSheet.getLastColumn();
  const row2vals = accountsSheet.getRange(2, 1, 1, lastCol).getValues()[0]; // 1D array
  const unassignedColIndex = row2vals.indexOf("Unassigned Accounts:");
  if (unassignedColIndex < 0) {
    Logger.log(`No "Unassigned Accounts:" label found in row 2 of "Accounts" for ${userSs.getName()}.`);
    return;
  }

  // Row 3, same column => unassignedColIndex + 1 (1-based column index)
  const targetCell = accountsSheet.getRange(3, unassignedColIndex + 1);

  // Replace commas with semicolons in the arrayLiteral
  const verticalArrayLiteral = arrayLiteral.replace(/,/g, ";");

  // Build the FILTER() formula
  const formula = `=IFERROR(FILTER(ARRAYFORMULA(${verticalArrayLiteral}), ISNA(MATCH(ARRAYFORMULA(${verticalArrayLiteral}), AssignedAccount, 0))))`;

  targetCell.setValue(formula);
  Logger.log(`Set "Unassigned Accounts" formula in row 3, col ${unassignedColIndex + 1} of "Accounts" sheet: ${formula}`);
}

function createAndLinkLegalDocsFolders() {
  try {
    // 1. Open the Admin Dashboard and locate the "Users" sheet
    const adminSs    = getAdminSpreadsheet();
    const usersSheet = adminSs.getSheetByName("Users");
    if (!usersSheet) throw new Error("Users sheet not found in Admin Dashboard.");

    // 2. Figure out where "Dashboard" and "Primary Email" live in row 2
    const hdrs = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
    const dashboardColIndex = hdrs.indexOf("Dashboard");
    const emailColIndex     = hdrs.indexOf("Primary Email");
    if (dashboardColIndex < 0) throw new Error(`"Dashboard" column not found in Users sheet.`);
    if (emailColIndex     < 0) throw new Error(`"Primary Email" column not found in Users sheet.`);

    // 3. Grab all dashboard URLs & corresponding emails
    const lastRow = usersSheet.getLastRow();
    const dashboards = usersSheet
      .getRange(3, dashboardColIndex+1, lastRow - 2, 1)
      .getValues()
      .flat();
    const emails = usersSheet
      .getRange(3, emailColIndex+1, lastRow - 2, 1)
      .getValues()
      .flat();

    // 4. Process each user
    dashboards.forEach((dashboardUrl, i) => {
      const email = emails[i];
      if (!dashboardUrl || !email) return;

      const dashboardId = extractIdFromUrl(dashboardUrl);
      if (!dashboardId) {
        Logger.log(`Row ${i+3}: invalid Dashboard URL, skipping.`);
        return;
      }

      try {
        // open user dashboard & find its parent folder
        const userSs        = SpreadsheetApp.openById(dashboardId);
        const dashboardFile = DriveApp.getFileById(dashboardId);
        const parents       = dashboardFile.getParents();
        if (!parents.hasNext()) {
          Logger.log(`Row ${i+3}: no parent folder, skipping.`);
          return;
        }
        const parentFolder = parents.next();

        // open Entities sheet
        const entitiesSheet = userSs.getSheetByName("Entities");
        if (!entitiesSheet) {
          Logger.log(`Row ${i+3}: no "Entities" sheet, skipping.`);
          return;
        }

        // locate columns
        const eHdrs              = entitiesSheet.getRange(2,1,1,entitiesSheet.getLastColumn()).getValues()[0];
        const entityNameColIndex = eHdrs.indexOf("Entity Name:");
        const legalDocsColIndex  = eHdrs.indexOf("Legal Documents:");
        if (entityNameColIndex < 0 || legalDocsColIndex < 0) {
          Logger.log(`Row ${i+3}: missing headers, skipping.`);
          return;
        }

        // for each entity row
        const lastEntRow = entitiesSheet.getLastRow();
        const data       = entitiesSheet.getRange(3, 1, lastEntRow - 2, entitiesSheet.getLastColumn()).getValues();
        data.forEach((row, j) => {
          const entityName     = row[entityNameColIndex];
          const existingLink   = row[legalDocsColIndex];
          if (!entityName || existingLink) return;

          // create folder, share it with the user as Editor
          const folderName       = `${entityName} - Legal Docs`;
          const newFolder        = parentFolder.createFolder(folderName);
          newFolder.addEditor(email);

          // write a hyperlink into the sheet
          const folderUrl        = newFolder.getUrl();
          const linkFormula      = `=HYPERLINK("${folderUrl}","Open")`;
          const targetRow        = j + 3;  // because data[0] is row 3
          entitiesSheet
            .getRange(targetRow, legalDocsColIndex + 1)
            .setValue(linkFormula);

          Logger.log(`Row ${targetRow}: created & shared "${folderName}" with ${email}`);
        });

      } catch (err) {
        Logger.log(`Row ${i+3}: error processing dashboard – ${err.message}`);
      }
    });

    Logger.log("createAndLinkLegalDocsFolders: done.");
  } catch (err) {
    Logger.log(`Error in createAndLinkLegalDocsFolders: ${err.message}`);
  }
}



function addNicknamesToTransactionsData() {
  const adminSs = getAdminSpreadsheet();

  //-------------------------------
  // 1) Build accountId -> combinedName from "Balance Data"
  //-------------------------------
  const balanceSheet = adminSs.getSheetByName("Balance Data");
  if (!balanceSheet) {
    Logger.log('No "Balance Data" sheet found.');
    return;
  }

  // Identify columns for "Account ID", "Account Name", "Official Name"
  const headers = balanceSheet.getRange(1, 1, 1, balanceSheet.getLastColumn()).getValues()[0];
  const accountIdColIndex = headers.indexOf("Account ID");
  const accountNameColIndex = headers.indexOf("Account Name");
  const officialNameColIndex = headers.indexOf("Official Name");
  if (accountIdColIndex < 0 || accountNameColIndex < 0 || officialNameColIndex < 0) {
    Logger.log('Could not find "Account ID", "Account Name", or "Official Name" columns in "Balance Data".');
    return;
  }

  const lastRowBalance = balanceSheet.getLastRow();
  if (lastRowBalance < 2) {
    Logger.log('No account rows in "Balance Data".');
    return;
  }

  // Read the data from row 2 downward
  const dataRangeBalance = balanceSheet.getRange(2, 1, lastRowBalance - 1, balanceSheet.getLastColumn()).getValues();
  const accountIdToCombinedName = {};

  dataRangeBalance.forEach(row => {
    const accountId = row[accountIdColIndex];
    const acctName = row[accountNameColIndex] || "";
    const offName = row[officialNameColIndex] || "";

    // Combine them the same way you do for dropdowns
    let combinedName;
    if (acctName && offName) {
      if (acctName === offName) {
        combinedName = acctName;
      } else {
        combinedName = acctName + " - " + offName;
      }
    } else if (acctName && !offName) {
      combinedName = acctName;
    } else if (!acctName && offName) {
      combinedName = offName;
    } else {
      // both blank => skip
      return;
    }

    if (accountId) {
      accountIdToCombinedName[accountId] = combinedName;
    }
  });

  //-------------------------------
  // 2) Build accountId -> nickname across all users
  //-------------------------------
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found.');
    return;
  }

  const lastRowUsers = usersSheet.getLastRow();
  if (lastRowUsers < 3) {
    Logger.log('No user rows in "Users" sheet.');
    return;
  }

  const row2vals = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashboardColIndex = row2vals.indexOf("Dashboard");
  if (dashboardColIndex < 0) {
    Logger.log('No "Dashboard" label found in row 2 of "Users" sheet.');
    return;
  }

  // We'll combine everything into one big dictionary: { accountId -> nickname }
  const accountIdToNickname = {};

  // For each user
  const userDashUrls = usersSheet.getRange(3, dashboardColIndex + 1, lastRowUsers - 2, 1).getValues();
  userDashUrls.forEach((rowVal, i) => {
    const dashboardUrl = rowVal[0];
    if (!dashboardUrl) return;
    const dashboardId = extractIdFromUrl(dashboardUrl);
    if (!dashboardId) return;

    try {
      const userSs = SpreadsheetApp.openById(dashboardId);
      const accountsSheet = userSs.getSheetByName("Accounts");
      if (!accountsSheet) {
        Logger.log(`No "Accounts" sheet in user dashboard: row ${i+3}`);
        return;
      }

      // Find columns "Assigned Account:" and "Nickname:"
      const row2 = accountsSheet.getRange(2, 1, 1, accountsSheet.getLastColumn()).getValues()[0];
      const assignedIndex = row2.indexOf("Assigned Account:");
      const nicknameIndex = row2.indexOf("Nickname:");
      if (assignedIndex < 0 || nicknameIndex < 0) {
        Logger.log(`Missing "Assigned Account:" or "Nickname:" in row 2 of user's "Accounts" sheet. Row ${i+3}`);
        return;
      }

      const lastRowAcc = accountsSheet.getLastRow();
      if (lastRowAcc < 3) return;

      const userAccountsData = accountsSheet.getRange(3, 1, lastRowAcc - 2, accountsSheet.getLastColumn()).getValues();

      userAccountsData.forEach(accRow => {
        const assignedName = accRow[assignedIndex] || "";
        const nickname = accRow[nicknameIndex] || "";
        if (assignedName && nickname) {
          // Find the accountId that has this combinedName
          for (let acctId in accountIdToCombinedName) {
            if (accountIdToCombinedName[acctId] === assignedName) {
              accountIdToNickname[acctId] = nickname;
            }
          }
        }
      });
    } catch (e) {
      Logger.log(`Error opening user dashboard row ${i+3}: ${e.message}`);
    }
  });

  //-------------------------------
  // 3) Add or update a "Nickname" column in "Transactions Data"
  //-------------------------------
  const txSheet = adminSs.getSheetByName("Transactions Data");
  if (!txSheet) {
    Logger.log('No "Transactions Data" sheet found.');
    return;
  }

  // Find or create the "Nickname" column at the far right
  let txHeaders = txSheet.getRange(1, 1, 1, txSheet.getLastColumn()).getValues()[0];
  let nicknameColIndex = txHeaders.indexOf("Nickname");
  if (nicknameColIndex < 0) {
    // Insert a new column after the last column
    nicknameColIndex = txHeaders.length; // 0-based
    txSheet.insertColumnAfter(nicknameColIndex + 1);
    txSheet.getRange(1, nicknameColIndex + 2).setValue("Nickname"); // 1-based
    // Refresh headers
    txHeaders = txSheet.getRange(1, 1, 1, txSheet.getLastColumn()).getValues()[0];
    nicknameColIndex = txHeaders.indexOf("Nickname");
  }

  // Also find "Account ID" column
  const accountIdTxIndex = txHeaders.indexOf("Account ID");
  if (accountIdTxIndex < 0) {
    Logger.log('No "Account ID" column found in "Transactions Data". Cannot map nicknames.');
    return;
  }

  const lastRowTx = txSheet.getLastRow();
  if (lastRowTx < 2) {
    Logger.log(`No transaction rows in "Transactions Data".`);
    return;
  }

  // Read all transaction rows
  const txRange = txSheet.getRange(2, 1, lastRowTx - 1, txSheet.getLastColumn());
  const txData = txRange.getValues();

  let changes = 0;
  txData.forEach((row, i) => {
    const acctIdCellValue = row[accountIdTxIndex];
    if (accountIdToNickname.hasOwnProperty(acctIdCellValue)) {
      // Write that nickname into the "Nickname" column
      row[nicknameColIndex] = accountIdToNickname[acctIdCellValue];
      changes++;
    } else {
      // If there's no nickname, you can either blank it or leave it as is
      row[nicknameColIndex] = ""; 
    }
  });

  if (changes > 0) {
    txRange.setValues(txData);
    Logger.log(`Wrote ${changes} nicknames into "Transactions Data" sheet (Nickname column).`);
  } else {
    Logger.log(`No nicknames added; either no matches or no nicknames exist.`);
  }

  transferPlaidTransactionsToUserDashboards_Optimized();

}

function transferPlaidTransactionsToUserDashboards_Optimized() {
  const adminSs = getAdminSpreadsheet();
  const tz      = Session.getScriptTimeZone();
  const txSheet = adminSs.getSheetByName("Transactions Data");
  if (!txSheet) {
    Logger.log('No "Transactions Data" sheet found in Admin Dashboard.');
    return;
  }

  // ─── Prep Transactions-Data headers ─────────────────────────
  let txHdrs = txSheet.getRange(1,1,1,txSheet.getLastColumn()).getValues()[0];
  const txnIdIdx        = txHdrs.indexOf("Transaction ID");
  const dateIdx         = txHdrs.indexOf("Date");
  const amtIdx          = txHdrs.indexOf("Amount");
  const currIdx         = txHdrs.indexOf("Currency");
  const merchIdx        = txHdrs.indexOf("Merchant Name");
  const nameIdx         = txHdrs.indexOf("Name");
  const catIdx          = txHdrs.indexOf("Category");
  const nickIdx         = txHdrs.indexOf("Nickname");
  const pendingIdx      = txHdrs.indexOf("Pending");
  let   pendingTxnIdIdx = txHdrs.indexOf("Pending Transaction ID");
  let   transferredIdx  = txHdrs.indexOf("Transferred?");

  if ([txnIdIdx,dateIdx,amtIdx,nickIdx,pendingIdx].some(i=>i<0)) {
    Logger.log("Missing required columns in Transactions Data. Aborting.");
    return;
  }

  if (pendingTxnIdIdx < 0) {
    txSheet.insertColumnAfter(txHdrs.length);
    txSheet.getRange(1, txHdrs.length+1).setValue("Pending Transaction ID");
    txHdrs = txSheet.getRange(1,1,1,txSheet.getLastColumn()).getValues()[0];
    pendingTxnIdIdx = txHdrs.indexOf("Pending Transaction ID");
  }
  if (transferredIdx < 0) {
    txSheet.insertColumnAfter(txHdrs.length);
    txSheet.getRange(1, txHdrs.length+1).setValue("Transferred?");
    txHdrs = txSheet.getRange(1,1,1,txSheet.getLastColumn()).getValues()[0];
    transferredIdx = txHdrs.indexOf("Transferred?");
  }

  // ─── Load all TX rows ────────────────────────────────────────
  const lastRow = txSheet.getLastRow();
  if (lastRow < 2) return;
  const txRange = txSheet.getRange(2,1,lastRow-1,txHdrs.length);
  const txData  = txRange.getValues();

  // ─── Build nickname→dashboard map ───────────────────────────
  const usersSheet            = adminSs.getSheetByName("Users");
  const nicknameToDashboardId = buildNicknameToDashboardMap_Optimized(usersSheet);

  const rowsToMarkPending     = [];
  const rowsToMarkTransferred = [];

  txData.forEach((row,i) => {
    if (row[transferredIdx] === "Yes") return; // already done

    const nickname   = row[nickIdx];
    const dashId     = nicknameToDashboardId[nickname];
    if (!nickname || !dashId) return;

    const txnId      = row[txnIdIdx].toString();
    const isPending  = row[pendingIdx] === "Yes";
    const oldPendId  = (row[pendingTxnIdIdx]||"").toString();

    // format date & amount
    let userDate = "";
    if (row[dateIdx] instanceof Date) {
      userDate = Utilities.formatDate(row[dateIdx], tz, "M/d/yyyy");
    } else {
      const [y,m,d] = (""+row[dateIdx]).split("-");
      userDate = `${m}/${d}/${y}`;
    }
    const rawAmt = Number(row[amtIdx])||0;
    const amtStr = (row[currIdx]==="USD")
      ? "$"+Math.abs(rawAmt).toFixed(2)
      : Math.abs(rawAmt).toFixed(2)+" "+row[currIdx];
    let fromVal = nickname, toVal = row[merchIdx]||row[nameIdx]||"";
    if (rawAmt < 0) [fromVal,toVal] = [toVal,fromVal];

    let cat="", subcat="";
    if (catIdx>=0 && row[catIdx]) {
      [cat, subcat] = row[catIdx].split(" > ");
    }

    // ─── Open user dashboard & ensure Plaid-ID col ─────────────
    const userSs   = SpreadsheetApp.openById(dashId);
    const finSheet = userSs.getSheetByName("Financial Journal");
    if (!finSheet) return;

    // pull headers from row 2 (label row)
    let finHdrs    = finSheet.getRange(2,1,1,finSheet.getLastColumn()).getValues()[0];
    let plaidIdCol = finHdrs.indexOf("Plaid Txn ID");
    if (plaidIdCol < 0) {
      // insert it at the far right
      const insertCol = finHdrs.length + 1;
      finSheet.insertColumnAfter(finHdrs.length);
      finSheet.getRange(2, insertCol).setValue("Plaid Txn ID");
      finSheet.hideColumn(finSheet.getRange(1, insertCol));
      // re-load
      finHdrs    = finSheet.getRange(2,1,1,finSheet.getLastColumn()).getValues()[0];
      plaidIdCol = finHdrs.indexOf("Plaid Txn ID");
    }

    // build simple colMap + our new plaidTxnId index
    const colMap = buildFinJournalColMap(finHdrs);
    colMap.plaidTxnId = plaidIdCol;

    // helper to insert a full-width blank row
    function insertFullRowWith(valuesByIndex) {
      const template = Array(finHdrs.length).fill("");
      Object.keys(valuesByIndex).forEach(idx => {
        template[idx] = valuesByIndex[idx];
      });
      const nextRow = finSheet.getLastRow() + 1;
      finSheet.insertRowAfter(finSheet.getLastRow());
      finSheet.getRange(nextRow, 1, 1, template.length).setValues([template]);
    }

    // —— 1) brand-new pending: insert & track its ID
    if (isPending && !oldPendId) {
      insertFullRowWith({
        [colMap.date]:         userDate,
        [colMap.fromCol]:      fromVal,
        [colMap.toCol]:        toVal,
        [colMap.amount]:       amtStr,
        [colMap.category]:     cat,
        [colMap.subcat]:       subcat,
        [colMap.plaidTxnId]:   txnId
      });
      txData[i][pendingTxnIdIdx] = txnId;
      rowsToMarkPending.push(i);
      return;
    }

    // —— 2) pending → posted: update the existing row
    if (!isPending && oldPendId) {
      const body = finSheet.getRange(3,1,finSheet.getLastRow()-2,finHdrs.length).getValues();
      for (let j=0; j<body.length; j++) {
        if (""+body[j][plaidIdCol] === oldPendId) {
          finSheet.getRange(3+j, colMap.date+1)         .setValue(userDate);
          finSheet.getRange(3+j, colMap.amount+1)       .setValue(amtStr);
          finSheet.getRange(3+j, colMap.plaidTxnId+1)   .setValue(txnId);
          rowsToMarkTransferred.push(i);
          return;
        }
      }
      // else falls through to #3
    }

    // —— 3) brand-new posted with no prior pending
    if (!isPending && !oldPendId) {
      insertFullRowWith({
        [colMap.date]:         userDate,
        [colMap.fromCol]:      fromVal,
        [colMap.toCol]:        toVal,
        [colMap.amount]:       amtStr,
        [colMap.category]:     cat,
        [colMap.subcat]:       subcat,
        [colMap.plaidTxnId]:   txnId
      });
      rowsToMarkTransferred.push(i);
    }
  });

  // ─── Write flags back ────────────────────────────────────────
  rowsToMarkPending.forEach(r =>    txData[r][pendingTxnIdIdx]   = txData[r][pendingTxnIdIdx]);
  rowsToMarkTransferred.forEach(r => txData[r][transferredIdx]   = "Yes");
  txRange.setValues(txData);

  Logger.log("Bulk transfer/update complete.");
}





/**
 * Return { Nickname -> userDashboardId } by scanning each user’s 
 * "Accounts" sheet in their dashboard, 
 * similar to your existing approach but in one pass.
 */
function buildNicknameToDashboardMap_Optimized(usersSheet) {
  const adminSs = getAdminSpreadsheet();
  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) return {};

  // Find "Dashboard" col in row 2
  const row2vals = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol = row2vals.indexOf("Dashboard");
  if (dashCol < 0) return {};

  let map = {};

  // For each user row
  const userData = usersSheet.getRange(3, dashCol + 1, lastRow - 2, 1).getValues();
  userData.forEach((rowVal, i) => {
    let dashboardUrl = rowVal[0];
    if (!dashboardUrl) return;
    let dashId = extractIdFromUrl(dashboardUrl);
    if (!dashId) return;

    // Try to read "Accounts" -> Nicknames
    try {
      let userSs = SpreadsheetApp.openById(dashId);
      let accSheet = userSs.getSheetByName("Accounts");
      if (!accSheet) return;

      const row2 = accSheet.getRange(2, 1, 1, accSheet.getLastColumn()).getValues()[0];
      const assignedIdx = row2.indexOf("Assigned Account:");
      const nicknameIdx = row2.indexOf("Nickname:");
      if (assignedIdx < 0 || nicknameIdx < 0) return;

      const lr = accSheet.getLastRow();
      if (lr < 3) return;
      const accData = accSheet.getRange(3, 1, lr - 2, accSheet.getLastColumn()).getValues();
      accData.forEach(accRow => {
        let nicknameVal = accRow[nicknameIdx];
        if (nicknameVal) {
          map[nicknameVal] = dashId;
        }
      });
    } catch (e) {
      // skip
      Logger.log(`Error building nickname map for row ${i+3}: ${e.message}`);
    }
  });

  return map;
}


/**
 * Convert a date from either a Date object or "YYYY-MM-DD" to "M/D/YYYY" (no leading zeros).
 */
function convertDateString_MDY(plaidDateStr) {
  if (!plaidDateStr) return "";

  // If it's already a Date
  if (plaidDateStr instanceof Date) {
    // "M/d/yyyy" => no leading zero
    return Utilities.formatDate(plaidDateStr, "GMT", "M/d/yyyy");
  }

  // If it's a string
  if (typeof plaidDateStr === "string") {
    let parts = plaidDateStr.split("-");
    if (parts.length === 3) {
      let year = +parts[0];
      let month = +parts[1];
      let day = +parts[2];
      if (year && month && day) {
        return `${month}/${day}/${year}`; // M/D/YYYY
      }
    }
    return plaidDateStr; // fallback
  }

  // Otherwise fallback
  return plaidDateStr.toString();
}


/**
 * Splits a Plaid category string (e.g. "Food and Drink > Restaurants > Mexican").
 * Returns { topCategory, subCategory }.
 */
function parsePlaidCategory(catString) {
  let topCategory = "";
  let subCategory = "";
  if (!catString) return { topCategory, subCategory };

  let splitted = catString.split(" > ");
  if (splitted.length >= 1) {
    topCategory = splitted[0];
  }
  if (splitted.length >= 3) {
    // prefer the 3rd item
    subCategory = splitted[2];
  } else if (splitted.length === 2) {
    // no 3rd
    subCategory = splitted[1];
  }

  return { topCategory, subCategory };
}















/***************************************************************
 * discoverRecurringTransactions()
 * 
 * Scans the Admin Dashboard's "Transactions Data" sheet for
 * recurring transactions and logs them to the user's dashboard.
 * The recurring transaction string is placed in the "Unassigned"
 * column and an empty checkbox in the "✔" column (headers in row 2).
 ***************************************************************/
function discoverRecurringTransactions() {
  const adminSs = getAdminSpreadsheet();
  const txSheet = adminSs.getSheetByName("Transactions Data");
  if (!txSheet) {
    Logger.log('No "Transactions Data" sheet found in Admin Dashboard.');
    return;
  }
  
  // Assume headers are in row 1.
  let txHeaders = txSheet.getRange(1, 1, 1, txSheet.getLastColumn()).getValues()[0];
  const nameIdx          = txHeaders.indexOf("Name");
  const merchantNameIdx  = txHeaders.indexOf("Merchant Name");
  const amountIdx        = txHeaders.indexOf("Amount");
  const dateIdx          = txHeaders.indexOf("Date");
  const categoryIdx      = txHeaders.indexOf("Category");
  const nicknameIdx      = txHeaders.indexOf("Nickname");
  
  // Ensure a "Recurring Imported?" column exists.
  let recImportedIdx = txHeaders.indexOf("Recurring Imported?");
  if (recImportedIdx < 0) {
    recImportedIdx = txHeaders.length;
    txSheet.insertColumnAfter(recImportedIdx);
    txSheet.getRange(1, recImportedIdx + 1).setValue("Recurring Imported?");
    txHeaders = txSheet.getRange(1, 1, 1, txSheet.getLastColumn()).getValues()[0];
    recImportedIdx = txHeaders.indexOf("Recurring Imported?");
  }
  
  if (merchantNameIdx < 0 || amountIdx < 0 || dateIdx < 0 || nicknameIdx < 0) {
    Logger.log("Missing required columns in Transactions Data.");
    return;
  }
  
  const lastRow = txSheet.getLastRow();
  const txDataRange = txSheet.getRange(2, 1, lastRow - 1, txSheet.getLastColumn());
  const txData = txDataRange.getValues();
  
  // Group recurring transactions by composite key:
  // normalized effective merchant (Merchant Name if available; else Name) + "|" + amount.
  let recurringGroups = {};
  for (let i = 0; i < txData.length; i++) {
    let row = txData[i];
    if (row[recImportedIdx] === "Yes") continue;
    
    // Determine effective merchant (prefer Merchant Name)
    let effectiveMerchant = row[merchantNameIdx] || row[nameIdx] || "";
    let normEffective = effectiveMerchant.replace(/\s/g, "").toLowerCase();
    
    // Use the exact amount for matching.
    let amountVal = row[amountIdx];
    let amountKey = amountVal.toString().trim();
    
    let key = normEffective + "|" + amountKey;
    if (!recurringGroups[key]) {
      recurringGroups[key] = {
        rows: [],
        effectiveMerchant: effectiveMerchant,
        amount: amountVal,
        dates: [],
        category: "",
        subcategory: "",
        nickname: row[nicknameIdx] || ""
      };
    }
    recurringGroups[key].rows.push(i);
    recurringGroups[key].dates.push(row[dateIdx]);
    if (!recurringGroups[key].category && categoryIdx >= 0) {
      let parsed = parsePlaidCategory(row[categoryIdx] || "");
      recurringGroups[key].category    = parsed.topCategory;
      recurringGroups[key].subcategory = parsed.subCategory;
    }
    if (!recurringGroups[key].nickname && nicknameIdx >= 0) {
      recurringGroups[key].nickname = row[nicknameIdx];
    }
  }
  
  // Only process groups with at least 2 occurrences.
  const groupsToImport = Object.values(recurringGroups).filter(g => g.rows.length >= 2);
  if (groupsToImport.length === 0) {
    Logger.log("No recurring transactions found.");
    return;
  }
  
  // Build a mapping from Nickname to User Dashboard ID.
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found in Admin Dashboard.');
    return;
  }
  const nicknameToDashboard = buildNicknameToDashboardMap_Optimized(usersSheet);
  
  // Process each recurring group.
  groupsToImport.forEach(group => {
    // Format dates
    let formattedDates = group.dates
      .map(dt => (dt instanceof Date ? formatMdy(dt) : dt))
      .filter((v,i,a) => a.indexOf(v) === i)
      .sort((a,b) => new Date(a) - new Date(b));
    let datesStr = formattedDates.join(", ");
    
    // Format amount as USD
    let absAmount = Math.abs(parseFloat(group.amount));
    let amountStr = "$" + absAmount.toFixed(2);
    
    // Build description
    let positive = parseFloat(group.amount) >= 0;
    let recurringTxStr = positive
      ? `${amountStr} to ${group.effectiveMerchant} from ${group.nickname} on ${datesStr} - ${group.category} > ${group.subcategory}`
      : `${amountStr} from ${group.effectiveMerchant} to ${group.nickname} on ${datesStr} - ${group.category} > ${group.subcategory}`;
    
    // Find user's dashboard
    let dashboardId = nicknameToDashboard[group.nickname];
    if (!dashboardId) {
      Logger.log(`No dashboard for nickname ${group.nickname}, skipping ${group.effectiveMerchant}`);
      return;
    }
    
    try {
      let userSs   = SpreadsheetApp.openById(dashboardId);
      let recSheet = userSs.getSheetByName("Recurring Transactions");
      if (!recSheet) {
        Logger.log(`No "Recurring Transactions" sheet in ${userSs.getName()}, skipping.`);
        return;
      }
      
      const recHeaders = recSheet.getRange(2, 1, 1, recSheet.getLastColumn()).getValues()[0];
      const recTxCol    = recHeaders.indexOf("Unassigned");
      const checkCol    = recHeaders.indexOf("✔");
      if (recTxCol < 0 || checkCol < 0) {
        Logger.log(`Missing columns in ${userSs.getName()}!`);
        return;
      }
      
      // Check for existing
      let exists = false;
      if (recSheet.getLastRow() >= 3) {
        let existing = recSheet.getRange(3, recTxCol + 1, recSheet.getLastRow() - 2, 1)
                              .getValues().flat();
        let normEff = group.effectiveMerchant.replace(/\s/g, "").toLowerCase().substring(0,5);
        exists = existing.some(val =>
          typeof val === "string" &&
          val.toLowerCase().includes(normEff) &&
          val.indexOf(amountStr) > -1
        );
      }
      
      if (!exists) {
        let newRow = new Array(recSheet.getLastColumn()).fill("");
        newRow[recTxCol] = recurringTxStr;
        newRow[checkCol] = false;
        let nextRow = recSheet.getLastRow() + 1;
        recSheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);
        recSheet.getRange(nextRow, checkCol+1).insertCheckboxes().setValue(false);
        Logger.log(`Imported recurring tx "${recurringTxStr}" into ${userSs.getName()} row ${nextRow}`);
      }
      
      // Mark source rows as imported
      group.rows.forEach(idx => txData[idx][recImportedIdx] = "Yes");
      
    } catch (e) {
      Logger.log(`Error in group ${group.effectiveMerchant}: ${e.message}`);
    }
  });
  
  // Write back to Transactions Data
  txSheet.getRange(2, 1, txData.length, txSheet.getLastColumn()).setValues(txData);
  Logger.log("Recurring transactions discovery complete.");
}



/***************************************************************
 * Helper: buildNicknameToDashboardMap_Optimized
 ***************************************************************/
function buildNicknameToDashboardMap_Optimized(usersSheet) {
  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) return {};
  
  const row2vals = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol = row2vals.indexOf("Dashboard");
  if (dashCol < 0) return {};
  
  let map = {};
  const userData = usersSheet.getRange(3, dashCol + 1, lastRow - 2, 1).getValues();
  userData.forEach((rowVal, i) => {
    let dashboardUrl = rowVal[0];
    if (!dashboardUrl) return;
    let dashId = extractIdFromUrl(dashboardUrl);
    if (!dashId) return;
    try {
      let userSs = SpreadsheetApp.openById(dashId);
      let accSheet = userSs.getSheetByName("Accounts");
      if (!accSheet) return;
      const row2 = accSheet.getRange(2, 1, 1, accSheet.getLastColumn()).getValues()[0];
      const assignedIdx = row2.indexOf("Assigned Account:");
      const nicknameIdx = row2.indexOf("Nickname:");
      if (assignedIdx < 0 || nicknameIdx < 0) return;
      const lr = accSheet.getLastRow();
      if (lr < 3) return;
      const accData = accSheet.getRange(3, 1, lr - 2, accSheet.getLastColumn()).getValues();
      accData.forEach(accRow => {
        let nicknameVal = accRow[nicknameIdx];
        if (nicknameVal) {
          map[nicknameVal] = dashId;
        }
      });
    } catch (e) {
      Logger.log(`Error building nickname map for row ${i+3}: ${e.message}`);
    }
  });
  return map;
}

/***************************************************************
 * formatMdy: Formats a Date as M/D/YYYY (no leading zeros)
 ***************************************************************/
function formatMdy(dateVal) {
  if (!(dateVal instanceof Date)) return dateVal;
  let m = dateVal.getMonth() + 1;
  let d = dateVal.getDate();
  let y = dateVal.getFullYear();
  return `${m}/${d}/${y}`;
}

/***************************************************************
 * parsePlaidCategory: Splits a category string into top and subcategories.
 ***************************************************************/
function parsePlaidCategory(catString) {
  let topCategory = "";
  let subCategory = "";
  if (!catString) return { topCategory, subCategory };
  let parts = catString.split(" > ");
  if (parts.length >= 1) {
    topCategory = parts[0];
  }
  if (parts.length >= 3) {
    subCategory = parts[2];
  } else if (parts.length === 2) {
    subCategory = parts[1];
  }
  return { topCategory, subCategory };
}

/***************************************************************
 * extractIdFromUrl: Extracts a document ID from a Google URL.
 ***************************************************************/
function extractIdFromUrl(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (let r of patterns) {
    const match = r.exec(url);
    if (match && match[1]) return match[1];
  }
  return null;
}













/**
 * clearRecurringMergedEntries()
 * 
 * From the Admin Dashboard, this function loops through each user (from the "Users" sheet),
 * opens that user's dashboard, and on the "Recurring Transactions" sheet looks for any row
 * (starting from row 3) where the cell in the "✔" column is checked.
 * For each such row, it deletes the cell in the "Unassigned" column and the cell in the "✔" column,
 * shifting all cells in that column upward—so that no gap remains—while leaving the rest of the row intact.
 */
function clearRecurringMergedEntries() {
  const adminSs = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found in Admin Dashboard.');
    return;
  }
  
  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('No user rows found in "Users" sheet.');
    return;
  }
  
  // Get header row (assumed to be row 2) to determine the "Dashboard" column.
  const headerRow = 2;
  const usersHeaders = usersSheet.getRange(headerRow, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol = usersHeaders.indexOf("Dashboard");
  if (dashCol < 0) {
    Logger.log('No "Dashboard" label found in row 2 of "Users" sheet.');
    return;
  }
  
  // Get all user dashboard URLs from row 3 downward.
  const dashboardUrls = usersSheet.getRange(3, dashCol + 1, lastRow - 2, 1).getValues();
  
  dashboardUrls.forEach((rowVal, i) => {
    const dashboardUrl = rowVal[0];
    if (!dashboardUrl) return;
    
    const userDashboardId = extractIdFromUrl(dashboardUrl);
    if (!userDashboardId) {
      Logger.log(`Row ${i + 3}: Invalid Dashboard URL. Skipping.`);
      return;
    }
    
    try {
      let userSs = SpreadsheetApp.openById(userDashboardId);
      let recSheet = userSs.getSheetByName("Recurring Transactions");
      if (!recSheet) {
        Logger.log(`No "Recurring Transactions" sheet found in dashboard ${userSs.getName()}. Skipping.`);
        return;
      }
      
      // Read header row (row 2) to get column indexes.
      const recHeaders = recSheet.getRange(2, 1, 1, recSheet.getLastColumn()).getValues()[0];
      const unassignedCol = recHeaders.indexOf("Unassigned") + 1; // 1-based
      const checkCol = recHeaders.indexOf("✔") + 1;
      
      if (unassignedCol < 1 || checkCol < 1) {
        Logger.log(`Missing "Unassigned" or "✔" columns in ${userSs.getName()}. Skipping.`);
        return;
      }
      
      // Process from bottom to top (to avoid shifting issues) starting at row 3.
      const recLastRow = recSheet.getLastRow();
      for (let row = recLastRow; row >= 3; row--) {
        let checkCell = recSheet.getRange(row, checkCol);
        let checkValue = checkCell.getValue();
        // If the checkbox is checked (true)
        if (checkValue === true) {
          // Delete the cell in the "Unassigned" column (shifting cells upward)
          recSheet.getRange(row, unassignedCol).deleteCells(SpreadsheetApp.Dimension.ROWS);
          // Delete the cell in the "✔" column (shifting cells upward)
          recSheet.getRange(row, checkCol).deleteCells(SpreadsheetApp.Dimension.ROWS);
          Logger.log(`Cleared recurring entry in ${userSs.getName()} at row ${row}.`);
        }
      }
      
    } catch (e) {
      Logger.log(`Error processing dashboard ${userDashboardId}: ${e.message}`);
    }
  });
  
  Logger.log("Recurring merged entries cleared on all user dashboards.");
}

/***************************************************************
 * extractIdFromUrl:
 * Extracts a document ID from a typical Google Drive URL.
 ***************************************************************/
function extractIdFromUrl(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (let r of patterns) {
    const match = r.exec(url);
    if (match && match[1]) return match[1];
  }
  return null;
}













/***************************************************************
 * MASTER FUNCTION: mergeRecurringTransactionsToFinancialJournal
 * 
 * 1) For each user in "Users" sheet (Admin Dashboard),
 *    open their "Recurring Transactions" & "Financial Journal".
 * 2) Build an accountStatusMap (nickname -> Manual/Imported).
 * 3) For each recurring transaction row:
 *    - If "Manual", check if today is due; if so, unshift a new row 
 *      onto finData with Merge="→" (if no duplicate).
 *    - If "Imported", do the old merging logic in handleImportedRecurring().
 * 4) At the end, we do ONE setValues(...) to rewrite the entire finData 
 *    back into the sheet (row 3 onward).
 * 5) We copy formatting (excluding the Receipt column) from row 3 (template)
 *    onto all the newly updated rows so date/amount columns keep consistent 
 *    formatting. Then we re-apply receipt RichText so hyperlinks stay #1155cc.
 ***************************************************************/
function mergeRecurringTransactionsToFinancialJournal() {
  Logger.log("=== Starting mergeRecurringTransactionsToFinancialJournal ===");
  const adminSs = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found => Aborting.');
    return;
  }

  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("No user rows found => Aborting.");
    return;
  }
  
  // Identify the "Dashboard" column
  const usersHeaders = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol = usersHeaders.indexOf("Dashboard");
  if (dashCol < 0) {
    Logger.log('No "Dashboard" label found => Aborting.');
    return;
  }
  
  // Read all user dashboard URLs
  const dashboardUrls = usersSheet.getRange(3, dashCol + 1, lastRow - 2, 1).getValues();
  Logger.log(`Found ${dashboardUrls.length} user rows to process.`);

  dashboardUrls.forEach((rowVal, i) => {
    const dashboardUrl = rowVal[0];
    if (!dashboardUrl) {
      Logger.log(`Row ${i + 3}: blank dashboard URL => skip`);
      return;
    }

    const userDashboardId = extractIdFromUrl(dashboardUrl);
    if (!userDashboardId) {
      Logger.log(`Row ${i+3}: invalid dashboard URL => skip`);
      return;
    }
    
    try {
      Logger.log(`Row ${i+3}: Opening user dashboard ID: ${userDashboardId}`);
      const userSs  = SpreadsheetApp.openById(userDashboardId);
      const recSheet= userSs.getSheetByName("Recurring Transactions");
      const finSheet= userSs.getSheetByName("Financial Journal");
      if (!recSheet || !finSheet) {
        Logger.log(`Missing sheets => skip user`);
        return;
      }

      // Build { nickname -> status }
      const accountStatusMap = getAccountStatusMap(userSs);
      Logger.log(`Built accountStatusMap => ${JSON.stringify(accountStatusMap)}`);

      // Build col maps
      const recHeaders = recSheet.getRange(2, 1, 1, recSheet.getLastColumn()).getValues()[0];
      const finHeaders = finSheet.getRange(2, 1, 1, finSheet.getLastColumn()).getValues()[0];
      const recMap = buildRecurringColMap(recHeaders);
      const finMap = buildFinJournalColMap(finHeaders);

      // Read "Recurring Transactions"
      const recLastRow = recSheet.getLastRow();
      if (recLastRow < 3) {
        Logger.log(`No data in "Recurring Transactions" => skip user`);
        return;
      }
      const recDataRange = recSheet.getRange(3, 1, recLastRow - 2, recSheet.getLastColumn());
      const recData = recDataRange.getValues();
      let recReceiptRT = null;
      if (recMap.receipt >= 0) {
        recReceiptRT = recSheet.getRange(3, recMap.receipt + 1, recLastRow - 2, 1).getRichTextValues();
      }

      // Read "Financial Journal"
      const finLastRow = finSheet.getLastRow();
      let finData = [];
      let finReceiptRT = null;
      if (finLastRow >= 3) {
        const finDataRange = finSheet.getRange(3, 1, finLastRow - 2, finSheet.getLastColumn());
        finData = finDataRange.getValues();
        if (finMap.receipt >= 0) {
          finReceiptRT = finSheet.getRange(3, finMap.receipt + 1, finLastRow - 2, 1).getRichTextValues();
        }
      }

      Logger.log(`Found ${recData.length} recurring rows, ${finData.length} existing FJ rows.`);

      // For each recurring row
      recData.forEach((recRow, recIdx) => {
        const unassignedVal = recRow[recMap.unassigned];
        if (!unassignedVal || unassignedVal.toString().trim() === "") return;

        const recFromVal = (recRow[recMap.from]||"").trim();
        const recToVal   = (recRow[recMap.to]  ||"").trim();
        const fromStatus = accountStatusMap[recFromVal] || "";
        const toStatus   = accountStatusMap[recToVal]   || "";

        if (fromStatus==="Manual" || toStatus==="Manual") {
          Logger.log(`Row ${recIdx+3}: "Manual" => handleManualRecurring => from="${recFromVal}", to="${recToVal}"`);
          handleManualRecurring(
            recIdx, recRow, recData, recReceiptRT, recMap,
            finMap, finData, finReceiptRT
          );
        } else {
          Logger.log(`Row ${recIdx+3}: "Imported" => handleImportedRecurring => from="${recFromVal}", to="${recToVal}"`);
          handleImportedRecurring(
            recIdx, recRow, recData, recReceiptRT, recMap,
            finMap, finData, finReceiptRT
          );
        }
      });

      // (F) Write final data
      if (finData.length > 0) {
        Logger.log(`Writing updated finData of length ${finData.length} to row 3...`);
        finSheet.getRange(3, 1, finData.length, finSheet.getLastColumn()).setValues(finData);

        // If we have matching finReceiptRT length, write them
        if (finReceiptRT && finReceiptRT.length === finData.length) {
          Logger.log(`Applying finReceiptRT for ${finReceiptRT.length} rows (initial).`);
          finSheet
            .getRange(3, finMap.receipt + 1, finData.length, 1)
            .setRichTextValues(finReceiptRT);
        }

        // (G) Copy formatting from row 3 template, EXCEPT skipping the receipt column
        const templateRow = finSheet.getRange(3, 1, 1, finSheet.getLastColumn());
        const rowStart = 3;
        const rowEnd   = 3 + finData.length - 1;

        // copy columns BEFORE receipt column
        const recColIndex = (finMap.receipt >= 0) ? (finMap.receipt + 1) : null;
        if (recColIndex && recColIndex > 1) {
          // from col 1 .. recColIndex-1
          Logger.log(`Copying format from row 3, col 1..${recColIndex-1} => rows ${rowStart}..${rowEnd}`);
          templateRow
            .offset(0, 0, 1, recColIndex-1)
            .copyFormatToRange(finSheet, 1, recColIndex-1, rowStart, rowEnd);
        }

        // copy columns AFTER receipt column
        if (recColIndex && recColIndex < finSheet.getLastColumn()) {
          Logger.log(`Copying format from row 3, col ${recColIndex+1}..end => rows ${rowStart}..${rowEnd}`);
          templateRow
            .offset(0, recColIndex, 1, finSheet.getLastColumn()-recColIndex)
            .copyFormatToRange(
              finSheet,
              recColIndex+1, finSheet.getLastColumn(),
              rowStart, rowEnd
            );
        } else if (recColIndex===null) {
          // no receipt column => just copy entire row
          Logger.log(`No "Receipt" col => copying entire row format from 3 => rows ${rowStart}..${rowEnd}`);
          templateRow.copyFormatToRange(finSheet, 1, finSheet.getLastColumn(), rowStart, rowEnd);
        }

        // (H) Finally re-apply the receipt rich text to ensure #1155cc color
        if (finReceiptRT && finReceiptRT.length === finData.length && recColIndex) {
          Logger.log(`Re-applying finReceiptRT after copyFormat, for ${finReceiptRT.length} rows`);
          finSheet
            .getRange(rowStart, recColIndex, finData.length, 1)
            .setRichTextValues(finReceiptRT);
        }

      }

      Logger.log(`Done merging recurring transactions for ${userSs.getName()}.`);

    } catch (err) {
      Logger.log(`Row ${i+3}: Error => ${err.message}`);
    }
  });

  Logger.log("=== Finished mergeRecurringTransactionsToFinancialJournal ===");
}

/***************************************************************
 * handleManualRecurring
 * 
 * Only insert if "today" is in the schedule AND we have
 * no existing row with the same date/from/to/amount/merge="→".
 * Then we unshift a new row into finData so it appears at "top".
 ***************************************************************/
function handleManualRecurring(
  recIdx, recRow, recData, recReceiptRT, recMap,
  finMap, finData, finReceiptRT
) {
  const recStart = parseDate(recRow[recMap.startDate]);
  if (!recStart) return;

  const recEnd   = parseDate(recRow[recMap.endDate]) || null;
  const freq     = (recRow[recMap.frequency] || "").trim();
  const recFrom  = (recRow[recMap.from] || "").trim();
  const recTo    = (recRow[recMap.to]   || "").trim();
  const recDesc  = recRow[recMap.desc]       || "";
  const recTags  = recRow[recMap.tags]       || "";
  const recCat   = recRow[recMap.category]   || "";
  const recSub   = recRow[recMap.subcat]     || "";
  const recAmt   = recRow[recMap.amount]     || "";

  const today = new Date(); 
  today.setHours(0,0,0,0);
  const todayStr = formatMdy(today);
  const expectedDates = getExpectedRecurringDates(recStart, freq, recEnd, today);
  if (!expectedDates.includes(todayStr)) {
    Logger.log(`ManualRecurring => not due today => skip`);
    return;
  }

  // Check duplicates
  const recAmtNum = parseAmount(recAmt);
  let foundDup = false;
  for (let i=0; i<finData.length; i++){
    const row = finData[i];

    // unify date
    let rowDateVal = row[finMap.date];
    let finDateStr = "";
    if (rowDateVal instanceof Date) {
      finDateStr = formatMdy(rowDateVal);
    } else if (rowDateVal) {
      finDateStr = rowDateVal.toString();
    }

    // unify from/to as lowercase
    const finFrom = (row[finMap.fromCol]||"").toString().trim().toLowerCase();
    const finTo   = (row[finMap.toCol]  ||"").toString().trim().toLowerCase();

    let rowMergeVal = (row[finMap.merge] == null) ? "" : row[finMap.merge].toString();
    const finMerge  = rowMergeVal.trim();
    const finAmt    = parseAmount(row[finMap.amount]||"");

    if (finDateStr===todayStr
        && finFrom===recFrom.toLowerCase()
        && finTo===recTo.toLowerCase()
        && Math.abs(finAmt - recAmtNum) < 0.000001
        && finMerge==="→") {
      Logger.log(`=> found duplicate => skip insertion`);
      foundDup=true; 
      break;
    }
  }
  if (foundDup) return;

  Logger.log(`=> Adding new manual row => date="${todayStr}" from="${recFrom}" to="${recTo}" amt="${recAmt}"`);

  // Build new row
  let finalReceipt="";
  let finalRT=null;
  if (recReceiptRT && recReceiptRT[recIdx] && recReceiptRT[recIdx][0]) {
    finalRT=applyHyperlinkStyle(recReceiptRT[recIdx][0], true);
    finalReceipt=finalRT.getText();
  } else {
    finalReceipt=(recRow[recMap.receipt]||"").toString();
  }

  const newRow=[
    recTags,      // Tags
    todayStr,     // Date
    recDesc,      // Description
    recFrom,      // From
    recTo,        // To
    recAmt,       // Amount
    recCat,       // Category
    recSub,       // Subcategory
    finalReceipt, // Receipt
    "→"           // Merge
  ];

  // Insert at front
  finData.unshift(newRow);
  if (finReceiptRT) {
    if (finalRT) {
      finReceiptRT.unshift([finalRT]);
    } else {
      finReceiptRT.unshift([null]);
    }
  }
}

/***************************************************************
 * handleImportedRecurring
 ***************************************************************/
function handleImportedRecurring(
  recIdx, recRow, recData, recReceiptRT, recMap,
  finMap, finData, finReceiptRT
) {
  const recStart = parseDate(recRow[recMap.startDate]);
  if (!recStart) return;
  const recEnd   = parseDate(recRow[recMap.endDate]) || null;
  const freq     = (recRow[recMap.frequency] || "").trim();
  const recAmt   = recRow[recMap.amount];
  const recTo    = (recRow[recMap.to]||"").trim().toLowerCase();
  const recFrom  = (recRow[recMap.from]||"").trim().toLowerCase();

  for (let i=0; i<finData.length; i++){
    let row=finData[i];
    let rowMergeVal = row[finMap.merge];
    if (rowMergeVal && rowMergeVal.toString().trim()==="⇆") continue;

    let finDateVal=row[finMap.date];
    let dateObj=(finDateVal instanceof Date)?finDateVal: parseDate(finDateVal);
    if (!dateObj) continue;
    let finDateStr=formatMdy(dateObj);
    let expectedDates=getExpectedRecurringDates(recStart,freq,recEnd,dateObj);
    if (!expectedDates.includes(finDateStr)) continue;

    if(Math.abs(parseAmount(recAmt)-parseAmount(row[finMap.amount]))>0.000001) continue;

    let recToShort=(recTo.length<5?recTo:recTo.substring(0,5));
    let finTo=(row[finMap.toCol]||"").toString().trim().toLowerCase();
    if(!finTo.includes(recToShort)) continue;

    let finFrom=(row[finMap.fromCol]||"").toString().trim().toLowerCase();
    if(finFrom!==recFrom) continue;

    Logger.log(`=> doMergeRecurring => recRow=${recIdx+3}, finRow=${i+3}`);
    doMergeRecurring(recIdx,i, recData, finData, recReceiptRT, recMap, finMap, finReceiptRT);
  }
}

/***************************************************************
 * doMergeRecurring
 ***************************************************************/
function doMergeRecurring(
  recIdx, finIdx,
  recData, finData, recReceiptRT,
  recMap, finMap, finReceiptRT
) {
  const recRow=recData[recIdx];
  const finRow=finData[finIdx];

  if(finMap.tags>=0 && recMap.tags>=0 && !finRow[finMap.tags]){
    finRow[finMap.tags]=recRow[recMap.tags];
  }
  if(finMap.fromCol>=0 && recMap.from>=0 && !finRow[finMap.fromCol]){
    finRow[finMap.fromCol]=recRow[recMap.from];
  }
  if(finMap.toCol>=0 && recMap.to>=0 && !finRow[finMap.toCol]){
    finRow[finMap.toCol]=recRow[recMap.to];
  }
  if(finMap.amount>=0 && recMap.amount>=0 && !finRow[finMap.amount]){
    finRow[finMap.amount]=recRow[recMap.amount];
  }
  if(finMap.category>=0 && recMap.category>=0 && !finRow[finMap.category]){
    finRow[finMap.category]=recRow[recMap.category];
  }
  if(finMap.subcat>=0 && recMap.subcat>=0 && !finRow[finMap.subcat]){
    finRow[finMap.subcat]=recRow[recMap.subcat];
  }
  // override desc
  if(finMap.desc>=0 && recMap.desc>=0){
    finRow[finMap.desc]=recRow[recMap.desc];
  }
  // override receipt
  if(finMap.receipt>=0 && recMap.receipt>=0){
    if(recReceiptRT && recReceiptRT[recIdx] && recReceiptRT[recIdx][0]){
      let recurringRT=recReceiptRT[recIdx][0];
      let styledRT=applyHyperlinkStyle(recurringRT,true);
      if(finReceiptRT){
        finReceiptRT[finIdx][0]=styledRT;
      }
      finRow[finMap.receipt]=styledRT.getText();
    } else {
      finRow[finMap.receipt]=recRow[recMap.receipt];
    }
  }
  // set merge="⇆"
  if(finMap.merge>=0){
    finRow[finMap.merge]="⇆";
  }
}

/***************************************************************
 * getExpectedRecurringDates
 ***************************************************************/
function getExpectedRecurringDates(recStart,frequency,recEnd,cutoffDate){
  let dates=[];
  let cutoff=recEnd?(recEnd<cutoffDate?recEnd:cutoffDate):cutoffDate;
  let current=new Date(recStart.getTime());
  frequency=frequency.toLowerCase();
  while(current<=cutoff){
    dates.push(formatMdy(current));
    if(frequency==="semi-monthly"){
      let second=addDays(current,14);
      if(second.getMonth()===current.getMonth()&&second<=cutoff){
        dates.push(formatMdy(second));
      }
      current=new Date(current.getFullYear(),current.getMonth()+1,recStart.getDate());
    }else{
      switch(frequency){
        case "daily":
          current=addDays(current,1);
          break;
        case "weekly":
          current=addDays(current,7);
          break;
        case "biweekly":
          current=addDays(current,14);
          break;
        case "monthly":
          current=new Date(current.getFullYear(),current.getMonth()+1,recStart.getDate());
          break;
        case "quarterly":
          current=new Date(current.getFullYear(),current.getMonth()+3,recStart.getDate());
          break;
        case "annually":
          current=new Date(current.getFullYear()+1,current.getMonth(),recStart.getDate());
          break;
        default:
          current=new Date(cutoff.getTime()+1);
          break;
      }
    }
  }
  return dates;
}

function addDays(d,n){
  let r=new Date(d);
  r.setDate(r.getDate()+n);
  return r;
}

/***************************************************************
 * getAccountStatusMap
 ***************************************************************/
function getAccountStatusMap(userSs){
  const accSheet=userSs.getSheetByName("Accounts");
  if(!accSheet)return{};
  const headerRow=accSheet.getRange(2,1,1,accSheet.getLastColumn()).getValues()[0];
  const nickIdx=headerRow.indexOf("Nickname:");
  const statIdx=headerRow.indexOf("Status:");
  if(nickIdx<0||statIdx<0)return{};
  let map={};
  const lr=accSheet.getLastRow();
  if(lr<3)return map;
  const data=accSheet.getRange(3,1,lr-2,accSheet.getLastColumn()).getValues();
  data.forEach(r=>{
    const n=(r[nickIdx]||"").toString().trim();
    const s=(r[statIdx]||"").toString().trim();
    if(n&&s)map[n]=s;
  });
  return map;
}

/***************************************************************
 * buildRecurringColMap
 ***************************************************************/
function buildRecurringColMap(headers){
  return{
    unassigned: headers.indexOf("Unassigned"),
    tags:       headers.indexOf("Tags"),
    frequency:  headers.indexOf("Frequency"),
    startDate:  headers.indexOf("Start Date"),
    endDate:    headers.indexOf("End Date"),
    desc:       headers.indexOf("Description"),
    from:       headers.indexOf("From"),
    to:         headers.indexOf("To"),
    amount:     headers.indexOf("Amount"),
    category:   headers.indexOf("Category"),
    subcat:     headers.indexOf("Subcategory"),
    receipt:    headers.indexOf("Receipt"),
    check:      headers.indexOf("✔")
  };
}

/***************************************************************
 * buildFinJournalColMap
 ***************************************************************/
function buildFinJournalColMap(headers){
  return{
    tags:     headers.indexOf("Tags"),
    date:     headers.indexOf("Date"),
    desc:     headers.indexOf("Description"),
    fromCol:  headers.indexOf("From"),
    toCol:    headers.indexOf("To"),
    amount:   headers.indexOf("Amount"),
    category: headers.indexOf("Category"),
    subcat:   headers.indexOf("Subcategory"),
    receipt:  headers.indexOf("Receipt"),
    merge:    headers.indexOf("Merge")
  };
}

/***************************************************************
 * formatMdy
 ***************************************************************/
function formatMdy(dateVal){
  if(!(dateVal instanceof Date))return dateVal;
  const m=dateVal.getMonth()+1;
  const d=dateVal.getDate();
  const y=dateVal.getFullYear();
  return`${m}/${d}/${y}`;
}

/***************************************************************
 * parseAmount / parseDate
 ***************************************************************/
function parseAmount(val){
  if(!val)return 0;
  let str=val.toString().replace(/\$/g,"").trim();
  let num=parseFloat(str);
  return isNaN(num)?0:num;
}
function parseDate(val){
  if(!val)return null;
  if(val instanceof Date)return val;
  let dt=new Date(val);
  return isNaN(dt.getTime())?null:dt;
}

/***************************************************************
 * extractIdFromUrl
 ***************************************************************/
function extractIdFromUrl(url){
  const patterns=[
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for(let r of patterns){
    let m=r.exec(url);
    if(m&&m[1])return m[1];
  }
  return null;
}

/***************************************************************
 * applyLinkStyleToAll
 ***************************************************************/
function applyLinkStyleToAll(rt2D){
  return rt2D.map(rowArr=>{
    const rtv=rowArr[0];
    if(!rtv)return[null];
    const styled=applyHyperlinkStyle(rtv,true);
    return[styled];
  });
}

/***************************************************************
 * applyHyperlinkStyle
 ***************************************************************/
function applyHyperlinkStyle(rtv, underlineOn){
  const runs=rtv.getRuns();
  if(!runs||!runs.length)return rtv;
  let builder=rtv.copy();
  const baseStyle=SpreadsheetApp.newTextStyle()
    .setForegroundColor("#1155cc")
    .setBold(false)
    .setUnderline(underlineOn)
    .build();
  runs.forEach(run=>{
    if(run.getLinkUrl()){
      builder.setTextStyle(run.getStartIndex(),run.getEndIndex(),baseStyle);
    }
  });
  return builder.build();
}
























/***************************************************************
 * MAIN function from Admin Dashboard:
 * merges for ALL users and logs merges in "Merged" sheet.
 ***************************************************************/
function bulkMergeManualAndPlaidAllUsers() {
  const adminSs = getAdminSpreadsheet(); // or SpreadsheetApp.getActiveSpreadsheet(), etc.
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found in Admin Dashboard.');
    return;
  }

  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("No user rows found in 'Users' sheet.");
    return;
  }

  // Identify "Dashboard" column in row 2 of "Users" sheet
  const usersHeaders = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol = usersHeaders.indexOf("Dashboard");
  if (dashCol < 0) {
    Logger.log('No "Dashboard" label found in row 2 of "Users" sheet.');
    return;
  }

  // For each user
  const urlValues = usersSheet.getRange(3, dashCol + 1, lastRow - 2, 1).getValues();
  urlValues.forEach((rowVal, i) => {
    const dashboardUrl = rowVal[0];
    if (!dashboardUrl) return;

    const userDashboardId = extractIdFromUrl(dashboardUrl);
    if (!userDashboardId) {
      Logger.log(`Row ${i+3}: Invalid Dashboard URL => skipping.`);
      return;
    }

    try {
      const userSs = SpreadsheetApp.openById(userDashboardId);
      // Merge & rewrite in the user’s “Financial Journal”
      mergeManualAndPlaidInDashboard_Rewrite(userSs);
    } catch (err) {
      Logger.log(`Row ${i+3}: Could not open user dashboard => ${err.message}`);
    }
  });

  Logger.log("Done merging manual vs. plaid transactions for all users.");
}

/***************************************************************
 * MERGE & REWRITE approach in a single user's "Financial Journal"
 *  - Preserves receipt hyperlinks
 *  - Logs merges in "Merged" sheet
 *  - Carries forward the original Plaid Transaction ID
 ***************************************************************/
function mergeManualAndPlaidInDashboard_Rewrite(userSs) {
  const finSheet = userSs.getSheetByName("Financial Journal");
  if (!finSheet) {
    Logger.log(`No "Financial Journal" sheet in ${userSs.getName()}.`);
    return;
  }

  // 1) Ensure "Plaid Txn ID" column exists, then build fresh colMap
  const headerRow = 2;
  let headers     = finSheet.getRange(headerRow, 1, 1, finSheet.getLastColumn()).getValues()[0];
  let colMap      = buildFinJournalColMap(headers);

  if (colMap.plaidId < 0) {
    finSheet.insertColumnAfter(headers.length);
    finSheet.getRange(headerRow, headers.length + 1).setValue("Plaid Txn ID");
    headers = finSheet.getRange(headerRow, 1, 1, finSheet.getLastColumn()).getValues()[0];
    colMap  = buildFinJournalColMap(headers);
  }

  // 2) Pull in all data + receipt RichText
  const lastRow      = finSheet.getLastRow();
  const dataStartRow = headerRow + 1;
  const numRows      = lastRow - headerRow;
  if (numRows < 1) return;

  const allData    = finSheet.getRange(dataStartRow, 1, numRows, finSheet.getLastColumn()).getValues();
  let allReceiptRT = null;
  if (colMap.receipt >= 0) {
    allReceiptRT = finSheet
      .getRange(dataStartRow, colMap.receipt + 1, numRows, 1)
      .getRichTextValues();
  }

  // 3) Classify each row
  const rowObjs = allData.map((rowData, idx) => ({
    rowIndex: idx,
    rowData:  rowData,
    mergeVal: (rowData[colMap.merge] || "") + ""
  }));
  const manualRows       = [];
  const plaidRows        = [];
  const userLabeledGroups= {};
  const deletedSet       = new Set();

  rowObjs.forEach(r => {
    const m = r.mergeVal;
    if (m === "⇆" || m === "→" || m === "╌") return;   // already merged
    if (m === "-") manualRows.push(r);                // manual
    else if (m === "") plaidRows.push(r);             // plaid
    else {                                            // user-labeled
      if (!userLabeledGroups[m]) userLabeledGroups[m] = [];
      userLabeledGroups[m].push(r);
    }
  });

  // 4a) Merge manual → plaid
  manualRows.forEach(man => {
    if (deletedSet.has(man.rowIndex)) return;
    const match = plaidRows.find(pl =>
      !deletedSet.has(pl.rowIndex) &&
      transactionsAreMatch(man.rowData, pl.rowData, colMap)
    );
    if (match) {
      doMergeAndMarkDeleted(
        man.rowIndex, match.rowIndex,
        allData, allReceiptRT,
        colMap, deletedSet,
        userSs
      );
    }
  });

  // 4b) Merge user-labeled pairs
  for (const label in userLabeledGroups) {
    const group = userLabeledGroups[label];
    for (let i = 0; i + 1 < group.length; i += 2) {
      const a = group[i], b = group[i+1];
      if (!deletedSet.has(a.rowIndex) && !deletedSet.has(b.rowIndex)) {
        doMergeAndMarkDeleted(
          a.rowIndex, b.rowIndex,
          allData, allReceiptRT,
          colMap, deletedSet,
          userSs
        );
      }
    }
  }

  // 5) Rewrite only the non-deleted rows
  const finalRows = [], finalRT = allReceiptRT ? [] : null;
  allData.forEach((r,i) => {
    if (!deletedSet.has(i)) {
      finalRows.push(r);
      if (finalRT) finalRT.push(allReceiptRT[i]);
    }
  });

  const newCount = finalRows.length;
  if (newCount) {
    finSheet
      .getRange(dataStartRow, 1, newCount, finalRows[0].length)
      .setValues(finalRows);
    if (finalRT) {
      finSheet
        .getRange(dataStartRow, colMap.receipt+1, newCount, 1)
        .setRichTextValues(applyLinkStyleToAll(finalRT));
    }
  }

  // 6) Clear leftover rows
  const leftover = numRows - newCount;
  if (leftover > 0) {
    finSheet
      .getRange(dataStartRow + newCount, 1, leftover, finSheet.getLastColumn())
      .clearContent();
  }

  Logger.log(`Merges done in ${userSs.getName()}, final row count: ${newCount}.`);
}

/***************************************************************
 * doMergeAndMarkDeleted: merges two rows, preserves Plaid Txn ID
 ***************************************************************/
function doMergeAndMarkDeleted(
  manIdx,
  plaidIdx,
  allData,
  allReceiptRT,
  colMap,
  deletedSet,
  userSs
) {
  const manRow = allData[manIdx], plRow = allData[plaidIdx];

  // Decide which to keep
  const keepIdx = (countNonEmpty(manRow,colMap) >= countNonEmpty(plRow,colMap))
                  ? manIdx : plaidIdx;
  const delIdx  = (keepIdx === manIdx) ? plaidIdx : manIdx;

  const keepRow = allData[keepIdx], delRow = allData[delIdx];

  // Earliest date
  const kd = parseDate(keepRow[colMap.date]), dd = parseDate(delRow[colMap.date]);
  if (dd && kd && dd < kd) keepRow[colMap.date] = delRow[colMap.date];

  // Fill blanks
  [
    colMap.tags, colMap.date, colMap.desc,
    colMap.fromCol, colMap.toCol, colMap.amount,
    colMap.category, colMap.subcat, colMap.receipt
  ].forEach(ci => {
    if (ci < 0) return;
    if (!keepRow[ci]) keepRow[ci] = delRow[ci];
  });

  // —— Preserve Plaid Txn ID —— 
  if (colMap.plaidId >= 0) {
    if (!keepRow[colMap.plaidId] && delRow[colMap.plaidId]) {
      keepRow[colMap.plaidId] = delRow[colMap.plaidId];
    }
  }

  // Mark merged
  keepRow[colMap.merge] = "⇆";
  delRow[colMap.merge]  = "⇆";

  // Preserve receipt hyperlink
  if (allReceiptRT && colMap.receipt >= 0) {
    const keepRT = allReceiptRT[keepIdx][0], delRT = allReceiptRT[delIdx][0];
    if (isPlainRichText(keepRT) && !isPlainRichText(delRT)) {
      allReceiptRT[keepIdx][0] = delRT;
    }
  }

  // Log merge
  logMergedTransaction(userSs, keepRow, delRow, colMap, allReceiptRT, keepIdx);

  // Mark for deletion
  deletedSet.add(delIdx);
}


/***************************************************************
 * logMergedTransaction => writes a row in "Merged" sheet (row3 headers)
 ***************************************************************/
function logMergedTransaction(
  userSs,
  keepRow,
  delRow,
  colMap,
  allReceiptRT,
  keepIdx
) {
  const mergedSheet = userSs.getSheetByName("Merged");
  if (!mergedSheet) {
    Logger.log('No "Merged" sheet found => skipping logging merges.');
    return;
  }

  // "Merged" sheet headers in row 3
  const mergedHeaderRow = 3;
  const mergedHeaders = mergedSheet
    .getRange(mergedHeaderRow, 1, 1, mergedSheet.getLastColumn())
    .getValues()[0];

  const tagsIdx     = mergedHeaders.indexOf("Tags");
  const dateIdx     = mergedHeaders.indexOf("Date");
  const descIdx     = mergedHeaders.indexOf("Description");
  const fromIdx     = mergedHeaders.indexOf("From");
  const toIdx       = mergedHeaders.indexOf("To");
  const amtIdx      = mergedHeaders.indexOf("Amount");
  const catIdx      = mergedHeaders.indexOf("Category");
  const subIdx      = mergedHeaders.indexOf("Subcategory");
  const recIdx      = mergedHeaders.indexOf("Receipt");
  const mergedTxIdx = mergedHeaders.indexOf("Merged Transaction");
  const unmergeIdx  = mergedHeaders.indexOf("Un-Merge?");

  if (tagsIdx < 0 || dateIdx < 0 || mergedTxIdx < 0) {
    Logger.log('Missing some columns in "Merged" => cannot log properly.');
    return;
  }

  // Build final row
  const rowData = new Array(mergedHeaders.length).fill("");

  // Fill from keepRow
  rowData[tagsIdx] = keepRow[colMap.tags] || "";
  rowData[dateIdx] = keepRow[colMap.date] || "";
  rowData[descIdx] = keepRow[colMap.desc] || "";
  rowData[fromIdx] = keepRow[colMap.fromCol] || "";
  rowData[toIdx]   = keepRow[colMap.toCol]   || "";
  rowData[amtIdx]  = keepRow[colMap.amount]  || "";
  rowData[catIdx]  = keepRow[colMap.category]|| "";
  rowData[subIdx]  = keepRow[colMap.subcat]  || "";

  // Build "Merged Transaction" => data from the "deleted" row
  const dDate = formatMdy(delRow[colMap.date]); 
  const dFrom = delRow[colMap.fromCol] || "";
  const dTo   = delRow[colMap.toCol]   || "";
  const dAmt  = delRow[colMap.amount]  || "";
  const dCat  = delRow[colMap.category]|| "";
  const dSub  = delRow[colMap.subcat]  || "";
  const mergedTxString = `${dDate} > ${dFrom} > ${dTo} > ${dAmt} > ${dCat} > ${dSub}`;
  rowData[mergedTxIdx] = mergedTxString;

  // "Un-Merge?" => place false + a checkbox
  if (unmergeIdx >= 0) {
    rowData[unmergeIdx] = false; 
  }

  // Insert a new row in "Merged"
  const nextRow = mergedSheet.getLastRow() + 1;
  mergedSheet.getRange(nextRow, 1, 1, rowData.length).setValues([ rowData ]);

  // If "Un-Merge?" col exist => checkbox
  if (unmergeIdx >= 0) {
    const cell = mergedSheet.getRange(nextRow, unmergeIdx + 1);
    cell.insertCheckboxes();
    cell.setValue(false);
  }

  // Copy keep row's "Receipt" hyperlink
  if (recIdx >= 0 && allReceiptRT && colMap.receipt >= 0) {
    const keepRT = allReceiptRT[keepIdx][0];
    if (keepRT && !isPlainRichText(keepRT)) {
      let styled = applyHyperlinkStyle(keepRT, true);
      mergedSheet
        .getRange(nextRow, recIdx + 1)
        .setRichTextValue(styled);
    } else {
      let receiptText = keepRow[colMap.receipt] || "";
      mergedSheet
        .getRange(nextRow, recIdx + 1)
        .setValue(receiptText);
    }
  }

  Logger.log(`Logged merge at row ${nextRow} in "Merged" of ${userSs.getName()}.`);
}


/***************************************************************
 * The main difference from your old code is that we do NOT call
 * range.clearContent() before rewriting finalRows. Instead we
 * directly setValues(finalRows) over the top portion, then
 * only clear leftover if the new result is shorter.
 ***************************************************************/


/***************************************************************
 * The rest is the same (utilities, matching logic, etc.)
 ***************************************************************/
function formatMdy(dateVal) {
  if (!(dateVal instanceof Date)) {
    return dateVal;
  }
  let m = dateVal.getMonth() + 1;
  let d = dateVal.getDate();
  let y = dateVal.getFullYear();
  return `${m}/${d}/${y}`;
}

function applyLinkStyleToAll(rt2D) {
  return rt2D.map(rowArr => {
    const rtv = rowArr[0];
    if (!rtv) return [ null ];
    const styled = applyHyperlinkStyle(rtv, false); 
    return [ styled ];
  });
}

function applyHyperlinkStyle(rtv, underlineOn = true) {
  const runs = rtv.getRuns();
  if (!runs || runs.length === 0) return rtv;

  let builder = rtv.copy();

  runs.forEach(run => {
    if (run.getLinkUrl()) {
      const style = SpreadsheetApp.newTextStyle()
        .copy(run.getTextStyle())
        .setForegroundColor("#1155cc")
        .setBold(false)
        .setUnderline(underlineOn)
        .build();

      builder.setTextStyle(run.getStartIndex(), run.getEndIndex(), style);
    }
  });

  return builder.build();
}


function isPlainRichText(rtv) {
  if (!rtv) return true;
  let runs = rtv.getRuns();
  if (!runs || runs.length === 0) return true;
  for (let r of runs) {
    if (r.getLinkUrl()) return false;
  }
  return true;
}

function transactionsAreMatch(mRow, pRow, colMap) {
  const mAmt = parseAmount(mRow[colMap.amount]);
  const pAmt = parseAmount(pRow[colMap.amount]);
  if (Math.abs(mAmt - pAmt) > 0.000001) return false;

  const mDate = parseDate(mRow[colMap.date]);
  const pDate = parseDate(pRow[colMap.date]);
  if (!mDate || !pDate) return false;
  const diffDays = (pDate - mDate) / 86400000;
  if (diffDays < -1 || diffDays > 5) return false;

  const mTo = (mRow[colMap.toCol] || "").trim().toLowerCase();
  const pTo = (pRow[colMap.toCol] || "").trim().toLowerCase();
  let shortStr = mTo.substring(0, 4);
  if (mTo.length < 4) shortStr = mTo;
  if (!shortStr || !pTo.includes(shortStr)) return false;

  const mFrom = (mRow[colMap.fromCol] || "").trim().toLowerCase();
  const pFrom = (pRow[colMap.fromCol] || "").trim().toLowerCase();
  return (mFrom === pFrom);
}

function parseAmount(val) {
  if (!val) return 0;
  let str = val.toString().replace(/\$/g, "").trim();
  let num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === "string") {
    let dt = new Date(val);
    if (!isNaN(dt.getTime())) return dt;
  }
  return null;
}

function countNonEmpty(row, colMap) {
  let c = 0;
  for (let k in colMap) {
    if (k === "merge") continue;
    const idx = colMap[k];
    if (idx < 0) continue;
    let val = row[idx];
    if (val !== "" && val != null) {
      c++;
    }
  }
  return c;
}

/***************************************************************
 * buildFinJournalColMap — now includes a `plaidId` field
 ***************************************************************/
function buildFinJournalColMap(headers) {
  return {
    tags:     headers.indexOf("Tags"),
    date:     headers.indexOf("Date"),
    desc:     headers.indexOf("Description"),
    fromCol:  headers.indexOf("From"),
    toCol:    headers.indexOf("To"),
    amount:   headers.indexOf("Amount"),
    category: headers.indexOf("Category"),
    subcat:   headers.indexOf("Subcategory"),
    receipt:  headers.indexOf("Receipt"),
    merge:    headers.indexOf("Merge"),
    plaidId:  headers.indexOf("Plaid Txn ID")
  };
}

function extractIdFromUrl(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (let r of patterns) {
    const match = r.exec(url);
    if (match && match[1]) return match[1];
  }
  return null;
}






















/********************************************************************
 * 1) TOP-LEVEL function in the Admin Control Panel
 ********************************************************************/
function bulkUnMergeAllCheckedRows() {
  const adminSs = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet found in Admin Dashboard.');
    return;
  }

  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("No user rows found in 'Users' sheet.");
    return;
  }

  // Identify "Dashboard" column in row 2
  const usersHeaders = usersSheet
    .getRange(2, 1, 1, usersSheet.getLastColumn())
    .getValues()[0];
  const dashCol = usersHeaders.indexOf("Dashboard");
  if (dashCol < 0) {
    Logger.log('No "Dashboard" label found in row 2 of "Users" sheet.');
    return;
  }

  // For each user row, un-merge
  const dashboardUrls = usersSheet.getRange(3, dashCol+1, lastRow-2, 1).getValues();
  dashboardUrls.forEach((rowVal, i) => {
    const dashUrl = rowVal[0];
    if (!dashUrl) return;

    const dashId = extractIdFromUrl(dashUrl);
    if (!dashId) {
      Logger.log(`Row ${i+3}: Invalid Dashboard URL => skipping.`);
      return;
    }

    try {
      const userSs = SpreadsheetApp.openById(dashId);
      unMergeAllCheckedRowsInDashboard(userSs);
    } catch (err) {
      Logger.log(`Row ${i+3}: Could not open user dashboard => ${err.message}`);
    }
  });

  Logger.log("Done un-merging for all users.");
}

/********************************************************************
 * 2) In each user's Dashboard, un-merge rows in "Merged" that have 
 *    "Un-Merge?"=true, REPLACING the merged row in the "Financial Journal" 
 *    with 2 separate rows. Then we DELETE the row from "Merged" entirely.
 ********************************************************************/
function unMergeAllCheckedRowsInDashboard(userSs) {
  const mergedSheet = userSs.getSheetByName("Merged");
  if (!mergedSheet) {
    Logger.log(`No "Merged" sheet in ${userSs.getName()} => skipping un-merge.`);
    return;
  }

  // Identify columns in row 3
  const mergedHeaderRow = 3;
  const mergedHeaders = mergedSheet
    .getRange(mergedHeaderRow, 1, 1, mergedSheet.getLastColumn())
    .getValues()[0];

  const tagsIdx     = mergedHeaders.indexOf("Tags");
  const dateIdx     = mergedHeaders.indexOf("Date");
  const descIdx     = mergedHeaders.indexOf("Description");
  const fromIdx     = mergedHeaders.indexOf("From");
  const toIdx       = mergedHeaders.indexOf("To");
  const amtIdx      = mergedHeaders.indexOf("Amount");
  const catIdx      = mergedHeaders.indexOf("Category");
  const subIdx      = mergedHeaders.indexOf("Subcategory");
  const recIdx      = mergedHeaders.indexOf("Receipt");
  const mergedTxIdx = mergedHeaders.indexOf("Merged Transaction");
  const unmergeIdx  = mergedHeaders.indexOf("Un-Merge?");

  if (dateIdx < 0 || mergedTxIdx < 0 || unmergeIdx < 0) {
    Logger.log(`Missing columns in "Merged" for ${userSs.getName()}.`);
    return;
  }

  // read row4 downward
  const mergedLastRow = mergedSheet.getLastRow();
  const mergedNumData = mergedLastRow - mergedHeaderRow;
  if (mergedNumData < 1) {
    Logger.log(`No data in "Merged" for ${userSs.getName()}.`);
    return;
  }

  // normal values
  const mergedRange = mergedSheet.getRange(
    mergedHeaderRow+1, 1, mergedNumData, mergedSheet.getLastColumn()
  );
  const mergedData = mergedRange.getValues();

  // Also read RichText for "Receipt" if recIdx >= 0
  let mergedReceiptRT = null;
  if (recIdx >= 0) {
    mergedReceiptRT = mergedSheet.getRange(
      mergedHeaderRow+1, recIdx+1, mergedNumData, 1
    ).getRichTextValues();
  }

  // Find which rows have "Un-Merge?"=true
  const rowsToUnmerge = [];
  mergedData.forEach((row, i) => {
    if (row[unmergeIdx] === true) {
      rowsToUnmerge.push(i);
    }
  });
  if (rowsToUnmerge.length === 0) {
    Logger.log(`No 'Un-Merge?' checks in "Merged" for ${userSs.getName()}.`);
    return;
  }

  // Now open "Financial Journal"
  const finSheet = userSs.getSheetByName("Financial Journal");
  if (!finSheet) {
    Logger.log(`No "Financial Journal" in ${userSs.getName()}. Skipping un-merge.`);
    return;
  }

  const finHeaders = finSheet
    .getRange(2,1,1,finSheet.getLastColumn())
    .getValues()[0];
  const colMap = buildFinJournalColMap(finHeaders);
  if (colMap.merge < 0) {
    Logger.log(`No "Merge" col in "Financial Journal" => can't un-merge for ${userSs.getName()}.`);
    return;
  }

  const finLastRow = finSheet.getLastRow();
  const finNum = finLastRow - 2; // row3 onward
  if (finNum < 1) {
    Logger.log(`No data in "Financial Journal" for ${userSs.getName()}.`);
    return;
  }
  const finRange = finSheet.getRange(3,1, finNum, finSheet.getLastColumn());
  const finValues = finRange.getValues();

  // Build a dictionary of { nickname => status } from "Accounts"
  const accountStatusMap = buildAccountStatusMap(userSs);

  // We'll store merges to process
  const mergesToProcess = [];

  rowsToUnmerge.forEach(rowIdx => {
    const rowData = mergedData[rowIdx];
    // build keepObj
    const keepObj = {
      tags: rowData[tagsIdx],
      date: rowData[dateIdx],
      desc: rowData[descIdx],
      from: rowData[fromIdx],
      to:   rowData[toIdx],
      amt:  rowData[amtIdx],
      cat:  rowData[catIdx],
      sub:  rowData[subIdx],
      rec:  rowData[recIdx] // text version
    };
    let keepRT = null;
    if (mergedReceiptRT && mergedReceiptRT[rowIdx]) {
      keepRT = mergedReceiptRT[rowIdx][0];
    }

    // parse Merged Transaction
    const mergedTx = rowData[mergedTxIdx] || "";
    const delObj = parseMergedTransactionString(mergedTx);

    // find merged row in finValues
    const foundIndex = findMergedRowIndex(finValues, keepObj, colMap);
    if (foundIndex < 0) {
      Logger.log(`Row ${rowIdx+4}: no matching "⇆" row => skipping unmerge.`);
      return;
    }

    mergesToProcess.push({
      mergedRowIndex: foundIndex,
      keepObj,
      keepRT,
      delObj,
      mergedSheetRow: mergedHeaderRow + 1 + rowIdx // actual row in "Merged" to DELETE
    });
  });

  // sort by descending mergedRowIndex so we do bottom-first
  mergesToProcess.sort((a,b)=> b.mergedRowIndex - a.mergedRowIndex);

  // We'll track rows in "Merged" to remove
  const mergedRowsToRemove = [];

  mergesToProcess.forEach(item => {
    const rowIndex = item.mergedRowIndex;
    const sheetRow = 3 + rowIndex;

    // Insert 2 blank rows above sheetRow
    finSheet.insertRowsBefore(sheetRow, 2);

    // Make 2 new row arrays
    // The "del" row -> no special logic
    const delArr  = buildNewRowFromObj(
      item.delObj, 
      finHeaders, 
      colMap, 
      "", 
      accountStatusMap
    );

    // The "keep" row => previously "unmerged manual transaction"
    // We pass a placeholder merge val "manual?" but buildNewRowFromObj 
    // will decide "╌" or "→" based on the account status
    const keepArr = buildNewRowFromObj(
      item.keepObj, 
      finHeaders, 
      colMap, 
      "manual?", 
      accountStatusMap
    );

    // place them
    finSheet.getRange(sheetRow, 1, 2, finHeaders.length)
      .setValues([ delArr, keepArr ]);

    // set RichText for the keep row's receipt if we have keepRT
    if (colMap.receipt>=0 && item.keepRT && !isPlainRichText(item.keepRT)) {
      let styled = applyHyperlinkStyle(item.keepRT, true /* underline */);
      finSheet
        .getRange(sheetRow+1, colMap.receipt+1)
        .setRichTextValue(styled);
    }

    // old merged row is now 2 rows down => delete that row
    finSheet.deleteRow(sheetRow+2);

    // remember we want to remove the row from "Merged"
    mergedRowsToRemove.push(item.mergedSheetRow);
  });

  // remove "Merged" rows from bottom up
  mergedRowsToRemove.sort((a,b)=> b - a);
  mergedRowsToRemove.forEach(r => {
    mergedSheet.deleteRow(r);
  });

  Logger.log(`In-place un-merge done in ${userSs.getName()}. Processed: ${mergesToProcess.length} merges, removed rows from Merged.`);
}

/**
 * buildNewRowFromObj => for "Financial Journal" columns
 * 
 * If mergeVal === "manual?", we look at 'From' and 'To' 
 * in the accountStatusMap:
 *   (1) If from and to are both non-blank and 
 *       at least one is "Manual" => "→" else "╌"
 *   (2) else if from is non-blank => check from only
 *   (3) else if to is non-blank => check to only
 *   (4) if none found => "╌"
 * 
 * Otherwise we just set mergeVal as-is.
 */
function buildNewRowFromObj(obj, finHeaders, colMap, mergeVal, accountStatusMap) {
  let rowArr = new Array(finHeaders.length).fill("");

  // Fill standard columns
  if (colMap.tags     >= 0) rowArr[colMap.tags]     = obj.tags || "";
  if (colMap.date     >= 0) rowArr[colMap.date]     = obj.date || "";
  if (colMap.desc     >= 0) rowArr[colMap.desc]     = obj.desc || "";
  if (colMap.fromCol  >= 0) rowArr[colMap.fromCol]  = obj.from || "";
  if (colMap.toCol    >= 0) rowArr[colMap.toCol]    = obj.to || "";
  if (colMap.amount   >= 0) rowArr[colMap.amount]   = obj.amt || "";
  if (colMap.category >= 0) rowArr[colMap.category] = obj.cat || "";
  if (colMap.subcat   >= 0) rowArr[colMap.subcat]   = obj.sub || "";
  if (colMap.receipt  >= 0) rowArr[colMap.receipt]  = obj.rec || "";

  // Now figure out what to do with the Merge col
  if (colMap.merge >= 0) {
    // If not our special placeholder, use it directly
    if (mergeVal !== "manual?") {
      rowArr[colMap.merge] = mergeVal;
    } else {
      // We want to check both from & to in accountStatusMap
      const fromAcct = obj.from || "";
      const toAcct   = obj.to   || "";

      let finalVal = "╌"; // default
      if (fromAcct && toAcct) {
        // Both are non-blank => if EITHER is "Manual," we do "→"
        const fromStat = accountStatusMap[fromAcct] || "";
        const toStat   = accountStatusMap[toAcct]   || "";
        if (fromStat === "Manual" || toStat === "Manual") {
          finalVal = "→";
        }
      }
      else if (fromAcct) {
        // check the from
        const fromStat = accountStatusMap[fromAcct] || "";
        if (fromStat === "Manual") {
          finalVal = "→";
        }
      }
      else if (toAcct) {
        // check the to
        const toStat = accountStatusMap[toAcct] || "";
        if (toStat === "Manual") {
          finalVal = "→";
        }
      }

      rowArr[colMap.merge] = finalVal;
    }
  }

  return rowArr;
}


/********************************************************************
 * parseMergedTransactionString => "M/D/YYYY > from > to > amt > cat > sub"
 ********************************************************************/
function parseMergedTransactionString(str) {
  const parts = str.split(">");
  let obj = {
    date: "",
    from: "",
    to: "",
    amt: "",
    cat: "",
    sub: "",
    rec: ""
  };
  if (parts.length >= 6) {
    obj.date = parts[0].trim();
    obj.from = parts[1].trim();
    obj.to   = parts[2].trim();
    obj.amt  = parts[3].trim();
    obj.cat  = parts[4].trim();
    obj.sub  = parts[5].trim();
  }
  return obj;
}

/********************************************************************
 * findMergedRowIndex => find row in finValues => "⇆" + matching fields
 ********************************************************************/
function findMergedRowIndex(finValues, keepObj, colMap) {
  for (let i=0; i<finValues.length; i++) {
    const row = finValues[i];
    if (row[colMap.merge] !== "⇆") continue;

    const dateMatch = (row[colMap.date]||"").toString() === (keepObj.date||"").toString();
    const fromMatch = (row[colMap.fromCol]||"") === (keepObj.from||"");
    const toMatch   = (row[colMap.toCol]||"")   === (keepObj.to||"");
    const amtMatch  = (row[colMap.amount]||"").toString() === (keepObj.amt||"").toString();
    const catMatch  = (row[colMap.category]||"") === (keepObj.cat||"");
    const subMatch  = (row[colMap.subcat]||"")   === (keepObj.sub||"");
    if (dateMatch && fromMatch && toMatch && amtMatch && catMatch && subMatch) {
      return i;
    }
  }
  return -1;
}

/********************************************************************
 * Checking if RichTextValue is plain or has link
 ********************************************************************/
function isPlainRichText(rtv) {
  if (!rtv) return true;
  const runs = rtv.getRuns();
  if (!runs || runs.length===0) return true;
  for (let r of runs) {
    if (r.getLinkUrl()) return false;
  }
  return true;
}

/********************************************************************
 * applyHyperlinkStyle => color=#1155cc, bold=false, underline=...
 ********************************************************************/
function applyHyperlinkStyle(rtv, underlineOn) {
  const runs = rtv.getRuns();
  if (!runs || runs.length === 0) return rtv;

  const builder = rtv.copy();
  let baseStyleBuilder = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#1155cc")
    .setBold(false);
  if (underlineOn) {
    baseStyleBuilder.setUnderline(true);
  } else {
    baseStyleBuilder.setUnderline(false);
  }
  const baseStyle = baseStyleBuilder.build();

  runs.forEach(run => {
    if (run.getLinkUrl()) {
      builder.setTextStyle(run.getStartIndex(), run.getEndIndex(), baseStyle);
    }
  });
  return builder.build();
}

/********************************************************************
 * buildFinJournalColMap => same as merges
 ********************************************************************/
function buildFinJournalColMap(headers) {
  return {
    tags: headers.indexOf("Tags"),
    date: headers.indexOf("Date"),
    desc: headers.indexOf("Description"),
    fromCol: headers.indexOf("From"),
    toCol: headers.indexOf("To"),
    amount: headers.indexOf("Amount"),
    category: headers.indexOf("Category"),
    subcat: headers.indexOf("Subcategory"),
    receipt: headers.indexOf("Receipt"),
    merge: headers.indexOf("Merge")
  };
}

/********************************************************************
 * Build a dictionary { nickname => status } from the "Accounts" sheet.
 * We assume row2 has "Nickname:" and "Status:" headers somewhere.
 ********************************************************************/
function buildAccountStatusMap(userSs) {
  const sheet = userSs.getSheetByName("Accounts");
  if (!sheet) {
    Logger.log(`No "Accounts" sheet in ${userSs.getName()}. Returning empty map.`);
    return {};
  }
  // find columns
  const row2vals = sheet.getRange(2,1,1,sheet.getLastColumn()).getValues()[0];
  const nicknameCol = row2vals.indexOf("Nickname:");
  const statusCol   = row2vals.indexOf("Status:");
  if (nicknameCol < 0 || statusCol < 0) {
    Logger.log(`Could not find "Nickname:" or "Status:" in row 2 => empty map.`);
    return {};
  }
  const lastRow = sheet.getLastRow();
  if (lastRow<3) {
    Logger.log(`No data in "Accounts" for ${userSs.getName()}.`);
    return {};
  }
  // read data from row3 down
  const numRows = lastRow - 2;
  const dataRange = sheet.getRange(3,1,numRows, sheet.getLastColumn()).getValues();
  let map = {};
  dataRange.forEach(row => {
    const nick = row[nicknameCol];
    const stat = row[statusCol];
    if (nick && stat) {
      map[nick.toString()] = stat.toString();
    }
  });
  return map;
}

/********************************************************************
 * Typical doc ID extraction from a URL
 ********************************************************************/
function extractIdFromUrl(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (let r of patterns) {
    const match = r.exec(url);
    if (match && match[1]) return match[1];
  }
  return null;
}






























//PRESERVE AND REAPPLY FILTERS TO FILTERED DATA

function whoOwsWhoAllScenarios() {
  importToWhoOwesWho_Scenario1();
  importToWhoOwesWho_Scenario2();
  importToWhoOwesWho_Scenario3();
  importToWhoOwesWho_Scenario4();
}

function preserveFilters(sheet) {
  var filter = sheet.getFilter();
  if (filter) {
    var filterSettings = {
      range: filter.getRange().getA1Notation(),
      criteria: {}
    };
    
    for (var i = 1; i <= filter.getRange().getNumColumns(); i++) {
      var criteria = filter.getColumnFilterCriteria(i);
      if (criteria) {
        filterSettings.criteria[i] = criteria.copy();
      }
    }

    // Remove the existing filter
    filter.remove();

    return filterSettings;
  } else {
    // If there is no filter, return null
    return null;
  }
}

function reapplyFilters(sheet, filterSettings) {
  if (filterSettings) {
    // Recreate the filter
    var range = sheet.getRange(filterSettings.range);
    var filter = range.createFilter();
    for (var columnIndex in filterSettings.criteria) {
      filter.setColumnFilterCriteria(parseInt(columnIndex), filterSettings.criteria[columnIndex]);
    }
  }
}

function extractIdFromUrl(url) {
  // This regex pattern is designed to match the various possible formats of Google Docs and Forms URLs
  const regexPatterns = [
    /\/d\/([a-zA-Z0-9_-]+)/, // Matches most Google Drive and Docs formats
    /\/forms\/d\/e\/([a-zA-Z0-9_-]+)/, // Matches some Forms URLs
    /id=([a-zA-Z0-9_-]+)/ // Matches URLs that end with an ID parameter
  ];

  for (let regex of regexPatterns) {
    const matches = regex.exec(url);
    if (matches && matches[1]) {
      return matches[1];
    }
  }

  return null; // Return null if no ID could be extracted
}


















/**
 * Main entry point from the Admin Dashboard.
 * For each user, we:
 *   1) Read the user’s "Dashboard" & "Transaction Form" columns from "Users" sheet.
 *   2) Open their User Dashboard, read named ranges:
 *        - AccountNicknames
 *        - Categories
 *        - Subcategories
 *        - EntitiesList
 *        - Tags
 *   3) Open the user’s Transaction Form.
 *   4) Update *all* matching questions for:
 *        - "Where is the transaction coming From?"   (LIST)
 *        - "Where is the transaction going To?"      (LIST)
 *        - "Please categorize your transaction:"     (LIST)
 *        - "Please add a subcategory:"               (LIST)
 *        - "Associated Entity:"                      (LIST)
 *        - "Select applicable tags:"                 (CHECKBOX)
 */
function updateUsersTransactionForms() {
  
  syncTagsListAcrossSheets();
  alphabetizeListsAllUsers();

  const adminSs    = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) return;

  // locate Dashboard & Form columns in row 2
  const hdrs         = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashColIndex = hdrs.indexOf("Dashboard");
  const formColIndex = hdrs.indexOf("Transaction Form");
  if (dashColIndex < 0 || formColIndex < 0) {
    Logger.log("Missing Dashboard or Transaction Form columns");
    return;
  }

  // loop user rows
  const rows = usersSheet.getRange(3, 1, usersSheet.getLastRow() - 2, usersSheet.getLastColumn()).getValues();
  rows.forEach((r, i) => {
    const dashUrl = r[dashColIndex];
    const formUrl = r[formColIndex];
    if (!dashUrl || !formUrl) return;

    const dashId = extractIdFromUrl(dashUrl);
    const formId = extractIdFromUrl(formUrl);
    if (!dashId || !formId) {
      Logger.log(`Row ${i+3}: bad URL`);
      return;
    }

    try {
      const userSs    = SpreadsheetApp.openById(dashId);
      const form      = FormApp.openById(formId);
      const sheet     = userSs.getSheetByName("YOUR_SHEET_WITH_TAGS_NAMED_RANGE"); // <- usually the “Settings” or similar

      // grab the named range object
      const tagsRange = userSs.getRangeByName("Tags");
      let tagsList = [];
      if (tagsRange) {
        // must call getRange on the sheet itself
        const tagsSheet   = tagsRange.getSheet();
        const names       = tagsRange.getValues().flat();
        const shows       = tagsSheet
          .getRange(tagsRange.getRow(), tagsRange.getColumn()+1,
                    tagsRange.getNumRows(), 1)
          .getValues().flat();
        names.forEach((tag,j) => {
          if (shows[j] === true) tagsList.push(tag);
        });
      }

      // the rest is unchanged…
      const accountNicknames = getNamedRangeValuesAsList(userSs, "AccountNicknames");
      const categories       = getNamedRangeValuesAsList(userSs, "Categories");
      const subcategories    = getNamedRangeValuesAsList(userSs, "Subcategories");
      const entitiesList     = getNamedRangeValuesAsList(userSs, "EntitiesList");

      updateAllFormDropdowns(form, "Where is the transaction coming From?", accountNicknames);
      updateAllFormDropdowns(form, "Where is the transaction going To?",   accountNicknames);
      updateAllFormDropdowns(form, "Please categorize your transaction:",  categories);
      updateAllFormDropdowns(form, "Please add a subcategory:",            subcategories);
      updateAllFormDropdowns(form, "Associated Entity:",                   entitiesList);

      updateAllFormCheckboxes(form, "Select applicable tags:", tagsList);

      Logger.log(`Row ${i+3}: form updated`);
    } catch (e) {
      Logger.log(`Row ${i+3}: Error updating form — ${e.message}`);
    }
  });
}


/**
 * Helper: Attempt to read a named range from userSs, returning as a string array.
 * If the named range doesn’t exist or is empty, we return [].
 */
function getNamedRangeValuesAsList(userSs, rangeName) {
  try {
    const range = userSs.getRangeByName(rangeName);
    if (!range) {
      Logger.log(`Named range '${rangeName}' not found in ${userSs.getName()}. Returning [].`);
      return [];
    }
    // Flatten the 2D array => 1D list
    const vals2D = range.getValues();
    let list = [];
    vals2D.forEach(row => {
      row.forEach(cell => {
        if (cell !== "" && cell != null) {
          list.push(cell.toString());
        }
      });
    });
    return list;
  } catch (e) {
    Logger.log(`Error reading named range '${rangeName}': ${e.message}`);
    return [];
  }
}

/**
 * For each user dashboard listed on the Admin “Users” sheet,
 * scan the Tags column on FJ, Recurring Transactions, and Who Owes Who?
 * and append any new tag names to the master Tags list (named range “Tags”),
 * inserting an empty checkbox in the adjacent “Show in Form?” column.
 */
function syncTagsListAcrossSheets() {
  const adminSs    = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) throw new Error("Users sheet not found");

  // find “Dashboard” column in row 2
  const hdrs      = usersSheet.getRange(2,1,1,usersSheet.getLastColumn()).getValues()[0];
  const dashCol   = hdrs.indexOf("Dashboard");
  if (dashCol < 0) throw new Error("‘Dashboard’ column not found in Users sheet");

  // pull each dashboard URL
  const urls      = usersSheet.getRange(3, dashCol+1, usersSheet.getLastRow()-2).getValues();
  urls.forEach((r,i) => {
    const url = r[0];
    if (!url) return;
    const dashId = extractIdFromUrl(url);
    if (!dashId) {
      Logger.log(`Row ${i+3}: invalid dashboard URL`);
      return;
    }
    try {
      const userSs     = SpreadsheetApp.openById(dashId);
      const tagsRange  = userSs.getRangeByName("Tags");
      if (!tagsRange) {
        Logger.log(`Row ${i+3}: no named range “Tags”`);
        return;
      }
      const tagSheet   = tagsRange.getSheet();
      const tagCol     = tagsRange.getColumn();
      const showCol    = tagCol + 1;

      // existing tags
      let existing = tagsRange.getValues().flat().filter(String);
      let tagSet   = new Set(existing);

      // sheets to scan
      const toScan = [
        {name: "Financial Journal",    headerRow: 2},
        {name: "Recurring Transactions", headerRow: 2},
        {name: "Who Owes Who?",         headerRow: 3}
      ];

      toScan.forEach(spec => {
        const sh = userSs.getSheetByName(spec.name);
        if (!sh) return;
        const row2 = sh.getRange(spec.headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
        const tagsColIndex = row2.indexOf("Tags");
        if (tagsColIndex < 0) return;
        const found = sh
          .getRange(spec.headerRow+1, tagsColIndex+1, sh.getLastRow()-spec.headerRow)
          .getValues()
          .flat()
          .filter(String);
        found.forEach(tag => {
          if (!tagSet.has(tag)) {
            tagSet.add(tag);
            existing.push(tag);
            // append at bottom of Tags list
            const appendRow = tagsRange.getRow() + existing.length - 1;
            tagSheet.getRange(appendRow, tagCol).setValue(tag);
            // insert unchecked checkbox
            const cell = tagSheet.getRange(appendRow, showCol);
            cell.insertCheckboxes();
            cell.setValue(true);
          }
        });
      });

      Logger.log(`Row ${i+3}: synced ${existing.length - tagSet.size} new tag(s).`);
    } catch (e) {
      Logger.log(`Row ${i+3}: error syncing tags → ${e}`);
    }
  });
}

function alphabetizeListsAllUsers() {
  const adminSs    = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) return;

  // find "Dashboard" column in row 2
  const headers = usersSheet.getRange(2,1,1,usersSheet.getLastColumn()).getValues()[0];
  const dashCol = headers.indexOf("Dashboard");
  if (dashCol < 0) {
    Logger.log("No Dashboard column in Users sheet.");
    return;
  }

  const urls = usersSheet
    .getRange(3, dashCol+1, usersSheet.getLastRow()-2, 1)
    .getValues()
    .flat();

  urls.forEach((url,i) => {
    if (!url) return;
    const id = extractIdFromUrl(url);
    if (!id) {
      Logger.log(`Row ${i+3}: invalid Dashboard URL`);
      return;
    }
    try {
      const userSs    = SpreadsheetApp.openById(id);
      const listSheet = userSs.getSheetByName("Lists");
      if (listSheet) _alphabetizeListsInSheet(listSheet);
      else Logger.log(`Row ${i+3}: no Lists sheet`);
    } catch (e) {
      Logger.log(`Row ${i+3}: ${e.message}`);
    }
  });
}

function _alphabetizeListsInSheet(sheet) {
  const LABEL_ROW = 3;
  const labels = sheet.getRange(LABEL_ROW,1,1,sheet.getLastColumn())
                      .getValues()[0];
  const lists = [
    "Account Types:",
    "Account Subtypes:",
    "Categories:",
    "Subcategories:",
    "Tags:"
  ];

  lists.forEach(label => {
    const colIdx = labels.indexOf(label);
    if (colIdx < 0) return;

    const startRow = LABEL_ROW + 1;
    const lastRow  = sheet.getLastRow();
    const rowCount = lastRow - LABEL_ROW;
    if (rowCount < 1) return;

    if (label === "Tags:") {
      // two columns: [Tags] + [Show in Form?]
      const range = sheet.getRange(startRow, colIdx+1, rowCount, 2);
      let data = range.getValues().filter(r => r[0] !== "");
      data.sort((a,b) =>
        a[0].toString().toLowerCase().localeCompare(b[0].toString().toLowerCase())
      );
      range.clearContent();
      if (data.length) {
        sheet.getRange(startRow, colIdx+1, data.length, 2).setValues(data);
      }
    } else {
      // single column
      const range = sheet.getRange(startRow, colIdx+1, rowCount, 1);
      let data = range.getValues().flat().filter(v => v !== "");
      data.sort((a,b) =>
        a.toString().toLowerCase().localeCompare(b.toString().toLowerCase())
      );
      range.clearContent();
      if (data.length) {
        sheet.getRange(startRow, colIdx+1, data.length, 1)
             .setValues(data.map(v => [v]));
      }
    }
  });
}

/**
 * Update **all** single-select dropdown questions in the given form
 * with the specified title. If multiple items share the same title,
 * we update them all (rather than stopping after the first).
 */
function updateAllFormDropdowns(form, questionTitle, optionsArray) {
  // Access every ITEM of type LIST
  const items = form.getItems(FormApp.ItemType.LIST);
  let foundCount = 0;
  for (let item of items) {
    if (item.getTitle() === questionTitle) {
      const listItem = item.asListItem();
      // Convert the array to choices
      const choices = optionsArray.map(opt => listItem.createChoice(opt));
      listItem.setChoices(choices);
      foundCount++;
    }
  }
  if (foundCount === 0) {
    Logger.log(`Dropdown question not found: "${questionTitle}"`);
  } else {
    Logger.log(`Updated ${foundCount} dropdown(s) for: "${questionTitle}"`);
  }
}

/**
 * Update **all** checkbox questions in the given form
 * with the specified title. If multiple items share the same title,
 * we update them all (rather than stopping after the first).
 */
function updateAllFormCheckboxes(form, questionTitle, optionsArray) {
  // Access every ITEM of type CHECKBOX
  const items = form.getItems(FormApp.ItemType.CHECKBOX);
  let foundCount = 0;
  for (let item of items) {
    if (item.getTitle() === questionTitle) {
      const checkItem = item.asCheckboxItem();
      // Convert array to multiple-choice "checkbox" items
      const choices = optionsArray.map(opt => checkItem.createChoice(opt));
      checkItem.setChoices(choices);
      foundCount++;
    }
  }
  if (foundCount === 0) {
    Logger.log(`Checkbox question not found: "${questionTitle}"`);
  } else {
    Logger.log(`Updated ${foundCount} checkbox(es) for: "${questionTitle}"`);
  }
}

/**
 * Extract the ID portion from a typical Google Drive URL, ignoring query params, etc.
 * e.g. "https://docs.google.com/spreadsheets/d/ABcD12345/edit#gid=0" => "ABcD12345"
 */
function extractIdFromUrl(url) {
  if (!url) return "";
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (let r of patterns) {
    const match = r.exec(url);
    if (match && match[1]) return match[1];
  }
  return "";
}
































function getColumnMapping(headers) {
  return {
    timestamp: headers.indexOf("Timestamp"),
    description: headers.indexOf("Description of Transaction:"),
    description2: headers.indexOf("Description of Transaction:", headers.indexOf("Description of Transaction:") + 1),
    description3: headers.indexOf("Description of Transaction:", headers.indexOf("Description of Transaction:", headers.indexOf("Description of Transaction:") + 1) + 1),
    from: headers.indexOf("Where is the transaction coming From?"),
    fromFallback: headers.indexOf("Not seeing the appropriate From account? Please write it here:"),
    to: headers.indexOf("Where is the transaction going To?"),
    toFallback: headers.indexOf("Not seeing the appropriate To account? Please write it here:"),
    amount: headers.indexOf("Transaction Amount:"),
    category: headers.indexOf("Please categorize your transaction:"),
    subcategory: headers.indexOf("Please add a subcategory:"),
    receipt: headers.indexOf("Please upload your receipt or proof here:"),
    receipt2: headers.indexOf("Please upload your receipt or proof here:", headers.indexOf("Please upload your receipt or proof here:") + 1),
    receipt3: headers.indexOf("Please upload your receipt or proof here:", headers.indexOf("Please upload your receipt or proof here:", headers.indexOf("Please upload your receipt or proof here:") + 1) + 1),
    transactionType: headers.indexOf("What type of transaction are you recording?"),
    associatedEntity7: headers.indexOf("Associated Entity:"), 
    associatedEntity8: headers.indexOf("Associated Entity:", headers.indexOf("Associated Entity:") + 1),
    oweWho: headers.indexOf("Who do you owe?"),
    dueDate1: headers.indexOf("Due date:"),
    amountOwed1: headers.indexOf("How much do you owe them?"),
    terms1: headers.indexOf("Are there any specific terms?"),
    oweWho2: headers.indexOf("Who is the second counterparty you owe?"),
    dueDate2: headers.indexOf("Due date:", headers.indexOf("Due date:") + 1),
    amountOwed2: headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?") + 1),
    terms2: headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?") + 1),
    oweWho3: headers.indexOf("Who is the third counterparty you owe?"),
    dueDate3: headers.indexOf("Due date:", headers.indexOf("Due date:", headers.indexOf("Due date:") + 1) + 1),
    amountOwed3: headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?") + 1) + 1),
    terms3: headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?") + 1) + 1),
    oweWho4: headers.indexOf("Who is the fourth counterparty you owe?"),
    dueDate4: headers.indexOf("Due date:", headers.indexOf("Due date:", headers.indexOf("Due date:", headers.indexOf("Due date:") + 1) + 1) + 1),
    amountOwed4: headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?") + 1) + 1) + 1),
    terms4: headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?") + 1) + 1) + 1),
    oweWho5: headers.indexOf("Who is the fifth counterparty you owe?"),
    dueDate5: headers.indexOf("Due date:", headers.indexOf("Due date:", headers.indexOf("Due date:", headers.indexOf("Due date:", headers.indexOf("Due date:") + 1) + 1) + 1) + 1),
    amountOwed5: headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?", headers.indexOf("How much do you owe them?") + 1) + 1) + 1) + 1),
    terms5: headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?", headers.indexOf("Are there any specific terms?") + 1) + 1) + 1) + 1),
    oweYou: headers.indexOf("Who owes you money?"),
    dueDateYou1: headers.indexOf("Due date:", headers.indexOf("Who owes you money?") + 1),
    amountYou1: headers.indexOf("How much do they owe you?", headers.indexOf("Who owes you money?") + 1),
    termsYou1: headers.indexOf("Are there any specific terms?", headers.indexOf("Who owes you money?") + 1),
    oweYou2: headers.indexOf("Who's the second counterparty that owes you money?"),
    dueDateYou2: headers.indexOf("Due date:", headers.indexOf("Who's the second counterparty that owes you money?") + 1),
    amountYou2: headers.indexOf("How much do they owe you?", headers.indexOf("Who's the second counterparty that owes you money?") + 1),
    termsYou2: headers.indexOf("Are there any specific terms?", headers.indexOf("Who's the second counterparty that owes you money?") + 1),
    oweYou3: headers.indexOf("Who's the third counterparty that owes you money?"),
    dueDateYou3: headers.indexOf("Due date:", headers.indexOf("Who's the third counterparty that owes you money?") + 1),
    amountYou3: headers.indexOf("How much do they owe you?", headers.indexOf("Who's the third counterparty that owes you money?") + 1),
    termsYou3: headers.indexOf("Are there any specific terms?", headers.indexOf("Who's the third counterparty that owes you money?") + 1),
    oweYou4: headers.indexOf("Who's the fourth counterparty that owes you money?"),
    dueDateYou4: headers.indexOf("Due date:", headers.indexOf("Who's the fourth counterparty that owes you money?") + 1),
    amountYou4: headers.indexOf("How much do they owe you?", headers.indexOf("Who's the fourth counterparty that owes you money?") + 1),
    termsYou4: headers.indexOf("Are there any specific terms?", headers.indexOf("Who's the fourth counterparty that owes you money?") + 1),
    oweYou5: headers.indexOf("Who's the fifth counterparty that owes you money?"),
    dueDateYou5: headers.indexOf("Due date:", headers.indexOf("Who's the fifth counterparty that owes you money?") + 1),
    amountYou5: headers.indexOf("How much do they owe you?", headers.indexOf("Who's the fifth counterparty that owes you money?") + 1),
    termsYou5: headers.indexOf("Are there any specific terms?", headers.indexOf("Who's the fifth counterparty that owes you money?") + 1),
    importedToWhoOwesWho: headers.indexOf("Who Owes Who Status"), // Adjusted for "Who Owes Who?" status column
    tags: headers.indexOf("Select applicable tags:")
  };
}

function formatDate(timestamp) {
  // Check if timestamp is already a Date object
  if (timestamp instanceof Date) {
    // Format Date object to MM/DD/YYYY string
    const year = timestamp.getFullYear();
    const month = timestamp.getMonth() + 1; // getMonth() returns 0-11
    const day = timestamp.getDate();
    return `${month}/${day}/${year}`;
  } else if (typeof timestamp === 'string') {
    // Assuming the timestamp is a string and contains "MM/DD/YYYY HH:mm:ss"
    return timestamp.split(' ')[0]; // Extracts and returns "MM/DD/YYYY"
  } else {
    // If the timestamp is neither a Date object nor a string, log an error or handle as appropriate
    Logger.log("Timestamp is not in an expected format: " + timestamp);
    return "";
  }
}

function formatReceipt(receiptLink) {
  // Returns a formula string that shows "View" and links to the receipt
  if (receiptLink) {
    return `=HYPERLINK("${receiptLink}", "View")`;
  }
  return ""; // Return empty string if no link provided
}

function shareReceiptFolderWithUsers() {
  const ss = getAdminSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) throw new Error("Users sheet not found.");

  // — find or create the needed columns in row 2 —
  const hdr = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const emailCol       = hdr.indexOf("Primary Email")      + 1;
  const formRespCol    = hdr.indexOf("Form Response")      + 1;
  let sharedCol        = hdr.indexOf("Receipt Folder Shared?") + 1;
  let folderUrlCol     = hdr.indexOf("Receipt Folder URL") + 1;

  if (!sharedCol) {
    sharedCol = usersSheet.getLastColumn() + 1;
    usersSheet.getRange(2, sharedCol).setValue("Receipt Folder Shared?");
  }
  if (!folderUrlCol) {
    folderUrlCol = usersSheet.getLastColumn() + 1 + (sharedCol > hdr.length ? 0 : 1);
    usersSheet.getRange(2, folderUrlCol).setValue("Receipt Folder URL");
  }

  const lastRow = usersSheet.getLastRow();
  const data = usersSheet
    .getRange(3, 1, lastRow - 2, usersSheet.getLastColumn())
    .getValues();

  data.forEach((row, idx) => {
    const email       = row[emailCol - 1];
    const formRespUrl = row[formRespCol - 1];
    const already     = row[sharedCol - 1];

    // skip if missing info or already done
    if (!email || !formRespUrl || already) return;

    const formRespId = extractIdFromUrl(formRespUrl);
    if (!formRespId) return;

    const respSs    = SpreadsheetApp.openById(formRespId);
    const respSheet = respSs.getSheets()[0];
    const headers   = respSheet.getRange(1, 1, 1, respSheet.getLastColumn()).getValues()[0];

    // find the “Please upload your receipt or proof here:” column
    const receiptIdx = headers.indexOf("Please upload your receipt or proof here:") + 1;
    if (!receiptIdx) return;

    // grab the first non‑empty receipt URL
    const receipts = respSheet
      .getRange(2, receiptIdx, respSheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(String);
    if (receipts.length === 0) return;

    const receiptUrl = receipts[0];
    // extract file ID from either “/d/ID/” or “id=ID”
    let m = receiptUrl.match(/\/d\/([^\/]+)/);
    let fileId = m ? m[1] : (receiptUrl.split("id=")[1] || "").split("&")[0];
    if (!fileId) return;

    // find the parent folder and share it
    const file    = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    if (!parents.hasNext()) return;
    const folder  = parents.next();
    folder.addViewer(email);

    // mark as shared & record folder URL
    const rowNum = idx + 3;
    usersSheet.getRange(rowNum, sharedCol).setValue("✔");
    usersSheet.getRange(rowNum, folderUrlCol).setValue(folder.getUrl());

    Logger.log(`Shared folder ${folder.getName()} with ${email}`);
  });
}


function importToFinancialJournal() {
  const ss = getAdminSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log("No 'Users' sheet found.");
    return;
  }
  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("No users data found in 'Users' sheet.");
    return;
  }

  // For each user in the "Users" sheet
  const usersData = usersSheet.getRange("A3:F" + lastRow).getValues();
  usersData.forEach(row => {
    const userName         = row[0];
    const dashboardUrl     = row[1];
    const formResponsesUrl = row[3];
    Logger.log(`Processing Financial Journal for ${userName}`);

    if (!dashboardUrl || !formResponsesUrl) {
      Logger.log(`Dashboard or Form Responses URL missing for ${userName}. Skipping.`);
      return;
    }

    const dashboardId     = extractIdFromUrl(dashboardUrl);
    const formResponsesId = extractIdFromUrl(formResponsesUrl);
    if (!dashboardId || !formResponsesId) {
      Logger.log(`Invalid Dashboard/FormResponses URL for ${userName}. Skipping.`);
      return;
    }

    try {
      const dashboardSs         = SpreadsheetApp.openById(dashboardId);
      const financialJournalSheet = dashboardSs.getSheetByName("Financial Journal");
      if (!financialJournalSheet) {
        Logger.log(`No "Financial Journal" sheet found for ${userName}. Skipping.`);
        return;
      }

      // 1) Build a map { nickname -> status } from the user's "Accounts" sheet
      const accountStatusMap = buildAccountStatusMap(dashboardSs);

      // 2) Ensure the "Merge" column exists (by name). We'll find or create it.
      const finHeadersRange = financialJournalSheet.getRange(2, 1, 1, financialJournalSheet.getLastColumn());
      let finHeaders = finHeadersRange.getValues()[0];
      let mergeColIndex = finHeaders.indexOf("Merge");
      if (mergeColIndex < 0) {
        // Insert a new column at the end
        mergeColIndex = finHeaders.length;
        financialJournalSheet.insertColumnAfter(mergeColIndex);
        // Write "Merge" label in row 2
        financialJournalSheet.getRange(2, mergeColIndex + 1).setValue("Merge");
        // Refresh headers
        finHeaders = financialJournalSheet.getRange(2, 1, 1, financialJournalSheet.getLastColumn()).getValues()[0];
        mergeColIndex = finHeaders.indexOf("Merge");
      }

      // 3) Access the Form Responses
      const formResponsesSs     = SpreadsheetApp.openById(formResponsesId);
      const formResponsesSheet  = formResponsesSs.getSheets()[0];
      const headers             = formResponsesSheet.getRange("A1:ZZ1").getValues()[0];
      let statusColIndex        = headers.indexOf("Financial Journal Status") + 1;

      // If "Financial Journal Status" col doesn't exist, create it
      if (statusColIndex === 0) {
        statusColIndex = headers.filter(String).length + 1;
        formResponsesSheet.getRange(1, statusColIndex).setValue("Financial Journal Status");
      }

      // 4) Gather form responses
      const responsesRange = `A2:ZZ${formResponsesSheet.getLastRow()}`;
      const responsesData  = formResponsesSheet.getRange(responsesRange).getValues();
      const colMapping     = getColumnMapping(headers);

      // 5) For each response row, import if not "Imported"
      responsesData.forEach((response, index) => {
        if (!response[statusColIndex - 1]) {
          // Gather data
          const date        = formatDate(response[colMapping.timestamp]);
          const description = response[colMapping.description] || "";
          const from        = response[colMapping.from] || response[colMapping.fromFallback] || "";
          const to          = response[colMapping.to]   || response[colMapping.toFallback]   || "";
          const amount      = response[colMapping.amount] || "";
          const category    = response[colMapping.category] || response[colMapping.categoryFallback] || "";
          const receipt     = formatReceipt(response[colMapping.receipt]);

          // New fields
          const rawTags     = response[colMapping.tags] || "";
          const tags        = rawTags; // if multiple, assume comma-separated
          const subcategory = response[colMapping.subcategory] || "";

          // Minimal check for meaningful data
          if (date || description || from || to || amount || category || tags || subcategory) {
            // 6) Decide the "Merge" column value based on from/to
            //    Check accountStatusMap. If we find "Manual" => "→", if "Imported" => "-", else default to "-"
            let mergeVal = "-"; // default
            let fromStat = accountStatusMap[from] || "";
            let toStat   = accountStatusMap[to]   || "";
            // If EITHER is "Manual", then "→"
            if (fromStat === "Manual" || toStat === "Manual") {
              mergeVal = "→";
            }
            // If EITHER is "Imported", keep it as "-" (the default).
            // (If you need a tie-break rule, add logic here.)

            // 7) Insert a new row at the very top (row 3)
            financialJournalSheet.insertRowsBefore(3, 1);

            // We build the row array (assuming 9 columns plus Merge at the end). 
            // Adjust if your columns differ.
            // A=Tag, B=Date, C=Desc, D=From, E=To, F=Amount, G=Category, H=Subcat, I=Receipt, J=Merge
            const rowArray = [
              tags,        // A
              date,        // B
              description, // C
              from,        // D
              to,          // E
              amount,      // F
              category,    // G
              subcategory, // H
              receipt,     // I
              mergeVal     // J
            ];

            // 8) Write row at row 3
            financialJournalSheet
              .getRange(3, 1, 1, rowArray.length)
              .setValues([rowArray]);

            // 9) Mark as "Imported"
            formResponsesSheet.getRange(index + 2, statusColIndex).setValue("Imported");
            Logger.log(`Imported row ${index + 2} for ${userName} to Financial Journal (top row).`);
          }
        }
      });
    } catch (e) {
      Logger.log(`Error processing ${userName}: ${e.message}`);
    }
  });

  shareReceiptFolderWithUsers();

}


/***************************************************************
 * A helper function that builds { nickname -> associatedEntity }
 * from the user's "Accounts" sheet. 
 * We assume row2 has "Nickname:" and "Associated Entity:" columns.
 ***************************************************************/
function buildNicknameToEntityMap(userSs) {
  let map = {};
  const accSheet = userSs.getSheetByName("Accounts");
  if (!accSheet) {
    Logger.log(`No "Accounts" sheet => empty nickname->entity map`);
    return map;
  }

  // read row2 to find columns
  const row2vals = accSheet.getRange(2, 1, 1, accSheet.getLastColumn()).getValues()[0];
  const nickIdx   = row2vals.indexOf("Nickname:");
  const entIdx    = row2vals.indexOf("Associated Entity:");
  if (nickIdx < 0 || entIdx < 0) {
    Logger.log(`No "Nickname:" or "Associated Entity:" in row2 => cannot build map`);
    return map;
  }

  const lastRow = accSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log(`No data rows in "Accounts" => empty map`);
    return map;
  }

  const dataRange = accSheet.getRange(3, 1, lastRow - 2, accSheet.getLastColumn()).getValues();
  dataRange.forEach(row => {
    let nickname = (row[nickIdx] || "").toString().trim();
    let entity   = (row[entIdx]  || "").toString().trim();
    if (nickname) {
      map[nickname] = entity;
    }
  });

  return map;
}

/***************************************************************
 * SCENARIO 1 => "Money Exchanged" with oweWho
 * 
 * => user owes => negative amounts
 * => we check the "To" nickname first in the Accounts map; 
 *    if found, that's the entity. If not found, fallback to "From" nickname.
 ***************************************************************/
function importToWhoOwesWho_Scenario1() {
  const ss = getAdminSpreadsheet(); 
  const usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) return;

  const usersData = usersSheet.getRange("A3:F" + usersSheet.getLastRow()).getValues();
  usersData.forEach(row => {
    const dashboardUrl     = row[1];
    const formResponsesUrl = row[3];
    if (!dashboardUrl || !formResponsesUrl) return;

    const dashboardId     = extractIdFromUrl(dashboardUrl);
    const formResponsesId = extractIdFromUrl(formResponsesUrl);
    if (!dashboardId || !formResponsesId) return;

    const dashboardSs      = SpreadsheetApp.openById(dashboardId);
    const whoOwesWhoSheet  = dashboardSs.getSheetByName("Who Owes Who?");
    if (!whoOwesWhoSheet) return;

    // Build the nickname->entity map from "Accounts"
    const nicknameEntityMap = buildNicknameToEntityMap(dashboardSs);

    const formResponsesSs   = SpreadsheetApp.openById(formResponsesId);
    const formResponsesSheet= formResponsesSs.getSheets()[0];

    // Ensure "Who Owes Who Status" col
    let headers       = formResponsesSheet.getRange("A1:ZZ1").getValues()[0];
    let statusColIndex= headers.indexOf("Who Owes Who Status") + 1;
    if (statusColIndex === 0) {
      statusColIndex = headers.filter(String).length + 1;
      formResponsesSheet.getRange(1, statusColIndex).setValue("Who Owes Who Status");
    }
    
    const colMapping = getColumnMapping(headers);
    const responses  = formResponsesSheet.getRange(`A2:ZZ${formResponsesSheet.getLastRow()}`).getValues();
    
    responses.forEach((response, i) => {
      if (
        response[colMapping.transactionType] === "Money Exchanged" &&
        response[colMapping.oweWho] &&
        !response[statusColIndex - 1] // not "Done"
      ) {
        let dateCreated = formatDate(response[colMapping.timestamp]);
        let rawTags     = response[colMapping.tags] || "";
        let tags        = rawTags;
        let entity7     = response[colMapping.associatedEntity7] || "";
        let entity8     = response[colMapping.associatedEntity8] || "";
        let fallbackAssociatedEntity = entity7 || entity8;

        // "oweWho" => negative amounts
        const cpFields  = [colMapping.oweWho, colMapping.oweWho2, colMapping.oweWho3, colMapping.oweWho4, colMapping.oweWho5];
        const ddFields  = [colMapping.dueDate1, colMapping.dueDate2, colMapping.dueDate3, colMapping.dueDate4, colMapping.dueDate5];
        const amtFields = [colMapping.amountOwed1, colMapping.amountOwed2, colMapping.amountOwed3, colMapping.amountOwed4, colMapping.amountOwed5];
        const termFields= [colMapping.terms1, colMapping.terms2, colMapping.terms3, colMapping.terms4, colMapping.terms5];

        for (let j=0; j<cpFields.length; j++){
          if (response[cpFields[j]]) {
            let dueDate     = formatDate(response[ddFields[j]]);
            let counterparty= response[cpFields[j]];
            let rawAmt      = response[amtFields[j]] || "";
            let amount      = rawAmt ? -Math.abs(rawAmt) : "";
            let term        = response[termFields[j]] || "";
            let receipt     = formatReceipt(response[colMapping.receipt]);

            // *** Determine "entity" by looking up the "To" nickname first, fallback to "From" ***
            // We'll see if they answered "Where is the transaction going To?"
            // or the fallback question. 
            let toNickname   = response[colMapping.to]   || response[colMapping.toFallback]   || "";
            let fromNickname = response[colMapping.from] || response[colMapping.fromFallback] || "";

            let matchedEntity = "";
            // first try the "toNickname"
            if (toNickname && nicknameEntityMap[toNickname]) {
              matchedEntity = nicknameEntityMap[toNickname];
            } else if (fromNickname && nicknameEntityMap[fromNickname]) {
              matchedEntity = nicknameEntityMap[fromNickname];
            } else {
              // fallback from the form itself
              matchedEntity = fallbackAssociatedEntity;
            }

            whoOwesWhoSheet.appendRow([
              tags,
              matchedEntity,       // merged entity
              dateCreated,
              dueDate,
              counterparty,
              amount,
              response[colMapping.description],
              term,
              receipt
            ]);
          }
        }
        // Mark as Done
        formResponsesSheet.getRange(i+2, statusColIndex).setValue("Done");
      }
    });
  });
}

/**
 * Scenario 2 => "I Owe Money"
 * => user owes => negative amounts
 */
function importToWhoOwesWho_Scenario2() {
  const ss         = getAdminSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) return;

  const usersData  = usersSheet.getRange("A3:F" + usersSheet.getLastRow()).getValues();
  usersData.forEach(row => {
    const dashboardUrl     = row[1];
    const formResponsesUrl = row[3];
    if (!dashboardUrl || !formResponsesUrl) return;

    const dashboardId     = extractIdFromUrl(dashboardUrl);
    const formResponsesId = extractIdFromUrl(formResponsesUrl);
    if (!dashboardId || !formResponsesId) return;

    const dashboardSs     = SpreadsheetApp.openById(dashboardId);
    const whoOwesWhoSheet = dashboardSs.getSheetByName("Who Owes Who?");
    if (!whoOwesWhoSheet) return;

    const formResponsesSs   = SpreadsheetApp.openById(formResponsesId);
    const formResponsesSheet= formResponsesSs.getSheets()[0];

    // Ensure "Who Owes Who Status"
    let headers       = formResponsesSheet.getRange("A1:ZZ1").getValues()[0];
    let statusColIndex= headers.indexOf("Who Owes Who Status") + 1;
    if (statusColIndex === 0) {
      statusColIndex = headers.filter(String).length + 1;
      formResponsesSheet.getRange(1, statusColIndex).setValue("Who Owes Who Status");
    }
    
    const colMapping = getColumnMapping(headers);
    const responses  = formResponsesSheet.getRange(`A2:ZZ${formResponsesSheet.getLastRow()}`).getValues();
    
    responses.forEach((response, i) => {
      if (response[colMapping.transactionType] === "I Owe Money" &&
          response[colMapping.oweWho] &&
          !response[statusColIndex - 1]) {

        let dateCreated = formatDate(response[colMapping.timestamp]);
        let rawTags     = response[colMapping.tags] || "";
        let tags        = rawTags;
        let entity7     = response[colMapping.associatedEntity7] || "";
        let entity8     = response[colMapping.associatedEntity8] || "";
        let associatedEntity = entity7 || entity8;

        // "I Owe Money" => negative
        const counterpartyFields = ["oweWho","oweWho2","oweWho3","oweWho4","oweWho5"];
        const dueDateFields      = ["dueDate1","dueDate2","dueDate3","dueDate4","dueDate5"];
        const amountFields       = ["amountOwed1","amountOwed2","amountOwed3","amountOwed4","amountOwed5"];
        const termFields         = ["terms1","terms2","terms3","terms4","terms5"];

        counterpartyFields.forEach((field, index2) => {
          if (response[colMapping[field]]) {
            let dueDate     = formatDate(response[colMapping[dueDateFields[index2]]]);
            let counterparty= response[colMapping[field]];
            let rawAmt      = response[colMapping[amountFields[index2]]] || "";
            // negative
            let amount      = rawAmt ? -Math.abs(rawAmt) : "";
            let term        = response[colMapping[termFields[index2]]]   || "";
            let receipt     = formatReceipt(response[colMapping.receipt2]);

            whoOwesWhoSheet.appendRow([
              tags,
              associatedEntity,
              dateCreated,
              dueDate,
              counterparty,
              amount, // negative
              response[colMapping.description2], 
              term,
              receipt
            ]);
          }
        });
        
        // Mark as done
        formResponsesSheet.getRange(i + 2, statusColIndex).setValue("Done");
      }
    });
  });
}

/**
 * Scenario 3 => "Money Exchanged" but user filled "oweYou" => 
 * means they owe me => positive amounts
 */
function importToWhoOwesWho_Scenario3() {
  const ss         = getAdminSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) return;

  const usersData  = usersSheet.getRange("A3:F" + usersSheet.getLastRow()).getValues();
  
  usersData.forEach(row => {
    const dashboardUrl     = row[1];
    const formResponsesUrl = row[3];
    if (!dashboardUrl || !formResponsesUrl) return;

    const dashboardId     = extractIdFromUrl(dashboardUrl);
    const formResponsesId = extractIdFromUrl(formResponsesUrl);
    if (!dashboardId || !formResponsesId) return;

    const dashboardSs     = SpreadsheetApp.openById(dashboardId);
    const whoOwesWhoSheet = dashboardSs.getSheetByName("Who Owes Who?");
    if (!whoOwesWhoSheet) return;

    // Build nickname->entity map from "Accounts"
    const nicknameEntityMap = buildNicknameToEntityMap(dashboardSs);

    const formResponsesSs   = SpreadsheetApp.openById(formResponsesId);
    const formResponsesSheet= formResponsesSs.getSheets()[0];
    
    // "Who Owes Who Status" col
    let headers       = formResponsesSheet.getRange("A1:ZZ1").getValues()[0];
    let statusColIndex= headers.indexOf("Who Owes Who Status") + 1;
    if (statusColIndex === 0) {
      statusColIndex = headers.filter(String).length + 1;
      formResponsesSheet.getRange(1, statusColIndex).setValue("Who Owes Who Status");
    }
    
    const colMapping = getColumnMapping(headers);
    const responses  = formResponsesSheet.getRange(`A2:ZZ${formResponsesSheet.getLastRow()}`).getValues();
    
    responses.forEach((response, i) => {
      if (
        response[colMapping.transactionType] === "Money Exchanged" &&
        response[colMapping.oweYou] &&
        !response[statusColIndex - 1]
      ) {
        let dateCreated = formatDate(response[colMapping.timestamp]);
        let rawTags     = response[colMapping.tags] || "";
        let tags        = rawTags;
        let entity7     = response[colMapping.associatedEntity7] || "";
        let entity8     = response[colMapping.associatedEntity8] || "";
        let fallbackAssociatedEntity = entity7 || entity8;

        // "oweYou" => user is owed => positive amounts
        const cpFields   = ["oweYou","oweYou2","oweYou3","oweYou4","oweYou5"];
        const ddFields   = ["dueDateYou1","dueDateYou2","dueDateYou3","dueDateYou4","dueDateYou5"];
        const amtFields  = ["amountYou1","amountYou2","amountYou3","amountYou4","amountYou5"];
        const termFields = ["termsYou1","termsYou2","termsYou3","termsYou4","termsYou5"];

        cpFields.forEach((field,index3) => {
          if (response[colMapping[field]]) {
            let dueDate     = formatDate(response[colMapping[ddFields[index3]]]);
            let counterparty= response[colMapping[field]];
            let rawAmt      = response[colMapping[amtFields[index3]]] || "";
            // positive
            let amount      = rawAmt ? Math.abs(rawAmt) : "";
            let term        = response[colMapping[termFields[index3]]] || "";
            let receipt     = formatReceipt(response[colMapping.receipt]);

            // *** Determine "entity" => look up the "From" nickname first, fallback to "To" ***
            let fromNickname = response[colMapping.from] || response[colMapping.fromFallback] || "";
            let toNickname   = response[colMapping.to]   || response[colMapping.toFallback]   || "";

            let matchedEntity = "";
            // first try from
            if (fromNickname && nicknameEntityMap[fromNickname]) {
              matchedEntity = nicknameEntityMap[fromNickname];
            } else if (toNickname && nicknameEntityMap[toNickname]) {
              matchedEntity = nicknameEntityMap[toNickname];
            } else {
              matchedEntity = fallbackAssociatedEntity;
            }

            whoOwesWhoSheet.appendRow([
              tags,
              matchedEntity,
              dateCreated,
              dueDate,
              counterparty,
              amount, // positive
              response[colMapping.description],
              term,
              receipt
            ]);
          }
        });

        // Mark as done
        formResponsesSheet.getRange(i + 2, statusColIndex).setValue("Done");
      }
    });
  });
}


/**
 * Scenario 4 => "Money Owed to Me"
 * => user is owed => positive amounts
 */
function importToWhoOwesWho_Scenario4() {
  const ss         = getAdminSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) return;

  const usersData  = usersSheet.getRange("A3:F" + usersSheet.getLastRow()).getValues();
  usersData.forEach(row => {
    const dashboardUrl     = row[1];
    const formResponsesUrl = row[3];
    if (!dashboardUrl || !formResponsesUrl) return;

    const dashboardId     = extractIdFromUrl(dashboardUrl);
    const formResponsesId = extractIdFromUrl(formResponsesUrl);
    if (!dashboardId || !formResponsesId) return;

    const dashboardSs     = SpreadsheetApp.openById(dashboardId);
    const whoOwesWhoSheet = dashboardSs.getSheetByName("Who Owes Who?");
    if (!whoOwesWhoSheet) return;

    const formResponsesSs   = SpreadsheetApp.openById(formResponsesId);
    const formResponsesSheet= formResponsesSs.getSheets()[0];
    
    // Ensure "Who Owes Who Status"
    let headers       = formResponsesSheet.getRange("A1:ZZ1").getValues()[0];
    let statusColIndex= headers.indexOf("Who Owes Who Status") + 1;
    if (statusColIndex === 0) {
      statusColIndex = headers.filter(String).length + 1;
      formResponsesSheet.getRange(1, statusColIndex).setValue("Who Owes Who Status");
    }
    
    const colMapping = getColumnMapping(headers);
    const responses  = formResponsesSheet.getRange(`A2:ZZ${formResponsesSheet.getLastRow()}`).getValues();
    
    responses.forEach((response, i) => {
      if (response[colMapping.transactionType] === "Money Owed to Me" &&
          response[colMapping.oweYou] &&
          !response[statusColIndex - 1]) {

        let dateCreated    = formatDate(response[colMapping.timestamp]);
        let rawTags        = response[colMapping.tags] || "";
        let tags           = rawTags;
        let entity7        = response[colMapping.associatedEntity7] || "";
        let entity8        = response[colMapping.associatedEntity8] || "";
        let associatedEntity = entity7 || entity8;

        const counterpartyFields = ["oweYou","oweYou2","oweYou3","oweYou4","oweYou5"];
        const dueDateFields      = ["dueDateYou1","dueDateYou2","dueDateYou3","dueDateYou4","dueDateYou5"];
        const amountFields       = ["amountYou1","amountYou2","amountYou3","amountYou4","amountYou5"];
        const termFields         = ["termsYou1","termsYou2","termsYou3","termsYou4","termsYou5"];

        counterpartyFields.forEach((field, index4) => {
          if (response[colMapping[field]]) {
            let dueDate = formatDate(response[colMapping[dueDateFields[index4]]]);
            let counterparty = response[colMapping[field]];
            let rawAmt = response[colMapping[amountFields[index4]]] || "";
            // scenario4 => also positive if "Money Owed to Me"
            let amount = rawAmt ? +Math.abs(rawAmt) : "";
            let term   = response[colMapping[termFields[index4]]] || "";
            let receipt= formatReceipt(response[colMapping.receipt3]);

            whoOwesWhoSheet.appendRow([
              tags,
              associatedEntity,
              dateCreated,
              dueDate,
              counterparty,
              amount,  // positive
              response[colMapping.description3],
              term,
              receipt
            ]);
          }
        });

        // Mark as done
        formResponsesSheet.getRange(i + 2, statusColIndex).setValue("Done");
      }
    });
  });
}




































/********************************************************************
 * MAIN ENTRY: 
 * 1) Reads the "Users" sheet, row by row.
 * 2) For each row, finds "Dashboard" URL in col 2, 
 *    "Form Responses" URL in col 4, just like importToFinancialJournal.
 * 3) If both exist, we open them. Then process "Who Owes Who?" partial payments.
 ********************************************************************/
function processWhoOwesWhoPaymentsAllUsers() {
  const adminSs = getAdminSpreadsheet();
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log('No "Users" sheet in Admin Dashboard.');
    return;
  }

  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("No user rows found in 'Users' sheet.");
    return;
  }

  // We'll read all from A3:F...
  const usersData = usersSheet.getRange("A3:F" + lastRow).getValues();
  usersData.forEach((rowData, i) => {
    const userName         = rowData[0]; // col A
    const dashboardUrl     = rowData[1]; // col B
    const formResponsesUrl = rowData[3]; // col D

    Logger.log(`Processing Who Owes Who for user: ${userName}`);

    if (!dashboardUrl || !formResponsesUrl) {
      Logger.log(`Missing Dashboard or Form Responses URL => skipping user ${userName}.`);
      return;
    }

    // Extract IDs
    const dashboardId = extractIdFromUrl(dashboardUrl);
    const formResponsesId = extractIdFromUrl(formResponsesUrl);

    if (!dashboardId || !formResponsesId) {
      Logger.log(`Invalid Dashboard/FormResponses URL => skipping user ${userName}`);
      return;
    }

    try {
      const dashboardSs = SpreadsheetApp.openById(dashboardId);
      const formResponsesSs = SpreadsheetApp.openById(formResponsesId);

      // We'll read the entire form responses data
      const formRespSheet = formResponsesSs.getSheets()[0]; // assume first sheet has data
      const formRespRange = formRespSheet.getDataRange();
      const formRespData  = formRespRange.getValues();

      // Now process partial payments in "Who Owes Who?"
      processWhoOwesWhoPaymentsSingleUser(dashboardSs, formRespData);
    } catch (err) {
      Logger.log(`Error processing ${userName}: ${err.message}`);
    }
  });

  Logger.log("Done processing Who Owes Who for all users.");
}


/********************************************************************
 * MAIN LOGIC: partial payment / paid in full / waive debt 
 * in a single user's "Who Owes Who?" sheet. 
 * 
 * Key changes:
 *  - We insert new transactions at the very top of "Financial Journal"
 *    (row 3), right under the header row (row 2).
 *  - We rely on `dateCreatedStr` for cat/sub matching in the form data.
 ********************************************************************/
function processWhoOwesWhoPaymentsSingleUser(userSs, formRespData) {
  const wowSheet = userSs.getSheetByName("Who Owes Who?");
  const finSheet = userSs.getSheetByName("Financial Journal");
  if (!wowSheet || !finSheet) {
    Logger.log(`Missing "Who Owes Who?" or "Financial Journal" in ${userSs.getName()}.`);
    return;
  }

  // Identify columns in "Who Owes Who?" (headers in row 3)
  const wowHeaders = wowSheet.getRange(3, 1, 1, wowSheet.getLastColumn()).getValues()[0];
  const wowMap = buildWhoOwesWhoMap(wowHeaders);

  // Identify columns in "Financial Journal" (headers in row 2)
  const finHeaders = finSheet.getRange(2, 1, 1, finSheet.getLastColumn()).getValues()[0];
  const finMap = buildFinJournalMap(finHeaders);

  // We'll read from row4 down
  const wowLastRow = wowSheet.getLastRow();
  if (wowLastRow < 4) {
    Logger.log(`No data in "Who Owes Who?" for ${userSs.getName()}.`);
    return;
  }
  const numDataRows = wowLastRow - 3; // row4..last
  if (numDataRows < 1) return;

  // We'll store the row numbers to delete
  const rowsToDelete = [];

  // We iterate from top to bottom, collecting rows to delete,
  // but do NOT delete until after the loop (so we don't shift rows).
  for (let i = 0; i < numDataRows; i++) {
    const sheetRow = 4 + i; // actual row on the sheet
    // read that row
    let rowVals = wowSheet.getRange(sheetRow, 1, 1, wowHeaders.length).getValues()[0];

    // also read the "Receipt" as RichText if we want to preserve a hyperlink
    let rowRT = null;
    if (wowMap.receipt >= 0) {
      rowRT = wowSheet.getRange(sheetRow, wowMap.receipt + 1).getRichTextValue();
    }

    // Payment or action
    const dateVal    = rowVals[wowMap.date];    // date of partial payment
    const paymentVal = rowVals[wowMap.payment]; // partial payment
    const actionVal  = (rowVals[wowMap.action] || "").toString().trim();

    let payDate = parseDate(dateVal);
    if (!payDate) {
      // no valid pay date => skip
      continue;
    }

    // Original debt info
    const originalAmount = parseFloat(rowVals[wowMap.amount]) || 0;
    if (!originalAmount) continue; // skip if 0 or empty
    const isPositive     = (originalAmount > 0); // sign indicates who owes who
    const tags           = rowVals[wowMap.tags]       || "";
    const entity         = rowVals[wowMap.entity]     || "";
    const dateCreatedStr = rowVals[wowMap.dateCreated]|| "";
    const counterpty     = rowVals[wowMap.counterparty]|| "";
    const descr          = rowVals[wowMap.description]|| "";
    const terms          = rowVals[wowMap.terms]      || "";

    let shortDateCreated = formatMdy(parseDate(dateCreatedStr));
    let originalReceiptFormula = "";
    if (rowRT) {
      originalReceiptFormula = convertRichTextToHyperlinkFormula(rowRT);
    }

    // Payment scenarios
    const paymentNum = parseFloat(paymentVal) || 0;
    let partialPayment = (paymentNum !== 0 && !actionVal);
    let paidInFull    = (actionVal.toLowerCase() === "paid in full" && paymentNum === 0);
    let waiveDebt     = (actionVal.toLowerCase() === "waive debt"   && paymentNum === 0);

    if (!partialPayment && !paidInFull && !waiveDebt) {
      // No scenario => skip
      continue;
    }

    // Build a base descriptor
    const baseDesc = `${descr} ${shortDateCreated} ${terms} - ${entity}`.trim();

    if (partialPayment) {
      const partialDesc = `Payment towards $${formatUsd(Math.abs(originalAmount))} ${baseDesc}`;
      let fromVal = isPositive ? counterpty : "";
      let toVal   = isPositive ? "" : counterpty;

      // find cat/sub from form data
      let {foundCat, foundSub} = findCatSubFromFormResp(formRespData, dateCreatedStr, descr, counterpty);
      if (!foundCat && !foundSub) {
        const fallback = findCategorySubcategory(finSheet, finMap, dateCreatedStr, descr);
        foundCat = fallback.foundCat;
        foundSub= fallback.foundSub;
      }

      // Insert transaction
      prependFinancialJournalRow(
        finSheet, finMap,
        tags,
        payDate,
        partialDesc,
        fromVal,
        toVal,
        paymentNum,
        foundCat,
        foundSub,
        originalReceiptFormula
      );

      // reduce the "Amount"
      let newAmount = isPositive ? (originalAmount - paymentNum) : (originalAmount + paymentNum);
      if (Math.abs(newAmount) < 0.000001) {
        // effectively zero => row is fully satisfied => mark for deletion
        rowsToDelete.push(sheetRow);
      } else {
        // update the leftover in the "Amount" cell
        wowSheet.getRange(sheetRow, wowMap.amount + 1).setValue(newAmount);
        // clear Payment, Date, Action
        clearWowCells(wowSheet, sheetRow, wowMap, false);
      }
    }
    else if (paidInFull) {
      const fullDesc = `Payment towards $${formatUsd(Math.abs(originalAmount))} ${baseDesc}`;
      let fromVal = isPositive ? counterpty : "";
      let toVal   = isPositive ? "" : counterpty;

      let {foundCat, foundSub} = findCatSubFromFormResp(formRespData, dateCreatedStr, descr, counterpty);
      if (!foundCat && !foundSub) {
        const fallback = findCategorySubcategory(finSheet, finMap, dateCreatedStr, descr);
        foundCat = fallback.foundCat;
        foundSub= fallback.foundSub;
      }

      prependFinancialJournalRow(
        finSheet, finMap,
        tags,
        payDate,
        fullDesc,
        fromVal,
        toVal,
        Math.abs(originalAmount),
        foundCat,
        foundSub,
        originalReceiptFormula
      );

      // row is fully satisfied => mark for deletion
      rowsToDelete.push(sheetRow);
    }
    else if (waiveDebt) {
      const waivedDesc = `Waived payment towards $${formatUsd(Math.abs(originalAmount))} ${baseDesc}`;
      let fromVal = isPositive ? counterpty : "";
      let toVal   = isPositive ? "" : counterpty;

      let {foundCat, foundSub} = findCatSubFromFormResp(formRespData, dateCreatedStr, descr, counterpty);
      if (!foundCat && !foundSub) {
        const fallback = findCategorySubcategory(finSheet, finMap, dateCreatedStr, descr);
        foundCat = fallback.foundCat; 
        foundSub= fallback.foundSub;
      }

      prependFinancialJournalRow(
        finSheet, finMap,
        tags,
        payDate,
        waivedDesc,
        fromVal,
        toVal,
        0,
        foundCat,
        foundSub,
        originalReceiptFormula
      );
      // row is fully satisfied => mark for deletion
      rowsToDelete.push(sheetRow);
    }
  } // end for

  // After we finish the loop, we delete rows in descending order
  rowsToDelete.sort((a,b) => b - a);
  for (let rowToDel of rowsToDelete) {
    wowSheet.deleteRow(rowToDel);
  }

  Logger.log(`Done partial payments with row deletion for "Who Owes Who?" => ${userSs.getName()}`);
}

/**
 * Clears only Payment, Date, Action columns
 * (If we want to remove leftover from the scenario).
 */
function clearWowCells(wowSheet, row, wowMap, fullClear) {
  // Payment, Date, Action
  let paymentCol = wowMap.payment,
      dateCol    = wowMap.date,
      actionCol  = wowMap.action;
  if (paymentCol>=0) wowSheet.getRange(row, paymentCol+1).clearContent();
  if (dateCol>=0)    wowSheet.getRange(row, dateCol+1).clearContent();
  if (actionCol>=0)  wowSheet.getRange(row, actionCol+1).clearContent();

  // If we had something else to clear, do it here
  // For partial leftover = not zero => we do not delete the row, so we keep "Amount"
}

/********************************************************************
 * Instead of 'appendRow', we insert at the TOP (row 3) 
 * so new entries appear under the header row (row 2).
 *
 * - We'll set the row data with a "-" in the Merge column
 ********************************************************************/
function prependFinancialJournalRow(
  finSheet, finMap,
  tags, payDate, desc, fromVal, toVal, amountVal,
  catVal, subVal, receiptFormula
) {
  // Format date as M/D/YYYY
  let dateStr = formatMdy(payDate);
  // Format amount as $xx.xx
  let amountStr = `$${formatUsd(amountVal)}`;

  // Build the row array
  let rowArr = [
    tags,        // Tags
    dateStr,     // Date
    desc,        // Description
    fromVal,     // From
    toVal,       // To
    amountStr,   // Amount
    catVal || "",// Category
    subVal || "",// Subcategory
    receiptFormula || "", // Receipt => e.g. =HYPERLINK("url","View")
    "○"          // Merge => place a dash
  ];

  // Insert a new blank row at row 3
  const insertRowIndex = 3; 
  finSheet.insertRowBefore(insertRowIndex);

  // Write our row data into that brand-new row 3
  finSheet.getRange(insertRowIndex, 1, 1, rowArr.length).setValues([ rowArr ]);

  // If your "Receipt" might need RichText styling, you can check if it's a hyperlink formula.
  // Typically, =HYPERLINK("...","View") is recognized automatically as a link by Sheets,
  // so no special styling needed. If you do want to enforce #1155cc with underline, 
  // you can do a style pass here similarly to styleHyperlinksInRange, 
  // but only for the newly inserted cell.
}


/********************************************************************
 * Searching for Category/Subcategory in the form responses 
 * by matching the entire "Timestamp" date with "dateCreated" 
 * plus partial substring matching on 'descr' + 'counterpty'.
 * Detailed logging version.
 ********************************************************************/
function findCatSubFromFormResp(formRespData, dateCreated, descr, counterpty) {
  let foundCat = "";
  let foundSub = "";

  Logger.log("========== START findCatSubFromFormResp ==========");
  Logger.log(`1) Provided dateCreated='${dateCreated}', descr='${descr}', counterpty='${counterpty}'.`);

  // If we have no data or only headers
  if (!formRespData || formRespData.length < 2) {
    Logger.log("2) formRespData is empty/insufficient => returning empty cat/sub.");
    Logger.log("========== END findCatSubFromFormResp (failed early) ==========");
    return {foundCat, foundSub};
  }

  if (!dateCreated) {
    Logger.log("3) dateCreated is empty => returning empty cat/sub.");
    Logger.log("========== END findCatSubFromFormResp (failed early) ==========");
    return {foundCat, foundSub};
  }

  // parse dateCreated
  Logger.log(`4) parse dateCreated='${dateCreated}'...`);
  let createdDateObj = parseDate(dateCreated);
  if (!createdDateObj) {
    Logger.log("   => parseDate returned null => can't compare => returning empty cat/sub.");
    Logger.log("========== END findCatSubFromFormResp (failed) ==========");
    return {foundCat, foundSub};
  } else {
    Logger.log(`   => parseDate OK => ${createdDateObj.toString()}`);
  }

  let headers = formRespData[0].map(h => (h||"").toString().trim());
  Logger.log(`5) We have ${formRespData.length} rows & ${headers.length} columns. Headers:`);
  Logger.log(headers);

  let timestampCol = headers.indexOf("Timestamp");
  Logger.log(`6) timestampCol index = ${timestampCol}`);
  if (timestampCol<0) {
    Logger.log("   => no 'Timestamp' col => returning empty cat/sub.");
    Logger.log("========== END findCatSubFromFormResp ==========");
    return {foundCat, foundSub};
  }

  let catCols = [], subCols = [];
  for (let c=0; c<headers.length; c++){
    let low = headers[c].toLowerCase();
    if (low.includes("please categorize your transaction")) catCols.push(c);
    if (low.includes("please add a subcategory")) subCols.push(c);
  }
  Logger.log(`7) catCols=${catCols}, subCols=${subCols}`);

  let lowDescr = (descr||"").toLowerCase();
  let lowCount = (counterpty||"").toLowerCase();
  Logger.log(`8) Searching each row for same day + rowStr.includes('${lowDescr}') + rowStr.includes('${lowCount}')`);

  for (let r=1; r<formRespData.length; r++){
    let row = formRespData[r];
    let stampVal = (row[timestampCol]||"").toString().trim();
    Logger.log(`\n--- Checking row ${r+1}: stampVal='${stampVal}' ---`);
    if (!stampVal) {
      Logger.log("   => empty => skip row");
      continue;
    }
    let userRespDateObj = parseDate(stampVal);
    if (!userRespDateObj) {
      Logger.log(`   => can't parse => skip row`);
      continue;
    }
    if (!sameDay(createdDateObj, userRespDateObj)) {
      Logger.log(`   => not sameDay => skip row (userRespDateObj='${userRespDateObj.toString()}')`);
      continue;
    }
    Logger.log(`   => day match => both are ${userRespDateObj.toDateString()}`);

    let rowStr = row.join(" ").toLowerCase();
    if (!rowStr.includes(lowDescr) || !rowStr.includes(lowCount)) {
      Logger.log(`   => missing descr/counterpty => skip row`);
      continue;
    }
    Logger.log("   => found date+descr+counterpty => extracting cat/sub now.");

    for (let cc of catCols) {
      let val = (row[cc]||"").toString().trim();
      if (val) {
        foundCat=val;
        Logger.log(`   => foundCat='${foundCat}' col ${cc+1}`);
        break;
      }
    }
    for (let sc of subCols) {
      let val = (row[sc]||"").toString().trim();
      if (val) {
        foundSub=val;
        Logger.log(`   => foundSub='${foundSub}' col ${sc+1}`);
        break;
      }
    }
    Logger.log(`   => returning foundCat='${foundCat}', foundSub='${foundSub}'`);
    Logger.log("========== END findCatSubFromFormResp (success) ==========");
    return {foundCat, foundSub};
  }
  Logger.log("=> no matching row => return empty cat/sub");
  Logger.log("========== END findCatSubFromFormResp (none) ==========");
  return {foundCat, foundSub};
}


/********************************************************************
 * Fallback approach => search existing "Financial Journal" 
 *   for a row with the same date + description
 ********************************************************************/
function findCategorySubcategory(finSheet, finMap, dateCreated, originalDesc) {
  let foundCat="", foundSub="";
  if (!dateCreated||!originalDesc) return {foundCat, foundSub};
  let lr = finSheet.getLastRow();
  if (lr<3) return {foundCat, foundSub};

  let dataRange = finSheet.getRange(3,1, lr-2, finSheet.getLastColumn());
  let finData   = dataRange.getValues();

  for (let i=0; i<finData.length; i++){
    let row = finData[i];
    let finDate = (row[finMap.date]||"").toString().trim();
    let finDesc = (row[finMap.description]||"").toString().trim();
    if (finDate===dateCreated && finDesc===originalDesc.trim()) {
      if (finMap.category>=0)   foundCat=(row[finMap.category]||"");
      if (finMap.subcategory>=0)foundSub=(row[finMap.subcategory]||"");
      return {foundCat, foundSub};
    }
  }
  return {foundCat, foundSub};
}


/********************************************************************
 * Utilities
 ********************************************************************/
function buildWhoOwesWhoMap(headers) {
  return {
    tags:          headers.indexOf("Tags"),
    entity:        headers.indexOf("Entity"),
    dateCreated:   headers.indexOf("Date Created"),
    dueDate:       headers.indexOf("Due Date"),
    counterparty:  headers.indexOf("Counterparty"),
    amount:        headers.indexOf("Amount"),
    description:   headers.indexOf("Description"),
    terms:         headers.indexOf("Terms"),
    receipt:       headers.indexOf("Receipt"),
    payment:       headers.indexOf("Payment"),
    date:          headers.indexOf("Date"),
    action:        headers.indexOf("Action")
  };
}

function buildFinJournalMap(headers) {
  return {
    tags:         headers.indexOf("Tags"),
    date:         headers.indexOf("Date"),
    description:  headers.indexOf("Description"),
    from:         headers.indexOf("From"),
    to:           headers.indexOf("To"),
    amount:       headers.indexOf("Amount"),
    category:     headers.indexOf("Category"),
    subcategory:  headers.indexOf("Subcategory"),
    receipt:      headers.indexOf("Receipt"),
    merge:        headers.indexOf("Merge")
  };
}

// Simple parseDate helper
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  let d = new Date(val);
  if (isNaN(d.getTime())) return null;
  return d;
}

// sameDay => ignoring time
function sameDay(d1, d2) {
  return (
    d1 &&
    d2 &&
    d1.getFullYear()===d2.getFullYear() &&
    d1.getMonth()===d2.getMonth() &&
    d1.getDate()===d2.getDate()
  );
}

function formatMdy(d) {
  if (!(d instanceof Date)) return "";
  return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`;
}

function formatUsd(num) {
  return num.toFixed(2);
}

/**
 * Convert a RichTextValue's first link to a 
 * =HYPERLINK("...","View") formula. 
 */
function convertRichTextToHyperlinkFormula(rtv) {
  if (!rtv || !rtv.getRuns) return "";
  let runs = rtv.getRuns();
  for (let run of runs) {
    let link = run.getLinkUrl();
    if (link) {
      return `=HYPERLINK("${link}","View")`;
    }
  }
  return "";
}

/**
 * Make all hyperlinks #1155cc, not bold, with underline
 * Usually Google Sheets auto-detects =HYPERLINK formulas,
 * so you might not strictly need this. But in case you do:
 */
function styleHyperlinksInRange(range) {
  let rtVals = range.getRichTextValues();
  let newVals = [];
  for (let r=0; r<rtVals.length; r++){
    let rtv = rtVals[r][0];
    if (!rtv) {
      newVals.push([null]);
      continue;
    }
    let styled = applyHyperlinkStyle(rtv, true /* underline */);
    newVals.push([styled]);
  }
  range.setRichTextValues(newVals);
}

function applyHyperlinkStyle(rtv, underlineOn) {
  let runs = rtv.getRuns();
  if (!runs || runs.length===0) return rtv;
  let builder = rtv.copy();
  let baseStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#1155cc")
    .setBold(false)
    .setUnderline(underlineOn)
    .build();
  for (let run of runs) {
    if (run.getLinkUrl()) {
      builder.setTextStyle(run.getStartIndex(), run.getEndIndex(), baseStyle);
    }
  }
  return builder.build();
}


















/**
 * For a single user's spreadsheet:
 *   1) Build { nickname -> status } from "Accounts"
 *   2) Open "Financial Journal"
 *   3) Identify "Merge"/"From"/"To" columns
 *   4) For each row with Merge ∈ {"⚠","○","-","→"}, re-check "From"/"To"
 *      and set a new Merge value. Then only setValues for that Merge column.
 *      This way, we do NOT overwrite or lose hyperlinks in "Receipt."
 */
function syncMergeColBasedOnAccounts(userSs) {
  // 1) Build accountStatusMap
  const accountStatusMap = buildAccountStatusMap(userSs);

  // 2) Open "Financial Journal"
  const finSheet = userSs.getSheetByName("Financial Journal");
  if (!finSheet) {
    Logger.log(`No "Financial Journal" sheet in ${userSs.getName()} => skipping.`);
    return;
  }

  // Identify columns from row 2
  const headerRow = 2;
  const finHeaders = finSheet.getRange(headerRow, 1, 1, finSheet.getLastColumn()).getValues()[0];
  const fromIdx   = finHeaders.indexOf("From");
  const toIdx     = finHeaders.indexOf("To");
  const mergeIdx  = finHeaders.indexOf("Merge");

  if (fromIdx < 0 || toIdx < 0 || mergeIdx < 0) {
    Logger.log(`Missing "From"/"To"/"Merge" in "Financial Journal" => skipping.`);
    return;
  }

  // 3) Read row 3 downward
  const lastRow = finSheet.getLastRow();
  const numDataRows = lastRow - headerRow;
  if (numDataRows < 1) {
    Logger.log(`No data in "Financial Journal" of ${userSs.getName()}.`);
    return;
  }

  // We'll read the entire block of data, but we'll only re-write the "Merge" column
  const dataStartRow = headerRow + 1; // 3
  const dataRange = finSheet.getRange(dataStartRow, 1, numDataRows, finSheet.getLastColumn());
  const data = dataRange.getValues();

  // We'll also keep track of which Merge cells we actually change
  let changes = 0;

  for (let r = 0; r < data.length; r++) {
    // row r => actual row on sheet is dataStartRow + r
    const rowArr = data[r];
    const mergeVal = rowArr[mergeIdx];

    // Only care if mergeVal ∈ {"⚠","○","-","→"}
    if (!["⚠","○","-","→"].includes(mergeVal)) {
      continue;
    }

    // read from/to
    const fromVal = rowArr[fromIdx] || "";
    const toVal   = rowArr[toIdx]   || "";

    // Decide new Merge symbol
    let newVal = decideMergeSymbol(fromVal, toVal, accountStatusMap);

    if (newVal !== mergeVal) {
      rowArr[mergeIdx] = newVal;
      changes++;
    }
  }

  if (changes > 0) {
    // We only need to rewrite the "Merge" column
    // Build a 2D array for just that column
    const mergeUpdates = data.map(r => [ r[mergeIdx] ]);

    // The Merge column's range:
    const mergeRange = finSheet.getRange(dataStartRow, mergeIdx + 1, numDataRows, 1);

    // Now we setValues just that single column
    mergeRange.setValues(mergeUpdates);
    Logger.log(`Updated ${changes} Merge cells in "Financial Journal" of ${userSs.getName()}.`);
  } else {
    Logger.log(`No Merge cells needed update in "Financial Journal" of ${userSs.getName()}.`);
  }
}


/**
 * Loops through your Admin Dashboard "Users" sheet,
 * calls syncMergeColBasedOnAccounts(...) for each user.
 * We do NOT rewrite the entire row => so we keep hyperlinks in Receipt column intact.
 */
function syncMergeColBasedOnAccountsForAllUsers() {
  const adminSs = getAdminSpreadsheet(); // or SpreadsheetApp.getActiveSpreadsheet()
  const usersSheet = adminSs.getSheetByName("Users");
  if (!usersSheet) {
    Logger.log("No 'Users' sheet found in Admin Dashboard.");
    return;
  }

  const lastRow = usersSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("No user data found in 'Users' sheet.");
    return;
  }

  // Identify "Dashboard" col in row 2
  const headers = usersSheet.getRange(2, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const dashCol = headers.indexOf("Dashboard");
  if (dashCol < 0) {
    Logger.log("No 'Dashboard' label found in row 2 of 'Users' sheet.");
    return;
  }

  // For each user row
  const userData = usersSheet.getRange(3, dashCol + 1, lastRow - 2, 1).getValues();
  userData.forEach((rowVal, i) => {
    const dashUrl = rowVal[0];
    if (!dashUrl) return; // skip blank

    const dashId = extractIdFromUrl(dashUrl);
    if (!dashId) {
      Logger.log(`Row ${i + 3}: invalid dashboard URL => skipping.`);
      return;
    }

    try {
      const userSs = SpreadsheetApp.openById(dashId);
      syncMergeColBasedOnAccounts(userSs); 
    } catch (err) {
      Logger.log(`Row ${i + 3}: error opening user dashboard => ${err.message}`);
    }
  });

  Logger.log("Done syncing 'Merge' column for all users.");
}


/*******************************************************************
 * The rest: decideMergeSymbol, buildAccountStatusMap, etc.
 * 
 * Because we do NOT re-write the entire row, "Receipt" hyperlinks 
 * remain intact. We only update the single 'Merge' column.
 *******************************************************************/
function decideMergeSymbol(fromVal, toVal, accountStatusMap) {
  if (!fromVal || !toVal) {
    return "○";
  }
  const fromStat = accountStatusMap[fromVal] || "";
  const toStat   = accountStatusMap[toVal]   || "";
  const noneFound = (!fromStat && !toStat);
  if (noneFound) {
    return "⚠";
  }
  if (fromStat === "Imported" || toStat === "Imported") {
    return "-";
  }
  return "→";
}

function buildAccountStatusMap(userSs) {
  const accSheet = userSs.getSheetByName("Accounts");
  if (!accSheet) {
    Logger.log(`No "Accounts" sheet => returning empty map.`);
    return {};
  }
  const row2vals = accSheet.getRange(2,1,1,accSheet.getLastColumn()).getValues()[0];
  const nickIdx  = row2vals.indexOf("Nickname:");
  const statIdx  = row2vals.indexOf("Status:");
  if (nickIdx < 0 || statIdx < 0) return {};
  
  const lr = accSheet.getLastRow();
  if (lr<3) return {};

  const dataRange = accSheet.getRange(3,1, lr-2, accSheet.getLastColumn()).getValues();
  let map = {};
  dataRange.forEach(row => {
    const nick = row[nickIdx];
    const st   = row[statIdx];
    if (nick && st) {
      map[nick.toString()] = st.toString();
    }
  });
  return map;
}

function extractIdFromUrl(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /\/forms\/d\/e\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (let r of patterns) {
    const match = r.exec(url);
    if (match && match[1]) return match[1];
  }
  return null;
}

