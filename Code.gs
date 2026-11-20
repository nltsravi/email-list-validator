/**
 * Google Apps Script Web App & Sheet Validator:
 * Strict User-Level Email Validator (100% Zero-Notification Guarantee)
 *
 * Security Policy:
 * - Only authenticated users belonging to `@tirwin.in` are permitted to access
 *   the Web App UI or invoke any backend RPC function.
 * - Supports importing email lists from ANY Google Doc or Google Sheet in the
 *   user's Google Drive (or by pasting a Google Doc/Sheet URL/ID or raw text).
 * - Processes Domain + User mailbox verification in batches of 100 without sending
 *   any email to the validated addresses.
 */

const CONFIG = {
  ALLOWED_DOMAIN: 'tirwin.in',
  BATCH_SIZE: 100,

  // For legacy standalone spreadsheet execution:
  SOURCE_SHEET_NAME: 'emailids',
  EMAIL_COLUMN_INDEX: 1,
  HAS_HEADER_ROW: true,

  // Strict mode: An email is ONLY marked "Active" if the specific user mailbox is confirmed to exist
  REQUIRE_USER_VERIFICATION_FOR_ACTIVE: true,

  // Built-in silent user checks (No API key needed, sends zero emails)
  ENABLE_MICROSOFT_USER_CHECK: true,
  ENABLE_GOOGLE_USER_CHECK: true,

  // Optional SMTP Handshake API for custom non-Google/non-Microsoft SMTP servers:
  // 'NONE' | 'HUNTER' | 'ABSTRACT' | 'ZEROBOUNCE' | 'REOON'
  VERIFICATION_PROVIDER: 'NONE',
  API_KEY: ''
};

/**
 * Helper to log execution status to both Apps Script Logger and Cloud Execution Console.
 *
 * @param {string} message
 */
function logStatus(message) {
  Logger.log(message);
  console.log(message);
}

/**
 * Note for Apps Script OAuth Scope detection:
 * Ensures Apps Script grants Drive, Docs, and Sheets OAuth permissions required by the Web App.
 *//** @private */
function _ensureScopes() {
  // DriveApp.getRootFolder();
  // DocumentApp.getActiveDocument();
  // SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Retrieves runtime configuration property (checks ScriptProperties first, then CONFIG fallback).
 *
 * @param {string} key
 * @param {*} fallbackValue
 * @return {string}
 */
function getRuntimeSecretOrConfig(key, fallbackValue) {
  try {
    const props = PropertiesService.getScriptProperties();
    const val = props.getProperty(key);
    if (val !== null && val !== undefined && String(val).trim() !== '') {
      return String(val).trim();
    }
  } catch (e) {
    // Fallback to CONFIG if PropertiesService is unavailable
  }
  return String(fallbackValue !== undefined ? fallbackValue : '');
}

/**
 * Enforces that the current user is authenticated with a `@tirwin.in` account.
 * Returns a session status object and throws an Error if `throwOnUnauthorized` is true.
 *
 * @param {boolean} throwOnUnauthorized
 * @return {{authorized: boolean, email: string, domain: string, reason: string}}
 */
function verifyTirwinAccess(throwOnUnauthorized) {
  const allowedDomain = getRuntimeSecretOrConfig('ALLOWED_DOMAIN', CONFIG.ALLOWED_DOMAIN).toLowerCase();
  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const userEmail = activeEmail || effectiveEmail;

  if (!userEmail || !userEmail.includes('@')) {
    const msg = `Authentication required. Please sign in with your @${allowedDomain} Google Workspace account.`;
    if (throwOnUnauthorized) throw new Error(msg);
    return { authorized: false, email: '', domain: '', reason: msg };
  }

  const userDomain = userEmail.split('@')[1];
  if (userDomain !== allowedDomain) {
    const msg = `Access Denied: "${userEmail}" belongs to "@${userDomain}". Only "@${allowedDomain}" users are permitted.`;
    logStatus(`[SECURITY] Blocked unauthorized access attempt by: ${userEmail}`);
    if (throwOnUnauthorized) throw new Error(msg);
    return { authorized: false, email: userEmail, domain: userDomain, reason: msg };
  }

  return {
    authorized: true,
    email: userEmail,
    domain: userDomain,
    reason: `Authenticated as ${userEmail}`
  };
}

/**
 * Web App HTTP GET Entry Point.
 * Serves the Email Validator web application (embeddable in Google Sites or standalone).
 */
function doGet(e) {
  const auth = verifyTirwinAccess(false);

  const template = HtmlService.createTemplateFromFile('Index');
  template.authJson = JSON.stringify(auth);
  template.allowedDomain = CONFIG.ALLOWED_DOMAIN;

  return template
    .evaluate()
    .setTitle('Tirwin Email List Validator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Web App RPC: Returns current authenticated session info and SMTP configuration status.
 */
function getWebAppSessionInfo() {
  const auth = verifyTirwinAccess(false);
  let smtpProvider = 'NONE';
  if (auth.authorized) {
    const cfg = getSmtpApiConfig();
    smtpProvider = cfg.provider && cfg.apiKey ? cfg.provider : 'NONE';
  }
  return {
    authorized: auth.authorized,
    email: auth.email,
    domain: auth.domain,
    reason: auth.reason,
    allowedDomain: CONFIG.ALLOWED_DOMAIN,
    batchSize: CONFIG.BATCH_SIZE,
    smtpProvider: smtpProvider
  };
}

/**
 * Web App RPC: Lists recent Google Docs, Google Sheets, and Text/CSV files from the user's Drive
 * so the user can pick any document to extract email lists from.
 *
 * @param {string} searchQuery Optional search keyword
 * @return {Array<{id: string, name: string, type: string, updated: string, url: string}>}
 */
function listRecentUserDocsAndSheets(searchQuery) {
  const auth = verifyTirwinAccess(true);
  logStatus(`[WEBAPP] User ${auth.email} listing Google Docs/Sheets (search="${searchQuery || ''}")`);

  const filesList = [];
  const safeQuery = String(searchQuery || '').replace(/'/g, "\\'").trim();

  let q =
    "(mimeType = 'application/vnd.google-apps.document' or " +
    "mimeType = 'application/vnd.google-apps.spreadsheet' or " +
    "mimeType = 'text/plain' or mimeType = 'text/csv') and trashed = false";

  if (safeQuery) {
    q += ` and title contains '${safeQuery}'`;
  }

  const iterator = DriveApp.searchFiles(q);
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';

  while (iterator.hasNext() && filesList.length < 40) {
    const file = iterator.next();
    const mime = file.getMimeType();
    let typeLabel = 'Google Doc';
    if (mime === 'application/vnd.google-apps.spreadsheet') {
      typeLabel = 'Google Sheet';
    } else if (mime === 'text/csv') {
      typeLabel = 'CSV File';
    } else if (mime === 'text/plain') {
      typeLabel = 'Text File';
    }

    filesList.push({
      id: file.getId(),
      name: file.getName(),
      type: typeLabel,
      mimeType: mime,
      updated: Utilities.formatDate(file.getLastUpdated(), timeZone, 'yyyy-MM-dd HH:mm'),
      url: file.getUrl()
    });
  }

  return filesList;
}

/**
 * Web App RPC: Extracts all candidate email entries from a selected Google Doc, Google Sheet,
 * or Text/CSV file (accepted by File ID or full Google Docs/Sheets URL).
 *
 * Designed to preserve candidate strings that have internal whitespace or special character typos
 * (such as "john doe@company.com" or "user#1@domain.com") so the UI initial vetting stage can
 * display and flag them accurately.
 *
 * @param {string} docUrlOrId
 * @return {{fileName: string, fileType: string, rawCandidates: Array<string>, totalExtracted: number}}
 */
function extractEmailsFromGoogleDocOrSheet(docUrlOrId) {
  const auth = verifyTirwinAccess(true);
  const input = String(docUrlOrId || '').trim();
  if (!input) {
    throw new Error('Please select a file or enter a valid Google Doc / Google Sheet URL or File ID.');
  }

  const fileId = extractDriveFileId(input);
  let file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (err) {
    throw new Error(`Unable to open file (ID: ${fileId}). Please check that the file exists and you have permission to view it.`);
  }

  const mimeType = file.getMimeType();
  const fileName = file.getName();
  let rawText = '';
  let fileType = 'Google Doc';

  logStatus(`[WEBAPP] User ${auth.email} extracting email list from "${fileName}" (${mimeType})`);

  if (mimeType === 'application/vnd.google-apps.document') {
    fileType = 'Google Doc';
    const doc = DocumentApp.openById(fileId);
    const bodyText = doc.getBody() ? doc.getBody().getText() : '';
    const headerText = doc.getHeader() ? doc.getHeader().getText() : '';
    const footerText = doc.getFooter() ? doc.getFooter().getText() : '';
    rawText = [headerText, bodyText, footerText].join('\n');
  } else if (mimeType === 'application/vnd.google-apps.spreadsheet') {
    fileType = 'Google Sheet';
    const ss = SpreadsheetApp.openById(fileId);
    const targetSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME) || ss.getSheets()[0];
    const values = targetSheet.getDataRange().getDisplayValues();
    const cellLines = [];
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cellVal = String(values[r][c] || '');
        if (cellVal.trim() !== '') {
          cellLines.push(cellVal);
        }
      }
    }
    rawText = cellLines.join('\n');
  } else {
    fileType = 'Text / CSV File';
    rawText = file.getBlob().getDataAsString('UTF-8');
  }

  const rawCandidates = parseEmailCandidatesFromText(rawText);
  logStatus(`[WEBAPP] Extracted ${rawCandidates.length} email candidate(s) from "${fileName}".`);

  return {
    fileName: fileName,
    fileType: fileType,
    rawCandidates: rawCandidates,
    totalExtracted: rawCandidates.length
  };
}

/**
 * Extracts a Google Drive File ID from a full Google Docs/Sheets/Drive URL or returns the raw ID.
 *
 * @param {string} urlOrId
 * @return {string}
 */
function extractDriveFileId(urlOrId) {
  const trimmed = String(urlOrId || '').trim();
  const dMatch = trimmed.match(/\/d\/([a-zA-Z0-9-_]{15,})/);
  if (dMatch && dMatch[1]) return dMatch[1];

  const idParamMatch = trimmed.match(/[?&]id=([a-zA-Z0-9-_]{15,})/);
  if (idParamMatch && idParamMatch[1]) return idParamMatch[1];

  return trimmed;
}

/**
 * Parses candidate email strings from unstructured or structured Google Doc / Sheet text.
 * Preserves whitespace or invalid special character anomalies on lines/items intended as emails
 * so initial vetting can catch and highlight them.
 *
 * @param {string} rawText
 * @return {Array<string>}
 */
function parseEmailCandidatesFromText(rawText) {
  if (!rawText) return [];

  // Split by newlines, commas, semicolons, or tabs
  const chunks = String(rawText).split(/[\r\n,;|\t]+/);
  const candidates = [];

  for (let i = 0; i < chunks.length; i++) {
    const rawChunk = chunks[i];
    if (!rawChunk || rawChunk.trim() === '') continue;

    const trimmed = rawChunk.trim();

    // Skip common header labels if they don't contain '@'
    if (
      !trimmed.includes('@') &&
      /^(email|email\s*id|email\s*ids|e-mail|email\s*address|sl\.?\s*no|s\.?\s*no|name|status)$/i.test(trimmed)
    ) {
      continue;
    }

    // If the line is a short single entry (e.g. a bullet item or table cell), keep it intact
    // (including any leading/trailing/internal spaces) so basic vetting can flag whitespace/format issues.
    const wordCount = trimmed.split(/\s+/).length;
    if (trimmed.includes('@') && wordCount <= 3 && trimmed.length <= 120) {
      // Strip common bullet markers or angle brackets if present, while preserving whitespace checks
      const cleaned = rawChunk.replace(/^[\s•*\-–—>]+(?=\S)/, '').replace(/^[<("']+|[>)"']+$/g, '');
      candidates.push(cleaned);
    } else if (trimmed.includes('@')) {
      // Long paragraph in a Google Doc: extract all email-like tokens containing '@'
      const matches = trimmed.match(/[^\s<>"',;()[\]]+@[^\s<>"',;()[\]]+/g);
      if (matches) {
        for (let m = 0; m < matches.length; m++) {
          candidates.push(matches[m].replace(/[.,;:!?]+$/, ''));
        }
      }
    } else if (wordCount === 1 && trimmed.length >= 3 && trimmed.includes('.')) {
      // Single token with a dot but missing '@' (e.g. "johndoe.gmail.com") -> include so format check catches missing '@'
      candidates.push(rawChunk);
    }
  }

  return candidates;
}

/**
 * Web App RPC: Validates a batch of up to 100 email addresses for Domain + User availability
 * without sending any email to the users, and logs progress to the Apps Script execution log.
 *
 * Also optionally logs the results to a new timestamped sheet if the script is bound to a spreadsheet.
 *
 * @param {Array<string>} emailBatch Array of pre-vetted email strings (max 100 per call)
 * @param {number} batchNumber 1-based batch index for logging
 * @param {number} totalBatches Total number of batches
 * @return {Array<Object>} Array of detailed validation result objects
 */
function validateEmailBatchWeb(emailBatch, batchNumber, totalBatches) {
  const auth = verifyTirwinAccess(true);

  if (!Array.isArray(emailBatch)) {
    throw new Error('Invalid email batch payload.');
  }

  const batch = emailBatch.slice(0, CONFIG.BATCH_SIZE);
  const bNum = batchNumber || 1;
  const bTotal = totalBatches || 1;

  logStatus('--------------------------------------------------------------------------------');
  logStatus(`[BATCH ${bNum}/${bTotal}] User ${auth.email} starting deep validation for ${batch.length} email(s)...`);

  const domainCache = {};
  const probeFileId = getOrCreateSilentProbeFileId();
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const results = [];

  for (let i = 0; i < batch.length; i++) {
    const email = String(batch[i] || '').trim();
    const startMs = Date.now();
    const checkedAt = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');

    logStatus(`[Batch ${bNum}/${bTotal} - ${i + 1}/${batch.length}] Validating "${email}"...`);
    const validation = validateSingleEmail(email, domainCache, probeFileId);
    const elapsedMs = Date.now() - startMs;

    const record = {
      email: email,
      status: validation.status,
      isActive: Boolean(validation.isActive),
      domainValid: Boolean(validation.domainValid),
      userActive: validation.userActive,
      method: validation.method,
      syntaxValid: Boolean(validation.syntaxValid),
      isDisposable: Boolean(validation.isDisposable),
      primaryMx: validation.primaryMx || 'N/A',
      mxRecords: validation.mxRecords || [],
      reason: validation.reason,
      checkedAt: checkedAt,
      elapsedMs: elapsedMs
    };

    results.push(record);

    logStatus(
      `[Batch ${bNum}/${bTotal} - ${i + 1}/${batch.length}] COMPLETED: "${email}" | ` +
      `Status: ${record.status} | Active: ${record.isActive ? 'Yes' : 'No'} | ` +
      `Domain Valid: ${record.domainValid ? 'Yes' : 'No'} | User Active: ${record.userActive} | ` +
      `Method: ${record.method} | MX: ${record.primaryMx} (${elapsedMs} ms)`
    );
  }

  logStatus(`[BATCH ${bNum}/${bTotal}] Completed ${results.length} email(s).`);
  return results;
}

/**
 * Web App RPC: Saves the completed validation run to a newly created date-and-timestamped
 * sheet tab if bound to a Google Sheet (or creates/updates a master log spreadsheet in Drive).
 *
 * @param {Array<Object>} allResults
 * @return {{loggedToSheet: boolean, sheetName: string, spreadsheetUrl: string}}
 */
function saveValidationRunToTimestampedSheet(allResults) {
  const auth = verifyTirwinAccess(true);
  if (!Array.isArray(allResults) || allResults.length === 0) {
    return { loggedToSheet: false, sheetName: '', spreadsheetUrl: '' };
  }

  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      const props = PropertiesService.getScriptProperties();
      const savedSsId = props.getProperty('VALIDATOR_LOG_SPREADSHEET_ID');
      if (savedSsId) {
        try {
          ss = SpreadsheetApp.openById(savedSsId);
        } catch (e) {
          ss = null;
        }
      }
      if (!ss) {
        ss = SpreadsheetApp.create('Tirwin Email Validation Logs');
        props.setProperty('VALIDATOR_LOG_SPREADSHEET_ID', ss.getId());
      }
    }

    const timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Kolkata';
    const executionTime = new Date();
    const formattedSheetTimestamp = Utilities.formatDate(executionTime, timeZone, 'yyyy-MM-dd_HH-mm-ss');

    let resultSheetName = `Validated_${formattedSheetTimestamp}`;
    if (ss.getSheetByName(resultSheetName)) {
      resultSheetName += `_${executionTime.getMilliseconds()}`;
    }
    const resultSheet = ss.insertSheet(resultSheetName);

    const headers = [
      'Email ID',
      'Overall Status',
      'Is Active',
      'Basic Vetting Passed',
      'Domain Valid',
      'User Active in Domain',
      'Verification Method',
      'Disposable Domain',
      'Primary Mail Server (MX)',
      'Validation Details',
      'Validated By',
      'Validated At'
    ];
    const outputRows = [headers];

    for (let i = 0; i < allResults.length; i++) {
      const r = allResults[i];
      outputRows.push([
        r.email || r.rawEmail || '',
        r.status || '',
        r.isActive ? 'Yes' : 'No',
        r.basicValid ? 'Yes' : 'No',
        r.domainValid ? 'Yes' : 'No',
        r.userActive || 'No',
        r.method || '',
        r.isDisposable ? 'Yes' : 'No',
        r.primaryMx || 'N/A',
        r.reason || '',
        auth.email,
        r.checkedAt || Utilities.formatDate(executionTime, timeZone, 'yyyy-MM-dd HH:mm:ss')
      ]);
    }

    resultSheet.getRange(1, 1, outputRows.length, headers.length).setValues(outputRows);
    formatResultSheet(resultSheet, outputRows.length, headers.length);

    logStatus(`[SHEET LOG] Saved ${allResults.length} validated email(s) to new sheet "${resultSheetName}".`);
    return {
      loggedToSheet: true,
      sheetName: resultSheetName,
      spreadsheetUrl: ss.getUrl()
    };
  } catch (err) {
    logStatus(`[WARN] Could not log run to Google Sheet: ${err.message}`);
    return { loggedToSheet: false, sheetName: '', spreadsheetUrl: '' };
  }
}

/**
 * Web App RPC: Saves an optional external SMTP API provider & key from the Web UI.
 *
 * @param {string} provider
 * @param {string} apiKey
 */
function saveSmtpApiConfigWeb(provider, apiKey) {
  const auth = verifyTirwinAccess(true);
  const props = PropertiesService.getScriptProperties();
  const cleanProvider = String(provider || 'NONE').trim().toUpperCase();
  const cleanKey = String(apiKey || '').trim();

  if (cleanProvider === 'NONE' || !cleanKey) {
    props.deleteProperty('VERIFICATION_PROVIDER');
    props.deleteProperty('API_KEY');
    logStatus(`[CONFIG] User ${auth.email} cleared external SMTP API configuration.`);
    return { provider: 'NONE' };
  }

  props.setProperty('VERIFICATION_PROVIDER', cleanProvider);
  props.setProperty('API_KEY', cleanKey);
  logStatus(`[CONFIG] User ${auth.email} configured SMTP API provider: ${cleanProvider}`);
  return { provider: cleanProvider };
}

/**
 * Returns a valid Google Drive file ID owned/editable by the user so the Drive v3 REST API
 * can test permissions with `?sendNotificationEmail=false` (never uses DriveApp.addViewer).
 *
 * @return {string|null}
 */
function getOrCreateSilentProbeFileId() {
  try {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) return activeSs.getId();
  } catch (e) {
    // Running as standalone web app
  }

  try {
    const props = PropertiesService.getUserProperties();
    const cachedId = props.getProperty('SILENT_REST_PROBE_FILE_ID');
    if (cachedId) {
      try {
        const f = DriveApp.getFileById(cachedId);
        if (!f.isTrashed()) return cachedId;
      } catch (e) {
        // File was deleted; recreate below
      }
    }

    const newFile = DriveApp.createFile('_tirwin_email_validator_anchor.txt', 'Internal anchor for silent REST lookup.');
    props.setProperty('SILENT_REST_PROBE_FILE_ID', newFile.getId());
    return newFile.getId();
  } catch (err) {
    logStatus(`[WARN] Unable to obtain probe file ID for Drive REST lookup: ${err.message}`);
    return null;
  }
}

/**
 * Adds a custom menu to the Google Spreadsheet when opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Email Validator')
    .addItem('Validate Emails & Users Now', 'validateEmailList')
    .addSeparator()
    .addItem('Configure SMTP API Key (For Custom Domains)', 'configureSmtpApiPrompt')
    .addToUi();
}

/**
 * Interactive prompt allowing the user to save an SMTP Verifier API key in ScriptProperties.
 */
function configureSmtpApiPrompt() {
  verifyTirwinAccess(true);
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const providerResp = ui.prompt(
    'Configure SMTP Mailbox Verifier (Step 1/2)',
    'Enter provider name (HUNTER, ABSTRACT, ZEROBOUNCE, REOON, or NONE to clear):',
    ui.ButtonSet.OK_CANCEL
  );
  if (providerResp.getSelectedButton() !== ui.Button.OK) return;

  const provider = providerResp.getResponseText().trim().toUpperCase();
  if (!provider || provider === 'NONE') {
    props.deleteProperty('VERIFICATION_PROVIDER');
    props.deleteProperty('API_KEY');
    ui.alert('Cleared external SMTP API configuration. Using built-in zero-email user checks.');
    return;
  }

  const keyResp = ui.prompt(
    'Configure SMTP Mailbox Verifier (Step 2/2)',
    `Enter your API Key for ${provider}:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (keyResp.getSelectedButton() !== ui.Button.OK) return;

  props.setProperty('VERIFICATION_PROVIDER', provider);
  props.setProperty('API_KEY', keyResp.getResponseText().trim());
  ui.alert(`Saved ${provider} API key for deep SMTP user verification.`);
}

/**
 * Standalone Sheet Function: Reads emails from "emailids", strictly validates whether the user is
 * valid and active in the domain without sending any email, logs execution progress for
 * every email ID, and writes results to a new date-and-timestamped sheet.
 */
function validateEmailList() {
  const startTimeMs = Date.now();
  logStatus('================================================================================');
  logStatus('[START] Email List Validation Execution Started');

  const auth = verifyTirwinAccess(true);
  logStatus(`[AUTH] Verified user: ${auth.email}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    const errMsg = `Source sheet "${CONFIG.SOURCE_SHEET_NAME}" not found. Please create a sheet named "${CONFIG.SOURCE_SHEET_NAME}".`;
    logStatus(`[ERROR] ${errMsg}`);
    throw new Error(errMsg);
  }

  logStatus(`[INFO] Reading email IDs from sheet "${CONFIG.SOURCE_SHEET_NAME}" (Column ${CONFIG.EMAIL_COLUMN_INDEX})...`);

  const lastRow = sourceSheet.getLastRow();
  if (lastRow === 0) {
    logStatus(`[WARNING] The sheet "${CONFIG.SOURCE_SHEET_NAME}" is empty. Aborting execution.`);
    SpreadsheetApp.getUi().alert(`The sheet "${CONFIG.SOURCE_SHEET_NAME}" is empty.`);
    return;
  }

  const rawValues = sourceSheet
    .getRange(1, CONFIG.EMAIL_COLUMN_INDEX, lastRow, 1)
    .getValues()
    .map(row => String(row[0]).trim());

  let startIndex = 0;
  if (CONFIG.HAS_HEADER_ROW && rawValues.length > 0 && !rawValues[0].includes('@')) {
    logStatus(`[INFO] Skipping header row: "${rawValues[0]}"`);
    startIndex = 1;
  }

  const emails = rawValues.slice(startIndex).filter(val => val !== '');
  if (emails.length === 0) {
    logStatus(`[WARNING] No email addresses found in "${CONFIG.SOURCE_SHEET_NAME}". Aborting execution.`);
    SpreadsheetApp.getUi().alert(`No email addresses found in "${CONFIG.SOURCE_SHEET_NAME}".`);
    return;
  }

  // Sort the email list alphabetically
  emails.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
  logStatus(`[INFO] Found and sorted ${emails.length} email ID(s) to validate.`);

  const domainCache = {};
  const timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'UTC';
  const executionTime = new Date();
  const formattedSheetTimestamp = Utilities.formatDate(executionTime, timeZone, 'yyyy-MM-dd_HH-mm-ss');

  let resultSheetName = `Validated_${formattedSheetTimestamp}`;
  if (ss.getSheetByName(resultSheetName)) {
    resultSheetName += `_${executionTime.getMilliseconds()}`;
  }
  const resultSheet = ss.insertSheet(resultSheetName);
  logStatus(`[INFO] Created new output sheet: "${resultSheetName}"`);
  logStatus('--------------------------------------------------------------------------------');

  const probeFileId = ss.getId();

  const headers = [
    'Email ID',
    'Overall Status',
    'Is Active',
    'Domain Valid',
    'User Active in Domain',
    'Verification Method',
    'Syntax Valid',
    'Disposable Domain',
    'Primary Mail Server (MX)',
    'Validation Details',
    'Validated At'
  ];
  const outputRows = [headers];

  let activeCount = 0;
  let inactiveCount = 0;
  let unverifiedCount = 0;

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    const progressPrefix = `[${i + 1}/${emails.length}]`;
    const emailStartMs = Date.now();

    logStatus(`${progressPrefix} Validating email: "${email}" ...`);

    const checkedAt = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
    const validation = validateSingleEmail(email, domainCache, probeFileId);
    const elapsedMs = Date.now() - emailStartMs;

    if (validation.isActive) {
      activeCount++;
    } else if (String(validation.userActive).includes('Unverified')) {
      unverifiedCount++;
    } else {
      inactiveCount++;
    }

    outputRows.push([
      email,
      validation.status,
      validation.isActive ? 'Yes' : 'No',
      validation.domainValid ? 'Yes' : 'No',
      validation.userActive,
      validation.method,
      validation.syntaxValid ? 'Yes' : 'No',
      validation.isDisposable ? 'Yes' : 'No',
      validation.primaryMx || 'N/A',
      validation.reason,
      checkedAt
    ]);

    logStatus(
      `${progressPrefix} COMPLETED: "${email}" | Status: ${validation.status} | ` +
      `Is Active: ${validation.isActive ? 'Yes' : 'No'} | Domain Valid: ${validation.domainValid ? 'Yes' : 'No'} | ` +
      `User Active: ${validation.userActive} | Method: ${validation.method} | ` +
      `MX: ${validation.primaryMx || 'N/A'} | Details: ${validation.reason} (${elapsedMs} ms)`
    );
  }

  logStatus('--------------------------------------------------------------------------------');
  logStatus(`[INFO] Writing ${outputRows.length - 1} validated record(s) to sheet "${resultSheetName}"...`);

  resultSheet.getRange(1, 1, outputRows.length, headers.length).setValues(outputRows);
  formatResultSheet(resultSheet, outputRows.length, headers.length);

  const totalDurationSec = ((Date.now() - startTimeMs) / 1000).toFixed(2);
  logStatus(
    `[SUMMARY] Total Processed: ${emails.length} | Active: ${activeCount} | ` +
    `Inactive/Invalid: ${inactiveCount} | Unverified: ${unverifiedCount}`
  );
  logStatus(`[DONE] Execution completed in ${totalDurationSec}s. Results saved to "${resultSheetName}".`);
  logStatus('================================================================================');
}

/**
 * Validates syntax, domain MX records, and strictly checks whether the specific user
 * is valid and active in that domain without sending any email.
 *
 * @param {string} email
 * @param {Object} domainCache
 * @param {string} probeFileId
 * @return {Object} Validation result details
 */
function validateSingleEmail(email, domainCache, probeFileId) {
  const result = {
    status: 'Inactive / Invalid',
    isActive: false,
    domainValid: false,
    userActive: 'No',
    method: 'Syntax Check',
    syntaxValid: false,
    isDisposable: false,
    primaryMx: '',
    mxRecords: [],
    reason: ''
  };

  // 1. Basic & RFC 5322 Syntax Validation (whitespace, special chars, local-part, domain format)
  if (/\s/.test(email)) {
    result.reason = 'Contains invalid whitespace characters';
    return result;
  }

  const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/;
  if (!emailRegex.test(email)) {
    result.reason = 'Invalid email syntax or disallowed special characters';
    return result;
  }

  const parts = email.split('@');
  const localPart = parts[0];
  const domain = parts[1].toLowerCase();

  if (localPart.startsWith('.') || localPart.endsWith('.') || localPart.includes('..')) {
    result.reason = 'Invalid local part (consecutive or leading/trailing dots)';
    return result;
  }

  result.syntaxValid = true;

  // 2. Check for known disposable / temporary email domains
  if (DISPOSABLE_DOMAINS.has(domain)) {
    result.isDisposable = true;
    result.domainValid = true;
    result.status = 'Inactive (Disposable Domain)';
    result.userActive = 'No (Disposable)';
    result.method = 'Disposable Filter';
    result.reason = 'Domain is valid, but belongs to a temporary/disposable email provider';
    return result;
  }

  // 3. Domain & Mail Server (MX Record) Validation via Google Public DNS
  if (!domainCache[domain]) {
    logStatus(`   -> [DNS] Querying MX records for domain "${domain}"...`);
    domainCache[domain] = checkDomainDns(domain);
  }
  const dnsInfo = domainCache[domain];

  result.primaryMx = dnsInfo.primaryMx;
  result.mxRecords = dnsInfo.mxList || [];
  result.method = 'DNS MX Lookup';

  if (!dnsInfo.domainExists) {
    result.domainValid = false;
    result.reason = 'Domain does not exist (NXDOMAIN)';
    return result;
  }

  if (!dnsInfo.hasMx) {
    result.domainValid = false;
    result.reason = 'Domain exists, but has no active mail servers (MX records) configured';
    return result;
  }

  // The domain itself is valid and accepts mail
  result.domainValid = true;

  // 4. Strictly validate whether the USER is valid and active inside this valid domain
  const userCheck = verifyUserActiveInValidDomain(email, domain, dnsInfo, domainCache, probeFileId);

  result.method = userCheck.method;
  result.userActive = userCheck.userActive;
  result.isActive = CONFIG.REQUIRE_USER_VERIFICATION_FOR_ACTIVE
    ? userCheck.userActive === 'Yes'
    : Boolean(userCheck.isActive);
  result.status = userCheck.status;
  result.reason = userCheck.reason;

  return result;
}

/**
 * Strictly verifies whether the specific user exists and is active inside a valid domain
 * without sending any email.
 */
function verifyUserActiveInValidDomain(email, domain, dnsInfo, domainCache, probeFileId) {
  // 1. Check if an external SMTP verification API is configured (in CONFIG or ScriptProperties)
  const smtpConfig = getSmtpApiConfig();
  if (smtpConfig.provider !== 'NONE' && smtpConfig.apiKey) {
    const apiCheck = checkMailboxWithSmtpApi(email, smtpConfig.provider, smtpConfig.apiKey);
    if (apiCheck) {
      return apiCheck;
    }
  }

  // 2. Check Google Workspace Directory / People API (if running inside a Workspace org & enabled)
  const directoryCheck = checkWorkspaceDirectoryUser(email, domain);
  if (directoryCheck) {
    return directoryCheck;
  }

  const mxLower = (dnsInfo.primaryMx || '').toLowerCase();
  const isMicrosoftDomain =
    domain === 'outlook.com' ||
    domain === 'hotmail.com' ||
    domain === 'live.com' ||
    domain === 'msn.com' ||
    mxLower.includes('outlook.com') ||
    mxLower.includes('protection.outlook.com');

  const isGoogleDomain =
    domain === 'gmail.com' ||
    domain === 'googlemail.com' ||
    mxLower.includes('google.com') ||
    mxLower.includes('googlemail.com') ||
    mxLower.includes('aspmx.l.google.com');

  // 3. Microsoft 365 / Exchange Online / Outlook.com / Hotmail user existence + canary check
  if (CONFIG.ENABLE_MICROSOFT_USER_CHECK && isMicrosoftDomain) {
    const msCheck = checkMicrosoftUserExistsStrict(email, domain, domainCache);
    if (msCheck) {
      return msCheck;
    }
  }

  // 4. Google / Gmail / Google Workspace user existence check via Drive REST API (sendNotificationEmail=false)
  if (CONFIG.ENABLE_GOOGLE_USER_CHECK && probeFileId) {
    const googleCheck = checkGoogleUserViaSilentRestApi(email, domain, isGoogleDomain, probeFileId);
    if (googleCheck) {
      return googleCheck;
    }
  }

  // 5. Fallback if domain is valid (e.g. Yahoo, Zoho, custom cPanel SMTP), but no SMTP API key is configured
  return {
    isActive: false,
    userActive: 'Unverified (Needs SMTP API)',
    status: 'Domain Valid / User Unverified',
    method: 'DNS MX Only (No SMTP Probe)',
    reason: `Domain "${domain}" is valid (MX: ${dnsInfo.primaryMx}), but verifying user "${email}" on this mail server requires an SMTP Verifier API key`
  };
}

/**
 * Checks if a user exists and is active in Microsoft 365 / Azure AD / Outlook / Hotmail
 * without sending any email. Also performs a per-domain canary test with a guaranteed
 * non-existent user to make sure federated tenants don't return false positives.
 */
function checkMicrosoftUserExistsStrict(email, domain, domainCache) {
  try {
    const realmUrl = `https://login.microsoftonline.com/getuserrealm.srf?login=${encodeURIComponent(email)}&json=1`;
    const realmRes = UrlFetchApp.fetch(realmUrl, { muteHttpExceptions: true });
    if (realmRes.getResponseCode() === 200) {
      const realmData = JSON.parse(realmRes.getContentText());
      if (realmData.NameSpaceType === 'Unknown') {
        return null; // Not a managed Microsoft tenant
      }
    }

    const userCode = queryMicrosoftCredentialType(email);
    if (userCode === null) return null;

    if (userCode === 1) {
      return {
        isActive: false,
        userActive: 'No',
        status: 'Inactive (User Not Found in Domain)',
        method: 'Microsoft 365 User Lookup',
        reason: `Domain "${domain}" is valid, but user "${email}" does not exist or is inactive`
      };
    }

    if (userCode === 0 || userCode === 5 || userCode === 6) {
      if (domain !== 'outlook.com' && domain !== 'hotmail.com' && domain !== 'live.com' && domain !== 'msn.com') {
        if (typeof domainCache[domain].msCanaryAcceptsAll === 'undefined') {
          const fakeUser = `nonexistent_canary_${new Date().getTime()}_9981@${domain}`;
          const canaryCode = queryMicrosoftCredentialType(fakeUser);
          domainCache[domain].msCanaryAcceptsAll = (canaryCode === 0 || canaryCode === 5 || canaryCode === 6);
        }

        if (domainCache[domain].msCanaryAcceptsAll) {
          return {
            isActive: false,
            userActive: 'Unverified (Catch-All Federated)',
            status: 'Domain Valid / Catch-All Tenant',
            method: 'Microsoft 365 Canary Check',
            reason: `Domain "${domain}" is a federated catch-all tenant that accepts any username during lookup`
          };
        }
      }

      return {
        isActive: true,
        userActive: 'Yes',
        status: 'Active',
        method: 'Microsoft 365 User Lookup',
        reason: `Domain "${domain}" is valid and user mailbox "${email}" is verified as active`
      };
    }
  } catch (err) {
    logStatus(`   -> [WARN] Microsoft user check error for ${email}: ${err.message}`);
  }
  return null;
}

/**
 * Helper to query Microsoft's silent GetCredentialType endpoint.
 */
function queryMicrosoftCredentialType(email) {
  const url = 'https://login.microsoftonline.com/common/GetCredentialType';
  const payload = {
    Username: email,
    isOtherIdpSupported: true,
    checkPhones: false,
    isRemoteNGCSupported: true,
    isCookieBannerShown: false,
    isFidoSupported: true,
    originalRequest: '',
    country: 'US',
    forceotclogin: false,
    isExternalFederationDisallowed: false,
    isRemoteConnectSupported: false,
    federationFlags: 0,
    isSignup: false,
    isAccessPassSupported: true
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    if (typeof data.IfExistsResult === 'number') {
      return data.IfExistsResult;
    }
  }
  return null;
}

/**
 * Strictly checks if a user account exists and is active on Gmail / Google Workspace
 * using the Google Drive v3 REST API with `?sendNotificationEmail=false`.
 */
function checkGoogleUserViaSilentRestApi(email, domain, isGoogleDomain, probeFileId) {
  const currentUserEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (currentUserEmail && email.toLowerCase() === currentUserEmail) {
    return {
      isActive: true,
      userActive: 'Yes',
      status: 'Active',
      method: 'Active Session User',
      reason: `User "${email}" is the currently authenticated active Google account`
    };
  }

  const token = ScriptApp.getOAuthToken();
  const createUrl =
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(probeFileId)}/permissions` +
    `?sendNotificationEmail=false&supportsAllDrives=true&fields=id,displayName,photoLink,deleted`;

  const response = UrlFetchApp.fetch(createUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      role: 'reader',
      type: 'user',
      emailAddress: email
    }),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode === 200 || statusCode === 201) {
    let perm = {};
    try {
      perm = JSON.parse(responseText);
    } catch (e) {
      // Ignore parse error
    }

    // Immediately remove the temporary reader permission silently
    if (perm && perm.id) {
      const deleteUrl =
        `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(probeFileId)}/permissions/${encodeURIComponent(perm.id)}` +
        `?supportsAllDrives=true`;
      UrlFetchApp.fetch(deleteUrl, {
        method: 'delete',
        headers: {
          Authorization: `Bearer ${token}`
        },
        muteHttpExceptions: true
      });
    }

    if (perm && perm.deleted === true) {
      return {
        isActive: false,
        userActive: 'No (Deleted Account)',
        status: 'Inactive (Deleted Google Account)',
        method: 'Google Identity REST Lookup',
        reason: `Domain "${domain}" is valid, but Google user "${email}" is deleted`
      };
    }

    return {
      isActive: true,
      userActive: 'Yes',
      status: 'Active',
      method: 'Google Identity REST Lookup',
      reason: `Domain "${domain}" is valid and user "${email}" is verified as an active Google/Workspace account (zero email sent)`
    };
  }

  // If HTTP 400 / 404 is returned, the email address is NOT a valid/active Google Account
  if (isGoogleDomain) {
    return {
      isActive: false,
      userActive: 'No',
      status: 'Inactive (User Not Found in Domain)',
      method: 'Google Identity REST Lookup',
      reason: `Domain "${domain}" is a valid Google-hosted domain, but user "${email}" does not exist or is inactive`
    };
  }

  return null;
}

/**
 * Checks if the user exists and is active in your own Google Workspace organization
 * via AdminDirectory API (if enabled in Services).
 */
function checkWorkspaceDirectoryUser(email, domain) {
  if (typeof AdminDirectory !== 'undefined' && AdminDirectory.Users) {
    try {
      const user = AdminDirectory.Users.get(email);
      if (user) {
        const isSuspended = Boolean(user.suspended) || Boolean(user.archived);
        return {
          isActive: !isSuspended,
          userActive: isSuspended ? 'No (Suspended/Archived)' : 'Yes',
          status: isSuspended ? 'Inactive (Suspended User)' : 'Active',
          method: 'Google Workspace Admin Directory',
          reason: isSuspended
            ? `Domain "${domain}" is valid, but user "${email}" is suspended or archived`
            : `Domain "${domain}" is valid and user "${email}" is active in Workspace Directory`
        };
      }
    } catch (err) {
      const msg = String(err.message || err).toLowerCase();
      if (msg.includes('not found')) {
        return {
          isActive: false,
          userActive: 'No',
          status: 'Inactive (User Not Found in Domain)',
          method: 'Google Workspace Admin Directory',
          reason: `Domain "${domain}" is valid, but user "${email}" does not exist in Workspace Directory`
        };
      }
    }
  }
  return null;
}

/**
 * Queries Google Public DNS-over-HTTPS (https://dns.google/resolve) to verify
 * whether the domain exists and has active MX records.
 */
function checkDomainDns(domain) {
  const dnsResult = {
    domainExists: false,
    hasMx: false,
    mxCount: 0,
    primaryMx: '',
    mxList: []
  };

  try {
    const mxUrl = `https://dns.google/resolve?name=${encodeURIComponent(domain)}&type=MX`;
    const mxResponse = UrlFetchApp.fetch(mxUrl, { muteHttpExceptions: true });
    const mxData = JSON.parse(mxResponse.getContentText());

    if (mxData.Status === 3) {
      return dnsResult;
    }

    dnsResult.domainExists = true;

    if (mxData.Status === 0 && Array.isArray(mxData.Answer)) {
      const mxRecords = mxData.Answer
        .filter(record => record.type === 15 && record.data && record.data.trim() !== '0 .')
        .map(record => {
          const segments = record.data.trim().split(/\s+/);
          return {
            priority: parseInt(segments[0], 10) || 0,
            exchange: (segments[1] || record.data).replace(/\.$/, '')
          };
        })
        .sort((a, b) => a.priority - b.priority);

      if (mxRecords.length > 0) {
        dnsResult.hasMx = true;
        dnsResult.mxCount = mxRecords.length;
        dnsResult.primaryMx = mxRecords[0].exchange;
        dnsResult.mxList = mxRecords.map(r => `${r.priority} ${r.exchange}`);
      }
    }
  } catch (err) {
    logStatus(`   -> [WARN] DNS lookup error for ${domain}: ${err.message}`);
  }

  return dnsResult;
}

/**
 * Retrieves SMTP Verifier API configuration from CONFIG or ScriptProperties.
 */
function getSmtpApiConfig() {
  let provider = (CONFIG.VERIFICATION_PROVIDER || 'NONE').toUpperCase();
  let apiKey = CONFIG.API_KEY || '';

  if (provider === 'NONE' || !apiKey) {
    try {
      const props = PropertiesService.getScriptProperties();
      const savedProvider = props.getProperty('VERIFICATION_PROVIDER');
      const savedKey = props.getProperty('API_KEY');
      if (savedProvider && savedKey) {
        provider = savedProvider.toUpperCase();
        apiKey = savedKey;
      }
    } catch (e) {
      // Ignore properties error if unauthorized
    }
  }
  return { provider, apiKey };
}

/**
 * Mailbox-level SMTP handshake (`RCPT TO`) verification via external API.
 */
function checkMailboxWithSmtpApi(email, provider, apiKey) {
  try {
    if (provider === 'HUNTER') {
      const url = `https://api.hunter.io/v2/email-verifier?email=${encodeURIComponent(email)}&api_key=${encodeURIComponent(apiKey)}`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const data = JSON.parse(res.getContentText()).data;
        const isCatchAll = data.accept_all === true || data.status === 'accept_all';
        const userStrictlyValid = data.smtp_check === true && data.status === 'valid' && !isCatchAll;
        return {
          isActive: userStrictlyValid,
          userActive: isCatchAll ? 'Unverified (Catch-All Domain)' : (userStrictlyValid ? 'Yes' : 'No'),
          status: userStrictlyValid
            ? 'Active'
            : (isCatchAll ? 'Domain Valid / Catch-All Unverified' : 'Inactive (User Not Found in Domain)'),
          method: 'SMTP Handshake (Hunter API)',
          reason: `Domain valid; SMTP RCPT TO User Status: ${data.status} (SMTP Check: ${data.smtp_check})`
        };
      }
    } else if (provider === 'ABSTRACT') {
      const url = `https://emailvalidation.abstractapi.com/v1/?api_key=${encodeURIComponent(apiKey)}&email=${encodeURIComponent(email)}`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const data = JSON.parse(res.getContentText());
        const isSmtpValid = data.is_smtp_valid && data.is_smtp_valid.value === true;
        const isCatchAll = data.is_catchall_email && data.is_catchall_email.value === true;
        const deliverable = data.deliverability === 'DELIVERABLE' && isSmtpValid && !isCatchAll;
        return {
          isActive: deliverable,
          userActive: isCatchAll ? 'Unverified (Catch-All Domain)' : (deliverable ? 'Yes' : 'No'),
          status: deliverable
            ? 'Active'
            : (isCatchAll ? 'Domain Valid / Catch-All Unverified' : 'Inactive (User Not Found in Domain)'),
          method: 'SMTP Handshake (AbstractAPI)',
          reason: `Domain valid; User Deliverability: ${data.deliverability} (SMTP Valid: ${isSmtpValid})`
        };
      }
    } else if (provider === 'ZEROBOUNCE') {
      const url = `https://api.zerobounce.net/v2/validate?api_key=${encodeURIComponent(apiKey)}&email=${encodeURIComponent(email)}&ip_address=`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const data = JSON.parse(res.getContentText());
        const valid = data.status === 'valid';
        const catchAll = data.status === 'catch-all';
        return {
          isActive: valid,
          userActive: catchAll ? 'Unverified (Catch-All Domain)' : (valid ? 'Yes' : 'No'),
          status: valid
            ? 'Active'
            : (catchAll ? 'Domain Valid / Catch-All Unverified' : 'Inactive (User Not Found in Domain)'),
          method: 'SMTP Handshake (ZeroBounce)',
          reason: `Domain valid; ZeroBounce User Status: ${data.status} (${data.sub_status || 'checked'})`
        };
      }
    } else if (provider === 'REOON') {
      const url = `https://emailverifier.reoon.com/api/v1/verify?email=${encodeURIComponent(email)}&key=${encodeURIComponent(apiKey)}&mode=power`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const data = JSON.parse(res.getContentText());
        const catchAll = data.is_catch_all === true || data.status === 'catch_all';
        const strictlyValid = data.is_deliverable === true && !catchAll && (data.status === 'safe' || data.status === 'valid');
        return {
          isActive: strictlyValid,
          userActive: catchAll ? 'Unverified (Catch-All Domain)' : (strictlyValid ? 'Yes' : 'No'),
          status: strictlyValid
            ? 'Active'
            : (catchAll ? 'Domain Valid / Catch-All Unverified' : 'Inactive (User Not Found in Domain)'),
          method: 'SMTP Handshake (Reoon API)',
          reason: `Domain valid; User Mailbox Status: ${data.status} (Deliverable: ${data.is_deliverable})`
        };
      }
    }
  } catch (err) {
    logStatus(`   -> [WARN] External API verification error for ${email}: ${err.message}`);
  }
  return null;
}

/**
 * Formats the newly created result sheet with header styling and conditional colors.
 */
function formatResultSheet(sheet, totalRows, totalCols) {
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');

  sheet.setFrozenRows(1);

  if (totalRows > 1) {
    const statusRange = sheet.getRange(2, 2, totalRows - 1, 5);

    const activeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active')
      .setBackground('#e6f4ea')
      .setFontColor('#137333')
      .setRanges([statusRange])
      .build();

    const yesRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Yes')
      .setBackground('#e6f4ea')
      .setFontColor('#137333')
      .setRanges([statusRange])
      .build();

    const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Inactive')
      .setBackground('#fce8e6')
      .setFontColor('#c5221f')
      .setRanges([statusRange])
      .build();

    const noRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('No')
      .setBackground('#fce8e6')
      .setFontColor('#c5221f')
      .setRanges([statusRange])
      .build();

    const unverifiedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Unverified')
      .setBackground('#fef7e0')
      .setFontColor('#b06000')
      .setRanges([statusRange])
      .build();

    sheet.setConditionalFormatRules([activeRule, yesRule, inactiveRule, noRule, unverifiedRule]);
  }

  sheet.autoResizeColumns(1, totalCols);
}

/**
 * Common disposable / temporary email domains to flag inactive/throwaway accounts.
 */
const DISPOSABLE_DOMAINS = new Set([
  '10minutemail.com',
  'guerrillamail.com',
  'mailinator.com',
  'tempmail.com',
  'temp-mail.org',
  'trashmail.com',
  'yopmail.com',
  'sharklasers.com',
  'getnada.com',
  'maildrop.cc',
  'dispostable.com',
  'fakeinbox.com',
  'throwawaymail.com',
  'mintemail.com',
  'mohmal.com'
]);
