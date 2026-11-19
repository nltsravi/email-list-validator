/**
 * Google Apps Script: Strict User-Level Email Validator (100% Zero-Notification Guarantee)
 *
 * Reads email addresses from a sheet named "emailids", validates both the domain
 * AND whether the specific user/mailbox is valid and active (even when the domain
 * itself is valid!), WITHOUT sending any email or Drive notification to the user.
 * Logs the results in a newly created sheet with the date and timestamp on every execution,
 * and prints real-time execution status to the Apps Script Execution Log (`Logger.log` & `console.log`)
 * as each email ID validation completes.
 */

const CONFIG = {
  SOURCE_SHEET_NAME: 'emailids',
  EMAIL_COLUMN_INDEX: 1, // 1 = Column A
  HAS_HEADER_ROW: true,  // Automatically skips row 1 if it is a header

  // Strict mode: When true (recommended), an email is ONLY marked "Active" if the
  // specific user mailbox is confirmed to exist (not just the domain!).
  REQUIRE_USER_VERIFICATION_FOR_ACTIVE: true,

  // Enable built-in silent user checks (No API key needed, sends zero emails)
  ENABLE_MICROSOFT_USER_CHECK: true, // Checks Outlook/Hotmail/Live/Office365 custom domains
  ENABLE_GOOGLE_USER_CHECK: true,    // Checks Gmail & Google Workspace domains via REST API with sendNotificationEmail=false

  // Optional SMTP Handshake API for custom non-Google/non-Microsoft SMTP servers (Yahoo, Zoho, cPanel, custom MX):
  // Supported values: 'NONE' | 'HUNTER' | 'ABSTRACT' | 'ZEROBOUNCE' | 'REOON'
  // (You can also configure the API key from the sheet menu: Email Validator -> Configure SMTP API Key)
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
 * Keeping a commented reference to DriveApp.getRootFolder() ensures Apps Script grants the
 * Drive OAuth token used by ScriptApp.getOAuthToken() for the REST call with sendNotificationEmail=false.
 *//** @private */
function _ensureDriveScope() {
  // DriveApp.getRootFolder();
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
 * Interactive prompt allowing the user to save an SMTP Verifier API key in ScriptProperties
 * so custom non-Google/non-Microsoft domains can also verify individual mailboxes via SMTP RCPT TO.
 */
function configureSmtpApiPrompt() {
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
 * Main function to read emails from "emailids", strictly validate whether the user is
 * valid and active in the domain without sending any email, log execution progress for
 * every email ID, and write results to a new date-and-timestamped sheet.
 */
function validateEmailList() {
  const startTimeMs = Date.now();
  logStatus('================================================================================');
  logStatus('[START] Email List Validation Execution Started');

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

  // Read column data
  const rawValues = sourceSheet
    .getRange(1, CONFIG.EMAIL_COLUMN_INDEX, lastRow, 1)
    .getValues()
    .map(row => String(row[0]).trim());

  // Automatically skip header row if configured and first cell doesn't look like an email
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

  logStatus(`[INFO] Found ${emails.length} email ID(s) to validate.`);

  const domainCache = {};
  const timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'UTC';
  const executionTime = new Date();
  const formattedSheetTimestamp = Utilities.formatDate(executionTime, timeZone, 'yyyy-MM-dd_HH-mm-ss');

  // Create a new sheet with Date and Timestamp every time the script executes
  let resultSheetName = `Validated_${formattedSheetTimestamp}`;
  if (ss.getSheetByName(resultSheetName)) {
    resultSheetName += `_${executionTime.getMilliseconds()}`;
  }
  const resultSheet = ss.insertSheet(resultSheetName);
  logStatus(`[INFO] Created new output sheet: "${resultSheetName}"`);
  logStatus('--------------------------------------------------------------------------------');

  // Use the active spreadsheet's own file ID for REST API permission probes with sendNotificationEmail=false
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

    // Log completion details for this email ID
    logStatus(
      `${progressPrefix} COMPLETED: "${email}" | Status: ${validation.status} | ` +
      `Is Active: ${validation.isActive ? 'Yes' : 'No'} | Domain Valid: ${validation.domainValid ? 'Yes' : 'No'} | ` +
      `User Active: ${validation.userActive} | Method: ${validation.method} | ` +
      `MX: ${validation.primaryMx || 'N/A'} | Details: ${validation.reason} (${elapsedMs} ms)`
    );
  }

  logStatus('--------------------------------------------------------------------------------');
  logStatus(`[INFO] Writing ${outputRows.length - 1} validated record(s) to sheet "${resultSheetName}"...`);

  // Write results in bulk
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
    reason: ''
  };

  // 1. Syntax Validation (RFC 5322 compliant check)
  const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/;
  if (!emailRegex.test(email)) {
    result.reason = 'Invalid email syntax format';
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
  result.userActive = userCheck.userActive; // 'Yes' | 'No' | 'No (Suspended)' | 'Unverified (Catch-All)' | 'Unverified (Needs SMTP API)'
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

  // 5. If the domain is valid (e.g. Yahoo, Zoho, custom cPanel SMTP), but the user could not be
  // verified without an SMTP handshake API key, do NOT falsely mark the user as Active.
  return {
    isActive: false,
    userActive: 'Unverified (Needs SMTP API)',
    status: 'Domain Valid / User Unverified',
    method: 'DNS MX Only (No SMTP Probe)',
    reason: `Domain "${domain}" is valid (MX: ${dnsInfo.primaryMx}), but verifying user "${email}" on this mail server requires an SMTP Verifier API key (Email Validator -> Configure SMTP API Key)`
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

    // IfExistsResult === 1 means the user definitely does NOT exist even though the domain is valid
    if (userCode === 1) {
      return {
        isActive: false,
        userActive: 'No',
        status: 'Inactive (User Not Found in Domain)',
        method: 'Microsoft 365 User Lookup',
        reason: `Domain "${domain}" is valid, but user "${email}" does not exist or is inactive`
      };
    }

    // IfExistsResult === 0, 5, or 6 indicates the user exists.
    // Run a one-time canary test on custom domains to ensure the domain doesn't return 0 for every random user.
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
 * Returns IfExistsResult integer (0 = exists, 1 = does not exist) or null on error.
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
    primaryMx: ''
  };

  try {
    const mxUrl = `https://dns.google/resolve?name=${encodeURIComponent(domain)}&type=MX`;
    const mxResponse = UrlFetchApp.fetch(mxUrl, { muteHttpExceptions: true });
    const mxData = JSON.parse(mxResponse.getContentText());

    // Status 3 = NXDOMAIN (Non-Existent Domain)
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
 * Strictly verifies that the individual user exists and is not on an unverifiable catch-all.
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
    // Apply conditional formatting to Overall Status (Col B), Is Active (Col C), Domain Valid (Col D), User Active in Domain (Col E)
    const statusRange = sheet.getRange(2, 2, totalRows - 1, 4);

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
