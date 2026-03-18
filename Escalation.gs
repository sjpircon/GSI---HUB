// ============================================================
// ESCALATION.GS — Hub Web App
// Handles form submissions from BSI, PSI, and FSI feedback pages.
// Routes each submission to the correct tab on the shared escalation sheet.
//
// REQUIRED SCRIPT PROPERTIES (Hub Apps Script → Project Settings → Script Properties):
//   ESCALATION_SHEET_ID   — ID of the shared Google Sheet
//                           (the portion of the URL between /d/ and /edit)
//
// OPTIONAL SCRIPT PROPERTIES:
//   SLACK_WEBHOOK_URL     — Shared fallback webhook for all tools
//   SLACK_WEBHOOK_BSI     — BSI-specific webhook (overrides shared)
//   SLACK_WEBHOOK_PSI     — PSI-specific webhook (overrides shared)
//   SLACK_WEBHOOK_FSI     — FSI-specific webhook (overrides shared)
//
// Sheet tabs are auto-created on first submission if they don't exist.
// Tab names must match TAB_MAP exactly (case-sensitive):
//   BSI Escalation Log
//   PSI Escalation Log
//   FSI Escalation Log
// ============================================================

const ESCALATION_DRAFT_KEY = 'ESCALATION_DRAFT';

const ESCALATION_HEADERS = [
  'Customer Name',
  'PMREO Link',
  'Spreadsheet Link',
  'Database Link',
  'Issue Type',
  'Object Type',
  'Description',
  'Screenshots',
  'Submitted By',
  'Timestamp',
  'Status'
];

const TAB_MAP = {
  BSI: 'BSI Escalation Log',
  PSI: 'PSI Escalation Log',
  FSI: 'FSI Escalation Log'
};


// ── Step 1: Save draft to Script Properties ──────────────────
//
// Called by the Hub feedback form before submitEscalation().
// Avoids google.script.run argument serialization issues by
// staging the data server-side first.
//
// @param {string} tool            - 'BSI' | 'PSI' | 'FSI'
// @param {string} customerName
// @param {string} pmreoLink
// @param {string} spreadsheetLink
// @param {string} databaseLink
// @param {string} issueType       - auto-classified by the UI category
// @param {string} objectType
// @param {string} description
// @param {string} screenshots     - Drive link or 'N/A'
// @param {string} timestamp       - ISO string from client
// ─────────────────────────────────────────────────────────────
function saveEscalationDraft(
  tool,
  customerName,
  pmreoLink,
  spreadsheetLink,
  databaseLink,
  issueType,
  objectType,
  description,
  screenshots,
  timestamp
) {
  const draft = {
    tool:            String(tool            || 'BSI').trim().toUpperCase(),
    customerName:    String(customerName    || '').trim(),
    pmreoLink:       String(pmreoLink       || '').trim(),
    spreadsheetLink: String(spreadsheetLink || '').trim(),
    databaseLink:    String(databaseLink    || '').trim(),
    issueType:       String(issueType       || '').trim(),
    objectType:      String(objectType      || '').trim(),
    description:     String(description     || '').trim(),
    screenshots:     String(screenshots     || 'N/A').trim(),
    timestamp:       String(timestamp       || new Date().toISOString()),
    status:          'New'
  };

  PropertiesService.getScriptProperties()
    .setProperty(ESCALATION_DRAFT_KEY, JSON.stringify(draft));
}


// ── Step 2: Read draft and submit ────────────────────────────
//
// Reads the staged draft, appends it to the correct tool tab
// on the escalation sheet, then sends a Slack notification.
// Cleans up the draft key on success.
//
// Throws on missing configuration so the UI can surface the error.
// ─────────────────────────────────────────────────────────────
function submitEscalation() {
  const props = PropertiesService.getScriptProperties();

  // ── Read and parse draft ────────────────────────────────
  const raw = props.getProperty(ESCALATION_DRAFT_KEY);
  if (!raw) throw new Error(
    'No escalation draft found. Please fill out the form again.'
  );

  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    throw new Error('Could not read escalation draft: ' + e.message);
  }

  // ── Resolve sheet ───────────────────────────────────────
  const sheetId = props.getProperty('ESCALATION_SHEET_ID');
  if (!sheetId) throw new Error(
    'ESCALATION_SHEET_ID is not set. ' +
    'Go to Hub Apps Script → Project Settings → Script Properties and add it.'
  );

  const tool    = data.tool || 'BSI';
  const tabName = TAB_MAP[tool] || 'BSI Escalation Log';
  const ss      = SpreadsheetApp.openById(sheetId);
  let   sheet   = ss.getSheetByName(tabName);

  // Auto-create tab with styled headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    const hdrRange = sheet.getRange(1, 1, 1, ESCALATION_HEADERS.length);
    hdrRange.setValues([ESCALATION_HEADERS])
            .setFontWeight('bold')
            .setBackground('#1E2430')
            .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(7, 320); // Description column — wider for readability
  }

  const submittedBy = Session.getActiveUser().getEmail();

  // ── Append row (order matches ESCALATION_HEADERS) ───────
  sheet.appendRow([
    data.customerName,
    data.pmreoLink,
    data.spreadsheetLink,
    data.databaseLink,
    data.issueType,
    data.objectType,
    data.description,
    data.screenshots,
    submittedBy,
    data.timestamp,
    data.status
  ]);

  // ── Slack notification (non-fatal) ──────────────────────
  // Per-tool webhook takes priority over the shared fallback.
  const webhookUrl =
    props.getProperty('SLACK_WEBHOOK_' + tool) ||
    props.getProperty('SLACK_WEBHOOK_URL')      ||
    '';

  if (webhookUrl) {
    _sendSlackNotification(webhookUrl, data, submittedBy);
  }

  // ── Clean up draft ──────────────────────────────────────
  props.deleteProperty(ESCALATION_DRAFT_KEY);
}


// ── Slack notification ───────────────────────────────────────
//
// Sends a payload to the configured Slack Workflow trigger URL.
// Keys must exactly match the variable names defined in the
// Slack Workflow — any mismatch renders that variable blank.
//
// Non-fatal: a Slack failure never blocks the sheet write.
// ─────────────────────────────────────────────────────────────
function _sendSlackNotification(webhookUrl, data, submittedBy) {
  const payload = {
    tool:            data.tool,
    customerName:    data.customerName,
    pmreoLink:       data.pmreoLink,
    spreadsheetLink: data.spreadsheetLink,
    databaseLink:    data.databaseLink,
    issueType:       data.issueType,
    objectType:      data.objectType,
    description:     data.description,
    screenshots:     data.screenshots,
    submittedBy:     submittedBy,
    timestamp:       data.timestamp,
    status:          data.status
  };

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('Slack notification failed: ' + e.message);
  }
}
