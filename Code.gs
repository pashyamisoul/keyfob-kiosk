/**
 * Keyfob Kiosk: Google Apps Script backend
 *
 * Bound to a Google Sheet that acts as the database.
 * Deployed as a Web App; the kiosk tablet just opens the URL.
 *
 * Sheets used:
 *   - "Fob Log"  one row per borrow event; return updates the same row.
 *                Fobs currently out are derived from this log. There is
 *                no fixed inventory; the fob number is typed in from the
 *                number printed on the fob itself.
 *   - "Teams"    list of team names shown in the kiosk dropdown
 *
 * Borrowers can be company employees (company domain enforced) or external
 * visitors (any valid email; their company and a contact person recorded).
 *
 * Emails: confirmation on check-out and check-in. A reminder goes out once
 * the fob has been out for 24 h, and again every further 24 h until it is
 * returned (hourly trigger, see installTriggers()).
 *
 * First-time setup: run setup() once, then installTriggers() once, from
 * the Apps Script editor.
 */

const CFG = {
  LOG_SHEET: 'Fob Log',
  TEAMS_SHEET: 'Teams',
  EMAIL_DOMAIN: 'example.com', // your company domain, enforced for employees only
  SEND_CONFIRMATION_EMAIL: true,
  REMINDER_EVERY_HOURS: 24,    // first reminder after 24 h out, then every 24 h
  // GDPR: external visitors' name/email are redacted from the log this many
  // days after the fob is returned (daily job, see installTriggers()).
  EXTERNAL_RETENTION_DAYS: 30,
};

// Log sheet columns (0-based indexes into a row)
const COL = {
  LOG_ID: 0, FOB: 1, NAME: 2, EMAIL: 3, TYPE: 4, TEAM: 5, MANAGER: 6,
  OUT_AT: 7, STATUS: 8, RETURNED_AT: 9, RETURNED_BY: 10,
  RETURNED_BY_EMAIL: 11, NOTES: 12, REMINDERS: 13, CONSENT: 14,
};

const LOG_HEADER = [
  'Log ID', 'Fob Number', 'Name', 'Email', 'Type', 'Team / Company',
  'Manager / Contact', 'Checked Out At', 'Status', 'Returned At',
  'Returned By Name', 'Returned By Email', 'Notes', 'Reminders Sent',
  'Consent',
];

// ---------------------------------------------------------------------------
// Web app entry point
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Keyfob Kiosk')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------------------
// One-time setup, run manually from the Apps Script editor.
// Safe to re-run: it also migrates an existing Fob Log to the newest layout.
// ---------------------------------------------------------------------------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(CFG.LOG_SHEET);
    log.appendRow(LOG_HEADER);
    log.setFrozenRows(1);
    log.getRange('1:1').setFontWeight('bold');
  } else {
    migrateLog_(log);
  }
  formatTimestampColumns_(log);

  let teams = ss.getSheetByName(CFG.TEAMS_SHEET);
  if (!teams) {
    teams = ss.insertSheet(CFG.TEAMS_SHEET);
    teams.appendRow(['Team']);
    teams.setFrozenRows(1);
    teams.getRange('1:1').setFontWeight('bold');
    [
      'Engineering', 'Product', 'Design', 'Marketing', 'Sales',
      'People & Culture', 'Finance', 'Office Management', 'IT', 'Other',
    ].forEach((t) => teams.appendRow([t]));
  }
}

/** Bring an older Fob Log sheet up to the current column layout. */
function migrateLog_(log) {
  let header = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0].map(String);
  if (header.indexOf('Type') === -1) {
    log.insertColumnAfter(4); // after Email
    log.getRange(1, 5).setValue('Type').setFontWeight('bold');
  }
  ['Reminders Sent', 'Consent'].forEach(function (name) {
    const h = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0].map(String);
    if (h.indexOf(name) === -1) {
      log.getRange(1, h.length + 1).setValue(name).setFontWeight('bold');
    }
  });
  // Refresh the descriptive header names (harmless if unchanged)
  log.getRange(1, 1, 1, LOG_HEADER.length).setValues([LOG_HEADER]).setFontWeight('bold');
}

const TIMESTAMP_FORMAT = 'yyyy-mm-dd hh:mm:ss';

/** Show full date AND time in the Checked Out At / Returned At columns. */
function formatTimestampColumns_(log) {
  [COL.OUT_AT, COL.RETURNED_AT].forEach(function (c) {
    log.getRange(2, c + 1, Math.max(log.getMaxRows() - 1, 1), 1)
      .setNumberFormat(TIMESTAMP_FORMAT);
  });
}

/**
 * Run once from the editor: installs the hourly trigger that sends
 * return reminders. Safe to re-run (replaces any existing trigger).
 */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'sendReturnReminders' || fn === 'redactExpiredExternalData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendReturnReminders').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('redactExpiredExternalData').timeBased().everyDays(1).atHour(3).create();
}

// ---------------------------------------------------------------------------
// Data the kiosk UI needs to render its dropdowns
// ---------------------------------------------------------------------------

/**
 * Look up a single fob typed on the return form. Returns the borrower's
 * name/email so the form can prefill them. Privacy: this only answers for
 * one specific fob number the person already has in hand; the full list of
 * outstanding fobs is never sent to the client.
 */
function lookupFob(fobRaw) {
  const fob = normalizeFob_(fobRaw);
  if (!fob) return { found: false };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = getOutFobs_(mustGetSheet_(ss, CFG.LOG_SHEET))[fob];
  if (!out) return { found: false };
  return { found: true, name: out.name, email: out.email };
}

// Privacy: the kiosk URL may be publicly reachable, so this must never
// expose who currently holds which fob. Only the team list leaves the server.
function getKioskData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const teamsSheet = ss.getSheetByName(CFG.TEAMS_SHEET);
  const teams = teamsSheet
    ? teamsSheet.getDataRange().getValues().slice(1).map((r) => String(r[0]).trim()).filter(Boolean)
    : [];

  return { teams: teams };
}

// ---------------------------------------------------------------------------
// Check-out
// ---------------------------------------------------------------------------

function issueFob(form) {
  const name = clean_(form.name);
  const email = clean_(form.email).toLowerCase();
  const isEmployee = clean_(form.type) !== 'external';
  const type = isEmployee ? 'Employee' : 'External';
  const team = clean_(form.team);      // team for employees, company for externals
  const manager = clean_(form.manager); // manager for employees, contact person for externals
  const fob = normalizeFob_(form.fob);
  const notes = clean_(form.notes);
  const consent = form.consent === true || form.consent === 'true';

  if (!name || !email || !team || !manager || !fob) {
    throw new Error('Please fill in all fields.');
  }
  validateEmail_(email, isEmployee);
  if (!isEmployee && !consent) {
    throw new Error('Please accept the privacy notice to continue.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = mustGetSheet_(ss, CFG.LOG_SHEET);

    const out = getOutFobs_(log)[fob];
    if (out) {
      // Do not name the current holder: this message is shown to anyone
      throw new Error('Fob ' + fob + ' is already checked out. Please double-check the number on the fob.');
    }

    const now = new Date();
    const logId = 'FOB-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + fob;

    log.appendRow([
      logId, fob, name, email, type, team, manager, now, 'OUT', '', '', '', notes, 0,
      isEmployee ? '' : 'Yes (kiosk checkbox)',
    ]);
    // appendRow can grow the grid past the pre-formatted range
    log.getRange(log.getLastRow(), COL.OUT_AT + 1).setNumberFormat(TIMESTAMP_FORMAT);

    if (CFG.SEND_CONFIRMATION_EMAIL) {
      trySendEmail_(email, 'Keyfob ' + fob + ' has been issued to you',
        'Hi ' + name + ',\n\n' +
        'Keyfob ' + fob + ' has been issued to you on ' +
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'd MMM yyyy, HH:mm') + '.\n\n' +
        'Please return it at reception when you are done. You will get a ' +
        'confirmation email once it is checked back in.\n\n' +
        'Best regards,\nOffice Management');
    }

    return { ok: true, fob: fob, name: name, email: email };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Check-in
// ---------------------------------------------------------------------------

function returnFob(form) {
  const name = clean_(form.name);
  const email = clean_(form.email).toLowerCase();
  const fob = normalizeFob_(form.fob);
  const returnNotes = clean_(form.notes);

  if (!name || !email || !fob) {
    throw new Error('Please fill in all fields.');
  }
  validateEmail_(email, false); // any valid email may return a fob

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = mustGetSheet_(ss, CFG.LOG_SHEET);
    const out = getOutFobs_(log)[fob];
    if (!out) {
      throw new Error('Fob ' + fob + ' is not marked as checked out.');
    }

    const now = new Date();
    let note = '';
    if (out.email && out.email !== email) {
      note = 'Returned by a different person than the borrower.';
    }
    // Merge: existing checkout note + mismatch note + note typed on return
    const existingNote = clean_(log.getRange(out.row, COL.NOTES + 1).getValue());
    const mergedNotes = [existingNote, note, returnNotes ? 'On return: ' + returnNotes : '']
      .filter(Boolean).join(' | ');
    log.getRange(out.row, COL.STATUS + 1, 1, 5).setValues([['RETURNED', now, name, email, mergedNotes]]);
    log.getRange(out.row, COL.RETURNED_AT + 1).setNumberFormat(TIMESTAMP_FORMAT);

    if (CFG.SEND_CONFIRMATION_EMAIL) {
      const when = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd MMM yyyy, HH:mm');
      trySendEmail_(email, 'Keyfob ' + fob + ' has been returned',
        'Hi ' + name + ',\n\n' +
        'Keyfob ' + fob + ' has been marked as returned on ' + when + '. Thank you!\n\n' +
        'Best regards,\nOffice Management');
      // If someone else returned it, also close the loop with the borrower
      if (note && out.email) {
        trySendEmail_(out.email, 'Keyfob ' + fob + ' has been returned',
          'Hi ' + out.name + ',\n\n' +
          'Keyfob ' + fob + ', which was issued to you, was returned by ' + name +
          ' on ' + when + '. Nothing more to do.\n\n' +
          'Best regards,\nOffice Management');
      }
    }

    return { ok: true, fob: fob, name: name, email: email };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Reminders: runs hourly via the trigger from installTriggers().
// First reminder once a fob has been out 24 h, then another every 24 h,
// until it is returned. "Reminders Sent" in the log tracks what went out.
// ---------------------------------------------------------------------------

function sendReturnReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = mustGetSheet_(ss, CFG.LOG_SHEET);
  const values = log.getDataRange().getValues();
  const now = Date.now();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[COL.STATUS]).toUpperCase() !== 'OUT') continue;
    const email = String(row[COL.EMAIL] || '').trim();
    const outAt = row[COL.OUT_AT] ? new Date(row[COL.OUT_AT]).getTime() : 0;
    if (!email || !outAt) continue;

    const hoursOut = (now - outAt) / 3600000;
    const due = Math.floor(hoursOut / CFG.REMINDER_EVERY_HOURS);
    const sent = Number(row[COL.REMINDERS] || 0);
    if (due > sent) {
      const days = Math.round(hoursOut / 24 * 10) / 10;
      trySendEmail_(email, 'Reminder: keyfob ' + row[COL.FOB] + ' return is pending',
        'Hi ' + row[COL.NAME] + ',\n\n' +
        'Keyfob ' + row[COL.FOB] + ' was issued to you ' + days + ' day(s) ago and has ' +
        'not been returned yet. Kindly return it at reception soon.\n\n' +
        'If you believe you already returned it, please contact Office Management.\n\n' +
        'Best regards,\nOffice Management');
      log.getRange(i + 1, COL.REMINDERS + 1).setValue(due);
    }
  }
}

// ---------------------------------------------------------------------------
// GDPR retention: runs daily via the trigger from installTriggers().
// External visitors' personal data (name, email, returned-by, notes) is
// redacted EXTERNAL_RETENTION_DAYS after the fob was returned. The row
// itself stays for statistics (fob, company, dates). Employee rows
// are internal records and are left untouched.
// ---------------------------------------------------------------------------

function redactExpiredExternalData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = mustGetSheet_(ss, CFG.LOG_SHEET);
  const values = log.getDataRange().getValues();
  const cutoff = Date.now() - CFG.EXTERNAL_RETENTION_DAYS * 24 * 60 * 60 * 1000;
  const REDACTED = 'REDACTED (GDPR retention)';

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[COL.TYPE]) !== 'External') continue;
    if (String(row[COL.STATUS]).toUpperCase() !== 'RETURNED') continue;
    if (String(row[COL.NAME]) === REDACTED) continue; // already done
    const returnedAt = row[COL.RETURNED_AT] ? new Date(row[COL.RETURNED_AT]).getTime() : 0;
    if (!returnedAt || returnedAt > cutoff) continue;

    log.getRange(i + 1, COL.NAME + 1, 1, 2).setValues([[REDACTED, REDACTED]]);
    log.getRange(i + 1, COL.RETURNED_BY + 1, 1, 3).setValues([[REDACTED, REDACTED, REDACTED]]);
  }
}

// ---------------------------------------------------------------------------
// Presentation: run styleSheets() once from the editor (and again anytime,
// it is safe to re-run). Builds a live Dashboard tab and makes the Fob Log
// readable for non-technical users. All looks live in code so they survive
// handover and can be restored with one run.
// ---------------------------------------------------------------------------

const BRAND = {
  primary: '#1E3A8A', primaryDark: '#16306F', cream: '#F8FAFC',
  grey: '#CBD5E1', black: '#0F172A', muted: '#475569',
  dandelion: '#FACC15', dandelionLight: '#FEF9C3', tangerineLight: '#FEE2E2',
  danger: '#B91C1C', emerald: '#16A34A', emeraldDark: '#14532D',
  emeraldLight: '#F0FDF4', redactedGrey: '#EEF1F4', redactedText: '#94A3B8',
  white: '#FFFFFF',
};

function styleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = mustGetSheet_(ss, CFG.LOG_SHEET);
  styleLog_(log);
  buildDashboard_(ss);
  const teams = ss.getSheetByName(CFG.TEAMS_SHEET);
  if (teams) teams.setTabColor(BRAND.emerald);
}

/** Traffic-light formatting, tidy columns, and gentle protection for the log. */
function styleLog_(log) {
  const rows = log.getMaxRows();

  // Header: white on the primary color, frozen; keep fob + name in view
  log.getRange(1, 1, 1, LOG_HEADER.length)
    .setBackground(BRAND.primary).setFontColor(BRAND.white).setFontWeight('bold');
  log.setFrozenRows(1);
  log.setFrozenColumns(3);
  log.setTabColor(BRAND.grey);

  // Readable column widths (A stays narrow, it gets hidden below)
  const widths = { 2: 100, 3: 170, 4: 220, 5: 90, 6: 170, 7: 170, 8: 155, 9: 110, 10: 155, 11: 170, 12: 220, 13: 240 };
  Object.keys(widths).forEach(function (c) { log.setColumnWidth(Number(c), widths[c]); });

  // Hide the technical columns the team never needs
  log.showColumns(1, LOG_HEADER.length); // reset first so re-runs stay correct
  log.hideColumns(COL.LOG_ID + 1);
  log.hideColumns(COL.REMINDERS + 1, 2); // Reminders Sent + Consent

  // Permanent filter row
  const existingFilter = log.getFilter();
  if (existingFilter) existingFilter.remove();
  log.getRange(1, 1, rows, LOG_HEADER.length).createFilter();

  // Traffic lights
  const dataRange = log.getRange(2, 1, rows - 1, LOG_HEADER.length);
  const statusRange = log.getRange(2, COL.STATUS + 1, rows - 1, 1);
  const rules = [
    // Redacted GDPR rows read as archived
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="REDACTED (GDPR retention)"')
      .setBackground(BRAND.redactedGrey).setFontColor(BRAND.redactedText)
      .setRanges([dataRange]).build(),
    // Whole row tinted when a fob is out for more than 24 h
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($I2="OUT", $H2<>"", $H2<NOW()-1)')
      .setBackground(BRAND.tangerineLight)
      .setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OUT')
      .setBackground(BRAND.dandelionLight).setFontColor(BRAND.primaryDark).setBold(true)
      .setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('RETURNED')
      .setBackground(BRAND.emeraldLight).setFontColor(BRAND.emeraldDark)
      .setRanges([statusRange]).build(),
  ];
  log.setConditionalFormatRules(rules);

  protectWarnOnly_(log, 'The kiosk manages this data. Edit only if you know what you are doing.');
}

/** Live overview tab: counters, currently-out list, latest check-outs. */
function buildDashboard_(ss) {
  const old = ss.getSheetByName('Dashboard');
  if (old) ss.deleteSheet(old);
  const dash = ss.insertSheet('Dashboard', 0);
  dash.setTabColor(BRAND.primary);
  dash.setHiddenGridlines(true);
  dash.getRange('A1:M60').setBackground(BRAND.cream);

  // Layout: A is a margin column; cards sit in B, D, F
  dash.setColumnWidth(1, 24);
  [[2, 150], [3, 210], [4, 210], [5, 165], [6, 150], [7, 30], [8, 165], [9, 90], [10, 200], [11, 120]]
    .forEach(function (w) { dash.setColumnWidth(w[0], w[1]); });

  dash.getRange('B2').setValue('Keyfob Dashboard')
    .setFontSize(22).setFontWeight('bold').setFontColor(BRAND.black);
  dash.getRange('B3').setValue('Live overview. Updates automatically from the Fob Log tab, nothing to maintain.')
    .setFontColor(BRAND.muted).setFontSize(10);

  // KPI cards
  const cards = [
    ['B', 'FOBS OUT NOW', '=COUNTIF(\'Fob Log\'!$I$2:$I,"OUT")', BRAND.primaryDark],
    ['D', 'OVERDUE (OUT > 24 H)', '=COUNTIFS(\'Fob Log\'!$I$2:$I,"OUT",\'Fob Log\'!$H$2:$H,"<"&NOW()-1)', BRAND.danger],
    ['F', 'RETURNED THIS WEEK', '=COUNTIFS(\'Fob Log\'!$I$2:$I,"RETURNED",\'Fob Log\'!$J$2:$J,">="&TODAY()-WEEKDAY(TODAY(),2)+1)', BRAND.emeraldDark],
  ];
  cards.forEach(function (card) {
    const colLetter = card[0];
    dash.getRange(colLetter + '5:' + colLetter + '6').setBackground(BRAND.white)
      .setBorder(true, true, true, true, false, false, BRAND.grey, SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(colLetter + '5').setValue(card[1])
      .setFontSize(9).setFontWeight('bold').setFontColor(BRAND.muted)
      .setVerticalAlignment('bottom').setHorizontalAlignment('center');
    dash.getRange(colLetter + '6').setFormula(card[2])
      .setFontSize(30).setFontWeight('bold').setFontColor(card[3])
      .setHorizontalAlignment('center').setVerticalAlignment('top').setNumberFormat('0');
  });
  dash.setRowHeight(6, 52);

  // Currently out (left)
  dash.getRange('B9').setValue('CURRENTLY OUT')
    .setFontWeight('bold').setFontColor(BRAND.muted).setFontSize(10);
  const outHeaders = ['Fob', 'Who', 'Team / Company', 'Checked out at', 'Days out'];
  dash.getRange('B10:F10').setValues([outHeaders])
    .setFontWeight('bold').setFontColor(BRAND.black)
    .setBorder(false, false, true, false, false, false, BRAND.grey, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  dash.getRange('B11').setFormula(
    '=IF(COUNTIF(\'Fob Log\'!$I$2:$I,"OUT")=0,"All fobs are in.",' +
    'QUERY(\'Fob Log\'!$A$2:$O,"select B, C, F, H where I = \'OUT\' order by H asc",0))');
  dash.getRange('F11').setFormula('=ARRAYFORMULA(IF(LEN($E$11:$E$40)=0,,ROUND(NOW()-$E$11:$E$40,1)))');
  dash.getRange('E11:E40').setNumberFormat('yyyy-mm-dd hh:mm');
  dash.getRange('F11:F40').setNumberFormat('0.0');

  // Latest check-outs (right)
  dash.getRange('H9').setValue('LATEST CHECK-OUTS')
    .setFontWeight('bold').setFontColor(BRAND.muted).setFontSize(10);
  dash.getRange('H10:K10').setValues([['When', 'Fob', 'Who', 'Status']])
    .setFontWeight('bold').setFontColor(BRAND.black)
    .setBorder(false, false, true, false, false, false, BRAND.grey, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  dash.getRange('H11').setFormula(
    '=IFERROR(QUERY(\'Fob Log\'!$A$2:$O,"select H, B, C, I where B is not null order by H desc limit 10",0),)');
  dash.getRange('H11:H21').setNumberFormat('yyyy-mm-dd hh:mm');

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(1)
      .setFontColor(BRAND.danger).setBold(true)
      .setRanges([dash.getRange('F11:F40')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OUT')
      .setFontColor(BRAND.primaryDark).setBold(true)
      .setRanges([dash.getRange('K11:K21')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('RETURNED')
      .setFontColor(BRAND.emeraldDark)
      .setRanges([dash.getRange('K11:K21')]).build(),
  ];
  dash.setConditionalFormatRules(rules);

  protectWarnOnly_(dash, 'This dashboard is generated. Re-run styleSheets() to rebuild it.');
}

function protectWarnOnly_(sheet, description) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(function (p) { p.remove(); });
  sheet.protect().setDescription(description).setWarningOnly(true);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Scan the log for fobs currently OUT.
 * Returns { fobNumber: { row (1-based sheet row), name, email, outAt (ms) } }
 */
function getOutFobs_(log) {
  const values = log.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const fob = normalizeFob_(values[i][COL.FOB]);
    if (!fob) continue;
    if (String(values[i][COL.STATUS]).toUpperCase() === 'OUT') {
      out[fob] = {
        row: i + 1,
        name: String(values[i][COL.NAME] || ''),
        email: String(values[i][COL.EMAIL] || '').toLowerCase(),
        outAt: values[i][COL.OUT_AT] ? new Date(values[i][COL.OUT_AT]).getTime() : 0,
      };
    } else if (out[fob]) {
      // A later RETURNED row shouldn't normally exist after an OUT row,
      // but if the sheet was edited by hand, trust the most recent state.
      delete out[fob];
    }
  }
  return out;
}

function clean_(v) {
  return String(v == null ? '' : v).trim();
}

/** Fob numbers are typed by hand: trim, uppercase, collapse inner spaces. */
function normalizeFob_(v) {
  return clean_(v).toUpperCase().replace(/\s+/g, '');
}

function validateEmail_(email, requireCompanyDomain) {
  if (requireCompanyDomain) {
    const re = new RegExp('^[A-Za-z0-9._%+-]+@' + CFG.EMAIL_DOMAIN.replace('.', '\\.') + '$', 'i');
    if (!re.test(email)) {
      throw new Error('Please use your @' + CFG.EMAIL_DOMAIN + ' email address.');
    }
  } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
    throw new Error('Please enter a valid email address.');
  }
}

function mustGetSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Run setup() first.');
  return sheet;
}

function trySendEmail_(to, subject, body) {
  try {
    MailApp.sendEmail(to, subject, body);
  } catch (e) {
    // Never let a mail failure break a check-out/in
    console.warn('Email to ' + to + ' failed: ' + e);
  }
}
