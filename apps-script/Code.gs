/**
 * Synced AI onboarding form → Google Sheet
 *
 * This file is a copy for version tracking. The live version runs in Google Apps Script
 * (open the Google Sheet → Extensions → Apps Script). After changing it there:
 * Deploy → Manage deployments → pencil icon → Version: "New version" → Deploy.
 * That keeps the same URL, so the form doesn't need to change.
 */

const SHEET_NAME = 'Onboarding';
const NOTIFY_EMAIL = 'info@syncedlead.com';  // who gets an email when a client finishes; '' turns it off

// Column headers, in order. `key` matches the field names the form sends.
const COLUMNS = [
  { key: 'id',        header: 'Submission ID' },
  { key: 'status',    header: 'Status' },
  { key: 'started',   header: 'Started' },
  { key: 'updated',   header: 'Last updated' },
  { key: 'fname',     header: 'First name' },
  { key: 'lname',     header: 'Last name' },
  { key: 'bizname',   header: 'Business name' },
  { key: 'phone',     header: 'Phone' },
  { key: 'email',     header: 'Email' },
  { key: 'ein',       header: 'EIN / Tax ID' },
  { key: 'meta',      header: 'Meta access' },
  { key: 'calendar',  header: 'Calendar' },
  { key: 'goal',      header: 'Clients goal (4 months)' },
  { key: 'marketing', header: 'Recent marketing' },
];

const META_LABELS = {
  shared: 'Client says shared',
  needs_help: 'Needs help',
  later: 'Will do later',
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const data = JSON.parse(e.postData.contents);
    const id = saveSubmission(data);
    return json({ ok: true, id: id });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function saveSubmission(data) {
  const sheet = getSheet();
  const now = new Date();
  const id = data.id || Utilities.getUuid();  // older versions of the form don't send an id

  // Find this submission's row, if an earlier step already saved it
  const lastRow = sheet.getLastRow();
  let rowIndex = -1;
  let existing = null;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const found = ids.findIndex(function (r) { return r[0] === id; });
    if (found !== -1) {
      rowIndex = found + 2;
      existing = sheet.getRange(rowIndex, 1, 1, COLUMNS.length).getValues()[0];
    }
  }

  const row = COLUMNS.map(function (col, i) {
    const old = existing ? existing[i] : '';
    switch (col.key) {
      case 'id':      return id;
      case 'started': return old || now;
      case 'updated': return now;
      case 'status':
        // Never downgrade a finished submission if a slower "partial" save arrives late
        if (old === 'Complete') return old;
        return data.status === 'partial' ? 'In progress' : 'Complete';
      case 'meta':
        return data.meta ? (META_LABELS[data.meta] || data.meta) : old;
      default: {
        const value = data[col.key];
        // Keep what's already saved if this save didn't include the field
        return value !== undefined && value !== '' ? asText(value) : old;
      }
    }
  });

  if (rowIndex === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  }

  // Email the team the first time a submission is finished
  const statusCol = COLUMNS.findIndex(function (c) { return c.key === 'status'; });
  const wasComplete = existing && existing[statusCol] === 'Complete';
  if (NOTIFY_EMAIL && row[statusCol] === 'Complete' && !wasComplete) {
    try {
      sendNotification(row);
    } catch (err) {
      console.error('Notification email failed:', err);  // the row is saved either way
    }
  }
  return id;
}

function sendNotification(row) {
  const get = function (key) {
    return String(row[COLUMNS.findIndex(function (c) { return c.key === key; })] || '').replace(/^'/, '');
  };
  const name = (get('fname') + ' ' + get('lname')).trim();
  const lines = COLUMNS
    .filter(function (c) { return ['id', 'status', 'started', 'updated', 'ein'].indexOf(c.key) === -1; })
    .map(function (c) { return c.header + ': ' + (get(c.key) || '-'); });

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    replyTo: get('email'),
    subject: 'New onboarding: ' + (get('bizname') || name),
    body: name + ' just finished the onboarding form.\n\n' + lines.join('\n') +
      '\n\nFull details (including EIN) in the sheet:\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS.map(function (c) { return c.header; }));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    // Keep phone numbers and EINs as text so Sheets doesn't strip leading zeros
    sheet.getRange('H:H').setNumberFormat('@');
    sheet.getRange('J:J').setNumberFormat('@');
  }
  return sheet;
}

// Stop answers like "=IMPORTXML(...)" from running as spreadsheet formulas
function asText(value) {
  const s = String(value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
