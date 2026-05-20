// ─────────────────────────────────────────────────────────────────
//  DWELLY LEAVE TRACKER — Google Apps Script
//  Serves the HTML frontend AND backend in one project.
// ─────────────────────────────────────────────────────────────────
//
//  SETUP:
//  1. https://script.google.com → open your existing project (or new one)
//  2. Replace Code.gs with this file's contents
//  3. File menu (top-left, ☰ icon) → click + → HTML → name it `index`
//     → paste the entire contents of index.html into it → save
//  4. Edit SHEET_ID below (from your Sheet URL)
//  5. Deploy → Manage deployments → pencil ✎ on the active one →
//     Version: New version → Deploy
//  6. Open the Web app URL — that's your live tracker
//
// ─────────────────────────────────────────────────────────────────

const SHEET_ID   = 'YOUR_SHEET_ID';            // from your Sheet URL
const SHEET_NAME = 'Absences';
const NOTIFY     = [
  'katya@dwelly.group'
  // Add agency director email here later, e.g.:
  // 'director@dwelly.group'
];

// ─── SERVE THE FRONTEND ──────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dwelly Leave Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── SHEET HELPERS ───────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','person','team','type','startDate','endDate','notes','calendarEventId','submittedBy','submittedAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRow(sheet, id) {
  const ids = sheet.getRange('A:A').getValues().flat();
  const i   = ids.indexOf(id);
  return i > 0 ? i + 1 : -1;
}

function workingDays(startStr, endStr) {
  let count = 0;
  const cur = new Date(startStr + 'T12:00:00');
  const end = new Date(endStr   + 'T12:00:00');
  while (cur <= end) {
    const d = cur.getDay();
    if (d !== 0 && d !== 6) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

// ─── API EXPOSED TO google.script.run ────────────────────────────
function fetchAbsences() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  return rows.slice(1).map(r => ({
    id:              String(r[0] || ''),
    person:          String(r[1] || ''),
    team:            String(r[2] || ''),
    type:            String(r[3] || ''),
    startDate:       String(r[4] || ''),
    endDate:         String(r[5] || ''),
    notes:           String(r[6] || ''),
    calendarEventId: String(r[7] || ''),
    submittedBy:     String(r[8] || ''),
    submittedAt:     String(r[9] || '')
  })).filter(a => a.id);
}

function createAbsence(data) {
  const sheet = getSheet();
  sheet.appendRow([
    data.id, data.person, data.team, data.type,
    data.startDate, data.endDate, data.notes || '',
    '', data.submittedBy, data.submittedAt
  ]);

  const calId = createCalEvent(data);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 8).setValue(calId);

  sendMail(data, 'logged');
  return { calendarEventId: calId };
}

function updateAbsence(data) {
  const sheet = getSheet();
  const row   = findRow(sheet, data.id);
  if (row < 0) throw new Error('Absence not found: ' + data.id);

  sheet.getRange(row, 3, 1, 5).setValues([[
    data.team, data.type, data.startDate, data.endDate, data.notes || ''
  ]]);

  if (data.calendarEventId) {
    try {
      const cal   = CalendarApp.getDefaultCalendar();
      const event = cal.getEventById(data.calendarEventId);
      if (event) {
        const endDate = new Date(data.endDate + 'T12:00:00');
        endDate.setDate(endDate.getDate() + 1);
        event.setTitle(`${data.type} – ${data.person}`);
        event.setAllDayDates(new Date(data.startDate + 'T12:00:00'), endDate);
        const wd = workingDays(data.startDate, data.endDate);
        event.setDescription(
          `Team: ${data.team}\nType: ${data.type}\n` +
          `Dates: ${data.startDate} → ${data.endDate} (${wd} working days)\n` +
          `Notes: ${data.notes || 'None'}`
        );
      }
    } catch (e) {
      Logger.log('Calendar update skipped: ' + e.message);
    }
  }

  sendMail(data, 'updated');
  return { ok: true };
}

function deleteAbsence(data) {
  const sheet = getSheet();
  const row   = findRow(sheet, data.id);
  if (row > 0) sheet.deleteRow(row);

  if (data.calendarEventId) {
    try {
      const event = CalendarApp.getDefaultCalendar().getEventById(data.calendarEventId);
      if (event) event.deleteEvent();
    } catch (e) {
      Logger.log('Calendar delete skipped: ' + e.message);
    }
  }

  sendMail(data, 'cancelled');
  return { ok: true };
}

// ─── CALENDAR ────────────────────────────────────────────────────
function createCalEvent(data) {
  const wd    = workingDays(data.startDate, data.endDate);
  const start = new Date(data.startDate + 'T12:00:00');
  const end   = new Date(data.endDate   + 'T12:00:00');
  end.setDate(end.getDate() + 1);

  const event = CalendarApp.getDefaultCalendar().createAllDayEvent(
    `${data.type} – ${data.person}`,
    start,
    end,
    {
      description:
        `Team: ${data.team}\nType: ${data.type}\n` +
        `Dates: ${data.startDate} → ${data.endDate} (${wd} working days)\n` +
        `Notes: ${data.notes || 'None'}\nLogged by: ${data.submittedBy}`,
      guests:      NOTIFY.join(','),
      sendInvites: true
    }
  );
  return event.getId();
}

// ─── EMAIL ───────────────────────────────────────────────────────
function sendMail(data, action) {
  const wd   = workingDays(data.startDate, data.endDate);
  const verb = { logged: 'logged', updated: 'UPDATED', cancelled: 'CANCELLED' }[action] || action;

  const subject =
    action === 'logged'    ? `[Leave Tracker] ${data.type} logged — ${data.person}, ${data.startDate} → ${data.endDate}` :
    action === 'updated'   ? `[Leave Tracker] UPDATED — ${data.type}, ${data.person}, ${data.startDate} → ${data.endDate}` :
                             `[Leave Tracker] CANCELLED — ${data.type}, ${data.person}, ${data.startDate} → ${data.endDate}`;

  const html = `
<div style="font-family:sans-serif;font-size:14px;color:#0f172a;max-width:480px;">
  <p style="margin:0 0 16px;">An absence has been <strong>${verb}</strong> on the Dwelly Leave Tracker.</p>
  <table style="border-collapse:collapse;width:100%;">
    <tr><td style="padding:6px 16px 6px 0;color:#64748b;white-space:nowrap">Person</td>
        <td style="padding:6px 0"><strong>${data.person}</strong></td></tr>
    <tr><td style="padding:6px 16px 6px 0;color:#64748b">Team</td>
        <td>${data.team}</td></tr>
    <tr><td style="padding:6px 16px 6px 0;color:#64748b">Type</td>
        <td>${data.type}</td></tr>
    <tr><td style="padding:6px 16px 6px 0;color:#64748b">Dates</td>
        <td>${data.startDate} → ${data.endDate}</td></tr>
    <tr><td style="padding:6px 16px 6px 0;color:#64748b">Duration</td>
        <td>${wd} working day${wd !== 1 ? 's' : ''}</td></tr>
    <tr><td style="padding:6px 16px 6px 0;color:#64748b">Notes</td>
        <td>${data.notes || 'None'}</td></tr>
  </table>
</div>`;

  GmailApp.sendEmail(
    NOTIFY.join(','),
    subject,
    `${data.type} ${verb} — ${data.person}\n${data.startDate} → ${data.endDate} (${wd} days)\nNotes: ${data.notes || 'None'}`,
    { htmlBody: html, name: 'Dwelly Leave Tracker' }
  );
}
