
/**
 * MOSAIC 2026 — Mission Results Recorder
 * Receives completed mission results from the React application.
 */

const SPREADSHEET_ID = '1EDjRRfjk7X2kZKeccGe4dKGB0zQB7acFMNsVAGvL29c';
const SHEET_NAME = 'Mission Results';

const HEADERS = [
  'Session ID',
  'Team Name',
  'Team Code',
  'Game Started At',
  'Stage 1 Completed At',
  'Stage 2 Completed At',
  'Game Ended At',
  'Overall Raw Time',
  'Overall Raw Seconds',
  'Stage 1 Raw Time',
  'Stage 1 Raw Seconds',
  'Stage 1 Penalties',
  'Stage 1 Penalty Seconds',
  'Stage 1 Adjusted Time',
  'Stage 1 Adjusted Seconds',
  'Stage 2 Raw Time',
  'Stage 2 Raw Seconds',
  'Stage 2 Penalties',
  'Stage 2 Penalty Seconds',
  'Stage 2 Adjusted Time',
  'Stage 2 Adjusted Seconds',
  'Stage 3 Raw Time',
  'Stage 3 Raw Seconds',
  'Stage 3 Overtime',
  'Stage 3 Overtime Seconds',
  'Stage 3 Overtime Penalty Seconds',
  'Stage 3 Penalty Seconds',
  'Total Penalty Count',
  'Total Penalty Time',
  'Total Penalty Seconds',
  'Adjusted Final Time',
  'Adjusted Seconds',
  'Displayed Result',
  'Recorded At'
];


function doGet() {
  return ContentService
    .createTextOutput(
      JSON.stringify({
        ok: true,
        service: 'MOSAIC 2026 Mission Results Recorder',
        version: 'diagnostic-v1',
        timestamp: new Date().toISOString()
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}


function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const body = e && e.postData
      ? e.postData.contents
      : '';

    if (!body) throw new Error('Missing request body.');

    const p = JSON.parse(body);
    validatePayload(p);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getResultsSheet(spreadsheet);
    const row = toRow(p);

    const lastRow = sheet.getLastRow();

    // Update an existing session instead of duplicating it.
    if (lastRow > 1) {
      const ids = sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues()
        .flat();

      const index = ids.indexOf(String(p.sessionId));

      if (index !== -1) {
        const rowNumber = index + 2;
        sheet.getRange(rowNumber, 1, 1, HEADERS.length)
          .setValues([row]);
        formatResultRow(sheet, rowNumber);
        SpreadsheetApp.flush();

        return jsonOutput({
          ok: true,
          updated: true,
          sessionId: String(p.sessionId)
        });
      }
    }

    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();

    formatResultRow(sheet, rowNumber);
    SpreadsheetApp.flush();

    return jsonOutput({
      ok: true,
      appended: true,
      sessionId: String(p.sessionId)
    });

  } catch (error) {
    return jsonOutput({
      ok: false,
      error: String(error && error.message ? error.message : error)
    });

  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function validatePayload(p) {
  if (!p || !p.sessionId) throw new Error('Missing sessionId.');
  if (!p.teamName) throw new Error('Missing teamName.');
  if (!p.startedAt) throw new Error('Missing startedAt.');
  if (!p.completedAt) throw new Error('Missing completedAt.');
}

function getResultsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      HEADERS.length - sheet.getMaxColumns()
    );
  }

  // This resets the header row to the current schema.
  // Existing data rows are not deleted.
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setWrap(true);

  return sheet;
}

function toRow(p) {
  return [
    p.sessionId || '',
    p.teamName || '',
    p.teamCode || '',
    toDate(p.startedAt),
    toDate(p.stage1CompletedAt),
    toDate(p.stage2CompletedAt),
    toDate(p.completedAt),

    p.rawTime || '00:00:00',
    numberOrZero(p.rawSeconds),

    p.stage1Duration || '00:00:00',
    numberOrZero(p.stage1DurationSeconds),
    numberOrZero(p.stage1Penalties),
    numberOrZero(p.stage1PenaltySeconds),
    p.stage1AdjustedTime || '00:00:00',
    numberOrZero(p.stage1AdjustedSeconds),

    p.stage2Duration || '00:00:00',
    numberOrZero(p.stage2DurationSeconds),
    numberOrZero(p.stage2Penalties),
    numberOrZero(p.stage2PenaltySeconds),
    p.stage2AdjustedTime || '00:00:00',
    numberOrZero(p.stage2AdjustedSeconds),

    p.stage3Duration || '00:00:00',
    numberOrZero(p.stage3DurationSeconds),
    p.stage3Overtime || '00:00:00',
    numberOrZero(p.stage3OvertimeSeconds),
    numberOrZero(p.stage3Penalties),
    numberOrZero(p.stage3PenaltySeconds),

    numberOrZero(p.totalPenaltyCount),
    p.penaltyTime || '00:00:00',
    numberOrZero(p.totalPenaltySeconds),

    p.adjustedTime || '00:00:00',
    numberOrZero(p.adjustedSeconds),
    p.displayedResult || '',
    new Date()
  ];
}

function toDate(value) {
  if (!value) return '';

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : date;
}

function numberOrZero(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function formatResultRow(sheet, rowNumber) {
  [4, 5, 6, 7, 34].forEach(function(column) {
    sheet.getRange(rowNumber, column)
      .setNumberFormat('dd/MM/yyyy hh:mm:ss AM/PM');
  });

  // Keep duration strings as text, not spreadsheet clock values.
  [8, 10, 14, 16, 20, 22, 24, 29, 31].forEach(function(column) {
    sheet.getRange(rowNumber, column).setNumberFormat('@');
  });
}

function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
