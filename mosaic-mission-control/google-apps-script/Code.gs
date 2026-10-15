```javascript
/**
 * MOSAIC 2026 — Mission Results Recorder
 *
 * Receives completed mission results from the React application
 * and records them in the "Mission Results" spreadsheet tab.
 *
 * Deployment:
 * 1. Deploy as a Web App.
 * 2. Execute as the spreadsheet owner.
 * 3. Set access to Anyone, if available for your account.
 * 4. Put the deployed /exec URL in VITE_GOOGLE_APPS_SCRIPT_URL.
 *
 * IMPORTANT:
 * Existing rows are preserved. Headers are updated to the current
 * schema. Existing records are not automatically recalculated.
 */

// Use the spreadsheet ID, NOT the full Google Sheets URL.
const SPREADSHEET_ID =
  '1EDjRRfjk7X2kZKeccGe4dKGB0zQB7acFMNsVAGvL29c';

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


/**
 * GET endpoint for checking whether the web app is reachable.
 * Open the deployed /exec URL in a browser to test it.
 */
function doGet() {
  return jsonOutput({
    ok: true,
    service: 'MOSAIC 2026 Mission Results Recorder',
    spreadsheet: SHEET_NAME,
    message: 'The Apps Script web app is reachable.'
  });
}


/**
 * Receives the completed mission result from the frontend.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const body =
      e && e.postData && e.postData.contents
        ? e.postData.contents
        : '{}';

    const payload = JSON.parse(body);

    validatePayload(payload);

    const spreadsheet =
      SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheet = getOrCreateResultsSheet(spreadsheet);

    const rowValues = toRow(payload);
    const sessionId = String(payload.sessionId);

    // Update an existing row if this session was already recorded.
    // Otherwise, append a new result row.
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const existingIds = sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues()
        .flat();

      const existingIndex = existingIds.indexOf(sessionId);

      if (existingIndex !== -1) {
        const rowNumber = existingIndex + 2;

        sheet
          .getRange(rowNumber, 1, 1, HEADERS.length)
          .setValues([rowValues]);

        formatResultRow(sheet, rowNumber);

        SpreadsheetApp.flush();

        return jsonOutput({
          ok: true,
          updated: true,
          sessionId: sessionId
        });
      }
    }

    sheet.appendRow(rowValues);

    const newRow = sheet.getLastRow();

    formatResultRow(sheet, newRow);

    SpreadsheetApp.flush();

    return jsonOutput({
      ok: true,
      appended: true,
      sessionId: sessionId
    });

  } catch (error) {
    return jsonOutput({
      ok: false,
      error: String(
        error && error.message
          ? error.message
          : error
      )
    });

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


/**
 * Checks that the essential result fields are present.
 */
function validatePayload(payload) {
  if (!payload.sessionId) {
    throw new Error('Missing sessionId.');
  }

  if (!payload.teamName) {
    throw new Error('Missing teamName.');
  }

  if (!payload.completedAt) {
    throw new Error('Missing completedAt. Only completed missions can be recorded.');
  }

  if (!payload.startedAt) {
    throw new Error('Missing startedAt.');
  }
}


/**
 * Creates the result tab if needed and synchronizes its headers.
 *
 * Existing rows are preserved. If the tab already exists, the
 * current header row is replaced with the new schema.
 */
function getOrCreateResultsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      HEADERS.length - sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS]);

  sheet.setFrozenRows(1);

  sheet
    .getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setWrap(true);

  return sheet;
}


/**
 * Converts the incoming payload into the spreadsheet column order.
 */
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


/**
 * Converts an ISO timestamp into a Date object so Sheets can
 * display it using a real date/time format.
 */
function toDate(value) {
  if (!value) {
    return '';
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return '';
  }

  return date;
}


/**
 * Prevents missing numeric values from becoming NaN.
 */
function numberOrZero(value) {
  const number = Number(value);

  return Number.isFinite(number) ? number : 0;
}


/**
 * Formats the timestamp columns in 12-hour format with AM/PM.
 *
 * The spreadsheet stores actual dates and times, not just strings.
 * The spreadsheet's timezone controls the displayed local time.
 */
function formatResultRow(sheet, rowNumber) {
  const timestampColumns = [
    4,  // Game Started At
    5,  // Stage 1 Completed At
    6,  // Stage 2 Completed At
    7,  // Game Ended At
    34  // Recorded At
  ];

  timestampColumns.forEach(function(column) {
    sheet
      .getRange(rowNumber, column)
      .setNumberFormat('dd/MM/yyyy hh:mm:ss AM/PM');
  });

  // Keep all duration values as text such as 00:08:42.
  // This avoids Sheets interpreting a duration as a clock time.
  const durationColumns = [
    8,  // Overall Raw Time
    11, // Stage 1 Raw Time
    14, // Stage 1 Adjusted Time
    17, // Stage 2 Raw Time
    20, // Stage 2 Adjusted Time
    23, // Stage 3 Raw Time
    24, // Stage 3 Overtime
    28, // Total Penalty Time
    30  // Adjusted Final Time
  ];

  durationColumns.forEach(function(column) {
    sheet
      .getRange(rowNumber, column)
      .setNumberFormat('@');
  });
}


/**
 * Returns a JSON response.
 */
function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
```
