function renderPaySlipList(summaryList) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payslip List');
  cleanContent(sheet);

  rc = createRowCursor(2);
  const { startOfWeek, endOfWeek } = getWeekRange(
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('Shift Entry')
      .getRange('A2')
      .getValue()
  );

  for (const [key, value] of Object.entries(summaryList)) {
    renderPaySlip(
      key,
      startOfWeek,
      endOfWeek,
      value.shiftLog,
      value.summary,
      value.weeklyTotal,
      sheet,
      rc
    );
  }
}

function renderPaySlip(
  name,
  startOfWeek,
  endOfWeek,
  parsedShiftData,
  summary,
  weeklyTotal,
  sheet,
  rc = createRowCursor(2)
) {
  // --- Payslip Header ---
  const titleRow = rc.peek();
  writeRow(sheet, rc, ['Weekly Payslip'], 3);
  const metaRow = rc.peek();
  writeRow(sheet, rc, [
    `Name:`,
    `${name}`,
    `Week:`,
    `${formatDate(startOfWeek)} to ${formatDate(endOfWeek)}`,
  ]);
  rc.skip(1);

  // --- Table Headers ---
  const baseHeaders = ['Date', 'Day', 'Start', 'End', 'Break', 'Total'];
  const allWageTypes = collectAllWageTypes(parsedShiftData, summary);
  const headers = baseHeaders.concat(allWageTypes);

  const shiftLogTitleRow = rc.peek();
  writeRow(sheet, rc, ['Shift Logs']);
  const headerRow = rc.peek();
  writeRow(sheet, rc, headers);

  // --- Render Shift Logs ---
  parsedShiftData.forEach((records, key) => {
    records.forEach((record) => {
      const row = [
        formatDate(record.date),
        getDayName(record.date),
        formatTime(record.start),
        formatTime(record.finish),
        record.break,
        getDurationHours(record.start, record.finish) - record.break,
      ];

      // Fill wage categories dynamically
      allWageTypes.forEach((type) => {
        const entry = record.parsedShift?.[type];
        row.push(
          entry
            ? roundToTwo(entry.hours) === 0
              ? ''
              : roundToTwo(entry.hours)
            : ''
        );
      });

      writeRow(sheet, rc, row);
    });
  });
  rc.skip(1);

  // --- Summary Block ---
  const summaryTitleRow = rc.peek();
  writeRow(sheet, rc, ['Summary']);

  const sortedKeys = Object.keys(summary).sort(
    (a, b) => summary[a].wage - summary[b].wage
  );
  sortedKeys.forEach((key) => {
    const item = summary[key];
    if (roundToTwo(item.total) > 0) {
      let row = [];
      if (WAGE_CONFIG.BOOL_SHOW_WAGE === true) {
        row = [
          `${key}:`,
          `${roundToTwo(item.hours)} hours`,
          `at $${roundToTwo(item.wage)}`,
          `$${roundToTwo(item.total)}`,
        ];
      } else {
        row = [`${key}:`, `${roundToTwo(item.hours)} hours`];
      }
      writeRow(sheet, rc, row);
    }
  });

  // --- Total Pay ---
  const totalRow = rc.peek();
  if (WAGE_CONFIG.BOOL_SHOW_WAGE === true) {
    writeRow(sheet, rc, [
      'Total',
      `${roundToTwo(weeklyTotal.hours)} hours`,
      '',
      `$${roundToTwo(weeklyTotal.total)}`,
    ]);
  } else {
    writeRow(sheet, rc, ['Total', `${roundToTwo(weeklyTotal.hours)} hours`]);
  }
  rc.skip(2);

  // --- Format ---
  boldRow(sheet, titleRow, 3);
  boldRow(sheet, metaRow, 4);
  boldRow(sheet, shiftLogTitleRow, 1);
  boldRow(sheet, headerRow, 6 + allWageTypes.length);
  boldRow(sheet, summaryTitleRow, 1);
  boldRow(sheet, totalRow, 4);
  border(sheet, titleRow, 1, totalRow, 6 + allWageTypes.length);
}

function collectAllWageTypes(parsedShiftData, summary) {
  const typesSet = new Set();

  parsedShiftData.forEach((records) => {
    records.forEach((record) => {
      if (!record.parsedShift) return;
      Object.keys(record.parsedShift).forEach((key) => {
        if (record.parsedShift[key]?.hours > 0) typesSet.add(key);
      });
    });
  });

  const typesArray = Array.from(typesSet);

  // Sort by actual wage value from summary
  typesArray.sort((a, b) => {
    const wageA = summary[a]?.wage ?? 0;
    const wageB = summary[b]?.wage ?? 0;
    return wageA - wageB;
  });

  return typesArray;
}

/* Helper function */
function cleanContent(sheet) {
  // Clear content after header (Row 1)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clear();
}

/* Cursor function for tracking the row index */
function createRowCursor(start = 2) {
  let row = start;
  return {
    peek: () => row, // read current row (no move)
    next: (n = 1) => {
      // return current row, then advance
      const r = row;
      row += n;
      return r;
    },
    skip: (n = 1) => {
      row += n;
    }, // move down n rows
    set: (r) => {
      row = r;
    }, // jump to a specific row
  };
}

/**
 * Helper to write a single row (1D array).
 */
function writeRow(sheet, rc, values, col = 1) {
  sheet.getRange(rc.next(), col, 1, values.length).setValues([values]);
}

/**
 * Helper to write a block (2D array). Returns rows written.
 */
function writeBlock(sheet, startRow, startCol, rows2D) {
  if (!rows2D || !rows2D.length) return 0;
  sheet
    .getRange(startRow, startCol, rows2D.length, rows2D[0].length)
    .setValues(rows2D);
  return rows2D.length;
}

function getDayName(date) {
  return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][date.getDay()];
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'h:mm a');
}

/*
Format Helper Functions
*/
function boldRow(sheet, row, cols = 12) {
  sheet.getRange(row, 1, 1, cols).setFontWeight('bold');
}

function border(sheet, r1, c1 = 1, r2, c2 = 12) {
  const rng = sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  // outer frame
  rng.setBorder(
    true,
    null,
    true,
    null,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  return rng;
}
