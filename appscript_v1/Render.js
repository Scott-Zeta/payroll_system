function renderPaySlip(
  name,
  startOfWeek,
  endOfWeek,
  parsedShiftData,
  summary,
  weeklyTotal
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payslip');

  // Clear content after header (Row 1)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // --- Payslip Header ---
  sheet.getRange(2, 1).setValue('Weekly Payslip');
  sheet.getRange(3, 1).setValue(`Name: ${name}`);
  sheet
    .getRange(4, 1)
    .setValue(`Week: ${formatDate(startOfWeek)} to ${formatDate(endOfWeek)}`);

  // --- Table Headers ---
  const baseHeaders = ['Date', 'Day', 'Start', 'End', 'Break', 'Total'];
  const allWageTypes = collectAllWageTypes(parsedShiftData, summary);
  const headers = baseHeaders.concat(allWageTypes);
  sheet.getRange(6, 1).setValue('Shift Logs');
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]);

  // --- Render Shift Logs ---
  let rowIndex = 8;
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

      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      rowIndex++;
    });
  });

  // --- Summary Block ---
  rowIndex += 2;
  sheet.getRange(rowIndex, 1).setValue('Summary');
  rowIndex++;

  const sortedKeys = Object.keys(summary).sort(
    (a, b) => summary[a].wage - summary[b].wage
  );
  sortedKeys.forEach((key) => {
    const item = summary[key];
    if (roundToTwo(item.total) > 0) {
      const line = [
        `${key}: ${roundToTwo(item.hours)} hours`,
        `at $${roundToTwo(item.wage)}`,
        `$${roundToTwo(item.total)}`,
      ];
      sheet.getRange(rowIndex, 1, 1, line.length).setValues([line]);
      rowIndex++;
    }
  });

  // --- Total Pay ---
  sheet
    .getRange(rowIndex, 1, 1, 3)
    .setValues([['Total', '', `$${roundToTwo(weeklyTotal.total)}`]]);
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

function getDayName(date) {
  return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][date.getDay()];
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'h:mm a');
}
