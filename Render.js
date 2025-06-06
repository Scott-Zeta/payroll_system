function renderPaySlip(name, startOfWeek, endOfWeek, parsedShiftData, summary, weeklyTotal) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payslip");
  
  // Clean the all content from row 2 to last row
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  // Payslip Header rows
  sheet.getRange(2,1).setValue("Weekly Payslip");
  sheet.getRange(3,1).setValue(`Name: ${name}`);
  sheet.getRange(4,1).setValue(`Week: ${formatDate(startOfWeek)} to ${formatDate(endOfWeek)}`)

  // Shift log Header rows
  sheet.getRange(6,1).setValue("Shift Logs")
  headers = ["Date", "Day", "Start","End","Break", "Total", "Regular Hours", "Early Overtime", "Late Overtime"]
  CONFIG.OT_DAILY_TIME_THRESHOLD.forEach(
    (value) => headers.push(`Daily OT${value}`)
  );
  CONFIG.OT_WEEKLY_TIME_THRESHOLD.forEach(
    (value) => headers.push(`Weekly OT${value}`)
  );
  sheet.getRange(7,1,1,headers.length).setValues([headers]);

  const dayPrefix = { 6: "SAT", 0: "SUN" };
  const dayArr = ["Sun", "Mon", "Tue", "Wed", "Thr", "Fri", "Sat"];
  let rowIndex = 8;
  parsedShiftData.forEach((value,key) => 
    value.forEach((record) =>{
      const row = [formatDate(record.date), dayArr[`${record.date.getDay()}`],formatTime(record.start),formatTime(record.finish), record.break, getDurationHours(record.start, record.finish) - record.break];
      
      const parsedShift = record.parsedShift
      // Non-OT hours
      const prefix = dayPrefix[`${record.date.getDay()}`] || "WD";
      const nonOTHoursArr = [parsedShift[`${prefix}_Regular`]?.hours || "", parsedShift[`${prefix}_Early_OT`]?.hours || "", parsedShift[`${prefix}_Late_OT`]?.hours || ""];
      row.push(...nonOTHoursArr);
      // Daily OT hours
      const OTHoursArr = [];
      CONFIG.OT_DAILY_TIME_THRESHOLD.forEach(
        (value) => {
          const hours = parsedShift[`Daily_OT_${value}`]?.hours || "";
          OTHoursArr.push(hours)
        }
      )
      CONFIG.OT_WEEKLY_TIME_THRESHOLD.forEach(
        (value) => {
          const hours = parsedShift[`Weekly_OT_${value}`]?.hours || "";
          OTHoursArr.push(hours)
        }
      )
      row.push(...OTHoursArr);
      // Render the shift log
      Logger.log(row)
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
      rowIndex++;
    }));

    //Render Summary rows
    rowIndex++;
    sheet.getRange(rowIndex, 1).setValue("Summary")
    rowIndex++;
    // Render the Summary in wage's order
    const sortedSummaryKey = Object.keys(summary).sort((a,b) => summary[a].wage - summary[b].wage);
    sortedSummaryKey.forEach((key) => {
      if (roundToTwo(summary[key].total) !== 0) {
        row = [`${key}: ${roundToTwo(summary[key].hours)} hours`, `at ${roundToTwo(summary[key].wage)} dollars`, `$ ${roundToTwo(summary[key].total)}`];
        sheet.getRange(rowIndex,1,1,row.length).setValues([row]);
        rowIndex++;
      }
    });

    const totalRow = ["Total","",`$ ${roundToTwo(weeklyTotal.total)}`];
    sheet.getRange(rowIndex,1,1,totalRow.length).setValues([totalRow]);
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "h:mm a");
}