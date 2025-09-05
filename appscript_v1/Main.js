let CONFIG;

function main() {
  // Load Global Config
  CONFIG = getConfig();
  if(!CONFIG) Logger.log("Get Config failed");
  Logger.log(CONFIG);

  const data = getValidatedShiftData()

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payslip");
  const inputName = sheet.getRange("A1").getValue();
  const inputDate = sheet.getRange("B1").getValue();
  const {startOfWeek, endOfWeek} = getWeekRange(inputDate)
  const filterData = data.filter(row => 
    row.name === inputName &&
    row.date >= startOfWeek &&
    row.date <= endOfWeek);

  const groupedShiftMap = sortAndGroupByDate(filterData);
  const summary = parseShift(groupedShiftMap);
  logOutMap(groupedShiftMap);

  const sortedSummaryKey = Object.keys(summary).sort((a,b) => summary[a].wage - summary[b].wage);
  const weeklyTotal = {
    hours: 0,
    total: 0
  }
  sortedSummaryKey.forEach((key) => {
    const item = summary[key];
    weeklyTotal.hours += item.hours;
    const total = roundToTwo(item.wage * item.hours);
    weeklyTotal.total += total;
    // Logger.log(`${key}: ${roundToTwo(item.hours)} hours, wage in ${roundToTwo(item.wage)} dollars/h, total ${total} dollars`);
  });
  Logger.log(summary);
  Logger.log(weeklyTotal);
  Logger.log(`${inputName} worked for ${weeklyTotal.hours} hours, with salary in total ${roundToTwo(weeklyTotal.total)} dollars, during the week from ${startOfWeek.toDateString()} to ${endOfWeek.toDateString()}`);
  
  renderPaySlip(inputName, startOfWeek, endOfWeek, groupedShiftMap,summary,weeklyTotal);
}
