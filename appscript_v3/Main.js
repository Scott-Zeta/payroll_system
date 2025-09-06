let CONFIG;
let TIME_CONFIG;
let WAGE_CONFIG;

function main() {
  // Load Global Config
  CONFIG = getConfig('Config');
  TIME_CONFIG = getConfig('Time Config');
  WAGE_CONFIG = getConfig('Wage Config');
  if (!CONFIG) Logger.log('Get Config failed');
  Logger.log(CONFIG);
  Logger.log(TIME_CONFIG);
  Logger.log(WAGE_CONFIG);

  const data = getValidatedShiftData();
  Logger.log(data);
  const groupedData = groupByName(data);
  Logger.log(groupedData);

  const summaryList = {};
  for (const [name, shifts] of Object.entries(groupedData)) {
    // Parse working hours
    const groupedShiftMap = sortAndGroupByDate(shifts);
    const summary = parseShift(groupedShiftMap);

    // Calculate the Total
    const sortedSummaryKey = Object.keys(summary).sort(
      (a, b) => summary[a].wage - summary[b].wage
    );
    const weeklyTotal = {
      hours: 0,
      total: 0,
    };
    sortedSummaryKey.forEach((key) => {
      const item = summary[key];
      weeklyTotal.hours += item.hours;
      const total = roundToTwo(item.wage * item.hours);
      weeklyTotal.total += total;
      // Logger.log(`${key}: ${roundToTwo(item.hours)} hours, wage in ${roundToTwo(item.wage)} dollars/h, total ${total} dollars`);
    });
    summary['Weekly_Total'] = weeklyTotal;
    summaryList[name] = summary;
  }

  Logger.log(summaryList);

  /*
  old version, deprecate later
  */
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payslip');
  const inputName = sheet.getRange('A1').getValue();
  const inputDate = sheet.getRange('B1').getValue();
  const { startOfWeek, endOfWeek } = getWeekRange(inputDate);

  const filterData = query(data, inputName, inputDate);
  // Logger.log(filterData)

  const groupedShiftMap = sortAndGroupByDate(filterData);
  const summary = parseShift(groupedShiftMap);
  logOutMap(groupedShiftMap);

  const sortedSummaryKey = Object.keys(summary).sort(
    (a, b) => summary[a].wage - summary[b].wage
  );
  const weeklyTotal = {
    hours: 0,
    total: 0,
  };
  sortedSummaryKey.forEach((key) => {
    const item = summary[key];
    weeklyTotal.hours += item.hours;
    const total = roundToTwo(item.wage * item.hours);
    weeklyTotal.total += total;
    // Logger.log(`${key}: ${roundToTwo(item.hours)} hours, wage in ${roundToTwo(item.wage)} dollars/h, total ${total} dollars`);
  });
  Logger.log(summary);
  Logger.log(weeklyTotal);
  Logger.log(
    `${inputName} worked for ${
      weeklyTotal.hours
    } hours, with salary in total ${roundToTwo(
      weeklyTotal.total
    )} dollars, during the week from ${startOfWeek.toDateString()} to ${endOfWeek.toDateString()}`
  );

  renderPaySlip(
    inputName,
    startOfWeek,
    endOfWeek,
    groupedShiftMap,
    summary,
    weeklyTotal
  );
}
