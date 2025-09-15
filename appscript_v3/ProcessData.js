function groupByName(data) {
  const groupedData = data.reduce((acc, item) => {
    if (item.name) {
      const key = item.name;
      if (!acc[key]) {
        acc[key] = [];
      }
      acc[key].push(item);
    }
    return acc;
  }, {});
  return groupedData;
}

function sortAndGroupByDate(data) {
  // Sort the record by Date, then by Start Time
  data.sort((a, b) => {
    if (a.date - b.date !== 0) {
      return a.date - b.date;
    }
    return a.start - b.start;
  });

  // Use Map to group sorted data by date to keep the time sequence
  const groupedMap = new Map();
  data.forEach((entry) => {
    // Use the day in week key (1 = Monday, ...., 7 = Sunday)
    const key = entry.date.getDay() === 0 ? 7 : entry.date.getDay();
    if (!groupedMap.has(key)) {
      groupedMap.set(key, []);
    }
    groupedMap.get(key).push(entry);
  });

  return groupedMap;
}

function parseShift(groupedShiftMap) {
  const summary = {};

  let weeklyTotal = 0;
  groupedShiftMap.forEach((value, key) => {
    // Logger.log(`Key: ${key}`);
    let dailyTotal = 0;
    value.forEach((record) => {
      // Logger.log(record);

      let parseResult = {};
      let breakTime = Math.max(record.break, 0);
      let duration = getDurationHours(record.start, record.finish) - breakTime;
      weeklyTotal += duration;
      dailyTotal += duration;
      // Logger.log("WeeklyTotal:" + weeklyTotal);
      // Logger.log("DailyTotal:" + dailyTotal);

      // First Parse WeeklyOvertime
      const { hoursRemain: weeklyRemain, overtimeResult: weeklyOTResult } =
        parseOvertime(
          'Weekly',
          weeklyTotal,
          duration,
          TIME_CONFIG.OT_WEEKLY_TIME_THRESHOLD,
          WAGE_CONFIG.OT_WEEKLY_THRESHOLD_WAGE
        );
      parseResult = { ...parseResult, ...weeklyOTResult };
      // If there is time remain, parse for if there is daily overtime
      // Logger.log(
      //   'Key: ' + key + 'Time Remain after parsing weeklyOT: ' + weeklyRemain
      // );
      if (weeklyRemain > 0) {
        const { hoursRemain: dailyRemain, overtimeResult: dailyOTResult } =
          parseOvertime(
            'Daily',
            weeklyRemain,
            weeklyRemain,
            TIME_CONFIG.OT_DAILY_TIME_THRESHOLD,
            WAGE_CONFIG.OT_DAILY_THRESHOLD_WAGE
          );
        parseResult = { ...parseResult, ...dailyOTResult };

        // If there are still hours remain, parse them as regular working hours based on opening and closing time
        if (dailyRemain > 0) {
          const regularWorkResult = parseRegularWork(
            record.date,
            record.start,
            dailyRemain,
            record.break
          );
          parseResult = { ...parseResult, ...regularWorkResult };
        }
      }
      record['parsedShift'] = parseResult;

      // caculate the summary
      for (const key in parseResult) {
        if (!summary[key])
          summary[key] = {
            wage: parseResult[key].wage,
            hours: 0,
            total: 0,
          };
        summary[key].hours += parseResult[key].hours;
        summary[key].total += parseResult[key].total;
      }
    });
  });

  return summary;
}

function parseRegularWork(date, shiftStart, remainHours, breakHours) {
  /* By rules, breakHours shall be deducted from working period with lowest wage
    Need to add the break time temporarily to calculate Early/Late overtime.
  */
  const duration = remainHours + breakHours;
  const shiftStartDecimal = timeToDecimal(shiftStart);
  const shiftEndDecimal = (shiftStartDecimal + duration) % 24;
  const openDecimal = timeToDecimal(TIME_CONFIG.OPEN_TIME);
  const closeDecimal =
    timeToDecimal(TIME_CONFIG.CLOSE_TIME) === 0
      ? 24
      : timeToDecimal(TIME_CONFIG.CLOSE_TIME);
  // Split the regular time into blocks by day(if the shift crosses midnight)
  const timeBlocks = splitRegularTimeByDay(
    date,
    shiftStartDecimal,
    shiftEndDecimal
  );
  const allWorkSegments = [];

  timeBlocks.forEach((block) => {
    // Seperate Early overtime, Late overtime, Work in opening hours
    const blockDuration = block.end - block.start;

    const earlyOT =
      Math.max(openDecimal - block.start, 0) -
      Math.max(openDecimal - block.end, 0);
    const lateOT =
      Math.max(block.end - closeDecimal, 0) -
      Math.max(block.start - closeDecimal, 0);
    const workInOpening = blockDuration - earlyOT - lateOT;

    const dayKey = block.date.getDay() === 0 ? 7 : block.date.getDay();
    const prefixMap = { 6: 'SAT', 7: 'SUN' };
    const prefix = prefixMap[dayKey] || 'WD';

    /* As new requirements, no need to split working hours in weekends into Early_OT/Late_OT/Regular */
    const isWeekend = dayKey === 6 || dayKey === 7;
    if (isWeekend && TIME_CONFIG.BOOL_WEEKEND_OT === false) {
      allWorkSegments.push({
        name: `${prefix}_Regular`,
        wage: WAGE_CONFIG[`${prefix}_BASE_WAGE`],
        hours: blockDuration,
      });
      return;
    }

    if (workInOpening > 0) {
      allWorkSegments.push({
        name: `${prefix}_Regular`,
        wage: WAGE_CONFIG[`${prefix}_BASE_WAGE`],
        hours: workInOpening,
      });
    }
    if (earlyOT > 0) {
      allWorkSegments.push({
        name: `${prefix}_Early_OT`,
        wage: WAGE_CONFIG[`${prefix}_EARLY_OT_WAGE`],
        hours: earlyOT,
      });
    }
    if (lateOT > 0) {
      allWorkSegments.push({
        name: `${prefix}_Late_OT`,
        wage: WAGE_CONFIG[`${prefix}_LATE_OT_WAGE`],
        hours: lateOT,
      });
    }
  });

  // Sort by the wage, deduct break time from lowest wage periods first
  allWorkSegments.sort((a, b) => a.wage - b.wage);
  let reamainingBreak = breakHours;

  for (let segment of allWorkSegments) {
    if (reamainingBreak <= 0) break;
    const diff = segment.hours - reamainingBreak;
    segment.hours = Math.max(diff, 0);
    reamainingBreak = -diff;
  }

  // Assemble result in object format
  const result = {};
  for (const segment of allWorkSegments) {
    if (!result[segment.name]) {
      result[segment.name] = {
        wage: segment.wage,
        hours: 0,
        total: 0,
      };
    }
    result[segment.name].hours += segment.hours;
    result[segment.name].total += segment.hours * segment.wage;
  }

  return result;
}

function parseOvertime(
  prefix,
  totalHours,
  duration,
  thresholds,
  thresholds_wage
) {
  const splitArray = new Array(thresholds.length + 1).fill(0);
  let start = totalHours - duration;
  let remaining = duration;
  // Logger.log(remaining);
  for (let i = 0; i <= thresholds.length; i++) {
    const end = thresholds[i] ?? Infinity;

    // Determine how much of the shift fits into this range
    const rangeStart = Math.max(start, i === 0 ? 0 : thresholds[i - 1]);
    const rangeEnd = Math.min(end, start + remaining);
    const hoursInRange = Math.max(0, rangeEnd - rangeStart);

    splitArray[i] = hoursInRange;
    remaining -= hoursInRange;
    start += hoursInRange;

    if (remaining <= 0) break;
  }

  const hoursRemain = splitArray[0];
  const overtimeArray = splitArray.slice(1);
  // Logger.log(`${prefix}: ` + splitArray);
  const overtimeResult = {};
  thresholds.forEach((threshold, index) => {
    overtimeResult[`${prefix}_OT_${threshold}`] = {
      wage: thresholds_wage[index],
      hours: overtimeArray[index],
      total: overtimeArray[index] * thresholds_wage[index],
    };
  });
  // Logger.log('Hours Remain for next stage: ' + hoursRemain);
  return { hoursRemain, overtimeResult };
}

function splitRegularTimeByDay(date, startDecimal, endDecimal) {
  const blocks = [];

  if (endDecimal <= startDecimal) {
    // Handle Shift Crosses midnight by sepearte into two blocks
    blocks.push({ date: date, start: startDecimal, end: 24 });

    const nextDate = new Date(date);
    nextDate.setDate(date.getDate() + 1);
    blocks.push({ date: nextDate, start: 0, end: endDecimal });
  } else {
    blocks.push({ date: date, start: startDecimal, end: endDecimal });
  }

  return blocks;
}
