function timeToDecimal(time){
  return time.getHours() + time.getMinutes()/60
}

function getDurationHours(start, end){
  const startDecimal = timeToDecimal(start)
  let endDecimal = timeToDecimal(end)
  if (endDecimal === 0 && end < start){
    // Handle time to the midnight.
    // By rules, Time across the midnight must be record under next day. So only 0:00 is allowed as end time.
    endDecimal = 24;
  }
  return endDecimal - startDecimal;
}

function getWeekRange(date) {
  // Date is mutable
  const inputDate = new Date(date);
  // Get the day of the week(0 = Sunday, 1 = Monday,...)
  const day = inputDate.getDay();
  const diffToMonday = (day === 0 ? -6 : 1 - day);

  const startOfWeek = new Date(inputDate);
  startOfWeek.setDate(inputDate.getDate() + diffToMonday);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  
  return {
    startOfWeek,
    endOfWeek
  };
}

function roundToTwo(num) {
  // Helper function for more accurate round up due to the float calculation issue.
  return Math.round((num + Number.EPSILON) * 100) / 100;
}
