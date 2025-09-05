function query(dataset, inputName, inputDate) {
  const {startOfWeek, endOfWeek} = getWeekRange(inputDate)
  const filterData = dataset.filter(row => 
    row.name === inputName &&
    row.date >= startOfWeek &&
    row.date <= endOfWeek);
  
  return filterData;
}
