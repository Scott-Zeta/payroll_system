function readSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return { headers, rows };
}

function getConfig() {
  const { headers, rows } = readSheet('Config');
  const config = {};
  const errors = [];

  headers.forEach((header, index) => {
    let value = rows[0][index];
    if (typeof value === 'string') value = value.trim();
    try {
      switch (true) {
        case header === 'OPEN_TIME':
          if (!(value instanceof Date)) throw `Invalid Time`;
          config[header] = value;
          break;
        case header === 'CLOSE_TIME':
          if (!(value instanceof Date)) throw `Invalid Time`;
          config[header] = value;
          break;
        case header.includes('THRESHOLD'):
          value = value.toString().split(',').map(Number);
          config[header] = value;
          break;
        case header.includes('WAGE'):
          if (isNaN(value) || value < 0) throw `Invalid threshold value`;
          config[header] = value;
          break;
      }
    } catch (e) {
      errors.push(`Error on column ${header}: ${e}`);
    }
  });

  if (errors.length) {
    raiseErrors(
      'There are some critical Config Errors, calculation can not procced:\n\n',
      errors
    );
    return {};
  }
  return config;
}

function getValidatedShiftData() {
  const { headers, rows } = readSheet('Shift Entry');
  const errors = []; //Error found in validation
  const result = []; //Data passed validation

  rows.forEach((row, i) => {
    const rowIndex = i + 2; //Row number in sheet
    const rowObject = {};
    let hasError = false;

    headers.forEach((header, index) => {
      let value = row[index];
      if (typeof value === 'string') value = value.trim();

      try {
        switch (header) {
          case 'Date':
            if (!(value instanceof Date)) throw `Invalid Date`;
            rowObject.date = value;
            break;

          case 'Name':
            if (!value) throw `Name is Missing`;
            rowObject.name = value.toString();
            break;

          case 'Start Time':
            if (!(value instanceof Date)) throw `Invalid Time`;
            rowObject.start = value;
            break;

          case 'Finish Time':
            if (!(value instanceof Date)) throw `Invalid Time`;
            rowObject.finish = value;
            break;

          case 'Break(Hours)':
            if (value === '' || value === null || value === undefined) {
              value = 0;
            }
            value = parseFloat(value);
            if (isNaN(value) || value < 0) throw `Invalid Break Value`;
            if (value > getDurationHours(rowObject.start, rowObject.finish))
              throw `Break Time is larger than Shift Hours`;
            rowObject.break = value;
            break;
        }
      } catch (e) {
        hasError = true;
        errors.push(`Errors on Row ${rowIndex}, Column "${header}": ${e}`);
      }
    });

    if (!hasError) result.push(rowObject);
  });

  if (errors.length) {
    raiseErrors(
      'Some records have validation errors and were excluded from the calculation:\n\n',
      errors
    );
    return result;
  }

  return result;
}
