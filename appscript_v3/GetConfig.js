function getConfig(configSheetName) {
  const { headers, rows } = readSheet(configSheetName);
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
