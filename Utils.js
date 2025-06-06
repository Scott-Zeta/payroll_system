function raiseErrors(message, errors) {
  Logger.log(message + errors.join('\n'));
  try {
    SpreadsheetApp.getUi().alert(message + errors.join('\n'));
  } catch (e) {
    Logger.log('UI alert skipped (not in UI context).');
  }
}

function logOutMap(mapInput) {
  mapInput.forEach((value, key) => {
    Logger.log(`Key: ${key}`);

    if (Array.isArray(value)) {
      if (value.every((item) => typeof item === 'object')) {
        value.forEach((item) => Logger.log(item));
      } else {
        Logger.log(value);
      }
    } else {
      Logger.log(value);
    }
  });
}
