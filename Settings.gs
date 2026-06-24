// ============================================================
// SETTINGS API
// ============================================================

function getSettings() {
  return sheetToObjects(getSheet('Settings'));
}

function updateSetting(key, value) {
  const sheet = getSheet('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  sheet.appendRow([key, value]);
  return true;
}
