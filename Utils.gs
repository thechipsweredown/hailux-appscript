// ============================================================
// UTILITIES
// ============================================================

function genId() {
  return Utilities.getUuid();
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script chưa được gắn vào Google Sheet.');
  return ss.getSheetByName(name);
}

function requireSheet(name) {
  const sheet = getSheet(name);
  if (!sheet) throw new Error('Sheet "' + name + '" không tồn tại. Vui lòng chạy setupSheets() trước.');
  return sheet;
}

function sheetToObjects(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      var val = row[i];
      if (val instanceof Date) val = val.toISOString();
      obj[h] = val;
    });
    return obj;
  });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function normLabel(s) {
  return (s || '').toLowerCase().replace(/\s/g, '');
}

function findStatusByNorm(statuses, norm) {
  return statuses.find(s => normLabel(s.label) === norm) || null;
}
