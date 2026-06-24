// ============================================================
// USERS API
// ============================================================

function getUsers() {
  return sheetToObjects(getSheet('Users'));
}

function createUser(data) {
  const sheet = getSheet('Users');
  const id = genId();
  const now = new Date().toISOString();
  sheet.appendRow([id, data.name, data.role || 'employee', data.email || '', now]);
  return { id, ...data, created_at: now };
}

function updateUser(id, data) {
  const sheet = getSheet('Users');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);
  Object.keys(data).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });
  return { id, ...data };
}

function deleteUser(id) {
  const sheet = getSheet('Users');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}
