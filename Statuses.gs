// ============================================================
// STATUSES API
// ============================================================

function getStatuses() {
  return sheetToObjects(getSheet('Statuses'));
}

function createStatus(data) {
  const sheet = getSheet('Statuses');
  const id = genId();
  sheet.appendRow([id, data.entity_type, data.label, data.color, data.order]);
  return { id, ...data };
}

function updateStatus(id, data) {
  const sheet = getSheet('Statuses');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);
  Object.keys(data).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });
  return { id, ...data };
}

function deleteStatus(id) {
  const sheet = getSheet('Statuses');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}

function getDefaultStatusId(entityType) {
  const statuses = getStatuses();
  const s = statuses.filter(s => s.entity_type === entityType).sort((a, b) => a.order - b.order)[0];
  return s ? s.id : '';
}

function _ensureJobStatus(label, color, order) {
  var sheet = getSheet('Statuses');
  var existing = sheetToObjects(sheet).find(function(s) {
    return s.entity_type === 'job' && s.label === label;
  });
  if (existing) return existing.id;
  var id = genId();
  sheet.appendRow([id, 'job', label, color, order]);
  return id;
}
