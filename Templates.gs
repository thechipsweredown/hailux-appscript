// ============================================================
// TASK TEMPLATES API
// ============================================================

function getTaskTemplates() {
  return sheetToObjects(getSheet('TaskTemplates'));
}

function createTaskTemplate(data) {
  const sheet = getSheet('TaskTemplates');
  const id = genId();
  sheet.appendRow([id, data.name, data.description || '']);
  return { id, ...data };
}

function deleteTaskTemplate(id) {
  const sheet = getSheet('TaskTemplates');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}
