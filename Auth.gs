// ============================================================
// AUTH
// ============================================================

function verifyManagerPassword(password) {
  const settings = sheetToObjects(getSheet('Settings'));
  const entry = settings.find(s => s.key === 'manager_password');
  return entry && String(entry.value) === String(password);
}
