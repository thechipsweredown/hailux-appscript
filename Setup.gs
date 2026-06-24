// ============================================================
// SETUP & SEED
// ============================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = {
    Users:         ['id','name','role','email','created_at'],
    Jobs:          ['id','code','name','category','customer_name','customer_contact','received_date','deadline','revenue','repair_scope','status_id','notes','created_at','avatar_id'],
    Tasks:         ['id','job_id','name','order','assignee_id','deadline','status_id','completed_at','notes','created_at','evidence_id','emp_notes'],
    Statuses:      ['id','entity_type','label','color','order'],
    TaskTemplates: ['id','name','description'],
    Settings:      ['key','value'],
  };

  for (const [name, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      styleHeaderRow(sheet);
    }
  }

  seedDefaultData(ss);
  return 'Setup hoàn tất!';
}

function styleHeaderRow(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground('#0d1b4b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function seedDefaultData(ss) {
  const statusSheet = ss.getSheetByName('Statuses');
  if (statusSheet.getLastRow() <= 1) {
    const statuses = [
      [genId(), 'job',  'Chưa làm',   '#9E9E9E', 1],
      [genId(), 'job',  'Đang làm',   '#2196F3', 2],
      [genId(), 'job',  'Hoàn thành', '#4CAF50', 3],
      [genId(), 'job',  'Hủy',        '#F44336', 4],
      [genId(), 'task', 'Chưa làm',   '#9E9E9E', 1],
      [genId(), 'task', 'Đang làm',   '#FF9800', 2],
      [genId(), 'task', 'Hoàn thành', '#4CAF50', 3],
    ];
    statuses.forEach(r => statusSheet.appendRow(r));
  }

  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['manager_password', '1234']);
    settingsSheet.appendRow(['app_name', 'HaiLux']);
    settingsSheet.appendRow(['drive_folder_id', '']);
  }

  const tmplSheet = ss.getSheetByName('TaskTemplates');
  if (tmplSheet.getLastRow() <= 1) {
    const templates = ['Tách túi','Dựng form','Đi viền','Thay da','Lắp ráp vào','Vệ sinh','Mạ khóa','Dán seal','Phục hồi màu','Xử lý xước'];
    templates.forEach(name => tmplSheet.appendRow([genId(), name, '']));
  }
}

// ============================================================
// MAINTENANCE
// ============================================================

function migrateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const migrations = {
    'Jobs':  ['avatar_id'],
    'Tasks': ['evidence_id', 'emp_notes'],
  };

  Object.entries(migrations).forEach(([name, cols]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    cols.forEach(col => {
      if (!headers.includes(col)) {
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(col);
        sheet.getRange(1, nextCol).setBackground('#0d1b4b').setFontColor('#fff').setFontWeight('bold');
        Logger.log('Added column: ' + col + ' to ' + name);
      } else {
        Logger.log('Column already exists: ' + col + ' in ' + name);
      }
    });
  });

  return 'Migration hoàn tất! Xem Nhật ký thực thi để biết chi tiết.';
}

function cleanEmptyTasks() {
  const sheet = getSheet('Tasks');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('name');
  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][nameCol] || !String(data[i][nameCol]).trim()) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return 'Đã xóa ' + deleted + ' task rỗng';
}

// ============================================================
// DEBUG
// ============================================================

function getJobsDebug() {
  var sheet = getSheet('Jobs');
  var lastRow = sheet ? sheet.getLastRow() : -1;
  var jobs = getJobs();
  return { sheetExists: !!sheet, lastRow: lastRow, jobsLength: jobs.length, first: jobs[0] || null };
}

function debugAll() {
  var results = {};
  var sheets = ['Users','Jobs','Tasks','Statuses','TaskTemplates','Settings'];
  sheets.forEach(function(name) {
    try {
      var sheet = getSheet(name);
      if (!sheet) { results[name] = 'MISSING'; return; }
      results[name] = sheet.getLastRow() + ' rows';
    } catch(e) { results[name] = 'ERROR: ' + e.message; }
  });
  try { results['_getInitialData'] = JSON.stringify(getInitialData()).substring(0, 200); } catch(e) { results['_getInitialData'] = 'ERROR: ' + e.message; }
  try { results['_getDashboardStats'] = 'OK: ' + JSON.stringify(getDashboardStats()).substring(0,100); } catch(e) { results['_getDashboardStats'] = 'ERROR: ' + e.message; }
  return JSON.stringify(results, null, 2);
}
