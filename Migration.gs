// ============================================================
// MIGRATION FROM XLSX
// Chạy trực tiếp trong file khách gửi (container-bound script).
// ============================================================

function migrateFromXlsx() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var tz   = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';

  // 1. Tìm sheet nguồn (không phải system sheet của HaiLux)
  var sysNames = ['Jobs','Tasks','Statuses','Users','TaskTemplates','Settings'];
  var srcSheet = ss.getSheets().filter(function(s) {
    return sysNames.indexOf(s.getName()) < 0;
  })[0];
  if (!srcSheet) throw new Error('Không tìm thấy sheet dữ liệu nguồn');

  // 2. Thử export xlsx để lấy ảnh — bỏ qua nếu file quá lớn (>50MB giới hạn GAS)
  var entryMap = {};
  var rowToImg = {};
  try {
    var xlsxBlob = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
    ).getBlob().setContentType('application/zip');
    Utilities.unzip(xlsxBlob).forEach(function(e) { entryMap[e.getName()] = e; });

    var sheetIdx  = ss.getSheets().indexOf(srcSheet) + 1;
    var sheetRels = entryMap['xl/worksheets/_rels/sheet' + sheetIdx + '.xml.rels'];
    if (sheetRels) {
      var dm = sheetRels.getDataAsString().match(/Target="\.\.\/drawings\/(drawing\d+\.xml)"/);
      if (dm) {
        var drawKey = 'xl/drawings/' + dm[1];
        var relsKey = 'xl/drawings/_rels/' + dm[1] + '.rels';
        if (entryMap[drawKey]) {
          rowToImg = _xlsxDrawingMap(
            entryMap[drawKey].getDataAsString(),
            entryMap[relsKey] ? entryMap[relsKey].getDataAsString()
              : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
          );
        }
      }
    }
  } catch(e) {
    Logger.log('Không lấy được ảnh (file quá lớn hoặc lỗi export): ' + e.message);
  }

  // 3. Đọc data thẳng từ sheet
  var data = srcSheet.getDataRange().getValues();
  var headerRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i].some(function(c) { return String(c).trim() === 'Mã SP'; })) {
      headerRow = i; break;
    }
  }
  if (headerRow < 0) throw new Error('Không tìm thấy header row (cần có cột "Mã SP")');

  var headers = data[headerRow].map(function(h) { return String(h).trim(); });
  var col = {
    code:     headers.indexOf('Mã SP'),
    name:     headers.indexOf('Tên Sản Phẩm'),
    received: headers.indexOf('Ngày nhận'),
    deadline: headers.indexOf('Deadline'),
    category: headers.indexOf('Phân loại'),
    customer: headers.indexOf('Tên khách hàng'),
    contact:  headers.indexOf('SĐT + Địa chỉ'),
    scope:    headers.indexOf('Hạng Mục Sửa'),
    revenue:  headers.indexOf('Báo giá'),
    status:   headers.indexOf('Trạng Thái')
  };

  // 4. Khởi tạo hệ thống HaiLux
  setupSheets();
  _ensureJobStatus('Đã thanh toán', '#9C27B0', 5);

  var statusByLabel = {};
  sheetToObjects(getSheet('Statuses')).forEach(function(s) {
    if (s.entity_type === 'job') statusByLabel[s.label] = s.id;
  });

  var settings  = getSettings();
  var folderRow = settings.find(function(r) { return r.key === 'drive_folder_id'; });
  var imgFolder = (folderRow && folderRow.value)
    ? DriveApp.getFolderById(folderRow.value)
    : DriveApp.getRootFolder();

  function fmtDate(v) {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return String(v);
  }

  // 5. Migrate từng dòng
  var jobSheet = getSheet('Jobs');
  var count = 0;

  for (var r = headerRow + 1; r < data.length; r++) {
    var row  = data[r];
    var code = col.code >= 0 ? String(row[col.code] || '').trim() : '';
    var name = col.name >= 0 ? String(row[col.name] || '').trim() : '';
    if (!code || !name) continue;

    var avatarId = '';
    var imgName  = rowToImg[r];
    if (imgName && entryMap['xl/media/' + imgName]) {
      try {
        var ext  = imgName.split('.').pop().toLowerCase();
        var mime = ext === 'png' ? 'image/png' : 'image/jpeg';
        var imgFile = imgFolder.createFile(
          entryMap['xl/media/' + imgName].setContentType(mime).setName(code + '_avatar.' + ext)
        );
        imgFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        avatarId = imgFile.getId();
      } catch(e) {
        Logger.log('Row ' + (r + 1) + ' image error: ' + e.message);
      }
    }

    var statusId = statusByLabel[_mapJobStatus(col.status >= 0 ? String(row[col.status] || '') : '')]
                || statusByLabel['Chưa làm'] || '';

    jobSheet.appendRow([
      genId(),
      code,
      name,
      col.category >= 0 ? (String(row[col.category] || '').trim() || 'Khách lẻ') : 'Khách lẻ',
      col.customer >= 0 ? String(row[col.customer] || '') : '',
      col.contact  >= 0 ? String(row[col.contact]  || '') : '',
      fmtDate(col.received >= 0 ? row[col.received] : ''),
      fmtDate(col.deadline >= 0 ? row[col.deadline] : ''),
      col.revenue  >= 0 ? (row[col.revenue] || '') : '',
      col.scope    >= 0 ? String(row[col.scope]    || '') : '',
      statusId,
      '',
      new Date().toISOString(),
      avatarId
    ]);
    count++;
  }

  return 'Migration xong! Đã tạo ' + count + ' jobs.';
}

function _xlsxDrawingMap(drawingXml, relsXml) {
  var rIdToFile = {};
  var relsNs = XmlService.getNamespace('http://schemas.openxmlformats.org/package/2006/relationships');
  XmlService.parse(relsXml).getRootElement().getChildren('Relationship', relsNs).forEach(function(rel) {
    rIdToFile[rel.getAttribute('Id').getValue()] = rel.getAttribute('Target').getValue().split('/').pop();
  });
  var rowToFile = {};
  (drawingXml.match(/<xdr:oneCellAnchor>[\s\S]*?<\/xdr:oneCellAnchor>/g) || []).forEach(function(a) {
    var rm = a.match(/<xdr:row>(\d+)<\/xdr:row>/);
    var em = a.match(/r:embed="(rId\d+)"/);
    if (rm && em && rIdToFile[em[1]]) rowToFile[parseInt(rm[1])] = rIdToFile[em[1]];
  });
  return rowToFile;
}

function _mapJobStatus(raw) {
  var map = {
    'đang sửa':      'Đang làm',
    'đang làm':      'Đang làm',
    'chưa làm':      'Chưa làm',
    'hoàn thành':    'Hoàn thành',
    'đã thanh toán': 'Đã thanh toán',
    'hủy':           'Hủy'
  };
  return map[(raw || '').toLowerCase().trim()] || 'Chưa làm';
}

// Chạy sau khi upload thư mục job_images lên Drive.
function migrateImages(folderId) {
  if (!folderId) throw new Error('Cần truyền folderId của thư mục ảnh trên Drive');

  var folder   = DriveApp.getFolderById(folderId);
  var jobs     = getJobs();
  var jobSheet = getSheet('Jobs');
  var headers  = getHeaders(jobSheet);
  var avatarCol = headers.indexOf('avatar_id') + 1;

  var codeToRow = {};
  jobs.forEach(function(j, i) { codeToRow[j.code] = i + 2; });

  var files = folder.getFiles();
  var updated = 0, skipped = 0;

  while (files.hasNext()) {
    var file = files.next();
    var m = file.getName().match(/^(.+)_avatar\./i);
    if (!m) { skipped++; continue; }

    var row = codeToRow[m[1]];
    if (!row) { Logger.log('Không tìm thấy job với code: ' + m[1]); skipped++; continue; }

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    jobSheet.getRange(row, avatarCol).setValue(file.getId());
    updated++;
  }

  return 'Xong! Cập nhật ' + updated + ' ảnh, bỏ qua ' + skipped + ' file.';
}
