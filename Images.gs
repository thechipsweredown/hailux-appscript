// ============================================================
// IMAGES API
// ============================================================

function _getOrCreateImageFolder() {
  const settings = getSettings();
  const folderSetting = settings.find(s => s.key === 'drive_folder_id');
  if (folderSetting && folderSetting.value) {
    try { return DriveApp.getFolderById(folderSetting.value); } catch(e) {}
  }
  const folders = DriveApp.getFoldersByName('HaiLux_Images');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('HaiLux_Images');
  updateSetting('drive_folder_id', folder.getId());
  return folder;
}

function uploadJobAvatarOnly(base64Data, mimeType, filename) {
  const folder = _getOrCreateImageFolder();
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { drive_file_id: file.getId() };
}

function uploadJobAvatar(base64Data, mimeType, filename, jobId) {
  const folder = _getOrCreateImageFolder();
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  updateJob(jobId, { avatar_id: fileId });
  return { drive_file_id: fileId, cdn_url: 'https://lh3.googleusercontent.com/d/' + fileId };
}

function uploadTaskEvidence(base64Data, mimeType, filename, taskId) {
  const folder = _getOrCreateImageFolder();
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();

  const task = getTaskById(taskId);
  const existing = (task && task.evidence_id) ? String(task.evidence_id).split(',').filter(Boolean) : [];
  existing.push(fileId);
  updateTask(taskId, { evidence_id: existing.join(','), job_id: task ? task.job_id : '' });

  return { drive_file_id: fileId, cdn_url: 'https://lh3.googleusercontent.com/d/' + fileId };
}

function deleteTaskEvidence(taskId, fileId) {
  const task = getTaskById(taskId);
  if (!task) return false;
  const ids = String(task.evidence_id || '').split(',').filter(id => id && id !== fileId);
  updateTask(taskId, { evidence_id: ids.join(','), job_id: task.job_id });
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
  return true;
}
