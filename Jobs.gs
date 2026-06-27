// ============================================================
// JOBS API
// ============================================================

function getJobs() {
  return sheetToObjects(getSheet('Jobs'));
}

function getJob(id) {
  return getJobs().find(j => j.id === id) || null;
}

function generateJobCode() {
  const jobs = getJobs();
  const parsed = jobs
    .map(j => { const m = String(j.code || '').match(/^([A-Za-z]+)(\d+)$/); return m ? { prefix: m[1].toUpperCase(), num: parseInt(m[2], 10), pad: m[2].length } : null; })
    .filter(Boolean);
  if (!parsed.length) return 'JOB0001';
  const top = parsed.reduce((a, b) => b.num > a.num ? b : a);
  return top.prefix + String(top.num + 1).padStart(top.pad, '0');
}

function getJobsWithStats() {
  const jobs  = getJobs();
  const tasks = getTasks();
  return jobs
    .sort((a, b) => {
      const da = b.received_date || '';
      const db = a.received_date || '';
      if (!da && !db) return (b.created_at || '').localeCompare(a.created_at || '');
      if (!da) return 1;
      if (!db) return -1;
      if (da !== db) return da.localeCompare(db);
      const numA = parseInt((a.code || '').replace(/\D/g, ''), 10) || 0;
      const numB = parseInt((b.code || '').replace(/\D/g, ''), 10) || 0;
      return numB - numA;
    })
    .map(j => {
      const jobTasks  = tasks.filter(t => t.job_id === j.id);
      const assignees = [...new Set(jobTasks.map(t => t.assignee_id).filter(Boolean))];
      return { ...j, task_count: jobTasks.length, assignee_count: assignees.length };
    });
}

function createJob(data) {
  const sheet = getSheet('Jobs');
  const id = genId();
  const now = new Date().toISOString();
  const statusId = data.status_id || getDefaultStatusId('job');
  sheet.appendRow([
    id,
    data.code || '',
    data.name || '',
    data.category || '',
    data.customer_name || '',
    data.customer_contact || '',
    data.received_date || '',
    data.deadline || '',
    data.revenue || 0,
    data.repair_scope || '',
    statusId,
    data.notes || '',
    now,
    data.avatar_id || '',
  ]);
  return { id, ...data, status_id: statusId, created_at: now };
}

function updateJob(id, data) {
  const sheet = getSheet('Jobs');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);
  Object.keys(data).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });
  return { id, ...data };
}

function deleteJob(id) {
  const sheet = getSheet('Jobs');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  const tasks = getTasks().filter(t => t.job_id === id);
  tasks.forEach(t => deleteTask(t.id));
  return true;
}
