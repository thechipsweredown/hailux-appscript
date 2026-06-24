// ============================================================
// TASKS API
// ============================================================

function getTasks() {
  return sheetToObjects(getSheet('Tasks'));
}

function getTasksByJob(jobId) {
  return getTasks().filter(t => t.job_id === jobId).sort((a, b) => Number(a.order) - Number(b.order));
}

function getTasksByAssignee(assigneeId) {
  return getTasks().filter(t => t.assignee_id === assigneeId);
}

function getTaskById(id) {
  return getTasks().find(t => t.id === id) || null;
}

function createTask(data) {
  const sheet = getSheet('Tasks');
  const id = genId();
  const now = new Date().toISOString();
  const statusId = data.status_id || getDefaultStatusId('task');
  sheet.appendRow([
    id,
    data.job_id || '',
    data.name || '',
    data.order || 1,
    data.assignee_id || '',
    data.deadline || '',
    statusId,
    data.completed_at || '',
    data.notes || '',
    now,
  ]);
  return { id, ...data, status_id: statusId, created_at: now };
}

function updateTask(id, data) {
  const sheet = getSheet('Tasks');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);

  if (data.status_id) {
    const statuses = getStatuses();
    const s = statuses.find(s => s.id === data.status_id);
    if (s && s.label === 'Hoàn thành' && !data.completed_at) {
      data.completed_at = new Date().toISOString();
    }
  }

  Object.keys(data).forEach(key => {
    if ((key === 'job_id' || key === 'assignee_id') && !data[key]) return;
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });

  autoUpdateJobStatus(data.job_id || getTaskById(id).job_id);
  return { id, ...data };
}

function deleteTask(id) {
  const sheet = getSheet('Tasks');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}

function reorderTasks(jobId, orderedIds) {
  orderedIds.forEach((id, index) => {
    updateTask(id, { order: index + 1, job_id: jobId });
  });
  return true;
}

// ============================================================
// TASK BUSINESS LOGIC
// ============================================================

function autoUpdateJobStatus(jobId) {
  if (!jobId) return;
  const tasks = getTasksByJob(jobId);
  if (tasks.length === 0) return;

  const statuses = getStatuses();
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');
  const doneStatus       = findStatusByNorm(taskStatuses, 'hoànthành') || findStatusByNorm(taskStatuses, 'đãxong');
  const inProgressStatus = findStatusByNorm(taskStatuses, 'đanglàm');

  const jobStatuses   = statuses.filter(s => s.entity_type === 'job');
  const jobDone       = findStatusByNorm(jobStatuses, 'hoànthành');
  const jobInProgress = findStatusByNorm(jobStatuses, 'đanglàm');
  const jobNotStarted = findStatusByNorm(jobStatuses, 'chưalàm');

  const allDone       = tasks.every(t => doneStatus    && t.status_id === doneStatus.id);
  const anyInProgress = tasks.some(t  => inProgressStatus && t.status_id === inProgressStatus.id);
  const anyDone       = tasks.some(t  => doneStatus    && t.status_id === doneStatus.id);

  let newStatusId;
  if (allDone && jobDone)                          newStatusId = jobDone.id;
  else if ((anyInProgress || anyDone) && jobInProgress) newStatusId = jobInProgress.id;
  else if (jobNotStarted)                          newStatusId = jobNotStarted.id;

  if (newStatusId) updateJob(jobId, { status_id: newStatusId });
}

function getEnrichedTasksForAssignee(assigneeId) {
  const allTasks = getTasks();
  const statuses = getStatuses();
  const jobs = getJobs();
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');

  const doneStatus    = findStatusByNorm(taskStatuses, 'hoànthành')
                     || findStatusByNorm(taskStatuses, 'đãxong');
  const ipStatus      = findStatusByNorm(taskStatuses, 'đanglàm');
  const pendingStatus = findStatusByNorm(taskStatuses, 'treo');

  const tasksByJob = {};
  allTasks.forEach(t => {
    if (!tasksByJob[t.job_id]) tasksByJob[t.job_id] = [];
    tasksByJob[t.job_id].push(t);
  });
  Object.values(tasksByJob).forEach(arr =>
    arr.sort((a, b) => Number(a.order) - Number(b.order))
  );

  const myTasks = allTasks.filter(t => t.assignee_id === assigneeId && t.name && String(t.name).trim());

  return myTasks.map(task => {
    const jobTasks = tasksByJob[task.job_id] || [];
    const myOrder  = Number(task.order);

    const prevTask     = jobTasks.find(t => Number(t.order) === myOrder - 1) || null;
    const prevPrevTask = jobTasks.find(t => Number(t.order) === myOrder - 2) || null;

    const isDone    = doneStatus    && task.status_id === doneStatus.id;
    const isIP      = ipStatus      && task.status_id === ipStatus.id;
    const isPassable = t => doneStatus && t.status_id === doneStatus.id;
    const prevIsDone = !prevTask    || isPassable(prevTask);
    const ppIsDone   = prevPrevTask && isPassable(prevPrevTask);

    let availability;
    if (isDone)          availability = 'done';
    else if (isIP)       availability = 'in_progress';
    else if (prevIsDone) availability = 'ready';
    else if (ppIsDone)   availability = 'upcoming';
    else                 availability = 'blocked';

    const job    = jobs.find(j => j.id === task.job_id) || null;
    const status = taskStatuses.find(s => s.id === task.status_id) || null;
    const effectiveDeadline = task.deadline || (job ? job.deadline : '');

    return { ...task, job, status, effective_deadline: effectiveDeadline, availability, prev_task: prevTask };
  });
}

function getTaskWithImages(taskId) {
  const task = getTaskById(taskId);
  const ids = task && task.evidence_id ? String(task.evidence_id).split(',').filter(Boolean) : [];
  const images = ids.map(id => ({ drive_file_id: id, cdn_url: 'https://lh3.googleusercontent.com/d/' + id }));

  let availability = 'ready';
  let doneStatus = null;
  if (task) {
    const allTasks = getTasksByJob(task.job_id).sort((a,b) => Number(a.order)-Number(b.order));
    const statuses = getStatuses().filter(s => s.entity_type === 'task');
    doneStatus = findStatusByNorm(statuses, 'hoànthành') || findStatusByNorm(statuses, 'đãxong');
    const myOrder = Number(task.order);
    const prevTask     = allTasks.find(t => Number(t.order) === myOrder - 1) || null;
    const prevPrevTask = allTasks.find(t => Number(t.order) === myOrder - 2) || null;
    const isPassable   = t => doneStatus && t.status_id === doneStatus.id;
    const prevIsDone   = !prevTask || isPassable(prevTask);
    const ppIsDone     = prevPrevTask && isPassable(prevPrevTask);
    const ipStatus     = findStatusByNorm(statuses, 'đanglàm');
    const isDone       = doneStatus && task.status_id === doneStatus.id;
    const isIP         = ipStatus   && task.status_id === ipStatus.id;
    if (isDone)          availability = 'done';
    else if (isIP)       availability = 'in_progress';
    else if (prevIsDone) availability = 'ready';
    else if (ppIsDone)   availability = 'upcoming';
    else                 availability = 'blocked';
  }

  var prevTaskImages = [];
  if (task && doneStatus) {
    const allTasks = getTasksByJob(task.job_id).sort((a,b) => Number(a.order)-Number(b.order));
    allTasks.forEach(function(t) {
      if (t.id === taskId) return;
      if (Number(t.order) >= Number(task.order)) return;
      if (t.status_id !== doneStatus.id) return;
      const eIds = t.evidence_id ? String(t.evidence_id).split(',').filter(Boolean) : [];
      if (eIds.length) prevTaskImages.push({
        task_name: t.name,
        images: eIds.map(function(id) { return { drive_file_id: id, cdn_url: 'https://lh3.googleusercontent.com/d/' + id }; })
      });
    });
  }

  return { task, images, availability, prev_task_images: prevTaskImages };
}

function getAllTasksWithDetails(filters) {
  filters = filters || {};
  const tasks    = getTasks();
  const jobs     = getJobs();
  const users    = getUsers();
  const statuses = getStatuses();
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');

  const jobMap  = {};  jobs.forEach(j => jobMap[j.id] = j);
  const userMap = {};  users.forEach(u => userMap[u.id] = u);
  const stMap   = {};  taskStatuses.forEach(s => stMap[s.id] = s);

  return tasks
    .map(t => ({
      ...t,
      job:      jobMap[t.job_id]       || null,
      assignee: userMap[t.assignee_id] || null,
      status:   stMap[t.status_id]     || null,
      effective_deadline: t.deadline || (jobMap[t.job_id] ? jobMap[t.job_id].deadline : ''),
    }))
    .filter(t => {
      if (!t.name || !String(t.name).trim()) return false;
      if (filters.assignee_id && t.assignee_id !== filters.assignee_id) return false;
      if (filters.status_id   && t.status_id   !== filters.status_id)   return false;
      if (filters.deadline_from && t.effective_deadline && t.effective_deadline < filters.deadline_from) return false;
      if (filters.deadline_to   && t.effective_deadline && t.effective_deadline > filters.deadline_to + 'T23:59:59') return false;
      return true;
    })
    .sort((a, b) => {
      if (!a.effective_deadline && !b.effective_deadline) return 0;
      if (!a.effective_deadline) return 1;
      if (!b.effective_deadline) return -1;
      return a.effective_deadline.localeCompare(b.effective_deadline);
    });
}
