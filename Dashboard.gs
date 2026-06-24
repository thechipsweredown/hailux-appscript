// ============================================================
// DASHBOARD & COMPOSITE
// ============================================================

function getDashboardStats(filters) {
  filters = filters || {};
  const dateFrom = filters.dateFrom ? new Date(filters.dateFrom) : null;
  const dateTo   = filters.dateTo   ? new Date(filters.dateTo + 'T23:59:59') : null;
  const category = filters.category || '';

  let jobs = getJobs();
  if (dateFrom) jobs = jobs.filter(j => j.received_date && new Date(j.received_date) >= dateFrom);
  if (dateTo)   jobs = jobs.filter(j => j.received_date && new Date(j.received_date) <= dateTo);
  if (category) jobs = jobs.filter(j => (j.category||'') === category);

  const jobIds = new Set(jobs.map(j => j.id));
  const allTasks = getTasks();
  const tasks = allTasks.filter(t => jobIds.has(t.job_id));

  const statuses = getStatuses();
  const users = getUsers();
  const jobStatuses  = statuses.filter(s => s.entity_type === 'job');
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');

  const jobStatusCounts = {};
  jobStatuses.forEach(s => {
    jobStatusCounts[s.label] = jobs.filter(j => j.status_id === s.id).length;
  });

  const taskStatusCounts = {};
  taskStatuses.forEach(s => {
    taskStatusCounts[s.label] = tasks.filter(t => t.status_id === s.id).length;
  });

  const totalRevenue = jobs.reduce((sum, j) => sum + (Number(j.revenue) || 0), 0);
  const doneStatus   = findStatusByNorm(jobStatuses, 'hoànthành');
  const completedRevenue = jobs
    .filter(j => doneStatus && j.status_id === doneStatus.id)
    .reduce((sum, j) => sum + (Number(j.revenue) || 0), 0);

  const now = new Date();
  const overdueJobs = jobs.filter(j => {
    if (!j.deadline) return false;
    return (!doneStatus || j.status_id !== doneStatus.id) && new Date(j.deadline) < now;
  }).length;

  const doneTSId = findStatusByNorm(taskStatuses, 'hoànthành')?.id || findStatusByNorm(taskStatuses, 'đãxong')?.id;
  const ipTSId   = findStatusByNorm(taskStatuses, 'đanglàm')?.id;

  const employeeStats = users
    .filter(u => u.role === 'employee')
    .map(u => {
      const myTasks = tasks.filter(t => t.assignee_id === u.id);
      return {
        id: u.id, name: u.name,
        total: myTasks.length,
        done: myTasks.filter(t => t.status_id === doneTSId).length,
        in_progress: myTasks.filter(t => t.status_id === ipTSId).length,
        todo: myTasks.filter(t => t.status_id !== doneTSId && t.status_id !== ipTSId).length,
      };
    })
    .filter(e => e.total > 0);

  return {
    jobs:     { total: jobs.length, by_status: jobStatusCounts, overdue: overdueJobs },
    tasks:    { total: tasks.length, by_status: taskStatusCounts },
    revenue:  { total: totalRevenue, completed: completedRevenue },
    employees: employeeStats,
  };
}

function getJobWithDetails(jobId) {
  const job = getJob(jobId);
  if (!job) return null;
  const tasks    = getTasksByJob(jobId);
  const users    = getUsers();
  const statuses = getStatuses();
  const enrichedTasks = tasks.map(t => {
    const assignee = users.find(u => u.id === t.assignee_id);
    const status   = statuses.find(s => s.id === t.status_id);
    return { ...t, assignee, status };
  });
  const jobStatus = statuses.find(s => s.id === job.status_id);
  return { ...job, status: jobStatus, tasks: enrichedTasks };
}

function getInitialData() {
  const jobs = getJobs();
  const categories = [...new Set(jobs.map(j => j.category).filter(Boolean))].sort();
  return {
    users: getUsers(),
    statuses: getStatuses(),
    taskTemplates: getTaskTemplates(),
    settings: getSettings(),
    categories,
  };
}
