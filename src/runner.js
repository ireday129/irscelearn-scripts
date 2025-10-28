/** Generic batch runner used by markCleanAsReported() and friends.
 * Persists an offset in Script Properties, calls stepFn(offset, limit),
 * and loops until either the step reports done OR we’re near exec time limit.
 * Re-invokes itself via a one-off time-based trigger if more work remains.
 */
function runJob(jobKey, stepFn, limit) {
  const props = PropertiesService.getScriptProperties();
  let offset = Number(props.getProperty(jobKey + '/offset') || 0);

  const start = Date.now();
  const MAX_MS = 280000; // ~4m40s to stay under 6-min limit

  while (true) {
    const res = stepFn(offset, limit); // MUST return {processed:number, done:boolean}
    const processed = (res && typeof res.processed === 'number') ? res.processed : 0;
    const done = !!(res && res.done);

    if (processed > 0) {
      offset += processed;
      props.setProperty(jobKey + '/offset', String(offset));
    }

    if (done) {
      props.deleteProperty(jobKey + '/offset');
      return; // finished
    }

    if ((Date.now() - start) > MAX_MS) {
      // Schedule a continuation soon and bail out of this execution
      ScriptApp.newTrigger(getSelfInvokerName_(jobKey))
               .timeBased()
               .after(30 * 1000) // 30s
               .create();
      return;
    }
    // If the step said not done but also didn't report progress, avoid a tight loop
    if (processed === 0) {
      // Safety: schedule and exit
      ScriptApp.newTrigger(getSelfInvokerName_(jobKey))
               .timeBased()
               .after(30 * 1000)
               .create();
      return;
    }
  }
}

/** Map job keys to the function that should be called to continue the job. */
function getSelfInvokerName_(jobKey) {
  // Extend as you add more jobs
  switch (jobKey) {
    case 'JOB/MARK_REPORTED': return 'runMarkReportedBatch';
    // e.g., case 'JOB/BUILD_CLEAN': return 'runBuildCleanBatch';
    default: return 'runMarkReportedBatch';
  }
}