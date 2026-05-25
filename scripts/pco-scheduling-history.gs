/**
 * PCO Services_Scheduling — full history backfill (Apps Script).
 *
 * Uses the service_types → teams → plan_people route, which is the only
 * plan_people path that returns data (the plan-level route returns 404).
 *
 * The full scheduling history takes longer than the 6-minute Apps Script
 * limit. This script is resumable and self-scheduling:
 *
 *   1. Run pcoSchedulingHistoryStart() ONCE from the editor.
 *      - Writes headers and runs the first batch immediately.
 *      - Creates a 5-minute time-driven trigger that keeps calling
 *        pcoSchedulingHistoryRun() until all teams are done.
 *      - Deletes the trigger automatically when complete.
 *
 *   2. Watch the Executions tab — the last run will log
 *      "COMPLETE — all X teams processed."
 *
 * To start over from scratch:
 *   Run pcoSchedulingHistoryReset() then pcoSchedulingHistoryStart() again.
 *
 * SETUP:
 *   Same sheet + Apps Script project as pco-apps-script.gs (Code.gs).
 *   Uses the same PCO_APP_ID / PCO_SECRET Script Properties.
 */

const SCHED_HISTORY_KEY   = 'sched_history_progress';
const SCHED_TAB_NAME      = 'Services_Scheduling';
const SCHED_STATUS_LABELS = { C: 'Confirmed', U: 'Unconfirmed', D: 'Declined' };
const SCHED_HEADERS = [
  'pp_id', 'service_type', 'plan_date', 'plan_dates', 'plan_title', 'series_title',
  'plan_id', 'person_name', 'status', 'team', 'position',
  'responds_to', 'decline_reason', 'notes', 'status_updated_at', 'prepared_at',
];

// Each run stops after 4.5 minutes so it finishes before the 5-minute trigger
// re-fires and well clear of Apps Script's 6-minute hard kill.
const MAX_RUN_MS = 4.5 * 60 * 1000;

// ─── Entry points ─────────────────────────────────────────────────────────────

/**
 * Run this ONCE to kick off the full backfill.
 * Creates a 5-minute trigger and immediately runs the first batch.
 */
function pcoSchedulingHistoryStart() {
  deleteTrigger_();
  ScriptApp.newTrigger('pcoSchedulingHistoryRun')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('Auto-trigger created (every 5 min). Running first batch now...');
  pcoSchedulingHistoryRun();
}

/**
 * Called automatically by the trigger. Safe to run manually too.
 * Resumes from wherever the last run stopped.
 */
function pcoSchedulingHistoryRun() {
  const startTime = Date.now();
  const props     = PropertiesService.getScriptProperties();
  const appId     = props.getProperty('PCO_APP_ID');
  const secret    = props.getProperty('PCO_SECRET');
  if (!appId || !secret) throw new Error('Set PCO_APP_ID and PCO_SECRET in Script Properties.');
  const auth = 'Basic ' + Utilities.base64Encode(appId + ':' + secret);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(SCHED_TAB_NAME);

  let progress = loadProgress(props);

  if (progress.status === 'complete') {
    Logger.log('COMPLETE — already finished. Run pcoSchedulingHistoryReset() to start over.');
    deleteTrigger_();
    return;
  }

  // Phase 1: build plan metadata cache (planId → service_type, date, dates, title, series).
  if (!progress.planMeta) {
    Logger.log('Building plan metadata cache (service types → plans)...');
    progress.planMeta = buildPlanMeta(auth);
    saveProgress(props, progress);
    Logger.log('Cached metadata for ' + Object.keys(progress.planMeta).length + ' plans.');
  }

  // Phase 2: build team list (all service types → all their teams).
  if (!progress.teams) {
    Logger.log('Building team list (service types → teams)...');
    progress.teams       = buildServiceTypeTeams(auth);
    progress.teamIndex   = 0;
    progress.rowsWritten = 0;
    Logger.log('Found ' + progress.teams.length + ' teams across all service types.');

    if (!sh) sh = ss.insertSheet(SCHED_TAB_NAME);
    sh.clear();
    sh.getRange(1, 1, 1, SCHED_HEADERS.length).setValues([SCHED_HEADERS]);
    sh.setFrozenRows(1);
    saveProgress(props, progress);
  }

  if (!sh) sh = ss.getSheetByName(SCHED_TAB_NAME);

  const teams = progress.teams;
  let   ti    = progress.teamIndex;

  Logger.log('Resuming at team ' + (ti + 1) + ' of ' + teams.length
    + ' (' + progress.rowsWritten + ' rows written so far)');

  while (ti < teams.length) {
    // Stop before the 4.5-minute wall so the next trigger run starts clean.
    if (Date.now() - startTime > MAX_RUN_MS) {
      progress.teamIndex = ti;
      saveProgress(props, progress);
      Logger.log('Time limit reached — saved at team ' + (ti + 1)
        + ' of ' + teams.length + '. Rows written so far: ' + progress.rowsWritten
        + '. Trigger will resume in ~5 min.');
      return;
    }

    const team  = teams[ti];
    const batch = [];

    // Page through ALL past plan_people for this team.
    // include=plan pulls plan attributes so we get dates/title without extra requests.
    let ppPath = '/services/v2/service_types/' + team.stId + '/teams/' + team.teamId
      + '/plan_people?filter=past&include=plan&per_page=100';
    while (ppPath) {
      const ppResp = pcoGetHistory(PCO_HOST + ppPath, auth);

      // Build plan data lookup from JSON:API included array.
      const planData = {};
      for (const inc of (ppResp.included || [])) {
        if (inc.type === 'Plan') {
          const a = inc.attributes;
          planData[inc.id] = {
            date:   a.sort_date    ? a.sort_date.slice(0, 10) : '',
            dates:  a.dates        || '',
            title:  a.title        || '',
            series: a.series_title || '',
          };
        }
      }

      for (const pp of (ppResp.data || [])) {
        const p      = pp.attributes;
        const planId = pp.relationships && pp.relationships.plan && pp.relationships.plan.data
          ? pp.relationships.plan.data.id : null;

        // Prefer included plan data; fall back to pre-built planMeta cache.
        const meta = (planId && planData[planId])
          || (planId && progress.planMeta[planId])
          || {};
        const serviceType = (planId && progress.planMeta[planId]
          && progress.planMeta[planId].serviceType) || team.stName;

        batch.push([
          pp.id,
          serviceType,
          meta.date   || '',
          meta.dates  || '',
          meta.title  || '',
          meta.series || '',
          planId      || '',
          p.name      || '',
          SCHED_STATUS_LABELS[p.status] || p.status || '',
          team.teamName,
          p.team_position_name           || '',
          p.responds_to_name             || '',
          p.decline_reason               || '',
          p.notes                        || '',
          p.status_updated_at            || '',
          p.prepare_notification_sent_at || '',
        ]);
      }

      const ppNext = ppResp.links && ppResp.links.next;
      ppPath = ppNext ? ppNext.replace(PCO_HOST, '') : null;
    }

    if (batch.length) {
      const lastRow = sh.getLastRow();
      sh.getRange(lastRow + 1, 1, batch.length, SCHED_HEADERS.length).setValues(batch);
      progress.rowsWritten += batch.length;
    }

    ti++;
    // Save after every team so a crash or timeout doesn't lose progress.
    progress.teamIndex = ti;
    saveProgress(props, progress);
  }

  // All teams done.
  progress.status    = 'complete';
  progress.teamIndex = ti;
  saveProgress(props, progress);
  deleteTrigger_();
  Logger.log('COMPLETE — all ' + teams.length + ' teams processed. Total rows written: '
    + progress.rowsWritten);
}

/**
 * Clears saved progress and the tab so the next start is fully fresh.
 * Also deletes any running trigger.
 */
function pcoSchedulingHistoryReset() {
  deleteTrigger_();
  PropertiesService.getScriptProperties().deleteProperty(SCHED_HISTORY_KEY);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHED_TAB_NAME);
  if (sh) sh.clear();
  Logger.log('Progress reset. Run pcoSchedulingHistoryStart() to begin the backfill.');
}

// ─── Trigger management ───────────────────────────────────────────────────────

function deleteTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'pcoSchedulingHistoryRun')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ─── Progress persistence ─────────────────────────────────────────────────────

function loadProgress(props) {
  const raw = props.getProperty(SCHED_HISTORY_KEY);
  return raw ? JSON.parse(raw) : {};
}

function saveProgress(props, progress) {
  props.setProperty(SCHED_HISTORY_KEY, JSON.stringify(progress));
}

// ─── Fetch helpers ────────────────────────────────────────────────────────────

// Returns {planId: {serviceType, date, dates, title, series}} for every plan.
function buildPlanMeta(auth) {
  const meta = {};
  let stPath = '/services/v2/service_types?per_page=100';
  while (stPath) {
    const stResp = pcoGetHistory(PCO_HOST + stPath, auth);
    for (const st of (stResp.data || [])) {
      const serviceType = st.attributes.name || st.id;
      let planPath = '/services/v2/service_types/' + st.id + '/plans?per_page=100';
      while (planPath) {
        const planResp = pcoGetHistory(PCO_HOST + planPath, auth);
        for (const plan of (planResp.data || [])) {
          const a = plan.attributes;
          meta[plan.id] = {
            serviceType,
            date:   a.sort_date    ? a.sort_date.slice(0, 10) : '',
            dates:  a.dates        || '',
            title:  a.title        || '',
            series: a.series_title || '',
          };
        }
        const next = planResp.links && planResp.links.next;
        planPath = next ? next.replace(PCO_HOST, '') : null;
      }
    }
    const next = stResp.links && stResp.links.next;
    stPath = next ? next.replace(PCO_HOST, '') : null;
  }
  return meta;
}

// Returns [{stId, stName, teamId, teamName}] for every team in every service type.
function buildServiceTypeTeams(auth) {
  const teams = [];
  let stPath = '/services/v2/service_types?per_page=100';
  while (stPath) {
    const stResp = pcoGetHistory(PCO_HOST + stPath, auth);
    for (const st of (stResp.data || [])) {
      const stName = st.attributes.name || st.id;
      let teamPath = '/services/v2/service_types/' + st.id + '/teams?per_page=100';
      while (teamPath) {
        const teamResp = pcoGetHistory(PCO_HOST + teamPath, auth);
        for (const team of (teamResp.data || [])) {
          teams.push({
            stId:     st.id,
            stName:   stName,
            teamId:   team.id,
            teamName: team.attributes.name || team.id,
          });
        }
        const next = teamResp.links && teamResp.links.next;
        teamPath = next ? next.replace(PCO_HOST, '') : null;
      }
    }
    const next = stResp.links && stResp.links.next;
    stPath = next ? next.replace(PCO_HOST, '') : null;
  }
  return teams;
}

function pcoGetHistory(url, auth, depth, retries) {
  depth   = depth   || 0;
  retries = retries || 0;
  if (depth > 5) throw new Error('Too many redirects: ' + url);

  let resp;
  try {
    resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: auth, Accept: 'application/json' },
      muteHttpExceptions: true,
      followRedirects: false,
    });
  } catch (e) {
    // UrlFetchApp throws Exception: Timeout on network-level timeouts.
    if (retries < 3) {
      Utilities.sleep((retries + 1) * 5000); // 5s, 10s, 15s backoff
      return pcoGetHistory(url, auth, depth, retries + 1);
    }
    throw new Error('Timeout after 3 retries: ' + url);
  }
  const code = resp.getResponseCode();
  const body = resp.getContentText();

  if (code === 429) {
    Utilities.sleep(3000);
    return pcoGetHistory(url, auth, depth);
  }
  if (code === 302) {
    let loc = '';
    try { loc = JSON.parse(body).location || ''; } catch (e) {}
    if (!loc) throw new Error('302 with no location body for ' + url);
    return pcoGetHistory(loc, auth, depth + 1, retries);
  }
  if (code !== 200) throw new Error('PCO ' + code + ' ' + url + ': ' + body.slice(0, 200));
  return JSON.parse(body);
}
