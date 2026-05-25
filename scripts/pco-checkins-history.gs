/**
 * PCO Check-Ins — full history backfill (Apps Script).
 *
 * Writes every individual check-in record (all kinds: Regular, Guest, Volunteer)
 * to a CheckIns_People tab. Resumable and self-scheduling — same pattern as
 * pco-scheduling-history.gs.
 *
 *   1. Run pcoCheckInsHistoryStart() ONCE from the editor.
 *      - Creates the tab, writes headers, and starts the first batch immediately.
 *      - Creates a 5-minute trigger that keeps calling pcoCheckInsHistoryRun()
 *        until all records are written.
 *
 *   2. Watch Executions log — the last run logs
 *      "COMPLETE — X check-in records written."
 *
 *   3. To start over: run pcoCheckInsHistoryReset() then pcoCheckInsHistoryStart().
 *
 * SETUP:
 *   Same sheet + Apps Script project as pco-apps-script.gs (Code.gs).
 *   Uses the same PCO_APP_ID / PCO_SECRET Script Properties.
 *   PCO_HOST and pcoGetHistory() are defined in the companion scripts —
 *   all three files must live in the same Apps Script project.
 *
 * After the backfill is complete, wire up incremental sync in pcoPullAll()
 * using a where[updated_at][gte] filter on /check-ins/v2/check_ins.
 */

const CI_HISTORY_KEY = 'ci_history_progress';
const CI_TAB_NAME    = 'CheckIns_People';
const CI_HEADERS     = [
  'ci_id', 'event_name', 'event_id', 'period_id', 'period_starts_at',
  'first_name', 'last_name', 'kind', 'checked_in_at', 'checked_out_at',
  'one_time_guest', 'security_code', 'number',
  'emergency_contact_name', 'emergency_contact_phone_number', 'medical_notes',
];

// Stop fetching after 4.5 minutes to stay clear of the 6-minute hard kill
// and finish before the 5-minute trigger re-fires.
const CI_MAX_RUN_MS = 4.5 * 60 * 1000;

// ─── Entry points ─────────────────────────────────────────────────────────────

/**
 * Run ONCE to kick off the full backfill.
 */
function pcoCheckInsHistoryStart() {
  deleteCiTrigger_();
  ScriptApp.newTrigger('pcoCheckInsHistoryRun')
    .timeBased().everyMinutes(5).create();
  Logger.log('Trigger created (every 5 min). Running first batch now...');
  pcoCheckInsHistoryRun();
}

/**
 * Called automatically by the trigger. Safe to run manually too.
 * Resumes from the saved next-page URL.
 */
function pcoCheckInsHistoryRun() {
  const startTime = Date.now();
  const props     = PropertiesService.getScriptProperties();
  const appId     = props.getProperty('PCO_APP_ID');
  const secret    = props.getProperty('PCO_SECRET');
  if (!appId || !secret) throw new Error('Set PCO_APP_ID and PCO_SECRET in Script Properties.');
  const auth = 'Basic ' + Utilities.base64Encode(appId + ':' + secret);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let   sh = ss.getSheetByName(CI_TAB_NAME);

  let progress = loadCiProgress_(props);

  if (progress.status === 'complete') {
    Logger.log('COMPLETE — already finished. Run pcoCheckInsHistoryReset() to start over.');
    deleteCiTrigger_();
    return;
  }

  // First run: create/clear the sheet and set the starting URL.
  if (!progress.nextUrl) {
    if (!sh) sh = ss.insertSheet(CI_TAB_NAME);
    sh.clear();
    sh.getRange(1, 1, 1, CI_HEADERS.length).setValues([CI_HEADERS]);
    sh.setFrozenRows(1);
    // include=person gets first/last name; event_period gets service date;
    // event gets event name. order=created_at gives consistent pagination.
    progress.nextUrl     = PCO_HOST + '/check-ins/v2/check_ins'
      + '?per_page=100&include=person,event_period,event&order=created_at';
    progress.rowsWritten = 0;
    saveCiProgress_(props, progress);
    Logger.log('Sheet created. Starting paginated fetch...');
  }

  if (!sh) sh = ss.getSheetByName(CI_TAB_NAME);
  Logger.log('Resuming — ' + progress.rowsWritten + ' rows written so far.');

  while (progress.nextUrl) {
    if (Date.now() - startTime > CI_MAX_RUN_MS) {
      saveCiProgress_(props, progress);
      Logger.log('Time limit reached — saved at ' + progress.rowsWritten
        + ' rows. Trigger will resume in ~5 min.');
      return;
    }

    const resp = pcoGetHistory(progress.nextUrl, auth);

    // Build in-memory lookups from the included array for this page.
    const personMap  = {};   // person_id  → {first, last}
    const periodMap  = {};   // period_id  → starts_at
    const eventMap   = {};   // event_id   → name

    for (const inc of (resp.included || [])) {
      if (inc.type === 'Person') {
        personMap[inc.id] = {
          first: inc.attributes.first_name || '',
          last:  inc.attributes.last_name  || '',
        };
      } else if (inc.type === 'EventPeriod') {
        periodMap[inc.id] = inc.attributes.starts_at || '';
      } else if (inc.type === 'Event') {
        eventMap[inc.id] = inc.attributes.name || '';
      }
    }

    const batch = [];
    for (const ci of (resp.data || [])) {
      const a        = ci.attributes;
      const rel      = ci.relationships || {};
      const personId = rel.person       && rel.person.data       ? rel.person.data.id       : null;
      const periodId = rel.event_period && rel.event_period.data ? rel.event_period.data.id : null;
      const eventId  = rel.event        && rel.event.data        ? rel.event.data.id        : null;

      const person     = (personId && personMap[personId]) || null;
      const startsAt   = (periodId && periodMap[periodId]) || '';
      const eventName  = (eventId  && eventMap[eventId])   || '';

      batch.push([
        ci.id,
        eventName,
        eventId  || '',
        periodId || '',
        startsAt,
        person ? person.first : (a.first_name || ''),
        person ? person.last  : (a.last_name  || ''),
        a.kind                           || '',
        a.created_at                     || '',
        a.checked_out_at                 || '',
        a.one_time_guest ? 'true' : 'false',
        a.security_code                  || '',
        a.number != null ? a.number : '',
        a.emergency_contact_name         || '',
        a.emergency_contact_phone_number || '',
        a.medical_notes                  || '',
      ]);
    }

    if (batch.length) {
      sh.getRange(sh.getLastRow() + 1, 1, batch.length, CI_HEADERS.length).setValues(batch);
      progress.rowsWritten += batch.length;
    }

    // Advance to next page or finish.
    const nextLink   = resp.links && resp.links.next;
    progress.nextUrl = nextLink || null;
    saveCiProgress_(props, progress);
  }

  // Fell through the while loop — all pages consumed.
  progress.status = 'complete';
  saveCiProgress_(props, progress);
  deleteCiTrigger_();
  Logger.log('COMPLETE — all ' + progress.rowsWritten + ' check-in records written.');
}

/**
 * Clears saved progress and the tab. Run before starting a fresh backfill.
 */
function pcoCheckInsHistoryReset() {
  deleteCiTrigger_();
  PropertiesService.getScriptProperties().deleteProperty(CI_HISTORY_KEY);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CI_TAB_NAME);
  if (sh) sh.clear();
  Logger.log('Reset complete. Run pcoCheckInsHistoryStart() to begin the backfill.');
}

// ─── Trigger management ───────────────────────────────────────────────────────

function deleteCiTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'pcoCheckInsHistoryRun')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ─── Progress persistence ─────────────────────────────────────────────────────

function loadCiProgress_(props) {
  const raw = props.getProperty(CI_HISTORY_KEY);
  return raw ? JSON.parse(raw) : {};
}

function saveCiProgress_(props, progress) {
  props.setProperty(CI_HISTORY_KEY, JSON.stringify(progress));
}
