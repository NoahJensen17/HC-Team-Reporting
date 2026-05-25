/**
 * Planning Center -> Google Sheets pull (Apps Script).
 *
 * Incremental upsert strategy: on every run, only records created or updated
 * since the last run are fetched from PCO. Existing rows are updated in place
 * if changed; new rows are appended. Nothing is ever cleared wholesale.
 *
 * Designed to run every 5 minutes between 5am–9pm. Most runs take only a few
 * seconds because PCO returns near-zero records when nothing has changed.
 *
 * SETUP:
 *   1. Open the target Google Sheet → Extensions → Apps Script.
 *   2. Paste this file as Code.gs (add the history script as a second file).
 *   3. Project Settings → Script Properties → add:
 *        PCO_APP_ID  = <your app id>
 *        PCO_SECRET  = <your secret>
 *   4. Run pcoPullAll() once manually to approve OAuth scopes and seed data.
 *   5. Run setupSyncTrigger() to start the every-5-minute schedule.
 *
 * INITIAL SERVICES_SCHEDULING LOAD:
 *   Run pcoSchedulingHistoryStart() from the companion script BEFORE running
 *   pcoPullAll() for the first time. After that, pcoPullAll() handles
 *   incremental updates automatically.
 *
 * Column order is locked after the first run. Hide any column and it stays
 * hidden on every subsequent run. New PCO fields are appended on the right.
 * To reset the locked order for a tab: run resetTabHeaders('Tab Name').
 */

const PCO_HOST = 'https://api.planningcenteronline.com';

const PCO_SOURCES = [
  { tab: 'People',          path: '/people/v2/people?per_page=100' },
  { tab: 'Households',      path: '/people/v2/households?per_page=100' },
  { tab: 'Services_Plans',  path: '/services/v2/service_types?per_page=100' },
  { tab: 'CheckIns_Events', path: '/check-ins/v2/events?per_page=100' },
  { tab: 'Groups',          path: '/groups/v2/groups?per_page=100' },
  { tab: 'Group_Types',     path: '/groups/v2/group_types?per_page=100' },
  { tab: 'Registrations',   path: '/registrations/v2/signups?per_page=100' },
  // Calendar, Giving, and Publishing intentionally omitted.
];

const HEADER_KEY_PREFIX = 'headers_';
const LAST_RUN_KEY      = 'pco_last_run_at';
const ALERT_EMAIL       = 'ninjanobe@gmail.com';

// ─── Trigger setup ────────────────────────────────────────────────────────────

/**
 * Run once to set up the :00/:30 aligned sync cadence.
 * Schedules a one-time anchor trigger at the next :00 or :30 mark, which then
 * creates the recurring every-30-minute trigger from that aligned start point.
 */
function setupSyncTrigger() {
  // Clear any existing sync and alignment triggers.
  ScriptApp.getProjectTriggers()
    .filter(t => ['pcoPullAll', 'syncAligner_'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Find the next :00 or :30 boundary.
  const now    = new Date();
  const anchor = new Date(now);
  anchor.setSeconds(0, 0);
  const m = now.getMinutes();
  anchor.setMinutes(m < 30 ? 30 : 60); // :30 this hour, or :00 next hour (minute 60 rolls over)

  ScriptApp.newTrigger('syncAligner_').timeBased().at(anchor).create();
  Logger.log('Sync aligner scheduled for ' + anchor.toLocaleTimeString()
    + '. Recurring :00/:30 cadence (4am–10pm CST) starts then.');
}

/**
 * One-time trigger fired at the first :00 or :30 boundary after setupSyncTrigger().
 * Deletes itself, then creates the recurring 30-minute trigger anchored to this moment.
 */
function syncAligner_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncAligner_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pcoPullAll').timeBased().everyMinutes(30).create();
  Logger.log('Recurring sync trigger created from aligned start: ' + new Date());
  pcoPullAll();
}

// Run to stop the scheduled sync completely.
function removeSyncTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['pcoPullAll', 'syncAligner_'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Sync triggers removed.');
}

// ─── Main entry point ────────────────────────────────────────────────────────

function pcoPullAll() {
  // Only run between 4:00am and 10:00pm CST (America/Chicago).
  const nowCST = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Chicago' }));
  const hour   = nowCST.getHours();
  if (hour < 4 || hour >= 22) {
    Logger.log('Outside active window (4am–10pm CST) — skipping sync.');
    return;
  }

  const errors = [];

  try {
    const props  = PropertiesService.getScriptProperties();
    const appId  = props.getProperty('PCO_APP_ID');
    const secret = props.getProperty('PCO_SECRET');
    if (!appId || !secret) throw new Error('PCO_APP_ID or PCO_SECRET missing from Script Properties.');
    const auth = 'Basic ' + Utilities.base64Encode(appId + ':' + secret);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    // ISO timestamp from the last successful run.
    // When set, every fetch filters to only records updated since then —
    // most 5-minute runs will return 0 records and finish in seconds.
    const lastRunAt = props.getProperty(LAST_RUN_KEY) || '';

    // Flat endpoints — incremental fetch by updated_at, upsert by id.
    for (const src of PCO_SOURCES) {
      try {
        const rows = fetchPaged(src.path, auth, lastRunAt);
        upsertTab(ss, src.tab, rows, 'id');
      } catch (e) {
        Logger.log(src.tab + ' FAILED: ' + e.message);
        errors.push('• ' + src.tab + ': ' + e.message);
      }
    }

    // CheckIns_People — incremental fetch by updated_at, upsert by ci_id.
    try {
      if (!lastRunAt) {
        Logger.log('CheckIns_People: no prior run timestamp — skipping. '
          + 'Run pcoCheckInsHistoryStart() for the initial load first.');
      } else {
        const rows = fetchCheckInsPeople(auth, lastRunAt);
        upsertTab(ss, 'CheckIns_People', rows, 'ci_id');
      }
    } catch (e) {
      Logger.log('CheckIns_People FAILED: ' + e.message);
      errors.push('• CheckIns_People: ' + e.message);
    }

    // Services_Scheduling — incremental fetch by updated_at, upsert by pp_id.
    try {
      if (!lastRunAt) {
        Logger.log('Services_Scheduling: no prior run timestamp — skipping incremental sync. '
          + 'Run pcoSchedulingHistoryStart() for the initial load first.');
      } else {
        const rows = fetchServicesScheduling(auth);
        upsertTab(ss, 'Services_Scheduling', rows, 'pp_id');
      }
    } catch (e) {
      Logger.log('Services_Scheduling FAILED: ' + e.message);
      errors.push('• Services_Scheduling: ' + e.message);
    }

    // Save timestamp so the next run only fetches changes since right now.
    props.setProperty(LAST_RUN_KEY, new Date().toISOString());

  } catch (fatal) {
    errors.push('• Fatal: ' + fatal.message);
  }

  if (errors.length) {
    const body = 'PCO sync ran into ' + errors.length + ' error(s) on '
      + new Date().toLocaleString() + ':\n\n'
      + errors.join('\n')
      + '\n\nOpen the Apps Script editor → Executions to see the full log.';
    MailApp.sendEmail(ALERT_EMAIL, 'PCO Sync Error', body);
  }
}

// ─── One-time data fix ────────────────────────────────────────────────────────

// The history backfill wrote 15 columns (no pp_id). The main sync later added
// pp_id as column 1, shifting the header forward by one. This left ~22k historical
// rows one column off. Run this ONCE to prepend a blank pp_id to every shifted row.
//
// A row is "shifted" if column A contains letters (a service type name) rather
// than a numeric PCO ID. Safe to re-run — already-fixed rows are left untouched.
function fixSchedulingHistoryShift() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Services_Scheduling');
  if (!sh) { Logger.log('Services_Scheduling tab not found.'); return; }

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  if (headers[0] !== 'pp_id') {
    Logger.log('Unexpected header[0]: "' + headers[0] + '" — aborting for safety.');
    return;
  }

  const ncols = headers.length; // 16
  let fixed = 0;

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const ppId = row[0] ? row[0].toString().trim() : '';
    // Shifted rows have a service type name (letters) in col A instead of a numeric PCO ID.
    if (ppId && !/^\d+$/.test(ppId)) {
      // Prepend blank pp_id and keep the first (ncols-1) existing values.
      data[i] = [''].concat(row.slice(0, ncols - 1));
      fixed++;
    }
  }

  if (fixed === 0) {
    Logger.log('No shifted rows found — data already looks correct.');
    return;
  }

  // Write all data rows back in one call.
  sh.getRange(2, 1, data.length - 1, ncols).setValues(data.slice(1));
  Logger.log('fixSchedulingHistoryShift: corrected ' + fixed + ' shifted rows.');
}

// Removes duplicate rows from Services_Scheduling.
//
// Two rows are duplicates when they share the same plan_id + person_name +
// team + position (the four fields that uniquely identify a single PCO
// plan_people assignment). Status is intentionally excluded so that a
// historical "Unconfirmed" row is still treated as a duplicate of a current
// "Confirmed" row for the same assignment.
//
// When duplicates are found, the row with a numeric pp_id is kept (current
// sync data). If both rows have numeric pp_ids, the first one in the sheet
// is kept. Rows with no plan_id or person_name are left untouched.
//
// Rewrites the entire sheet in one call — much faster than row-by-row
// deletion for 40k+ rows.
function deduplicateServicesScheduling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Services_Scheduling');
  if (!sh) { Logger.log('Services_Scheduling tab not found.'); return; }

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const col     = {};
  headers.forEach((h, i) => col[h.toString().trim()] = i);

  const ci   = col['pp_id'];
  const cpid = col['person_id'];
  const cp   = col['plan_id'];
  const cn   = col['person_name'];
  const ct   = col['team'];
  const cpos = col['position'];

  if ([ci, cp, cn, ct, cpos].some(x => x === undefined)) {
    Logger.log('Missing required column. Headers found: ' + headers.join(', '));
    return;
  }

  // Build a dedup key for each row.
  // Rows with a numeric pp_id use it directly as the key — it is the canonical
  // PCO unique identifier for a person-on-a-plan assignment.
  // Rows missing pp_id fall back to the composite plan_id+person_id+person_name+team+position.
  function rowKey(row) {
    const ppId = (row[ci] || '').toString().trim();
    if (/^\d+$/.test(ppId)) return 'pp:' + ppId;
    const pid  = cpid !== undefined ? (row[cpid] || '').toString().trim() : '';
    const base = pid
      ? [row[cp], pid, row[ct], row[cpos]]
      : [row[cp], row[cn], row[ct], row[cpos]];
    return 'comp:' + base.map(v => (v || '').toString().trim()).join('|');
  }

  // First pass: for each key keep the "best" row.
  // Best = has a numeric pp_id; ties keep the first occurrence.
  const best = {};
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const planId     = (row[cp] || '').toString().trim();
    const personName = (row[cn] || '').toString().trim();
    if (!planId || !personName) continue;

    const key    = rowKey(row);
    const ppId   = (row[ci] || '').toString().trim();
    const hasNum = /^\d+$/.test(ppId);

    if (best[key] === undefined) {
      best[key] = i;
    } else {
      const prevHasNum = /^\d+$/.test((data[best[key]][ci] || '').toString().trim());
      if (!prevHasNum && hasNum) best[key] = i;
    }
  }

  // Second pass: keep only best-row indices, plus unkeyed rows.
  const keepIdx = new Set(Object.values(best));
  const kept    = [headers];
  let   removed = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const planId     = (row[cp] || '').toString().trim();
    const personName = (row[cn] || '').toString().trim();
    if (!planId || !personName) { kept.push(row); continue; }
    if (keepIdx.has(i)) { kept.push(row); } else { removed++; }
  }

  if (removed === 0) {
    Logger.log('deduplicateServicesScheduling: no duplicates found.');
    return;
  }

  Logger.log('Removing ' + removed + ' duplicate rows — rewriting sheet with '
    + (kept.length - 1) + ' data rows...');
  sh.clearContents();
  sh.getRange(1, 1, kept.length, headers.length).setValues(kept);
  sh.setFrozenRows(1);
  Logger.log('deduplicateServicesScheduling: done. Removed ' + removed + ' duplicates.');
}

// ─── Reset helpers ────────────────────────────────────────────────────────────

function resetTabHeaders(tabName) {
  PropertiesService.getScriptProperties().deleteProperty(HEADER_KEY_PREFIX + tabName);
  Logger.log('Header order reset for tab: ' + tabName);
}
function resetSchedulingHeaders()  { resetTabHeaders('Services_Scheduling'); }
function resetCheckInPeopleHeaders(){ resetTabHeaders('CheckIns_People'); }

// Clears the last-run timestamp. The next pcoPullAll() will do a full fetch
// of all tabs (except Services_Scheduling, which needs the history script).
function resetLastRunAt() {
  PropertiesService.getScriptProperties().deleteProperty(LAST_RUN_KEY);
  Logger.log('Last-run timestamp cleared. Next run will do a full fetch.');
}

// Sets the last-run timestamp to right now so the next pcoPullAll() only
// picks up changes from this point forward. Run this once after the initial
// history backfill is complete to start incremental syncing cleanly.
function setLastRunAtToNow() {
  const ts = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(LAST_RUN_KEY, ts);
  Logger.log('pco_last_run_at set to: ' + ts);
}

// ─── Fetch ───────────────────────────────────────────────────────────────────

// When `since` is provided, appends ?where[updated_at][gte]=since so only
// recently changed records are returned. Returns [] on a quiet 5-min window.
function fetchPaged(startPath, auth, since) {
  const out  = [];
  const filter = since ? '&where[updated_at][gte]=' + encodeURIComponent(since) : '';
  let   path = startPath + filter;
  while (path) {
    const json = pcoGet(PCO_HOST + path, auth);
    for (const rec of (json.data || [])) out.push(flatten(rec));
    const next = json.links && json.links.next;
    path = next ? next.replace(PCO_HOST, '') : null;
  }
  return out;
}

// Fetches event_periods across all check-in events.
// When `since` is set, only periods updated after that timestamp are returned.
function fetchCheckInAttendance(auth, since) {
  const out    = [];
  const filter = since ? '&where[updated_at][gte]=' + encodeURIComponent(since) : '';

  let evPath = '/check-ins/v2/events?per_page=100';
  while (evPath) {
    const evResp = pcoGet(PCO_HOST + evPath, auth);
    for (const ev of (evResp.data || [])) {
      const eventName = ev.attributes.name || ev.id;

      let pPath = '/check-ins/v2/events/' + ev.id
        + '/event_periods?per_page=100&order=-starts_at' + filter;
      while (pPath) {
        const pResp = pcoGet(PCO_HOST + pPath, auth);
        for (const ep of (pResp.data || [])) {
          const regular   = ep.attributes.regular_count   || 0;
          const guest     = ep.attributes.guest_count     || 0;
          const volunteer = ep.attributes.volunteer_count || 0;
          const total     = regular + guest + volunteer;
          if (!total) continue;
          out.push({
            event_name:      eventName,
            event_id:        ev.id,
            period_id:       ep.id,
            starts_at:       ep.attributes.starts_at || '',
            ends_at:         ep.attributes.ends_at   || '',
            regular_count:   regular,
            guest_count:     guest,
            volunteer_count: volunteer,
            total_count:     total,
          });
        }
        const pNext = pResp.links && pResp.links.next;
        pPath = pNext ? pNext.replace(PCO_HOST, '') : null;
      }
    }
    const evNext = evResp.links && evResp.links.next;
    evPath = evNext ? evNext.replace(PCO_HOST, '') : null;
  }
  return out;
}

// Fetches individual check-in records updated since `since`.
// Incremental runs are small so include=person,event_period,event is safe.
function fetchCheckInsPeople(auth, since) {
  const out    = [];
  const filter = since ? '&where[updated_at][gte]=' + encodeURIComponent(since) : '';
  let   path   = '/check-ins/v2/check_ins?per_page=100&include=person,event_period,event' + filter;

  while (path) {
    const resp = pcoGet(PCO_HOST + path, auth);

    const personMap = {};
    const periodMap = {};
    const eventMap  = {};
    for (const inc of (resp.included || [])) {
      if (inc.type === 'Person')      personMap[inc.id] = { first: inc.attributes.first_name || '', last: inc.attributes.last_name || '' };
      else if (inc.type === 'EventPeriod') periodMap[inc.id] = inc.attributes.starts_at || '';
      else if (inc.type === 'Event')  eventMap[inc.id]  = inc.attributes.name || '';
    }

    for (const ci of (resp.data || [])) {
      const a        = ci.attributes;
      const rel      = ci.relationships || {};
      const personId = rel.person       && rel.person.data       ? rel.person.data.id       : null;
      const periodId = rel.event_period && rel.event_period.data ? rel.event_period.data.id : null;
      const eventId  = rel.event        && rel.event.data        ? rel.event.data.id        : null;
      const person   = (personId && personMap[personId]) || null;

      out.push({
        ci_id:            ci.id,
        person_id:        personId  || '',
        event_name:       (eventId  && eventMap[eventId])  || '',
        event_id:         eventId   || '',
        period_id:        periodId  || '',
        period_starts_at: (periodId && periodMap[periodId]) || '',
        first_name:       person ? person.first : (a.first_name || ''),
        last_name:        person ? person.last  : (a.last_name  || ''),
        kind:             a.kind             || '',
        checked_in_at:    a.created_at       || '',
        checked_out_at:   a.checked_out_at   || '',
        one_time_guest:   a.one_time_guest ? 'true' : 'false',
      });
    }

    const next = resp.links && resp.links.next;
    path = next ? next.replace(PCO_HOST, '') : null;
  }
  return out;
}

// Fetches plan_people for all upcoming plans across every team.
//
// Always runs a full parallel fetch — PCO exposes no endpoint that reliably
// filters plan_people by updated_at, so there is no fast gate to detect
// volunteer confirm/decline changes without downloading the full future window.
// UrlFetchApp.fetchAll fires all 55 team requests simultaneously so the sweep
// completes in ~20-30 s regardless of history size. Client-side date filter
// discards past records so historical rows are never re-appended.
function fetchServicesScheduling(auth) {
  const out         = [];
  const statusLabel = { C: 'Confirmed', U: 'Unconfirmed', D: 'Declined' };
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);
  const cutoffStr = cutoff.toISOString().slice(0, 10); // 7 days ago, 'YYYY-MM-DD'

  // Fetch all service types.
  const allSt = [];
  let stPath = '/services/v2/service_types?per_page=100';
  while (stPath) {
    const r = pcoGet(PCO_HOST + stPath, auth);
    allSt.push(...(r.data || []));
    const next = r.links && r.links.next;
    stPath = next ? next.replace(PCO_HOST, '') : null;
  }

  // Collect all teams then fetch plan_people in parallel.
  const pending = []; // { url, stName, teamName }
  for (const st of allSt) {
    const stName = st.attributes.name || st.id;
    let teamPath = '/services/v2/service_types/' + st.id + '/teams?per_page=100';
    while (teamPath) {
      const r = pcoGet(PCO_HOST + teamPath, auth);
      for (const team of (r.data || [])) {
        pending.push({
          url:      PCO_HOST + '/services/v2/service_types/' + st.id
                    + '/teams/' + team.id + '/plan_people?include=plan&per_page=100',
          stName:   stName,
          teamName: team.attributes.name || team.id,
        });
      }
      const next = r.links && r.links.next;
      teamPath = next ? next.replace(PCO_HOST, '') : null;
    }
  }
  Logger.log('Services_Scheduling: ' + allSt.length + ' service type(s), '
    + pending.length + ' teams — fetching plan_people in parallel...');

  // Shared request options (URL varies per request).
  const baseOpts = {
    method:             'get',
    headers:            { Authorization: auth, Accept: 'application/json' },
    muteHttpExceptions: true,
    followRedirects:    false,
  };

  // Process pages in parallel batches until all queues are empty.
  while (pending.length > 0) {
    const batch = pending.splice(0, pending.length); // take everything queued
    const reqs  = batch.map(b => Object.assign({ url: b.url }, baseOpts));
    const resps = UrlFetchApp.fetchAll(reqs);

    for (let i = 0; i < resps.length; i++) {
      const code  = resps[i].getResponseCode();
      const body  = resps[i].getContentText();
      const item  = batch[i];

      if (code === 429) {
        Utilities.sleep(2000);
        pending.push(item); // retry next round
        continue;
      }
      if (code === 302) {
        let loc = '';
        try { loc = JSON.parse(body).location || ''; } catch (e) {}
        if (loc) pending.push({ url: loc, stName: item.stName, teamName: item.teamName });
        continue;
      }
      if (code !== 200) {
        Logger.log('Services_Scheduling: HTTP ' + code + ' for team ' + item.teamName + ' — skipping');
        continue;
      }

      const json     = JSON.parse(body);
      const planData = {};
      for (const inc of (json.included || [])) {
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

      for (const pp of (json.data || [])) {
        const planId = pp.relationships && pp.relationships.plan && pp.relationships.plan.data
          ? pp.relationships.plan.data.id : null;
        const meta   = (planId && planData[planId]) || {};

        // Keep plans from the past 7 days forward (volunteers may confirm/decline after service).
        if (meta.date && meta.date < cutoffStr) continue;

        const p        = pp.attributes;
        const personId = pp.relationships && pp.relationships.person && pp.relationships.person.data
          ? pp.relationships.person.data.id : null;
        out.push({
          pp_id:             pp.id,
          person_id:         personId    || '',
          service_type:      item.stName,
          plan_date:         meta.date   || '',
          plan_dates:        meta.dates  || '',
          plan_title:        meta.title  || '',
          series_title:      meta.series || '',
          plan_id:           planId      || '',
          person_name:       p.name                         || '',
          status:            statusLabel[p.status]          || p.status || '',
          team:              item.teamName,
          position:          p.team_position_name           || '',
          responds_to:       p.responds_to_name             || '',
          decline_reason:    p.decline_reason               || '',
          notes:             p.notes                        || '',
          status_updated_at: p.status_updated_at            || '',
          prepared_at:       p.prepare_notification_sent_at || '',
        });
      }

      // Queue next page if present.
      const nextUrl = json.links && json.links.next;
      if (nextUrl) pending.push({ url: nextUrl, stName: item.stName, teamName: item.teamName });
    }

    if (pending.length > 0) Utilities.sleep(500); // brief pause between rounds
  }

  Logger.log('Services_Scheduling: fetched ' + out.length + ' future assignment(s).');
  return out;
}

function pcoGet(url, auth, depth) {
  depth = depth || 0;
  if (depth > 5) throw new Error('Too many redirects: ' + url);
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: auth, Accept: 'application/json' },
    muteHttpExceptions: true,
    followRedirects: false,
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code === 429) { Utilities.sleep(3000); return pcoGet(url, auth, depth); }
  if (code === 302) {
    let loc = '';
    try { loc = JSON.parse(body).location || ''; } catch (e) {}
    if (!loc) throw new Error('302 with no location body for ' + url);
    return pcoGet(loc, auth, depth + 1);
  }
  if (code !== 200) throw new Error('PCO ' + code + ' ' + url + ': ' + body.slice(0, 200));
  return JSON.parse(body);
}

// ─── Flatten ─────────────────────────────────────────────────────────────────

function flatten(rec) {
  const row = { id: rec.id, type: rec.type };
  const attrs = rec.attributes || {};
  for (const k of Object.keys(attrs)) {
    const v = attrs[k];
    row[k] = (v && typeof v === 'object') ? JSON.stringify(v) : v;
  }
  return row;
}

// ─── Upsert ───────────────────────────────────────────────────────────────────
//
// Never clears the sheet. Rows whose keyField value already exists are updated
// in place only when a value has changed. New key values are appended.
// Column order is stable (locked after first run; new PCO fields go on the right).

function upsertTab(ss, name, rows, keyField) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (!rows.length) {
    Logger.log(name + ': nothing to upsert.');
    return;
  }

  // Build stable header order.
  const props = PropertiesService.getScriptProperties();
  const hKey  = HEADER_KEY_PREFIX + name;
  const stored = props.getProperty(hKey);
  let lockedHeaders = stored ? JSON.parse(stored) : [];

  // If no stored order but the sheet already has data (e.g. written by the
  // history backfill script), bootstrap from the existing header row so we
  // don't overwrite it.
  if (!stored && sh.getLastRow() > 0) {
    const sheetRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    lockedHeaders = sheetRow.filter(h => h !== '' && h != null);
  }

  const incomingFields = Array.from(rows.reduce((s, r) => {
    Object.keys(r).forEach(k => s.add(k)); return s;
  }, new Set()));
  const lockedSet = new Set(lockedHeaders);
  const headers = lockedHeaders.concat(incomingFields.filter(f => !lockedSet.has(f)));
  props.setProperty(hKey, JSON.stringify(headers));

  const keyCol = headers.indexOf(keyField);

  // Write / refresh header row.
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const lastRow = sh.getLastRow();

  // Empty sheet or key column not yet in headers — write everything as a batch.
  if (lastRow <= 1 || keyCol === -1) {
    const matrix = rows.map(r => headers.map(h => r[h] == null ? '' : r[h]));
    if (matrix.length) sh.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
    Logger.log(name + ': wrote ' + matrix.length + ' rows (first run)');
    return;
  }

  // Read all existing rows once to build key → row-index map.
  const numDataRows  = lastRow - 1;
  const numCols      = Math.max(sh.getLastColumn(), headers.length);
  const existingData = sh.getRange(2, 1, numDataRows, numCols).getValues();

  const idToIdx = {};
  existingData.forEach((row, i) => {
    const v = String(row[keyCol] == null ? '' : row[keyCol]);
    if (v !== '') idToIdx[v] = i;
  });

  const toUpdate = []; // { sheetRow (1-based), values }
  const toAppend = [];

  for (const row of rows) {
    const values = headers.map(h => row[h] == null ? '' : row[h]);
    const key    = String(row[keyField] == null ? '' : row[keyField]);

    if (key !== '' && idToIdx[key] !== undefined) {
      const existing = existingData[idToIdx[key]].slice(0, headers.length);
      while (existing.length < headers.length) existing.push('');
      if (values.some((v, i) => String(v) !== String(existing[i]))) {
        toUpdate.push({ sheetRow: idToIdx[key] + 2, values });
      }
    } else {
      toAppend.push(values);
    }
  }

  for (const u of toUpdate) {
    sh.getRange(u.sheetRow, 1, 1, headers.length).setValues([u.values]);
  }
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
  }

  Logger.log(name + ': ' + toUpdate.length + ' updated, ' + toAppend.length + ' new');
}

// ─── Check-Ins per-person test ────────────────────────────────────────────────

// Run this ONCE to verify per-person check-in access.
// Uses the most recent event period and tries two endpoint approaches.
// Check the Executions log to see what fields come back.
function testCheckInsPeople() {
  const props  = PropertiesService.getScriptProperties();
  const appId  = props.getProperty('PCO_APP_ID');
  const secret = props.getProperty('PCO_SECRET');
  if (!appId || !secret) throw new Error('Set PCO_APP_ID and PCO_SECRET in Script Properties.');
  const auth = 'Basic ' + Utilities.base64Encode(appId + ':' + secret);

  // Step 1: find the most recently updated event.
  const evResp = pcoGet(PCO_HOST + '/check-ins/v2/events?per_page=10&order=-updated_at', auth);
  const events = evResp.data || [];
  if (!events.length) { Logger.log('No events found.'); return; }
  Logger.log('Most recent events:');
  events.forEach(e => Logger.log('  ' + e.id + ' — ' + e.attributes.name));

  // Use the first (most recently updated) event.
  const ev = events[0];
  Logger.log('\nUsing event: ' + ev.attributes.name + ' (id ' + ev.id + ')');

  // Step 2: grab the most recent period for that event.
  const pResp = pcoGet(
    PCO_HOST + '/check-ins/v2/events/' + ev.id + '/event_periods?per_page=5&order=-starts_at', auth
  );
  const periods = pResp.data || [];
  if (!periods.length) { Logger.log('No periods found for this event.'); return; }
  Logger.log('Most recent periods:');
  periods.forEach(p => Logger.log('  ' + p.id + ' — ' + p.attributes.starts_at
    + '  regular=' + p.attributes.regular_count
    + ' volunteer=' + p.attributes.volunteer_count));

  const period = periods[0];
  Logger.log('\nUsing period: ' + period.attributes.starts_at + ' (id ' + period.id + ')');

  // Step 3a: try event_period sub-resource.
  Logger.log('\n--- Approach A: /event_periods/{id}/check_ins ---');
  try {
    const ciResp = pcoGet(
      PCO_HOST + '/check-ins/v2/event_periods/' + period.id + '/check_ins?per_page=5&include=person',
      auth
    );
    Logger.log('Total in period: ' + (ciResp.meta && ciResp.meta.total_count));
    logCheckInSample_(ciResp);
  } catch (e) {
    Logger.log('Approach A failed: ' + e.message);
  }

  // Step 3b: try top-level check_ins filtered by event_period_id.
  Logger.log('\n--- Approach B: /check_ins?where[event_period_id]={id} ---');
  try {
    const ciResp2 = pcoGet(
      PCO_HOST + '/check-ins/v2/check_ins?where[event_period_id]=' + period.id
        + '&per_page=5&include=person',
      auth
    );
    Logger.log('Total: ' + (ciResp2.meta && ciResp2.meta.total_count));
    logCheckInSample_(ciResp2);
  } catch (e) {
    Logger.log('Approach B failed: ' + e.message);
  }
}

function logCheckInSample_(ciResp) {
  const people = {};
  for (const inc of (ciResp.included || [])) {
    if (inc.type === 'Person') {
      people[inc.id] = (inc.attributes.first_name || '') + ' ' + (inc.attributes.last_name || '');
    }
  }
  for (const ci of (ciResp.data || [])) {
    const a = ci.attributes;
    const personId = ci.relationships && ci.relationships.person && ci.relationships.person.data
      ? ci.relationships.person.data.id : null;
    const name = (personId && people[personId]) || a.name || '(anonymous)';
    Logger.log('  id=' + ci.id + ' name=' + name + ' kind=' + a.kind
      + ' checked_in_at=' + a.created_at + ' attrs=' + JSON.stringify(a));
  }
}
