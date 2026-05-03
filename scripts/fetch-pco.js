#!/usr/bin/env node
// Fetches PCO Services + Check-Ins at deploy time and writes pco-data.js.
// Run via GitHub Actions (no CORS restrictions). Reads credentials from env.

const https = require('https');
const fs = require('fs');
const path = require('path');

const APP_ID = process.env.PCO_APP_ID;
const SECRET = process.env.PCO_SECRET;

if (!APP_ID || !SECRET) {
  console.log('PCO: credentials not set — writing empty pco-data.js');
  fs.writeFileSync(path.join(__dirname, '..', 'pco-data.js'),
    'window._PCO_STATIC={servicesByDate:{},checkInsByDate:{}};\n');
  process.exit(0);
}

function pcoGet(urlPath) {
  return new Promise((resolve, reject) => {
    const auth = Buffer.from(`${APP_ID}:${SECRET}`).toString('base64');
    const opts = {
      hostname: 'api.planningcenteronline.com',
      path: urlPath,
      headers: { 'Authorization': `Basic ${auth}`, 'Accept': 'application/json' }
    };
    https.get(opts, res => {
      let body = '';
      res.on('data', c => body += c);
      res.on('end', () => {
        if (res.statusCode === 429) {
          setTimeout(() => pcoGet(urlPath).then(resolve).catch(reject), 3000);
          return;
        }
        // PCO returns 302 with redirect location in JSON body (not HTTP Location header)
        if (res.statusCode === 302) {
          try {
            const loc = JSON.parse(body).location;
            if (loc) {
              const redirectPath = loc.replace('https://api.planningcenteronline.com', '');
              pcoGet(redirectPath).then(resolve).catch(reject);
              return;
            }
          } catch (e) {}
          return reject(new Error(`PCO 302 with no location: ${body.slice(0, 200)}`));
        }
        if (res.statusCode !== 200) return reject(new Error(`PCO ${res.statusCode} ${urlPath}: ${body.slice(0,200)}`));
        try { resolve(JSON.parse(body)); } catch (e) { reject(e); }
      });
    }).on('error', reject);
  });
}

function dateKey(iso) {
  if (!iso) return null;
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return `${parseInt(m[2])}/${parseInt(m[3])}/${m[1]}`;
}

async function main() {
  const servicesByDate = {};
  const checkInsByDate = {};

  // --- Services + Scheduling ---
  // Fetch plans for each service type. For plans within the last 6 months,
  // also fetch plan_people (with team names) to show per-person scheduling status.
  try {
    const currentYear = new Date().getFullYear();

    // Fetch all service types (no page limit — main Sunday type may be beyond first 8)
    let typesPath = '/services/v2/service_types?per_page=100';
    const allTypes = [];
    while (typesPath) {
      const tr = await pcoGet(typesPath);
      allTypes.push(...(tr.data || []));
      const next = tr.links && tr.links.next;
      typesPath = next ? next.replace('https://api.planningcenteronline.com', '') : null;
    }
    console.log(`PCO Services: found ${allTypes.length} service types`);
    let schedulingFetched = 0;

    for (const st of allTypes) {
      try {
        const plansResp = await pcoGet(
          `/services/v2/service_types/${st.id}/plans?filter=past&per_page=50&order=-sort_date`
        );
        for (const plan of (plansResp.data || [])) {
          const key = dateKey(plan.attributes.sort_date);
          if (!key) continue;

          const entry = { serviceType: st.attributes.name, title: plan.attributes.title || '', people: [] };

          // Fetch per-person scheduling for plans in the last ~12 months (current year or prior year)
          const planYear = parseInt(key.split('/')[2]);
          if (planYear >= currentYear - 1) {
            try {
              const ppResp = await pcoGet(
                `/services/v2/plans/${plan.id}/plan_people?include=team&per_page=100`
              );
              // Build team name lookup from JSON:API included array
              const teamNames = {};
              for (const inc of (ppResp.included || [])) {
                if (inc.type === 'Team') teamNames[inc.id] = inc.attributes.name;
              }
              entry.people = (ppResp.data || [])
                .filter(p => p.attributes.status) // skip unscheduled/declined-before-scheduling
                .map(p => ({
                  name: p.attributes.name || '',
                  status: p.attributes.status,           // 'C'=confirmed, 'U'=unconfirmed, 'D'=declined
                  position: p.attributes.team_position_name || '',
                  team: (p.relationships && p.relationships.team && p.relationships.team.data
                    ? teamNames[p.relationships.team.data.id] : '') || ''
                }))
                .filter(p => p.name);
              schedulingFetched++;
            } catch (e) { console.warn(`    plan_people ${plan.id} failed:`, e.message); }
          }

          if (!servicesByDate[key]) servicesByDate[key] = [];
          servicesByDate[key].push(entry);
        }
      } catch (e) { console.warn(`  Services type ${st.id} error:`, e.message); }
    }
    console.log(`PCO Services: ${Object.keys(servicesByDate).length} dates, scheduling fetched for ${schedulingFetched} recent plans`);
  } catch (e) { console.warn('PCO Services fetch failed:', e.message); }

  // --- Check-Ins ---
  // Paginate through ALL events and collect all event_periods with non-zero attendance.
  // No server-side date filtering — the app handles date windowing at render time.
  try {
    const allEvents = [];
    let evPath = '/check-ins/v2/events?per_page=100';
    while (evPath) {
      const evResp = await pcoGet(evPath);
      allEvents.push(...(evResp.data || []));
      const next = evResp.links && evResp.links.next;
      evPath = next ? next.replace('https://api.planningcenteronline.com', '') : null;
    }
    console.log(`PCO Check-Ins: scanning ${allEvents.length} events...`);

    let activeEvents = 0;
    for (const ev of allEvents) {
      try {
        const periodsResp = await pcoGet(
          `/check-ins/v2/events/${ev.id}/event_periods?per_page=100&order=-starts_at`
        );
        const periods = periodsResp.data || [];
        let gotData = false;
        for (const ep of periods) {
          const key = dateKey(ep.attributes.starts_at);
          if (!key) continue;
          const total = (ep.attributes.regular_count || 0)
            + (ep.attributes.guest_count || 0)
            + (ep.attributes.volunteer_count || 0);
          if (!total) continue;
          if (!checkInsByDate[key]) checkInsByDate[key] = { total: 0, regular: 0, guest: 0 };
          checkInsByDate[key].total += total;
          checkInsByDate[key].regular += (ep.attributes.regular_count || 0);
          checkInsByDate[key].guest += (ep.attributes.guest_count || 0);
          gotData = true;
        }
        if (gotData) {
          activeEvents++;
          console.log(`  Event "${ev.attributes.name}" (${ev.id}): newest period starts_at = ${periods[0] && periods[0].attributes.starts_at}`);
        }
      } catch (e) { console.warn(`  Event ${ev.id} error:`, e.message); }
    }
    console.log(`PCO Check-Ins: ${activeEvents} active events, ${Object.keys(checkInsByDate).length} total dates loaded`);
  } catch (e) { console.warn('PCO Check-Ins fetch failed:', e.message); }

  const output = `window._PCO_STATIC=${JSON.stringify({ servicesByDate, checkInsByDate })};\n`;
  fs.writeFileSync(path.join(__dirname, '..', 'pco-data.js'), output);
  console.log('pco-data.js written successfully.');
}

main().catch(e => {
  console.error('PCO fetch script failed:', e);
  fs.writeFileSync(path.join(__dirname, '..', 'pco-data.js'),
    'window._PCO_STATIC={servicesByDate:{},checkInsByDate:{}};\n');
  process.exit(0);
});
