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

  // --- Services ---
  try {
    const typesResp = await pcoGet('/services/v2/service_types?per_page=25');
    const types = (typesResp.data || []).slice(0, 8);
    for (const st of types) {
      try {
        const plansResp = await pcoGet(
          `/services/v2/service_types/${st.id}/plans?filter=past&per_page=50&order=-sort_date`
        );
        for (const plan of (plansResp.data || [])) {
          const key = dateKey(plan.attributes.sort_date);
          if (!key) continue;
          if (!servicesByDate[key]) servicesByDate[key] = [];
          servicesByDate[key].push({
            serviceType: st.attributes.name,
            title: plan.attributes.title || ''
          });
        }
      } catch (e) { console.warn(`  Services type ${st.id} error:`, e.message); }
    }
    console.log(`PCO Services: ${Object.keys(servicesByDate).length} dates loaded`);
  } catch (e) { console.warn('PCO Services fetch failed:', e.message); }

  // --- Check-Ins ---
  // Paginate through ALL events and check each one for recent event_periods.
  // Can't use top-level event_periods (PCO returns 404) and can't rely on
  // event updated_at ordering (metadata changes, not check-in activity).
  try {
    const since = new Date();
    since.setFullYear(since.getFullYear() - 2);
    const sinceMs = since.getTime();

    // Collect all events across pages
    const allEvents = [];
    let evPath = '/check-ins/v2/events?per_page=100';
    while (evPath) {
      const evResp = await pcoGet(evPath);
      allEvents.push(...(evResp.data || []));
      const next = evResp.links && evResp.links.next;
      evPath = next ? next.replace('https://api.planningcenteronline.com', '') : null;
    }
    console.log(`PCO Check-Ins: scanning ${allEvents.length} events for recent activity...`);

    for (const ev of allEvents) {
      try {
        // Fetch most-recent-first; skip this event entirely if newest period is older than 2 years
        const periodsResp = await pcoGet(
          `/check-ins/v2/events/${ev.id}/event_periods?per_page=100&order=-starts_at`
        );
        const periods = periodsResp.data || [];
        if (!periods.length) continue;
        // Quick bail: if the newest period is outside our window, skip this event
        const newestMs = new Date((periods[0].attributes.starts_at || '').split('T')[0]).getTime();
        if (newestMs < sinceMs) continue;

        for (const ep of periods) {
          const key = dateKey(ep.attributes.starts_at);
          if (!key) continue;
          // Client-side date filter
          const epMs = new Date(ep.attributes.starts_at.split('T')[0]).getTime();
          if (epMs < sinceMs) continue;
          const total = (ep.attributes.regular_count || 0)
            + (ep.attributes.guest_count || 0)
            + (ep.attributes.volunteer_count || 0);
          if (!total) continue;
          if (!checkInsByDate[key]) checkInsByDate[key] = { total: 0, regular: 0, guest: 0 };
          checkInsByDate[key].total += total;
          checkInsByDate[key].regular += (ep.attributes.regular_count || 0);
          checkInsByDate[key].guest += (ep.attributes.guest_count || 0);
        }
      } catch (e) { /* skip events that error */ }
    }
    console.log(`PCO Check-Ins: ${Object.keys(checkInsByDate).length} dates loaded`);
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
