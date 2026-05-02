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
  // Fetch event_periods directly (top-level) filtered to last 2 years.
  // Avoids the event updated_at ordering problem where inactive old events
  // sort ahead of the currently-active Sunday service event.
  try {
    const since = new Date();
    since.setFullYear(since.getFullYear() - 2);
    const sinceStr = since.toISOString().split('T')[0];
    let path = `/check-ins/v2/event_periods?where[starts_at][gte]=${sinceStr}&per_page=100&order=-starts_at`;
    let pages = 0;
    while (path && pages < 10) {
      const resp = await pcoGet(path);
      for (const ep of (resp.data || [])) {
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
      }
      // Follow pagination if more pages exist
      const next = resp.links && resp.links.next;
      path = next ? next.replace('https://api.planningcenteronline.com', '') : null;
      pages++;
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
