require('dotenv').config({ path: '../.env' });
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');
const fs = require('fs');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SEEN_FILE = '/tmp/seen.json';

process.on('uncaughtException', e => console.error('ERR:', e.message));
process.on('unhandledRejection', e => console.error('REJ:', e?.message || e));

let seen = new Set();
try { seen = new Set(JSON.parse(fs.readFileSync(SEEN_FILE))); } catch {}
const saveSeen = () => { try { fs.writeFileSync(SEEN_FILE, JSON.stringify([...seen])); } catch {} };

function getSheets() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.JWT(creds.client_email, null, creds.private_key,
    ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

// ─── HEADERS ───────────────────────────────────────────────────────────────
const HEADERS = [
  'Date Received',
  'Address', 'City', 'State', 'Zip', 'County', 'Subdivision',
  'Beds', 'Baths', 'Half Baths', 'Sqft', 'Lot Sqft', 'Lot Acres', 'Year Built', 'Stories',
  'Construction', 'Foundation', 'Pool', 'Pool Notes', 'Garage', 'Garage Spaces',
  'Carport', 'Basement', 'Attic',
  'Overall Condition', 'Roof Type', 'Roof Age / Year', 'AC Year / Age',
  'Water Heater', 'Electrical', 'Plumbing', 'Windows', 'Flooring',
  'Kitchen Notes', 'Bath Notes',
  'Asking Price', 'ARV', 'Repairs Estimate', 'Assignment Fee', 'Equity',
  'Rent Current', 'Rent Market', 'Annual Taxes', 'HOA Fee',
  'Close Date', 'Inspection Period', 'Earnest Money', 'Financing Terms', 'Cash Only',
  'Contact 1 Name', 'Contact 1 Title', 'Contact 1 Company',
  'Contact 1 Phone', 'Contact 1 Phone 2', 'Contact 1 Email', 'Contact 1 Website',
  'Contact 2 Name', 'Contact 2 Title', 'Contact 2 Company',
  'Contact 2 Phone', 'Contact 2 Email',
  'Contact 3 Name', 'Contact 3 Phone', 'Contact 3 Email',
  'ALL Phones Found', 'ALL Emails Found', 'ALL Names Found',
  'Seller Name', 'Seller Phone', 'Seller Situation', 'Seller Motivation', 'Occupancy',
  'Flood Zone', 'HOA', 'School District',
  'Google Drive Link', 'Zillow Link', 'Google Maps Link', 'All Other Links',
  'Photos Included', 'Photo Count', 'Photo Links',
  'Comp 1', 'Comp 2', 'Comp 3',
  'What Is Updated', 'What Needs Work', 'Highlights', 'Red Flags', 'Additional Notes',
  'Wholesaler Company', 'List Name',
  // ── NEW: Duplicate & Archive ──
  'Duplicate Flag', 'First Seen Date', 'Times Seen',
  'Email Subject', 'Email UID'
];

// ─── SHEET INIT ─────────────────────────────────────────────────────────────
async function initSheet() {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A1:A1' }).catch(() => null);
  const hasHeaders = res?.data?.values?.[0]?.[0] === 'Date Received';
  const hasData = (await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A2:A2' }).catch(() => null))?.data?.values;

  if (!hasHeaders) {
    await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals' }).catch(() => {});
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [HEADERS] }
    });
    // Also init Archive and Brain sheets
    await ensureTab(s, 'Deal Archive', 3);
    await ensureTab(s, "Derek's Brain", 4);
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Deal Archive!A1',
      valueInputOption: 'RAW', requestBody: { values: [HEADERS] }
    }).catch(() => {});
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A1",
      valueInputOption: 'RAW', requestBody: { values: [['Wholesaler Email', 'Wholesaler Company', 'Times Emailed', 'Typical Format', 'Avg Properties Per Email', 'Avg ARV Accuracy', 'Last Seen', 'Notes']] }
    }).catch(() => {});
    seen = new Set();
    saveSeen();
    console.log('Sheet initialized —', HEADERS.length, 'columns');
  } else if (!hasData) {
    seen = new Set();
    saveSeen();
    console.log('Sheet empty — seen IDs reset');
  } else {
    console.log('Resuming normally');
  }
}

async function ensureTab(s, title, index) {
  const sheet = await s.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const exists = sheet.data.sheets.some(sh => sh.properties.title === title);
  if (!exists) {
    await s.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title, index } } }] }
    });
  }
}

async function writeRow(range, row) {
  const s = getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `${range}!A:A`,
    valueInputOption: 'RAW', requestBody: { values: [row] }
  });
}

// ─── DUPLICATE DETECTION ─────────────────────────────────────────────────────
// Load all addresses currently in Active Deals
async function loadActiveAddresses() {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!B:G'  // Address through Subdivision
  }).catch(() => null);
  const rows = res?.data?.values || [];
  const map = new Map(); // normalized address → { row_index, date, times_seen }
  rows.slice(1).forEach((row, i) => {
    const addr = normalize(row[0], row[1], row[3]); // address, city, zip
    if (addr) map.set(addr, { rowIndex: i + 2, date: rows[i + 1]?.[0] });
  });
  return map;
}

function normalize(address, city, zip) {
  if (!address) return null;
  return `${address} ${city || ''} ${zip || ''}`.toLowerCase().replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim();
}

// Check if address exists in Active Deals — returns { isDupe, firstSeen, timesSeen } 
async function checkDuplicate(address, city, zip) {
  const s = getSheets();
  const nn = normalize(address, city, zip);
  if (!nn) return { isDupe: false };

  // Check Active Deals
  const active = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CP' }).catch(() => null);
  const rows = active?.data?.values || [];
  const headers = rows[0] || [];
  const addrIdx = headers.indexOf('Address');
  const cityIdx = headers.indexOf('City');
  const zipIdx = headers.indexOf('Zip');
  const dateIdx = headers.indexOf('Date Received');
  const timesIdx = headers.indexOf('Times Seen');

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rowAddr = normalize(row[addrIdx], row[cityIdx], row[zipIdx]);
    if (rowAddr === nn) {
      const timesSeen = parseInt(row[timesIdx] || '1') + 1;
      // Update times seen in that row
      const timesCol = String.fromCharCode(65 + timesIdx);
      await s.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `Active Deals!${timesCol}${i + 1}`,
        valueInputOption: 'RAW', requestBody: { values: [[timesSeen]] }
      }).catch(() => {});
      return { isDupe: true, firstSeen: row[dateIdx], timesSeen };
    }
  }

  // Also check Deal Archive
  const archive = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Deal Archive!A:CP' }).catch(() => null);
  const archRows = archive?.data?.values || [];
  for (let i = 1; i < archRows.length; i++) {
    const row = archRows[i];
    const rowAddr = normalize(row[addrIdx], row[cityIdx], row[zipIdx]);
    if (rowAddr === nn) {
      return { isDupe: true, firstSeen: row[dateIdx], timesSeen: 2, inArchive: true };
    }
  }

  return { isDupe: false };
}

// ─── 7-DAY AUTO-ARCHIVE ───────────────────────────────────────────────────────
async function archiveOldDeals() {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:A' }).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const cutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const toArchive = [];

  for (let i = 1; i < rows.length; i++) {
    const dateStr = rows[i][0];
    if (dateStr && new Date(dateStr) < cutoff) {
      toArchive.push(i + 1); // 1-indexed sheet row
    }
  }

  if (toArchive.length === 0) return;

  // Get full data for rows to archive
  const full = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CP' }).catch(() => null);
  const fullRows = full?.data?.values || [];

  const archiveRows = toArchive.map(idx => fullRows[idx - 1]).filter(Boolean);

  // Write to Deal Archive
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Deal Archive!A:A',
    valueInputOption: 'RAW', requestBody: { values: archiveRows }
  }).catch(() => {});

  // Delete from Active Deals (in reverse order to preserve indices)
  const activeSheetId = await getSheetId(s, 'Active Deals');
  const requests = toArchive.slice().reverse().map(rowIdx => ({
    deleteDimension: {
      range: {
        sheetId: activeSheetId,
        dimension: 'ROWS',
        startIndex: rowIdx - 1,
        endIndex: rowIdx
      }
    }
  }));

  await s.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests } }).catch(() => {});
  console.log(`📦 Archived ${toArchive.length} deals older than 7 days`);
}

async function getSheetId(s, title) {
  const sheet = await s.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const match = sheet.data.sheets.find(sh => sh.properties.title === title);
  return match?.properties?.sheetId || 0;
}

// ─── DEREK'S BRAIN ────────────────────────────────────────────────────────────
async function updateBrain(fromEmail, company, propertyCount, body) {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "Derek's Brain!A:H" }).catch(() => null);
  const rows = res?.data?.values || [];

  const now = new Date().toISOString();
  let found = false;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) {
      const timesEmailed = parseInt(rows[i][2] || '0') + 1;
      const avgProps = ((parseFloat(rows[i][4] || '0') * (timesEmailed - 1)) + propertyCount) / timesEmailed;
      // Update row
      const s2 = getSheets();
      await s2.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `Derek's Brain!A${i + 1}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[fromEmail, company || rows[i][1], timesEmailed, rows[i][3], avgProps.toFixed(1), rows[i][5], now, rows[i][7]]] }
      }).catch(() => {});
      found = true;
      break;
    }
  }

  if (!found) {
    // Detect format type from body
    const fmt = body.includes('ARV') ? 'ARV-listed' : body.includes('Sales Price') ? 'Sales-Price' : 'Unknown';
    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A:A",
      valueInputOption: 'RAW',
      requestBody: { values: [[fromEmail, company || '', 1, fmt, propertyCount, 'N/A', now, '']] }
    }).catch(() => {});
  }
}

// ─── KEYWORD SCREEN ───────────────────────────────────────────────────────────
const DEAL_WORDS = ['off market','wholesale','arv','flip','motivated','for sale','equity',
  'distressed','foreclosure','probate','deal','available deals','fix flip','cash buyer',
  'assignment','sales price','asking price','beds','baths','sqft','under roof','bungalow',
  'rehab','investment property','property available','price reduction','reduced'];
const isDeal = (subj, from) => DEAL_WORDS.some(k => `${subj} ${from}`.toLowerCase().includes(k));

// ─── EXTRACTION ───────────────────────────────────────────────────────────────
const v = x => (x === null || x === undefined) ? '' : String(x);

async function extractProperties(from, subject, body) {
  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4000,
    messages: [{ role: 'user', content: `You are a real estate data extraction engine for Coralstone Capital Group.
Extract EVERY property from this wholesale deal email. Return a JSON ARRAY — one object per property.

Each object must have ALL these fields (null if not found):
address, city, state, zip, county, subdivision, beds, baths, half_baths, sqft, lot_sqft, lot_acres,
year_built, stories, construction, foundation, pool, pool_notes, garage, garage_spaces, carport,
basement, attic, overall_condition, roof_type, roof_age, ac_year, water_heater, electrical, plumbing,
windows, flooring, kitchen_notes, bath_notes, asking_price, arv, repairs_estimate, assignment_fee,
equity, rent_current, rent_market, annual_taxes, hoa_fee, close_date, inspection_period, earnest_money,
financing_terms, cash_only, contact_1_name, contact_1_title, contact_1_company, contact_1_phone,
contact_1_phone_2, contact_1_email, contact_1_website, contact_2_name, contact_2_title, contact_2_company,
contact_2_phone, contact_2_email, contact_3_name, contact_3_phone, contact_3_email, all_phones,
all_emails, all_names, seller_name, seller_phone, seller_situation, seller_motivation, occupancy,
flood_zone, hoa, school_district, drive_link, zillow_link, google_maps_link, all_other_links,
photos_included, photo_count, photo_links, comp_1, comp_2, comp_3, what_is_updated, what_needs_work,
highlights, red_flags, additional_notes, wholesaler_company, list_name

FROM: ${from}
SUBJECT: ${subject}
BODY:
${body.slice(0, 12000)}

Return ONLY valid JSON array. No markdown.` }]
  });

  try {
    const text = res.content[0].text.trim().replace(/^```json\s*/,'').replace(/^```\s*/,'').replace(/\s*```$/,'').trim();
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    console.error('Parse error:', e.message);
    return null;
  }
}

function buildRow(p, subject, uid, dupeFlag, firstSeen, timesSeen) {
  return [
    new Date().toISOString(),
    v(p.address), v(p.city), v(p.state), v(p.zip), v(p.county), v(p.subdivision),
    v(p.beds), v(p.baths), v(p.half_baths), v(p.sqft), v(p.lot_sqft), v(p.lot_acres),
    v(p.year_built), v(p.stories), v(p.construction), v(p.foundation),
    v(p.pool), v(p.pool_notes), v(p.garage), v(p.garage_spaces),
    v(p.carport), v(p.basement), v(p.attic),
    v(p.overall_condition), v(p.roof_type), v(p.roof_age), v(p.ac_year),
    v(p.water_heater), v(p.electrical), v(p.plumbing), v(p.windows), v(p.flooring),
    v(p.kitchen_notes), v(p.bath_notes),
    v(p.asking_price), v(p.arv), v(p.repairs_estimate), v(p.assignment_fee), v(p.equity),
    v(p.rent_current), v(p.rent_market), v(p.annual_taxes), v(p.hoa_fee),
    v(p.close_date), v(p.inspection_period), v(p.earnest_money),
    v(p.financing_terms), v(p.cash_only),
    v(p.contact_1_name), v(p.contact_1_title), v(p.contact_1_company),
    v(p.contact_1_phone), v(p.contact_1_phone_2), v(p.contact_1_email), v(p.contact_1_website),
    v(p.contact_2_name), v(p.contact_2_title), v(p.contact_2_company),
    v(p.contact_2_phone), v(p.contact_2_email),
    v(p.contact_3_name), v(p.contact_3_phone), v(p.contact_3_email),
    v(p.all_phones), v(p.all_emails), v(p.all_names),
    v(p.seller_name), v(p.seller_phone), v(p.seller_situation), v(p.seller_motivation), v(p.occupancy),
    v(p.flood_zone), v(p.hoa), v(p.school_district),
    v(p.drive_link), v(p.zillow_link), v(p.google_maps_link), v(p.all_other_links),
    v(p.photos_included), v(p.photo_count), v(p.photo_links),
    v(p.comp_1), v(p.comp_2), v(p.comp_3),
    v(p.what_is_updated), v(p.what_needs_work), v(p.highlights), v(p.red_flags), v(p.additional_notes),
    v(p.wholesaler_company), v(p.list_name),
    // Duplicate fields
    dupeFlag ? `DUPE (first seen ${firstSeen})` : '',
    dupeFlag ? firstSeen : new Date().toISOString(),
    timesSeen || 1,
    v(subject), v(uid)
  ];
}

// ─── MAIN POLL ────────────────────────────────────────────────────────────────
async function poll() {
  console.log(`\n[${new Date().toLocaleTimeString()}] Polling ${process.env.IMAP_USER}...`);

  const client = new ImapFlow({
    host: process.env.IMAP_HOST, port: 993, secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false, socketTimeout: 120000, connectionTimeout: 60000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', e => console.error('IMAP err:', e.message));

  let written = 0, dupes = 0, archived = 0;

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');

    try {
      const since = new Date(Date.now() - 72 * 60 * 60 * 1000);

      // Phase 1: envelopes only
      const candidates = [];
      for await (const msg of client.fetch({ since }, { envelope: true, uid: true })) {
        const uid = String(msg.uid);
        if (seen.has(uid)) continue;
        const subj = msg.envelope?.subject || '';
        const from = msg.envelope?.from?.[0]?.address || '';
        if (isDeal(subj, from)) candidates.push({ uid, subj, from });
        else seen.add(uid);
      }
      saveSeen();
      console.log(`${candidates.length} deal email(s) to process`);

      // Phase 2: full body + extract
      for (const c of candidates) {
        console.log(`\n📬 ${c.subj}`);
        try {
          const msgData = await client.fetchOne(c.uid, { source: true }, { uid: true });
          const parsed = await simpleParser(msgData.source);
          const body = parsed.text || (parsed.html || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

          const properties = await extractProperties(c.from, c.subj, body);
          if (!properties || properties.length === 0) {
            console.log('  → No properties extracted');
          } else {
            console.log(`  → ${properties.length} propert${properties.length > 1 ? 'ies' : 'y'}`);

            for (const p of properties) {
              if (!p.address) { console.log('  → Skipped (no address)'); continue; }

              // Duplicate check
              const dup = await checkDuplicate(p.address, p.city, p.zip);

              if (dup.isDupe) {
                dupes++;
                console.log(`  🔁 DUPE: ${p.address} (first seen ${dup.firstSeen}, now seen ${dup.timesSeen}x)`);
                // Still write it but flagged
                await writeRow('Active Deals', buildRow(p, c.subj, c.uid, true, dup.firstSeen, dup.timesSeen));
              } else {
                await writeRow('Active Deals', buildRow(p, c.subj, c.uid, false, null, 1));
                console.log(`  ✅ ${p.address}, ${p.city}`);
                written++;
              }
            }

            // Update Derek's Brain
            await updateBrain(c.from, properties[0]?.wholesaler_company, properties.length, body);
          }

          seen.add(c.uid);
          saveSeen();
          await new Promise(r => setTimeout(r, 800));
        } catch (e) {
          console.error(`  Error on UID ${c.uid}:`, e.message);
          seen.add(c.uid);
          saveSeen();
        }
      }
    } finally { lock.release(); }

    // Auto-archive deals older than 7 days
    await archiveOldDeals();

  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    try { await client.logout(); } catch {}
  }

  console.log(`\n✅ Done — ${written} new | ${dupes} dupes flagged | archived`);
}

// ─── BOOT ─────────────────────────────────────────────────────────────────────
const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS / 60000}min`);
console.log(`   Features: duplicate detection | 7-day archive | wholesaler brain`);

initSheet().then(() => {
  poll();
  setInterval(poll, POLL_MS);
}).catch(e => { console.error('Init error:', e.message); poll(); setInterval(poll, POLL_MS); });
