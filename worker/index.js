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

// ─── COLUMN SCHEMA ───────────────────────────────────────────────────────────
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
  'Email Subject', 'Email UID',
  // Feature columns
  'Duplicate Of',        // address of original if dupe
  'Archive Date',        // filled when archived
];

const COL = {};
HEADERS.forEach((h, i) => { COL[h] = i; });

// ─── SHEET HELPERS ───────────────────────────────────────────────────────────
async function getTabId(s, title) {
  const ss = await s.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const tab = ss.data.sheets.find(t => t.properties.title === title);
  return tab?.properties?.sheetId;
}

async function ensureTab(s, title, index) {
  const ss = await s.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const exists = ss.data.sheets.some(t => t.properties.title === title);
  if (!exists) {
    await s.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title, index } } }] }
    });
  }
}

async function initSheet() {
  const s = getSheets();

  // Ensure all tabs exist
  await ensureTab(s, 'Active Deals', 0);
  await ensureTab(s, 'Deal Storage', 1);
  await ensureTab(s, 'Derek\'s Brain', 2);

  // Check if Active Deals has correct headers
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1:A1'
  }).catch(() => null);

  const hasHeaders = res?.data?.values?.[0]?.[0] === 'Date Received';
  if (!hasHeaders) {
    await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals' }).catch(() => {});
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [HEADERS] }
    });
    console.log(`Sheet initialized — ${HEADERS.length} columns`);
  }

  // Check if Deal Storage has headers
  const res2 = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A1:A1'
  }).catch(() => null);
  if (!res2?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A1',
      valueInputOption: 'RAW', requestBody: { values: [HEADERS] }
    });
  }

  // Check if Active Deals has data — if not, reset seen IDs
  const data = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A2:A2'
  }).catch(() => null);
  if (!data?.data?.values) {
    seen = new Set();
    saveSeen();
    console.log('Sheet empty — seen IDs reset, will reprocess all emails');
  }

  // Ensure Derek's Brain headers
  const brainRes = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A1:A1"
  }).catch(() => null);
  if (!brainRes?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A1",
      valueInputOption: 'RAW',
      requestBody: { values: [['Wholesaler Company', 'Email Domain', 'Total Deals Sent', 'Avg Asking Price', 'Price Range', 'Common Areas', 'Email Format Notes', 'Quality Notes', 'Last Seen', 'Sample Subject']] }
    });
    console.log("Derek's Brain initialized");
  }
}

async function appendToTab(tab, row) {
  const s = getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `${tab}!A:A`,
    valueInputOption: 'RAW', requestBody: { values: [row] }
  });
}

// ─── DUPLICATE DETECTION ────────────────────────────────────────────────────
function normalizeAddr(address, city, zip) {
  if (!address) return null;
  return `${address} ${city || ''} ${zip || ''}`.toLowerCase()
    .replace(/[^\w\s]/g, '').replace(/\s+/g, ' ').trim();
}

async function findDuplicate(address, city, zip) {
  const key = normalizeAddr(address, city, zip);
  if (!key) return null;
  const s = getSheets();

  for (const tab of ['Active Deals', 'Deal Storage']) {
    const res = await s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${tab}!B:G`  // Address through Zip cols
    }).catch(() => null);
    const rows = res?.data?.values || [];
    for (let i = 1; i < rows.length; i++) {
      const rowAddr = rows[i][0]; // Address
      const rowCity = rows[i][1]; // City
      const rowZip  = rows[i][4]; // Zip
      if (normalizeAddr(rowAddr, rowCity, rowZip) === key) {
        return `${rowAddr}, ${rowCity} (${tab} row ${i + 1})`;
      }
    }
  }
  return null;
}

// ─── 7-DAY AUTO-ARCHIVE ──────────────────────────────────────────────────────
async function archiveExpired() {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  }).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const now = new Date();
  const header = rows[0];
  const active = [header];
  const toArchive = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const dateStr = row[COL['Date Received']];
    if (!dateStr) { active.push(row); continue; }

    const received = new Date(dateStr);
    const ageDays = (now - received) / (1000 * 60 * 60 * 24);

    if (ageDays > 7) {
      // Mark archive date
      while (row.length < HEADERS.length) row.push('');
      row[COL['Archive Date']] = now.toISOString();
      toArchive.push(row);
    } else {
      active.push(row);
    }
  }

  if (toArchive.length > 0) {
    // Rewrite Active Deals without expired rows
    await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals' });
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: active }
    });
    // Append expired to Deal Storage
    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
      valueInputOption: 'RAW', requestBody: { values: toArchive }
    });
    console.log(`Archived ${toArchive.length} expired deal(s) to Deal Storage`);
  }
}

// ─── DEREK'S BRAIN ──────────────────────────────────────────────────────────
async function updateBrain(wholesalerCompany, emailDomain, dealCount, askingPrices, areas, subject) {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  }).catch(() => null);
  const rows = res?.data?.values || [];

  const avgPrice = askingPrices.length > 0
    ? Math.round(askingPrices.reduce((a, b) => a + b, 0) / askingPrices.length)
    : null;
  const priceRange = askingPrices.length > 0
    ? `$${Math.min(...askingPrices).toLocaleString()} - $${Math.max(...askingPrices).toLocaleString()}`
    : '';
  const commonAreas = [...new Set(areas.filter(Boolean))].join(', ');

  // Find existing row for this wholesaler
  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === wholesalerCompany || rows[i][1] === emailDomain) {
      existingRow = i + 1; // 1-indexed sheet row
      break;
    }
  }

  const brainRow = [
    wholesalerCompany || '',
    emailDomain || '',
    dealCount,
    avgPrice ? `$${avgPrice.toLocaleString()}` : '',
    priceRange,
    commonAreas,
    '', // Email Format Notes — Derek fills this over time
    '', // Quality Notes
    new Date().toISOString(),
    subject || ''
  ];

  if (existingRow > 0) {
    // Preserve existing notes (cols F and G = index 5, 6)
    const existing = rows[existingRow - 1];
    brainRow[6] = existing[6] || ''; // preserve format notes
    brainRow[7] = existing[7] || ''; // preserve quality notes
    // Update total deals — increment
    const prevCount = parseInt(existing[2]) || 0;
    brainRow[2] = prevCount + dealCount;

    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Derek's Brain!A${existingRow}`,
      valueInputOption: 'RAW', requestBody: { values: [brainRow] }
    });
  } else {
    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A:A",
      valueInputOption: 'RAW', requestBody: { values: [brainRow] }
    });
  }
}

// ─── KEYWORD SCREEN ──────────────────────────────────────────────────────────
const DEAL_WORDS = ['off market','wholesale','arv','flip','motivated','for sale','equity',
  'distressed','foreclosure','probate','deal','available deals','fix flip','cash buyer',
  'assignment','sales price','asking price','beds','baths','sqft','under roof','bungalow',
  'rehab','investment property','property available','price reduction','reduced'];
const isDeal = (subj, from) => DEAL_WORDS.some(k => `${subj} ${from}`.toLowerCase().includes(k));

// ─── EXTRACTION ───────────────────────────────────────────────────────────────
async function extractProperties(from, subject, body) {
  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4000,
    messages: [{ role: 'user', content: `You are a real estate data extraction engine. Extract EVERY property from this wholesale deal email. Return a JSON array — one object per property. Miss nothing.

Each object must have ALL these fields (null if not found):
address, city, state, zip, county, subdivision,
beds (number), baths (number), half_baths (number), sqft (number), lot_sqft (number), lot_acres (number),
year_built (number), stories (number), construction, foundation,
pool ("YES"/"NO"), pool_notes, garage, garage_spaces (number), carport, basement, attic,
overall_condition, roof_type, roof_age, ac_year, water_heater, electrical, plumbing, windows, flooring,
kitchen_notes, bath_notes,
asking_price (number only), arv (number only), repairs_estimate (number only), assignment_fee (number only), equity,
rent_current (number), rent_market (number), annual_taxes (number), hoa_fee,
close_date, inspection_period, earnest_money, financing_terms, cash_only,
contact_1_name, contact_1_title, contact_1_company, contact_1_phone, contact_1_phone_2, contact_1_email, contact_1_website,
contact_2_name, contact_2_title, contact_2_company, contact_2_phone, contact_2_email,
contact_3_name, contact_3_phone, contact_3_email,
all_phones, all_emails, all_names,
seller_name, seller_phone, seller_situation, seller_motivation, occupancy,
flood_zone, hoa, school_district,
drive_link, zillow_link, google_maps_link, all_other_links,
photos_included, photo_count, photo_links,
comp_1, comp_2, comp_3,
what_is_updated, what_needs_work, highlights, red_flags, additional_notes,
wholesaler_company, list_name

FROM: ${from}
SUBJECT: ${subject}
BODY:
${body.slice(0, 12000)}

Return ONLY a valid JSON array. No markdown. No explanation.` }]
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

const v = x => (x === null || x === undefined) ? '' : String(x);

function buildRow(p, subject, uid, dupeOf) {
  const row = new Array(HEADERS.length).fill('');
  row[COL['Date Received']] = new Date().toISOString();
  row[COL['Address']] = v(p.address);
  row[COL['City']] = v(p.city);
  row[COL['State']] = v(p.state);
  row[COL['Zip']] = v(p.zip);
  row[COL['County']] = v(p.county);
  row[COL['Subdivision']] = v(p.subdivision);
  row[COL['Beds']] = v(p.beds);
  row[COL['Baths']] = v(p.baths);
  row[COL['Half Baths']] = v(p.half_baths);
  row[COL['Sqft']] = v(p.sqft);
  row[COL['Lot Sqft']] = v(p.lot_sqft);
  row[COL['Lot Acres']] = v(p.lot_acres);
  row[COL['Year Built']] = v(p.year_built);
  row[COL['Stories']] = v(p.stories);
  row[COL['Construction']] = v(p.construction);
  row[COL['Foundation']] = v(p.foundation);
  row[COL['Pool']] = v(p.pool);
  row[COL['Pool Notes']] = v(p.pool_notes);
  row[COL['Garage']] = v(p.garage);
  row[COL['Garage Spaces']] = v(p.garage_spaces);
  row[COL['Carport']] = v(p.carport);
  row[COL['Basement']] = v(p.basement);
  row[COL['Attic']] = v(p.attic);
  row[COL['Overall Condition']] = v(p.overall_condition);
  row[COL['Roof Type']] = v(p.roof_type);
  row[COL['Roof Age / Year']] = v(p.roof_age);
  row[COL['AC Year / Age']] = v(p.ac_year);
  row[COL['Water Heater']] = v(p.water_heater);
  row[COL['Electrical']] = v(p.electrical);
  row[COL['Plumbing']] = v(p.plumbing);
  row[COL['Windows']] = v(p.windows);
  row[COL['Flooring']] = v(p.flooring);
  row[COL['Kitchen Notes']] = v(p.kitchen_notes);
  row[COL['Bath Notes']] = v(p.bath_notes);
  row[COL['Asking Price']] = v(p.asking_price);
  row[COL['ARV']] = v(p.arv);
  row[COL['Repairs Estimate']] = v(p.repairs_estimate);
  row[COL['Assignment Fee']] = v(p.assignment_fee);
  row[COL['Equity']] = v(p.equity);
  row[COL['Rent Current']] = v(p.rent_current);
  row[COL['Rent Market']] = v(p.rent_market);
  row[COL['Annual Taxes']] = v(p.annual_taxes);
  row[COL['HOA Fee']] = v(p.hoa_fee);
  row[COL['Close Date']] = v(p.close_date);
  row[COL['Inspection Period']] = v(p.inspection_period);
  row[COL['Earnest Money']] = v(p.earnest_money);
  row[COL['Financing Terms']] = v(p.financing_terms);
  row[COL['Cash Only']] = v(p.cash_only);
  row[COL['Contact 1 Name']] = v(p.contact_1_name);
  row[COL['Contact 1 Title']] = v(p.contact_1_title);
  row[COL['Contact 1 Company']] = v(p.contact_1_company);
  row[COL['Contact 1 Phone']] = v(p.contact_1_phone);
  row[COL['Contact 1 Phone 2']] = v(p.contact_1_phone_2);
  row[COL['Contact 1 Email']] = v(p.contact_1_email);
  row[COL['Contact 1 Website']] = v(p.contact_1_website);
  row[COL['Contact 2 Name']] = v(p.contact_2_name);
  row[COL['Contact 2 Title']] = v(p.contact_2_title);
  row[COL['Contact 2 Company']] = v(p.contact_2_company);
  row[COL['Contact 2 Phone']] = v(p.contact_2_phone);
  row[COL['Contact 2 Email']] = v(p.contact_2_email);
  row[COL['Contact 3 Name']] = v(p.contact_3_name);
  row[COL['Contact 3 Phone']] = v(p.contact_3_phone);
  row[COL['Contact 3 Email']] = v(p.contact_3_email);
  row[COL['ALL Phones Found']] = v(p.all_phones);
  row[COL['ALL Emails Found']] = v(p.all_emails);
  row[COL['ALL Names Found']] = v(p.all_names);
  row[COL['Seller Name']] = v(p.seller_name);
  row[COL['Seller Phone']] = v(p.seller_phone);
  row[COL['Seller Situation']] = v(p.seller_situation);
  row[COL['Seller Motivation']] = v(p.seller_motivation);
  row[COL['Occupancy']] = v(p.occupancy);
  row[COL['Flood Zone']] = v(p.flood_zone);
  row[COL['HOA']] = v(p.hoa);
  row[COL['School District']] = v(p.school_district);
  row[COL['Google Drive Link']] = v(p.drive_link);
  row[COL['Zillow Link']] = v(p.zillow_link);
  row[COL['Google Maps Link']] = v(p.google_maps_link);
  row[COL['All Other Links']] = v(p.all_other_links);
  row[COL['Photos Included']] = v(p.photos_included);
  row[COL['Photo Count']] = v(p.photo_count);
  row[COL['Photo Links']] = v(p.photo_links);
  row[COL['Comp 1']] = v(p.comp_1);
  row[COL['Comp 2']] = v(p.comp_2);
  row[COL['Comp 3']] = v(p.comp_3);
  row[COL['What Is Updated']] = v(p.what_is_updated);
  row[COL['What Needs Work']] = v(p.what_needs_work);
  row[COL['Highlights']] = v(p.highlights);
  row[COL['Red Flags']] = v(p.red_flags);
  row[COL['Additional Notes']] = v(p.additional_notes);
  row[COL['Wholesaler Company']] = v(p.wholesaler_company);
  row[COL['List Name']] = v(p.list_name);
  row[COL['Email Subject']] = v(subject);
  row[COL['Email UID']] = v(uid);
  row[COL['Duplicate Of']] = dupeOf ? v(dupeOf) : '';
  row[COL['Archive Date']] = '';
  return row;
}

// ─── MAIN POLL ───────────────────────────────────────────────────────────────
async function poll() {
  console.log(`\n[${new Date().toLocaleTimeString()}] Polling ${process.env.IMAP_USER}...`);

  // Run archive check on every poll
  await archiveExpired().catch(e => console.error('Archive error:', e.message));

  const client = new ImapFlow({
    host: process.env.IMAP_HOST, port: 993, secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false, socketTimeout: 120000, connectionTimeout: 60000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', e => console.error('IMAP err (handled):', e.message));

  let written = 0, dupes = 0;

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

      // Phase 2: fetch + extract
      for (const c of candidates) {
        console.log(`\n📬 ${c.subj}`);
        try {
          const msgData = await client.fetchOne(c.uid, { source: true }, { uid: true });
          const parsed = await simpleParser(msgData.source);
          const body = parsed.text || (parsed.html || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

          const properties = await extractProperties(c.from, c.subj, body);
          if (!properties || properties.length === 0) {
            console.log('  → No properties found');
          } else {
            console.log(`  → ${properties.length} propert${properties.length > 1 ? 'ies' : 'y'}`);

            // Collect brain data for this email
            const askingPrices = [];
            const areas = [];

            for (const p of properties) {
              if (!p.address) { console.log('  → Skipped (no address)'); continue; }

              // Duplicate detection
              const dupeOf = await findDuplicate(p.address, p.city, p.zip);

              if (dupeOf) {
                console.log(`  ⚠️ DUPE: ${p.address} already exists — ${dupeOf}`);
                // Still write it but flag it
                const row = buildRow(p, c.subj, c.uid, dupeOf);
                await appendToTab('Active Deals', row);
                dupes++;
              } else {
                const row = buildRow(p, c.subj, c.uid, null);
                await appendToTab('Active Deals', row);
                console.log(`  ✅ ${p.address}, ${p.city} | Ask: $${p.asking_price} | ARV: $${p.arv}`);
                written++;
              }

              if (p.asking_price && !isNaN(Number(p.asking_price))) askingPrices.push(Number(p.asking_price));
              if (p.city) areas.push(p.city);

              await new Promise(r => setTimeout(r, 400));
            }

            // Update Derek's Brain
            const domain = c.from.split('@')[1] || '';
            const company = properties[0]?.wholesaler_company || domain;
            if (company) {
              await updateBrain(company, domain, properties.length, askingPrices, areas, c.subj)
                .catch(e => console.error('Brain update error:', e.message));
            }
          }

          seen.add(c.uid);
          saveSeen();
        } catch (e) {
          console.error(`  Error on UID ${c.uid}:`, e.message);
          seen.add(c.uid);
          saveSeen();
        }
      }
    } finally { lock.release(); }

    console.log(`\n✅ Poll complete — ${written} new | ${dupes} dupes flagged`);
  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    try { await client.logout(); } catch {}
  }
}

// ─── BOOT ────────────────────────────────────────────────────────────────────
const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS/60000}min | ${HEADERS.length} columns`);
console.log('Features: duplicate detection | 7-day auto-archive | Derek\'s Brain');

initSheet().then(() => {
  poll();
  setInterval(poll, POLL_MS);
}).catch(e => {
  console.error('Init error:', e.message);
  poll();
  setInterval(poll, POLL_MS);
});
