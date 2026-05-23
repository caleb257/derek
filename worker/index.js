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

// Persist seen email IDs across restarts
let seen = new Set();
try { seen = new Set(JSON.parse(fs.readFileSync(SEEN_FILE))); } catch {}
const saveSeen = () => { try { fs.writeFileSync(SEEN_FILE, JSON.stringify([...seen])); } catch {} };

function getSheets() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.JWT(creds.client_email, null, creds.private_key,
    ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

// ─── SHEET COLUMNS ───────────────────────────────────────────────────────────
const ACTIVE_HEADERS = [
  'Date Received', 'Expires',
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
  'Wholesaler Company', 'List Name', 'Email Subject', 'Email UID'
];

const BRAIN_HEADERS = [
  'Wholesaler Email', 'Wholesaler Company', 'Last Seen', 'Times Sent',
  'Format Notes', 'Typical Fields', 'What Works', 'Watch Out For',
  'Avg Properties Per Email', 'Notes'
];

const v = x => (x === null || x === undefined) ? '' : String(x);

// ─── SHEET INIT ───────────────────────────────────────────────────────────────
async function initSheet() {
  const s = getSheets();

  // Check Active Deals
  const activeCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1:A1'
  }).catch(() => null);
  const hasActiveHeaders = activeCheck?.data?.values?.[0]?.[0] === 'Date Received';
  const hasActiveData = !!(await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A2:A2'
  }).catch(() => null))?.data?.values;

  if (!hasActiveHeaders) {
    await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals' }).catch(() => {});
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [ACTIVE_HEADERS] }
    });
    seen = new Set(); saveSeen();
    console.log(`Sheet initialized: ${ACTIVE_HEADERS.length} columns`);
  } else if (!hasActiveData) {
    seen = new Set(); saveSeen();
    console.log('Sheet empty — resetting seen IDs');
  } else {
    console.log('Sheet has existing data — resuming');
  }

  // Check Deal Storage
  const storageCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A1:A1'
  }).catch(() => null);
  if (!storageCheck?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A1',
      valueInputOption: 'RAW', requestBody: { values: [ACTIVE_HEADERS] }
    });
    console.log('Deal Storage tab initialized');
  }

  // Check Derek's Brain
  const brainCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A1:A1"
  }).catch(() => null);
  if (!brainCheck?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A1",
      valueInputOption: 'RAW', requestBody: { values: [BRAIN_HEADERS] }
    });
    console.log("Derek's Brain tab initialized");
  }
}

// ─── DUPLICATE DETECTION ─────────────────────────────────────────────────────
function normalizeAddress(addr) {
  if (!addr) return '';
  return addr.toLowerCase()
    .replace(/\./g, '').replace(/,/g, '').replace(/\s+/g, ' ')
    .replace(/\bstreet\b/g, 'st').replace(/\bavenue\b/g, 'ave')
    .replace(/\bdrive\b/g, 'dr').replace(/\bboulevard\b/g, 'blvd')
    .replace(/\broad\b/g, 'rd').replace(/\bcourt\b/g, 'ct')
    .replace(/\blane\b/g, 'ln').replace(/\bplace\b/g, 'pl')
    .trim();
}

async function isDuplicate(address, city, zip) {
  if (!address) return { isDupe: false };
  const key = normalizeAddress(`${address} ${city} ${zip}`);
  const s = getSheets();

  for (const tab of ['Active Deals', 'Deal Storage']) {
    const res = await s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${tab}!C:F` // Address, City, State, Zip
    }).catch(() => null);
    const rows = res?.data?.values || [];
    for (let i = 1; i < rows.length; i++) {
      const existing = normalizeAddress(`${rows[i][0]} ${rows[i][1]} ${rows[i][3]}`);
      if (existing && existing === key) {
        return { isDupe: true, tab, row: i + 1 };
      }
    }
  }
  return { isDupe: false };
}

// ─── AUTO-ARCHIVE ─────────────────────────────────────────────────────────────
async function archiveExpired() {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  }).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const now = new Date();
  const headers = rows[0];
  const active = [headers];
  const expired = [];

  for (let i = 1; i < rows.length; i++) {
    const expires = rows[i][1]; // column B = Expires
    if (expires && new Date(expires) < now) {
      expired.push(rows[i]);
    } else {
      active.push(rows[i]);
    }
  }

  if (expired.length === 0) return;

  // Clear Active Deals and rewrite without expired rows
  await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ' });
  await s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: active }
  });

  // Append expired to Deal Storage
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
    valueInputOption: 'RAW', requestBody: { values: expired }
  });

  console.log(`📦 Archived ${expired.length} expired deal(s) to Deal Storage`);
}

// ─── DEREK'S BRAIN ────────────────────────────────────────────────────────────
async function updateBrain(fromEmail, company, subject, propertyCount, body) {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  }).catch(() => null);
  const rows = res?.data?.values || [];

  const now = new Date().toISOString();

  // Find existing entry for this sender
  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) { existingRow = i + 1; break; }
  }

  // Quick analysis of what format they use
  const hasEmoji = /[\u{1F300}-\u{1FFFF}]/u.test(body);
  const hasBullets = /^[-•*]/m.test(body);
  const hasNumbered = /^\d+\./m.test(body);
  const hasTable = body.includes('\t') || body.includes(' | ');
  const formatType = hasEmoji ? 'Emoji bullets' : hasNumbered ? 'Numbered list' : hasBullets ? 'Bullet list' : hasTable ? 'Table' : 'Plain text';

  if (existingRow > 0) {
    // Update existing entry
    const existing = rows[existingRow - 1];
    const timesSent = parseInt(existing[3] || '0') + 1;
    const avgProps = ((parseFloat(existing[8] || '0') * (timesSent - 1)) + propertyCount) / timesSent;

    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `Derek's Brain!A${existingRow}`,
      valueInputOption: 'RAW', requestBody: {
        values: [[
          fromEmail,
          company || existing[1] || '',
          now,
          timesSent,
          formatType,
          existing[5] || '', // Typical fields — preserved
          existing[6] || '', // What works — preserved
          existing[7] || '', // Watch out for — preserved
          avgProps.toFixed(1),
          existing[9] || ''  // Notes — preserved
        ]]
      }
    });
    console.log(`🧠 Brain updated: ${fromEmail} (${timesSent} emails seen)`);
  } else {
    // New wholesaler
    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A:A",
      valueInputOption: 'RAW', requestBody: {
        values: [[
          fromEmail,
          company || '',
          now,
          1,
          formatType,
          'address, beds, baths, sqft, asking, arv, drive link, phone', // defaults
          'Address + emoji bullets common format',
          '',
          propertyCount.toString(),
          `First seen via: ${subject}`
        ]]
      }
    });
    console.log(`🧠 Brain: new wholesaler logged — ${fromEmail}`);
  }
}

// ─── DEAL KEYWORDS ───────────────────────────────────────────────────────────
const DEAL_WORDS = ['off market','wholesale','arv','flip','motivated','for sale','equity',
  'distressed','foreclosure','probate','deal','available deals','fix flip','cash buyer',
  'assignment','sales price','asking price','beds','baths','sqft','under roof','bungalow',
  'rehab','investment property','property available','price reduction','reduced'];
const isDealEmail = (subj, from) => DEAL_WORDS.some(k => `${subj} ${from}`.toLowerCase().includes(k));

// ─── EXTRACTION ───────────────────────────────────────────────────────────────
async function extractProperties(from, subject, body) {
  const prompt = `You are a real estate data extraction engine for Coralstone Capital Group.
Extract EVERY property from this wholesale deal email. Return a JSON array — one object per property. Miss nothing.

Each object must have ALL these fields (null if not found):
address, city, state, zip, county, subdivision,
beds, baths, half_baths, sqft, lot_sqft, lot_acres, year_built, stories,
construction, foundation, pool, pool_notes, garage, garage_spaces,
carport, basement, attic,
overall_condition, roof_type, roof_age, ac_year,
water_heater, electrical, plumbing, windows, flooring,
kitchen_notes, bath_notes,
asking_price (number only), arv (number only), repairs_estimate (number only),
assignment_fee (number only), equity, rent_current, rent_market,
annual_taxes (number only), hoa_fee,
close_date, inspection_period, earnest_money, financing_terms, cash_only,
contact_1_name, contact_1_title, contact_1_company,
contact_1_phone, contact_1_phone_2, contact_1_email, contact_1_website,
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

Return ONLY a valid JSON array. No markdown. No explanation.`;

  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4000,
    messages: [{ role: 'user', content: prompt }]
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

// ─── ROW BUILDER ─────────────────────────────────────────────────────────────
function buildRow(p, subject, uid) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000); // 7 days
  return [
    now.toISOString(),
    expires.toISOString(),
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
    v(subject), v(uid)
  ];
}

async function appendRow(row) {
  const s = getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:A',
    valueInputOption: 'RAW', requestBody: { values: [row] }
  });
}

// ─── MAIN POLL ────────────────────────────────────────────────────────────────
async function poll() {
  console.log(`\n[${new Date().toLocaleTimeString()}] Polling ${process.env.IMAP_USER}...`);

  // Archive expired deals first
  await archiveExpired();

  const client = new ImapFlow({
    host: process.env.IMAP_HOST, port: 993, secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false, socketTimeout: 120000, connectionTimeout: 60000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', e => console.error('IMAP err (handled):', e.message));

  let written = 0, dupes = 0, skipped = 0;

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');

    try {
      const since = new Date(Date.now() - 72 * 60 * 60 * 1000);

      // Phase 1: headers only — zero body download
      const candidates = [];
      for await (const msg of client.fetch({ since }, { envelope: true, uid: true })) {
        const uid = String(msg.uid);
        if (seen.has(uid)) continue;
        const subj = msg.envelope?.subject || '';
        const from = msg.envelope?.from?.[0]?.address || '';
        if (isDealEmail(subj, from)) {
          candidates.push({ uid, subj, from });
        } else {
          seen.add(uid);
        }
      }
      saveSeen();
      console.log(`${candidates.length} deal email(s) to process`);

      // Phase 2: full body + extract for deal emails only
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
            console.log(`  → ${properties.length} propert${properties.length > 1 ? 'ies' : 'y'} found`);

            // Update brain for this wholesaler
            const company = properties[0]?.wholesaler_company || '';
            await updateBrain(c.from, company, c.subj, properties.length, body);

            for (const p of properties) {
              if (!p.address) { console.log('  → Skipped (no address)'); continue; }

              // Duplicate check
              const dupeCheck = await isDuplicate(p.address, p.city, p.zip);
              if (dupeCheck.isDupe) {
                console.log(`  🔁 DUPE: ${p.address}, ${p.city} (already in ${dupeCheck.tab})`);
                dupes++;
                continue;
              }

              const row = buildRow(p, c.subj, c.uid);
              await appendRow(row);
              console.log(`  ✅ ${p.address}, ${p.city} | Ask: $${p.asking_price} | ARV: $${p.arv}`);
              written++;
              await new Promise(r => setTimeout(r, 300));
            }
          }

          seen.add(c.uid);
          saveSeen();
          await new Promise(r => setTimeout(r, 500));
        } catch (e) {
          console.error(`  Error on UID ${c.uid}:`, e.message);
          seen.add(c.uid);
          saveSeen();
        }
      }
    } finally { lock.release(); }

    console.log(`\n✅ Poll done — ${written} new, ${dupes} dupes, ${skipped} skipped`);
  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    try { await client.logout(); } catch {}
  }
}

// ─── BOOT ─────────────────────────────────────────────────────────────────────
const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS / 60000}min`);
console.log(`Features: duplicate detection ✓ | 7-day auto-archive ✓ | wholesaler brain ✓`);

initSheet().then(() => {
  poll();
  setInterval(poll, POLL_MS);
}).catch(e => {
  console.error('Init error:', e.message);
  poll();
  setInterval(poll, POLL_MS);
});
