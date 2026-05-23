require('dotenv').config({ path: '../.env' });
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');
const fs = require('fs');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
process.on('uncaughtException', e => console.error('ERR:', e.message));
process.on('unhandledRejection', e => console.error('REJ:', e?.message || e));

// Seen IDs stored in Google Sheets (persistent across Railway restarts/container moves)
// Falls back to in-memory if sheet not ready yet
let seen = new Set();
let seenSheetReady = false;

async function loadSeenFromSheet() {
  try {
    const s = getSheets();
    const res = await s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: 'Seen!A:A'
    }).catch(() => null);
    const rows = res?.data?.values || [];
    seen = new Set(rows.slice(1).map(r => r[0]).filter(Boolean));
    seenSheetReady = true;
    console.log(`📂 Loaded ${seen.size} seen IDs from sheet`);
  } catch (e) {
    console.error('Could not load seen IDs from sheet:', e.message);
  }
}

async function saveSeen() {
  if (!seenSheetReady) return;
  try {
    const s = getSheets();
    const values = [['Email UID'], ...[...seen].map(uid => [uid])];
    await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Seen!A:A' });
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Seen!A1',
      valueInputOption: 'RAW', requestBody: { values }
    });
  } catch (e) {
    console.error('Could not save seen IDs to sheet:', e.message);
  }
}

function getSheets() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.JWT(creds.client_email, null, creds.private_key,
    ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

// ─── HEADERS ──────────────────────────────────────────────────────────────────
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
  'Format Type', 'Typical Fields', 'What Works', 'Watch Out For',
  'Avg Properties Per Email', 'Notes'
];

const v = x => (x === null || x === undefined) ? '' : String(x);

// ─── IMPROVEMENT 1: BIGGER KEYWORD LIST ──────────────────────────────────────
const DEAL_WORDS = [
  // Property type signals
  'off market', 'wholesale', 'flip', 'rehab', 'fixer', 'as-is', 'as is',
  'investment property', 'investment opportunity', 'bungalow', 'sfr',
  // Financial signals
  'arv', 'asking price', 'asking:', 'sales price', 'assignment fee',
  'equity', 'cash buyer', 'cash only', 'price reduction', 'price drop', 'reduced',
  'motivated', 'must sell', 'below market', 'deal alert',
  // Bedroom/bath patterns (e.g. "3/2", "4/3/2")
  '3/2', '4/2', '4/3', '2/2', '2/1', '3/1', '5/3', '5/4', '4/4', '3/3',
  // Situation signals
  'distressed', 'foreclosure', 'pre-foreclosure', 'probate', 'divorce',
  'inherited', 'estate sale', 'vacant', 'absentee',
  // Action signals
  'available deal', 'available now', 'property available', 'new deal',
  'deal alert', 'just listed', 'hot deal', 'assignment', 'subject to',
  'seller financing', 'seller finance', 'sub to',
  // Size signals
  'sqft', 'sq ft', 'square feet', 'beds', 'baths', 'bedroom', 'bathroom',
  'under roof', 'living area',
  // Market signals
  'wholesaler', 'acquisitions', 'off-market', 'pocket listing'
];

const isDealEmail = (subj, from) => {
  const combined = `${subj} ${from}`.toLowerCase();
  return DEAL_WORDS.some(k => combined.includes(k));
};

// ─── IMPROVEMENT 2: BETTER ADDRESS NORMALIZATION ─────────────────────────────
function normalizeAddress(addr) {
  if (!addr) return '';
  return addr.toLowerCase()
    .replace(/\./g, '').replace(/,/g, '').replace(/#/g, '')
    // Full → abbreviated
    .replace(/\bstreet\b/g, 'st').replace(/\bavenue\b/g, 'ave')
    .replace(/\bdrive\b/g, 'dr').replace(/\bboulevard\b/g, 'blvd')
    .replace(/\broad\b/g, 'rd').replace(/\bcourt\b/g, 'ct')
    .replace(/\blane\b/g, 'ln').replace(/\bplace\b/g, 'pl')
    .replace(/\bcircle\b/g, 'cir').replace(/\bterrace\b/g, 'ter')
    .replace(/\btrail\b/g, 'trl').replace(/\bway\b/g, 'wy')
    .replace(/\bnorth\b/g, 'n').replace(/\bsouth\b/g, 's')
    .replace(/\beast\b/g, 'e').replace(/\bwest\b/g, 'w')
    .replace(/\bsaint\b/g, 'st').replace(/\bmt\b/g, 'mount')
    // Ordinals
    .replace(/\b1st\b/g, 'first').replace(/\b2nd\b/g, 'second')
    .replace(/\b3rd\b/g, 'third').replace(/\b4th\b/g, 'fourth')
    .replace(/\s+/g, ' ').trim();
}

async function isDuplicate(address, city, zip) {
  if (!address) return { isDupe: false };
  const key = normalizeAddress(`${address} ${city} ${zip}`);
  const s = getSheets();
  for (const tab of ['Active Deals', 'Deal Storage']) {
    const res = await s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${tab}!C:F`
    }).catch(() => null);
    const rows = res?.data?.values || [];
    for (let i = 1; i < rows.length; i++) {
      const existing = normalizeAddress(`${rows[i][0]} ${rows[i][1]} ${rows[i][3]}`);
      if (existing && existing === key) return { isDupe: true, tab };
    }
  }
  return { isDupe: false };
}

// ─── IMPROVEMENT 5: STRIP EMAIL FOOTERS ──────────────────────────────────────
function stripFooters(text) {
  if (!text) return '';
  // Common footer markers — cut everything after these
  const cutMarkers = [
    /^unsubscribe/im, /^to unsubscribe/im, /^you (are|were) receiving/im,
    /^this email was sent/im, /^you received this/im, /^if you no longer/im,
    /^confidentiality notice/im, /^disclaimer:/im, /^legal notice/im,
    /^this message (is|was) sent/im, /^remove me/im, /^manage (your )?preferences/im,
    /^privacy policy/im, /^view in browser/im, /^having trouble viewing/im,
    /^update your email preferences/im, /^©\s*20/im,
    /^sent from my/im, /^get outlook for/im
  ];

  const lines = text.split('\n');
  for (let i = 0; i < lines.length; i++) {
    if (cutMarkers.some(rx => rx.test(lines[i].trim()))) {
      return lines.slice(0, i).join('\n').trim();
    }
  }
  return text.trim();
}

// ─── IMPROVEMENT 6: BRAIN-INFORMED EXTRACTION ────────────────────────────────
async function getBrainContext(fromEmail) {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  }).catch(() => null);
  const rows = res?.data?.values || [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) {
      return {
        formatType: rows[i][4] || '',
        typicalFields: rows[i][5] || '',
        whatWorks: rows[i][6] || '',
        watchOutFor: rows[i][7] || '',
        timesSent: parseInt(rows[i][3] || '0')
      };
    }
  }
  return null;
}

// ─── EXTRACTION ───────────────────────────────────────────────────────────────
async function extractProperties(from, subject, body) {
  // Improvement 5: strip footer before sending
  const cleanBody = stripFooters(body);

  // Improvement 6: get brain context for this sender
  const brain = await getBrainContext(from);
  let brainHint = '';
  if (brain && brain.timesSent > 0) {
    brainHint = `\n\nKNOWN SENDER INTELLIGENCE (${brain.timesSent} emails seen):
- Format: ${brain.formatType}
- What works: ${brain.whatWorks}
- Watch out for: ${brain.watchOutFor}
- Typical fields present: ${brain.typicalFields}
Use this to improve your extraction accuracy.`;
  }

  const prompt = `You are a real estate data extraction engine for Coralstone Capital Group.
Extract EVERY property from this wholesale deal email. Return a JSON array — one object per property. Miss nothing.${brainHint}

Each object must have ALL these fields (null if not found):
address, city, state, zip, county, subdivision,
beds, baths, half_baths, sqft, lot_sqft, lot_acres, year_built, stories,
construction, foundation, pool, pool_notes, garage, garage_spaces,
carport, basement, attic, overall_condition,
roof_type, roof_age, ac_year, water_heater, electrical, plumbing, windows, flooring,
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
${cleanBody.slice(0, 12000)}

Return ONLY a valid JSON array. No markdown. No explanation.`;

  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4000,
    messages: [{ role: 'user', content: prompt }]
  });

  try {
    const text = res.content[0].text.trim()
      .replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '').trim();
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
  const expires = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  return [
    now.toISOString(), expires.toISOString(),
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

// ─── SHEET INIT ───────────────────────────────────────────────────────────────
async function initSheet() {
  const s = getSheets();
  const activeCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1:A1'
  }).catch(() => null);
  const hasHeaders = activeCheck?.data?.values?.[0]?.[0] === 'Date Received';
  const hasData = !!(await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A2:A2'
  }).catch(() => null))?.data?.values;

  if (!hasHeaders) {
    await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals' }).catch(() => {});
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [ACTIVE_HEADERS] }
    });
    seen = new Set(); await saveSeen();
    console.log(`Sheet initialized: ${ACTIVE_HEADERS.length} columns`);
  } else if (!hasData) {
    seen = new Set(); await saveSeen();
    console.log('Sheet empty — resetting seen IDs');
  } else {
    console.log('Sheet has data — resuming');
  }

  // Deal Storage tab
  const storageCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A1:A1'
  }).catch(() => null);
  if (!storageCheck?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A1',
      valueInputOption: 'RAW', requestBody: { values: [ACTIVE_HEADERS] }
    });
  }

  // Errors tab — improvement 3
  const errCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Errors!A1:A1'
  }).catch(() => null);
  if (!errCheck?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Errors!A1',
      valueInputOption: 'RAW', requestBody: { values: [['Date', 'From', 'Subject', 'UID', 'Error']] }
    });
    console.log('Errors tab initialized');
  }

  // Seen UIDs tab (persistent across restarts)
  const seenCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Seen!A1:A1'
  }).catch(() => null);
  if (!seenCheck?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Seen!A1',
      valueInputOption: 'RAW', requestBody: { values: [['Email UID']] }
    });
    console.log('Seen tab initialized');
  }

  // Derek's Brain tab
  const brainCheck = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A1:A1"
  }).catch(() => null);
  if (!brainCheck?.data?.values?.[0]?.[0]) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A1",
      valueInputOption: 'RAW', requestBody: { values: [BRAIN_HEADERS] }
    });
    console.log("Derek's Brain initialized");
  }
}

// ─── IMPROVEMENT 3: LOG ERRORS TO SHEET ──────────────────────────────────────
async function logError(from, subject, uid, error) {
  const s = getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Errors!A:A',
    valueInputOption: 'RAW',
    requestBody: { values: [[new Date().toISOString(), from, subject, uid, error]] }
  }).catch(() => {});
  console.log(`⚠️ Error logged to sheet: ${error}`);
}

// ─── AUTO-ARCHIVE (7 days) ────────────────────────────────────────────────────
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
    const expires = rows[i][1];
    if (expires && new Date(expires) < now) expired.push(rows[i]);
    else active.push(rows[i]);
  }

  if (expired.length === 0) return;

  await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ' });
  await s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: active }
  });
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
    valueInputOption: 'RAW', requestBody: { values: expired }
  });
  console.log(`📦 Archived ${expired.length} expired deal(s)`);
}

// ─── DEREK'S BRAIN UPDATE ────────────────────────────────────────────────────
async function updateBrain(fromEmail, company, subject, propertyCount, body) {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  }).catch(() => null);
  const rows = res?.data?.values || [];
  const now = new Date().toISOString();

  const hasEmoji = /[\u{1F300}-\u{1FFFF}]/u.test(body);
  const hasBullets = /^[-•*]/m.test(body);
  const hasNumbered = /^\d+\./m.test(body);
  const hasTable = body.includes('\t') || / \| /.test(body);
  const formatType = hasEmoji ? 'Emoji bullets' : hasNumbered ? 'Numbered list'
    : hasBullets ? 'Bullet list' : hasTable ? 'Table' : 'Plain text';

  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) { existingRow = i + 1; break; }
  }

  if (existingRow > 0) {
    const existing = rows[existingRow - 1];
    const timesSent = parseInt(existing[3] || '0') + 1;
    const prevAvgRaw = parseFloat(existing[8] || '0');
    const prevAvg = isNaN(prevAvgRaw) ? 0 : prevAvgRaw;
    const avgProps = ((prevAvg * (timesSent - 1)) + propertyCount) / timesSent;

    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `Derek's Brain!A${existingRow}`,
      valueInputOption: 'RAW', requestBody: { values: [[
        fromEmail, company || existing[1] || '', now, timesSent,
        formatType, existing[5] || '', existing[6] || '',
        existing[7] || '', avgProps.toFixed(1), existing[9] || ''
      ]] }
    });
    console.log(`🧠 Brain: ${fromEmail} updated (${timesSent} emails, avg ${avgProps.toFixed(1)} props)`);
  } else {
    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A:A",
      valueInputOption: 'RAW', requestBody: { values: [[
        fromEmail, company || '', now, 1, formatType,
        'address, beds, baths, sqft, asking, arv, drive link, phone',
        'Check for Google Drive links per property',
        '', propertyCount.toString(), `First seen: ${subject}`
      ]] }
    });
    console.log(`🧠 Brain: new wholesaler logged — ${fromEmail}`);
  }
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
  await archiveExpired();

  const client = new ImapFlow({
    host: process.env.IMAP_HOST, port: 993, secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false, socketTimeout: 120000, connectionTimeout: 60000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', e => console.error('IMAP err (handled):', e.message));

  let written = 0, dupes = 0, skipped = 0, errors = 0;

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');

    try {
      // Improvement 4: 7-day lookback instead of 72 hours
      const since = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);

      // Phase 1: envelopes only — zero cost
      const candidates = [];
      for await (const msg of client.fetch({ since }, { envelope: true, uid: true })) {
        const uid = String(msg.uid);
        if (seen.has(uid)) continue;
        const subj = msg.envelope?.subject || '';
        const from = msg.envelope?.from?.[0]?.address || '';
        if (isDealEmail(subj, from)) candidates.push({ uid, subj, from });
        else seen.add(uid);
      }
      await saveSeen();
      console.log(`${candidates.length} deal email(s) to process`);

      // Phase 2: full body + extract
      for (const c of candidates) {
        console.log(`\n📬 ${c.subj}`);
        try {
          const msgData = await client.fetchOne(c.uid, { source: true }, { uid: true });
          const parsed = await simpleParser(msgData.source);
          const rawBody = parsed.text || (parsed.html || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

          const properties = await extractProperties(c.from, c.subj, rawBody);

          // Improvement 3: log failed extractions
          if (!properties || properties.length === 0) {
            console.log('  → No properties extracted');
            await logError(c.from, c.subj, c.uid, 'Extraction returned 0 properties');
            errors++;
          } else {
            console.log(`  → ${properties.length} propert${properties.length > 1 ? 'ies' : 'y'}`);
            const company = properties[0]?.wholesaler_company || '';
            await updateBrain(c.from, company, c.subj, properties.length, rawBody);

            for (const p of properties) {
              if (!p.address) { console.log('  → Skipped (no address)'); continue; }
              const dupeCheck = await isDuplicate(p.address, p.city, p.zip);
              if (dupeCheck.isDupe) {
                console.log(`  🔁 DUPE: ${p.address} (already in ${dupeCheck.tab})`);
                dupes++;
              } else {
                await appendRow(buildRow(p, c.subj, c.uid));
                console.log(`  ✅ ${p.address}, ${p.city} | Ask: $${p.asking_price} | ARV: $${p.arv}`);
                written++;
              }
              await new Promise(r => setTimeout(r, 300));
            }
          }

          seen.add(c.uid);
          await saveSeen();
          await new Promise(r => setTimeout(r, 500));
        } catch (e) {
          console.error(`  Error on UID ${c.uid}:`, e.message);
          await logError(c.from, c.subj, c.uid, e.message);
          seen.add(c.uid);
          await saveSeen();
          errors++;
        }
      }
    } finally { lock.release(); }

    console.log(`\n✅ Poll done — ${written} new | ${dupes} dupes | ${errors} errors | ${skipped} skipped`);
  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    try { await client.logout(); } catch {}
  }
}

// ─── BOOT ─────────────────────────────────────────────────────────────────────
const POLL_MS = parseInt(process.env.POLL_INTERVAL || '300000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS / 60000}min`);
console.log(`✓ Extended keywords | ✓ Better dupe detection | ✓ Error logging`);
console.log(`✓ 7-day lookback | ✓ Footer stripping | ✓ Brain-informed extraction`);

initSheet()
  .then(() => loadSeenFromSheet())
  .then(() => {
    poll();
    setInterval(poll, POLL_MS);
  }).catch(e => {
    console.error('Init error:', e.message);
    poll();
    setInterval(poll, POLL_MS);
  });
