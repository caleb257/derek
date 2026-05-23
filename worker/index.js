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

// Seen IDs in memory + persisted to Sheets
let seen = new Set();
let seenSheetReady = false;

function getSheets() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.JWT(creds.client_email, null, creds.private_key,
    ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

// ── Sheets retry wrapper ──────────────────────────────────────────────────────
async function sheetsCall(fn, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try { return await fn(); }
    catch (e) {
      if (i === retries - 1) throw e;
      const wait = (i + 1) * 2000;
      console.error(`  Sheets error (retry ${i+1}/${retries-1} in ${wait/1000}s):`, e.message);
      await new Promise(r => setTimeout(r, wait));
    }
  }
}

// ── Seen ID persistence ───────────────────────────────────────────────────────
async function loadSeenFromSheet() {
  try {
    const res = await sheetsCall(() => getSheets().spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: 'Seen!A:A'
    }));
    const rows = res?.data?.values || [];
    seen = new Set(rows.slice(1).map(r => r[0]).filter(Boolean));
    seenSheetReady = true;
    console.log(`📂 Loaded ${seen.size} seen IDs from sheet`);
  } catch (e) { console.error('Could not load seen IDs:', e.message); }
}

async function saveSeen() {
  if (!seenSheetReady) return;
  try {
    const values = [['Email UID'], ...[...seen].map(uid => [uid])];
    await sheetsCall(() => getSheets().spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Seen!A:A' }));
    await sheetsCall(() => getSheets().spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Seen!A1',
      valueInputOption: 'RAW', requestBody: { values }
    }));
  } catch (e) { console.error('Could not save seen IDs:', e.message); }
}

// ── Headers ───────────────────────────────────────────────────────────────────
const ACTIVE_HEADERS = [
  'Date Received', 'Pass', 'Sold', 'Review', 'Expires', 'Property Type',
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

const PRICE_CHANGE_HEADERS = [
  'Date Detected', 'Address', 'City', 'State', 'Zip',
  'Old Asking Price', 'New Asking Price', 'Change ($)', 'Change (%)',
  'ARV', 'Contact', 'Phone', 'Email', 'Wholesaler', 'Email Subject'
];

const BRAIN_HEADERS = [
  'Wholesaler Email', 'Wholesaler Company', 'Last Seen', 'Times Sent',
  'Format Type', 'Typical Fields', 'What Works', 'Watch Out For',
  'Avg Properties Per Email', 'Notes'
];

const v = x => (x === null || x === undefined) ? '' : String(x);

// ── Keywords ──────────────────────────────────────────────────────────────────
const DEAL_WORDS = [
  'off market', 'off-market', 'wholesale', 'flip', 'rehab', 'fixer', 'as-is', 'as is',
  'investment property', 'investment opportunity', 'bungalow', 'sfr',
  'arv', 'asking price', 'asking:', 'sales price', 'assignment fee',
  'equity', 'cash buyer', 'cash only', 'price reduction', 'price drop', 'reduced',
  'motivated', 'must sell', 'below market', 'deal alert',
  '3/2', '4/2', '4/3', '2/2', '2/1', '3/1', '5/3', '5/4', '4/4', '3/3', '2/1/1', '3/2/2',
  'distressed', 'foreclosure', 'pre-foreclosure', 'probate', 'divorce',
  'inherited', 'estate sale', 'vacant', 'absentee',
  'available deal', 'available now', 'property available', 'new deal',
  'deal alert', 'just listed', 'hot deal', 'assignment', 'subject to',
  'seller financing', 'seller finance', 'sub to',
  'sqft', 'sq ft', 'square feet', 'beds', 'baths', 'bedroom', 'bathroom',
  'under roof', 'living area', 'wholesaler', 'acquisitions', 'pocket listing',
  'manufactured', 'mobile home', 'duplex', 'triplex', 'fourplex', 'multi-family',
  'commercial', 'land', 'lot for sale', 'tear down', 'teardown'
];
const isDealEmail = (subj, from) => {
  const combined = `${subj} ${from}`.toLowerCase();
  return DEAL_WORDS.some(k => combined.includes(k));
};

// ── Property type detection ───────────────────────────────────────────────────
function detectPropertyType(p) {
  const text = `${v(p.address)} ${v(p.additional_notes)} ${v(p.highlights)} ${v(p.what_needs_work)}`.toLowerCase();
  if (/duplex|triplex|fourplex|multi.?family|multi.?unit|apt|apartment/.test(text)) return 'Multi-Family';
  if (/manufactured|mobile home|modular/.test(text)) return 'Manufactured';
  if (/commercial|warehouse|retail|office|industrial/.test(text)) return 'Commercial';
  if (/\blot\b|vacant land|acreage|\bacres\b|land/.test(text) && !p.beds) return 'Land';
  if (/condo|townhome|townhouse/.test(text)) return 'Condo/Townhome';
  return 'Single Family';
}

// ── Address normalization ─────────────────────────────────────────────────────
function normalizeAddress(addr) {
  if (!addr) return '';
  return addr.toLowerCase()
    .replace(/\./g, '').replace(/,/g, '').replace(/#/g, '')
    .replace(/\bstreet\b/g, 'st').replace(/\bavenue\b/g, 'ave')
    .replace(/\bdrive\b/g, 'dr').replace(/\bboulevard\b/g, 'blvd')
    .replace(/\broad\b/g, 'rd').replace(/\bcourt\b/g, 'ct')
    .replace(/\blane\b/g, 'ln').replace(/\bplace\b/g, 'pl')
    .replace(/\bcircle\b/g, 'cir').replace(/\bterrace\b/g, 'ter')
    .replace(/\btrail\b/g, 'trl').replace(/\bway\b/g, 'wy')
    .replace(/\bnorth\b/g, 'n').replace(/\bsouth\b/g, 's')
    .replace(/\beast\b/g, 'e').replace(/\bwest\b/g, 'w')
    .replace(/\bsaint\b/g, 'st').replace(/\b1st\b/g, 'first')
    .replace(/\b2nd\b/g, 'second').replace(/\b3rd\b/g, 'third')
    .replace(/\s+/g, ' ').trim();
}

// ── Duplicate + price change check ───────────────────────────────────────────
async function checkDuplicate(address, city, zip, newPrice) {
  if (!address) return { isDupe: false };
  const key = normalizeAddress(`${address} ${city} ${zip}`);
  for (const tab of ['Active Deals', 'Deal Storage']) {
    const res = await sheetsCall(() => getSheets().spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${tab}!C:AK` // C=Address...AK includes Asking Price col
    })).catch(() => null);
    const rows = res?.data?.values || [];
    // headers: C=Address D=City E=State F=Zip ... Asking Price is col 37 (0-indexed from C = index 34)
    for (let i = 1; i < rows.length; i++) {
      const existing = normalizeAddress(`${rows[i][0]} ${rows[i][1]} ${rows[i][3]}`);
      if (existing && existing === key) {
        // Price change detection — col index 34 = Asking Price (C offset)
        const oldPrice = parseFloat(String(rows[i][34] || '0').replace(/[^0-9.]/g, ''));
        const np = parseFloat(String(newPrice || '0').replace(/[^0-9.]/g, ''));
        if (np && oldPrice && Math.abs(np - oldPrice) > 100) {
          return { isDupe: true, isPriceChange: true, tab, oldPrice, newPrice: np };
        }
        return { isDupe: true, isPriceChange: false, tab };
      }
    }
  }
  return { isDupe: false };
}

// ── Footer stripper ───────────────────────────────────────────────────────────
function stripFooters(text) {
  if (!text) return '';
  const cutMarkers = [
    /^unsubscribe/im, /^to unsubscribe/im, /^you (are|were) receiving/im,
    /^this email was sent/im, /^you received this/im, /^if you no longer/im,
    /^confidentiality notice/im, /^disclaimer:/im, /^legal notice/im,
    /^this message (is|was) sent/im, /^remove me/im,
    /^manage (your )?preferences/im, /^privacy policy/im,
    /^view in browser/im, /^having trouble viewing/im,
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

// ── Brain context ─────────────────────────────────────────────────────────────
async function getBrainContext(fromEmail) {
  const res = await sheetsCall(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  })).catch(() => null);
  const rows = res?.data?.values || [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) {
      return {
        formatType: rows[i][4] || '', typicalFields: rows[i][5] || '',
        whatWorks: rows[i][6] || '', watchOutFor: rows[i][7] || '',
        timesSent: parseInt(rows[i][3] || '0')
      };
    }
  }
  return null;
}

// ── Extraction ────────────────────────────────────────────────────────────────
async function extractProperties(from, subject, body) {
  const cleanBody = stripFooters(body);
  const brain = await getBrainContext(from);
  let brainHint = '';
  if (brain && brain.timesSent > 0) {
    brainHint = `\n\nKNOWN SENDER (${brain.timesSent} emails seen):
- Format: ${brain.formatType}
- What works: ${brain.whatWorks}
- Watch out for: ${brain.watchOutFor}
Use this context to improve extraction accuracy.`;
  }

  const prompt = `You are a real estate data extraction engine for Coralstone Capital Group.
Extract EVERY property from this wholesale deal email. Return a JSON array — one object per property.${brainHint}

Each object must have ALL these fields (null if not found):
address, city, state, zip, county, subdivision,
beds (number), baths (number), half_baths (number),
sqft (number), lot_sqft (number), lot_acres (number),
year_built (number), stories (number),
construction, foundation, pool, pool_notes, garage, garage_spaces,
carport, basement, attic, overall_condition,
roof_type, roof_age, ac_year, water_heater, electrical, plumbing, windows, flooring,
kitchen_notes, bath_notes,
asking_price (number — search for "asking", "price", "sales price", dollar amounts),
arv (number — search for "ARV", "after repair value", "after repairs", "retail value"),
repairs_estimate (number), assignment_fee (number), equity,
rent_current, rent_market, annual_taxes (number), hoa_fee,
close_date, inspection_period, earnest_money, financing_terms, cash_only,
contact_1_name, contact_1_title, contact_1_company,
contact_1_phone, contact_1_phone_2, contact_1_email, contact_1_website,
contact_2_name, contact_2_title, contact_2_company, contact_2_phone, contact_2_email,
contact_3_name, contact_3_phone, contact_3_email,
all_phones (ALL phone numbers found anywhere), all_emails (ALL emails found anywhere), all_names,
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

Return ONLY a valid JSON array. No markdown.`;

  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001', max_tokens: 4000,
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

// ── Row builder ───────────────────────────────────────────────────────────────
function buildRow(p, subject, uid, propertyType) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  return [
    now.toISOString(), false, false, false, expires.toISOString(), propertyType,
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

// ── Sheet init ────────────────────────────────────────────────────────────────
async function initSheet() {
  const s = getSheets();
  const activeCheck = await sheetsCall(() => s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1:C1'
  })).catch(() => null);
  const firstCell = activeCheck?.data?.values?.[0]?.[0];
  const hasPropertyType = activeCheck?.data?.values?.[0]?.[2] === 'Property Type';
  const hasData = !!(await sheetsCall(() => s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A2:A2'
  })).catch(() => null))?.data?.values;

  if (firstCell !== 'Date Received' || !hasPropertyType) {
    await sheetsCall(() => s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals' }));
    await sheetsCall(() => s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [ACTIVE_HEADERS] }
    }));
    seen = new Set(); seenSheetReady = false;
    console.log(`Sheet initialized: ${ACTIVE_HEADERS.length} columns`);
  } else if (!hasData) {
    seen = new Set(); seenSheetReady = false;
    console.log('Sheet empty — resetting seen IDs');
  } else {
    console.log('Sheet has data — resuming');
  }

  for (const [tab, headers] of [
    ['Deal Storage', ACTIVE_HEADERS],
    ['Price Changes', PRICE_CHANGE_HEADERS],
    ['Errors', ['Date', 'From', 'Subject', 'UID', 'Error']],
    ["Derek's Brain", BRAIN_HEADERS],
    ['Seen', ['Email UID']]
  ]) {
    const check = await sheetsCall(() => s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${tab}!A1:A1`
    })).catch(() => null);
    const existingHeader = check?.data?.values?.[0]?.[0];
    if (!existingHeader || existingHeader !== headers[0]) {
      await sheetsCall(() => s.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `${tab}!A1`,
        valueInputOption: 'RAW', requestBody: { values: [headers] }
      }));
      console.log(`${tab} tab initialized`);
    }
  }
}

// ── Error logging ─────────────────────────────────────────────────────────────
async function logError(from, subject, uid, error) {
  await sheetsCall(() => getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Errors!A:A',
    valueInputOption: 'RAW',
    requestBody: { values: [[new Date().toISOString(), from, subject, uid, error]] }
  })).catch(() => {});
  console.log(`⚠️ Logged to Errors tab: ${error}`);
}

// ── Price change logger ───────────────────────────────────────────────────────
async function logPriceChange(p, subject, oldPrice, newPrice) {
  const change = newPrice - oldPrice;
  const changePct = oldPrice ? ((change / oldPrice) * 100).toFixed(1) : '?';
  await sheetsCall(() => getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Price Changes!A:A',
    valueInputOption: 'RAW',
    requestBody: { values: [[
      new Date().toISOString(),
      v(p.address), v(p.city), v(p.state), v(p.zip),
      oldPrice, newPrice, change, `${changePct}%`,
      v(p.arv), v(p.contact_1_name), v(p.contact_1_phone), v(p.contact_1_email),
      v(p.wholesaler_company), subject
    ]] }
  }));
  const arrow = change < 0 ? '📉' : '📈';
  console.log(`  ${arrow} Price change: ${p.address} $${oldPrice} → $${newPrice} (${changePct}%)`);
}

// ── Process checkboxes (Pass / Sold) ─────────────────────────────────────────
async function processCheckboxes() {
  const s = getSheets();
  const res = await sheetsCall(() => s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  })).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const headers = rows[0];
  const keep = [headers];
  const passed = [];
  const sold = [];

  // B=Pass (index 1), C=Sold (index 2)
  for (let i = 1; i < rows.length; i++) {
    const isPass = rows[i][1] === 'TRUE' || rows[i][1] === true;
    const isSold = rows[i][2] === 'TRUE' || rows[i][2] === true;
    if (isPass) {
      passed.push(rows[i]);
    } else if (isSold) {
      const soldRow = [...rows[i]];
      soldRow[1] = false; // clear Pass
      soldRow[2] = false; // clear Sold
      sold.push(soldRow);
    } else {
      keep.push(rows[i]);
    }
  }

  if (passed.length === 0 && sold.length === 0) return;

  // Rewrite Active Deals without passed/sold rows
  await sheetsCall(() => s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ' }));
  await sheetsCall(() => s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: keep }
  }));

  // Archive passed rows to Deal Storage
  if (passed.length > 0) {
    await sheetsCall(() => s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
      valueInputOption: 'RAW', requestBody: { values: passed }
    }));
    console.log(`🗑️  Moved ${passed.length} PASS deal(s) to Deal Storage`);
  }

  // Archive sold rows to Deal Storage with SOLD tag in Review column (D, index 3)
  if (sold.length > 0) {
    const taggedSold = sold.map(r => {
      const row = [...r];
      row[3] = 'SOLD'; // mark Review column
      return row;
    });
    await sheetsCall(() => s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
      valueInputOption: 'RAW', requestBody: { values: taggedSold }
    }));
    console.log(`🏠 Moved ${sold.length} SOLD deal(s) to Deal Storage`);
  }
}

// ── Auto-archive ──────────────────────────────────────────────────────────────
async function archiveExpired() {
  const res = await sheetsCall(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  })).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const now = new Date(), headers = rows[0];
  const active = [headers], expired = [];
  for (let i = 1; i < rows.length; i++) {
    const expires = rows[i][1];
    if (expires && new Date(expires) < now) expired.push(rows[i]);
    else active.push(rows[i]);
  }
  if (expired.length === 0) return;

  await sheetsCall(() => getSheets().spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ' }));
  await sheetsCall(() => getSheets().spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: active }
  }));
  await sheetsCall(() => getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
    valueInputOption: 'RAW', requestBody: { values: expired }
  }));
  console.log(`📦 Archived ${expired.length} expired deal(s)`);
}

// ── Derek's Brain update ──────────────────────────────────────────────────────
async function updateBrain(fromEmail, company, subject, propertyCount, body) {
  const res = await sheetsCall(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  })).catch(() => null);
  const rows = res?.data?.values || [];
  const now = new Date().toISOString();

  const hasEmoji = /[\u{1F300}-\u{1FFFF}]/u.test(body);
  const hasBullets = /^[-•*]/m.test(body);
  const hasNumbered = /^\d+\./m.test(body);
  const hasTable = / \| /.test(body);
  const formatType = hasEmoji ? 'Emoji bullets' : hasNumbered ? 'Numbered list'
    : hasBullets ? 'Bullet list' : hasTable ? 'Table' : 'Plain text';

  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) { existingRow = i + 1; break; }
  }

  if (existingRow > 0) {
    const ex = rows[existingRow - 1];
    const timesSent = parseInt(ex[3] || '0') + 1;
    const prevAvg = isNaN(parseFloat(ex[8])) ? 0 : parseFloat(ex[8]);
    const avgProps = ((prevAvg * (timesSent - 1)) + propertyCount) / timesSent;
    await sheetsCall(() => getSheets().spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `Derek's Brain!A${existingRow}`,
      valueInputOption: 'RAW', requestBody: { values: [[
        fromEmail, company || ex[1] || '', now, timesSent,
        formatType, ex[5] || '', ex[6] || '', ex[7] || '',
        avgProps.toFixed(1), ex[9] || ''
      ]] }
    }));
    console.log(`🧠 Brain: ${fromEmail} (${timesSent} emails, avg ${avgProps.toFixed(1)} props)`);
  } else {
    await sheetsCall(() => getSheets().spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A:A",
      valueInputOption: 'RAW', requestBody: { values: [[
        fromEmail, company || '', now, 1, formatType,
        'address, beds, baths, sqft, asking, arv, drive link, phone',
        'Check for Google Drive links per property', '',
        propertyCount.toString(), `First seen: ${subject}`
      ]] }
    }));
    console.log(`🧠 Brain: new wholesaler — ${fromEmail}`);
  }
}

// ── Checkbox processor (Pass/Sold → archive) ─────────────────────────────────
async function processCheckboxes() {
  const s = getSheets();
  const res = await sheetsCall(() => s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  })).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const headers = rows[0];
  const keep = [headers];
  const toArchive = [];
  let passed = 0, sold = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const isPass = row[1] === 'TRUE' || row[1] === true;
    const isSold = row[2] === 'TRUE' || row[2] === true;

    if (isPass) {
      // Mark as PASS in the archived copy
      const archived = [...row];
      archived[1] = 'PASS';
      toArchive.push(archived);
      passed++;
    } else if (isSold) {
      const archived = [...row];
      archived[2] = 'SOLD';
      toArchive.push(archived);
      sold++;
    } else {
      keep.push(row);
    }
  }

  if (toArchive.length === 0) return;

  // Rewrite Active Deals without the archived rows
  await sheetsCall(() => s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ' }));
  await sheetsCall(() => s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: keep }
  }));
  // Append to Deal Storage
  await sheetsCall(() => s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
    valueInputOption: 'RAW', requestBody: { values: toArchive }
  }));

  if (passed > 0) console.log(`❌ Passed ${passed} deal(s) → Deal Storage`);
  if (sold > 0) console.log(`🤝 Sold ${sold} deal(s) → Deal Storage`);
}

// ── IMAP connect with retry ───────────────────────────────────────────────────
async function connectImap(retries = 3) {
  for (let i = 0; i < retries; i++) {
    const client = new ImapFlow({
      host: process.env.IMAP_HOST, port: 993, secure: true,
      auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
      logger: false, socketTimeout: 120000, connectionTimeout: 60000,
      tls: { rejectUnauthorized: false }
    });
    client.on('error', e => console.error('IMAP socket error (handled):', e.message));
    try {
      await client.connect();
      return client;
    } catch (e) {
      if (i === retries - 1) throw e;
      console.error(`IMAP connect failed (retry ${i+1}/${retries-1}):`, e.message);
      await new Promise(r => setTimeout(r, 3000));
    }
  }
}

// ── Main poll ─────────────────────────────────────────────────────────────────
async function poll() {
  console.log(`\n[${new Date().toLocaleTimeString()}] Polling ${process.env.IMAP_USER}...`);
  await archiveExpired();

  // Process Pass/Sold checkboxes before polling IMAP
  await processCheckboxes();

  let client;
  let written = 0, dupes = 0, priceChanges = 0, errors = 0;

  try {
    client = await connectImap();
    const lock = await client.getMailboxLock('INBOX');

    try {
      const since = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);

      // Phase 1: envelopes only
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

              const propertyType = detectPropertyType(p);
              const dupeCheck = await checkDuplicate(p.address, p.city, p.zip, p.asking_price);

              if (dupeCheck.isDupe && dupeCheck.isPriceChange) {
                await logPriceChange(p, c.subj, dupeCheck.oldPrice, dupeCheck.newPrice);
                priceChanges++;
              } else if (dupeCheck.isDupe) {
                console.log(`  🔁 DUPE: ${p.address} (in ${dupeCheck.tab})`);
                dupes++;
              } else {
                await sheetsCall(() => getSheets().spreadsheets.values.append({
                  spreadsheetId: SHEET_ID, range: 'Active Deals!A:A',
                  valueInputOption: 'RAW',
                  requestBody: { values: [buildRow(p, c.subj, c.uid, propertyType)] }
                }));
                console.log(`  ✅ [${propertyType}] ${p.address}, ${p.city} | Ask: $${p.asking_price} | ARV: $${p.arv}`);
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

    console.log(`\n✅ Poll done — ${written} new | ${dupes} dupes | ${priceChanges} price changes | ${errors} errors`);
  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    if (client) { try { await client.logout(); } catch {} }
  }
}

// ── Boot ──────────────────────────────────────────────────────────────────────
const POLL_MS = parseInt(process.env.POLL_INTERVAL || '300000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS / 60000}min`);
console.log(`✓ Price changes | ✓ Sheets retry | ✓ Property type | ✓ IMAP reconnect`);

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
