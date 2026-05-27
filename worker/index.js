require('dotenv').config({ path: '../.env' });
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const URBAN_URL = process.env.URBAN_URL || 'https://urban-production-cffb.up.railway.app';
const URBAN_TOKEN = process.env.URBAN_TOKEN || process.env.URBAN_PASSWORD || 'coralstone2025';

process.on('uncaughtException', e => console.error('ERR:', e.message));
process.on('unhandledRejection', e => console.error('REJ:', e?.message || e));

// Seen IDs — in-memory, saved to Sheets once per poll (not per email)
let seen = new Set();
let seenSheetReady = false;
let seenDirty = false; // track if seen changed this poll

function getSheets() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.JWT(creds.client_email, null, creds.private_key,
    ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

// Sheets retry wrapper
async function sc(fn, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try { return await fn(); }
    catch (e) {
      if (i === retries - 1) throw e;
      const wait = (i + 1) * 2000;
      console.error(`  Sheets retry ${i+1} in ${wait/1000}s:`, e.message);
      await new Promise(r => setTimeout(r, wait));
    }
  }
}

// Convert 0-based index to spreadsheet column letter (A, B, ..., Z, AA, AB...)
function colLetter(idx) {
  let letter = '', n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// Seen persistence — loads once at boot, saves once at end of each poll
// Known wholesaler emails from Brain — bypass keyword filter for these
let knownSenders = new Set();

async function loadKnownSenders() {
  try {
    const res = await sc(() => getSheets().spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A:A"
    }));
    const rows = res?.data?.values || [];
    knownSenders = new Set(rows.slice(1).map(r => r[0]).filter(Boolean));
    console.log(`📧 Loaded ${knownSenders.size} known senders from Brain`);
  } catch (e) { console.error('loadKnownSenders error:', e.message); }
}

async function loadSeenFromSheet() {
  try {
    const res = await sc(() => getSheets().spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: 'Seen!A:A'
    }));
    seen = new Set((res?.data?.values || []).slice(1).map(r => r[0]).filter(Boolean));
    seenSheetReady = true;
    console.log(`📂 Loaded ${seen.size} seen IDs`);
  } catch (e) { console.error('loadSeen error:', e.message); }
}

async function flushSeen() {
  if (!seenSheetReady || !seenDirty) return;
  try {
    // Only keep the last 2000 UIDs — prune oldest to keep set lean
    // UIDs are IMAP sequential so highest = newest
    const seenArr = [...seen].map(Number).sort((a,b) => a-b);
    if (seenArr.length > 2000) {
      const pruned = seenArr.slice(-2000); // keep newest 2000
      seen = new Set(pruned.map(String));
      console.log(`🧹 Pruned seen set to ${seen.size} UIDs`);
    }
    const values = [['Email UID'], ...[...seen].map(uid => [uid])];
    await sc(() => getSheets().spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Seen!A:A' }));
    await sc(() => getSheets().spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Seen!A1',
      valueInputOption: 'RAW', requestBody: { values }
    }));
    seenDirty = false;
  } catch (e) { console.error('flushSeen error:', e.message); }
}

// ── SCHEMA ────────────────────────────────────────────────────────────────────
const ACTIVE_HEADERS = [
  'Date Received', 'Pass', 'Sold', 'Review', 'Asking Price', 'ARV', 'Expires', 'Property Type',
  'Address', 'City', 'State', 'Zip', 'County', 'Subdivision',
  'Beds', 'Baths', 'Half Baths', 'Sqft', 'Lot Sqft', 'Lot Acres',
  'Year Built', 'Stories', 'Construction', 'Foundation',
  'Pool', 'Pool Notes', 'Garage', 'Garage Spaces', 'Carport', 'Basement', 'Attic',
  'Overall Condition', 'Roof Type', 'Roof Age / Year', 'AC Year / Age',
  'Water Heater', 'Electrical', 'Plumbing', 'Windows', 'Flooring',
  'Kitchen Notes', 'Bath Notes',
  'Repairs Estimate', 'Assignment Fee', 'Equity',
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
  'Wholesaler Company', 'List Name', 'Days Active', 'Email Subject', 'Email UID'
];

// COL map — immune to column reordering
const COL = {};
ACTIVE_HEADERS.forEach((h, i) => { COL[h] = i; });

const PRICE_CHANGE_HEADERS = ['Date Detected','Address','City','State','Zip',
  'Old Asking Price','New Asking Price','Change ($)','Change (%)','ARV',
  'Contact','Phone','Email','Wholesaler','Email Subject'];

const BRAIN_HEADERS = ['Wholesaler Email','Wholesaler Company','Last Seen','Times Sent',
  'Format Type','Typical Fields','What Works','Watch Out For','Avg Properties Per Email','Notes'];

const DIR_HEADERS = ['Company','Contact Name','Email','Phone','Website',
  'Deals Sent','Avg Properties/Email','Last Deal Date','Property Types','Avg Asking Price','Notes'];

const v = x => (x === null || x === undefined) ? '' : String(x);
const safeNum = x => { const n = parseFloat(String(x||'').replace(/[^0-9.]/g,'')); return isNaN(n) ? 0 : n; };
const safeInt = x => { const n = parseInt(String(x||'')); return isNaN(n) ? 0 : n; };

// ── KEYWORDS ──────────────────────────────────────────────────────────────────
const DEAL_WORDS = [
  'off market','off-market','wholesale','flip','rehab','fixer','as-is','as is',
  'investment property','investment opportunity','bungalow','sfr',
  'arv','asking price','asking:','sales price','assignment fee',
  'equity','cash buyer','cash only','price reduction','price drop','reduced',
  'motivated','must sell','below market','deal alert',
  '3/2','4/2','4/3','2/2','2/1','3/1','5/3','5/4','4/4','3/3',
  'distressed','foreclosure','pre-foreclosure','probate','divorce',
  'inherited','estate sale','vacant','absentee',
  'available deal','available now','property available','new deal',
  'just listed','hot deal','assignment','subject to',
  'seller financing','seller finance','sub to',
  'sqft','sq ft','square feet','beds','baths','bedroom','bathroom',
  'under roof','living area','wholesaler','acquisitions','pocket listing',
  'manufactured','mobile home','duplex','triplex','fourplex','multi-family',
  'commercial','land','lot for sale','tear down','teardown'
];
// Real estate company domain patterns — any email from these gets through
const DEAL_DOMAINS = [
  'properties', 'homes', 'realty', 'realtors', 'investments', 'capital',
  'acquisitions', 'buyers', 'sellers', 'wholesal', 'flipper', 'investor',
  'estate', 'housing', 'proptech', 'homesolutions', 'cashbuyer', 'offmarket',
  '27.pro', 'prophethomes', 'mitchelldean', 'investmdr', 'keygleehomes', 'alpinetrustproperties',
  'automation.podio', '21propertygroup', 'sell4no', 'flbuyers', 'coralstone',
  'hubs.ly', 'hubspotemail', 'sendgrid', 'mailchimp', 'constantcontact'
];

// Broad subject keywords — much wider net than before
const DEAL_WORDS_EXTRA = [
  'property', 'properties', 'home', 'house', 'deal', 'flip', 'rehab', 'invest',
  'wholesale', 'off market', 'off-market', 'listing', 'available', 'sale',
  'buyer', 'seller', 'motivated', 'equity', 'arv', 'asking', 'price',
  '3/2', '4/2', '4/3', '2/2', '3/1', '5/3', '5/4', '4/4', '3/3', '2/1',
  'sqft', 'sq ft', 'beds', 'bath', 'bedroom', 'single family', 'sfr',
  'foreclosure', 'probate', 'vacant', 'estate', 'cash', 'assignment',
  'subject to', 'sub to', 'seller financ', 'opportunity', 'acquisition',
  'portfolio', 'distressed', 'reduced', 'new deal', 'hot deal', 'just listed',
  'multi', 'duplex', 'triplex', 'commercial', 'land', 'lot', 'acre',
  'manufactured', 'mobile', 'condo', 'townhome', 'investment', 'rental',
  'bungalow', 'fixer', 'as-is', 'as is', 'tear down', 'teardown',
  'below market', 'must sell', 'check this out', 'check out', 'new listing',
  'just came in', 'just hit', 'fresh deal', 'pocket listing', 'new one',
  'new property', 'new properties', 'have a deal', 'got a deal', 'good deal',
  'great deal', 'price drop', 'price reduction', 'reduced price', 'price cut',
  'back on market', 'back on the market', 'back on', 'price change', 'updated',
  'contract', 'under contract', 'close', 'closing', 'escrow', 'title',
  'inspect', 'walk', 'drive by', 'showing', 'view', 'visit', 'tour',
  'tampa', 'st pete', 'saint pete', 'clearwater', 'brandon', 'lakeland',
  'orlando', 'miami', 'jacksonville', 'sarasota', 'naples', 'fort myers',
  'spring hill', 'brooksville', 'zephyrhills', 'holiday', 'new port richey',
  'port richey', 'hudson', 'hernando', 'pasco', 'hillsborough', 'pinellas',
  'manatee', 'polk', 'florida', ' fl ', ',fl', 'fl 3'
];

const isDealEmail = (subj, from) => {
  // 1. Always process known senders regardless of subject
  if (knownSenders.has(from)) { return true; }

  // 2. Domain pattern matching — real estate company domains
  const fromLower = from.toLowerCase();
  const matchedDomain = DEAL_DOMAINS.find(d => fromLower.includes(d));
  if (matchedDomain) { return true; }

  // 3. Broad keyword matching on subject + from
  const t = `${subj} ${from}`.toLowerCase();
  const matchedWord = DEAL_WORDS.find(k => t.includes(k)) || DEAL_WORDS_EXTRA.find(k => t.includes(k));
  if (matchedWord) { return true; }

  // Nothing matched — log it so we can see what's being filtered
  console.log(`  🔍 Filter miss: "${subj}" from ${from} — sending to Haiku pre-screen`);
  return false;
};

// ── PROPERTY TYPE ─────────────────────────────────────────────────────────────
function detectPropertyType(p) {
  const t = `${v(p.address)} ${v(p.additional_notes)} ${v(p.highlights)} ${v(p.what_needs_work)}`.toLowerCase();
  if (/duplex|triplex|fourplex|multi.?family|multi.?unit/.test(t)) return 'Multi-Family';
  if (/manufactured|mobile home|modular/.test(t)) return 'Manufactured';
  if (/commercial|warehouse|retail|office/.test(t)) return 'Commercial';
  if (/\blot\b|vacant land|acreage|\bacres\b/.test(t) && !p.beds) return 'Land';
  if (/condo|townhome|townhouse/.test(t)) return 'Condo/Townhome';
  return 'Single Family';
}

// ── ADDRESS NORMALIZATION ─────────────────────────────────────────────────────
function normalizeAddr(addr, city, zip) {
  return `${addr||''} ${city||''} ${zip||''}`.toLowerCase()
    .replace(/[.,#]/g,'')
    .replace(/\bstreet\b/g,'st').replace(/\bavenue\b/g,'ave')
    .replace(/\bdrive\b/g,'dr').replace(/\bboulevard\b/g,'blvd')
    .replace(/\broad\b/g,'rd').replace(/\bcourt\b/g,'ct')
    .replace(/\blane\b/g,'ln').replace(/\bplace\b/g,'pl')
    .replace(/\bcircle\b/g,'cir').replace(/\bterrace\b/g,'ter')
    .replace(/\bnorth\b/g,'n').replace(/\bsouth\b/g,'s')
    .replace(/\beast\b/g,'e').replace(/\bwest\b/g,'w')
    .replace(/\bsaint\b/g,'st').replace(/\b1st\b/g,'first')
    .replace(/\b2nd\b/g,'second').replace(/\b3rd\b/g,'third')
    .replace(/\s+/g,' ').trim();
}

// ── DUPE CHECK — COL map, never drifts ───────────────────────────────────────
async function checkDuplicate(address, city, zip, newPrice) {
  if (!address) return { isDupe: false };
  const key = normalizeAddr(address, city, zip);
  const isXXXX = /^x+$/i.test((address||'').trim()) || (address||'').trim() === 'XXXX';
  const addrIdx = COL['Address'], cityIdx = COL['City'];
  const zipIdx = COL['Zip'], priceIdx = COL['Asking Price'];

  for (const tab of ['Active Deals', 'Deal Storage']) {
    const res = await sc(() => getSheets().spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${tab}!A:CZ`
    })).catch(() => null);
    const rows = res?.data?.values || [];
    for (let i = 1; i < rows.length; i++) {
      const rowAddr = rows[i][addrIdx] || '';
      const rowCity = rows[i][cityIdx] || '';
      const rowZip = rows[i][zipIdx] || '';
      const existing = normalizeAddr(rowAddr, rowCity, rowZip);

      // For XXXX addresses: match on city + zip + price (same wholesaler, same area, same price = same deal)
      if (isXXXX) {
        const rowIsXXXX = /^x+$/i.test(rowAddr.trim()) || rowAddr.trim() === 'XXXX';
        if (rowIsXXXX) {
          const cityMatch = rowCity.toLowerCase().trim() === (city||'').toLowerCase().trim();
          const zipMatch = rowZip.trim() === (zip||'').trim();
          if (cityMatch && zipMatch) {
            const oldPrice = safeNum(rows[i][priceIdx]);
            const np = safeNum(newPrice);
            if (np && oldPrice && Math.abs(np - oldPrice) > 1000) {
              return { isDupe: true, isPriceChange: true, tab, oldPrice, newPrice: np };
            }
            return { isDupe: true, isPriceChange: false, tab };
          }
        }
        continue; // XXXX vs real address = not a dupe
      }

      // Normal address matching
      if (!existing || existing !== key) continue;
      const oldPrice = safeNum(rows[i][priceIdx]);
      const np = safeNum(newPrice);
      if (np && oldPrice && Math.abs(np - oldPrice) > 100) {
        return { isDupe: true, isPriceChange: true, tab, oldPrice, newPrice: np };
      }
      return { isDupe: true, isPriceChange: false, tab };
    }
  }
  return { isDupe: false };
}

// ── FOOTER STRIP ──────────────────────────────────────────────────────────────
function stripFooters(text) {
  if (!text) return '';
  const cuts = [/^unsubscribe/im,/^to unsubscribe/im,/^you (are|were) receiving/im,
    /^this email was sent/im,/^you received this/im,/^if you no longer/im,
    /^confidentiality notice/im,/^disclaimer:/im,/^this message (is|was) sent/im,
    /^remove me/im,/^manage (your )?preferences/im,/^privacy policy/im,
    /^view in browser/im,/^having trouble viewing/im,/^©\s*20/im,/^sent from my/im];
  const lines = text.split('\n');
  for (let i = 0; i < lines.length; i++) {
    if (cuts.some(rx => rx.test(lines[i].trim()))) return lines.slice(0, i).join('\n').trim();
  }
  return text.trim();
}

// ── BRAIN CONTEXT ─────────────────────────────────────────────────────────────
async function getBrainContext(fromEmail) {
  const res = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  })).catch(() => null);
  const rows = res?.data?.values || [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) {
      return { formatType: rows[i][4]||'', whatWorks: rows[i][6]||'',
        watchOutFor: rows[i][7]||'', timesSent: safeInt(rows[i][3]) };
    }
  }
  return null;
}

// ── LIGHTWEIGHT DEAL PRE-SCREEN (only for unknown senders) ───────────────────
async function isActuallyADeal(from, subject, body) {
  // Only pre-screen unknown senders with short ambiguous subjects
  // Known senders and clear keyword matches skip this
  const preview = body.slice(0, 1500);
  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001', max_tokens: 10,
    messages: [{ role: 'user', content: `Does this email contain a real estate property for sale (wholesale deal, investment property, fix and flip, etc)? Reply only YES or NO.

From: ${from}
Subject: ${subject}
Body preview: ${preview}` }]
  }).catch(() => null);
  if (!res) return true; // if check fails, assume it's a deal
  const answer = res.content[0].text.trim().toUpperCase();
  return answer.startsWith('Y');
}

// ── EXTRACTION ────────────────────────────────────────────────────────────────
async function extractProperties(from, subject, body) {
  const cleanBody = stripFooters(body);
  const brain = await getBrainContext(from);
  let hint = '';
  if (brain?.timesSent > 0) {
    hint = `\n\nKNOWN SENDER (${brain.timesSent} prior emails): format=${brain.formatType}. ${brain.whatWorks}. Watch: ${brain.watchOutFor}.`;
  }

  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001', max_tokens: 4000,
    messages: [{ role: 'user', content: `You are a real estate data extraction engine for Coralstone Capital Group.
Extract EVERY property from this wholesale deal email as a JSON array — one object per property.
If this is a forwarded email (contains "Fwd:" or "---------- Forwarded message"), extract from the forwarded content.
If no specific properties are mentioned but a URL or link is present, still extract whatever data is available.
If there are no extractable properties at all, return an empty array [].${hint}

Fields per object (null if missing):
address, city, state, zip, county, subdivision,
beds(number), baths(number), half_baths(number), sqft(number), lot_sqft(number), lot_acres(number),
year_built(number), stories(number), construction, foundation,
pool, pool_notes, garage, garage_spaces, carport, basement, attic, overall_condition,
roof_type, roof_age, ac_year, water_heater, electrical, plumbing, windows, flooring,
kitchen_notes, bath_notes,
asking_price(number — find: asking/price/sales price/$amounts),
arv(number — find: ARV/after repair value/after repairs/retail value),
repairs_estimate(number), assignment_fee(number), equity,
rent_current, rent_market, annual_taxes(number), hoa_fee,
close_date, inspection_period, earnest_money, financing_terms, cash_only,
contact_1_name, contact_1_title, contact_1_company, contact_1_phone, contact_1_phone_2,
contact_1_email, contact_1_website,
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

Return ONLY valid JSON array. No markdown.` }]
  });

  try {
    const text = res.content[0].text.trim().replace(/^```json\s*/,'').replace(/^```\s*/,'').replace(/\s*```$/,'').trim();
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) { console.error('Parse error:', e.message); return null; }
}

// ── ROW BUILDER ───────────────────────────────────────────────────────────────
function buildRow(p, subject, uid, propType) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7*24*60*60*1000);
  return [
    now.toISOString(), false, false, false, v(p.asking_price), v(p.arv), expires.toISOString(), propType,
    v(p.address), v(p.city), v(p.state), v(p.zip), v(p.county), v(p.subdivision),
    v(p.beds), v(p.baths), v(p.half_baths), v(p.sqft), v(p.lot_sqft), v(p.lot_acres),
    v(p.year_built), v(p.stories), v(p.construction), v(p.foundation),
    v(p.pool), v(p.pool_notes), v(p.garage), v(p.garage_spaces),
    v(p.carport), v(p.basement), v(p.attic),
    v(p.overall_condition), v(p.roof_type), v(p.roof_age), v(p.ac_year),
    v(p.water_heater), v(p.electrical), v(p.plumbing), v(p.windows), v(p.flooring),
    v(p.kitchen_notes), v(p.bath_notes),
    v(p.repairs_estimate), v(p.assignment_fee), v(p.equity),
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
    '0',  // Days Active — updated by updateDaysActive() each poll
    v(subject), v(uid)
  ];
}

// ── SHEET INIT ────────────────────────────────────────────────────────────────
async function initSheet() {
  const s = getSheets();
  // Check current col E header — if it's NOT Asking Price, rewrite headers
  const headerRow = await sc(() => s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1:CV1'
  })).catch(() => null);
  const currentHeaders = headerRow?.data?.values?.[0] || [];
  const a1val = currentHeaders[0];
  const colEval = currentHeaders[4]; // index 4 = column E

  if (!a1val) {
    // Empty sheet — write headers fresh
    await sc(() => s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [ACTIVE_HEADERS] }
    }));
    seen = new Set(); seenSheetReady = false;
    console.log(`Sheet initialized: ${ACTIVE_HEADERS.length} columns`);
  } else if (colEval !== 'Asking Price') {
    // Headers are in old order — rewrite header row only (don't touch data)
    // Clear row 1 and rewrite with correct order
    await sc(() => s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [ACTIVE_HEADERS] }
    }));
    // Also clear data rows — they're in wrong column order
    // Reset seen so Derek reprocesses and writes fresh with correct layout
    await sc(() => s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A2:CV' }));
    seen = new Set(); seenSheetReady = false;
    console.log(`Headers fixed. Data cleared for reprocessing with correct column order.`);
  } else {
    console.log(`Sheet active — col E="${colEval}" ✓`);
  }

  // Get existing tabs, create any missing ones
  const meta = await sc(() => s.spreadsheets.get({ spreadsheetId: SHEET_ID })).catch(() => null);
  const existingTabs = new Set((meta?.data?.sheets||[]).map(sh => sh.properties.title));

  const requiredTabs = [
    ['Deal Storage', ACTIVE_HEADERS],
    ['Price Changes', PRICE_CHANGE_HEADERS],
    ['Errors', ['Date','From','Subject','UID','Error']],
    ["Derek's Brain", BRAIN_HEADERS],
    ['Seen', ['Email UID']],
    ['Wholesaler Directory', DIR_HEADERS]
  ];

  for (const [tab, headers] of requiredTabs) {
    if (!existingTabs.has(tab)) {
      await sc(() => s.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: tab } } }] }
      }));
      console.log(`Created tab: ${tab}`);
    }
    const check = await sc(() => s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${tab}!A1:A1`
    })).catch(() => null);
    if (!check?.data?.values?.[0]?.[0]) {
      await sc(() => s.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `${tab}!A1`,
        valueInputOption: 'RAW', requestBody: { values: [headers] }
      }));
      console.log(`${tab} headers written`);
    }
  }
}

// ── ERROR LOG ─────────────────────────────────────────────────────────────────
async function logError(from, subject, uid, error) {
  try {
    await sc(() => getSheets().spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Errors!A:A',
      valueInputOption: 'RAW',
      requestBody: { values: [[new Date().toISOString(), from, subject, uid, error]] }
    }));
  } catch (e) {
    if (e.message?.includes('Unable to parse range')) {
      try {
        const s = getSheets();
        await sc(() => s.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: { requests: [{ addSheet: { properties: { title: 'Errors' } } }] }
        }));
        await sc(() => s.spreadsheets.values.update({
          spreadsheetId: SHEET_ID, range: 'Errors!A1',
          valueInputOption: 'RAW',
          requestBody: { values: [['Date','From','Subject','UID','Error']] }
        }));
      } catch {}
    }
  }
  console.log(`⚠️ Error logged: ${error}`);
}

// ── PRICE CHANGE ──────────────────────────────────────────────────────────────
async function logPriceChange(p, subject, oldPrice, newPrice) {
  const change = newPrice - oldPrice;
  const pct = oldPrice ? ((change/oldPrice)*100).toFixed(1) : '?';
  await sc(() => getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Price Changes!A:A',
    valueInputOption: 'RAW',
    requestBody: { values: [[new Date().toISOString(), v(p.address), v(p.city), v(p.state), v(p.zip),
      oldPrice, newPrice, change, `${pct}%`, v(p.arv),
      v(p.contact_1_name), v(p.contact_1_phone), v(p.contact_1_email),
      v(p.wholesaler_company), subject]] }
  }));
  console.log(`  ${change<0?'📉':'📈'} Price change: ${p.address} $${oldPrice}→$${newPrice} (${pct}%)`);
}

// ── CHECKBOX PROCESSOR ────────────────────────────────────────────────────────
async function processCheckboxes() {
  const res = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  })).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const passIdx = COL['Pass'], soldIdx = COL['Sold'];
  const headers = rows[0], keep = [headers], passed = [], sold = [];

  for (let i = 1; i < rows.length; i++) {
    const isPass = rows[i][passIdx] === 'TRUE' || rows[i][passIdx] === true;
    const isSold = rows[i][soldIdx] === 'TRUE' || rows[i][soldIdx] === true;
    if (isPass) passed.push(rows[i]);
    else if (isSold) { const r=[...rows[i]]; r[passIdx]=false; r[soldIdx]=false; r[COL['Review']]='SOLD'; sold.push(r); }
    else keep.push(rows[i]);
  }
  if (!passed.length && !sold.length) return;

  await sc(() => getSheets().spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ' }));
  await sc(() => getSheets().spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: keep }
  }));
  if (passed.length) {
    await sc(() => getSheets().spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
      valueInputOption: 'RAW', requestBody: { values: passed }
    }));
    console.log(`🗑️  ${passed.length} passed → Deal Storage`);
  }
  if (sold.length) {
    await sc(() => getSheets().spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
      valueInputOption: 'RAW', requestBody: { values: sold }
    }));
    console.log(`🏠 ${sold.length} sold → Deal Storage`);
  }
}

// ── DAYS ACTIVE (batch update, correct column letter) ─────────────────────────
async function updateDaysActive() {
  const res = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  })).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const dateIdx = COL['Date Received'];
  const daysIdx = COL['Days Active'];
  const col = colLetter(daysIdx);  // correct even for col 97+
  const now = new Date();
  const updates = [];

  for (let i = 1; i < rows.length; i++) {
    const received = rows[i][dateIdx];
    if (!received) continue;
    const days = Math.floor((now - new Date(received)) / (1000*60*60*24));
    const current = safeInt(rows[i][daysIdx]);
    if (days !== current) updates.push({ range: `Active Deals!${col}${i+1}`, values: [[days]] });
  }
  if (!updates.length) return;

  await sc(() => getSheets().spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { valueInputOption: 'RAW', data: updates }
  }));
  console.log(`📅 Days Active updated: ${updates.length} row(s) | col ${col}`);
}

// ── AUTO-ARCHIVE (7 days) ─────────────────────────────────────────────────────
async function archiveExpired() {
  const res = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ'
  })).catch(() => null);
  const rows = res?.data?.values || [];
  if (rows.length <= 1) return;

  const expiresIdx = COL['Expires'];
  const now = new Date(), headers = rows[0], active = [headers], expired = [];
  for (let i = 1; i < rows.length; i++) {
    const exp = rows[i][expiresIdx];
    if (exp && new Date(exp) < now) expired.push(rows[i]);
    else active.push(rows[i]);
  }
  if (!expired.length) return;

  await sc(() => getSheets().spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:CZ' }));
  await sc(() => getSheets().spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: active }
  }));
  await sc(() => getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Deal Storage!A:A',
    valueInputOption: 'RAW', requestBody: { values: expired }
  }));
  console.log(`📦 ${expired.length} expired → Deal Storage`);
}

// ── WHOLESALER DIRECTORY (match by email col C=index 2) ───────────────────────
async function updateWholesalerDirectory(fromEmail, company, contactName, phone, website, propType, askingPrice) {
  const res = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Wholesaler Directory!A:K'
  })).catch(() => null);
  const rows = res?.data?.values || [];
  const today = new Date().toISOString().split('T')[0];
  let existingRow = -1;

  // Match by email (column C = index 2)
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === fromEmail) { existingRow = i + 1; break; }
  }

  if (existingRow > 0) {
    const ex = rows[existingRow - 1];
    const deals = safeInt(ex[5]) + 1;
    const prevAvg = safeNum(ex[6]);
    const avg = prevAvg === 0 ? 1 : ((prevAvg * (deals-1)) + 1) / deals;
    const types = (ex[8]||'').split(',').map(t=>t.trim()).filter(Boolean);
    if (propType && !types.includes(propType)) types.push(propType);
    const prevAvgAsk = safeNum(ex[9]);
    const newAvgAsk = askingPrice ? ((prevAvgAsk*(deals-1)) + safeNum(askingPrice)) / deals : prevAvgAsk;

    await sc(() => getSheets().spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `Wholesaler Directory!A${existingRow}`,
      valueInputOption: 'RAW', requestBody: { values: [[
        company || ex[0] || '',
        contactName || ex[1] || '',
        fromEmail,
        phone || ex[3] || '',
        website || ex[4] || '',
        deals,
        avg.toFixed(1),
        today,
        types.join(', '),
        newAvgAsk ? Math.round(newAvgAsk).toString() : ex[9]||'',
        ex[10]||''
      ]] }
    }));
  } else {
    await sc(() => getSheets().spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Wholesaler Directory!A:A',
      valueInputOption: 'RAW', requestBody: { values: [[
        company||'', contactName||'', fromEmail, phone||'', website||'',
        1, '1.0', today, propType||'',
        askingPrice ? String(Math.round(safeNum(askingPrice))) : '', ''
      ]] }
    }));
  }
}

// ── BRAIN UPDATE ──────────────────────────────────────────────────────────────
async function updateBrain(fromEmail, company, subject, propCount, body) {
  const res = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  })).catch(() => null);
  const rows = res?.data?.values || [];
  const now = new Date().toISOString();
  const hasEmoji = /[\u{1F300}-\u{1FFFF}]/u.test(body);
  const hasBullets = /^[-•*]/m.test(body);
  const hasNumbered = /^\d+\./m.test(body);
  const hasTable = / \| /.test(body);
  const fmt = hasEmoji?'Emoji bullets':hasNumbered?'Numbered list':hasBullets?'Bullet list':hasTable?'Table':'Plain text';

  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === fromEmail) { existingRow = i+1; break; }
  }

  if (existingRow > 0) {
    const ex = rows[existingRow-1];
    const times = safeInt(ex[3]) + 1;
    const prevAvg = safeNum(ex[8]) || 0;
    const avg = ((prevAvg*(times-1)) + propCount) / times;
    await sc(() => getSheets().spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `Derek's Brain!A${existingRow}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[fromEmail, company||ex[1]||'', now, times, fmt, ex[5]||'', ex[6]||'', ex[7]||'', avg.toFixed(1), ex[9]||'']] }
    }));
    console.log(`🧠 ${fromEmail} (${times} emails, avg ${avg.toFixed(1)} props)`);
  } else {
    await sc(() => getSheets().spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: "Derek's Brain!A:A",
      valueInputOption: 'RAW',
      requestBody: { values: [[fromEmail, company||'', now, 1, fmt, 'address,beds,baths,sqft,asking,arv,drive link,phone', 'Check Google Drive links per property', '', propCount.toString(), `First: ${subject}`]] }
    }));
    console.log(`🧠 New wholesaler: ${fromEmail}`);
    knownSenders.add(fromEmail); // add to in-memory set immediately
  }
}

// ── BACKFILL WHOLESALER DIRECTORY FROM BRAIN (one-time on boot if empty) ──────
async function backfillWholesalerDirectory() {
  const dirCheck = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Wholesaler Directory!A2:C2'
  })).catch(() => null);
  if (dirCheck?.data?.values?.length) {
    console.log('📋 Wholesaler Directory has data — skipping backfill');
    return;
  }

  const brain = await sc(() => getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: "Derek's Brain!A:J"
  })).catch(() => null);
  const rows = brain?.data?.values || [];
  if (rows.length <= 1) return;

  const today = new Date().toISOString().split('T')[0];
  const entries = [];

  for (let i = 1; i < rows.length; i++) {
    const email = v(rows[i][0]);
    const company = v(rows[i][1]);
    const timesSent = v(rows[i][3]);
    const avgProps = v(rows[i][8]);
    if (!email || email.includes('@') === false) continue; // skip non-emails

    entries.push([
      company,                                    // A: Company
      '',                                         // B: Contact Name
      email,                                      // C: Email  ← key match field
      '',                                         // D: Phone
      '',                                         // E: Website
      safeInt(timesSent) || 1,                    // F: Deals Sent
      (safeNum(avgProps) || 1).toFixed(1),        // G: Avg Props/Email
      today,                                      // H: Last Deal Date
      '',                                         // I: Property Types
      '',                                         // J: Avg Asking Price
      ''                                          // K: Notes
    ]);
  }

  if (!entries.length) return;
  await sc(() => getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Wholesaler Directory!A:A',
    valueInputOption: 'RAW', requestBody: { values: entries }
  }));
  console.log(`📋 Backfilled ${entries.length} wholesalers into Directory`);
}

// ── IMAP CONNECT WITH RETRY ───────────────────────────────────────────────────
async function connectImap(retries = 3) {
  for (let i = 0; i < retries; i++) {
    const client = new ImapFlow({
      host: process.env.IMAP_HOST, port: 993, secure: true,
      auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
      logger: false, socketTimeout: 120000, connectionTimeout: 60000,
      tls: { rejectUnauthorized: false }
    });
    client.on('error', e => console.error('IMAP socket (handled):', e.message));
    try { await client.connect(); return client; }
    catch (e) {
      if (i === retries-1) throw e;
      console.error(`IMAP retry ${i+1}:`, e.message);
      await new Promise(r => setTimeout(r, 3000));
    }
  }
}

// ── MAIN POLL ─────────────────────────────────────────────────────────────────
async function poll() {
  console.log(`\n[${new Date().toLocaleTimeString()}] Polling...`);
  await processCheckboxes();
  await updateDaysActive();
  await archiveExpired();

  let client, written = 0, dupes = 0, priceChanges = 0, errors = 0;
  seenDirty = false;

  try {
    client = await connectImap();
    const lock = await client.getMailboxLock('INBOX');

    try {
      const since = new Date(Date.now() - 7*24*60*60*1000);

      // Phase 1: envelopes only — queue ALL new unseen emails for phase 2 screening
      // We never filter at envelope level for unknown senders — Haiku screens them in phase 2
      const candidates = [];
      for await (const msg of client.fetch({ since }, { envelope: true, uid: true })) {
        const uid = String(msg.uid);
        if (seen.has(uid)) continue;
        const subj = msg.envelope?.subject || '';
        const from = msg.envelope?.from?.[0]?.address || '';
        candidates.push({ uid, subj, from });
      }
      console.log(`${candidates.length} new email(s) to screen`);

      // Phase 2: full body + extract
      for (const c of candidates) {
        console.log(`\n📬 ${c.subj}`);
        try {
          const msgData = await client.fetchOne(c.uid, { source: true }, { uid: true });
          const parsed = await simpleParser(msgData.source);
          // Extract text body — also pull HTML text nodes for image-heavy emails
          const htmlBody = parsed.html || '';
          let rawBody = parsed.text || '';

          // Always try to extract text from HTML — even if text body exists
          // Image-heavy emails often have critical data only in HTML attributes/links
          if (htmlBody && rawBody.trim().length < 5000) {
            const htmlText = htmlBody
              .replace(/<style[^>]*>.*?<\/style>/gis, '')
              .replace(/<script[^>]*>.*?<\/script>/gis, '')
              .replace(/<img[^>]+alt="([^"]+)"/gi, ' [IMG: $1] ')
              .replace(/<td[^>]*>/gi, ' ')
              .replace(/<tr[^>]*>/gi, '\n')
              .replace(/<br\s*\/?>/gi, '\n')
              .replace(/<p[^>]*>/gi, '\n')
              .replace(/<div[^>]*>/gi, '\n')
              .replace(/<a[^>]+href="(https?[^"]+)"[^>]*>([^<]*)<\/a>/gi, ' $2 (link: $1) ')
              .replace(/<[^>]+>/g, ' ')
              .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
              .replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
            if (htmlText.length > rawBody.trim().length) {
              rawBody = htmlText;
              if (rawBody.length > 200) console.log(`  📄 Using HTML text (${rawBody.length} chars)`);
            }
          }

          // Extract image URLs from HTML for Vision processing
          const imgUrls = [];
          if (htmlBody) {
            const imgMatches = [...htmlBody.matchAll(/<img[^>]+src="(https?:\/\/[^"]+)"/gi)];
            imgUrls.push(...imgMatches.map(m => m[1]).slice(0, 3)); // max 3 images
          }

          // If body still thin AND we have images, use Vision to read them
          const isBodyThin = rawBody.trim().length < 300;
          if (isBodyThin && imgUrls.length > 0) {
            console.log(`  🖼️ Thin body (${rawBody.trim().length} chars) + ${imgUrls.length} images — using Vision`);
            try {
              const visionContent = [];
              for (const url of imgUrls) {
                try {
                  const imgRes = await fetch(url, {
                    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Coralstone/1.0)' },
                    timeout: 8000
                  });
                  if (!imgRes.ok) continue;
                  const arrBuf = await imgRes.arrayBuffer();
                  const buf = Buffer.from(arrBuf);
                  const mime = (imgRes.headers.get('content-type') || 'image/jpeg').split(';')[0].trim();
                  // Only use images (skip non-image content types)
                  if (!mime.startsWith('image/')) continue;
                  // Only use if reasonable size (< 4MB base64 limit)
                  if (buf.length < 4 * 1024 * 1024) {
                    visionContent.push({
                      type: 'image',
                      source: { type: 'base64', media_type: mime, data: buf.toString('base64') }
                    });
                  }
                } catch(imgErr) { console.log(`  Image fetch failed: ${imgErr.message}`); }
              }
              if (visionContent.length > 0) {
                visionContent.push({
                  type: 'text',
                  text: `Email subject: ${c.subj}\nFrom: ${c.from}\nExtract ALL visible text from this real estate email image — addresses, prices, beds/baths, contact info, ARV, repairs, everything you can read.`
                });
                const visionRes = await anthropic.messages.create({
                  model: 'claude-haiku-4-5-20251001', max_tokens: 2000,
                  messages: [{ role: 'user', content: visionContent }]
                });
                const visionText = visionRes.content[0].text;
                console.log(`  🖼️ Vision extracted ${visionText.length} chars`);
                rawBody = visionText + '\n\n' + rawBody; // prepend vision text
              }
            } catch(visionErr) {
              console.log(`  Vision error: ${visionErr.message}`);
            }
          }

          // Smart screening: known senders + strong keywords → skip Haiku entirely
          // Everyone else gets a fast Haiku YES/NO check (~$0.0001 each)
          const fastPass = isDealEmail(c.subj, c.from);
          if (!fastPass) {
            console.log(`  🤖 Haiku screening: "${c.subj}" from ${c.from}`);
            const isDeal = await isActuallyADeal(c.from, c.subj, rawBody);
            if (!isDeal) {
              console.log(`  ⏩ Not a deal — skipping`);
              seen.add(c.uid); seenDirty = true;
              await new Promise(r => setTimeout(r, 200));
              continue;
            }
            console.log(`  ✅ Haiku confirmed deal — extracting`);
          } else {
            console.log(`  ✅ Fast pass — extracting`);
          }

          let properties = await extractProperties(c.from, c.subj, rawBody);

          // Fallback: if extraction fails, try with ONLY the subject line + sender
          // This catches cases where body parsing fails but subject has key info
          if (!properties || properties.length === 0) {
            console.log('  → Extraction failed — trying subject-only fallback');
            properties = await extractProperties(c.from, c.subj,
              `Subject: ${c.subj}\nFrom: ${c.from}\nNote: Email body could not be parsed. Extract what you can from subject line only.`
            );
          }

          if (!properties || properties.length === 0) {
            console.log('  → No properties extracted');
            await logError(c.from, c.subj, c.uid, 'Extraction returned 0 properties');
            errors++;
          } else {
            console.log(`  → ${properties.length} propert${properties.length>1?'ies':'y'}`);
            await updateBrain(c.from, properties[0]?.wholesaler_company||'', c.subj, properties.length, rawBody);

            for (const p of properties) {
              // Accept redacted addresses (e.g. "XXXX, City FL") — wholesaler withholds until called
              const isRedacted = !p.address || /^x+$/i.test((p.address || '').trim()) || (p.address || '').trim().length < 4;
              if (isRedacted && !p.city) { console.log('  → Skipped (no address or city)'); continue; }
              if (isRedacted) {
                p.address = 'XXXX'; // redacted — call wholesaler
                p.highlights = (p.highlights ? p.highlights + ' | ' : '') + 'ADDRESS REDACTED — CALL WHOLESALER FOR ADDRESS';
                console.log(`  ⚠️ Redacted address — saving with XXXX placeholder for ${p.city}, ${p.state}`);
              }
              const propType = detectPropertyType(p);
              const dupeAddr = isRedacted ? `XXXX-${p.city}-${p.zip}` : p.address;
              const dupe = await checkDuplicate(dupeAddr, p.city, p.zip, p.asking_price);

              if (dupe.isDupe && dupe.isPriceChange) {
                await logPriceChange(p, c.subj, dupe.oldPrice, dupe.newPrice);
                priceChanges++;
              } else if (dupe.isDupe) {
                console.log(`  🔁 DUPE: ${p.address} (in ${dupe.tab})`);
                dupes++;
              } else {
                // Find first empty row in col A (ignores checkbox-only rows)
                const colARes = await sc(() => getSheets().spreadsheets.values.get({
                  spreadsheetId: SHEET_ID, range: 'Active Deals!A:A'
                }));
                const colAVals = colARes?.data?.values || [];
                const nextRow = colAVals.length + 1; // 1-indexed, after last filled A cell
                await sc(() => getSheets().spreadsheets.values.update({
                  spreadsheetId: SHEET_ID, range: `Active Deals!A${nextRow}`,
                  valueInputOption: 'RAW',
                  requestBody: { values: [buildRow(p, c.subj, c.uid, propType)] }
                }));
                await updateWholesalerDirectory(c.from, v(p.wholesaler_company),
                  v(p.contact_1_name), v(p.contact_1_phone), v(p.contact_1_website),
                  propType, v(p.asking_price));
                console.log(`  ✅ [${propType}] ${p.address}, ${p.city} | Ask:$${p.asking_price} ARV:$${p.arv}`);
                written++;
                // Auto-underwrite via Urban (non-blocking)
                // Use address as the lookup key — Urban matches deals by address
                triggerUrbanUnderwrite(
                  v(p.address), v(p.city), v(p.state), v(p.zip), v(p.address)
                ).catch(e => console.log(`  Urban trigger err: ${e.message}`));
              }
              await new Promise(r => setTimeout(r, 300));
            }
          }

          seen.add(c.uid); seenDirty = true;
          await new Promise(r => setTimeout(r, 500));
        } catch (e) {
          console.error(`  Error UID ${c.uid}:`, e.message);
          // Don't mark seen on connection errors — retry next poll
          const isConnErr = e.message?.includes('Connection not available') ||
                            e.message?.includes('ECONNRESET') ||
                            e.message?.includes('ETIMEDOUT') ||
                            e.message?.includes('socket') ||
                            e.message?.includes('network');
          if (isConnErr) {
            console.log(`  ↩️ Connection error — will retry UID ${c.uid} next poll`);
          } else {
            await logError(c.from, c.subj, c.uid, e.message);
            seen.add(c.uid); seenDirty = true;
          }
          errors++;
        }
      }
    } finally { lock.release(); }

    // Flush seen ONCE at end of poll — not after every email
    await flushSeen();
    console.log(`\n✅ Done — ${written} new | ${dupes} dupes | ${priceChanges} price changes | ${errors} errors`);
  } catch (e) {
    console.error('Poll error:', e.message);
    await flushSeen();
  } finally {
    if (client) { try { await client.logout(); } catch {} }
  }
}

// ── TRIGGER URBAN AUTO-UNDERWRITE ────────────────────────────────────────────
async function triggerUrbanUnderwrite(address, city, state, zip, uid) {
  try {
    console.log(`  🏙️ Triggering Urban auto-underwrite for ${address}...`);
    // Hit Urban's underwrite endpoint — fire and forget (don't await full response)
    // Urban uses SSE streaming so we just open the connection and let it run
    // Encode the address as the uid — Urban will match by address
    const encUid = encodeURIComponent(uid);
    const res = await fetch(`${URBAN_URL}/api/underwrite-by-address/${encUid}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-urban-token': URBAN_TOKEN
      },
      body: JSON.stringify({ forceRefresh: false, deep: false }),
      signal: AbortSignal.timeout(180000) // 3 min timeout
    });
    if (!res.ok) {
      console.log(`  ⚠️ Urban returned ${res.status}`);
      return;
    }
    // Drain the SSE stream
    const reader = res.body.getReader();
    const dec = new TextDecoder();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = dec.decode(value, { stream: true });
      const lines = chunk.split('\n').filter(l => l.startsWith('data: '));
      for (const line of lines) {
        try {
          const data = JSON.parse(line.slice(6));
          if (data.status) console.log(`  🏙️ Urban: ${data.status}`);
          if (data.done) {
            const uw = data.underwrite;
            console.log(`  ✅ Urban verdict: ${uw.verdict} (${uw.score}/10) — ${uw.verdictReason}`);
          }
          if (data.error) console.log(`  ❌ Urban error: ${data.error}`);
        } catch {}
      }
    }
  } catch (e) {
    console.log(`  ⚠️ Urban trigger failed: ${e.message}`);
  }
}

// ── FORCE REPROCESS ──────────────────────────────────────────────────────────
// Clears the top N highest UIDs from the seen set so Derek re-evaluates them
// UIDs are IMAP sequential — highest = most recent
// Safe: dupe detection prevents re-writing already-saved deals
function clearRecentFromSeen(count = 50) {
  try {
    if (seen.size === 0) return;
    // Sort all seen UIDs numerically, take the top `count`
    const sorted = [...seen].map(Number).filter(n => !isNaN(n)).sort((a, b) => b - a);
    const toRemove = sorted.slice(0, count);
    let removed = 0;
    for (const uid of toRemove) {
      seen.delete(String(uid));
      removed++;
    }
    if (removed > 0) {
      seenDirty = true;
      console.log(`🔄 Cleared top ${removed} recent UIDs from seen — will reprocess`);
    }
  } catch(e) {
    console.log('clearRecentFromSeen error:', e.message);
  }
}

// ── BOOT ──────────────────────────────────────────────────────────────────────
const POLL_MS = parseInt(process.env.POLL_INTERVAL || '300000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS/60000}min`);
console.log(`Days Active col: ${colLetter(COL['Days Active'])} | Address col: ${colLetter(COL['Address'])} | Asking col: ${colLetter(COL['Asking Price'])}`);

initSheet()
  .then(() => loadSeenFromSheet())
  .then(() => {
    // Clear the most recent UIDs so failed extractions get retried
    // Dupe detection prevents re-writing already-saved deals
    clearRecentFromSeen(150); // clear top 150 most recent UIDs
    return loadKnownSenders();
  })
  .then(() => backfillWholesalerDirectory())
  .then(() => { poll(); setInterval(poll, POLL_MS); })
  .catch(e => { console.error('Init error:', e.message); poll(); setInterval(poll, POLL_MS); });
