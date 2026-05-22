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

// ALL the columns — everything Derek should capture
const HEADERS = [
  'Date Received',
  // Property
  'Address', 'City', 'State', 'Zip', 'County', 'Subdivision',
  // Structure
  'Beds', 'Baths', 'Half Baths', 'Sqft', 'Lot Sqft', 'Lot Acres', 'Year Built', 'Stories',
  'Construction', 'Foundation', 'Pool', 'Pool Notes', 'Garage', 'Garage Spaces',
  'Carport', 'Basement', 'Attic',
  // Condition / Systems
  'Overall Condition', 'Roof Type', 'Roof Age / Year', 'AC Year / Age',
  'Water Heater', 'Electrical', 'Plumbing', 'Windows', 'Flooring',
  'Kitchen Notes', 'Bath Notes',
  // Financials
  'Asking Price', 'ARV', 'Repairs Estimate', 'Assignment Fee', 'Equity',
  'Rent Current', 'Rent Market', 'Annual Taxes', 'HOA Fee',
  // Deal terms
  'Close Date', 'Inspection Period', 'Earnest Money', 'Financing Terms', 'Cash Only',
  // Contacts — ALL of them
  'Contact 1 Name', 'Contact 1 Title', 'Contact 1 Company',
  'Contact 1 Phone', 'Contact 1 Phone 2', 'Contact 1 Email', 'Contact 1 Website',
  'Contact 2 Name', 'Contact 2 Title', 'Contact 2 Company',
  'Contact 2 Phone', 'Contact 2 Email',
  'Contact 3 Name', 'Contact 3 Phone', 'Contact 3 Email',
  'ALL Phones Found', 'ALL Emails Found', 'ALL Names Found',
  // Seller
  'Seller Name', 'Seller Phone', 'Seller Situation', 'Seller Motivation', 'Occupancy',
  // Location flags
  'Flood Zone', 'HOA', 'School District',
  // Links
  'Google Drive Link', 'Zillow Link', 'Google Maps Link', 'All Other Links',
  // Photos
  'Photos Included', 'Photo Count', 'Photo Links',
  // Comps
  'Comp 1', 'Comp 2', 'Comp 3',
  // Highlights
  'What Is Updated', 'What Needs Work', 'Highlights', 'Red Flags', 'Additional Notes',
  // Meta
  'Wholesaler Company', 'List Name', 'Email Subject', 'Email UID'
];

async function initSheet() {
  const s = getSheets();
  const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A1:A1' }).catch(() => null);
  if (!res?.data?.values) {
    await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals' }).catch(() => {});
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [HEADERS] }
    });
    console.log('Sheet initialized with', HEADERS.length, 'columns');
  }
}

async function writeRow(values) {
  const s = getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:A',
    valueInputOption: 'RAW', requestBody: { values: [values] }
  });
}

// Keyword screen — zero API calls
const DEAL_WORDS = ['off market','wholesale','arv','flip','motivated','for sale','equity',
  'distressed','foreclosure','probate','deal','available deals','fix flip','cash buyer',
  'assignment','sales price','asking price','beds','baths','sqft','under roof','bungalow',
  'rehab','investment property','property available','price reduction','reduced'];
const isDeal = (subj, from) => DEAL_WORDS.some(k => `${subj} ${from}`.toLowerCase().includes(k));

const v = x => (x === null || x === undefined) ? '' : String(x);

async function extractProperties(from, subject, body) {
  const prompt = `You are a real estate data extraction engine for Coralstone Capital Group. 
Your job: extract EVERY single piece of information from this wholesale deal email into structured data.
This email may contain MULTIPLE properties — extract ALL of them as separate objects in the array.
Miss nothing. If a field isn't in the email, use null.

Return a JSON array (one object per property). Each object must have ALL these fields:

PROPERTY LOCATION:
- address (street only, e.g. "11256 Holbrook St")
- city
- state (2-letter)
- zip
- county
- subdivision (neighborhood/subdivision name if mentioned)

STRUCTURE:
- beds (number)
- baths (number)
- half_baths (number)
- sqft (number — square footage under roof/living area)
- lot_sqft (number)
- lot_acres (number)
- year_built (number)
- stories (number)
- construction (e.g. "Block", "Frame", "Stucco/Frame", "CBS")
- foundation (e.g. "Slab", "Crawl")
- pool ("YES" or "NO", include notes like "Recently Resurfaced")
- pool_notes (details about pool condition, equipment, lanai, etc.)
- garage (e.g. "2 Car", "1 Car", "None")
- garage_spaces (number)
- carport ("YES" or "NO" or sqft)
- basement ("YES" or "NO")
- attic ("YES" or "NO")

CONDITION / SYSTEMS:
- overall_condition (e.g. "Fair", "Good", "Needs Work")
- roof_type (e.g. "Shingle", "Tile", "Metal")
- roof_age (e.g. "2021 Roof", "Needs New Roof", "2022 Roof")
- ac_year (e.g. "2013 A/C", "2020 A/C", "2025 A/C")
- water_heater (e.g. "2024 Tankless", "Standard")
- electrical (e.g. "Good", "Updated", "2025 Panel Upgrade")
- plumbing (e.g. "Good", "Copper", "PVC")
- windows (e.g. "Hurricane Impact 2023", "Standard", "New")
- flooring (e.g. "Tile", "Carpet", "Mixed")
- kitchen_notes (condition/features)
- bath_notes (condition/features)

FINANCIALS:
- asking_price (number only, no symbols)
- arv (number only — After Repair Value stated)
- repairs_estimate (number only if stated)
- assignment_fee (number only if stated)
- equity (number or % if stated)
- rent_current (monthly rent if tenant-occupied)
- rent_market (market rent estimate if stated)
- annual_taxes (number)
- hoa_fee (number, monthly or annual — note which)

DEAL TERMS:
- close_date (e.g. "05/29/2026")
- inspection_period (e.g. "7 days")
- earnest_money (e.g. "$1,000", "1%")
- financing_terms (e.g. "Cash Only", "Conventional OK")
- cash_only ("YES" or "NO")

CONTACTS (extract EVERY person mentioned):
- contact_1_name
- contact_1_title (e.g. "Acquisition Manager")
- contact_1_company
- contact_1_phone
- contact_1_phone_2 (second number if any)
- contact_1_email
- contact_1_website
- contact_2_name
- contact_2_title
- contact_2_company
- contact_2_phone
- contact_2_email
- contact_3_name
- contact_3_phone
- contact_3_email
- all_phones (ALL phone numbers found ANYWHERE in the email, comma separated)
- all_emails (ALL email addresses found ANYWHERE in the email, comma separated)
- all_names (ALL person names found ANYWHERE in the email, comma separated)

SELLER INFO:
- seller_name (if mentioned)
- seller_phone (if mentioned)
- seller_situation (e.g. "Motivated", "Probate", "Divorce", "Pre-foreclosure")
- seller_motivation (description)
- occupancy (e.g. "Vacant", "Owner Occupied", "Tenant Occupied")

LOCATION FLAGS:
- flood_zone ("YES" / "NO" / zone code like "AE")
- hoa ("YES" or "NO")
- school_district

LINKS (copy exactly as they appear):
- drive_link (Google Drive link for THIS specific property)
- zillow_link
- google_maps_link
- all_other_links (any other URLs, comma separated)

PHOTOS:
- photos_included ("YES" or "NO")
- photo_count (number if mentioned)
- photo_links (direct photo URLs if any)

COMPS:
- comp_1 (address + price if comps provided)
- comp_2
- comp_3

HIGHLIGHTS:
- what_is_updated (things recently done — e.g. "2022 Roof, 2020 AC, Hurricane Windows 2023")
- what_needs_work (repairs needed — e.g. "Needs Roof, 2013 AC, Full Remodel")
- highlights (standout features — e.g. "Pool, Huge Lot, No Flood Zone, Heart of Spring Hill")
- red_flags (concerns)
- additional_notes (anything else important not captured above)

META:
- wholesaler_company (company sending the email)
- list_name (name of the list/campaign if mentioned)

FROM: ${from}
SUBJECT: ${subject}
BODY:
${body.slice(0, 12000)}

IMPORTANT: Return ONLY a valid JSON array. No markdown. No explanation. No text before or after the array.`;

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
    console.error('JSON parse error:', e.message);
    console.error('Raw response:', res.content[0].text.slice(0, 500));
    return null;
  }
}

function buildRow(p, subject, uid) {
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
    v(subject), v(uid)
  ];
}

async function poll() {
  console.log(`\n[${new Date().toLocaleTimeString()}] Polling ${process.env.IMAP_USER}...`);

  const client = new ImapFlow({
    host: process.env.IMAP_HOST, port: 993, secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false, socketTimeout: 120000, connectionTimeout: 60000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', e => console.error('IMAP socket error (handled):', e.message));

  let written = 0;

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');

    try {
      const since = new Date(Date.now() - 72 * 60 * 60 * 1000);

      // Phase 1: fetch envelopes only (no body, no cost)
      const candidates = [];
      for await (const msg of client.fetch({ since }, { envelope: true, uid: true })) {
        const uid = String(msg.uid);
        if (seen.has(uid)) continue;
        const subj = msg.envelope?.subject || '';
        const from = msg.envelope?.from?.[0]?.address || '';
        if (isDeal(subj, from)) {
          candidates.push({ uid, subj, from });
        } else {
          seen.add(uid);
        }
      }
      saveSeen();
      console.log(`${candidates.length} deal email(s) to process`);

      // Phase 2: fetch full body + extract only for deals
      for (const c of candidates) {
        console.log(`\n📬 Processing: ${c.subj}`);
        try {
          const msgData = await client.fetchOne(c.uid, { source: true }, { uid: true });
          const parsed = await simpleParser(msgData.source);
          const body = parsed.text || (parsed.html || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

          const properties = await extractProperties(c.from, c.subj, body);

          if (!properties || properties.length === 0) {
            console.log('  → No properties extracted');
          } else {
            console.log(`  → ${properties.length} propert${properties.length > 1 ? 'ies' : 'y'} found`);
            for (const p of properties) {
              if (!p.address) { console.log('  → Skipped (no address)'); continue; }
              const row = buildRow(p, c.subj, c.uid);
              await writeRow(row);
              console.log(`  ✅ ${p.address}, ${p.city} | Ask: $${p.asking_price} | ARV: $${p.arv}`);
              written++;
            }
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

    console.log(`\n✅ Poll complete — ${written} new deal${written !== 1 ? 's' : ''} written to sheet`);
  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    try { await client.logout(); } catch {}
  }
}

const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS / 60000}min | ${HEADERS.length} columns`);

// Init sheet then start polling
initSheet().then(() => {
  poll();
  setInterval(poll, POLL_MS);
}).catch(e => {
  console.error('Init error:', e.message);
  poll();
  setInterval(poll, POLL_MS);
});
