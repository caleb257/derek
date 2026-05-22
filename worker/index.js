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
  const auth = new google.auth.JWT(creds.client_email, null, creds.private_key, ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

const HEADERS = [
  'Date','Address','City','State','Zip',
  'Beds','Baths','Sqft','Lot Sqft','Year Built',
  'Pool','Garage','Construction','Roof Type','Roof Age','AC Year',
  'Asking Price','ARV','Repairs Needed','Close Date',
  'Contact Name','Contact Phone','Contact Email','Contact Company',
  'All Phones','All Emails',
  'Seller Situation','HOA','Flood Zone','Annual Taxes',
  'Drive Link','All Links','Photo Count',
  'Highlights','What Needs Work','Notes',
  'Email Subject','Email UID'
];

async function ensureHeaders() {
  const s = getSheets();
  const check = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A1:A1' }).catch(() => null);
  if (!check?.data?.values) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
      valueInputOption: 'RAW', requestBody: { values: [HEADERS] }
    });
    console.log('✅ Headers written');
  }
}

async function clearAndReset() {
  const s = getSheets();
  await s.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: 'Active Deals!A:AZ' });
  await s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A1',
    valueInputOption: 'RAW', requestBody: { values: [HEADERS] }
  });
  console.log('✅ Sheet cleared and headers written');
}

async function appendRow(row) {
  const s = getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:A',
    valueInputOption: 'RAW', requestBody: { values: [row] }
  });
}

const KEYWORDS = ['off market','wholesale','arv','flip','motivated','for sale','opportunity','equity','distressed','foreclosure','probate','deal','available deals','fix flip','cash buyer','assignment','sales price','asking price','beds','baths','sqft','under roof','bungalow','rehab'];
const isDeal = (subject, from) => KEYWORDS.some(k => `${subject} ${from}`.toLowerCase().includes(k));

async function extract(from, subject, body) {
  const prompt = `You are extracting real estate deal data from a wholesale email. Extract EVERY property listed. Return a JSON array — one object per property. Do not skip any property.

For each property extract ALL of these fields (null if not found):
- address (street address only)
- city
- state  
- zip
- beds (number)
- baths (number)
- sqft (number, square feet under roof)
- lot_sqft (number)
- year_built (number)
- pool (YES/NO)
- garage (e.g. "2 Car", "1 Car", "None")
- construction (e.g. "Block", "Stucco/Frame", "Frame")
- roof_type (e.g. "Shingle", "Tile")
- roof_age (e.g. "Needs new roof", "2022 Roof", "2021 Roof")
- ac_year (e.g. "2013 A/C", "2020 A/C")
- asking_price (number only, no $ sign)
- arv (number only, no $ sign — the ARV stated)
- repairs_needed (brief description of what needs work)
- close_date (e.g. "05/29/2026")
- contact_name (person's name)
- contact_phone (phone number)
- contact_email (email address)
- contact_company (company name)
- all_phones (ALL phone numbers found anywhere in email, comma separated)
- all_emails (ALL email addresses found anywhere in email, comma separated)
- seller_situation (e.g. motivated seller, probate, divorce, etc.)
- hoa (YES/NO or amount)
- flood_zone (YES/NO or zone code)
- annual_taxes (number)
- drive_link (Google Drive or Dropbox link for THIS property)
- all_links (ALL URLs found for this property)
- photo_count (number of photos mentioned)
- highlights (notable features, comma separated — e.g. "Pool Recently Resurfaced, New Hurricane Windows 2023, HUGE Lot")
- what_needs_work (e.g. "Needs Roof, 2013 AC")
- notes (anything else important)

FROM: ${from}
SUBJECT: ${subject}
BODY:
${body.slice(0, 10000)}

Return ONLY a valid JSON array. No markdown. No explanation.`;

  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 3000,
    messages: [{ role: 'user', content: prompt }]
  });

  try {
    const text = res.content[0].text.trim().replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    console.error('Parse error:', e.message);
    return null;
  }
}

const v = x => (x === null || x === undefined) ? '' : String(x);

function buildRow(p, subject, uid) {
  return [
    new Date().toISOString(),
    v(p.address), v(p.city), v(p.state), v(p.zip),
    v(p.beds), v(p.baths), v(p.sqft), v(p.lot_sqft), v(p.year_built),
    v(p.pool), v(p.garage), v(p.construction), v(p.roof_type), v(p.roof_age), v(p.ac_year),
    v(p.asking_price), v(p.arv), v(p.repairs_needed), v(p.close_date),
    v(p.contact_name), v(p.contact_phone), v(p.contact_email), v(p.contact_company),
    v(p.all_phones), v(p.all_emails),
    v(p.seller_situation), v(p.hoa), v(p.flood_zone), v(p.annual_taxes),
    v(p.drive_link), v(p.all_links), v(p.photo_count),
    v(p.highlights), v(p.what_needs_work), v(p.notes),
    v(subject), v(uid)
  ];
}

async function poll() {
  console.log(`\n[${new Date().toLocaleTimeString()}] Polling...`);
  const client = new ImapFlow({
    host: process.env.IMAP_HOST, port: 993, secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false, socketTimeout: 120000, connectionTimeout: 60000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', e => console.error('IMAP err (handled):', e.message));

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');
    let newDeals = 0;

    try {
      const since = new Date(Date.now() - 72 * 60 * 60 * 1000);

      // Phase 1: headers only — no body download, zero cost
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

      // Phase 2: fetch full body only for deal emails
      for (const c of candidates) {
        try {
          const msg = await client.fetchOne(c.uid, { source: true }, { uid: true });
          const parsed = await simpleParser(msg.source);
          const body = parsed.text || (parsed.html || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

          console.log(`📬 ${c.subj}`);
          const properties = await extract(c.from, c.subj, body);

          if (!properties || properties.length === 0) {
            console.log('  No properties extracted');
          } else {
            console.log(`  ${properties.length} propert${properties.length > 1 ? 'ies' : 'y'}`);
            for (const p of properties) {
              if (!p.address) { console.log('  Skipping (no address)'); continue; }
              await appendRow(buildRow(p, c.subj, c.uid));
              console.log(`  ✅ ${p.address}, ${p.city} — $${p.asking_price} / ARV $${p.arv}`);
              newDeals++;
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

      console.log(`✅ Done — ${newDeals} new deal${newDeals !== 1 ? 's' : ''} written to sheet`);
    } finally { lock.release(); }
  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    try { await client.logout(); } catch {}
  }
}

const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | polling every ${POLL_MS/60000}min`);

// Clear sheet on startup so old misaligned data is gone
ensureHeaders().then(() => {
  poll();
  setInterval(poll, POLL_MS);
});
