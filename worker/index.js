require('dotenv').config({ path: '../.env' });
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');
const fs = require('fs');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SEEN_FILE = '/tmp/seen.json';

// Never crash
process.on('uncaughtException', e => console.error('ERR:', e.message));
process.on('unhandledRejection', e => console.error('REJ:', e?.message || e));

// Track seen emails across restarts
let seen = new Set();
try { seen = new Set(JSON.parse(fs.readFileSync(SEEN_FILE))); } catch {}
const saveSeen = () => { try { fs.writeFileSync(SEEN_FILE, JSON.stringify([...seen])); } catch {} };

// Google Sheets client
function sheets() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.JWT(creds.client_email, null, creds.private_key, ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

// Write a row to the sheet
async function writeRow(row) {
  const s = sheets();
  // Ensure headers exist
  const check = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Active Deals!A1' }).catch(() => null);
  if (!check?.data?.values) {
    await s.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'Active Deals!A1', valueInputOption: 'RAW',
      requestBody: { values: [['Date','Address','City','State','Zip','Beds','Baths','Sqft','Asking Price','ARV','Repairs','Close Date','Contact Name','Contact Phone','Contact Email','Contact Company','All Phones','All Emails','Pool','Garage','Lot Sqft','Year Built','Roof Age','HVAC Age','Seller Situation','HOA','Flood Zone','Drive Link','All Links','Notes','Subject','UID']] }
    });
  }
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'Active Deals!A:A', valueInputOption: 'RAW',
    requestBody: { values: [row] }
  });
}

// Deal keywords — no API call needed
const KEYWORDS = ['off market','wholesale','arv','flip','motivated','for sale','opportunity','equity','distressed','foreclosure','probate','deal','property available','available deals','fix flip','cash buyer','assignment','sales price','asking price','beds','baths','sqft','under roof'];
const isDeal = (subject, from) => KEYWORDS.some(k => `${subject} ${from}`.toLowerCase().includes(k));

// Extract all properties from one email body — ONE Claude call
async function extract(from, subject, body) {
  const res = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001', // cheapest model
    max_tokens: 2000,
    messages: [{ role: 'user', content: `You are a real estate data extractor. Extract ALL properties from this wholesale deal email. Return a JSON array (one object per property). Fields for each: address, city, state, zip, beds, baths, sqft, asking_price, arv, repairs, close_date, contact_name, contact_phone, contact_email, contact_company, all_phones, all_emails, pool, garage, lot_sqft, year_built, roof_age, hvac_age, seller_situation, hoa, flood_zone, drive_link, all_links, notes. Use null for missing. Return ONLY the JSON array.\n\nFROM: ${from}\nSUBJECT: ${subject}\nBODY:\n${body.slice(0, 8000)}` }]
  });
  try {
    const text = res.content[0].text.trim().replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch { return null; }
}

const v = x => (x === null || x === undefined) ? '' : String(x);

async function poll() {
  console.log(`[${new Date().toLocaleTimeString()}] Polling...`);
  const client = new ImapFlow({
    host: process.env.IMAP_HOST, port: 993, secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false, socketTimeout: 120000, connectionTimeout: 60000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', e => console.error('IMAP err:', e.message));

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');
    let newDeals = 0;

    try {
      const since = new Date(Date.now() - 72 * 60 * 60 * 1000);

      // Phase 1: get envelopes only (cheap)
      const candidates = [];
      for await (const msg of client.fetch({ since }, { envelope: true, uid: true })) {
        const uid = String(msg.uid);
        if (seen.has(uid)) continue;
        const subj = msg.envelope?.subject || '';
        const from = msg.envelope?.from?.[0]?.address || '';
        if (isDeal(subj, from)) candidates.push({ uid, subj, from });
        else { seen.add(uid); } // mark non-deals as seen immediately
      }

      saveSeen();
      console.log(`${candidates.length} deal emails to process`);

      // Phase 2: fetch body only for deal emails
      for (const c of candidates) {
        try {
          const msg = await client.fetchOne(c.uid, { source: true }, { uid: true });
          const parsed = await simpleParser(msg.source);
          const body = parsed.text || (parsed.html || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

          console.log(`📬 ${c.subj}`);
          const properties = await extract(c.from, c.subj, body);

          if (!properties || properties.length === 0) {
            console.log('  No properties found');
          } else {
            console.log(`  ${properties.length} properties`);
            for (const p of properties) {
              if (!p.address) continue;
              const row = [new Date().toISOString(), v(p.address), v(p.city), v(p.state), v(p.zip), v(p.beds), v(p.baths), v(p.sqft), v(p.asking_price), v(p.arv), v(p.repairs), v(p.close_date), v(p.contact_name), v(p.contact_phone), v(p.contact_email), v(p.contact_company), v(p.all_phones), v(p.all_emails), v(p.pool), v(p.garage), v(p.lot_sqft), v(p.year_built), v(p.roof_age), v(p.hvac_age), v(p.seller_situation), v(p.hoa), v(p.flood_zone), v(p.drive_link), v(p.all_links), v(p.notes), c.subj, c.uid];
              await writeRow(row);
              console.log(`  ✅ ${p.address}, ${p.city}`);
              newDeals++;
            }
          }

          seen.add(c.uid);
          saveSeen();
        } catch (e) {
          console.error(`  Error on ${c.uid}:`, e.message);
          seen.add(c.uid); // don't retry broken emails
          saveSeen();
        }
      }

      console.log(`Done. ${newDeals} new deals written.`);
    } finally { lock.release(); }
  } catch (e) {
    console.error('Poll error:', e.message);
  } finally {
    try { await client.logout(); } catch {}
  }
}

const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log(`🤙 Derek | ${process.env.IMAP_USER} | every ${POLL_MS/60000}min`);
poll();
setInterval(poll, POLL_MS);
