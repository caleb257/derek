require('dotenv').config({ path: '../.env' });
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { getSheetsClient } = require('./sheets');
const { loadBrain, logLesson } = require('./brain');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Never crash
process.on('uncaughtException', (err) => console.error('💥 Uncaught:', err.message));
process.on('unhandledRejection', (reason) => console.error('💥 Rejected:', reason?.message || reason));

function s(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'YES' : 'NO';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

function normalizeAddress(address, city, zip) {
  if (!address) return null;
  return `${address} ${city || ''} ${zip || ''}`.toLowerCase().replace(/\s+/g, ' ').replace(/[^\w\s]/g, '').trim();
}

async function isDealEmail(subject, from, bodyPreview) {
  const combined = `${subject} ${from} ${bodyPreview}`.toLowerCase();
  const keywords = ['off market','wholesale','arv','asking','flip','motivated','investment','sqft','beds','baths','price drop','for sale','opportunity','equity','distressed','vacant','foreclosure','probate','inherited','cash buyer','assignment','contract','closing','sales price'];
  if (keywords.some(kw => combined.includes(kw))) return true;
  const res = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514', max_tokens: 50,
    messages: [{ role: 'user', content: `Is this a real estate deal email? YES or NO.\nFrom: ${from}\nSubject: ${subject}\nPreview: ${bodyPreview.substring(0, 400)}` }]
  });
  return res.content[0].text.trim().toUpperCase().startsWith('YES');
}

async function extractDeals(from, subject, textBody, htmlBody) {
  const body = textBody || (htmlBody || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
  const brainContext = await loadBrain();
  const brainNote = brainContext ? `\nKNOWN WHOLESALER FORMATS:\n${brainContext}` : '';

  const res = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514', max_tokens: 8000,
    system: `You are Derek the Dealer, acquisitions analyst for Coralstone Capital Group, Tampa Bay FL. Extract ALL properties from wholesale deal emails. Return a JSON ARRAY — one object per property. If one email has 7 properties, return 7 objects. Return ONLY valid JSON array, no markdown.${brainNote}`,
    messages: [{ role: 'user', content: `Extract every property from this email as a JSON array. Each object must have these fields (null if not found):\naddress, city, state, zip, county, beds, baths, half_baths, sqft_living, sqft_total, lot_size_sqft, lot_size_acres, year_built, stories, garage, garage_spaces, pool, construction_type, roof_type, roof_age, hvac_type, hvac_age, asking_price, arv_stated, repair_estimate_stated, repair_scope_notes, assignment_fee, equity_stated, close_timeline, close_date_target, offer_deadline, earnest_money, financing_terms, cash_only, occupancy_status, contact_1_name, contact_1_company, contact_1_email, contact_1_phone, contact_1_phone_2, contact_1_website, contact_2_name, contact_2_company, contact_2_email, contact_2_phone, contact_3_name, contact_3_phone, all_phones_found, all_emails_found, all_names_found, seller_name, seller_phone, seller_situation, seller_motivation, hoa, hoa_fee, flood_zone, school_district, taxes_annual, zoning, parcel_id, mls_number, property_link_main, drive_link, zillow_link, google_maps_link, all_links_found, photos_included, photo_count, all_photo_links, comp_1_address, comp_1_price, comp_2_address, comp_2_price, comp_3_address, comp_3_price, marketing_headline, what_needs_work, what_is_updated, red_flags, highlights, additional_notes, wholesaler_email_format, what_worked_in_parsing, watch_out_for, raw_asking_price_number, raw_arv_number\n\nFROM: ${from}\nSUBJECT: ${subject}\nBODY:\n${body}` }]
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

function getHeaders() {
  return ['Date Received','Expires','Address','City','State','Zip','County','Beds','Baths','Half Baths','Sqft Living','Sqft Total','Lot Sqft','Lot Acres','Year Built','Stories','Garage','Garage Spaces','Pool','Construction','Roof Type','Roof Age','HVAC Type','HVAC Age','Asking Price','ARV Stated','Repairs Stated','Repair Scope','Assignment Fee','Equity','Close Timeline','Close Target','Offer Deadline','Earnest Money','Financing','Cash Only','Occupancy','Contact 1 Name','Contact 1 Company','Contact 1 Email','Contact 1 Phone','Contact 1 Phone 2','Contact 1 Website','Contact 2 Name','Contact 2 Company','Contact 2 Email','Contact 2 Phone','Contact 3 Name','Contact 3 Phone','All Phones','All Emails','All Names','Seller Name','Seller Phone','Seller Situation','Seller Motivation','HOA','HOA Fee','Flood Zone','School District','Annual Taxes','Zoning','Parcel ID','MLS #','Main Link','Drive Link','Zillow','Google Maps','All Links','Photos','Photo Count','All Photo Links','Comp 1 Address','Comp 1 Price','Comp 2 Address','Comp 2 Price','Comp 3 Address','Comp 3 Price','Headline','What Needs Work','What Updated','Red Flags','Highlights','Notes','Email Subject','Email ID'];
}

function buildRow(deal, subject, uid) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7*24*60*60*1000);
  const d = k => s(deal[k]);
  return [now.toISOString(), expires.toISOString(), d('address'), d('city'), d('state'), d('zip'), d('county'), d('beds'), d('baths'), d('half_baths'), d('sqft_living'), d('sqft_total'), d('lot_size_sqft'), d('lot_size_acres'), d('year_built'), d('stories'), d('garage'), d('garage_spaces'), d('pool'), d('construction_type'), d('roof_type'), d('roof_age'), d('hvac_type'), d('hvac_age'), d('asking_price'), d('arv_stated'), d('repair_estimate_stated'), d('repair_scope_notes'), d('assignment_fee'), d('equity_stated'), d('close_timeline'), d('close_date_target'), d('offer_deadline'), d('earnest_money'), d('financing_terms'), d('cash_only'), d('occupancy_status'), d('contact_1_name'), d('contact_1_company'), d('contact_1_email'), d('contact_1_phone'), d('contact_1_phone_2'), d('contact_1_website'), d('contact_2_name'), d('contact_2_company'), d('contact_2_email'), d('contact_2_phone'), d('contact_3_name'), d('contact_3_phone'), d('all_phones_found'), d('all_emails_found'), d('all_names_found'), d('seller_name'), d('seller_phone'), d('seller_situation'), d('seller_motivation'), d('hoa'), d('hoa_fee'), d('flood_zone'), d('school_district'), d('taxes_annual'), d('zoning'), d('parcel_id'), d('mls_number'), d('property_link_main'), d('drive_link'), d('zillow_link'), d('google_maps_link'), d('all_links_found'), d('photos_included'), d('photo_count'), d('all_photo_links'), d('comp_1_address'), d('comp_1_price'), d('comp_2_address'), d('comp_2_price'), d('comp_3_address'), d('comp_3_price'), d('marketing_headline'), d('what_needs_work'), d('what_is_updated'), d('red_flags'), d('highlights'), d('additional_notes'), s(subject), s(uid)];
}

async function ensureHeaders(sheets) {
  const check = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A1:A1' }).catch(() => ({ data: {} }));
  if (!check.data.values) {
    const h = getHeaders();
    for (const tab of ['Active Deals', 'Deal Storage', 'Price Changes']) {
      await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A1`, valueInputOption: 'RAW', requestBody: { values: [h] } }).catch(() => {});
    }
    console.log('✅ Headers written');
  }
}

async function checkDuplicate(sheets, deal) {
  const nn = normalizeAddress(deal.address, deal.city, deal.zip);
  if (!nn) return { isDuplicate: false };
  for (const tab of ['Active Deals', 'Deal Storage']) {
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A:E` }).catch(() => ({ data: { values: [] } }));
    for (const row of (r.data.values || []).slice(1)) {
      if (nn === normalizeAddress(row[2], row[3], row[5])) {
        const np = deal.raw_asking_price_number;
        const op = parseFloat((row[25] || '').toString().replace(/[^0-9.]/g, ''));
        if (np && op && Math.abs(np - op) > 500) return { isDuplicate: true, isPriceChange: true, oldPrice: row[25], newPrice: deal.asking_price };
        return { isDuplicate: true, isPriceChange: false };
      }
    }
  }
  return { isDuplicate: false };
}

const processedIds = new Set();

async function pollInbox() {
  console.log(`\n🔍 [${new Date().toLocaleTimeString()}] Polling...`);

  let client;
  try {
    client = new ImapFlow({
      host: process.env.IMAP_HOST, port: parseInt(process.env.IMAP_PORT || '993'), secure: true,
      auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
      logger: false, socketTimeout: 60000, connectionTimeout: 60000, greetingTimeout: 30000,
      tls: { rejectUnauthorized: false }
    });
    client.on('error', err => console.error('📡 IMAP error (caught):', err.message));
  } catch (e) { console.error('IMAP init error:', e.message); return; }

  let newDeals = 0, dupes = 0, skipped = 0;

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');

    try {
      const since = new Date(Date.now() - 48*60*60*1000);
      for await (const msg of client.fetch({ since }, { source: true, uid: true })) {
        const uid = msg.uid.toString();
        if (processedIds.has(uid)) continue;

        let parsed;
        try { parsed = await simpleParser(msg.source); } catch { processedIds.add(uid); continue; }

        const subject = parsed.subject || '';
        const from = parsed.from?.text || '';
        const text = parsed.text || '';
        const html = parsed.html || '';
        const preview = text || html.replace(/<[^>]*>/g, ' ');

        const isDeal = await isDealEmail(subject, from, preview);
        if (!isDeal) { skipped++; processedIds.add(uid); continue; }

        console.log(`📬 ${subject}`);

        const deals = await extractDeals(from, subject, text, html);
        if (!deals || deals.length === 0) { console.warn('⚠️ No deals extracted'); processedIds.add(uid); continue; }
        console.log(`📦 ${deals.length} propert${deals.length > 1 ? 'ies' : 'y'}`);

        const sheets = getSheetsClient();
        await ensureHeaders(sheets);

        for (const deal of deals) {
          if (!deal?.address) { console.warn('⚠️ No address, skipping'); continue; }
          const dup = await checkDuplicate(sheets, deal);

          if (dup.isDuplicate && !dup.isPriceChange) {
            console.log(`🔁 Dupe: ${deal.address}`); dupes++;
          } else if (dup.isPriceChange) {
            const row = buildRow(deal, subject, uid);
            await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Price Changes!A:A', valueInputOption: 'RAW', requestBody: { values: [row] } });
            console.log(`💰 Price change: ${deal.address}`);
          } else {
            const row = buildRow(deal, subject, uid);
            await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:A', valueInputOption: 'RAW', requestBody: { values: [row] } });
            console.log(`✅ ${deal.address}, ${deal.city}`); newDeals++;
          }
          await new Promise(r => setTimeout(r, 500));
        }

        await logLesson(deals[0]?.contact_1_email || from, deals[0]?.contact_1_company || '', deals[0]?.wholesaler_email_format || '', deals[0]?.what_worked_in_parsing || '', deals[0]?.watch_out_for || '', deals.length);
        processedIds.add(uid);
        await new Promise(r => setTimeout(r, 1000));
      }
    } finally { lock.release(); }

    // Archive expired
    const sheets = getSheetsClient();
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:Z' }).catch(() => ({ data: { values: [] } }));
    const rows = r.data.values || [];
    if (rows.length > 1) {
      const now = new Date(), active = [rows[0]], expired = [];
      for (let i = 1; i < rows.length; i++) { new Date(rows[i][1]) < now ? expired.push(rows[i]) : active.push(rows[i]); }
      if (expired.length > 0) {
        await sheets.spreadsheets.values.clear({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:Z' });
        await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A1', valueInputOption: 'RAW', requestBody: { values: active } });
        await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Deal Storage!A:A', valueInputOption: 'RAW', requestBody: { values: expired } });
        console.log(`📦 Archived ${expired.length}`);
      }
    }
  } catch (err) {
    console.error('❌ Poll error:', err.message);
  } finally {
    try { await client.logout(); } catch {}
  }

  console.log(`📊 ${newDeals} new | ${dupes} dupes | ${skipped} skipped`);
}

const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log('🤙 Derek — Coralstone Capital Group');
console.log(`📬 ${process.env.IMAP_USER} | every ${POLL_MS/60000}min\n`);
pollInbox();
setInterval(pollInbox, POLL_MS);
