require('dotenv').config({ path: '../.env' });
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { getSheetsClient } = require('./sheets');
const { loadBrain, logLesson } = require('./brain');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

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

function isDealByKeywords(subject, from) {
  const combined = `${subject} ${from}`.toLowerCase();
  const keywords = ['off market','wholesale','arv','flip','motivated','investment','for sale','opportunity','equity','distressed','vacant','foreclosure','probate','inherited','cash buyer','assignment','deal','property available','available deals','fix flip','fix and flip'];
  return keywords.some(kw => combined.includes(kw));
}

async function isDealEmail(subject, from, bodyPreview) {
  if (isDealByKeywords(subject, from)) return true;
  const res = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514', max_tokens: 50,
    messages: [{ role: 'user', content: `Is this a real estate deal email? YES or NO.\nFrom: ${from}\nSubject: ${subject}\nPreview: ${bodyPreview.substring(0, 300)}` }]
  });
  return res.content[0].text.trim().toUpperCase().startsWith('YES');
}

async function extractDeals(from, subject, body) {
  const brainContext = await loadBrain();
  const brainNote = brainContext ? `\nKNOWN WHOLESALER FORMATS:\n${brainContext}` : '';

  const res = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514', max_tokens: 8000,
    system: `You are Derek the Dealer, acquisitions analyst for Coralstone Capital Group, Tampa Bay FL. Extract ALL properties from wholesale deal emails. Return a JSON ARRAY — one object per property. Return ONLY valid JSON array, no markdown.${brainNote}`,
    messages: [{ role: 'user', content: `Extract every property as a JSON array. Each object must have (null if not found): address, city, state, zip, beds, baths, sqft_living, lot_size_sqft, year_built, pool, garage, construction_type, roof_type, roof_age, hvac_age, asking_price, arv_stated, repair_estimate_stated, close_date_target, contact_1_name, contact_1_company, contact_1_email, contact_1_phone, contact_2_name, contact_2_phone, all_phones_found, all_emails_found, seller_situation, hoa, flood_zone, taxes_annual, drive_link, all_links_found, photos_included, marketing_headline, what_needs_work, what_is_updated, red_flags, additional_notes, wholesaler_email_format, what_worked_in_parsing, watch_out_for, raw_asking_price_number, raw_arv_number\n\nFROM: ${from}\nSUBJECT: ${subject}\nBODY:\n${body.substring(0, 12000)}` }]
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
  return ['Date Received','Expires','Address','City','State','Zip','Beds','Baths','Sqft','Lot Sqft','Year Built','Pool','Garage','Construction','Roof Type','Roof Age','HVAC Age','Asking Price','ARV','Repairs','Close Date','Contact 1 Name','Contact 1 Company','Contact 1 Email','Contact 1 Phone','Contact 2 Name','Contact 2 Phone','All Phones','All Emails','Seller Situation','HOA','Flood Zone','Annual Taxes','Drive Link','All Links','Photos','Headline','What Needs Work','What Updated','Red Flags','Notes','Email Subject','Email ID'];
}

function buildRow(deal, subject, uid) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7*24*60*60*1000);
  const d = k => s(deal[k]);
  return [now.toISOString(), expires.toISOString(), d('address'), d('city'), d('state'), d('zip'), d('beds'), d('baths'), d('sqft_living'), d('lot_size_sqft'), d('year_built'), d('pool'), d('garage'), d('construction_type'), d('roof_type'), d('roof_age'), d('hvac_age'), d('asking_price'), d('arv_stated'), d('repair_estimate_stated'), d('close_date_target'), d('contact_1_name'), d('contact_1_company'), d('contact_1_email'), d('contact_1_phone'), d('contact_2_name'), d('contact_2_phone'), d('all_phones_found'), d('all_emails_found'), d('seller_situation'), d('hoa'), d('flood_zone'), d('taxes_annual'), d('drive_link'), d('all_links_found'), d('photos_included'), d('marketing_headline'), d('what_needs_work'), d('what_is_updated'), d('red_flags'), d('additional_notes'), s(subject), s(uid)];
}

async function ensureHeaders(sheets) {
  const check = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A1:A1' }).catch(() => ({ data: {} }));
  if (!check.data.values) {
    const h = getHeaders();
    for (const tab of ['Active Deals','Deal Storage','Price Changes']) {
      await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A1`, valueInputOption: 'RAW', requestBody: { values: [h] } }).catch(() => {});
    }
    console.log('✅ Headers written');
  }
}

async function checkDuplicate(sheets, deal) {
  const nn = normalizeAddress(deal.address, deal.city, deal.zip);
  if (!nn) return { isDuplicate: false };
  for (const tab of ['Active Deals','Deal Storage']) {
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A:F` }).catch(() => ({ data: { values: [] } }));
    for (const row of (r.data.values || []).slice(1)) {
      if (nn === normalizeAddress(row[2], row[3], row[5])) return { isDuplicate: true };
    }
  }
  return { isDuplicate: false };
}

function createImapClient() {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST,
    port: parseInt(process.env.IMAP_PORT || '993'),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false,
    socketTimeout: 120000,
    connectionTimeout: 60000,
    greetingTimeout: 30000,
    tls: { rejectUnauthorized: false }
  });
  client.on('error', err => console.error('📡 IMAP error (handled):', err.message));
  return client;
}

const processedIds = new Set();

async function pollInbox() {
  console.log(`\n🔍 [${new Date().toLocaleTimeString()}] Polling...`);
  let newDeals = 0, dupes = 0, skipped = 0;
  const client = createImapClient();

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');

    try {
      const since = new Date(Date.now() - 72*60*60*1000); // 72hr lookback

      // Step 1: Fetch headers only first (fast, small)
      const candidates = [];
      for await (const msg of client.fetch({ since }, { envelope: true, uid: true })) {
        const uid = msg.uid.toString();
        if (processedIds.has(uid)) continue;
        const subject = msg.envelope?.subject || '';
        const from = msg.envelope?.from?.[0]?.address || '';
        if (isDealByKeywords(subject, from)) {
          candidates.push({ uid, subject, from });
        } else {
          candidates.push({ uid, subject, from, needsCheck: true });
        }
      }

      console.log(`📋 Found ${candidates.length} emails to evaluate`);

      // Step 2: For keyword matches, fetch full body; for others, fetch text preview only
      for (const candidate of candidates) {
        if (processedIds.has(candidate.uid)) continue;

        let body = '';
        let isConfirmedDeal = !candidate.needsCheck;

        if (candidate.needsCheck) {
          // Fetch just text part for quick check
          try {
            const msgData = await client.fetchOne(candidate.uid, { bodyParts: ['TEXT'] }, { uid: true });
            const preview = msgData?.bodyParts?.get('TEXT')?.toString() || '';
            isConfirmedDeal = await isDealEmail(candidate.subject, candidate.from, preview);
            if (isConfirmedDeal) body = preview;
          } catch (e) {
            console.error(`Preview fetch error ${candidate.uid}:`, e.message);
            processedIds.add(candidate.uid);
            continue;
          }
        }

        if (!isConfirmedDeal) {
          skipped++;
          processedIds.add(candidate.uid);
          continue;
        }

        // Fetch full source for deal emails
        if (!body) {
          try {
            const msgData = await client.fetchOne(candidate.uid, { source: true }, { uid: true });
            const parsed = await simpleParser(msgData.source);
            body = parsed.text || (parsed.html || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
          } catch (e) {
            console.error(`Full fetch error ${candidate.uid}:`, e.message);
            processedIds.add(candidate.uid);
            continue;
          }
        }

        console.log(`📬 ${candidate.subject}`);

        const deals = await extractDeals(candidate.from, candidate.subject, body);
        if (!deals || deals.length === 0) {
          console.warn('⚠️ No deals extracted');
          processedIds.add(candidate.uid);
          continue;
        }
        console.log(`📦 ${deals.length} propert${deals.length > 1 ? 'ies' : 'y'}`);

        const sheets = getSheetsClient();
        await ensureHeaders(sheets);

        for (const deal of deals) {
          if (!deal?.address) { console.warn('⚠️ No address, skipping'); continue; }
          const dup = await checkDuplicate(sheets, deal);
          if (dup.isDuplicate) {
            console.log(`🔁 Dupe: ${deal.address}`); dupes++;
          } else {
            await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:A', valueInputOption: 'RAW', requestBody: { values: [buildRow(deal, candidate.subject, candidate.uid)] } });
            console.log(`✅ ${deal.address}, ${deal.city}`); newDeals++;
          }
          await new Promise(r => setTimeout(r, 500));
        }

        await logLesson(deals[0]?.contact_1_email || candidate.from, deals[0]?.contact_1_company || '', deals[0]?.wholesaler_email_format || '', deals[0]?.what_worked_in_parsing || '', deals[0]?.watch_out_for || '', deals.length);
        processedIds.add(candidate.uid);
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
