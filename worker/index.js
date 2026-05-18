require('dotenv').config({ path: '../.env' });

const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');
const { loadBrain, logLesson } = require('./brain');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

function getSheetsClient() {
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

async function isDealEmail(subject, from, bodyPreview) {
  const res = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 50,
    messages: [{ role: 'user', content: `Is this a wholesale real estate deal email being offered for sale? YES or NO only.\nFrom: ${from}\nSubject: ${subject}\nPreview: ${bodyPreview.substring(0, 300)}` }]
  });
  return res.content[0].text.trim().toUpperCase().startsWith('YES');
}

async function extractDealData(from, subject, textBody, htmlBody, attachmentNames) {
  const body = textBody || htmlBody?.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ') || '';

  // Load Derek's brain — what he already knows about this wholesaler and others
  const brainContext = await loadBrain();
  const brainSection = brainContext ? `\n\nWHAT YOU ALREADY KNOW ABOUT WHOLESALER FORMATS:\n${brainContext}\n\nUse this knowledge to parse more accurately.` : '';

  const res = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4000,
    system: `You are Derek the Dealer, an expert real estate acquisitions analyst for Coralstone Capital Group in Tampa Bay, FL. You read wholesale deal emails from ANY format and extract EVERY piece of data. You get smarter with every email you process. Return ONLY valid JSON, no markdown.${brainSection}`,
    messages: [{
      role: 'user',
      content: `Extract ALL data from this wholesale deal email. Be thorough.\n\nFROM: ${from}\nSUBJECT: ${subject}\nATTACHMENTS: ${attachmentNames.join(', ') || 'none'}\n\nFULL BODY:\n${body}\n\nReturn JSON:\n{"address":null,"address2":null,"city":null,"state":null,"zip":null,"county":null,"neighborhood":null,"subdivision":null,"property_type":null,"beds":null,"baths":null,"half_baths":null,"sqft":null,"lot_size":null,"year_built":null,"stories":null,"garage":null,"pool":null,"basement":null,"construction_type":null,"roof_type":null,"hvac":null,"asking_price":null,"arv_stated":null,"repair_estimate_stated":null,"repair_scope_notes":null,"assignment_fee":null,"net_to_seller":null,"price_per_sqft":null,"arv_per_sqft":null,"equity_stated":null,"wholesaler_name":null,"wholesaler_company":null,"wholesaler_email":null,"wholesaler_phone":null,"wholesaler_website":null,"second_contact_name":null,"second_contact_phone":null,"second_contact_email":null,"seller_name":null,"seller_situation":null,"seller_motivation":null,"occupancy_status":null,"tenant_info":null,"days_on_market":null,"list_date":null,"close_timeline":null,"inspection_period":null,"earnest_money":null,"financing_terms":null,"title_company":null,"hoa":null,"hoa_fee":null,"flood_zone":null,"flood_zone_code":null,"school_district":null,"taxes_annual":null,"insurance_estimate":null,"utility_notes":null,"zoning":null,"parcel_id":null,"mls_number":null,"property_links":null,"zillow_link":null,"redfin_link":null,"mls_link":null,"propstream_link":null,"google_maps_link":null,"photos_included":false,"photo_count":null,"photo_links":null,"virtual_tour_link":null,"comps_provided":null,"comp_1":null,"comp_2":null,"comp_3":null,"deal_source":null,"list_name":null,"marketing_headline":null,"additional_notes":null,"red_flags":null,"wholesaler_email_format":null,"what_worked_in_parsing":null,"watch_out_for":null,"fields_extracted_count":0,"raw_asking_price_number":null,"raw_arv_number":null,"raw_repair_number":null,"raw_assignment_fee_number":null}`
    }]
  });

  try {
    const text = res.content[0].text.trim().replace(/```json|```/g, '').trim();
    return JSON.parse(text);
  } catch (e) {
    console.error('Parse error:', e.message);
    return null;
  }
}

function normalizeAddress(address, city, zip) {
  if (!address) return null;
  return `${address} ${city || ''} ${zip || ''}`.toLowerCase().replace(/\s+/g, ' ').replace(/[^\w\s]/g, '').trim();
}

function getHeaders() {
  return ['Date Received','Expires','Address','Address2','City','State','Zip','County','Neighborhood','Subdivision','Property Type','Beds','Baths','Half Baths','Sqft','Lot Size','Year Built','Stories','Garage','Pool','Basement','Construction Type','Roof Type','HVAC','Asking Price','ARV (Stated)','Repairs (Stated)','Repair Scope Notes','Assignment Fee','Net to Seller','$/Sqft','ARV/Sqft','Equity (Stated)','Wholesaler Name','Wholesaler Company','Wholesaler Email','Wholesaler Phone','Wholesaler Website','Contact 2 Name','Contact 2 Phone','Contact 2 Email','Seller Name','Seller Situation','Seller Motivation','Occupancy','Tenant Info','Days on Market','List Date','Close Timeline','Inspection Period','Earnest Money','Financing Terms','Title Company','HOA','HOA Fee','Flood Zone','Flood Zone Code','School District','Annual Taxes','Insurance Estimate','Utility Notes','Zoning','Parcel ID','MLS #','Property Links','Zillow','Redfin','MLS Link','PropStream','Google Maps','Photos Included','Photo Count','Photo Links','Virtual Tour','Comps Provided','Comp 1','Comp 2','Comp 3','Deal Source','List Name','Marketing Headline','Additional Notes','Red Flags','Wholesaler Format','Email Subject','Email ID'];
}

function buildRow(deal, subject, messageId) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  return [now.toISOString(),expires.toISOString(),deal.address,deal.address2,deal.city,deal.state,deal.zip,deal.county,deal.neighborhood,deal.subdivision,deal.property_type,deal.beds,deal.baths,deal.half_baths,deal.sqft,deal.lot_size,deal.year_built,deal.stories,deal.garage,deal.pool,deal.basement,deal.construction_type,deal.roof_type,deal.hvac,deal.asking_price,deal.arv_stated,deal.repair_estimate_stated,deal.repair_scope_notes,deal.assignment_fee,deal.net_to_seller,deal.price_per_sqft,deal.arv_per_sqft,deal.equity_stated,deal.wholesaler_name,deal.wholesaler_company,deal.wholesaler_email,deal.wholesaler_phone,deal.wholesaler_website,deal.second_contact_name,deal.second_contact_phone,deal.second_contact_email,deal.seller_name,deal.seller_situation,deal.seller_motivation,deal.occupancy_status,deal.tenant_info,deal.days_on_market,deal.list_date,deal.close_timeline,deal.inspection_period,deal.earnest_money,deal.financing_terms,deal.title_company,deal.hoa,deal.hoa_fee,deal.flood_zone,deal.flood_zone_code,deal.school_district,deal.taxes_annual,deal.insurance_estimate,deal.utility_notes,deal.zoning,deal.parcel_id,deal.mls_number,deal.property_links,deal.zillow_link,deal.redfin_link,deal.mls_link,deal.propstream_link,deal.google_maps_link,deal.photos_included?'YES':'NO',deal.photo_count,deal.photo_links,deal.virtual_tour_link,deal.comps_provided,deal.comp_1,deal.comp_2,deal.comp_3,deal.deal_source,deal.list_name,deal.marketing_headline,deal.additional_notes,deal.red_flags,deal.wholesaler_email_format,subject,messageId];
}

async function ensureSheetSetup(sheets) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID });
  const existing = spreadsheet.data.sheets.map(s => s.properties.title);
  const requests = [];
  if (!existing.includes('Active Deals')) requests.push({ addSheet: { properties: { title: 'Active Deals', index: 0 } } });
  if (!existing.includes('Deal Storage')) requests.push({ addSheet: { properties: { title: 'Deal Storage', index: 1 } } });
  if (!existing.includes('Price Changes')) requests.push({ addSheet: { properties: { title: 'Price Changes', index: 2 } } });
  if (!existing.includes("Derek's Brain")) requests.push({ addSheet: { properties: { title: "Derek's Brain", index: 3 } } });
  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: process.env.GOOGLE_SHEET_ID, requestBody: { requests } });
  }
  const check = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A1:A1' }).catch(() => ({ data: {} }));
  if (!check.data.values) {
    const headers = getHeaders();
    for (const tab of ['Active Deals', 'Deal Storage', 'Price Changes']) {
      await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A1`, valueInputOption: 'RAW', requestBody: { values: [headers] } });
    }
    await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: "Derek's Brain!A1", valueInputOption: 'RAW', requestBody: { values: [['Wholesaler Email','Company','Format Type','What Worked','Watch Out For','Fields Extracted','Last Seen']] } });
    console.log('✅ All tabs and headers ready');
  }
}

async function checkDuplicate(sheets, deal) {
  const normalizedNew = normalizeAddress(deal.address, deal.city, deal.zip);
  if (!normalizedNew) return { isDuplicate: false };
  for (const tab of ['Active Deals', 'Deal Storage']) {
    const result = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A:AZ` }).catch(() => ({ data: { values: [] } }));
    const rows = result.data.values || [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (normalizedNew === normalizeAddress(row[2], row[4], row[6])) {
        const newPrice = deal.raw_asking_price_number;
        const oldPriceNum = parseFloat((row[24] || '').toString().replace(/[^0-9.]/g, ''));
        if (newPrice && oldPriceNum && Math.abs(newPrice - oldPriceNum) > 500) {
          return { isDuplicate: true, isPriceChange: true, oldPrice: row[24], newPrice: deal.asking_price };
        }
        return { isDuplicate: true, isPriceChange: false };
      }
    }
  }
  return { isDuplicate: false };
}

const processedIds = new Set();

async function pollInbox() {
  console.log(`\n🔍 [${new Date().toLocaleTimeString()}] Polling...`);
  const client = new ImapFlow({
    host: process.env.IMAP_HOST,
    port: parseInt(process.env.IMAP_PORT || '993'),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASSWORD },
    logger: false
  });

  let newDeals = 0, dupes = 0, priceChanges = 0, skipped = 0;

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');
    try {
      const since = new Date(Date.now() - 48 * 60 * 60 * 1000);
      for await (const msg of client.fetch({ since }, { source: true, uid: true })) {
        const uid = msg.uid.toString();
        if (processedIds.has(uid)) continue;

        let parsed;
        try { parsed = await simpleParser(msg.source); } catch { processedIds.add(uid); continue; }

        const subject = parsed.subject || '';
        const from = parsed.from?.text || '';
        const textBody = parsed.text || '';
        const htmlBody = parsed.html || '';
        const bodyPreview = textBody || htmlBody?.replace(/<[^>]*>/g, ' ') || '';
        const attachmentNames = (parsed.attachments || []).map(a => a.filename).filter(Boolean);

        const isDeal = await isDealEmail(subject, from, bodyPreview);
        if (!isDeal) { skipped++; processedIds.add(uid); continue; }

        console.log(`📬 ${subject}`);

        const deal = await extractDealData(from, subject, textBody, htmlBody, attachmentNames);
        if (!deal?.address) { console.warn(`⚠️ No address: ${subject}`); processedIds.add(uid); continue; }

        const sheets = getSheetsClient();
        await ensureSheetSetup(sheets);
        const dup = await checkDuplicate(sheets, deal);

        if (dup.isDuplicate && !dup.isPriceChange) {
          console.log(`🔁 Dupe: ${deal.address}`);
          dupes++;
        } else if (dup.isPriceChange) {
          const row = buildRow(deal, subject, uid);
          row[24] = `${dup.newPrice} (was ${dup.oldPrice})`;
          await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Price Changes!A:A', valueInputOption: 'RAW', requestBody: { values: [row] } });
          console.log(`💰 Price change: ${deal.address}`);
          priceChanges++;
        } else {
          await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:A', valueInputOption: 'RAW', requestBody: { values: [buildRow(deal, subject, uid)] } });
          console.log(`✅ New: ${deal.address}, ${deal.city}`);
          newDeals++;
        }

        // Update Derek's brain after every successful parse
        await logLesson(
          deal.wholesaler_email || from,
          deal.wholesaler_company || '',
          deal.wholesaler_email_format || 'unknown',
          deal.what_worked_in_parsing || '',
          deal.watch_out_for || '',
          deal.fields_extracted_count || 0
        );

        processedIds.add(uid);
        await new Promise(r => setTimeout(r, 1000));
      }
    } finally { lock.release(); }

    // Archive expired
    const sheets = getSheetsClient();
    const result = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:AZ' }).catch(() => ({ data: { values: [] } }));
    const rows = result.data.values || [];
    if (rows.length > 1) {
      const now = new Date();
      const active = [rows[0]], expired = [];
      for (let i = 1; i < rows.length; i++) {
        new Date(rows[i][1]) < now ? expired.push(rows[i]) : active.push(rows[i]);
      }
      if (expired.length > 0) {
        await sheets.spreadsheets.values.clear({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:AZ' });
        await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A1', valueInputOption: 'RAW', requestBody: { values: active } });
        await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Deal Storage!A:A', valueInputOption: 'RAW', requestBody: { values: expired } });
        console.log(`📦 Archived ${expired.length}`);
      }
    }
  } catch (err) {
    console.error('❌', err.message);
  } finally {
    try { await client.logout(); } catch {}
  }

  console.log(`📊 ${newDeals} new | ${dupes} dupes | ${priceChanges} price changes | ${skipped} skipped`);
}

const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log('🤙 Derek — Coralstone Capital Group');
console.log(`📬 ${process.env.IMAP_USER} | every ${POLL_MS / 60000}min\n`);

pollInbox();
setInterval(pollInbox, POLL_MS);
