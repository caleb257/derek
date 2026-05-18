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
  const brainContext = await loadBrain();
  const brainSection = brainContext ? `\n\nWHAT YOU ALREADY KNOW ABOUT WHOLESALER FORMATS:\n${brainContext}\n\nUse this to parse more accurately.` : '';

  const res = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 8000,
    system: `You are Derek the Dealer, expert acquisitions analyst for Coralstone Capital Group in Tampa Bay FL. You read wholesale real estate emails in ANY format and extract LITERALLY EVERYTHING — every number, every name, every phone, every email, every link, every note, every detail no matter how small. Leave NOTHING behind. If a piece of information exists anywhere in the email, you find it and extract it. Return ONLY valid JSON, no markdown, no explanation.${brainSection}`,
    messages: [{
      role: 'user',
      content: `Extract EVERY SINGLE piece of information from this wholesale deal email. Be absolutely exhaustive — get everything including every contact, every phone number, every email address, every link, every number mentioned anywhere.

FROM: ${from}
SUBJECT: ${subject}
ATTACHMENTS: ${attachmentNames.join(', ') || 'none'}

FULL EMAIL BODY:
${body}

Return this complete JSON (null for truly missing fields — never guess critical numbers, but extract everything that exists):
{
  "address": null,
  "address2": null,
  "city": null,
  "state": null,
  "zip": null,
  "county": null,
  "neighborhood": null,
  "subdivision": null,
  "cross_streets": null,
  "google_maps_address": null,

  "property_type": null,
  "beds": null,
  "baths": null,
  "half_baths": null,
  "sqft_living": null,
  "sqft_total": null,
  "lot_size_sqft": null,
  "lot_size_acres": null,
  "year_built": null,
  "stories": null,
  "units": null,
  "garage": null,
  "garage_spaces": null,
  "carport": null,
  "pool": null,
  "basement": null,
  "basement_sqft": null,
  "attic": null,
  "construction_type": null,
  "foundation_type": null,
  "roof_type": null,
  "roof_age": null,
  "hvac_type": null,
  "hvac_age": null,
  "water_heater": null,
  "water_heater_age": null,
  "electrical_panel": null,
  "plumbing_type": null,
  "windows": null,
  "flooring": null,
  "kitchen_condition": null,
  "bath_condition": null,
  "exterior_condition": null,
  "interior_condition": null,
  "overall_condition": null,

  "asking_price": null,
  "arv_stated": null,
  "repair_estimate_stated": null,
  "repair_scope_notes": null,
  "assignment_fee": null,
  "net_to_seller": null,
  "price_per_sqft": null,
  "arv_per_sqft": null,
  "equity_stated": null,
  "equity_percent": null,
  "profit_potential_stated": null,
  "monthly_rent_current": null,
  "monthly_rent_market": null,
  "gross_rent_multiplier": null,
  "cap_rate_stated": null,
  "noi_stated": null,
  "annual_gross_income": null,
  "annual_expenses_stated": null,
  "cash_on_cash_stated": null,
  "price_reduction_from": null,
  "price_reduction_amount": null,

  "contact_1_name": null,
  "contact_1_title": null,
  "contact_1_company": null,
  "contact_1_email": null,
  "contact_1_phone": null,
  "contact_1_phone_2": null,
  "contact_1_website": null,
  "contact_1_social": null,

  "contact_2_name": null,
  "contact_2_title": null,
  "contact_2_company": null,
  "contact_2_email": null,
  "contact_2_phone": null,
  "contact_2_phone_2": null,
  "contact_2_website": null,
  "contact_2_social": null,

  "contact_3_name": null,
  "contact_3_title": null,
  "contact_3_company": null,
  "contact_3_email": null,
  "contact_3_phone": null,
  "contact_3_website": null,

  "all_emails_found": null,
  "all_phones_found": null,
  "all_names_found": null,
  "all_companies_found": null,

  "seller_name": null,
  "seller_phone": null,
  "seller_email": null,
  "seller_situation": null,
  "seller_motivation": null,
  "seller_timeline": null,
  "seller_asking_originally": null,
  "how_seller_was_found": null,

  "occupancy_status": null,
  "tenant_name": null,
  "tenant_phone": null,
  "tenant_email": null,
  "tenant_lease_end": null,
  "tenant_monthly_rent": null,
  "tenant_is_section_8": null,

  "days_on_market": null,
  "list_date": null,
  "offer_deadline": null,
  "close_timeline": null,
  "close_date_target": null,
  "inspection_period": null,
  "earnest_money": null,
  "earnest_money_type": null,
  "financing_terms": null,
  "cash_only": null,
  "title_company": null,
  "title_contact": null,
  "title_phone": null,
  "closing_agent": null,
  "contract_type": null,
  "assignment_allowed": null,
  "double_close_ok": null,

  "hoa": null,
  "hoa_name": null,
  "hoa_fee": null,
  "hoa_fee_frequency": null,
  "hoa_phone": null,
  "hoa_includes": null,
  "hoa_special_assessment": null,

  "flood_zone": null,
  "flood_zone_code": null,
  "flood_insurance_required": null,
  "flood_insurance_cost": null,
  "school_district": null,
  "elementary_school": null,
  "middle_school": null,
  "high_school": null,
  "taxes_annual": null,
  "taxes_monthly": null,
  "tax_year": null,
  "homestead_exempt": null,
  "insurance_annual": null,
  "insurance_monthly": null,
  "utility_water": null,
  "utility_sewer": null,
  "utility_electric": null,
  "utility_gas": null,
  "utility_trash": null,
  "utility_notes": null,
  "zoning": null,
  "zoning_description": null,
  "parcel_id": null,
  "legal_description": null,
  "mls_number": null,
  "mls_status": null,

  "property_link_main": null,
  "zillow_link": null,
  "redfin_link": null,
  "mls_link": null,
  "propstream_link": null,
  "google_maps_link": null,
  "virtual_tour_link": null,
  "video_link": null,
  "dropbox_link": null,
  "drive_link": null,
  "all_links_found": null,

  "photos_included": false,
  "photo_count": null,
  "photo_link_1": null,
  "photo_link_2": null,
  "photo_link_3": null,
  "all_photo_links": null,

  "comp_1_address": null,
  "comp_1_price": null,
  "comp_1_sqft": null,
  "comp_1_date": null,
  "comp_1_price_sqft": null,
  "comp_2_address": null,
  "comp_2_price": null,
  "comp_2_sqft": null,
  "comp_2_date": null,
  "comp_2_price_sqft": null,
  "comp_3_address": null,
  "comp_3_price": null,
  "comp_3_sqft": null,
  "comp_3_date": null,
  "comp_3_price_sqft": null,
  "comp_4_address": null,
  "comp_4_price": null,
  "comp_4_sqft": null,
  "comp_4_date": null,
  "comp_5_address": null,
  "comp_5_price": null,
  "comp_5_sqft": null,
  "comp_5_date": null,
  "arv_comp_notes": null,

  "deal_source": null,
  "list_name": null,
  "campaign_name": null,
  "marketing_headline": null,
  "wholesaler_deal_id": null,
  "wholesaler_notes": null,
  "repair_items_detailed": null,
  "what_needs_work": null,
  "what_is_updated": null,
  "permits_required": null,
  "code_violations": null,
  "liens_known": null,
  "title_issues_known": null,
  "probate": null,
  "bankruptcy": null,
  "divorce": null,
  "foreclosure_status": null,
  "auction_info": null,
  "additional_notes": null,
  "red_flags": null,
  "highlights": null,

  "wholesaler_email_format": null,
  "what_worked_in_parsing": null,
  "watch_out_for": null,
  "fields_extracted_count": 0,
  "raw_asking_price_number": null,
  "raw_arv_number": null,
  "raw_repair_number": null,
  "raw_assignment_fee_number": null
}`
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
  return [
    'Date Received','Expires',
    'Address','Address2','City','State','Zip','County','Neighborhood','Subdivision','Cross Streets','Maps Address',
    'Property Type','Beds','Baths','Half Baths','Sqft Living','Sqft Total','Lot Sqft','Lot Acres','Year Built','Stories','Units',
    'Garage','Garage Spaces','Carport','Pool','Basement','Basement Sqft','Attic',
    'Construction','Foundation','Roof Type','Roof Age','HVAC Type','HVAC Age','Water Heater','WH Age','Electrical','Plumbing','Windows','Flooring',
    'Kitchen Condition','Bath Condition','Exterior Condition','Interior Condition','Overall Condition',
    'Asking Price','ARV Stated','Repairs Stated','Repair Scope','Assignment Fee','Net to Seller',
    '$/Sqft','ARV/Sqft','Equity','Equity %','Profit Potential','Rent Current','Rent Market','GRM','Cap Rate','NOI','Annual Income','Annual Expenses','CoC','Price Was','Price Reduction',
    'Contact 1 Name','Contact 1 Title','Contact 1 Company','Contact 1 Email','Contact 1 Phone','Contact 1 Phone 2','Contact 1 Website','Contact 1 Social',
    'Contact 2 Name','Contact 2 Title','Contact 2 Company','Contact 2 Email','Contact 2 Phone','Contact 2 Phone 2','Contact 2 Website','Contact 2 Social',
    'Contact 3 Name','Contact 3 Title','Contact 3 Company','Contact 3 Email','Contact 3 Phone','Contact 3 Website',
    'All Emails','All Phones','All Names','All Companies',
    'Seller Name','Seller Phone','Seller Email','Seller Situation','Seller Motivation','Seller Timeline','Seller Original Ask','How Found',
    'Occupancy','Tenant Name','Tenant Phone','Tenant Email','Lease End','Tenant Rent','Section 8',
    'Days on Market','List Date','Offer Deadline','Close Timeline','Close Target','Inspection Period','Earnest Money','EMD Type','Financing','Cash Only',
    'Title Company','Title Contact','Title Phone','Closing Agent','Contract Type','Assignment OK','Double Close OK',
    'HOA','HOA Name','HOA Fee','HOA Frequency','HOA Phone','HOA Includes','Special Assessment',
    'Flood Zone','Flood Code','Flood Insurance Required','Flood Insurance Cost',
    'School District','Elementary','Middle','High School',
    'Annual Taxes','Monthly Taxes','Tax Year','Homestead',
    'Annual Insurance','Monthly Insurance',
    'Water','Sewer','Electric','Gas','Trash','Utility Notes',
    'Zoning','Zoning Description','Parcel ID','Legal Description','MLS #','MLS Status',
    'Main Link','Zillow','Redfin','MLS Link','PropStream','Google Maps','Virtual Tour','Video','Dropbox','Drive','All Links',
    'Photos','Photo Count','Photo 1','Photo 2','Photo 3','All Photo Links',
    'Comp 1 Address','Comp 1 Price','Comp 1 Sqft','Comp 1 Date','Comp 1 $/Sqft',
    'Comp 2 Address','Comp 2 Price','Comp 2 Sqft','Comp 2 Date','Comp 2 $/Sqft',
    'Comp 3 Address','Comp 3 Price','Comp 3 Sqft','Comp 3 Date','Comp 3 $/Sqft',
    'Comp 4 Address','Comp 4 Price','Comp 4 Sqft','Comp 4 Date','Comp 4 $/Sqft',
    'Comp 5 Address','Comp 5 Price','Comp 5 Sqft','Comp 5 Date','Comp 5 $/Sqft','Comp Notes',
    'Deal Source','List Name','Campaign','Headline','Wholesaler Deal ID','Wholesaler Notes',
    'Repair Items Detail','What Needs Work','What Is Updated','Permits Required',
    'Code Violations','Liens','Title Issues','Probate','Bankruptcy','Divorce','Foreclosure','Auction Info',
    'Additional Notes','Red Flags','Highlights',
    'Email Subject','Email ID'
  ];
}

function buildRow(deal, subject, messageId) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  return [
    now.toISOString(), expires.toISOString(),
    deal.address, deal.address2, deal.city, deal.state, deal.zip, deal.county, deal.neighborhood, deal.subdivision, deal.cross_streets, deal.google_maps_address,
    deal.property_type, deal.beds, deal.baths, deal.half_baths, deal.sqft_living, deal.sqft_total, deal.lot_size_sqft, deal.lot_size_acres, deal.year_built, deal.stories, deal.units,
    deal.garage, deal.garage_spaces, deal.carport, deal.pool, deal.basement, deal.basement_sqft, deal.attic,
    deal.construction_type, deal.foundation_type, deal.roof_type, deal.roof_age, deal.hvac_type, deal.hvac_age, deal.water_heater, deal.water_heater_age, deal.electrical_panel, deal.plumbing_type, deal.windows, deal.flooring,
    deal.kitchen_condition, deal.bath_condition, deal.exterior_condition, deal.interior_condition, deal.overall_condition,
    deal.asking_price, deal.arv_stated, deal.repair_estimate_stated, deal.repair_scope_notes, deal.assignment_fee, deal.net_to_seller,
    deal.price_per_sqft, deal.arv_per_sqft, deal.equity_stated, deal.equity_percent, deal.profit_potential_stated, deal.monthly_rent_current, deal.monthly_rent_market, deal.gross_rent_multiplier, deal.cap_rate_stated, deal.noi_stated, deal.annual_gross_income, deal.annual_expenses_stated, deal.cash_on_cash_stated, deal.price_reduction_from, deal.price_reduction_amount,
    deal.contact_1_name, deal.contact_1_title, deal.contact_1_company, deal.contact_1_email, deal.contact_1_phone, deal.contact_1_phone_2, deal.contact_1_website, deal.contact_1_social,
    deal.contact_2_name, deal.contact_2_title, deal.contact_2_company, deal.contact_2_email, deal.contact_2_phone, deal.contact_2_phone_2, deal.contact_2_website, deal.contact_2_social,
    deal.contact_3_name, deal.contact_3_title, deal.contact_3_company, deal.contact_3_email, deal.contact_3_phone, deal.contact_3_website,
    deal.all_emails_found, deal.all_phones_found, deal.all_names_found, deal.all_companies_found,
    deal.seller_name, deal.seller_phone, deal.seller_email, deal.seller_situation, deal.seller_motivation, deal.seller_timeline, deal.seller_asking_originally, deal.how_seller_was_found,
    deal.occupancy_status, deal.tenant_name, deal.tenant_phone, deal.tenant_email, deal.tenant_lease_end, deal.tenant_monthly_rent, deal.tenant_is_section_8,
    deal.days_on_market, deal.list_date, deal.offer_deadline, deal.close_timeline, deal.close_date_target, deal.inspection_period, deal.earnest_money, deal.earnest_money_type, deal.financing_terms, deal.cash_only,
    deal.title_company, deal.title_contact, deal.title_phone, deal.closing_agent, deal.contract_type, deal.assignment_allowed, deal.double_close_ok,
    deal.hoa, deal.hoa_name, deal.hoa_fee, deal.hoa_fee_frequency, deal.hoa_phone, deal.hoa_includes, deal.hoa_special_assessment,
    deal.flood_zone, deal.flood_zone_code, deal.flood_insurance_required, deal.flood_insurance_cost,
    deal.school_district, deal.elementary_school, deal.middle_school, deal.high_school,
    deal.taxes_annual, deal.taxes_monthly, deal.tax_year, deal.homestead_exempt,
    deal.insurance_annual, deal.insurance_monthly,
    deal.utility_water, deal.utility_sewer, deal.utility_electric, deal.utility_gas, deal.utility_trash, deal.utility_notes,
    deal.zoning, deal.zoning_description, deal.parcel_id, deal.legal_description, deal.mls_number, deal.mls_status,
    deal.property_link_main, deal.zillow_link, deal.redfin_link, deal.mls_link, deal.propstream_link, deal.google_maps_link, deal.virtual_tour_link, deal.video_link, deal.dropbox_link, deal.drive_link, deal.all_links_found,
    deal.photos_included ? 'YES' : 'NO', deal.photo_count, deal.photo_link_1, deal.photo_link_2, deal.photo_link_3, deal.all_photo_links,
    deal.comp_1_address, deal.comp_1_price, deal.comp_1_sqft, deal.comp_1_date, deal.comp_1_price_sqft,
    deal.comp_2_address, deal.comp_2_price, deal.comp_2_sqft, deal.comp_2_date, deal.comp_2_price_sqft,
    deal.comp_3_address, deal.comp_3_price, deal.comp_3_sqft, deal.comp_3_date, deal.comp_3_price_sqft,
    deal.comp_4_address, deal.comp_4_price, deal.comp_4_sqft, deal.comp_4_date, deal.comp_4_price_sqft,
    deal.comp_5_address, deal.comp_5_price, deal.comp_5_sqft, deal.comp_5_date, deal.comp_5_price_sqft, deal.arv_comp_notes,
    deal.deal_source, deal.list_name, deal.campaign_name, deal.marketing_headline, deal.wholesaler_deal_id, deal.wholesaler_notes,
    deal.repair_items_detailed, deal.what_needs_work, deal.what_is_updated, deal.permits_required,
    deal.code_violations, deal.liens_known, deal.title_issues_known, deal.probate, deal.bankruptcy, deal.divorce, deal.foreclosure_status, deal.auction_info,
    deal.additional_notes, deal.red_flags, deal.highlights,
    subject, messageId
  ];
}

async function ensureSheetSetup(sheets) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID });
  const existing = spreadsheet.data.sheets.map(s => s.properties.title);
  const requests = [];
  if (!existing.includes('Active Deals')) requests.push({ addSheet: { properties: { title: 'Active Deals', index: 0 } } });
  if (!existing.includes('Deal Storage')) requests.push({ addSheet: { properties: { title: 'Deal Storage', index: 1 } } });
  if (!existing.includes('Price Changes')) requests.push({ addSheet: { properties: { title: 'Price Changes', index: 2 } } });
  if (!existing.includes("Derek's Brain")) requests.push({ addSheet: { properties: { title: "Derek's Brain", index: 3 } } });
  if (requests.length > 0) await sheets.spreadsheets.batchUpdate({ spreadsheetId: process.env.GOOGLE_SHEET_ID, requestBody: { requests } });
  const check = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A1:A1' }).catch(() => ({ data: {} }));
  if (!check.data.values) {
    const h = getHeaders();
    for (const tab of ['Active Deals', 'Deal Storage', 'Price Changes']) {
      await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A1`, valueInputOption: 'RAW', requestBody: { values: [h] } });
    }
    await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: "Derek's Brain!A1", valueInputOption: 'RAW', requestBody: { values: [['Wholesaler Email','Company','Format Type','What Worked','Watch Out For','Fields Extracted','Last Seen']] } });
    console.log('✅ All tabs ready');
  }
}

async function checkDuplicate(sheets, deal) {
  const nn = normalizeAddress(deal.address, deal.city, deal.zip);
  if (!nn) return { isDuplicate: false };
  for (const tab of ['Active Deals', 'Deal Storage']) {
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: `${tab}!A:AZ` }).catch(() => ({ data: { values: [] } }));
    const rows = r.data.values || [];
    for (let i = 1; i < rows.length; i++) {
      if (nn === normalizeAddress(rows[i][2], rows[i][4], rows[i][6])) {
        const np = deal.raw_asking_price_number;
        const op = parseFloat((rows[i][47] || '').toString().replace(/[^0-9.]/g, ''));
        if (np && op && Math.abs(np - op) > 500) return { isDuplicate: true, isPriceChange: true, oldPrice: rows[i][47], newPrice: deal.asking_price };
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
          console.log(`🔁 Dupe: ${deal.address}`); dupes++;
        } else if (dup.isPriceChange) {
          const row = buildRow(deal, subject, uid);
          row[47] = `${dup.newPrice} (was ${dup.oldPrice})`;
          await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Price Changes!A:A', valueInputOption: 'RAW', requestBody: { values: [row] } });
          console.log(`💰 Price change: ${deal.address}`); priceChanges++;
        } else {
          await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:A', valueInputOption: 'RAW', requestBody: { values: [buildRow(deal, subject, uid)] } });
          console.log(`✅ ${deal.address}, ${deal.city}`); newDeals++;
        }

        await logLesson(deal.contact_1_email || from, deal.contact_1_company || '', deal.wholesaler_email_format || 'unknown', deal.what_worked_in_parsing || '', deal.watch_out_for || '', deal.fields_extracted_count || 0);

        processedIds.add(uid);
        await new Promise(r => setTimeout(r, 1500));
      }
    } finally { lock.release(); }

    const sheets = getSheetsClient();
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:AZ' }).catch(() => ({ data: { values: [] } }));
    const rows = r.data.values || [];
    if (rows.length > 1) {
      const now = new Date(), active = [rows[0]], expired = [];
      for (let i = 1; i < rows.length; i++) { new Date(rows[i][1]) < now ? expired.push(rows[i]) : active.push(rows[i]); }
      if (expired.length > 0) {
        await sheets.spreadsheets.values.clear({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A:AZ' });
        await sheets.spreadsheets.values.update({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Active Deals!A1', valueInputOption: 'RAW', requestBody: { values: active } });
        await sheets.spreadsheets.values.append({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: 'Deal Storage!A:A', valueInputOption: 'RAW', requestBody: { values: expired } });
        console.log(`📦 Archived ${expired.length}`);
      }
    }
  } catch (err) { console.error('❌', err.message); }
  finally { try { await client.logout(); } catch {} }

  console.log(`📊 ${newDeals} new | ${dupes} dupes | ${priceChanges} price changes | ${skipped} skipped`);
}

const POLL_MS = parseInt(process.env.POLL_INTERVAL || '900000');
console.log('🤙 Derek — Coralstone Capital Group');
console.log(`📬 ${process.env.IMAP_USER} | every ${POLL_MS / 60000}min\n`);
pollInbox();
setInterval(pollInbox, POLL_MS);
