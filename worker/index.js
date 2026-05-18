require('dotenv').config({ path: '../.env' });

const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');
const TelegramBot = require('node-telegram-bot-api');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const bot = new TelegramBot(process.env.TELEGRAM_BOT_TOKEN);

// ─── Google Sheets Auth ───────────────────────────────────────────────────────
function getSheetsClient() {
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// ─── Claude: Is this a wholesale deal email? ──────────────────────────────────
async function isDealEmail(subject, body) {
  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 100,
    messages: [{
      role: 'user',
      content: `Is this email a wholesale real estate deal? Answer only YES or NO.

Subject: ${subject}
Body: ${body.substring(0, 500)}`
    }]
  });
  return response.content[0].text.trim().toUpperCase().startsWith('YES');
}

// ─── Claude: Extract all deal data ───────────────────────────────────────────
async function extractDealData(subject, body, from) {
  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 2000,
    messages: [{
      role: 'user',
      content: `You are Derek the Dealer, an expert real estate acquisitions analyst for Coralstone Capital Group in Tampa Bay, FL.

Extract every piece of data from this wholesale deal email. Return ONLY valid JSON, no markdown, no explanation.

Email From: ${from}
Subject: ${subject}
Body: ${body}

Return this exact JSON structure (use null for missing fields):
{
  "address": null,
  "city": null,
  "state": null,
  "zip": null,
  "county": null,
  "property_type": null,
  "beds": null,
  "baths": null,
  "sqft": null,
  "year_built": null,
  "asking_price": null,
  "arv_stated": null,
  "repair_estimate_stated": null,
  "assignment_fee": null,
  "wholesaler_name": null,
  "wholesaler_company": null,
  "wholesaler_email": null,
  "wholesaler_phone": null,
  "seller_motivation": null,
  "days_on_market": null,
  "close_timeline": null,
  "property_links": null,
  "photos_included": false,
  "photo_links": null,
  "comps_provided": null,
  "flood_zone": null,
  "hoa": null,
  "occupancy_status": null,
  "title_notes": null,
  "additional_notes": null,
  "deal_source": null,
  "raw_asking_price_number": null,
  "raw_arv_number": null,
  "raw_repair_number": null
}`
    }]
  });

  try {
    return JSON.parse(response.content[0].text.trim());
  } catch (e) {
    console.error('JSON parse error:', e.message);
    return null;
  }
}

// ─── Google Sheets: Ensure tabs exist ────────────────────────────────────────
async function ensureSheetTabs(sheets) {
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID
  });

  const existingSheets = spreadsheet.data.sheets.map(s => s.properties.title);
  const requests = [];

  if (!existingSheets.includes('Active Deals')) {
    requests.push({ addSheet: { properties: { title: 'Active Deals' } } });
  }
  if (!existingSheets.includes('Deal Storage')) {
    requests.push({ addSheet: { properties: { title: 'Deal Storage' } } });
  }

  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      requestBody: { requests }
    });
  }

  // Add headers if Active Deals is new
  const activeData = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Active Deals!A1:A1'
  });

  if (!activeData.data.values) {
    const headers = [
      'Date Received', 'Expires', 'Address', 'City', 'State', 'Zip', 'County',
      'Property Type', 'Beds', 'Baths', 'Sqft', 'Year Built',
      'Asking Price', 'ARV (Stated)', 'Repairs (Stated)', 'Assignment Fee',
      'Wholesaler Name', 'Wholesaler Company', 'Wholesaler Email', 'Wholesaler Phone',
      'Seller Motivation', 'Days on Market', 'Close Timeline',
      'Property Links', 'Photos Included', 'Photo Links', 'Comps Provided',
      'Flood Zone', 'HOA', 'Occupancy', 'Title Notes',
      'Deal Source', 'Additional Notes', 'Email Subject', 'Raw Email ID'
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: 'Active Deals!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [headers] }
    });

    // Same headers for storage
    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: 'Deal Storage!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [headers] }
    });
  }
}

// ─── Google Sheets: Write deal row ───────────────────────────────────────────
async function writeDealToSheet(sheets, deal, subject, messageId) {
  const now = new Date();
  const expires = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  const row = [
    now.toISOString(),
    expires.toISOString(),
    deal.address,
    deal.city,
    deal.state,
    deal.zip,
    deal.county,
    deal.property_type,
    deal.beds,
    deal.baths,
    deal.sqft,
    deal.year_built,
    deal.asking_price,
    deal.arv_stated,
    deal.repair_estimate_stated,
    deal.assignment_fee,
    deal.wholesaler_name,
    deal.wholesaler_company,
    deal.wholesaler_email,
    deal.wholesaler_phone,
    deal.seller_motivation,
    deal.days_on_market,
    deal.close_timeline,
    deal.property_links,
    deal.photos_included ? 'YES' : 'NO',
    deal.photo_links,
    deal.comps_provided,
    deal.flood_zone,
    deal.hoa,
    deal.occupancy_status,
    deal.title_notes,
    deal.deal_source,
    deal.additional_notes,
    subject,
    messageId
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Active Deals!A:A',
    valueInputOption: 'RAW',
    requestBody: { values: [row] }
  });

  console.log(`✅ Deal written to sheet: ${deal.address}`);
}

// ─── Google Sheets: Move expired deals to storage ────────────────────────────
async function archiveExpiredDeals(sheets) {
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Active Deals!A:AJ'
  });

  const rows = result.data.values;
  if (!rows || rows.length <= 1) return;

  const now = new Date();
  const headers = rows[0];
  const activeRows = [headers];
  const expiredRows = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const expiresAt = new Date(row[1]);
    if (expiresAt < now) {
      expiredRows.push(row);
    } else {
      activeRows.push(row);
    }
  }

  if (expiredRows.length === 0) return;

  // Write active rows back
  await sheets.spreadsheets.values.clear({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Active Deals!A:AJ'
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Active Deals!A1',
    valueInputOption: 'RAW',
    requestBody: { values: activeRows }
  });

  // Append expired to storage
  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Deal Storage!A:A',
    valueInputOption: 'RAW',
    requestBody: { values: expiredRows }
  });

  console.log(`📦 Archived ${expiredRows.length} expired deal(s) to storage`);
}

// ─── Telegram: Send new deal notification ────────────────────────────────────
async function sendTelegramAlert(deal, subject) {
  const address = [deal.address, deal.city, deal.state, deal.zip]
    .filter(Boolean).join(', ') || 'Address TBD';

  const msg = `🏠 *NEW DEAL — Derek the Dealer*

📍 *${address}*
💰 Ask: ${deal.asking_price || 'Not listed'}
📊 ARV (stated): ${deal.arv_stated || 'Not provided'}
🔨 Repairs (stated): ${deal.repair_estimate_stated || 'Not provided'}
🏗 Type: ${deal.property_type || 'Unknown'} | ${deal.beds || '?'}bd/${deal.baths || '?'}ba | ${deal.sqft || '?'} sqft
👤 Source: ${deal.wholesaler_name || 'Unknown'} ${deal.wholesaler_company ? `@ ${deal.wholesaler_company}` : ''}
📸 Photos: ${deal.photos_included ? 'YES' : 'NO'}
🔗 ${deal.property_links || 'No link provided'}

_Check the sheet for full data._`;

  await bot.sendMessage(process.env.TELEGRAM_CHAT_ID, msg, { parse_mode: 'Markdown' });
  console.log(`📱 Telegram alert sent for: ${address}`);
}

// ─── Track processed emails ───────────────────────────────────────────────────
const processedIds = new Set();

// ─── Main poll function ───────────────────────────────────────────────────────
async function pollInbox() {
  console.log(`\n🔍 [${new Date().toLocaleTimeString()}] Polling inbox...`);

  const client = new ImapFlow({
    host: process.env.IMAP_HOST,
    port: parseInt(process.env.IMAP_PORT || '993'),
    secure: true,
    auth: {
      user: process.env.IMAP_USER,
      pass: process.env.IMAP_PASSWORD
    },
    logger: false
  });

  try {
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');

    try {
      // Get emails from last 24 hours
      const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
      const messages = client.fetch({ since }, { source: true, envelope: true });

      for await (const msg of messages) {
        const uid = msg.uid.toString();
        if (processedIds.has(uid)) continue;

        const parsed = await simpleParser(msg.source);
        const subject = parsed.subject || '';
        const body = parsed.text || parsed.html || '';
        const from = parsed.from?.text || '';

        // Skip if not a deal
        const isDeal = await isDealEmail(subject, body);
        if (!isDeal) {
          processedIds.add(uid);
          continue;
        }

        console.log(`📬 Deal found: ${subject}`);

        // Extract all data
        const dealData = await extractDealData(subject, body, from);
        if (!dealData) {
          console.error('Failed to extract deal data');
          processedIds.add(uid);
          continue;
        }

        // Write to Google Sheets
        const sheets = getSheetsClient();
        await ensureSheetTabs(sheets);
        await writeDealToSheet(sheets, dealData, subject, uid);

        // Send Telegram alert
        await sendTelegramAlert(dealData, subject);

        processedIds.add(uid);
      }
    } finally {
      lock.release();
    }

    // Archive expired deals
    const sheets = getSheetsClient();
    await archiveExpiredDeals(sheets);

  } catch (err) {
    console.error('Poll error:', err.message);
  } finally {
    await client.logout();
  }
}

// ─── Start ────────────────────────────────────────────────────────────────────
console.log('🤙 Derek the Dealer — Deal Harvesting Worker');
console.log(`📬 Monitoring: ${process.env.IMAP_USER}`);
console.log(`⏱  Polling every ${parseInt(process.env.POLL_INTERVAL || '600000') / 60000} minutes\n`);

pollInbox();
setInterval(pollInbox, parseInt(process.env.POLL_INTERVAL || '600000'));
