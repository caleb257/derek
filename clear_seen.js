
require('dotenv').config({ path: '../.env' });
const { google } = require('googleapis');

const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
const auth = new google.auth.JWT(creds.client_email, null, creds.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']);
const sheets = google.sheets({ version: 'v4', auth });

async function run() {
  // Get UIDs already written to Active Deals (col CV = index 99 = col CV)
  const dealsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Active Deals!CV:CV'
  });
  const writtenUIDs = new Set((dealsRes.data.values || []).slice(1).map(r => r[0]).filter(Boolean));
  console.log('UIDs already in sheet:', [...writtenUIDs].join(', '));

  // Get all seen UIDs
  const seenRes = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Seen!A:A'
  });
  const allSeen = (seenRes.data.values || []).slice(1).map(r => r[0]).filter(Boolean);
  console.log('Total seen UIDs:', allSeen.length, '| Max:', Math.max(...allSeen.map(Number)));

  // Keep only UIDs that are already written as deals (protect against re-extraction)
  // Clear everything else so Derek re-screens with the new logic
  const keepUIDs = allSeen.filter(uid => writtenUIDs.has(uid));
  console.log('Keeping', keepUIDs.length, 'UIDs (already extracted deals)');
  console.log('Clearing', allSeen.length - keepUIDs.length, 'UIDs to re-screen');

  // Write back just the kept UIDs
  const newValues = [['Email UID'], ...keepUIDs.map(uid => [uid])];
  await sheets.spreadsheets.values.clear({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Seen!A:A'
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: 'Seen!A1',
    valueInputOption: 'RAW',
    requestBody: { values: newValues }
  });
  console.log('Done — Derek will rescan', allSeen.length - keepUIDs.length, 'emails on next poll');
}
run().catch(console.error);
