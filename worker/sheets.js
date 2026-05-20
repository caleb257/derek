const { google } = require('googleapis');

function getSheetsClient() {
  let credentials;

  // Method 1: Full JSON credentials (preferred)
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try {
      credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    } catch (e) {
      console.error('Failed to parse GOOGLE_CREDENTIALS_JSON:', e.message);
    }
  }

  // Method 2: Individual env vars fallback
  if (!credentials) {
    let key = process.env.GOOGLE_PRIVATE_KEY || '';
    key = key.replace(/\\n/g, '\n').replace(/^["']|["']$/g, '');
    credentials = {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: key,
    };
  }

  const auth = new google.auth.JWT({
    email: credentials.client_email,
    key: credentials.private_key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return google.sheets({ version: 'v4', auth });
}

module.exports = { getSheetsClient };
