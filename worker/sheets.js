const { google } = require('googleapis');

function getSheetsClient() {
  // Handle private key whether it has literal newlines or \n escape sequences
  let privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
  
  // If it doesn't already have real newlines, convert \n to actual newlines
  if (!privateKey.includes('\n')) {
    privateKey = privateKey.replace(/\\n/g, '\n');
  }

  // Strip surrounding quotes if present
  privateKey = privateKey.replace(/^["']|["']$/g, '');

  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

module.exports = { getSheetsClient };
