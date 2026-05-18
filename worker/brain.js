const { google } = require('googleapis');

function getSheetsClient() {
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// Load everything Derek has learned so far
async function loadBrain() {
  try {
    const sheets = getSheetsClient();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Derek's Brain!A:F"
    }).catch(() => ({ data: { values: [] } }));

    const rows = result.data.values || [];
    if (rows.length <= 1) return '';

    // Build a plain text summary of what Derek knows
    const lessons = rows.slice(1).map(row => {
      return `- Wholesaler: ${row[0] || '?'} | Format: ${row[1] || '?'} | What worked: ${row[2] || '?'} | Watch out for: ${row[3] || '?'}`;
    }).join('\n');

    return lessons;
  } catch (e) {
    return '';
  }
}

// After a successful parse, log what Derek learned
async function logLesson(wholesalerEmail, wholesalerCompany, formatType, whatWorked, watchOutFor, fieldsExtracted) {
  try {
    const sheets = getSheetsClient();

    // Ensure brain tab exists
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID });
    const existing = spreadsheet.data.sheets.map(s => s.properties.title);

    if (!existing.includes("Derek's Brain")) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: "Derek's Brain", index: 3 } } }] }
      });
      // Write headers
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: "Derek's Brain!A1",
        valueInputOption: 'RAW',
        requestBody: { values: [['Wholesaler Email', 'Company', 'Format Type', 'What Worked', 'Watch Out For', 'Fields Extracted', 'Last Seen']] }
      });
    }

    // Check if this wholesaler already has a row and update it
    const existing_data = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Derek's Brain!A:A"
    }).catch(() => ({ data: { values: [] } }));

    const rows = existing_data.data.values || [];
    let existingRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === wholesalerEmail) { existingRow = i + 1; break; }
    }

    const rowData = [wholesalerEmail, wholesalerCompany || '', formatType || '', whatWorked || '', watchOutFor || '', fieldsExtracted.toString(), new Date().toISOString()];

    if (existingRow > 0) {
      // Update existing row
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: `Derek's Brain!A${existingRow}`,
        valueInputOption: 'RAW',
        requestBody: { values: [rowData] }
      });
    } else {
      // New wholesaler — append
      await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: "Derek's Brain!A:A",
        valueInputOption: 'RAW',
        requestBody: { values: [rowData] }
      });
    }

    console.log(`🧠 Brain updated for: ${wholesalerEmail}`);
  } catch (e) {
    console.error('Brain write error:', e.message);
  }
}

module.exports = { loadBrain, logLesson };
