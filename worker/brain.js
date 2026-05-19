const { getSheetsClient } = require('./sheets');

async function loadBrain() {
  try {
    const sheets = getSheetsClient();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Derek's Brain!A:G"
    }).catch(() => ({ data: { values: [] } }));

    const rows = result.data.values || [];
    if (rows.length <= 1) return '';

    return rows.slice(1).map(row =>
      `- Wholesaler: ${row[0] || '?'} | Format: ${row[2] || '?'} | What worked: ${row[3] || '?'} | Watch out for: ${row[4] || '?'}`
    ).join('\n');
  } catch (e) {
    console.error('Brain load error:', e.message);
    return '';
  }
}

async function logLesson(wholesalerEmail, wholesalerCompany, formatType, whatWorked, watchOutFor, fieldsExtracted) {
  try {
    const sheets = getSheetsClient();
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID });
    const existing = spreadsheet.data.sheets.map(s => s.properties.title);

    if (!existing.includes("Derek's Brain")) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: "Derek's Brain", index: 3 } } }] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: "Derek's Brain!A1",
        valueInputOption: 'RAW',
        requestBody: { values: [['Wholesaler Email', 'Company', 'Format Type', 'What Worked', 'Watch Out For', 'Fields Extracted', 'Last Seen']] }
      });
    }

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
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: `Derek's Brain!A${existingRow}`,
        valueInputOption: 'RAW',
        requestBody: { values: [rowData] }
      });
    } else {
      await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: "Derek's Brain!A:A",
        valueInputOption: 'RAW',
        requestBody: { values: [rowData] }
      });
    }
    console.log(`🧠 Brain updated: ${wholesalerEmail}`);
  } catch (e) {
    console.error('Brain write error:', e.message);
  }
}

module.exports = { loadBrain, logLesson };
