import { GoogleAuth } from 'google-auth-library';
import { google } from 'googleapis';

async function test() {
  const auth = new GoogleAuth({
    keyFile: 'gmail_auth_config.txt'
  });
  
  const sheets = google.sheets({ version: 'v4', auth });
  const TARGET_SPREADSHEET_ID = '1OhhnD-9R_876ehw1xROZpyd0VcCLTwuuUUnh3A1Alv4';
  
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: TARGET_SPREADSHEET_ID,
      range: 'Cleaned!A1:H100'
    });
    
    const rows = res.data.values || [];
    console.log('Saturday 11:30 AM entries in Cleaned sheet:\n');
    // Find Saturday 11:30 AM rows
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === 'Saturday' && rows[i][1] && rows[i][1].includes('11:30')) {
        console.log(`Row ${i+1}:`);
        console.log(`  Day: ${rows[i][0]}`);
        console.log(`  Time: ${rows[i][1]}`);
        console.log(`  Location: ${rows[i][2]}`);
        console.log(`  Class: ${rows[i][3]}`);
        console.log(`  Trainer: ${rows[i][4]}`);
        console.log(`  Notes: ${rows[i][5]}`);
        console.log(`  Date: ${rows[i][6]}`);
        console.log(`  Theme: ${rows[i][7]}\n`);
      }
    }
  } catch (err) {
    console.error('Error:', err.message);
  }
}

test();
