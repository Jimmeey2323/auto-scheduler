import { google } from 'googleapis';
import 'dotenv/config';

async function checkCleanedSheet() {
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET
  );
  
  const tokenResponse = await fetch(process.env.GOOGLE_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: process.env.GOOGLE_CLIENT_ID,
      client_secret: process.env.GOOGLE_CLIENT_SECRET,
      refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
      grant_type: 'refresh_token'
    })
  });
  
  const tokens = await tokenResponse.json();
  client.setCredentials({ access_token: tokens.access_token });
  
  const sheets = google.sheets({ version: 'v4', auth: client });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID,
    range: 'Cleaned!A1:H200'
  });
  
  const values = res.data.values || [];
  
  // Check for SOLD OUT or hosted classes
  console.log('\nSearching for SOLD OUT or hosted classes in Kemps...');
  let soldOutCount = 0;
  values.forEach((row, i) => {
    if (i === 0) return; // skip header
    const location = row[2] || '';
    if (!location.includes('Kwality House')) return;
    
    if (row[5] && row[5].includes('SOLD OUT')) {
      soldOutCount++;
      console.log(`Row ${i}: Day=${row[0]}, Time=${row[1]}, Class=${row[3]}, Trainer=${row[4]}, Notes=${row[5]}`);
    }
  });
  console.log(`\nTotal SOLD OUT classes at Kemps: ${soldOutCount}`);
}

checkCleanedSheet().catch(console.error);
