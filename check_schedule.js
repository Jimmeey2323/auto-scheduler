import { google } from 'googleapis';
import 'dotenv/config';

async function checkScheduleSheet() {
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
  
  // Check Schedule sheet
  console.log('\n=== SCHEDULE SHEET ===');
  const scheduleRes = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID,
    range: 'Schedule!A1:AO50'
  });
  
  const scheduleValues = scheduleRes.data.values || [];
  console.log('Row 2 (Dates):', scheduleValues[1]?.slice(18, 28));
  console.log('Row 3 (Days):', scheduleValues[2]?.slice(18, 28));
  console.log('Row 4 (Headers):', scheduleValues[3]?.slice(18, 28));
  
  console.log('\nSaturday data (columns around index 23-27):');
  for (let i = 4; i < Math.min(20, scheduleValues.length); i++) {
    const row = scheduleValues[i];
    const satLocation = row[23];
    const satClass = row[24];
    const satTrainer1 = row[25];
    const satTrainer2 = row[26];
    const satCover = row[27];
    
    if (satClass && satClass.trim()) {
      console.log(`Row ${i+1}: Loc="${satLocation}" Class="${satClass}" T1="${satTrainer1}" T2="${satTrainer2}" Cover="${satCover}"`);
    }
  }
  
  // Check Cleaned sheet
  console.log('\n=== CLEANED SHEET ===');
  const cleanedRes = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID,
    range: 'Cleaned!A1:H200'
  });
  
  const cleanedValues = cleanedRes.data.values || [];
  console.log('\nSaturday 11:30 classes at Kemps:');
  cleanedValues.forEach((row, i) => {
    if (i === 0) return;
    const day = row[0];
    const time = row[1];
    const location = row[2];
    const className = row[3];
    const trainer = row[4];
    const notes = row[5];
    
    if (day === 'Saturday' && time && time.includes('11:30') && location && location.includes('Kwality')) {
      console.log(`Row ${i}: ${day} ${time} - ${className} - ${trainer} | Notes: ${notes}`);
    }
  });
}

checkScheduleSheet().catch(console.error);
