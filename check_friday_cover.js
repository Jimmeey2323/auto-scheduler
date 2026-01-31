import { google } from 'googleapis';
import 'dotenv/config';

async function checkFridayCover() {
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
  
  // Check Schedule sheet Friday columns
  console.log('=== SCHEDULE SHEET - FRIDAY ===');
  const scheduleRes = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID,
    range: 'Schedule!A1:AO50'
  });
  
  const scheduleValues = scheduleRes.data.values || [];
  console.log('Row 3 (Days) cols 23-28:', scheduleValues[2]?.slice(23, 28));
  console.log('Row 4 (Headers) cols 23-28:', scheduleValues[3]?.slice(23, 28));
  
  console.log('\nFriday 8:30 AM rows:');
  for (let i = 4; i < Math.min(25, scheduleValues.length); i++) {
    const row = scheduleValues[i];
    const time = row[0];
    if (time && time.includes('8:30')) {
      const friLocation = row[23];
      const friClass = row[24];
      const friTrainer1 = row[25];
      const friTrainer2 = row[26];
      const friCover = row[27];
      console.log(`Row ${i+1}: Time="${time}" Loc="${friLocation}" Class="${friClass}" T1="${friTrainer1}" T2="${friTrainer2}" Cover="${friCover}"`);
    }
  }
  
  // Check Cleaned sheet
  console.log('\n=== CLEANED SHEET - FRIDAY 8:30 ===');
  const cleanedRes = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID,
    range: 'Cleaned!A1:H200'
  });
  
  const cleanedValues = cleanedRes.data.values || [];
  cleanedValues.forEach((row, i) => {
    if (i === 0) return;
    const day = row[0];
    const time = row[1];
    const location = row[2];
    const className = row[3];
    const trainer = row[4];
    const notes = row[5];
    
    if (day === 'Friday' && time && time.includes('8:30') && location && location.includes('Kwality')) {
      console.log(`Row ${i}: ${day} ${time} - ${className} - ${trainer} | Notes: ${notes}`);
    }
  });
}

checkFridayCover().catch(console.error);
