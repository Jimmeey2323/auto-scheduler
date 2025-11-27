require('dotenv').config();
const { google } = require('googleapis');

async function checkSaturday1130() {
  const accessToken = process.env.GOOGLE_ACCESS_TOKEN;
  const oauth2Client = new google.auth.OAuth2();
  oauth2Client.setCredentials({ access_token: accessToken });
  
  const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
  
  // Check Schedule sheet
  const scheduleResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: '1OhhnD-9R_876ehw1xROZpyd0VcCLTwuuUUnh3A1Alv4',
    range: 'Schedule!A1:AR125'
  });
  
  const scheduleData = scheduleResponse.data.values;
  console.log('\n=== SCHEDULE SHEET - Saturday 11:30 AM ===');
  
  // Find Saturday column (around column 28-33)
  const headerRow = scheduleData[3];
  const saturdayIdx = headerRow.findIndex(h => String(h).toLowerCase().includes('saturday'));
  console.log(`Saturday starts at column index: ${saturdayIdx}`);
  
  // Row 52 should be 11:30 AM based on previous logs
  const row52 = scheduleData[51]; // 0-indexed
  console.log(`Row 52 data (columns ${saturdayIdx} to ${saturdayIdx+5}):`);
  console.log(`  Time: ${scheduleData[51][0]}`);
  console.log(`  Location: ${row52[saturdayIdx]}`);
  console.log(`  Class: ${row52[saturdayIdx+1]}`);
  console.log(`  Trainer 1: ${row52[saturdayIdx+2]}`);
  console.log(`  Trainer 2: ${row52[saturdayIdx+3]}`);
  console.log(`  Cover: ${row52[saturdayIdx+4]}`);
  
  // Check Cleaned sheet
  const cleanedResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: '1OhhnD-9R_876ehw1xROZpyd0VcCLTwuuUUnh3A1Alv4',
    range: 'Cleaned!A1:I200'
  });
  
  const cleanedData = cleanedResponse.data.values;
  console.log('\n=== CLEANED SHEET - Saturday 11:30 AM ===');
  
  const saturday1130 = cleanedData.filter(row => 
    row[0] && row[0].includes('Saturday') && 
    row[1] && row[1].includes('11:30') &&
    row[2] && row[2].toLowerCase().includes('kwality')
  );
  
  saturday1130.forEach((row, idx) => {
    console.log(`\nEntry ${idx + 1}:`);
    console.log(`  Day: ${row[0]}`);
    console.log(`  Time: ${row[1]}`);
    console.log(`  Location: ${row[2]}`);
    console.log(`  Class: ${row[3]}`);
    console.log(`  Trainer: ${row[4]}`);
    console.log(`  Notes: ${row[5]}`);
  });
}

checkSaturday1130().catch(console.error);
