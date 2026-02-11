const fs = require('fs');
const path = require('path');

// Load credentials
const credFile = './gmail_auth_config.txt';
const cred = require(path.resolve(credFile));

const { google } = require('googleapis');

async function main() {
  const auth = new google.auth.OAuth2(
    cred.client_id,
    cred.client_secret,
    cred.redirect_uri
  );
  
  auth.setCredentials({ access_token: cred.access_token });
  
  const sheets = google.sheets({ version: 'v4', auth });
  
  // Fetch all data
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: '1sYLMlPVLfFSmBVl8_d_iUWP7MKq-1Pq3BdMQO7qC1Bw',
    range: 'Schedule!A:H'
  });
  
  const rows = result.data.values || [];
  
  // Get header row
  const headers = rows[0] || [];
  console.log('Headers:', headers);
  console.log('');
  
  // Find Back Body Blaze classes
  const backBodyClasses = rows.filter((row, idx) => {
    if (idx === 0) return false;
    return row[7] && row[7].toLowerCase().includes('back body');
  });
  
  // Group by location
  const byLocation = {};
  backBodyClasses.forEach(row => {
    const location = row[0] || 'Unknown';
    if (!byLocation[location]) byLocation[location] = [];
    byLocation[location].push(row);
  });
  
  // Print grouped results
  for (const [location, classes] of Object.entries(byLocation)) {
    console.log(`\n${location}:`);
    classes.forEach(row => {
      const day = row[1] || '';
      const time = row[2] || '';
      const trainer = row[6] || '';
      const classType = row[7] || '';
      console.log(`  ${day} ${time}: ${classType} - ${trainer}`);
    });
  }
}

main().catch(console.error);
