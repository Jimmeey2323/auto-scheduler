const { google } = require('googleapis');
const path = require('path');

async function debugTuesday() {
    const auth = new google.auth.GoogleAuth({
        keyFile: path.join(__dirname, 'service-account.json'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = '1pHfT7i-MOQWVFNhgiHmchVG2J9cNoYMgDLFNKfO4dx8';
    
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'Schedule!A1:AO130'
    });

    const rows = response.data.values || [];
    console.log('Total rows:', rows.length);
    
    // Row 3 (index 2) has days
    const dayRow = rows[2] || [];
    console.log('Day row (index 2):', dayRow.slice(7, 13)); // Tuesday columns start at 7
    
    // Check rows 110-125 for Tuesday data (columns 7-12)
    // Tuesday: locationCol: 7, classCol: 8, trainer1Col: 9, trainer2Col: 10, coverCol: 12
    console.log('\n--- Tuesday data in rows 110-125 ---');
    for (let i = 110; i < Math.min(125, rows.length); i++) {
        const row = rows[i] || [];
        const time = row[0]; // Time column
        const location = row[7];
        const className = row[8];
        const trainer = row[9];
        
        if (location || className) {
            console.log(`Row ${i}: Time=${time}, Location=${location}, Class=${className}, Trainer=${trainer}`);
        }
    }
    
    // Also check if Recovery class appears anywhere
    console.log('\n--- Searching for Recovery class ---');
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i] || [];
        const rowStr = row.join(' ').toLowerCase();
        if (rowStr.includes('recovery') || rowStr.includes('8:00 pm') || rowStr.includes('8 pm')) {
            console.log(`Row ${i}: ${row.slice(0, 15).join(' | ')}`);
        }
    }
}

debugTuesday().catch(console.error);
