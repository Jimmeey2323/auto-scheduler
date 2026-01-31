import { google } from 'googleapis';
import fs from 'fs';

async function checkThemes() {
    const keys = JSON.parse(fs.readFileSync('keys.json'));
    const SPREADSHEET_ID = '18uHEPdM1JFNq1A_LGg0LnLhIc0Eo7wEE_gGBbfrlVZk';
    
    const auth = new google.auth.GoogleAuth({
        credentials: keys,
        scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    
    const sheets = google.sheets({ version: 'v4', auth });
    
    // Get Cleaned sheet data
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Cleaned!A1:H200',
    });
    
    const rows = response.data.values;
    if (!rows) {
        console.log('No data');
        return;
    }
    
    // Find Saturday and Sunday cycle classes
    console.log('Saturday/Sunday cycle classes with themes:');
    rows.forEach((row, i) => {
        const day = (row[0] || '').toString().toLowerCase();
        const classType = (row[3] || '').toString().toLowerCase();
        const theme = (row[7] || '').toString();
        
        if ((day.includes('saturday') || day.includes('sunday')) && 
            classType.includes('cycle') && 
            theme) {
            console.log(`Row ${i + 1}: Day=${row[0]}, Time=${row[1]}, Class=${row[3]}, Theme=${theme}`);
        }
    });
    
    // Also check all classes with themes
    console.log('\n\nAll classes with themes (column H):');
    rows.forEach((row, i) => {
        const theme = (row[7] || '').toString().trim();
        if (theme) {
            console.log(`Row ${i + 1}: Day=${row[0]}, Time=${row[1]}, Class=${row[3]}, Theme=${theme}`);
        }
    });
}

checkThemes().catch(console.error);
