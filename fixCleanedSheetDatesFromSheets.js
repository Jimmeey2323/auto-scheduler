const fs = require('fs');
const { parse } = require('csv-parse/sync');
const path = require('path');
const ScheduleUpdater = require('./updateKempsSchedule.js');

/**
 * Fix the cleaned sheet to populate proper dates starting from the first date found in Google Sheets
 */

class DatePopulatorFromSheets {
    constructor(csvPath) {
        this.csvPath = csvPath;
        this.dayOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    }

    /**
     * Parse date string and return a Date object
     */
    parseDate(dateStr) {
        if (!dateStr) return null;
        
        // Handle formats like "24 Nov 2025" or "Nov 24, 2025"
        const cleanDate = dateStr.trim();
        
        // Try parsing different formats
        const formats = [
            // "24 Nov 2025" format
            /(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})/,
            // "Nov 24, 2025" format
            /([A-Za-z]{3})\s+(\d{1,2}),?\s+(\d{4})/
        ];

        for (const format of formats) {
            const match = cleanDate.match(format);
            if (match) {
                let day, month, year;
                
                if (format === formats[0]) { // "24 Nov 2025"
                    [, day, month, year] = match;
                } else { // "Nov 24, 2025"
                    [, month, day, year] = match;
                }
                
                // Convert month abbreviation to number
                const monthMap = {
                    'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5,
                    'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11
                };
                
                const monthNum = monthMap[month.toLowerCase()];
                if (monthNum !== undefined) {
                    return new Date(parseInt(year), monthNum, parseInt(day));
                }
            }
        }
        
        // Try direct parsing as fallback
        const parsed = new Date(cleanDate);
        return isNaN(parsed.getTime()) ? null : parsed;
    }

    /**
     * Format date to "DD MMM YYYY" format
     */
    formatDate(date) {
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        
        const day = date.getDate().toString().padStart(2, '0');
        const month = months[date.getMonth()];
        const year = date.getFullYear();
        
        return `${day} ${month} ${year}`;
    }

    /**
     * Get the date for a specific day of week based on a reference date
     */
    getDateForDay(dayName, referenceDate) {
        const targetDayIndex = this.dayOrder.indexOf(dayName);
        if (targetDayIndex === -1) return null;

        const referenceDayIndex = referenceDate.getDay() === 0 ? 6 : referenceDate.getDay() - 1; // Convert Sunday=0 to Monday=0
        const dayDiff = targetDayIndex - referenceDayIndex;
        
        const targetDate = new Date(referenceDate);
        targetDate.setDate(referenceDate.getDate() + dayDiff);
        
        return targetDate;
    }

    /**
     * Get reference date from row 2 of the Google Sheets
     */
    async getReferenceDateFromSheets() {
        try {
            console.log('📊 Fetching reference dates from row 2 of Google Sheets...');
            const updater = new ScheduleUpdater('', '', '');
            
            // Get raw sheet data to read row 2 (dates row)
            const sheetData = await this.getRawSheetData();
            
            if (sheetData && sheetData.length > 1 && sheetData[1]) {
                // Row 2 (index 1) should contain dates
                const dateRow = sheetData[1];
                console.log('📅 Found row 2 with dates:', dateRow.slice(0, 10));
                
                // Find the first valid date in row 2
                for (const cellValue of dateRow) {
                    if (cellValue && cellValue.toString().trim()) {
                        const parsedDate = this.parseDate(cellValue.toString().trim());
                        if (parsedDate) {
                            console.log(`✅ Found reference date from row 2: ${cellValue}`);
                            return parsedDate;
                        }
                    }
                }
            }
            
            console.log('⚠️  No valid date found in row 2, falling back to sheet records');
            
            // Fallback to old method
            await updater.readSheet();
            for (const record of updater.allSheetRecords) {
                if (record.Date) {
                    const parsedDate = this.parseDate(record.Date);
                    if (parsedDate) {
                        console.log(`✅ Found reference date in sheets: ${record.Date} (${record.Day})`);
                        return parsedDate;
                    }
                }
            }
            
            console.log('⚠️  No valid date found in sheets, using current Monday');
            return this.getCurrentMonday();
            
        } catch (error) {
            console.error('❌ Error fetching from sheets:', error.message);
            console.log('⚠️  Falling back to current Monday');
            return this.getCurrentMonday();
        }
    }

    /**
     * Get raw sheet data to access row 2 dates
     */
    async getRawSheetData() {
        try {
            const { google } = require('googleapis');
            const fs = require('fs');
            
            // Same config as ScheduleUpdater
            const GOOGLE_CONFIG = {
                CLIENT_ID: '977135500395-pggchnpjr8ujupkk9t3f04ck86bqmdm5.apps.googleusercontent.com',
                CLIENT_SECRET: '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6UDp',
                REDIRECT_URI: 'urn:ietf:wg:oauth:2.0:oob',
                TOKEN_PATH: './gmail_token.json',
                SPREADSHEET_ID: '1_1m5V2QUB24L2mQkkYgBxrIw4K1Sn3r6MDd2TXtRNH4',
                SHEET_NAME: 'Cleaned'
            };
            
            // Get access token
            const tokenData = fs.readFileSync(GOOGLE_CONFIG.TOKEN_PATH, 'utf8');
            const tokens = JSON.parse(tokenData);
            
            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET,
                GOOGLE_CONFIG.REDIRECT_URI
            );
            
            oauth2Client.setCredentials({ refresh_token: tokens.refresh_token });
            const { credentials } = await oauth2Client.refreshAccessToken();
            oauth2Client.setCredentials({ access_token: credentials.access_token });
            
            const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
            
            // Read from Schedule sheet to get row 2 dates
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: GOOGLE_CONFIG.SPREADSHEET_ID,
                range: 'Schedule!A1:AS10' // Get first 10 rows including row 2
            });
            
            return response.data.values || [];
            
        } catch (error) {
            console.log('⚠️  Could not access raw sheet data:', error.message);
            return null;
        }
    }

    /**
     * Process the CSV and add proper dates based on Google Sheets reference
     */
    async processCSV() {
        try {
            console.log('📄 Reading CSV file...');
            const content = fs.readFileSync(this.csvPath, 'utf8');
            
            // Parse as tab-delimited
            const records = parse(content, { 
                columns: true, 
                skip_empty_lines: true,
                delimiter: '\t'
            });

            console.log(`✅ Found ${records.length} records`);
            
            // Get reference date from Google Sheets
            const referenceDate = await this.getReferenceDateFromSheets();
            console.log(`📅 Using reference date: ${this.formatDate(referenceDate)}`);

            // Process each record to add proper dates
            const processedRecords = records.map((record, index) => {
                const dayName = record.Day?.trim();
                if (dayName && this.dayOrder.includes(dayName)) {
                    const dateForDay = this.getDateForDay(dayName, referenceDate);
                    if (dateForDay) {
                        record.Date = this.formatDate(dateForDay);
                    }
                } else {
                    record.Date = ''; // Keep empty if day is not recognized
                }
                
                return record;
            });

            // Create new CSV content with Date column
            console.log('📝 Creating updated CSV...');
            const headers = ['Day', 'Time', 'Location', 'Class', 'Trainer', 'Date', 'Notes'];
            let csvContent = headers.join('\t') + '\n';
            
            processedRecords.forEach(record => {
                const row = headers.map(header => {
                    const value = record[header] || '';
                    return value.toString().trim();
                });
                csvContent += row.join('\t') + '\n';
            });

            // Save the updated CSV
            const backupPath = this.csvPath.replace('.csv', '.backup.csv');
            if (!fs.existsSync(backupPath)) {
                fs.writeFileSync(backupPath, content);
                console.log(`💾 Created backup at: ${backupPath}`);
            }

            fs.writeFileSync(this.csvPath, csvContent);
            console.log(`✅ Updated CSV saved to: ${this.csvPath}`);
            
            // Show sample of updated records
            console.log('\n📋 Sample updated records:');
            processedRecords.slice(0, 5).forEach((record, i) => {
                console.log(`  ${i + 1}. ${record.Day} ${record.Date} - ${record.Time} - ${record.Class}`);
            });

            return processedRecords;
            
        } catch (error) {
            console.error('❌ Error processing CSV:', error.message);
            throw error;
        }
    }

    /**
     * Get the current Monday as reference date
     */
    getCurrentMonday() {
        const now = new Date();
        const dayOfWeek = now.getDay(); // 0 = Sunday, 1 = Monday, ...
        const daysToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1; // How many days back to Monday
        
        const monday = new Date(now);
        monday.setDate(now.getDate() - daysToMonday);
        monday.setHours(0, 0, 0, 0); // Reset time to start of day
        
        return monday;
    }
}

// If running directly
if (require.main === module) {
    const csvPath = path.join(__dirname, 'Schedule Views - Cleaned (2).csv');
    const populator = new DatePopulatorFromSheets(csvPath);
    
    populator.processCSV()
        .then(() => {
            console.log('🎉 Date population from Google Sheets complete!');
        })
        .catch(error => {
            console.error('💥 Failed:', error.message);
            process.exit(1);
        });
}

module.exports = DatePopulatorFromSheets;