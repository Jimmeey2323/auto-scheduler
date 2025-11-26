const fs = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');
const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const { google } = require('googleapis');
const axios = require('axios');
const { PDFDocument } = require('pdf-lib');
require('dotenv').config();

// Google OAuth Configuration
const GOOGLE_CONFIG = {
  CLIENT_ID: process.env.GOOGLE_CLIENT_ID,
  CLIENT_SECRET: process.env.GOOGLE_CLIENT_SECRET,
  REFRESH_TOKEN: process.env.GOOGLE_REFRESH_TOKEN,
  TOKEN_URL: process.env.GOOGLE_TOKEN_URL || "https://oauth2.googleapis.com/token",
  FOLDER_ID: process.env.GOOGLE_FOLDER_ID || "1PPaEKOBcPtjSUpFZZArkRLEBcGO5h108",
  SPREADSHEET_ID: process.env.GOOGLE_SPREADSHEET_ID || "1OhhnD-9R_876ehw1xROZpyd0VcCLTwuuUUnh3A1Alv4",
  SHEET_NAME: process.env.GOOGLE_SHEET_NAME || "Cleaned",
  TARGET_SPREADSHEET_ID: process.env.GOOGLE_SPREADSHEET_ID || "1OhhnD-9R_876ehw1xROZpyd0VcCLTwuuUUnh3A1Alv4",
  TARGET_SHEET_NAME: "Schedule"
};

// Email processing configuration
const EMAIL_CONFIG = {
  SENDER_EMAIL: "mrigakshi@physique57mumbai.com",
  SUBJECT_KEYWORD: "Mumbai schedule", // More specific to target current week format
  TRAINERS: [
    "Saniya", "Anisha", "Rohan", "Debs", "Richard", "Atulan", "Pranjali", 
    "Mriga", "Janhavi", "Reshma", "Vivaran", "Simran", "Karan", "Cauveri", 
    "Simonelle", "Anmol", "Raunak", "Bret", "Sovena"
  ],
  CLASS_TYPES: [
    "Barre57", "Barre57 exp", "Cardio B", "Cardio B exp", "Amped Up", "HIIT", 
    "BBB", "MAT57", "Recovery", "Mansee", "Anandita", "Kajal", "Taarika", 
    "Pooja", "Hosted", "amped Up exp", "Mat 57 exp", "BBB exp", "Megha", 
    "Smita Parekh", "Foundations", "Neeta", "Trainer's Choice", "Namrata", 
    "FIT", "Cardio B+", "Sweat", "Shuchi", "PreNatal", "CYCLE", "CYCLE EXP", 
    "Sakshi", "Nandini", "Strength - FB", "Strength - Push", "Strength - Pull"
  ],
  LOCATIONS: [
    "Kemps", "Colaba", "Virtual-S", "Juhu", "PVT - Virt", "Bandra", "Bbay Gym", 
    "Annex", "Bay club", "Sound Rise"
  ]
};

/**
 * Advanced Node.js Script to Update Kemps.html with CSV Data
 * Reads class data from CSV and updates HTML in accurate positions
 * without altering styling, layout or structure
 * Generates PDF and uploads to Google Drive
 */

class ScheduleUpdater {
    constructor(htmlPath, outputPath, location = 'kemps') {
        this.htmlPath = htmlPath;
        this.outputPath = outputPath || htmlPath;
        this.kwalityClasses = [];
        this.allSheetRecords = [];
        this.$ = null;
        this.currentLocation = location.toLowerCase(); // Track current location for theme badge styling
        this.locationName = this.currentLocation.charAt(0).toUpperCase() + this.currentLocation.slice(1); // 'Kemps' or 'Bandra'
    }

    /**
     * Convert a string to Title Case while preserving common acronyms
     * - Keeps ALL-CAPS words of length 2-5 (e.g., HIIT, TRX)
     * - Title-cases hyphenated and apostrophe words (e.g., full-body, o'connor)
     * - Leaves short conjunctions/articles lowercase unless first word
     */
    toTitleCase(str) {
        if (!str) return '';
        const lowerExceptions = new Set(['and', 'or', 'the', 'a', 'an', 'of', 'for', 'with', 'on', 'in', 'to', 'at', 'by', 'from']);
        const isAllCapsAcronym = (w) => /^[A-Z]{2,5}$/.test(w);
        const smartCap = (s) => {
            if (!s) return s;
            return s.replace(/^(\P{L}*)(\p{L})([\p{L}']*)/u, (_m, lead, first, rest) => {
                return lead + first.toUpperCase() + (rest ? rest.toLowerCase() : '');
            });
        };

        const capWord = (word, isFirst) => {
            if (!word) return word;
            if (isAllCapsAcronym(word)) return word; // Preserve acronyms

            // Handle mixed punctuation like hyphens and apostrophes
            return word
                .split('-')
                .map(segment => segment
                    .split("'")
                    .map((part, idx) => {
                        const base = idx === 0 && !isFirst && lowerExceptions.has(part.toLowerCase()) ? part.toLowerCase() : smartCap(part);
                        return base;
                    })
                    .join("'"))
                .join('-');
        };

        // Collapse whitespace, then process
        const tokens = String(str).trim().replace(/\s+/g, ' ').split(' ');
        const titled = tokens.map((w, i) => capWord(w, i === 0)).join(' ');
        return this.applyBrandCasing(titled);
    }

    /**
     * Fix known brand/style casing after Title Case
     */
    applyBrandCasing(str) {
        if (!str) return str;
        return str
            .replace(/\bPowercycle\b/g, 'PowerCycle');
    }

    /**
     * Read schedule data from Google Sheets ("Cleaned" tab)
     */
    async readSheet() {
        console.log('📄 Reading Google Sheet data...');
        try {
            const accessToken = await this.getAccessToken();

            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });

            const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
            // Read a wide range to include all columns; adjust if needed
            const range = `${GOOGLE_CONFIG.SHEET_NAME}!A1:Z1000`;
            const res = await sheets.spreadsheets.values.get({
                spreadsheetId: GOOGLE_CONFIG.SPREADSHEET_ID,
                range
            });

            const values = res.data.values || [];
            if (values.length === 0) {
                console.warn('⚠️  No data found in sheet.');
                this.kwalityClasses = [];
                return this.kwalityClasses;
            }

            // First row is header
            const headers = values[0].map(h => String(h).trim());
            console.log(`📋 Headers found in Google Sheet: ${headers.join(', ')}`);
            console.log(`   Total columns: ${headers.length}`);
            if (headers.length >= 8) {
                console.log(`   Column H (index 7): "${headers[7]}"`);
            }
            
            const records = values.slice(1)
                .filter(row => row && row.some(cell => String(cell || '').trim() !== ''))
                .map(row => {
                    const obj = {};
                    headers.forEach((h, idx) => {
                        obj[h] = (row[idx] !== undefined && row[idx] !== null) ? String(row[idx]).trim() : '';
                    });
                    return obj;
                });

            // Filter only Kwality House, Kemps Corner classes
            this.kwalityClasses = records.filter(record =>
                record.Location && record.Location.includes('Kwality House')
            );
            // Preserve full set for other location usage
            this.allSheetRecords = records;

            console.log(`✅ Found ${this.kwalityClasses.length} classes for Kwality House`);
            return this.kwalityClasses;
        } catch (error) {
            console.error('❌ Error reading Google Sheet:', error.response?.data || error.message);
            throw error;
        }
    }

    /**
     * Read schedule data directly from Google Sheets Cleaned sheet (replaces CSV reading)
     */
    async readCleanedSheet() {
        console.log('📋 Reading schedule data from Google Sheets Cleaned tab...');
        try {
            const accessToken = await this.getAccessToken();

            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });

            const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
            
            // Read from Cleaned sheet instead of CSV
            const range = 'Cleaned!A1:Z1000';
            const res = await sheets.spreadsheets.values.get({
                spreadsheetId: GOOGLE_CONFIG.SPREADSHEET_ID,
                range
            });

            const values = res.data.values || [];
            if (values.length === 0) {
                console.warn('⚠️  No data found in Cleaned sheet.');
                this.kwalityClasses = [];
                return this.kwalityClasses;
            }

            // First row is header
            const headers = values[0].map(h => String(h).trim());
            console.log(`📋 Headers found in Cleaned sheet: ${headers.join(', ')}`);
            
            const records = values.slice(1)
                .filter(row => row && row.some(cell => String(cell || '').trim() !== ''))
                .map(row => {
                    const obj = {};
                    headers.forEach((h, idx) => {
                        obj[h] = (row[idx] !== undefined && row[idx] !== null) ? String(row[idx]).trim() : '';
                    });
                    return obj;
                });

            // Filter only Kwality House, Kemps Corner classes for Kemps HTML
            this.kwalityClasses = records.filter(record =>
                record.Location && record.Location.includes('Kwality House')
            );
            
            // Store all records for other purposes
            this.allSheetRecords = records;

            console.log(`✅ Found ${this.kwalityClasses.length} classes for Kwality House from Cleaned sheet`);
            console.log(`✅ Total ${records.length} classes in Cleaned sheet`);
            return this.kwalityClasses;
        } catch (error) {
            console.error('❌ Error reading from Cleaned sheet:', error.response?.data || error.message);
            throw error;
        }
    }

    /**
     * Process emails to extract schedule data and update target spreadsheet
     */
    async processEmailAndUpdateSchedule() {
        console.log('📧 Starting email processing...');
        
        try {
            // Step 1: Find and fetch the latest schedule email
            console.log('🔍 Step 1: Finding latest schedule email...');
            const emailData = await this.findLatestScheduleEmail();
            if (!emailData) {
                console.log('⚠️  No schedule email found');
                return;
            }

            console.log('✅ Found email:', emailData.subject);
            console.log('📧 Email body preview:', emailData.body.substring(0, 200) + '...');

            // Step 2: Extract Google Sheets link from email
            console.log('🔗 Step 2: Extracting Google Sheets link...');
            const sheetsLink = this.extractSheetsLink(emailData.body);
            if (!sheetsLink) {
                console.log('⚠️  No Google Sheets link found in email');
                console.log('🔍 Email body search preview:', emailData.body.substring(0, 500));
                return;
            }

            console.log('✅ Found Google Sheets link:', sheetsLink);

            // Step 3: Extract schedule data from the linked spreadsheet
            console.log('📋 Step 3: Extracting data from linked spreadsheet...');
            const scheduleData = await this.fetchDataFromLinkedSheet(sheetsLink);
            
            if (!scheduleData || scheduleData.length === 0) {
                console.log('❌ No schedule data retrieved from linked sheet');
                return;
            }

            console.log(`✅ Retrieved ${scheduleData.length} schedule records from linked sheet`);
            
            // Step 4: Parse email for covers and themes
            console.log('🎨 Step 4: Parsing email for covers and themes...');
            const emailInfo = this.parseEmailForScheduleInfo(emailData.allMessages);
            
            console.log(`✅ Parsed ${emailInfo.covers.length} covers and ${emailInfo.themes.length} themes from email`);

            // Step 5: Update target spreadsheet with combined data
            console.log('📊 Step 5: Updating target spreadsheet...');
            await this.updateTargetSpreadsheet(scheduleData, emailInfo);
            
            console.log('✅ Email processing completed successfully');
            
        } catch (error) {
            console.error('❌ Error processing email:', error.message);
            console.error('🔍 Full error:', error);
            throw error;
        }
    }

    /**
     * Find the latest email from the specified sender with Schedule in subject
     */
    /**
     * Get the current week's date range for email search
     */
    getCurrentWeekDateRange() {
        const today = new Date();
        const currentDay = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
        
        // Calculate Monday of current week
        const monday = new Date(today);
        monday.setDate(today.getDate() - currentDay + 1);
        
        // Calculate Sunday of current week  
        const sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6);
        
        return { monday, sunday };
    }

    /**
     * Generate email subject pattern for current week
     */
    getCurrentWeekSubjectPattern() {
        const { monday, sunday } = this.getCurrentWeekDateRange();
        
        const startDay = monday.getDate();
        const endDay = sunday.getDate();
        const month = sunday.toLocaleDateString('en-GB', { month: 'short' });
        const year = sunday.getFullYear().toString().slice(-2);
        
        // Format: "1- 7th Dec '25" or "29 Nov - 5th Dec '25" (cross-month)
        if (monday.getMonth() === sunday.getMonth()) {
            // Same month
            return `${startDay}- ${endDay}${this.getOrdinalSuffix(endDay)} ${month} '${year}`;
        } else {
            // Cross month
            const startMonth = monday.toLocaleDateString('en-GB', { month: 'short' });
            return `${startDay} ${startMonth} - ${endDay}${this.getOrdinalSuffix(endDay)} ${month} '${year}`;
        }
    }

    /**
     * Get ordinal suffix for day (1st, 2nd, 3rd, 4th, etc.)
     */
    getOrdinalSuffix(day) {
        if (day >= 11 && day <= 13) return 'th';
        switch (day % 10) {
            case 1: return 'st';
            case 2: return 'nd'; 
            case 3: return 'rd';
            default: return 'th';
        }
    }

    async findLatestScheduleEmail() {
        const expectedSubject = this.getCurrentWeekSubjectPattern();
        console.log(`🔍 Searching for current week email with subject containing: "${expectedSubject}"`);
        console.log(`📅 Current week: ${this.getCurrentWeekDateRange().monday.toDateString()} - ${this.getCurrentWeekDateRange().sunday.toDateString()}`);
        
        try {
            const accessToken = await this.getAccessToken();
            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });

            const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
            
            // Search for current week's email with specific date pattern
            const specificQuery = `from:${EMAIL_CONFIG.SENDER_EMAIL} subject:"${EMAIL_CONFIG.SUBJECT_KEYWORD}" subject:"${expectedSubject}"`;
            console.log(`🔍 Specific email search query: ${specificQuery}`);
            
            let response = await gmail.users.messages.list({
                userId: 'me',
                q: specificQuery,
                maxResults: 1
            });
            
            if (!response.data.messages || response.data.messages.length === 0) {
                console.log('⚠️ No email found with exact current week date pattern');
                console.log('🔍 Searching with broader pattern...');
                
                // Fallback: search for any Mumbai schedule email from past week
                const fallbackQuery = `from:${EMAIL_CONFIG.SENDER_EMAIL} subject:"${EMAIL_CONFIG.SUBJECT_KEYWORD}" newer_than:7d`;
                console.log(`🔍 Fallback query: ${fallbackQuery}`);
                
                response = await gmail.users.messages.list({
                    userId: 'me',
                    q: fallbackQuery,
                    maxResults: 3
                });
                
                if (!response.data.messages || response.data.messages.length === 0) {
                    console.log('❌ No Mumbai schedule emails found in past 7 days');
                }
            }

            console.log(`📬 Found ${response.data.messages.length} potential emails`);
            
            // Get the most recent email (first in the list)
            const messageId = response.data.messages[0].id;
            console.log(`📬 Using email with ID: ${messageId}`);

            // Get the email thread
            const message = await gmail.users.messages.get({
                userId: 'me',
                id: messageId,
                format: 'full'
            });

            // Get all messages in the thread
            const threadId = message.data.threadId;
            const thread = await gmail.users.threads.get({
                userId: 'me',
                id: threadId,
                format: 'full'
            });

            // Extract email bodies from all messages in thread
            const allMessages = thread.data.messages.map(msg => this.extractEmailBody(msg));
            const latestMessage = this.extractEmailBody(message.data);
            const emailSubject = this.getHeader(message.data, 'Subject');
            const emailDate = this.getHeader(message.data, 'Date');

            console.log(`✅ Found email thread with ${thread.data.messages.length} messages`);
            console.log(`📧 Email subject: "${emailSubject}"`);
            console.log(`📅 Email date: ${emailDate}`);
            console.log(`📧 Email preview: ${latestMessage.substring(0, 200)}...`);
            
            return {
                body: latestMessage,
                allMessages: allMessages,
                subject: emailSubject,
                date: emailDate
            };

        } catch (error) {
            console.error('❌ Error searching for emails:', error);
            throw error;
        }
    }

    /**
     * Extract email body from Gmail message data
     */
    extractEmailBody(messageData) {
        let body = '';
        
        if (messageData.payload.body && messageData.payload.body.data) {
            body = Buffer.from(messageData.payload.body.data, 'base64').toString();
        } else if (messageData.payload.parts) {
            for (const part of messageData.payload.parts) {
                if (part.mimeType === 'text/plain' && part.body.data) {
                    body += Buffer.from(part.body.data, 'base64').toString();
                } else if (part.mimeType === 'text/html' && part.body.data) {
                    // Convert HTML to text (basic conversion)
                    const htmlBody = Buffer.from(part.body.data, 'base64').toString();
                    body += htmlBody.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
                }
            }
        }
        
        return body;
    }

    /**
     * Get header value from email message
     */
    getHeader(messageData, headerName) {
        const headers = messageData.payload.headers;
        const header = headers.find(h => h.name === headerName);
        return header ? header.value : '';
    }

    /**
     * Extract Google Sheets link from email body
     */
    extractSheetsLink(emailBody) {
        console.log('🔗 Extracting Google Sheets link from email...');
        
        // Look for Google Sheets URLs
        const sheetsRegex = /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/g;
        const matches = emailBody.match(sheetsRegex);
        
        if (matches && matches.length > 0) {
            console.log(`✅ Found Google Sheets link: ${matches[0]}`);
            return matches[0];
        }
        
        console.log('❌ No Google Sheets link found in email body');
        return null;
    }

    /**
     * Fetch data from linked Google Sheet
     */
    async fetchDataFromLinkedSheet(sheetsLink) {
        console.log('📊 Fetching data from linked spreadsheet...');
        
        try {
            // Extract spreadsheet ID from URL
            const match = sheetsLink.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
            if (!match) {
                throw new Error('Invalid spreadsheet URL');
            }
            
            const spreadsheetId = match[1];
            console.log(`📋 Spreadsheet ID: ${spreadsheetId}`);
            
            const accessToken = await this.getAccessToken();
            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });

            const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
            
            // Try to get data from 'Schedule' sheet
            const range = 'Schedule!A1:Z1000';
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: range
            });

            const values = response.data.values || [];
            if (values.length === 0) {
                console.log('⚠️  No data found in Schedule sheet');
                return [];
            }

            // Convert to objects
            const headers = values[0].map(h => String(h).trim());
            const records = values.slice(1).map(row => {
                const obj = {};
                headers.forEach((h, idx) => {
                    obj[h] = (row[idx] !== undefined && row[idx] !== null) ? String(row[idx]).trim() : '';
                });
                return obj;
            });

            console.log(`✅ Retrieved ${records.length} schedule records`);
            return records;
            
        } catch (error) {
            console.error('❌ Error fetching data from linked sheet:', error);
            throw error;
        }
    }

    /**
     * Parse email content for covers and themes information
     */
    parseEmailForScheduleInfo(allMessages) {
        console.log('🔍 Parsing email content for schedule information...');
        
        const result = {
            covers: [],
            themes: [],
            hostedClasses: []
        };
        
        // Combine all messages for parsing
        const fullContent = allMessages.join('\n\n');
        console.log('📧 Email content length:', fullContent.length);
        console.log('📧 First 500 chars:', fullContent.substring(0, 500));
        
        // Parse covers section
        const coversMatch = fullContent.match(/Covers\s*:?(.*?)(?=\n\n|\nAmped Up theme|Bandra cycle themes|FIT theme|Best,|$)/is);
        if (coversMatch) {
            console.log('🎯 Found covers section:', coversMatch[1].substring(0, 200));
            result.covers = this.parseCoversSection(coversMatch[1]);
        } else {
            console.log('❌ No covers section found');
            console.log('🔍 Looking for alternative covers pattern...');
            
            // Try alternative pattern
            const altCoversMatch = fullContent.match(/(Kemps\s*-.*?)(?=\n\nBandra|Bandra\s*-)/is);
            if (altCoversMatch) {
                console.log('🎯 Found alternative covers section:', altCoversMatch[1].substring(0, 200));
                result.covers = this.parseCoversSection(altCoversMatch[1]);
            }
        }
        
        // Parse themes sections - using simpler approach to avoid matchAll issues
        const themeSections = [];
        
        // Look for theme patterns individually with more precise boundaries
        const amped_theme = fullContent.match(/Amped Up theme\s*\*?\s*:\s*(.*?)(?=\n\*?Bandra cycle themes|\nBest,|$)/is);
        if (amped_theme) {
            // Clean up the captured text to remove any trailing content
            let ampedContent = amped_theme[1].trim();
            // Remove any text that starts with "Bandra cycle themes" or similar
            ampedContent = ampedContent.split(/Bandra cycle themes/i)[0].trim();
            themeSections.push({ type: 'Amped Up', content: ampedContent });
        }
        
        const bandra_themes = fullContent.match(/Bandra cycle themes\s*-\s*\*?\s*(.*?)(?=\nBest,|$)/is);
        if (bandra_themes) {
            themeSections.push({ type: 'Bandra cycle', content: bandra_themes[1] });
        }
        
        const fit_theme = fullContent.match(/FIT theme\s*:\s*All classes,\s*all week\s*-\s*(TABATA)/i);
        if (fit_theme) {
            themeSections.push({ type: 'FIT', content: `All classes, all week - ${fit_theme[1].trim()}` });
        }
        
        console.log(`📊 Found ${themeSections.length} theme sections`);
        
        for (const section of themeSections) {
            console.log(`🎨 Processing ${section.type} themes`);
            const themes = this.parseThemesSection(section.type, section.content);
            result.themes.push(...themes);
        }
        
        // Parse hosted classes
        const hostedMatch = fullContent.match(/-Hosted Classes\s*-(.*?)(?=\n\n|\n[A-Z]|$)/is);
        if (hostedMatch) {
            console.log('🏢 Found hosted classes section');
            result.hostedClasses = this.parseHostedClasses(hostedMatch[1]);
        }
        
        console.log(`✅ Parsed ${result.covers.length} covers, ${result.themes.length} themes, ${result.hostedClasses.length} hosted classes`);
        return result;
    }

    /**
     * Parse covers section from email
     */
    parseCoversSection(coversText) {
        console.log('🎯 Parsing covers section, length:', coversText.length);
        console.log('🎯 First 300 chars:', coversText.substring(0, 300));
        
        const covers = [];
        const lines = coversText.split('\n').filter(line => line.trim());
        
        let currentLocation = '';
        let previousDay = null;
        
        for (let i = 0; i < lines.length; i++) {
            const trimmedLine = lines[i].trim();
            console.log(`📝 Processing line ${i + 1}: "${trimmedLine}"`);
            
            if (!trimmedLine || trimmedLine === '*') continue;
            
            // Check if this is a location header
            if (this.isLocationHeader(trimmedLine)) {
                currentLocation = this.extractLocation(trimmedLine);
                console.log(`📍 Found location: ${currentLocation}`);
                previousDay = null; // Reset previous day when location changes
                continue;
            }
            
            // Parse cover entries
            const coverInfo = this.parseCoverLine(trimmedLine, currentLocation, previousDay);
            if (coverInfo) {
                console.log(`✅ Parsed cover:`, coverInfo);
                covers.push(coverInfo);
                // Update previousDay for potential continuation lines
                previousDay = coverInfo.day;
            } else {
                console.log(`❌ Could not parse cover line: "${trimmedLine}"`);
            }
        }
        
        console.log(`📊 Parsed ${covers.length} total covers`);
        return covers;
    }

    /**
     * Check if a line is a location header
     */
    isLocationHeader(line) {
        // Handle markdown-style headers like "*Kemps -*" or "Bandra -"
        const cleanLine = line.replace(/\*/g, '').trim().replace(/-$/, '').trim();
        const locationPattern = new RegExp(`^(${EMAIL_CONFIG.LOCATIONS.join('|')})\\s*-?\\s*$`, 'i');
        return locationPattern.test(cleanLine);
    }

    /**
     * Extract location from header line
     */
    extractLocation(line) {
        // Remove markdown formatting and cleanup
        const cleanLine = line.replace(/\*/g, '').trim().replace(/-$/, '').trim();
        for (const location of EMAIL_CONFIG.LOCATIONS) {
            if (cleanLine.toLowerCase().includes(location.toLowerCase())) {
                return location;
            }
        }
        return '';
    }

    /**
     * Parse individual cover line
     */
    parseCoverLine(line, location, previousDay = null) {
        // Pattern: Day - time(s) - trainer
        // Example: "Mon - 8,9.15, 11.30 am - Richard"
        const coverPattern = /^([A-Za-z]+)\s*-\s*(.*?)\s*-\s*(.+)$/;
        const match = line.match(coverPattern);
        
        if (match) {
            const day = match[1].trim();
            const timeText = match[2].trim();
            const trainer = match[3].trim();
            
            // Parse multiple times if present
            const timeInfo = this.parseTimeText(timeText);
            
            // Check if this is a pattern-based cover (morning cycles, evening barre)
            if (timeInfo.timePattern) {
                return {
                    location: location,
                    day: this.expandDayName(day),
                    timePattern: timeInfo.timePattern,
                    classType: timeInfo.classType,
                    trainer: trainer,
                    type: 'cover'
                };
            } else {
                // Regular time-based cover
                return {
                    location: location,
                    day: this.expandDayName(day),
                    times: Array.isArray(timeInfo) ? timeInfo : [timeInfo],
                    trainer: trainer,
                    type: 'cover'
                };
            }
        }
        
        // Try pattern for continuation lines (time - trainer)
        const continuationPattern = /^([\d\.,:\s]+\s*(?:am|pm|lab|B57|Barre|cycle))\s*-\s*(.+)$/i;
        const continuationMatch = line.match(continuationPattern);
        
        if (continuationMatch && previousDay) {
            const timeText = continuationMatch[1].trim();
            const trainer = continuationMatch[2].trim();
            
            const timeInfo = this.parseTimeText(timeText);
            
            return {
                location: location,
                day: previousDay,
                times: Array.isArray(timeInfo) ? timeInfo : [timeInfo],
                trainer: trainer,
                type: 'cover'
            };
        }
        
        return null;
    }

    /**
     * Parse time text that might contain multiple times
     */
    parseTimeText(timeText) {
        // Handle patterns like "8,9.15, 11.30 am" or "6,7.30 pm" or "Morning cycles"
        const times = [];
        
        // Handle descriptive times like "Morning cycles", "Evening Barre classes"
        if (/morning.*cycle/i.test(timeText)) {
            return {
                timePattern: 'morning',
                classType: 'CYCLE',
                description: timeText.trim()
            };
        }
        
        if (/evening.*barre/i.test(timeText)) {
            return {
                timePattern: 'evening', 
                classType: 'Barre',
                description: timeText.trim()
            };
        }
        
        // For other descriptive times, return as-is
        if (/morning|evening|afternoon/i.test(timeText)) {
            times.push(timeText.trim());
            return times;
        }
        
        // Extract AM/PM suffix
        const ampmMatch = timeText.match(/\b(am|pm)\b/i);
        const suffix = ampmMatch ? ampmMatch[1].toLowerCase() : '';
        
        // Split by commas and parse each time
        const timeSegments = timeText.split(',');
        
        for (let segment of timeSegments) {
            segment = segment.trim()
                .replace(/\b(am|pm)\b/i, '') // Remove AM/PM
                .replace(/lab|B57|Barre|cycle/gi, '') // Remove class type indicators
                .trim();
            
            if (segment) {
                // Convert . to : for time format
                let normalizedTime = segment.replace(/(\d+)\.(\d+)/, '$1:$2');
                
                // Add suffix if we have one
                if (suffix && normalizedTime.match(/^\d+:?\d*$/)) {
                    times.push(`${normalizedTime} ${suffix}`.trim());
                } else {
                    times.push(normalizedTime.trim());
                }
            }
        }
        
        return times;
    }

    /**
     * Expand abbreviated day names
     */
    expandDayName(day) {
        const dayMap = {
            'mon': 'Monday',
            'tue': 'Tuesday', 'tues': 'Tuesday',
            'wed': 'Wednesday',
            'thu': 'Thursday', 'thurs': 'Thursday',
            'fri': 'Friday',
            'sat': 'Saturday',
            'sun': 'Sunday'
        };
        
        return dayMap[day.toLowerCase()] || day;
    }

    /**
     * Parse themes section from email
     */
    parseThemesSection(themeType, themeContent) {
        const themes = [];
        const lines = themeContent.split('\n').filter(line => line.trim());
        
        console.log(`🎨 Parsing ${themeType} theme section with ${lines.length} lines`);
        
        for (const line of lines) {
            const trimmedLine = line.trim();
            if (!trimmedLine) continue;
            
            // Parse different theme patterns
            const themeInfo = this.parseThemeLine(trimmedLine, themeType);
            if (themeInfo) {
                themes.push(themeInfo);
                console.log(`✅ Parsed theme:`, themeInfo);
            }
        }
        
        return themes;
    }

    /**
     * Parse individual theme line
     */
    parseThemeLine(line, themeType) {
        console.log(`🔍 Parsing theme line: "${line}" for type: ${themeType}`);
        
        if (themeType === 'Amped Up') {
            // Pattern: "Tuesday - Icy Isometric"
            const pattern = /([A-Za-z]+)\s*-\s*(.+)/;
            const match = line.match(pattern);
            if (match) {
                return {
                    day: this.expandDayName(match[1]),
                    theme: match[2].trim(),
                    classType: 'Amped Up',
                    location: 'Kemps', // Amped Up is a Kemps class
                    type: 'theme'
                };
            }
        } else if (themeType === 'Bandra cycle') {
            // Pattern: "1. Monday 10 am - Taylor Swift vs Kendrick lamar" or "2. Tuesday 8am Retro -Future Nostalgia"
            // Handle special characters by cleaning them first
            const cleanedLine = line.replace(/[^\w\s\.:–-]/g, ' ').trim();
            
            // Try main pattern first (with dash)
            let pattern = /^\d+\.\s*([A-Za-z]+)\s+([\d:]+\s*[ap]m)\s*[-–]\s*(.+)$/i;
            let match = cleanedLine.match(pattern);
            
            if (!match) {
                // Try alternative pattern (without dash requirement)
                pattern = /^\d+\s*\.?\s*([A-Za-z]+)\s+(\d+\s*[ap]m|[\d:]+\s*[ap]m)\s+(.+)$/i;
                match = cleanedLine.match(pattern);
            }
            
            if (match) {
                return {
                    day: this.expandDayName(match[1]),
                    time: match[2].trim(),
                    theme: match[3].trim(),
                    classType: 'CYCLE',
                    location: 'Bandra', // Specifically Bandra cycles
                    type: 'theme'
                };
            }
        } else if (themeType === 'FIT') {
            // Pattern: "All classes, all week - TABATA"
            if (line.toLowerCase().includes('all classes') && line.toLowerCase().includes('all week')) {
                const themeMatch = line.match(/all week\s*[-–]\s*(.+)$/i);
                if (themeMatch) {
                    return {
                        day: 'All',
                        theme: themeMatch[1].trim(),
                        classType: 'FIT',
                        location: 'All',
                        type: 'theme'
                    };
                }
            }
        }
        
        // Fallback pattern for other formats
        const pattern = /^([A-Za-z]+)\s*[-–]\s*(.+)$/;
        const match = line.match(pattern);
        if (match) {
            return {
                day: this.expandDayName(match[1]),
                theme: match[2].trim(),
                themeType: themeType,
                type: 'theme'
            };
        }
        
        return null;
    }

    /**
     * Parse hosted classes section
     */
    parseHostedClasses(hostedText) {
        const hosted = [];
        const lines = hostedText.split('\n').filter(line => line.trim());
        
        for (const line of lines) {
            const trimmedLine = line.trim();
            if (!trimmedLine || trimmedLine.startsWith('-')) continue;
            
            const hostedInfo = this.parseHostedLine(trimmedLine);
            if (hostedInfo) {
                hosted.push(hostedInfo);
            }
        }
        
        return hosted;
    }

    /**
     * Parse individual hosted class line
     */
    parseHostedLine(line) {
        // Pattern: "Friday - Bay club BKC - cycle - 7.30 & 9 am - Vivaran"
        const pattern = /^([A-Za-z]+)\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*(.+)$/;
        const match = line.match(pattern);
        
        if (match) {
            return {
                day: this.expandDayName(match[1]),
                location: match[2].trim(),
                classType: match[3].trim(),
                time: match[4].trim(),
                trainer: match[5].trim(),
                type: 'hosted'
            };
        }
        
        return null;
    }

    /**
     * Update target spreadsheet with parsed schedule data
     */
    async updateTargetSpreadsheet(scheduleData, emailInfo) {
        console.log('📝 Updating target spreadsheet...');
        
        try {
            const accessToken = await this.getAccessToken();
            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });

            const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
            
            // Step 1: Copy complete sheet data from linked spreadsheet to target spreadsheet
            console.log('📋 Step 1: Copying complete sheet data from linked spreadsheet...');
            const copiedValues = await this.copyScheduleDataToTargetSheet(scheduleData, sheets);
            
            if (!copiedValues || copiedValues.length === 0) {
                console.log('❌ No data to copy from linked spreadsheet');
                return;
            }

            console.log(`✅ Successfully prepared ${copiedValues.length} rows from linked sheet`);

            // Step 2: First update the target sheet with the fresh data
            console.log('📊 Step 2: Writing fresh data to target spreadsheet...');
            await this.writeDataToTargetSheet(copiedValues, sheets);
            
            // Step 3: Now read the updated data back for theme/cover application
            console.log('📖 Step 3: Reading updated data for theme application...');
            const updatedRange = `${GOOGLE_CONFIG.TARGET_SHEET_NAME}!A1:ZZ${copiedValues.length}`;
            const updatedResponse = await sheets.spreadsheets.values.get({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                range: updatedRange
            });
            const currentValues = updatedResponse.data.values || [];
            
            // Step 4: Analyze the structure of the updated data
            const sheetStructure = this.analyzeSheetStructure(currentValues);
            console.log('📋 Sheet structure:', JSON.stringify(sheetStructure, null, 2));
            
            // Step 5: Apply covers and themes to the freshly copied data
            console.log('🎨 Step 5: Applying email themes and covers to copied data...');
            const finalValues = this.applyEmailDataToSheet(currentValues, emailInfo, sheetStructure);
            
            // Step 6: Update the sheet with the final modified data
            if (finalValues.length > 0) {
                console.log('📝 Step 6: Writing final data with themes and covers...');
                await this.writeDataToTargetSheet(finalValues, sheets);
            }
            
            // Step 7: Clean the updated data and populate the Cleaned sheet
            console.log('🧹 Step 7: Cleaning data and populating Cleaned sheet...');
            await this.cleanAndPopulateCleanedSheet(sheets);
            
            console.log(`✅ Target spreadsheet updated with fresh data and email themes/covers applied`);
            
        } catch (error) {
            console.error('❌ Error updating target spreadsheet:', error);
            throw error;
        }
    }

    /**
     * Copy schedule data from linked spreadsheet to target spreadsheet
     * This completely replaces the target sheet data with fresh data from linked sheet
     */
    async copyScheduleDataToTargetSheet(scheduleData, sheets) {
        console.log(`📋 Copying schedule data from linked spreadsheet...`);
        
        try {
            // Get the raw data from linked sheet to maintain exact structure
            console.log('📥 Fetching raw data from linked sheet to maintain structure...');
            const linkedSheetData = await this.getRawDataFromLinkedSheet(sheets);
            
            if (!linkedSheetData || linkedSheetData.length === 0) {
                console.log('❌ No raw data available from linked sheet');
                return null;
            }

            console.log(`📊 Retrieved ${linkedSheetData.length} rows from linked sheet`);
            console.log('🔍 Sample headers:', linkedSheetData[0]?.slice(0, 10));
            
            // Clean and format the data (especially time columns)
            const cleanedData = this.cleanSheetData(linkedSheetData);
            
            console.log(`✅ Prepared ${cleanedData.length} rows for target sheet`);
            
            return cleanedData;
            
        } catch (error) {
            console.error('❌ Error copying schedule data:', error);
            throw error;
        }
    }

    /**
     * Get raw data from linked sheet maintaining exact structure
     */
    async getRawDataFromLinkedSheet(sheets) {
        console.log('📥 Fetching raw data from linked sheet...');
        
        try {
            // Get the linked sheet ID from the last processed email
            const emailData = await this.findLatestScheduleEmail();
            if (!emailData) {
                console.log('❌ No email data to extract sheet ID from');
                return null;
            }

            const sheetsLink = this.extractSheetsLink(emailData.body);
            if (!sheetsLink) {
                console.log('❌ No sheets link found in email');
                return null;
            }

            const spreadsheetId = this.extractSpreadsheetId(sheetsLink);
            
            // Get all data from the Schedule sheet in the linked spreadsheet
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: 'Schedule!A:ZZ' // Get all data from Schedule sheet
            });
            
            const values = response.data.values || [];
            console.log(`📊 Retrieved ${values.length} rows from linked Schedule sheet`);
            
            return values;
            
        } catch (error) {
            console.error('❌ Error fetching raw data from linked sheet:', error);
            throw error;
        }
    }

    /**
     * Extract spreadsheet ID from Google Sheets URL
     */
    extractSpreadsheetId(url) {
        const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!match) {
            throw new Error('Invalid Google Sheets URL format');
        }
        return match[1];
    }

    /**
     * Normalize class names specifically for the cleaned sheet output
     * This uses the comprehensive mapping provided by the user
     */
    normalizeClassNameForCleaned(raw) {
        if (!raw) return '';
        const val = raw.toString().trim().replace(/\s+/g, ' ').toLowerCase();
        const map = {
            // Direct mappings to new "Studio" format
            'hosted class': 'Studio Hosted Class',
            'fit': 'Studio FIT',
            'back body blaze': 'Studio Back Body Blaze',
            'bbb': 'Studio Back Body Blaze',
            'barre 57': 'Studio Barre 57',
            'barre57': 'Studio Barre 57',
            'mat 57': 'Studio Mat 57',
            'mat57': 'Studio Mat 57',
            "trainer's choice": "Studio Trainer's Choice",
            'amped up': 'Studio Amped Up!',
            'amped up!': 'Studio Amped Up!',
            'hiit': 'Studio HIIT',
            'foundations': 'Studio Foundations',
            'sweat in 30': 'Studio SWEAT In 30',
            'sweat': 'Studio SWEAT In 30',
            'cardio barre plus': 'Studio Cardio Barre Plus',
            'cardio b+': 'Studio Cardio Barre Plus',
            'cardio barre': 'Studio Cardio Barre',
            'cardio b': 'Studio Cardio Barre',
            'recovery': 'Studio Recovery',
            'pre/post natal': 'Studio Pre/Post Natal',
            'prenatal': 'Studio Pre/Post Natal',
            'cycle': 'Studio PowerCycle',
            'powercycle': 'Studio PowerCycle',
            'strength lab': 'Studio Strength Lab',
            'strength lab (full body)': 'Studio Strength Lab',
            'strength (pull)': 'Studio Strength Lab (Pull)',
            'strength (push)': 'Studio Strength Lab (Push)',
            'strength - fb': 'Studio Strength Lab (Full Body)',
            'strength - pull': 'Studio Strength Lab (Pull)',
            'strength - push': 'Studio Strength Lab (Push)',
            // Express versions
            'cardio barre express': 'Studio Cardio Barre Express',
            'cardio barre exp': 'Studio Cardio Barre Express',
            'cardio b exp': 'Studio Cardio Barre Express',
            'barre 57 express': 'Studio Barre 57 Express',
            'barre 57 exp': 'Studio Barre 57 Express',
            'barre57 exp': 'Studio Barre 57 Express',
            'back body blaze express': 'Studio Back Body Blaze Express',
            'bbb exp': 'Studio Back Body Blaze Express',
            'mat 57 express': 'Studio Mat 57 Express',
            'mat 57 exp': 'Studio Mat 57 Express',
            'mat57 exp': 'Studio Mat 57 Express',
            'cycle exp': 'Studio PowerCycle Express',
            'powercycle express': 'Studio PowerCycle Express',
        };
        if (map[val]) return map[val];
        return raw.toString().trim().replace(/\b\w/g, c => c.toUpperCase());
    }

    /**
     * Clean and normalize schedule data and populate the Cleaned sheet
     */
    async cleanAndPopulateCleanedSheet(sheets) {
        console.log('🧹 Cleaning schedule data and creating Cleaned sheet...');
        
        try {
            // Get the current data from the Schedule sheet
            const scheduleRange = `${GOOGLE_CONFIG.TARGET_SHEET_NAME}!A:ZZ`;
            const scheduleResponse = await sheets.spreadsheets.values.get({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                range: scheduleRange
            });
            
            const rows = scheduleResponse.data.values || [];
            
            if (rows.length < 5) {
                console.log('❌ Schedule sheet must have at least 5 rows (header structure required)');
                return;
            }

            const dayRow = rows[2] || [];    // Row 3 has days
            const headerRow = rows[3] || []; // Row 4 has headers
            const dateRow = rows[1] || [];   // Row 2 has dates
            const dataRows = rows.slice(4);
            
            console.log('📋 Processing schedule data for cleaning...');
            console.log(`🔢 Found ${dataRows.length} data rows to process`);
            console.log('📅 Date row (row 2):', dateRow.slice(0, 10));
            
            // Define column mappings based on your Google Apps Script
            const locationCols = [1, 7, 13, 18, 23, 28, 34];
            const dayCols = locationCols;
            const classCols = [2, 8, 14, 19, 24, 29, 35];
            const trainer1Cols = [3, 9, 15, 20, 25, 30, 36];
            const trainer2Cols = [4, 10, 16, 21, 26, 31, 37]; // For themes
            const coverCols = [6, 12, 17, 22, 27, 32, 38];

            // Find time column
            const timeColIndex = headerRow.findIndex(h => 
                (h || '').toString().trim().toLowerCase() === 'time'
            );
            
            if (timeColIndex === -1) {
                console.log('❌ Time column header not found in row 4');
                return;
            }

            const daysOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
            const allClasses = [];

            // Process each data row
            for (let r = 0; r < dataRows.length; r++) {
                const row = dataRows[r];
                
                // Process each location column set
                for (let setIdx = 0; setIdx < locationCols.length; setIdx++) {
                    const location = this.normalizeLocationName(row[locationCols[setIdx]]);
                    if (!location) continue;

                    const dayRaw = dayRow[dayCols[setIdx]] || 'Unknown';
                    const day = daysOrder.find(d => 
                        d.toLowerCase() === dayRaw.toString().toLowerCase()
                    ) || dayRaw;

                    const classNameRaw = row[classCols[setIdx]];
                    const className = this.normalizeClassNameForCleaned(classNameRaw);
                    if (!className || !this.isValidClassName(className)) continue;

                    const trainerRaw = row[trainer1Cols[setIdx]];
                    const coverRaw = row[coverCols[setIdx]];
                    const themeRaw = row[trainer2Cols[setIdx]]; // Theme from Trainer 2 column
                    
                    let trainer = this.normalizeTrainerName(trainerRaw);
                    let notes = '';
                    
                    // Handle covers
                    if (coverRaw) {
                        const coverNorm = this.normalizeTrainerName(coverRaw);
                        notes = coverNorm ? `Cover: ${coverNorm}` : 'Cover noted';
                        if (coverNorm) trainer = coverNorm;
                    }
                    
                    // Exclude classes without a trainer
                    if (!trainer) continue;

                    const timeRaw = row[timeColIndex];
                    const timeDate = this.parseTimeToDate(timeRaw);
                    const time = timeDate ? this.formatTime(timeDate) : timeRaw;
                    
                    // Get actual date from row 2 (same column as location)
                    const rawDate = dateRow[locationCols[setIdx]];
                    const date = rawDate && rawDate.toString().trim() ? 
                        this.formatDateFromSheet(rawDate.toString().trim()) : 
                        this.getDateForDay(day); // fallback to calculated date
                    
                    // Add theme to notes if present and not a trainer name
                    let theme = '';
                    if (themeRaw && themeRaw.toString().trim()) {
                        const themeValue = themeRaw.toString().trim();
                        // Only use as theme if it's not a trainer name
                        if (!this.isTrainerName(themeValue)) {
                            theme = themeValue;
                        }
                    }

                    allClasses.push({
                        Day: day,
                        Time: time,
                        Location: location,
                        Class: className,
                        Trainer: trainer,
                        Notes: notes,
                        Date: date,
                        Theme: theme
                    });
                }
            }

            // Sort by day and then by time
            allClasses.sort((a, b) => {
                const dayDiff = daysOrder.indexOf(a.Day) - daysOrder.indexOf(b.Day);
                if (dayDiff !== 0) return dayDiff;
                const aDate = this.parseTimeToDate(a.Time);
                const bDate = this.parseTimeToDate(b.Time);
                return (aDate && bDate) ? aDate - bDate : 0;
            });

            console.log(`✅ Processed ${allClasses.length} valid classes`);

            // Write cleaned data to Cleaned sheet
            const headers = ['Day', 'Time', 'Location', 'Class', 'Trainer', 'Notes', 'Date', 'Theme'];
            const values = [
                headers, 
                ...allClasses.map(obj => headers.map(h => obj[h] || ''))
            ];

            // Clear and write to Cleaned sheet
            await this.writeDataToSheet('Cleaned', values, sheets);
            
            console.log(`✅ Successfully populated Cleaned sheet with ${allClasses.length} classes`);
            
        } catch (error) {
            console.error('❌ Error cleaning and populating Cleaned sheet:', error);
            throw error;
        }
    }

    /**
     * Write data to a specific sheet (creates if doesn't exist)
     */
    async writeDataToSheet(sheetName, data, sheets) {
        console.log(`📝 Writing ${data.length} rows to ${sheetName} sheet...`);
        
        try {
            // Calculate range
            const maxCols = Math.max(...data.map(row => row.length));
            const endCol = this.numberToColumnLetter(maxCols);
            const range = `${sheetName}!A1:${endCol}${data.length}`;
            
            // Clear the sheet first
            await sheets.spreadsheets.values.clear({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                range: `${sheetName}!A:ZZ`
            });
            
            // Write data
            await sheets.spreadsheets.values.update({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                range: range,
                valueInputOption: 'RAW',
                resource: {
                    values: data
                }
            });
            
            console.log(`✅ Successfully wrote data to ${sheetName} sheet`);
            
        } catch (error) {
            console.error(`❌ Error writing data to ${sheetName} sheet:`, error);
            throw error;
        }
    }

    /**
     * Convert column number to Excel column letter
     */
    numberToColumnLetter(num) {
        let result = '';
        while (num > 0) {
            num--; 
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26);
        }
        return result;
    }

    /**
     * Get date string for a given day of the week (format: DD-MMM-YYYY)
     */
    getDateForDay(dayName) {
        const daysOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        const today = new Date();
        const currentDay = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
        const targetDayIndex = daysOrder.indexOf(dayName);
        
        if (targetDayIndex === -1) return today.toLocaleDateString('en-GB');
        
        // Calculate days to add to get to target day
        const mondayIndex = 1;
        const daysToAdd = (targetDayIndex + mondayIndex - currentDay + 7) % 7;
        
        const targetDate = new Date(today);
        targetDate.setDate(today.getDate() + daysToAdd);
        
        // Format as DD-MMM-YYYY
        const day = targetDate.getDate().toString().padStart(2, '0');
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
            'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const month = monthNames[targetDate.getMonth()];
        const year = targetDate.getFullYear();
        
        return `${day}-${month}-${year}`;
    }

    /**
     * Normalize location names
     */
    normalizeLocationName(raw) {
        if (!raw) return '';
        const val = raw.toString().trim().toLowerCase();
        const map = {
            'kemps': 'Kwality House, Kemps Corner',
            'kemps corner': 'Kwality House, Kemps Corner',
            'bandra': 'Supreme HQ, Bandra',
            'kenkere': 'Kenkere House',
            'south united': 'South United Football Club',
            'copper cloves': 'The Studio by Copper + Cloves',
            'wework galaxy': 'WeWork Galaxy',
            'wework prestige': 'WeWork Prestige Central',
            'physique': 'Physique Outdoor Pop-up',
            'annex': 'Kwality House, Kemps Corner',
        };
        for (const key in map) {
            if (val.includes(key)) return map[key];
        }
        return raw.toString().trim();
    }

    /**
     * Normalize class names
     */
    normalizeClassName(raw) {
        if (!raw) return '';
        const val = raw.toString().trim().replace(/\s+/g, ' ').toLowerCase();
        const map = {
            // Direct mappings to new "Studio" format
        'hosted class': 'Studio Hosted Class',
        'fit': 'Studio FIT',
        'back body blaze': 'Studio Back Body Blaze',
        'bbb': 'Studio Back Body Blaze',
        'barre 57': 'Studio Barre 57',
        'barre57': 'Studio Barre 57',
        'mat 57': 'Studio Mat 57',
        'mat57': 'Studio Mat 57',
        "trainer's choice": "Studio Trainer's Choice",
        'amped up': 'Studio Amped Up!',
        'amped up!': 'Studio Amped Up!',
        'hiit': 'Studio HIIT',
        'foundations': 'Studio Foundations',
        'sweat in 30': 'Studio SWEAT In 30',
        'sweat': 'Studio SWEAT In 30',
        'cardio barre plus': 'Studio Cardio Barre Plus',
        'cardio b+': 'Studio Cardio Barre Plus',
        'cardio barre': 'Studio Cardio Barre',
        'cardio b': 'Studio Cardio Barre',
        'recovery': 'Studio Recovery',
        'pre/post natal': 'Studio Pre/Post Natal',
        'prenatal': 'Studio Pre/Post Natal',
        'cycle': 'Studio PowerCycle',
        'powercycle': 'Studio PowerCycle',
        'strength lab': 'Studio Strength Lab',
        'strength lab (full body)': 'Studio Strength Lab',
        'strength (pull)': 'Studio Strength Lab (Pull)',
        'strength (push)': 'Studio Strength Lab (Push)',
        'strength - fb': 'Studio Strength Lab (Full Body)',
        'strength - pull': 'Studio Strength Lab (Pull)',
        'strength - push': 'Studio Strength Lab (Push)',

        // Express versions
        'cardio barre express': 'Studio Cardio Barre Express',
        'cardio barre exp': 'Studio Cardio Barre Express',
        'cardio b exp': 'Studio Cardio Barre Express',
        'barre 57 express': 'Studio Barre 57 Express',
        'barre 57 exp': 'Studio Barre 57 Express',
        'barre57 exp': 'Studio Barre 57 Express',
        'back body blaze express': 'Studio Back Body Blaze Express',
        'bbb exp': 'Studio Back Body Blaze Express',
        'mat 57 express': 'Studio Mat 57 Express',
        'mat 57 exp': 'Studio Mat 57 Express',
        'mat57 exp': 'Studio Mat 57 Express',
        'cycle exp': 'Studio PowerCycle Express',
        'powercycle express': 'Studio PowerCycle Express',
        };
        if (map[val]) return map[val];
        return raw.toString().trim().replace(/\b\w/g, c => c.toUpperCase());
    }

    /**
     * Check if a value is a trainer name
     */
    isTrainerName(value) {
        if (!value) return false;
        const val = value.toString().trim().toLowerCase();
        
        // Check against mapping first
        const map = {
            'mriga': 'Mrigakshi Jaiswal',
            'nishant': 'Nishanth Raj',
            'raunaq': 'Raunak Khemuka',
            'richy': "Richard D'Costa"
        };
        if (map[val]) return true;
        
        const teachers = [
            'Anisha Shah', 'Atulan Purohit', 'Janhavi Jain', 'Karanvir Bhatia', 'Mrigakshi Jaiswal',
            'Pranjali Jain', 'Reshma Sharma', "Richard D'Costa", 'Rohan Dahima', 'Upasna Paranjpe',
            'Karan Bhatia', 'Saniya Jaiswal', 'Vivaran Dhasmana', 'Nishanth Raj', 'Cauveri Vikrant',
            'Kabir Varma', 'Simonelle De Vitre', 'Simran Dutt', 'Anmol Sharma', 'Bret Saldanha',
            'Raunak Khemuka', 'Kajol Kanchan', 'Pushyank Nahar', 'Shruti Kulkarni',
            'Shruti Suresh', 'Poojitha Bhaskar', 'Siddhartha Kusuma', 'Chaitanya Nahar', 'Veena Narasimhan',
            // Short names for compatibility
            'Rohan', 'Anisha', 'Richard', 'Pranjali', 'Reshma', 'Atulan', 'Karanvir', 'Cauveri',
            'Mrigakshi', 'Vivaran', 'Karan', 'Nishanth', 'Pushyank', 'Kajol', 'Siddhartha', 'Shruti K', 'Veena', 'Chaitanya', 'Raunak'
        ];
        
        for (const t of teachers) {
            const low = t.toLowerCase();
            if (low === val || low.startsWith(val + ' ') || val.startsWith(low.split(' ')[0])) {
                return true;
            }
        }
        return false;
    }

    /**
     * Normalize trainer names
     */
    normalizeTrainerName(raw) {
        if (!raw) return '';
        const val = raw.toString().trim().toLowerCase();
        const map = {
            'mriga': 'Mrigakshi Jaiswal',
            'nishant': 'Nishanth Raj',
            'raunaq': 'Raunak Khemuka',
            'richy': "Richard D'Costa"
        };
        if (map[val]) return map[val];
        
        const teachers = [
            'Anisha Shah', 'Atulan Purohit', 'Janhavi Jain', 'Karanvir Bhatia', 'Mrigakshi Jaiswal',
            'Pranjali Jain', 'Reshma Sharma', "Richard D'Costa", 'Rohan Dahima', 'Upasna Paranjpe',
            'Karan Bhatia', 'Saniya Jaiswal', 'Vivaran Dhasmana', 'Nishanth Raj', 'Cauveri Vikrant',
            'Kabir Varma', 'Simonelle De Vitre', 'Simran Dutt', 'Anmol Sharma', 'Bret Saldanha',
            'Raunak Khemuka', 'Kajol Kanchan', 'Pushyank Nahar', 'Shruti Kulkarni',
            'Shruti Suresh', 'Poojitha Bhaskar', 'Siddhartha Kusuma', 'Chaitanya Nahar', 'Veena Narasimhan',
            // Short names for compatibility
            'Rohan', 'Anisha', 'Richard', 'Pranjali', 'Reshma', 'Atulan', 'Karanvir', 'Cauveri',
            'Mrigakshi', 'Vivaran', 'Karan', 'Nishanth', 'Pushyank', 'Kajol', 'Siddhartha', 'Shruti K', 'Veena', 'Chaitanya', 'Raunak'
        ];
        
        for (const t of teachers) {
            const low = t.toLowerCase();
            if (low === val || low.startsWith(val + ' ')) return t;
        }
        return raw.toString().trim();
    }

    /**
     * Check if class name is valid
     */
    isValidClassName(name) {
        if (!name) return false;
        const val = name.toString().trim().toLowerCase();
        const invalid = ['smita', 'anandita', 'host', 'cover', 'replacement', 'sakshi', 'parekh', 'taarika'];
        if (invalid.some(i => val.includes(i))) return false;
        if (/^\d+$/.test(val)) return false;
        if (val.split(' ').length === 1 && val.length < 3) return false;
        return true;
    }

    /**
     * Parse time string to Date object
     */
    parseTimeToDate(timeStr) {
        if (!timeStr) return null;
        const t = this.normalizeTimeString(timeStr);
        const match = t.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
        if (!match) return null;
        let h = parseInt(match[1]);
        const m = parseInt(match[2]);
        const ampm = match[3].toUpperCase();
        if (ampm === 'PM' && h !== 12) h += 12;
        if (ampm === 'AM' && h === 12) h = 0;
        const d = new Date();
        d.setHours(h, m, 0, 0);
        return d;
    }

    /**
     * Normalize time string format
     */
    normalizeTimeString(timeStr) {
        if (!timeStr) return '';
        let t = timeStr.toString().trim().replace(/\s*[:,.]\s*/g, ':');
        t = t.replace(/(\d)(AM|PM)/gi, '$1 $2').toUpperCase();
        return t;
    }

    /**
     * Format Date object to time string
     */
    formatTime(date) {
        if (!date) return '';
        let h = date.getHours();
        const m = date.getMinutes();
        const ampm = h >= 12 ? 'PM' : 'AM';
        h = h % 12 || 12;
        const mm = m < 10 ? '0' + m : m;
        return `${h}:${mm} ${ampm}`;
    }

    /**
     * Clean and format sheet data, especially time values
     */
    cleanSheetData(sheetData) {
        console.log('🧹 Cleaning and formatting sheet data...');
        
        if (!sheetData || sheetData.length === 0) return [];
        
        const cleanedData = sheetData.map((row, rowIndex) => {
            return row.map((cell, colIndex) => {
                // Clean time values in first column (typical time column)
                if (colIndex === 0) {
                    return this.cleanAndFormatTime(cell);
                }
                
                // Clean other cells
                return typeof cell === 'string' ? cell.trim() : (cell || '');
            });
        });
        
        console.log(`✅ Cleaned ${cleanedData.length} rows of data`);
        return cleanedData;
    }

    /**
     * Clean and format time values like the Google Apps Script version
     */
    cleanAndFormatTime(timeStr) {
        if (!timeStr || String(timeStr).trim() === '') return '';

        // Convert to string and remove extra spaces
        let cleaned = String(timeStr).trim();

        // Replace common issues: commas, dots (except in time), extra spaces
        cleaned = cleaned.replace(/,/g, ':');  // Replace comma with colon
        cleaned = cleaned.replace(/\s+/g, ' '); // Replace multiple spaces with single space

        // Extract time pattern: handles formats like "7:00 AM", "7.00 AM", "700 AM", "7 AM"
        const timePattern = /(\d{1,2})[\s:.-](\d{0,2})\s(AM|PM|am|pm)/i;
        const match = cleaned.match(timePattern);

        if (match) {
            let hours = parseInt(match[1]);
            let minutes = match[2] ? match[2].padStart(2, '0') : '00';
            let period = match[3].toUpperCase();

            // Validate hours and minutes
            if (hours < 1 || hours > 12) return cleaned; // Invalid hour, return original
            if (parseInt(minutes) > 59) return cleaned; // Invalid minutes, return original

            // Format as HH:MM AM/PM
            return hours + ':' + minutes + ' ' + period;
        }

        // If no match, return original cleaned value
        return cleaned;
    }

    /**
     * Write data to target spreadsheet
     */
    async writeDataToTargetSheet(data, sheets) {
        console.log(`📝 Writing ${data.length} rows to target spreadsheet...`);
        
        try {
            // Calculate the range needed
            const maxCols = Math.max(...data.map(row => row.length));
            
            // Convert column number to Excel column letter
            function numberToColumnLetter(num) {
                let result = '';
                while (num > 0) {
                    num--; 
                    result = String.fromCharCode(65 + (num % 26)) + result;
                    num = Math.floor(num / 26);
                }
                return result;
            }
            
            const endCol = numberToColumnLetter(maxCols);
            const updateRange = `${GOOGLE_CONFIG.TARGET_SHEET_NAME}!A1:${endCol}${data.length}`;
            
            console.log(`📊 Writing to range: ${updateRange} (${data.length} rows, ${maxCols} columns)`);
            
            // Clear the sheet first
            await sheets.spreadsheets.values.clear({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                range: `${GOOGLE_CONFIG.TARGET_SHEET_NAME}!A:ZZ`
            });
            
            // Write the new data
            await sheets.spreadsheets.values.update({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                range: updateRange,
                valueInputOption: 'RAW',
                resource: {
                    values: data
                }
            });
            
            console.log(`✅ Successfully wrote data to target spreadsheet`);
            
        } catch (error) {
            console.error('❌ Error writing data to target sheet:', error);
            throw error;
        }
    }

    /**
     * Create a basic sheet structure when target sheet is empty
     */
    createBasicSheetStructure(scheduleData) {
        console.log('🏗️ Creating basic sheet structure from schedule data...');
        
        if (!scheduleData || scheduleData.length === 0) {
            console.log('❌ No schedule data to create structure from');
            return null;
        }

        // Get headers from the first record
        const headers = Object.keys(scheduleData[0]);
        console.log('📝 Available headers from schedule data:', headers.slice(0, 10));
        
        // Create basic structure rows
        const row1 = new Array(headers.length).fill(''); // Empty row 1
        const row2 = new Array(headers.length).fill(''); // Empty row 2
        const row3 = new Array(headers.length).fill(''); // Days row (will be filled based on headers)
        const row4 = headers.slice(); // Headers row
        
        // Fill row 3 with day names and row 2 with dates based on headers
        headers.forEach((header, index) => {
            const dayMatch = header.match(/(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)/i);
            if (dayMatch) {
                const dayName = dayMatch[1];
                row3[index] = dayName.toUpperCase();
                
                // Add corresponding date to row 2 
                const date = this.getDateForDay(dayName);
                row2[index] = date;
            }
        });
        
        // Convert schedule data to rows
        const dataRows = scheduleData.map(record => {
            return headers.map(header => record[header] || '');
        });
        
        const finalValues = [row1, row2, row3, row4, ...dataRows];
        
        console.log(`✅ Created basic structure with ${finalValues.length} total rows`);
        console.log('📅 Created row 2 (dates):', row2.slice(0, 10));
        console.log('📅 Created row 3 (days):', row3.slice(0, 10));
        console.log('📅 Created row 4 (headers):', row4.slice(0, 10));
        
        return finalValues;
    }

    /**
     * Convert record objects back to sheet format matching target structure
     */
    convertRecordsToSheetFormat(records, structureRows) {
        console.log('🔄 Converting records to sheet format...');
        
        if (!structureRows || structureRows.length < 4) {
            console.log('❌ Invalid structure rows');
            return [];
        }

        // Get the header structure from row 4 (index 3)
        const targetHeaders = structureRows[3] || [];
        console.log('📝 Target sheet headers (first 10):', targetHeaders.slice(0, 10));
        
        // Map each record to match the target sheet structure
        const dataRows = records.map(record => {
            const row = new Array(targetHeaders.length).fill('');
            
            // Map record properties to correct columns
            Object.keys(record).forEach(key => {
                const headerIndex = targetHeaders.findIndex(header => 
                    header && header.toString().trim().toLowerCase() === key.toLowerCase()
                );
                
                if (headerIndex !== -1) {
                    row[headerIndex] = record[key] || '';
                }
            });
            
            return row;
        });
        
        console.log(`✅ Converted ${dataRows.length} records to sheet format`);
        return dataRows;
    }

    /**
     * Analyze the complex multi-day spreadsheet structure
     */
    analyzeSheetStructure(values) {
        const structure = {
            dayColumns: {},
            coverColumns: {},
            trainer2Columns: {},
            headerRows: 4 // Headers are in row 4 (index 3)
        };

        // Row 3 (index 2) has day names, row 4 (index 3) has headers
        const dayRow = values[2] || [];    // Row 3 has days
        const headerRow = values[3] || []; // Row 4 has headers like "Trainer 2", "Cover"
        
        console.log('📅 Day row (first 20):', dayRow.slice(0, 20));
        console.log('📅 Header row (first 20):', headerRow.slice(0, 20));
        
        // DEBUG: Show ALL headers with their indices
        console.log('🔍 All headers:');
        for (let i = 0; i < Math.min(headerRow.length, 45); i++) {
            const header = String(headerRow[i] || '').trim();
            if (header) {
                console.log(`  Col ${i}: "${header}"`);
            }
        }

        // Find day columns and their associated data columns
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        
        for (let colIndex = 0; colIndex < dayRow.length; colIndex++) {
            const cellValue = String(dayRow[colIndex] || '').trim();
            
            for (const day of days) {
                if (cellValue.toLowerCase().includes(day.toLowerCase())) {
                    if (!structure.dayColumns[day]) {
                        structure.dayColumns[day] = [];
                    }
                    
                    // Use known pattern-based mapping based on the header structure we observed
                    // From debug: Col 0: Time, Col 1: Location, Col 2: Class, Col 3: Trainer 1, Col 4: Trainer 2, Col 6: Cover
                    // Then: Col 7: Location, Col 8: Class, Col 9: Trainer 1, Col 11: Trainer 2, Col 12: Cover, etc.
                    
                    let locationCol = -1;
                    let classCol = -1;
                    let trainer1Col = -1;
                    let trainer2Col = -1;
                    let coverCol = -1;
                    
                    if (day === 'Monday' && colIndex === 1) {
                        locationCol = 1; classCol = 2; trainer1Col = 3; trainer2Col = 4; coverCol = 6;
                    } else if (day === 'Tuesday' && colIndex === 7) {
                        locationCol = 7; classCol = 8; trainer1Col = 9; trainer2Col = 11; coverCol = 12;
                    } else if (day === 'Wednesday' && colIndex === 13) {
                        locationCol = 13; classCol = 14; trainer1Col = 15; trainer2Col = 16; coverCol = 17;
                    } else if (day === 'Thursday' && colIndex === 18) {
                        locationCol = 18; classCol = 19; trainer1Col = 20; trainer2Col = 21; coverCol = 22;
                    } else if (day === 'Friday' && colIndex === 23) {
                        locationCol = 23; classCol = 24; trainer1Col = 25; trainer2Col = 26; coverCol = 27;
                    } else if (day === 'Saturday' && colIndex === 28) {
                        locationCol = 28; classCol = 29; trainer1Col = 30; trainer2Col = 31; coverCol = 32;
                    } else if (day === 'Sunday' && colIndex === 34) {
                        locationCol = 34; classCol = 35; trainer1Col = 36; trainer2Col = 37; coverCol = 38;
                    } else {
                        // Fallback: search for headers if the pattern doesn't match
                        for (let searchCol = colIndex - 6; searchCol <= colIndex + 6; searchCol++) {
                            if (searchCol >= 0 && searchCol < headerRow.length) {
                                const headerValue = String(headerRow[searchCol] || '').toLowerCase().trim();
                                
                                if (headerValue === 'location' && locationCol === -1 && Math.abs(searchCol - colIndex) <= 6) {
                                    locationCol = searchCol;
                                }
                                if (headerValue === 'class' && classCol === -1 && Math.abs(searchCol - colIndex) <= 6) {
                                    classCol = searchCol;
                                }
                                if (headerValue === 'trainer 1' && trainer1Col === -1 && Math.abs(searchCol - colIndex) <= 6) {
                                    trainer1Col = searchCol;
                                }
                                if (headerValue === 'trainer 2' && trainer2Col === -1 && Math.abs(searchCol - colIndex) <= 6) {
                                    trainer2Col = searchCol;
                                }
                                if (headerValue === 'cover' && coverCol === -1 && Math.abs(searchCol - colIndex) <= 6) {
                                    coverCol = searchCol;
                                }
                            }
                        }
                    }
                    
                    
                    // Ensure we have valid column mappings
                    if (locationCol !== -1 && classCol !== -1 && trainer1Col !== -1 && trainer2Col !== -1) {
                        structure.dayColumns[day].push({
                            dayCol: colIndex,
                            locationCol: locationCol,
                            classCol: classCol,
                            trainer1Col: trainer1Col,
                            trainer2Col: trainer2Col,
                            coverCol: coverCol
                        });
                        
                        // Update the structure collections
                        structure.trainer2Columns[day] = trainer2Col;
                        structure.coverColumns[day] = coverCol;
                        
                        console.log(`📋 Found ${day} at column ${colIndex}, location: ${locationCol}, class: ${classCol}, trainer1: ${trainer1Col}, trainer2: ${trainer2Col}, cover: ${coverCol}`);
                    } else {
                        console.log(`❌ Invalid column mapping for ${day} - skipping`);
                    }
                    break;
                }
            }
        }

        return structure;
    }

    /**
     * Apply email data (covers and themes) to the existing sheet structure
     */
    applyEmailDataToSheet(currentValues, emailInfo, structure) {
        console.log(`🔄 Applying ${emailInfo.covers.length} covers and ${emailInfo.themes.length} themes to sheet...`);
        
        // Create a copy of current values to modify
        const updatedValues = currentValues.map(row => [...row]);
        
        let coversApplied = 0;
        let themesApplied = 0;
        
        // Apply covers
        console.log('🔍 Starting cover application...');
        for (const cover of emailInfo.covers) {
            console.log(`📝 Processing cover for ${cover.day} at ${cover.location}: ${cover.trainer}`);
            console.log(`🔍 Cover details:`, JSON.stringify(cover, null, 2));
            
            const dayColumns = structure.dayColumns[cover.day];
            if (!dayColumns) {
                console.log(`❌ No columns found for day: ${cover.day}`);
                continue;
            }
            
            console.log(`🔍 Found ${dayColumns.length} column configs for ${cover.day}`);
            
            // Find rows that match this cover's criteria
            for (let rowIndex = structure.headerRows; rowIndex < updatedValues.length; rowIndex++) {
                const row = updatedValues[rowIndex];
                if (!row) continue;
                
                // Check each day column configuration for this day
                for (const colConfig of dayColumns) {
                    if (!colConfig.coverCol || colConfig.coverCol >= row.length) continue;
                    
                    const timeCell = String(row[0] || '').trim(); // Time is usually in first column
                    const locationCell = String(row[colConfig.locationCol] || '').trim().toLowerCase();
                    const classCell = String(row[colConfig.classCol] || '').toLowerCase();
                    
                    // Debug: Show what we're checking
                    if (timeCell && locationCell.includes(cover.location.toLowerCase())) {
                        console.log(`🔍 Checking row ${rowIndex + 1}: Time="${timeCell}" Location="${locationCell}" Class="${classCell}" vs Cover Location="${cover.location}"`);
                    }
                    
                    // Match location - improved matching
                    if (!this.matchLocation(locationCell, cover.location)) continue;
                    
                    let shouldApplyCover = false;
                    
                    if (cover.timePattern && cover.classType) {
                        // Handle pattern-based covers (morning cycles, evening barre)
                        console.log(`🔍 Checking pattern cover: ${cover.timePattern} ${cover.classType} against ${timeCell} ${classCell}`);
                        
                        if (cover.timePattern === 'morning') {
                            // Check if it's AM and matches class type
                            if (timeCell.toLowerCase().includes('am') && 
                                classCell.includes(cover.classType.toLowerCase())) {
                                shouldApplyCover = true;
                                console.log(`✅ Morning ${cover.classType} match found`);
                            }
                        } else if (cover.timePattern === 'evening') {
                            // Check if it's PM and matches class type
                            if (timeCell.toLowerCase().includes('pm') && 
                                classCell.includes('barre')) { // More flexible for barre matching
                                shouldApplyCover = true;
                                console.log(`✅ Evening Barre match found`);
                            }
                        }
                    } else if (cover.times && cover.times.length > 0) {
                        // Handle specific time-based covers
                        for (const coverTime of cover.times) {
                            // Normalize time formats for comparison
                            const normalizedCoverTime = this.normalizeTime(coverTime);
                            const normalizedCellTime = this.normalizeTime(timeCell);
                            
                            console.log(`🔍 Time comparison: "${normalizedCellTime}" vs "${normalizedCoverTime}"`);
                            
                            // Flexible time matching - handle slight variations
                            if (this.timeMatches(normalizedCellTime, normalizedCoverTime)) {
                                shouldApplyCover = true;
                                console.log(`✅ Time match found!`);
                                break;
                            }
                        }
                    }
                    
                    if (shouldApplyCover) {
                        row[colConfig.coverCol] = cover.trainer;
                        coversApplied++;
                        console.log(`✅ Applied cover: ${cover.trainer} to ${cover.day} ${timeCell} ${classCell} at row ${rowIndex + 1}, col ${colConfig.coverCol + 1}`);
                    }
                }
            }
        }

        // DEBUG: Show comprehensive sheet structure analysis
        console.log('\n🔍 DEBUG: Sheet Structure Analysis:');
        console.log('Day columns:', Object.keys(structure.dayColumns));
        for (const [day, columns] of Object.entries(structure.dayColumns)) {
            console.log(`${day}:`, columns.map(col => `{day:${col.dayCol}, location:${col.locationCol}, class:${col.classCol}, trainer1:${col.trainer1Col}, trainer2:${col.trainer2Col}, cover:${col.coverCol}}`));
        }
        
        // DEBUG: Show what themes we're trying to apply
        console.log('\n🔍 DEBUG: Themes to apply:');
        for (const theme of emailInfo.themes) {
            console.log(`  - ${theme.classType} theme "${theme.theme}" for ${theme.day} ${theme.time || ''} at ${theme.location || 'any location'}`);
        }
        
        // DEBUG: Show sample data for each day to understand the actual content
        console.log('\n🔍 DEBUG: Sample data by day:');
        for (const [dayName, dayColumns] of Object.entries(structure.dayColumns)) {
            if (!dayColumns || !dayColumns[0]) continue;
            const colConfig = dayColumns[0];
            
            console.log(`\n${dayName} column structure:`);
            console.log(`  locationCol: ${colConfig.locationCol}, classCol: ${colConfig.classCol}, trainer2Col: ${colConfig.trainer2Col}`);
            
            // Show some sample rows for this day
            console.log(`  Sample ${dayName} data:`);
            let count = 0;
            for (let i = structure.headerRows; i < updatedValues.length && count < 5; i++) {
                const row = updatedValues[i];
                if (row && row[0]) {
                    const time = String(row[0] || '').trim();
                    const location = String(row[colConfig.locationCol] || '').trim();
                    const classType = String(row[colConfig.classCol] || '').trim();
                    const trainer1 = String(row[colConfig.trainer1Col] || '').trim();
                    const trainer2 = String(row[colConfig.trainer2Col] || '').trim();
                    
                    if (time) {
                        console.log(`    Row ${i+1}: Time="${time}" Location="${location}" Class="${classType}" T1="${trainer1}" T2="${trainer2}"`);
                        count++;
                    }
                }
            }
        }
        console.log('');
        
        // Apply themes with fixed matching logic for actual class names
        for (const theme of emailInfo.themes) {
            console.log(`🎨 Processing theme for ${theme.day}: ${theme.theme} (Type: ${theme.classType})`);
            
            if (theme.classType === 'FIT' && theme.location === 'All') {
                // Apply theme to all FIT classes across all days
                console.log(`🔍 Looking for FIT classes across all days for theme: ${theme.theme}`);
                
                for (const [dayName, dayColumns] of Object.entries(structure.dayColumns)) {
                    if (!dayColumns) continue;
                    
                    for (const colConfig of dayColumns) {
                        for (let rowIndex = structure.headerRows; rowIndex < updatedValues.length; rowIndex++) {
                            const row = updatedValues[rowIndex];
                            if (!row || row.length <= colConfig.trainer2Col) continue;
                            
                            const classCell = String(row[colConfig.classCol] || '').trim();
                            const timeCell = String(row[0] || '').trim();
                            const locationCell = String(row[colConfig.locationCol] || '').toLowerCase().trim();
                            
                            // Match FIT classes using dynamic matching
                            if (this.classNamesMatch(classCell, 'fit') && timeCell) {
                                row[colConfig.trainer2Col] = theme.theme;
                                console.log(`✅ Applied FIT theme: ${theme.theme} to ${dayName} row ${rowIndex + 1}, col ${colConfig.trainer2Col + 1} (Class: "${classCell}", Time: "${timeCell}", Location: "${locationCell}")`);
                                themesApplied++;
                            }
                        }
                    }
                }
            } else if (theme.classType === 'Amped Up' && theme.day) {
                // Apply to Amped Up classes on specific day at specific location
                console.log(`🔍 Looking for Amped Up classes on ${theme.day} at ${theme.location}`);
                
                const dayColumns = structure.dayColumns[theme.day];
                if (dayColumns) {
                    console.log(`🔍 Found ${dayColumns.length} column configs for ${theme.day}`);
                    for (const colConfig of dayColumns) {
                        for (let rowIndex = structure.headerRows; rowIndex < updatedValues.length; rowIndex++) {
                            const row = updatedValues[rowIndex];
                            if (!row || row.length <= colConfig.trainer2Col) continue;
                            
                            const classCell = String(row[colConfig.classCol] || '').trim();
                            const locationCell = String(row[colConfig.locationCol] || '').toLowerCase().trim();
                            const timeCell = String(row[0] || '').trim();
                            
                            // Show all classes for this day for debugging
                            if (classCell && this.matchLocation(locationCell, 'kemps')) {
                                console.log(`🔍 ${theme.day} Kemps class found: Time="${timeCell}" Class="${classCell}" Location="${locationCell}"`);
                            }
                            
                            // Match Amped Up classes at correct location using dynamic matching
                            if (this.classNamesMatch(classCell, 'amped up') && 
                                this.matchLocation(locationCell, theme.location) && 
                                timeCell) {
                                console.log(`🎯 FOUND AMPED UP MATCH: Class="${classCell}" normalized="${this.normalizeClassName(classCell)}" Location="${locationCell}" matches="${this.matchLocation(locationCell, theme.location)}"`);
                                row[colConfig.trainer2Col] = theme.theme;
                                console.log(`✅ Applied Amped Up theme: ${theme.theme} to ${theme.day} row ${rowIndex + 1}, col ${colConfig.trainer2Col + 1} (Class: "${classCell}", Time: "${timeCell}", Location: "${locationCell}")`);
                                themesApplied++;
                            } else if (classCell && this.matchLocation(locationCell, theme.location) && timeCell) {
                                // Debug: Show why it didn't match
                                const classMatch = this.classNamesMatch(classCell, 'amped up');
                                const locationMatch = this.matchLocation(locationCell, theme.location);
                                console.log(`❌ AMPED UP NO MATCH: Class="${classCell}" (normalized="${this.normalizeClassName(classCell)}") classMatch=${classMatch}, locationMatch=${locationMatch}, theme.location="${theme.location}"`);
                            }
                        }
                    }
                }
            } else if (theme.classType === 'CYCLE' && theme.location === 'Bandra' && theme.time && theme.day) {
                // Apply to specific cycle classes at Bandra with exact time and day matching
                console.log(`🔍 Looking for cycle classes on ${theme.day} at ${theme.location} at time ${theme.time}`);
                
                const dayColumns = structure.dayColumns[theme.day];
                if (dayColumns) {
                    console.log(`🔍 Found ${dayColumns.length} column configs for ${theme.day}`);
                    for (const colConfig of dayColumns) {
                        for (let rowIndex = structure.headerRows; rowIndex < updatedValues.length; rowIndex++) {
                            const row = updatedValues[rowIndex];
                            if (!row || row.length <= colConfig.trainer2Col) continue;
                            
                            const timeCell = String(row[0] || '').trim();
                            const classCell = String(row[colConfig.classCol] || '').trim();
                            const locationCell = String(row[colConfig.locationCol] || '').toLowerCase().trim();
                            
                            // Debug: Show all cycle classes for this day
                            if (this.classNamesMatch(classCell, 'cycle') && timeCell) {
                                console.log(`🚴‍♂️ ${theme.day} cycle class found: Time="${timeCell}" Class="${classCell}" Location="${locationCell}"`);
                            }
                            
                            // Show all Bandra cycle classes for this day for debugging
                            if (this.classNamesMatch(classCell, 'cycle') && this.matchLocation(locationCell, 'bandra')) {
                                console.log(`🔍 ${theme.day} Bandra cycle: Time="${timeCell}" Class="${classCell}" Location="${locationCell}" (looking for time: ${theme.time})`);
                                console.log(`   Normalized times: cellTime="${this.normalizeTimeFormat(timeCell)}" vs themeTime="${this.normalizeTimeFormat(theme.time)}"`);
                            }                            // Match cycle classes at Bandra with matching time using dynamic matching
                            if (this.classNamesMatch(classCell, 'cycle') && 
                                this.matchLocation(locationCell, 'bandra') && 
                                timeCell) {
                                
                                // Normalize times for exact comparison
                                const themeTime = this.normalizeTimeFormat(theme.time);
                                const cellTime = this.normalizeTimeFormat(timeCell);
                                
                                console.log(`🔍 Time comparison: "${cellTime}" vs "${themeTime}" for theme: ${theme.theme}`);
                                
                                // Exact time match required
                                if (cellTime === themeTime) {
                                    row[colConfig.trainer2Col] = theme.theme;
                                    console.log(`✅ Applied Bandra cycle theme: ${theme.theme} to ${theme.day} row ${rowIndex + 1}, col ${colConfig.trainer2Col + 1} (Class: "${classCell}", Time: "${timeCell}", Location: "${locationCell}")`);
                                    themesApplied++;
                                } else {
                                    console.log(`❌ Time mismatch: "${cellTime}" ≠ "${themeTime}" for ${theme.theme}`);
                                }
                            }
                        }
                    }
                }
            } else {
                // Handle other specific theme types with flexible validation
                console.log(`🔍 Looking for ${theme.classType || 'general'} classes on ${theme.day}`);
                
                if (theme.day && theme.day !== 'All') {
                    const dayColumns = structure.dayColumns[theme.day];
                    if (dayColumns) {
                        for (const colConfig of dayColumns) {
                            for (let rowIndex = structure.headerRows; rowIndex < updatedValues.length; rowIndex++) {
                                const row = updatedValues[rowIndex];
                                if (!row || row.length <= colConfig.trainer2Col) continue;
                                
                                const classCell = String(row[colConfig.classCol] || '').toLowerCase().trim();
                                const timeCell = String(row[0] || '').trim();
                                const locationCell = String(row[colConfig.locationCol] || '').toLowerCase().trim();
                                
                                // Only apply if we have exact class type match
                                if (classCell && theme.classType && 
                                    classCell.includes(theme.classType.toLowerCase()) && 
                                    timeCell) {
                                    row[colConfig.trainer2Col] = theme.theme;
                                    console.log(`✅ Applied ${theme.classType} theme: ${theme.theme} to ${theme.day} row ${rowIndex + 1}, col ${colConfig.trainer2Col + 1} (Class: "${classCell}", Time: "${timeCell}", Location: "${locationCell}")`);
                                    themesApplied++;
                                }
                            }
                        }
                    }
                }
            }
        }

        console.log(`📊 Applied ${coversApplied} covers and ${themesApplied} themes to spreadsheet`);
        return updatedValues;
    }

    /**
     * Check if two time strings match (allowing for format variations)
     */
    timesMatch(time1, time2) {
        if (!time1 || !time2) return false;
        
        // Normalize both times
        const norm1 = this.normalizeTime(time1).toLowerCase();
        const norm2 = this.normalizeTime(time2).toLowerCase();
        
        return norm1.includes(norm2.split(' ')[0]) || norm2.includes(norm1.split(' ')[0]);
    }

    /**
     * Normalize time format for comparison
     */
    normalizeTimeFormat(timeStr) {
        if (!timeStr) return '';
        
        // Remove extra spaces and convert to lowercase
        let normalized = timeStr.toLowerCase().trim();
        
        // Handle dot notation (e.g., "10.15 am" -> "10:15 am")
        normalized = normalized.replace(/(\d+)\.(\d+)/, '$1:$2');
        
        // Handle 24-hour format (e.g., "07:30", "18:00")
        if (/^\d{1,2}:\d{2}$/.test(normalized)) {
            const [hours, minutes] = normalized.split(':');
            const hour = parseInt(hours);
            
            if (hour === 0) {
                normalized = `12:${minutes} am`;
            } else if (hour < 12) {
                normalized = `${hour}:${minutes} am`;
            } else if (hour === 12) {
                normalized = `12:${minutes} pm`;
            } else {
                normalized = `${hour - 12}:${minutes} pm`;
            }
        }
        
        // Handle formats like "8 am", "9:15 am" etc.
        // Add :00 if missing minutes for consistency
        normalized = normalized.replace(/^(\d{1,2})\s+([ap]m)/, '$1:00 $2');
        
        // Convert "7:30 am" to "07:30 am" format for consistency
        normalized = normalized.replace(/^(\d)(:?\d*)(\s*[ap]m)/, '0$1$2$3');
        
        // Ensure consistent spacing
        normalized = normalized.replace(/(\d+:?\d*)\s*([ap]m)/, '$1 $2');
        
        // Final step: ensure all times have :XX format (add :00 if missing)
        normalized = normalized.replace(/^(\d{2})\s+([ap]m)/, '$1:00 $2');
        
        return normalized;
    }

    /**
     * Normalize class names for consistent matching
     */
    normalizeClassName(className) {
        if (!className) return '';
        
        return className
            .toLowerCase()
            .trim()
            .replace(/^studio\s+/, '') // Remove "Studio" prefix
            .replace(/\s+/g, ' ') // Normalize spaces
            .replace(/!$/, '') // Remove trailing !
            .trim();
    }

    /**
     * Check if two class names match (supports partial matching)
     */
    classNamesMatch(className1, className2) {
        if (!className1 || !className2) return false;
        
        const norm1 = this.normalizeClassName(className1);
        const norm2 = this.normalizeClassName(className2);
        
        // Direct match
        if (norm1 === norm2) return true;
        
        // Partial matches for common variations
        if (norm1.includes('powercycle') && norm2.includes('cycle')) return true;
        if (norm1.includes('cycle') && norm2.includes('powercycle')) return true;
        if (norm1.includes('amped') && norm2.includes('amped')) return true;
        if (norm1.includes('fit') && norm2.includes('fit')) return true;
        
        return false;
    }

    /**
     * Better location matching for covers
     */
    matchLocation(locationInSheet, coverLocation) {
        if (!locationInSheet || !coverLocation) return false;
        
        const sheetLoc = locationInSheet.toLowerCase().trim();
        const coverLoc = coverLocation.toLowerCase().trim();
        
        // Direct match
        if (sheetLoc.includes(coverLoc)) return true;
        
        // Location aliases
        const aliases = {
            'kemps': ['kwality', 'kemps corner', 'kemps'],
            'bandra': ['supreme', 'bandra'],
            'kenkere': ['kenkere'],
            'south united': ['south', 'united'],
            'copper': ['copper', 'cloves'],
            'wework': ['wework'],
            'physique': ['physique']
        };
        
        for (const [key, values] of Object.entries(aliases)) {
            if (coverLoc.includes(key)) {
                return values.some(alias => sheetLoc.includes(alias));
            }
        }
        
        return false;
    }

    /**
     * Flexible time matching for covers
     */
    timeMatches(time1, time2) {
        if (!time1 || !time2) return false;
        
        // Direct match
        if (time1 === time2) return true;
        
        // Extract hour and minute from both times
        const extractTime = (timeStr) => {
            const match = timeStr.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)/i);
            if (match) {
                const hour = parseInt(match[1]);
                const minute = match[2] ? parseInt(match[2]) : 0;
                const period = match[3].toLowerCase();
                return { hour, minute, period };
            }
            return null;
        };
        
        const t1 = extractTime(time1);
        const t2 = extractTime(time2);
        
        if (t1 && t2) {
            // Compare hour, minute and period
            return t1.hour === t2.hour && 
                   t1.minute === t2.minute && 
                   t1.period === t2.period;
        }
        
        return false;
    }

    /**
     * Format date from sheet to DD MMM YYYY format
     */
    formatDateFromSheet(dateStr) {
        if (!dateStr) return '';
        
        try {
            // Handle various date formats from sheets
            let cleanDate = dateStr.trim();
            
            // Try parsing as a date
            const parsedDate = new Date(cleanDate);
            
            if (!isNaN(parsedDate.getTime())) {
                const day = parsedDate.getDate().toString().padStart(2, '0');
                const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                const month = months[parsedDate.getMonth()];
                const year = parsedDate.getFullYear();
                
                return `${day} ${month} ${year}`;
            }
            
            // If it's already in DD MMM YYYY format, return as-is
            if (/\d{1,2}\s+[A-Za-z]{3}\s+\d{4}/.test(cleanDate)) {
                return cleanDate;
            }
            
            // Fallback: return original
            return cleanDate;
            
        } catch (error) {
            console.log(`⚠️ Could not format date "${dateStr}", using as-is`);
            return dateStr;
        }
    }

    /**
     * Read HTML file
     */
    readHTML() {
        console.log('📄 Reading HTML file...');
        const htmlContent = fs.readFileSync(this.htmlPath, 'utf-8');
        this.$ = cheerio.load(htmlContent, {
            xmlMode: false,
            decodeEntities: false,
            _useHtmlParser2: true
        });
        console.log('✅ HTML loaded successfully');
        
        // DEBUG: Count PDF-related elements before processing
        console.log('\n🔍 DEBUG: PDF Elements Count Before Processing:');
        console.log('  - <script> tags:', this.$('script').length);
        console.log('  - <div id="pg1">:', this.$('#pg1').length);
        console.log('  - <div id="pg2">:', this.$('#pg2').length);
        console.log('  - <img id="pdf1">:', this.$('#pdf1').length);
        console.log('  - <img id="pdf2">:', this.$('#pdf2').length);
        console.log('  - metadata script:', this.$('script#metadata').length);
        console.log('  - annotations script:', this.$('script#annotations').length);
        console.log('  - Total spans:', this.$('span').length);
    }

    /**
     * Normalize time format (handle inconsistencies like "7,15 PM")
     * Ensures format is always HH:MM AM/PM with space before AM/PM
     */
    normalizeTime(time) {
        if (!time) return '';
        // Replace comma with colon and trim
        let normalized = time.replace(',', ':').trim();
        
        // Ensure space before AM/PM
        normalized = normalized.replace(/([0-9])(AM|PM)/i, '$1 $2');
        
        // Pad single digit hours with leading zero
        normalized = normalized.replace(/^([0-9]):/, '0$1:');
        
        return normalized;
    }

    /**
     * Get trainer first name
     */
    getTrainerFirstName(trainerFullName) {
        if (!trainerFullName) return '';
        return trainerFullName.split(' ')[0];
    }

    /**
     * Get formatted date range from column G data
     * Returns string like "November 17th - November 23rd"
     */
    getDateRangeFromSheet() {
        if (!this.allSheetRecords || this.allSheetRecords.length === 0) {
            console.warn('⚠️  No sheet records available for date range extraction');
            return 'November 17th - November 23rd'; // Fallback
        }

        const dates = [];
        this.allSheetRecords.forEach(record => {
            const dateStr = record['Date']; // Column G header is "Date"
            if (dateStr && dateStr.trim()) {
                try {
                    // Parse various date formats (MM/DD/YYYY, M/D/YYYY, etc.)
                    const parsed = new Date(dateStr);
                    if (!isNaN(parsed.getTime())) {
                        dates.push(parsed);
                    }
                } catch (e) {
                    // Skip invalid dates
                }
            }
        });

        if (dates.length === 0) {
            console.warn('⚠️  No valid dates found in column G');
            return 'November 17th - November 23rd'; // Fallback
        }

        // Find min and max dates
        const minDate = new Date(Math.min(...dates));
        const maxDate = new Date(Math.max(...dates));

        // Format with ordinal suffixes
        const formatWithOrdinal = (date) => {
            const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December'];
            const day = date.getDate();
            const month = months[date.getMonth()];
            
            // Get ordinal suffix
            let suffix = 'th';
            if (day === 1 || day === 21 || day === 31) suffix = 'st';
            else if (day === 2 || day === 22) suffix = 'nd';
            else if (day === 3 || day === 23) suffix = 'rd';
            
            return `${month} ${day}${suffix}`;
        };

        const startStr = formatWithOrdinal(minDate);
        const endStr = formatWithOrdinal(maxDate);

        console.log(`📅 Date range extracted: ${startStr} - ${endStr}`);
        return `${startStr} - ${endStr}`;
    }

    /**
     * Create organized schedule data structure
     */
    organizeScheduleByDay() {
        const scheduleByDay = {
            'Monday': [],
            'Tuesday': [],
            'Wednesday': [],
            'Thursday': [],
            'Friday': [],
            'Saturday': [],
            'Sunday': []
        };

        let themeDebugCount = 0;
        this.kwalityClasses.forEach(classData => {
            const day = classData.Day;
            if (scheduleByDay[day]) {
                // Check for theme data in various possible column names
                let theme = classData.Theme || classData.theme || classData['Theme Name'] || 
                           classData['theme_name'] || classData['Class Theme'] || 
                           classData['class_theme'] || classData.H || classData['Column H'] || '';
                
                // If no theme found in data, apply known theme patterns
                if (!theme || !theme.trim()) {
                    theme = this.getThemeForClass(classData);
                }
                
                // Debug: Log first 5 entries with theme data
                if (theme && theme.trim() && themeDebugCount < 5) {
                    console.log(`🎨 Theme found: "${theme}" for ${classData.Class} on ${day} at ${classData.Time}`);
                    themeDebugCount++;
                }
                
                scheduleByDay[day].push({
                    time: this.normalizeTime(classData.Time),
                    class: classData.Class,
                    trainer: classData.Trainer,
                    notes: classData.Notes || '',
                    theme: theme.trim() // Add theme information
                });
            }
        });
        
        if (themeDebugCount === 0) {
            console.log('⚠️  WARNING: No theme data found in any records! Check if theme patterns are properly configured.');
        }

        return scheduleByDay;
    }

    /**
     * Get theme for a class based on known patterns when no theme data is in spreadsheet
     */
    getThemeForClass(classData) {
        const day = classData.Day ? classData.Day.toLowerCase() : '';
        const time = classData.Time ? classData.Time.toLowerCase() : '';
        const location = classData.Location ? classData.Location.toLowerCase() : '';
        const classType = classData.Class ? classData.Class.toLowerCase() : '';
        
        // All FIT classes get TABATA theme
        if (classType.includes('fit')) {
            return 'TABATA';
        }
        
        // Bandra PowerCycle specific themes based on parsed email data
        if (location.includes('bandra') && classType.includes('powercycle')) {
            // Monday 10:00 AM - Taylor Swift vs Kendrick Lamar
            if (day === 'monday' && (time.includes('10:00') || time.includes('10 am'))) {
                return 'Taylor Swift vs Kendrick lamar';
            }
            // Tuesday 7:15 PM - Taylor Swift Vs Kendrick
            if (day === 'tuesday' && (time.includes('7:15 pm') || time.includes('19:15'))) {
                return 'Taylor Swift Vs Kendrick';
            }
            // Wednesday 6:00 PM - Latin Heat
            if (day === 'wednesday' && (time.includes('6:00 pm') || time.includes('18:00') || time.includes('6 pm'))) {
                return 'Latin Heat';
            }
            // Thursday 8:00 AM - Latin Heat
            if (day === 'thursday' && (time.includes('8:00 am') || time.includes('8 am') || time.includes('08:00'))) {
                return 'Latin Heat';
            }
            // Saturday 9:30 AM - Taylor Swift Vs Kendrick Lamar
            if (day === 'saturday' && (time.includes('9:30 am') || time.includes('09:30'))) {
                return 'Taylor Swift Vs Kendrick Lamar';
            }
            // Sunday 10:00 AM - Latin Heat
            if (day === 'sunday' && (time.includes('10:00 am') || time.includes('10 am') || time.includes('10:00'))) {
                return 'Latin Heat';
            }
        }
        
        // Kemps Amped Up classes get specific themes
        if (location.includes('kemps') && (classType.includes('amped up') || classType === 'fit')) {
            if (day === 'tuesday') {
                return 'Icy Isometric';
            }
        }
        
        return ''; // No theme found
    }

    /**
     * Create a neat theme badge for display
     */
    createThemeBadge(theme, location = 'kemps') {
        // Clean up the theme name
        let cleanTheme = theme.trim().toUpperCase();
        
        // Standardized background colors with improved contrast
        const bgColor = location.toLowerCase().includes('bandra') ? 
            'linear-gradient(135deg, #022c22 0%, #064e3b 50%, #065f46 100%)' : 
            'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
        
        // Use consistent ⚡️ icon for all badges
        const icon = '⚡️';
        
        // Standardized styling for both locations
        const standardStyle = {
            background: bgColor,
            color: 'white',
            padding: '3px 8px',
            borderRadius: '12px', // Consistent rounded corners
            fontSize: '8px',      // Consistent font size
            fontWeight: '700',    // Consistent font weight
            marginLeft: '8px',
            display: 'inline-block',
            verticalAlign: 'middle',
            lineHeight: '1.3',
            boxShadow: '0 2px 6px rgba(0,0,0,0.4)',
            letterSpacing: '0.1px',
            textTransform: 'uppercase',
            minWidth: 'fit-content',
            maxWidth: '180px',
            textAlign: 'center',
            whiteSpace: 'normal',
            wordWrap: 'break-word',
            border: '1px solid rgba(255,255,255,0.2)',
            position: 'relative',
            top: '-1px'
        };
        
        // Convert style object to CSS string
        const styleString = Object.entries(standardStyle)
            .map(([key, value]) => `${key.replace(/[A-Z]/g, match => '-' + match.toLowerCase())}: ${value}`)
            .join('; ');
        
        return `<span class="theme-badge" style="${styleString}">${icon} ${cleanTheme}</span>`;
    }

    /**
     * Comprehensive cleanup of all existing theme badges from HTML content
     * This only removes actual theme badge elements, not class name spans
     */
    cleanupAllThemeBadges() {
        console.log('🧹 Starting comprehensive theme badge cleanup...');
        
        // Remove by CSS class - this removes the actual theme badge spans
        const themeBadges = this.$('.theme-badge');
        console.log(`   Found ${themeBadges.length} theme-badge elements to remove`);
        themeBadges.remove();
        
        // Get all spans and check for standalone theme badge content
        const allSpans = this.$('span');
        let removedCount = 0;
        
        allSpans.each((index, element) => {
            const $span = this.$(element);
            const spanText = $span.text().trim();
            
            // Only remove spans that are clearly standalone theme indicators,
            // NOT class names like "PowerCycle - Instructor"
            const hasLightningEmoji = /[⚡️⚡]/.test(spanText);
            const isStandaloneThemeKeyword = /^\s*(?:POWER|THEME|SPECIAL)\s*$/i.test(spanText);
            
            // Only remove if it's clearly a theme badge, not a class description
            if (hasLightningEmoji || isStandaloneThemeKeyword) {
                console.log(`   Removing span with theme badge content: "${spanText.substring(0, 50)}${spanText.length > 50 ? '...' : ''}"`);
                $span.remove();
                removedCount++;
            }
        });
        
        console.log(`   Removed ${removedCount} spans with theme badge content`);
        console.log('✅ Theme badge cleanup complete');
    }

    /**
     * Check if a span is a header/title element that should be protected
     */
    isHeaderElement($span) {
        const text = $span.text().trim().toUpperCase();
        const protectedKeywords = [
            'STUDIO',
            'SCHEDULE',
            'KEMPS',
            'CORNER',
            'BEGINNER',
            'INTERMEDIATE',
            'ADVANCED',
            'FOUNDATION',
            'THEMED CLASSES',
            'STATIC MAGIC',
            'MONDAY',
            'TUESDAY',
            'WEDNESDAY',
            'THURSDAY',
            'FRIDAY',
            'SATURDAY',
            'SUNDAY'
        ];
        
        // Check if text contains any protected keywords
        const containsKeyword = protectedKeywords.some(keyword => text.includes(keyword));
        
        // Check if it looks like a date (contains "th" and month names)
        const looksLikeDate = /(?:january|february|march|april|may|june|july|august|september|october|november|december).*\d{1,2}(?:st|nd|rd|th)/i.test(text);
        
        // Check if text is very long (likely a header)
        const isLongText = text.length > 30;
        
        return containsKeyword || looksLikeDate || (isLongText && text.includes(':'));
    }

    /**
     * Update positioned span elements (visual layout)
     */
    updatePositionedSpans() {
        console.log('🔄 Updating positioned span elements...');
        const scheduleByDay = this.organizeScheduleByDay();
        let updateCount = 0;

        // Helper to parse left/top from inline style
        const getSpanPosition = ($span) => {
            const style = $span.attr('style') || '';
            // Extract left/top values like left:123px; top:456px;
            const leftMatch = style.match(/left\s*:\s*([\d.]+)px/i);
            const topMatch = style.match(/top\s*:\s*([\d.]+)px/i);
            return {
                left: leftMatch ? parseFloat(leftMatch[1]) : NaN,
                top: topMatch ? parseFloat(topMatch[1]) : NaN
            };
        };

        // Collect all time spans
        const timeSpans = this.$('span').filter((_i, elem) => {
            return /^\d{1,2}:\d{2}\s*(?:AM|PM)$/i.test(this.$(elem).text().trim());
        }).get();

        console.log(`\n🔍 DEBUG: Found ${timeSpans.length} time spans to process`);

        // Build column clusters by x-position (left). Tolerance ~ 20px
        const lefts = [];
        timeSpans.forEach((el) => {
            const pos = getSpanPosition(this.$(el));
            if (!Number.isNaN(pos.left)) lefts.push(pos.left);
        });

        const clusterPositions = (vals, tolerance = 20) => {
            const sorted = vals.slice().sort((a, b) => a - b);
            const clusters = [];
            for (const v of sorted) {
                const last = clusters[clusters.length - 1];
                if (!last) {
                    clusters.push({ values: [v] });
                } else {
                    const mean = last.values.reduce((s, x) => s + x, 0) / last.values.length;
                    if (Math.abs(v - mean) <= tolerance) {
                        last.values.push(v);
                    } else {
                        clusters.push({ values: [v] });
                    }
                }
            }
            return clusters.map(cluster => ({
                center: cluster.values.reduce((s, x) => s + x, 0) / cluster.values.length,
                count: cluster.values.length
            }));
        };

        let clusters = clusterPositions(lefts, 20);
        // If too many clusters, keep the 7 most populated, then sort by center
        if (clusters.length > 7) {
            clusters.sort((a, b) => b.count - a.count);
            clusters = clusters.slice(0, 7);
        }
        // Sort by x position ascending
        clusters.sort((a, b) => a.center - b.center);

        const dayOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        if (clusters.length !== 7) {
            console.warn(`⚠️  Expected 7 day columns but found ${clusters.length}. Proceeding with left-to-right mapping for available columns.`);
        }
        const dayByColumnIndex = {};
        clusters.forEach((_c, idx) => {
            if (idx < dayOrder.length) dayByColumnIndex[idx] = dayOrder[idx];
        });

        const findDayForSpan = ($span) => {
            const pos = getSpanPosition($span);
            if (Number.isNaN(pos.left) || clusters.length === 0) return null;
            // Find nearest cluster center
            let bestIdx = 0;
            let bestDist = Math.abs(pos.left - clusters[0].center);
            for (let i = 1; i < clusters.length; i++) {
                const d = Math.abs(pos.left - clusters[i].center);
                if (d < bestDist) {
                    bestDist = d;
                    bestIdx = i;
                }
            }
            return dayByColumnIndex[bestIdx] || null;
        };

        // First, clean up all existing theme badges to prevent duplicates
        this.cleanupAllThemeBadges();

        // Track updated day+time to prevent duplicates within a day
        const updatedCombos = new Set();

        timeSpans.forEach((timeElem, index) => {
            const $timeSpan = this.$(timeElem);
            const timeText = $timeSpan.text().trim().replace(/\s/g, ' ').trim();

            console.log(`\n  Processing time span #${index + 1}: "${timeText}"`);
            
            // Determine day by nearest x-position cluster
            const detectedDay = findDayForSpan($timeSpan);
            if (!detectedDay) {
                console.log('    ⚠ Could not determine day for time span; skipping safe update');
                return;
            }

            const classesForDay = scheduleByDay[detectedDay] || [];
            
            // First, scan siblings to get the class name from HTML for better matching
            let htmlClassName = '';
            let tempCurrent = $timeSpan[0].nextSibling;
            while (tempCurrent && htmlClassName === '') {
                if (tempCurrent.type === 'tag' && tempCurrent.name === 'span') {
                    const $tempSpan = this.$(tempCurrent);
                    const tempText = $tempSpan.text().trim();
                    // Stop if we hit another time span
                    if (/^\d{1,2}:\d{2}\s*(?:AM|PM)$/i.test(tempText)) {
                        break;
                    }
                    // Extract class name (before hyphen and trainer name)
                    if (tempText && !this.isHeaderElement($tempSpan)) {
                        htmlClassName = tempText.split('-')[0].trim().toUpperCase();
                        break;
                    }
                }
                tempCurrent = tempCurrent.nextSibling;
            }
            
            // Find matching class - prefer exact class name match if available
            let matchingClass = null;
            const timeMatches = classesForDay.filter(c => {
                const normalizedCsvTime = this.normalizeTime(c.time);
                const csvTimeCompare = normalizedCsvTime.replace(/^0/, '');
                const htmlTimeCompare = timeText.replace(/^0/, '');
                return csvTimeCompare.toLowerCase() === htmlTimeCompare.toLowerCase();
            });
            
            if (timeMatches.length > 1 && htmlClassName) {
                // Multiple classes at same time - match by class name too
                matchingClass = timeMatches.find(c => {
                    const csvClassName = this.normalizeClassName(c.class).toUpperCase();
                    return csvClassName.includes(htmlClassName) || htmlClassName.includes(csvClassName);
                });
            }
            
            // Fall back to first time match if no class name match found
            if (!matchingClass && timeMatches.length > 0) {
                matchingClass = timeMatches[0];
            }

            if (matchingClass) {
                // Create a unique key for this day+time+class+trainer combination to prevent duplicates
                const normalizedClass = this.normalizeClassName(matchingClass.class).toUpperCase();
                const normalizedTrainer = (matchingClass.trainer || '').trim().toUpperCase();
                const elementIndex = $timeSpan.index();
                const combinationKey = `${detectedDay}|${timeText}|${normalizedClass}|${normalizedTrainer}|${elementIndex}`;
                if (updatedCombos.has(combinationKey)) {
                    console.log(`    ⚠ Skipping duplicate: already updated this day/time/class/trainer combination`);
                    return; // Skip this duplicate
                }
                
                console.log(`    ✓ Found matching class in ${detectedDay}: ${matchingClass.class}`);
                console.log(`    🔍 DEBUG matchingClass:`, JSON.stringify({
                    class: matchingClass.class,
                    trainer: matchingClass.trainer,
                    theme: matchingClass.theme,
                    time: matchingClass.time
                }));
                
                let current = $timeSpan[0].nextSibling;
                const spansToRemove = [];
                let firstContentSpan = null;
                let siblingsProcessed = 0;

                console.log(`    Scanning siblings after time span...`);
                
                while (current) {
                    siblingsProcessed++;
                    
                    if (current.type === 'tag' && current.name === 'span') {
                        const $currentSpan = this.$(current);
                        const spanText = $currentSpan.text().trim();
                        const spanId = $currentSpan.attr('id');
                        const spanClass = $currentSpan.attr('class');
                        
                        console.log(`      Sibling #${siblingsProcessed}: <span${spanId ? ' id="'+spanId+'"' : ''}${spanClass ? ' class="'+spanClass+'"' : ''}> text: "${spanText.substring(0, 50)}${spanText.length > 50 ? '...' : ''}"`);
                        
                        if (/^\d{1,2}:\d{2}\s*(?:AM|PM)$/i.test(spanText)) {
                            console.log(`      ↳ Next time span found, stopping scan`);
                            break; // Stop at the next time span
                        }
                        
                        // Check if this is a header/protected element
                        if (this.isHeaderElement($currentSpan)) {
                            console.log(`      ↳ PROTECTED HEADER ELEMENT - stopping scan to preserve`);
                            break;
                        }
                        
                        // Enhanced badge removal - check for CSS classes, inline patterns, and content
                        const hasThemeClass = $currentSpan.hasClass('theme-badge');
                        const hasOldTheme = /[⚡️⚡]/.test(spanText);
                        const hasOldThemeText = /\b(?:theme|special)\b/i.test(spanText);
                        
                        // Only mark for removal if it's not the first span (which contains class info) 
                        // OR if it's clearly a theme badge
                        if (!firstContentSpan) {
                            firstContentSpan = $currentSpan;
                            console.log(`      ↳ Marked as firstContentSpan (class content)`);
                            // For the first span, we'll replace its content entirely, so always add to removal list
                            spansToRemove.push($currentSpan);
                            console.log(`      ↳ First span will be replaced with updated content`);
                        } else if (hasThemeClass || hasOldTheme || hasOldThemeText) {
                            // Remove subsequent spans only if they contain actual theme badge content
                            spansToRemove.push($currentSpan);
                            console.log(`      ↳ Subsequent span contains theme badge content (class: ${hasThemeClass}, emoji: ${hasOldTheme}, text: ${hasOldThemeText}), added to removal list`);
                        } else {
                            console.log(`      ↳ Clean subsequent span, keeping unchanged`);
                        }
                    } else if (current.type === 'tag' && current.name !== 'span') {
                        console.log(`      Sibling #${siblingsProcessed}: <${current.name}> (non-span tag) - stopping scan`);
                        break;
                    }
                    current = current.nextSibling;
                }

                console.log(`    Summary: ${spansToRemove.length} spans marked for removal`);

                if (firstContentSpan && spansToRemove.length > 0) {
                    const normalizedCSVClass = this.normalizeClassName(matchingClass.class);
                    const classDisplay = this.toTitleCase(normalizedCSVClass);
                    const trainerFirstName = this.getTrainerFirstName(matchingClass.trainer);
                    const trainerDisplay = this.toTitleCase(trainerFirstName);
                    let newText = classDisplay;
                    if (trainerDisplay) {
                        newText += ` - ${trainerDisplay}`;
                    }

                    // Add theme badge if theme exists
                    let themeBadge = '';
                    if (matchingClass.theme && matchingClass.theme.trim()) {
                        themeBadge = this.createThemeBadge(matchingClass.theme.trim(), this.currentLocation);
                    }

                    console.log(`    Creating new span with text: "${newText}" and theme: "${matchingClass.theme}"`);

                    // Create a new span with the content, preserving the original's attributes
                    const newSpan = firstContentSpan.clone().text(newText);
                    
                    // Add theme badge as HTML if it exists
                    if (themeBadge) {
                        newSpan.append(themeBadge);
                    }
                    
                    // Apply consistent Montserrat font with darker weight and color for all days
                    newSpan.css('font-family', 'Montserrat, sans-serif');
                    newSpan.css('font-weight', '600');
                    newSpan.css('color', '#1a1a1a');
                    newSpan.css('letter-spacing', '-0.1px');
                    newSpan.css('font-style', 'normal');
                    newSpan.css('text-transform', 'none');

                    // Insert the new span after the time span
                    $timeSpan.after(newSpan);
                    console.log(`    ✓ New span inserted after time span`);

                    // Remove all the old content spans
                    spansToRemove.forEach(($span, idx) => {
                        console.log(`    Removing span #${idx + 1}...`);
                        $span.remove();
                    });
                    console.log(`    ✓ ${spansToRemove.length} spans removed`);
                    
                    // Mark this combination as updated
                    updatedCombos.add(combinationKey);
                    updateCount++;
                } else {
                    console.log(`    ⚠ No content span found after time span`);
                }
            } else {
                console.log(`    ✗ No matching class found for ${detectedDay} at time: "${timeText}"`);
            }
        });

        console.log(`\n✅ Updated ${updateCount} positioned span elements`);
    }

    /**
     * Update dynamically generated schedule entries
     */
    updateScheduleEntries() {
        console.log('🔄 Updating schedule-entry spans...');
        const scheduleByDay = this.organizeScheduleByDay();
        let updateCount = 0;

        this.$('.schedule-entry').each((_i, elem) => {
            const $entry = this.$(elem);
            const day = $entry.attr('data-day');
            const time = $entry.attr('data-time');

            if (day && time && scheduleByDay[day]) {
                const matchingClass = scheduleByDay[day].find(c => 
                    this.normalizeTime(c.time).toLowerCase() === time.toLowerCase()
                );

                if (matchingClass) {
                    // Update data attributes
                    const classDisplay = this.toTitleCase(this.normalizeClassName(matchingClass.class));
                    const trainerDisplay = this.toTitleCase(matchingClass.trainer);
                    $entry.attr('data-class', classDisplay);
                    $entry.attr('data-trainer', trainerDisplay);
                    
                    if (matchingClass.notes) {
                        $entry.attr('data-notes', matchingClass.notes);
                    }
                    
                    if (matchingClass.theme && matchingClass.theme.trim()) {
                        $entry.attr('data-theme', matchingClass.theme.trim());
                    }

                    // Update text content with theme badge
                    let newText = `${time} – ${classDisplay} – ${trainerDisplay}`;
                    if (matchingClass.theme && matchingClass.theme.trim()) {
                        const themeBadge = this.createThemeBadge(matchingClass.theme.trim(), this.currentLocation);
                        newText += ` ${themeBadge}`;
                    }
                    if (matchingClass.notes) {
                        newText += ` – [${matchingClass.notes}]`;
                    }
                    
                    $entry.text(newText);
                    updateCount++;
                }
            }
        });

        console.log(`✅ Updated ${updateCount} schedule-entry spans`);
    }

    /**
     * Update date range headers in HTML with dynamic dates from column G
     */
    updateDateHeaders() {
        console.log('\n📅 Updating date range headers...');
        const dateRange = this.getDateRangeFromSheet();
        
        // Find all spans that contain date-like text (e.g., "November 17th - November 23rd")
        const datePattern = /(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)\s*-\s*(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)/i;
        
        let updatedCount = 0;
        this.$('span').each((_i, elem) => {
            const $span = this.$(elem);
            const text = $span.text().trim();
            
            if (datePattern.test(text)) {
                console.log(`  Found date header: "${text}"`);
                $span.text(dateRange);
                console.log(`  ✓ Updated to: "${dateRange}"`);
                updatedCount++;
            }
        });
        
        console.log(`✅ Updated ${updatedCount} date header spans`);
    }

    /**
     * Main update method
     */
    update() {
        try {
            // Prefer Google Sheet; fallback to CSV if needed
            if (GOOGLE_CONFIG.SPREADSHEET_ID) {
                console.log('ℹ️  Using Google Sheet as data source');
            }
            this.readHTML();
            // Note: readSheet is async; use updateWithPDF for full async flow
            console.warn('⚠️  update() is sync and expects CSV. Use updateWithPDF() for Sheets.');
            
            // Replace background image with custom image (PNG)
            const customImagePath = path.join(__dirname, 'Bandra.png');
            this.replaceBackgroundImage(customImagePath);
            
            this.updatePositionedSpans();
            this.updateScheduleEntries();
            this.updateDateHeaders();
            this.save();
            console.log('🎉 Schedule update completed successfully!');
        } catch (error) {
            console.error('❌ Error updating schedule:', error.message);
            throw error;
        }
    }

    /**
     * Save updated HTML
     */
    /**
     * Replace background image with a custom image
     */
    replaceBackgroundImage(imagePath) {
        console.log('\n🖼️  Replacing background image...');
        
        if (!fs.existsSync(imagePath)) {
            console.error(`❌ Image file not found: ${imagePath}`);
            return;
        }
        
        // Read the image and convert to base64
        const imageBuffer = fs.readFileSync(imagePath);
        const base64Image = imageBuffer.toString('base64');
        const imageExtension = path.extname(imagePath).toLowerCase().substring(1);
        const mimeType = imageExtension === 'jpg' ? 'jpeg' : imageExtension;
        const dataUrl = `data:image/${mimeType};base64,${base64Image}`;
        
        // Replace the img src for pdf1 (first page background)
        const $pdf1 = this.$('#pdf1');
        if ($pdf1.length > 0) {
            $pdf1.attr('src', dataUrl);
            console.log('✅ Replaced background image for page 1');
        } else {
            console.log('⚠️  #pdf1 element not found');
        }
        
        // Also replace pdf2 if it exists
        const $pdf2 = this.$('#pdf2');
        if ($pdf2.length > 0) {
            $pdf2.attr('src', dataUrl);
            console.log('✅ Replaced background image for page 2');
        }
    }

    save() {
        console.log('\n💾 Saving updated HTML...');
        
        // DEBUG: Count PDF-related elements after processing
        console.log('\n🔍 DEBUG: PDF Elements Count After Processing:');
        console.log('  - <script> tags:', this.$('script').length);
        console.log('  - <div id="pg1">:', this.$('#pg1').length);
        console.log('  - <div id="pg2">:', this.$('#pg2').length);
        console.log('  - <img id="pdf1">:', this.$('#pdf1').length);
        console.log('  - <img id="pdf2">:', this.$('#pdf2').length);
        console.log('  - metadata script:', this.$('script#metadata').length);
        console.log('  - annotations script:', this.$('script#annotations').length);
        console.log('  - Total spans:', this.$('span').length);
        
        const updatedHTML = this.$.html();
        
        // DEBUG: Check if PDF elements exist in final HTML string
        console.log('\n🔍 DEBUG: PDF Elements in Final HTML String:');
        console.log('  - Contains "pg1":', updatedHTML.includes('id="pg1"'));
        console.log('  - Contains "pg2":', updatedHTML.includes('id="pg2"'));
        console.log('  - Contains "pdf1":', updatedHTML.includes('id="pdf1"'));
        console.log('  - Contains "pdf2":', updatedHTML.includes('id="pdf2"'));
        console.log('  - Contains "metadata":', updatedHTML.includes('id="metadata"'));
        console.log('  - Contains "annotations":', updatedHTML.includes('id="annotations"'));
        
        fs.writeFileSync(this.outputPath, updatedHTML, 'utf-8');
        console.log(`✅ Saved to: ${this.outputPath}`);
    }

    /**
     * Generate detailed report
     */
    generateReport() {
        const scheduleByDay = this.organizeScheduleByDay();
        console.log('\n📊 Schedule Report for Kwality House, Kemps Corner:\n');
        
        Object.entries(scheduleByDay).forEach(([day, classes]) => {
            if (classes.length > 0) {
                console.log(`${day}:`);
                classes.forEach(c => {
                    console.log(`  ${c.time} - ${c.class} - ${c.trainer}${c.notes ? ' [' + c.notes + ']' : ''}`);
                });
                console.log('');
            }
        });
    }

    /**
     * Get Google OAuth access token
     */
    async getAccessToken() {
        console.log('\n🔐 Getting Google OAuth access token...');
        try {
            const response = await axios.post(GOOGLE_CONFIG.TOKEN_URL, {
                client_id: GOOGLE_CONFIG.CLIENT_ID,
                client_secret: GOOGLE_CONFIG.CLIENT_SECRET,
                refresh_token: GOOGLE_CONFIG.REFRESH_TOKEN,
                grant_type: 'refresh_token'
            });
            console.log('✅ Access token obtained');
            return response.data.access_token;
        } catch (error) {
            console.error('❌ Error getting access token:', error.response?.data || error.message);
            throw error;
        }
    }

    /**
     * Check if file exists in Google Drive folder
     */
    async checkFileExists(drive, fileName) {
        console.log(`\n🔍 Checking if ${fileName} already exists in Drive...`);
        try {
            const response = await drive.files.list({
                q: `name='${fileName}' and '${GOOGLE_CONFIG.FOLDER_ID}' in parents and trashed=false`,
                fields: 'files(id, name)',
                spaces: 'drive'
            });
            
            if (response.data.files && response.data.files.length > 0) {
                const fileId = response.data.files[0].id;
                console.log(`✅ File found: ${response.data.files[0].name} (ID: ${fileId})`);
                console.log(`🔄 Will update existing file to preserve ID and sharing settings`);
                return fileId; // Return file ID for updating
            }
            console.log('📝 File does not exist yet, will create new file');
            return null;
        } catch (error) {
            console.error('❌ Error checking file existence:', error.message);
            throw error;
        }
    }

    /**
     * Generate PDF from HTML file
     */
    async generatePDF() {
        console.log('\n📄 Generating PDF from HTML...');
        const pdfPath = path.join(__dirname, `Schedule-${this.locationName}.pdf`);
        
        try {
            // Read the updated HTML file - ensure we're reading the latest version
            console.log(`  Reading from: ${this.outputPath}`);
            let htmlContent = fs.readFileSync(this.outputPath, 'utf-8');
            console.log(`  HTML content size: ${htmlContent.length} bytes`);
            
            // Convert background image to data URL for better PDF compatibility
            const imagePath = path.join(__dirname, `${this.locationName}.png`);
            if (fs.existsSync(imagePath)) {
                const imageBuffer = fs.readFileSync(imagePath);
                const base64Image = imageBuffer.toString('base64');
                const dataUrl = `data:image/png;base64,${base64Image}`;
                
                // Replace the relative URL with data URL (support both Kemps.png and Bandra.png)
                htmlContent = htmlContent.replace(
                    new RegExp(`background-image: url\\\\('\\\\.\\\\/${this.locationName}\\\\.png'\\\\);`, 'g'),
                    `background-image: url('${dataUrl}');`
                );
                console.log('  ✅ Converted background image to data URL for PDF');
            }
            
            // Add CSS to hide annotations and ensure consistent styling
            const pdfSpecificCSS = `
                <style>
                    /* Completely remove annotations container */
                    .annotations-container,
                    .annotations-container * {
                        display: none !important;
                        visibility: hidden !important;
                        opacity: 0 !important;
                        position: absolute !important;
                        width: 0 !important;
                        height: 0 !important;
                        overflow: hidden !important;
                    }
                    
                    /* Hide IDR Solutions links */
                    a[href*="idrsolutions.com"],
                    a[href*="idrsolutions"] {
                        display: none !important;
                        visibility: hidden !important;
                    }
                    
                    /* Ensure consistent font rendering for all text */
                    * {
                        -webkit-font-smoothing: antialiased !important;
                        -moz-osx-font-smoothing: grayscale !important;
                        text-rendering: optimizeLegibility !important;
                    }
                    
                    /* Force background and color rendering */
                    body, .page, * {
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                        color-adjust: exact !important;
                    }
                    
                    /* Remove any potential overlays or popups */
                    [class*="popup"],
                    [class*="overlay"],
                    [class*="modal"] {
                        display: none !important;
                    }
                    
                    /* Ensure theme-related elements are completely hidden (except badges) */
                    .theme-index,
                    #theme-index,
                    .legend,
                    .theme-legend,
                    .index-legend {
                        display: none !important;
                        visibility: hidden !important;
                        position: absolute !important;
                        width: 0 !important;
                        height: 0 !important;
                        overflow: hidden !important;
                    }
                    
                    /* Hide any page beyond the first 2 pages */
                    .page-container:nth-child(n+3) {
                        display: none !important;
                    }
                </style>
            `;
            
            // Insert the PDF-specific CSS before the closing </head> tag
            htmlContent = htmlContent.replace('</head>', pdfSpecificCSS + '</head>');
            
            // Also add a script to remove annotations via JavaScript
            const removeAnnotationsScript = `
                <script>
                    document.addEventListener('DOMContentLoaded', function() {
                        // Remove all annotation containers
                        const annotations = document.querySelectorAll('.annotations-container');
                        annotations.forEach(el => el.remove());
                        
                        // Remove IDR links
                        const idrLinks = document.querySelectorAll('a[href*="idrsolutions"]');
                        idrLinks.forEach(el => el.remove());
                    });
                </script>
            `;
            
            htmlContent = htmlContent.replace('</body>', removeAnnotationsScript + '</body>');
            
            // Write temporary HTML file
            const tempHtmlPath = path.join(__dirname, 'temp_for_pdf.html');
            fs.writeFileSync(tempHtmlPath, htmlContent, 'utf-8');
            
            // Launch Puppeteer with proper configuration
            const browser = await puppeteer.launch({
                headless: true,
                args: [
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-web-security',
                    '--font-render-hinting=none'
                ]
            });
            
            const page = await browser.newPage();
            
            // Load the HTML file
            await page.goto(`file://${tempHtmlPath}`, {
                waitUntil: 'networkidle0',
                timeout: 30000
            });
            
            // Wait for fonts to load
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Execute script to remove annotations in the page context
            await page.evaluate(() => {
                const annotations = document.querySelectorAll('.annotations-container');
                annotations.forEach(el => {
                    if (el.parentNode) {
                        el.parentNode.removeChild(el);
                    }
                });
                
                const idrLinks = document.querySelectorAll('a[href*="idrsolutions"]');
                idrLinks.forEach(el => {
                    if (el.parentNode) {
                        el.parentNode.removeChild(el);
                    }
                });
            });
            
            // Generate PDF with exact styling preservation
            await page.pdf({
                path: pdfPath,
                format: 'A4',
                printBackground: true,
                preferCSSPageSize: false,
                displayHeaderFooter: false,
                margin: {
                    top: '0mm',
                    right: '0mm',
                    bottom: '0mm',
                    left: '0mm'
                }
            });
            
            await browser.close();
            
            // Clean up temporary file
            fs.unlinkSync(tempHtmlPath);
            
            console.log(`✅ PDF generated: ${pdfPath}`);
            return pdfPath;
        } catch (error) {
            console.error('❌ Error generating PDF:', error.message);
            throw error;
        }
    }

    /**
     * Upload PDF to Google Drive
     */
    async uploadToGoogleDrive(pdfPath) {
        console.log('\n☁️  Uploading PDF to Google Drive...');
        
        try {
            const accessToken = await this.getAccessToken();
            
            // Create OAuth2 client
            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });
            
            const drive = google.drive({ version: 'v3', auth: oauth2Client });
            
            const fileName = `Schedule-${this.locationName}.pdf`;
            const fileMetadata = {
                name: fileName,
                parents: [GOOGLE_CONFIG.FOLDER_ID]
            };
            
            const media = {
                mimeType: 'application/pdf',
                body: fs.createReadStream(pdfPath)
            };
            
            // Check if file already exists
            const existingFileId = await this.checkFileExists(drive, fileName);
            
            let response;
            if (existingFileId) {
                // Update existing file to preserve ID and sharing settings
                console.log('📤 Updating existing file...');
                response = await drive.files.update({
                    fileId: existingFileId,
                    media: media,
                    fields: 'id, name, webViewLink'
                });
                console.log(`✅ File updated successfully!`);
            } else {
                // Create new file
                console.log('📤 Creating new file...');
                response = await drive.files.create({
                    requestBody: fileMetadata,
                    media: media,
                    fields: 'id, name, webViewLink'
                });
                console.log(`✅ File created successfully!`);
            }
            console.log(`   ID: ${response.data.id}`);
            console.log(`   Name: ${response.data.name}`);
            console.log(`   Link: ${response.data.webViewLink || 'https://drive.google.com/file/d/' + response.data.id}`);
            return response.data;
        } catch (error) {
            console.error('❌ Error uploading to Google Drive:', error.message);
            if (error.response) {
                console.error('   Response data:', error.response.data);
            }
            throw error;
        }
    }

    /**
     * Upload a named PDF (e.g., Bandra.pdf) to Drive in same folder.
     */
    async uploadNamedPDF(pdfPath, fileName) {
        console.log(`\n☁️  Uploading ${fileName} to Google Drive...`);
        try {
            const accessToken = await this.getAccessToken();
            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });
            const drive = google.drive({ version: 'v3', auth: oauth2Client });
            const fileMetadata = { name: fileName, parents: [GOOGLE_CONFIG.FOLDER_ID] };
            const media = { mimeType: 'application/pdf', body: fs.createReadStream(pdfPath) };
            const existingFileId = await this.checkFileExists(drive, fileName);
            
            let response;
            if (existingFileId) {
                // Update existing file
                console.log('📤 Updating existing file...');
                response = await drive.files.update({
                    fileId: existingFileId,
                    media: media,
                    fields: 'id, name, webViewLink'
                });
                console.log(`✅ ${fileName} updated: ${response.data.id}`);
            } else {
                // Create new file
                console.log('📤 Creating new file...');
                response = await drive.files.create({ requestBody: fileMetadata, media, fields: 'id, name, webViewLink' });
                console.log(`✅ ${fileName} created: ${response.data.id}`);
            }
            return response.data;
        } catch (err) {
            console.error(`❌ Error uploading ${fileName}:`, err.message);
            throw err;
        }
    }

    /**
     * Main update method with PDF generation and upload (reads from Google Sheets)
     */
    async updateWithPDF() {
        try {
            // Create single backup before updating
            if (fs.existsSync(this.outputPath)) {
                const backupPath = this.outputPath.replace('.html', '.backup.html');
                fs.copyFileSync(this.outputPath, backupPath);
                console.log(`🗂️  Created backup: ${path.basename(backupPath)}`);
            }
            
            // Step 1: Update HTML schedule from Google Sheets Cleaned sheet
            await this.readCleanedSheet(); // Read from Cleaned sheet instead of raw Schedule sheet
            this.readHTML();
            
            const customImagePath = path.join(__dirname, `${this.locationName}.png`);
            this.replaceBackgroundImage(customImagePath);
            
            this.updatePositionedSpans();
            this.updateScheduleEntries();
            this.updateDateHeaders();
            this.save();
            console.log('✅ Schedule HTML updated from Google Sheets Cleaned data!');
            
            // Step 2: Generate PDF
            const pdfPath = await this.generatePDF();
            
            // Step 3: Upload to Google Drive
            await this.uploadToGoogleDrive(pdfPath);
            
            console.log('\n🎉 Complete! Schedule updated from Google Sheets, PDF generated, and uploaded to Google Drive!');
        } catch (error) {
            console.error('❌ Error in updateWithPDF:', error.message);
            throw error;
        }
    }

    /**
     * Complete workflow: Email processing -> Google Sheets -> HTML/PDF (No CSV dependency)
     */
    async completeGoogleSheetsWorkflow() {
        console.log('🚀 Starting complete Google Sheets workflow (no CSV)...');
        
        try {
            // STEP 1: Process email and update Google Sheets
            console.log('📧 Step 1: Processing email and updating Google Sheets...');
            await this.processEmailAndUpdateSchedule();
            console.log('✅ Google Sheets updated with email data\n');
            
            // STEP 2: Update HTML and PDF directly from Google Sheets
            console.log('📄 Step 2: Updating HTML and generating PDF from Google Sheets...');
            await this.updateWithPDF(); // Now uses readCleanedSheet internally
            
            console.log('🎉 Complete Google Sheets workflow finished successfully!');
            console.log('🔍 Summary of updates:');
            console.log('   - Google Sheets updated with email covers and themes');
            console.log('   - Cleaned sheet populated with correct dates from Schedule sheet');
            console.log('   - HTML updated directly from Google Sheets Cleaned data');
            console.log('   - PDF generated and uploaded');
            console.log('   - No CSV files used in the process');
            
        } catch (error) {
            console.error('❌ Workflow failed:', error.message);
            throw error;
        }
    }

    /**
     * Generate PDF with custom file name (does not upload). Uses current outputPath HTML.
     */
    async generatePDFNamed(pdfFileName) {
        console.log(`\n📄 Generating PDF (${pdfFileName}) from HTML...`);
        const pdfPath = path.join(__dirname, pdfFileName);
        try {
            let htmlContent = fs.readFileSync(this.outputPath, 'utf-8');
            
            // Convert background image to data URL
            const isBandraFile = pdfFileName.toLowerCase().includes('bandra');
            const isKempsFileCheck = pdfFileName.toLowerCase().includes('kemps') || pdfFileName.toLowerCase().includes('schedule');
            
            if (isBandraFile) {
                const imagePath = path.join(__dirname, 'Bandra.png');
                if (fs.existsSync(imagePath)) {
                    const imageBuffer = fs.readFileSync(imagePath);
                    const base64Image = imageBuffer.toString('base64');
                    const dataUrl = `data:image/png;base64,${base64Image}`;
                    
                    // Replace the relative URL with data URL for Bandra
                    htmlContent = htmlContent.replace(
                        /background-image: url\('\.\/Bandra\.png'\);/g,
                        `background-image: url('${dataUrl}');`
                    );
                    console.log('  ✅ Converted background image to data URL for Bandra PDF');
                }
            } else if (isKempsFileCheck) {
                const imagePath = path.join(__dirname, 'Kemps.png');
                if (fs.existsSync(imagePath)) {
                    const imageBuffer = fs.readFileSync(imagePath);
                    const base64Image = imageBuffer.toString('base64');
                    const dataUrl = `data:image/png;base64,${base64Image}`;
                    
                    // Replace the relative URL with data URL
                    htmlContent = htmlContent.replace(
                        /background-image: url\('\.\/Kemps\.png'\);/g,
                        `background-image: url('${dataUrl}');`
                    );
                    console.log('  ✅ Converted background image to data URL for Kemps PDF');
                }
            }
            
            // Determine if this is a Kemps file (which needs page limiting)
            const shouldLimitPages = pdfFileName.toLowerCase().includes('kemps') || pdfFileName.toLowerCase().includes('schedule');
            
            const pdfSpecificCSS = `
                <style>
                    .annotations-container, .annotations-container * { display:none !important; }
                    a[href*="idrsolutions"], a[href*="idrsolutions.com"] { display:none !important; }
                    body, .page, * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
                    
                    /* Ensure theme-related elements are completely hidden (except badges) */
                    .theme-index,
                    #theme-index,
                    .legend,
                    .theme-legend,
                    .index-legend {
                        display: none !important;
                    }
                    
                    /* Keep theme badges visible and properly styled */
                    .theme-badge {
                        display: inline-block !important;
                        visibility: visible !important;
                    }
                    
                    ${shouldLimitPages ? `
                    /* Hide any page beyond the first 2 pages (only for Kemps) */
                    .page-container:nth-child(n+3) {
                        display: none !important;
                    }
                    ` : ''}
                </style>
            `;
            htmlContent = htmlContent.replace('</head>', pdfSpecificCSS + '</head>');
            const tempHtmlPath = path.join(__dirname, 'temp_for_pdf_' + pdfFileName + '.html');
            fs.writeFileSync(tempHtmlPath, htmlContent, 'utf-8');
            const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox','--disable-setuid-sandbox'] });
            const page = await browser.newPage();
            await page.goto(`file://${tempHtmlPath}`, { waitUntil: 'networkidle0', timeout: 30000 });
            await new Promise(r => setTimeout(r, 1000));
            await page.pdf({ path: pdfPath, format: 'A4', printBackground: true, margin:{top:'0mm',right:'0mm',bottom:'0mm',left:'0mm'} });
            await browser.close();
            fs.unlinkSync(tempHtmlPath);
            console.log(`✅ PDF generated: ${pdfPath}`);
            return pdfPath;
        } catch (err) {
            console.error('❌ Error generating custom PDF:', err.message);
            throw err;
        }
    }

    /**
     * Generate a combined PDF with both Kemps and Bandra schedules
     * This merges the actual PDF files to preserve each schedule's independent styling
     */
    async generateCombinedPDF() {
        console.log('\n📄 Generating Combined PDF with Kemps and Bandra schedules...');
        
        try {
            const kempsPdfPath = path.join(__dirname, 'Schedule-Kemps.pdf');
            const bandraPdfPath = path.join(__dirname, 'Bandra.pdf');
            const combinedPdfPath = path.join(__dirname, 'Schedule-Combined.pdf');
            
            // Check if both PDFs exist
            if (!fs.existsSync(kempsPdfPath)) {
                throw new Error('Kemps PDF not found. Please run the script to generate it first.');
            }
            if (!fs.existsSync(bandraPdfPath)) {
                throw new Error('Bandra PDF not found. Please run the script to generate it first.');
            }
            
            console.log('📑 Merging PDFs while preserving individual styling...');
            
            // Create a new PDF document
            const mergedPdf = await PDFDocument.create();
            
            // Load the Kemps PDF and copy schedule pages (exclude theme/legend pages)
            console.log('  ➤ Adding Kemps schedule pages...');
            const kempsPdfBytes = fs.readFileSync(kempsPdfPath);
            const kempsPdf = await PDFDocument.load(kempsPdfBytes);
            const totalKempsPages = kempsPdf.getPageCount();
            // For Kemps, copy first 2 pages if available (skip pages 3+ which are theme pages)
            const kempsPageIndices = Array.from({length: Math.min(2, totalKempsPages)}, (_, i) => i);
            const kempsPages = await mergedPdf.copyPages(kempsPdf, kempsPageIndices);
            kempsPages.forEach(page => mergedPdf.addPage(page));
            console.log(`    ✓ Added ${kempsPages.length} pages from Kemps schedule (total available: ${totalKempsPages})`);
            
            // Load the Bandra PDF and copy all available pages (Bandra doesn't seem to have theme pages)
            console.log('  ➤ Adding Bandra schedule pages...');
            const bandraPdfBytes = fs.readFileSync(bandraPdfPath);
            const bandraPdf = await PDFDocument.load(bandraPdfBytes);
            const totalBandraPages = bandraPdf.getPageCount();
            // For Bandra, copy all pages since it doesn't have the theme page issue
            const bandraPageIndices = Array.from({length: totalBandraPages}, (_, i) => i);
            const bandraPages = await mergedPdf.copyPages(bandraPdf, bandraPageIndices);
            bandraPages.forEach(page => mergedPdf.addPage(page));
            console.log(`    ✓ Added ${bandraPages.length} pages from Bandra schedule (total available: ${totalBandraPages})`);
            
            // Save the merged PDF
            const mergedPdfBytes = await mergedPdf.save();
            fs.writeFileSync(combinedPdfPath, mergedPdfBytes);
            
            console.log(`✅ Combined PDF generated: ${combinedPdfPath}`);
            console.log(`   Total pages: ${kempsPages.length + bandraPages.length} (${kempsPages.length} Kemps + ${bandraPages.length} Bandra)`);
            console.log(`   Theme/legend pages excluded from Kemps PDF only`);
            
            // Upload combined PDF to Google Drive
            await this.uploadNamedPDF(combinedPdfPath, 'Schedule-Combined.pdf');
            
            return combinedPdfPath;
        } catch (error) {
            console.error('❌ Error generating combined PDF:', error.message);
            throw error;
        }
    }

    /**
     * Atomic file update process - ensures all three files are updated simultaneously
     */
    async updateAllFilesAtomically() {
        console.log('🔄 Starting atomic file update process...');
        
        try {
            // Step 1: Update Kemps HTML and PDF
            console.log('📄 Step 1: Updating Kemps schedule...');
            await this.updateWithPDF();
            const kempsPdfPath = path.join(__dirname, 'Schedule-Kemps.pdf');
            console.log('   ✓ Kemps schedule updated');
            
            // Step 2: Update Bandra HTML and PDF
            console.log('📄 Step 2: Updating Bandra schedule...');
            await this.updateBandra();
            const bandraPdfPath = path.join(__dirname, 'Bandra.pdf');
            console.log('   ✓ Bandra schedule updated');
            
            // Step 3: Generate combined PDF
            console.log('📑 Step 3: Generating combined PDF...');
            const combinedPdfPath = await this.generateCombinedPDF();
            console.log('   ✓ Combined PDF generated');
            
            // Step 4: Upload all files to Google Drive atomically
            console.log('☁️  Step 4: Uploading all files to Google Drive...');
            await Promise.all([
                this.uploadNamedPDF(kempsPdfPath, 'Schedule-Kemps.pdf'),
                this.uploadNamedPDF(bandraPdfPath, 'Bandra.pdf'),
                this.uploadNamedPDF(combinedPdfPath, 'Schedule-Combined.pdf')
            ]);
            
            console.log('🎉 Atomic file update completed successfully!');
            console.log('📊 Updated files:');
            console.log('   - Kemps.html & Schedule-Kemps.pdf');
            console.log('   - Bandra.html & Bandra.pdf');
            console.log('   - Schedule-Combined.pdf');
            console.log('   - All files uploaded to Google Drive');
            
        } catch (error) {
            console.error('❌ Atomic update failed:', error.message);
            console.log('🔄 Attempting rollback...');
            
            // Restore from backups if they exist
            const kempsBackup = path.join(__dirname, 'Kemps.backup.html');
            const bandraBackup = path.join(__dirname, 'Bandra.backup.html');
            
            if (fs.existsSync(kempsBackup)) {
                fs.copyFileSync(kempsBackup, this.outputPath);
                console.log('   ↳ Restored Kemps.html from backup');
            }
            
            if (fs.existsSync(bandraBackup)) {
                fs.copyFileSync(bandraBackup, path.join(__dirname, 'Bandra.html'));
                console.log('   ↳ Restored Bandra.html from backup');
            }
            
            throw error;
        }
    }

    /**
     * Populate Bandra.html using sheet filtered for Supreme HQ, Bandra and generate Bandra.pdf
     * Leaves Kemps logic untouched.
     */
    async updateBandra() {
        this.currentLocation = 'bandra'; // Set location for theme badge styling
        console.log('\n🚀 Starting Bandra schedule update...');
        // Load all sheet records (does not disturb Kemps filtering already done)  
        const records = await this.readSheet();
        const bandraClasses = records.filter(r => r.Location && /Supreme HQ.*Bandra|Supreme HQ,\s*Bandra/i.test(r.Location));
        console.log(`✅ Found ${bandraClasses.length} classes for Supreme HQ, Bandra`);
        // Temporarily switch context
        const originalHtmlPath = this.htmlPath;
        const originalOutputPath = this.outputPath;
        const bandraHtmlPath = path.join(__dirname, 'Bandra.html');
        
        // Create single backup before updating
        if (fs.existsSync(bandraHtmlPath)) {
            const backupPath = bandraHtmlPath.replace('.html', '.backup.html');
            fs.copyFileSync(bandraHtmlPath, backupPath);
            console.log(`🗂️  Created backup: ${path.basename(backupPath)}`);
        }
        
        this.htmlPath = bandraHtmlPath;
        this.outputPath = bandraHtmlPath; // write in place
        this.kwalityClasses = bandraClasses; // reuse existing downstream logic
        this.readHTML();
        
        // Replace background image with Bandra.png
        const bandraImagePath = path.join(__dirname, 'Bandra.png');
        this.replaceBackgroundImage(bandraImagePath);
        
        this.updatePositionedSpans();
        this.updateScheduleEntries();
        this.updateDateHeaders();
        this.save();
        await this.generatePDFNamed('Bandra.pdf');
        // Upload Bandra.pdf to Drive
        const bandraPdfPath = path.join(__dirname, 'Bandra.pdf');
        await this.uploadNamedPDF(bandraPdfPath, 'Bandra.pdf');
        // Restore original context for safety
        this.htmlPath = originalHtmlPath;
        this.outputPath = originalOutputPath;
        console.log('🎉 Bandra schedule update complete (HTML + Bandra.pdf)');
    }
}

// Main execution
if (require.main === module) {
    // No CSV path needed - everything reads from Google Sheets
    const htmlPath = path.join(__dirname, 'Kemps.html');
    const outputPath = path.join(__dirname, 'Kemps.html');

    const updater = new ScheduleUpdater(htmlPath, outputPath); // No CSV needed
    
    (async () => {
        try {
            console.log('🚀 Starting Google Sheets-only schedule update workflow...');
            console.log('📊 No CSV files will be used - all data from Google Sheets');
            
            // Use the complete Google Sheets workflow
            await updater.completeGoogleSheetsWorkflow();
            
            // Optional: Also update Bandra if needed
            console.log('\n📄 Updating Bandra schedule...');
            await updater.updateBandra();
            
            console.log('\n✨ All tasks completed successfully!');
            console.log('📄 Generated files:');
            console.log('   - Kemps.html (updated from Google Sheets)');
            console.log('   - Kemps_Updated.pdf (uploaded to Drive)');
            console.log('   - Bandra.pdf (if applicable)');
            
        } catch (error) {
            console.error('Failed to update schedule:', error);
            process.exit(1);
        }
    })();
}

module.exports = ScheduleUpdater;
