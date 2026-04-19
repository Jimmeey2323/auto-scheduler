import fs from 'fs';
import http from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { parse } from 'csv-parse/sync';
import cheerio from 'cheerio';
import beautify from 'js-beautify';
import puppeteer from 'puppeteer';
import { google } from 'googleapis';
import axios from 'axios';
import open from 'open';
import { PDFDocument } from 'pdf-lib';
import 'dotenv/config';
import OpenAI from 'openai';

// Enhanced Schedule Mapping System
import EnhancedScheduleMapper from './enhancedScheduleMapper.js';
import { BANDRA_STATIC_STRIP_ASSETS } from './bandraStaticThemeAssets.js';

const beautifyHtml = typeof beautify?.html === 'function' ? beautify.html : (html => html);

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
  SENDER_EMAILS: ["mrigakshi@physique57mumbai.com", "vivaran@physique57mumbai.com"],
  SUBJECT_KEYWORD: "Schedule", // Search for emails with 'Schedule' in subject
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

// Dynamic row management configuration
// Set to true to dynamically add/remove rows based on sheet data
// Set to false to use the existing update logic (preserves HTML row count)
const DYNAMIC_ROW_MODE = true;
const SERVE_TABS_HOST = process.env.SCHEDULE_PREVIEW_HOST || '127.0.0.1';
const SERVE_TABS_OPEN_DELAY_MS = 250;

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function createPreviewHtmlServer(previewFiles, host = SERVE_TABS_HOST, port = 0) {
    const previewFileMap = new Map(
        previewFiles.map(file => [file.routeName, file])
    );

    const server = http.createServer((req, res) => {
        const requestUrl = new URL(req.url || '/', `http://${host}`);

        if (requestUrl.pathname === '/') {
            const linksMarkup = previewFiles.map(file => {
                const href = `/${encodeURIComponent(file.routeName)}`;
                return `<li><a href="${href}" target="_blank" rel="noreferrer">${file.label}</a></li>`;
            }).join('');

            const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Schedule Preview</title>
  <style>
    body {
      margin: 0;
      padding: 32px;
      font: 500 16px/1.6 Inter, system-ui, sans-serif;
      color: #e2e8f0;
      background: linear-gradient(180deg, #0f172a, #020617);
    }
    h1 { margin-top: 0; font-size: 24px; }
    p { color: #94a3b8; }
    ul { padding-left: 20px; }
    a { color: #7dd3fc; }
  </style>
</head>
<body>
  <h1>Generated schedule previews</h1>
  <p>The latest HTML outputs are being served locally. Open either file below in its own tab.</p>
  <ul>${linksMarkup}</ul>
</body>
</html>`;

            res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
            res.end(html);
            return;
        }

        const routeName = decodeURIComponent(requestUrl.pathname.replace(/^\//, ''));
        const previewFile = previewFileMap.get(routeName);

        if (!previewFile) {
            res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            res.end('Preview not found.');
            return;
        }

        if (!fs.existsSync(previewFile.filePath)) {
            res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            res.end(`${previewFile.routeName} does not exist yet.`);
            return;
        }

        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        fs.createReadStream(previewFile.filePath).pipe(res);
    });

    return new Promise((resolve, reject) => {
        server.once('error', reject);
        server.listen(port, host, () => {
            server.off('error', reject);
            resolve(server);
        });
    });
}

async function serveOutputFilesInTabs(previewFiles, options = {}) {
    const host = options.host || SERVE_TABS_HOST;
    const port = Number.isInteger(options.port) ? options.port : 0;
    const server = await createPreviewHtmlServer(previewFiles, host, port);
    const address = server.address();
    const resolvedPort = typeof address === 'object' && address ? address.port : port;
    const baseUrl = `http://${host}:${resolvedPort}`;

    console.log(`\n🌐 Serving generated HTML previews at ${baseUrl}`);
    console.log('📑 Opening each generated HTML file in a separate browser tab...');

    for (const file of previewFiles) {
        const previewUrl = `${baseUrl}/${encodeURIComponent(file.routeName)}`;
        console.log(`   - ${file.label}: ${previewUrl}`);
        await open(previewUrl, { wait: false, newInstance: false });
        await sleep(SERVE_TABS_OPEN_DELAY_MS);
    }

    console.log('🛑 Press Ctrl+C when you are done previewing to stop the local preview server.');

    const closeServer = () => {
        if (!server.listening) {
            return;
        }

        console.log('\n🧹 Shutting down preview server...');
        server.close(() => process.exit(0));
    };

    process.once('SIGINT', closeServer);
    process.once('SIGTERM', closeServer);

    return { server, baseUrl };
}

/**
 * Advanced Node.js Script to Update Kemps.html with CSV Data
 * Reads class data from CSV and updates HTML in accurate positions
 * without altering styling, layout or structure
 * Generates PDF and uploads to Google Drive
 */

class ScheduleUpdater {
    constructor(htmlPath, outputPath, location = 'kemps', options = {}) {
        this.htmlPath = htmlPath;
        this.outputPath = outputPath || htmlPath;
        this.kwalityClasses = [];
        this.allSheetRecords = [];
        this.$ = null;
        this.currentLocation = location.toLowerCase(); // Track current location for theme badge styling
        this.locationName = this.currentLocation.charAt(0).toUpperCase() + this.currentLocation.slice(1); // 'Kemps' or 'Bandra'
        this.themeRenderMode = options.themeRenderMode === 'static' ? 'static' : 'badge';
        this.skipPdf = options.skipPdf === true; // Skip PDF generation if set (for static deployments)
        this.staticThemeRows = [];
        this.staticThemeColorMap = new Map();
        this.staticThemeLegendAssetMap = new Map();
        this.googleAccessTokenCache = null;
        this.googleAccessTokenExpiry = 0;
        this.googleAccessTokenPromise = null;
        this.currentScheduleSubject = '';
        this.currentScheduleWeek = null;
        
        // Initialize enhanced mapping system
        this.enhancedMapper = new EnhancedScheduleMapper();
        
        // Initialize AI parser
        this.initializeAIParser();
    }

    /**
     * Initialize OpenAI client for AI-powered email parsing
     */
    initializeAIParser() {
        if (process.env.OPENAI_API_KEY) {
            this.openai = new OpenAI({
                apiKey: process.env.OPENAI_API_KEY
            });
            this.aiModel = process.env.OPENAI_MODEL || 'gpt-4o-mini';
            this.aiEnabled = true;
            console.log('✅ AI Email Parser initialized with model:', this.aiModel);
        } else {
            this.aiEnabled = false;
            console.log('ℹ️  AI Email Parser disabled (OPENAI_API_KEY not set)');
        }
    }

    /**
     * Parse email thread using AI (GPT-4)
     */
    async parseEmailWithAI(emailMessages, isFirstEmail = false) {
        if (!this.aiEnabled) {
            return null;
        }

        console.log('\n🤖 AI Email Parser: Analyzing email thread...');
        console.log(`📧 Processing ${emailMessages.length} message(s)`);

        try {
            const fullThread = emailMessages.join('\n\n---MESSAGE SEPARATOR---\n\n');
            
            // Parse different types of information in parallel
            const [covers, themes, changes] = await Promise.all([
                isFirstEmail ? Promise.resolve([]) : this.aiParseCovers(fullThread),
                this.aiParseThemes(fullThread),
                this.aiParseChanges(fullThread)
            ]);

            const result = {
                covers: covers || [],
                themes: themes || [],
                changes: changes || [],
                aiParsed: true
            };

            console.log(`✅ AI Parsing Complete:`);
            console.log(`   📝 Covers: ${result.covers.length}`);
            console.log(`   🎨 Themes: ${result.themes.length}`);
            console.log(`   🔄 Changes: ${result.changes.length}`);

            return result;

        } catch (error) {
            console.error('❌ AI parsing error:', error.message);
            return null;
        }
    }

    /**
     * AI: Extract cover information
     */
    async aiParseCovers(emailContent) {
        const prompt = `Extract class cover/substitution information from this email thread.

A cover is when one trainer substitutes for another trainer.

Return a JSON object with this structure:
{
  "covers": [
    {
      "day": "Monday",
      "time": "7:30 AM",
      "location": "Kemps",
      "class": "Barre 57",
      "originalTrainer": "Pranjali",
      "coverTrainer": "Rohan"
    }
  ]
}

If no covers found, return: {"covers": []}

Known locations: Kemps, Bandra, Annex, Kwality House, Supreme HQ
Known trainers: Mrigakshi, Anisha, Pranjali, Richard, Rohan, Karan, Simonelle, Reshma, Cauveri, Vivaran, Atulan, Raunak, Bret, Anmol, Simran

Email thread:
${emailContent.substring(0, 8000)}

Return ONLY valid JSON, no other text.`;

        try {
            const response = await this.openai.chat.completions.create({
                model: this.aiModel,
                messages: [
                    { role: 'system', content: 'You are a precise schedule data extractor. Always return valid JSON.' },
                    { role: 'user', content: prompt }
                ],
                temperature: 0.1,
                response_format: { type: 'json_object' }
            });

            const parsed = JSON.parse(response.choices[0].message.content);
            return parsed.covers || [];
        } catch (error) {
            console.error('   ❌ AI cover parsing failed:', error.message);
            return [];
        }
    }

    /**
     * AI: Extract theme information
     */
    async aiParseThemes(emailContent) {
        const prompt = `Extract ALL fitness class themes from this email thread including PowerCycle, Amped Up, FIT, and Bandra cycle themes.

Look for patterns like:
- "Power Cycle themes:" followed by location sections (Bandra, Kemps)
- "Amped Up theme:" with day-theme pairs
- "FIT Theme:" with weekly themes
- "Bandra cycle themes:" with day-theme pairs

Return a JSON object with this structure:
{
  "themes": [
    {
      "day": "Monday",
      "time": "8:00 AM",
      "location": "Kemps",
      "classType": "PowerCycle",
      "theme": "Lady Gaga vs Bruno Mars"
    },
    {
      "day": "Tuesday",
      "time": "",
      "location": "Kemps",
      "classType": "Amped Up",
      "theme": "Heart Rate & Heartbreak"
    }
  ]
}

If no themes found, return: {"themes": []}

Common theme names: Lady Gaga vs Bruno Mars, Rihanna + Friends, Teen Crush, Love Pop, Heart Rate & Heartbreak, Taylor Swift, Super Sets, Tabata
Known class types: PowerCycle, Amped Up, CYCLE, FIT
Known locations: Kemps, Bandra, Annex, Kwality House, Supreme HQ

Email thread:
${emailContent.substring(0, 12000)}

Return ONLY valid JSON, no other text.`;

        try {
            const response = await this.openai.chat.completions.create({
                model: this.aiModel,
                messages: [
                    { role: 'system', content: 'You are a precise schedule data extractor. Always return valid JSON.' },
                    { role: 'user', content: prompt }
                ],
                temperature: 0.1,
                response_format: { type: 'json_object' }
            });

            const parsed = JSON.parse(response.choices[0].message.content);
            const themes = parsed.themes || [];
            
            if (themes.length > 0) {
                console.log('\n📋 AI EXTRACTED THEMES:');
                console.log('='.repeat(80));
                themes.forEach((theme, idx) => {
                    const timeStr = theme.time ? ` ${theme.time}` : '';
                    const locationStr = theme.location ? ` at ${theme.location}` : '';
                    console.log(`${idx + 1}. [${theme.classType}] ${theme.day}${timeStr}${locationStr}: "${theme.theme}"`);
                });
                console.log('='.repeat(80) + '\n');
            } else {
                console.log('\n⚠️  AI found no themes in email thread\n');
            }
            
            return themes;
        } catch (error) {
            console.error('   ❌ AI theme parsing failed:', error.message);
            return [];
        }
    }

    /**
     * AI: Extract schedule changes (sold out, cancellations, time changes)
     */
    async aiParseChanges(emailContent) {
        const prompt = `Extract schedule changes from this email thread.

Look for:
- Classes marked as "sold out"
- Cancellations
- Time changes
- Location changes
- Trainer changes (permanent, not covers)

Return a JSON object:
{
  "changes": [
    {
      "type": "sold_out",
      "day": "Monday",
      "time": "7:30 AM",
      "location": "Kemps",
      "class": "Barre 57",
      "description": "Class marked as sold out"
    }
  ]
}

Valid types: sold_out, cancellation, time_change, location_change, trainer_change

If no changes found, return: {"changes": []}

Email thread:
${emailContent.substring(0, 8000)}

Return ONLY valid JSON, no other text.`;

        try {
            const response = await this.openai.chat.completions.create({
                model: this.aiModel,
                messages: [
                    { role: 'system', content: 'You are a precise schedule data extractor. Always return valid JSON.' },
                    { role: 'user', content: prompt }
                ],
                temperature: 0.1,
                response_format: { type: 'json_object' }
            });

            const parsed = JSON.parse(response.choices[0].message.content);
            
            // Log changes for visibility
            if (parsed.changes && parsed.changes.length > 0) {
                console.log('   🔍 AI detected changes:');
                parsed.changes.forEach((change, idx) => {
                    console.log(`      ${idx + 1}. [${change.type}] ${change.day} ${change.time} - ${change.description}`);
                });
            }
            
            return parsed.changes || [];
        } catch (error) {
            console.error('   ❌ AI change parsing failed:', error.message);
            return [];
        }
    }

    /**
     * Apply AI-detected changes to schedule data
     */
    applyAIChangesToSchedule(scheduleData, changes) {
        if (!changes || changes.length === 0) {
            return scheduleData;
        }

        console.log(`\n🔄 Applying ${changes.length} AI-detected changes to schedule...`);
        const updatedData = [...scheduleData];
        let appliedCount = 0;

        for (const change of changes) {
            // Find matching class
            const matchIndex = updatedData.findIndex(item => {
                const dayMatch = item.Day && item.Day.toLowerCase().includes(change.day.toLowerCase());
                const timeMatch = this.normalizeTime(item.Time || '') === this.normalizeTime(change.time || '');
                const locationMatch = !change.location || 
                    (item.Location && item.Location.toLowerCase().includes(change.location.toLowerCase()));
                
                return dayMatch && timeMatch && locationMatch;
            });

            if (matchIndex >= 0) {
                const item = updatedData[matchIndex];
                
                switch (change.type) {
                    case 'sold_out':
                        item.Theme = 'Sold Out';
                        console.log(`   ✅ Marked sold out: ${change.day} ${change.time} ${change.class}`);
                        appliedCount++;
                        break;
                        
                    case 'cancellation':
                        item.Notes = (item.Notes || '') + ' CANCELLED';
                        console.log(`   ✅ Marked cancelled: ${change.day} ${change.time} ${change.class}`);
                        appliedCount++;
                        break;
                }
            }
        }

        console.log(`✅ Successfully applied ${appliedCount}/${changes.length} changes`);
        return updatedData;
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
     * Format class name with proper casing:
     * - PowerCycle displays as 'powerCycle'
     * - All other classes display in UPPERCASE
     */
    formatClassName(className) {
        if (!className) return '';
        const lower = className.toLowerCase().trim();
        
        // PowerCycle special case - display as 'powerCycle'
        if (lower.includes('powercycle')) {
            let formatted = lower.replace(/powercycle/g, 'powerCycle');
            // Handle Express suffix
            formatted = formatted.replace(/express/g, 'Express');
            // Capitalize 'Studio' if present
            formatted = formatted.replace(/^studio\s+/i, 'STUDIO ');
            return formatted;
        }
        
        // All other classes in UPPERCASE
        return className.toUpperCase();
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

            // Filter classes based on current location
            let filteredClasses;
            if (this.currentLocation === 'kemps') {
                filteredClasses = records.filter(record =>
                    record.Location && record.Location.includes('Kwality House')
                );
                console.log(`✅ Found ${filteredClasses.length} classes for Kwality House from Cleaned sheet`);
            } else if (this.currentLocation === 'bandra') {
                filteredClasses = records.filter(record =>
                    record.Location && /Supreme HQ.*Bandra|Supreme HQ,\s*Bandra/i.test(record.Location)
                );
                console.log(`✅ Found ${filteredClasses.length} classes for Supreme HQ, Bandra from Cleaned sheet`);
            } else {
                // Default fallback
                filteredClasses = records.filter(record =>
                    record.Location && record.Location.includes('Kwality House')
                );
                console.log(`⚠️  Unknown location '${this.currentLocation}', defaulting to Kwality House. Found ${filteredClasses.length} classes`);
            }
            
            // Assign filtered classes to instance variable
            this.kwalityClasses = filteredClasses;
            
            // Store all records for other purposes
            this.allSheetRecords = records;

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
            
            // Step 2: Extract Google Sheets link from the selected newest email only.
            // Do not scan older thread messages, otherwise previous-week links can win.
            console.log('🔗 Step 2: Extracting Google Sheets link from the newest selected email only...');
            const sheetsLink = this.extractSheetsLink(emailData.body);
            
            if (!sheetsLink) {
                console.log('⚠️  No Google Sheets link found in the newest selected email');
                console.log('🔍 Latest message preview:', emailData.body.substring(0, 500));
                return;
            }

            console.log('✅ Found Google Sheets link:', sheetsLink);
            this.currentSourceSheetsLink = sheetsLink;

            // Step 3: Extract schedule data from the linked spreadsheet
            console.log('📋 Step 3: Extracting data from linked spreadsheet...');
            const scheduleData = await this.fetchDataFromLinkedSheet(sheetsLink);
            
            if (!scheduleData || scheduleData.length === 0) {
                console.log('❌ No schedule data retrieved from linked sheet');
                return;
            }

            console.log(`✅ Retrieved ${scheduleData.length} schedule records from linked sheet`);
            
            // Step 4: Determine if this is first email (initial schedule) or subsequent email (changes)
            console.log('🎨 Step 4: Parsing email for covers and themes...');
            
            // Determine email type based only on the selected newest email.
            const isFirstEmail = this.hasSpreadsheetLink(emailData.body);
            
            console.log(`📧 Email type: ${isFirstEmail ? 'FIRST EMAIL (using spreadsheet covers only)' : 'SUBSEQUENT EMAIL (using email body covers)'}`);

            const emailMessagesToParse = emailData.body ? [emailData.body] : [];
            const emailInfo = await this.parseEmailForScheduleInfo(emailMessagesToParse, isFirstEmail);
            
            // Add the email subject to emailInfo for date calculation
            emailInfo.subject = emailData.subject;
            
            console.log(`✅ Parsed ${emailInfo.covers.length} covers and ${emailInfo.themes.length} themes from email`);
            
            // Log covers from email body (only if subsequent email)
            if (!isFirstEmail && emailInfo.covers.length > 0) {
                console.log('\n📧 ===== COVERS FROM EMAIL BODY =====');
                this.logEmailCovers(emailInfo.covers);
                console.log('====================================\n');
            } else if (isFirstEmail) {
                console.log('\nℹ️  First email detected - using covers from spreadsheet only, ignoring email body covers\n');
            }

            // Step 5: Update target spreadsheet with combined data
            console.log('📊 Step 5: Updating target spreadsheet...');
            await this.updateTargetSpreadsheet(scheduleData, emailInfo, isFirstEmail);
            
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
     * Get the current week's date range for email search - Enhanced Version
     */
    getCurrentWeekDateRange() {
        console.log('📅 Calculating week date range with enhanced algorithm...');
        const result = this.enhancedMapper.calculateWeekendDates();
        
        return {
            monday: result.monday,
            sunday: result.sunday,
            saturday: result.saturday
        };
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
     * Check whether a subject line looks like the weekly Mumbai schedule email.
     */
    isWeeklyMumbaiScheduleSubject(subject = '') {
        const normalized = String(subject).trim().toLowerCase();
        if (!normalized) return false;

        return normalized.includes('mumbai schedule') || /\b\d{1,2}\s*-\s*\d{1,2}(?:st|nd|rd|th)?\s+[a-z]{3}'?\d{2}\b/i.test(subject);
    }

    /**
     * Score a candidate schedule email subject so weekly schedule threads outrank
     * quarterly schedules, meetings, and other noisy schedule-related emails.
     */
    scoreScheduleEmailSubject(subject = '') {
        const normalized = String(subject).trim().toLowerCase();
        if (!normalized) return -1000;

        const currentWeekPattern = this.getCurrentWeekSubjectPattern().toLowerCase();
        let score = 0;

        if (normalized.includes(currentWeekPattern)) score += 1000;
        if (normalized.includes('mumbai schedule')) score += 500;
        if (this.isWeeklyMumbaiScheduleSubject(subject)) score += 250;
        if (/\b(re:|fwd:)\b/i.test(subject)) score += 25;

        if (normalized.includes('quarterly')) score -= 1000;
        if (normalized.includes('trail')) score -= 500;
        if (normalized.includes('meeting')) score -= 500;
        if (normalized.includes('invitation')) score -= 500;
        if (normalized.includes('schedule format')) score -= 250;

        return score;
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
        console.log(`🔍 Searching for the most recent schedule email...`);
        
        try {
            const accessToken = await this.getAccessToken();
            const oauth2Client = new google.auth.OAuth2(
                GOOGLE_CONFIG.CLIENT_ID,
                GOOGLE_CONFIG.CLIENT_SECRET
            );
            oauth2Client.setCredentials({ access_token: accessToken });

            const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
            
            // Search for schedule emails from approved senders.
            // We always use the newest matching email received.
            const senderQuery = EMAIL_CONFIG.SENDER_EMAILS.map(e => `from:${e}`).join(' OR ');
            const searchQuery = `(${senderQuery}) subject:Schedule newer_than:21d`;
            console.log(`🔍 Email search query: ${searchQuery}`);
            
            let response = await gmail.users.messages.list({
                userId: 'me',
                q: searchQuery,
                maxResults: 10
            });
            
            if (!response.data.messages || response.data.messages.length === 0) {
                console.log('❌ No schedule emails found matching criteria');
                return null;
            }

            console.log(`📬 Found ${response.data.messages.length} potential emails`);

            // Get metadata so we can deterministically pick the newest matching email.
            const messagesWithDates = await Promise.all(
                response.data.messages.map(async (msg) => {
                    const detail = await gmail.users.messages.get({
                        userId: 'me',
                        id: msg.id,
                        format: 'metadata',
                        metadataHeaders: ['Subject', 'Date']
                    });
                    const subject = detail.data.payload.headers.find(h => h.name === 'Subject')?.value;
                    const date = detail.data.payload.headers.find(h => h.name === 'Date')?.value;
                    const parsedWeek = this.extractWeekFromEmailSubject(subject);
                    return {
                        id: msg.id,
                        subject: subject,
                        date: date,
                        internalDate: detail.data.internalDate,
                        parsedWeek,
                        scheduleWeekEndTs: parsedWeek?.endDate ? new Date(parsedWeek.endDate).getTime() : Number.NEGATIVE_INFINITY,
                        isMumbaiSchedule: /mumbai\s+schedule/i.test(subject || '')
                    };
                })
            );

            const preferredMessages = messagesWithDates.filter(message => message.isMumbaiSchedule);
            const candidateMessages = preferredMessages.length > 0 ? preferredMessages : messagesWithDates;

            if (preferredMessages.length > 0) {
                console.log(`📬 Narrowed candidates to ${preferredMessages.length} Mumbai schedule email(s)`);
            } else {
                console.log('⚠️  No explicit "Mumbai schedule" subjects found, falling back to all Schedule matches');
            }

            console.log('📬 Candidate schedule emails sorted by schedule week, then received time:');
            candidateMessages
                .slice()
                .sort((a, b) => {
                    if (b.scheduleWeekEndTs !== a.scheduleWeekEndTs) {
                        return b.scheduleWeekEndTs - a.scheduleWeekEndTs;
                    }

                    return parseInt(b.internalDate) - parseInt(a.internalDate);
                })
                .forEach((message, index) => {
                    const parsedRange = message.parsedWeek?.weekFound
                        ? `${new Date(message.parsedWeek.startDate).toDateString()} → ${new Date(message.parsedWeek.endDate).toDateString()}`
                        : 'no subject week parsed';
                    console.log(`   ${index + 1}. ${message.date || 'Unknown date'} | ${message.subject} | ${parsedRange}`);
                });

            // Prefer the email for the latest schedule week; break ties by received time.
            candidateMessages.sort((a, b) => {
                if (b.scheduleWeekEndTs !== a.scheduleWeekEndTs) {
                    return b.scheduleWeekEndTs - a.scheduleWeekEndTs;
                }

                return parseInt(b.internalDate) - parseInt(a.internalDate);
            });
            
            const mostRecentMessage = candidateMessages[0];
            console.log(`\n✅ Using selected schedule email: "${mostRecentMessage.subject}"\n`);
            
            // Get the most recent email (already sorted by date)
            const messageId = mostRecentMessage.id;
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
                date: emailDate,
                id: messageId,
                threadId: threadId
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
        const collectedParts = [];

        const normalizeBase64 = (data) => {
            if (!data) return '';
            const normalized = data.replace(/-/g, '+').replace(/_/g, '/');
            const paddingNeeded = normalized.length % 4;
            return paddingNeeded ? normalized + '='.repeat(4 - paddingNeeded) : normalized;
        };

        const decodePart = (data) => {
            if (!data) return '';
            try {
                return Buffer.from(normalizeBase64(data), 'base64').toString('utf-8');
            } catch (err) {
                console.warn('⚠️  Failed to decode email segment:', err.message);
                return '';
            }
        };

        const htmlToText = (html) => {
            if (!html) return '';
            let text = html;

            // Preserve natural line breaks from common block elements before stripping tags
            text = text.replace(/<\s*(br|\/p|\/div|\/li|\/tr)\b[^>]*>/gi, '\n');
            text = text.replace(/<script[\s\S]*?<\/script>/gi, '');
            text = text.replace(/<style[\s\S]*?<\/style>/gi, '');
            text = text.replace(/<[^>]+>/g, ' ');

            // Decode a few common HTML entities (enough for schedule parsing)
            text = text
                .replace(/&nbsp;/gi, ' ')
                .replace(/&amp;/gi, '&')
                .replace(/&quot;/gi, '"')
                .replace(/&#39;/gi, "'")
                .replace(/&lt;/gi, '<')
                .replace(/&gt;/gi, '>');

            // Normalise whitespace but preserve deliberate line breaks
            text = text.replace(/\r\n/g, '\n');
            text = text.replace(/[ \t]+\n/g, '\n').replace(/\n[ \t]+/g, '\n');
            text = text.replace(/\n{3,}/g, '\n\n');
            text = text.replace(/[ \t]{2,}/g, ' ');
            return text.trim();
        };

        const collectParts = (part) => {
            if (!part) return;
            const mimeType = part.mimeType || '';
            const data = part.body?.data;

            if (mimeType === 'text/plain' && data) {
                collectedParts.push({ type: 'text/plain', content: decodePart(data) });
                return;
            }

            if (mimeType === 'text/html' && data) {
                collectedParts.push({ type: 'text/html', content: htmlToText(decodePart(data)) });
                return;
            }

            if (part.parts && part.parts.length) {
                part.parts.forEach(collectParts);
            }
        };

        if (messageData?.payload) {
            collectParts(messageData.payload);
        }

        // Fallback: sometimes the payload body is populated even without parts
        if (collectedParts.length === 0 && messageData?.payload?.body?.data) {
            collectedParts.push({ type: 'text/plain', content: decodePart(messageData.payload.body.data) });
        }

        const plainTextSegments = collectedParts.filter(part => part.type === 'text/plain' && part.content.trim());
        const htmlSegments = collectedParts.filter(part => part.type === 'text/html' && part.content.trim());

        // Prefer rendered HTML text when available since it contains the full weekly bulletin
        let combined = '';
        if (htmlSegments.length) {
            combined = htmlSegments.map(part => part.content).join('\n');
        } else if (plainTextSegments.length) {
            combined = plainTextSegments.map(part => part.content).join('\n');
        } else {
            combined = collectedParts.map(part => part.content).join('\n');
        }

        return combined
            .replace(/\r\n/g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
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
        console.log(`📄 Email body preview (first 300 chars): ${emailBody.substring(0, 300)}...`);
        
        // Look for Google Sheets URLs
        const sheetsRegex = /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/g;
        const matches = emailBody.match(sheetsRegex);
        
        if (matches && matches.length > 0) {
            console.log(`✅ Found Google Sheets link: ${matches[0]}`);
            console.log(`📊 Spreadsheet ID: ${matches[0].match(/\/d\/([a-zA-Z0-9-_]+)/)[1]}`);
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
     * @param {Array} allMessages - Array of email message bodies
     * @param {Boolean} isFirstEmail - If true, only extract themes (covers come from spreadsheet)
     */
    async parseEmailForScheduleInfo(allMessages, isFirstEmail = true) {
        console.log('🔍 Parsing email content for schedule information...');
        
        // Use AI parsing exclusively for theme detection if enabled
        if (this.aiEnabled) {
            try {
                const aiResult = await this.parseEmailWithAI(allMessages, isFirstEmail);
                
                if (aiResult && aiResult.aiParsed) {
                    console.log('✅ Using AI-parsed results EXCLUSIVELY for themes');
                    console.log(`🎨 AI extracted ${aiResult.themes.length} themes`);
                    
                    // Only get covers and hosted classes from regex if needed
                    const regexResult = await this.parseEmailForScheduleInfoRegex(allMessages, isFirstEmail, true); // Skip themes
                    
                    return {
                        covers: aiResult.covers.length > 0 ? aiResult.covers : regexResult.covers,
                        themes: aiResult.themes, // AI themes ONLY - no regex themes
                        changes: aiResult.changes,
                        hostedClasses: regexResult.hostedClasses,
                        aiParsed: true
                    };
                }
            } catch (error) {
                console.warn('⚠️  AI parsing failed, using regex fallback:', error.message);
            }
        }

        // Fallback to regex-based parsing
        return this.parseEmailForScheduleInfoRegex(allMessages, isFirstEmail);
    }

    /**
     * Regex-based email parsing (original implementation)
     * @param {Array} allMessages - Array of email message bodies
     * @param {Boolean} isFirstEmail - If true, only extract themes (covers come from spreadsheet)
     */
    async parseEmailForScheduleInfoRegex(allMessages, isFirstEmail = true, skipThemes = false) {
        console.log('🔍 Parsing email content for schedule information...');
        console.log(`📧 Total messages in thread: ${(allMessages || []).length}`);
        if (skipThemes) {
            console.log('🎨 Skipping theme extraction - AI handling themes exclusively');
        }
        
        const result = {
            covers: [],
            themes: [],
            hostedClasses: []
        };
        
        // CRITICAL: Parse covers and themes from most recent SCHEDULE EMAIL THREAD
        // We want the most recent schedule email thread, but themes/covers could be in ANY message within that thread
        let fullContent = '';
        let mostRecentMessageIndex = -1;
        
        if (allMessages && allMessages.length > 0) {
            // For covers parsing, still use the most recent message in the thread
            mostRecentMessageIndex = allMessages.length - 1;
            fullContent = allMessages[mostRecentMessageIndex].replace(/\r\n/g, '\n');
            
            console.log('\n' + '='.repeat(80));
            console.log('📧 USING MOST RECENT SCHEDULE EMAIL THREAD');
            console.log(`📧 Thread has ${allMessages.length} messages total`);
            console.log(`📧 Most recent message (#${mostRecentMessageIndex + 1}) content length: ${fullContent.length}`);
            console.log('📧 Most recent message preview:', fullContent.substring(0, 400));
            console.log('='.repeat(80) + '\n');
        } else {
            console.log('⚠️  No messages found in thread');
            return result;
        }
        
        // Parse covers section ONLY if this is NOT the first email
        // First email: covers come from spreadsheet
        // Subsequent emails: covers come from email body (changes/updates)
        if (!isFirstEmail) {
            console.log('\n' + '='.repeat(80));
            console.log('🔍 PARSING COVERS FROM MOST RECENT MESSAGE (subsequent email)');
            console.log('='.repeat(80) + '\n');
            
            // Parse covers section - improved regex to capture entire section
            // The covers section typically goes from "Covers :" until the next major section like "Amped Up theme" or "Bandra cycle themes"
            const coversMatch = fullContent.match(/Covers\s*:?\s*([\s\S]*?)(?=\n\s*(?:Amped Up theme|Bandra cycle themes|FIT theme|Best,\s*$))/i);
            if (coversMatch) {
                console.log('🎯 Found covers section, length:', coversMatch[1].length);
                console.log('🎯 Covers preview:', coversMatch[1].substring(0, 300));
                result.covers = this.parseCoversSection(coversMatch[1]);
            } else {
                console.log('❌ No covers section found in most recent message');
                console.log('🔍 Looking for alternative covers pattern...');
                
                // Try alternative pattern - capture everything between "Covers" and next section
                const altCoversMatch = fullContent.match(/Covers\s*:?\s*([\s\S]*?)(?=\n\s*(?:Amped|Bandra cycle|FIT theme|Best))/i);
                if (altCoversMatch) {
                    console.log('🎯 Found alternative covers section:', altCoversMatch[1].substring(0, 200));
                    result.covers = this.parseCoversSection(altCoversMatch[1]);
                }
            }
        } else {
            console.log('\nℹ️  Skipping email body covers parsing (first email - using spreadsheet covers only)\n');
        }
        
        // Parse themes sections - using simpler approach to avoid matchAll issues
        const themeSections = [];
        
        // Look for theme patterns individually with more precise boundaries
        const amped_theme = fullContent.match(/Amped Up theme\s*\*?\s*:\s*(.*?)(?=\n\s*(?:FIT\s*Theme|Bandra cycle themes|Best,|$))/is);
        if (amped_theme) {
            // Clean up the captured text to remove any trailing content
            let ampedContent = amped_theme[1].trim();
            // Remove any text that starts with "Bandra cycle themes", "FIT Theme" or similar
            ampedContent = ampedContent.split(/Bandra cycle themes|FIT\s*Theme/i)[0].trim();
            themeSections.push({ type: 'Amped Up', content: ampedContent });
        }
        
        // Match Bandra cycle themes section, stopping at common email reply markers
        // This prevents capturing content from email replies in the thread
        const bandra_themes = fullContent.match(/Bandra cycle themes\s*[-–:]\s*\*?\s*(.*?)(?=\n(?:Best,|Thanks\s+and\s+regards|Warm\s+Regards|Regards,|On\s+(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)|\n--\s*\n))/is);
        if (bandra_themes) {
            // Also truncate at any line that looks like email metadata
            let bandraContent = bandra_themes[1];
            // Remove content after "Thanks and regards" or email signatures
            bandraContent = bandraContent.split(/Thanks\s+and\s+regards|Warm\s+Regards|Regards,/i)[0].trim();
            themeSections.push({ type: 'Bandra cycle', content: bandraContent });
        }
        
        // FIT Theme patterns: "FIT Theme : ALL classes of the week : SUPER SETS" or "FIT theme: All classes, all week - TABATA"
        const fit_theme = fullContent.match(/FIT\s*Theme\s*:\s*(?:ALL\s*classes\s*(?:of\s*the\s*week|,\s*all\s*week))\s*[:\-–]\s*(.+?)(?=\n|$)/i);
        if (fit_theme) {
            const fitThemeName = fit_theme[1].trim();
            console.log(`🏋️ Found FIT theme: ${fitThemeName}`);
            themeSections.push({ type: 'FIT', content: `All classes, all week - ${fitThemeName}` });
        }
        
        // PowerCycle Themes: Skip regex extraction if AI is handling themes
        if (!skipThemes) {
            let extractedPowerCycleThemes = [];
            
            console.log('\n' + '='.repeat(80));
            console.log('🚴 EXTRACTING POWERCYCLE THEMES FROM ALL MESSAGES IN MOST RECENT THREAD (REGEX)');
            console.log('='.repeat(80));
            
            // CORRECTED: Search through ALL messages in the most recent thread
            // The thread itself is the most recent, but themes could be in any message within it
            console.log(`📧 Analyzing ALL ${allMessages.length} messages in most recent thread for PowerCycle themes:`);
            
            allMessages.forEach((message, idx) => {
                console.log(`   Message ${idx + 1}: ${message.substring(0, 100).replace(/\n/g, ' ')}...`);
            });
            console.log('');
            
            // Use enhanced mapper with ALL messages from the most recent thread
            extractedPowerCycleThemes = this.enhancedMapper.extractPowerCycleThemes(allMessages);
            
            console.log(`\n🔍 PowerCycle extraction result: ${extractedPowerCycleThemes.length} themes found`);
            
            if (extractedPowerCycleThemes.length > 0) {
                console.log('\n📋 EXTRACTED POWERCYCLE THEMES FROM MOST RECENT EMAIL THREAD:');
                console.log('='.repeat(80));
                extractedPowerCycleThemes.forEach((theme, idx) => {
                    console.log(`${idx + 1}. ${theme.day} ${theme.time} at ${theme.location}: "${theme.theme}"`);
                });
                console.log('='.repeat(80) + '\n');
            } else {
                console.log('\n⚠️  NO PowerCycle themes found in most recent email thread');
                console.log('🔍 Falling back to DEFAULT themes (this may be incorrect!)\n');
                extractedPowerCycleThemes = this.enhancedMapper.getDefaultPowerCycleThemes();
                
                console.log('\n📋 USING DEFAULT POWERCYCLE THEMES:');
                console.log('='.repeat(80));
                extractedPowerCycleThemes.forEach((theme, idx) => {
                    console.log(`${idx + 1}. ${theme.day} ${theme.time} at ${theme.location}: "${theme.theme}"`);
                });
                console.log('='.repeat(80) + '\n');
            }
            
            // Store extracted themes for later mapping
            this.extractedPowerCycleThemes = extractedPowerCycleThemes;
            
            // Add PowerCycle themes directly to result (they're already parsed by enhancedMapper)
            if (extractedPowerCycleThemes.length > 0) {
                result.themes.push(...extractedPowerCycleThemes);
            }
        } else {
            console.log('\n🎨 Skipping regex PowerCycle theme extraction - AI is handling themes');
        }
        
        console.log(`📊 Found ${themeSections.length} theme sections`);
        
        if (!skipThemes) {
            for (const section of themeSections) {
                console.log(`🎨 Processing ${section.type} themes`);
                const themes = this.parseThemesSection(section.type, section.content);
                result.themes.push(...themes);
            }
        } else {
            console.log('🎨 Skipping regex theme section processing - AI is handling themes');
        }
        
        // Deduplicate themes only if not skipping themes (AI handles its own deduplication)
        if (!skipThemes) {
            const seenThemes = new Set();
            result.themes = result.themes.filter(theme => {
                const key = `${theme.day}|${theme.time || ''}|${theme.theme}|${theme.classType || ''}`;
                if (seenThemes.has(key)) {
                    return false; // Skip duplicate
                }
                seenThemes.add(key);
                return true;
            });
            
            console.log(`📊 After deduplication: ${result.themes.length} unique themes`);
        }
        
        // Parse hosted classes - handle various formats like "- Hosted Classes -" or "- Hosted  Classes -"
        const hostedMatch = fullContent.match(/-\s*Hosted\s+Classes\s*-\s*(.*?)(?=\n\s*(?:Covers|Amped|Bandra|FIT|Best|Thanks)|$)/is);
        if (hostedMatch) {
            console.log('🏢 Found hosted classes section');
            console.log('🏢 Hosted section content:', hostedMatch[1].substring(0, 300));
            result.hostedClasses = this.parseHostedClasses(hostedMatch[1]);
        } else {
            console.log('⚠️ No hosted classes section found');
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
                // Check if it's just a day marker (sets context for following lines)
                if (coverInfo.isDayMarker) {
                    console.log(`📅 Set day context to: ${coverInfo.day}`);
                    previousDay = coverInfo.day;
                } else {
                    console.log(`✅ Parsed cover:`, coverInfo);
                    covers.push(coverInfo);
                    // Update previousDay for potential continuation lines
                    previousDay = coverInfo.day;
                }
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
        // Pattern 1: Day with times but no trainer (sets context for next lines)
        // Example: "Wed - 8 am, 9.15 am" or "Thurs - 7.30,9, 11 am"
        // This pattern matches: Day - <anything that looks like times with am/pm but no dash after>
        const dayOnlyPattern = /^([A-Za-z]+)\s*-\s*([^-]+(?:am|pm)[^-]*)$/i;
        const dayOnlyMatch = line.match(dayOnlyPattern);
        
        if (dayOnlyMatch && !dayOnlyMatch[2].includes('-')) {
            // Verify it has time-like patterns (numbers with am/pm)
            if (/\d.*(?:am|pm)/i.test(dayOnlyMatch[2])) {
                const day = this.expandDayName(dayOnlyMatch[1].trim());
                console.log(`📅 Found day declaration (no trainer): ${day} with times: ${dayOnlyMatch[2].trim()}`);
                // Return a marker object to update previousDay
                return {
                    isDayMarker: true,
                    day: day
                };
            }
        }
        
        // Pattern 2: Day - time(s) - trainer
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
                // Regular time-based cover with class types
                const timeArray = Array.isArray(timeInfo) ? timeInfo : [timeInfo];
                
                return {
                    location: location,
                    day: this.expandDayName(day),
                    timesWithClasses: timeArray, // Array of {time, classType}
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
            const timeArray = Array.isArray(timeInfo) ? timeInfo : [timeInfo];
            
            return {
                location: location,
                day: previousDay,
                timesWithClasses: timeArray, // Array of {time, classType}
                trainer: trainer,
                type: 'cover'
            };
        }
        
        // Try pattern for descriptive continuation lines like "Evening cycles - Raunak", "Evening Barre classes - Pranjali"
        const descriptiveContinuationPattern = /^((?:morning|evening|afternoon)\s+(?:cycle|cycles|barre|barre classes|classes))\s*-\s*(.+)$/i;
        const descriptiveMatch = line.match(descriptiveContinuationPattern);
        
        if (descriptiveMatch && previousDay) {
            const description = descriptiveMatch[1].trim();
            const trainer = descriptiveMatch[2].trim();
            
            const timeInfo = this.parseTimeText(description);
            
            if (timeInfo.timePattern) {
                console.log(`✅ Parsed descriptive continuation cover: ${description} -> ${trainer} for ${previousDay}`);
                return {
                    location: location,
                    day: previousDay,
                    timePattern: timeInfo.timePattern,
                    classType: timeInfo.classType,
                    trainer: trainer,
                    type: 'cover'
                };
            }
        }
        
        return null;
    }

    /**
     * Parse time text that might contain multiple times
     */
    parseTimeText(timeText) {
        // Handle patterns like "8,9.15, 11.30 am" or "6,7.30 pm" or "Morning cycles"
        // NEW: Also handle "9 am lab, 10.15 B57, 11.30 am Lab" with class types
        const timesWithClasses = [];
        
        // Handle descriptive times like "Morning cycles", "Evening Barre classes", "Evening cycles"
        if (/morning.*cycle/i.test(timeText)) {
            return {
                timePattern: 'morning',
                classType: 'CYCLE',
                description: timeText.trim()
            };
        }
        
        if (/evening.*cycle/i.test(timeText)) {
            return {
                timePattern: 'evening', 
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
        
        if (/evening.*class/i.test(timeText)) {
            // Generic evening classes - could be any class type
            return {
                timePattern: 'evening', 
                classType: 'ALL',
                description: timeText.trim()
            };
        }
        
        if (/morning.*class/i.test(timeText)) {
            // Generic morning classes - could be any class type
            return {
                timePattern: 'morning', 
                classType: 'ALL',
                description: timeText.trim()
            };
        }
        
        // For other descriptive times, return as-is
        if (/morning|evening|afternoon/i.test(timeText)) {
            timesWithClasses.push({ time: timeText.trim() });
            return timesWithClasses;
        }
        
        // Extract AM/PM suffix
        const ampmMatch = timeText.match(/\b(am|pm)\b/i);
        const suffix = ampmMatch ? ampmMatch[1].toLowerCase() : '';
        
        // Split by commas and parse each time with its class type
        const timeSegments = timeText.split(',');
        
        for (let segment of timeSegments) {
            segment = segment.trim();
            
            // Extract class type indicator if present
            let classType = null;
            if (/\blab\b/i.test(segment)) {
                classType = 'lab';
            } else if (/\bB57\b/i.test(segment)) {
                classType = 'barre57';
            } else if (/\bbarre\b/i.test(segment)) {
                classType = 'barre';
            } else if (/\bcycle\b/i.test(segment)) {
                classType = 'cycle';
            }
            
            // Remove AM/PM and class type indicators to get clean time
            const cleanTime = segment
                .replace(/\b(am|pm)\b/i, '')
                .replace(/\b(lab|B57|barre|cycle)\b/gi, '')
                .trim();
            
            if (cleanTime) {
                // Convert . to : for time format
                let normalizedTime = cleanTime.replace(/(\d+)\.(\d+)/, '$1:$2');
                
                // Add suffix if we have one
                if (suffix && normalizedTime.match(/^\d+:?\d*$/)) {
                    normalizedTime = `${normalizedTime} ${suffix}`;
                }
                
                timesWithClasses.push({
                    time: normalizedTime.trim(),
                    classType: classType
                });
            }
        }
        
        return timesWithClasses;
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
            // Pattern: "Tuesday - Icy Isometric" or "Tuesday : Progression Overload"
            const pattern = /([A-Za-z]+)\s*[:-]\s*(.+)/;
            const match = line.match(pattern);
            if (match && !match[2].toLowerCase().includes('sending this shortly')) {
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
        console.log(`🔍 Parsing hosted line: "${line}"`);
        
        // Pattern 1: "Kemps - Saturday - 11.30 am - B57 - SOLD OUT - for Tanzire- Pranjali"
        // Format: Location - Day - Time - Class - SOLD OUT - for <someone> - Trainer
        const soldOutPattern = /^([^-]+?)\s*-\s*([A-Za-z]+day)\s*-\s*([\d.:]+\s*(?:am|pm))\s*-\s*([^-]+?)\s*-\s*SOLD\s*OUT\s*-\s*for\s+([^-]+?)\s*-\s*(.+)$/i;
        const soldOutMatch = line.match(soldOutPattern);
        
        if (soldOutMatch) {
            const result = {
                location: soldOutMatch[1].trim(),
                day: this.expandDayName(soldOutMatch[2].trim()),
                time: soldOutMatch[3].trim(),
                classType: soldOutMatch[4].trim(),
                soldOut: true,
                hostedFor: soldOutMatch[5].trim(),
                trainer: soldOutMatch[6].trim(),
                type: 'hosted'
            };
            console.log(`✅ Parsed hosted (sold out):`, result);
            return result;
        }
        
        // Pattern 2: "Sunday - 8.45 am - Barre hosted - Kin club - Rohan" (no location prefix)
        // Format: Day - Time - Class - Venue - Trainer
        const noLocationPattern = /^([A-Za-z]+day)\s*-\s*([\d.:]+\s*(?:am|pm))\s*-\s*([^-]+?)\s*-\s*([^-]+?)\s*-\s*(.+)$/i;
        const noLocationMatch = line.match(noLocationPattern);
        
        if (noLocationMatch) {
            const result = {
                location: 'Bandra', // Default to Bandra if no location specified
                day: this.expandDayName(noLocationMatch[1].trim()),
                time: noLocationMatch[2].trim(),
                classType: noLocationMatch[3].trim(),
                venue: noLocationMatch[4].trim(),
                trainer: noLocationMatch[5].trim(),
                type: 'hosted'
            };
            console.log(`✅ Parsed hosted (no location):`, result);
            return result;
        }
        
        // Pattern 3: "Location - Day - Time - Class - Trainer" (standard format)
        const standardPattern = /^([^-]+?)\s*-\s*([A-Za-z]+day?)\s*-\s*([\d.:]+\s*(?:am|pm)?)\s*-\s*([^-]+?)\s*-\s*(.+)$/i;
        const standardMatch = line.match(standardPattern);
        
        if (standardMatch) {
            const result = {
                location: standardMatch[1].trim(),
                day: this.expandDayName(standardMatch[2].trim()),
                time: standardMatch[3].trim(),
                classType: standardMatch[4].trim(),
                trainer: standardMatch[5].trim(),
                type: 'hosted'
            };
            console.log(`✅ Parsed hosted (standard):`, result);
            return result;
        }
        
        console.log(`⚠️ Could not parse hosted line: "${line}"`);
        return null;
    }

    /**
     * Update target spreadsheet with parsed schedule data
     */
    async updateTargetSpreadsheet(scheduleData, emailInfo, isFirstEmail = true) {
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
            const copiedValues = await this.copyScheduleDataToTargetSheet(scheduleData, sheets, emailInfo);
            
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
            
            // Step 5: Apply additional covers from email (themes stay from sheet Theme columns)
            console.log('🎨 Step 5: Applying email covers to copied data (themes retained from sheet columns)...');
            const finalValues = this.applyEmailDataToSheet(currentValues, emailInfo, sheetStructure);
            
            // Step 6: Update the sheet with the final modified data
            if (finalValues.length > 0) {
                console.log('📝 Step 6: Writing final data with themes and covers...');
                await this.writeDataToTargetSheet(finalValues, sheets);
            }
            
            // Store emailInfo for use in cleaning process
            this.currentEmailInfo = emailInfo;
            this.currentScheduleSubject = emailInfo?.subject || '';
            this.currentScheduleWeek = this.resolveScheduleWeekReference(this.currentScheduleSubject);
            
            // Apply AI-detected changes if available
            if (emailInfo.changes && emailInfo.changes.length > 0 && emailInfo.aiParsed) {
                console.log(`\n🤖 Applying ${emailInfo.changes.length} AI-detected changes to schedule...`);
                // Note: Changes will be applied during the cleanAndPopulateCleanedSheet step
                // where we have access to the normalized schedule data
            }
            
            // Step 7: Populate the Covers sheet with all covers
            console.log('📋 Step 7: Populating Covers sheet...');
            // Only include email covers if this is NOT the first email
            const emailCoversToUse = isFirstEmail ? [] : emailInfo.covers;
            console.log(`📊 Using ${emailCoversToUse.length} email covers (first email: ${isFirstEmail})`);
            await this.populateCoversSheet(sheets, emailCoversToUse);
            
            // Step 8: Clean the updated data and populate the Cleaned sheet
            console.log('🧹 Step 8: Cleaning data and populating Cleaned sheet...');
            await this.cleanAndPopulateCleanedSheet(sheets);
            
            console.log(`✅ Target spreadsheet updated with fresh data, email covers applied, and sheet themes retained`);
            
        } catch (error) {
            console.error('❌ Error updating target spreadsheet:', error);
            throw error;
        }
    }

    /**
     * Copy schedule data from linked spreadsheet to target spreadsheet
     * This completely replaces the target sheet data with fresh data from linked sheet
     */
    async copyScheduleDataToTargetSheet(scheduleData, sheets, emailInfo) {
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
            let cleanedData = this.cleanSheetData(linkedSheetData);
            
            // Update date headers in row 2 to current week dates
            if (cleanedData.length >= 3) {
                console.log('📅 Updating date headers to current week...');
                cleanedData = this.updateSheetDateHeaders(cleanedData, emailInfo?.subject);
            }
            
            console.log(`✅ Prepared ${cleanedData.length} rows for target sheet`);
            
            return cleanedData;
            
        } catch (error) {
            console.error('❌ Error copying schedule data:', error);
            throw error;
        }
    }

    /**
     * Update date headers in sheet data to current week dates
     */
    updateSheetDateHeaders(sheetData, emailSubject) {
        if (!sheetData || sheetData.length < 3) {
            console.log('⚠️  Not enough rows to update date headers');
            return sheetData;
        }
        
        const updatedData = [...sheetData];
        const scheduleLayout = this.detectScheduleHeaderRows(updatedData);
        const dayRow = updatedData[scheduleLayout.dayRowIndex] || [];
        const dateRow = updatedData[scheduleLayout.dateRowIndex] || [];
        
        console.log('📅 Original date row:', dateRow.slice(0, 10));
        console.log('📅 Day row for reference:', dayRow.slice(0, 10));
        console.log(`📐 Date/day rows detected at: ${scheduleLayout.dateRowIndex + 1}/${scheduleLayout.dayRowIndex + 1}`);
        
        if (emailSubject) {
            console.log('📅 Using email subject for date calculation:', emailSubject);
        }
        
        // Update dates for each day column
        dayRow.forEach((cell, index) => {
            if (cell && typeof cell === 'string') {
                const dayMatch = cell.match(/(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)/i);
                if (dayMatch) {
                    const dayName = dayMatch[1];
                    const currentWeekDate = this.getDateForDay(dayName, emailSubject);
                    
                    // Convert format from DD-MMM-YYYY to "D MMM YYYY" (to match existing format)
                    const parts = currentWeekDate.split('-');
                    const day = parseInt(parts[0], 10); // Remove leading zero
                    const month = parts[1];
                    const year = parts[2];
                    const formattedDate = `${day} ${month} ${year}`;
                    
                    if (!updatedData[scheduleLayout.dateRowIndex]) {
                        updatedData[scheduleLayout.dateRowIndex] = [];
                    }
                    updatedData[scheduleLayout.dateRowIndex][index] = formattedDate;
                    
                    console.log(`📅 Updated ${dayName} date: ${dateRow[index]} → ${formattedDate}`);
                }
            }
        });
        
        console.log('✅ Updated date row:', (updatedData[scheduleLayout.dateRowIndex] || []).slice(0, 10));
        
        return updatedData;
    }

    /**
     * Get raw data from linked sheet maintaining exact structure
     */
    async getRawDataFromLinkedSheet(sheets) {
        console.log('📥 Fetching raw data from linked sheet...');
        
        try {
            // Reuse the exact sheet link chosen earlier in this run.
            let sheetsLink = this.currentSourceSheetsLink;
            if (!sheetsLink) {
                console.log('⚠️  No cached source sheet link found, falling back to the latest selected email body');

                const emailData = await this.findLatestScheduleEmail();
                if (!emailData) {
                    console.log('❌ No email data to extract sheet ID from');
                    return null;
                }

                sheetsLink = this.extractSheetsLink(emailData.body);
                if (!sheetsLink) {
                    console.log('❌ No sheets link found in email');
                    return null;
                }
            }

            const spreadsheetId = this.extractSpreadsheetId(sheetsLink);
            
            // Get all data from the Schedule sheet in the linked spreadsheet
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: 'Schedule!A:ZZ' // Get all data from Schedule sheet
            });
            
            const values = response.data.values || [];
            console.log(`📊 Retrieved ${values.length} rows from linked Schedule sheet`);
            
            // Store raw spreadsheet data for pattern matching
            this.rawSpreadsheetData = values;
            
            // Extract and store covers from spreadsheet
            this.spreadsheetCovers = this.extractSpreadsheetCovers(values);
            
            // Log covers found in spreadsheet
            console.log('\n📋 ===== COVERS FROM SPREADSHEET =====');
            this.logSpreadsheetCovers(values);
            console.log('=====================================\n');
            
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
     * Check if email contains a Google Sheets link (indicates first/initial email)
     */
    hasSpreadsheetLink(emailBody) {
        const sheetsLinkPattern = /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
        return sheetsLinkPattern.test(emailBody);
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
            'hosted': 'Studio Barre 57',  // Hosted classes default to Barre 57
            'hosted class': 'Studio Barre 57',
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
     * Normalize header label for structure matching
     */
    normalizeHeaderLabel(value) {
        return (value || '')
            .toString()
            .trim()
            .toLowerCase()
            .replace(/\s+/g, ' ');
    }

    /**
     * Extract canonical day name from a cell value
     */
    extractDayNameFromCell(value) {
        const text = (value || '').toString().trim();
        if (!text) return '';
        const match = text.match(/\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b/i);
        if (!match) return '';
        const day = match[1].toLowerCase();
        return day.charAt(0).toUpperCase() + day.slice(1);
    }

    /**
     * Find time column index from header row
     */
    getTimeColumnIndexFromHeader(headerRow) {
        for (let i = 0; i < (headerRow || []).length; i++) {
            if (this.normalizeHeaderLabel(headerRow[i]) === 'time') {
                return i;
            }
        }
        return -1;
    }

    /**
     * Detect where date/day/header rows are in the schedule sheet
     * Supports both:
     * - [blank, dates, days, headers, data...]
     * - [dates, days, headers, data...]
     */
    detectScheduleHeaderRows(rows) {
        const fallback = {
            dateRowIndex: 1,
            dayRowIndex: 2,
            headerRowIndex: 3,
            dataStartRowIndex: 4
        };

        if (!rows || rows.length === 0) return fallback;

        let headerRowIndex = -1;
        let bestScore = -1;
        const scanLimit = Math.min(rows.length, 8);

        for (let i = 0; i < scanLimit; i++) {
            const labels = (rows[i] || []).map(cell => this.normalizeHeaderLabel(cell));
            const hasTime = labels.includes('time');
            if (!hasTime) continue;

            let score = 0;
            if (labels.includes('location')) score += 2;
            if (labels.includes('class')) score += 2;
            if (labels.includes('trainer 1') || labels.includes('trainer1')) score += 2;
            if (labels.includes('cover')) score += 1;
            if (labels.includes('theme')) score += 1;

            if (score > bestScore) {
                bestScore = score;
                headerRowIndex = i;
            }
        }

        if (headerRowIndex === -1) {
            return fallback;
        }

        const dayRowIndex = Math.max(0, headerRowIndex - 1);
        const dateRowIndex = Math.max(0, headerRowIndex - 2);

        return {
            dateRowIndex,
            dayRowIndex,
            headerRowIndex,
            dataStartRowIndex: headerRowIndex + 1
        };
    }

    /**
     * Detect day block column mappings dynamically from day/header rows
     */
    buildDayColumnMappings(rows) {
        const mappings = {};
        const daysOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        daysOrder.forEach(day => {
            mappings[day] = [];
        });

        if (!rows || rows.length < 4) {
            return mappings;
        }

        const scheduleLayout = this.detectScheduleHeaderRows(rows);
        const dayRow = rows[scheduleLayout.dayRowIndex] || [];
        const headerRow = rows[scheduleLayout.headerRowIndex] || [];
        const maxCols = Math.max(dayRow.length, headerRow.length);

        const dayStarts = [];
        for (let col = 0; col < maxCols; col++) {
            const dayName = this.extractDayNameFromCell(dayRow[col]);
            if (!dayName) continue;

            const prevDay = col > 0 ? this.extractDayNameFromCell(dayRow[col - 1]) : '';
            if (col === 0 || prevDay !== dayName) {
                dayStarts.push({ day: dayName, startCol: col });
            }
        }

        // Fallback when day row is missing labels: infer from repeating "Location" blocks
        if (dayStarts.length === 0) {
            let dayIdx = 0;
            for (let col = 0; col < maxCols && dayIdx < daysOrder.length; col++) {
                if (this.normalizeHeaderLabel(headerRow[col]) === 'location') {
                    dayStarts.push({ day: daysOrder[dayIdx], startCol: col });
                    dayIdx++;
                }
            }
        }

        for (let idx = 0; idx < dayStarts.length; idx++) {
            const current = dayStarts[idx];
            const next = dayStarts[idx + 1];
            const endCol = next ? next.startCol - 1 : maxCols - 1;

            const config = {
                dayCol: current.startCol,
                startCol: current.startCol,
                endCol,
                locationCol: -1,
                classCol: -1,
                trainer1Col: -1,
                trainer2Col: -1,
                coverCol: -1,
                themeCol: -1
            };

            for (let col = current.startCol; col <= endCol; col++) {
                const header = this.normalizeHeaderLabel(headerRow[col]);
                if (header === 'location' && config.locationCol === -1) config.locationCol = col;
                if (header === 'class' && config.classCol === -1) config.classCol = col;
                if ((header === 'trainer 1' || header === 'trainer1') && config.trainer1Col === -1) config.trainer1Col = col;
                if ((header === 'trainer 2' || header === 'trainer2') && config.trainer2Col === -1) config.trainer2Col = col;
                if (header === 'cover' && config.coverCol === -1) config.coverCol = col;
                if (header === 'theme' && config.themeCol === -1) config.themeCol = col;
            }

            // Relative fallback based on common block order
            if (config.locationCol === -1 && current.startCol <= endCol) config.locationCol = current.startCol;
            if (config.classCol === -1 && current.startCol + 1 <= endCol) config.classCol = current.startCol + 1;
            if (config.trainer1Col === -1 && current.startCol + 2 <= endCol) config.trainer1Col = current.startCol + 2;
            if (config.trainer2Col === -1 && current.startCol + 3 <= endCol) config.trainer2Col = current.startCol + 3;
            if (config.coverCol === -1 && current.startCol + 4 <= endCol) config.coverCol = current.startCol + 4;
            if (config.themeCol === -1 && current.startCol + 5 <= endCol &&
                this.normalizeHeaderLabel(headerRow[current.startCol + 5]) === 'theme') {
                config.themeCol = current.startCol + 5;
            }

            if (config.locationCol >= 0 && config.classCol >= 0 && config.trainer1Col >= 0 && mappings[current.day]) {
                mappings[current.day].push(config);
            }
        }

        return mappings;
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

            const scheduleLayout = this.detectScheduleHeaderRows(rows);
            const headerRow = rows[scheduleLayout.headerRowIndex] || [];
            const dateRow = rows[scheduleLayout.dateRowIndex] || [];
            const dataRows = rows.slice(scheduleLayout.dataStartRowIndex);
            
            console.log('📋 Processing schedule data for cleaning...');
            console.log(`🔢 Found ${dataRows.length} data rows to process`);
            console.log(`📐 Detected layout: date row=${scheduleLayout.dateRowIndex + 1}, day row=${scheduleLayout.dayRowIndex + 1}, header row=${scheduleLayout.headerRowIndex + 1}, data starts row=${scheduleLayout.dataStartRowIndex + 1}`);
            
            // Dynamically detect all day blocks and column positions from sheet headers
            const dayColumnMappings = this.buildDayColumnMappings(rows);
            const daysOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

            // Find time column
            const timeColIndex = this.getTimeColumnIndexFromHeader(headerRow);
            
            if (timeColIndex === -1) {
                console.log('❌ Time column header not found in row 4');
                return;
            }

            console.log('📋 Day column mapping for cleaning:');
            for (const dayName of daysOrder) {
                const configs = dayColumnMappings[dayName] || [];
                if (configs.length === 0) continue;
                for (const cfg of configs) {
                    console.log(
                        `   ${dayName}: day=${cfg.dayCol}, location=${cfg.locationCol}, class=${cfg.classCol}, trainer1=${cfg.trainer1Col}, trainer2=${cfg.trainer2Col}, cover=${cfg.coverCol}, theme=${cfg.themeCol}`
                    );
                }
            }

            const allClasses = [];

            // Process each data row
            for (let r = 0; r < dataRows.length; r++) {
                const row = dataRows[r];
                
                // Process each configured day block
                for (const day of daysOrder) {
                    const dayConfigs = dayColumnMappings[day] || [];
                    for (const colConfig of dayConfigs) {
                        const location = this.normalizeLocationName(row[colConfig.locationCol]);
                        if (!location) continue;

                        const classNameRaw = row[colConfig.classCol];
                        const className = this.normalizeClassNameForCleaned(classNameRaw);
                        if (!className || !this.isValidClassName(className)) continue;

                        const trainerRaw = row[colConfig.trainer1Col];
                        const trainer2Raw = colConfig.trainer2Col >= 0 ? row[colConfig.trainer2Col] : '';
                        const coverRaw = colConfig.coverCol >= 0 ? row[colConfig.coverCol] : '';
                        const themeRaw = colConfig.themeCol >= 0 ? row[colConfig.themeCol] : '';
                        
                        // Parse time first so it's available for logging
                        const timeRaw = row[timeColIndex];
                        const timeDate = this.parseTimeToDate(timeRaw);
                        let time = timeDate ? this.formatTime(timeDate) : timeRaw;
                        
                        let trainer = this.normalizeTrainerName(trainerRaw);
                        let notes = '';

                        if (!trainer) {
                            const secondaryTrainer = this.normalizeTrainerName(trainer2Raw);
                            if (secondaryTrainer && this.isTrainerName(secondaryTrainer)) {
                                trainer = secondaryTrainer;
                                console.log(`  ↪ Using Trainer 2 at ${day} ${time}: ${secondaryTrainer}`);
                            }
                        }
                        
                        // **STEP 1: Check if this is a hosted class (class name = "hosted") - CHECK THIS FIRST**
                        const isHostedClass = (trainerRaw && trainerRaw.toString().toLowerCase().includes('hosted')) || 
                                             (classNameRaw && classNameRaw.toString().toLowerCase().includes('hosted'));
                        
                        // If hosted class, mark as SOLD OUT and use cover trainer if available
                        if (isHostedClass) {
                            notes = 'SOLD OUT';
                            // For hosted classes, if there's a cover, that's the actual trainer doing the class
                            if (coverRaw && coverRaw.toString().trim() && coverRaw.toString().toLowerCase() !== 'undefined') {
                                trainer = this.normalizeTrainerName(coverRaw);
                                console.log(`  Hosted class at ${day} ${time} - marked as SOLD OUT - Trainer: ${trainer}`);
                            } else {
                                console.log(`  Hosted class at ${day} ${time} - marked as SOLD OUT`);
                            }
                        } else {
                            // **STEP 2: For non-hosted classes, check if Cover column has a value**
                            // If Cover column has a value, replace Trainer 1 with the cover trainer
                            if (coverRaw && coverRaw.toString().trim() && coverRaw.toString().toLowerCase() !== 'undefined') {
                                const coverNorm = this.normalizeTrainerName(coverRaw);
                                // Skip if cover is just "Yes" or similar indicators (not a real trainer name)
                                const invalidCoverValues = ['yes', 'no', 'true', 'false', 'y', 'n', 'x', '1', '0'];
                                if (coverNorm && !invalidCoverValues.includes(coverNorm.toLowerCase())) {
                                    const originalTrainer = trainer || 'regular instructor';
                                    notes = `Cover: ${coverNorm} for ${originalTrainer}`;
                                    trainer = coverNorm; // Replace trainer with cover
                                    console.log(`  ✓ Applied cover at ${day} ${time}: ${coverNorm} covering for ${originalTrainer}`);
                                }
                            }
                        }

                        // Exclude classes without a trainer (unless hosted and will be filled from email)
                        if (!trainer && !isHostedClass) continue;
                        // Normalize time for consistent alignment
                        time = this.normalizeTimeDisplay(time);
                        
                        // Get actual date from row 2 (same column as day header if possible)
                        const rawDate = dateRow[colConfig.dayCol] || dateRow[colConfig.locationCol];
                        const date = rawDate && rawDate.toString().trim() ? 
                            this.formatDateFromSheet(rawDate.toString().trim()) : 
                            this.getDateForDay(day); // fallback to calculated date
                        
                        // Use theme directly from Theme column when present
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
            }

            // Build a map of day -> date from the classes we've already processed (from sheet)
            // This ensures hosted classes get dates consistent with the schedule week
            const dateByDay = {};
            for (const cls of allClasses) {
                if (cls.Date && !dateByDay[cls.Day]) {
                    dateByDay[cls.Day] = cls.Date;
                }
            }
            console.log('📅 Date map from sheet:', dateByDay);

            // Build a set of existing class keys to prevent duplicates
            // Also build a set of SOLD OUT classes by day+time+trainer for hosted class deduplication
            const existingClassKeys = new Set();
            const existingSoldOutKeys = new Set();
            
            for (const cls of allClasses) {
                // Normalize time for comparison (remove special chars, standardize format)
                const normalizedTime = this.normalizeTimeForComparison(cls.Time);
                const key = `${cls.Day}|${normalizedTime}|${this.normalizeClassName(cls.Class)}`.toLowerCase();
                existingClassKeys.add(key);
                
                // For SOLD OUT classes, also track by day+time+trainer (to catch hosted variants)
                if (cls.Notes && cls.Notes.includes('SOLD OUT')) {
                    const soldOutKey = `${cls.Day}|${normalizedTime}|${this.normalizeTrainerName(cls.Trainer)}`.toLowerCase();
                    existingSoldOutKeys.add(soldOutKey);
                    console.log(`  📝 Tracking SOLD OUT: ${soldOutKey}`);
                }
            }

            // Add hosted classes from email info if available
            if (this.currentEmailInfo && this.currentEmailInfo.hostedClasses) {
                console.log(`📋 Adding ${this.currentEmailInfo.hostedClasses.length} hosted classes from email...`);
                for (const hosted of this.currentEmailInfo.hostedClasses) {
                    // Normalize location to match our format
                    const normalizedLocation = this.normalizeLocationName(hosted.location);
                    
                    // Normalize time for the hosted class
                    const normalizedTime = this.normalizeTimeDisplay(hosted.time);
                    const normalizedTimeForCompare = this.normalizeTimeForComparison(hosted.time);
                    const normalizedTrainer = this.normalizeTrainerName(hosted.trainer);
                    
                    // Check if this class already exists by class name
                    const hostedKey = `${hosted.day}|${normalizedTimeForCompare}|${this.normalizeClassName(hosted.classType)}`.toLowerCase();
                    
                    // Also check by day+time+trainer for SOLD OUT duplicates
                    const soldOutKey = `${hosted.day}|${normalizedTimeForCompare}|${normalizedTrainer}`.toLowerCase();
                    
                    if (existingClassKeys.has(hostedKey)) {
                        console.log(`  ⚠️ Skipping duplicate hosted class (class match): ${hosted.day} ${hosted.time} ${hosted.classType}`);
                        continue;
                    }
                    
                    if (existingSoldOutKeys.has(soldOutKey)) {
                        console.log(`  ⚠️ Skipping duplicate hosted class (SOLD OUT match): ${hosted.day} ${hosted.time} ${hosted.trainer}`);
                        continue;
                    }
                    
                    // Use the date from sheet data for the same day, fall back to calculated
                    const hostedDate = dateByDay[hosted.day] || this.getDateForDay(hosted.day);
                    console.log(`  📅 Hosted class on ${hosted.day} using date: ${hostedDate}`);
                    
                    existingClassKeys.add(hostedKey);
                    existingSoldOutKeys.add(soldOutKey);
                    
                    allClasses.push({
                        Day: hosted.day,
                        Time: normalizedTime,
                        Location: normalizedLocation,
                        Class: this.normalizeClassNameForCleaned(hosted.classType),
                        Trainer: normalizedTrainer,
                        Notes: 'SOLD OUT',
                        Date: hostedDate,
                        Theme: ''
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
            
            // Apply AI-detected changes if available
            let finalClasses = allClasses;
            if (this.currentEmailInfo && this.currentEmailInfo.changes && this.currentEmailInfo.changes.length > 0) {
                console.log(`\n🤖 Applying AI-detected changes to cleaned schedule data...`);
                finalClasses = this.applyAIChangesToSchedule(allClasses, this.currentEmailInfo.changes);
            }

            // Write cleaned data to Cleaned sheet
            const headers = ['Day', 'Time', 'Location', 'Class', 'Trainer', 'Notes', 'Date', 'Theme'];
            const values = [
                headers, 
                ...finalClasses.map(obj => headers.map(h => obj[h] || ''))
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
    getDateForDay(dayName, emailSubject) {
        var daysOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        var targetDayIndex = daysOrder.indexOf(dayName);
        
        if (targetDayIndex === -1) return new Date().toLocaleDateString('en-GB');

        var weekReference = this.resolveScheduleWeekReference(emailSubject);
        var monday = new Date(weekReference.monday);
        console.log('📅 Using Monday for schedule week: ' + monday.toDateString());
        
        // Calculate target day from Monday
        var targetDate = new Date(monday);
        targetDate.setDate(monday.getDate() + targetDayIndex); // Monday=0, Tuesday=1, ..., Sunday=6
        
        // Format as DD-MMM-YYYY
        var day = targetDate.getDate().toString();
        if (day.length === 1) day = '0' + day;
        var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
            'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        var month = monthNames[targetDate.getMonth()];
        var year = targetDate.getFullYear();
        
        return day + '-' + month + '-' + year;
    }

    /**
     * Resolve the Monday/Sunday boundaries for the active schedule week.
     * Priority:
     * 1. Explicit email subject passed to the method
     * 2. Current run's stored email subject
     * 3. Previously cached schedule week
     * 4. Current calendar week
     */
    resolveScheduleWeekReference(emailSubject) {
        var subjectToUse = emailSubject || this.currentScheduleSubject || this.currentEmailInfo?.subject || '';

        if (subjectToUse && this.currentScheduleWeek?.subject === subjectToUse && this.currentScheduleWeek?.monday) {
            return {
                ...this.currentScheduleWeek,
                monday: new Date(this.currentScheduleWeek.monday),
                sunday: new Date(this.currentScheduleWeek.sunday)
            };
        }

        if (subjectToUse) {
            var parsedWeek = this.extractWeekFromEmailSubject(subjectToUse);
            if (parsedWeek && parsedWeek.weekFound) {
                var mondayFromSubject = new Date(parsedWeek.startDate);
                mondayFromSubject.setHours(0, 0, 0, 0);

                var sundayFromSubject = new Date(mondayFromSubject);
                sundayFromSubject.setDate(mondayFromSubject.getDate() + 6);

                var resolvedFromSubject = {
                    source: 'email-subject',
                    subject: subjectToUse,
                    monday: mondayFromSubject,
                    sunday: sundayFromSubject
                };

                this.currentScheduleSubject = subjectToUse;
                this.currentScheduleWeek = resolvedFromSubject;
                return resolvedFromSubject;
            }
        }

        if (this.currentScheduleWeek?.monday) {
            return {
                ...this.currentScheduleWeek,
                monday: new Date(this.currentScheduleWeek.monday),
                sunday: new Date(this.currentScheduleWeek.sunday)
            };
        }

        var today = new Date();
        var currentDay = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
        var monday = new Date(today);
        var daysFromMonday = currentDay === 0 ? -6 : (1 - currentDay);
        monday.setDate(today.getDate() + daysFromMonday);
        monday.setHours(0, 0, 0, 0);

        var sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6);

        var fallbackWeek = {
            source: 'current-week',
            subject: '',
            monday,
            sunday
        };

        this.currentScheduleWeek = fallbackWeek;
        return fallbackWeek;
    }

    /**
     * Extract week date range from email subject
     * Handles formats like "9 -15th Feb'26" or "29 Nov - 5th Dec '25"
     */
    extractWeekFromEmailSubject(emailSubject) {
        if (!emailSubject) return null;

        emailSubject = String(emailSubject)
            .replace(/[–—]/g, '-')
            .replace(/^\s*(?:re|fwd?)\s*:\s*/i, '')
            .trim();

        const getMonthIndex = (monthName) => {
            if (!monthName) return -1;
            const normalized = monthName.toLowerCase().trim();
            const monthMap = {
                jan: 0,
                january: 0,
                feb: 1,
                february: 1,
                mar: 2,
                march: 2,
                apr: 3,
                april: 3,
                may: 4,
                jun: 5,
                june: 5,
                jul: 6,
                july: 6,
                aug: 7,
                august: 7,
                sep: 8,
                sept: 8,
                september: 8,
                oct: 9,
                october: 9,
                nov: 10,
                november: 10,
                dec: 11,
                december: 11
            };

            return monthMap[normalized] ?? -1;
        };
        
        // Pattern for "9 -15th Feb'26", "23 - 29th Mar'26", or "9th - 15th February'26"
        var sameMonthPattern = /(\d{1,2})(?:st|nd|rd|th)?\s*-\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\s*'(\d{2})/i;
        var sameMonthMatch = emailSubject.match(sameMonthPattern);
        
        if (sameMonthMatch) {
            var startDay = parseInt(sameMonthMatch[1]);
            var endDay = parseInt(sameMonthMatch[2]);
            var month = sameMonthMatch[3];
            var year = 2000 + parseInt(sameMonthMatch[4]);
            
            console.log('📅 Extracted week from subject: ' + startDay + '-' + endDay + ' ' + month + ' ' + year);
            
            var monthIndex = getMonthIndex(month);
            
            if (monthIndex === -1) return null;
            
            // Create start date (should be Monday)
            var startDate = new Date(year, monthIndex, startDay);
            
            return {
                startDate: startDate,
                endDate: new Date(year, monthIndex, endDay),
                weekFound: true
            };
        }

        // Pattern for month-first format like
        // "Mumbai schedule for April 6-12 '26" or "April 6 - 12th '26"
        var monthFirstPattern = /([A-Za-z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?\s*-\s*(\d{1,2})(?:st|nd|rd|th)?\s*'?(\d{2})/i;
        var monthFirstMatch = emailSubject.match(monthFirstPattern);

        if (monthFirstMatch) {
            var monthFirstMonth = monthFirstMatch[1];
            var monthFirstStartDay = parseInt(monthFirstMatch[2]);
            var monthFirstEndDay = parseInt(monthFirstMatch[3]);
            var monthFirstYear = 2000 + parseInt(monthFirstMatch[4]);

            console.log('📅 Extracted month-first week from subject: ' + monthFirstMonth + ' ' + monthFirstStartDay + '-' + monthFirstEndDay + ' ' + monthFirstYear);

            var monthFirstMonthIndex = getMonthIndex(monthFirstMonth);
            if (monthFirstMonthIndex === -1) return null;

            return {
                startDate: new Date(monthFirstYear, monthFirstMonthIndex, monthFirstStartDay),
                endDate: new Date(monthFirstYear, monthFirstMonthIndex, monthFirstEndDay),
                weekFound: true
            };
        }
        
        // Pattern for "29 Nov - 5th Dec '25" or "30th Mar - 5th April'26" format (cross-month)
        var crossMonthPattern = /(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\s*-\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\s*'(\d{2})/i;
        var crossMonthMatch = emailSubject.match(crossMonthPattern);
        
        if (crossMonthMatch) {
            var startDay = parseInt(crossMonthMatch[1]);
            var startMonth = crossMonthMatch[2];
            var endDay = parseInt(crossMonthMatch[3]);
            var endMonth = crossMonthMatch[4];
            var year = 2000 + parseInt(crossMonthMatch[5]);
            
            console.log('📅 Extracted cross-month week: ' + startDay + ' ' + startMonth + ' - ' + endDay + ' ' + endMonth + ' ' + year);
            
            var startMonthIndex = getMonthIndex(startMonth);
            var endMonthIndex = getMonthIndex(endMonth);
            
            if (startMonthIndex === -1 || endMonthIndex === -1) return null;
            
            var startDate = new Date(year, startMonthIndex, startDay);
            var endDate = new Date(year, endMonthIndex, endDay);
            
            return {
                startDate: startDate,
                endDate: endDate,
                weekFound: true
            };
        }
        
        console.log('📅 No week date pattern found in email subject');
        return null;
    }

    /**
     * Normalize location names
     */
    /**
     * Enhanced location normalization using dynamic mapping
     */
    normalizeLocationName(raw) {
        if (!raw) return '';
        
        const locationResult = this.enhancedMapper.identifyLocation(raw);
        return locationResult.canonical;
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
     * Check whether notes or theme mark a class as sold out.
     */
    isSoldOutClass(classData = {}) {
        const notes = String(classData.Notes || classData.notes || '').trim();
        const theme = String(classData.Theme || classData.theme || '').trim();
        return /\bsold\s*out\b/i.test(notes) || /\bsold\s*out\b/i.test(theme);
    }

    /**
     * Hosted classes and rows without any resolved trainer should never be rendered into PDFs.
     */
    shouldExcludeClassFromPdf(classData = {}) {
        const className = String(classData.Class || classData.class || '').trim();
        const trainer = String(classData.Trainer || classData.trainer || '').trim();
        const notes = String(classData.Notes || classData.notes || '').trim();

        if (/\bhosted\b/i.test(className) || /\bhosted\b/i.test(notes)) {
            return true;
        }

        if (!trainer) {
            return true;
        }

        if (this.isSoldOutClass(classData)) {
            return true;
        }

        return false;
    }

    /**
     * Check if class name is valid
     */
    isValidClassName(name) {
        if (!name) return false;
        const val = name.toString().trim().toLowerCase();
        // Invalid class names - classes from these should be skipped as they come from trainer/notes fields
        // Fixed: Use word boundary matching instead of substring matching to avoid false positives
        const invalid = ['smita', 'anandita', 'cover', 'replacement', 'sakshi', 'parekh', 'taarika', 'host'];
        const words = val.split(/\s+/);
        if (invalid.some(invalidWord => words.includes(invalidWord))) return false;
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
     * Normalize time string format - remove special characters, standardize format
     */
    normalizeTimeString(timeStr) {
        if (!timeStr) return '';
        let t = timeStr.toString().trim();
        
        // Remove invisible/special Unicode characters
        t = t.replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '');
        
        // Replace all separators (. , : etc) with :
        t = t.replace(/[.,;]/g, ':');
        
        // Remove extra colons
        t = t.replace(/:+/g, ':');
        
        // Ensure space between time and AM/PM
        t = t.replace(/(\d)(AM|PM)/gi, '$1 $2');
        
        // Clean up multiple spaces
        t = t.replace(/\s+/g, ' ').trim();
        
        return t.toUpperCase();
    }

    /**
     * Normalize time string for consistent display alignment
     */
    normalizeTimeDisplay(timeStr) {
        if (!timeStr) return '';
        const normalized = this.normalizeTimeString(timeStr);
        const match = normalized.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
        if (!match) return timeStr;
        
        let hour = parseInt(match[1]);
        const minute = match[2];
        const ampm = match[3].toUpperCase();
        
        // Format with consistent spacing: "HH:MM AM" (2-digit hour, no leading space)
        const paddedHour = hour.toString().padStart(2, '0');
        return `${paddedHour}:${minute} ${ampm}`;
    }

    /**
     * Normalize time for comparison (strip all formatting, just get HH:MM AM/PM)
     * Used to detect duplicate classes with different time formats (8.45 am vs 8:45 AM)
     */
    normalizeTimeForComparison(timeStr) {
        if (!timeStr) return '';
        
        // Remove all special characters except digits and letters
        let t = timeStr.toString().trim();
        
        // Replace . , : with : for consistent parsing
        t = t.replace(/[.,]/g, ':');
        
        // Remove extra spaces
        t = t.replace(/\s+/g, ' ').trim();
        
        // Extract time components
        const match = t.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
        if (!match) return t.toLowerCase();
        
        let hour = parseInt(match[1]);
        const minute = match[2] ? parseInt(match[2]) : 0;
        let ampm = (match[3] || '').toLowerCase();
        
        // If no AM/PM specified, try to infer from hour
        if (!ampm) {
            ampm = hour >= 7 && hour < 12 ? 'am' : 'pm';
        }
        
        // Return standardized format for comparison: "h:mm am" (no padding)
        return `${hour}:${minute.toString().padStart(2, '0')} ${ampm}`;
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
        const scheduleLayout = this.detectScheduleHeaderRows(values);
        const structure = {
            dayColumns: {},
            coverColumns: {},
            trainer2Columns: {},
            themeColumns: {},
            headerRows: scheduleLayout.dataStartRowIndex,
            timeColumn: 0
        };

        const headerRow = values[scheduleLayout.headerRowIndex] || [];
        const detectedTimeCol = this.getTimeColumnIndexFromHeader(headerRow);
        if (detectedTimeCol >= 0) {
            structure.timeColumn = detectedTimeCol;
        }

        const dayMappings = this.buildDayColumnMappings(values);
        for (const [day, configs] of Object.entries(dayMappings)) {
            if (!configs || configs.length === 0) continue;
            structure.dayColumns[day] = configs;

            const primary = configs[0];
            structure.trainer2Columns[day] = primary.trainer2Col;
            structure.coverColumns[day] = primary.coverCol;
            structure.themeColumns[day] = primary.themeCol;

            for (const config of configs) {
                console.log(
                    `📋 Found ${day}: day=${config.dayCol}, location=${config.locationCol}, class=${config.classCol}, trainer1=${config.trainer1Col}, trainer2=${config.trainer2Col}, cover=${config.coverCol}, theme=${config.themeCol}`
                );
            }
        }

        return structure;
    }

    /**
     * Apply email cover data to the existing sheet structure.
     * Theme values are sourced directly from sheet Theme columns.
     */
    applyEmailDataToSheet(currentValues, emailInfo, structure) {
        console.log(`🔄 Applying ${emailInfo.covers.length} covers to sheet...`);
        
        // Create a copy of current values to modify
        const updatedValues = currentValues.map(row => [...row]);
        
        let coversApplied = 0;
        let themesApplied = 0;
        
        // NOTE: DO NOT CLEAR COVER COLUMNS
        // The spreadsheet data from the linked sheet already contains covers.
        // We only add ADDITIONAL covers from the email body, not replace existing ones.
        console.log('ℹ️  Preserving existing covers from spreadsheet, will only add additional covers from email...');
        
        // Apply additional covers from email
        console.log('🔍 Starting additional cover application from email...');
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
                    if (typeof colConfig.coverCol !== 'number' || colConfig.coverCol < 0 || colConfig.coverCol >= row.length) continue;
                    
                    const timeCell = String(row[structure.timeColumn] || '').trim();
                    const locationCell = String(row[colConfig.locationCol] || '').trim().toLowerCase();
                    const classCell = String(row[colConfig.classCol] || '').toLowerCase();
                    
                    // Skip hosted classes - they should not receive covers from email
                    if (classCell.includes('hosted')) {
                        continue;
                    }
                    
                    // Match location - improved matching
                    if (!this.matchLocation(locationCell, cover.location)) continue;
                    
                    let shouldApplyCover = false;
                    
                    if (cover.timePattern && cover.classType) {
                        // Handle pattern-based covers (morning cycles, evening barre, evening cycles, etc.)
                        console.log(`🔍 Checking pattern cover: ${cover.timePattern} ${cover.classType} against ${timeCell} ${classCell}`);
                        
                        const isPM = timeCell.toLowerCase().includes('pm');
                        const isAM = timeCell.toLowerCase().includes('am');
                        
                        // Determine if the class type matches
                        let classTypeMatches = false;
                        if (cover.classType === 'ALL') {
                            classTypeMatches = true; // Match any class type
                        } else if (cover.classType.toLowerCase() === 'cycle') {
                            classTypeMatches = classCell.includes('cycle') || classCell.includes('powercycle');
                        } else if (cover.classType.toLowerCase() === 'barre') {
                            classTypeMatches = classCell.includes('barre') || classCell.includes('b57');
                        } else {
                            classTypeMatches = classCell.includes(cover.classType.toLowerCase());
                        }
                        
                        if (cover.timePattern === 'morning' && isAM && classTypeMatches) {
                            shouldApplyCover = true;
                            console.log(`✅ Morning ${cover.classType} match found`);
                        } else if (cover.timePattern === 'evening' && isPM && classTypeMatches) {
                            shouldApplyCover = true;
                            console.log(`✅ Evening ${cover.classType} match found`);
                        }
                    } else if (cover.timesWithClasses && cover.timesWithClasses.length > 0) {
                        // NEW: Handle specific time-based covers with class types
                        // Example: [{time: "9 am", classType: "lab"}, {time: "10:15 am", classType: "barre57"}, {time: "11:30 am", classType: "lab"}]
                        for (const timeWithClass of cover.timesWithClasses) {
                            const coverTime = timeWithClass.time || timeWithClass;
                            const coverClassType = timeWithClass.classType;
                            
                            // Normalize time formats for comparison
                            const normalizedCoverTime = this.normalizeTime(coverTime);
                            const normalizedCellTime = this.normalizeTime(timeCell);
                            
                            console.log(`🔍 Time+Class comparison: "${normalizedCellTime}" vs "${normalizedCoverTime}" | Cell class="${classCell}" vs Cover class="${coverClassType || 'any'}"`);
                            
                            // Check time match first
                            const timeMatches = this.timeMatches(normalizedCellTime, normalizedCoverTime);
                            
                            if (timeMatches) {
                                // If class type is specified in cover, must match the actual class
                                if (coverClassType) {
                                    let classMatches = false;
                                    
                                    if (coverClassType === 'lab') {
                                        // Match strength/lab classes
                                        classMatches = classCell.includes('strength') || classCell.includes('lab');
                                    } else if (coverClassType === 'barre57' || coverClassType === 'barre') {
                                        // Match barre classes
                                        classMatches = classCell.includes('barre') || classCell.includes('b57');
                                    } else if (coverClassType === 'cycle') {
                                        // Match cycle classes
                                        classMatches = classCell.includes('cycle') || classCell.includes('powercycle');
                                    } else {
                                        // Generic match
                                        classMatches = classCell.includes(coverClassType.toLowerCase());
                                    }
                                    
                                    if (classMatches) {
                                        shouldApplyCover = true;
                                        console.log(`✅ Time+Class match found! ${normalizedCellTime} ${coverClassType}`);
                                        break;
                                    } else {
                                        console.log(`❌ Time matches but class type doesn't: expected "${coverClassType}", got "${classCell}"`);
                                    }
                                } else {
                                    // No class type specified, just time match is enough
                                    shouldApplyCover = true;
                                    console.log(`✅ Time match found (no class filter)!`);
                                    break;
                                }
                            }
                        }
                    } else if (cover.times && cover.times.length > 0) {
                        // Legacy: Handle old format without class types
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
                        // Only apply cover if the cell is currently empty
                        // This preserves existing covers from the spreadsheet
                        const existingCover = String(row[colConfig.coverCol] || '').trim();
                        
                        if (!existingCover || existingCover.toLowerCase() === 'undefined') {
                            row[colConfig.coverCol] = cover.trainer;
                            coversApplied++;
                            console.log(`✅ Applied additional cover from email: ${cover.trainer} to ${cover.day} ${timeCell} ${classCell} at row ${rowIndex + 1}, col ${colConfig.coverCol + 1}`);
                        } else {
                            console.log(`ℹ️  Skipping - existing cover already present: "${existingCover}" at ${cover.day} ${timeCell}`);
                        }
                    }
                }
            }
        }

        if (emailInfo.themes && emailInfo.themes.length > 0) {
            console.log('ℹ️  Skipping email theme application. Theme values are sourced directly from the sheet Theme columns.');
        }

        console.log(`📊 Applied ${coversApplied} covers and ${themesApplied} themes to spreadsheet`);
        return updatedValues;
    }

    /**
     * Update theme in the in-memory class data to keep it in sync with Google Sheets
     */
    updateInMemoryClassTheme(dayName, timeString, className, theme) {
        if (!this.kwalityClasses) return;
        
        // Normalize values for matching
        const normalizedDay = dayName.charAt(0).toUpperCase() + dayName.slice(1).toLowerCase();
        const normalizedTime = this.normalizeTimeForComparison(timeString);
        const normalizedClass = this.normalizeClassName(className);
        
        // Find matching class record(s) and update theme
        let updatedCount = 0;
        for (const classData of this.kwalityClasses) {
            if (classData.Day === normalizedDay &&
                this.normalizeTimeForComparison(classData.Time) === normalizedTime &&
                this.normalizeClassName(classData.Class) === normalizedClass) {
                
                classData.Theme = theme;
                updatedCount++;
                console.log(`  🔄 Updated in-memory class: ${normalizedDay} ${timeString} ${className} -> Theme: "${theme}"`);
            }
        }
        
        if (updatedCount === 0) {
            console.log(`  ⚠️  No matching in-memory class found for: ${normalizedDay} ${timeString} ${className}`);
        }
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
    normalizeClassNameForDisplay(className) {
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
        
        const norm1 = this.normalizeClassNameForDisplay(className1);
        const norm2 = this.normalizeClassNameForDisplay(className2);
        
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
        
        // Location aliases - IMPORTANT: Annex is treated as Kemps, not Bandra!
        const aliases = {
            'kemps': ['kwality', 'kemps corner', 'kemps', 'annex'],
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

        this.resetStaticThemeState();
        this.cleanupStaticThemeArtifacts();
        this.ensureThemeBadgeCSS();
        this.ensureStaticThemeArtifactCSS();
        
        // Ensure sold-out badge CSS is present
        this.ensureSoldOutBadgeCSS();
        
    }

    /**
     * Reset in-memory static theme rendering state.
     */
    resetStaticThemeState() {
        this.staticThemeRows = [];
        this.staticThemeColorMap = new Map();
        this.staticThemeLegendAssetMap = new Map();
        this.staticThemeSharedWidth = null;
        this.staticThemeIndexBandWidth = 80.2;
    }

    /**
     * Remove previously generated static theme overlays and index entries.
     */
    cleanupStaticThemeArtifacts() {
        if (!this.$) return;

        this.$('.theme-row-highlight, .theme-index-band, .theme-index-entry, .theme-index-title, .theme-static-label').remove();
        if (this.shouldRenderStaticTheme()) {
            this.$('.theme-badge').remove();
        }
        this.$('span').filter((_index, element) => this.$(element).text().trim().toUpperCase() === 'STATIC MAGIC').remove();
    }

    /**
     * Ensure the base theme badge CSS matches the current pastel badge treatment.
     */
    ensureThemeBadgeCSS() {
        const styleTag = this.$('style').first();
        if (!styleTag.length) {
            console.warn('⚠️  No style tag found, skipping theme badge CSS check');
            return;
        }

        const isBandra = String(this.currentLocation || '').toLowerCase().includes('bandra');
        const badgePadding = isBandra ? '3px 10px' : '3px 11px';
        const badgeBorderRadius = isBandra ? '8px' : '6px';
        const themeBadgeCSS = `
        .theme-badge {
            background: linear-gradient(135deg, rgba(245, 188, 208, 0.82) 0%, rgba(245, 188, 208, 0.62) 100%);
            color: #2C2D2D;
            padding: ${badgePadding};
            border-radius: ${badgeBorderRadius};
            font-size: ${isBandra ? '8.5px' : '8px'};
            font-weight: 700;
            font-family: 'Montserrat', sans-serif;
            margin-left: ${isBandra ? '9px' : '6px'};
            display: inline-block;
            vertical-align: middle;
            line-height: 1.3;
            box-shadow: 0 6px 14px rgba(69, 59, 42, 0.12);
            border: none;
            letter-spacing: ${isBandra ? '0.28px' : '0.1px'};
            text-transform: uppercase;
            min-width: fit-content;
            max-width: ${isBandra ? '148px' : '180px'};
            text-align: center;
            white-space: normal;
            word-wrap: break-word;
            position: relative;
            top: -1px;
            backdrop-filter: blur(12px) saturate(140%);
            -webkit-backdrop-filter: blur(12px) saturate(140%);
        }`;

        let existingStyle = styleTag.html() || '';

        if (existingStyle.includes('.theme-badge')) {
            console.log('🎨 Updating theme badge CSS...');
            existingStyle = existingStyle.replace(/\.theme-badge\s*\{[^}]+\}/g, '');
            styleTag.html(existingStyle + themeBadgeCSS);
        } else {
            console.log('🎨 Injecting theme badge CSS...');
            styleTag.append(themeBadgeCSS);
        }
        console.log('✅ Theme badge CSS applied');
    }

    /**
     * Ensure static theme highlight/index CSS matches the current borderless surface treatment.
     */
    ensureStaticThemeArtifactCSS() {
        const styleTag = this.$('style').first();
        if (!styleTag.length) {
            console.warn('⚠️  No style tag found, skipping static theme CSS check');
            return;
        }

        const staticThemeCSS = `
        .theme-row-highlight,
        .theme-index-band {
            border-radius: 8px !important;
            border: none !important;
            box-shadow: 0 10px 22px rgba(69, 59, 42, 0.10) !important;
            opacity: 1;
        }

        .theme-row-highlight {
            transform: none !important;
            transform-origin: left center;
        }

        .theme-index-band {
            transform: none !important;
            transform-origin: left center;
        }

        .theme-index-entry {
            color: #453b2a !important;
            font-family: 'Montserrat-Bold_1z', 'Montserrat-Bold_21', 'Montserrat', sans-serif !important;
            font-size: 12px !important;
            font-style: normal !important;
            font-weight: 700 !important;
            letter-spacing: 0.32px !important;
            text-transform: uppercase;
            text-shadow: none !important;
        }`;

        let existingStyle = styleTag.html() || '';
        existingStyle = existingStyle
            .replace(/\.theme-row-highlight,\s*\.theme-index-band\s*\{[^}]+\}/gs, '')
            .replace(/\.theme-row-highlight\s*\{[^}]+\}/gs, '')
            .replace(/\.theme-index-band\s*\{[^}]+\}/gs, '')
            .replace(/\.theme-index-entry\s*\{[^}]+\}/gs, '');

        styleTag.html(existingStyle + staticThemeCSS);
        console.log('✅ Static theme artifact CSS applied');
    }

    /**
     * Ensure sold-out badge CSS is present in the HTML with premium styling
     */
    ensureSoldOutBadgeCSS() {
        const styleTag = this.$('style').first();
        if (!styleTag.length) {
            console.warn('⚠️  No style tag found, skipping CSS check');
            return;
        }

        // Premium SOLD OUT badge CSS
        const soldOutBadgeCSS = `
        .sold-out-badge {
            background: linear-gradient(135deg, #dc2626 0%, #b91c1c 50%, #991b1b 100%);
            color: white;
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 8px;
            font-weight: 800;
            margin-left: 10px;
            display: inline-block;
            vertical-align: middle;
            line-height: 1;
            box-shadow: 0 2px 8px rgba(220, 38, 38, 0.4), 0 1px 3px rgba(0, 0, 0, 0.2);
            border: none;
            letter-spacing: 1.2px;
            text-transform: uppercase;
            position: relative;
            top: 0px;
            font-family: 'Montserrat', sans-serif;
        }
        `;

        let existingStyle = styleTag.html() || '';
        
        // Remove any existing sold-out-badge CSS block
        if (existingStyle.includes('.sold-out-badge')) {
            console.log('🎨 Updating sold-out badge CSS with premium styling...');
            // Remove old CSS block (handle different formatting)
            existingStyle = existingStyle.replace(/\.sold-out-badge\s*\{[^}]+\}/g, '');
            styleTag.html(existingStyle + soldOutBadgeCSS);
        } else {
            console.log('🎨 Injecting sold-out badge CSS...');
            styleTag.append(soldOutBadgeCSS);
        }
        console.log('✅ Premium sold-out badge CSS applied');
    }

    /**
     * Normalize time format (handle inconsistencies like "7,15 PM")
     * Ensures format is always HH:MM AM/PM with space before AM/PM
     */
    normalizeTime(time) {
        if (!time) return '';
        // Replace comma with colon and trim
        let normalized = time.replace(',', ':').trim();
        
        // Remove space before colon (e.g., "10 :00 AM" -> "10:00 AM")
        normalized = normalized.replace(/\s+:/, ':');
        
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

        console.log(`📊 Processing ${this.kwalityClasses.length} classes from kwalityClasses array...`);
        if (this.kwalityClasses.length > 0) {
            console.log(`📋 Sample classData keys:`, Object.keys(this.kwalityClasses[0]));
        }

        let excludedFromPdfCount = 0;
        
        this.kwalityClasses.forEach(classData => {
            if (this.shouldExcludeClassFromPdf(classData)) {
                excludedFromPdfCount++;
                return;
            }

            // Normalize day value to Title Case keys used in scheduleByDay
            const rawDay = String(classData.Day || classData.day || classData.DayName || '').trim();
            if (!rawDay) return; // skip records without day
            const day = rawDay.length <= 3 ? this.expandDayName(rawDay) : (rawDay.charAt(0).toUpperCase() + rawDay.slice(1).toLowerCase());
            if (scheduleByDay[day]) {
                // Check for theme data in explicit theme-related columns only
                // (Avoid using generic columns like Column H/G which can contain trainer names)
                let theme = classData.Theme || classData.theme || classData['Theme Name'] || 
                           classData['theme_name'] || classData['Class Theme'] || 
                           classData['class_theme'] || classData.Themes || classData['Theme(s)'] || '';
                
                // If no theme found in data, apply known theme patterns
                if (!theme || !theme.trim()) {
                    // Try to extract inline theme info from Notes or other free-text fields
                    const notesField = String(classData.Notes || classData.notes || classData.Note || '').trim();
                    const notesThemeMatch = notesField.match(/theme\s*[:\-]\s*(.+)$/i);
                    if (notesThemeMatch) {
                        theme = notesThemeMatch[1].trim();
                    } else {
                        theme = this.getThemeForClass(classData);
                    }
                }

                // Guardrail: if theme looks like trainer/class text, ignore it
                if (theme && theme.trim()) {
                    const themeLower = theme.trim().toLowerCase();
                    const trainerRaw = String(classData.Trainer || classData.trainer || '').trim();
                    const trainerFirst = this.getTrainerFirstName(trainerRaw).toLowerCase();
                    const classRaw = String(classData.Class || classData.class || '').trim();
                    const classLower = this.normalizeClassNameForDisplay(classRaw);

                    if ((trainerRaw && themeLower === trainerRaw.toLowerCase()) ||
                        (trainerFirst && themeLower === trainerFirst) ||
                        (classLower && themeLower === classLower)) {
                        theme = '';
                    }
                }
                
                scheduleByDay[day].push({
                    time: this.normalizeTime(classData.Time || classData.time || ''),
                    class: classData.Class || classData.class || '',
                    trainer: classData.Trainer || classData.trainer || '',
                    notes: classData.Notes || classData.notes || '',
                    theme: theme.trim() // Add theme information
                });
            }
        });

        if (excludedFromPdfCount > 0) {
            console.log(`🚫 Excluded ${excludedFromPdfCount} hosted / sold-out / trainerless classes from HTML/PDF output`);
        }

        return scheduleByDay;
    }

    /**
     * Get theme for a class based on known patterns when no theme data is in spreadsheet
     * NOTE: This function now returns empty - all themes should come from email parsing
     */
    getThemeForClass(classData) {
        // Only use themes that are explicitly parsed from the email body
        // No hardcoded themes - return empty string
        return '';
    }

    /**
     * Check whether theme text should render as static inline text instead of a badge.
     */
    shouldRenderStaticTheme() {
        return this.themeRenderMode === 'static';
    }

    /**
     * Create theme markup based on the current render mode.
     */
    createThemeMarkup(theme, location = 'kemps') {
        if (!theme || !theme.trim()) {
            return '';
        }

        return this.shouldRenderStaticTheme()
            ? ''
            : this.createThemeBadge(theme, location);
    }

    /**
     * Shared static-theme sizing should only consider rows that actually render a visible theme.
     */
    hasRenderableStaticTheme(theme = '') {
        const cleanTheme = String(theme || '').trim();
        return !!cleanTheme && cleanTheme.toLowerCase() !== 'sold out';
    }

    /**
     * Bandra static exports render the theme inline inside the row text.
     */
    shouldInlineStaticTheme(location = '') {
        // Bandra static schedules now mirror Kemps by keeping theme names in the
        // bottom legend instead of appending them inline next to trainer names.
        return false;
    }

    /**
     * Format the inline theme suffix appended to the trainer name.
     */
    formatInlineStaticThemeSuffix(theme, location = '') {
        if (!this.shouldInlineStaticTheme(location)) return '';

        const cleanTheme = String(theme || '').trim();
        if (!cleanTheme || cleanTheme.toLowerCase() === 'sold out') return '';

        return ` [${cleanTheme.toUpperCase()}]`;
    }

    /**
     * Inline theme markup for Bandra static rows uses slightly smaller, tighter type.
     */
    createInlineStaticThemeMarkup(theme, location = '') {
        const textSuffix = this.formatInlineStaticThemeSuffix(theme, location);
        if (!textSuffix) return '';

        return ` <span class="theme-inline-label" style="font-size:0.88em;letter-spacing:-0.32px;font-weight:500;display:inline-block;white-space:nowrap;">${textSuffix.trim()}</span>`;
    }

    /**
     * Themed Bandra static rows use white text on the highlight strip.
     */
    shouldUseWhiteStaticThemeText(theme, location = '') {
        const cleanTheme = String(theme || '').trim();
        return this.shouldInlineStaticTheme(location) && !!cleanTheme && cleanTheme.toLowerCase() !== 'sold out';
    }

    /**
     * Rough text-width estimate so static highlight strips can cover inline theme text.
     */
    estimateStaticTextWidth(text = '') {
        let width = 0;
        let insideThemeBrackets = false;

        for (const char of String(text)) {
            let charWidth;

            if (/[WMQGO0-9@#%&]/.test(char)) {
                charWidth = 9.2;
            } else if (/[A-Z]/.test(char)) {
                charWidth = 7.6;
            } else if (/[a-z]/.test(char)) {
                charWidth = 6.6;
            } else if (/[\[\](){}]/.test(char)) {
                charWidth = 4.8;
            } else if (/[-–]/.test(char)) {
                charWidth = 4.6;
            } else if (/\s/.test(char)) {
                charWidth = 3.7;
            } else {
                charWidth = 6;
            }

            if (char === '[') insideThemeBrackets = true;

            if (insideThemeBrackets) {
                charWidth *= 0.88;
            }

            width += charWidth;

            if (char === ']') insideThemeBrackets = false;
        }

        return Math.ceil(width);
    }

    hexToRgba(hex, alpha = 1) {
        const normalized = String(hex || '').trim().replace('#', '');
        if (!normalized) return `rgba(255, 255, 255, ${alpha})`;

        const expanded = normalized.length === 3
            ? normalized.split('').map((char) => char + char).join('')
            : normalized;

        const value = expanded.padEnd(6, 'f').slice(0, 6);
        const r = parseInt(value.slice(0, 2), 16);
        const g = parseInt(value.slice(2, 4), 16);
        const b = parseInt(value.slice(4, 6), 16);
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    getThemeSurfaceStyles(color, variant = 'row') {
        const baseColor = color || '#fff0a6';
        const background = variant === 'badge'
            ? `linear-gradient(135deg, ${this.hexToRgba(baseColor, 0.82)} 0%, ${this.hexToRgba(baseColor, 0.64)} 100%)`
            : `linear-gradient(135deg, ${this.hexToRgba(baseColor, 0.74)} 0%, ${this.hexToRgba(baseColor, 0.54)} 100%)`;

        return {
            background,
            border: 'none',
            boxShadow: variant === 'badge'
                ? '0 6px 14px rgba(69,59,42,0.12)'
                : '0 10px 22px rgba(69,59,42,0.10)',
            backdropFilter: 'blur(12px) saturate(140%)',
            WebkitBackdropFilter: 'blur(12px) saturate(140%)'
        };
    }

    /**
     * Estimate the sold-out strike width so it covers time + class + trainer text
     * and stops just before the SOLD OUT badge begins.
     */
    getSoldOutLineWidth({ dayLeft = 0, classLeft = 0, classText = '', maxWidth = 365 } = {}) {
        const prefixWidth = Math.max(0, Number(classLeft || 0) - Number(dayLeft || 0));
        const textWidth = this.estimateStaticTextWidth(classText || '');
            const desiredWidth = Math.ceil(prefixWidth + (textWidth * 1.06) + 16);
            return Math.max(150, Math.min(Number(maxWidth || 365), desiredWidth));
    }

    /**
     * Compute the highlight width needed to cover the full visible row text.
     */
    getDesiredStaticHighlightWidth({ highlightLeft = 0, textLeft = 0, text = '', baseWidth = 365 } = {}) {
        const prefixWidth = Math.max(0, Number(textLeft || 0) - Number(highlightLeft || 0));
        const textWidth = this.estimateStaticTextWidth(text);
        return Math.max(baseWidth, Math.ceil(prefixWidth + textWidth + 18));
    }

    getStaticThemeMaxWidthForRow(row = {}) {
        if (this.usesBandraStaticThemeGeometry(row.location)) {
            const timeAnchor = Number(row.timeLeft || row.highlightLeft || 0);
            const rowBottom = Number(row.rowBottom || 0);
            const isRightColumn = timeAnchor >= 430;
            const isLowerSection = rowBottom < 400;
            return isRightColumn ? (isLowerSection ? 396 : 412) : (isLowerSection ? 395 : 384);
        }

        const highlightLeft = Number(row.highlightLeft || 0);
        const isRightColumn = Number(row.timeLeft || highlightLeft || 0) >= 430;
        const maxRightBoundary = isRightColumn ? 860 : 451;
        return Math.max(365, Math.floor(maxRightBoundary - highlightLeft));
    }

    computeSharedStaticThemeWidth(rows = []) {
        const themedRows = rows.filter((row) => this.hasRenderableStaticTheme(row?.theme));
        if (!themedRows.length) return 365;

        const longestDesiredWidth = themedRows.reduce((maxWidth, row) => {
            const desiredWidth = this.getDesiredStaticHighlightWidth({
                highlightLeft: row.highlightLeft,
                textLeft: row.textLeft,
                text: row.classText || '',
                baseWidth: 365
            });
            return Math.max(maxWidth, desiredWidth);
        }, 365);

        const sharedSafeMaxWidth = themedRows.reduce((minWidth, row) => {
            return Math.min(minWidth, this.getStaticThemeMaxWidthForRow(row));
        }, Number.POSITIVE_INFINITY);

        return Math.max(365, Math.min(Math.ceil(longestDesiredWidth + 8), sharedSafeMaxWidth));
    }

    /**
     * Bandra static highlights should sit a touch tighter than the previous full-width version.
     */
    getReducedInlineStaticHighlightWidth(width = 0, location = '') {
        if (!this.shouldInlineStaticTheme(location)) return Math.round(Number(width || 0));
        return Math.round(Number(width || 0) * 0.95);
    }

    /**
     * Register a themed row so static mode can render a background bar and theme index entry later.
     */
    registerStaticThemeRow(rowConfig) {
        if (!this.shouldRenderStaticTheme()) return;
        if (!rowConfig || !this.hasRenderableStaticTheme(rowConfig.theme)) return;

        this.staticThemeRows.push({
            theme: String(rowConfig.theme).trim(),
            page: rowConfig.page,
            location: rowConfig.location || this.currentLocation,
            timeLeft: rowConfig.timeLeft,
            textLeft: rowConfig.textLeft,
            rowBottom: rowConfig.rowBottom,
            highlightWidth: rowConfig.highlightWidth,
            highlightHeight: rowConfig.highlightHeight || 23.2,
            highlightLeft: rowConfig.highlightLeft,
            classText: rowConfig.classText || ''
        });
    }

    /**
     * Build a stable theme->color map in first-seen order to mimic the fixed index file.
     */
    buildStaticThemeColorMap() {
        const palette = [
            '#ffd5e3', '#fff0a6', '#ecdeff', '#d0f1ff', '#d3fbc9', '#ffe4ff',
            '#ffddc4', '#ffd6e1', '#bafecf', '#dee3ff', '#B7D86F', '#ecebea',
            '#B886E4', '#f6fedb', '#ffd7cc', '#82C2EE', '#ffcf94', '#f6ffb6'
        ];
        const colorMap = new Map();
        const legendAssetMap = new Map();

        this.staticThemeRows.forEach((row) => {
            const themeKey = row.theme.trim().toUpperCase();
            if (!colorMap.has(themeKey)) {
                const nextIndex = colorMap.size;
                if (nextIndex < palette.length) {
                    colorMap.set(themeKey, palette[nextIndex]);
                } else {
                    const hue = (nextIndex * 47) % 360;
                    colorMap.set(themeKey, `hsl(${hue}, 85%, 78%)`);
                }

                if (this.usesBandraStaticThemeGeometry(row.location)) {
                    const highlightGeometry = this.resolveStaticThemeHighlightGeometry(row);
                    const legendAsset = this.getBandraStaticHighlightAsset(highlightGeometry);
                    legendAssetMap.set(themeKey, legendAsset);
                }
            }
        });

        this.staticThemeColorMap = colorMap;
        this.staticThemeLegendAssetMap = legendAssetMap;
    }

    /**
     * Bandra / Supreme HQ use the legacy static-strip layout from the reference export.
     */
    usesBandraStaticThemeGeometry(location = '') {
        const normalized = String(location || '').toLowerCase();
        return normalized.includes('bandra') || normalized.includes('supreme');
    }

    /**
     * Resolve the final geometry for a static theme highlight.
     * Bandra/Supreme use fixed strip positions that match the legacy export.
     */
    resolveStaticThemeHighlightGeometry(row) {
        const fallbackHeight = row.highlightHeight || 23.2;
        const sharedWidth = Number(this.staticThemeSharedWidth || 0);
        // rowBottom is tracked as the text baseline for the themed row.
        // Position the highlight to sit on that same row instead of shifting it
        // down by a full highlight height, which makes it appear on the next row.
        const fallbackBottom = (row.rowBottom || 0) - 1;
        const geometry = {
            left: row.highlightLeft,
            width: sharedWidth > 0 ? sharedWidth : row.highlightWidth,
            height: fallbackHeight,
            bottom: fallbackBottom
        };

        if (!this.usesBandraStaticThemeGeometry(row.location)) {
            return geometry;
        }

        const timeAnchor = Number(row.timeLeft || row.highlightLeft || 0);
        const rowBottom = Number(row.rowBottom || 0);
        const isRightColumn = timeAnchor >= 430;
        const isLowerSection = rowBottom < 400;
        const desiredWidth = sharedWidth > 0 ? sharedWidth : Number(row.highlightWidth || 0);
        const legacyDefaultWidth = isRightColumn
            ? (isLowerSection ? 396 : 412)
            : (isLowerSection ? 395 : 384);
        const maxWidth = isRightColumn ? 412 : ((Number(row.page || 1) >= 2) ? 420 : 412);
        const hasInlineThemeLabel = this.shouldInlineStaticTheme(row.location) && /\[[^\]]+\]/.test(row.classText || '');
        const reducedMaxWidth = this.getReducedInlineStaticHighlightWidth(maxWidth, row.location);

        if (isRightColumn) {
            geometry.left = isLowerSection ? 475 : 467;
            geometry.width = isLowerSection ? 396 : 412;
            geometry.height = 22;
        } else {
            geometry.left = 61;
            geometry.width = isLowerSection ? 395 : 384;
            geometry.height = 23;
        }

        if (desiredWidth > 0) {
            geometry.width = Math.min(
                reducedMaxWidth,
                Math.max(365, this.getReducedInlineStaticHighlightWidth(desiredWidth, row.location))
            );
        } else {
            geometry.width = legacyDefaultWidth;
        }

        if (hasInlineThemeLabel) {
            geometry.width = reducedMaxWidth;
        }

        // In the reference file the strip sits on the row baseline as a background,
        // not below the row like a separate block.
        geometry.bottom = rowBottom - 1;

        return geometry;
    }

    /**
     * Return the exact strip asset used by the legacy Bandra export for a given geometry.
     */
    getBandraStaticHighlightAsset(geometry) {
        if (!geometry) return null;
        const width = Math.round(Number(geometry.width || 0));
        const height = Math.round(Number(geometry.height || 0));
        const key = `${width}x${height}`;
        if (BANDRA_STATIC_STRIP_ASSETS[key]) {
            return BANDRA_STATIC_STRIP_ASSETS[key];
        }

        if (height === 23) {
            return width >= 395
                ? (BANDRA_STATIC_STRIP_ASSETS['395x23'] || BANDRA_STATIC_STRIP_ASSETS['384x23'] || null)
                : (BANDRA_STATIC_STRIP_ASSETS['384x23'] || BANDRA_STATIC_STRIP_ASSETS['395x23'] || null);
        }

        if (height === 22) {
            return width >= 412
                ? (BANDRA_STATIC_STRIP_ASSETS['412x22'] || BANDRA_STATIC_STRIP_ASSETS['396x22'] || null)
                : (BANDRA_STATIC_STRIP_ASSETS['396x22'] || BANDRA_STATIC_STRIP_ASSETS['412x22'] || null);
        }

        return null;
    }

    /**
     * Return the gradient strip to use in the static legend/index for a given theme.
     */
    getStaticThemeLegendAsset(themeKey) {
        return this.staticThemeLegendAssetMap.get(themeKey) || null;
    }

    /**
     * Determine which page contains a given span.
     */
    getPageNumberForSpan($span) {
        const $section = $span.closest('section.page');
        if ($section.length > 0) {
            const ariaLabel = $section.attr('aria-label') || '';
            const match = ariaLabel.match(/Page\s+(\d+)/i);
            if (match) return parseInt(match[1], 10);
        }

        if ($span.closest('#pg2, #pg2Overlay').length > 0) return 2;
        return 1;
    }

    /**
     * Render the static row highlights and theme index after rows are updated.
     */
    renderStaticThemeArtifacts() {
        if (!this.shouldRenderStaticTheme()) return;

        this.cleanupStaticThemeArtifacts();

        const validRows = this.staticThemeRows.filter(row => this.hasRenderableStaticTheme(row.theme));
        if (validRows.length === 0) {
            console.log('ℹ️  No themed rows found for static theme rendering');
            return;
        }

        this.staticThemeSharedWidth = this.computeSharedStaticThemeWidth(validRows);
        this.staticThemeIndexBandWidth = Math.max(82, Math.min(96, Math.round(this.staticThemeSharedWidth * 0.24)));

        this.buildStaticThemeColorMap();
        this.renderStaticThemeHighlights(validRows);
        const shouldRenderIndex = this.staticThemeColorMap.size > 0;
        if (shouldRenderIndex) {
            this.renderStaticThemeIndex(validRows);
        }
        const renderedIndexCount = shouldRenderIndex ? this.staticThemeColorMap.size : 0;
        console.log(`✅ Rendered ${validRows.length} static theme highlights and ${renderedIndexCount} index entries`);
    }

    /**
     * Render flat highlight bars behind themed rows.
     */
    renderStaticThemeHighlights(rows) {
        rows.forEach((row) => {
            const themeKey = row.theme.trim().toUpperCase();
            const color = this.staticThemeColorMap.get(themeKey) || '#fff0a6';
            const $page = this.$('section.page').eq(Math.max(0, (row.page || 1) - 1));
            const $container = $page.find('.text-container').first();
            if (!$container.length) return;

            const geometry = this.resolveStaticThemeHighlightGeometry(row);
            const surfaceStyles = this.getThemeSurfaceStyles(color, 'row');
            const highlightHtml = `<span class="theme-row-highlight" style="position:absolute;left:${geometry.left}px;bottom:${geometry.bottom}px;width:${geometry.width}px;height:${geometry.height}px;background:${surfaceStyles.background};z-index:1;display:block;border-radius:8px;border:${surfaceStyles.border};box-shadow:${surfaceStyles.boxShadow};backdrop-filter:${surfaceStyles.backdropFilter};-webkit-backdrop-filter:${surfaceStyles.WebkitBackdropFilter};opacity:1;"></span>`;
            $container.append(highlightHtml);
        });
    }

    /**
     * Render the theme names as an index near the end of the document, like index 7.html.
     */
    renderStaticThemeIndex(rows) {
        const orderedThemes = Array.from(this.staticThemeColorMap.keys());
        if (orderedThemes.length === 0) return;

        const $lastPage = this.$('section.page').last();
        const $container = $lastPage.find('.text-container').first();
        if (!$container.length) return;

        const startBandLeft = 482.7;
        const startBottom = 304;
        const lineGap = 37;
        const columnGap = 185;
        const maxRowsPerColumn = 5;
        const bandWidth = this.staticThemeIndexBandWidth || 80.2;
        const bandHeight = 23.2;
        const legendLabelGap = 18;

        orderedThemes.forEach((themeName, index) => {
            const column = Math.floor(index / maxRowsPerColumn);
            const row = index % maxRowsPerColumn;
            const bandLeft = startBandLeft + (column * columnGap);
            const textLeft = bandLeft + bandWidth + legendLabelGap;
            const bottom = startBottom - (row * lineGap);
            const bandBottom = bottom - 2;
            const color = this.staticThemeColorMap.get(themeName) || '#fff0a6';
            const surfaceStyles = this.getThemeSurfaceStyles(color, 'index');
            const bandHtml = `<span class="theme-index-band" style="position:absolute;left:${bandLeft}px;bottom:${bandBottom}px;width:${bandWidth}px;height:${bandHeight}px;background:${surfaceStyles.background};z-index:1;display:block;border-radius:8px;border:${surfaceStyles.border};box-shadow:${surfaceStyles.boxShadow};backdrop-filter:${surfaceStyles.backdropFilter};-webkit-backdrop-filter:${surfaceStyles.WebkitBackdropFilter};opacity:1;"></span>`;
            const entryHtml = `<span class="t v0 s8 theme-index-entry" style="left:${textLeft}px;bottom:${bottom}px;letter-spacing:0.32px;z-index:2;font-family:Montserrat-Bold_1z,Montserrat-Bold_21,Montserrat,sans-serif;font-size:12px;font-style:normal;font-weight:700;color:#453b2a;">${themeName}</span>`;
            $container.append(bandHtml);
            $container.append(entryHtml);
        });
    }

    /**
     * Pick one of the flat pastel highlight colors used in index 7.html.
     */
    getStaticThemeHighlightColor(theme) {
        const palette = ['#ffd5e3', '#fff0a6', '#ecdeff', '#d0f1ff', '#d3fbc9', '#ffe4ff'];
        const value = String(theme || '');
        let hash = 0;

        for (let i = 0; i < value.length; i++) {
            hash = ((hash << 5) - hash) + value.charCodeAt(i);
            hash |= 0;
        }

        return palette[Math.abs(hash) % palette.length];
    }

    /**
     * Create static inline theme text inspired by the older fixed-layout exports.
     */
    createStaticThemeLabel(theme, location = 'kemps') {
        const cleanTheme = theme.trim().toUpperCase();
        const highlightColor = this.getStaticThemeHighlightColor(cleanTheme);

        const staticStyle = {
            background: highlightColor,
            color: '#2C2D2D',
            marginLeft: '10px',
            display: 'inline-block',
            verticalAlign: 'middle',
            lineHeight: '1.5',
            letterSpacing: '0.21px',
            textTransform: 'uppercase',
            fontSize: '13px',
            fontWeight: '500',
            fontFamily: 'Montserrat, sans-serif',
            position: 'relative',
            top: '-1px',
            padding: '1px 12px 1px 12px',
            borderRadius: '0',
            whiteSpace: 'nowrap',
            boxShadow: 'none',
            border: 'none',
            transform: 'scaleX(1.006)',
            transformOrigin: 'left center'
        };

        const styleString = Object.entries(staticStyle)
            .map(([key, value]) => `${key.replace(/[A-Z]/g, match => '-' + match.toLowerCase())}: ${value}`)
            .join('; ');

        return `<span class="theme-static-label" style="${styleString}">${cleanTheme}</span>`;
    }

    /**
     * Create a neat theme badge for display
     */
    createThemeBadge(theme, location = 'kemps') {
        // Clean up the theme name
        let cleanTheme = theme.trim().toUpperCase();

        // Use the shared pastel palette for visual theme badges as well.
        const badgeColor = this.getStaticThemeHighlightColor(cleanTheme);
        
        // Use consistent ⚡️ icon for all badges
        const icon = '⚡️';
        
        const isBandra = location.toLowerCase().includes('bandra');
        const surfaceStyles = this.getThemeSurfaceStyles(badgeColor, 'badge');

        // Standardized styling for both locations using the requested flat pastel badge palette.
        const standardStyle = {
            background: surfaceStyles.background,
            color: '#2C2D2D',
            padding: isBandra ? '3px 10px' : '3px 11px',
            borderRadius: isBandra ? '8px' : '6px',
            fontSize: isBandra ? '8.5px' : '8px',
            fontWeight: '700',
            marginLeft: isBandra ? '9px' : '6px',
            display: 'inline-block',
            verticalAlign: 'middle',
            lineHeight: '1.3',
            boxShadow: surfaceStyles.boxShadow,
            letterSpacing: isBandra ? '0.28px' : '0.1px',
            textTransform: 'uppercase',
            minWidth: 'fit-content',
            maxWidth: isBandra ? '148px' : '180px',
            textAlign: 'center',
            whiteSpace: 'normal',
            wordWrap: 'break-word',
            border: surfaceStyles.border,
            position: 'relative',
            top: '-1px',
            backdropFilter: surfaceStyles.backdropFilter,
            WebkitBackdropFilter: surfaceStyles.WebkitBackdropFilter
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
        this.cleanupStaticThemeArtifacts();
        
        // Remove by CSS class - this removes the actual theme badge spans
        const themeDecorations = this.$('.theme-badge, .theme-static-label');
        console.log(`   Found ${themeDecorations.length} theme decoration elements to remove`);
        themeDecorations.remove();
        
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
            const hasStaticThemeClass = $span.hasClass('theme-static-label');
            
            // Only remove if it's clearly a theme badge, not a class description
            if (hasLightningEmoji || isStandaloneThemeKeyword || hasStaticThemeClass) {
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
        
        // Don't treat class names starting with "STUDIO " as headers
        // because they're actual class names that need to be updated
        if (/^STUDIO\s+/.test(text) && text.includes('-')) {
            // This looks like "STUDIO CLASSNAME - TRAINER", not a header
            return false;
        }
        
        const protectedKeywords = [
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
        
        // Check if text contains any protected keywords (but not if it's part of class-trainer format)
        if (!text.includes('-')) {
            // Only if there's NO hyphen (which would indicate trainer name)
            const containsKeyword = protectedKeywords.some(keyword => text.includes(keyword));
            if (containsKeyword) return true;
        }
        
        // Check if it looks like a date (contains "th" and month names)
        const looksLikeDate = /(?:january|february|march|april|may|june|july|august|september|october|november|december).*\d{1,2}(?:st|nd|rd|th)/i.test(text);
        
        // Check if text is very long (likely a header) - but not if it contains typical trainer separators
        const isLongText = text.length > 30 && !text.includes('-');
        
        return looksLikeDate || (isLongText && text.includes(':'));
    }

    /**
     * DYNAMIC ROW MANAGEMENT: Adds/removes rows based on sheet data
     * This method replaces the static update approach with a data-driven approach
     * where the number of rows per day matches exactly what's in the sheet.
     */
    dynamicUpdatePositionedSpans() {
        console.log('🔄 DYNAMIC UPDATE: Rebuilding schedule rows from sheet data...');
        const scheduleByDay = this.organizeScheduleByDay();

        // Configuration for row spacing and positioning
        const ROW_HEIGHT = 25; // pixels between rows (vertical spacing)
        const TIME_CLASS_OFFSET = 86; // horizontal offset from time to class span
        
        // Day header configuration - maps day names to their column positions
        const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        const dayHeaderPattern = /^(MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)\s*$/i;

        // Helper to determine which page a span is on
        const getPageForSpan = ($span) => {
            const $section = $span.closest('section.page');
            if ($section.length > 0) {
                const ariaLabel = $section.attr('aria-label') || '';
                if (ariaLabel.includes('Page 2')) return 2;
                if (ariaLabel.includes('Page 1')) return 1;
            }
            if ($span.closest('[id*="pg2"]').length > 0) return 2;
            if ($span.closest('[id*="pg1"]').length > 0) return 1;
            return 1;
        };

        // Build a list of all day header spans with their positions
        const dayHeaders = [];
        this.$('span').each((_i, elem) => {
            const $span = this.$(elem);
            const text = $span.text().trim();
            if (dayHeaderPattern.test(text)) {
                const style = $span.attr('style') || '';
                const leftMatch = style.match(/left:\s*([\d.]+)px/);
                const bottomMatch = style.match(/bottom:\s*([\d.]+)px/);
                const left = leftMatch ? parseFloat(leftMatch[1]) : 0;
                const bottom = bottomMatch ? parseFloat(bottomMatch[1]) : 0;
                const dayName = text.trim().charAt(0).toUpperCase() + text.trim().slice(1).toLowerCase();
                const page = getPageForSpan($span);
                const $container = $span.closest('section.page, [id^="pg"]');
                dayHeaders.push({ elem, $span, text: dayName, left, bottom, page, $container });
            }
        });

        console.log(`📅 Found ${dayHeaders.length} day headers`);
        dayHeaders.forEach(dh => {
            console.log(`  - ${dh.text}: left=${dh.left}px, bottom=${dh.bottom}px, page=${dh.page}`);
        });

        // For each day, find and remove all existing time/class spans, then regenerate
        let totalAdded = 0;
        let totalRemoved = 0;

        for (const dayHeader of dayHeaders) {
            const dayName = dayHeader.text;
            const classesForDay = scheduleByDay[dayName] || [];
            
            console.log(`\n📆 Processing ${dayName}: ${classesForDay.length} classes in sheet`);

            // Find all existing time spans for this day column
            // Time spans are within ~100px of the day header's left position
            const dayLeft = dayHeader.left;
            const dayBottom = dayHeader.bottom;
            const dayPage = dayHeader.page;
            const $dayContainer = dayHeader.$container;

            // Collect all time spans in this day's column
            const existingTimeSpans = [];
            this.$('span').each((_i, elem) => {
                const $span = this.$(elem);
                const text = $span.text().trim();
                
                // Check if it's a time span
                if (!/^\d{1,2}:\d{2}\s*(?:AM|PM)?$/i.test(text)) return;
                
                const style = $span.attr('style') || '';
                const leftMatch = style.match(/left:\s*([\d.]+)px/);
                const bottomMatch = style.match(/bottom:\s*([\d.]+)px/);
                const spanLeft = leftMatch ? parseFloat(leftMatch[1]) : 0;
                const spanBottom = bottomMatch ? parseFloat(bottomMatch[1]) : 0;
                const spanPage = getPageForSpan($span);

                // Must be on same page and in same column (within 20px tolerance for time spans)
                // Time spans are typically slightly to the left of day headers
                if (spanPage !== dayPage) return;
                if (Math.abs(spanLeft - dayLeft) > 30) return;
                
                // Must be below the day header (lower bottom value)
                if (spanBottom >= dayBottom) return;

                existingTimeSpans.push({
                    $span,
                    left: spanLeft,
                    bottom: spanBottom,
                    text
                });
            });

            // Sort by bottom position (highest first = top of column)
            existingTimeSpans.sort((a, b) => b.bottom - a.bottom);

            console.log(`  Found ${existingTimeSpans.length} existing time spans in HTML`);

            // Calculate the starting position for new rows
            // Use the highest existing time span position, or calculate from day header
            let startBottom;
            if (existingTimeSpans.length > 0) {
                startBottom = existingTimeSpans[0].bottom;
            } else {
                // Start below the day header with reduced spacing
                startBottom = dayBottom - 40; // Reduced offset from header to first class (was 60)
            }

            // Find the class spans associated with each time span and remove them
            const spansToRemove = new Set();
            
            for (const timeData of existingTimeSpans) {
                const $timeSpan = timeData.$span;
                spansToRemove.add($timeSpan[0]);

                // Find associated class/trainer span(s) - typically immediately after the time span
                let current = $timeSpan[0].nextSibling;
                while (current) {
                    if (current.type === 'tag' && current.name === 'span') {
                        const $currentSpan = this.$(current);
                        const currentText = $currentSpan.text().trim();
                        
                        // Stop at next time span
                        if (/^\d{1,2}:\d{2}\s*(?:AM|PM)?$/i.test(currentText)) break;
                        
                        // Stop at header elements
                        if (this.isHeaderElement($currentSpan)) break;

                        // Get position to verify it's in the same row
                        const currentStyle = $currentSpan.attr('style') || '';
                        const currentBottomMatch = currentStyle.match(/bottom:\s*([\d.]+)px/);
                        const currentBottom = currentBottomMatch ? parseFloat(currentBottomMatch[1]) : 0;

                        // If within same row (5px tolerance), mark for removal
                        if (Math.abs(currentBottom - timeData.bottom) <= 5) {
                            spansToRemove.add(current);
                        } else {
                            // Different row, stop scanning
                            break;
                        }
                    }
                    current = current.nextSibling;
                }
            }

            // Also remove any orphaned class/trainer spans at the same position
            // These might be left over from incomplete removals
            const timePositions = new Map();
            existingTimeSpans.forEach(td => {
                if (!timePositions.has(td.bottom)) {
                    timePositions.set(td.bottom, []);
                }
                timePositions.get(td.bottom).push(td);
            });

            this.$('span').each((_i, elem) => {
                const $span = this.$(elem);
                const style = $span.attr('style') || '';
                const leftMatch = style.match(/left:\s*([\d.]+)px/);
                const bottomMatch = style.match(/bottom:\s*([\d.]+)px/);
                const left = leftMatch ? parseFloat(leftMatch[1]) : 0;
                const bottom = bottomMatch ? parseFloat(bottomMatch[1]) : 0;
                const text = $span.text().trim();
                
                // Check if this is a class/trainer span (not a time span, not already marked)
                if (!spansToRemove.has(elem) && 
                    !/^\d{1,2}:\d{2}\s*(?:AM|PM)?$/i.test(text) &&
                    !this.isHeaderElement($span)) {
                    
                    // If there's a time span at this bottom position, and this span is at the class offset position
                    if (timePositions.has(bottom) && Math.abs(left - dayLeft - TIME_CLASS_OFFSET) <= 5) {
                        spansToRemove.add(elem);
                    }
                }
            });

            // Remove all marked spans
            const removedCount = spansToRemove.size;
            spansToRemove.forEach(elem => {
                this.$(elem).remove();
            });
            totalRemoved += removedCount;
            console.log(`  Removed ${removedCount} existing spans`);

            // Now generate new spans for each class in the sheet
            if (classesForDay.length === 0) {
                console.log(`  No classes for ${dayName}, skipping generation`);
                continue;
            }

            // Sort classes by time for proper ordering
            const sortedClasses = [...classesForDay].sort((a, b) => {
                const timeA = this.parseTimeToMinutes(a.time);
                const timeB = this.parseTimeToMinutes(b.time);
                return timeA - timeB;
            });

            // Find the annotation container or content container to insert new spans
            // Instead of trying to append to a container, insert directly after the day header
            // This ensures the spans appear in the right location in the final HTML
            const $dayHeader = dayHeader.$span;

            // Generate new time and class spans for each class
            let currentBottom = startBottom;
            let lastInsertedElement = $dayHeader[0];
            
            for (let i = 0; i < sortedClasses.length; i++) {
                const classData = sortedClasses[i];
                const rawTime = classData.time;
                // Normalize time: remove extra spaces, special characters
                const time = this.normalizeTime(rawTime).replace(/\s+/g, ' ').trim();
                const className = this.formatClassName(this.normalizeClassName(classData.class)).replace(/^STUDIO\s+/i, '');
                const trainerName = this.getTrainerFirstName(classData.trainer).toUpperCase();
                const theme = classData.theme || '';
                const inlineThemeSuffix = this.formatInlineStaticThemeSuffix(theme, this.currentLocation);
                const inlineThemeMarkup = this.createInlineStaticThemeMarkup(theme, this.currentLocation);
                const useWhiteThemedText = this.shouldUseWhiteStaticThemeText(theme, this.currentLocation);
                
                // Debug logging to verify data
                console.log(`  📝 [${dayName}] ${time} - Class: "${className}", Trainer: "${trainerName}", Theme: "${theme}"`);
                if (!className || !trainerName) {
                    console.warn(`  ⚠️  Empty values detected - classData:`, classData);
                }
                // Check if sold out - either from notes or theme
                const isSoldOut = (classData.notes && classData.notes.includes('SOLD OUT')) || 
                                  (theme && theme.toLowerCase().trim() === 'sold out');



                // Create time span with normalized time
                const timeSpanHtml = `<span class="t s9" style="left:${dayLeft}px;bottom:${currentBottom}px;color:${useWhiteThemedText ? '#ffffff' : '#1a1a1a'};">${time}</span>`;
                
                // Create class/trainer span
                let classText = className;
                if (trainerName) {
                    classText += ` - ${trainerName}`;
                }
                classText += inlineThemeSuffix;
                const classTextHtml = inlineThemeMarkup
                    ? `${className}${trainerName ? ` - ${trainerName}` : ''}${inlineThemeMarkup}`
                    : classText;
                
                // Build badge HTML
                let badgeHtml = '';
                // Only add theme badge if it's not "Sold Out" (sold out badge added separately)
                if (theme && theme.trim() && theme.toLowerCase().trim() !== 'sold out' && !this.shouldInlineStaticTheme(this.currentLocation)) {
                    badgeHtml += this.createThemeMarkup(theme.trim(), this.currentLocation);
                }
                if (isSoldOut) {
                    badgeHtml += ' <span class="sold-out-badge">SOLD OUT</span>';
                }

                const classLeft = dayLeft + TIME_CLASS_OFFSET;
                const highlightLeft = Math.max(dayLeft - 2, 0);
                const highlightWidth = this.shouldInlineStaticTheme(this.currentLocation)
                    ? this.getDesiredStaticHighlightWidth({
                        highlightLeft,
                        textLeft: classLeft,
                        text: classText,
                        baseWidth: 365
                    })
                    : 365;
                // Create red strikethrough line for sold out classes that covers the time and class name (but not the badge)
                let strikethroughHtml = '';
                if (isSoldOut) {
                    const soldOutLineWidth = this.getSoldOutLineWidth({
                        dayLeft,
                        classLeft,
                        classText,
                        maxWidth: highlightWidth
                    });
                    strikethroughHtml = `<span class="sold-out-line" style="position: absolute; left:${dayLeft}px; bottom:${currentBottom + 8}px; width: ${soldOutLineWidth}px; height: 2px; background-color: #dc143c; z-index: 10; pointer-events: none;"></span>`;
                }
                
                const classSpanHtml = `<span class="t v0 s5" style="left:${classLeft}px;bottom:${currentBottom}px;font-family:Montserrat,sans-serif;font-weight:400;color:${useWhiteThemedText ? '#ffffff' : '#1a1a1a'};">${classTextHtml}${badgeHtml}</span>`;

                if (theme && theme.trim() && theme.toLowerCase().trim() !== 'sold out' && this.shouldRenderStaticTheme()) {
                    this.registerStaticThemeRow({
                        theme,
                        page: dayPage,
                        location: this.currentLocation,
                        timeLeft: dayLeft,
                        textLeft: classLeft,
                        rowBottom: currentBottom,
                        highlightLeft,
                        highlightWidth,
                        classText
                    });
                }

                // Parse and insert time span
                const timeSpan = this.$(timeSpanHtml);
                this.$(lastInsertedElement).after(timeSpan);
                lastInsertedElement = timeSpan[0];
                
                // Insert red strikethrough line if sold out (before class span so it appears behind)
                if (isSoldOut) {
                    const strikethroughSpan = this.$(strikethroughHtml);
                    this.$(lastInsertedElement).after(strikethroughSpan);
                    lastInsertedElement = strikethroughSpan[0];
                }
                
                // Parse and insert class span
                const classSpan = this.$(classSpanHtml);
                this.$(lastInsertedElement).after(classSpan);
                lastInsertedElement = classSpan[0];

                totalAdded += isSoldOut ? 3 : 2;

                // Move to next row
                currentBottom -= ROW_HEIGHT;
            }

            console.log(`  Generated ${sortedClasses.length * 2} new spans for ${sortedClasses.length} classes`);
        }

        console.log(`\n✅ DYNAMIC UPDATE COMPLETE:`);
        console.log(`   - Removed: ${totalRemoved} old spans`);
        console.log(`   - Added: ${totalAdded} new spans`);

        // Clean up any remaining malformed times in the HTML
        this.cleanupMalformedTimes();

        // Post-processing
        this.normalizeAllContentSpans();
    }

    /**
     * OLD METHOD: Generate new time and class spans for each class
     * (Kept for reference - not actively used)
     */
    generateNewSpansOLD() {
        // This old code is deprecated
        return;
    }

    /**
     * Clean up any remaining malformed times in the HTML
     * Removes spans that contain malformed times like "10 :00 AM" (space before colon)
     */
    cleanupMalformedTimes() {
        console.log('🧹 Cleaning up malformed times from HTML...');
        const malformedPattern = /\d{1,2}\s+:\d{2}\s*(AM|PM)/i;
        let removedCount = 0;

        this.$('span').each((_i, elem) => {
            const $span = this.$(elem);
            const text = $span.text().trim();
            
            // Check if this is a malformed time span (has space before colon)
            if (malformedPattern.test(text) && /^\d{1,2}\s+:\d{2}\s*(AM|PM)$/.test(text)) {
                // Only remove if it's exactly a time span (not mixed with other text)
                $span.remove();
                removedCount++;
            }
        });
        
        if (removedCount > 0) {
            console.log(`  ✓ Removed ${removedCount} malformed time spans`);
        } else {
            console.log('  ✓ No malformed times found');
        }
    }

    /**
     * Parse time string to minutes since midnight for sorting
     */
    parseTimeToMinutes(timeStr) {
        if (!timeStr) return 0;
        const normalized = this.normalizeTime(timeStr);
        const match = normalized.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
        if (!match) return 0;
        
        let hours = parseInt(match[1], 10);
        const minutes = parseInt(match[2], 10);
        const period = match[3].toUpperCase();
        
        if (period === 'PM' && hours !== 12) hours += 12;
        if (period === 'AM' && hours === 12) hours = 0;
        
        return hours * 60 + minutes;
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

        // Helper to determine which page a span is on
        const getPageForSpan = ($span) => {
            // Check section.page parent with aria-label
            const $section = $span.closest('section.page');
            if ($section.length > 0) {
                const ariaLabel = $section.attr('aria-label') || '';
                if (ariaLabel.includes('Page 2')) return 2;
                if (ariaLabel.includes('Page 1')) return 1;
            }
            // Fallback: check for pg1Overlay or pg2Overlay ancestor
            if ($span.closest('[id*="pg2"]').length > 0) return 2;
            if ($span.closest('[id*="pg1"]').length > 0) return 1;
            return 1; // Default to page 1
        };

        // NEW APPROACH: Find all day headers and build a map of spans to days
        // Day headers are spans containing day names like "MONDAY", "TUESDAY", etc.
        const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        const dayHeaderPattern = /^(MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)\s*$/i;
        
        // Build a list of all day header spans with their positions
        const dayHeaders = [];
        this.$('span').each((_i, elem) => {
            const $span = this.$(elem);
            const text = $span.text().trim();
            if (dayHeaderPattern.test(text)) {
                const style = $span.attr('style') || '';
                const leftMatch = style.match(/left:\s*([\d.]+)px/);
                const bottomMatch = style.match(/bottom:\s*([\d.]+)px/);
                const left = leftMatch ? parseFloat(leftMatch[1]) : 0;
                const bottom = bottomMatch ? parseFloat(bottomMatch[1]) : 0;
                const dayName = text.trim().charAt(0).toUpperCase() + text.trim().slice(1).toLowerCase();
                const page = getPageForSpan($span);
                dayHeaders.push({ elem, $span, text: dayName, left, bottom, page });
            }
        });
        
        console.log(`📅 Found ${dayHeaders.length} day headers in HTML`);
        dayHeaders.forEach(dh => {
            console.log(`  - ${dh.text} at left:${dh.left}px, bottom:${dh.bottom}px, page:${dh.page}`);
        });
        
        // Helper to find the day for a time span based on day headers
        // Logic: Find the day header that is closest to this span (same column, higher bottom value)
        const findDayByHeader = ($timeSpan) => {
            const style = $timeSpan.attr('style') || '';
            const leftMatch = style.match(/left:\s*([\d.]+)px/);
            const bottomMatch = style.match(/bottom:\s*([\d.]+)px/);
            const spanLeft = leftMatch ? parseFloat(leftMatch[1]) : 0;
            const spanBottom = bottomMatch ? parseFloat(bottomMatch[1]) : 0;
            const spanPage = getPageForSpan($timeSpan);
            
            // Find day headers in the same column (within 50px tolerance) and same page
            // that are above this span (higher bottom value)
            let bestMatch = null;
            let bestDist = Infinity;
            
            for (const dh of dayHeaders) {
                // Must be on same page
                if (dh.page !== spanPage) continue;
                
                // Must be in same column (similar left position, within 100px tolerance)
                const leftDiff = Math.abs(dh.left - spanLeft);
                if (leftDiff > 100) continue;
                
                // Day header should be above or at the same level as the time span (higher or equal bottom)
                if (dh.bottom < spanBottom) continue;
                
                // Calculate distance - prefer headers that are closest above
                const dist = dh.bottom - spanBottom + leftDiff * 0.1; // Weight left difference slightly
                if (dist < bestDist) {
                    bestDist = dist;
                    bestMatch = dh;
                }
            }
            
            return bestMatch ? bestMatch.text : null;
        };

        // Check if this is a multi-page PDF (Bandra style with pg1 and pg2, or section.page elements)
        const hasPage1 = this.$('#pg1').length > 0 || this.$('section.page[aria-label*="Page 1"]').length > 0;
        const hasPage2 = this.$('#pg2').length > 0 || this.$('section.page[aria-label*="Page 2"]').length > 0;
        const isMultiPagePDF = hasPage1 && hasPage2;
        console.log(`📄 Multi-page PDF detection: hasPage1=${hasPage1}, hasPage2=${hasPage2}, isMultiPagePDF=${isMultiPagePDF}`);
        
        // For Bandra multi-page PDF:
        // Page 1: Mon, Tue, Wed, Thu (4 days)
        // Page 2: Fri, Sat, Sun (3 days)
        const page1Days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday'];
        const page2Days = ['Friday', 'Saturday', 'Sunday'];

        // Collect all time spans
        const timeSpans = this.$('span').filter((_i, elem) => {
            return /^\d{1,2}:\d{2}\s*(?:AM|PM)$/i.test(this.$(elem).text().trim());
        }).get();

        console.log(`📄 Multi-page PDF detected: ${isMultiPagePDF}`);

        // Build column clusters by x-position (left). Tolerance ~ 20px
        // For multi-page, cluster separately for each page
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

        let clusters, dayByColumnIndex;
        let page1Clusters = [], page2Clusters = [];
        let page1DayByColumn = {}, page2DayByColumn = {};

        if (isMultiPagePDF) {
            // Separate clustering for each page
            const page1Lefts = [], page2Lefts = [];
            timeSpans.forEach((el) => {
                const $el = this.$(el);
                const pos = getSpanPosition($el);
                if (!Number.isNaN(pos.left)) {
                    if (getPageForSpan($el) === 1) {
                        page1Lefts.push(pos.left);
                    } else {
                        page2Lefts.push(pos.left);
                    }
                }
            });

            page1Clusters = clusterPositions(page1Lefts, 20);
            page2Clusters = clusterPositions(page2Lefts, 20);
            
            // Keep most populated clusters for each page
            if (page1Clusters.length > page1Days.length) {
                page1Clusters.sort((a, b) => b.count - a.count);
                page1Clusters = page1Clusters.slice(0, page1Days.length);
            }
            page1Clusters.sort((a, b) => a.center - b.center);
            
            if (page2Clusters.length > page2Days.length) {
                page2Clusters.sort((a, b) => b.count - a.count);
                page2Clusters = page2Clusters.slice(0, page2Days.length);
            }
            page2Clusters.sort((a, b) => a.center - b.center);

            // Map clusters to days
            page1Clusters.forEach((c, idx) => {
                if (idx < page1Days.length) page1DayByColumn[idx] = page1Days[idx];
            });
            page2Clusters.forEach((c, idx) => {
                if (idx < page2Days.length) page2DayByColumn[idx] = page2Days[idx];
            });

            console.log(`📄 Page 1: ${page1Clusters.length} columns detected for ${page1Days.join(', ')}`);
            console.log(`📄 Page 2: ${page2Clusters.length} columns detected for ${page2Days.join(', ')}`);
        } else {
            // Original single-page logic
            const lefts = [];
            timeSpans.forEach((el) => {
                const pos = getSpanPosition(this.$(el));
                if (!Number.isNaN(pos.left)) lefts.push(pos.left);
            });

            clusters = clusterPositions(lefts, 20);
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
            dayByColumnIndex = {};
            clusters.forEach((_c, idx) => {
                if (idx < dayOrder.length) dayByColumnIndex[idx] = dayOrder[idx];
            });
        }

        const findDayForSpan = ($span) => {
            // First try day header-based detection (more accurate for Bandra-style layouts)
            const dayFromHeader = findDayByHeader($span);
            if (dayFromHeader) {
                return dayFromHeader;
            }
            
            // Fall back to column-based detection
            const pos = getSpanPosition($span);
            if (Number.isNaN(pos.left)) return null;

            if (isMultiPagePDF) {
                const page = getPageForSpan($span);
                const pageClusters = page === 1 ? page1Clusters : page2Clusters;
                const pageDayByColumn = page === 1 ? page1DayByColumn : page2DayByColumn;
                
                if (pageClusters.length === 0) return null;
                
                // Find nearest cluster center
                let bestIdx = 0;
                let bestDist = Math.abs(pos.left - pageClusters[0].center);
                for (let i = 1; i < pageClusters.length; i++) {
                    const d = Math.abs(pos.left - pageClusters[i].center);
                    if (d < bestDist) {
                        bestDist = d;
                        bestIdx = i;
                    }
                }
                return pageDayByColumn[bestIdx] || null;
            } else {
                if (!clusters || clusters.length === 0) return null;
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
            }
        };

        // First, clean up all existing theme badges to prevent duplicates
        this.cleanupAllThemeBadges();

        // Track updated day+time to prevent duplicates within a day
        const updatedCombos = new Set();
        
        // Track which data records have been used (to handle multiple classes at same time)
        const usedRecordKeys = new Set();

        timeSpans.forEach((timeElem, index) => {
            const $timeSpan = this.$(timeElem);
            const timeText = $timeSpan.text().trim().replace(/\s/g, ' ').trim();

            // Determine day by nearest x-position cluster
            const detectedDay = findDayForSpan($timeSpan);
            if (!detectedDay) {
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
                if (csvTimeCompare.toLowerCase() !== htmlTimeCompare.toLowerCase()) {
                    return false;
                }
                // Exclude already-used records
                const recordKey = `${detectedDay}|${c.time}|${c.class}|${c.trainer}`;
                return !usedRecordKeys.has(recordKey);
            });
            
            if (timeMatches.length > 1 && htmlClassName) {
                // Multiple classes at same time - match by class name too
                matchingClass = timeMatches.find(c => {
                    const csvClassName = this.normalizeClassName(c.class).toUpperCase();
                    const matches = csvClassName.includes(htmlClassName) || htmlClassName.includes(csvClassName);
                    return matches;
                });
            }
            
            // If still no match and multiple options, prefer non-sold-out over sold-out
            if (!matchingClass && timeMatches.length > 1) {
                const nonSoldOut = timeMatches.find(c => !c.notes || !c.notes.includes('SOLD OUT'));
                if (nonSoldOut) {
                    matchingClass = nonSoldOut;
                } else {
                    matchingClass = timeMatches[0];
                }
            }
            
            // Fall back to first unused time match if no class name match found
            if (!matchingClass && timeMatches.length > 0) {
                matchingClass = timeMatches[0];
            }

            if (matchingClass) {
                // Mark this record as used
                const recordKey = `${detectedDay}|${matchingClass.time}|${matchingClass.class}|${matchingClass.trainer}`;
                usedRecordKeys.add(recordKey);
                
                // Create a unique key for this day+time+class+trainer combination to prevent duplicates
                const normalizedClass = this.normalizeClassName(matchingClass.class).toUpperCase();
                const normalizedTrainer = (matchingClass.trainer || '').trim().toUpperCase();
                const elementIndex = $timeSpan.index();
                const combinationKey = `${detectedDay}|${timeText}|${normalizedClass}|${normalizedTrainer}|${elementIndex}`;
                if (updatedCombos.has(combinationKey)) {
                    console.log(`    ⚠ Skipping duplicate: already updated this day/time/class/trainer combination`);
                    return; // Skip this duplicate
                }
                
                // Get time span's position to find related content spans
                const timeSpanStyle = $timeSpan.attr('style') || '';
                const timeSpanBottomMatch = timeSpanStyle.match(/bottom:\s*([\d.]+)px/);
                const timeSpanBottom = timeSpanBottomMatch ? parseFloat(timeSpanBottomMatch[1]) : 0;
                const timeSpanLeftMatch = timeSpanStyle.match(/left:\s*([\d.]+)px/);
                const timeSpanLeft = timeSpanLeftMatch ? parseFloat(timeSpanLeftMatch[1]) : 0;
                
                let current = $timeSpan[0].nextSibling;
                const spansToRemove = [];
                let firstContentSpan = null;
                let siblingsProcessed = 0;

                while (current) {
                    siblingsProcessed++;
                    
                    if (current.type === 'tag' && current.name === 'span') {
                        const $currentSpan = this.$(current);
                        const spanText = $currentSpan.text().trim();
                        const spanId = $currentSpan.attr('id');
                        const spanClass = $currentSpan.attr('class');
                        
                        // Get this span's position
                        const currentStyle = $currentSpan.attr('style') || '';
                        const currentBottomMatch = currentStyle.match(/bottom:\s*([\d.]+)px/);
                        const currentBottom = currentBottomMatch ? parseFloat(currentBottomMatch[1]) : 0;
                        const currentLeftMatch = currentStyle.match(/left:\s*([\d.]+)px/);
                        const currentLeft = currentLeftMatch ? parseFloat(currentLeftMatch[1]) : 0;
                        
                        if (/^\d{1,2}:\d{2}\s*(?:AM|PM)$/i.test(spanText)) {
                            break; // Stop at the next time span
                        }
                        
                        // Check if this is a header/protected element
                        if (this.isHeaderElement($currentSpan)) {
                            break;
                        }
                        
                        // Check if this span is at the same bottom position (same row, within 5px tolerance)
                        const sameRow = Math.abs(currentBottom - timeSpanBottom) <= 5;
                        
                        // Enhanced badge removal - check for CSS classes, inline patterns, and content
                        const hasThemeClass = $currentSpan.hasClass('theme-badge') || $currentSpan.hasClass('theme-static-label');
                        const hasOldTheme = /[⚡️⚡]/.test(spanText);
                        const hasOldThemeText = /\b(?:theme|special)\b/i.test(spanText);
                        
                        // Check if this span is a trainer-only span (starts with "- " or just a name after a hyphen)
                        const isTrainerSpan = /^-\s*[A-Za-z]+/.test(spanText) || /^[A-Z][a-z]+\s*$/.test(spanText);
                        
                        // Only mark for removal if it's not the first span (which contains class info) 
                        // OR if it's clearly a theme badge
                        // OR if it's in the same row and appears to be part of the class info
                        if (!firstContentSpan) {
                            firstContentSpan = $currentSpan;
                            // For the first span, we'll replace its content entirely, so always add to removal list
                            spansToRemove.push($currentSpan);
                        } else if (sameRow && (isTrainerSpan || spanText.length === 0)) {
                            // Same row trainer span or empty span - mark for removal
                            spansToRemove.push($currentSpan);
                        } else if (hasThemeClass || hasOldTheme || hasOldThemeText) {
                            // Remove subsequent spans only if they contain actual theme badge content
                            spansToRemove.push($currentSpan);
                        }
                    } else if (current.type === 'tag' && current.name !== 'span') {
                        break;
                    }
                    current = current.nextSibling;
                }

                if (firstContentSpan && spansToRemove.length > 0) {
                    const normalizedCSVClass = this.normalizeClassName(matchingClass.class);
                    // Remove "Studio " prefix for display in HTML/PDF
                    let classDisplay = this.formatClassName(normalizedCSVClass)
                        .replace(/^STUDIO\s+/i, ''); // Remove "STUDIO " prefix
                    const trainerFirstName = this.getTrainerFirstName(matchingClass.trainer);
                    const trainerDisplay = trainerFirstName.toUpperCase();
                    
                    // Check if this is a sold-out/hosted class
                    const isSoldOut = matchingClass.notes && matchingClass.notes.includes('SOLD OUT');
                    const useWhiteThemedText = this.shouldUseWhiteStaticThemeText(matchingClass.theme, this.currentLocation);
                    $timeSpan.css('color', useWhiteThemedText ? '#ffffff' : '#1a1a1a');
                    
                    
                    let newText = classDisplay;
                    if (trainerDisplay) {
                        newText += ` - ${trainerDisplay}`;
                    }
                    const inlineThemeSuffix = this.formatInlineStaticThemeSuffix(matchingClass.theme, this.currentLocation);
                    const inlineThemeMarkup = this.createInlineStaticThemeMarkup(matchingClass.theme, this.currentLocation);
                    newText += inlineThemeSuffix;

                    // Add theme badge if theme exists
                    let themeBadge = '';
                    if (matchingClass.theme && matchingClass.theme.trim() && !this.shouldInlineStaticTheme(this.currentLocation)) {
                        themeBadge = this.createThemeMarkup(matchingClass.theme.trim(), this.currentLocation);
                    }

                    if (matchingClass.theme && matchingClass.theme.trim() && matchingClass.theme.toLowerCase().trim() !== 'sold out' && this.shouldRenderStaticTheme()) {
                        this.registerStaticThemeRow({
                            theme: matchingClass.theme.trim(),
                            page: this.getPageNumberForSpan($timeSpan),
                            location: this.currentLocation,
                            timeLeft: timeSpanLeft,
                            textLeft: timeSpanLeft + 86,
                            rowBottom: timeSpanBottom,
                            highlightLeft: Math.max(timeSpanLeft - 2, 0),
                            highlightWidth: this.shouldInlineStaticTheme(this.currentLocation)
                                ? this.getDesiredStaticHighlightWidth({
                                    highlightLeft: Math.max(timeSpanLeft - 2, 0),
                                    textLeft: timeSpanLeft + 86,
                                    text: newText,
                                    baseWidth: 365
                                })
                                : 365,
                            classText: newText
                        });
                    }

                    // Create a new span with the content, preserving the original's attributes
                    const newSpan = firstContentSpan.clone().text(trainerDisplay ? `${classDisplay} - ${trainerDisplay}` : classDisplay);
                    
                    // Remove old sold-out styling
                    newSpan.removeClass('sold-out');
                    
                    // Remove old sold-out badges
                    newSpan.find('.sold-out-badge').remove();
                    
                    // Apply sold-out styling if needed
                    if (isSoldOut) {
                        newSpan.addClass('sold-out');
                    }
                    
                    // Add theme badge as HTML if it exists
                    let badgeHTML = themeBadge || '';
                    if (isSoldOut) {
                        badgeHTML += ' <span class="sold-out-badge">SOLD OUT</span>';
                    }
                    
                    if (inlineThemeMarkup) {
                        const currentHTML = newSpan.html();
                        newSpan.html(currentHTML + inlineThemeMarkup);
                    }

                    if (badgeHTML) {
                        // Append badges to the span
                        const currentHTML = newSpan.html();
                        newSpan.html(currentHTML + badgeHTML);
                    }
                    
                    // Apply consistent Montserrat font with regular weight for all days
                    newSpan.css('font-family', 'Montserrat, sans-serif');
                    newSpan.css('font-weight', '400');
                    newSpan.css('color', useWhiteThemedText ? '#ffffff' : '#1a1a1a');
                    newSpan.css('letter-spacing', '-0.1px');
                    newSpan.css('font-style', 'normal');
                    newSpan.css('text-transform', 'none');

                    // Insert the new span after the time span
                    $timeSpan.after(newSpan);

                    // Remove all the old content spans
                    spansToRemove.forEach(($span, idx) => {
                        $span.remove();
                    });
                    
                    // Mark this combination as updated
                    updatedCombos.add(combinationKey);
                    updateCount++;
                }
            }
        });

        console.log(`\n✅ Updated ${updateCount} positioned span elements`);
        
        // Clean up old sold-out badges and red lines from classes that are no longer sold out
        this.cleanupOldSoldOutElements();
        
        // Post-processing: Normalize all class/trainer content spans for consistent styling
        this.normalizeAllContentSpans();
    }

    /**
     * Clean up sold-out badges and red lines from classes that are no longer sold out
     */
    cleanupOldSoldOutElements() {
        console.log(`🧹 Cleaning up old sold-out badges and red lines... (Current file: ${this.outputPath})`);
        
        // Use the same data source as the creation logic - kwalityClasses array
        const currentSoldOutClasses = new Set();
        
        // Process the same data that was used for span creation
        this.kwalityClasses.forEach(classData => {
            const theme = classData.Theme || '';
            const notes = classData.Notes || '';
            const isSoldOut = (notes && notes.includes('SOLD OUT')) || 
                             (theme && theme.toLowerCase().trim() === 'sold out');
            
            if (isSoldOut) {
                const day = classData.Day;
                const time = this.normalizeTime(classData.Time);
                const key = `${day}-${time}`;
                currentSoldOutClasses.add(key);
            }
        });
        
        let badgeCleanupCount = 0;
        let lineCleanupCount = 0;

        console.log(`📊 Found ${currentSoldOutClasses.size} currently sold out classes`);
        
        // Clean up sold-out badges from spans that no longer match sold-out classes
        this.$('span.sold-out-badge').each((_, elem) => {
            const $badge = this.$(elem);
            let $parentSpan = $badge.closest('span');

            // If the badge appears to be a standalone span (closest span is itself)
            // or its parent has no positioning style, try to attach it to the
            // nearest styled content span so positioning and cleanup heuristics work.
            const parentStyle = $parentSpan.attr('style') || '';
            if (!$parentSpan.attr('style') || $parentSpan.is($badge)) {
                // Prefer the next styled sibling (likely the class span), else previous
                const $nextStyled = $badge.nextAll('span[style]').first();
                const $prevStyled = $badge.prevAll('span[style]').first();
                const $target = $nextStyled.length ? $nextStyled : ($prevStyled.length ? $prevStyled : null);
                if ($target) {
                    console.log(`    ⚙️  Moving standalone sold-out badge into nearest styled span`);
                    $target.append($badge);
                    $parentSpan = $target;
                }
            }
            
            // Try to determine day and time from parent span position and nearby time spans
            const day = this.extractDayFromSpanPosition($parentSpan);
            const time = this.extractTimeFromNearbySpans($parentSpan);
            
            if (day && time) {
                const key = `${day}-${this.normalizeTime(time)}`;
                if (!currentSoldOutClasses.has(key)) {
                    console.log(`    Removing sold-out badge from: ${day} ${time}`);
                    $badge.remove();
                    badgeCleanupCount++;
                }
            } else {
                // If we can't determine day/time, do NOT remove the badge automatically.
                // Removing here has been observed to delete legitimate sold-out badges
                // because position parsing can be brittle across templates. Keep the
                // badge and log a warning so issues can be investigated manually.
                console.log(`    ⚠️  Could not determine day/time for sold-out badge — keeping it (parent span style: ${$parentSpan.attr('style') || 'none'})`);
            }
        });
        
        // Clean up red strikethrough lines
        // Deduplicate obvious duplicate red lines (same style attribute)
        let lineCleanupDupCount = 0;
        const soldOutLines = this.$('span.sold-out-line');
        console.log(`🔍 Found ${soldOutLines.length} sold-out-line elements to check`);
        const seenLineStyles = new Set();
        soldOutLines.each((_, el) => {
            const $line = this.$(el);
            const style = $line.attr('style') || '';
            if (seenLineStyles.has(style)) {
                // Remove exact-duplicate visual lines
                $line.remove();
                lineCleanupDupCount++;
            } else {
                seenLineStyles.add(style);
            }
        });
        if (lineCleanupDupCount > 0) {
            console.log(`    Removed ${lineCleanupDupCount} duplicate sold-out-line elements`);
            lineCleanupCount += lineCleanupDupCount;
        }

        const seenSoldOutLineKeys = new Set();
        this.$('span.sold-out-line').each((_, elem) => {
            const $line = this.$(elem);
            console.log(`  Checking line: ${$line.attr('style')}`);
            
            // Try to determine day and time from line position
            const day = this.extractDayFromLinePosition($line);
            const time = this.extractTimeFromLinePosition($line);
            console.log(`  Extracted day: "${day}", time: "${time}"`);
            
            if (day && time) {
                const key = `${day}-${this.normalizeTime(time)}`;
                console.log(`  Checking key "${key}" in currentSoldOutClasses:`, currentSoldOutClasses.has(key));
                if (seenSoldOutLineKeys.has(key)) {
                    console.log(`    Removing duplicate sold-out line for: ${day} ${time}`);
                    $line.remove();
                    lineCleanupCount++;
                    return;
                }

                if (!currentSoldOutClasses.has(key)) {
                    console.log(`    Removing sold-out line from: ${day} ${time}`);
                    $line.remove();
                    lineCleanupCount++;
                } else {
                    seenSoldOutLineKeys.add(key);
                }
            } else {
                // If we can't determine day/time, remove the line to be safe
                console.log(`    Removing orphaned sold-out line (couldn't determine day/time)`);
                $line.remove();
                lineCleanupCount++;
            }
        });
        
        console.log(`✅ Cleaned up ${badgeCleanupCount} sold-out badges and ${lineCleanupCount} red lines`);
    }
    
    /**
     * Extract day from span position using column layout
     */
    extractDayFromSpanPosition($span) {
        const style = $span.attr('style') || '';
        const leftMatch = style.match(/left:\s*(\d+\.?\d*)px/);
        if (!leftMatch) return null;

        const left = parseFloat(leftMatch[1]);
        const bottom = this.extractBottomFromSpan($span);
        const page = this.getPageNumberForSpan($span);

        return this.resolveDayFromLayoutPosition(left, bottom, page);
    }
    
    /**
     * Extract day from line position using column layout
     */
    extractDayFromLinePosition($line) {
        const style = $line.attr('style') || '';
        const leftMatch = style.match(/left:\s*(\d+\.?\d*)px/);
        if (!leftMatch) return null;

        const left = parseFloat(leftMatch[1]);
        const bottom = this.extractBottomFromSpan($line);
        const page = this.getPageNumberForSpan($line);

        return this.resolveDayFromLayoutPosition(left, bottom, page);
    }

    /**
     * Resolve a day name using the actual day headers on the current HTML page.
     */
    resolveDayFromLayoutPosition(left, bottom, page) {
        if (typeof left !== 'number' || Number.isNaN(left)) return null;

        const dayHeaderPattern = /^(MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)\s*$/i;
        let bestMatch = null;
        let bestDistance = Infinity;

        this.$('span').each((_, elem) => {
            const $header = this.$(elem);
            const text = $header.text().trim();
            if (!dayHeaderPattern.test(text)) return;

            const headerPage = this.getPageNumberForSpan($header);
            if (headerPage !== page) return;

            const headerStyle = $header.attr('style') || '';
            const leftMatch = headerStyle.match(/left:\s*(\d+\.?\d*)px/);
            const bottomMatch = headerStyle.match(/bottom:\s*(\d+\.?\d*)px/);
            if (!leftMatch || !bottomMatch) return;

            const headerLeft = parseFloat(leftMatch[1]);
            const headerBottom = parseFloat(bottomMatch[1]);
            const leftDiff = Math.abs(headerLeft - left);

            if (leftDiff > 110) return;
            if (typeof bottom === 'number' && !Number.isNaN(bottom) && headerBottom < bottom) return;

            const verticalDistance = typeof bottom === 'number' && !Number.isNaN(bottom)
                ? Math.abs(headerBottom - bottom)
                : 0;
            const weightedDistance = verticalDistance + (leftDiff * 0.25);

            if (weightedDistance < bestDistance) {
                bestDistance = weightedDistance;
                bestMatch = text.charAt(0).toUpperCase() + text.slice(1).toLowerCase();
            }
        });

        return bestMatch;
    }
    
    /**
     * Extract time from nearby time spans
     */
    extractTimeFromNearbySpans($span) {
        // Look for time spans near this class span
        const spanBottom = this.extractBottomFromSpan($span);
        if (spanBottom === null) return null;

        const style = $span.attr('style') || '';
        const leftMatch = style.match(/left:\s*(\d+\.?\d*)px/);
        const spanLeft = leftMatch ? parseFloat(leftMatch[1]) : null;
        const spanPage = this.getPageNumberForSpan($span);
        
        // Find time span with similar bottom position (within 5px tolerance)
        let bestMatch = null;
        let bestDistance = Infinity;
        
        this.$('span').each((_, elem) => {
            const $timeSpan = this.$(elem);
            const text = $timeSpan.text().trim();
            
            // Check if this looks like a time
            if (/^\d{1,2}:\d{2}\s*(?:AM|PM)?$/i.test(text)) {
                const timePage = this.getPageNumberForSpan($timeSpan);
                if (timePage !== spanPage) return;

                if (spanLeft !== null) {
                    const timeStyle = $timeSpan.attr('style') || '';
                    const timeLeftMatch = timeStyle.match(/left:\s*(\d+\.?\d*)px/);
                    if (timeLeftMatch) {
                        const timeLeft = parseFloat(timeLeftMatch[1]);
                        if (Math.abs(timeLeft - spanLeft) > 110) return;
                    }
                }

                const timeBottom = this.extractBottomFromSpan($timeSpan);
                if (timeBottom !== null) {
                    const distance = Math.abs(spanBottom - timeBottom);
                    if (distance < 10 && distance < bestDistance) { // Within 10px tolerance
                        bestMatch = text;
                        bestDistance = distance;
                    }
                }
            }
        });
        
        return bestMatch;
    }
    
    /**
     * Extract time from line position by finding nearby time spans
     */
    extractTimeFromLinePosition($line) {
        const lineBottom = this.extractBottomFromSpan($line);
        if (lineBottom === null) return null;

        const style = $line.attr('style') || '';
        const leftMatch = style.match(/left:\s*(\d+\.?\d*)px/);
        const lineLeft = leftMatch ? parseFloat(leftMatch[1]) : null;
        const linePage = this.getPageNumberForSpan($line);
        
        // Find time span with similar bottom position (the line should be slightly above the text)
        let bestMatch = null;
        let bestDistance = Infinity;
        
        this.$('span').each((_, elem) => {
            const $timeSpan = this.$(elem);
            const text = $timeSpan.text().trim();
            
            // Check if this looks like a time
            if (/^\d{1,2}:\d{2}\s*(?:AM|PM)?$/i.test(text)) {
                const timePage = this.getPageNumberForSpan($timeSpan);
                if (timePage !== linePage) return;

                if (lineLeft !== null) {
                    const timeStyle = $timeSpan.attr('style') || '';
                    const timeLeftMatch = timeStyle.match(/left:\s*(\d+\.?\d*)px/);
                    if (timeLeftMatch) {
                        const timeLeft = parseFloat(timeLeftMatch[1]);
                        if (Math.abs(timeLeft - lineLeft) > 50) return;
                    }
                }

                const timeBottom = this.extractBottomFromSpan($timeSpan);
                if (timeBottom !== null) {
                    // Line should be about 8px above the text (based on creation logic)
                    const expectedLineBottom = timeBottom + 8;
                    const distance = Math.abs(lineBottom - expectedLineBottom);
                    if (distance < 15 && distance < bestDistance) { // Allow some tolerance
                        bestMatch = text;
                        bestDistance = distance;
                    }
                }
            }
        });
        
        return bestMatch;
    }
    
    /**
     * Extract bottom position from span style attribute
     */
    extractBottomFromSpan($span) {
        const style = $span.attr('style') || '';
        const bottomMatch = style.match(/bottom:\s*(\d+\.?\d*)px/);
        return bottomMatch ? parseFloat(bottomMatch[1]) : null;
    }

    /**
     * Normalize all content spans to have consistent font-weight and casing
     * This ensures any spans not matched by the main update loop still get proper styling
     */
    normalizeAllContentSpans() {
        console.log('🎨 Normalizing all content spans for consistent styling...');
        let normalizedCount = 0;
        const badgeSelector = this.shouldRenderStaticTheme()
            ? '.sold-out-badge'
            : '.theme-badge, .theme-static-label, .sold-out-badge';
        
        // Find all spans that look like class content (contain trainer names or class names)
        this.$('span').each((_, elem) => {
            const $span = this.$(elem);
            const text = $span.text().trim();
            const style = $span.attr('style') || '';
            
            // Skip time spans, theme badges, and headers
            if (/^\d{1,2}:\d{2}\s*(?:AM|PM)?$/i.test(text)) return;
            if ($span.hasClass('theme-badge') || $span.hasClass('theme-static-label')) return;
            if (text.length > 50) return; // Skip long text (likely headers)
            if (text.length < 5) return; // Skip very short text
            
            // Check if this looks like a class-trainer entry (contains hyphen separator)
            if (text.includes(' - ') && style.includes('font-family')) {
                // Normalize font-weight to 400
                if (style.includes('font-weight: 600')) {
                    $span.css('font-weight', '400');
                    
                    // Also fix the casing of the text content
                    const parts = text.split(' - ');
                    if (parts.length >= 2) {
                        const className = parts[0].trim();
                        const trainerName = parts.slice(1).join(' - ').trim();
                        
                        // Format class name (powerCycle vs UPPERCASE)
                        const formattedClass = this.formatClassName(className);
                        // Trainer name always uppercase
                        const formattedTrainer = trainerName.toUpperCase();
                        
                        // Remove "STUDIO " prefix from class display
                        let formattedClassForDisplay = formattedClass.replace(/^STUDIO\s+/i, '');
                        const newText = `${formattedClassForDisplay} - ${formattedTrainer}`;
                        
                        // Preserve any child elements (like theme badges and sold-out badges)
                        const childBadges = $span.find(badgeSelector).clone();
                        const hasSoldOut = $span.find('.sold-out-badge').length > 0;
                        
                        $span.text(newText);
                        
                        // Re-append the badges
                        if (childBadges.length) {
                            childBadges.each((_, badge) => {
                                $span.append(this.$(badge));
                            });
                        }
                        
                        // Restore sold-out class if badge exists
                        if (hasSoldOut) {
                            $span.addClass('sold-out');
                        }
                    }
                    
                    normalizedCount++;
                }
            }
        });

        if (this.shouldRenderStaticTheme()) {
            this.$('.theme-badge, .theme-static-label').remove();
        }
        
        console.log(`✅ Normalized ${normalizedCount} content spans`);
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
                    const classDisplay = this.formatClassName(this.normalizeClassName(matchingClass.class));
                    const trainerDisplay = this.getTrainerFirstName(matchingClass.trainer).toUpperCase();
                    $entry.attr('data-class', classDisplay);
                    $entry.attr('data-trainer', trainerDisplay);
                    
                    if (matchingClass.notes) {
                        $entry.attr('data-notes', matchingClass.notes);
                    }
                    
                    if (matchingClass.theme && matchingClass.theme.trim()) {
                        $entry.attr('data-theme', matchingClass.theme.trim());
                    }

                    // Update text content with theme badge and sold-out status
                    let newText = `${time} – ${classDisplay} – ${trainerDisplay}`;
                    const inlineThemeSuffix = this.formatInlineStaticThemeSuffix(matchingClass.theme, this.currentLocation);
                    newText += inlineThemeSuffix;
                    
                    // Check if this is a sold-out/hosted class
                    const isSoldOut = matchingClass.notes && matchingClass.notes.includes('SOLD OUT');
                    
                    if (matchingClass.theme && matchingClass.theme.trim() && !this.shouldInlineStaticTheme(this.currentLocation)) {
                        const themeBadge = this.createThemeMarkup(matchingClass.theme.trim(), this.currentLocation);
                        newText += ` ${themeBadge}`;
                    }
                    
                    if (isSoldOut) {
                        // Add sold-out badge
                        newText += ` <span class="sold-out-badge">SOLD OUT</span>`;
                    } else if (matchingClass.notes && !matchingClass.notes.includes('SOLD OUT')) {
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
            
            // Use dynamic or static update based on configuration
            if (DYNAMIC_ROW_MODE) {
                console.log('📊 DYNAMIC_ROW_MODE enabled: Rows will be added/removed based on sheet data');
                this.dynamicUpdatePositionedSpans();
            } else {
                this.updatePositionedSpans();
            }
            this.updateScheduleEntries();
            this.updateDateHeaders();
            this.renderStaticThemeArtifacts();
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
        const updatedHTML = this.formatOutputHtml(this.$.html());
        fs.writeFileSync(this.outputPath, updatedHTML, 'utf-8');
        console.log(`✅ Saved to: ${this.outputPath}`);
    }

    formatOutputHtml(html) {
        return beautifyHtml(String(html || ''), {
            indent_size: 2,
            indent_char: ' ',
            preserve_newlines: true,
            max_preserve_newlines: 2,
            end_with_newline: true,
            wrap_line_length: 0,
            wrap_attributes: 'auto',
            wrap_attributes_indent_size: 2,
            unformatted: [
                'code', 'pre', 'em', 'strong', 'span', 'i', 'b', 'u', 'svg', 'path', 'image',
                'defs', 'clipPath', 'style', 'script', 'textarea'
            ],
            content_unformatted: ['script', 'style'],
            inline: [
                'a', 'abbr', 'acronym', 'b', 'bdo', 'big', 'br', 'button', 'cite', 'code', 'dfn',
                'em', 'i', 'img', 'input', 'kbd', 'label', 'map', 'object', 'output', 'q', 'samp',
                'script', 'select', 'small', 'span', 'strong', 'sub', 'sup', 'textarea', 'time',
                'tt', 'var'
            ]
        });
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
        const now = Date.now();
        if (this.googleAccessTokenCache && now < this.googleAccessTokenExpiry) {
            return this.googleAccessTokenCache;
        }

        if (this.googleAccessTokenPromise) {
            return this.googleAccessTokenPromise;
        }

        console.log('\n🔐 Getting Google OAuth access token...');
        this.googleAccessTokenPromise = (async () => {
            const maxAttempts = 3;
            const retryDelayMs = 1500;

            for (let attempt = 1; attempt <= maxAttempts; attempt++) {
                try {
                    const response = await axios.post(GOOGLE_CONFIG.TOKEN_URL, {
                        client_id: GOOGLE_CONFIG.CLIENT_ID,
                        client_secret: GOOGLE_CONFIG.CLIENT_SECRET,
                        refresh_token: GOOGLE_CONFIG.REFRESH_TOKEN,
                        grant_type: 'refresh_token'
                    }, {
                        timeout: 15000
                    });

                    const accessToken = response.data.access_token;
                    const expiresInSeconds = Number(response.data.expires_in || 3600);
                    this.googleAccessTokenCache = accessToken;
                    this.googleAccessTokenExpiry = Date.now() + Math.max(0, expiresInSeconds - 60) * 1000;
                    console.log('✅ Access token obtained');
                    return accessToken;
                } catch (error) {
                    const isRetryable = this.isRetryableGoogleTokenError(error);
                    const errorDetails = error.response?.data || error.message;

                    if (!isRetryable || attempt === maxAttempts) {
                        console.error('❌ Error getting access token:', errorDetails);
                        throw error;
                    }

                    console.warn(`⚠️  Access token request failed (attempt ${attempt}/${maxAttempts}) — retrying in ${retryDelayMs * attempt}ms...`, errorDetails);
                    await this.sleep(retryDelayMs * attempt);
                }
            }

            throw new Error('Failed to obtain Google OAuth access token after retries');
        })();

        try {
            return await this.googleAccessTokenPromise;
        } finally {
            this.googleAccessTokenPromise = null;
        }
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    isRetryableGoogleTokenError(error) {
        if (!error || error.response) return false;

        const retryableCodes = new Set([
            'ETIMEDOUT',
            'ECONNABORTED',
            'ECONNRESET',
            'ENOTFOUND',
            'EAI_AGAIN'
        ]);

        const codesToCheck = [
            error.code,
            error.cause?.code,
            ...(Array.isArray(error.cause?.errors) ? error.cause.errors.map(innerError => innerError?.code) : [])
        ].filter(Boolean);

        return codesToCheck.some(code => retryableCodes.has(code));
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
            
            // Debug: Check if sold-out elements are in the HTML
            const soldOutBadges = (htmlContent.match(/sold-out-badge/g) || []).length;
            const soldOutLines = (htmlContent.match(/sold-out-line/g) || []).length;
            console.log(`  Found ${soldOutBadges} sold-out badges and ${soldOutLines} sold-out lines in HTML`);
            
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
                    }

                    @page {
                        size: A4 portrait;
                        margin: 0;
                    }

                    :root {
                        --print-page-scale: 0.87245;
                        --print-page-width: 793px;
                        --print-page-height: 1122px;
                    }

                    html, body {
                        margin: 0 !important;
                        padding: 0 !important;
                        min-width: auto !important;
                        background: #fff !important;
                    }

                    .page-container {
                        margin: 0 auto !important;
                        width: var(--print-page-width) !important;
                        height: var(--print-page-height) !important;
                        overflow: hidden !important;
                        break-inside: avoid !important;
                        page-break-inside: avoid !important;
                    }

                    .page-container:has(+ .page-container) {
                        break-after: page !important;
                        page-break-after: always !important;
                    }

                    .page-container:not(:has(+ .page-container)) {
                        break-after: auto !important;
                        page-break-after: auto !important;
                    }

                    .page {
                        margin: 0 !important;
                        zoom: var(--print-page-scale) !important;
                        transform: none !important;
                        break-inside: avoid-page !important;
                        page-break-inside: avoid !important;
                    }
                    
                    /* Ensure sold-out badges and lines render properly */
                    .sold-out-badge {
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                        color-adjust: exact !important;
                        background: #dc2626 !important;
                        color: white !important;
                        display: inline-block !important;
                        font-weight: 800 !important;
                    }
                    
                    .sold-out-line {
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                        color-adjust: exact !important;
                        background-color: #dc2626 !important;
                        display: block !important;
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
            
            // Use dynamic or static update based on configuration
            if (DYNAMIC_ROW_MODE) {
                console.log('📊 DYNAMIC_ROW_MODE enabled: Rows will be added/removed based on sheet data');
                this.dynamicUpdatePositionedSpans();
            } else {
                this.updatePositionedSpans();
            }
            this.updateScheduleEntries();
            this.updateDateHeaders();
            this.renderStaticThemeArtifacts();
            this.save();
            console.log('✅ Schedule HTML updated from Google Sheets Cleaned data!');
            
            // Clean up old sold-out elements now that HTML is saved
            this.cleanupOldSoldOutElements();
            
            // Save again after cleanup to persist the cleanup changes
            this.save();
            console.log('✅ Sold-out element cleanup completed and saved!');
            
            // Add small delay to ensure HTML file is fully written to disk
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            // Step 2: Generate PDF (skip if skipPdf flag is set, e.g., for static Vercel deployments)
            if (this.skipPdf) {
                console.log('⏭️  Skipping PDF generation (skipPdf mode enabled)');
            } else {
                const pdfPath = await this.generatePDF();
                
                // Step 3: Upload to Google Drive
                await this.uploadToGoogleDrive(pdfPath);
            }
            
            if (this.skipPdf) {
                console.log('\n🎉 Complete! Schedule HTML updated from Google Sheets!');
            } else {
                console.log('\n🎉 Complete! Schedule updated from Google Sheets, PDF generated, and uploaded to Google Drive!');
            }
        } catch (error) {
            console.error('❌ Error in updateWithPDF:', error.message);
            throw error;
        }
    }

    /**
     * Complete workflow: Email processing -> Google Sheets -> HTML/PDF (No CSV dependency)
     * @param {boolean} skipEmail - If true, skip email processing and use existing Cleaned sheet data
     */
    async completeGoogleSheetsWorkflow(skipEmail = false) {
        console.log('🚀 Starting complete Google Sheets workflow (no CSV)...');
        
        try {
            if (skipEmail) {
                console.log('⏭️  Skipping email processing - using existing Cleaned sheet data\n');
            } else {
                // STEP 1: Process email and update Google Sheets
                console.log('📧 Step 1: Processing email and updating Google Sheets...');
                await this.processEmailAndUpdateSchedule();
            console.log('✅ Google Sheets updated with latest linked sheet and email cover updates\n');
            }
            
            // STEP 2: Update HTML and PDF directly from Google Sheets
            console.log('📄 Step 2: Updating HTML and generating PDF from Google Sheets...');
            await this.updateWithPDF(); // Now uses readCleanedSheet internally
            
            console.log('🎉 Complete Google Sheets workflow finished successfully!');
            console.log('🔍 Summary of updates:');
            console.log('   - Google Sheets updated with email covers');
            console.log('   - Themes sourced directly from Schedule sheet Theme columns');
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
     * Extract covers from spreadsheet data using known column mappings
     */
    extractSpreadsheetCovers(sheetData) {
        if (!sheetData || sheetData.length < 5) {
            return [];
        }

        const covers = [];
        const daysOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        const structure = this.analyzeSheetStructure(sheetData);
        const timeColIndex = structure.timeColumn;
        
        // Scan rows for covers (starting from row 5, index 4)
        for (let rowIndex = structure.headerRows; rowIndex < sheetData.length; rowIndex++) {
            const row = sheetData[rowIndex];
            if (!row || !row[timeColIndex]) continue; // Skip rows without time
            
            const time = String(row[timeColIndex] || '').trim();
            
            // Check each day block's cover column
            for (const dayName of daysOrder) {
                const dayConfigs = structure.dayColumns[dayName] || [];
                for (const columns of dayConfigs) {
                    if (typeof columns.coverCol !== 'number' || columns.coverCol < 0) continue;

                    const location = String(row[columns.locationCol] || '').trim();
                    const className = String(row[columns.classCol] || '').trim();
                    const trainer1 = String(row[columns.trainer1Col] || '').trim();
                    const cover = String(row[columns.coverCol] || '').trim();
                    
                    if (cover && cover.toLowerCase() !== 'undefined' && location && time) {
                        covers.push({
                            source: 'spreadsheet',
                            day: dayName,
                            time: time,
                            location: location,
                            className: className,
                            originalTrainer: trainer1,
                            coverTrainer: cover
                        });
                    }
                }
            }
        }
        
        return covers;
    }

    /**
    * Log covers found in spreadsheet
     */
    logSpreadsheetCovers(sheetData) {
        if (!sheetData || sheetData.length < 5) {
            console.log('No spreadsheet data to analyze');
            return;
        }

        const daysOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        const structure = this.analyzeSheetStructure(sheetData);
        const timeColIndex = structure.timeColumn;

        console.log('Analyzing spreadsheet for covers...\n');
        
        // Scan rows for covers (starting from row 5, index 4)
        for (let rowIndex = structure.headerRows; rowIndex < sheetData.length; rowIndex++) {
            const row = sheetData[rowIndex];
            if (!row || !row[timeColIndex]) continue; // Skip rows without time
            
            const time = String(row[timeColIndex] || '').trim();
            
            // Check each day block's cover column
            for (const dayName of daysOrder) {
                const dayConfigs = structure.dayColumns[dayName] || [];
                for (const columns of dayConfigs) {
                    if (typeof columns.coverCol !== 'number' || columns.coverCol < 0) continue;

                    const location = String(row[columns.locationCol] || '').trim();
                    const className = String(row[columns.classCol] || '').trim();
                    const trainer1 = String(row[columns.trainer1Col] || '').trim();
                    const cover = String(row[columns.coverCol] || '').trim();
                    
                    // Only log rows with covers
                    if (cover && cover.toLowerCase() !== 'undefined' && location) {
                        console.log(`📍 ${dayName.padEnd(10)} | ${time.padEnd(10)} | ${location.padEnd(10)} | ${className.padEnd(20)} | Trainer: ${trainer1.padEnd(15)} | Cover: ${cover}`);
                    }
                }
            }
        }
    }

    /**
    * Log covers from email body
     */
    logEmailCovers(covers) {
        if (!covers || covers.length === 0) {
            console.log('No covers found in email body');
            return;
        }

        console.log(`Found ${covers.length} cover entries:\n`);
        
        for (const cover of covers) {
            const location = cover.location || 'N/A';
            const day = cover.day || 'N/A';
            const trainer = cover.trainer || 'N/A';
            
            if (cover.timePattern && cover.classType) {
                // Pattern-based cover (e.g., "morning cycles", "evening barre")
                console.log(`📧 ${day.padEnd(10)} | ${cover.timePattern.toUpperCase()} ${cover.classType.padEnd(10)} | ${location.padEnd(10)} | Trainer: ${trainer}`);
            } else if (cover.timesWithClasses && cover.timesWithClasses.length > 0) {
                // Time-based covers with class types
                for (const timeWithClass of cover.timesWithClasses) {
                    const time = timeWithClass.time || timeWithClass;
                    const classType = timeWithClass.classType || 'any';
                    console.log(`📧 ${day.padEnd(10)} | ${String(time).padEnd(10)} | ${location.padEnd(10)} | Class: ${classType.padEnd(15)} | Trainer: ${trainer}`);
                }
            } else if (cover.times && cover.times.length > 0) {
                // Legacy format without class types
                for (const time of cover.times) {
                    console.log(`📧 ${day.padEnd(10)} | ${String(time).padEnd(10)} | ${location.padEnd(10)} | Class: any${' '.padEnd(15)} | Trainer: ${trainer}`);
                }
            } else {
                console.log(`📧 ${day.padEnd(10)} | [unknown format] | ${location.padEnd(10)} | Trainer: ${trainer}`);
            }
        }
    }

    /**
     * Get detailed cover information for email covers by looking up class names from spreadsheet
     */
    getEmailCoverDetails(emailCover) {
        const details = [];
        
        if (!this.spreadsheetCovers || this.spreadsheetCovers.length === 0) {
            // No spreadsheet data available, return basic info
            if (emailCover.timesWithClasses && emailCover.timesWithClasses.length > 0) {
                for (const timeWithClass of emailCover.timesWithClasses) {
                    details.push({
                        day: emailCover.day,
                        time: timeWithClass.time,
                        location: emailCover.location,
                        className: this.mapClassTypeToFullName(timeWithClass.classType || ''),
                        originalTrainer: '',
                        coverTrainer: emailCover.trainer
                    });
                }
            } else if (emailCover.times && emailCover.times.length > 0) {
                for (const time of emailCover.times) {
                    details.push({
                        day: emailCover.day,
                        time: time,
                        location: emailCover.location,
                        className: '',
                        originalTrainer: '',
                        coverTrainer: emailCover.trainer
                    });
                }
            } else if (emailCover.timePattern) {
                details.push({
                    day: emailCover.day,
                    time: emailCover.timePattern + (emailCover.classType ? ' ' + emailCover.classType : ''),
                    location: emailCover.location,
                    className: emailCover.classType ? this.mapClassTypeToFullName(emailCover.classType) : '',
                    originalTrainer: '',
                    coverTrainer: emailCover.trainer
                });
            }
            return details;
        }
        
        // Look up class names from spreadsheet covers
        if (emailCover.timesWithClasses && emailCover.timesWithClasses.length > 0) {
            for (const timeWithClass of emailCover.timesWithClasses) {
                const matchingCovers = this.findMatchingSpreadsheetCovers(
                    emailCover.day,
                    timeWithClass.time,
                    emailCover.location,
                    timeWithClass.classType
                );
                
                if (matchingCovers.length > 0) {
                    // Found matching classes in spreadsheet
                    for (const match of matchingCovers) {
                        details.push({
                            day: emailCover.day,
                            time: match.time,
                            location: match.location,
                            className: match.className,
                            originalTrainer: match.originalTrainer,
                            coverTrainer: emailCover.trainer
                        });
                    }
                } else {
                    // No match found, use basic info
                    details.push({
                        day: emailCover.day,
                        time: timeWithClass.time,
                        location: emailCover.location,
                        className: this.mapClassTypeToFullName(timeWithClass.classType || ''),
                        originalTrainer: '',
                        coverTrainer: emailCover.trainer
                    });
                }
            }
        } else if (emailCover.times && emailCover.times.length > 0) {
            for (const time of emailCover.times) {
                const matchingCovers = this.findMatchingSpreadsheetCovers(
                    emailCover.day,
                    time,
                    emailCover.location
                );
                
                if (matchingCovers.length > 0) {
                    for (const match of matchingCovers) {
                        details.push({
                            day: emailCover.day,
                            time: match.time,
                            location: match.location,
                            className: match.className,
                            originalTrainer: match.originalTrainer,
                            coverTrainer: emailCover.trainer
                        });
                    }
                } else {
                    details.push({
                        day: emailCover.day,
                        time: time,
                        location: emailCover.location,
                        className: '',
                        originalTrainer: '',
                        coverTrainer: emailCover.trainer
                    });
                }
            }
        } else if (emailCover.timePattern) {
            // Pattern-based cover (morning/evening)
            const matchingCovers = this.findMatchingSpreadsheetCoversByPattern(
                emailCover.day,
                emailCover.timePattern,
                emailCover.location,
                emailCover.classType
            );
            
            if (matchingCovers.length > 0) {
                for (const match of matchingCovers) {
                    details.push({
                        day: emailCover.day,
                        time: match.time,
                        location: match.location,
                        className: match.className,
                        originalTrainer: match.originalTrainer,
                        coverTrainer: emailCover.trainer
                    });
                }
            } else {
                details.push({
                    day: emailCover.day,
                    time: emailCover.timePattern + (emailCover.classType ? ' ' + emailCover.classType : ''),
                    location: emailCover.location,
                    className: emailCover.classType ? this.mapClassTypeToFullName(emailCover.classType) : '',
                    originalTrainer: '',
                    coverTrainer: emailCover.trainer
                });
            }
        }
        
        return details.length > 0 ? details : [{
            day: emailCover.day,
            time: 'All',
            location: emailCover.location,
            className: '',
            originalTrainer: '',
            coverTrainer: emailCover.trainer
        }];
    }
    
    /**
     * Find matching spreadsheet covers by day, time, location, and optional class type
     */
    findMatchingSpreadsheetCovers(day, time, location, classType = null) {
        const matches = [];
        const normalizedTime = this.normalizeTime(time);
        
        // First try: match with location
        for (const cover of this.spreadsheetCovers) {
            if (cover.day !== day) continue;
            if (!this.matchLocation(cover.location, location)) continue;
            if (!this.timeMatches(cover.time, normalizedTime)) continue;
            
            // If class type is specified, check if it matches
            if (classType) {
                if (!this.classTypeMatches(cover.className, classType)) continue;
            }
            
            matches.push(cover);
        }
        
        // If no matches found with location, try ANY location for that day/time
        // This handles cases where email groups covers under one location but they span multiple locations
        if (matches.length === 0) {
            for (const cover of this.spreadsheetCovers) {
                if (cover.day !== day) continue;
                if (!this.timeMatches(cover.time, normalizedTime)) continue;
                
                // If class type is specified, check if it matches
                if (classType) {
                    if (!this.classTypeMatches(cover.className, classType)) continue;
                }
                
                matches.push(cover);
            }
        }
        
        return matches;
    }
    
    /**
     * Find matching spreadsheet covers by pattern (morning/evening) and class type
     * Also searches in ALL spreadsheet data (not just covers) to find all matching classes
     */
    findMatchingSpreadsheetCoversByPattern(day, pattern, location, classType) {
        const matches = [];
        const isMorning = pattern.toLowerCase().includes('morning');
        const isEvening = pattern.toLowerCase().includes('evening');
        
        // First, check existing spreadsheet covers
        for (const cover of this.spreadsheetCovers) {
            if (cover.day !== day) continue;
            
            // Check time pattern
            const time = cover.time.toLowerCase();
            const isMorningClass = time.includes('am') && !time.match(/^(12):/);
            const isEveningClass = time.includes('pm') && !time.match(/^(12):/);
            
            if (isMorning && !isMorningClass) continue;
            if (isEvening && !isEveningClass) continue;
            
            // Check class type if specified
            if (classType && !this.classTypeMatches(cover.className, classType)) continue;
            
            // For location: try exact match first, then any location
            const locationMatches = this.matchLocation(cover.location, location);
            if (locationMatches || !location) {
                matches.push(cover);
            }
        }
        
        // If no matches in covers, search the raw spreadsheet data for all matching classes
        if (matches.length === 0 && this.rawSpreadsheetData) {
            matches.push(...this.findClassesInSpreadsheet(day, pattern, classType, location));
        }
        
        return matches;
    }
    
    /**
     * Find all classes in spreadsheet that match pattern and class type
     */
    findClassesInSpreadsheet(day, pattern, classType, location = null) {
        const matches = [];
        if (!this.rawSpreadsheetData || this.rawSpreadsheetData.length < 5) return matches;
        
        const isMorning = pattern.toLowerCase().includes('morning');
        const isEvening = pattern.toLowerCase().includes('evening');
        
        const structure = this.analyzeSheetStructure(this.rawSpreadsheetData);
        const timeColIndex = structure.timeColumn;
        const dayKey = Object.keys(structure.dayColumns).find(d => d.toLowerCase() === String(day).toLowerCase()) || day;
        const dayConfigs = structure.dayColumns[dayKey] || [];

        if (dayConfigs.length === 0) return matches;

        // Scan all rows for matching classes in this day block
        for (let rowIndex = structure.headerRows; rowIndex < this.rawSpreadsheetData.length; rowIndex++) {
            const row = this.rawSpreadsheetData[rowIndex];
            if (!row || !row[timeColIndex]) continue;

            const time = String(row[timeColIndex] || '').trim();
            if (!time) continue;

            for (const cfg of dayConfigs) {
                const cellLocation = String(row[cfg.locationCol] || '').trim();
                const className = String(row[cfg.classCol] || '').trim();
                const trainer = String(row[cfg.trainer1Col] || '').trim();
                
                if (!cellLocation || !className) continue;
                
                // Check time pattern
                const timeLower = time.toLowerCase();
                const isMorningClass = timeLower.includes('am') && !timeLower.match(/^(12):/);
                const isEveningClass = timeLower.includes('pm') && !timeLower.match(/^(12):/);
                
                if (isMorning && !isMorningClass) continue;
                if (isEvening && !isEveningClass) continue;
                
                // Check class type
                if (classType && !this.classTypeMatches(className, classType)) continue;
                
                // Check location if specified
                if (location && !this.matchLocation(cellLocation, location)) continue;
                
                matches.push({
                    source: 'spreadsheet',
                    day: dayKey,
                    time: time,
                    location: cellLocation,
                    className: className,
                    originalTrainer: trainer,
                    coverTrainer: '' // Will be filled in by caller
                });
            }
        }
        
        return matches;
    }
    
    /**
     * Map short class type codes to full names
     */
    mapClassTypeToFullName(classType) {
        const lowerType = classType.toLowerCase();
        const mapping = {
            'lab': 'Strength Lab',
            'barre57': 'Barre57',
            'b57': 'Barre57',
            'barre': 'Barre57',
            'mat57': 'Mat57',
            'm57': 'Mat57',
            'mat': 'Mat57',
            'cycle': 'PowerCycle',
            'bbb': 'Back Body Blaze',
            'fit': 'FIT',
            'cardio': 'Cardio B'
        };
        
        return mapping[lowerType] || classType;
    }
    
    /**
     * Check if class name matches the class type
     */
    classTypeMatches(className, classType) {
        const lowerClassName = className.toLowerCase();
        const lowerClassType = classType.toLowerCase();
        
        // Direct match
        if (lowerClassName.includes(lowerClassType)) return true;
        
        // Check aliases
        if (lowerClassType === 'lab' && lowerClassName.includes('strength')) return true;
        if (lowerClassType === 'barre57' || lowerClassType === 'b57' || lowerClassType === 'barre') {
            if (lowerClassName.includes('barre') && lowerClassName.includes('57')) return true;
            if (lowerClassName === 'barre57') return true;
        }
        if (lowerClassType === 'mat57' || lowerClassType === 'm57' || lowerClassType === 'mat') {
            if (lowerClassName.includes('mat') && lowerClassName.includes('57')) return true;
            if (lowerClassName === 'mat57') return true;
        }
        if (lowerClassType === 'cycle' && lowerClassName.includes('cycle')) return true;
        if (lowerClassType === 'bbb' && lowerClassName.includes('back body blaze')) return true;
        if (lowerClassType === 'fit' && lowerClassName === 'fit') return true;
        if (lowerClassType === 'cardio' && lowerClassName.includes('cardio')) return true;
        
        return false;
    }

    /**
     * Ensure a sheet exists in the spreadsheet, create if it doesn't
     */
    async ensureSheetExists(sheets, sheetName) {
        try {
            // Get spreadsheet metadata to check if sheet exists
            const spreadsheet = await sheets.spreadsheets.get({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID
            });
            
            const sheetExists = spreadsheet.data.sheets.some(
                sheet => sheet.properties.title === sheetName
            );
            
            if (!sheetExists) {
                console.log(`📄 Creating new sheet: ${sheetName}`);
                await sheets.spreadsheets.batchUpdate({
                    spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                    resource: {
                        requests: [{
                            addSheet: {
                                properties: {
                                    title: sheetName
                                }
                            }
                        }]
                    }
                });
                console.log(`✅ Created sheet: ${sheetName}`);
            } else {
                console.log(`✓ Sheet "${sheetName}" already exists`);
            }
        } catch (error) {
            console.error(`❌ Error ensuring sheet exists: ${error.message}`);
            throw error;
        }
    }

    /**
     * Populate the Covers sheet with all covers from spreadsheet and email
     */
    async populateCoversSheet(sheets, emailCovers) {
        console.log('📋 Populating Covers sheet...');
        
        try {
            // First, ensure the Covers sheet exists
            await this.ensureSheetExists(sheets, 'Covers');
            
            // Clear existing data from the Covers sheet
            console.log('🧹 Clearing existing data from Covers sheet...');
            try {
                await sheets.spreadsheets.values.clear({
                    spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                    range: 'Covers!A:Z'
                });
                console.log('✅ Cleared existing data');
            } catch (clearError) {
                console.log('⚠️  Could not clear existing data (sheet might be empty):', clearError.message);
            }
            
            // Prepare headers
            const headers = ['Source', 'Day', 'Time', 'Location', 'Class', 'Original Trainer', 'Cover Trainer'];
            const rows = [headers];
            
            // Add spreadsheet covers
            if (this.spreadsheetCovers && this.spreadsheetCovers.length > 0) {
                console.log(`📊 Adding ${this.spreadsheetCovers.length} covers from spreadsheet`);
                for (const cover of this.spreadsheetCovers) {
                    rows.push([
                        cover.source,
                        cover.day,
                        cover.time,
                        cover.location,
                        cover.className,
                        cover.originalTrainer,
                        cover.coverTrainer
                    ]);
                }
            }
            
            // Add email covers with class names looked up from spreadsheet
            if (emailCovers && emailCovers.length > 0) {
                console.log(`📧 Adding ${emailCovers.length} covers from email`);
                for (const cover of emailCovers) {
                    // Look up class names from spreadsheet for each time
                    const coverDetails = this.getEmailCoverDetails(cover);
                    
                    // Add a row for each time/class combination
                    for (const detail of coverDetails) {
                        rows.push([
                            'email',
                            detail.day,
                            detail.time,
                            detail.location,
                            detail.className,
                            detail.originalTrainer,
                            detail.coverTrainer
                        ]);
                    }
                }
            }
            
            console.log(`📝 Writing ${rows.length - 1} total cover entries to Covers sheet`);
            
            // Write to Covers sheet
            await sheets.spreadsheets.values.update({
                spreadsheetId: GOOGLE_CONFIG.TARGET_SPREADSHEET_ID,
                range: 'Covers!A1',
                valueInputOption: 'RAW',
                resource: {
                    values: rows
                }
            });
            
            console.log('✅ Successfully populated Covers sheet');
            
        } catch (error) {
            console.error('❌ Error populating Covers sheet:', error.message);
            // Don't throw - this is not critical to the main workflow
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
                    @page { size: A4 portrait; margin: 0; }
                    :root {
                        --print-page-scale: 0.87245;
                        --print-page-width: 793px;
                        --print-page-height: 1122px;
                    }
                    html, body {
                        margin: 0 !important;
                        padding: 0 !important;
                        min-width: auto !important;
                        background: #fff !important;
                    }
                    .page-container {
                        margin: 0 auto !important;
                        width: var(--print-page-width) !important;
                        height: var(--print-page-height) !important;
                        overflow: hidden !important;
                        break-inside: avoid !important;
                        page-break-inside: avoid !important;
                    }
                    .page-container:has(+ .page-container) {
                        break-after: page !important;
                        page-break-after: always !important;
                    }
                    .page-container:not(:has(+ .page-container)) {
                        break-after: auto !important;
                        page-break-after: auto !important;
                    }
                    .page {
                        margin: 0 !important;
                        zoom: var(--print-page-scale) !important;
                        transform: none !important;
                        break-inside: avoid-page !important;
                        page-break-inside: avoid !important;
                    }
                    
                    /* Ensure theme-related elements are completely hidden (except badges) */
                    .theme-index,
                    #theme-index,
                    .legend,
                    .theme-legend,
                    .index-legend {
                        display: none !important;
                    }
                    
                    ${this.shouldRenderStaticTheme() ? `
                    /* Static mode uses strip overlays and legend bars, not inline badges */
                    .theme-badge,
                    .theme-static-label {
                        display: none !important;
                        visibility: hidden !important;
                    }
                    .theme-row-highlight,
                    .theme-index-band {
                        display: block !important;
                        visibility: visible !important;
                    }
                    ` : `
                    /* Keep theme markup visible in exported PDFs */
                    .theme-badge,
                    .theme-static-label {
                        display: inline-block !important;
                        visibility: visible !important;
                    }
                    `}
                    
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
            const bandraPdfPath = path.join(__dirname, 'Schedule-Bandra.pdf');
            const combinedPdfPath = path.join(__dirname, 'Schedule-Mumbai.pdf');
            
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
            await this.uploadNamedPDF(combinedPdfPath, 'Schedule-Mumbai.pdf');
            
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
            const bandraPdfPath = path.join(__dirname, 'Schedule-Bandra.pdf');
            console.log('   ✓ Bandra schedule updated');
            
            // Step 3: Generate combined PDF
            console.log('📑 Step 3: Generating combined PDF...');
            const combinedPdfPath = await this.generateCombinedPDF();
            console.log('   ✓ Combined PDF generated');
            
            // Step 4: Upload all files to Google Drive atomically
            console.log('☁️  Step 4: Uploading all files to Google Drive...');
            await Promise.all([
                this.uploadNamedPDF(kempsPdfPath, 'Schedule-Kemps.pdf'),
                this.uploadNamedPDF(bandraPdfPath, 'Schedule-Bandra.pdf'),
                this.uploadNamedPDF(combinedPdfPath, 'Schedule-Mumbai.pdf')
            ]);
            
            console.log('🎉 Atomic file update completed successfully!');
            console.log('📊 Updated files:');
            console.log('   - Kemps.html & Schedule-Kemps.pdf');
            console.log('   - Bandra.html & Schedule-Bandra.pdf');
            console.log('   - Schedule-Mumbai.pdf');
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
        await this.readSheet();
        // Filter Bandra classes from allSheetRecords (not kwalityClasses which is Kemps only)
        const bandraClasses = (this.allSheetRecords || []).filter(r => r.Location && /Supreme HQ.*Bandra|Supreme HQ,\s*Bandra/i.test(r.Location));
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
        
        if (DYNAMIC_ROW_MODE) {
            console.log('📊 DYNAMIC_ROW_MODE enabled for Bandra: rebuilding rows from sheet data');
            this.dynamicUpdatePositionedSpans();
        } else {
            this.updatePositionedSpans();
        }
        this.updateScheduleEntries();
        this.updateDateHeaders();
        this.renderStaticThemeArtifacts();
        this.save();
        
        // Skip PDF generation if skipPdf flag is set (e.g., for static Vercel deployments)
        if (this.skipPdf) {
            console.log('⏭️  Skipping PDF generation for Bandra (skipPdf mode enabled)');
        } else {
            await this.generatePDFNamed('Schedule-Bandra.pdf');
            // Upload Bandra.pdf to Drive
            const bandraPdfPath = path.join(__dirname, 'Schedule-Bandra.pdf');
            await this.uploadNamedPDF(bandraPdfPath, 'Schedule-Bandra.pdf');
        }
        
        // Restore original context for safety
        this.htmlPath = originalHtmlPath;
        this.outputPath = originalOutputPath;
        
        if (this.skipPdf) {
            console.log('🎉 Bandra schedule update complete (HTML only)');
        } else {
            console.log('🎉 Bandra schedule update complete (HTML + Bandra.pdf)');
        }
    }
}

// Main execution
// Robust main-check that works for CommonJS and ESM (even when invoked with a relative path)
let isMain = false;
if (typeof require !== 'undefined' && require.main === module) {
    isMain = true;
} else if (typeof import.meta !== 'undefined') {
    try {
        const entryPath = process.argv[1] ? path.resolve(process.argv[1]) : null;
        const scriptPath = fileURLToPath(import.meta.url);
        if (entryPath && path.resolve(scriptPath) === entryPath) isMain = true;
    } catch (e) {
        // fallback: leave isMain false
    }
}

// Provide CommonJS-like __filename and __dirname when running as ESM
let __filename = typeof fileURLToPath === 'function' ? fileURLToPath(import.meta.url) : undefined;
let __dirname = __filename ? path.dirname(__filename) : undefined;
if (isMain) {
    // Check for command-line flags
    const args = process.argv.slice(2);
    const skipEmail = args.includes('--skip-email') || args.includes('--html-only');
    const staticThemeMode = args.includes('--static');
    const badgeThemeMode = args.includes('--badge');
    const skipPdf = args.includes('--skip-pdf');
    const serveTabs = args.includes('--serve-tabs');
    const resolvedStaticThemeMode = badgeThemeMode ? false : (staticThemeMode || serveTabs);
    
    // Show help if requested
    if (args.includes('--help') || args.includes('-h')) {
        console.log('\n📋 Usage: node updateKempsSchedule.js [options]\n');
        console.log('Options:');
        console.log('  --skip-email, --html-only   Skip email processing, use existing Cleaned sheet data');
        console.log('  --static                    Render themes as plain static text instead of badges');
        console.log('  --badge                     Force badge rendering even in local preview mode');
        console.log('  --serve-tabs                Serve generated HTML locally and open each file in its own browser tab');
        console.log('  --help, -h                  Show this help message\n');
        console.log('Examples:');
        console.log('  node updateKempsSchedule.js                    # Full workflow (email + sheets + HTML)');
        console.log('  node updateKempsSchedule.js --skip-email       # Skip email, update HTML from existing data');
        console.log('  node updateKempsSchedule.js --static           # Full workflow with static theme text');
        console.log('  node updateKempsSchedule.js --serve-tabs       # Generate local previews that match the Vercel static build');
        console.log('  node updateKempsSchedule.js --serve-tabs --badge # Generate local preview tabs with theme badges');
        console.log('  npm run update                                 # Full workflow');
        console.log('  npm run update -- --static                     # Full workflow with static theme text');
        console.log('  npm run update -- --serve-tabs                 # Full workflow plus local browser preview tabs');
        console.log('  npm run preview                                # Local preview tabs matching Vercel output');
        console.log('  npm run preview:badge                          # Local preview tabs with badge rendering');
        console.log('  npm run update -- --skip-email                 # Skip email via npm\n');
        process.exit(0);
    }
    
    // Read from backup (clean template) and write to Kemps.html
    const backupPath = path.join(__dirname, 'Kemps.backup.html');
    const htmlPath = fs.existsSync(backupPath) ? backupPath : path.join(__dirname, 'Kemps.html');
    const outputPath = path.join(__dirname, 'Kemps.html');
    
    console.log(`📄 Source file: ${path.basename(htmlPath)}`);
    console.log(`📄 Output file: ${path.basename(outputPath)}`);
    
    if (skipEmail) {
        console.log('⏭️  Mode: Skip email processing (using existing Cleaned sheet data)\n');
    }

    console.log(`🎨 Theme render mode: ${resolvedStaticThemeMode ? 'static text' : 'badge'}`);
    if (serveTabs) {
        console.log('🌐 Preview mode: serve generated HTML and open each file in a separate browser tab');
        if (!staticThemeMode && !badgeThemeMode) {
            console.log('🪞 Local preview parity: --serve-tabs defaults to Vercel-style static theme rendering');
        }
    }

    const updater = new ScheduleUpdater(htmlPath, outputPath, 'kemps', {
        themeRenderMode: resolvedStaticThemeMode ? 'static' : 'badge',
        skipPdf: skipPdf
    }); // No CSV needed
    
    (async () => {
        try {
            console.log('🚀 Starting Google Sheets-only schedule update workflow...');
            console.log('📊 No CSV files will be used - all data from Google Sheets');
            
            // Use the complete Google Sheets workflow
            await updater.completeGoogleSheetsWorkflow(skipEmail);
            
            // Optional: Also update Bandra if needed
            console.log('\n📄 Updating Bandra schedule...');
            await updater.updateBandra();
            
            console.log('\n✨ All tasks completed successfully!');
            if (skipEmail) {
                console.log('📄 Mode: HTML-only update (email processing skipped)');
            }
            console.log('📄 Generated files:');
            console.log('   - Kemps.html (updated from Google Sheets)');
            console.log('   - Kemps_Updated.pdf (uploaded to Drive)');
            console.log('   - Bandra.pdf (if applicable)');

            if (serveTabs) {
                await serveOutputFilesInTabs([
                    {
                        label: 'Kemps',
                        routeName: 'Kemps.html',
                        filePath: path.join(__dirname, 'Kemps.html')
                    },
                    {
                        label: 'Bandra',
                        routeName: 'Bandra.html',
                        filePath: path.join(__dirname, 'Bandra.html')
                    }
                ]);
            }
            
        } catch (error) {
            console.error('Failed to update schedule:', error);
            process.exit(1);
        }
    })();
}

export default ScheduleUpdater;
