// Debug script to test email search
import ScheduleUpdater from './updateKempsSchedule.js';

async function debugEmailSearch() {
    const updater = new ScheduleUpdater('Kemps.html', 'Kemps.html', 'kemps');
    
    console.log('=== Testing Email Search ===\n');
    
    try {
        // Get access token
        const accessToken = await updater.getAccessToken();
        console.log('✅ Got access token\n');
        
        import { google } from 'googleapis';
        const oauth2Client = new google.auth.OAuth2();
        oauth2Client.setCredentials({ access_token: accessToken });
        const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
        
        // Test 1: Search from vivaran only
        console.log('--- Test 1: Emails from vivaran@physique57mumbai.com ---');
        let response = await gmail.users.messages.list({
            userId: 'me',
            q: 'from:vivaran@physique57mumbai.com newer_than:30d',
            maxResults: 5
        });
        console.log(`Found: ${response.data.messages?.length || 0} emails\n`);
        
        if (response.data.messages?.length > 0) {
            for (const msg of response.data.messages.slice(0, 3)) {
                const detail = await gmail.users.messages.get({
                    userId: 'me',
                    id: msg.id,
                    format: 'metadata',
                    metadataHeaders: ['Subject', 'From', 'Date']
                });
                const headers = detail.data.payload.headers;
                const subject = headers.find(h => h.name === 'Subject')?.value;
                const from = headers.find(h => h.name === 'From')?.value;
                const date = headers.find(h => h.name === 'Date')?.value;
                console.log(`  From: ${from}`);
                console.log(`  Subject: ${subject}`);
                console.log(`  Date: ${date}\n`);
            }
        }
        
        // Test 2: Search with "Schedule" in subject from vivaran
        console.log('--- Test 2: Emails from vivaran with "Schedule" in subject ---');
        response = await gmail.users.messages.list({
            userId: 'me',
            q: 'from:vivaran@physique57mumbai.com subject:Schedule newer_than:30d',
            maxResults: 5
        });
        console.log(`Found: ${response.data.messages?.length || 0} emails\n`);
        
        // Test 3: Search both senders with OR
        console.log('--- Test 3: Both senders with OR ---');
        response = await gmail.users.messages.list({
            userId: 'me',
            q: '(from:mrigakshi@physique57mumbai.com OR from:vivaran@physique57mumbai.com) subject:Schedule newer_than:30d',
            maxResults: 5
        });
        console.log(`Found: ${response.data.messages?.length || 0} emails\n`);
        
        if (response.data.messages?.length > 0) {
            for (const msg of response.data.messages.slice(0, 3)) {
                const detail = await gmail.users.messages.get({
                    userId: 'me',
                    id: msg.id,
                    format: 'metadata',
                    metadataHeaders: ['Subject', 'From', 'Date']
                });
                const headers = detail.data.payload.headers;
                const subject = headers.find(h => h.name === 'Subject')?.value;
                const from = headers.find(h => h.name === 'From')?.value;
                const date = headers.find(h => h.name === 'Date')?.value;
                console.log(`  From: ${from}`);
                console.log(`  Subject: ${subject}`);
                console.log(`  Date: ${date}\n`);
            }
        }
        
        // Test 4: Use curly braces for OR (Gmail alternative syntax)
        console.log('--- Test 4: Using {from:a from:b} syntax ---');
        response = await gmail.users.messages.list({
            userId: 'me',
            q: '{from:mrigakshi@physique57mumbai.com from:vivaran@physique57mumbai.com} subject:Schedule newer_than:30d',
            maxResults: 5
        });
        console.log(`Found: ${response.data.messages?.length || 0} emails\n`);
        
        if (response.data.messages?.length > 0) {
            for (const msg of response.data.messages.slice(0, 3)) {
                const detail = await gmail.users.messages.get({
                    userId: 'me',
                    id: msg.id,
                    format: 'metadata',
                    metadataHeaders: ['Subject', 'From', 'Date']
                });
                const headers = detail.data.payload.headers;
                const subject = headers.find(h => h.name === 'Subject')?.value;
                const from = headers.find(h => h.name === 'From')?.value;
                const date = headers.find(h => h.name === 'Date')?.value;
                console.log(`  From: ${from}`);
                console.log(`  Subject: ${subject}`);
                console.log(`  Date: ${date}\n`);
            }
        }
        
    } catch (error) {
        console.error('Error:', error.message);
    }
}

debugEmailSearch();
