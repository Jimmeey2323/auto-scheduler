import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';

const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
const TOKEN_PATH = path.join(__dirname, 'token.json');

async function getAuth() {
    const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
}

async function main() {
    const auth = await getAuth();
    const gmail = google.gmail({ version: 'v1', auth });
    
    // Search for latest schedule email
    const messages = await gmail.users.messages.list({
        userId: 'me',
        q: 'from:hi@thestudio.co.in subject:"Next week schedule"',
        maxResults: 1
    });
    
    if (!messages.data.messages || messages.data.messages.length === 0) {
        console.log('No messages found');
        return;
    }
    
    const msgId = messages.data.messages[0].id;
    const msg = await gmail.users.messages.get({
        userId: 'me',
        id: msgId,
        format: 'full'
    });
    
    // Decode and print the email body
    function getBody(payload) {
        let body = '';
        if (payload.body && payload.body.data) {
            body = Buffer.from(payload.body.data, 'base64').toString('utf-8');
        }
        if (payload.parts) {
            for (const part of payload.parts) {
                if (part.mimeType === 'text/plain' && part.body && part.body.data) {
                    body += Buffer.from(part.body.data, 'base64').toString('utf-8');
                }
                if (part.parts) {
                    body += getBody(part);
                }
            }
        }
        return body;
    }
    
    const body = getBody(msg.data.payload);
    console.log('=== EMAIL BODY ===');
    console.log(body);
    console.log('=== END EMAIL BODY ===');
    
    // Extract and show the Covers section specifically
    const coversMatch = body.match(/Covers\s*:?\s*([\s\S]*?)(?=\n\s*(?:Amped Up theme|Bandra cycle themes|FIT theme|Best,\s*$))/i);
    if (coversMatch) {
        console.log('\n=== COVERS SECTION ===');
        console.log(coversMatch[1]);
        console.log('=== END COVERS SECTION ===');
    } else {
        console.log('\n❌ No covers section found with main regex');
        // Try alternative
        const altMatch = body.match(/Covers\s*:?\s*([\s\S]*?)(?=\n\s*(?:Amped|Bandra cycle|FIT theme|Best))/i);
        if (altMatch) {
            console.log('\n=== COVERS SECTION (ALT) ===');
            console.log(altMatch[1]);
            console.log('=== END COVERS SECTION ===');
        }
    }
}

main().catch(console.error);
