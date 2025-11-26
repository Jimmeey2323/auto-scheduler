const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

// OAuth2 configuration with Gmail scope
const OAUTH2_CONFIG = {
  CLIENT_ID: process.env.GOOGLE_CLIENT_ID,
  CLIENT_SECRET: process.env.GOOGLE_CLIENT_SECRET,
  REDIRECT_URI: "http://localhost:8080/oauth2callback",
  SCOPES: [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
  ]
};

async function setupGmailAuth() {
  console.log('🔐 Setting up Gmail authentication...\n');
  
  const oauth2Client = new google.auth.OAuth2(
    OAUTH2_CONFIG.CLIENT_ID,
    OAUTH2_CONFIG.CLIENT_SECRET,
    OAUTH2_CONFIG.REDIRECT_URI
  );

  // Generate authorization URL
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: OAUTH2_CONFIG.SCOPES,
    prompt: 'consent' // Force consent screen to get refresh token
  });

  console.log('📋 Required scopes:');
  OAUTH2_CONFIG.SCOPES.forEach(scope => {
    console.log(`   - ${scope}`);
  });
  console.log('\n🌐 Opening browser for authentication...');
  console.log('Authorization URL:', authUrl);
  
  // Open browser
  try {
    const open = (await import('open')).default;
    await open(authUrl);
  } catch (error) {
    console.log('❌ Could not open browser automatically.');
    console.log('📋 Please manually open the following URL in your browser:');
    console.log(authUrl);
  }
  
  console.log('\n📝 Instructions:');
  console.log('1. Complete the authentication in your browser');
  console.log('2. You will be redirected to localhost:8080/oauth2callback');
  console.log('3. Copy the authorization code from the URL');
  console.log('4. The URL will look like: http://localhost:8080/oauth2callback?code=YOUR_CODE_HERE');
  console.log('5. Copy just the code part and paste it below');
  
  // Wait for user input
  const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
  });
  
  return new Promise((resolve) => {
    readline.question('\n🔑 Enter the authorization code: ', async (code) => {
      readline.close();
      
      try {
        console.log('\n🔄 Exchanging code for tokens...');
        const { tokens } = await oauth2Client.getToken(code);
        
        console.log('✅ Tokens obtained successfully!');
        console.log('\n📋 Your new OAuth configuration:');
        console.log('CLIENT_ID:', OAUTH2_CONFIG.CLIENT_ID);
        console.log('CLIENT_SECRET:', OAUTH2_CONFIG.CLIENT_SECRET);
        console.log('REFRESH_TOKEN:', tokens.refresh_token);
        
        // Update the main script with new refresh token
        const configUpdate = `
// Updated Google OAuth Configuration with Gmail scope
const GOOGLE_CONFIG = {
  CLIENT_ID: "${OAUTH2_CONFIG.CLIENT_ID}",
  CLIENT_SECRET: "${OAUTH2_CONFIG.CLIENT_SECRET}",
  REFRESH_TOKEN: "${tokens.refresh_token}",
  TOKEN_URL: "https://oauth2.googleapis.com/token",
  FOLDER_ID: "1PPaEKOBcPtjSUpFZZArkRLEBcGO5h108",
  SPREADSHEET_ID: "1OhhnD-9R_876ehw1xROZpyd0VcCLTwuuUUnh3A1Alv4",
  SHEET_NAME: "Cleaned",
  TARGET_SPREADSHEET_ID: "1OhhnD-9R_876ehw1xROZpyd0VcCLTwuuUUnh3A1Alv4",
  TARGET_SHEET_NAME: "Schedule"
};`;
        
        console.log('\n💾 Configuration to update in your main script:');
        console.log(configUpdate);
        
        // Save to a file for easy copying
        fs.writeFileSync('gmail_auth_config.txt', configUpdate);
        console.log('\n📄 Configuration saved to gmail_auth_config.txt');
        console.log('📋 Copy the REFRESH_TOKEN value and update your main script');
        
        resolve();
      } catch (error) {
        console.error('❌ Error exchanging code for tokens:', error);
        resolve();
      }
    });
  });
}

// Run the setup
if (require.main === module) {
  setupGmailAuth().then(() => {
    console.log('\n✨ Gmail authentication setup complete!');
    process.exit(0);
  }).catch((error) => {
    console.error('❌ Setup failed:', error);
    process.exit(1);
  });
}

module.exports = { setupGmailAuth };