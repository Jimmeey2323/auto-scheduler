import ScheduleUpdater from './updateKempsSchedule.js';

async function debugLatestEmail() {
    const updater = new ScheduleUpdater('./Kemps.html', './Kemps.html', 'kemps');
    
    console.log('🔍 Getting latest email data...\n');
    
    const emailData = await updater.findLatestScheduleEmail();
    if (!emailData) {
        console.log('❌ No email found');
        return;
    }
    
    console.log('=== EMAIL THREAD ANALYSIS ===');
    console.log(`📧 Subject: ${emailData.subject}`);
    console.log(`📅 Date: ${emailData.date}`);
    console.log(`📨 Thread has ${emailData.allMessages.length} messages\n`);
    
    console.log('=== LATEST MESSAGE (emailData.body) ===');
    console.log(emailData.body);
    console.log('\n' + '='.repeat(80) + '\n');
    
    console.log('=== ALL MESSAGES IN THREAD ===');
    emailData.allMessages.forEach((message, index) => {
        console.log(`--- Message ${index + 1} ---`);
        console.log(message.substring(0, 500));
        console.log('...\n');
    });
    
    // Check for Google Sheets links
    console.log('=== GOOGLE SHEETS LINK SEARCH ===');
    console.log('🔍 Checking latest message:');
    const linkInLatest = updater.extractSheetsLink(emailData.body);
    
    console.log('\n🔍 Checking all messages:');
    emailData.allMessages.forEach((message, index) => {
        console.log(`\nMessage ${index + 1}:`);
        const link = updater.extractSheetsLink(message);
        if (link) console.log(`  ✅ FOUND LINK: ${link}`);
    });
}

debugLatestEmail().catch(console.error);