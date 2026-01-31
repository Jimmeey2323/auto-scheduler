import path from 'path';
import ScheduleUpdater from './updateKempsSchedule.js';

/**
 * Simple script to update HTML and PDF from existing Google Sheets data
 * Assumes Google Sheets already has the latest themes/covers
 */
async function updateHTMLFromSheets() {
    console.log('📄 Updating HTML files and generating PDFs from Google Sheets Cleaned tab...\n');
    
    try {
        // Update Kemps
        console.log('🏢 Processing Kemps schedule...');
        const kempsUpdater = new ScheduleUpdater(path.join(__dirname, 'Kemps.html'), path.join(__dirname, 'Kemps.html'), 'kemps');
        await kempsUpdater.updateWithPDF();
        console.log('✅ Kemps HTML and PDF updated\n');
        
        // Update Bandra  
        console.log('🏢 Processing Bandra schedule...');
        const bandraUpdater = new ScheduleUpdater(path.join(__dirname, 'Bandra.html'), path.join(__dirname, 'Bandra.html'), 'bandra');
        await bandraUpdater.updateWithPDF();
        console.log('✅ Bandra HTML and PDF updated\n');
        
        console.log('🎉 All HTML files and PDFs have been updated and uploaded to Google Drive!');
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        console.error(error.stack);
        process.exit(1);
    }
}

// Run if executed directly
if (require.main === module) {
    updateHTMLFromSheets();
}

export default { updateHTMLFromSheets };
