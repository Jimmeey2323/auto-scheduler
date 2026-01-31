import path from 'path';
import ScheduleUpdater from './updateKempsSchedule.js';

/**
 * Complete workflow to update schedule from email to final HTML/PDF
 */
async function fullScheduleUpdate() {
    console.log('🚀 Starting complete schedule update workflow...');
    
    try {
        const htmlPath = path.join(__dirname, 'Kemps.html');
        const outputPath = path.join(__dirname, 'Kemps.html');
        
        // Create updater without CSV dependency
        const updater = new ScheduleUpdater(htmlPath, outputPath);
        
        // STEP 1: Process email and update Google Sheets
        console.log('📧 Step 1: Processing email and updating Google Sheets...');
        await updater.processEmailAndUpdateSchedule();
        console.log('✅ Google Sheets updated with email data\\n');
        
        // STEP 2: Fix CSV dates from updated Google Sheets
        console.log('📅 Step 2: Updating CSV with correct dates from Google Sheets...');
        const datePopulator = new DatePopulatorFromSheets(csvPath);
        await datePopulator.processCSV();
        console.log('✅ CSV updated with correct dates\\n');
        
        // STEP 3: Update all HTML files and generate PDFs atomically
        console.log('📄 Step 3: Updating all HTML files and generating PDFs atomically...');
        await updater.updateAllFilesAtomically();
        console.log('✅ All HTML and PDF files updated atomically\n');
        
        console.log('🎉 Complete schedule update workflow finished successfully!');
        console.log('🔍 Summary of updates:');
        console.log('   - Google Sheets updated with email covers and themes');
        console.log('   - All HTML and PDF files updated atomically');
        console.log('   - All files uploaded to Google Drive');
        
    } catch (error) {
        console.error('❌ Workflow failed:', error.message);
        console.error('🔍 Error details:', error);
        process.exit(1);
    }
}

// Run the complete workflow
if (require.main === module) {
    fullScheduleUpdate();
}

export default fullScheduleUpdate;