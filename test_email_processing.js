import ScheduleUpdater from './updateKempsSchedule.js';

async function testEmailProcessing() {
    const updater = new ScheduleUpdater('./Kemps.html', './Kemps.html', 'kemps');
    
    console.log('🔍 Testing complete email processing workflow...\n');
    
    try {
        await updater.processEmailAndUpdateSchedule();
    } catch (error) {
        console.error('❌ Error:', error);
    }
}

testEmailProcessing();