import ScheduleUpdater from './updateKempsSchedule.js';

const updater = new ScheduleUpdater('./Kemps.html', './Kemps.html', 'kemps');

const testLine = "Kemps - Saturday - 11.30 am - B57 - SOLD OUT - for Raman Lamba - Pranjali";
console.log('Testing line:', testLine);
const result = updater.parseHostedLine(testLine);
console.log('Result:', JSON.stringify(result, null, 2));
