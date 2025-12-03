const ScheduleUpdater = require('./updateKempsSchedule.js');
const path = require('path');

(async () => {
  const updater = new ScheduleUpdater(path.join(__dirname, 'Kemps.html'), path.join(__dirname, 'Kemps.html'));
  const emailData = await updater.findLatestScheduleEmail();
  if (!emailData) { 
    console.log('No email found'); 
    process.exit(1); 
  }
  
  console.log('=== EMAIL THEMES ===');
  const info = updater.parseEmailForScheduleInfo(emailData.allMessages);
  console.log('Bandra themes from email:');
  info.themes.filter(t => t.location === 'Bandra').forEach(t => {
    console.log('  ' + t.day + ' ' + (t.time || '') + ': ' + t.theme);
  });
  
  console.log('');
  console.log('=== CLEANED SHEET THEMES ===');
  await updater.readSheet();
  const bandraWithThemes = (updater.allSheetRecords || []).filter(r => 
    r.Location && r.Location.includes('Bandra') && r.Theme && r.Theme.trim()
  );
  console.log('Bandra classes with themes in sheet (' + bandraWithThemes.length + '):');
  bandraWithThemes.forEach(r => {
    console.log('  ' + r.Day + ' ' + r.Time + ' ' + r.Class + ': ' + r.Theme);
  });
  
  console.log('');
  console.log('=== SATURDAY & SUNDAY BANDRA CYCLE CLASSES ===');
  const weekendBandra = (updater.allSheetRecords || []).filter(r => 
    r.Location && r.Location.toLowerCase().includes('bandra') && 
    (r.Day === 'Saturday' || r.Day === 'Sunday') &&
    r.Class && r.Class.toLowerCase().includes('cycle')
  );
  console.log('Weekend Bandra cycle classes (' + weekendBandra.length + '):');
  weekendBandra.forEach(r => {
    console.log('  ' + r.Day + ' ' + r.Time + ' ' + r.Class + ' Theme: ' + (r.Theme || 'NONE'));
  });
  
  console.log('');
  console.log('=== WEDNESDAY BANDRA CLASSES ===');
  const wednesdayBandra = (updater.allSheetRecords || []).filter(r => 
    r.Location && r.Location.includes('Bandra') && r.Day === 'Wednesday'
  );
  console.log('Wednesday Bandra classes (' + wednesdayBandra.length + '):');
  wednesdayBandra.forEach(r => {
    console.log('  ' + r.Time + ' ' + r.Class + ' - ' + r.Trainer);
  });
})();
