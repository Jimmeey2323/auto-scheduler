import ScheduleUpdater from './updateKempsSchedule.js';
import path from 'path';

(async () => {
  const updater = new ScheduleUpdater(path.join(__dirname, 'Kemps.html'), path.join(__dirname, 'Kemps.html'));
  await updater.readCleanedSheet();
  
  console.log('=== ALL RECORDS BY LOCATION ===');
  const byLocation = {};
  (updater.allSheetRecords || []).forEach(r => {
    const loc = r.Location || 'Unknown';
    if (!byLocation[loc]) byLocation[loc] = 0;
    byLocation[loc]++;
  });
  console.log(byLocation);
  
  console.log('');
  console.log('=== SAMPLE KEMPS WEDNESDAY ENTRIES ===');
  const kempsWed = (updater.allSheetRecords || []).filter(r => 
    r.Location && r.Location.toLowerCase().includes('kwality') && 
    r.Day === 'Wednesday'
  );
  console.log('Count:', kempsWed.length);
  kempsWed.forEach(r => {
    console.log(`  ${r.Time} ${r.Class} - ${r.Trainer}`);
  });
  
  console.log('');
  console.log('=== SAMPLE BANDRA WEDNESDAY ENTRIES ===');
  const bandraWed = (updater.allSheetRecords || []).filter(r => 
    r.Location && r.Location.toLowerCase().includes('bandra') && 
    r.Day === 'Wednesday'
  );
  console.log('Count:', bandraWed.length);
  bandraWed.forEach(r => {
    console.log(`  ${r.Time} ${r.Class} - ${r.Trainer}`);
  });
  
  console.log('');
  console.log('=== CHECK FOR DUPLICATES ===');
  const seen = {};
  const duplicates = [];
  (updater.allSheetRecords || []).forEach(r => {
    const key = `${r.Day}|${r.Time}|${r.Location}|${r.Class}`;
    if (seen[key]) {
      duplicates.push(key);
    }
    seen[key] = true;
  });
  console.log('Duplicate entries:', duplicates.length);
  if (duplicates.length > 0) {
    duplicates.slice(0, 10).forEach(d => console.log('  ' + d));
  }
})();
