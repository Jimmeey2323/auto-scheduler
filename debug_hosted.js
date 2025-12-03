const ScheduleUpdater = require('./updateKempsSchedule.js');

async function debugHosted() {
  const updater = new ScheduleUpdater('./Kemps.html', './Kemps.html', 'kemps');
  
  const emailData = await updater.findLatestScheduleEmail();
  const fullContent = emailData.allMessages.join('\n\n').replace(/\r\n/g, '\n');
  
  console.log('\n=== LOOKING FOR HOSTED CLASSES SECTION ===\n');
  
  const hostedMatch = fullContent.match(/-Hosted Classes\s*-(.*?)(?=\n\n|\n[A-Z]|Covers|$)/is);
  
  if (hostedMatch) {
    console.log('Found hosted section!');
    console.log('Content:', hostedMatch[1].substring(0, 500));
    
    const lines = hostedMatch[1].split('\n').filter(line => line.trim());
    console.log('\nLines found:', lines.length);
    lines.forEach((line, i) => {
      console.log(`Line ${i}: "${line}"`);
      if (line.trim() && !line.trim().startsWith('-')) {
        const result = updater.parseHostedLine(line.trim());
        console.log(`  Parsed:`, result);
      }
    });
  } else {
    console.log('No hosted section found');
    console.log('\nSearching for "Hosted" in email...');
    const idx = fullContent.indexOf('Hosted');
    if (idx > -1) {
      console.log('Found at index:', idx);
      console.log('Context:', fullContent.substring(Math.max(0, idx - 100), idx + 300));
    }
  }
}

debugHosted().catch(console.error);
