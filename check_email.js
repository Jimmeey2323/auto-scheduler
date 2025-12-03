const ScheduleUpdater = require('./updateKempsSchedule.js');

async function checkEmail() {
  const updater = new ScheduleUpdater('./Kemps.html', './Kemps.html', 'kemps');
  
  console.log('Finding latest email...');
  const emailData = await updater.findLatestScheduleEmail();
  
  if (!emailData) {
    console.log('No email found');
    return;
  }
  
  console.log('\n=== EMAIL CONTENT ===');
  console.log('Subject:', emailData.subject);
  console.log('\nBody (first 2000 chars):');
  console.log(emailData.body.substring(0, 2000));
  
  console.log('\n\n=== PARSING HOSTED CLASSES ===');
  const emailInfo = updater.parseEmailForScheduleInfo(emailData.allMessages);
  console.log('Hosted classes found:', emailInfo.hostedClasses.length);
  emailInfo.hostedClasses.forEach(h => {
    console.log(`  ${h.day} - ${h.location} - ${h.classType} - ${h.time} - ${h.trainer}`);
  });
}

checkEmail().catch(console.error);
