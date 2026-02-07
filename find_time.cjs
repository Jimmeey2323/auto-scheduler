const fs = require('fs');
const cheerio = require('cheerio');
const html = fs.readFileSync('Kemps.html','utf8');
const $ = cheerio.load(html);
$('span.s9').each((i,el)=>{
  const txt = $(el).text().trim();
  if(txt === '05:15 PM'){
    console.log('Found time span:', i, $(el).attr('style'));
    const next = $(el).nextAll('span').filter((j,sp)=> $(sp).hasClass('v0') || $(sp).hasClass('v0 s5') || $(sp).hasClass('v0 s5'));
    if(next.length>0){
      console.log('Next span HTML:', $(next[0]).prop('outerHTML'));
    } else {
      console.log('No next span');
    }
  }
});
