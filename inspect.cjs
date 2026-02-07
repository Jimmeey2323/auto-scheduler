const fs=require('fs');
const cheerio=require('cheerio');
const html=fs.readFileSync('Kemps.html','utf8');
const $=cheerio.load(html);
const badges=$('span.sold-out-badge');
console.log('badges:', badges.length);
badges.each((i,el)=>{
  const p=$(el).closest('span');
  console.log(i, 'parent text:', p.text().replace(/\s+/g,' '));
  console.log('outerHTML:', $(el).prop('outerHTML'));
});
const lines=$('span.sold-out-line');
console.log('lines:', lines.length);
lines.each((i,el)=>console.log(i, $(el).attr('style')));
