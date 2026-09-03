import JSZip from 'jszip';
import fs from 'fs';

async function getSharedStrings() {
  const buf = fs.readFileSync('./public/正式员工绩效表单套表-JX1.7.xlsx');
  const zip = await JSZip.loadAsync(buf);

  const ssFile = zip.file('xl/sharedStrings.xml');
  const ssXml = await ssFile.async('text');
  
  // Extract all <si> entries
  const siMatches = ssXml.match(/<si>.*?<\/si>/gs) || [];
  console.log(`Total shared strings: ${siMatches.length}`);
  siMatches.slice(0, 30).forEach((si, i) => {
    const tMatches = si.match(/<t[^>]*>([^<]*)<\/t>/g) || [];
    const text = tMatches.map(t => t.replace(/<[^>]+>/g, '')).join('');
    console.log(`[${i}] "${text.substring(0, 80)}"`);
  });
  
  console.log('\n=== Row 10-13 XML ===');
  const sheetFile = zip.file('xl/worksheets/sheet1.xml');
  const xml = await sheetFile.async('text');
  
  // Extract rows 10-13
  const row10Match = xml.match(/<row r="10"[^>]*>.*?<\/row>/s);
  const row11Match = xml.match(/<row r="11"[^>]*>.*?<\/row>/s);
  const row12Match = xml.match(/<row r="12"[^>]*>.*?<\/row>/s);
  const row13Match = xml.match(/<row r="13"[^>]*>.*?<\/row>/s);
  const row14Match = xml.match(/<row r="14"[^>]*>.*?<\/row>/s);
  
  console.log('Row 10:', row10Match ? row10Match[0].substring(0, 300) : 'NOT FOUND');
  console.log('Row 11:', row11Match ? row11Match[0].substring(0, 300) : 'NOT FOUND');
  console.log('Row 12:', row12Match ? row12Match[0].substring(0, 300) : 'NOT FOUND');
  console.log('Row 13:', row13Match ? row13Match[0].substring(0, 300) : 'NOT FOUND');
  console.log('Row 14:', row14Match ? row14Match[0].substring(0, 300) : 'NOT FOUND');
}

getSharedStrings().catch(console.error);
