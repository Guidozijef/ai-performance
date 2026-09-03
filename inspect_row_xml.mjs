import JSZip from 'jszip';
import fs from 'fs';

// Inspect the exact row XML format for task rows 10-13 and row 14 (fixed item)
async function inspectRowXml() {
  const buf = fs.readFileSync('./public/正式员工绩效表单套表-JX1.7.xlsx');
  const zip = await JSZip.loadAsync(buf);
  
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('text');
  const ssXml = await zip.file('xl/sharedStrings.xml').async('text');
  
  // Parse all shared strings
  const siMatches = ssXml.match(/<si>.*?<\/si>/gs) || [];
  
  // Show all rows 10-14 fully
  const rowMatches = [...sheetXml.matchAll(/<row r="(1[0-4])"[^>]*>.*?<\/row>/gs)];
  rowMatches.forEach(m => {
    console.log(`\n=== Row ${m[1]} ===`);
    console.log(m[0]);
  });
  
  // Show shared strings relevant to rows 10-13
  console.log('\n=== Shared Strings 26-45 ===');
  for (let i = 26; i <= 50 && i < siMatches.length; i++) {
    const tMatches = siMatches[i].match(/<t[^>]*>([^<]*)<\/t>/g) || [];
    const text = tMatches.map(t => t.replace(/<[^>]+>/g, '')).join('');
    console.log(`[${i}] "${text.substring(0, 100)}"`);
  }
}

inspectRowXml().catch(console.error);
