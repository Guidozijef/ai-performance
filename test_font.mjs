import fs from 'fs';
import JSZip from 'jszip';

async function check() {
  const buf = fs.readFileSync('src/1、正式员工绩效表单套表-JX1.7xlsx.xlsx');
  const zip = await JSZip.loadAsync(buf);
  const stylesXml = await zip.file('xl/styles.xml').async('text');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('text');

  const fonts = [...stylesXml.matchAll(/<font>([\s\S]*?)<\/font>/g)].map(m => m[1]);
  const xfs = [...stylesXml.matchAll(/<xf ([^>]+)>(?:<alignment ([^>]+)\/>)?<\/xf>/g)];

  const ssMatches = [...ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)];
  ssMatches.forEach((m, idx) => {
    if (m[1].includes('<rPr>') || m[1].includes('FFFF0000') || m[1].includes('color')) {
      console.log(`si [${idx}] has rich text/color:`, m[1]);
    }
  });

  // Check all cells in row 10 and row 3, 4, 5 again
  for (let r = 1; r <= 13; r++) {
    const rowMatch = sheetXml.match(new RegExp(`<row r="${r}"[\\s\\S]*?<\\/row>`));
    if (rowMatch) {
      const redCells = [...rowMatch[0].matchAll(/<c r="([A-Z0-9]+)"(?: s="(\d+)")?[^>]*>/g)]
        .filter(c => {
          const sId = c[2] ? parseInt(c[2]) : null;
          if (sId === null) return false;
          const xf = xfs[sId];
          return xf && (xf[1].includes('fontId="18"') || xf[1].includes('fontId="20"') || xf[1].includes('fontId="7"'));
        });
      if (redCells.length > 0) {
        console.log(`Row ${r} has red cells:`, redCells.map(c => c[1]).join(', '));
      }
    }
  }


  console.log('\n--- Row 11 cells font & alignment ---');
  ['A11','B11','C11','D11','E11','F11','G11','H11','I11','J11','K11','L11'].forEach(ref => {
    const m = sheetXml.match(new RegExp('<c r="' + ref + '"(?: s="(\\d+)")?'));
    const sId = m ? parseInt(m[1]) : null;
    const xf = sId !== null ? xfs[sId] : null;
    const fontIdMatch = xf ? xf[1].match(/fontId="(\d+)"/) : null;
    const fontId = fontIdMatch ? parseInt(fontIdMatch[1]) : null;
    const fontXml = fontId !== null ? fonts[fontId] : null;
    console.log(ref, 'styleId:', sId, 'align:', xf ? xf[2] : 'none', 'fontXml:', fontXml ? fontXml.replace(/\s+/g, ' ') : 'none');
  });
}
check().catch(console.error);
