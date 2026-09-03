import JSZip from 'jszip';
import fs from 'fs';

async function inspectRawXml() {
  const buf = fs.readFileSync('./public/正式员工绩效表单套表-JX1.7.xlsx');
  const zip = await JSZip.loadAsync(buf);

  // 查看 xlsx 内部结构
  console.log('=== Files in xlsx ZIP ===');
  Object.keys(zip.files).forEach(name => {
    if (!zip.files[name].dir) {
      console.log(name);
    }
  });

  // 找到 sheet1
  const sheetFile = zip.file('xl/worksheets/sheet1.xml');
  if (sheetFile) {
    const xml = await sheetFile.async('text');
    // 仅显示前 5000 字符
    console.log('\n=== sheet1.xml (first 5000 chars) ===');
    console.log(xml.substring(0, 5000));
  }
}

inspectRawXml().catch(console.error);
