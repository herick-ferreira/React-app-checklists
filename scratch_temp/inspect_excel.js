const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const filePath = path.join(__dirname, '..', 'public', 'Exemplo.xlsx');
console.log('Lendo arquivo:', filePath);

if (!fs.existsSync(filePath)) {
  console.error('Arquivo não encontrado!');
  process.exit(1);
}

const fileBuffer = fs.readFileSync(filePath);
const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });

const parseDate = (Data) => {
  let dateObj = null;
  if (Data instanceof Date) {
    dateObj = Data;
  } else if (typeof Data === 'number') {
    dateObj = new Date((Data - 25569) * 86400 * 1000);
  } else if (typeof Data === 'string') {
    const cleaned = Data.trim();
    const match = cleaned.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
    if (match) {
      const day = parseInt(match[1], 10);
      const month = parseInt(match[2], 10) - 1;
      const year = parseInt(match[3], 10);
      dateObj = new Date(year, month, day);
    } else {
      const d = new Date(cleaned);
      if (!isNaN(d.getTime())) dateObj = d;
    }
  }
  return dateObj;
};

const anosEncontrados = new Set();
const mesesEncontrados = new Set();
let invalidCount = 0;

raw.forEach(r => {
  const parsed = parseDate(r.Data);
  if (parsed && !isNaN(parsed.getTime())) {
    anosEncontrados.add(parsed.getFullYear());
    mesesEncontrados.add(parsed.getMonth() + 1);
  } else {
    invalidCount++;
  }
});

console.log('Anos encontrados após a correção:', Array.from(anosEncontrados).sort());
console.log('Meses encontrados após a correção:', Array.from(mesesEncontrados).sort((a,b) => a-b));
console.log('Registros inválidos:', invalidCount);
