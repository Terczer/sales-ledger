const http = require('node:http');
const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');

const ROOT = __dirname;
const TEMPLATE_FILENAME = 'ewidencja sprzedaży bezrachunkowej template.xlsx';
const TEMPLATE_PATH = path.join(ROOT, 'template', TEMPLATE_FILENAME);
const DATA_DIR = path.join(ROOT, 'dane');
const EXPORTS_DIR = path.join(ROOT, 'exports');
const DATA_FILENAME_PATTERN = /^transakcje .+\.xlsx$/i;
const PORT = Number(process.env.PORT) || 3000;

const SHEET_NAMES = {
  ledger: 'Ewidencja Sprzedaży',
  tpay: 'Tpay Transakcje',
  ebok: 'Ebok PlanetPay Transakcje',
};

const LEDGER_HEADERS = {
  number: ['Lp'],
  date: ['Data transakcji'],
  vat: ['Stawka VAT'],
  gross: ['Kwota brutto'],
  notes: ['Uwagi'],
  transactionId: ['Id transakcji', 'ID transakcji'],
};

const SOURCES = [
  {
    sheetName: SHEET_NAMES.tpay,
    operator: 'TPAY',
    headers: {
      date: ['Data płatności'],
      amount: ['Kwota'],
      transactionId: ['Tytuł'],
    },
  },
  {
    sheetName: SHEET_NAMES.ebok,
    operator: 'eBok ITCARD',
    headers: {
      date: ['Data i godzina transakcji'],
      amount: ['Kwota całkowita transakcji'],
      transactionId: ['Nr transkacji (RRN)', 'Nr transakcji (RRN)'],
    },
  },
];

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLocaleLowerCase('pl-PL');
}

function getCellValue(cell) {
  const value = cell.value;
  if (value == null) return '';
  if (value instanceof Date) return value;
  if (typeof value === 'object') {
    if ('result' in value) return value.result ?? '';
    if ('text' in value) return value.text ?? '';
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
  }
  return value;
}

function getCellText(cell) {
  const value = getCellValue(cell);
  if (value instanceof Date) return value.toISOString();
  return String(value).trim();
}

function findWorksheet(workbook, expectedName) {
  const expected = normalizeText(expectedName);
  const worksheet = workbook.worksheets.find((sheet) => normalizeText(sheet.name) === expected);
  if (!worksheet) {
    throw new Error(`Brak wymaganego arkusza „${expectedName}”.`);
  }
  return worksheet;
}

function findHeaderRow(worksheet, requiredHeaders) {
  const required = Object.entries(requiredHeaders).map(([key, aliases]) => ({
    key,
    aliases: aliases.map(normalizeText),
  }));
  const lastRow = Math.min(worksheet.rowCount, 30);

  for (let rowNumber = 1; rowNumber <= lastRow; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const columns = new Map();
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      const value = normalizeText(getCellValue(cell));
      if (value) columns.set(value, columnNumber);
    });

    const result = {};
    for (const header of required) {
      const match = header.aliases.find((alias) => columns.has(alias));
      if (match) result[header.key] = columns.get(match);
    }

    if (Object.keys(result).length === required.length) {
      return { rowNumber, columns: result };
    }
  }

  throw new Error(`Nie znaleziono wymaganych nagłówków w arkuszu „${worksheet.name}”.`);
}

function dateParts(value, date1904 = false) {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return {
      year: value.getUTCFullYear(),
      month: value.getUTCMonth() + 1,
      day: value.getUTCDate(),
    };
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    const date = new Date(epoch + Math.floor(value) * 86400000);
    return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
  }

  const text = String(value ?? '').trim();
  let match = text.match(/^(\d{1,2})[-./](\d{1,2})[-./](\d{4})(?:\s.*)?$/);
  if (match) {
    return validateDateParts(Number(match[3]), Number(match[2]), Number(match[1]));
  }

  match = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:[T\s].*)?$/);
  if (match) {
    return validateDateParts(Number(match[1]), Number(match[2]), Number(match[3]));
  }
  return null;
}

function validateDateParts(year, month, day) {
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() + 1 !== month ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return { year, month, day };
}

function parseAmount(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.round(value * 100) / 100;
  }

  let text = String(value ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/[^\d,.-]/g, '')
    .replace(/\s/g, '');

  if (!text || !/^-?[\d.,]+$/.test(text)) return null;

  if (text.includes(',')) {
    text = text.replace(/\./g, '').replace(',', '.');
  }
  const amount = Number(text);
  return Number.isFinite(amount) ? Math.round(amount * 100) / 100 : null;
}

function getTransactionId(cell) {
  const value = getCellValue(cell);
  if (value === '') return '';
  if (typeof value === 'number' && Number.isFinite(value)) {
    return cell.text || String(value);
  }
  return String(value).trim();
}

function readTransactions(workbook) {
  const transactions = [];

  for (const source of SOURCES) {
    const worksheet = findWorksheet(workbook, source.sheetName);
    const header = findHeaderRow(worksheet, source.headers);

    for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      const dateValue = getCellValue(row.getCell(header.columns.date));
      const amountValue = getCellValue(row.getCell(header.columns.amount));
      const idCell = row.getCell(header.columns.transactionId);
      const transactionId = getTransactionId(idCell);

      if (dateValue === '' && amountValue === '' && transactionId === '') continue;

      const date = dateParts(dateValue, Boolean(workbook.properties.date1904));
      const amount = parseAmount(amountValue);
      if (!date) {
        throw new Error(`Nieprawidłowa data w arkuszu „${source.sheetName}”, wiersz ${rowNumber}.`);
      }
      if (amount === null) {
        throw new Error(`Nieprawidłowa kwota w arkuszu „${source.sheetName}”, wiersz ${rowNumber}.`);
      }
      if (!transactionId) {
        throw new Error(`Brak identyfikatora transakcji w arkuszu „${source.sheetName}”, wiersz ${rowNumber}.`);
      }

      transactions.push({
        ...date,
        amount,
        transactionId,
        notes: `Usługi logopedyczne – rozliczenie operatora płatności (${source.operator})`,
      });
    }
  }

  if (transactions.length === 0) {
    throw new Error('Nie znaleziono transakcji do ujęcia w ewidencji.');
  }
  return transactions;
}

function formatMonth(monthKey) {
  const [year, month] = monthKey.split('-');
  return `${month}/${year}`;
}

function getMonthKey(transaction) {
  return `${transaction.year}-${String(transaction.month).padStart(2, '0')}`;
}

function getOutputBaseName(originalName) {
  const basename = path.basename(String(originalName).replace(/\\/g, '/'));
  const withoutExtension = basename.replace(/\.xlsx$/i, '');
  const cleaned = withoutExtension
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '');
  return cleaned || 'Ewidencja sprzedaży';
}

function copyRowStyle(sourceRow, targetRow, columnCount) {
  for (let column = 1; column <= columnCount; column += 1) {
    const sourceCell = sourceRow.getCell(column);
    targetRow.getCell(column).style = { ...sourceCell.style };
  }
  if (sourceRow.height) targetRow.height = sourceRow.height;
}

function createLedger(workbook, transactions, monthKey) {
  const worksheet = findWorksheet(workbook, SHEET_NAMES.ledger);
  const { rowNumber: headerRowNumber, columns } = findHeaderRow(worksheet, LEDGER_HEADERS);
  const firstDataRow = headerRowNumber + 1;
  const styleRow = worksheet.getRow(firstDataRow);

  for (let rowNumber = firstDataRow; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (const column of Object.values(columns)) {
      row.getCell(column).value = null;
    }
  }

  const titleCell = worksheet.getCell(headerRowNumber - 1, 1);
  const existingTitle = getCellText(titleCell);
  const titleBase = existingTitle.replace(/\s+\d{1,2}[/-]\d{4}\s*$/, '').trim()
    || 'Ewidencja Sprzedaży Bezrachunkowej';
  titleCell.value = `${titleBase} ${formatMonth(monthKey)}`;

  let total = 0;
  transactions.forEach((transaction, index) => {
    const rowNumber = firstDataRow + index;
    const row = worksheet.getRow(rowNumber);
    if (rowNumber !== firstDataRow) copyRowStyle(styleRow, row, worksheet.columnCount);

    const date = new Date(Date.UTC(transaction.year, transaction.month - 1, transaction.day));
    total += transaction.amount;
    row.getCell(columns.number).value = index + 1;
    row.getCell(columns.date).value = date;
    row.getCell(columns.vat).value = 'zw';
    row.getCell(columns.gross).value = transaction.amount;
    row.getCell(columns.notes).value = transaction.notes;
    row.getCell(columns.transactionId).value = transaction.transactionId;
  });
  total = Math.round(total * 100) / 100;

  const totalRowNumber = firstDataRow + transactions.length;
  const totalRow = worksheet.getRow(totalRowNumber);
  copyRowStyle(styleRow, totalRow, worksheet.columnCount);
  totalRow.getCell(columns.number).value = 'SUMA';
  const grossCell = totalRow.getCell(columns.gross);
  const firstAmountRow = firstDataRow;
  const lastAmountRow = totalRowNumber - 1;
  const grossColumnLetter = worksheet.getColumn(columns.gross).letter;
  grossCell.value = {
    formula: `SUM(${grossColumnLetter}${firstAmountRow}:${grossColumnLetter}${lastAmountRow})`,
    result: total,
  };

  for (const column of Object.values(columns)) {
    const cell = totalRow.getCell(column);
    cell.font = { ...cell.font, bold: true };
    cell.border = { ...cell.border, top: { style: 'thin' } };
  }

  return { total, worksheet };
}

async function listTransactionFiles(dataDir = DATA_DIR) {
  let entries;
  try {
    entries = await fs.readdir(dataDir, { withFileTypes: true });
  } catch (error) {
    if (error.code === 'ENOENT') return [];
    throw error;
  }

  const files = await Promise.all(entries
    .filter((entry) => entry.isFile() && DATA_FILENAME_PATTERN.test(entry.name))
    .map(async (entry) => {
      const filePath = path.join(dataDir, entry.name);
      const stat = await fs.stat(filePath);
      return { filename: entry.name, modifiedAt: stat.mtime.toISOString() };
    }));

  return files.sort((a, b) => b.modifiedAt.localeCompare(a.modifiedAt)
    || b.filename.localeCompare(a.filename, 'pl'));
}

function readJson(request) {
  return new Promise((resolve, reject) => {
    let body = '';
    let tooLarge = false;
    request.setEncoding('utf8');
    request.on('data', (chunk) => {
      if (tooLarge) return;
      body += chunk;
      if (body.length > 4096) {
        tooLarge = true;
        body = '';
      }
    });
    request.on('error', reject);
    request.on('end', () => {
      if (tooLarge) {
        reject(new Error('Nieprawidłowe żądanie.'));
        return;
      }
      try {
        resolve(JSON.parse(body));
      } catch {
        reject(new Error('Nieprawidłowe żądanie.'));
      }
    });
  });
}

async function generateWorkbooks(sourceFilename, {
  dataDir = DATA_DIR,
  templatePath = TEMPLATE_PATH,
  exportsDir = EXPORTS_DIR,
} = {}) {
  if (typeof sourceFilename !== 'string' || !DATA_FILENAME_PATTERN.test(sourceFilename)) {
    throw new Error('Wybierz plik transakcji z folderu dane.');
  }

  const availableFiles = await listTransactionFiles(dataDir);
  if (!availableFiles.some((file) => file.filename === sourceFilename)) {
    throw new Error('Wybrany plik nie znajduje się już w folderze dane. Odśwież listę i spróbuj ponownie.');
  }

  const sourcePath = path.join(dataDir, sourceFilename);
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(sourcePath);
  const transactions = readTransactions(sourceWorkbook);
  const groups = new Map();

  for (const transaction of transactions) {
    const key = getMonthKey(transaction);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(transaction);
  }

  try {
    await fs.access(templatePath);
  } catch (error) {
    if (error.code === 'ENOENT') {
      throw new Error(`Nie znaleziono szablonu „${TEMPLATE_FILENAME}” w folderze template.`);
    }
    throw error;
  }

  await fs.mkdir(exportsDir, { recursive: true });
  const baseName = getOutputBaseName(TEMPLATE_FILENAME).replace(/\s+template$/i, '');
  const results = [];

  for (const [monthKey, monthTransactions] of [...groups.entries()].sort(([a], [b]) => a.localeCompare(b))) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const { total } = createLedger(workbook, monthTransactions, monthKey);
    const [year, month] = monthKey.split('-');
    const filename = `${baseName} - ${month}-${year}.xlsx`;
    const outputPath = path.join(exportsDir, filename);
    await workbook.xlsx.writeFile(outputPath);
    results.push({
      filename,
      month: formatMonth(monthKey),
      count: monthTransactions.length,
      total,
      url: `/exports/${encodeURIComponent(filename)}`,
    });
  }

  return results;
}

async function sendStatic(response, filename, contentType) {
  const content = await fs.readFile(path.join(ROOT, 'public', filename));
  response.writeHead(200, {
    'Content-Type': contentType,
    'Content-Length': content.length,
    'Cache-Control': 'no-cache',
  });
  response.end(content);
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  response.end(JSON.stringify(payload));
}

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url, 'http://localhost');

  if (request.method === 'GET' && url.pathname === '/') {
    await sendStatic(response, 'index.html', 'text/html; charset=utf-8');
    return;
  }
  if (request.method === 'GET' && url.pathname === '/app.js') {
    await sendStatic(response, 'app.js', 'text/javascript; charset=utf-8');
    return;
  }
  if (request.method === 'GET' && url.pathname === '/styles.css') {
    await sendStatic(response, 'styles.css', 'text/css; charset=utf-8');
    return;
  }
  if (request.method === 'GET' && url.pathname === '/api/files') {
    try {
      const files = await listTransactionFiles();
      sendJson(response, 200, { files });
    } catch (error) {
      sendJson(response, 500, { error: error.message || 'Nie udało się odczytać folderu dane.' });
    }
    return;
  }
  if (request.method === 'POST' && url.pathname === '/api/generate') {
    try {
      const { filename } = await readJson(request);
      const files = await generateWorkbooks(filename);
      sendJson(response, 200, { files });
    } catch (error) {
      sendJson(response, 400, { error: error.message || 'Nie udało się wygenerować pliku.' });
    }
    return;
  }
  if (request.method === 'GET' && url.pathname.startsWith('/exports/')) {
    let filename;
    try {
      filename = decodeURIComponent(url.pathname.slice('/exports/'.length));
    } catch {
      response.writeHead(400).end('Nieprawidłowa nazwa pliku.');
      return;
    }
    const safeName = path.basename(filename);
    const filePath = path.resolve(EXPORTS_DIR, safeName);
    if (safeName !== filename || !safeName.toLowerCase().endsWith('.xlsx')) {
      response.writeHead(400).end('Nieprawidłowa nazwa pliku.');
      return;
    }
    try {
      const content = await fs.readFile(filePath);
      response.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Length': content.length,
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(safeName)}`,
      });
      response.end(content);
    } catch (error) {
      if (error.code === 'ENOENT') {
        response.writeHead(404).end('Nie znaleziono pliku.');
      } else {
        response.writeHead(500).end('Nie udało się pobrać pliku.');
      }
    }
    return;
  }

  response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  response.end('Nie znaleziono strony.');
});

if (require.main === module) {
  server.listen(PORT, '127.0.0.1', () => {
    console.log(`Generator ewidencji działa: http://127.0.0.1:${PORT}`);
  });
}

module.exports = {
  createLedger,
  dateParts,
  findHeaderRow,
  generateWorkbooks,
  listTransactionFiles,
  parseAmount,
  readTransactions,
};
