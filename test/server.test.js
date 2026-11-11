const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const ExcelJS = require('exceljs');
const {
  createLedger,
  dateParts,
  generateWorkbooks,
  listTransactionFiles,
  parseAmount,
  readTransactions,
} = require('../server');

function createTemplateWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const ledger = workbook.addWorksheet(' Ewidencja Sprzedaży ');
  ledger.addRow(['Ewidencja Sprzedaży Bezrachunkowej']);
  ledger.addRow(['Lp', 'Data transakcji', 'Stawka VAT', 'Kwota brutto', 'Uwagi', 'Id transakcji']);
  ledger.addRow([null, null, null, null, null, null]);
  ledger.getRow(3).getCell(2).numFmt = 'dd.mm.yyyy';
  ledger.getRow(3).getCell(4).numFmt = '#,##0.00';
  return workbook;
}

function createTransactionsWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const tpay = workbook.addWorksheet('Tpay Transakcje');
  tpay.addRow(['Tytuł', 'Data płatności', 'Kwota']);
  tpay.addRow(['TP-100', new Date(Date.UTC(2026, 9, 28, 12)), '750,00 PLN']);
  tpay.addRow(['TP-200', new Date(Date.UTC(2026, 10, 2, 10)), 125.5]);

  const ebok = workbook.addWorksheet('Ebok PlanetPay Transakcje');
  ebok.addRow(['Nr transkacji (RRN)', 'Data i godzina transakcji', 'Kwota całkowita transakcji']);
  ebok.addRow(['630118149005', '02-11-2026 19:39:32', 200]);
  return workbook;
}

test('reads transaction values and formats Polish currency amounts', () => {
  assert.equal(parseAmount('750,00 PLN'), 750);
  assert.equal(parseAmount('1.234,56 zł'), 1234.56);
  assert.equal(parseAmount('125.50'), 125.5);
  assert.equal(parseAmount('błędna kwota'), null);
  assert.deepEqual(dateParts('28-10-2026 19:39:32'), { year: 2026, month: 10, day: 28 });

  const transactions = readTransactions(createTransactionsWorkbook());
  assert.equal(transactions.length, 3);
  assert.equal(transactions[0].amount, 750);
  assert.equal(transactions[0].notes, 'Usługi logopedyczne – rozliczenie operatora płatności (TPAY)');
  assert.equal(transactions[2].transactionId, '630118149005');
});

test('writes ledger rows, VAT, operator notes, and total into the template', () => {
  const workbook = createTemplateWorkbook();
  const transactions = readTransactions(createTransactionsWorkbook()).filter((item) => item.month === 10);
  const result = createLedger(workbook, transactions, '2026-10');
  const sheet = workbook.worksheets[0];

  assert.equal(result.total, 750);
  assert.equal(sheet.getCell('A1').value, 'Ewidencja Sprzedaży Bezrachunkowej 10/2026');
  assert.equal(sheet.getCell('A3').value, 1);
  assert.equal(sheet.getCell('B3').value.toISOString(), '2026-10-28T00:00:00.000Z');
  assert.equal(sheet.getCell('C3').value, 'zw');
  assert.equal(sheet.getCell('E3').value, 'Usługi logopedyczne – rozliczenie operatora płatności (TPAY)');
  assert.equal(sheet.getCell('F3').value, 'TP-100');
  assert.equal(sheet.getCell('A4').value, 'SUMA');
  assert.equal(sheet.getCell('D4').value.result, 750);
});

test('lists only transaction workbooks from the data folder', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'ledger-files-'));
  try {
    await Promise.all([
      fs.writeFile(path.join(root, 'transakcje 10-2026.xlsx'), ''),
      fs.writeFile(path.join(root, 'transakcje 11-2026.xlsx'), ''),
      fs.writeFile(path.join(root, 'inne.xlsx'), ''),
      fs.writeFile(path.join(root, 'transakcje 12-2026.csv'), ''),
    ]);

    const files = await listTransactionFiles(root);
    assert.deepEqual(files.map((file) => file.filename).sort(), [
      'transakcje 10-2026.xlsx',
      'transakcje 11-2026.xlsx',
    ]);
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('uses the fixed template and creates separate exports for each transaction month', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'ledger-workbooks-'));
  const dataDir = path.join(root, 'dane');
  const templateDir = path.join(root, 'template');
  const exportsDir = path.join(root, 'exports');
  const sourceFilename = 'transakcje 10-2026.xlsx';
  const outputNames = [
    'ewidencja sprzedaży bezrachunkowej - 10-2026.xlsx',
    'ewidencja sprzedaży bezrachunkowej - 11-2026.xlsx',
  ];

  try {
    await fs.mkdir(dataDir);
    await fs.mkdir(templateDir);
    await fs.writeFile(
      path.join(dataDir, sourceFilename),
      Buffer.from(await createTransactionsWorkbook().xlsx.writeBuffer()),
    );
    const templatePath = path.join(templateDir, 'ewidencja sprzedaży bezrachunkowej template.xlsx');
    await fs.writeFile(templatePath, Buffer.from(await createTemplateWorkbook().xlsx.writeBuffer()));

    const results = await generateWorkbooks(sourceFilename, { dataDir, templatePath, exportsDir });
    assert.deepEqual(results.map((result) => result.filename), outputNames);
    assert.deepEqual(results.map((result) => result.count), [1, 2]);
    assert.deepEqual(results.map((result) => result.total), [750, 325.5]);

    const november = new ExcelJS.Workbook();
    await november.xlsx.readFile(path.join(exportsDir, outputNames[1]));
    const sheet = november.worksheets[0];
    assert.equal(sheet.getCell('A1').value, 'Ewidencja Sprzedaży Bezrachunkowej 11/2026');
    assert.equal(sheet.getCell('A5').value, 'SUMA');
    assert.equal(sheet.getCell('D5').value.result, 325.5);

    await assert.rejects(
      generateWorkbooks('..\\outside.xlsx', { dataDir, templatePath, exportsDir }),
      /Wybierz plik transakcji/,
    );
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});
