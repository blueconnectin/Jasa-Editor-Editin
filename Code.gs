const TX = ['ID','Periode','Tanggal','Editor','Jasa','Harga','Tarif %','Metode bayar','Status bayar','Catatan','Perlu ditinjau','Baris Excel','Versi','Diarsipkan'];
const COST = ['Periode','Biaya lain'];
const ZONE = 'Asia/Jakarta';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Pendapatan Editor')
    .addItem('Siapkan dan impor', 'setup').addItem('Buka aplikasi', 'openPanel').addToUi();
}

function page_() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Pendapatan Editor');
}

function openPanel() {
  SpreadsheetApp.getUi().showModalDialog(page_().setWidth(1150).setHeight(730), 'Pendapatan Editor');
}

function doGet() {
  return page_().addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function locked_(action) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const result = action();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function sheet_(book, name, headers) {
  const sheet = book.getSheetByName(name) || book.insertSheet(name);
  if (!sheet.getLastRow()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#183b35').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:C').setNumberFormat('@');
  }
  if (sheet.getRange(1, 1, 1, headers.length).getValues()[0].join('|') !== headers.join('|')) {
    throw new Error('Header tab ' + name + ' berbeda. Periksa nama tab dan kolom.');
  }
  return sheet;
}

function rows_(sheet, width) {
  return sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues() : [];
}

function tables_() {
  const id = PropertiesService.getScriptProperties().getProperty('BOOK_ID');
  if (!id) throw new Error('Jalankan setup dari editor Apps Script terlebih dahulu.');
  const book = SpreadsheetApp.openById(id);
  return [sheet_(book, 'Transaksi App', TX), sheet_(book, 'Biaya Bulanan App', COST)];
}

function setup() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Buka Apps Script dari spreadsheet hasil konversi Excel.');
  const source = book.getSheets().find(sheet =>
    sheet.getName().trim().replace(/\s+/g, ' ').toLowerCase() === 'fee editor'
  );
  const count = locked_(() => {
    const tx = sheet_(book, 'Transaksi App', TX);
    const costs = sheet_(book, 'Biaya Bulanan App', COST);
    const existing = new Set(rows_(tx, TX.length).map(row => row[0]));
    const editorSheet = sheet_(book, 'Daftar Editor', ['Nama Editor']);
    const pending = source && source.getLastRow() >= 3
      ? importRows_(
          source.getRange(3, 1, source.getLastRow() - 2, 10).getValues(),
          source.getRange(3, 1, source.getLastRow() - 2, 10).getFormulas(),
          book.getSpreadsheetTimeZone()
        ).filter(row => !existing.has(row[0]))
      : [];
    const names = new Set(rows_(editorSheet, 1).map(row => row[0]));
    const newNames = [...new Set(pending.map(row => row[3]))].filter(name => !names.has(name));
    if (newNames.length) {
      editorSheet.getRange(editorSheet.getLastRow() + 1, 1, newNames.length, 1)
        .setValues(newNames.map(name => [name]));
    }
    if (pending.length) tx.getRange(tx.getLastRow() + 1, 1, pending.length, TX.length).setValues(pending);
    PropertiesService.getScriptProperties().setProperty('BOOK_ID', book.getId());
    return pending.length;
  });
  book.toast(count + ' transaksi baru diimpor.', 'Pendapatan Editor', 8);
}

function importRows_(values, formulas, zone) {
  let period = '';
  const periods = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  return values.flatMap((row, index) => {
    const [month, rawDate, editor, service, price, method,,,, note] = row;
    const sourceRow = index + 3;
    const date = rawDate instanceof Date && !isNaN(rawDate)
      ? Utilities.formatDate(rawDate, zone, 'yyyy-MM-dd') : '';
    const monthIndex = periods.indexOf(String(month).trim());
    if (monthIndex >= 0) period = '2026-' + String(monthIndex + 1).padStart(2, '0');
    if (date && validDate_(date)) period = date.slice(0, 7);
    if (!editor || !service || typeof price !== 'number') return [];
    const rates = [60,70,80].filter((rate, offset) => formulas[index][6 + offset] || typeof row[6 + offset] === 'number');
    const effectiveRates = rates.filter(rate => Number(row[[60,70,80].indexOf(rate) + 6]) > 0);
    const rate = effectiveRates.length === 1 ? effectiveRates[0] : rates.length === 1 ? rates[0] : '';
    const issues = [];
    if (!date || !validDate_(date)) issues.push('Tanggal sumber kosong atau tidak valid');
    if (rates.length !== 1) issues.push('Periksa pilihan tarif dari kolom G–I');
    const status = /belum\s+bayar/i.test(String(note)) ? 'Belum bayar' : 'Belum diketahui';
    return [[
      'excel-' + sourceRow, period, validDate_(date) ? date : '', safe_(editor), safe_(service),
      price, rate, safe_(method), status, safe_(note), issues.join('. '), sourceRow, 'awal', false
    ]];
  });
}

function validDate_(value) {
  if (!/^20\d\d-\d\d-\d\d$/.test(value)) return false;
  const date = new Date(value + 'T00:00:00Z');
  return !isNaN(date) && date.toISOString().slice(0, 10) === value;
}

function safe_(value) {
  const text = String(value ?? '').trim();
  return /^[=+\-@']/.test(text) ? "'" + text : text;
}

function readApp() {
  return locked_(() => {
    const [tx, costs] = tables_();
    const records = rows_(tx, TX.length).filter(row => row[0] && !row[13]).map(row => ({
      id: row[0], period: row[1], date: row[2], editor: row[3], service: row[4], price: row[5],
      rate: row[6], method: row[7], status: row[8], note: row[9], issue: row[10],
      source: row[11], version: row[12]
    }));
    const editors = rows_(sheet_(tx.getParent(), 'Daftar Editor', ['Nama Editor']), 1)
      .map(row => String(row[0]).trim()).filter(Boolean);
    const withdrawals = withdrawals_(withdrawalSheet_(tx.getParent()));
    const wallets = wallets_(rows_(tx, TX.length), withdrawals, editors);
    return {records, editors, withdrawals, wallets, costs: Object.fromEntries(rows_(costs, COST.length)), today: Utilities.formatDate(new Date(), ZONE, 'yyyy-MM-dd')};
  });
}

function saveTransaction(input) {
  const date = String(input.date || '');
  const price = Number(input.price);
  const rate = Number(input.rate);
  const editor = safe_(input.editor);
  const service = safe_(input.service);
  const method = safe_(input.method);
  const note = safe_(input.note);
  const status = String(input.status);
  if (!validDate_(date) || !editor || !service || String(input.price ?? '').trim() === '' ||
      !Number.isFinite(price) || price < 0 || price > 1e12 || ![60,70,80].includes(rate) ||
      !['Lunas','Belum bayar','Belum diketahui'].includes(status) ||
      editor.length > 151 || service.length > 501 || method.length > 151 || note.length > 2001) {
    throw new Error('Periksa tanggal, editor, jasa, harga, dan tarif.');
  }
  return locked_(() => {
    const [tx] = tables_();
    const rows = rows_(tx, TX.length);
    const id = String(input.id || input.requestId || '');
    if (!id || (!input.id && !/^[a-f0-9-]{36}$/i.test(id))) throw new Error('ID transaksi tidak valid.');
    const index = rows.findIndex(row => row[0] === id);
    const old = rows[index];
    if (old && !input.id) return id;
    if (input.id && (!old || old[13] || old[12] !== input.version)) {
      throw new Error('Transaksi sudah berubah. Muat ulang sebelum menyimpan.');
    }
    const record = [id, date.slice(0, 7), date, editor, service, price, rate, method,
      status, note, '', old ? old[11] : '', Utilities.getUuid(), false];
    tx.getRange(old ? index + 2 : tx.getLastRow() + 1, 1, 1, TX.length).setValues([record]);
    return id;
  });
}

function archiveTransaction(id, version) {
  locked_(() => {
    const [tx] = tables_();
    const rows = rows_(tx, TX.length);
    const index = rows.findIndex(row => row[0] === id);
    if (index < 0 || rows[index][12] !== version) throw new Error('Transaksi sudah berubah. Muat ulang.');
    tx.getRange(index + 2, TX.length).setValue(true);
  });
}

function saveCosts(period, amount) {
  const value = Number(amount);
  if (!/^20\d\d-(0[1-9]|1[0-2])$/.test(period) || String(amount ?? '').trim() === '' ||
      !Number.isFinite(value) || value < 0 || value > 1e12) throw new Error('Periksa periode dan biaya.');
  locked_(() => {
    const [, sheet] = tables_();
    const index = rows_(sheet, COST.length).findIndex(row => row[0] === period);
    sheet.getRange(index < 0 ? sheet.getLastRow() + 1 : index + 2, 1, 1, 2).setValues([[period,value]]);
  });
}

// Pengambilan disimpan sebagai riwayat. Pembatalan mempertahankan catatan asli.
const WITHDRAWAL_HEADERS = [
  'ID', 'Tanggal', 'Akun', 'Jumlah', 'Keterangan',
  'Dibuat', 'Dibatalkan', 'Alasan pembatalan'
];
const ALLOCATION_ACCOUNTS = [
  ['Investasi', 3], ['Dana darurat', 3], ['Iklan', 10],
  ['Infaq', 2], ['Mumu', 41], ['Hilmi', 10], ['Gina', 31]
];

function cents_(amount) {
  return Math.round((Number(amount) + Number.EPSILON) * 100);
}

function withdrawalSheet_(book) {
  return sheet_(book, 'Pengambilan App', WITHDRAWAL_HEADERS);
}

function withdrawals_(sheet) {
  return rows_(sheet, WITHDRAWAL_HEADERS.length).filter(row => row[0]).map(row => ({
    id: String(row[0]), date: String(row[1]), account: String(row[2]),
    amount: Number(row[3]), note: String(row[4]), created: String(row[5]),
    cancelled: row[6] === true, reason: String(row[7])
  }));
}

function wallets_(transactions, withdrawals, editors) {
  const active = transactions.filter(row => row[0] && !row[13]);
  let revenue = 0;
  let fees = 0;
  const editorFees = new Map(editors.map(name => [name, 0]));
  for (const row of active) {
    revenue += cents_(row[5]);
    const fee = row[6] ? cents_(Number(row[5]) * Number(row[6]) / 100) : 0;
    fees += fee;
    editorFees.set(String(row[3]), (editorFees.get(String(row[3])) || 0) + fee);
  }
  const basis = Math.max(revenue - fees, 0);
  let distributed = 0;
  const accounts = ALLOCATION_ACCOUNTS.map(([name, percent], index) => {
    const income = index === ALLOCATION_ACCOUNTS.length - 1
      ? basis - distributed : Math.round(basis * percent / 100);
    distributed += income;
    return {id: 'alokasi:' + name, name, type: 'Alokasi hasil usaha', income};
  });
  for (const [name, income] of editorFees) {
    accounts.push({id: 'editor:' + name, name, type: 'Fee editor', income});
  }
  // Akun lama tetap terlihat walaupun nama editornya sudah diubah.
  for (const entry of withdrawals) {
    if (!accounts.some(account => account.id === entry.account)) {
      accounts.push({
        id: entry.account, name: entry.account.replace(/^editor:/, ''),
        type: 'Fee editor', income: 0
      });
    }
  }
  const used = new Map();
  for (const entry of withdrawals.filter(entry => !entry.cancelled)) {
    used.set(entry.account, (used.get(entry.account) || 0) + cents_(entry.amount));
  }
  return accounts.map(account => ({
    ...account, income: account.income / 100,
    withdrawn: (used.get(account.id) || 0) / 100,
    balance: (account.income - (used.get(account.id) || 0)) / 100
  }));
}

function saveWithdrawal(input) {
  const id = String(input.requestId || '');
  const date = String(input.date || '');
  const account = String(input.account || '');
  const amount = Number(input.amount);
  const note = safe_(input.note);
  const today = Utilities.formatDate(new Date(), ZONE, 'yyyy-MM-dd');
  if (!/^[a-f0-9-]{36}$/i.test(id) || !validDate_(date) || date > today ||
      !Number.isFinite(amount) || amount <= 0 || amount > 1e12 ||
      Math.abs(amount * 100 - Math.round(amount * 100)) > 0.001 ||
      !note || note.length > 1001) {
    throw new Error('Isi tanggal hingga hari ini, jumlah positif, dan keterangan pengambilan.');
  }
  return locked_(() => {
    const [tx] = tables_();
    const sheet = withdrawalSheet_(tx.getParent());
    const entries = withdrawals_(sheet);
    const existing = entries.find(entry => entry.id === id);
    if (existing) {
      if (existing.account !== account || existing.amount !== amount ||
          existing.date !== date || existing.note !== note) {
        throw new Error('Permintaan ini sudah tersimpan. Muat ulang riwayat.');
      }
      return id;
    }
    const accounts = wallets_(rows_(tx, TX.length), entries, []);
    const selected = accounts.find(item => item.id === account);
    if (!selected || cents_(amount) > cents_(selected.balance)) {
      throw new Error('Saldo tidak mencukupi. Muat ulang untuk melihat saldo terbaru.');
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, WITHDRAWAL_HEADERS.length).setValues([[
      id, date, account, amount, note, new Date().toISOString(), false, ''
    ]]);
    return id;
  });
}

function cancelWithdrawal(id, reason) {
  const note = safe_(reason);
  if (!note || note.length > 1001) throw new Error('Isi alasan pembatalan.');
  return locked_(() => {
    const [tx] = tables_();
    const sheet = withdrawalSheet_(tx.getParent());
    const entries = rows_(sheet, WITHDRAWAL_HEADERS.length);
    const index = entries.findIndex(row => row[0] === id);
    if (index < 0) throw new Error('Catatan pengambilan tidak ditemukan.');
    if (entries[index][6] === true) return;
    sheet.getRange(index + 2, 7, 1, 2).setValues([[true, note]]);
  });
}
