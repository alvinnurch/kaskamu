const SHEET_NAME = 'Transaksi';
const HEADER_ROW = 6;           // baris "ID, Tanggal, Uraian..."
const RECAP_SHEET_NAME = 'Dashboard'; // ganti kalau nama tab recap kamu beda

// Cari sel "TOTAL TABUNGAN" (di sheet recap, atau kalau tidak ada, di semua tab),
// lalu baca baris label (2 baris di bawah) dan baris nilai (3 baris di bawah), kolom demi kolom.
// Hapus SEMUA spasi (sheet ini menulis "TOTAL TABUNGAN" dengan spasi di antara tiap huruf)
function normTight_(s) { return String(s).replace(/\s+/g, '').toUpperCase(); }

function findSavingsHeader_(sh) {
  const data = sh.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (normTight_(data[r][c]).indexOf('TOTALTABUNGAN') !== -1) return { data, headerRow: r, startCol: c };
    }
  }
  return null;
}

// Log tabungan/investasi: tab "Tabungan", header di baris 6,
// kolom B=Tahun, C=Bulan, D=Jenis, E=Tipe, F=Jumlah. Tiap setoran = baris baru.
const TABUNGAN_SHEET_NAME = 'Tabungan';
const TABUNGAN_HEADER_ROW = 6;
const TABUNGAN_TAHUN_COL = 2;  // B
const TABUNGAN_BULAN_COL = 3;  // C
const TABUNGAN_JENIS_COL = 4;  // D
const TABUNGAN_TIPE_COL = 5;   // E
const TABUNGAN_JUMLAH_COL = 6; // F
const BULAN_GAS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

function readSavingsLog_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TABUNGAN_SHEET_NAME);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const totals = {};
  const order = [];
  for (let r = TABUNGAN_HEADER_ROW; r < data.length; r++) {
    const jenis = String(data[r][TABUNGAN_JENIS_COL - 1] || '').trim();
    if (!jenis) continue;
    const jumlah = Number(data[r][TABUNGAN_JUMLAH_COL - 1]) || 0;
    if (!(jenis in totals)) { totals[jenis] = 0; order.push(jenis); }
    totals[jenis] += jumlah;
  }
  return order.map(label => ({ label, total: totals[label] }));
}

function readSavingsThisMonth_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TABUNGAN_SHEET_NAME);
  if (!sh) return 0;
  const data = sh.getDataRange().getValues();
  const now = new Date();
  const bulanNow = now.getMonth() + 1, tahunNow = now.getFullYear();
  let total = 0;
  for (let r = TABUNGAN_HEADER_ROW; r < data.length; r++) {
    const jenis = String(data[r][TABUNGAN_JENIS_COL - 1] || '').trim();
    if (!jenis) continue;
    const tahun = Number(data[r][TABUNGAN_TAHUN_COL - 1]);
    const bulan = Number(data[r][TABUNGAN_BULAN_COL - 1]);
    if (tahun === tahunNow && bulan === bulanNow) total += Number(data[r][TABUNGAN_JUMLAH_COL - 1]) || 0;
  }
  return total;
}

function appendSavingLog_(label, amount, tipe) {
  const sh = SpreadsheetApp.getActive().getSheetByName(TABUNGAN_SHEET_NAME);
  if (!sh) return { ok: false, error: 'Tab "Tabungan" tidak ditemukan' };
  const now = new Date();
  const data = sh.getDataRange().getValues();
  let lastFilledRow = TABUNGAN_HEADER_ROW - 1;
  for (let r = TABUNGAN_HEADER_ROW; r < data.length; r++) {
    if (String(data[r][TABUNGAN_JENIS_COL - 1] || '').trim()) lastFilledRow = r + 1;
  }
  const row = lastFilledRow + 1;
  sh.getRange(row, TABUNGAN_TAHUN_COL, 1, 5).setValues([[now.getFullYear(), now.getMonth() + 1, label, tipe || 'Tabungan', amount]]);
  return { ok: true };
}

// Baca tabel "ANGSURAN / TAGIHAN" di tab Dashboard (cari header "Jenis" + "Sisa Utang").
function readInstallmentsDashboard_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(RECAP_SHEET_NAME);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  let headerRow = -1, labelCol = -1, valueCol = -1;
  for (let r = 0; r < data.length && headerRow === -1; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (normTight_(data[r][c]) === 'JENIS') {
        for (let c2 = c + 1; c2 < data[r].length; c2++) {
          const t = normTight_(data[r][c2]);
          if (t.indexOf('SISAUTANG') !== -1 || (t.indexOf('SISA') !== -1 && t.indexOf('UTANG') !== -1)) { headerRow = r; labelCol = c; valueCol = c2; break; }
        }
        if (headerRow !== -1) break;
      }
    }
  }
  if (headerRow === -1) return null;
  const rows = [];
  for (let r = headerRow + 1; r < data.length; r++) {
    const label = String(data[r][labelCol] || '').trim();
    if (!label) break;
    const val = data[r][valueCol];
    if (val === '' || val === null || typeof val === 'undefined') continue;
    rows.push({ label, total: val });
  }
  return rows.length ? rows : null;
}

// Baca saldo tabungan utk ditampilkan: tab "Dashboard", tabel "TABUNGAN & TAGIHAN BULAN INI"
// (header baris 41, kolom G=Jenis, J=Saldo/Sisa — sheet sudah otomatis menjumlah dari tab Tabungan).
// Baca daftar kategori tabungan/investasi yang masih aktif dari tab "Setting"
// (cari sel berisi "Tabungan"/"Investasi" lalu baca daftar nama ke bawah sampai kosong).
function readActiveSavingCategories_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Setting');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  let found = null;
  for (let r = 0; r < data.length && !found; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const t = normTight_(data[r][c]);
      if (t.indexOf('TABUNGAN') !== -1 || t.indexOf('INVESTASI') !== -1) { found = { r, c }; break; }
    }
  }
  if (!found) return null;
  const list = [];
  for (let r = found.r + 1; r < data.length; r++) {
    const v = String(data[r][found.c] || '').trim();
    if (!v) break;
    list.push(v);
  }
  return list.length ? list : null;
}

function readSavingsDashboard_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(RECAP_SHEET_NAME);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  let headerRow = -1, labelCol = -1, valueCol = -1;
  for (let r = 0; r < data.length && headerRow === -1; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (normTight_(data[r][c]) === 'JENIS') {
        for (let c2 = c + 1; c2 < data[r].length; c2++) {
          const t = normTight_(data[r][c2]);
          if (t.indexOf('SALDO') !== -1 || t.indexOf('SISA') !== -1) { headerRow = r; labelCol = c; valueCol = c2; break; }
        }
        if (headerRow !== -1) break;
      }
    }
  }
  if (headerRow === -1) { headerRow = 40; labelCol = 6; valueCol = 9; }
  const rows = [];
  for (let r = headerRow + 1; r < data.length; r++) {
    const label = String(data[r][labelCol] || '').trim();
    if (!label) break;
    const val = data[r][valueCol];
    if (val === '' || val === null || typeof val === 'undefined') continue;
    rows.push({ label, total: val });
  }
  return rows.length ? rows : null;
}

function readSavings_() {
  return readSavingsDashboard_() || readSavingsLog_() || [];
}

function findIdCols_(sh) {
  const headers = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  const cols = [];
  headers.forEach((v, i) => { const t = String(v).trim(); if (t === 'ID' || t === 'ID Trs') cols.push(i + 1); });
  return cols; // [kolomPengeluaran, kolomPemasukan]
}

function findRowById_(sh, col, id) {
  if (id && String(id).indexOf('row:') === 0) return Number(String(id).slice(4));
  const lastRow = sh.getLastRow();
  const ids = lastRow > HEADER_ROW ? sh.getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 1).getValues() : [];
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return HEADER_ROW + 1 + i;
  return null;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function respond_(e, obj) {
  const cb = e.parameter.callback;
  if (cb) return ContentService.createTextOutput(cb + '(' + JSON.stringify(obj) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(obj);
}

function handleWrite_(e) {
  const action = e.parameter.action;

  if (action === 'addSaving') {
    const label = String(e.parameter.label || '').trim();
    const amount = Number(e.parameter.amount) || 0;
    if (!label || !amount) return { ok: false, error: 'Kategori atau nominal tidak valid' };
    return appendSavingLog_(label, amount, e.parameter.tipe);
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const [outCol, inCol] = findIdCols_(sh);
  const isExpense = e.parameter.type === 'expense';
  const col = isExpense ? outCol : inCol;
  const rupiah = ' Rp ' + Number(e.parameter.amount).toLocaleString('id-ID') + ',00 ';

  if (action === 'delete') {
    const row = findRowById_(sh, col, e.parameter.id);
    if (row) deleteAndShiftUp_(sh, col, row);
    return { ok: true };
  }

  if (action === 'update') {
    const row = findRowById_(sh, col, e.parameter.id);
    if (row) {
      const existingId = sh.getRange(row, col).getValue();
      sh.getRange(row, col, 1, 6).setValues([[existingId, toDateObj_(e.parameter.date), e.parameter.note, e.parameter.category, rupiah, e.parameter.method]]);
      sh.getRange(row, col + 1).setNumberFormat('dd mmm yyyy');
    }
    return { ok: true };
  }

  const newDate = toDateObj_(e.parameter.date);
  let row, count;
  if (isExpense) {
    const place = chronoPlace_(sh, outCol, newDate);
    row = place.row;
    count = place.count;
    if (place.insertNeeded) sh.insertRowBefore(row);
  } else {
    const inPlace = chronoPlace_(sh, inCol, newDate);
    count = inPlace.count;
    row = inPlace.row;
    if (inPlace.insertNeeded) sh.insertRowBefore(row);
    if (count === 0 && !inPlace.isBackdated) {
      // Bulan baru & ditambahkan di ujung (bukan backdate) — sejajarkan dgn baris awal bulan baru di Pengeluaran.
      const outInfo = monthCountAndRow_(sh, outCol, newDate);
      if (outInfo.count > 0) {
        const outMonthStartRow = outInfo.nextRow - outInfo.count;
        if (outMonthStartRow > row) {
          sh.getRange(row, inCol, outMonthStartRow - row, 6).setBackground('#000000');
          row = outMonthStartRow;
        }
      }
    }
  }
  const id = (isExpense ? 'TR/o-' : 'TR/i-') + String(count + 1).padStart(2, '0');
  sh.getRange(row, col, 1, 6).setBackground(null);
  sh.getRange(row, col, 1, 6).setValues([[id, newDate, e.parameter.note, e.parameter.category, rupiah, e.parameter.method]]);
  sh.getRange(row, col + 1).setNumberFormat('dd mmm yyyy');
  return { ok: true, id, rowRef: 'row:' + row };
}
// Cari posisi kronologis yg benar utk tanggal baru di kolom ini: baris tepat setelah
// entri terakhir yg tanggalnya <= tanggal baru (jadi backdate disisipkan di tengah, bukan di ujung).
// Juga hitung jumlah entri di bulan yg sama (di posisi manapun) utk nomor ID.
function chronoPlace_(sh, col, newDate) {
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return { row: HEADER_ROW + 1, count: 0, insertNeeded: false, isBackdated: false };
  const numRows = lastRow - HEADER_ROW;
  const ids = sh.getRange(HEADER_ROW + 1, col, numRows, 1).getValues();
  const dates = sh.getRange(HEADER_ROW + 1, col + 1, numRows, 1).getValues();
  let count = 0, anchorRow = null, lastFilledRow = HEADER_ROW;
  for (let i = 0; i < numRows; i++) {
    if (ids[i][0] === '') continue;
    lastFilledRow = HEADER_ROW + 1 + i;
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;
    if (d.getFullYear() === newDate.getFullYear() && d.getMonth() === newDate.getMonth()) count++;
    if (d <= newDate) anchorRow = HEADER_ROW + 1 + i;
  }
  const insertRow = anchorRow ? anchorRow + 1 : HEADER_ROW + 1;
  const isBackdated = anchorRow !== null && anchorRow < lastFilledRow;
  return { row: insertRow, count, insertNeeded: insertRow <= lastFilledRow, isBackdated };
}
function usedCount_(sh, c) {
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return 0;
  const vals = lastRow > HEADER_ROW ? sh.getRange(HEADER_ROW + 1, c, lastRow - HEADER_ROW, 1).getValues() : [];
  let n = 0;
  for (let i = 0; i < vals.length; i++) if (vals[i][0] !== '') n = i + 1;
  return n;
}
// Hitung berapa entri sudah ada utk bulan (dari newDate) ini di kolom tsb (utk reset nomor ID per bulan),
// serta baris kosong berikutnya.
function monthCountAndRow_(sh, col, newDate) {
  const usedTotal = usedCount_(sh, col);
  const nextRow = HEADER_ROW + 1 + usedTotal;
  let count = 0;
  for (let r = nextRow - 1; r > HEADER_ROW; r--) {
    const v = sh.getRange(r, col + 1).getValue();
    if (!(v instanceof Date)) break;
    if (v.getFullYear() === newDate.getFullYear() && v.getMonth() === newDate.getMonth()) count++;
    else break;
  }
  return { count, nextRow };
}
// Hapus transaksi di baris "row", lalu geser semua baris di bawahnya (hanya kolom ini) naik satu,
// supaya tidak ada baris kosong di tengah — kolom sisi lain (pemasukan/pengeluaran) tidak disentuh.
function deleteAndShiftUp_(sh, col, row) {
  const lastRow = sh.getLastRow();
  if (lastRow <= row) { sh.getRange(row, col, 1, 6).clearContent(); return; }
  const numRows = lastRow - row;
  const values = sh.getRange(row + 1, col, numRows, 6).getValues();
  sh.getRange(row, col, numRows, 6).setValues(values);
  sh.getRange(lastRow, col, 1, 6).clearContent();
}
function toDateObj_(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0)) : new Date();
}

function doGet(e) {
  try {
    return doGetInner_(e);
  } catch (err) {
    return respond_(e, { ok: false, error: 'Server error: ' + (err && err.message ? err.message : String(err)) });
  }
}

function doGetInner_(e) {
  const action = e.parameter.action;
  if (!action && !e.parameter.callback) {
    return ContentService.createTextOutput('Akses melalui Google Sheets menu, atau buka aplikasi dari halaman yang sudah di-hosting.');
  }
  if (action === 'resetAll') {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    const [outCol, inCol] = findIdCols_(sh);
    const lastRow = sh.getLastRow();
    if (lastRow > HEADER_ROW) {
      sh.getRange(HEADER_ROW + 1, outCol, lastRow - HEADER_ROW, 6).clearContent();
      sh.getRange(HEADER_ROW + 1, inCol, lastRow - HEADER_ROW, 6).clearContent();
    }
    return respond_(e, { ok: true });
  }
  if (action === 'create' || action === 'update' || action === 'delete' || action === 'addSaving') {
    return respond_(e, handleWrite_(e));
  }
  if (action === 'debugSavings') {
    const sh2 = SpreadsheetApp.getActive().getSheetByName(TABUNGAN_SHEET_NAME);
    return respond_(e, {
      codeVersion: 'v6-setting-filter',
      allSheetNames: SpreadsheetApp.getActive().getSheets().map(s => s.getName()),
      tabunganSheetFound: !!sh2,
      headerRow: sh2 ? sh2.getRange(TABUNGAN_HEADER_ROW - 1, 1, 1, 6).getValues()[0] : null,
      savingsLog: readSavingsLog_(),
      activeSavingCategories: readActiveSavingCategories_(),
      savingsDashboard: readSavingsDashboard_(),
    });
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const [outCol, inCol] = findIdCols_(sh);
  const lastRow = sh.getLastRow();

  function readBlock(col, type) {
    if (lastRow <= HEADER_ROW) return [];
    const values = sh.getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 6).getValues();
    const display = sh.getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 6).getDisplayValues();
    const rows = [];
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === '') continue;
      rows.push({
        id: values[i][0],
        rowRef: 'row:' + (HEADER_ROW + 1 + i),
        date: display[i][1], // teks persis seperti tampil di sel, hindari salah baca locale
        note: values[i][2], category: values[i][3], amount: values[i][4], method: values[i][5], type,
      });
    }
    return rows;
  }

  const transactions = readBlock(outCol, 'expense').concat(readBlock(inCol, 'income'));
  return respond_(e, { transactions, savings: readSavings_(), installments: readInstallmentsDashboard_(), categories: readCategories_(), summary: readDashboardSummary_(), savingsThisMonth: readSavingsThisMonth_(), methods: readMethods_(), codeVersion: 'v6-setting-filter' });
}

// Baca daftar kategori dari tab Dashboard, blok "PEMASUKAN PER KATEGORI" dan
// "PENGELUARAN PER KATEGORI" (cari 2 sel "Kategori" berurutan, baca ke bawah sampai kosong/TOTAL).
function readCategories_() {
  const dash = SpreadsheetApp.getActive().getSheetByName(RECAP_SHEET_NAME);
  if (!dash) return null;
  const data = dash.getDataRange().getValues();
  const headerCells = [];
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (normTight_(data[r][c]) === 'KATEGORI') headerCells.push({ r, c });
    }
  }
  if (headerCells.length < 2) return null;
  headerCells.sort((a, b) => a.c - b.c);
  function readList(r0, c0) {
    const labels = [], actuals = {}, budgets = {};
    for (let r = r0 + 1; r < data.length; r++) {
      const label = String(data[r][c0] || '').trim();
      if (!label || normTight_(label) === 'TOTAL') break;
      labels.push(label);
      budgets[label] = data[r][c0 + 1]; // kolom Budget (1 kolom setelah Kategori)
      actuals[label] = data[r][c0 + 2]; // kolom Actual (2 kolom setelah Kategori)
    }
    return { labels, actuals, budgets };
  }
  const income = readList(headerCells[0].r, headerCells[0].c);
  const expense = readList(headerCells[1].r, headerCells[1].c);
  return { income: income.labels, expense: expense.labels, incomeActuals: income.actuals, expenseActuals: expense.actuals, expenseBudgets: expense.budgets };
}

// Baca ringkasan bulan aktif (Saldo/Pemasukan/Pengeluaran) dari header Dashboard,
// bukan dari menjumlah seluruh transaksi (yang bisa lintas bulan).
// Baca daftar metode transaksi dari tab Dashboard (cari sel bertuliskan "Metode" lalu baca ke bawah sampai kosong).
function readMethods_() {
  const dash = SpreadsheetApp.getActive().getSheetByName(RECAP_SHEET_NAME);
  if (!dash) return null;
  const data = dash.getDataRange().getValues();
  let found = null;
  for (let r = 0; r < data.length && !found; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (normTight_(data[r][c]) === 'METODE') { found = { r, c }; break; }
    }
  }
  if (!found) return null;
  const list = [];
  for (let r = found.r + 1; r < data.length; r++) {
    const v = String(data[r][found.c] || '').trim();
    if (!v) break;
    list.push(v);
  }
  return list.length ? list : null;
}

function readDashboardSummary_() {
  const dash = SpreadsheetApp.getActive().getSheetByName(RECAP_SHEET_NAME);
  if (!dash) return null;
  const data = dash.getDataRange().getValues();
  function findExact(label) {
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        if (normTight_(data[r][c]).indexOf(normTight_(label)) !== -1) return { r, c };
      }
    }
    return null;
  }
  function findCombo(labels) {
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const t = normTight_(data[r][c]);
        if (labels.every(l => t.indexOf(normTight_(l)) !== -1)) return { r, c };
      }
    }
    return null;
  }
  const income = findExact('PEMASUKAN');
  const expense = findCombo(['PENGELUARAN', 'TABUNGAN']) || findExact('PENGELUARAN');
  const selisih = findExact('SELISIH') || findExact('SALDO');
  const periode = findExact('Periode aktif');
  if (!income || !expense) return null;
  const incomeVal = data[income.r + 1][income.c];
  const expenseVal = data[expense.r + 1][expense.c];
  let periodeText = '';
  if (periode) {
    const monthNum = Number(data[periode.r][periode.c + 1]);
    const yearNum = Number(data[periode.r][periode.c + 2]);
    if (monthNum >= 1 && monthNum <= 12 && yearNum) {
      const BULAN_ID = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
      periodeText = BULAN_ID[monthNum - 1] + ' ' + yearNum;
    }
  }
  return {
    monthLabel: periodeText,
    income: incomeVal,
    expense: expenseVal,
    balance: selisih ? data[selisih.r + 1][selisih.c] : (Number(incomeVal) || 0) - (Number(expenseVal) || 0),
  };
}

// Sementara untuk debug: tampilkan nama semua tab + sel yang mengandung kata TABUNGAN
function debugSavings_() {
  const ss = SpreadsheetApp.getActive();
  const sheetNames = ss.getSheets().map(s => s.getName());
  const hits = [];
  ss.getSheets().forEach(sh => {
    const data = sh.getDataRange().getValues();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const raw = String(data[r][c]);
        if (normTight_(raw).indexOf('TABUNGAN') !== -1) {
          hits.push({ sheet: sh.getName(), row: r + 1, col: c + 1, text: raw });
        }
      }
    }
  });
  const dump = [];
  const trSheet = ss.getSheetByName(SHEET_NAME);
  if (trSheet) {
    const all = trSheet.getDataRange().getValues();
    for (let r = 28; r < Math.min(all.length, 45); r++) dump.push(all[r]);
  }
  return { sheetNames, hits, dumpTransaksiRows29to45: dump };
}