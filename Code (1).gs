/****************************************************************
 *  MANDARIN CINCAY — Backend (Google Apps Script)
 *  ------------------------------------------------------------
 *  Cara pakai:
 *   1. Buka Google Sheet baru → Extensions → Apps Script
 *   2. Hapus semua isi Code.gs, paste file ini
 *   3. Jalankan fungsi  setupSheets()  sekali (izinkan akses)
 *   4. Deploy → New deployment → Web app
 *        Execute as        : Me
 *        Who has access    : Anyone
 *   5. Copy URL /exec-nya ke SCRIPT_URL di api.js
 *
 *  Kalau nanti ada perubahan kode:
 *   Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy
 *   (URL tidak berubah)
 ****************************************************************/

const SS          = SpreadsheetApp.getActiveSpreadsheet();
const FOLDER_NAME = 'MandarinCincay_BuktiBayar';

/* ── Master data ────────────────────────────────────────────────
 *  Semua isi dropdown diambil dari sheet "Master" supaya bisa diubah
 *  sendiri tanpa menyentuh kode. Nilai di bawah hanya dipakai sebagai
 *  isian awal saat sheet Master pertama kali dibuat.
 * ─────────────────────────────────────────────────────────────── */
const MASTER_COLS = ['PROGRAM', 'MODE', 'BATCH', 'PAKET', 'TOTAL_SESI', 'PROGRAM_PAKAI_BATCH'];

const MASTER_DEFAULT = {
  'PROGRAM'             : ['HSK 1','HSK 2','HSK 3','HSK 4','HSK 5','Conversation','Semi-Private','Other'],
  'MODE'                : ['Online','Offline'],
  'BATCH'               : ['SP I','SP II','SP III','SP IV'],
  'PAKET'               : ['1 Month (8x - 8h)','2 Months (16x - 16h)','3 Months (24x - 24h)','Trial (4x - 4h)'],
  'TOTAL_SESI'          : ['4','8','16','24'],
  'PROGRAM_PAKAI_BATCH' : ['Semi-Private'],
};

// Kategori materi — dipakai konsisten di frontend & laporan bulanan
// Menambah kategori materi = tambah key di sini + tambah nama kolom yang sama
// di schema sheet 'Sesi' + tambah entri di MATERI (api.js), lalu Run setupSheets().
const MATERI_KEYS  = ['mock_paper','review','dictation','vocabulary','homework','writing'];

/* ── Util ───────────────────────────────────────────────────── */
function uid() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function getSheet(name) {
  return SS.getSheetByName(name) || SS.insertSheet(name);
}

function pad2(n) { return ('0' + n).slice(-2); }

/** Bulatkan ke 2 desimal — mencegah 12.299999999 dari penjumlahan pecahan. */
function bulat2(n) { return Math.round(Number(n) * 100) / 100; }

function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate());
  }
  return String(v).trim();
}

/** Jam "HH:mm". Sheets sering mengubah "10:00" jadi nilai waktu (tanggal 1899-12-30),
 *  jadi kalau yang terbaca Date, ambil jam & menitnya — bukan tanggalnya. */
function fmtTime(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return pad2(v.getHours()) + ':' + pad2(v.getMinutes());
  const s = String(v).trim();
  const m = /^(\d{1,2})[:.](\d{2})/.exec(s);
  return m ? pad2(Number(m[1])) + ':' + m[2] : s;
}

/** Baca sheet jadi array of object, dengan normalisasi tanggal, jam & nomor WA. */
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers   = data[0];
  const WA_COLS   = ['wa_ortu', 'wa_laporan', 'wa'];
  const DATE_COLS = ['tanggal', 'tgl_mulai', 'tgl_berhenti', 'tgl_bergabung', 'tgl_kirim'];
  const TIME_COLS = ['jam_mulai', 'jam_selesai'];

  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].every(c => c === '' || c === null)) continue;
    const obj = {};
    headers.forEach((h, c) => {
      let v = data[i][c];
      if (TIME_COLS.indexOf(h) >= 0)      v = fmtTime(v);
      else if (DATE_COLS.indexOf(h) >= 0) v = fmtDate(v);
      else if (WA_COLS.indexOf(h) >= 0)   v = v === '' ? '' : String(v);
      else if (v instanceof Date)         v = fmtDate(v);
      obj[h] = v;
    });
    out.push(obj);
  }
  return out;
}

/** Tulis satu baris berdasarkan nama kolom (aman kalau urutan kolom berubah). */
function appendByHeader(sheet, obj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  headers.forEach((h, i) => { if (obj[h] !== undefined) row[i] = obj[h]; });
  sheet.appendRow(row);
}

/** Update baris berdasarkan kolom id. Hanya field yang dikirim yang diubah. */
function updateByHeader(sheet, id, obj) {
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx   = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) !== String(id)) continue;
    Object.keys(obj).forEach(k => {
      const c = headers.indexOf(k);
      if (c >= 0 && obj[k] !== undefined) sheet.getRange(i + 1, c + 1).setValue(obj[k]);
    });
    return { updated: true };
  }
  throw new Error('Data dengan id ' + id + ' tidak ditemukan');
}

function deleteByHeader(sheet, id) {
  const data  = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idIdx]) === String(id)) { sheet.deleteRow(i + 1); return { deleted: true }; }
  }
  throw new Error('Data tidak ditemukan');
}

function ok(data)  { return ContentService.createTextOutput(JSON.stringify({ status: 'ok',    data    })).setMimeType(ContentService.MimeType.JSON); }
function err(msg)  { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: msg })).setMimeType(ContentService.MimeType.JSON); }

/* ══════════════════════════════════════════════════════════════
   ROUTER
   ══════════════════════════════════════════════════════════════ */
function doGet(e)  { return route((e && e.parameter) || {}); }
function doPost(e) {
  try   { return route(JSON.parse(e.postData.contents)); }
  catch (ex) { return err('Payload tidak valid: ' + ex.message); }
}

function route(p) {
  try {
    switch (p.action) {
      /* setup & master */
      case 'setup':          return ok(setupSheets());
      case 'getMaster':      return ok(getMaster());

      /* murid */
      case 'getMurid':       return ok(getMurid(p));
      case 'getMuridByLink': return ok(getMuridByLink(p.link_id));
      case 'addMurid':       return ok(addMurid(p));
      case 'updateMurid':    return ok(updateByHeader(getSheet('Murid'), p.id, muridFields(p)));
      case 'deleteMurid':    return ok(deleteByHeader(getSheet('Murid'), p.id));

      /* guru / laoshi */
      case 'getGuru':        return ok(sheetToObjects(getSheet('Guru')));
      case 'addGuru':        return ok(addGuru(p));
      case 'updateGuru':     return ok(updateByHeader(getSheet('Guru'), p.id, guruFields(p)));
      case 'deleteGuru':     return ok(deleteByHeader(getSheet('Guru'), p.id));

      /* sesi (absensi + laporan materi jadi satu) */
      case 'getSesi':        return ok(getSesi(p));
      case 'addSesi':        return ok(addSesi(p));
      case 'addSesiBatch':   return ok(addSesiBatch(p));
      case 'updateSesi':     return ok(updateByHeader(getSheet('Sesi'), p.id, sesiFields(p)));
      case 'deleteSesi':     return ok(deleteByHeader(getSheet('Sesi'), p.id));

      /* laporan & rekap */
      case 'getRekap':          return ok(getRekap());
      case 'getLaporanBulanan': return ok(getLaporanBulanan(p.murid_id, p.bulan));

      /* invoice */
      case 'getInvoice':           return ok(getInvoice(p));
      case 'generateInvoice':      return ok(generateInvoice(p));
      case 'updateInvoiceStatus':  return ok(updateByHeader(getSheet('Invoice'), p.id, { status: p.status, tgl_kirim: p.tgl_kirim || '' }));
      case 'updateInvoiceNominal': return ok(updateByHeader(getSheet('Invoice'), p.id, { nominal: Number(p.nominal) || 0 }));
      case 'deleteInvoice':        return ok(deleteByHeader(getSheet('Invoice'), p.id));
      case 'uploadBukti':          return ok(uploadBukti(p));

      default: return err('Action tidak dikenal: ' + p.action);
    }
  } catch (ex) {
    return err(ex.message);
  }
}

/* ══════════════════════════════════════════════════════════════
   SETUP
   ══════════════════════════════════════════════════════════════ */
const SCHEMAS = {
  'Murid': [
    'id', 'nama', 'program', 'batch', 'mode', 'guru_id', 'tgl_mulai', 'paket', 'total_sesi',
    'fee_per_lesson', 'wa_ortu', 'wa_laporan', 'aktif', 'catatan', 'link_id'
  ],
  'Guru': [
    'id', 'nama', 'id_karyawan', 'jabatan', 'tgl_bergabung',
    'honor_onsite', 'honor_online', 'wa'
  ],
  'Sesi': [
    'id', 'murid_id', 'nama_murid', 'guru_id', 'tanggal', 'jam_mulai', 'jam_selesai',
    'program', 'tipe', 'mode',
    'durasi',
    'mock_paper', 'review', 'dictation', 'vocabulary', 'homework', 'writing',
    'skor', 'skor_max', 'catatan', 'timestamp'
  ],
  'Invoice': [
    'id', 'murid_id', 'nama_murid', 'bulan', 'sesi', 'nominal',
    'status', 'tgl_kirim', 'bukti_url', 'catatan'
  ],
  'Master': MASTER_COLS,
};

/* ══════════════════════════════════════════════════════════════
   MASTER — isi dropdown, bisa diedit langsung di sheet "Master"
   Tiap kolom = satu jenis dropdown. Tambah/hapus baris sesuka hati,
   baris kosong diabaikan. Perubahan muncul setelah halaman di-refresh.
   ══════════════════════════════════════════════════════════════ */
function getMaster() {
  const sheet = getSheet('Master');
  const out = { materi: MATERI_KEYS };
  const kosong = sheet.getLastRow() < 2;

  const data    = kosong ? [] : sheet.getDataRange().getValues();
  const headers = data.length ? data[0] : [];

  MASTER_COLS.forEach(col => {
    const key = col.toLowerCase();
    const c = headers.indexOf(col);
    let nilai = [];
    if (c >= 0) {
      for (let i = 1; i < data.length; i++) {
        const v = String(data[i][c] === null || data[i][c] === undefined ? '' : data[i][c]).trim();
        if (v && nilai.indexOf(v) === -1) nilai.push(v);
      }
    }
    // Kolom yang dikosongkan total jatuh kembali ke daftar bawaan,
    // supaya dropdown tidak pernah kosong sama sekali.
    out[key] = nilai.length ? nilai : MASTER_DEFAULT[col].slice();
  });
  return out;
}

/** Isi sheet Master dengan nilai bawaan — hanya kalau masih kosong. */
function seedMaster() {
  const sheet = getSheet('Master');
  if (sheet.getLastRow() >= 2) return 'Master: sudah ada isi, tidak diubah';

  const tinggi = Math.max.apply(null, MASTER_COLS.map(c => MASTER_DEFAULT[c].length));
  const rows = [];
  for (let i = 0; i < tinggi; i++) {
    rows.push(MASTER_COLS.map(c => MASTER_DEFAULT[c][i] || ''));
  }
  sheet.getRange(2, 1, rows.length, MASTER_COLS.length).setValues(rows);
  sheet.getRange(2, 5, rows.length, 1).setNumberFormat('@');   // TOTAL_SESI sebagai teks
  return 'Master: diisi nilai bawaan';
}

function setupSheets() {
  const results = [];
  Object.keys(SCHEMAS).forEach(name => {
    const headers = SCHEMAS[name];
    const sheet   = getSheet(name);
    const existing = sheet.getLastRow() > 0
      ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
      : [];

    if (existing.length === 0 || existing[0] === '') {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      results.push(name + ': dibuat');
    } else {
      // Tambahkan kolom yang belum ada (tanpa merusak data lama)
      const missing = headers.filter(h => existing.indexOf(h) === -1);
      if (missing.length) {
        sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
        results.push(name + ': +' + missing.join(', '));
      } else {
        results.push(name + ': sudah lengkap');
      }
    }

    // Styling header — maroon Mandarin Cincay
    const width = Math.max(sheet.getLastColumn(), headers.length);
    sheet.getRange(1, 1, 1, width)
         .setBackground('#7A1F2B').setFontColor('#FAF4E8').setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Kolom WA & jam disimpan sebagai teks — supaya "08xx" tidak kehilangan nol
    // dan "10:00" tidak diubah Sheets jadi nilai waktu 1899-12-30.
    const hdrNow = sheet.getRange(1, 1, 1, width).getValues()[0];
    ['wa_ortu', 'wa_laporan', 'wa', 'jam_mulai', 'jam_selesai'].forEach(k => {
      const c = hdrNow.indexOf(k);
      if (c >= 0) sheet.getRange(2, c + 1, 2000, 1).setNumberFormat('@');
    });
  });

  results.push(seedMaster());
  getSheet('Master').setFrozenRows(1);
  return results;
}

/* ══════════════════════════════════════════════════════════════
   MURID
   ══════════════════════════════════════════════════════════════ */
function muridFields(p) {
  const f = {};
  ['nama', 'program', 'batch', 'mode', 'guru_id', 'tgl_mulai', 'paket',
   'wa_ortu', 'wa_laporan', 'aktif', 'catatan']
    .forEach(k => { if (p[k] !== undefined) f[k] = p[k]; });
  if (p.total_sesi      !== undefined) f.total_sesi      = Number(p.total_sesi) || 0;
  if (p.fee_per_lesson  !== undefined) f.fee_per_lesson  = Number(p.fee_per_lesson) || 0;
  return f;
}

function addMurid(p) {
  const sheet = getSheet('Murid');
  const id      = uid();
  const link_id = String(p.nama || 'murid').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')
                  + '-' + id.slice(0, 6);
  const row = muridFields(p);
  row.id      = id;
  row.link_id = link_id;
  if (row.aktif === undefined || row.aktif === '') row.aktif = 'aktif';
  appendByHeader(sheet, row);
  return { id: id, link_id: link_id };
}

function getMurid(p) {
  let rows = sheetToObjects(getSheet('Murid'));
  if (p.id)      return rows.filter(r => r.id === p.id)[0] || null;
  if (p.program) rows = rows.filter(r => r.program === p.program);
  if (p.aktif)   rows = rows.filter(r => r.aktif   === p.aktif);
  return rows;
}

function getMuridByLink(link_id) {
  const murid = sheetToObjects(getSheet('Murid')).filter(r => r.link_id === link_id)[0];
  if (!murid) throw new Error('Murid tidak ditemukan');
  return murid;
}

/* ══════════════════════════════════════════════════════════════
   GURU / LAOSHI
   ══════════════════════════════════════════════════════════════ */
function guruFields(p) {
  const f = {};
  ['nama', 'id_karyawan', 'jabatan', 'tgl_bergabung', 'wa']
    .forEach(k => { if (p[k] !== undefined) f[k] = p[k]; });
  if (p.honor_onsite !== undefined) f.honor_onsite = Number(p.honor_onsite) || 0;
  if (p.honor_online !== undefined) f.honor_online = Number(p.honor_online) || 0;
  return f;
}

function addGuru(p) {
  const row = guruFields(p);
  row.id = uid();
  appendByHeader(getSheet('Guru'), row);
  return { id: row.id };
}

/* ══════════════════════════════════════════════════════════════
   SESI  — satu baris = satu pertemuan
   tipe : 'reguler' (memotong kuota) | 'free' (tidak memotong kuota)
   ══════════════════════════════════════════════════════════════ */
function sesiFields(p) {
  const f = {};
  ['murid_id', 'guru_id', 'tanggal', 'jam_mulai', 'jam_selesai', 'program', 'mode', 'catatan', 'skor_max']
    .forEach(k => { if (p[k] !== undefined) f[k] = p[k]; });
  MATERI_KEYS.forEach(k => { if (p[k] !== undefined) f[k] = p[k]; });
  if (p.tipe !== undefined) f.tipe = (String(p.tipe).toLowerCase() === 'free') ? 'free' : 'reguler';
  if (p.skor !== undefined) f.skor = p.skor;
  // Berapa banyak kuota yang dipotong sesi ini. 60 menit = 1, 90 menit = 1.5, dst.
  if (p.durasi !== undefined) f.durasi = durasiValid(p.durasi);
  return f;
}

/** Durasi selalu angka positif; kosong/aneh dianggap 1 sesi. */
function durasiValid(v) {
  const n = Number(v);
  return (isFinite(n) && n > 0) ? n : 1;
}

function addSesi(p) {
  const sheet = getSheet('Sesi');
  const murid = sheetToObjects(getSheet('Murid')).filter(r => r.id === p.murid_id)[0];
  const row = sesiFields(p);
  row.id         = uid();
  row.nama_murid = murid ? murid.nama : '';
  if (!row.program && murid) row.program = murid.program;
  if (!row.guru_id && murid) row.guru_id = murid.guru_id;
  if (!row.tipe)             row.tipe    = 'reguler';
  row.durasi = durasiValid(row.durasi);
  row.timestamp  = new Date().toISOString();
  appendByHeader(sheet, row);
  return { id: row.id };
}

/** Satu pertemuan Semi-Private dicatat untuk semua anggota batch sekaligus.
 *  p.murid_ids = array id murid. Tiap murid tetap dapat record sendiri,
 *  supaya kuota & invoice-nya terhitung masing-masing. */
function addSesiBatch(p) {
  const ids = p.murid_ids || [];
  if (!ids.length) throw new Error('Tidak ada murid pada batch ini');
  const hasil = ids.map(id => addSesi(Object.assign({}, p, { murid_id: id })));
  return { saved: hasil.length, ids: hasil.map(h => h.id) };
}

function getSesi(p) {
  let rows = sheetToObjects(getSheet('Sesi'));
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.guru_id)  rows = rows.filter(r => r.guru_id  === p.guru_id);
  if (p.program)  rows = rows.filter(r => r.program  === p.program);
  if (p.bulan)    rows = rows.filter(r => String(r.tanggal).indexOf(p.bulan) === 0);
  rows.sort((a, b) => String(a.tanggal) < String(b.tanggal) ? 1 : -1);
  return rows;
}

/* ══════════════════════════════════════════════════════════════
   REKAP — total sesi terpakai / sisa per murid
   ══════════════════════════════════════════════════════════════ */
function getRekap() {
  const murid = sheetToObjects(getSheet('Murid'));
  const sesi  = sheetToObjects(getSheet('Sesi'));
  const guru  = sheetToObjects(getSheet('Guru'));
  const guruMap = {};
  guru.forEach(g => { guruMap[g.id] = g.nama; });

  const byMurid = {};
  sesi.forEach(s => {
    const b = byMurid[s.murid_id] || (byMurid[s.murid_id] = { pakai: 0, free: 0, last: '' });
    const d = durasiValid(s.durasi);            // 90 menit = 1.5 sesi
    if (String(s.tipe).toLowerCase() === 'free') b.free += d; else b.pakai += d;
    if (String(s.tanggal) > b.last) b.last = String(s.tanggal);
  });

  return murid.map(m => {
    const b     = byMurid[m.id] || { pakai: 0, free: 0, last: '' };
    const total = Number(m.total_sesi) || 0;
    return {
      id: m.id, nama: m.nama, program: m.program,
      batch: m.batch || '', mode: m.mode || '',
      guru_id: m.guru_id, laoshi: guruMap[m.guru_id] || '',
      tgl_mulai: m.tgl_mulai, paket: m.paket,
      total_sesi: total,
      sesi_terpakai: bulat2(b.pakai),
      sesi_free: bulat2(b.free),
      sisa_sesi: bulat2(total - b.pakai),
      sesi_terakhir: b.last,
      fee_per_lesson: Number(m.fee_per_lesson) || 0,
      wa_ortu: m.wa_ortu, wa_laporan: m.wa_laporan,
      aktif: m.aktif, link_id: m.link_id, catatan: m.catatan
    };
  });
}

/* ══════════════════════════════════════════════════════════════
   LAPORAN BULANAN — data siap cetak
   ══════════════════════════════════════════════════════════════ */
function getLaporanBulanan(murid_id, bulan) {
  const murid = sheetToObjects(getSheet('Murid')).filter(r => r.id === murid_id)[0];
  if (!murid) throw new Error('Murid tidak ditemukan');

  const guru   = sheetToObjects(getSheet('Guru')).filter(g => g.id === murid.guru_id)[0];
  const semua  = sheetToObjects(getSheet('Sesi'))
                   .filter(s => s.murid_id === murid_id)
                   .sort((a, b) => String(a.tanggal) > String(b.tanggal) ? 1 : -1);

  const total = Number(murid.total_sesi) || 0;
  let pakai = 0, urut = 0;
  const semuaDenganSisa = semua.map(s => {
    const isFree = String(s.tipe).toLowerCase() === 'free';
    const dur    = durasiValid(s.durasi);
    if (!isFree) pakai += dur;
    urut++;
    return {
      no: urut,
      id: s.id, tanggal: s.tanggal,
      jam: (s.jam_mulai || '') + (s.jam_selesai ? ' - ' + s.jam_selesai : ''),
      program: s.program, tipe: isFree ? 'free' : 'reguler', durasi: dur,
      review: s.review, dictation: s.dictation, vocabulary: s.vocabulary,
      homework: s.homework, writing: s.writing,
      skor: s.skor, skor_max: s.skor_max, catatan: s.catatan,
      sisa: isFree ? 'FREE' : bulat2(total - pakai)
    };
  });

  const bulanRows = bulan
    ? semuaDenganSisa.filter(s => String(s.tanggal).indexOf(bulan) === 0)
    : semuaDenganSisa;

  return {
    murid: {
      nama: murid.nama, program: murid.program,
      batch: murid.batch || '', mode: murid.mode || '',
      laoshi: guru ? guru.nama : '',
      tgl_mulai: murid.tgl_mulai, paket: murid.paket,
      total_sesi: total, sesi_terpakai: bulat2(pakai), sisa_sesi: bulat2(total - pakai)
    },
    bulan: bulan || '',
    sesi: bulanRows
  };
}

/* ══════════════════════════════════════════════════════════════
   INVOICE
   ══════════════════════════════════════════════════════════════ */
function getInvoice(p) {
  let rows = sheetToObjects(getSheet('Invoice'));
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.bulan)    rows = rows.filter(r => r.bulan    === p.bulan);
  return rows.sort((a, b) => String(a.bulan) < String(b.bulan) ? 1 : -1);
}

function generateInvoice(p) {
  const murid = sheetToObjects(getSheet('Murid')).filter(r => r.id === p.murid_id)[0];
  if (!murid) throw new Error('Murid tidak ditemukan');

  // Sesi FREE tidak ditagih
  let sesi = Number(p.sesi);
  if (!sesi) {
    sesi = bulat2(sheetToObjects(getSheet('Sesi')).filter(s =>
      s.murid_id === p.murid_id &&
      String(s.tanggal).indexOf(p.bulan) === 0 &&
      String(s.tipe).toLowerCase() !== 'free'
    ).reduce(function(a, s){ return a + durasiValid(s.durasi); }, 0));
  }
  if (sesi <= 0) throw new Error('Tidak ada sesi berbayar di bulan ' + p.bulan);

  if (getInvoice({ murid_id: p.murid_id, bulan: p.bulan }).length)
    throw new Error('Invoice ' + p.bulan + ' untuk murid ini sudah ada');

  const fee = Number(murid.fee_per_lesson) || 0;
  const row = {
    id: uid(), murid_id: p.murid_id, nama_murid: murid.nama,
    bulan: p.bulan, sesi: sesi, nominal: fee * sesi,
    status: 'belum_bayar', tgl_kirim: '', bukti_url: '', catatan: p.catatan || ''
  };
  appendByHeader(getSheet('Invoice'), row);
  return row;
}

function uploadBukti(p) {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  const parts   = String(p.base64).split(',');
  const meta    = parts[0], b64 = parts[1];
  const mime    = meta.substring(meta.indexOf(':') + 1, meta.indexOf(';'));
  const blob    = Utilities.newBlob(Utilities.base64Decode(b64), mime, p.filename || 'bukti');
  const file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = file.getUrl();
  if (p.invoice_id) updateByHeader(getSheet('Invoice'), p.invoice_id, { bukti_url: url, status: 'menunggu_konfirmasi' });
  return { url: url };
}
