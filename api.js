// ============================================================
//  MANDARIN CINCAY — Client API Helper
//  Ganti SCRIPT_URL di bawah dengan URL /exec dari Apps Script
//  (Deploy → New deployment → Web app → Anyone)
// ============================================================

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw3e6PIPiqHlzFxAJ4b-wdrMLldExwFlvTx492MOPtiorQ9izGMNTR2mdtLwrvi4mt_Bw/exec';

// Penanda versi — muncul di header aplikasi. Kalau angkanya bukan yang terbaru,
// berarti browser masih memakai file lama: tekan Ctrl+Shift+R.
const APP_BUILD = 'v11 · 23 Sep 2026';

const API = {
  /** Satu percobaan panggilan. Dipakai oleh post() yang menambahkan retry. */
  async _sekaliPanggil(payload) {
    const res = await fetch(SCRIPT_URL, {
      method     : 'POST',
      headers    : { 'Content-Type': 'text/plain;charset=utf-8' },
      body       : JSON.stringify(payload),
      // Deployment di-set "Anyone" = boleh anonim. Tanpa baris ini browser
      // ikut mengirim sesi Google, dan Google bisa mengalihkan balasan ke
      // alamat yang tidak bisa diambil kembali (404 di /macros/echo).
      credentials: 'omit',
      redirect   : 'follow',
      cache      : 'no-store',
    });

    // Balasan dibaca sebagai teks dulu. Kalau Apps Script bermasalah, yang
    // datang halaman HTML — dan JSON.parse langsung gagal tanpa penjelasan.
    const teks = await res.text();
    let json;
    try {
      json = JSON.parse(teks);
    } catch (_) {
      const judul = (teks.match(/<title[^>]*>([^<]*)<\/title>/i) || [])[1] || '';
      throw new Error(
        'Google membalas halaman, bukan data (HTTP ' + res.status +
        (judul ? ' — "' + judul.trim() + '"' : '') + '). ' +
        'Biasanya deployment perlu dibuat ulang: Apps Script → Deploy → New deployment.'
      );
    }
    if (json.status === 'error') throw new Error(json.message);
    return json.data;
  },

  /** Apps Script butuh waktu menyebarkan versi baru setelah di-deploy, dan
   *  selama jeda itu URL-nya bisa membalas halaman 404. Sekali coba ulang
   *  menutup jeda tersebut; kalau tetap gagal, errornya diteruskan apa adanya. */
  async post(payload) {
    try {
      return await API._sekaliPanggil(payload);
    } catch (e) {
      if (!/membalas halaman/i.test(e.message)) throw e;
      await new Promise(r => setTimeout(r, 1800));
      return await API._sekaliPanggil(payload);
    }
  },

  // ── SETUP / MASTER ────────────────────────────────────────
  setup:     ()  => API.post({ action: 'setup' }),
  getBoot:   ()  => API.post({ action: 'getBoot' }),   // semua data awal sekali jalan
  getMaster: ()  => API.post({ action: 'getMaster' }),

  // ── MURID ─────────────────────────────────────────────────
  getMurid:       (opts = {}) => API.post({ action: 'getMurid', ...opts }),
  getMuridByLink: (link_id)   => API.post({ action: 'getMuridByLink', link_id }),
  addMurid:       (d)         => API.post({ action: 'addMurid', ...d }),
  updateMurid:    (d)         => API.post({ action: 'updateMurid', ...d }),
  deleteMurid:    (id)        => API.post({ action: 'deleteMurid', id }),

  // ── GURU / LAOSHI ─────────────────────────────────────────
  getGuru:    ()   => API.post({ action: 'getGuru' }),
  addGuru:    (d)  => API.post({ action: 'addGuru', ...d }),
  updateGuru: (d)  => API.post({ action: 'updateGuru', ...d }),
  deleteGuru: (id) => API.post({ action: 'deleteGuru', id }),

  // ── SESI (absensi + materi jadi satu record) ──────────────
  getSesi:    (opts = {}) => API.post({ action: 'getSesi', ...opts }),
  addSesi:      (d) => API.post({ action: 'addSesi', ...d }),
  addSesiBatch: (d) => API.post({ action: 'addSesiBatch', ...d }),
  updateSesi: (d)         => API.post({ action: 'updateSesi', ...d }),
  deleteSesi: (id)        => API.post({ action: 'deleteSesi', id }),

  // ── REKAP & LAPORAN ───────────────────────────────────────
  getRekap:          ()                 => API.post({ action: 'getRekap' }),
  getLaporanBulanan: (murid_id, bulan)  => API.post({ action: 'getLaporanBulanan', murid_id, bulan }),

  // ── INVOICE ───────────────────────────────────────────────
  getInvoice:           (opts = {})            => API.post({ action: 'getInvoice', ...opts }),
  generateInvoice:      (d)                    => API.post({ action: 'generateInvoice', ...d }),
  updateInvoiceStatus:  (id, status, tgl_kirim)=> API.post({ action: 'updateInvoiceStatus', id, status, tgl_kirim }),
  updateInvoiceNominal: (id, nominal)          => API.post({ action: 'updateInvoiceNominal', id, nominal }),
  deleteInvoice:        (id)                   => API.post({ action: 'deleteInvoice', id }),

  uploadBukti: (invoice_id, murid_id, bulan, file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          resolve(await API.post({
            action: 'uploadBukti', invoice_id, murid_id, bulan,
            base64: e.target.result, filename: file.name
          }));
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    }),
};

// ── Nilai cadangan (fallback) ───────────────────────────────
//  Isi dropdown yang sebenarnya diambil dari sheet "Master" lewat
//  API.getMaster(). Daftar di bawah hanya dipakai kalau sheet belum
//  terbaca — jadi form tetap bisa dipakai walau koneksi bermasalah.
//  Untuk menambah program / batch / paket: edit sheet Master, bukan file ini.
const DEF_PROGRAMS       = ['HSK 1', 'HSK 2', 'HSK 3', 'HSK 4', 'HSK 5', 'Conversation', 'Semi-Private', 'Other'];
const DEF_MODES          = ['Online', 'Offline'];
const DEF_BATCH_PRESETS  = ['SP I', 'SP II', 'SP III', 'SP IV'];
const DEF_BATCH_PROGRAMS = ['Semi-Private'];
const DEF_PAKET          = ['1 Month (8x - 8h)', '2 Months (16x - 16h)', '3 Months (24x - 24h)', 'Trial (4x - 4h)'];
const DEF_TOTAL_SESI     = ['4', '8', '16', '24'];

// Kategori materi. Urutan di sini = urutan di form input dan di laporan bulanan.
// Menambah kategori: tambah satu baris di sini, tambah key yang sama di
// MATERI_KEYS + schema sheet 'Sesi' pada Code.gs, lalu Run setupSheets().
// Kolom input di form dibuat otomatis dari daftar ini.
const MATERI = [
  { key: 'mock_paper', zh: '模拟卷', id: 'Mock Paper', ph: 'mis. H10901, H11220' },
  { key: 'material',   zh: '教材',   id: 'Material',   ph: 'mis. Huruf Vocal, Buku HSK 2 Bab 3' },
  { key: 'review',     zh: '复习',   id: 'Review',     ph: 'mis. 我 - 他, BPMF, 4+1 Tones' },
  { key: 'dictation',  zh: '听写',   id: 'Dictation',  ph: 'mis. 我 - 他' },
  { key: 'vocabulary', zh: '词汇',   id: 'Vocabulary', ph: 'mis. 你们 - 他, 她' },
  { key: 'homework',   zh: '作业',   id: 'Homework',   ph: 'mis. Buku halaman 12-14' },
  { key: 'writing',    zh: '写作',   id: 'Writing',    ph: 'mis. 她 - 那儿' },
];

// ── Tanda tangan & stempel pada laporan bulanan ─────────────
//  Yang menandatangani laporan adalah Head Teacher, bukan laoshi pengajar
//  murid — jadi tanda tangannya satu untuk semua laporan.
//  Ganti nama file di bawah kalau tanda tangannya berganti.
const TTD_LAPORAN = {
  nama    : '罗老师',
  jabatan : '主任老师 / Head Teacher',
  gambar  : 'ttd-luo.png',   // kosongkan ('') kalau tidak mau tanda tangan
  stempel : 'stempel.png',   // kosongkan ('') kalau tidak mau stempel
};
