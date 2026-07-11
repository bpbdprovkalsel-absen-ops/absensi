/**
 * ABSENSI BPBD PROVINSI KALIMANTAN SELATAN
 * Backend Google Apps Script — menghubungkan Web App ke Google Sheets + Drive
 *
 * Dokumentasi struktur data, cara deploy, dan catatan teknis lengkap: lihat README.md.
 */

const SHEET_PEGAWAI = 'Pegawai';
const SHEET_LOG = 'LogAbsensi';
const SHEET_IZIN = 'OpsiIzin';
const SHEET_ADMIN = 'AdminUsers';
const SHEET_SETTING = 'Pengaturan';
const DRIVE_FOLDER_NAME = 'Foto Absensi BPBD';

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ====================================================================
//  CACHE (OPTIMISASI)
// ====================================================================
// Sheet "Pegawai", "OpsiIzin", dan "Pengaturan" jarang berubah tapi sering
// DIBACA (tiap kali mesin absen mencari ID, tiap absen masuk, tiap admin
// buka panel). Simpan hasil bacanya di CacheService (bersama untuk semua
// user yang memakai Web App ini) supaya tidak selalu buka Spreadsheet.
// TTL pendek (60 detik) supaya perubahan dari admin tetap terlihat cepat.
const CACHE_TTL_DETIK = 60;
const SHEET_YANG_DI_CACHE = [SHEET_PEGAWAI, SHEET_IZIN, SHEET_SETTING];

function cacheKeySheet(sheetName) {
  return 'sheet_' + sheetName;
}

// Baca sheet lewat cache kalau sheet ini termasuk yang di-cache (lihat
// SHEET_YANG_DI_CACHE); sheet lain (mis. LogAbsensi, AdminUsers) selalu
// dibaca langsung karena datanya lebih sensitif/sering berubah.
function sheetToObjectsCached(sheetName) {
  if (SHEET_YANG_DI_CACHE.indexOf(sheetName) === -1) return sheetToObjects(sheetName);
  const cache = CacheService.getScriptCache();
  const key = cacheKeySheet(sheetName);
  try {
    const cached = cache.get(key);
    if (cached !== null) return JSON.parse(cached);
  } catch (e) {
    // kalau cache korup/tidak kebaca, jatuh ke pembacaan langsung di bawah
  }
  const data = sheetToObjects(sheetName);
  try {
    cache.put(key, JSON.stringify(data), CACHE_TTL_DETIK);
  } catch (e) {
    // data terlalu besar untuk cache (>100KB) -- abaikan saja, tetap benar
    // secara fungsional, cuma tidak ter-cache.
  }
  return data;
}

// Panggil ini setelah menulis ke salah satu sheet yang di-cache, supaya
// pembacaan berikutnya langsung dapat data terbaru (bukan nunggu TTL habis).
function invalidateCacheSheet(sheetName) {
  if (SHEET_YANG_DI_CACHE.indexOf(sheetName) === -1) return;
  CacheService.getScriptCache().remove(cacheKeySheet(sheetName));
}

function setupSheets() {
  const ss = getSS();
  // Kolom "Piket": lihat normalisasiPiket(). Kolom "PembatasanRadius": lihat
  // normalisasiPembatasanRadius() & tentukanGeofenceAktifUntukPegawai().
  const PEGAWAI_HEADERS = ['ID', 'Nama', 'Jabatan', 'Bidang', 'FotoURL', 'Piket', 'PembatasanRadius'];
  if (!ss.getSheetByName(SHEET_PEGAWAI)) {
    const sh = ss.insertSheet(SHEET_PEGAWAI);
    sh.appendRow(PEGAWAI_HEADERS);
  } else {
    // migrasi: tambahkan kolom yang belum ada tanpa mengubah data lama.
    const sh = ss.getSheetByName(SHEET_PEGAWAI);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    PEGAWAI_HEADERS.forEach(h => {
      if (headers.indexOf(h) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      }
    });
  }
  const LOG_HEADERS = ['Timestamp', 'ID', 'Nama', 'Status', 'Keterangan', 'Ketepatan', 'FotoAbsen', 'FotoBukti', 'LokasiLat', 'LokasiLng', 'LokasiAkurasi', 'KeteranganPiket'];
  if (!ss.getSheetByName(SHEET_LOG)) {
    const sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(LOG_HEADERS);
  } else {
    // migrasi: tambahkan kolom yang belum ada tanpa mengubah data lama.
    const sh = ss.getSheetByName(SHEET_LOG);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    LOG_HEADERS.forEach(h => {
      if (headers.indexOf(h) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      }
    });
  }
  if (!ss.getSheetByName(SHEET_IZIN)) {
    const sh = ss.insertSheet(SHEET_IZIN);
    sh.appendRow(['No', 'Teks']);
    sh.appendRow([1, 'Izin Sakit']);
    sh.appendRow([2, 'Izin Pribadi']);
    sh.appendRow([3, 'Izin Lapangan']);
  }
  if (!ss.getSheetByName(SHEET_ADMIN)) {
    const sh = ss.insertSheet(SHEET_ADMIN);
    sh.appendRow(['Username', 'PasswordHash', 'Role']);
    // Akun bawaan — ganti segera lewat sheet ini, lihat README.
    sh.appendRow(['admin', sha256Hex('GantiSekarang123'), 'Superadmin']);
  } else {
    // migrasi: tambahkan kolom "Role" kalau sheet ini dibuat sebelum fitur
    // peran admin ada. Akun lama yang kolom Role-nya kosong dianggap
    // "Superadmin" (perilaku sama seperti sebelumnya) supaya tidak ada
    // admin yang tiba-tiba terkunci aksesnya — lihat README untuk cara
    // mengatur ulang peran tiap akun.
    const sh = ss.getSheetByName(SHEET_ADMIN);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (headers.indexOf('Role') === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue('Role');
      const jumlahAkun = sh.getLastRow() - 1;
      if (jumlahAkun > 0) {
        const kolomRole = sh.getLastColumn();
        const nilaiRoleLama = sh.getRange(2, kolomRole, jumlahAkun, 1).getValues();
        const nilaiRoleBaru = nilaiRoleLama.map(r => [r[0] ? r[0] : 'Superadmin']);
        sh.getRange(2, kolomRole, jumlahAkun, 1).setValues(nilaiRoleBaru);
      }
    }
  }
  if (!ss.getSheetByName(SHEET_SETTING)) {
    const sh = ss.insertSheet(SHEET_SETTING);
    sh.appendRow(['Key', 'Value']);
    sh.appendRow(['GeofenceAktif', false]);
    sh.appendRow(['LokasiLat', '']);
    sh.appendRow(['LokasiLng', '']);
    sh.appendRow(['RadiusMeter', 100]);
    sh.appendRow(['JamBatasTelat', '08:00:01']);
    sh.appendRow(['JamBatasTelatPiketPagi', '08:00:01']);
    sh.appendRow(['JamBatasTelatPiketMalam', '19:00:01']);
    sh.appendRow(['JamBatasTelatPiketTKB', '08:00:01']);
  } else {
    // migrasi: tambahkan key yang belum ada.
    const sh = ss.getSheetByName(SHEET_SETTING);
    const keys = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues().map(r => r[0]);
    if (keys.indexOf('JamBatasTelat') === -1) sh.appendRow(['JamBatasTelat', '08:00:01']);
    if (keys.indexOf('JamBatasTelatPiketPagi') === -1) sh.appendRow(['JamBatasTelatPiketPagi', '08:00:01']);
    if (keys.indexOf('JamBatasTelatPiketMalam') === -1) sh.appendRow(['JamBatasTelatPiketMalam', '19:00:01']);
    if (keys.indexOf('JamBatasTelatPiketTKB') === -1) sh.appendRow(['JamBatasTelatPiketTKB', '08:00:01']);
  }
}

// Nilai "Piket" yang sah. Lihat README untuk penjelasan lengkap tiap jenis
// dan catatan migrasi dari "Piket Pagi"/"Piket Malam" lama.
const PIKET_VALID = ['Tidak Piket', 'Piket Radio', 'Piket TKB'];
const PIKET_LAMA_KE_RADIO = ['Piket Pagi', 'Piket Malam']; // nilai lama, lihat migrasi di README
function normalisasiPiket(v) {
  const s = String(v || 'Tidak Piket').trim();
  if (PIKET_LAMA_KE_RADIO.indexOf(s) !== -1) return 'Piket Radio';
  return PIKET_VALID.indexOf(s) !== -1 ? s : 'Tidak Piket';
}

// Nilai "PembatasanRadius" yang sah untuk pegawai. "Default" (atau nilai lain
// yang tidak dikenali/kosong) berarti pegawai ikut pengaturan geofencing global.
const PEMBATASAN_RADIUS_VALID = ['Default', 'Aktif', 'Nonaktif'];
function normalisasiPembatasanRadius(v) {
  const s = String(v || 'Default').trim();
  return PEMBATASAN_RADIUS_VALID.indexOf(s) !== -1 ? s : 'Default';
}

// Pengecualian radius per pegawai di atas pengaturan global. Dipakai di
// catatAbsen() dan diekspos ke Mesin Absensi lewat cariPegawai().
function tentukanGeofenceAktifUntukPegawai(geofenceAktifGlobal, pegawai) {
  const override = normalisasiPembatasanRadius(pegawai ? pegawai.PembatasanRadius : 'Default');
  if (override === 'Aktif') return true;
  if (override === 'Nonaktif') return false;
  return geofenceAktifGlobal;
}


// Peran admin yang sah & normalisasinya. Nilai kosong/tidak dikenali
// SENGAJA dianggap peran paling terbatas ("Admin Lihat-saja") supaya kalau
// ada baris baru di sheet AdminUsers yang kolom Role-nya lupa diisi, admin
// itu tidak otomatis mendapat akses penuh. (Migrasi akun LAMA yang sudah
// ada sebelum fitur ini ditambahkan ditangani terpisah di setupSheets(),
// supaya akun lama tidak mendadak terkunci — lihat README.)
const ROLE_VALID = ['Superadmin', 'Admin Operasional', 'Admin Lihat-saja'];
function normalisasiRole(v) {
  const s = String(v || '').trim();
  return ROLE_VALID.indexOf(s) !== -1 ? s : 'Admin Lihat-saja';
}

// Peran mana saja yang boleh menjalankan tiap aksi TULIS (POST) di panel
// admin. Aksi yang tidak terdaftar di sini (mis. "absen" dari mesin absen,
// atau "adminLogin" itu sendiri) tidak dibatasi peran. Lihat README untuk
// penjelasan lengkap tiap peran.
const IZIN_AKSI_ADMIN = {
  addPegawai: ['Superadmin', 'Admin Operasional'],
  updatePegawai: ['Superadmin', 'Admin Operasional'],
  deletePegawai: ['Superadmin', 'Admin Operasional'],
  addOpsiIzin: ['Superadmin', 'Admin Operasional'],
  updateOpsiIzin: ['Superadmin', 'Admin Operasional'],
  deleteOpsiIzin: ['Superadmin', 'Admin Operasional'],
  updateLogAbsensi: ['Superadmin', 'Admin Operasional'],
  simpanPengaturanLokasi: ['Superadmin'],
  simpanPengaturanWaktu: ['Superadmin']
};

// Ambil peran admin berdasarkan username, LANGSUNG dari sheet AdminUsers —
// tidak pernah percaya begitu saja pada "role" yang (mungkin) dikirim dari
// browser, supaya orang tidak bisa mengubah peran sendiri lewat Console
// browser. Dipakai oleh cekIzinAksiAdmin() sebelum tiap aksi tulis.
function getRoleUntukUsername(username) {
  const sh = getSS().getSheetByName(SHEET_ADMIN);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const akun = rows.find(r => String(r[0]).trim() === String(username || '').trim());
  return akun ? normalisasiRole(akun[2]) : null;
}

// Gerbang izin server-side untuk doPost(): mengembalikan { ok:false, ... }
// kalau aksi ini dibatasi Role dan admin yang mengirim request tidak punya
// izin (atau sesinya tidak valid), atau null kalau boleh lanjut. Ini
// berjalan DI SERVER, jadi tidak bisa dilewati hanya dengan mengubah/
// menyembunyikan tombol di admin.html — lihat README.
function cekIzinAksiAdmin(body, action) {
  const daftarRoleBoleh = IZIN_AKSI_ADMIN[action];
  if (!daftarRoleBoleh) return null; // aksi ini tidak dibatasi peran
  const role = getRoleUntukUsername(body.username);
  if (!role) {
    return { ok: false, error: 'Sesi admin tidak valid/kedaluwarsa. Silakan login ulang.' };
  }
  if (daftarRoleBoleh.indexOf(role) === -1) {
    return { ok: false, error: `Akun Anda (peran: ${role}) tidak memiliki izin untuk melakukan aksi ini.` };
  }
  return null;
}


function sheetToObjects(sheetName) {
  const sh = getSS().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return values.map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => (obj[h] = row[j]));
    return obj;
  });
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Menambah baris baru berdasarkan nama header sheet (bukan urutan tetap),
// supaya aman walau urutan/kolom di sheet berubah atau ada kolom tambahan.
function appendRowByHeaders(sheet, dataObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => (dataObj[h] !== undefined ? dataObj[h] : ''));
  sheet.appendRow(row);
}

// Hash SHA-256 dalam bentuk hex — hasilnya SAMA PERSIS dengan hasil
// crypto.subtle.digest('SHA-256', ...) di browser, jadi hash yang dibuat
// lewat Console browser (lihat penjelasan di adminLogin()) bisa langsung
// ditempel ke sheet "AdminUsers" di sini.
function sha256Hex(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

// ---------- PENGATURAN (geofencing radius absen) ----------
function getPengaturan() {
  const rows = sheetToObjectsCached(SHEET_SETTING);
  const map = {};
  rows.forEach(r => { map[r.Key] = r.Value; });
  return {
    geofenceAktif: map.GeofenceAktif === true || String(map.GeofenceAktif).toUpperCase() === 'TRUE',
    lokasiLat: map.LokasiLat !== '' && map.LokasiLat !== undefined ? Number(map.LokasiLat) : null,
    lokasiLng: map.LokasiLng !== '' && map.LokasiLng !== undefined ? Number(map.LokasiLng) : null,
    radiusMeter: map.RadiusMeter !== '' && map.RadiusMeter !== undefined ? Number(map.RadiusMeter) : 100,
    jamBatasTelat: map.JamBatasTelat !== '' && map.JamBatasTelat !== undefined ? String(map.JamBatasTelat) : '08:00:01'
  };
}

// ---------- PENGATURAN (batas jam masuk — Tepat Waktu / Telat, per Jenis Piket) ----------
function getPengaturanWaktu() {
  const p = getPengaturan();
  const rows = sheetToObjectsCached(SHEET_SETTING);
  const map = {};
  rows.forEach(r => { map[r.Key] = r.Value; });
  return {
    jamBatasTelat: p.jamBatasTelat,
    jamBatasTelatPiketPagi: map.JamBatasTelatPiketPagi !== '' && map.JamBatasTelatPiketPagi !== undefined ? String(map.JamBatasTelatPiketPagi) : '08:00:01',
    jamBatasTelatPiketMalam: map.JamBatasTelatPiketMalam !== '' && map.JamBatasTelatPiketMalam !== undefined ? String(map.JamBatasTelatPiketMalam) : '19:00:01',
    jamBatasTelatPiketTKB: map.JamBatasTelatPiketTKB !== '' && map.JamBatasTelatPiketTKB !== undefined ? String(map.JamBatasTelatPiketTKB) : '08:00:01'
  };
}

// Memilih batas jam sesuai Jenis Piket pegawai. Lihat README untuk aturan lengkap.
function pilihJamBatas(jenisPiket, waktuSetting, keteranganPiket) {
  if (jenisPiket === 'Piket Radio') {
    const shift = String(keteranganPiket || '').trim();
    return shift === 'Malam' ? waktuSetting.jamBatasTelatPiketMalam : waktuSetting.jamBatasTelatPiketPagi;
  }
  if (jenisPiket === 'Piket TKB') return waktuSetting.jamBatasTelatPiketTKB;
  return waktuSetting.jamBatasTelat;
}

function simpanPengaturanWaktu(body) {
  const polaJam = /^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$/;
  const jam = String(body.jamBatasTelat || '').trim();
  const jamPagi = String(body.jamBatasTelatPiketPagi || '').trim();
  const jamMalam = String(body.jamBatasTelatPiketMalam || '').trim();
  const jamTKB = String(body.jamBatasTelatPiketTKB || '').trim();
  if (!polaJam.test(jam) || !polaJam.test(jamPagi) || !polaJam.test(jamMalam) || !polaJam.test(jamTKB)) {
    return { ok: false, error: 'Format jam tidak valid pada salah satu kolom. Gunakan format JJ:MM:DD, contoh 08:00:01' };
  }
  const sh = getSS().getSheetByName(SHEET_SETTING);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const setValue = (key, value) => {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
    }
    sh.appendRow([key, value]);
  };
  setValue('JamBatasTelat', jam);
  setValue('JamBatasTelatPiketPagi', jamPagi);
  setValue('JamBatasTelatPiketMalam', jamMalam);
  setValue('JamBatasTelatPiketTKB', jamTKB);
  invalidateCacheSheet(SHEET_SETTING);
  return { ok: true };
}

// "Telat" jika jam absen MASUK lebih besar dari batas; selain itu "Tepat Waktu".
function hitungKetepatanWaktu(tanggal, jamBatasStr) {
  const cocok = /^(\d{1,2}):(\d{1,2}):(\d{1,2})$/.exec(String(jamBatasStr || '08:00:01').trim());
  const batasDetik = cocok
    ? (Number(cocok[1]) * 3600 + Number(cocok[2]) * 60 + Number(cocok[3]))
    : (8 * 3600 + 0 * 60 + 1);
  const detikSekarang = tanggal.getHours() * 3600 + tanggal.getMinutes() * 60 + tanggal.getSeconds();
  return detikSekarang > batasDetik ? 'Telat' : 'Tepat Waktu';
}

function simpanPengaturan(body) {
  const aktif = !!body.aktif;
  if (aktif) {
    const lat = Number(body.lat), lng = Number(body.lng), radius = Number(body.radius);
    if (!isFinite(lat) || !isFinite(lng) || lat === 0 && lng === 0) {
      return { ok: false, error: 'Titik lokasi kantor belum diisi / tidak valid' };
    }
    if (!isFinite(radius) || radius <= 0) {
      return { ok: false, error: 'Radius harus lebih besar dari 0 meter' };
    }
  }
  const sh = getSS().getSheetByName(SHEET_SETTING);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const setValue = (key, value) => {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
    }
    sh.appendRow([key, value]); // jaga-jaga kalau key belum ada
  };
  setValue('GeofenceAktif', aktif);
  setValue('LokasiLat', body.lat !== undefined ? Number(body.lat) : '');
  setValue('LokasiLng', body.lng !== undefined ? Number(body.lng) : '');
  setValue('RadiusMeter', body.radius !== undefined ? Number(body.radius) : 100);
  invalidateCacheSheet(SHEET_SETTING);
  return { ok: true };
}

// Jarak antar 2 koordinat GPS dalam meter (rumus Haversine)
function hitungJarakMeter(lat1, lng1, lat2, lng2) {
  const R = 6371000; // radius bumi dalam meter
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ---------- REKAP BULANAN (per pegawai: Masuk / Tepat Waktu / Telat / Izin / Alpha) ----------
// Definisi lengkap "Alpha" & hari kerja: lihat README.
function getRekapBulanan(bulanParam, tahunParam) {
  const now = new Date();
  const bulan = bulanParam ? Number(bulanParam) : (now.getMonth() + 1); // 1-12
  const tahun = tahunParam ? Number(tahunParam) : now.getFullYear();

  const HARI_KERJA = [1, 2, 3, 4, 5]; // Senin-Jumat. Tambahkan 6 kalau Sabtu juga hari kerja.

  const akanDatang = (tahun > now.getFullYear()) || (tahun === now.getFullYear() && bulan > now.getMonth() + 1);
  const bulanBerjalan = (tahun === now.getFullYear() && bulan === now.getMonth() + 1);
  const tanggalAkhir = bulanBerjalan ? now.getDate() : new Date(tahun, bulan, 0).getDate();

  let totalHariKerja = 0;
  if (!akanDatang) {
    for (let d = 1; d <= tanggalAkhir; d++) {
      if (HARI_KERJA.indexOf(new Date(tahun, bulan - 1, d).getDay()) !== -1) totalHariKerja++;
    }
  }

  const pegawaiList = sheetToObjectsCached(SHEET_PEGAWAI);
  const logList = sheetToObjects(SHEET_LOG);

  const hasil = pegawaiList.map(p => {
    const idPegawai = String(p.ID).trim();
    const hariMasuk = new Set();
    const hariTelat = new Set();
    const hariIzin = new Set();
    // Union hari yang sudah "tertangani" (MASUK ATAU IZIN), dipakai khusus
    // untuk menghitung Alpha supaya 1 hari tidak terhitung dobel.
    const hariTertangani = new Set();

    logList.forEach(log => {
      if (String(log.ID).trim() !== idPegawai) return;
      const ts = log.Timestamp instanceof Date ? log.Timestamp : new Date(log.Timestamp);
      if (!ts || isNaN(ts.getTime())) return;
      if (ts.getFullYear() !== tahun || (ts.getMonth() + 1) !== bulan) return;
      const tglKey = ts.getFullYear() + '-' + (ts.getMonth() + 1) + '-' + ts.getDate();
      if (log.Status === 'MASUK') {
        hariMasuk.add(tglKey);
        hariTertangani.add(tglKey);
        if (log.Ketepatan === 'Telat') hariTelat.add(tglKey);
      } else if (log.Status === 'IZIN') {
        hariIzin.add(tglKey);
        hariTertangani.add(tglKey);
      }
    });

    const totalMasuk = hariMasuk.size;
    const totalTelat = hariTelat.size;
    const totalIzin = hariIzin.size;

    return {
      ID: p.ID,
      Nama: p.Nama,
      Jabatan: p.Jabatan || '',
      Bidang: p.Bidang || '',
      Piket: normalisasiPiket(p.Piket),
      TotalMasuk: totalMasuk,
      TotalTepatWaktu: totalMasuk - totalTelat,
      TotalTelat: totalTelat,
      TotalIzin: totalIzin,
      TotalAlpha: Math.max(0, totalHariKerja - hariTertangani.size)
    };
  });

  return { bulan: bulan, tahun: tahun, totalHariKerja: totalHariKerja, data: hasil };
}

function doGet(e) {
  setupSheets();
  const action = e.parameter.action;
  try {
    if (action === 'getPegawai') {
      // Normalisasi Piket & PembatasanRadius untuk data lama — lihat README.
      const list = sheetToObjectsCached(SHEET_PEGAWAI).map(p => {
        p.Piket = normalisasiPiket(p.Piket);
        p.PembatasanRadius = normalisasiPembatasanRadius(p.PembatasanRadius);
        return p;
      });
      return jsonOut({ ok: true, data: list });
    }
    if (action === 'getLog') return jsonOut({ ok: true, data: sheetToObjects(SHEET_LOG).reverse() });
    if (action === 'getOpsiIzin') return jsonOut({ ok: true, data: sheetToObjectsCached(SHEET_IZIN) });
    if (action === 'getPengaturanLokasi') return jsonOut({ ok: true, data: getPengaturan() });
    if (action === 'getPengaturanWaktu') return jsonOut({ ok: true, data: getPengaturanWaktu() });
    if (action === 'getRekapBulanan') {
      const bulan = e.parameter.bulan ? Number(e.parameter.bulan) : '';
      const tahun = e.parameter.tahun ? Number(e.parameter.tahun) : '';
      return jsonOut({ ok: true, data: getRekapBulanan(bulan, tahun) });
    }
    if (action === 'cariPegawai') {
      const id = String(e.parameter.id || '').trim();
      const list = sheetToObjectsCached(SHEET_PEGAWAI);
      const found = list.find(p => String(p.ID).trim() === id);
      // Normalisasi + info geofencing untuk pegawai ini — lihat README.
      if (found) {
        found.Piket = normalisasiPiket(found.Piket);
        found.PembatasanRadius = normalisasiPembatasanRadius(found.PembatasanRadius);
        found.GeofenceAktifUntukPegawaiIni = tentukanGeofenceAktifUntukPegawai(getPengaturan().geofenceAktif, found);
      }
      return jsonOut({ ok: true, data: found || null });
    }
    return jsonOut({ ok: false, error: 'Aksi tidak dikenali' });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// Aksi yang MENGUBAH data dan rawan tabrakan kalau 2 admin klik simpan
// hampir bersamaan (mis. dua admin menambah pegawai dengan ID yang sama
// di detik yang sama). Aksi "absen" SENGAJA tidak dikunci — itu aksi
// paling sering dipanggil (tiap pegawai check-in/out) dan append-only,
// jadi tidak butuh lock; mengunci itu hanya akan memperlambat semua
// pegawai yang sedang antre absen.
const AKSI_PERLU_LOCK = [
  'addPegawai', 'updatePegawai', 'deletePegawai',
  'addOpsiIzin', 'updateOpsiIzin', 'deleteOpsiIzin',
  'simpanPengaturanLokasi', 'simpanPengaturanWaktu', 'updateLogAbsensi'
];

function doPost(e) {
  setupSheets();
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'Body tidak valid' });
  }
  const action = body.action;
  try {
    const penolakan = cekIzinAksiAdmin(body, action);
    if (penolakan) return jsonOut(penolakan);

    if (AKSI_PERLU_LOCK.indexOf(action) !== -1) {
      const lock = LockService.getScriptLock();
      const dapatLock = lock.tryLock(10000);
      if (!dapatLock) {
        return jsonOut({ ok: false, error: 'Server sedang memproses perubahan lain, coba lagi beberapa detik.' });
      }
      try {
        return jsonOut(jalankanAksiTulisAdmin(action, body));
      } finally {
        lock.releaseLock();
      }
    }

    switch (action) {
      case 'absen': return jsonOut(catatAbsen(body));
      case 'adminLogin': return jsonOut(adminLogin(body));
      default: return jsonOut({ ok: false, error: 'Aksi tidak dikenali' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// Dipanggil hanya lewat doPost, sudah di dalam LockService lock di atas.
function jalankanAksiTulisAdmin(action, body) {
  switch (action) {
    case 'addPegawai': return addPegawai(body);
    case 'updatePegawai': return updatePegawai(body);
    case 'deletePegawai': return deleteRow(SHEET_PEGAWAI, body.row);
    case 'addOpsiIzin': return addOpsiIzin(body);
    case 'updateOpsiIzin': return updateOpsiIzin(body);
    case 'deleteOpsiIzin': return deleteRow(SHEET_IZIN, body.row);
    case 'simpanPengaturanLokasi': return simpanPengaturan(body);
    case 'simpanPengaturanWaktu': return simpanPengaturanWaktu(body);
    case 'updateLogAbsensi': return updateLogAbsensi(body);
    default: return { ok: false, error: 'Aksi tidak dikenali' };
  }
}

// ---------- PEGAWAI ----------
function addPegawai(body) {
  const id = String(body.id || '').trim();
  const nama = String(body.nama || '').trim();
  if (!id || !nama) {
    return { ok: false, error: 'ID dan Nama wajib diisi' };
  }
  // Cegah ID ganda — kalau tidak dicek, cariPegawai() di mesin absen akan
  // selalu mengambil data yang PALING ATAS saja dan membingungkan admin.
  const bentrok = sheetToObjects(SHEET_PEGAWAI).find(p => String(p.ID).trim() === id);
  if (bentrok) {
    return { ok: false, error: `ID "${id}" sudah dipakai oleh ${bentrok.Nama}. Gunakan ID lain atau edit data yang sudah ada.` };
  }
  const sh = getSS().getSheetByName(SHEET_PEGAWAI);
  sh.appendRow([id, nama, body.jabatan || '', body.bidang || '', body.fotoUrl || '', normalisasiPiket(body.piket), normalisasiPembatasanRadius(body.pembatasanRadius)]);
  invalidateCacheSheet(SHEET_PEGAWAI);
  return { ok: true };
}

function updatePegawai(body) {
  const id = String(body.id || '').trim();
  const nama = String(body.nama || '').trim();
  if (!body.row) {
    return { ok: false, error: 'Baris data tidak valid' };
  }
  if (!id || !nama) {
    return { ok: false, error: 'ID dan Nama wajib diisi' };
  }
  // Cegah ID ganda dengan pegawai LAIN (baris yang sedang diedit sendiri dikecualikan)
  const bentrok = sheetToObjects(SHEET_PEGAWAI).find(p => p._row !== Number(body.row) && String(p.ID).trim() === id);
  if (bentrok) {
    return { ok: false, error: `ID "${id}" sudah dipakai oleh ${bentrok.Nama}. Gunakan ID lain.` };
  }
  const sh = getSS().getSheetByName(SHEET_PEGAWAI);
  sh.getRange(body.row, 1, 1, 7).setValues([[id, nama, body.jabatan || '', body.bidang || '', body.fotoUrl || '', normalisasiPiket(body.piket), normalisasiPembatasanRadius(body.pembatasanRadius)]]);
  invalidateCacheSheet(SHEET_PEGAWAI);
  return { ok: true };
}

// ---------- OPSI IZIN ----------
function addOpsiIzin(body) {
  const teks = String(body.teks || '').trim();
  if (!teks) return { ok: false, error: 'Teks izin wajib diisi' };
  const sh = getSS().getSheetByName(SHEET_IZIN);
  const data = sheetToObjects(SHEET_IZIN);
  // Nomor berikutnya dihitung dari NILAI "No" TERBESAR yang ada, bukan dari
  // jumlah baris fisik — supaya tidak terjadi nomor ganda kalau sebelumnya
  // ada opsi izin di tengah yang pernah dihapus.
  const nextNo = data.reduce((max, o) => Math.max(max, Number(o.No) || 0), 0) + 1;
  sh.appendRow([nextNo, teks]);
  invalidateCacheSheet(SHEET_IZIN);
  return { ok: true };
}

function updateOpsiIzin(body) {
  const teks = String(body.teks || '').trim();
  if (!body.row) return { ok: false, error: 'Baris data tidak valid' };
  if (!teks) return { ok: false, error: 'Teks izin wajib diisi' };
  const sh = getSS().getSheetByName(SHEET_IZIN);
  sh.getRange(body.row, 2).setValue(teks);
  invalidateCacheSheet(SHEET_IZIN);
  return { ok: true };
}

// ---------- GENERIC DELETE ----------
function deleteRow(sheetName, row) {
  if (!row || Number(row) < 2) {
    return { ok: false, error: 'Baris data tidak valid' };
  }
  getSS().getSheetByName(sheetName).deleteRow(Number(row));
  invalidateCacheSheet(sheetName);
  return { ok: true };
}

// ---------- LOGIN ADMIN ----------
// Cara menambah/mengganti akun admin (tanpa perlu ubah kode ini): lihat README.
function adminLogin(body) {
  const username = String(body.username || '').trim();
  const password = String(body.password || '');
  if (!username || !password) {
    return { ok: false, error: 'Username dan password wajib diisi' };
  }
  const sh = getSS().getSheetByName(SHEET_ADMIN);
  if (!sh || sh.getLastRow() < 2) {
    return { ok: false, error: 'Belum ada akun admin terdaftar di sheet AdminUsers' };
  }
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const hashInput = sha256Hex(password);
  const akun = rows.find(r => String(r[0]).trim() === username && String(r[1]).trim() === hashInput);
  if (!akun) {
    return { ok: false, error: 'Username atau password salah' };
  }
  return { ok: true, role: normalisasiRole(akun[2]) };
}

// ---------- ABSEN (dari mesin absensi) ----------
function catatAbsen(body) {
  // body: { id, nama, status: 'MASUK'|'KELUAR'|'IZIN', keterangan, fotoBase64,
  //         fotoBukti, lokasiLat, lokasiLng, lokasiAkurasi, keteranganPiket }
  const pegawai = sheetToObjectsCached(SHEET_PEGAWAI).find(p => String(p.ID).trim() === String(body.id || '').trim());

  // Validasi radius (geofencing) dicek di server, bukan hanya di HP pegawai.
  const pengaturan = getPengaturan();
  const geofenceAktifUntukPegawaiIni = tentukanGeofenceAktifUntukPegawai(pengaturan.geofenceAktif, pegawai);
  if (geofenceAktifUntukPegawaiIni) {
    const lat = Number(body.lokasiLat), lng = Number(body.lokasiLng);
    if (!body.lokasiLat || !body.lokasiLng || !isFinite(lat) || !isFinite(lng)) {
      return { ok: false, error: 'Lokasi GPS wajib aktif untuk absen. Aktifkan izin lokasi lalu ambil foto ulang.' };
    }
    // Kalau titik lokasi kantor belum/tidak valid, tolak absen dengan pesan
    // jelas (jangan diloloskan diam-diam — lihat README).
    if (pengaturan.lokasiLat === null || pengaturan.lokasiLng === null ||
        !isFinite(pengaturan.lokasiLat) || !isFinite(pengaturan.lokasiLng)) {
      return { ok: false, error: 'Pembatasan radius lokasi sedang aktif tetapi titik lokasi kantor belum diatur. Hubungi admin untuk mengatur ulang di Panel Admin > Lokasi Absen.' };
    }
    const jarak = hitungJarakMeter(lat, lng, pengaturan.lokasiLat, pengaturan.lokasiLng);
    if (jarak > pengaturan.radiusMeter) {
      return {
        ok: false,
        error: `Anda berada di luar radius absen yang diizinkan (jarak ${Math.round(jarak)} m, maksimal ${pengaturan.radiusMeter} m dari lokasi kantor).`
      };
    }
  }

  let fotoUrl = '';
  if (body.fotoBase64) {
    fotoUrl = simpanFotoKeDrive(body.fotoBase64, body.id, body.status);
  }
  let fotoBuktiUrl = '';
  if (body.fotoBukti) {
    fotoBuktiUrl = simpanFotoBuktiKeDrive(body.fotoBukti, body.id, body.nama);
  }
  const waktuAbsen = new Date();
  const jenisPiket = normalisasiPiket(pegawai ? pegawai.Piket : 'Tidak Piket');

  // Pegawai "Piket Radio" wajib memilih shift Pagi/Malam saat MASUK/KELUAR — lihat README.
  let keteranganPiket = '';
  if (jenisPiket === 'Piket Radio' && (body.status === 'MASUK' || body.status === 'KELUAR')) {
    keteranganPiket = String(body.keteranganPiket || '').trim();
    if (['Pagi', 'Malam'].indexOf(keteranganPiket) === -1) {
      return { ok: false, error: 'Pilih dulu apakah Piket Radio ini dilaksanakan Pagi atau Malam.' };
    }
  }

  // Ketepatan hanya dihitung untuk status MASUK — lihat pilihJamBatas() / README.
  let ketepatan = '';
  if (body.status === 'MASUK') {
    const jamBatas = pilihJamBatas(jenisPiket, getPengaturanWaktu(), keteranganPiket);
    ketepatan = hitungKetepatanWaktu(waktuAbsen, jamBatas);
  }
  const sh = getSS().getSheetByName(SHEET_LOG);
  appendRowByHeaders(sh, {
    Timestamp: waktuAbsen,
    ID: body.id,
    Nama: body.nama,
    Status: body.status,
    Keterangan: body.keterangan || '',
    Ketepatan: ketepatan,
    FotoAbsen: fotoUrl,
    FotoBukti: fotoBuktiUrl,
    LokasiLat: body.lokasiLat || '',
    LokasiLng: body.lokasiLng || '',
    LokasiAkurasi: body.lokasiAkurasi || '',
    KeteranganPiket: keteranganPiket
  });
  return { ok: true, fotoUrl: fotoUrl, fotoBuktiUrl: fotoBuktiUrl, ketepatan: ketepatan };
}

// ---------- EDIT LOG ABSENSI (koreksi manual oleh admin) ----------
// Hanya menulis ulang Timestamp/Status/Keterangan/Ketepatan/KeteranganPiket;
// foto & lokasi asli tidak disentuh. Detail: lihat README.
function updateLogAbsensi(body) {
  const row = Number(body.row);
  if (!row || row < 2) return { ok: false, error: 'Baris log tidak valid.' };

  const sh = getSS().getSheetByName(SHEET_LOG);
  if (!sh || row > sh.getLastRow()) {
    return { ok: false, error: 'Baris log tidak ditemukan (mungkin sudah berubah) — muat ulang halaman lalu coba lagi.' };
  }

  const status = String(body.status || '').trim().toUpperCase();
  if (['MASUK', 'KELUAR', 'IZIN'].indexOf(status) === -1) {
    return { ok: false, error: 'Status Absensi tidak valid.' };
  }

  const cocokTgl = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(body.tanggal || '').trim());
  const cocokJam = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(String(body.jam || '').trim());
  if (!cocokTgl || !cocokJam) {
    return { ok: false, error: 'Format Tanggal atau Jam Absen tidak valid.' };
  }
  const timestampBaru = new Date(
    Number(cocokTgl[1]), Number(cocokTgl[2]) - 1, Number(cocokTgl[3]),
    Number(cocokJam[1]), Number(cocokJam[2]), 0
  );
  if (isNaN(timestampBaru.getTime())) {
    return { ok: false, error: 'Tanggal/Jam tidak valid.' };
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rowValues = sh.getRange(row, 1, 1, headers.length).getValues()[0];
  const dataLama = {};
  headers.forEach((h, i) => { dataLama[h] = rowValues[i]; });

  // Ketepatan dihitung ulang seperti di catatAbsen() — lihat README.
  const pegawai = sheetToObjectsCached(SHEET_PEGAWAI).find(p => String(p.ID).trim() === String(dataLama.ID).trim());
  const jenisPiket = normalisasiPiket(pegawai ? pegawai.Piket : 'Tidak Piket');

  let keteranganPiket = '';
  if (jenisPiket === 'Piket Radio' && (status === 'MASUK' || status === 'KELUAR')) {
    keteranganPiket = String(body.keteranganPiket || '').trim();
    if (['Pagi', 'Malam'].indexOf(keteranganPiket) === -1) {
      return { ok: false, error: 'Pilih dulu apakah Piket Radio ini dilaksanakan Pagi atau Malam.' };
    }
  }

  let ketepatan = '';
  if (status === 'MASUK') {
    const jamBatas = pilihJamBatas(jenisPiket, getPengaturanWaktu(), keteranganPiket);
    ketepatan = hitungKetepatanWaktu(timestampBaru, jamBatas);
  }

  const setCell = (namaKolom, nilai) => {
    const idx = headers.indexOf(namaKolom);
    if (idx !== -1) sh.getRange(row, idx + 1).setValue(nilai);
  };
  setCell('Timestamp', timestampBaru);
  setCell('Status', status);
  setCell('Keterangan', body.keterangan || '');
  setCell('Ketepatan', ketepatan);
  setCell('KeteranganPiket', keteranganPiket);

  return { ok: true, ketepatan: ketepatan };
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function simpanFotoKeDrive(base64Data, id, status) {
  const folder = getOrCreateFolder();
  const cleaned = base64Data.replace(/^data:image\/\w+;base64,/, '');
  const bytes = Utilities.base64Decode(cleaned);
  const blob = Utilities.newBlob(bytes, 'image/jpeg',
    `absen_${id}_${status}_${new Date().getTime()}.jpg`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?id=${file.getId()}`;
}

// ---------- FOTO BUKTI IZIN (disimpan per nama karyawan, biar gampang dicari) ----------
const DRIVE_FOLDER_BUKTI_IZIN = 'Foto Bukti Izin BPBD';

function sanitizeFolderName(name) {
  // Rapikan nama supaya aman dipakai sebagai nama folder Drive
  const bersih = String(name || 'Tanpa Nama').trim().replace(/[\/\\]/g, '-').replace(/\s+/g, ' ');
  return bersih || 'Tanpa Nama';
}

function getOrCreateSubfolder(parentFolder, subfolderName) {
  const folders = parentFolder.getFoldersByName(subfolderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(subfolderName);
}

function getOrCreateBuktiIzinFolderForPegawai(nama) {
  const rootFolders = DriveApp.getFoldersByName(DRIVE_FOLDER_BUKTI_IZIN);
  const root = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(DRIVE_FOLDER_BUKTI_IZIN);
  return getOrCreateSubfolder(root, sanitizeFolderName(nama));
}

function simpanFotoBuktiKeDrive(base64Data, id, nama) {
  // Struktur Drive: "Foto Bukti Izin BPBD" / {Nama Karyawan} / bukti_{id}_{timestamp}.jpg
  const folder = getOrCreateBuktiIzinFolderForPegawai(nama);
  const cleaned = base64Data.replace(/^data:image\/\w+;base64,/, '');
  const bytes = Utilities.base64Decode(cleaned);
  const blob = Utilities.newBlob(bytes, 'image/jpeg',
    `bukti_${id}_${new Date().getTime()}.jpg`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?id=${file.getId()}`;
}
