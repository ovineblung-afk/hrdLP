/**
 * =========================================================
 * FILE: KaryawanSakitTerlambat.gs
 * FUNGSI: Menangani CRUD data Karyawan Sakit dan Terlambat
 *         (terhubung langsung ke Google Sheets).
 * =========================================================
 */

const SHEET_KARYAWAN_SAKIT = "Data_Karyawan_Sakit";
const SHEET_KARYAWAN_TERLAMBAT = "Data_Karyawan_Terlambat";

function getSakitTerlambatDB() {
  if (typeof getDB === 'function') {
    try {
      const db = getDB();
      if (db) return db;
    } catch (e) {
      console.warn("getDB() error in KaryawanSakitTerlambat:", e);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function formatSakitTerlambatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  let str = String(val).trim();
  if (str.length === 10 && /^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  let match = str.match(/\b(\d{4})-(\d{2})-(\d{2})\b/);
  if (match) return match[0];
  return str;
}

function formatSakitTerlambatTime(cellVal, defaultTime) {
  if (!cellVal) return defaultTime;
  if (cellVal instanceof Date) {
    return Utilities.formatDate(cellVal, Session.getScriptTimeZone(), "HH:mm");
  }
  let str = String(cellVal).trim();
  if (str.length === 5 && str.includes(':')) return str;
  let match = str.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
  if (match) {
    return match[1].padStart(2, '0') + ':' + match[2];
  }
  return str || defaultTime;
}

// =========================================================
// MODUL KARYAWAN SAKIT
// =========================================================

function initSheetKaryawanSakit() {
  const ss = getSakitTerlambatDB();
  let sheet = ss.getSheetByName(SHEET_KARYAWAN_SAKIT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KARYAWAN_SAKIT);
    sheet.appendRow([
      "ID", "NIK", "Nama", "Departemen", "Jabatan",
      "Tanggal Mulai", "Tanggal Selesai", "Total Hari",
      "Jenis", "Keterangan", "Waktu Update"
    ]);
    sheet.getRange("A1:K1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getKaryawanSakit() {
  try {
    const sheet = initSheetKaryawanSakit();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return [];

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let strMulai = formatSakitTerlambatDate(data[i][5]);
      let strSelesai = formatSakitTerlambatDate(data[i][6]);

      result.push({
        id: String(data[i][0]),
        nik: String(data[i][1] || ''),
        nama: String(data[i][2] || '-'),
        dept: String(data[i][3] || '-'),
        jabatan: String(data[i][4] || '-'),
        tanggalMulai: strMulai,
        tanggalSelesai: strSelesai,
        totalHari: parseInt(data[i][7]) || 1,
        jenis: String(data[i][8] || 'Surat Dokter'),
        keterangan: String(data[i][9] || '')
      });
    }
    return result;
  } catch (e) {
    throw new Error("Gagal mengambil data Karyawan Sakit: " + e.message);
  }
}

function simpanKaryawanSakit(item) {
  try {
    const sheet = initSheetKaryawanSakit();
    const data = sheet.getDataRange().getValues();
    const waktuUpdate = new Date();

    let isUpdated = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(item.id)) {
        sheet.getRange(i + 1, 2, 1, 10).setValues([[
          item.nik || '',
          item.nama || '-',
          item.dept || '-',
          item.jabatan || '-',
          item.tanggalMulai || '',
          item.tanggalSelesai || '',
          item.totalHari || 1,
          item.jenis || 'Surat Dokter',
          item.keterangan || '',
          waktuUpdate
        ]]);
        isUpdated = true;
        break;
      }
    }

    if (!isUpdated) {
      sheet.appendRow([
        item.id || ('KS_' + Date.now()),
        item.nik || '',
        item.nama || '-',
        item.dept || '-',
        item.jabatan || '-',
        item.tanggalMulai || '',
        item.tanggalSelesai || '',
        item.totalHari || 1,
        item.jenis || 'Surat Dokter',
        item.keterangan || '',
        waktuUpdate
      ]);
    }

    return "Data Karyawan Sakit berhasil disimpan.";
  } catch (e) {
    throw new Error("Gagal menyimpan data Karyawan Sakit: " + e.message);
  }
}

function hapusKaryawanSakit(id) {
  try {
    const sheet = initSheetKaryawanSakit();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Data Karyawan Sakit berhasil dihapus.";
      }
    }
    throw new Error("ID Karyawan Sakit tidak ditemukan.");
  } catch (e) {
    throw new Error("Gagal menghapus data Karyawan Sakit: " + e.message);
  }
}

// =========================================================
// MODUL KARYAWAN TERLAMBAT
// =========================================================

function initSheetKaryawanTerlambat() {
  const ss = getSakitTerlambatDB();
  let sheet = ss.getSheetByName(SHEET_KARYAWAN_TERLAMBAT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KARYAWAN_TERLAMBAT);
    sheet.appendRow([
      "ID", "NIK", "Nama", "Departemen", "Jabatan",
      "Tanggal", "Jam Masuk", "Jam Standar", "Jumlah Menit",
      "Alasan", "Keterangan", "Waktu Update"
    ]);
    sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getKaryawanTerlambat() {
  try {
    const sheet = initSheetKaryawanTerlambat();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return [];

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let strTanggal = formatSakitTerlambatDate(data[i][5]);
      let strJamMasuk = formatSakitTerlambatTime(data[i][6], '08:30');
      let strJamStandar = formatSakitTerlambatTime(data[i][7], '08:00');

      result.push({
        id: String(data[i][0]),
        nik: String(data[i][1] || ''),
        nama: String(data[i][2] || '-'),
        dept: String(data[i][3] || '-'),
        jabatan: String(data[i][4] || '-'),
        tanggal: strTanggal,
        jamMasuk: strJamMasuk,
        jamStandar: strJamStandar,
        jumlahMenit: parseInt(data[i][8]) || 0,
        alasan: String(data[i][9] || ''),
        keterangan: String(data[i][10] || '')
      });
    }
    return result;
  } catch (e) {
    throw new Error("Gagal mengambil data Karyawan Terlambat: " + e.message);
  }
}

function simpanKaryawanTerlambat(item) {
  try {
    const sheet = initSheetKaryawanTerlambat();
    const data = sheet.getDataRange().getValues();
    const waktuUpdate = new Date();

    let isUpdated = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(item.id)) {
        sheet.getRange(i + 1, 2, 1, 11).setValues([[
          item.nik || '',
          item.nama || '-',
          item.dept || '-',
          item.jabatan || '-',
          item.tanggal || '',
          item.jamMasuk || '08:30',
          item.jamStandar || '08:00',
          item.jumlahMenit || 0,
          item.alasan || '',
          item.keterangan || '',
          waktuUpdate
        ]]);
        isUpdated = true;
        break;
      }
    }

    if (!isUpdated) {
      sheet.appendRow([
        item.id || ('KT_' + Date.now()),
        item.nik || '',
        item.nama || '-',
        item.dept || '-',
        item.jabatan || '-',
        item.tanggal || '',
        item.jamMasuk || '08:30',
        item.jamStandar || '08:00',
        item.jumlahMenit || 0,
        item.alasan || '',
        item.keterangan || '',
        waktuUpdate
      ]);
    }

    return "Data Karyawan Terlambat berhasil disimpan.";
  } catch (e) {
    throw new Error("Gagal menyimpan data Karyawan Terlambat: " + e.message);
  }
}

function hapusKaryawanTerlambat(id) {
  try {
    const sheet = initSheetKaryawanTerlambat();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Data Karyawan Terlambat berhasil dihapus.";
      }
    }
    throw new Error("ID Karyawan Terlambat tidak ditemukan.");
  } catch (e) {
    throw new Error("Gagal menghapus data Karyawan Terlambat: " + e.message);
  }
}
