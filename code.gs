/**
 * =========================================================
 * FILE: Code.gs
 * FUNGSI: Entry point Web App (doGet) + util koneksi Spreadsheet.
 *
 * PERBAIKAN PENTING (anti-hang saat deploy):
 * File ini SEBELUMNYA berisi fungsi CRUD (simpanKaryawan, getKaryawan,
 * simpanTraining, getTraining, hapusTraining, hapusKaryawan,
 * getRiwayatLengkap, hapusRiwayat, dst) dengan NAMA SAMA PERSIS seperti
 * yang ada di Karyawan.gs, Training.gs, dan Summary.gs -- tapi memakai
 * nama sheet & struktur data yang BERBEDA ('data_karyawan'/'data_training'/
 * 'HR_JSON' vs 'Data_Karyawan'/'Data_Training'/'HRLP_Riwayat').
 *
 * Karena semua file .gs pada satu project Apps Script berbagi SATU scope
 * global, dua fungsi dengan nama sama akan saling menimpa tergantung
 * urutan file dimuat -- urutan ini TIDAK dijamin stabil oleh Apps Script.
 * Jika versi di file ini yang "menang", data akan terbaca/tersimpan ke
 * sheet yang salah, atau error tak tertangani (mis. hapusRiwayat dipanggil
 * frontend dengan teks Periode, padahal versi lama di file ini mengharapkan
 * nomor baris) -- inilah salah satu penyebab aplikasi terasa hang/error
 * saat dipakai.
 *
 * Fungsi-fungsi lama tsb TIDAK dihapus (agar tetap bisa ditelusuri/dipakai
 * sebagai referensi), hanya diberi prefix "LEGACY_UNUSED_" di bagian bawah
 * file ini sehingga TIDAK LAGI bentrok dengan Karyawan.gs / Training.gs /
 * Summary.gs. Frontend (Index.html) 100% memakai fungsi dari 3 file
 * modular tsb -- lihat komentar "Background Sync" pada masing-masing
 * pemanggilan google.script.run di Index.html.
 * =========================================================
 */

const SPREADSHEET_ID = ""; // Kosongkan jika script menempel di sheet

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('HR LP — Labor Cost & Productivity')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getDB() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.length > 20) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}


// =========================================================
// DI BAWAH INI: fungsi LEGACY, sengaja di-nonaktifkan lewat
// perubahan nama (prefix LEGACY_UNUSED_) agar tidak lagi
// bentrok/dipanggil tidak sengaja. Tidak dipakai oleh
// Index.html versi saat ini.
// =========================================================

function LEGACY_UNUSED_simpanDataUtama(data) {
  var ss = getDB();
  var sheet = ss.getSheetByName('HR_JSON');
  if (!sheet) {
    sheet = ss.insertSheet('HR_JSON');
    sheet.appendRow(['Periode', 'Timestamp', 'Raw_Data']);
  }

  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  var rawData = JSON.stringify(data);
  sheet.appendRow([data.periode, timestamp, rawData]);
  return "Data Laporan Performa Berhasil Disimpan!";
}

function LEGACY_UNUSED_getRiwayatLengkap() {
  var ss = getDB();
  var sheet = ss.getSheetByName('HR_JSON');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    result.push({
      id: i, // row index
      periode: row[0],
      timestamp: row[1],
      rawData: row[2]
    });
  }
  return result;
}

function LEGACY_UNUSED_hapusRiwayat(id) {
  var sheet = getDB().getSheetByName('HR_JSON');
  if(sheet) sheet.deleteRow(id + 1); // +1 karena ada header
  return true;
}

function LEGACY_UNUSED_getKaryawan() {
  var ss = getDB();
  var sheet = ss.getSheetByName('data_karyawan');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) { // Pastikan NIK ada
      result.push({
        id: data[i][0].toString(),
        nik: data[i][1].toString(),
        nama: data[i][2].toString(),
        dept: data[i][3].toString(),
        jabatan: data[i][4].toString(),
        targetJam: parseFloat(data[i][5]) || 0,
        mandatory: data[i][6] ? data[i][6].toString() : 'Belum',
        status: data[i][7] ? data[i][7].toString() : 'Aktif'
      });
    }
  }
  return result;
}

function LEGACY_UNUSED_simpanBanyakKaryawan(dataArray) {
  var ss = getDB();
  var sheet = ss.getSheetByName('data_karyawan');

  if (!sheet) {
    sheet = ss.insertSheet('data_karyawan');
    sheet.appendRow(['ID', 'NIK', 'Nama Lengkap', 'Departemen', 'Jabatan', 'Target Jam', 'Mandatory Training', 'Status']);
  }

  var rowsToAppend = [];
  for (var i = 0; i < dataArray.length; i++) {
    var k = dataArray[i];
    rowsToAppend.push([k.id, k.nik, k.nama, k.dept, k.jabatan, k.targetJam, k.mandatory, k.status]);
  }

  // Batch insert agar tidak timeout!
  if (rowsToAppend.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }
  return rowsToAppend.length;
}

function LEGACY_UNUSED_simpanKaryawan(k) {
  var ss = getDB();
  var sheet = ss.getSheetByName('data_karyawan');
  if (!sheet) {
    sheet = ss.insertSheet('data_karyawan');
    sheet.appendRow(['ID', 'NIK', 'Nama Lengkap', 'Departemen', 'Jabatan', 'Target Jam', 'Mandatory Training', 'Status']);
  }

  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == k.nik) {
      rowIndex = i + 1; break;
    }
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 8).setValues([[k.id, k.nik, k.nama, k.dept, k.jabatan, k.targetJam, k.mandatory, k.status]]);
  } else {
    sheet.appendRow([k.id, k.nik, k.nama, k.dept, k.jabatan, k.targetJam, k.mandatory, k.status]);
  }
  return "Sukses menyimpan " + k.nama;
}

function LEGACY_UNUSED_hapusKaryawan(id) {
  var sheet = getDB().getSheetByName('data_karyawan');
  if(!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function LEGACY_UNUSED_getTraining() {
  var ss = getDB();
  var sheet = ss.getSheetByName('data_training');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      result.push({
        id: data[i][0].toString(),
        nik: data[i][1].toString(),
        tanggal: data[i][2].toString(),
        topik: data[i][3].toString(),
        durasi: parseFloat(data[i][4]) || 0
      });
    }
  }
  return result;
}

function LEGACY_UNUSED_simpanTraining(t) {
  var ss = getDB();
  var sheet = ss.getSheetByName('data_training');
  if (!sheet) {
    sheet = ss.insertSheet('data_training');
    sheet.appendRow(['ID', 'NIK', 'Tanggal', 'Topik', 'Durasi']);
  }

  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == t.id) {
      rowIndex = i + 1; break;
    }
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 5).setValues([[t.id, t.nik, t.tanggal, t.topik, t.durasi]]);
  } else {
    sheet.appendRow([t.id, t.nik, t.tanggal, t.topik, t.durasi]);
  }
  return "Sukses menyimpan riwayat pelatihan";
}

function LEGACY_UNUSED_hapusTraining(id) {
  var sheet = getDB().getSheetByName('data_training');
  if(!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}
