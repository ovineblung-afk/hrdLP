/**
 * =========================================================
 * FILE: Karyawan.gs
 * FUNGSI: Menangani CRUD data karyawan (Target Training)
 * =========================================================
 */

const SHEET_KARYAWAN = "Data_Karyawan";

// Fungsi untuk memastikan sheet Data_Karyawan ada beserta headernya
function initSheetKaryawan() {
  // PENTING: pakai getDB() (dari Code.gs) BUKAN SpreadsheetApp.getActiveSpreadsheet()
  // langsung -- supaya kalau SPREADSHEET_ID diisi di Code.gs, semua file .gs
  // konsisten menulis ke spreadsheet target yang sama, bukan sebagian ke
  // spreadsheet aktif dan sebagian ke SPREADSHEET_ID (data akan kelihatan
  // "hilang" karena sebenarnya tersebar di 2 spreadsheet berbeda).
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_KARYAWAN);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KARYAWAN);
    // Membuat header
    sheet.appendRow(["ID", "NIK", "Nama", "Departemen", "Jabatan", "Target Jam", "Mandatory", "Status", "Waktu Update"]);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Mengambil seluruh data Karyawan untuk dikirim ke frontend
function getKaryawan() {
  try {
    const sheet = initSheetKaryawan();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return []; // Hanya header
    
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue; // Lewati baris kosong/tanpa NIK
      result.push({
        id: data[i][0],
        nik: String(data[i][1]),
        nama: data[i][2],
        dept: data[i][3],
        jabatan: data[i][4],
        targetJam: parseFloat(data[i][5]) || 0,
        mandatory: data[i][6] || 'Belum',
        status: data[i][7]
      });
    }
    return result;
  } catch (e) {
    throw new Error("Gagal mengambil data karyawan: " + e.message);
  }
}

// Menyimpan (Tambah baru) atau Memperbarui (Edit) data Karyawan
function simpanKaryawan(karyawan) {
  // LockService: mencegah 2 admin yang edit master karyawan bersamaan saling menimpa baris.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    throw new Error("Server sedang sibuk, silakan coba lagi.");
  }
  try {
    const sheet = initSheetKaryawan();
    const data = sheet.getDataRange().getValues();
    const waktuUpdate = new Date();
    
    let isUpdated = false;
    
    // Cek apakah ID sudah ada (untuk proses Edit)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === karyawan.id) {
        // Update baris yang sesuai (i + 1 karena array mulai dari 0, baris sheet mulai dari 1)
        sheet.getRange(i + 1, 2, 1, 8).setValues([[
          karyawan.nik,
          karyawan.nama,
          karyawan.dept,
          karyawan.jabatan,
          karyawan.targetJam,
          karyawan.mandatory || 'Belum',
          karyawan.status,
          waktuUpdate
        ]]);
        isUpdated = true;
        break;
      }
    }
    
    // Jika tidak diupdate (ID baru), maka tambahkan sebagai baris baru
    if (!isUpdated) {
      sheet.appendRow([
        karyawan.id,
        karyawan.nik,
        karyawan.nama,
        karyawan.dept,
        karyawan.jabatan,
        karyawan.targetJam,
        karyawan.mandatory || 'Belum',
        karyawan.status,
        waktuUpdate
      ]);
    }
    
    return "Data karyawan berhasil disimpan.";
  } catch (e) {
    throw new Error("Gagal menyimpan karyawan: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// Menghapus data karyawan berdasarkan ID
function hapusKaryawan(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    throw new Error("Server sedang sibuk, silakan coba lagi.");
  }
  try {
    const sheet = initSheetKaryawan();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return "Data karyawan berhasil dihapus.";
      }
    }
    throw new Error("ID Karyawan tidak ditemukan.");
  } catch (e) {
    throw new Error("Gagal menghapus karyawan: " + e.message);
  } finally {
    lock.releaseLock();
  }
}