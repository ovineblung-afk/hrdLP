/**
 * =========================================================
 * FILE: DailyReport.gs
 * FUNGSI: Menangani Panel Daily HR Report -- laporan harian
 *         bebas teks, SATU baris per tanggal (upsert by tanggal).
 *         Terpisah total dari Data_Karyawan, Data_Training,
 *         Data_Mandatory_Assignment, maupun HRLP_Riwayat.
 *
 * CARA PASANG:
 *  1. Buat file script baru di Apps Script Editor bernama "DailyReport.gs".
 *  2. Salin seluruh isi file ini ke dalamnya.
 *  3. Deploy ulang / refresh Web App. Sheet "Data_Daily_Report"
 *     akan otomatis dibuat saat fungsi pertama kali dipanggil.
 * =========================================================
 */

const SHEET_DAILY_REPORT = "Data_Daily_Report";

// Fungsi untuk memastikan sheet Data_Daily_Report ada beserta headernya
function initSheetDailyReport() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_DAILY_REPORT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DAILY_REPORT);
    sheet.appendRow(["Tanggal", "Konten Laporan", "Waktu Update"]);
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 500); // Kolom konten laporan dilebarkan agar mudah dibaca manual di Sheet
  }
  return sheet;
}

// Mengambil seluruh riwayat Daily HR Report untuk dikirim ke frontend
function getDailyReports() {
  try {
    const sheet = initSheetDailyReport();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return []; // Hanya header

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Lewati baris tanpa tanggal
      let tglCell = data[i][0];
      // Kolom Tanggal bisa tersimpan sebagai Date object (jika diketik manual di Sheet)
      // atau sebagai string "YYYY-MM-DD" (jika berasal dari <input type="date"> di frontend).
      let tanggal = (tglCell instanceof Date)
        ? Utilities.formatDate(tglCell, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(tglCell);
      result.push({
        tanggal: tanggal,
        konten: data[i][1] || "",
        waktuUpdate: data[i][2] ? String(data[i][2]) : ""
      });
    }
    return result;
  } catch (e) {
    throw new Error("Gagal mengambil data Daily HR Report: " + e.message);
  }
}

// Menyimpan (Tambah baru) atau Memperbarui (jika tanggal sudah ada) satu Daily HR Report.
// Setiap tanggal HANYA memiliki SATU baris (upsert by tanggal), sesuai kebutuhan
// "satu halaman laporan baru otomatis per hari" di frontend.
function simpanDailyReport(report) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    throw new Error("Server sedang sibuk, silakan coba lagi.");
  }
  try {
    const sheet = initSheetDailyReport();
    const data = sheet.getDataRange().getValues();

    let isUpdated = false;
    for (let i = 1; i < data.length; i++) {
      let tglCell = data[i][0];
      let tanggalRow = (tglCell instanceof Date)
        ? Utilities.formatDate(tglCell, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(tglCell);
      if (tanggalRow === report.tanggal) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[report.konten, report.waktuUpdate || new Date()]]);
        isUpdated = true;
        break;
      }
    }

    if (!isUpdated) {
      sheet.appendRow([report.tanggal, report.konten, report.waktuUpdate || new Date()]);
    }

    return "Daily HR Report berhasil disimpan.";
  } catch (e) {
    throw new Error("Gagal menyimpan Daily HR Report: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// Menghapus Daily HR Report berdasarkan tanggal (format "yyyy-MM-dd")
function hapusDailyReport(tanggal) {
  try {
    const sheet = initSheetDailyReport();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let tglCell = data[i][0];
      let tanggalRow = (tglCell instanceof Date)
        ? Utilities.formatDate(tglCell, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(tglCell);
      if (tanggalRow === tanggal) {
        sheet.deleteRow(i + 1);
        return "Daily HR Report berhasil dihapus.";
      }
    }
    throw new Error("Laporan untuk tanggal tersebut tidak ditemukan.");
  } catch (e) {
    throw new Error("Gagal menghapus Daily HR Report: " + e.message);
  }
}