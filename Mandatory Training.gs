/**
 * =========================================================
 * FILE: MandatoryTraining.gs
 * FUNGSI: Menangani sistem ASSIGNMENT untuk Mandatory Training
 *         (terpisah total dari Training Reguler / Data_Training).
 *
 * Setiap baris di sheet ini = SATU penugasan (assignment) training
 * wajib untuk SATU karyawan (identitas karyawan tetap pakai NIK yang
 * sama dengan master Data_Karyawan). Karyawan yang tidak punya baris
 * di sheet ini TIDAK dihitung sama sekali oleh dashboard Mandatory
 * Training di frontend.
 *
 * CARA PASANG:
 *  1. Buat file script baru di Apps Script Editor bernama "MandatoryTraining.gs".
 *  2. Salin seluruh isi file ini ke dalamnya.
 *  3. Deploy ulang / refresh Web App. Sheet "Data_Mandatory_Assignment"
 *     akan otomatis dibuat saat fungsi pertama kali dipanggil.
 * =========================================================
 */

const SHEET_MANDATORY_ASSIGNMENT = "Data_Mandatory_Assignment";

// Fungsi untuk memastikan sheet Data_Mandatory_Assignment ada beserta headernya
function initSheetMandatoryAssignment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MANDATORY_ASSIGNMENT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MANDATORY_ASSIGNMENT);
    // Header Data_Mandatory_Assignment
    sheet.appendRow(["ID", "NIK", "Training", "Due Date", "Status", "Durasi", "Trainer Type", "Trainer ID", "Trainer Name", "External Instansi", "Waktu Update"]);
    sheet.getRange("A1:K1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Mengambil seluruh data assignment Mandatory Training untuk dikirim ke frontend
function getMandatoryAssignments() {
  try {
    const sheet = initSheetMandatoryAssignment();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return []; // Hanya header

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue; // Lewati baris tanpa NIK
      result.push({
        id: data[i][0],
        nik: String(data[i][1]),
        training: data[i][2],
        // Format Due Date ke YYYY-MM-DD agar konsisten dengan <input type="date">
        dueDate: data[i][3] ? Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        status: data[i][4] || 'Belum',
        durasi: parseFloat(data[i][5]) || 2.0,
        trainerType: data[i][6] || 'Internal',
        trainerId: String(data[i][7] || ''),
        trainerName: data[i][8] || '',
        externalInstansi: data[i][9] || ''
      });
    }
    return result;
  } catch (e) {
    throw new Error("Gagal mengambil data assignment Mandatory Training: " + e.message);
  }
}

// Menyimpan (Tambah baru) atau Memperbarui (Edit) satu assignment Mandatory Training
function simpanMandatoryAssignment(assignment) {
  try {
    const sheet = initSheetMandatoryAssignment();
    const data = sheet.getDataRange().getValues();
    const waktuUpdate = new Date();

    let isUpdated = false;

    // Cek apakah ID sudah ada (untuk proses Edit)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(assignment.id)) {
        sheet.getRange(i + 1, 2, 1, 10).setValues([[
          assignment.nik,
          assignment.training,
          assignment.dueDate,
          assignment.status || 'Belum',
          assignment.durasi || 2.0,
          assignment.trainerType || 'Internal',
          assignment.trainerId || '',
          assignment.trainerName || '',
          assignment.externalInstansi || '',
          waktuUpdate
        ]]);
        isUpdated = true;
        break;
      }
    }

    // Jika tidak diupdate (ID baru), maka tambahkan sebagai baris baru
    if (!isUpdated) {
      sheet.appendRow([
        assignment.id,
        assignment.nik,
        assignment.training,
        assignment.dueDate,
        assignment.status || 'Belum',
        assignment.durasi || 2.0,
        assignment.trainerType || 'Internal',
        assignment.trainerId || '',
        assignment.trainerName || '',
        assignment.externalInstansi || '',
        waktuUpdate
      ]);
    }

    return "Assignment Mandatory Training berhasil disimpan.";
  } catch (e) {
    throw new Error("Gagal menyimpan assignment Mandatory Training: " + e.message);
  }
}

// Menyimpan banyak (Batch / Mass) assignment Mandatory Training sekaligus
function simpanBatchMandatoryAssignments(assignments) {
  try {
    if (!Array.isArray(assignments) || assignments.length === 0) return "Tidak ada data untuk disimpan.";
    const sheet = initSheetMandatoryAssignment();
    const data = sheet.getDataRange().getValues();
    const waktuUpdate = new Date();

    let count = 0;
    assignments.forEach(function(assignment) {
      let isUpdated = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(assignment.id)) {
          sheet.getRange(i + 1, 2, 1, 10).setValues([[
            assignment.nik,
            assignment.training,
            assignment.dueDate,
            assignment.status || 'Belum',
            assignment.durasi || 2.0,
            assignment.trainerType || 'Internal',
            assignment.trainerId || '',
            assignment.trainerName || '',
            assignment.externalInstansi || '',
            waktuUpdate
          ]]);
          isUpdated = true;
          break;
        }
      }
      if (!isUpdated) {
        sheet.appendRow([
          assignment.id,
          assignment.nik,
          assignment.training,
          assignment.dueDate,
          assignment.status || 'Belum',
          assignment.durasi || 2.0,
          assignment.trainerType || 'Internal',
          assignment.trainerId || '',
          assignment.trainerName || '',
          assignment.externalInstansi || '',
          waktuUpdate
        ]);
      }
      count++;
    });

    return count + " Assignment Mandatory Training berhasil disimpan.";
  } catch (e) {
    throw new Error("Gagal simpanBatchMandatoryAssignments: " + e.message);
  }
}

// Menghapus satu assignment Mandatory Training berdasarkan ID
function hapusMandatoryAssignment(id) {
  try {
    const sheet = initSheetMandatoryAssignment();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return "Assignment Mandatory Training berhasil dihapus.";
      }
    }
    throw new Error("ID Assignment tidak ditemukan.");
  } catch (e) {
    throw new Error("Gagal menghapus assignment Mandatory Training: " + e.message);
  }
}