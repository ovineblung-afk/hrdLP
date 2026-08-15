/**
 * =========================================================
 * FILE: Training.gs
 * FUNGSI: Menangani CRUD data riwayat pelatihan (Log Training),
 *         termasuk Batch Training Input (1 event -> banyak peserta).
 *
 * PERBAIKAN PENTING (dibanding versi sebelumnya):
 * 1. Sheet "Data_Training" sebelumnya cuma 6 kolom (ID, NIK, Tanggal,
 *    Topik, Durasi, Waktu Input) -- padahal form Batch Training Input
 *    di frontend sudah mengumpulkan data jauh lebih kaya (Event ID,
 *    Kategori, Trainer, Lokasi, dst). initSheetTraining() sekarang
 *    MEMPERBAIKI OTOMATIS header sheet yang sudah ada (menambah kolom
 *    yang kurang), TANPA menghapus/menggeser data lama.
 * 2. Frontend memanggil google.script.run.simpanBatchTraining(...) untuk
 *    Batch Training Input, tapi fungsi ini SEBELUMNYA TIDAK ADA sama
 *    sekali di server -- akibatnya SEMUA batch training gagal tersimpan
 *    ke Google Sheets (hanya sempat tampil di layar/localStorage lalu
 *    hilang saat refresh). Fungsi ini sekarang dibuat, menyimpan SEMUA
 *    baris dalam SATU eksekusi (bukan loop simpanTraining() satu-satu
 *    yang rawan race condition -- lihat kasus serupa di HRCalendar.gs).
 * 3. simpanBatchTraining() juga menyinkronkan status Assignment Mandatory
 *    Training (kalau ada) ke sheet Data_Mandatory_Assignment dalam
 *    eksekusi yang SAMA, supaya tidak ada lagi perubahan status yang
 *    cuma "hidup" di memori browser dan hilang saat refresh.
 * =========================================================
 */

const SHEET_TRAINING = "Data_Training";
const TRAINING_HEADERS = [
  "ID", "NIK", "Tanggal", "Topik Pelatihan", "Durasi (Jam)", "Waktu Input",
  "Event ID", "Kategori", "Trainer Name", "Trainer ID", "Trainer Dept",
  "Trainer Position", "Trainer Type", "Lokasi", "Deskripsi", "Materi",
  "Sertifikat", "Created At"
];

// Fungsi untuk memastikan sheet Data_Training ada beserta headernya (18 kolom).
// Kalau sheet sudah ada tapi header-nya masih versi lama (kolom lebih sedikit),
// header akan DIPERBAIKI OTOMATIS -- data yang sudah ada TIDAK dihapus/digeser.
function initSheetTraining() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_TRAINING);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TRAINING);
    sheet.appendRow(TRAINING_HEADERS);
    sheet.getRange(1, 1, 1, TRAINING_HEADERS.length).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    return sheet;
  }

  const currentWidth = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, Math.min(currentWidth, TRAINING_HEADERS.length)).getValues()[0];
  const needsFix = currentWidth < TRAINING_HEADERS.length
    || TRAINING_HEADERS.some((h, idx) => currentHeaders[idx] !== h);

  if (needsFix) {
    sheet.getRange(1, 1, 1, TRAINING_HEADERS.length).setValues([TRAINING_HEADERS]);
    sheet.getRange(1, 1, 1, TRAINING_HEADERS.length).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Mengambil seluruh riwayat training untuk dikirim ke frontend
function getTraining() {
  try {
    const sheet = initSheetTraining();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return []; // Hanya header

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Lewati baris tanpa ID
      let tglCell = data[i][2];
      result.push({
        id: data[i][0],
        nik: String(data[i][1]),
        tanggal: tglCell ? (tglCell instanceof Date
          ? Utilities.formatDate(tglCell, Session.getScriptTimeZone(), "yyyy-MM-dd")
          : String(tglCell)) : "",
        topik: data[i][3] || "",
        durasi: parseFloat(data[i][4]) || 0,
        eventId: data[i][6] || "",
        kategori: data[i][7] || "",
        trainerName: data[i][8] || "",
        trainer: data[i][8] || "",
        trainerId: data[i][9] || "",
        trainerDept: data[i][10] || "",
        trainerPos: data[i][11] || "",
        trainerType: data[i][12] || "",
        lokasi: data[i][13] || "",
        deskripsi: data[i][14] || "",
        materi: data[i][15] || "",
        sertifikat: data[i][16] || "",
        createdAt: data[i][17] || ""
      });
    }
    return result;
  } catch (e) {
    throw new Error("Gagal mengambil data riwayat training: " + e.message);
  }
}

// Menyimpan SATU riwayat training (dipakai oleh alur Edit satu-record).
// Untuk Batch Training Input (banyak peserta sekaligus), frontend memakai
// simpanBatchTraining() di bawah, BUKAN fungsi ini di dalam loop.
function simpanTraining(training) {
  try {
    const sheet = initSheetTraining();
    const data = sheet.getDataRange().getValues();
    const waktuInput = new Date();

    const rowValues = [
      training.nik, training.tanggal, training.topik, training.durasi, waktuInput,
      training.eventId || "", training.kategori || "", training.trainerName || training.trainer || "",
      training.trainerId || "", training.trainerDept || "", training.trainerPos || "",
      training.trainerType || "", training.lokasi || "", training.deskripsi || "",
      training.materi || "", training.sertifikat || "", training.createdAt || new Date().toISOString()
    ];

    let isUpdated = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === training.id) {
        sheet.getRange(i + 1, 2, 1, rowValues.length).setValues([rowValues]);
        isUpdated = true;
        break;
      }
    }

    if (!isUpdated) {
      sheet.appendRow([training.id].concat(rowValues));
    }

    return "Riwayat pelatihan berhasil disimpan.";
  } catch (e) {
    throw new Error("Gagal menyimpan riwayat training: " + e.message);
  }
}

// Menyimpan BANYAK riwayat training sekaligus (Batch Training Input) DALAM SATU
// EKSEKUSI TUNGGAL, dan sekaligus menyinkronkan status Assignment Mandatory
// Training yang ikut berubah (kalau ada) ke sheet Data_Mandatory_Assignment --
// semua dalam eksekusi yang sama supaya tidak ada langkah yang tertinggal.
//
// payload = {
//   trainingItems: [ {id, nik, tanggal, topik, durasi, eventId, kategori,
//                      trainerName, trainerId, trainerDept, trainerPos,
//                      trainerType, lokasi, deskripsi, materi, sertifikat,
//                      createdAt}, ... ],
//   mandatoryUpdates: [ {id, nik, training, dueDate, status}, ... ]   // opsional
// }
function simpanBatchTraining(payload) {
  // LockService: batch input bisa menulis puluhan baris sekaligus, jadi jendela
  // waktu antara "baca getLastRow()" dan "tulis appendRow" jauh lebih panjang
  // dibanding simpanTraining() satu-record. Kalau ada 2 admin yang sama-sama
  // klik "Simpan Riwayat (Batch)" hampir bersamaan, tanpa lock keduanya bisa
  // membaca getLastRow() yang sama lalu saling menimpa baris satu sama lain.
  // Lock ini memastikan satu eksekusi batch selesai total sebelum eksekusi lain
  // (dari sesi/user manapun) boleh mulai menulis ke sheet yang sama.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tunggu maks 30 detik sebelum menyerah
  } catch (lockErr) {
    throw new Error("Server sedang sibuk memproses batch training lain, silakan coba lagi dalam beberapa detik.");
  }

  try {
    const trainingItems = (payload && payload.trainingItems) ? payload.trainingItems : (Array.isArray(payload) ? payload : []);
    const mandatoryUpdates = (payload && payload.mandatoryUpdates) ? payload.mandatoryUpdates : [];

    if (!trainingItems.length) {
      throw new Error("Tidak ada data peserta yang dikirim (trainingItems kosong). Pastikan minimal 1 peserta dipilih.");
    }
    trainingItems.forEach(function(t, idx) {
      if (!t || !t.id || !t.nik || !t.tanggal || !t.topik) {
        throw new Error("Data peserta pada baris ke-" + (idx + 1) + " tidak lengkap (id/nik/tanggal/topik wajib diisi).");
      }
    });

    // --- 1. Tulis semua riwayat training dalam satu batch ---
    const sheet = initSheetTraining();
    const data = sheet.getDataRange().getValues();
    const idToRow = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) idToRow[data[i][0]] = i + 1;
    }

    const rowsToAppend = [];
    let added = 0, updated = 0;
    const waktuInput = new Date();

    trainingItems.forEach(function(t) {
      const rowValues = [
        t.id, t.nik, t.tanggal, t.topik, t.durasi, waktuInput,
        t.eventId || "", t.kategori || "", t.trainerName || t.trainer || "",
        t.trainerId || "", t.trainerDept || "", t.trainerPos || "",
        t.trainerType || "", t.lokasi || "", t.deskripsi || "",
        t.materi || "", t.sertifikat || "", t.createdAt || new Date().toISOString()
      ];
      if (idToRow[t.id]) {
        sheet.getRange(idToRow[t.id], 2, 1, rowValues.length - 1).setValues([rowValues.slice(1)]);
        updated++;
      } else {
        rowsToAppend.push(rowValues);
        added++;
      }
    });

    if (rowsToAppend.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }
    // Flush sekarang, sebelum lock dilepas -- memastikan baris training BENAR-BENAR
    // ter-commit ke Sheet walaupun langkah sinkron mandatory di bawah ini gagal.
    SpreadsheetApp.flush();

    // --- 2. Sinkron status Assignment Mandatory Training (kalau ada) ---
    // PENTING: dibungkus try/catch TERPISAH. Riwayat training di atas sudah
    // tersimpan (dan sudah di-flush) sebelum blok ini jalan, jadi kalau sinkron
    // mandatory gagal (misal sheet Data_Mandatory_Assignment bermasalah), batch
    // training TETAP dianggap berhasil -- yang gagal cuma sinkron mandatory-nya,
    // dan itu dilaporkan lewat mandatorySyncError, bukan membatalkan seluruh
    // penyimpanan training yang sudah sukses.
    let mandatoryAdded = 0, mandatoryUpdated = 0, mandatorySyncError = null;
    if (mandatoryUpdates.length > 0) {
      try {
        const maSheet = initSheetMandatoryAssignment();
        const maData = maSheet.getDataRange().getValues();
        const maIdToRow = {};
        for (let i = 1; i < maData.length; i++) {
          if (maData[i][0]) maIdToRow[maData[i][0]] = i + 1;
        }
        const maRowsToAppend = [];
        const maWaktu = new Date();

        mandatoryUpdates.forEach(function(a) {
          const rowValues = [a.nik, a.training, a.dueDate, a.status || 'Belum', maWaktu];
          if (maIdToRow[a.id]) {
            maSheet.getRange(maIdToRow[a.id], 2, 1, rowValues.length).setValues([rowValues]);
            mandatoryUpdated++;
          } else {
            maRowsToAppend.push([a.id].concat(rowValues));
            mandatoryAdded++;
          }
        });

        if (maRowsToAppend.length > 0) {
          const maStartRow = maSheet.getLastRow() + 1;
          maSheet.getRange(maStartRow, 1, maRowsToAppend.length, maRowsToAppend[0].length).setValues(maRowsToAppend);
        }
        SpreadsheetApp.flush();
      } catch (maErr) {
        mandatorySyncError = maErr.message;
      }
    }

    return {
      added: added, updated: updated,
      mandatoryAdded: mandatoryAdded, mandatoryUpdated: mandatoryUpdated,
      mandatorySyncError: mandatorySyncError
    };
  } catch (e) {
    throw new Error("Gagal menyimpan Batch Training: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// Menghapus data riwayat training berdasarkan ID
function hapusTraining(id) {
  try {
    const sheet = initSheetTraining();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return "Riwayat pelatihan berhasil dihapus.";
      }
    }
    throw new Error("ID Training tidak ditemukan.");
  } catch (e) {
    throw new Error("Gagal menghapus riwayat training: " + e.message);
  }
}
