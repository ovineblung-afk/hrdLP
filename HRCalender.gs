/**
 * =========================================================
 * FILE: HRCalendar.gs
 * FUNGSI: Menangani modul HR Calendar (sheet "HR_Calendar")
 * =========================================================
 */

const SHEET_HR_CALENDAR = "HR_Calendar";
const HR_CALENDAR_HEADERS = [
  "ID", "Tanggal", "Nama Kegiatan", "Departemen", "PIC", "Lokasi", "Status", "Keterangan",
  "Jam Mulai", "Jam Selesai", "Jenis Kegiatan", "Prioritas", "Catatan", "Dibuat Oleh",
  "Lampiran", "Tanggal Dibuat", "Waktu Update"
];

function initSheetHRCalendar() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_HR_CALENDAR);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HR_CALENDAR);
    sheet.appendRow(HR_CALENDAR_HEADERS);
    sheet.getRange(1, 1, 1, HR_CALENDAR_HEADERS.length).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    return sheet;
  }

  const currentHeaderWidth = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, Math.min(currentHeaderWidth, HR_CALENDAR_HEADERS.length)).getValues()[0];
  const headerNeedsFix = currentHeaderWidth < HR_CALENDAR_HEADERS.length
    || HR_CALENDAR_HEADERS.some((h, idx) => currentHeaders[idx] !== h);

  if (headerNeedsFix) {
    sheet.getRange(1, 1, 1, HR_CALENDAR_HEADERS.length).setValues([HR_CALENDAR_HEADERS]);
    sheet.getRange(1, 1, 1, HR_CALENDAR_HEADERS.length).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getHRCalendarEvents() {
  try {
    const sheet = initSheetHRCalendar();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return [];

    const result = [];
    const timeZone = Session.getScriptTimeZone() || "Asia/Jakarta";

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let tglCell = data[i][1];
      let tanggal = "";
      if (tglCell instanceof Date) {
        tanggal = Utilities.formatDate(tglCell, timeZone, "yyyy-MM-dd");
      } else if (tglCell) {
        tanggal = String(tglCell).trim();
      }

      result.push({
        id: String(data[i][0]),
        tanggal: tanggal,
        namaKegiatan: String(data[i][2] || ""),
        dept: String(data[i][3] || ""),
        pic: String(data[i][4] || ""),
        lokasi: String(data[i][5] || ""),
        status: String(data[i][6] || "Planning"),
        keterangan: String(data[i][7] || ""),
        jamMulai: String(data[i][8] || ""),
        jamSelesai: String(data[i][9] || ""),
        jenisKegiatan: String(data[i][10] || ""),
        prioritas: String(data[i][11] || ""),
        catatan: String(data[i][12] || ""),
        dibuatOleh: String(data[i][13] || ""),
        lampiran: String(data[i][14] || ""),
        tanggalDibuat: data[i][15] ? String(data[i][15]) : "",
        waktuUpdate: data[i][16] ? String(data[i][16]) : ""
      });
    }
    return result;
  } catch (e) {
    Logger.log("Error getHRCalendarEvents: " + e.message);
    throw new Error("Gagal mengambil data HR Calendar: " + e.message);
  }
}

function simpanHRCalendarEvent(ev) {
  if (!ev || typeof ev !== 'object') {
    throw new Error("Data kegiatan tidak valid.");
  }
  if (!ev.tanggal || !ev.namaKegiatan || !ev.dept || !ev.pic) {
    throw new Error("Tanggal, Nama Kegiatan, Departemen, dan PIC wajib diisi!");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error("Sistem sedang sibuk menyimpan data lain, coba lagi beberapa saat.");
  }

  try {
    const sheet = initSheetHRCalendar();
    const data = sheet.getDataRange().getValues();

    let isUpdated = false;
    const evIdStr = String(ev.id).trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === evIdStr) {
        sheet.getRange(i + 1, 2, 1, 16).setValues([[
          String(ev.tanggal || ""),
          String(ev.namaKegiatan || ""),
          String(ev.dept || ""),
          String(ev.pic || ""),
          String(ev.lokasi || ""),
          String(ev.status || "Planning"),
          String(ev.keterangan || ""),
          String(ev.jamMulai || ""),
          String(ev.jamSelesai || ""),
          String(ev.jenisKegiatan || ""),
          String(ev.prioritas || ""),
          String(ev.catatan || ""),
          String(ev.dibuatOleh || ""),
          String(ev.lampiran || ""),
          String(ev.tanggalDibuat || ""),
          String(ev.waktuUpdate || "")
        ]]);
        isUpdated = true;
        break;
      }
    }

    if (!isUpdated) {
      sheet.appendRow([
        evIdStr,
        String(ev.tanggal || ""),
        String(ev.namaKegiatan || ""),
        String(ev.dept || ""),
        String(ev.pic || ""),
        String(ev.lokasi || ""),
        String(ev.status || "Planning"),
        String(ev.keterangan || ""),
        String(ev.jamMulai || ""),
        String(ev.jamSelesai || ""),
        String(ev.jenisKegiatan || ""),
        String(ev.prioritas || ""),
        String(ev.catatan || ""),
        String(ev.dibuatOleh || ""),
        String(ev.lampiran || ""),
        String(ev.tanggalDibuat || ""),
        String(ev.waktuUpdate || "")
      ]);
    }

    return "Kegiatan HR Calendar berhasil disimpan.";
  } catch (e) {
    Logger.log("Error simpanHRCalendarEvent: " + e.message);
    throw new Error("Gagal menyimpan kegiatan HR Calendar: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function hapusHRCalendarEvent(id) {
  if (!id) throw new Error("ID kegiatan tidak boleh kosong.");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error("Sistem sedang sibuk, coba lagi beberapa saat.");
  }

  try {
    const sheet = initSheetHRCalendar();
    const data = sheet.getDataRange().getValues();
    const targetIdStr = String(id).trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetIdStr) {
        sheet.deleteRow(i + 1);
        return "Kegiatan HR Calendar berhasil dihapus.";
      }
    }
    throw new Error("ID kegiatan tidak ditemukan.");
  } catch (e) {
    Logger.log("Error hapusHRCalendarEvent: " + e.message);
    throw new Error("Gagal menghapus kegiatan HR Calendar: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

const DRIVE_FOLDER_NAME_KALENDER = "HR_Calendar_Lampiran";

function getOrCreateLampiranFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME_KALENDER);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME_KALENDER);
}

function uploadFileLampiranKalender(base64Data, fileName, mimeType) {
  try {
    if (!base64Data) throw new Error("Data file kosong.");
    const folder = getOrCreateLampiranFolder_();
    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType || "application/octet-stream", fileName || "lampiran");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log("Error uploadFileLampiranKalender: " + e.message);
    throw new Error("Gagal mengunggah lampiran ke Google Drive: " + e.message);
  }
}
