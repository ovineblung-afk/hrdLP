/**
 * =========================================================
 * FILE: Summary.gs
 * FUNGSI: Menangani penyimpanan JSON keseluruhan Dashboard 
 * ke Sheet "HRLP_Riwayat"
 * =========================================================
 */

function getLabelPeriode_(bulanVal, tahunVal) {
  const bNum = Number(bulanVal);
  if (!isNaN(bNum) && bNum >= 1 && bNum <= 12) {
    const namaBulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
    return namaBulan[bNum - 1] + " " + tahunVal;
  }
  return String(bulanVal).trim() + " " + tahunVal;
}


// Setup awal untuk Sheet Riwayat Laporan
function initSheetRiwayat() {
  const ss = getDB();
  let sheet = ss.getSheetByName("HRLP_Riwayat");
  
  if (!sheet) {
    sheet = ss.insertSheet("HRLP_Riwayat");
    // Header disesuaikan dengan struktur CSV Anda
    sheet.appendRow([
      "Periode", "Timestamp", "Rev_Actual", "Rev_Budget",
      "FO_Payroll", "FO_DW", "FO_Casual", "FO_Meals", "FO_Arrivals",
      "HK_Payroll", "HK_DW", "HK_Casual", "HK_Meals", "HK_Occupied",
      "FBS_Payroll", "FBS_DW", "FBS_Casual", "FBS_Meals", "FBS_Cover",
      "FBP_Payroll", "FBP_DW", "FBP_Casual", "FBP_Meals", "FBP_FoodCover",
      "ENG_Payroll", "ENG_DW", "ENG_Casual", "ENG_Meals",
      "HRD_Payroll", "HRD_DW", "HRD_Casual", "HRD_Meals", "HRD_Training", "HRD_TotalEmp",
      "Data_JSON"
    ]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Fungsi utama yang dipanggil oleh Dashboard
function simpanDataLaporan(appData) {
  // LockService: mencegah race condition kalau ada 2 sesi menyimpan periode yang
  // sama nyaris bersamaan (baca rowIdx lama lalu saling menimpa baris).
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    throw new Error("Server sedang sibuk menyimpan data lain, silakan coba lagi.");
  }
  try {
    const periode = appData.periode;
    if (!periode) throw new Error("Periode tidak ditemukan!");

    // Cek apakah periode target sudah CLOSED di MonthlyHistory
    const ss = getDB();
    const sheetHistory = ss.getSheetByName("MonthlyHistory");
    if (sheetHistory) {
      const dataHistory = sheetHistory.getDataRange().getValues();
      for (let i = 1; i < dataHistory.length; i++) {
        const labelPeriode = getLabelPeriode_(dataHistory[i][1], dataHistory[i][2]); // Bulan Tahun
        if (labelPeriode.trim() === String(periode).trim() && String(dataHistory[i][7]) === "CLOSED") {
          throw new Error("Periode " + periode + " telah CLOSED dan dikunci. Data tidak dapat diperbarui.");
        }
      }
    }

    const sheet = initSheetRiwayat();
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(periode).trim()) {
        rowIdx = i + 1;
        break;
      }
    }

    const rowData = [
      periode, new Date(), appData.umum?.totRevAct || 0, appData.umum?.totRevBud || 0,
      appData.fo?.payTot || 0, appData.fo?.qtyDW || 0, appData.fo?.qtyCas || 0, 0, appData.fo?.arrAct || 0,
      appData.hk?.payTot || 0, appData.hk?.qtyDW || 0, appData.hk?.qtyCas || 0, 0, appData.hk?.occRoom || 0,
      appData.fbs?.payTot || 0, appData.fbs?.qtyDW || 0, appData.fbs?.qtyCas || 0, 0, appData.fbs?.cover || 0,
      appData.fbp?.payTot || 0, appData.fbp?.qtyDW || 0, appData.fbp?.qtyCas || 0, 0, appData.fbp?.coverFood || 0,
      appData.eng?.payTot || 0, appData.eng?.qtyDW || 0, appData.eng?.qtyCas || 0, 0,
      appData.hrd?.payTot || 0, appData.hrd?.qtyDW || 0, appData.hrd?.qtyCas || 0, appData.hrd?.mealsAct || 0, appData.training?.totBud || 0, 0,
      JSON.stringify(appData)
    ];
    
    if (rowIdx !== -1) {
      // Overwrite existing row for this period
      sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // Append new row if not found
      sheet.appendRow(rowData);
    }
    return "Data berhasil disimpan.";
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

// Mengambil seluruh riwayat laporan dengan menggabungkan MonthlyHistory (CLOSED) dan HRLP_Riwayat (OPEN)
function getRiwayatLengkap() {
  try {
    const riwayatArray = [];
    const seenPeriods = {};

    // 1. Ambil data dari MonthlyHistory (Closed Snapshots)
    const ss = getDB();
    const sheetHistory = ss.getSheetByName("MonthlyHistory");
    if (sheetHistory) {
      const dataRange = sheetHistory.getDataRange();
      const displayValues = dataRange.getDisplayValues();
      const rawValues = dataRange.getValues();
      if (displayValues.length > 1) {
        for (let i = 1; i < displayValues.length; i++) {
          const rowDisplay = displayValues[i];
          const rowRaw = rawValues[i];
          const status = rowDisplay[7] || "CLOSED";
          const labelPeriode = getLabelPeriode_(rowRaw[1], rowRaw[2]); // e.g. "Juli 2026"
          
          // Coba cari kolom mana saja yang berisi JSON string valid (mulai dari kolom terakhir)
          let jsonString = "";
          for (let colIdx = rowRaw.length - 1; colIdx >= 0; colIdx--) {
            const valStr = String(rowRaw[colIdx] || "").trim();
            if (valStr.startsWith("{") && valStr.endsWith("}")) {
              jsonString = valStr;
              break;
            }
          }
          
          // Jika tidak ditemukan JSON, buat fallback JSON dari data tabular baris ini
          if (!jsonString) {
            const fallbackObj = {
              status: status,
              periode: labelPeriode,
              catatan: rowRaw[27] || "",
              umum: {
                totRevAct: parseFloat(rowRaw[17]) || 0,
                totRevBud: parseFloat(rowRaw[26]) || 0,
                revEmpAct: parseFloat(rowRaw[22]) || 0,
                revEmpBud: 0
              },
              fo: {
                payTot: 0,
                qtyDW: 0,
                qtyCas: 0,
                arrAct: 0,
                arrBud: 0,
                revRmAct: 0,
                revRmBud: 0
              },
              fnb: {
                payTot: 0,
                qtyDW: 0,
                qtyCas: 0,
                revFnbAct: 0,
                revFnbBud: 0,
                apMtdAct: 0,
                apMtdBud: 0
              },
              hrd: {
                payTot: parseFloat(rowRaw[18]) || 0,
                mealsAct: parseFloat(rowRaw[19]) || 0,
                mealsBud: 0,
                prodAct: parseFloat(rowRaw[24]) || 0,
                prodBud: 0,
                qtyKTNKT: parseFloat(rowRaw[15]) || 0,
                qtyDW: 0,
                qtyCas: 0,
                qtyOut: 0
              },
              kpiPter: {
                pterMtdAct: parseFloat(rowRaw[23]) || 0,
                pterMtdFcst: parseFloat(rowRaw[25]) || 0,
                pterMtdBud: 0
              }
            };
            jsonString = JSON.stringify(fallbackObj);
          }
          
          riwayatArray.push({
            tanggalLengkap: rowDisplay[3] + " " + rowDisplay[4], // Tanggal Closing + Jam Closing
            periode: labelPeriode,
            totRevAct: parseFloat(rowRaw[17]) || 0,
            rawData: jsonString,
            status: status
          });
          seenPeriods[labelPeriode] = true;
        }
      }
    }

    // 2. Ambil data dari HRLP_Riwayat (Open/Active Reports)
    const sheet = initSheetRiwayat();
    const dataRange = sheet.getDataRange();
    const displayValues = dataRange.getDisplayValues();
    const rawValues = dataRange.getValues();
    if (displayValues.length > 1) {
      for (let i = 1; i < displayValues.length; i++) {
        const rowDisplay = displayValues[i];
        const rowRaw = rawValues[i];
        const labelPeriode = rowDisplay[0] || "N/A";
        // Hanya tambahkan jika periode tersebut belum di-Closing (belum ada di seenPeriods)
        if (!seenPeriods[labelPeriode]) {
          const jsonString = rowRaw[rowRaw.length - 1]; // Asumsi JSON ada di kolom terakhir - gunakan raw value agar tidak terpotong
          let parsed = {};
          try { parsed = JSON.parse(jsonString); } catch(e) {}
          riwayatArray.push({
            tanggalLengkap: rowDisplay[1] || "N/A",
            periode: labelPeriode,
            totRevAct: parseFloat(rowRaw[2]) || 0,
            rawData: jsonString,
            status: parsed.status || "OPEN"
          });
        }
      }
    }

    return riwayatArray;
  } catch (e) {
    throw new Error("Gagal memuat riwayat: " + e.message);
  }
}

// Menghapus data riwayat dengan perlindungan untuk periode CLOSED
function hapusRiwayat(periodeTarget, timestampTarget) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    throw new Error("Server sedang sibuk, silakan coba lagi.");
  }
  try {
    // Cek apakah periodeTarget sudah CLOSED di MonthlyHistory
    const ss = getDB();
    const sheetHistory = ss.getSheetByName("MonthlyHistory");
    if (sheetHistory) {
      const dataHistory = sheetHistory.getDataRange().getValues();
      for (let i = 1; i < dataHistory.length; i++) {
        const labelPeriode = getLabelPeriode_(dataHistory[i][1], dataHistory[i][2]); // Bulan Tahun
        if (labelPeriode.trim() === String(periodeTarget).trim() && String(dataHistory[i][7]) === "CLOSED") {
          throw new Error("Periode " + periodeTarget + " telah CLOSED dan dikunci. Data historis tidak dapat dihapus.");
        }
      }
    }

    const sheet = initSheetRiwayat();
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      const matchPeriode = String(data[i][0]).trim() === String(periodeTarget).trim();
      const matchTimestamp = timestampTarget ? (String(data[i][1]).trim() === String(timestampTarget).trim()) : true;
      if (matchPeriode && matchTimestamp) {
        sheet.deleteRow(i + 1);
        return "Riwayat berhasil dihapus.";
      }
    }
    throw new Error("Data tidak ditemukan.");
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}