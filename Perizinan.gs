/**
 * =========================================================
 * FILE: Perizinan.gs
 * FUNGSI: CRUD Data Perizinan Hotel (Google Apps Script)
 * =========================================================
 */

var DRIVE_FOLDER_NAME_PERIZINAN = "HR_Perizinan_Dokumen";

function getOrCreatePerizinanFolder_() {
  if (typeof DriveApp === 'undefined') return null;
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME_PERIZINAN);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME_PERIZINAN);
}

function uploadFilePerizinanDrive(base64Data, fileName, mimeType, isAdmin) {
  if (!isAdmin) {
    throw new Error("Akses ditolak: Hanya Admin yang dapat mengunggah dokumen perizinan.");
  }
  try {
    if (!base64Data) throw new Error("Data file kosong.");
    if (typeof DriveApp === 'undefined') {
      return {
        fileId: 'FILE-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
        fileUrl: 'https://drive.google.com/file/d/mock_' + Date.now() + '/view',
        fileName: fileName || 'dokumen'
      };
    }
    var folder = getOrCreatePerizinanFolder_();
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || "application/octet-stream", fileName || "dokumen_perizinan");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      fileName: fileName || file.getName()
    };
  } catch (e) {
    Logger.log("Error uploadFilePerizinanDrive: " + e.message);
    throw new Error("Gagal mengunggah dokumen perizinan ke Google Drive: " + e.message);
  }
}

function formatDateCell_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return val.toString();
}

function getPerizinan(isAdmin) {
  try {
    var ss = getDB();
    var sheet = ss.getSheetByName('HRLP_Perizinan');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var headers = data[0].map(function(h) { return h ? h.toString().trim() : ''; });
    var biayaColIdx = headers.indexOf('Perkiraan Biaya (Rp)');

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        var biayaVal = 0;
        var ketVal = '';
        var fileNameVal = '';
        var fileIdVal = '';
        var fileUrlVal = '';
        var tglInputVal = '';
        var lastUpdatedVal = '';
        var updatedByVal = '';

        if (biayaColIdx > -1) {
          biayaVal = parseFloat(row[biayaColIdx]) || 0;
          ketVal = row[biayaColIdx + 1] ? row[biayaColIdx + 1].toString() : '';
          fileNameVal = row[biayaColIdx + 2] ? row[biayaColIdx + 2].toString() : '';
          fileIdVal = row[biayaColIdx + 3] ? row[biayaColIdx + 3].toString() : '';
          fileUrlVal = row[biayaColIdx + 4] ? row[biayaColIdx + 4].toString() : '';
          tglInputVal = row[biayaColIdx + 5] ? row[biayaColIdx + 5].toString() : '';
          lastUpdatedVal = row[biayaColIdx + 6] ? row[biayaColIdx + 6].toString() : '';
          updatedByVal = row[biayaColIdx + 7] ? row[biayaColIdx + 7].toString() : '';
        } else {
          // Backward compatibility for old sheet format
          biayaVal = 0;
          ketVal = row[8] ? row[8].toString() : '';
          fileNameVal = row[9] ? row[9].toString() : '';
          fileIdVal = row[10] ? row[10].toString() : '';
          fileUrlVal = row[11] ? row[11].toString() : '';
          tglInputVal = row[12] ? row[12].toString() : '';
          lastUpdatedVal = row[13] ? row[13].toString() : '';
          updatedByVal = row[14] ? row[14].toString() : '';
        }

        // Parse multi-files array
        var filesList = [];
        if (fileUrlVal) {
          try {
            if (fileUrlVal.trim().indexOf('[') === 0) {
              filesList = JSON.parse(fileUrlVal);
            }
          } catch(e) {}
        }

        if (!filesList || filesList.length === 0) {
          if (fileUrlVal || fileNameVal) {
            var names = fileNameVal ? fileNameVal.split(' | ') : [];
            var ids = fileIdVal ? fileIdVal.split(' | ') : [];
            var urls = fileUrlVal ? fileUrlVal.split(' | ') : [];
            var maxLen = Math.max(names.length, ids.length, urls.length);
            for (var k = 0; k < maxLen; k++) {
              if (urls[k] || names[k]) {
                filesList.push({
                  fileName: names[k] || 'Dokumen',
                  fileId: ids[k] || ('FILE-' + k),
                  fileUrl: urls[k] || ''
                });
              }
            }
          }
        }

        // Access Control: Non-admin users do NOT receive document files/URLs
        if (!isAdmin) {
          fileNameVal = '';
          fileIdVal = '';
          fileUrlVal = '';
          filesList = [];
        }

        result.push({
          id: row[0] ? row[0].toString() : '',
          namaPerizinan: row[1] ? row[1].toString() : '',
          instansi: row[2] ? row[2].toString() : '',
          nomorDokumen: row[3] ? row[3].toString() : '',
          tanggalMulai: formatDateCell_(row[4]),
          tanggalAkhir: formatDateCell_(row[5]),
          dueDate: formatDateCell_(row[6]),
          pic: row[7] ? row[7].toString() : '',
          biaya: biayaVal,
          keterangan: ketVal,
          fileName: fileNameVal,
          fileId: fileIdVal,
          fileUrl: fileUrlVal,
          files: filesList,
          tanggalInput: tglInputVal,
          lastUpdated: lastUpdatedVal,
          updatedBy: updatedByVal
        });
      }
    }
    return result;
  } catch (err) {
    Logger.log('Error getPerizinan: ' + err.message);
    return [];
  }
}

function simpanPerizinan(p, isAdmin) {
  if (!isAdmin) {
    throw new Error('Akses ditolak: Hanya Admin yang dapat menyimpan data perizinan.');
  }
  try {
    var ss = getDB();
    var sheet = ss.getSheetByName('HRLP_Perizinan');
    var expectedHeaders = [
      'ID Perizinan', 'Nama Perizinan', 'Instansi', 'Nomor Dokumen',
      'Tanggal Mulai', 'Tanggal Akhir', 'Due Date', 'PIC', 'Perkiraan Biaya (Rp)', 'Keterangan',
      'Nama File', 'File ID', 'File URL', 'Tanggal Input', 'Last Updated', 'Updated By'
    ];

    if (!sheet) {
      sheet = ss.insertSheet('HRLP_Perizinan');
      sheet.appendRow(expectedHeaders);
    } else {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.indexOf('Perkiraan Biaya (Rp)') === -1) {
        sheet.insertColumnAfter(8);
        sheet.getRange(1, 9).setValue('Perkiraan Biaya (Rp)');
      }
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var existingFileName = '';
    var existingFileId = '';
    var existingFileUrl = '';

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === p.id.toString()) {
        rowIndex = i + 1;
        var headers = data[0].map(function(h) { return h ? h.toString().trim() : ''; });
        var biayaColIdx = headers.indexOf('Perkiraan Biaya (Rp)');
        if (biayaColIdx > -1) {
          existingFileName = data[i][biayaColIdx + 2] ? data[i][biayaColIdx + 2].toString() : '';
          existingFileId = data[i][biayaColIdx + 3] ? data[i][biayaColIdx + 3].toString() : '';
          existingFileUrl = data[i][biayaColIdx + 4] ? data[i][biayaColIdx + 4].toString() : '';
        }
        if (!p.tanggalMulai && data[i][4]) {
          p.tanggalMulai = formatDateCell_(data[i][4]);
        }
        if (!p.tanggalAkhir && data[i][5] && data[i][5].toString().toLowerCase().indexOf('seumur') > -1) {
          p.tanggalAkhir = 'Seumur Hidup';
        }
        if (!p.dueDate && data[i][6]) {
          p.dueDate = formatDateCell_(data[i][6]);
        }
        break;
      }
    }

    var biayaNum = parseFloat(p.biaya) || 0;

    var filesArr = p.files || [];
    var fileNameStr = '';
    var fileIdStr = '';
    var fileUrlStr = '';

    if (filesArr.length > 0) {
      fileNameStr = filesArr.map(function(f) { return f.fileName || ''; }).join(' | ');
      fileIdStr = filesArr.map(function(f) { return f.fileId || ''; }).join(' | ');
      fileUrlStr = JSON.stringify(filesArr.map(function(f) {
        return {
          fileName: f.fileName || '',
          fileId: f.fileId || '',
          fileUrl: f.fileUrl || ''
        };
      }));
    } else if (p.fileUrl) {
      fileNameStr = p.fileName || '';
      fileIdStr = p.fileId || '';
      fileUrlStr = p.fileUrl || '';
    } else if (rowIndex > -1) {
      // Pertahankan dokumen yang sudah ada jika tidak diubah
      fileNameStr = existingFileName;
      fileIdStr = existingFileId;
      fileUrlStr = existingFileUrl;
    }

    var rowValues = [
      p.id, p.namaPerizinan || '', p.instansi || '', p.nomorDokumen || '',
      p.tanggalMulai || '', p.tanggalAkhir || '', p.dueDate || '',
      p.pic || '', biayaNum, p.keterangan || '', fileNameStr, fileIdStr,
      fileUrlStr, p.tanggalInput || '', p.lastUpdated || '', p.updatedBy || ''
    ];

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
    return "Berhasil menyimpan perizinan.";
  } catch (err) {
    Logger.log('Error simpanPerizinan: ' + err.message);
    throw new Error('Gagal menyimpan perizinan: ' + err.message);
  }
}

function hapusPerizinan(id, isAdmin) {
  if (!isAdmin) {
    throw new Error('Akses ditolak: Hanya Admin yang dapat menghapus data perizinan.');
  }
  try {
    if (!id) throw new Error('ID Perizinan tidak valid.');
    var ss = getDB();
    var sheet = ss.getSheetByName('HRLP_Perizinan');
    if (!sheet) throw new Error('Sheet HRLP_Perizinan tidak ditemukan.');

    var data = sheet.getDataRange().getValues();
    var deleted = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }
    if (!deleted) throw new Error('Record dengan ID ' + id + ' tidak ditemukan di database.');
    return "Berhasil menghapus perizinan.";
  } catch (err) {
    Logger.log('Error hapusPerizinan: ' + err.message);
    throw new Error('Gagal menghapus data perizinan: ' + err.message);
  }
}

function generateIdPerizinan(sheet) {
  var data = sheet.getDataRange().getValues();
  var maxIdNum = 0;
  for (var i = 1; i < data.length; i++) {
    var idStr = data[i][0] ? data[i][0].toString().trim() : '';
    if (idStr.indexOf('PER-') === 0) {
      var num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  }
  var nextNum = maxIdNum + 1;
  var nextNumStr = nextNum.toString();
  while (nextNumStr.length < 6) {
    nextNumStr = '0' + nextNumStr;
  }
  return 'PER-' + nextNumStr;
}

function validatePerizinanImport(row) {
  var errors = [];
  if (!row.namaPerizinan) errors.push('Nama Perizinan kosong');
  if (!row.instansi) errors.push('Dinas / Instansi Terkait kosong');
  if (!row.nomorDokumen) errors.push('Nomor Dokumen kosong');
  if (!row.tanggalMulai) errors.push('Tanggal Berlaku Mulai kosong');
  if (!row.dueDate) errors.push('Due Date kosong');
  if (!row.pic) errors.push('PIC kosong');
  return errors;
}

function importPerizinanExcel(importDataList, isAdmin) {
  if (!isAdmin) {
    return { success: false, message: "Akses ditolak. Fitur hanya tersedia untuk Admin." };
  }
  
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Server sedang sibuk. Silakan coba beberapa saat lagi.');
  }
  
  try {
    var ss = getDB();
    var sheet = ss.getSheetByName('HRLP_Perizinan');
    if (!sheet) {
      sheet = ss.insertSheet('HRLP_Perizinan');
      sheet.appendRow([
        'ID Perizinan', 'Nama Perizinan', 'Instansi', 'Nomor Dokumen',
        'Tanggal Mulai', 'Tanggal Akhir', 'Due Date', 'PIC', 'Keterangan',
        'Nama File', 'File ID', 'File URL', 'Tanggal Input', 'Last Updated', 'Updated By'
      ]);
    }
    
    var existingData = sheet.getDataRange().getValues();
    var existingDocs = {};
    for (var i = 1; i < existingData.length; i++) {
      var docNum = existingData[i][3] ? existingData[i][3].toString().trim().toLowerCase() : '';
      if (docNum) {
        existingDocs[docNum] = true;
      }
    }
    
    var result = {
      success: true,
      total: importDataList.length,
      imported: 0,
      duplicate: 0,
      invalid: 0,
      errors: []
    };
    
    var rowsToAppend = [];
    var currentId = generateIdPerizinan(sheet);
    var idNum = parseInt(currentId.substring(4), 10);
    
    var nowStr = new Date().toISOString().substring(0, 10);
    
    for (var j = 0; j < importDataList.length; j++) {
      var p = importDataList[j];
      var docKey = p.nomorDokumen ? p.nomorDokumen.toString().trim().toLowerCase() : '';
      
      var validationErrors = validatePerizinanImport(p);
      if (validationErrors.length > 0) {
        result.invalid++;
        result.errors.push({
          row: j + 2,
          nama: p.namaPerizinan || 'Tanpa Nama',
          masalah: validationErrors.join(', ')
        });
        continue;
      }
      
      if (docKey && existingDocs[docKey]) {
        result.duplicate++;
        continue;
      }
      
      var nextNumStr = idNum.toString();
      while (nextNumStr.length < 6) {
        nextNumStr = '0' + nextNumStr;
      }
      var newId = 'PER-' + nextNumStr;
      idNum++;
      
      rowsToAppend.push([
        newId, 
        p.namaPerizinan || '', 
        p.instansi || '', 
        p.nomorDokumen || '',
        p.tanggalMulai || '', 
        p.tanggalAkhir || '', 
        p.dueDate || '',
        p.pic || '', 
        p.keterangan || '', 
        p.fileName || '', 
        p.fileId || '',
        p.fileUrl || '', 
        nowStr, 
        nowStr, 
        'Admin Import'
      ]);
      
      if (docKey) existingDocs[docKey] = true;
      result.imported++;
    }
    
    if (rowsToAppend.length > 0) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }
    
    return result;
  } catch (err) {
    Logger.log('Error importPerizinanExcel: ' + err.message);
    throw new Error('Gagal import perizinan: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

// Access Control Helpers for Document Access (Rule 5)
function getDokumenPerizinan(id, isAdmin) {
  if (!isAdmin) {
    throw new Error("Akses ditolak: Dokumen perizinan hanya dapat diakses dalam Mode Admin.");
  }
  var all = getPerizinan(true);
  var item = all.find(function(x) { return String(x.id) === String(id); });
  if (!item) throw new Error("Data perizinan tidak ditemukan.");
  return item.files && item.files.length > 0 ? item.files : (item.fileUrl ? [{ fileName: item.fileName, fileId: item.fileId, fileUrl: item.fileUrl }] : []);
}

function viewDokumenPerizinan(id, isAdmin) {
  return getDokumenPerizinan(id, isAdmin);
}

function getFilePerizinan(id, isAdmin) {
  return getDokumenPerizinan(id, isAdmin);
}

function uploadDokumenPerizinan(base64Data, fileName, mimeType, isAdmin) {
  return uploadFilePerizinanDrive(base64Data, fileName, mimeType, isAdmin);
}

function deleteDokumenPerizinan(id, isAdmin) {
  return hapusPerizinan(id, isAdmin);
}
