/**
 * =====================================================================
 * FILE: CorporateKnowledge.gs
 * MODUL: Corporate Knowledge Management System
 * INTEGRASI: Google Sheets (Metadata Database) & Google Drive (File Storage)
 * 
 * Fitur:
 * 1. Manajemen Kategori (17 Kategori Default, Status Active/Inactive, Hitungan Aktual Dokumen)
 * 2. Upload Dokumen & Bulk Upload (Auto Folder 'Corporate Knowledge / [Kategori]')
 * 3. Nomor Dokumen Otomatis (001, 002, 003... Monotonik Unik tidak didaur ulang)
 * 4. Document ID Unik Internal (CK-DOC-XXXXXX)
 * 5. Versioning (v1.0 -> v1.1 -> v2.0 dengan riwayat lengkap di CK_Versions)
 * 6. Keamanan Akses: Karyawan = View-Only tanpa akses Download & tanpa Direct Drive URL
 * =====================================================================
 */

const CK_DEFAULT_CATEGORIES = [
  ['cat-sop-operasional', 'SOP Operasional', 'fa-clipboard-list', 'sky', 'ACTIVE', 'Standar Operasional Prosedur seluruh divisi operasional hotel', 1],
  ['cat-kebijakan-regulasi', 'Kebijakan & Regulasi', 'fa-scale-balanced', 'indigo', 'ACTIVE', 'Kebijakan resmi manajemen dan kepatuhan regulasi eksternal', 2],
  ['cat-formulir-template', 'Formulir & Template', 'fa-file-lines', 'emerald', 'ACTIVE', 'Template dokumen, form permohonan, dan formulir standar', 3],
  ['cat-santika-management-system', 'Santika Management System', 'fa-hotel', 'blue', 'ACTIVE', 'Sistem manajemen mutu dan standar operasional Santika', 4],
  ['cat-indonesian-home', 'Indonesian Home', 'fa-house-chimney-user', 'amber', 'ACTIVE', 'Konsep keramahtamahan khas budaya Indonesia (Hospitality Values)', 5],
  ['cat-hr-policy', 'HR Policy', 'fa-user-shield', 'purple', 'ACTIVE', 'Ketentuan SDM, etika kerja, hak & kewajiban karyawan', 6],
  ['cat-training-material', 'Training Material', 'fa-graduation-cap', 'teal', 'ACTIVE', 'Modul pelatihan, materi orientasi, dan pengembangan kompetensi', 7],
  ['cat-health-safety', 'Health & Safety', 'fa-heart-pulse', 'rose', 'ACTIVE', 'Keselamatan, kesehatan kerja, sanitasi, dan mitigasi darurat', 8],
  ['cat-peraturan-perusahaan', 'Peraturan Perusahaan', 'fa-book-bookmark', 'violet', 'ACTIVE', 'Buku peraturan perusahaan dan pedoman disiplin kerja', 9],
  ['cat-engineering', 'Engineering', 'fa-wrench', 'orange', 'ACTIVE', 'Pedoman pemeliharaan teknis gedung, genset, chiller, dan ME', 10],
  ['cat-front-office', 'Front Office', 'fa-bell-concierge', 'cyan', 'ACTIVE', 'SOP layanan tamu, reservasi, check-in/out, dan handling keluhan', 11],
  ['cat-housekeeping', 'Housekeeping', 'fa-broom', 'lime', 'ACTIVE', 'Standar kebersihan kamar, public area, laundry, dan linen', 12],
  ['cat-food-beverage', 'Food & Beverage', 'fa-utensils', 'amber', 'ACTIVE', 'Standar penyajian makanan, resep standar, sanitasi FB, dan bar', 13],
  ['cat-finance', 'Finance', 'fa-coins', 'emerald', 'ACTIVE', 'Prosedur keuangan, petty cash, purchasing, invoicing, dan audit', 14],
  ['cat-sales-marketing', 'Sales & Marketing', 'fa-bullhorn', 'pink', 'ACTIVE', 'Materi promosi, rate structure, paket event, dan marketing collateral', 15],
  ['cat-hom', 'HOM', 'fa-briefcase', 'slate', 'ACTIVE', 'Head of Maintenance / Hotel Operations Management guidelines', 16],
  ['cat-lainnya', 'Lainnya', 'fa-folder-open', 'gray', 'ACTIVE', 'Dokumen referensi dan pengetahuan umum perusahaan', 17]
];

function initCKSheets() {
  const ss = getDB();
  
  // 1. Sheet Categories
  let sheetCat = ss.getSheetByName('CK_Categories');
  if (!sheetCat) {
    sheetCat = ss.insertSheet('CK_Categories');
    sheetCat.appendRow(['ID', 'Nama', 'Icon', 'Warna', 'Status', 'Deskripsi', 'Urutan', 'Created_At']);
    CK_DEFAULT_CATEGORIES.forEach(c => {
      sheetCat.appendRow([c[0], c[1], c[2], c[3], c[4], c[5], c[6], new Date().toISOString()]);
    });
  }

  // 2. Sheet Documents
  let sheetDoc = ss.getSheetByName('CK_Documents');
  if (!sheetDoc) {
    sheetDoc = ss.insertSheet('CK_Documents');
    sheetDoc.appendRow([
      'Document_ID', 'Nomor', 'Nama_Dokumen', 'Kategori_ID', 'Kategori', 
      'File_ID', 'File_Name', 'File_Type', 'File_Size', 'Versi', 
      'Status', 'Tanggal_Upload', 'Uploaded_By', 'Deskripsi', 'Raw_Versions_JSON'
    ]);
  }

  // 3. Sheet Config (Sequence Counter)
  let sheetConfig = ss.getSheetByName('CK_Config');
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet('CK_Config');
    sheetConfig.appendRow(['Key', 'Value', 'Updated_At']);
    sheetConfig.appendRow(['NEXT_DOC_SEQUENCE', 1, new Date().toISOString()]);
  }

  return { ss, sheetCat, sheetDoc, sheetConfig };
}

function getCKNextDocNumber() {
  const { sheetConfig } = initCKSheets();
  const data = sheetConfig.getDataRange().getValues();
  let rowIdx = -1;
  let nextSeq = 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'NEXT_DOC_SEQUENCE') {
      rowIdx = i + 1;
      nextSeq = parseInt(data[i][1]) || 1;
      break;
    }
  }

  if (rowIdx === -1) {
    sheetConfig.appendRow(['NEXT_DOC_SEQUENCE', 2, new Date().toISOString()]);
  } else {
    sheetConfig.getRange(rowIdx, 2).setValue(nextSeq + 1);
    sheetConfig.getRange(rowIdx, 3).setValue(new Date().toISOString());
  }

  const padded = String(nextSeq).padStart(3, '0');
  return padded;
}

function getOrCreateCKDriveFolder(categoryName) {
  let rootFolder;
  const rootName = 'Corporate Knowledge';
  const rootIter = DriveApp.getFoldersByName(rootName);
  if (rootIter.hasNext()) {
    rootFolder = rootIter.next();
  } else {
    rootFolder = DriveApp.createFolder(rootName);
  }

  const catFolderName = (categoryName || 'Lainnya').trim();
  const subIter = rootFolder.getFoldersByName(catFolderName);
  if (subIter.hasNext()) {
    return subIter.next();
  } else {
    return rootFolder.createFolder(catFolderName);
  }
}

// ==========================================
// PUBLIC GAS API FUNCTIONS
// ==========================================

function getCKCategories(isAdmin) {
  const { sheetCat, sheetDoc } = initCKSheets();
  const catRows = sheetCat.getDataRange().getValues();
  const docRows = sheetDoc.getDataRange().getValues();

  // Calculate actual doc counts per category
  const docCountMap = {};
  for (let i = 1; i < docRows.length; i++) {
    const row = docRows[i];
    const catId = row[3];
    const catName = row[4];
    const status = row[10];
    if (status !== 'DELETED') {
      if (catId) docCountMap[catId] = (docCountMap[catId] || 0) + 1;
      if (catName) docCountMap[catName] = (docCountMap[catName] || 0) + 1;
    }
  }

  const result = [];
  for (let i = 1; i < catRows.length; i++) {
    const row = catRows[i];
    const id = row[0];
    const nama = row[1];
    const status = row[4] || 'ACTIVE';

    if (!isAdmin && status !== 'ACTIVE') {
      continue; // Karyawan tidak dapat melihat kategori INACTIVE
    }

    result.push({
      id: id,
      nama: nama,
      icon: row[2] || 'fa-folder',
      warna: row[3] || 'sky',
      status: status,
      deskripsi: row[5] || '',
      urutan: parseInt(row[6]) || i,
      docCount: docCountMap[id] || docCountMap[nama] || 0
    });
  }

  return result;
}

function saveCKCategory(catObj, isAdmin) {
  if (!isAdmin) throw new Error('Akses ditolak: Hanya Admin yang dapat mengelola kategori.');
  const { sheetCat } = initCKSheets();
  const data = sheetCat.getDataRange().getValues();
  const id = catObj.id || ('cat-' + catObj.nama.toLowerCase().replace(/[^a-z0-9]+/g, '-'));

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > -1) {
    sheetCat.getRange(foundRow, 2).setValue(catObj.nama);
    sheetCat.getRange(foundRow, 3).setValue(catObj.icon || 'fa-folder');
    sheetCat.getRange(foundRow, 4).setValue(catObj.warna || 'sky');
    sheetCat.getRange(foundRow, 5).setValue(catObj.status || 'ACTIVE');
    sheetCat.getRange(foundRow, 6).setValue(catObj.deskripsi || '');
  } else {
    sheetCat.appendRow([
      id,
      catObj.nama,
      catObj.icon || 'fa-folder',
      catObj.warna || 'sky',
      catObj.status || 'ACTIVE',
      catObj.deskripsi || '',
      data.length,
      new Date().toISOString()
    ]);
  }

  return getCKCategories(true);
}

function getCKDocuments(filters, isAdmin) {
  const { sheetDoc, sheetCat } = initCKSheets();
  const docRows = sheetDoc.getDataRange().getValues();
  const catRows = sheetCat.getDataRange().getValues();

  // Active categories map
  const activeCats = {};
  for (let i = 1; i < catRows.length; i++) {
    if (catRows[i][4] === 'ACTIVE') {
      activeCats[catRows[i][0]] = true;
      activeCats[catRows[i][1]] = true;
    }
  }

  const result = [];
  for (let i = 1; i < docRows.length; i++) {
    const row = docRows[i];
    const docId = row[0];
    const catId = row[3];
    const catName = row[4];
    const status = row[10] || 'ACTIVE';

    if (status === 'DELETED') continue;

    // Filter inactive category for employee
    if (!isAdmin) {
      if (!activeCats[catId] && !activeCats[catName]) continue;
      if (status !== 'ACTIVE') continue;
    }

    let parsedVersions = [];
    try {
      if (row[14]) parsedVersions = JSON.parse(row[14]);
    } catch(e) {}

    const name = row[2] || row[6] || "Nama dokumen tidak tersedia";
    const docNumber = row[1] || '000';
    const version = row[9] || 'v1.0';
    result.push({
      // New Standardized Fields
      documentId: docId,
      documentNumber: docNumber,
      documentName: name,
      category: catName,
      categoryId: catId,
      fileId: row[5],
      fileName: row[6],
      fileType: row[7],
      fileSize: row[8],
      version: version,
      status: status,
      uploadedAt: row[11],
      uploadedBy: row[12],
      description: row[13],
      versions: parsedVersions,
      
      // Fallback for UI stability
      id: docId,
      nomorDokumen: docNumber,
      nomor: docNumber,
      nama: name,
      namaDokumen: name,
      kategoriId: catId,
      kategoriNama: catName,
      kategori: catName,
      versi: version,
      tanggalUpload: row[11],
      createdAt: row[11],
      uploader: row[12],
      deskripsi: row[13]
    });
  }

  return result;
}

function uploadCKFileDrive(base64Data, fileName, mimeType, categoryName, customName, uploadedBy, isAdmin) {
  if (!isAdmin) throw new Error('Akses ditolak: Hanya Admin yang dapat mengunggah dokumen.');
  const { sheetDoc } = initCKSheets();
  const folder = getOrCreateCKDriveFolder(categoryName);
  
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setDescription(`Corporate Knowledge: ${categoryName} - ${customName || fileName}`);

  const autoNomor = getCKNextDocNumber();
  const docId = 'CK-DOC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const finalName = customName || fileName.replace(/\.[^/.]+$/, "");
  const nowStr = new Date().toISOString();

  const versionsArray = [
    {
      version: 'v1.0',
      fileId: file.getId(),
      fileName: fileName,
      fileType: mimeType,
      fileSize: file.getSize(),
      uploadedAt: nowStr,
      uploadedBy: uploadedBy || 'Admin HR',
      notes: 'Initial upload (Versi awal)'
    }
  ];

  sheetDoc.appendRow([
    docId,
    autoNomor,
    finalName,
    'cat-' + categoryName.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
    categoryName,
    file.getId(),
    fileName,
    mimeType,
    file.getSize(),
    'v1.0',
    'ACTIVE',
    nowStr,
    uploadedBy || 'Admin HR',
    '',
    JSON.stringify(versionsArray)
  ]);

  return {
    id: docId,
    nomor: autoNomor,
    namaDokumen: finalName,
    kategori: categoryName,
    fileId: file.getId(),
    fileName: fileName,
    versi: 'v1.0'
  };
}

function bulkUploadCKFilesDrive(filesArray, categoryName, uploadedBy, isAdmin) {
  if (!isAdmin) throw new Error('Akses ditolak: Hanya Admin yang dapat mengunggah dokumen.');
  const results = [];
  const errors = [];

  filesArray.forEach((item, idx) => {
    try {
      const res = uploadCKFileDrive(
        item.base64Data,
        item.fileName,
        item.mimeType,
        item.kategori || categoryName,
        item.customName,
        uploadedBy,
        true
      );
      results.push(res);
    } catch(err) {
      errors.push({ fileName: item.fileName, error: err.message });
    }
  });

  return {
    successCount: results.length,
    failedCount: errors.length,
    items: results,
    errors: errors
  };
}

function saveCKNewVersionDrive(docId, base64Data, fileName, mimeType, newVersion, notes, uploadedBy, isAdmin) {
  if (!isAdmin) throw new Error('Akses ditolak: Hanya Admin yang dapat memperbarui versi.');
  const { sheetDoc } = initCKSheets();
  const data = sheetDoc.getDataRange().getValues();

  let targetRow = -1;
  let catName = 'Lainnya';
  let prevVersions = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docId) {
      targetRow = i + 1;
      catName = data[i][4];
      try {
        if (data[i][14]) prevVersions = JSON.parse(data[i][14]);
      } catch(e) {}
      break;
    }
  }

  if (targetRow === -1) throw new Error('Dokumen tidak ditemukan.');

  const folder = getOrCreateCKDriveFolder(catName);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = folder.createFile(blob);
  const nowStr = new Date().toISOString();

  const verRecord = {
    version: newVersion,
    fileId: file.getId(),
    fileName: fileName,
    fileType: mimeType,
    fileSize: file.getSize(),
    uploadedAt: nowStr,
    uploadedBy: uploadedBy || 'Admin HR',
    notes: notes || 'Pembaruan versi'
  };

  prevVersions.push(verRecord);

  sheetDoc.getRange(targetRow, 6).setValue(file.getId());
  sheetDoc.getRange(targetRow, 7).setValue(fileName);
  sheetDoc.getRange(targetRow, 8).setValue(mimeType);
  sheetDoc.getRange(targetRow, 9).setValue(file.getSize());
  sheetDoc.getRange(targetRow, 10).setValue(newVersion);
  sheetDoc.getRange(targetRow, 15).setValue(JSON.stringify(prevVersions));

  return { success: true, version: newVersion };
}

function getCKData(isAdmin) {
  try {
    const categories = getCKCategories(isAdmin);
    const documents = getCKDocuments({}, isAdmin);
    const { sheetConfig } = initCKSheets();
    let nextSeq = 1;
    try {
      const data = sheetConfig.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'NEXT_DOC_SEQUENCE') {
          nextSeq = parseInt(data[i][1]) || 1;
          break;
        }
      }
    } catch(e) {}

    return {
      success: true,
      categories: categories,
      documents: documents,
      nextSeqNumber: nextSeq
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
      categories: [],
      documents: [],
      nextSeqNumber: 1
    };
  }
}

function simpanCKKategori(payload, isAdmin) {
  try {
    const res = saveCKCategory(payload, isAdmin);
    return { success: true, categories: res, message: 'Kategori berhasil disimpan' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function hapusCKKategori(catId, isAdmin) {
  if (!isAdmin) return { success: false, error: 'Akses ditolak' };
  try {
    const { sheetCat, sheetDoc } = initCKSheets();
    const docRows = sheetDoc.getDataRange().getValues();
    for (let i = 1; i < docRows.length; i++) {
      if ((docRows[i][3] === catId || docRows[i][4] === catId) && docRows[i][10] !== 'DELETED') {
        return { success: false, error: 'Kategori masih memiliki dokumen aktif' };
      }
    }
    const catRows = sheetCat.getDataRange().getValues();
    for (let i = 1; i < catRows.length; i++) {
      if (catRows[i][0] === catId) {
        sheetCat.deleteRow(i + 1);
        return { success: true, message: 'Kategori berhasil dihapus' };
      }
    }
    return { success: false, error: 'Kategori tidak ditemukan' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function unggahCKDokumen(payload, isAdmin) {
  try {
    if (!isAdmin) throw new Error('Akses ditolak: Hanya Admin yang dapat mengunggah dokumen.');
    const base64Data = payload.fileData || '';
    const fileName = payload.fileName || 'dokumen.docx';
    const mimeType = payload.fileType || 'application/octet-stream';
    const customName = payload.nama || fileName.replace(/\.[^/.]+$/, "");
    const categoryId = payload.kategoriId || '';
    const uploadedBy = payload.uploader || 'Admin HR';
    const deskripsi = payload.deskripsi || '';

    const { sheetDoc, sheetCat } = initCKSheets();
    
    // Find category name
    let catName = 'Lainnya';
    const catRows = sheetCat.getDataRange().getValues();
    for (let i = 1; i < catRows.length; i++) {
      if (catRows[i][0] === categoryId || catRows[i][1] === categoryId || catRows[i][0] === payload.kategori || catRows[i][1] === payload.kategori) {
        catName = catRows[i][1];
        break;
      }
    }

    const folder = getOrCreateCKDriveFolder(catName);
    let cleanBase64 = base64Data;
    if (base64Data.indexOf('base64,') > -1) {
      cleanBase64 = base64Data.split('base64,')[1];
    }
    const decoded = Utilities.base64Decode(cleanBase64);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setDescription(`Corporate Knowledge: ${catName} - ${customName}`);

    const autoNomor = getCKNextDocNumber();
    const docId = 'CK-DOC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const nowStr = new Date().toISOString();

    const versionsArray = [
      {
        version: 'v1.0',
        fileId: file.getId(),
        fileName: fileName,
        fileType: mimeType,
        fileSize: file.getSize(),
        uploadedAt: nowStr,
        uploadedBy: uploadedBy,
        notes: 'Initial upload (Versi awal)'
      }
    ];

    sheetDoc.appendRow([
      docId,
      autoNomor,
      customName,
      'cat-' + catName.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
      catName,
      file.getId(),
      fileName,
      mimeType,
      file.getSize(),
      'v1.0',
      'ACTIVE',
      nowStr,
      uploadedBy,
      deskripsi,
      JSON.stringify(versionsArray)
    ]);

    return {
      success: true,
      document: {
        id: docId,
        nomorDokumen: autoNomor,
        nomor: autoNomor,
        nama: customName,
        namaDokumen: customName,
        kategoriId: 'cat-' + catName.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
        kategoriNama: catName,
        kategori: catName,
        fileId: file.getId(),
        fileName: fileName,
        fileType: mimeType,
        fileSize: file.getSize(),
        version: 'v1.0',
        versi: 'v1.0',
        status: 'ACTIVE',
        uploader: uploadedBy,
        uploadedBy: uploadedBy,
        deskripsi: deskripsi,
        createdAt: nowStr,
        versions: versionsArray
      },
      message: 'Dokumen berhasil diunggah dengan nomor [' + autoNomor + '].'
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
      message: 'Gagal mengunggah dokumen: ' + err.message
    };
  }
}

function perbaruiVersiCKDokumen(payload, isAdmin) {
  try {
    const docId = payload.documentId || payload.id;
    const base64Data = payload.fileData;
    const fileName = payload.fileName;
    const mimeType = payload.fileType;
    const newVersion = payload.newVersion;
    const notes = payload.notes;
    const uploadedBy = payload.uploader;
    const res = saveCKNewVersionDrive(docId, base64Data, fileName, mimeType, newVersion, notes, uploadedBy, isAdmin);
    return res;
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function simpanCKDokumen(payload, isAdmin) {
  if (!isAdmin) return { success: false, error: 'Akses ditolak' };
  try {
    const { sheetDoc, sheetCat } = initCKSheets();
    const data = sheetDoc.getDataRange().getValues();
    const docId = payload.id;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === docId) {
        const rowNum = i + 1;
        if (payload.nama) sheetDoc.getRange(rowNum, 3).setValue(payload.nama);
        if (payload.kategoriId) {
          sheetDoc.getRange(rowNum, 4).setValue(payload.kategoriId);
          const catRows = sheetCat.getDataRange().getValues();
          for (let j = 1; j < catRows.length; j++) {
            if (catRows[j][0] === payload.kategoriId) {
              sheetDoc.getRange(rowNum, 5).setValue(catRows[j][1]);
              break;
            }
          }
        }
        if (payload.status) sheetDoc.getRange(rowNum, 11).setValue(payload.status);
        if (typeof payload.deskripsi === 'string') sheetDoc.getRange(rowNum, 14).setValue(payload.deskripsi);
        return { success: true, message: 'Dokumen berhasil diperbarui' };
      }
    }
    return { success: false, error: 'Dokumen tidak ditemukan' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function hapusCKDokumen(docId, isAdmin) {
  try {
    if (!isAdmin) {
      return { success: false, error: 'Akses ditolak: Hanya Admin yang dapat menghapus dokumen.' };
    }

    const { sheetDoc } = initCKSheets();
    const docRows = sheetDoc.getDataRange().getValues();
    
    let deletedDoc = null;
    let rowIndex = -1;

    for (let i = 1; i < docRows.length; i++) {
      if (docRows[i][0] === docId) {
        deletedDoc = {
          nomor: docRows[i][1],
          fileId: docRows[i][5]
        };
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Dokumen tidak ditemukan' };
    }

    // Try deleting file first
    if (deletedDoc.fileId) {
      try {
        const file = DriveApp.getFileById(deletedDoc.fileId);
        file.setTrashed(true);
      } catch (e) {
        return { success: false, error: 'Gagal menghapus file di Google Drive: ' + e.message };
      }
    }

    // Then delete from metadata
    try {
      sheetDoc.deleteRow(rowIndex);
    } catch (e) {
      return { success: false, error: 'File berhasil diproses tetapi metadata gagal diperbarui. Jangan membuat data menjadi tidak sinkron.' };
    }

    return { 
      success: true, 
      message: 'Dokumen berhasil dihapus', 
      documentId: docId 
    };
  } catch(err) {
    return { success: false, error: err.message };
  }
}
