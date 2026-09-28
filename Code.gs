function myFunction() {
  var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var payload = requestData.payload;
    function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);/**
 * TPK DESA PUSUK LESTARI - BACKEND APPS SCRIPT
 * Web App Endpoint & Database Controller
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('TPK DESA PUSUK LESTARI - Sistem Informasi Pendampingan Keluarga')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}
    var response = {};

    switch(action) {
      case 'setupDatabase':
        response = setupDatabase();
        break;
      case 'login':
        response = processLogin(payload);
        break;
      case 'getData':
        response = fetchAllData();
        break;
      case 'saveSasaran':
        response = saveSasaranData(payload);
        break;
      case 'addKunjungan':
        response = recordKunjungan(payload);
        break;
      case 'backupDatabase':
        response = createSpreadsheetBackup();
        break;
      default:
        response = { status: 'error', message: 'Aksi tidak dikenali' };
    }

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Inisialisasi Otomatis Google Sheets Database
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetsToCreate = [
    { name: 'CONFIG', headers: ['KEY', 'VALUE'] },
    { name: 'USERS', headers: ['ID', 'USERNAME', 'PASSWORD', 'ROLE', 'NAMA'] },
    { name: 'MASTER_DUSUN', headers: ['ID', 'NAMA_DUSUN'] },
    { name: 'MASTER_TPK', headers: ['ID', 'NAMA_TPK', 'KETUA'] },
    { name: 'BADUTA', headers: ['ID', 'NIK', 'NAMA', 'JK', 'TGL_LAHIR', 'WALI', 'HP', 'DUSUN', 'BB', 'TB', 'GIZI', 'KEMBANG', 'STATUS', 'TPK', 'KET'] },
    { name: 'BUMIL', headers: ['ID', 'NIK', 'NAMA', 'TGL_LAHIR', 'SUAMI', 'HP', 'DUSUN', 'USIA_HAMIL', 'HPL', 'LILA', 'HB', 'KEK', 'ANEMIA', 'RISIKO', 'STATUS', 'TPK', 'KET'] },
    { name: 'NIPAS', headers: ['ID', 'NIK', 'NAMA', 'TGL_LAHIR', 'TGL_SALIN', 'SUAMI', 'HP', 'DUSUN', 'ASI', 'KB', 'BAHAYA', 'STATUS', 'TPK', 'KET'] },
    { name: 'CATIN', headers: ['ID', 'NIK', 'NAMA', 'JK', 'TGL_LAHIR', 'PASANGAN', 'HP', 'DUSUN', 'TGL_NIKAH', 'HB', 'LILA', 'KESIAPAN', 'STATUS', 'TPK', 'KET'] },
    { name: 'KUNJUNGAN', headers: ['ID', 'TARGET_ID', 'TGL', 'KUNJUNGAN_KE', 'HASIL', 'TINDAK_LANJUT', 'TPK'] },
    { name: 'AUDIT_LOG', headers: ['TIMESTAMP', 'USER', 'AKSI', 'TARGET', 'ID_DATA', 'KETERANGAN'] }
  ];

  sheetsToCreate.forEach(function(item) {
    var sheet = ss.getSheetByName(item.name);
    if (!sheet) {
      sheet = ss.insertSheet(item.name);
      sheet.getRange(1, 1, 1, item.headers.length).setValues([item.headers])
        .setFontWeight('bold')
        .setBackground('#1F4E78')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  });

  // Default User Admin
  var userSheet = ss.getSheetByName('USERS');
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow(['USR-001', 'admin', '123', 'ADMIN', 'Administrator TPK']);
  }

  return { status: 'success', message: 'Database TPK Desa Pusuk Lestari Berhasil Disiapkan!' };
}

function processLogin(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('USERS');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var username = String(data[i][1]);
    var password = String(data[i][2]);
    if (username === payload.username && password === payload.password) {
      logAudit(payload.username, 'LOGIN', 'USERS', data[i][0], 'Login sukses');
      return {
        status: 'success',
        user: { username: data[i][1], role: data[i][3], nama: data[i][4] }
      };
    }
  }
  return { status: 'error', message: 'Username atau Password salah!' };
}

function fetchAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    status: 'success',
    baduta: getSheetDataValues(ss, 'BADUTA'),
    bumil: getSheetDataValues(ss, 'BUMIL'),
    nipas: getSheetDataValues(ss, 'NIPAS'),
    catin: getSheetDataValues(ss, 'CATIN'),
    kunjungan: getSheetDataValues(ss, 'KUNJUNGAN')
  };
}

function getSheetDataValues(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  
  var headers = rows[0];
  var results = [];
  
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j].toLowerCase()] = row[j];
    }
    results.push(obj);
  }
  return results;
}

function logAudit(user, aksi, target, idData, ket) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AUDIT_LOG');
  if (sheet) {
    var now = Utilities.formatDate(new Date(), 'Asia/Makassar', 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([now, user, aksi, target, idData, ket]);
  }
}

function createSpreadsheetBackup() {
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var folderName = "TPK_Pusuk_Lestari_Backups";
  var folders = DriveApp.getFoldersByName(folderName);
  var targetFolder;
  
  if (folders.hasNext()) {
    targetFolder = folders.next();
  } else {
    targetFolder = DriveApp.createFolder(folderName);
  }
  
  var backupName = "BACKUP_TPK_PUSUK_LESTARI_" + Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyyMMdd_HHmmss');
  file.makeCopy(backupName, targetFolder);
  
  return { status: 'success', message: 'Backup database berhasil disimpan di Google Drive folder ' + folderName };
}

/* PLACEHOLDER: FUNGSI CRUD BARU */
}
