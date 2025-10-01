function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    
    // Inisialisasi sheet
    var sheet = initializeSheet();
    
    // Format data sesuai dengan struktur sheet
    var hari = data.hari;
    var tanggal = data.tanggal; // Format: YYYY-MM-DD
    var jam = data.jam; // Format: HH:MM:SS
    var namaBidan = data.namaBidan;
    var kegiatan = data.kegiatan;
    var namaPasien = data.namaPasien;
    var lokasiHomeCare = data.lokasiHomeCare || '';
    var keterangan = data.keterangan || '';

    // Tambahkan baris baru di sheet
    sheet.appendRow([hari, tanggal, jam, namaBidan, kegiatan, namaPasien, lokasiHomeCare, keterangan]);

    return ContentService.createTextOutput(JSON.stringify({
      "result": "success",
      "message": "Data berhasil disimpan"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action;
    
    if (action === 'getData') {
      return getDataFromSheet();
    } else if (action === 'initialize') {
      return initializeSheetResponse();
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "message": "Action tidak dikenali"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function initializeSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "REKAPBIDAN";
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // Jika sheet tidak ada, buat baru
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    
    // Buat header
    var headers = [
      "Hari", 
      "Tanggal", 
      "Jam", 
      "Nama Bidan", 
      "Jenis Kegiatan", 
      "Nama Pasien",
      "Lokasi Home Care",
      "Keterangan"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#ff9ab4")
              .setFontColor("white")
              .setFontWeight("bold");
    
    // Bekukan baris header
    sheet.setFrozenRows(1);
    
    // Atur lebar kolom
    sheet.setColumnWidth(1, 100); // Hari
    sheet.setColumnWidth(2, 120); // Tanggal
    sheet.setColumnWidth(3, 100); // Jam
    sheet.setColumnWidth(4, 150); // Nama Bidan
    sheet.setColumnWidth(5, 180); // Jenis Kegiatan
    sheet.setColumnWidth(6, 150); // Nama Pasien
    sheet.setColumnWidth(7, 200); // Lokasi Home Care
    sheet.setColumnWidth(8, 250); // Keterangan
    
    // Tambahkan data contoh jika sheet baru
    addSampleData(sheet);
  }
  
  return sheet;
}

function addSampleData(sheet) {
  // Tambahkan beberapa data contoh untuk testing
  var sampleData = [
    ['Kamis', '2025-09-25', '16:30:00', 'Linda Warnii', 'Lembur', 'Yuli Ridata Hasibuan', '', 'Lembur malam'],
    ['Minggu', '2025-09-07', '09:39:00', 'aisyah', 'Home Care', 'beryl', 'Rumah Pasien - Jl. Merdeka No. 123', 'Kontrol rutin'],
    ['Senin', '2025-09-01', '09:40:00', 'Riska Wardana', 'Home Care', 'yasmine', 'Rumah Pasien - Jl. Sudirman No. 45', 'Perawatan pasca melahirkan']
  ];
  
  if (sheet.getLastRow() === 1) { // Hanya tambahkan jika hanya ada header
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

function initializeSheetResponse() {
  try {
    var sheet = initializeSheet();
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success",
      "message": "Sheet berhasil diinisialisasi"
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getDataFromSheet() {
  try {
    var sheet = initializeSheet();
    var data = sheet.getDataRange().getValues();
    
    // Hapus header
    data.shift();
    
    // Format data menjadi array of objects
    var formattedData = data.map(function(row) {
      return {
        hari: row[0] || '',
        tanggal: formatDateForDisplay(row[1]),
        jam: formatTimeForDisplay(row[2]),
        namaBidan: row[3] || '',
        kegiatan: row[4] || '',
        namaPasien: row[5] || '',
        lokasiHomeCare: row[6] || '',
        keterangan: row[7] || ''
      };
    }).filter(function(row) {
      // Hapus baris kosong
      return row.namaBidan !== '';
    });
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success",
      "data": formattedData
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function formatDateForDisplay(dateValue) {
  if (!dateValue) return '';
  
  try {
    if (typeof dateValue === 'string') {
      return dateValue;
    } else if (dateValue instanceof Date) {
      // Format: YYYY-MM-DD
      var year = dateValue.getFullYear();
      var month = ('0' + (dateValue.getMonth() + 1)).slice(-2);
      var day = ('0' + dateValue.getDate()).slice(-2);
      return year + '-' + month + '-' + day;
    }
  } catch (e) {
    return '';
  }
  
  return '';
}

function formatTimeForDisplay(timeValue) {
  if (!timeValue) return '';
  
  try {
    if (typeof timeValue === 'string') {
      return timeValue;
    } else if (timeValue instanceof Date) {
      // Format: HH:MM:SS
      var hours = ('0' + timeValue.getHours()).slice(-2);
      var minutes = ('0' + timeValue.getMinutes()).slice(-2);
      var seconds = ('0' + timeValue.getSeconds()).slice(-2);
      return hours + ':' + minutes + ':' + seconds;
    }
  } catch (e) {
    return '';
  }
  
  return '';
}