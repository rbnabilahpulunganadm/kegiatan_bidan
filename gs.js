function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Ganti "Sheet1" dengan nama sheet Anda jika berbeda
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REKAPBIDAN"); 

    // Ambil data dari aplikasi
    var hari = data.hari;
    var tanggal = data.tanggal;
    var jam = data.jam;
    var namaBidan = data.namaBidan;
    var kegiatan = data.kegiatan;
    var namaPasien = data.namaPasien;

    // Tambahkan baris baru di sheet
    sheet.appendRow([hari, tanggal, jam, namaBidan, kegiatan, namaPasien]);

    return ContentService.createTextOutput(JSON.stringify({
      "result": "success"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}