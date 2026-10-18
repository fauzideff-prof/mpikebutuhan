
/**
 * Serves the HTML Web App UI when accessed via URL.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Media Pembelajaran IPS VII - SMPN 3 PAKUHAJI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Saves student quiz scores, game results, and reflections to Google Sheets.
 * @param {Object} data - Student progress and answers.
 */
function saveStudentData(data) {
  try {
    if (!data || typeof data !== 'object') {
      throw new Error('Data siswa tidak valid.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error(
        'Spreadsheet tidak ditemukan. Pastikan script terhubung ke Google Sheets.'
      );
    }

    var sheet = ss.getSheetByName('Hasil_Siswa');

    // Create worksheet and header if it does not exist.
    if (!sheet) {
      sheet = ss.insertSheet('Hasil_Siswa');

      sheet.appendRow([
        'Waktu Tersimpan',
        'Nama Siswa',
        'Skor Kuis (0-100)',
        'Jumlah Benar',
        'Jumlah Salah',
        'Skor Game (0-100)',
        'Refleksi Q1 (Primer vs Sekunder)',
        'Refleksi Q2 (Alasan Skala Prioritas)',
        'Refleksi Q3 (3 Kebutuhan Belajar)'
      ]);

      var headerRange = sheet.getRange(1, 1, 1, 9);
      headerRange
        .setFontWeight('bold')
        .setBackground('#1E3A8A')
        .setFontColor('#FFFFFF');

      sheet.setFrozenRows(1);
    }

    // Save student record.
    sheet.appendRow([
      new Date(),
      data.studentName || 'Siswa Tanpa Nama',
      data.quizScore !== undefined ? data.quizScore : 0,
      data.correctAnswers !== undefined ? data.correctAnswers : 0,
      data.wrongAnswers !== undefined ? data.wrongAnswers : 0,
      data.gameScore !== undefined ? data.gameScore : 0,
      data.reflectionQ1 || '-',
      data.reflectionQ2 || '-',
      data.reflectionQ3 || '-'
    ]);

    sheet.autoResizeColumns(1, 9);

    return {
      success: true,
      message: 'Data hasil belajar siswa berhasil disimpan ke Google Sheets!'
    };

  } catch (err) {
    return {
      success: false,
      error: err.message || String(err)
    };
  }
}
