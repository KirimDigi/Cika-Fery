/***************************************************************
 * KIRIMDIGI - RSVP & Ucapan -> Google Spreadsheet
 * Cara pakai:
 * 1. Buka spreadsheet kamu > menu Extensions > Apps Script
 * 2. Hapus semua isi editor, copy-paste SELURUH file ini
 * 3. Save (Ctrl+S)
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy URL /exec > paste ke index.html pada var SHEET_API
 * Tab "Ucapan" akan dibuat otomatis (tab lama tidak diubah).
 ***************************************************************/

const TAB = 'Ucapan';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.appendRow(['Timestamp', 'Nama', 'Kehadiran', 'JumlahTamu', 'Ucapan']);
  }
  return sh;
}

// Dipanggil undangan untuk MENAMPILKAN ucapan
function doGet() {
  const v = sheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][1] && !v[i][4]) continue; // lewati baris kosong
    out.push({
      timestamp: v[i][0],
      at: v[i][0],
      name: v[i][1],
      attend: v[i][2],
      guests: v[i][3],
      message: v[i][4]
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Dipanggil undangan setiap tamu menekan Kirim
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  sheet_().appendRow([
    d.timestamp ? new Date(d.timestamp) : new Date(),
    d.name || '',
    d.attend || '',
    d.guests || '',
    d.message || ''
  ]);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
