const SPREADSHEET_TITLE = 'Rekod Kehadiran Program IoT dan Automasi Pintar AI';
const SHEET_NAME = 'Respons Kehadiran';

function setupAttendanceBackend() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('ATTENDANCE_SPREADSHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_TITLE);
    props.setProperty('ATTENDANCE_SPREADSHEET_ID', ss.getId());
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Tarikh/Masa Hantaran','Nama','No. Kad Pengenalan','No. Telefon','Organisasi / Nama Sekolah','Mata Pelajaran','Tahun / Tingkatan','Pengalaman Mengajar']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,8).setFontWeight('bold');
    sheet.autoResizeColumns(1,8);
    sheet.getRange('C:D').setNumberFormat('@');
  }
  Logger.log('Google Sheets respons: ' + ss.getUrl());
}

function doPost(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    if (p.website && p.website.trim()) return output_({ok:false,message:'Rejected'});
    const nama=clean_(p.nama,150), ic=clean_(p.ic,12), telefon=clean_(p.telefon,11);
    const organisasi=clean_(p.organisasi,200), subjek=clean_(p.subjek,200);
    const tahun=clean_(p.tahun,150), pengalaman=clean_(p.pengalaman,50);
    if (!nama || !/^[0-9]{12}$/.test(ic) || !/^01[0-9]{8,9}$/.test(telefon) || !organisasi || !subjek || !tahun || !pengalaman)
      return output_({ok:false,message:'Maklumat tidak lengkap atau format tidak sah.'});
    const allowed=['Kurang daripada 1 tahun','1–5 tahun','6–10 tahun','11–15 tahun','16–20 tahun','Lebih daripada 20 tahun'];
    if (!allowed.includes(pengalaman)) return output_({ok:false,message:'Pilihan pengalaman tidak sah.'});
    const lock=LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const id=PropertiesService.getScriptProperties().getProperty('ATTENDANCE_SPREADSHEET_ID');
      if (!id) throw new Error('Jalankan setupAttendanceBackend() terlebih dahulu.');
      const sheet=SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
      if (!sheet) throw new Error('Helaian respons tidak ditemui.');
      sheet.appendRow([new Date(),safeCell_(nama),"'"+ic,"'"+telefon,safeCell_(organisasi),safeCell_(subjek),safeCell_(tahun),safeCell_(pengalaman)]);
    } finally { lock.releaseLock(); }
    return output_({ok:true,message:'Kehadiran direkodkan.'});
  } catch(err) {
    console.error(err);
    return output_({ok:false,message:'Ralat pelayan. Sila hubungi urus setia.'});
  }
}
function doGet() { return output_({ok:true,message:'Attendance endpoint is active.'}); }
function clean_(value,maxLength) { return String(value||'').trim().slice(0,maxLength); }
function safeCell_(value) { const s=String(value||''); return /^[=+\-@]/.test(s) ? "'"+s : s; }
function output_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }