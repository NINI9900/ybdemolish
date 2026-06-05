// =============================================
// 청년철거 시공사례 API - Google Apps Script
// =============================================
// 배포 URL: https://script.google.com/macros/s/AKfycbwZkOmLSIertg7CaImERwnxXlqobBjeDeoTOjfK6MS_BD4gse3pr-L_Y73rbc27huNR/exec
// 연결 시트: https://docs.google.com/spreadsheets/d/1IxGqPQhEotHPwqf2CiQetn4tI9Sciymhx0PEyY2Kebc

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const result = rows
    .map((r, i) => ({ r, i: i + 2 }))
    .filter(({ r }) => {
      const name = String(r[1] || '').trim();
      const uploaded = String(r[4] || '').trim().toUpperCase();
      const thumb = String(r[12] || '').trim();
      const icon = String(r[11] || '').trim();
      return name && thumb && icon && (uploaded === 'Y' || uploaded === 'O' || uploaded === '완료' || uploaded === '업로드');
    })
    .map(({ r, i }) => ({
      id:       i,
      name:     String(r[1]  || '').trim(),
      category: String(r[2]  || '').trim(),
      size:     String(r[3]  || '').trim(),
      region:   String(r[5]  || '').trim(),
      type:     String(r[6]  || '').trim(),
      subsidy:  String(r[7]  || '').trim(),
      date:     String(r[8]  || '').trim(),
      period:   String(r[9]  || '').trim(),
      link:     String(r[10] || '').trim(),
      icon:     String(r[11] || '').trim(),
      thumb:    String(r[12] || '').trim(),
      img1:     String(r[13] || '').trim(),
      img2:     String(r[14] || '').trim(),
      img3:     String(r[15] || '').trim(),
      img4:     String(r[16] || '').trim(),
      img5:     String(r[17] || '').trim(),
      img6:     String(r[18] || '').trim(),
      img7:     String(r[19] || '').trim(),
      img8:     String(r[20] || '').trim(),
      img9:     String(r[21] || '').trim(),
      img10:    String(r[22] || '').trim(),
    }));

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
