// =============================================
// 청년철거 파트너 지원 - Google Apps Script
// =============================================
// 이 코드를 통째로 복사해서 Apps Script에 붙여넣으세요

const SHEET_ID = '1ZwtxHu_DyBlrtuh2kz2cRWIikrBvjqE-PUGENNMqSg0';
const NOTIFY_EMAIL = 'a62936559@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ── 스프레드시트 저장 ──
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

    // 첫 행이 비어있으면 헤더 추가
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출일시', '업체명', '담당자', '연락처', '이메일', '지역', '업종', '인원수', '한마디']);
      // 헤더 스타일
      const header = sheet.getRange(1, 1, 1, 9);
      header.setFontWeight('bold');
      header.setBackground('#1B6CF7');
      header.setFontColor('#ffffff');
    }

    // 데이터 추가
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      data.company || '',
      data.name    || '',
      data.phone   || '',
      data.email   || '',
      data.region  || '',
      data.type     || '',
      data.teamsize || '',
      data.message  || ''
    ]);

    // 열 너비 자동 조정
    sheet.autoResizeColumns(1, 9);

    // ── 이메일 발송 ──
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: `[청년철거] 파트너 지원 접수 - ${data.company} / ${data.name}`,
      htmlBody: `
        <div style="font-family: sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="background: #1B6CF7; padding: 24px 32px; border-radius: 12px 12px 0 0;">
            <h2 style="color: #fff; margin: 0; font-size: 20px;">파트너 지원 접수</h2>
          </div>
          <div style="background: #f9f9f9; padding: 32px; border-radius: 0 0 12px 12px; border: 1px solid #eee;">
            <table style="width: 100%; border-collapse: collapse; font-size: 15px;">
              <tr><td style="padding: 10px 0; color: #888; width: 100px;">업체명</td><td style="padding: 10px 0; font-weight: 600;">${data.company}</td></tr>
              <tr><td style="padding: 10px 0; color: #888;">담당자</td><td style="padding: 10px 0;">${data.name}</td></tr>
              <tr><td style="padding: 10px 0; color: #888;">연락처</td><td style="padding: 10px 0;">${data.phone}</td></tr>
              <tr><td style="padding: 10px 0; color: #888;">이메일</td><td style="padding: 10px 0;">${data.email || '-'}</td></tr>
              <tr><td style="padding: 10px 0; color: #888;">지역</td><td style="padding: 10px 0;">${data.region}</td></tr>
              <tr><td style="padding: 10px 0; color: #888;">업종</td><td style="padding: 10px 0;">${data.type}</td></tr>
              <tr><td style="padding: 10px 0; color: #888;">인원수</td><td style="padding: 10px 0;">${data.teamsize || '-'}</td></tr>
              <tr><td style="padding: 10px 0; color: #888; vertical-align: top;">한마디</td><td style="padding: 10px 0;">${data.message || '-'}</td></tr>
            </table>
            <div style="margin-top: 24px; padding-top: 16px; border-top: 1px solid #eee; color: #aaa; font-size: 13px;">
              제출 시각: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 MM월 dd일 HH:mm:ss')}
            </div>
          </div>
        </div>
      `
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
