// =============================================
// 청년철거 견적문의 - Google Apps Script
// =============================================

const SHEET_ID     = '1Tbv5An6wX5HpR4MPMh47xAKrf5mCr2Jf932CMHEqvxw';
const NOTIFY_EMAIL = 'a62936559@gmail.com';
const FOLDER_NAME  = '청년철거_견적문의_사진';

function doPost(e) {
  try {
    let data;
    // form submit 방식
    if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = {};
    }

    // ── 스프레드시트 저장 ──
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출일시','현장주소','연락처','철거유형','상세내용','폐업지원금','방문가능일자','사진링크']);
      const header = sheet.getRange(1, 1, 1, 8);
      header.setFontWeight('bold');
      header.setBackground('#1B6CF7');
      header.setFontColor('#ffffff');
    }

    // ── 사진 Drive 저장 ──
    let photoLinks = '-';
    if (data.photos && data.photos.length > 0) {
      const folder = getOrCreateFolder(FOLDER_NAME);
      const submitTime = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
      const subFolder = folder.createFolder(`${submitTime}_${data.phone}`);
      const links = [];
      data.photos.forEach((photo, i) => {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(photo.data),
          'image/jpeg',
          photo.name || `photo_${i+1}.jpg`
        );
        const file = subFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        links.push(file.getUrl());
      });
      photoLinks = links.join('\n');
    }

    // 데이터 저장
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      data.address  || '',
      data.phone    || '',
      data.type     || '',
      data.detail   || '',
      data.subsidy  || '',
      data.date     || '',
      photoLinks
    ]);
    sheet.autoResizeColumns(1, 8);

    // ── 이메일 발송 ──
    const photoHtml = data.photos && data.photos.length > 0
      ? `<tr><td style="padding:10px 0;color:#888;vertical-align:top;">사진</td><td style="padding:10px 0;">${data.photos.length}장 첨부 (Drive 저장)</td></tr>`
      : '';

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: `[청년철거] 견적문의 접수 - ${data.address} / ${data.phone}`,
      htmlBody: `
        <div style="font-family:sans-serif;max-width:600px;margin:0 auto;">
          <div style="background:#1B6CF7;padding:24px 32px;border-radius:12px 12px 0 0;">
            <h2 style="color:#fff;margin:0;font-size:20px;">견적 문의 접수</h2>
          </div>
          <div style="background:#f9f9f9;padding:32px;border-radius:0 0 12px 12px;border:1px solid #eee;">
            <table style="width:100%;border-collapse:collapse;font-size:15px;">
              <tr><td style="padding:10px 0;color:#888;width:120px;">현장 주소</td><td style="padding:10px 0;font-weight:600;">${data.address}</td></tr>
              <tr><td style="padding:10px 0;color:#888;">연락처</td><td style="padding:10px 0;">${data.phone}</td></tr>
              <tr><td style="padding:10px 0;color:#888;">철거 유형</td><td style="padding:10px 0;">${data.type}</td></tr>
              <tr><td style="padding:10px 0;color:#888;vertical-align:top;">상세 내용</td><td style="padding:10px 0;">${data.detail || '-'}</td></tr>
              <tr><td style="padding:10px 0;color:#888;">폐업지원금</td><td style="padding:10px 0;">${data.subsidy || '-'}</td></tr>
              <tr><td style="padding:10px 0;color:#888;">방문 가능일</td><td style="padding:10px 0;">${data.date || '-'}</td></tr>
              ${photoHtml}
            </table>
            <div style="margin-top:24px;padding-top:16px;border-top:1px solid #eee;color:#aaa;font-size:13px;">
              제출 시각: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 MM월 dd일 HH:mm:ss')}
            </div>
          </div>
        </div>
      `
    });

    return ContentService
      .createTextOutput('ok')
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService
      .createTextOutput('error: ' + err.toString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  return doPost(e);
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}
