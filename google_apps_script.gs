/**
 * KEEPINTOUCH 캠페인 신청 폼 → 구글 시트 연동용 Apps Script
 *
 * 설정 방법:
 * 1. 새 Google 시트를 만든다 (또는 기존 시트 사용).
 * 2. 시트 하단에 "신청자" 라는 이름의 탭을 만든다. (다른 이름을 쓰려면 아래 SHEET_NAME 값도 같이 바꿀 것)
 * 3. 상단 메뉴 확장 프로그램(Extensions) > Apps Script 를 연다.
 * 4. 기본 코드(Code.gs)를 지우고 이 파일 내용 전체를 붙여넣는다.
 * 5. 저장 후 배포(Deploy) > 새 배포(New deployment) 클릭.
 *    - 유형: 웹 앱(Web app)
 *    - 실행: 나(Me)
 *    - 액세스 권한: 모든 사용자(Anyone)
 * 6. 배포 후 나오는 "웹 앱 URL"을 복사한다.
 * 7. keepintouch_musinsa_space1_guide.html 파일에서
 *    const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL';
 *    이 줄의 값을 방금 복사한 URL로 교체한다.
 */

const SHEET_NAME = '신청자';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출 시각', '이름', '인스타그램 링크', '휴대폰 번호', '이메일', '방문 희망일', '타입', '고료']);
    }

    sheet.appendRow([
      data.submittedAt || new Date().toISOString(),
      data.name || '',
      data.instagram || '',
      data.phone || '',
      data.email || '',
      data.visitDate || '',
      data.type || '',
      data.fee || ''
    ]);

    return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
