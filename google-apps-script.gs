/**
 * 구로구 AI 마케팅 인턴십 파트너 신청서 데이터 처리
 * Google Sheets에 자동으로 데이터를 저장하는 Google Apps Script
 */

// 스프레드시트 ID와 시트 이름을 여기에 설정하세요
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // 스프레드시트 URL에서 /d/ 다음에 있는 ID
const SHEET_NAME = '신청서'; // 데이터를 저장할 시트 이름

/**
 * POST 요청을 처리하는 함수
 * 웹 앱으로 배포할 때 이 함수가 자동으로 실행됩니다
 */
function doPost(e) {
  try {
    // 요청 데이터 파싱
    const data = JSON.parse(e.postData.contents);

    // 스프레드시트 열기
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    // 시트가 없으면 생성하고 헤더 추가
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      const headers = [
        '제출일시',
        '회사명',
        '담당자 성함',
        '연락처',
        '이메일',
        '업종',
        '희망 인턴 수',
        '정규직 전환 의향'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // 헤더 스타일 설정
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1e3a8a');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');

      // 열 너비 자동 조정
      for (let i = 1; i <= headers.length; i++) {
        sheet.autoResizeColumn(i);
      }
    }

    // 새 행에 데이터 추가
    const newRow = [
      data.timestamp || new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
      data.companyName || '',
      data.contactName || '',
      data.phone || '',
      data.email || '',
      data.industry || '',
      data.internCount || '',
      data.conversion || ''
    ];

    sheet.appendRow(newRow);

    // 성공 응답
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'success',
      'message': '신청서가 성공적으로 제출되었습니다.'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // 에러 로깅
    Logger.log('Error: ' + error.toString());

    // 에러 응답
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'error',
      'message': '신청서 제출 중 오류가 발생했습니다: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET 요청을 처리하는 함수 (테스트용)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    'status': 'success',
    'message': '구로구 AI 마케팅 인턴십 파트너 신청서 API가 정상 작동 중입니다.'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 스프레드시트 초기 설정 함수 (수동 실행용)
 * Apps Script 편집기에서 이 함수를 실행하여 스프레드시트를 초기화할 수 있습니다
 */
function setupSpreadsheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  // 헤더 설정
  const headers = [
    '제출일시',
    '회사명',
    '담당자 성함',
    '연락처',
    '이메일',
    '업종',
    '희망 인턴 수',
    '정규직 전환 의향'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 헤더 스타일
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1e3a8a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 열 너비 자동 조정
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // 데이터 영역 테두리 설정
  sheet.getRange(1, 1, 1000, headers.length).setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );

  Logger.log('스프레드시트 초기 설정이 완료되었습니다.');
}
