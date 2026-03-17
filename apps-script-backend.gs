// =============================================
// 계성 동창 집들이 날짜 투표 - Apps Script
// =============================================

// 스프레드시트 시트 이름
const SHEET_NAME = '투표결과';

// 날짜 목록 (index.html과 동일하게)
const DATES = [
  '3/25(수)', '3/26(목)', '3/27(금)', '3/28(토)',
  '3/29(일)', '3/30(월)', '3/31(화)', '4/1(수)',
  '4/2(목)', '4/3(금)', '4/4(토)', '4/5(일)'
];

// ============================================
// GET 요청: 현재 투표 결과 집계 반환
// ============================================
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // 시트가 없으면 생성
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const header = ['타임스탬프', '이름', ...DATES];
    sheet.appendRow(header);
    // 헤더 스타일
    const headerRange = sheet.getRange(1, 1, 1, header.length);
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }

  const data = sheet.getDataRange().getValues();

  // 날짜별 득표수 집계
  const counts = {};
  DATES.forEach(d => counts[d] = 0);

  // 헤더 행 제외하고 집계
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // 3번째 열부터 날짜 선택값
    for (let j = 2; j < row.length; j++) {
      if (row[j] === '○' || row[j] === true || row[j] === 1) {
        const dateKey = DATES[j - 2];
        if (dateKey) counts[dateKey]++;
      }
    }
  }

  // 전체 참여자 수
  const total = data.length > 1 ? data.length - 1 : 0;

  const result = {
    success: true,
    total: total,
    counts: counts
  };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// POST 요청: 새 투표 응답 저장
// ============================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const name = params.name || '익명';
    const selectedDates = params.dates || []; // 선택한 날짜 배열

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // 시트가 없으면 생성
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // 헤더 행이 없으면 생성
    if (sheet.getLastRow() === 0) {
      const header = ['타임스탬프', '이름', ...DATES];
      sheet.appendRow(header);
      // 헤더 스타일
      const headerRange = sheet.getRange(1, 1, 1, header.length);
      headerRange.setBackground('#4a86e8');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
    }

    // 이미 투표한 사람인지 확인 (이름 기준 중복 체크)
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][1] === name) {
        // 기존 행 업데이트
        const rowValues = [new Date(), name, ...DATES.map(d => selectedDates.includes(d) ? '○' : '')];
        sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);

        // ○ 표시된 셀에 초록색 배경
        for (let j = 0; j < DATES.length; j++) {
          const cell = sheet.getRange(i + 1, j + 3);
          if (selectedDates.includes(DATES[j])) {
            cell.setBackground('#b7e1cd'); // 연초록
          } else {
            cell.setBackground('#ffffff'); // 흰색
          }
        }

        return ContentService
          .createTextOutput(JSON.stringify({ success: true, updated: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 새 행 추가
    const rowValues = [new Date(), name, ...DATES.map(d => selectedDates.includes(d) ? '○' : '')];
    sheet.appendRow(rowValues);

    // ○ 표시된 셀에 초록색 배경
    const lastRow = sheet.getLastRow();
    for (let i = 0; i < DATES.length; i++) {
      const cell = sheet.getRange(lastRow, i + 3);
      if (selectedDates.includes(DATES[i])) {
        cell.setBackground('#b7e1cd'); // 연초록
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, updated: false }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
