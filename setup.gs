/**
 * setup.gs
 * 스프레드시트 초기 설정 및 카테고리 시트 생성 스크립트
 */

/**
 * 메뉴 생성 (스프레드시트 열릴 때)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Certificate OCR')
    .addItem('1. 초기 설정 (시트 생성)', 'setupSpreadsheet')
    .addItem('2. 트리거 설정 (자동 실행 켜기)', 'setupTrigger')
    .addItem('Upstage API Key 설정', 'promptForApiKey')
    .addItem('미처리 응답 일괄 처리', 'processAllUnprocessed')
    .addToUi();
}

/**
 * 트리거 자동 설정
 */
function setupTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  // 기존 트리거 확인
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ui.alert('이미 트리거가 설정되어 있습니다.');
      return;
    }
  }
  
  // 새 트리거 생성
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
    
  ui.alert('자동 실행 트리거가 설정되었습니다. 이제 폼 제출 시 자동으로 실행됩니다.');
}

/**
 * 전체 스프레드시트 구조 초기화
 * - Config 시트 생성
 * - 카테고리별 시트 생성
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. Config 시트 생성
  setupConfigSheet(ss);
  
  // 2. 카테고리 시트 생성 (Config에 정의된 내용 기반 + 기본값)
  const categories = getCategoryList(ss);
  categories.forEach(cat => {
    setupCategorySheet(ss, cat);
  });
  
  ui.alert('초기 설정이 완료되었습니다.');
}

/**
 * Config 시트 설정
 */
function setupConfigSheet(ss) {
  const sheetName = 'Config';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // 헤더 설정
  const headers = ['시트이름', '키워드(쉼표구분)'];
  sheet.getRange('A1:B1').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#EEEEEE')
    .setHorizontalAlignment('center');
    
  // 초기 데이터가 비어있으면 기본값 채우기
  if (sheet.getLastRow() <= 1) {
    const defaultData = [
      ['법정의무연수1', '법정의무연수1,법정의무 연수1,법정 의무연수1,법정 의무 연수1,법정의무연수 1,법정의무 1,의무연수1,의무 연수1,법정의무연수l,의무연수I,법적의무연수1'],
      ['법정의무연수2', '법정의무연수2,법정의무 연수2,법정 의무연수2,법정 의무 연수2,법정의무연수 2,법정의무 2,의무연수2,의무 연수2'],
      ['성희롱성폭력', '성희롱,성폭력,성매매,가정폭력'],
      ['아동학대신고의무자', '아동학대,아동학대예방,아동학대신고의무자'],
      ['기초학력', '문해력,수해력,기초학력,기초 학력,학력 증진,기초학력증진'],
      ['학교안전교육', '학교안전,재난안전,생활안전'],
      ['주제미지정', ''] // Fallback
    ];
    sheet.getRange(2, 1, defaultData.length, 2).setValues(defaultData);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Config 시트에서 카테고리 목록 가져오기
 */
function getCategoryList(ss) {
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  // A열(시트이름)만 가져옴
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(r => r[0]).filter(String);
}

/**
 * 개별 카테고리 시트 생성 및 스타일 적용
 */
function setupCategorySheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // 헤더 정의
  const headers = [
    'Timestamp', '제출자명', '학교명', '생년월일', '이수번호', 
    '과정명', '연수종류', '연수기간', '이수시간', '원본이미지링크', 
    'OCR원문', '응답행번호', '상태'
  ];
  
  // 헤더 입력
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 스타일 적용
  headerRange
    .setFontWeight('bold')
    .setBackground('#EEEEEE')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
    
  // 전체 시트 기본 스타일
  const maxCols = headers.length;
  const fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), maxCols);
  
  fullRange
    .setFontFamily('Nanum Gothic') // 또는 Arial
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true); // 기본적으로 줄바꿈 활성화
    
  // OCR 원문 컬럼(11번째)은 줄바꿈 끄기 (너무 길어지는 것 방지)
  // 대신 셀 내용을 잘라서 보여주거나(Clip), 사용자가 더블클릭해서 보도록 유도
  sheet.getRange(1, 11, sheet.getMaxRows(), 1).setWrap(false); 

  // 틀 고정
  sheet.setFrozenRows(1);
  
  // 열 너비 조정
  sheet.setColumnWidth(1, 120); // Timestamp
  sheet.setColumnWidth(6, 200); // 과정명
  sheet.setColumnWidth(10, 250); // 원본이미지링크 (더 넓게)
  sheet.setColumnWidth(11, 400); // OCR원문 (더 넓게)
}
