/**
 * code.gs
 * 메인 로직: 폼 제출 트리거 처리, OCR 호출, 파싱, 카테고리 분류
 */

// --- 상수 및 설정 ---
const UPSTAGE_API_URL = 'https://api.upstage.ai/v1/document-digitization';
const SOLAR_API_URL = 'https://api.upstage.ai/v1/solar/chat/completions';
const SCRIPT_PROP_KEY = 'UPSTAGE_API_KEY';

/**
 * API Key 설정 (메뉴에서 호출)
 */
function promptForApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Upstage API Key 설정',
    'Upstage Console에서 발급받은 API Key를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );

  const button = result.getSelectedButton();
  const text = result.getResponseText();

  if (button == ui.Button.OK) {
    if (text) {
      setUpstageApiKey(text);
      ui.alert('API Key가 저장되었습니다.');
    } else {
      ui.alert('API Key가 입력되지 않았습니다.');
    }
  }
}

function setUpstageApiKey(key) {
  PropertiesService.getScriptProperties().setProperty(SCRIPT_PROP_KEY, key);
}

function getUpstageApiKey_() {
  return PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY);
}

/**
 * 트리거 함수: 폼 제출 시 실행
 */
function onFormSubmit(e) {
  console.log('onFormSubmit triggered');
  if (!e) {
    console.error('No event object passed to onFormSubmit');
    return;
  }
  
  try {
    const range = e.range;
    const row = range.getRow();
    console.log(`Processing row: ${row}`);
    processSingleResponse_(row);
  } catch (error) {
    console.error(`Error in onFormSubmit: ${error.toString()}`);
  }
}

/**
 * 미처리 응답 일괄 처리 (수동 실행용)
 */
function processAllUnprocessed() {
  console.log('Starting batch processing...');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 첫 번째 시트가 폼 응답이라고 가정
  const lastRow = sheet.getLastRow();
  const data = sheet.getDataRange().getValues();
  
  // 헤더 제외하고 2행부터 확인
  // 주의: 폼 응답 시트에는 '상태' 컬럼이 없을 수 있음. 
  // 여기서는 단순히 모든 행을 순회하며, 이미 처리된 내역을 별도로 기록하지 않는다면 중복 처리될 수 있음.
  // PRD에 따르면 "Setup.gs는 이 시트를 절대 수정하지 않음"이라고 되어 있으므로,
  // 처리 여부를 판단하기 어렵습니다. 
  // 따라서, 여기서는 사용자가 UI에서 특정 행을 지정하거나, 
  // 별도의 로직으로 처리 여부를 판단해야 하지만, 
  // 일단 요청대로 "미처리 응답 모두 처리"는 순차적으로 실행하도록 구현합니다.
  // 실무적으로는 처리된 행의 ID를 어딘가에 저장하거나, 결과 시트에 해당 응답행번호가 있는지 확인해야 합니다.
  
  // 여기서는 결과 시트들을 확인하여 이미 처리된 행인지 체크하는 로직을 추가합니다.
  const processedRows = getProcessedRowNumbers_();
  
  for (let i = 1; i < data.length; i++) { // 0-indexed, skip header
    const rowNum = i + 1;
    if (!processedRows.has(rowNum)) {
      try {
        processSingleResponse_(rowNum);
      } catch (e) {
        console.error(`Row ${rowNum} processing failed: ${e.message}`);
      }
    }
  }
  return '처리 완료';
}

/**
 * 이미 처리된 응답 행 번호 집합 반환
 */
function getProcessedRowNumbers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const processed = new Set();
  const sheets = ss.getSheets();
  
  // 첫 번째 시트(폼 응답)와 Config 시트를 제외한 나머지 시트(카테고리 시트)를 확인
  const exclude = [sheets[0].getName(), 'Config'];
  
  sheets.forEach(sheet => {
    if (exclude.includes(sheet.getName())) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    
    // '응답행번호' 컬럼 위치 찾기 (헤더에서)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idx = headers.indexOf('응답행번호');
    
    if (idx > -1) {
      const data = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues();
      data.forEach(r => {
        if (r[0]) processed.add(Number(r[0]));
      });
    }
  });
  
  return processed;
}

/**
 * 단일 응답 처리 메인 로직
 */
function processSingleResponse_(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheets()[0]; // 폼 응답 시트
  
  // 데이터 가져오기 (행 전체)
  const rowData = formSheet.getRange(rowNumber, 1, 1, formSheet.getLastColumn()).getValues()[0];
  
  // 필드 매핑 (폼 질문 순서에 따라 인덱스가 달라질 수 있음. 여기서는 PRD 기준 추정)
  // Timestamp(0), 제출자명(1), 학교명(2), 이수증 이미지(3) ...
  // 실제 폼 구조에 맞춰 인덱스 조정 필요. 여기서는 0, 1, 2, 3으로 가정.
  const timestamp = rowData[0];
  const submitter = rowData[1];
  const school = rowData[2];
  const fileUrls = rowData[3]; // 콤마로 구분된 URL 또는 줄바꿈
  
  if (!fileUrls) {
    console.warn(`Row ${rowNumber}: No file URLs found.`);
    return;
  }
  
  // 파일 ID 추출 (여러 개일 수 있음)
  const fileIds = extractFileIds_(fileUrls);
  
  // Config 로드
  const config = loadCategoryConfig_();
  
  fileIds.forEach(fileId => {
    let status = 'OK';
    let ocrText = '';
    let parsedData = {};
    
    console.log(`Processing fileId: ${fileId}`);
    
    try {
      // 1. OCR 호출
      ocrText = callUpstageOcr_(fileId);
      if (!ocrText) throw new Error('OCR returned empty text');
      console.log('OCR success');
      
      // 2. 파싱 (정규식)
      parsedData = parseCertificateText_(ocrText);
      
      // 3. 파싱 검증 및 LLM 보완
      if (isParsingIncomplete_(parsedData)) {
        console.log('Regex parsing incomplete, trying LLM...');
        // LLM 호출
        try {
          const llmData = parseWithSolar_(ocrText);
          // 기존 파싱 데이터에 병합 (LLM이 더 정확하다고 가정하거나, 빈 값만 채움)
          parsedData = { ...parsedData, ...llmData };
          console.log('LLM success');
        } catch (e) {
          console.error('LLM Parsing failed:', e);
          // LLM 실패해도 정규식 결과라도 있으면 진행, 없으면 NEEDS_REVIEW
          if (isParsingIncomplete_(parsedData)) {
             status = 'NEEDS_REVIEW';
          }
        }
      }
      
    } catch (e) {
      console.error(`OCR/Parsing Error for file ${fileId}:`, e);
      status = 'OCR_FAIL';
      ocrText = e.message; // 에러 메시지 저장
    }
    
    // 4. 카테고리 분류
    const categorySheetName = classifyCategory_(parsedData, ocrText, config);
    console.log(`Classified as: ${categorySheetName}`);
    
    // 5. 시트에 저장
    saveToSheet_(categorySheetName, {
      timestamp,
      submitter,
      school,
      parsedData,
      fileId,
      ocrText,
      rowNumber,
      status
    });
  });
}

/**
 * 파일 URL 문자열에서 ID 추출
 */
function extractFileIds_(urlStr) {
  // Google Form 업로드 URL 포맷: https://drive.google.com/open?id=...
  // 콤마 등으로 구분될 수 있음
  const ids = [];
  const regex = /id=([a-zA-Z0-9_-]+)/g;
  let match;
  while ((match = regex.exec(urlStr)) !== null) {
    ids.push(match[1]);
  }
  return ids;
}

/**
 * 디버깅용: 특정 파일 ID로 OCR 테스트
 */
function testOcr() {
  const fileId = '14HNkvZDwWSmmmCdBCsZDLrGbBwS_ab5g'; // 스크린샷에서 추출한 ID
  console.log(`Testing OCR with fileId: ${fileId}`);
  
  try {
    const text = callUpstageOcr_(fileId);
    console.log('OCR Result:', text);
    
    const parsed = parseCertificateText_(text);
    console.log('Parsed Data:', JSON.stringify(parsed, null, 2));
    
    const config = loadCategoryConfig_();
    const category = classifyCategory_(parsed, text, config);
    console.log('Classified Category:', category);
    
  } catch (e) {
    console.error('Test failed:', e);
  }
}

/**
 * Upstage OCR API 호출
 */
function callUpstageOcr_(fileId) {
  const apiKey = getUpstageApiKey_();
  if (!apiKey) throw new Error('API Key not set');
  
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  
  const payload = {
    document: blob,
    ocr: 'auto',
    model: 'document-parse', // 모델명 확인
    output_formats: JSON.stringify(['text']) 
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`
    },
    payload: payload,
    muteHttpExceptions: true
  };
  
  console.log('Sending request to Upstage API...');
  const response = UrlFetchApp.fetch(UPSTAGE_API_URL, options);
  const code = response.getResponseCode();
  const content = response.getContentText();
  
  console.log(`Response Code: ${code}`);
  console.log(`Response Content (First 500 chars): ${content.substring(0, 500)}`);
  
  if (code !== 200) {
    throw new Error(`Upstage API Error: ${code} - ${content}`);
  }
  
  const result = JSON.parse(content);
  console.log('Full API Response:', JSON.stringify(result, null, 2)); // 전체 응답 구조 확인
  
  // text 필드가 없으면 content.text나 content.html 등 다른 필드 확인
  if (result.content && result.content.text) return result.content.text;
  if (result.text) return result.text;
  
  return ''; 
}

/**
 * 정규식 기반 파싱
 */
function parseCertificateText_(text) {
  // 키워드 사이에 공백/줄바꿈이 있을 수 있으므로 \s*를 적극 활용
  // 예: "생 년 월 일" 또는 "성\n명" 등 대응
  
  const data = {
    issueNumber: extractByRegex_(text, /제\s*([\s\S]+?)\s*호/), // 줄바꿈 포함 가능성
    name: extractByRegex_(text, /성\s*명\s*[:：]?\s*([^\n]+)/),
    birthDate: extractByRegex_(text, /생\s*년\s*월\s*일\s*[:：]?\s*([0-9.\-\s]+)/),
    
    // 과정명은 '성명'이나 '생년월일' 키워드가 나오기 전까지의 모든 텍스트 (줄바꿈 포함)
    courseName: extractByRegex_(text, /과\s*정\s*명\s*[:：]?\s*([\s\S]+?)(?=\n\s*성\s*명|\n\s*생\s*년\s*월\s*일|\n\s*직\s*급)/),
    
    period: extractByRegex_(text, /연\s*수\s*기\s*간\s*[:：]?\s*([0-9.~\s]+)/),
    type: extractByRegex_(text, /연\s*수\s*종\s*류\s*[:：]?\s*([^\n]+)/),
    hours: extractByRegex_(text, /이\s*수\s*시\s*간\s*[:：]?\s*([^\n]+)/)
  };
  
  // 데이터 정제 (불필요한 줄바꿈/공백 제거)
  for (const key in data) {
    if (data[key]) {
      data[key] = data[key].replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
    }
  }
  
  return data;
}

function extractByRegex_(text, regex) {
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

/**
 * 파싱 결과 불완전 여부 확인
 */
function isParsingIncomplete_(data) {
  // 필수 필드: 이수번호, 성명, 과정명
  return !data.issueNumber || !data.name || !data.courseName;
}

/**
 * Solar LLM 파싱 호출
 */
function parseWithSolar_(text) {
  const apiKey = getUpstageApiKey_();
  
  const prompt = `
    다음 이수증 텍스트에서 정보를 추출하여 JSON 형식으로 반환해줘.
    필드명: issueNumber(이수번호), name(성명), birthDate(생년월일), courseName(과정명), period(연수기간), type(연수종류), hours(이수시간)
    값이 없으면 빈 문자열로 둬.
    
    텍스트:
    ${text}
  `;
  
  const payload = {
    model: 'solar-pro-2', // 모델명 확인 필요 (solar-pro2 or solar-1-mini-chat etc. PRD says solar-pro2)
    messages: [
      { role: 'system', content: 'You are a helpful assistant that extracts information from certificates to JSON.' },
      { role: 'user', content: prompt }
    ],
    response_format: { type: 'json_object' }
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(SOLAR_API_URL, options);
  const content = response.getContentText();
  const json = JSON.parse(content);
  
  if (json.error) throw new Error(json.error.message);
  
  const contentStr = json.choices[0].message.content;
  return JSON.parse(contentStr);
}

/**
 * Config 로드
 */
function loadCategoryConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  // 헤더 제외, [시트이름, 키워드스트링]
  const config = [];
  for (let i = 1; i < data.length; i++) {
    config.push({
      sheetName: data[i][0],
      keywords: data[i][1].split(',').map(k => k.trim()).filter(k => k)
    });
  }
  return config;
}

/**
 * 카테고리 분류
 */
function classifyCategory_(parsedData, ocrText, configList) {
  // 1. 연수종류나 과정명에서 키워드 검색
  const targetText = (parsedData.type + ' ' + parsedData.courseName + ' ' + ocrText).toLowerCase();
  
  for (const cfg of configList) {
    if (cfg.sheetName === '주제미지정') continue;
    
    for (const keyword of cfg.keywords) {
      if (targetText.includes(keyword.toLowerCase())) {
        return cfg.sheetName;
      }
    }
  }
  
  return '주제미지정';
}

/**
 * 시트에 저장
 */
function saveToSheet_(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  // 시트가 없으면 생성 (Setup 로직 재사용 권장하지만 여기선 간단히 생성)
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // 헤더 추가 등은 Setup.gs에 위임하거나 여기서 처리해야 함.
    // 안전을 위해 Setup.gs의 함수를 호출할 수도 있지만, 의존성 문제로 간단히 처리
    sheet.appendRow(['Timestamp', '제출자명', '학교명', '생년월일', '이수번호', '과정명', '연수종류', '연수기간', '이수시간', '원본이미지링크', 'OCR원문', '응답행번호', '상태']);
  }
  
  const row = [
    data.timestamp,
    data.submitter,
    data.school,
    data.parsedData.birthDate || '',
    data.parsedData.issueNumber || '',
    data.parsedData.courseName || '',
    data.parsedData.type || '',
    data.parsedData.period || '',
    data.parsedData.hours || '',
    `https://drive.google.com/file/d/${data.fileId}/view`,
    data.ocrText.substring(0, 4000), // 셀 제한 고려
    data.rowNumber,
    data.status
  ];
  
  sheet.appendRow(row);
}
