/** MBTI 테스트 날짜별 시트 저장 - Google Apps Script **/

const DEFAULT_HEADERS = ['타임스탬프','이름','MBTI 유형','특징','점수 분포'];
const HEADER_KEYS = {
  name:    ['이름','name'],
  title:   ['MBTI 유형','mbti유형','mbti_type','resulttitle'],
  content: ['특징','특성','feature','resultcontent'],
  score:   ['점수 분포','점수분포','점수','score'],
  ts:      ['타임스탬프','타임스태프','timestamp','시간','작성시각']
};

function norm_(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,''); }

// 날짜 기반 시트명 생성 함수
function getDateSheetName_(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `MBTI_${year}-${month}-${day}`;
}

// 날짜별 시트 가져오기 또는 생성
function getOrCreateDateSheet_(ss, date) {
  const sheetName = getDateSheetName_(date);
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // 새 시트 생성
    sheet = ss.insertSheet(sheetName);
    
    // 시트 탭 색상 설정 (MBTI용 색상)
    const colors = ['#667eea', '#764ba2', '#ff9a9e', '#fecfef', '#a8edea', '#fed6e3'];
    const colorIndex = Math.floor(Math.random() * colors.length);
    sheet.setTabColor(colors[colorIndex]);
  }
  
  return sheet;
}

function ensureHeaders_(sh){
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,DEFAULT_HEADERS.length).setValues([DEFAULT_HEADERS]);
    
    // 헤더 스타일링
    const headerRange = sh.getRange(1, 1, 1, DEFAULT_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#667eea');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    
    // 열 너비 자동 조정
    sh.autoResizeColumns(1, DEFAULT_HEADERS.length);
    return;
  }
  
  const lastCol = sh.getLastColumn();
  const row1 = sh.getRange(1,1,1,lastCol).getValues()[0];
  const normRow = row1.map(norm_);

  // 각 키의 대표 헤더가 없으면 추가, 비표준 변형이면 표준 명칭으로 교체
  const want = {
    name: '이름', title: 'MBTI 유형', content: '특징', score: '점수 분포', ts: '타임스탬프'
  };
  Object.entries(HEADER_KEYS).forEach(([k, alts])=>{
    const idx = normRow.findIndex(h => alts.includes(h));
    if (idx === -1) {
      // 맨 뒤에 새로 추가
      sh.getRange(1, sh.getLastColumn()+1).setValue(want[k]);
      sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
    } else {
      // 표준 명칭으로 교체
      sh.getRange(1, idx+1).setValue(want[k]);
    }
  });
}

function findCol_(sh, key){ // key in HEADER_KEYS
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const normHeaders = headers.map(norm_);
  const alts = HEADER_KEYS[key];
  const idx = normHeaders.findIndex(h => alts.includes(h));
  return idx === -1 ? null : idx+1; // 1-based
}

function doPost(e) {
  try {
    // 스프레드시트 ID (실제 ID로 변경 필요)
    const SPREADSHEET_ID = '스프레드시트 ID';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 현재 날짜로 시트 결정
    const currentDate = new Date();
    const sh = getOrCreateDateSheet_(ss, currentDate);

    ensureHeaders_(sh);

    // POST 데이터 파싱
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        Logger.log('JSON parse error: ' + parseErr);
        data = e.parameter || {};
      }
    } else {
      data = e ? (e.parameter || {}) : {};
    }

    Logger.log('Received data: ' + JSON.stringify(data));

    // 값 준비
    const rawTs = data.timestamp || new Date().toLocaleString('ko-KR');

    // 열 찾기
    const nameCol = findCol_(sh,'name');
    const titleCol = findCol_(sh,'title');
    const contentCol = findCol_(sh,'content');
    const scoreCol = findCol_(sh,'score');
    const tsCol = findCol_(sh,'ts');

    const rowLen = sh.getLastColumn();
    const row = new Array(rowLen).fill('');

    if (tsCol)     row[tsCol-1]     = rawTs;
    if (nameCol)   row[nameCol-1]   = data.name || '';
    if (titleCol)  row[titleCol-1]  = data.resultTitle || '';
    if (contentCol)row[contentCol-1]= data.resultContent || '';
    if (scoreCol)  row[scoreCol-1]  = data.score || '';

    sh.appendRow(row);

    // 데이터 행 스타일링
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const dataRange = sh.getRange(lastRow, 1, 1, rowLen);
      dataRange.setBorder(true, true, true, true, true, true);
      
      // 교대로 배경색 적용
      if (lastRow % 2 === 0) {
        dataRange.setBackground('#f8f9fa');
      }
    }

    // 성공 응답
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success', 
        message: '저장 완료!',
        sheetName: sh.getName(),
        row: sh.getLastRow()
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err);

    // 에러 응답
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error', 
        message: String(err)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// OPTIONS 요청 처리 (CORS 대응)
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}

// 디버그용 GET 요청
function doGet() {
  return ContentService
    .createTextOutput('MBTI Apps Script is working! 🚀')
    .setMimeType(ContentService.MimeType.TEXT);
}

// 모든 날짜 시트의 데이터를 통합하여 요약 시트 생성
function createMBTISummarySheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheetName = 'MBTI_전체요약';
    
    // 기존 요약 시트 삭제 후 새로 생성
    let summarySheet = ss.getSheetByName(summarySheetName);
    if (summarySheet) {
      ss.deleteSheet(summarySheet);
    }
    summarySheet = ss.insertSheet(summarySheetName);
    
    // 요약 시트 헤더 설정
    const summaryHeaders = ['날짜', '타임스탬프', '이름', 'MBTI 유형', '특징', '점수 분포'];
    summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
    summarySheet.getRange(1, 1, 1, summaryHeaders.length)
      .setFontWeight('bold')
      .setBackground('#667eea')
      .setFontColor('white')
      .setHorizontalAlignment('center');
    
    const allSheets = ss.getSheets();
    const dateSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return /^MBTI_\d{4}-\d{2}-\d{2}$/.test(name); // MBTI_YYYY-MM-DD 형식만
    });
    
    let summaryRow = 2;
    
    // 각 날짜 시트에서 데이터 수집
    dateSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const dateOnly = sheetName.replace('MBTI_', ''); // MBTI_ 접두사 제거
      const lastRow = sheet.getLastRow();
      
      if (lastRow > 1) { // 헤더 외에 데이터가 있는 경우
        const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        
        data.forEach(row => {
          if (row.some(cell => cell !== '')) { // 빈 행이 아닌 경우
            const summaryRowData = [dateOnly, ...row];
            summarySheet.getRange(summaryRow, 1, 1, summaryRowData.length).setValues([summaryRowData]);
            summaryRow++;
          }
        });
      }
    });
    
    // 열 너비 자동 조정
    summarySheet.autoResizeColumns(1, summaryHeaders.length);
    
    // 요약 시트를 첫 번째 위치로 이동
    ss.moveSheet(summarySheet, 1);
    
    SpreadsheetApp.getUi().alert(`MBTI 요약 시트가 생성되었습니다!\n총 ${dateSheets.length}개의 날짜 시트에서 데이터를 통합했습니다.`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('요약 시트 생성 중 오류 발생: ' + error.toString());
  }
}

// 특정 날짜의 MBTI 데이터 삭제
function deleteMBTIDataByDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('MBTI 날짜별 데이터 삭제', 'YYYY-MM-DD 형식으로 삭제할 날짜를 입력하세요:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const dateInput = response.getResponseText().trim();
    
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateInput)) {
      ui.alert('올바른 날짜 형식(YYYY-MM-DD)을 입력해주세요.');
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `MBTI_${dateInput}`;
    const sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      const confirmResponse = ui.alert(`${sheetName} 시트를 삭제하시겠습니까?`, ui.ButtonSet.YES_NO);
      if (confirmResponse == ui.Button.YES) {
        ss.deleteSheet(sheet);
        ui.alert(`${sheetName} 시트가 삭제되었습니다.`);
      }
    } else {
      ui.alert(`${sheetName} 시트를 찾을 수 없습니다.`);
    }
  }
}

// MBTI 유형별 통계 생성
function createMBTIStatsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statsSheetName = 'MBTI_통계';
    
    // 기존 통계 시트 삭제 후 새로 생성
    let statsSheet = ss.getSheetByName(statsSheetName);
    if (statsSheet) {
      ss.deleteSheet(statsSheet);
    }
    statsSheet = ss.insertSheet(statsSheetName);
    
    // 모든 MBTI 날짜 시트에서 데이터 수집
    const allSheets = ss.getSheets();
    const dateSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return /^MBTI_\d{4}-\d{2}-\d{2}$/.test(name);
    });
    
    const mbtiCount = {};
    
    dateSheets.forEach(sheet => {
      const titleCol = findCol_(sheet, 'title');
      if (!titleCol) return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      const titles = sheet.getRange(2, titleCol, lastRow - 1, 1).getValues().flat();
      
      titles.forEach(title => {
        if (title) {
          // MBTI 유형 추출 (예: "ENFP (재기발랄한 활동가)" -> "ENFP")
          const mbtiMatch = String(title).match(/^([A-Z]{4})/);
          if (mbtiMatch) {
            const mbtiType = mbtiMatch[1];
            mbtiCount[mbtiType] = (mbtiCount[mbtiType] || 0) + 1;
          }
        }
      });
    });
    
    // 통계 시트 헤더
    statsSheet.getRange(1, 1).setValue('MBTI 유형');
    statsSheet.getRange(1, 2).setValue('참여자 수');
    statsSheet.getRange(1, 3).setValue('비율(%)');
    
    const headerRange = statsSheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#667eea');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    
    // 데이터 입력
    const totalCount = Object.values(mbtiCount).reduce((sum, count) => sum + count, 0);
    const sortedMBTI = Object.entries(mbtiCount).sort((a, b) => b[1] - a[1]);
    
    sortedMBTI.forEach(([mbtiType, count], index) => {
      const percentage = totalCount > 0 ? Math.round((count / totalCount) * 100) : 0;
      const row = index + 2;
      
      statsSheet.getRange(row, 1).setValue(mbtiType);
      statsSheet.getRange(row, 2).setValue(count);
      statsSheet.getRange(row, 3).setValue(`${percentage}%`);
    });
    
    // 열 너비 자동 조정
    statsSheet.autoResizeColumns(1, 3);
    
    SpreadsheetApp.getUi().alert(`MBTI 통계 시트가 생성되었습니다!\n총 ${totalCount}명의 데이터를 분석했습니다.`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('통계 시트 생성 중 오류 발생: ' + error.toString());
  }
}

// 테스트 함수
function testMBTIInsert() {
  const testData = {
    timestamp: new Date().toLocaleString('ko-KR'),
    name: '테스트 유저',
    resultTitle: 'ENFP (재기발랄한 활동가)',
    resultContent: '아이디어 뱅크. 새로운 아이디어와 열정으로 채널에 활기를 불어넣고, 시청자와 소통한다.',
    score: 'E:3 I:0 S:1 N:2 T:1 F:2 J:1 P:2'
  };
  
  const e = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(e);
  Logger.log('Test result: ' + result.getContent());
}

// 메뉴 추가
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 MBTI 데이터 관리')
    .addItem('📋 전체 요약 시트 생성', 'createMBTISummarySheet')
    .addItem('📈 MBTI 통계 시트 생성', 'createMBTIStatsSheet')
    .addItem('🗑️ 날짜별 데이터 삭제', 'deleteMBTIDataByDate')
    .addItem('🧪 테스트 데이터 추가', 'testMBTIInsert')
    .addToUi();
}
