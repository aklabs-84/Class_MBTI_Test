function doPost(e) {
  try {
    // 스프레드시트 ID (URL에서 추출) - 실제 스프레드시트 ID로 변경 필요
    const SPREADSHEET_ID = '1WI20EWfNEgDa_UmvboQr09KXxcUdZpx8b4ZosvO4aZg';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName('MBTI_Results') || ss.insertSheet('MBTI_Results');

    // 헤더 설정 (첫 실행시에만)
    if (sh.getLastRow() === 0) {
      // 헤더 항목들
      sh.getRange(1,1,1,5).setValues([['타임스탬프','이름','MBTI 유형','특징','점수 분포']]);

      // 헤더 스타일링
      const headerRange = sh.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#667eea');
      headerRange.setFontColor('white');
      headerRange.setHorizontalAlignment('center');
      
      // 열 너비 자동 조정
      sh.autoResizeColumns(1, 5);
    }

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

    // 새 행에 데이터 추가
    sh.appendRow([
      data.timestamp || new Date().toLocaleString('ko-KR'),
      data.name || '',
      data.resultTitle || '',
      data.resultContent || '',
      data.score || ''
    ]);

    // 데이터 행 스타일링 (선택사항)
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const dataRange = sh.getRange(lastRow, 1, 1, 5);
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

// 테스트 함수 (선택사항 - 수동으로 데이터 추가해볼 때 사용)
function testInsert() {
  const testData = {
    timestamp: new Date().toLocaleString('ko-KR'),
    name: '테스트 유저',
    resultTitle: 'ENFP (재기발랄한 활동가)',
    resultContent: '아이디어 뱅크. 새로운 아이디어와 열정으로 채널에 활기를 불어넣고, 시청자와 소통한다.',
    score: 'E:3 I:0 S:1 N:2 T:1 F:2 J:1 P:2'
  };
  
  // doPost 함수를 시뮬레이션
  const e = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(e);
  Logger.log('Test result: ' + result.getContent());
}

// 스프레드시트 데이터 조회 함수 (선택사항)
function getAllResults() {
  try {
    const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // 실제 ID로 변경
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName('MBTI_Results');
    
    if (!sh) {
      return { status: 'error', message: 'Sheet not found' };
    }
    
    const data = sh.getDataRange().getValues();
    
    return {
      status: 'success',
      data: data,
      totalRows: data.length
    };
    
  } catch (err) {
    return { status: 'error', message: String(err) };
  }
}

// 특정 사용자의 결과만 조회하는 함수 (선택사항)
function getUserResults(userName) {
  try {
    const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // 실제 ID로 변경
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName('MBTI_Results');
    
    if (!sh) {
      return { status: 'error', message: 'Sheet not found' };
    }
    
    const data = sh.getDataRange().getValues();
    const userResults = data.filter((row, index) => {
      if (index === 0) return true; // 헤더 포함
      return row[1] === userName; // 이름 컬럼에서 검색
    });
    
    return {
      status: 'success',
      data: userResults,
      count: userResults.length - 1 // 헤더 제외
    };
    
  } catch (err) {
    return { status: 'error', message: String(err) };
  }
}
