// 구글 설문지 응답 자동 PDF 생성 및 메일 전송 시스템

// 전역 설정
const RESPONSE_SHEET_NAME = '설문지 응답 시트1';
const DOCUMENT_SHEET_NAME = '문서';
const CHECKBOX_COLUMN = 10; // J열
const AMOUNT_COLUMN = 9;    // I열
const EMAIL_COLUMN = 5;     // E열
const PDF_STATUS_COLUMN = 11; // K열
const MAIL_STATUS_COLUMN = 12; // L열

// 메인 트리거 함수: 셀 편집 감지
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    // 설문지 응답 시트1이고 J열(체크박스)일 때만 실행
    if (sheet.getName() !== RESPONSE_SHEET_NAME || range.getColumn() !== CHECKBOX_COLUMN) {
      return;
    }
    
    const row = range.getRow();
    const isChecked = range.getValue();
    
    // 체크박스가 체크되었을 때만 실행
    if (isChecked === true) {
      console.log(`체크박스 체크됨 - 행: ${row}`);
      processDataAndGeneratePDF(row);
    }
    
  } catch (error) {
    console.error('체크박스 변경 처리 중 오류:', error);
  }
}

// 새 응답에 체크박스 자동 생성
function createCheckboxForNewResponse() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!sheet) {
      console.error('설문지 응답 시트1을 찾을 수 없습니다');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    
    // 2행부터 시작 (1행은 헤더)
    for (let row = 2; row <= lastRow; row++) {
      const checkboxCell = sheet.getRange(row, CHECKBOX_COLUMN);
      
      // 체크박스가 없고 빈 행이 아닌 경우에만 체크박스 생성
      if (!checkboxCell.isChecked() && sheet.getRange(row, 1).getValue()) {
        checkboxCell.insertCheckboxes();
        console.log(`체크박스 생성 완료 - 행: ${row}`);
      }
    }
    
  } catch (error) {
    console.error('체크박스 생성 중 오류:', error);
  }
}

// 데이터 매핑 및 PDF 생성
function processDataAndGeneratePDF(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    const documentSheet = ss.getSheetByName(DOCUMENT_SHEET_NAME);
    
    if (!responseSheet || !documentSheet) {
      throw new Error('필요한 시트를 찾을 수 없습니다');
    }
    
    // 진행 중 상태 표시
    updateStatus(rowIndex, PDF_STATUS_COLUMN, '처리 중...', '#FFFF00');
    
    // 응답 데이터 추출
    const rowData = responseSheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
    
    // 필수 데이터 검증
    if (!validateRequiredData(rowData)) {
      throw new Error('필수 데이터가 누락되었습니다');
    }
    
    // 데이터 매핑 (PRD 기준)
    documentSheet.getRange('C8').setValue(rowData[2]); // C열 → C8
    documentSheet.getRange('H8').setValue(rowData[3]); // D열 → H8
    documentSheet.getRange('C9').setValue(rowData[5]); // F열 → C9
    documentSheet.getRange('J24').setValue(rowData[8]); // I열 → J24
    
    // PDF 생성
    const pdfBlob = createPDFFromSheet(documentSheet);
    const fileName = `기부영수증_${rowData[2]}_${new Date().getTime()}.pdf`;
    
    // Drive에 저장
    const folder = DriveApp.getRootFolder(); // 또는 특정 폴더 지정
    const file = folder.createFile(pdfBlob.setName(fileName));
    
    updateStatus(rowIndex, PDF_STATUS_COLUMN, 'PDF 저장 완료', '#00FF00');
    console.log(`PDF 생성 완료: ${fileName}`);
    
    // 메일 발송
    const email = rowData[4]; // E열
    sendEmailWithPDF(email, pdfBlob, rowIndex);
    
  } catch (error) {
    console.error(`PDF 생성 중 오류 (행 ${rowIndex}):`, error);
    updateStatus(rowIndex, PDF_STATUS_COLUMN, `PDF 저장 실패: ${error.message}`, '#FF0000');
  }
}

// PDF 첨부 메일 발송
function sendEmailWithPDF(email, pdfBlob, rowIndex) {
  try {
    // 진행 중 상태 표시
    updateStatus(rowIndex, MAIL_STATUS_COLUMN, '메일 전송 중...', '#FFFF00');
    
    // 이메일 형식 검증
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      throw new Error('잘못된 이메일 주소');
    }
    
    const subject = '기부영수증 발급 안내';
    const body = `
안녕하세요.

요청하신 기부영수증을 첨부 파일로 보내드립니다.

감사합니다.
    `;
    
    GmailApp.sendEmail(
      email,
      subject,
      body,
      {
        attachments: [pdfBlob]
      }
    );
    
    updateStatus(rowIndex, MAIL_STATUS_COLUMN, '메일 전송 완료', '#00FF00');
    console.log(`메일 전송 완료: ${email}`);
    
  } catch (error) {
    console.error(`메일 전송 중 오류 (행 ${rowIndex}):`, error);
    updateStatus(rowIndex, MAIL_STATUS_COLUMN, `메일 전송 실패: ${error.message}`, '#FF0000');
  }
}

// 상태 업데이트
function updateStatus(rowIndex, column, status, color) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    const cell = sheet.getRange(rowIndex, column);
    
    cell.setValue(status);
    if (color) {
      cell.setBackground(color);
    }
    
  } catch (error) {
    console.error('상태 업데이트 중 오류:', error);
  }
}

// 필수 데이터 유효성 검증
function validateRequiredData(rowData) {
  // C, D, E, F, I열 데이터 존재 여부 확인
  const requiredFields = [
    { index: 2, name: 'C열 데이터' },
    { index: 3, name: 'D열 데이터' },
    { index: 4, name: 'E열 이메일' },
    { index: 5, name: 'F열 데이터' },
    { index: 8, name: 'I열 금액' }
  ];
  
  for (const field of requiredFields) {
    if (!rowData[field.index] || rowData[field.index] === '') {
      console.error(`${field.name} 누락`);
      return false;
    }
  }
  
  // 금액 데이터 숫자 형식 확인
  if (isNaN(rowData[8])) {
    console.error('I열 금액이 숫자가 아닙니다');
    return false;
  }
  
  return true;
}

// 시트를 PDF로 변환 - 대안 방법
function createPDFFromSheet(sheet) {
  try {
    const spreadsheet = sheet.getParent();
    const sheetId = sheet.getSheetId();
    
    // 방법 1: DriveApp 사용 (권한 문제 회피)
    const tempFile = DriveApp.createFile(
      spreadsheet.getBlob().setName('temp_sheet.xlsx')
    );
    
    // PDF로 변환
    const pdfBlob = tempFile.getAs('application/pdf');
    
    // 임시 파일 삭제
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    
    return pdfBlob;
    
  } catch (error) {
    console.error('DriveApp 방법 실패, UrlFetch 시도:', error);
    
    // 방법 2: 기존 UrlFetch 방법
    return createPDFWithUrlFetch(sheet);
  }
}

// UrlFetch를 사용한 PDF 생성 (백업 방법)
function createPDFWithUrlFetch(sheet) {
  const spreadsheet = sheet.getParent();
  const sheetId = sheet.getSheetId();
  
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?` +
    `exportFormat=pdf&format=pdf&gid=${sheetId}&` +
    `size=A4&portrait=true&fitw=true&top_margin=0.75&bottom_margin=0.75&` +
    `left_margin=0.7&right_margin=0.7&horizontal_alignment=CENTER&` +
    `vertical_alignment=TOP`;
  
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });
  
  return response.getBlob();
}

// 폼 제출 시 자동 체크박스 생성
function onFormSubmit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    
    // 설문지 응답 시트가 아니면 종료
    if (sheet.getName() !== RESPONSE_SHEET_NAME) {
      return;
    }
    
    const row = e.range.getRow();
    console.log(`새 응답 접수 - 행: ${row}`);
    
    // 체크박스 생성
    const checkboxCell = sheet.getRange(row, CHECKBOX_COLUMN);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false); // 기본값: 체크 해제
    
    console.log(`체크박스 생성 완료 - 행: ${row}, J열`);
    
  } catch (error) {
    console.error('폼 제출 처리 중 오류:', error);
  }
}

// 수동 실행용 함수들
function setupTriggers() {
  console.log('onEdit, onFormSubmit 함수는 자동으로 감지됩니다.');
  console.log('기존 응답에 체크박스를 일괄 추가하려면 createCheckboxForNewResponse() 실행');
}

function testFunction() {
  console.log('테스트 함수 실행');
  createCheckboxForNewResponse();
}
