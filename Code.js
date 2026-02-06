/**
 * @OnlyCurrentDoc
 */

// ==========================================
// 1. 메뉴 및 메인 진입점
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('견적서')
    .addItem('PDF 저장', 'generateQuotePdfFromSelection')
    .addItem('미리보기', 'showPreview')
    .addToUi();
}

/**
 * 선택된 셀의 견적번호를 기반으로 견적서 PDF를 생성합니다.
 */
function generateQuotePdfFromSelection() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  try {
    // 1. 설정 로드
    const settings = getSettings_(ss);

    // 2. 현재 선택된 견적번호 추출
    const quoteNo = getSelectedQuoteNo_(sh);

    // 3. 견적 데이터 수집 및 정렬
    const quoteData = getQuoteData_(sh, quoteNo);

    // 4. 데이터 병합 (헤더 + 아이템 + 계산 + 공급자 정보)
    const data = buildRenderData_(quoteData, settings.supplier);

    // 5. PDF 생성 및 저장
    const pdfFile = createPdf_(data, quoteNo);

    // 6. 결과 모달 표시
    showLinkModal_(pdfFile.getUrl());

  } catch (e) {
    SpreadsheetApp.getUi().alert(`오류 발생: ${e.message}`);
    console.error(e.stack);
  }
}

/**
 * 선택된 셀의 견적번호를 기준으로 템플릿만 미리보기합니다.
 */
function showPreview() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  try {
    const settings = getSettings_(ss);
    const quoteNo = getSelectedQuoteNo_(sh);
    const quoteData = getQuoteData_(sh, quoteNo);
    const data = buildRenderData_(quoteData, settings.supplier);

    const template = HtmlService.createTemplateFromFile('template');
    template.data = data;

    const htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setWidth(1200)
      .setHeight(900);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '견적서 미리보기');
  } catch (e) {
    SpreadsheetApp.getUi().alert(`오류 발생: ${e.message}`);
    console.error(e.stack);
  }
}

// ==========================================
// 2. 데이터 처리 로직
// ==========================================

function getSelectedQuoteNo_(sheet) {
  const range = sheet.getActiveRange();
  if (!range) throw new Error('선택된 셀이 없습니다.');

  // 헤더 체크
  const headers = sheet.getDataRange().getValues()[0].map(h => String(h).trim());
  const cQuoteNo = headers.indexOf('견적번호');

  if (cQuoteNo === -1) throw new Error('"견적번호" 열을 찾을 수 없습니다.');

  // 선택된 행의 견적번호 가져오기
  const userRowIndex = range.getRow(); // 1-based
  if (userRowIndex < 2) throw new Error('데이터 행을 선택해주세요 (헤더 제외).');

  const val = sheet.getRange(userRowIndex, cQuoteNo + 1).getValue();
  const quoteNo = String(val || '').trim();

  if (!quoteNo) throw new Error('선택된 행에 견적번호가 비어있습니다.');
  return quoteNo;
}

function getQuoteData_(sheet, quoteNo) {
  const allValues = sheet.getDataRange().getValues();
  if (allValues.length < 2) throw new Error('데이터가 없습니다.');

  const headers = allValues[0].map(h => String(h).trim());
  const col = (name) => {
    const idx = headers.indexOf(name);
    if (idx === -1) throw new Error(`필수 열 "${name}"이(가) 없습니다.`);
    return idx;
  };

  // 컬럼 매핑
  const C = {
    NO: col('견적번호'),
    DATE: col('견적일'),
    COMPANY: col('업체명'),
    TO: col('담당자'),
    ITEM: col('품명'),
    SPEC: col('공정'),
    QTY: col('수량'),
    UNIT: col('단가'),
    NOTE: col('비고')
  };
   
  // 해당 견적번호 행 필터링
  const rows = allValues.slice(1).filter(r => String(r[C.NO] || '').trim() === quoteNo);

  if (rows.length === 0) throw new Error(`견적번호 "${quoteNo}"에 해당하는 데이터가 없습니다.`);

  // 품명 기준 정렬 (가나다순)
  rows.sort((a, b) => {
    const nameA = String(a[C.ITEM] || '').trim();
    const nameB = String(b[C.ITEM] || '').trim();
    return nameA.localeCompare(nameB, 'ko');
  });

  return { rows, C, firstRow: rows[0] };
}

function buildRenderData_({ rows, C, firstRow }, supplier) {
  // 헤더 정보
  const header = {
    NO: String(firstRow[C.NO]),
    견적일: firstRow[C.DATE],
    공상호: firstRow[C.COMPANY],
    수신처: firstRow[C.TO]
  };

  // 품목 리스트 (할인 행은 별도 처리)
  const items = [];
  let discountAmount = 0;
  let hasDiscount = false;

  rows.forEach(r => {
    const qty = Number(r[C.QTY] || 0);
    const unit = Number(r[C.UNIT] || 0);
    const name = String(r[C.ITEM] || '').trim();
    const item = {
      name,
      spec: String(r[C.SPEC] || ''),
      qty,
      unit,
      amount: qty * unit,
      note: String(r[C.NOTE] || '')
    };

    if (name === '할인') {
      hasDiscount = true;
      discountAmount += item.amount;
      return;
    }

    items.push(item);
  });

  // 합계 계산 (할인 금액을 VAT 계산 전에 반영)
  const supply = items.reduce((sum, item) => sum + item.amount, 0) + discountAmount;
  const vat = Math.floor(supply * 0.1); // 부가세: 절사 or 반올림 정책에 따라 조정 (여기선 내림/절사 예시)
  const total = supply + vat;

  return {
    header,
    items,
    discountAmount,
    hasDiscount,
    supply,
    vat,
    total,
    supplier
  };
}

function getSettings_(ss) {
  const sheet = ss.getSheetByName('설정');
  if (!sheet) throw new Error('"설정" 시트가 필요합니다.');

  const data = sheet.getDataRange().getValues();
  const map = {};

  data.forEach(row => {
    const key = String(row[0] || '').trim();
    if (key) map[key] = row[1];
  });

  // PDF 저장 폴더에서 직인 이미지 파일 찾기
  let sealImageData = '';
  
  try {
    const rootFolder = DriveApp.getRootFolder();
    const folderIterator = rootFolder.getFoldersByName('견적서');
    
    if (folderIterator.hasNext()) {
      const folder = folderIterator.next();
      const files = folder.getFilesByName('seal.jpeg');
      
      if (files.hasNext()) {
        const file = files.next();
        const blob = file.getBlob();
        const base64 = Utilities.base64Encode(blob.getBytes());
        const mimeType = blob.getContentType();
        sealImageData = `data:${mimeType};base64,${base64}`;
        Logger.log('직인 이미지 로드 완료. MimeType: ' + mimeType);
      } else {
        Logger.log('경고: 견적서 폴더에서 seal.jpeg를 찾을 수 없습니다.');
      }
    } else {
      Logger.log('경고: 견적서 폴더가 없습니다.');
    }
  } catch (e) {
    Logger.log('직인 이미지 로드 실패: ' + e.message);
  }

  // 필수값 체크 생략(없으면 빈값)
  return {
    supplier: {
      company: map['공급자_상호'] || '',
      ceo: map['공급자_대표자'] || '',
      regNo: map['공급자_등록번호'] || '',
      address: map['공급자_사업장주소'] || '',
      bizType: map['공급자_업태'] || '',
      bizItem: map['공급자_종목'] || '',
      contact: map['공급자_연락처'] || '',
      email: map['공급자_이메일'] || '',
      fax: map['공급자_팩스'] || '',
      sealImage: sealImageData
    }
  };
}

// ==========================================
// 3. PDF 생성 및 Drive 저장
// ==========================================

function createPdf_(data, quoteNo) {
  const template = HtmlService.createTemplateFromFile('template');
  template.data = data;

  const htmlOutput = template.evaluate();

  // HTML을 PDF Blob으로 변환
  const pdfBlob = htmlOutput.getBlob().getAs(MimeType.PDF).setName(`${quoteNo}.pdf`);

  // 폴더 지정 및 저장 (덮어쓰기 로직 포함)
  const folder = getOrCreateSubfolder_(DriveApp.getRootFolder(), '견적서');
  return overwriteFileInFolder_(folder, pdfBlob);
}

function getOrCreateSubfolder_(parent, name) {
  const iterator = parent.getFoldersByName(name);
  if (iterator.hasNext()) return iterator.next();
  return parent.createFolder(name);
}

function overwriteFileInFolder_(folder, blob) {
  const name = blob.getName();
  const files = folder.getFilesByName(name);
  while (files.hasNext()) {
    files.next().setTrashed(true); // 기존 파일 휴지통으로 이동
  }
  return folder.createFile(blob);
}

// ==========================================
// 4. UI 및 유틸
// ==========================================

function showLinkModal_(url) {
  const tpl = HtmlService.createTemplateFromFile('pdf-link');
  tpl.pdfUrl = url;

  const html = tpl.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(400)
    .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, 'PDF 생성 완료');
}
