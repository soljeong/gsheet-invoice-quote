/**
 * @OnlyCurrentDoc
 */

// ==========================================
// 1. 메뉴 및 메인 진입점
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('견적서')
    .addItem('선택 견적번호로 PDF 생성', 'generateQuotePdfFromSelection')
    .addToUi();
}

/**
 * 선택된 셀의 견적번호를 기반으로 견적서 PDF를 생성합니다.
 */
function generateQuotePdfFromSelection() {
  const ss = SpreadsheetApp.getActive();
  // 견적서 시트에서만 동작하도록 제한 (또는 활성 시트가 견적서라고 가정)
  const sheet = ss.getSheetByName('견적서');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('"견적서" 시트를 찾을 수 없습니다.');
    return;
  }

  try {
    // 1. 설정 로드
    const settings = getSettings_(ss);

    // 2. 현재 선택된 견적번호 추출 (활성 시트 기준)
    const activeSheet = ss.getActiveSheet();
    if (activeSheet.getName() !== '견적서') {
      throw new Error('"견적서" 시트에서 실행해 주세요.');
    }
    const quoteNo = getSelectedQuoteNo_(activeSheet);

    // 3. 견적 데이터 수집 (3개 테이블 조인)
    const quoteData = getQuoteData_(ss, quoteNo);

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

function getQuoteData_(ss, quoteNo) {
  // 1. 시트 가져오기
  const shQuote = ss.getSheetByName('견적서');
  const shItem = ss.getSheetByName('견적품목');
  const shProcess = ss.getSheetByName('견적공정');

  if (!shQuote || !shItem || !shProcess) {
    throw new Error('필수 시트(견적서, 견적품목, 견적공정)가 누락되었습니다.');
  }

  // 2. 헬퍼 함수: 시트 데이터 읽기 및 헤더 매핑
  const readSheet = (sheet) => {
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];
    const headers = values[0].map(h => String(h).trim());
    return values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  };

  const quotes = readSheet(shQuote);
  const items = readSheet(shItem);
  const processes = readSheet(shProcess);

  // 3. 견적서(Quote) 찾기
  // 컬럼 매핑: 견적번호, 등록일(->견적일), 고객사(->업체명), 고객사담당자(->수신자)
  const quoteRow = quotes.find(r => String(r['견적번호'] || '').trim() === quoteNo);
  if (!quoteRow) throw new Error(`견적번호 "${quoteNo}"에 해당하는 견적서 데이터가 없습니다.`);

  const header = {
    NO: String(quoteRow['견적번호'] || ''),
    DATE: quoteRow['등록일'],      // AppSheet 컬럼명 '등록일'
    COMPANY: quoteRow['고객사'],   // AppSheet 컬럼명 '고객사'
    TO: quoteRow['고객사담당자']   // AppSheet 컬럼명 '고객사담당자'
  };

  // 4. 견적품목(Item) 필터링 & 매핑
  // "견적번호" 컬럼으로 필터링
  const relatedItems = items.filter(r => String(r['견적번호'] || '').trim() === quoteNo);

  // 품목ID -> { 품명, 규격 } 맵 생성
  const itemMap = {};
  relatedItems.forEach(r => {
    const id = r['품목ID']; // AppSheet Key
    if (id) {
      itemMap[id] = {
        name: String(r['품명'] || ''),
        spec: String(r['견적설명'] || '') // '규격' 대신 '견적설명' 사용
      };
    }
  });

  // 5. 견적공정(Process) 필터링 및 조인
  // 품목ID가 위에서 찾은 relatedItems에 포함되는 공정만 필터링
  const joinedItems = [];

  // 공정 데이터 순회
  processes.forEach(p => {
    const itemId = p['품목ID'];
    const parentItem = itemMap[itemId];

    if (parentItem) {
      // 데이터 결합
      const processName = String(p['공정'] || '');
      const qty = Number(p['수량'] || 0);
      const unit = Number(p['단가'] || 0);
      const note = String(p['비고'] || '');

      joinedItems.push({
        name: `${parentItem.name} ${processName}`, // 품명 + 공정명
        spec: parentItem.spec,
        qty: qty,
        unit: unit,
        amount: qty * unit, // 단순 계산
        note: note,
        originalItemName: parentItem.name // 정렬용
      });
    }
  });

  if (joinedItems.length === 0) {
    throw new Error(`견적번호 "${quoteNo}"에 해당하는 공정 데이터(상세 품목)가 없습니다.`);
  }

  // 6. 정렬: 품명 가나다순 -> 공정명 순? (일단 품명 기준으로 정렬)
  joinedItems.sort((a, b) => {
    const nameA = a.originalItemName;
    const nameB = b.originalItemName;
    return nameA.localeCompare(nameB, 'ko');
  });

  return { header, items: joinedItems };
}

function buildRenderData_({ header, items }, supplier) {
  // 이미 items는 getQuoteData_에서 계산되어 넘어옴

  // 합계 계산
  const supply = items.reduce((sum, item) => sum + item.amount, 0);
  const vat = Math.floor(supply * 0.1);
  const total = supply + vat;

  return {
    header: {
      NO: header.NO,
      견적일: header.DATE,
      공상호: header.COMPANY,
      수신처: header.TO
    },
    items,
    supply,
    vat,
    total,
    supplier
  };
}

function getSettings_(ss) {
  const sheet = ss.getSheetByName('설정'); // 설정 시트는 그대로 유지된다고 가정 (또는 _Per User Settings 확인 필요하지만 기존 사용)
  // 기존 코드 유지
  if (!sheet) {
    // 설정 시트가 없으면 빈 값 반환하거나 에러 처리.
    // 여기서는 에러 없이 빈 객체 반환하도록 수정 (안전장치)
    return { supplier: {} };
  }

  const data = sheet.getDataRange().getValues();
  const map = {};

  data.forEach(row => {
    const key = String(row[0] || '').trim();
    if (key) map[key] = row[1];
  });

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
      fax: map['공급자_팩스'] || ''
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
