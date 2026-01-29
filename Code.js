function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('견적서')
    .addItem('선택 견적번호로 PDF 생성', 'generateQuotePdfFromSelection')
    .addToUi();
}

function generateQuotePdfFromSelection() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet(); // '견적서' 시트에서 실행한다고 가정 (원하면 이름 고정 가능)
  const settings = getSettings_(ss);

  const range = ss.getActiveRange();
  if (!range) throw new Error('선택된 셀이 없습니다.');

  // 1) 헤더 & 데이터 읽기
  const all = sh.getDataRange().getValues();
  if (all.length < 2) throw new Error('데이터가 없습니다.');

  const headers = all[0].map(h => String(h).trim());
  const col = (name) => headers.indexOf(name);

  const cQuoteNo = col('견적번호');
  const cDate    = col('견적일');
  const cCompany = col('업체명');
  const cTo      = col('수신자');
  const cItem    = col('품명');
  const cSpec    = col('규격');
  const cQty     = col('수량');
  const cUnit    = col('단가');
  const cNote    = col('비고');

  const must = [cQuoteNo, cDate, cCompany, cTo, cItem, cSpec, cQty, cUnit, cNote];
  if (must.some(i => i < 0)) {
    throw new Error('헤더가 정확히 있어야 합니다: 견적번호, 견적일, 업체명, 수신자, 품명, 규격, 수량, 단가, 비고');
  }

  // 2) 선택된 셀에서 견적번호 가져오기
  // - 선택한 셀이 '견적번호' 열이면 그대로 사용
  // - 아니면 선택한 행의 '견적번호' 값을 사용
  const r = range.getRow();
  if (r < 2) throw new Error('헤더(1행)가 아닌 데이터 행을 선택하세요.');

  const selectedQuoteNo = (() => {
    const selectedCol = range.getColumn();
    if (selectedCol === cQuoteNo + 1) return sh.getRange(r, selectedCol).getValue();
    return sh.getRange(r, cQuoteNo + 1).getValue();
  })();

  const quoteNo = String(selectedQuoteNo || '').trim();
  if (!quoteNo) throw new Error('선택된 행에서 견적번호를 찾지 못했습니다. (견적번호 셀이 비어있음)');

  // 3) 같은 견적번호인 행들만 모으기
  const rows = all.slice(1).filter(row => String(row[cQuoteNo] || '').trim() === quoteNo);
  if (rows.length === 0) throw new Error(`견적번호 "${quoteNo}"에 해당하는 행이 없습니다.`);

  // 4) 헤더 정보(첫 행 기준)
  const first = rows[0];
  const header = {
    NO: quoteNo,
    견적일: first[cDate],
    공상호: first[cCompany],   // 업체명 -> 공상호(양식 표기용)
    수신처: first[cTo],         // 수신자 -> 수신처(양식 표기용)
  };

  // 5) 품목 리스트 (품명 기준 정렬)
  const rowsSorted = rows.slice().sort((a, b) => {
    const aName = String(a[cItem] || '').trim();
    const bName = String(b[cItem] || '').trim();
    return aName.localeCompare(bName, 'ko');
  });

  const items = rowsSorted.map(row => {
    const qty = Number(row[cQty] || 0);
    const unit = Number(row[cUnit] || 0);
    return {
      name: row[cItem] || '',
      spec: row[cSpec] || '',
      qty,
      unit,
      amount: qty * unit,
      note: row[cNote] || ''
    };
  });

  // (선택) 합계/부가세
  const supply = items.reduce((sum, it) => sum + it.amount, 0);
  const vat = Math.round(supply * 0.1);
  const total = supply + vat;
  
  // 6) HTML → PDF 생성
  const tpl = HtmlService.createTemplateFromFile('template');
  tpl.data = { header, items, supply, vat, total, supplier: settings.supplier };

  const htmlOutput = tpl.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, '견적서 미리보기');
  const blob = htmlOutput.getBlob().setName(`견적서_${quoteNo}.pdf`);

  // 저장 폴더: 드라이브 루트 (원하면 폴더ID로 변경 가능)
  const file = DriveApp.getRootFolder().createFile(blob);

  SpreadsheetApp.getUi().alert(
    `PDF 생성 완료!\n\n견적번호: ${quoteNo}\n파일: ${file.getName()}\nURL: ${file.getUrl()}`
  );
}

function getSettings_(ss) {
  const sheet = ss.getSheetByName('설정');
  if (!sheet) {
    throw new Error('설정 시트를 찾을 수 없습니다. (시트명: "설정")');
  }

  const rows = sheet.getDataRange().getValues();
  const map = {};

  rows.forEach((row, i) => {
    const key = String(row[0] || '').trim();
    const value = row[1];
    if (!key || key === '키') return;
    map[key] = value;
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
