# 견적서 양식 앱 — Codex 전달용 기획서 (실행 가능한 코드 산출 목표)

## 1) 프로젝트 개요

Google Sheets에서 견적 데이터를 관리하고, **선택한 견적번호 기준으로 PDF 견적서를 생성**하는 Apps Script 기반 미니 앱.

핵심 기능
- 스프레드시트 메뉴에 `견적서 → 선택 견적번호로 PDF 생성`
- 선택된 행의 `견적번호` 기준으로 모든 행을 모아 템플릿에 주입
- **품명 기준 정렬**
- 템플릿에서 **품명 중복 시 셀 병합(rowSpan)처럼 표시**
- HTML을 PDF로 변환하여 **Drive 루트의 `견적서/` 폴더에 `견적번호.pdf`로 저장**
- 저장 완료 후 **PDF 링크만 있는 모달(`pdf-link.html`)** 표시
- 로컬에서 템플릿을 빠르게 개발하기 위한 **preview 템플릿 + 더미 데이터 + 동기화 스크립트** 포함

---

## 2) 현재 파일 구조 (필수 산출물)

- `Code.js`  
  - Apps Script 로직 (메뉴, 데이터 수집, 정렬, PDF 생성/저장, 링크 모달)
- `template.html`  
  - Apps Script용 HTML 템플릿 (rowSpan 병합 표시 포함)
- `template-preview.html`  
  - 로컬 미리보기 전용 HTML 템플릿 (Live Server용)
- `preview-data.js`  
  - 로컬 미리보기용 더미 데이터 (`window` 가드 필요)
- `pdf-link.html`  
  - PDF 링크만 표시하는 작은 모달 템플릿
- `scripts/sync-template-from-preview.js`  
  - preview → template 동기화 스크립트 (SYNC 마커 구간 치환)
- `README.md`  
  - 사용법/워크플로/동기화 스크립트 설명

---

## 3) 사용자 흐름

1. 사용자는 스프레드시트에서 견적 행을 선택  
2. 메뉴 `견적서 → 선택 견적번호로 PDF 생성` 실행  
3. 선택된 `견적번호`를 기준으로 해당 번호의 모든 행을 모음  
4. **품명 기준 정렬** 후 템플릿에 주입  
5. HTML → PDF 변환 후 Drive 저장  
6. `pdf-link.html` 모달에 “PDF 열기” 버튼 표시

---

## 4) 데이터 모델

### 시트 헤더 (정확히 일치 필요)
- `견적번호`, `견적일`, `업체명`, `수신자`, `품명`, `규격`, `수량`, `단가`, `비고`

### 템플릿 주입 데이터 구조
```
data = {
  header: {
    NO, 견적일, 상호, 수신처
  },
  items: [
    { name, spec, qty, unit, amount, note }
  ],
  supply, vat, total,
  supplier: {
    company, ceo, regNo, address, bizType, bizItem, contact, email, fax
  }
}
```

---

## 5) 핵심 기능 상세

### 5-1. 품명 기준 정렬
- Apps Script에서 `rows`를 `품명` 컬럼 기준으로 `localeCompare('ko')` 정렬

### 5-2. 품명 중복 병합 표시
- `template.html`에서 JS 렌더링 시 **연속된 동일 `name`**에 대해
  - 첫 행의 `<td class="name">`에 `rowSpan` 증가
  - 이후 행은 해당 `<td>` 제거
- `.name` 셀은 `vertical-align: middle` 적용

### 5-3. PDF 저장 규칙
- Drive 루트 하위 `견적서/` 폴더
- 파일명: `견적번호.pdf`
- 동일 파일명 존재 시 **휴지통 처리 후 재생성**

### 5-4. 링크 모달 (pdf-link.html)
- 저장 완료 후 작은 모달 표시
- 버튼 클릭 시 저장된 PDF URL 열기

---

## 6) 로컬 템플릿 개발 워크플로

### Live Server 미리보기
- `template-preview.html`을 Live Server로 열어 실시간 확인
- 더미 데이터는 `preview-data.js`
  - Apps Script 환경에서 오류 방지: `if (typeof window !== "undefined") { ... }`

### preview → template 동기화
- `scripts/sync-template-from-preview.js` 실행
```
node scripts/sync-template-from-preview.js
```
- SYNC 마커 구간만 복사
  - CSS: `/* SYNC:STYLE:START */` ~ `/* SYNC:STYLE:END */`
  - BODY: `<!-- SYNC:BODY:START -->` ~ `<!-- SYNC:BODY:END -->`

---

## 7) 코드 수준 요구사항 (Codex 구현 지침)

### Code.js 요구사항
- `onOpen()` 메뉴 추가
- `generateQuotePdfFromSelection()` 구현
  - 선택 셀 기준 `견적번호` 획득
  - 해당 번호의 모든 행 수집
  - **품명 기준 정렬**
  - 합계 계산(공급가액/부가세/합계)
  - `template.html`에 데이터 주입
  - HTML → PDF 변환: `getAs(MimeType.PDF)`
  - Drive `견적서/` 폴더에 저장, 같은 이름 덮어쓰기
  - `pdf-link.html` 모달로 링크 표시
- `getSettings_()` 구현 (설정 시트에서 공급자 정보 로딩)
- `getOrCreateSubfolder_()` / `overwriteFileInFolder_()` 유틸 포함

### template.html 요구사항
- 현재 HTML 템플릿 구조 유지
- 품명 중복 rowSpan 로직 반영
- `.name` 셀 세로 가운데 정렬

### pdf-link.html 요구사항
- 버튼 하나만 있는 간단 모달
- `pdfUrl` 템플릿 변수 사용

---

## 8) 비기능 요구사항

- Apps Script에서 실행 가능해야 함
- 미리보기 모달은 현재 주석 처리 상태 (필요시 재활성화 가능)
- PDF 생성 결과는 반드시 **PDF로 렌더링**되어야 함 (HTML 그대로 저장 금지)

---

## 9) 남은 TODO (선택 과제)

- PDF 저장 폴더 선택 UI
- 견적번호 자동 생성 규칙
- VAT 옵션(면세/수동 입력)
- 시트 유효성 검사 강화

---

## 10) Codex 에이전트용 실행 프롬프트 (요약)

```
목표:
Google Sheets + Apps Script로 견적서 PDF 생성 앱을 구현한다.
선택한 견적번호 기준으로 데이터를 모아 품명 기준 정렬 후 템플릿에 주입하고,
PDF를 Drive 루트의 '견적서/'에 '견적번호.pdf'로 저장(덮어쓰기 포함)한다.
저장 후 pdf-link.html 모달에 PDF 열기 버튼을 표시한다.

필수 파일:
Code.js, template.html, template-preview.html, preview-data.js, pdf-link.html,
scripts/sync-template-from-preview.js, README.md

핵심 구현 포인트:
- 품명 기준 정렬(ko localeCompare)
- 품명 중복 시 rowSpan 병합 표시 (template.html)
- PDF 저장 시 getAs(MimeType.PDF) 사용
- 동일 파일명 덮어쓰기
- 링크 모달은 pdf-link.html 사용
- preview 템플릿은 Live Server용, 더미데이터는 window 가드 사용
- preview → template 동기화 스크립트 (SYNC 마커 구간 치환)
```

