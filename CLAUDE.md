# 툴 검교정 정리 — 인계 문서

검교정 후 찍은 툴 사진(손글씨 메모 또는 QR 스티커 + 교정필증)으로 Tool Verification 엑셀(SharePoint)의
Cal Last / Cal Next와 Label Photo / SMAX Tool Photo를 고치는 도구입니다. 브랜치 `claude/exciting-lovelace-hs156x`.

## 사용자 / 작업 방식
- 사용자는 Philips 한국 현장 서비스 엔지니어, 프로그래밍 경험 없음. 코드는 Claude가 작성·디버깅. 대화는 한국어 존댓말, 짧고 구체적으로.
- 회사 Windows PC, Python 3.14.x, `pip` 대신 항상 `python -m pip`. 브랜치 ZIP → 압축 풀기 → `build.bat` → `dist\툴검교정.exe`.
- API 키는 쓰지 않기로 함(2026-10-15 "API는 필요없게"). 회사 데이터는 Claude 환경에서 볼 수 없음(SharePoint 403).
- "일단 만들어 보고 수정은 하나씩". 큰 구조 변경은 먼저 확인.

## 구성 (세 가지 쓰는 방법, 같은 화면)
1. `web/` — 브라우저 도구(순수 JS). 엑셀 읽기·쓰기(JSZip, XML 최소 수정), 교정필증 검출·크롭, Tesseract.js OCR(교정필증 날짜는 잘 읽고,
   손글씨 시리얼은 엑셀 시리얼 목록과 대조해 가장 가까운 툴 → 애매하면 '확인 필요'), QR 스티커(jsQR, PDF 생성), TEST 탭 시험 반영,
   Office Script 생성(Excel 웹 자동화 탭에서 실행 → 파일 교체 없이 반영). claude.ai 아티팩트: https://claude.ai/artifact/APBdgkG9sx6URYhA2a3WWp
2. `desktop/` + `desktop_main.py` → `툴검교정.exe` — FSN 프로그램(kts4u1/Customer)과 같은 Edge SSO 방식. 로컬 서버(127.0.0.1, 실행마다 토큰)가
   `web/`을 띄우고 Edge로 엶. SharePoint 받기 = 페이지 안 fetch `download.aspx?UniqueId=GUID`, 올리기 = REST `GetFileById(GUID)/$value`
   (X-HTTP-Method: PUT, contextinfo digest, If-Match). 받은 뒤 ETag가 바뀌었으면 거절(412), 편집 중 잠김은 423으로 안내.
3. `app.py`, `toolcal/` — 처음 만든 Python(Streamlit) 버전. Claude API 선택 사항. 지금은 1·2가 주력.

## SharePoint
- 대상 파일: `https://share.philips.com/:x:/r/sites/HSCSCCJ-jr-all/_layouts/15/Doc.aspx?…sourcedoc={5b1d0df6-6a62-49fb-92e3-3cb102684f84}…`
  (`desktop/app.py`의 DEFAULT_SHAREPOINT_URL, `%USERPROFILE%\ToolCal\config.json`의 `sharepoint_url`로 변경)
- 앱 폴더 `%USERPROFILE%\ToolCal`: `debug_log.txt`, `output\`(올린 파일 사본), `browser_profile\`.

## 테스트
- `python -m pytest` (Python 쪽 + desktop 서버 + 실제 Chromium으로 SharePoint 흉내 서버 받기/올리기: `tests/mock_sharepoint.py`).
- 브라우저 도구는 Playwright로 확인(스크래치 스크립트). Office Script는 문서 기준 타입 검사 + 모의 실행만, 실제 Excel 웹 실행은 사용자 확인 필요.

## 열린 항목 (2026-11-02)
- 사용자 PC에서 `툴검교정.exe` 첫 실행: SharePoint 받기 → TEST 탭 반영 → 올리기 결과(성공/423 잠김/권한) 확인.
- Office Script 첫 실제 실행 결과 확인.
- 손글씨 시리얼 인식률은 휴대폰 원본 해상도 사진으로 다시 확인. 확실하게 하려면 QR 스티커 사용 권장.
