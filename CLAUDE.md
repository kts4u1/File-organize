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
   손글씨 시리얼은 엑셀 시리얼 목록과 대조해 가장 가까운 툴 → 애매하면 '확인 필요'), QR 스티커(jsQR, PDF 생성),
   Office Script 생성(Excel 웹 자동화 탭에서 실행 → 파일 교체 없이 반영). claude.ai 아티팩트: https://claude.ai/artifact/APBdgkG9sx6URYhA2a3WWp
2. `desktop/` + `desktop_main.py` → `툴검교정.exe` — FSN 프로그램(kts4u1/Customer)과 같은 Edge SSO 방식. 로컬 서버(127.0.0.1, 실행마다 토큰)가
   `web/`을 띄우고 Edge로 엶. SharePoint 받기 = 페이지 안 fetch `download.aspx?UniqueId=GUID`, 올리기 = REST `GetFileById(GUID)/$value`
   (X-HTTP-Method: PUT, contextinfo digest, If-Match). 받은 뒤 ETag가 바뀌었으면 거절(412), 편집 중 잠김은 423으로 안내.
   흐름(2026-11-02 사용자 요청 "사진만 넣고 … 업로드하면 끝나게"): 시작하면 SharePoint 파일을 자동으로 받음(받는 동안 넣은 사진은 기다림,
   진행은 `/api/progress`로 표시) → 사진 → 버튼 하나 `SharePoint에 반영 (N건)` = 엑셀 만들기 + PC 사본 저장 + 올리기.
   412(그 사이 다른 사람 저장)면 `최신 파일로 다시 반영`(같은 사진을 최신본에 다시 반영). TEST 탭 시험 반영 기능은 화면에서 뺐음
   (`xlsx.js`의 TEST 시트 함수는 남아 있음, 안 씀).
   창은 Edge 하나(2026-11-02 사용자 요청 "창 2개를 하나로"): 예전 Tk 상태 창 없음. 상태 기록은 페이지 오른쪽 위 '프로그램 상태'
   (`/api/progress`의 `log`, 같은 단계의 진행 줄은 한 줄로 갱신), 저장/로그 폴더 버튼(`/api/open`), 도구 창을 닫으면 종료
   (`Program.wait_until_closed`, 테스트 `test_program_ends_when_the_tool_window_is_closed`). 오류로 Edge를 못 열 때만 Tk 메시지 상자.
3. `app.py`, `toolcal/` — 처음 만든 Python(Streamlit) 버전. Claude API 선택 사항. 지금은 1·2가 주력.

## SharePoint
- 대상 파일: `https://share.philips.com/:x:/r/sites/HSCSCCJ-jr-all/_layouts/15/Doc.aspx?…sourcedoc={5b1d0df6-6a62-49fb-92e3-3cb102684f84}…`
  (`desktop/app.py`의 DEFAULT_SHAREPOINT_URL, `%USERPROFILE%\ToolCal\config.json`의 `sharepoint_url`로 변경)
- 앱 폴더 `%USERPROFILE%\ToolCal`: `debug_log.txt`, `output\`(올린 파일 사본), `browser_profile\`.

## 테스트
- `python -m pytest` (Python 쪽 + desktop 서버 + 실제 Chromium으로 SharePoint 흉내 서버 받기/올리기: `tests/mock_sharepoint.py`).
- 브라우저 도구는 Playwright로 확인(스크래치 스크립트). Office Script는 문서 기준 타입 검사 + 모의 실행만, 실제 Excel 웹 실행은 사용자 확인 필요.

## 이미 해결한 문제 (되돌리지 말 것)
- 2026-11-02 첫 실행: SharePoint 로그인·파일 정보는 성공, 받기에서 `ReadTimeoutError … localhost:<드라이버 포트> (read timeout=30)`.
  원인: WebDriver 명령 대기 30초보다 페이지 안 다운로드가 오래 걸림(FSN의 `set_timeout(20)`은 신버전 selenium에서 실제로 적용되지 않아 안 겪었음).
  수정: 받기/올리기를 페이지에서 백그라운드로 시작하고 1초마다 상태만 확인(`DOWNLOAD_START_JS`/`UPLOAD_START_JS` + `_poll`), 진행 MB 표시,
  3분간 진행 없으면 중단, 명령 대기 120초. 테스트 `test_slow_transfers_outlast_the_webdriver_command_limit`(이전 코드는 같은 오류로 실패).
- 실제 SharePoint 파일 이름은 `Tool list for Field validation 20250325.xlsx` (처음 받은 사본과 다를 수 있음 → 시트/헤더 확인 필요).

- 2026-11-02 두 번째 실행: 받기 성공, TEST 반영 후 올리기(207.8 MB, 76초)에서 423
  `SPFileLockException … locked for shared use by taesu.kim@philips.com` — 다른 사람이 Excel에서 열어 둠. 팀 공용 파일이라 흔한 상황.
  → 423이면 잠근 사람을 표시하고 결과 칸에 'Excel 웹 스크립트로 반영' 버튼(같은 변경으로 Office Script 생성) + 'Excel 웹에서 파일 열기' 링크.
  파일째 교체는 아무도 안 열었을 때만 가능. 실제 파일이 약 208 MB라 받기·올리기가 각각 1분 이상 걸림.
- 같은 날 사용자 확인: taesu.kim = 사용자 본인. 원인은 프로그램 자신 — 로그인하려고 파일 링크(Doc.aspx?action=edit)를 열어
  Excel 웹이 편집 모드로 파일을 잠갔음. 수정: SharePoint 창은 '사이트 콘텐츠'(`/_layouts/15/viewlsts.aspx`, `landing_url`)로 로그인하고
  파일(Doc.aspx)은 절대 열지 않음, 창이 파일 화면에 있으면 먼저 벗어남. 테스트 `test_program_never_opens_the_workbook_itself`
  (흉내 서버는 Doc.aspx를 열면 잠김; 이전 코드는 실패).

- 같은 날 새 빌드로 올리기 성공(사용자 확인: SharePoint에 TEST 탭이 생김). 그 뒤 TEST 기능 제거·자동 받기·바로 올리기로 변경.
  SharePoint에 남은 TEST 탭은 사용자가 직접 지우면 됨(프로그램은 건드리지 않음).

- 같은 날 저녁: 자동 받기 + `SharePoint에 반영` 한 번에 올리기 성공(사용자 화면 확인). 이때 받은 파일은 10.7 MB(전엔 207.8 MB),
  시트 `Calweb등록필요`가 새로 생김 — 그 사이 누가 Excel에서 정리·저장한 것으로 보임.

## 열린 항목 (2026-11-02)
- 창 하나(상태 창 제거, Edge 창 닫으면 종료) — 새 빌드로 실제 확인 필요.
- 다른 사람이 실제로 열어 둔 경우엔 Office Script(Excel 웹 자동화 탭) 경로 — 실제 실행 결과 미확인.
- Office Script 첫 실제 실행 결과 확인.
- 손글씨 시리얼 인식률은 휴대폰 원본 해상도 사진으로 다시 확인. 확실하게 하려면 QR 스티커 사용 권장.
