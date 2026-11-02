# 툴 검교정 정리 — 인계 문서

검교정 후 찍은 툴 사진(손글씨 메모 또는 QR 스티커 + 교정필증)으로 Tool Verification 엑셀(SharePoint)의
Cal Last / Cal Next와 Label Photo / SMAX Tool Photo를 고치는 도구입니다.
저장소 `kts4u1/File-organize`, 브랜치 `claude/exciting-lovelace-hs156x`.

## 한눈에 (2026-11-02 기준)
- 주로 쓰는 것: 회사 PC 프로그램 `툴검교정.exe`. 실행 → SharePoint 파일 자동 받기 → 사진 넣기 → `SharePoint에 반영 (N건)` 한 번 → 올리기 →
  올린 파일이 SharePoint 창에 보기 모드로 열림. 실제 회사 SharePoint에서 받기·올리기 성공 확인됨(사용자 화면).
- 마지막으로 바꾼 것(실제 PC 확인 전): 창 하나(Edge), SharePoint 창 평소 최소화, 올린 뒤 파일 보기. → 아래 '열린 항목'.

## 사용자 / 작업 방식
- 사용자는 Philips 한국 현장 서비스 엔지니어, 프로그래밍 경험 없음. 코드는 Claude가 작성·디버깅. 대화는 한국어 존댓말, 짧고 구체적으로.
- 회사 Windows PC, Python 3.14.x, `pip` 대신 항상 `python -m pip`.
- 새 빌드 안내 때는 ZIP 주소를 링크 글자로 숨기지 말고 전체 주소 그대로 적기(2026-11-02 요청). 안내 형식:
  1. ZIP 받기: https://github.com/kts4u1/File-organize/archive/refs/heads/claude/exciting-lovelace-hs156x.zip
  2. 받은 ZIP 파일을 오른쪽 클릭 → **압축 풀기**
  3. 풀린 폴더 안의 `build.bat` 더블클릭 (1~2분)
  4. 같은 폴더의 `dist\툴검교정.exe` 실행
- API 키는 쓰지 않기로 함(2026-10-15 "API는 필요없게"). 회사 데이터는 Claude 환경에서 볼 수 없음(SharePoint 403) → 흉내 서버로 시험.
- "일단 만들어 보고 수정은 하나씩". 큰 구조 변경은 먼저 확인. 회사 엑셀·사진은 커밋하지 않음(.gitignore).

## 구성 (세 가지 쓰는 방법, 같은 화면)
1. `web/` — 브라우저 도구(순수 JS). 엑셀 읽기·쓰기(JSZip, XML 최소 수정), 교정필증 검출·크롭, Tesseract.js OCR(교정필증 날짜는 잘 읽고,
   손글씨 시리얼은 엑셀 시리얼 목록과 대조해 가장 가까운 툴 → 애매하면 '확인 필요'), QR 스티커(jsQR, PDF 생성),
   Office Script 생성(Excel 웹 자동화 탭에서 실행 → 파일 교체 없이 반영). claude.ai 아티팩트: https://claude.ai/artifact/APBdgkG9sx6URYhA2a3WWp
   (아티팩트는 `web/index.html` + `js/app.js`를 다시 게시해서 갱신)
2. `desktop/` + `desktop_main.py` → `툴검교정.exe` — FSN 프로그램(kts4u1/Customer)과 같은 Edge SSO 방식. 로컬 서버(127.0.0.1, 실행마다 토큰)가
   `web/`을 띄우고 Edge로 엶. SharePoint 받기 = 페이지 안 fetch `download.aspx?UniqueId=GUID`, 올리기 = REST `GetFileById(GUID)/$value`
   (X-HTTP-Method: PUT, contextinfo digest, If-Match). 받은 뒤 ETag가 바뀌었으면 거절(412), 편집 중 잠김은 423으로 안내.
   - 흐름(2026-11-02 "사진만 넣고 … 업로드하면 끝나게"): 시작하면 SharePoint 파일 자동 받기(받는 동안 넣은 사진은 기다림) → 사진 →
     버튼 하나 `SharePoint에 반영 (N건)` = 엑셀 만들기 + PC 사본 저장 + 올리기. 412면 `최신 파일로 다시 반영`(같은 사진을 최신본에 다시).
     TEST 탭 시험 반영은 화면에서 뺐음(`xlsx.js`의 TEST 시트 함수는 남아 있음, 안 씀).
   - 창은 Edge 하나(2026-11-02 "창 2개를 하나로"): Tk 상태 창 없음. 상태 기록은 페이지 오른쪽 위 '프로그램 상태'
     (`/api/progress`의 `log`, 같은 단계 진행 줄은 한 줄로 갱신), 저장/로그 폴더 버튼(`/api/open`), 도구 창을 닫으면 종료
     (`Program.wait_until_closed`). Edge를 못 열 때만 Tk 메시지 상자.
   - SharePoint 창(로그인된 별도 Edge 창, 받기·올리기를 이 창에서 함): 평소 최소화, 로그인 기다릴 때만 앞으로, 제목·노란 줄 "닫지 마세요".
     올리기 성공 뒤 이 창에 파일을 보기 모드로 열어 앞에 둠.
3. `app.py`, `toolcal/` — 처음 만든 Python(Streamlit) 버전. Claude API 선택 사항. 지금은 1·2가 주력.

## 주요 파일
- `web/index.html`(화면·CSS), `web/js/app.js`(화면 흐름, 데스크톱 연동: 자동 받기·올리기·프로그램 상태), `xlsx.js`(엑셀 읽기·쓰기),
  `vision.js`(교정필증 검출, OCR), `match.js`(시리얼 대조·판정), `officescript.js`(Office Script 생성), `stickers.js`(QR 스티커 PDF),
  `demo.js`/`demo-data.js`(예시로 해보기), `ocr-worker.js`(Tesseract 언어 데이터 로딩), `web/vendor`, `web/tesseract`(오프라인 라이브러리).
- `desktop/sharepoint.py`(`SharePointFile` 주소 해석·`view_url`·`landing_url`, `BrowserSession` Edge 창 관리, `SharePoint` 받기·올리기),
  `desktop/server.py`(로컬 서버, `/api/*`), `desktop/app.py`(`Program`, 설정·로그 경로, `DEFAULT_SHAREPOINT_URL`).
- `build.py` / `build.bat`(PyInstaller onefile, CRLF·ASCII만), `requirements-desktop.txt`(selenium, pyinstaller).
- `tests/mock_sharepoint.py`(SharePoint 흉내: 로그인 쿠키, digest, ETag, Doc.aspx 편집=잠금·보기=안 잠금, 느린 전송), `tests/test_desktop.py`.

## SharePoint
- 대상 파일: `https://share.philips.com/:x:/r/sites/HSCSCCJ-jr-all/_layouts/15/Doc.aspx?…sourcedoc={5b1d0df6-6a62-49fb-92e3-3cb102684f84}…`
  (`desktop/app.py`의 DEFAULT_SHAREPOINT_URL, `%USERPROFILE%\ToolCal\config.json`의 `sharepoint_url`로 변경)
- 실제 파일 이름 `Tool list for Field validation 20250325.xlsx`. 시트(2026-11-02 저녁): 광주, 대전, 분실, 사무실 UPS, 부외자산, Calweb등록필요.
  크기 10.7 MB(그날 낮엔 207.8 MB였음 — 누가 정리한 듯).
- 앱 폴더 `%USERPROFILE%\ToolCal`: `debug_log.txt`, `output\`(올린 파일 사본), `browser_profile\`.

## 빌드 / 테스트
- 빌드: `build.bat` → `dist\툴검교정.exe` (`--collect-submodules=selenium`, `--add-data web`).
- `python -m pytest` — 52개(2026-11-02). Python 쪽 + desktop 서버 + 실제 Chromium으로 SharePoint 흉내 서버 받기/올리기.
  Chromium은 `/opt/pw-browsers/chromium-1194`, PATH에서 node22(다른 버전 chromedriver)를 빼고 실행.
  headless Chromium은 브라우저의 첫 창을 닫을 때 가끔 15초 멈춤 → 창 닫기 테스트는 도구를 두 번째 창에 엶.
- 브라우저 도구는 Playwright(Node)로 확인(스크래치 스크립트). Office Script는 문서 기준 타입 검사 + 모의 실행만, 실제 Excel 웹 실행은 사용자 확인 필요.

## 이미 해결한 문제 (되돌리지 말 것)
- 2026-11-02 첫 실행: 받기에서 `ReadTimeoutError … localhost:<드라이버 포트> (read timeout=30)`.
  원인: WebDriver 명령 대기 30초보다 페이지 안 다운로드가 오래 걸림(FSN의 `set_timeout(20)`은 신버전 selenium에서 실제로 적용되지 않아 안 겪었음).
  수정: 받기/올리기를 페이지에서 백그라운드로 시작하고 1초마다 상태만 확인(`DOWNLOAD_START_JS`/`UPLOAD_START_JS` + `_poll`), 진행 MB 표시,
  3분간 진행 없으면 중단, 명령 대기 120초. 테스트 `test_slow_transfers_outlast_the_webdriver_command_limit`.
- 같은 날 올리기(207.8 MB)에서 423 `SPFileLockException … locked for shared use by taesu.kim@philips.com`.
  → 423이면 잠근 사람을 표시하고 'Excel 웹 스크립트로 반영' 버튼(같은 변경으로 Office Script) + 'Excel 웹에서 파일 열기' 링크.
  파일째 교체는 아무도 편집 중이 아닐 때만 가능(팀 공용 파일이라 흔한 상황).
- taesu.kim = 사용자 본인. 원인은 프로그램 자신 — 로그인하려고 파일 링크(Doc.aspx?action=edit)를 열어 Excel 웹이 편집 모드로 잠갔음.
  수정: SharePoint 창은 '사이트 콘텐츠'(`/_layouts/15/viewlsts.aspx`, `landing_url`)로 로그인, 받기·올리기 중에는 파일 화면을 열지 않고
  창이 파일 화면에 있으면 먼저 벗어남. 파일을 여는 건 올리기 성공 뒤 확인용 **보기 모드(action=view)** 뿐 — 편집 모드로 열면 안 됨.
  테스트 `test_program_never_opens_the_workbook_itself`.
- 로컬 서버 받기 한도가 200 MiB라 208 MB 파일에 2 MB 여유뿐이었음 → 1 GiB(`MAX_UPLOAD`). 테스트 `test_size_limit_leaves_room_for_the_real_workbook`.
- "왜 사이트 콘텐츠 창이 항상 떠 있어? 닫으면 업로드도 안 돼" + "마지막에는 파일(Doc.aspx) 창이 열려야 확인할 수 있어".
  → SharePoint 창 평소 최소화(`tuck_sharepoint`), 로그인 페이지가 3초 넘게 기다릴 때만 앞으로(`show_sharepoint`), 안내(`LABEL_JS`).
  창이 닫혀 명령이 실패하면 다시 열고 한 번 재시도(`_again_if_window_closed`; PUT을 보낸 뒤면 재시도하지 않고 '다시 올리기' 안내 —
  이미 올라갔으면 412 → '최신 파일로 다시 반영'). 10분 넘게 쉬었으면 사이트 콘텐츠를 다시 열어 로그인 갱신.
  올리기 성공 뒤 파일을 보기 모드(`view_url`)로 열고 앞에 둠(`keep_front`). 테스트 `test_sharepoint_window_stays_out_of_the_way`.
- 결과 칸에 "null" 글자가 찍히던 문제 → `put()`(null 빼고 append).

## 열린 항목 (2026-11-02)
- 창 하나 + SharePoint 창 최소화 + 올린 뒤 파일 보기 — 새 빌드로 실제 확인 필요.
  (Windows Edge에서 msedgedriver가 창을 바꿀 때 최소화된 창을 다시 띄우는지 미확인 — 띄우면 잠깐 보였다가 다시 최소화됨)
- Excel 웹 보기 모드(action=view)가 실제로 잠그지 않는지 확인(올린 파일을 띄워 둔 채 다음 사진 올리기).
- 다른 사람이 실제로 열어 둔 경우의 Office Script(Excel 웹 자동화 탭) 경로 — 실제 실행 결과 미확인.
- 손글씨 시리얼 인식률은 휴대폰 원본 해상도 사진으로 다시 확인. 확실하게 하려면 QR 스티커 사용 권장.
- SharePoint에 남은 TEST 탭은 사용자가 직접 지우기로 함(프로그램은 건드리지 않음).
