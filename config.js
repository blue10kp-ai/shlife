// 이 파일만 고치면 설정이 바뀝니다.
window.APP_CONFIG = {
  // Azure 앱 등록에서 받은 "애플리케이션(클라이언트) ID"
  clientId: "5af373d4-1eb9-4fc1-93d3-ba3074160333",

  // 고객 폴더들이 들어 있는 OneDrive 폴더 (OneDrive 최상위 기준)
  baseFolder: "Claude/고객",
  // 앱이 저장하는 파일(설정·권한)의 폴더. 앱마다 Claude/app/앱이름/설정
  appFolder: "Claude/app/신한FC/설정",

  // 관리자 아이디 (이 계정으로 로그인하면 '관리자' 메뉴가 보임)
  admins: ["blue10kp@gmail.com"],

  // 고객 폴더 소유자(관리자)의 OneDrive 정보 — 다른 사용자가 공유받은 '고객' 폴더를 찾을 때 사용
  owner: { driveId: "1C9CB085D01137F9", baseId: "1C9CB085D01137F9!sf69cb45e11a746d390c1d72f1911c277", appDirId: "1C9CB085D01137F9!sfbb7913784a04d6ca4efee053d78d36f" },

  // 저장소: "onedrive" 또는 "google" (주소 끝 ?drive=google 로 이 기기만 먼저 바꿔 시험할 수 있음)
  storage: "onedrive",
  // Google Drive 모드: 관리자는 내 드라이브의 Claude/고객 을 경로로 찾음.
  //   다른 사용자는 관리자가 공유해 준 폴더를 아래 ID로 찾음 (폴더 주소 drive.google.com/drive/folders/<ID>)
  gdrive: { baseId: "1IaRf-UbNqNCaI9Q_XdV3_Tt8bQVHnuNM", appDirId: "1ERsiC6gUGh0LljPSKxRx04Z5dZUGJGC0" },

  // 아이디별 FC 권한을 저장하는 파일 (고객 폴더 안). 관리자 페이지에서 편집
  permFile: "shlife-권한.json",

  // 고객 검색 범위 버튼 (전체 + 아래 순서)
  //   FC 폴더 안 "날짜 - FC - 리스트.xlsx"의 시트를 읽음 (FC마다 가장 최근 날짜 파일 하나)
  //   순서 = 우선순위 (같은 고객이 여러 시트에 있으면 앞쪽 우선, 일정에서도 이 순서로 표시)
  //   hidden: 버튼은 없고 '전체' 검색에만 포함
  lists: [
    { id: "contract", label: "계약", fileRegex: " - (리스트|계약)\\.xlsx$", sheet: "계약고객" },
    { id: "transfer", label: "이관", fileRegex: " - (리스트|계약)\\.xlsx$", sheet: "이관고객" },
    { id: "ddahaeng", label: "따행", fileRegex: " - 리스트\\.xlsx$", sheet: "따행고객" },
    { id: "db",       label: "DB",   fileRegex: " - 리스트\\.xlsx$", sheet: "DB고객" },
    { id: "plan",     label: "추가", fileRegex: " - 리스트\\.xlsx$", sheet: "예정고객" },
    { id: "progress", label: "진행", fileRegex: " - (리스트|기타)\\.xlsx$", sheet: "진행고객", hidden: true }
  ],

  // 계약 탭에 보여줄 계약 상세 ('계약자'·'주피보험자'·'피보'로 고객 매칭)
  contractLists: [
    { label: "계약", fileRegex: " - (리스트|계약)\\.xlsx$", sheet: "계약" },
    { label: "이관", fileRegex: " - (리스트|계약)\\.xlsx$", sheet: "이관" }
  ],

  // 기본정보 탭의 추가 정보 (고객 폴더 바로 안 엑셀). 담당 FC + 고객명으로 줄을 찾아 읽고/저장
  //   '기본' 시트: 담당 FC·고객명으로 줄을 찾음 / '상담기록' 시트: 고객명·담당 FC·저장날짜·내용
  extraInfo: {
    fileKeyword: "고객추가정보", sheet: "기본", fcCol: "담당 FC", nameCol: "고객명",
    primary: ["미팅장소", "메모"], carrierCol: "통신사",
    afterAddress: ["차량번호", "자동차보험갱신일"], monthDayCols: ["자동차보험갱신일"],
    consult: { sheet: "상담기록", nameCol: "고객명", fcCol: "담당 FC", dateCol: ["저장날짜시간", "저장날짜"], textCol: "내용", voiceCol: "음성파일", voiceText: "음성기록" }
  },

  // 세부항목 버튼: 파일/폴더 이름에 키워드가 들어 있으면 그 탭에 표시됩니다.
  tabs: [
    // 기본정보에는 첨부파일을 붙이지 않음 (상품명의 '기본형' 등에 걸려 PDF가 붙던 문제)
    { id: "basic",    label: "정보", keywords: [] },
    { id: "contract", label: "계약", keywords: ["계약", "증권", "청약"] },
    // 고객 계약 PDF("날짜 - 고객 - FC - 상품명 - 금액원.pdf")는 보장분석에서 제외
    { id: "analysis", label: "분석", keywords: ["보장", "분석", "covan"], exclude: "\\d[\\d,]*원\\.pdf$" },
    { id: "consult",  label: "상담", keywords: ["상담", "메모", "기록", "통화"] }
  ],

  // 상담 탭: 처음에 보여줄 상담기록 수 (나머지는 '더보기')
  consultPageSize: 10,
  // 일정: 구글 캘린더 (읽기 전용). 관리자 설정에서 고른 캘린더가 없으면 아래 이름의 캘린더를 사용
  google: {
    clientId: "482206908404-5sdbvetcrc8j7ctrhe6aa23ccmd3d05d.apps.googleusercontent.com",
    calendars: ["1팀", "김현성", "최수진"],
    teamCalendar: "1팀", // 이 캘린더 일정만 고객 정보·이동 시간 표시 (나머지는 개인 일정)
    // 이름이 이메일로 보이는 캘린더의 표시 이름
    calendarNames: { "blue10kp@gmail.com": "김현성", "nonmollarechoi@gmail.com": "최수진" }
  }
};
