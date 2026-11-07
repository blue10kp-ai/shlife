// 이 파일만 고치면 설정이 바뀝니다.
window.APP_CONFIG = {
  // Azure 앱 등록에서 받은 "애플리케이션(클라이언트) ID"
  clientId: "5af373d4-1eb9-4fc1-93d3-ba3074160333",

  // 고객 폴더들이 들어 있는 OneDrive 폴더 (OneDrive 최상위 기준)
  baseFolder: "Claude/고객",

  // 다른 사용자가 공유받은 폴더를 못 찾을 때만 채웁니다 (보통은 비워 둠).
  sharedFolder: { driveId: "", itemId: "" },

  // 고객 검색에 쓰는 리스트 엑셀 (baseFolder 바로 안에 있는 파일, 이름에 키워드가 들어간 파일 모두)
  lists: [
    { id: "contract", label: "계약", fileKeyword: "계약고객리스트" },
    { id: "transfer", label: "이관", fileKeyword: "이관고객리스트" }
  ],

  // 세부항목 버튼: 파일/폴더 이름에 키워드가 들어 있으면 그 탭에 표시됩니다.
  tabs: [
    { id: "basic",    label: "기본정보", keywords: ["기본", "정보", "인적", "고객정보"] },
    { id: "contract", label: "계약현황", keywords: ["계약", "증권", "청약"] },
    { id: "analysis", label: "보장분석", keywords: ["보장", "분석", "covan"] },
    { id: "consult",  label: "상담기록", keywords: ["상담", "메모", "기록", "통화"] },
    { id: "contact",  label: "연락하기", keywords: [] },
    { id: "all",      label: "전체파일", keywords: [] }
  ]
};
