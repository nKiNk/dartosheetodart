const CONFIG = {
  METADATA_SHEET_NAME: "__METADATA__",
  TEMPLATE_SPREADSHEET_ID: "", // 빈 구글 시트 템플릿 ID (앱스스크립트 속성에서도 설정 가능)
  APP_NAME: "DARToSheeToDART",
  WORKSPACE_FOLDERS: {
    SOURCE: "source",
    WORKSHEET: "worksheet",
    BACKUPSHEET: "backupsheet",
    TARGET: "target"
  },
  MAX_UPLOAD_MB_NOTICE: 50,
  DSD_DATA_END_COLUMN: 13,
  USER_AREA_START_COLUMN: 14,
  USER_AREA_PREVIEW_COLUMNS: 8,
  BOUNDARY_LINE_COLOR: "#9ca3af",
  TABLE_TYPES: {
    STATEMENTS: ["재무상태표", "손익계산서", "포괄손익계산서", "자본변동표", "현금흐름표"]
  },
  DART_TABLE_DEFAULTS: {
    TR: {
      ACOPY: "Y",
      ADELETE: "Y",
      HEIGHT: "30"
    },
    TD: {
      CLASS: "NORMAL",
      VALIGN: "MIDDLE"
    },
    TH: {
      CLASS: "NORMAL",
      VALIGN: "MIDDLE"
    }
  }
};
