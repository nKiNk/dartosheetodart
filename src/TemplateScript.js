/**
 * DARToSheeToDART 템플릿용 앱스 스크립트 코드
 * 이 코드를 복사하여 템플릿용으로 사용할 빈 구글 시트의 [확장 프로그램 > Apps Script] 에 붙여넣고 저장하세요.
 * 그리고 생성된 시트의 ID를 Config.js 또는 스크립트 속성의 TEMPLATE_SPREADSHEET_ID 로 지정하세요.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("DARToSheeToDART")
    .addItem("Sheet -> DSD 변환 (웹앱에서 열기)", "openWebAppForConversion")
    .addToUi();
}

function openWebAppForConversion() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const metaSheet = sheet.getSheetByName("__METADATA__");
    if (!metaSheet) {
      throw new Error("메타데이터 시트를 찾을 수 없습니다. 정상적으로 변환된 시트가 아닙니다.");
    }
    
    // __SUMMARY__ 파싱
    const data = metaSheet.getRange("A1:B30").getValues();
    let runtimeContext = null;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === "__SUMMARY__") {
        const summary = JSON.parse(data[i][1]);
        runtimeContext = summary.runtimeContext;
        break;
      }
    }

    if (!runtimeContext || !runtimeContext.webAppUrl) {
      throw new Error("웹앱 URL 정보가 메타데이터에 없습니다.");
    }

    let url = runtimeContext.webAppUrl;
    url += (url.indexOf("?") === -1 ? "?" : "&") + "sheetId=" + sheet.getId();

    const htmlString = `
      <div style="font-family: 'Malgun Gothic', sans-serif; padding: 20px; text-align: center;">
        <h3 style="color: #1a73e8; margin-top: 0;">DARToSheeToDART</h3>
        <p style="font-size: 14px; color: #333;">아래 버튼을 클릭하여 웹앱에서 변환을 진행하세요.</p>
        <a href="${url}" target="_blank" style="
          display: inline-block;
          margin-top: 15px;
          padding: 10px 24px;
          background-color: #1a73e8;
          color: white;
          text-decoration: none;
          border-radius: 4px;
          font-weight: bold;
          font-size: 14px;
        " onclick="google.script.host.close();">
          웹앱 열기 (새 창)
        </a>
      </div>
    `;

    const html = HtmlService.createHtmlOutput(htmlString)
      .setWidth(350)
      .setHeight(180);
    
    SpreadsheetApp.getUi().showModalDialog(html, "웹앱 연결");
  } catch (error) {
    SpreadsheetApp.getUi().alert("오류 발생", error.message || String(error), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

