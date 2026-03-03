function doGet(e) {
  const template = HtmlService.createTemplateFromFile("WebApp");
  template.initialSheetId = (e && e.parameter && e.parameter.sheetId) ? e.parameter.sheetId : "";
  
  return template.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setFaviconUrl("https://dart.fss.or.kr/favicon.ico")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen(e) {
  addSpreadsheetMenu_();
}

function onInstall(e) {
  onOpen(e);
}

function addSpreadsheetMenu_() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("DARToSheeToDART")
      .addItem("Sheet -> DSD 변환 (웹앱에서 열기)", "menuConvertSheetToDsd")
      .addToUi();
  } catch(e) {
    // getUi() fails when run from standalone script without active document
  }
}

function menuConvertSheetToDsd() {
  try {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSheet) {
      throw new Error("활성화된 시트가 없습니다.");
    }
    
    // readMetadataMap은 SheetToDsd.js에 정의되어 있음
    const metadata = readMetadataMap(activeSheet);
    const runtimeContext = metadata.runtimeContext || (metadata.summary ? metadata.summary.runtimeContext : null);
    
    if (!runtimeContext || !runtimeContext.webAppUrl) {
      throw new Error("웹앱 URL 정보가 없습니다. 웹앱에서 시트를 생성한 경우가 맞는지 확인하세요.");
    }

    const sheetId = activeSheet.getId();
    let url = runtimeContext.webAppUrl;
    url += (url.indexOf("?") === -1 ? "?" : "&") + "sheetId=" + sheetId;

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
    SpreadsheetApp.getUi().alert("오류", error.message || String(error), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppBootstrap() {
  return {
    appName: CONFIG.APP_NAME,
    slogan: "신속하고 투명한 전자공시시스템, DARToSheeToDART",
    maxUploadMbNotice: CONFIG.MAX_UPLOAD_MB_NOTICE,
    webAppUrl: ScriptApp.getService().getUrl() || ""
  };
}

function apiCreateWorkspace(parentFolderInput) {
  const parentFolderId = parseDriveId(parentFolderInput);
  if (!parentFolderId) {
    throw new Error("유효한 작업 폴더 URL 또는 ID를 입력하세요.");
  }

  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const workspace = ensureWorkspaceFolders(parentFolder);

  return {
    rootFolderId: workspace.rootFolder.getId(),
    rootFolderName: workspace.rootFolder.getName(),
    sourceFolderId: workspace.sourceFolder.getId(),
    worksheetFolderId: workspace.worksheetFolder.getId(),
    backupsheetFolderId: workspace.backupsheetFolder.getId(),
    targetFolderId: workspace.targetFolder.getId()
  };
}

function apiUploadDsdAndCreateWorksheet(payload) {
  if (!payload || !payload.rootFolderId || !payload.fileName || !payload.base64Data) {
    throw new Error("업로드 요청 값이 부족합니다.");
  }

  const workspace = ensureWorkspaceFolders(DriveApp.getFolderById(payload.rootFolderId));
  const mimeType = payload.mimeType || "application/octet-stream";
  const fileBlob = Utilities.newBlob(Utilities.base64Decode(payload.base64Data), mimeType, payload.fileName);
  const sourceFile = workspace.sourceFolder.createFile(fileBlob);

  const extracted = extractDsd(sourceFile.getId());
  const parsed = parseDsdStructure(extracted.contentsXml);

  const spreadsheetName = sourceFile.getName().replace(/\.dsd$/i, "") + "_변환";
  const runtimeContext = {
    appName: CONFIG.APP_NAME,
    webAppUrl: ScriptApp.getService().getUrl() || "",
    rootFolderId: workspace.rootFolder.getId(),
    sourceFolderId: workspace.sourceFolder.getId(),
    worksheetFolderId: workspace.worksheetFolder.getId(),
    backupsheetFolderId: workspace.backupsheetFolder.getId(),
    targetFolderId: workspace.targetFolder.getId(),
    sourceFileId: sourceFile.getId(),
    sourceFileName: sourceFile.getName(),
    createdAt: new Date().toISOString()
  };

  const templateId = PropertiesService.getScriptProperties().getProperty("TEMPLATE_SPREADSHEET_ID") || CONFIG.TEMPLATE_SPREADSHEET_ID || "";
  const spreadsheet = createSpreadsheetFromSections(parsed, spreadsheetName, extracted.contentsXml, runtimeContext, templateId, workspace.worksheetFolder);
  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  
  if (!templateId) {
    spreadsheetFile.moveTo(workspace.worksheetFolder);
  }

  const backupName = spreadsheet.getName() + "_baseline_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const backupFile = spreadsheetFile.makeCopy(backupName, workspace.backupsheetFolder);
  runtimeContext.backupSpreadsheetId = backupFile.getId();
  runtimeContext.backupSpreadsheetUrl = backupFile.getUrl();
  writeRuntimeContext(spreadsheet, runtimeContext);

  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sourceFileId: sourceFile.getId(),
    sourceFileName: sourceFile.getName(),
    backupSpreadsheetId: backupFile.getId(),
    backupSpreadsheetUrl: backupFile.getUrl(),
    targetFolderId: workspace.targetFolder.getId()
  };
}

function apiConvertSheetToDsd(spreadsheetInput) {
  const spreadsheetId = parseDriveId(spreadsheetInput);
  if (!spreadsheetId) {
    throw new Error("유효한 구글 시트 URL 또는 ID를 입력하세요.");
  }

  const result = runSheetsToDsd(spreadsheetId);
  return {
    fileId: result.fileId,
    fileName: result.fileName,
    fileUrl: result.fileUrl,
    folderId: result.folderId
  };
}

function runDsdToSheetsLegacy(fileId, rootFolderId) {
  return apiUploadDsdAndCreateWorksheet({
    rootFolderId: rootFolderId,
    fileName: DriveApp.getFileById(fileId).getName(),
    mimeType: DriveApp.getFileById(fileId).getMimeType(),
    base64Data: Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes())
  });
}

function ensureWorkspaceFolders(rootFolder) {
  const sourceFolder = getOrCreateChildFolder(rootFolder, CONFIG.WORKSPACE_FOLDERS.SOURCE);
  const worksheetFolder = getOrCreateChildFolder(rootFolder, CONFIG.WORKSPACE_FOLDERS.WORKSHEET);
  const backupsheetFolder = getOrCreateChildFolder(rootFolder, CONFIG.WORKSPACE_FOLDERS.BACKUPSHEET);
  const targetFolder = getOrCreateChildFolder(rootFolder, CONFIG.WORKSPACE_FOLDERS.TARGET);

  return {
    rootFolder: rootFolder,
    sourceFolder: sourceFolder,
    worksheetFolder: worksheetFolder,
    backupsheetFolder: backupsheetFolder,
    targetFolder: targetFolder
  };
}

function getOrCreateChildFolder(parentFolder, childName) {
  const folders = parentFolder.getFoldersByName(childName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(childName);
}

function writeRuntimeContext(spreadsheet, runtimeContext) {
  const metaSheet = spreadsheet.getSheetByName(CONFIG.METADATA_SHEET_NAME);
  if (!metaSheet) {
    return;
  }

  const lastRow = metaSheet.getLastRow();
  const key = "__RUNTIME_CONTEXT__";
  const rows = lastRow > 0 ? metaSheet.getRange(1, 1, lastRow, 2).getValues() : [];

  for (let i = 0; i < rows.length; i += 1) {
    if (String(rows[i][0] || "") === key) {
      metaSheet.getRange(i + 1, 2).setValue(JSON.stringify(runtimeContext));
      return;
    }
  }

  metaSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, JSON.stringify(runtimeContext)]]);
}

function parseDriveId(input) {
  if (!input) {
    return "";
  }

  const value = String(input).trim();
  if (!value) {
    return "";
  }

  const idMatch = value.match(/[\w-]{20,}/);
  return idMatch ? idMatch[0] : "";
}

function notifyUser(title, message) {
  Logger.log(title + ": " + message);
  return {
    title: title,
    message: message
  };
}
