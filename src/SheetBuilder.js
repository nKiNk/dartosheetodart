function createSpreadsheetFromSections(parsed, spreadsheetName, originalContentsXml, runtimeContext, templateId, targetFolder) {
  let spreadsheet;
  if (templateId && targetFolder) {
    const file = DriveApp.getFileById(templateId).makeCopy(spreadsheetName, targetFolder);
    spreadsheet = SpreadsheetApp.openById(file.getId());
  } else {
    spreadsheet = SpreadsheetApp.create(spreadsheetName);
  }

  const defaultSheet = spreadsheet.getSheets()[0];
  defaultSheet.setName("작업안내");
  renderGuideSheet(defaultSheet, runtimeContext || {});

  const metaSheet = spreadsheet.getSheetByName(CONFIG.METADATA_SHEET_NAME) || spreadsheet.insertSheet(CONFIG.METADATA_SHEET_NAME);
  metaSheet.hideSheet();

  let sectionOrder = 1;
  const sectionRecords = [];
  const blockRecords = [];

  function createContentSheet(baseName) {
    const name = createUniqueSheetName(spreadsheet, baseName);
    return spreadsheet.insertSheet(name);
  }

  const priority = ["표지", "목차", "감사보고서", "(첨부)연결재무제표"];
  const sections = parsed.sections || [];
  const notes = parsed.notes || [];

  for (let i = 0; i < priority.length; i += 1) {
    const section = sections.find(function (s) {
      return s.sheetName === priority[i];
    });
    if (!section) {
      continue;
    }
    sectionOrder = renderSectionToSheet(spreadsheet, createContentSheet, section, sectionOrder, sectionRecords, blockRecords);
  }

  for (let i = 0; i < sections.length; i += 1) {
    if (priority.indexOf(sections[i].sheetName) !== -1) {
      continue;
    }
    sectionOrder = renderSectionToSheet(spreadsheet, createContentSheet, sections[i], sectionOrder, sectionRecords, blockRecords);
  }

  for (let i = 0; i < notes.length; i += 1) {
    sectionOrder = renderSectionToSheet(spreadsheet, createContentSheet, notes[i], sectionOrder, sectionRecords, blockRecords);
  }

  writeMetadataSnapshot(metaSheet, {
    version: 4,
    createdAt: new Date().toISOString(),
    dataEndColumn: CONFIG.DSD_DATA_END_COLUMN,
    userAreaStartColumn: CONFIG.USER_AREA_START_COLUMN,
    sectionCount: sectionRecords.length,
    blockCount: blockRecords.length,
    runtimeContext: runtimeContext || {}
  }, originalContentsXml || "", sectionRecords, blockRecords);

  return spreadsheet;
}

function renderGuideSheet(sheet, runtimeContext) {
  const webAppUrl = runtimeContext.webAppUrl || "웹앱 URL이 없습니다. (배포 URL 확인 필요)";
  const testWebAppUrl = "https://script.google.com/macros/s/AKfycby6aHKq7d0_iOkhlF31FjBa5jy4sQI4Ez8Nm6LU2E_DKyRl74PstiiFkX7QtgZHGcV9wA/exec";
  const rows = [
    [CONFIG.APP_NAME],
    ["DART DSD <-> Google Sheets 양방향 변환 작업안내"],
    [""],
    ["[기본 정보]"],
    ["웹앱 주소"],
    [webAppUrl],
    ["테스트 웹앱 주소(참고)"],
    [testWebAppUrl],
    ["현재 작업 시트 주소"],
    [sheet.getParent().getUrl()],
    [""],
    ["[사용 방법 1: 웹앱 사용(권장)]"],
    ["1) 웹앱에서 DSD 파일 업로드 후 작업 시트 생성"],
    ["2) 시트 수정은 DSD 영역(A~M)에서 수행"],
    ["3) 우측 작업영역(N열 이후)은 메모/검증용으로 사용(DSD 반영 제외)"],
    ["4) 웹앱의 Sheet -> DSD 변환 실행 후 target 폴더 결과 확인"],
    [""],
    ["[사용 방법 2: 코드 복사 사용]"],
    ["1) Apps Script 프로젝트에 src 코드 반영"],
    ["2) onOpen 메뉴 또는 웹앱으로 변환 작업 실행"],
    [""],
    ["[사용 방법 3: 라이브러리 사용]"],
    ["호출 권장 API: apiInitializeWorkspace / apiUploadDsdAndCreateWorksheet / apiConvertSheetToDsd"],
    [""],
    ["[최근 반영 사항]"],
    ["- 북마크 기반 주석 분리 우선 적용(없으면 fallback)"],
    ["- 서술문 숫자 추출 시 다중 숫자 참조 오프셋 정확화"],
    ["- 표 span/빈셀 및 시트 렌더링 안정성 개선"],
    [""],
    ["[주의]"],
    ["- 본 도구는 외부 DB/분석 서버에 입력값을 기록하지 않습니다."],
    ["- 변환에 필요한 정보는 현재 스프레드시트와 사용자 Drive 폴더에만 저장됩니다."],
    ["- [IGNORE] 행은 수식 보조값이므로 삭제 시 문장 수식이 깨질 수 있습니다."],
    ["- __METADATA__ 및 __BASE__ 시트는 숨김 유지(수정/삭제 금지)."],
    ["- 금감원 로고는 데모용 요소입니다. 실사용 배포 시 자체 로고로 교체하세요."]
  ];

  sheet.clear();
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  sheet.getRange("A1:A2").setFontWeight("bold");
  sheet.getRange("A4:A4").setFontWeight("bold");
  sheet.getRange("A10:A10").setFontWeight("bold");
  sheet.getRange("A16:A16").setFontWeight("bold");
  sheet.getRange("A20:A20").setFontWeight("bold");
  sheet.getRange("A23:A23").setFontWeight("bold");
  sheet.getRange("A30:A30").setFontWeight("bold");
  sheet.setColumnWidth(1, 900);
  sheet.getRange(1, 1, rows.length, 1).setWrap(true);
  sheet.setFrozenRows(2);
}

function renderSectionToSheet(spreadsheet, createContentSheet, section, sectionOrder, sectionRecords, blockRecords) {
  const sheet = createContentSheet(section.sheetName);
  const rendered = renderBlocksToSheet(sheet, section.blocks || []);
  markUserWorkspace(sheet);
  applyWorkAreaBoundary(sheet, Math.max(1, rendered.usedRows));
  const baselineSheet = createBaselineSheet(spreadsheet, sheet, rendered.usedRows);

  const sectionRecord = {
    sectionKey: section.sectionKey,
    sectionType: section.sectionType,
    sheetName: sheet.getName(),
    baselineSheetName: baselineSheet.getName(),
    title: section.title,
    order: sectionOrder,
    sourceStart: section.sourceStart,
    sourceEnd: section.sourceEnd
  };
  sectionRecords.push(sectionRecord);

  for (let i = 0; i < rendered.blockRecords.length; i += 1) {
    const blockRecord = rendered.blockRecords[i];
    blockRecord.sectionKey = section.sectionKey;
    blockRecord.sectionType = section.sectionType;
    blockRecord.sheetName = sheet.getName();
    blockRecord.baselineSheetName = baselineSheet.getName();
    blockRecord.key = buildBlockKey(sheet.getName(), blockRecord.marker);
    blockRecords.push(blockRecord);
  }

  return sectionOrder + 1;
}

function createBaselineSheet(spreadsheet, sourceSheet, usedRows) {
  const baselineName = createUniqueSheetName(spreadsheet, "__BASE__" + sourceSheet.getName());
  const baselineSheet = spreadsheet.insertSheet(baselineName);
  const rows = Math.max(1, usedRows || sourceSheet.getLastRow());
  const cols = CONFIG.DSD_DATA_END_COLUMN;
  ensureSheetCapacity(baselineSheet, rows, cols);

  const values = sourceSheet.getRange(1, 1, rows, cols).getDisplayValues();
  baselineSheet.getRange(1, 1, rows, cols).setValues(values);
  baselineSheet.hideSheet();
  return baselineSheet;
}

function renderBlocksToSheet(sheet, blocks) {
  const maxCols = CONFIG.DSD_DATA_END_COLUMN;
  const allRows = [];
  const allWeights = [];
  const blockRecords = [];
  
  // To track ranges for formatting later
  const formattingQueue = [];

  for (let i = 0; i < blocks.length; i += 1) {
    const block = blocks[i];
    const contentRows = blockToSheetRows(block, maxCols);
    
    // Add Marker row
    const markerRowIdx = allRows.length + 1; // 1-based index for sheets
    allRows.push(buildBlankRow(maxCols));
    allRows[allRows.length - 1][0] = block.marker || "";
    allWeights.push(buildWeightRow(maxCols, 0));

    // Track content start for formatting
    const contentStartIdx = allRows.length + 1;
    
    for (let r = 0; r < contentRows.length; r += 1) {
      allRows.push(contentRows[r]);
      allWeights.push(buildWeightRow(maxCols, -1));
    }
    
    const contentEndIdx = allRows.length;

    // Save formatting tasks
    if (block.kind === "표" && contentRows.length > 0) {
      formattingQueue.push({
        type: "TABLE",
        startRow: contentStartIdx,
        numRows: contentRows.length,
        block: block
      });
    } else if (contentRows.length > 0) {
      // For narratives, titles, etc.
      formattingQueue.push({
        type: "TEXT",
        startRow: contentStartIdx,
        numRows: contentRows.length,
        contentRows: contentRows
      });
    }

    allRows.push(buildBlankRow(maxCols));
    allWeights.push(buildWeightRow(maxCols, -1));

    blockRecords.push({
      marker: block.marker,
      kind: block.kind,
      absStart: block.absStart,
      absEnd: block.absEnd,
      tableSchema: block.kind === "표" ? extractTableSchema(block) : null,
      originalPayloadHash: hashPayload(contentRows)
    });
  }

  if (allRows.length === 0) {
    return { usedRows: 0, blockRecords: blockRecords };
  }

  ensureSheetCapacity(sheet, allRows.length, maxCols);
  
  // Set main values
  const range = sheet.getRange(1, 1, allRows.length, maxCols);
  
  // If the cell contains a formula (starts with =), we must use setFormulas, 
  // otherwise setValues. Since we have a mix, we must separate them or use setValues first, 
  // then overwrite formulas.
  range.setValues(allRows);
  
  // Apply formulas explicitly where needed
  for (let r = 0; r < allRows.length; r++) {
    for (let c = 0; c < allRows[r].length; c++) {
      if (typeof allRows[r][c] === 'string' && allRows[r][c].startsWith('=')) {
        sheet.getRange(r + 1, c + 1).setFormula(allRows[r][c]);
      }
    }
  }
  
  range.setFontWeights(allWeights);

  // Apply UX Formatting (Borders, Backgrounds, Merges)
  applyUXFormatting(sheet, formattingQueue, maxCols);

  return { usedRows: allRows.length, blockRecords: blockRecords };
}

function applyUXFormatting(sheet, formattingQueue, maxCols) {
  for (let i = 0; i < formattingQueue.length; i++) {
    const task = formattingQueue[i];
    
    if (task.type === "TABLE") {
      const tableRows = task.block.rows || [];
      if (tableRows.length === 0) continue;

      const layout = buildTableLayout(tableRows, maxCols);
      const applyCols = Math.min(layout.maxLogicalCols, maxCols - 1);
      if (applyCols <= 0) continue;

      const tableVisualRows = Math.min(task.numRows, tableRows.length);
      const tableRange = sheet.getRange(task.startRow, 2, tableVisualRows, applyCols); // Start at B (col 2)
      tableRange.breakApart();

      tableRange.setBorder(true, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      
      for (let p = 0; p < layout.placements.length; p += 1) {
        const placed = layout.placements[p];
        const mergeRows = Math.max(1, Math.min(placed.rowSpan, tableVisualRows - placed.row));
        const mergeCols = Math.max(1, Math.min(placed.colSpan, applyCols - placed.col + 1));
        if (mergeRows > 1 || mergeCols > 1) {
          sheet.getRange(task.startRow + placed.row, placed.col + 1, mergeRows, mergeCols).merge();
        }
        if (placed.cell && placed.cell.tag === "TH") {
          const cellRange = sheet.getRange(task.startRow + placed.row, placed.col + 1, mergeRows, mergeCols);
          cellRange.setBackground("#f3f3f3");
          cellRange.setFontWeight("bold");
        }
      }

      for (let r = 0; r < task.numRows; r += 1) {
        const marker = String(sheet.getRange(task.startRow + r, 1).getDisplayValue() || "").trim();
        if (marker === "[IGNORE]") {
          const varRange = sheet.getRange(task.startRow + r, 2);
          varRange.setBackground("#e8f0fe");
          varRange.setHorizontalAlignment("right");
          varRange.setNumberFormat("#,##0.####");
        }
      }
    } 
    else if (task.type === "TEXT") {
      const applyCols = Math.min(maxCols - 1, 10);
      
      for (let r = 0; r < task.numRows; r++) {
        const rowData = task.contentRows[r];
        const marker = String(rowData[0] || "");
        
        // Style variables nicely
        if (marker === "[IGNORE]") {
          const varRange = sheet.getRange(task.startRow + r, 2); // B column
          varRange.setBackground("#e8f0fe"); // light blue
          varRange.setHorizontalAlignment("right");
          // Format based on value
          const val = rowData[1];
          if (typeof val === 'number') {
            // Percent vs Currency is handled in formula, but we can set generic format here
            varRange.setNumberFormat("#,##0.####");
          }
        } else if (String(rowData[1] || "").trim() !== "") {
          // Merge text content
          const textRange = sheet.getRange(task.startRow + r, 2, 1, applyCols);
          textRange.mergeAcross();
          textRange.setWrap(true);
          textRange.setHorizontalAlignment("left");
          textRange.setVerticalAlignment("top");
        }
      }
    }
  }
}

function blockToSheetRows(block, maxCols) {
  const rows = [];
  if (block.kind === "표") {
    const tableRows = block.rows || [];
    const layout = buildTableLayout(tableRows, maxCols);
    const grid = layout.grid;
    
    const extraVarRows = [];

    if (layout.placedCellCount === 1 && layout.placements.length === 1) {
      const only = layout.placements[0];
      if (only.colSpan === 1 && only.rowSpan === 1) {
        const sentence = String(grid[only.row][only.col] || "");
        const parsed = parseSentenceForNumbers(sentence);
        if (parsed.hasNumber) {
          grid[only.row][only.col] = parsed.formula;
          for (let v = 0; v < parsed.variables.length; v += 1) {
            const varRow = buildBlankRow(maxCols);
            varRow[0] = "[IGNORE]";
            varRow[1] = parsed.variables[v].value;
            extraVarRows.push(varRow);
          }
        }
      }
    }

    for (let r = 0; r < grid.length; r++) {
      rows.push(grid[r]);
    }
    
    // Push extra variable rows below the table if it was a single cell table
    for (let i = 0; i < extraVarRows.length; i++) {
      rows.push(extraVarRows[i]);
    }
    
    return rows;
  }

  const textRows = splitTextToSheetRows(block.text || "", maxCols);
  for (let i = 0; i < textRows.length; i += 1) {
    rows.push(textRows[i]);
  }
  return rows;
}

function splitTextToSheetRows(text, maxCols) {
  const rows = [];
  const paragraphs = String(text || "").split("\n").filter(function (line) {
    return line.trim() !== "";
  });

  for (let p = 0; p < paragraphs.length; p += 1) {
    const value = paragraphs[p].trim();
    if (!value) {
      continue;
    }
    
    // Split paragraph into sentences by looking for period followed by space
    const sentences = value.split(/(?<=[.?!])\s+/);
    
    for (let s = 0; s < sentences.length; s += 1) {
      const sentence = sentences[s].trim();
      if (!sentence) continue;
      
      const parsed = parseSentenceForNumbers(sentence);
      
      const row = buildBlankRow(maxCols);
      if (parsed.hasNumber) {
        row[1] = parsed.formula;
        rows.push(row);
        
        // Add variable rows below the formula
        for (let v = 0; v < parsed.variables.length; v += 1) {
          const varRow = buildBlankRow(maxCols);
          varRow[0] = "[IGNORE]";
          varRow[1] = parsed.variables[v].value;
          rows.push(varRow);
        }
      } else {
        row[1] = sentence;
        rows.push(row);
      }
    }

    if (p < paragraphs.length - 1) {
      rows.push(buildBlankRow(maxCols));
    }
  }

  return rows;
}

function parseSentenceForNumbers(sentence) {
  let hasNumber = false;
  const variables = [];
  let foundMatches = [];
  
  const isMathContext = function(idx) {
    const prev = sentence.substring(Math.max(0, idx - 10), idx);
    const next = sentence.substring(idx, Math.min(sentence.length, idx + 10));
    return /[=+\-×÷]/.test(prev) || /[=+\-×÷]/.test(next);
  };
  
  const rawNumRegex = /([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]+)?)/g;
  let rMatch;
  
  while ((rMatch = rawNumRegex.exec(sentence)) !== null) {
    const numStr = rMatch[1];
    const startIndex = rMatch.index;
    const endIndex = rMatch.index + numStr.length;
    
    const charBefore = startIndex > 0 ? sentence[startIndex - 1] : '';
    const charAfter = endIndex < sentence.length ? sentence[endIndex] : '';
    const twoBefore = startIndex > 1 ? sentence.substring(startIndex - 2, startIndex) : charBefore;
    const nextChars = sentence.substring(endIndex, endIndex + 3).trim();
    
    // Exclusions
    if (charBefore === '(' && charAfter === ')') continue;
    if (twoBefore === '(*' && charAfter === ')') continue;
    if (twoBefore.endsWith('*') && (charAfter === ' ' || charAfter === ',' || charAfter === '.')) continue;
    if (charBefore === '제' || twoBefore.endsWith('제')) continue;
    if (charAfter === '호' || charAfter === '조' || charAfter === '항') continue;
    if (charAfter === '년' || charAfter === '월' || charAfter === '일') continue;
    if (isInsideQuotedRange(sentence, startIndex, endIndex)) continue;
    const tokenWindow = sentence.substring(Math.max(0, startIndex - 8), Math.min(sentence.length, endIndex + 2));
    if (/[A-Za-z가-힣]-\d/.test(tokenWindow)) continue;
    if (sentence.substring(Math.max(0, startIndex - 5), startIndex).indexOf("회제이") !== -1) continue; // Exclude 회제이 specific references
    
    // Inclusions
    const isCurrency = nextChars.startsWith('원') || nextChars.startsWith('천원') || nextChars.startsWith('만원') || nextChars.startsWith('백만원') || nextChars.startsWith('억원');
    const isPercent = nextChars.startsWith('%');
    const isStock = nextChars.startsWith('주');
    const isMath = isMathContext(startIndex);
    
    if (isCurrency || isPercent || isStock || isMath) {
      foundMatches.push({
        numStr: numStr,
        startIndex: startIndex,
        endIndex: endIndex,
        isPercent: isPercent,
        isStock: isStock
      });
    }
  }
  
  if (foundMatches.length === 0) {
    return { hasNumber: false, formula: "", variables: [] };
  }
  
  let currentPos = 0;
  let formulaStr = '="';
  
  for (let i = 0; i < foundMatches.length; i++) {
    const m = foundMatches[i];
    const textPart = sentence.substring(currentPos, m.startIndex).replace(/"/g, '""');
    formulaStr += textPart + '"&';
    
    let val = parseFloat(m.numStr.replace(/,/g, ''));
    if (isNaN(val)) val = m.numStr;
    const rowOffset = variables.length + 1;
    
    variables.push({
      value: val,
      isPercent: m.isPercent
    });
    
    if (m.isPercent) {
      formulaStr += 'TEXT(INDIRECT("R[' + rowOffset + ']C[0]", FALSE), "General")&"';
    } else {
      formulaStr += 'TEXT(INDIRECT("R[' + rowOffset + ']C[0]", FALSE), "#,##0")&"';
    }
    
    currentPos = m.endIndex;
  }
  
  const lastPart = sentence.substring(currentPos).replace(/"/g, '""');
  formulaStr += lastPart + '"';
  
  formulaStr = formulaStr.replace('=""&', '=');
  formulaStr = formulaStr.replace('&""', '');
  
  return {
    hasNumber: true,
    formula: formulaStr,
    variables: variables
  };
}

function buildTableLayout(tableRows, maxCols) {
  const grid = [];
  const occupied = [];
  for (let i = 0; i < tableRows.length; i += 1) {
    grid.push(buildBlankRow(maxCols));
    occupied.push(new Array(maxCols).fill(false));
  }

  const placements = [];
  let maxLogicalCols = 0;
  let placedCellCount = 0;

  for (let r = 0; r < tableRows.length; r += 1) {
    const cells = tableRows[r].cells || [];
    let cIdx = 0;

    for (let c = 0; c < cells.length; c += 1) {
      while (cIdx + 1 < maxCols && occupied[r][cIdx + 1]) {
        cIdx += 1;
      }
      if (cIdx + 1 >= maxCols) {
        break;
      }

      const cell = cells[c] || {};
      const colSpan = Math.max(1, parseInt(cell.attrs && cell.attrs.COLSPAN ? cell.attrs.COLSPAN : "1", 10) || 1);
      const rowSpan = Math.max(1, parseInt(cell.attrs && cell.attrs.ROWSPAN ? cell.attrs.ROWSPAN : "1", 10) || 1);
      const col = cIdx + 1;

      grid[r][col] = cell.value || "";
      occupied[r][col] = true;
      placedCellCount += 1;
      placements.push({ row: r, col: col, colSpan: colSpan, rowSpan: rowSpan, cell: cell });
      maxLogicalCols = Math.max(maxLogicalCols, col + colSpan - 1);

      for (let rs = 0; rs < rowSpan; rs += 1) {
        for (let cs = 0; cs < colSpan; cs += 1) {
          if (rs === 0 && cs === 0) {
            continue;
          }
          if (r + rs < occupied.length && col + cs < maxCols) {
            occupied[r + rs][col + cs] = true;
          }
        }
      }

      cIdx += colSpan;
    }
  }

  return {
    grid: grid,
    placements: placements,
    maxLogicalCols: maxLogicalCols,
    placedCellCount: placedCellCount
  };
}

function isInsideQuotedRange(text, startIndex, endIndex) {
  return isInsideQuote(text, startIndex, "'") ||
    isInsideQuote(text, Math.max(startIndex, endIndex - 1), "'") ||
    isInsideQuote(text, startIndex, '"') ||
    isInsideQuote(text, Math.max(startIndex, endIndex - 1), '"');
}

function isInsideQuote(text, idx, quoteChar) {
  let open = false;
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === quoteChar) {
      open = !open;
      continue;
    }
    if (i === idx) {
      return open;
    }
  }
  return false;
}

function buildBlankRow(maxCols) {
  const row = [];
  for (let c = 0; c < maxCols; c += 1) {
    row.push("");
  }
  return row;
}

function buildWeightRow(maxCols, markerIndex) {
  const row = [];
  for (let c = 0; c < maxCols; c += 1) {
    row.push(c === markerIndex ? "bold" : "normal");
  }
  return row;
}

function ensureSheetCapacity(sheet, rows, cols) {
  if (rows > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  }
  if (cols > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  }
}

function extractTableSchema(block) {
  const schema = {
    attrs: block.tableAttrs || {},
    rows: []
  };
  const rows = block.rows || [];
  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r];
    const cells = row.cells || [];
    schema.rows.push({
      attrs: row.attrs || {},
      cells: cells.map(function (cell) {
        return {
          tag: cell.tag || "TD",
          attrs: cell.attrs || {}
        };
      })
    });
  }
  const colgroupMatch = block.rawXml ? block.rawXml.match(/<COLGROUP[^>]*>[\s\S]*?<\/COLGROUP>/i) : null;
  if (colgroupMatch) {
    schema.colgroup = colgroupMatch[0];
  }
  return schema;
}

function hashPayload(rows) {
  const normalized = JSON.stringify(rows);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalized, Utilities.Charset.UTF_8);
  return bytesToHex(bytes);
}

function bytesToHex(bytes) {
  const chars = [];
  for (let i = 0; i < bytes.length; i += 1) {
    const b = (bytes[i] + 256) % 256;
    chars.push((b < 16 ? "0" : "") + b.toString(16));
  }
  return chars.join("");
}

function buildBlockKey(sheetName, marker) {
  return "BLOCK::" + sheetName + "::" + marker;
}

function writeMetadataSnapshot(metaSheet, summary, originalContentsXml, sectionRecords, blockRecords) {
  const rows = [];
  rows.push(["__SUMMARY__", JSON.stringify(summary)]);

  for (let i = 0; i < sectionRecords.length; i += 1) {
    rows.push(["SECTION::" + sectionRecords[i].sectionKey, JSON.stringify(sectionRecords[i])]);
  }

  for (let i = 0; i < blockRecords.length; i += 1) {
    rows.push([blockRecords[i].key, JSON.stringify(blockRecords[i])]);
  }

  const chunkSize = 40000;
  const chunkCount = Math.ceil(originalContentsXml.length / chunkSize);
  rows.push(["__CONTENTS_CHUNKS__", String(chunkCount)]);
  for (let i = 0; i < chunkCount; i += 1) {
    const chunk = originalContentsXml.slice(i * chunkSize, (i + 1) * chunkSize);
    rows.push(["__CONTENTS_CHUNK__::" + i, chunk]);
  }

  metaSheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function markUserWorkspace(sheet) {
  sheet.getRange(1, CONFIG.USER_AREA_START_COLUMN).setValue("[User Work Space]: excluded in DSD");
}

function applyWorkAreaBoundary(sheet, usedRows) {
  const maxRows = Math.max(usedRows || 1, sheet.getMaxRows());
  const dsdBoundaryRange = sheet.getRange(1, CONFIG.DSD_DATA_END_COLUMN, maxRows, 1);
  dsdBoundaryRange.setBorder(false, false, false, false, false, false);
  dsdBoundaryRange.setBorder(null, null, null, true, null, null, CONFIG.BOUNDARY_LINE_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  const userBoundaryRange = sheet.getRange(1, CONFIG.USER_AREA_START_COLUMN, maxRows, 1);
  userBoundaryRange.setBorder(false, false, false, false, false, false);
  userBoundaryRange.setBorder(null, true, null, null, null, null, CONFIG.BOUNDARY_LINE_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const previewWidth = Math.max(1, CONFIG.USER_AREA_PREVIEW_COLUMNS);
  if (CONFIG.USER_AREA_START_COLUMN + previewWidth - 1 <= sheet.getMaxColumns()) {
    sheet.getRange(1, CONFIG.USER_AREA_START_COLUMN, maxRows, previewWidth).setBackground("#f9fafb");
  }
}

function createUniqueSheetName(spreadsheet, baseName) {
  let candidate = String(baseName).substring(0, 99);
  let serial = 1;
  while (spreadsheet.getSheetByName(candidate)) {
    candidate = (baseName + "_" + serial).substring(0, 99);
    serial += 1;
  }
  return candidate;
}
