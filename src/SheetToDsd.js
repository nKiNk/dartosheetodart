function runSheetsToDsd(spreadsheetInput) {
  const spreadsheet = resolveSpreadsheetForExport(spreadsheetInput);
  const metadata = readMetadataMap(spreadsheet);
  const runtimeContext = metadata.runtimeContext || (metadata.summary ? metadata.summary.runtimeContext : null);

  if (!runtimeContext || !runtimeContext.sourceFileId || !runtimeContext.targetFolderId) {
    throw new Error("워크시트 메타데이터에 source/target 정보가 없습니다. 웹앱에서 생성한 시트를 사용하세요.");
  }

  const exported = createDsdFromSpreadsheet(
    spreadsheet,
    runtimeContext.sourceFileId,
    runtimeContext.targetFolderId,
    metadata
  );

  notifyUser("DSD 생성 완료", "생성된 파일: " + exported.getName());
  return {
    fileId: exported.getId(),
    fileName: exported.getName(),
    fileUrl: exported.getUrl(),
    folderId: runtimeContext.targetFolderId
  };
}

function resolveSpreadsheetForExport(spreadsheetInput) {
  if (!spreadsheetInput) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      return active;
    }
    throw new Error("구글 시트 URL 또는 ID를 입력하세요.");
  }

  const spreadsheetId = parseDriveId(spreadsheetInput);
  if (!spreadsheetId) {
    throw new Error("유효한 구글 시트 URL 또는 ID를 입력하세요.");
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function createDsdFromSpreadsheet(spreadsheet, sourceDsdFileId, targetFolderId, metadata) {
  const meta = metadata || readMetadataMap(spreadsheet);
  const extracted = extractDsd(sourceDsdFileId);
  const patched = patchContentsXmlFromSheet(spreadsheet, meta);
  const outputName = spreadsheet.getName() + "_updated";

  if (patched.logs && patched.logs.length > 0) {
    for (let i = 0; i < patched.logs.length; i += 1) {
      Logger.log("PATCH_LOG: " + JSON.stringify(patched.logs[i]));
    }
  }

  return createDsdFromXml(patched.contentsXml, extracted.metaXml, extracted.resources, outputName, targetFolderId);
}

function patchContentsXmlFromSheet(spreadsheet, metadata) {
  const originalContentsXml = metadata.originalContentsXml;
  if (!originalContentsXml) {
    throw new Error("Missing original contents.xml snapshot in metadata.");
  }

  const patches = buildPatchPlan(spreadsheet, metadata);
  if (patches.plan.length === 0) {
    return {
      contentsXml: originalContentsXml,
      logs: [{ level: "info", message: "No changes detected. Export uses original contents.xml." }]
    };
  }

  const merged = applyPatchSet(originalContentsXml, patches.plan);
  const fullValidation = validatePatchedXml(merged);
  if (fullValidation.ok) {
    return {
      contentsXml: merged,
      logs: patches.logs.concat([{ level: "info", message: "Applied patches", applied: patches.plan.length }])
    };
  }

  const recovery = recoverWithSafePatches(originalContentsXml, patches.plan);
  const logs = patches.logs.concat([
    {
      level: "error",
      message: "Full patch validation failed. Falling back to safe patch mode.",
      detail: fullValidation.error
    }
  ]).concat(recovery.logs);

  return {
    contentsXml: recovery.contentsXml,
    logs: logs
  };
}

function buildPatchPlan(spreadsheet, metadata) {
  const logs = [];
  const plan = [];
  const blockMap = metadata.blocks;
  const keys = Object.keys(blockMap);

  for (let i = 0; i < keys.length; i += 1) {
    const block = blockMap[keys[i]];
    const sheet = spreadsheet.getSheetByName(block.sheetName);
    const baselineSheet = spreadsheet.getSheetByName(block.baselineSheetName || "");
    if (!sheet || !baselineSheet) {
      logs.push({ level: "warn", message: "Sheet or baseline missing", blockKey: keys[i], sheetName: block.sheetName, baselineSheetName: block.baselineSheetName });
      continue;
    }

    const workRows = extractBlockRowsFromSheet(sheet, block.marker);
    const baseRows = extractBlockRowsFromSheet(baselineSheet, block.marker);
    if (workRows === null || baseRows === null) {
      logs.push({ level: "warn", message: "Marker not found", blockKey: keys[i], marker: block.marker });
      continue;
    }

    if (hashPayload(workRows) === hashPayload(baseRows)) {
      continue;
    }

    const replacement = buildReplacementXml(block, workRows);
    plan.push({
      blockKey: keys[i],
      marker: block.marker,
      start: Number(block.absStart),
      end: Number(block.absEnd),
      replacement: replacement
    });

    logs.push({
      level: "info",
      message: "Block changed",
      blockKey: keys[i],
      marker: block.marker,
      baselineSheetName: block.baselineSheetName,
      start: Number(block.absStart),
      end: Number(block.absEnd)
    });
  }

  plan.sort(function (a, b) {
    return b.start - a.start;
  });
  return { plan: plan, logs: logs };
}

function extractBlockRowsFromSheet(sheet, marker) {
  const lastRow = Math.max(1, sheet.getLastRow());
  const width = CONFIG.DSD_DATA_END_COLUMN;
  const values = sheet.getRange(1, 1, lastRow, width).getDisplayValues();

  let start = -1;
  for (let r = 0; r < values.length; r += 1) {
    if (String(values[r][0] || "").trim() === marker) {
      start = r + 1;
      break;
    }
  }
  if (start < 0) {
    return null;
  }

  const rows = [];
  for (let r = start; r < values.length; r += 1) {
    const markerCell = String(values[r][0] || "").trim();
    if (markerCell.match(/^\[\d+_(표|서술문|제목|표지제목)_\d+\]$/)) {
      break;
    }
    // Skip the variable holding rows for text formula logic
    if (markerCell === "[IGNORE]") {
      continue;
    }
    rows.push(values[r]);
  }

  while (rows.length > 0 && rows[rows.length - 1].join("").trim() === "") {
    rows.pop();
  }

  return rows;
}

function buildReplacementXml(block, rows) {
  if (block.kind === "표") {
    const grid = [];
    for (let r = 0; r < rows.length; r += 1) {
      grid.push(rows[r].slice(1));
    }
    return gridToTableXml(normalizeGridWithSchema(grid, block.tableSchema || null), block.tableSchema || null);
  }

  if (block.kind === "제목" || block.kind === "표지제목") {
    const tag = block.kind === "표지제목" ? "COVER-TITLE" : "TITLE";
    const text = rows[0] ? String(rows[0][1] || "").trim() : "";
    return "<" + tag + ">" + escapeXml(text) + "</" + tag + ">";
  }

  const textRows = rows.map(function (row) {
    return row.slice(1);
  });
  return textRowsToParagraphs(textRows);
}

function normalizeGridWithSchema(values, schema) {
  const trimmed = trimGrid(values);
  const nonEmptyRows = trimmed.length;
  const nonEmptyCols = nonEmptyRows > 0 ? trimmed[0].length : 0;

  const schemaRows = schema && schema.rows ? schema.rows.length : 0;
  const schemaCols = inferSchemaLogicalCols(schema);
  const finalRows = Math.max(nonEmptyRows, schemaRows);
  const finalCols = Math.max(nonEmptyCols, schemaCols);

  if (finalRows === 0 || finalCols === 0) {
    return trimmed;
  }

  const out = [];
  for (let r = 0; r < finalRows; r += 1) {
    const src = values[r] || [];
    const row = [];
    for (let c = 0; c < finalCols; c += 1) {
      row.push(src[c] !== undefined ? src[c] : "");
    }
    out.push(row);
  }
  return out;
}

function inferSchemaLogicalCols(schema) {
  if (!schema || !schema.rows) {
    return 0;
  }
  let maxCols = 0;
  for (let r = 0; r < schema.rows.length; r += 1) {
    const cells = schema.rows[r] && schema.rows[r].cells ? schema.rows[r].cells : [];
    let logical = 0;
    for (let c = 0; c < cells.length; c += 1) {
      const attrs = cells[c] && cells[c].attrs ? cells[c].attrs : {};
      const colSpan = Math.max(1, parseInt(attrs.COLSPAN || "1", 10) || 1);
      logical += colSpan;
    }
    maxCols = Math.max(maxCols, logical);
  }
  return maxCols;
}

function applyPatchSet(baseXml, patchSet) {
  let output = baseXml;
  for (let i = 0; i < patchSet.length; i += 1) {
    const patch = patchSet[i];
    output = output.slice(0, patch.start) + patch.replacement + output.slice(patch.end);
  }
  return output;
}

function validatePatchedXml(xmlText) {
  try {
    XmlService.parse(xmlText);
    return { ok: true };
  } catch (error) {
    return { ok: false, error: String(error) };
  }
}

function recoverWithSafePatches(originalContentsXml, patchSet) {
  const accepted = [];
  const logs = [];

  for (let i = patchSet.length - 1; i >= 0; i -= 1) {
    const candidate = patchSet[i];
    const candidateSet = accepted.concat([candidate]).sort(function (a, b) {
      return b.start - a.start;
    });
    const testXml = applyPatchSet(originalContentsXml, candidateSet);
    const validation = validatePatchedXml(testXml);
    if (validation.ok) {
      accepted.push(candidate);
      logs.push({ level: "info", message: "Patch accepted in recovery", blockKey: candidate.blockKey, marker: candidate.marker });
    } else {
      logs.push({
        level: "error",
        message: "Patch rejected in recovery",
        blockKey: candidate.blockKey,
        marker: candidate.marker,
        error: validation.error
      });
    }
  }

  const finalSet = accepted.sort(function (a, b) {
    return b.start - a.start;
  });
  const finalXml = applyPatchSet(originalContentsXml, finalSet);
  return { contentsXml: finalXml, logs: logs };
}

function readMetadataMap(spreadsheet) {
  const metaSheet = spreadsheet.getSheetByName(CONFIG.METADATA_SHEET_NAME);
  if (!metaSheet) {
    throw new Error("Metadata sheet not found: " + CONFIG.METADATA_SHEET_NAME);
  }

  const lastRow = metaSheet.getLastRow();
  if (lastRow < 1) {
    throw new Error("Metadata sheet is empty.");
  }

  const rows = metaSheet.getRange(1, 1, lastRow, 2).getValues();
  const sections = {};
  const blocks = {};
  const chunks = [];
  let summary = {};
  let runtimeContext = null;

  for (let i = 0; i < rows.length; i += 1) {
    const key = String(rows[i][0] || "").trim();
    const value = rows[i][1];
    if (!key) {
      continue;
    }

    if (key === "__SUMMARY__") {
      summary = safeJsonParse(value);
      continue;
    }
    if (key === "__RUNTIME_CONTEXT__") {
      runtimeContext = safeJsonParse(value);
      continue;
    }
    if (key.indexOf("SECTION::") === 0) {
      sections[key] = safeJsonParse(value);
      continue;
    }
    if (key.indexOf("BLOCK::") === 0) {
      blocks[key] = safeJsonParse(value);
      continue;
    }
    if (key.indexOf("__CONTENTS_CHUNK__::") === 0) {
      const idx = Number(key.split("::")[1]);
      chunks[idx] = String(value || "");
    }
  }

  return {
    summary: summary,
    runtimeContext: runtimeContext,
    sections: sections,
    blocks: blocks,
    originalContentsXml: chunks.join("")
  };
}

function safeJsonParse(value) {
  if (typeof value === "object" && value !== null) {
    return value;
  }
  return JSON.parse(String(value || "{}"));
}

function trimGrid(values) {
  let lastRow = -1;
  let lastCol = -1;

  for (let r = 0; r < values.length; r += 1) {
    for (let c = 0; c < values[r].length; c += 1) {
      if (String(values[r][c]).trim() !== "") {
        lastRow = Math.max(lastRow, r);
        lastCol = Math.max(lastCol, c);
      }
    }
  }

  if (lastRow < 0 || lastCol < 0) {
    return [];
  }

  const out = [];
  for (let r = 0; r <= lastRow; r += 1) {
    out.push(values[r].slice(0, lastCol + 1));
  }
  return out;
}

function textRowsToParagraphs(rows) {
  const lines = [];
  let paragraph = [];

  for (let i = 0; i < rows.length; i += 1) {
    const line = String(rows[i] && rows[i][0] ? rows[i][0] : "").trim();
    if (line) {
      paragraph.push(line);
    } else if (paragraph.length > 0) {
      lines.push("<P>" + escapeXml(paragraph.join(" ")) + "</P>");
      paragraph = [];
    }
  }
  if (paragraph.length > 0) {
    lines.push("<P>" + escapeXml(paragraph.join(" ")) + "</P>");
  }

  return lines.join("\n");
}

function gridToTableXml(grid, schema) {
  if (!grid || grid.length === 0) {
    return "<TABLE></TABLE>";
  }

  const tableAttrs = schema ? schema.attrs : {};
  let xml = "<TABLE" + objectToXmlAttrs(tableAttrs) + ">\n";
  if (schema && schema.colgroup) {
    xml += schema.colgroup + "\n";
  }
  xml += "<TBODY>\n";

  // Create a mask to track which cells are covered by rowspans/colspans
  const skipMask = [];
  for (let i = 0; i < grid.length; i++) {
    skipMask.push(new Array(grid[i].length).fill(false));
  }

  for (let r = 0; r < grid.length; r += 1) {
    const rowSchema = schema && schema.rows && schema.rows[r] ? schema.rows[r] : null;
    const trAttrs = rowSchema ? rowSchema.attrs : CONFIG.DART_TABLE_DEFAULTS.TR;
    xml += "<TR" + objectToXmlAttrs(trAttrs) + ">\n";

    let schemaCellIdx = 0;
    
    for (let c = 0; c < grid[r].length; c += 1) {
      if (skipMask[r][c]) continue; // Skip cells that are merged into another cell

      const cellSchema = rowSchema && rowSchema.cells && rowSchema.cells[schemaCellIdx] ? rowSchema.cells[schemaCellIdx] : null;
      const tag = cellSchema ? cellSchema.tag : "TD";
      const cellAttrs = cellSchema ? cellSchema.attrs : CONFIG.DART_TABLE_DEFAULTS.TD;
      
      // Update skip mask based on colspan and rowspan
      const colSpan = parseInt(cellAttrs.COLSPAN || "1", 10);
      const rowSpan = parseInt(cellAttrs.ROWSPAN || "1", 10);
      
      for (let rs = 0; rs < rowSpan; rs++) {
        for (let cs = 0; cs < colSpan; cs++) {
          if (rs === 0 && cs === 0) continue;
          if (r + rs < skipMask.length && c + cs < skipMask[r + rs].length) {
            skipMask[r + rs][c + cs] = true;
          }
        }
      }

      const value = escapeXml(String(grid[r][c] || ""));
      xml += "<" + tag + objectToXmlAttrs(cellAttrs) + ">" + value + "</" + tag + ">\n";
      
      schemaCellIdx++;
    }

    xml += "</TR>\n";
  }

  xml += "</TBODY>\n</TABLE>";
  return xml;
}

function objectToXmlAttrs(obj) {
  const keys = Object.keys(obj || {});
  if (keys.length === 0) {
    return "";
  }
  const parts = [];
  for (let i = 0; i < keys.length; i += 1) {
    parts.push(keys[i] + '="' + escapeXml(String(obj[keys[i]])) + '"');
  }
  return " " + parts.join(" ");
}

function escapeXml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
