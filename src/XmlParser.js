function parseDsdStructure(xmlString) {
  const parsed = {
    sections: [],
    notes: []
  };
  const inlineNotes = [];

  const sectionRegex = /<(SECTION-1|SECTION-2|COVER|TOC)\b[^>]*>[\s\S]*?<\/\1>/gi;
  let sectionMatch = sectionRegex.exec(xmlString);
  let sectionIndex = 0;

  while (sectionMatch !== null) {
    const tag = sectionMatch[1].toUpperCase();
    const fullXml = sectionMatch[0];
    const absStart = sectionMatch.index;
    const absEnd = sectionRegex.lastIndex;
    const sectionKeyBase = "SEC_" + sectionIndex + "_" + tag;

    if (tag === "SECTION-2") {
      const noteGroups = splitSection2Notes(fullXml, absStart, sectionKeyBase);
      for (let i = 0; i < noteGroups.length; i += 1) {
        parsed.notes.push(noteGroups[i]);
      }
    } else {
      const sectionRecord = buildSectionRecord(tag, fullXml, absStart, absEnd, sectionKeyBase);
      if (sectionRecord) {
        if (sectionRecord.sectionType === "SECTION1_ATTACHED") {
          const splitResult = splitAttachedFinancialAndInlineNotes(sectionRecord);
          if (splitResult.financialBlocks.length > 0) {
            sectionRecord.blocks = splitResult.financialBlocks;
            parsed.sections.push(sectionRecord);
          }
          appendUniqueNotes(inlineNotes, splitResult.inlineNotes);
        } else {
          parsed.sections.push(sectionRecord);
        }
      }
    }

    sectionIndex += 1;
    sectionMatch = sectionRegex.exec(xmlString);
  }

  if (parsed.notes.length === 0) {
    const noteSectionRegex = /<SECTION-2\b[^>]*>[\s\S]*?<\/SECTION-2>/gi;
    let noteMatch = noteSectionRegex.exec(xmlString);
    let noteSectionIndex = 0;
    while (noteMatch !== null) {
      const sectionKeyBase = "SEC2_FALLBACK_" + noteSectionIndex;
      const fallbackNotes = splitSection2Notes(noteMatch[0], noteMatch.index, sectionKeyBase);
      const bookmarkNotes = fallbackNotes.filter(function (note) {
        return String(note.noteNumber) !== "0";
      });
      appendUniqueNotes(parsed.notes, bookmarkNotes);
      noteSectionIndex += 1;
      noteMatch = noteSectionRegex.exec(xmlString);
    }
  }

  if (parsed.notes.length === 0 && inlineNotes.length > 0) {
    appendUniqueNotes(parsed.notes, inlineNotes);
  }

  parsed.notes = dedupeNotes(parsed.notes);
  return parsed;
}

function splitAttachedFinancialAndInlineNotes(sectionRecord) {
  const blocks = sectionRecord.blocks || [];
  const financialBlocks = [];
  const inlineNotes = [];
  const noteTitleRegex = /^(?:주석\s*)?(\d+)[\.)]\s+(.+)$/;
  const statementNames = (CONFIG && CONFIG.STATEMENTS ? CONFIG.STATEMENTS : []).map(function (name) {
    return String(name || "").replace(/\s+/g, "");
  });
  let currentInlineNote = null;

  for (let i = 0; i < blocks.length; i += 1) {
    const block = blocks[i];
    const statementTitle = extractStatementTitleFromTable(block);
    if (statementTitle && statementNames.indexOf(statementTitle) !== -1) {
      currentInlineNote = null;
      financialBlocks.push(block);
      continue;
    }

    const text = block.kind === "서술문" ? String(block.text || "").replace(/\s+/g, " ").trim() : "";
    const noteMatch = text ? noteTitleRegex.exec(text) : null;
    if (noteMatch) {
      currentInlineNote = {
        sectionKey: sectionRecord.sectionKey + "_INLINE_NOTE_" + noteMatch[1] + "_" + (inlineNotes.length + 1),
        sectionType: "SECTION2_NOTE",
        noteNumber: noteMatch[1],
        sheetName: "주석_" + noteMatch[1],
        title: text,
        sourceStart: block.absStart,
        sourceEnd: block.absEnd,
        sourceXml: block.rawXml || "",
        blocks: [block]
      };
      inlineNotes.push(currentInlineNote);
      continue;
    }

    if (currentInlineNote) {
      currentInlineNote.blocks.push(block);
      currentInlineNote.sourceEnd = block.absEnd;
      currentInlineNote.sourceXml += block.rawXml || "";
    } else {
      financialBlocks.push(block);
    }
  }

  return {
    financialBlocks: financialBlocks,
    inlineNotes: inlineNotes
  };
}

function extractStatementTitleFromTable(block) {
  if (!block || block.kind !== "표") {
    return "";
  }
  const rows = block.rows || [];
  if (rows.length === 0 || !rows[0].cells || rows[0].cells.length === 0) {
    return "";
  }
  return String(rows[0].cells[0].value || "").replace(/\s+/g, "");
}

function appendUniqueNotes(target, incoming) {
  const seen = {};
  for (let i = 0; i < target.length; i += 1) {
    seen[String(target[i].sourceStart) + ":" + String(target[i].sourceEnd)] = true;
  }
  for (let i = 0; i < incoming.length; i += 1) {
    const key = String(incoming[i].sourceStart) + ":" + String(incoming[i].sourceEnd);
    if (seen[key]) {
      continue;
    }
    target.push(incoming[i]);
    seen[key] = true;
  }
}

function buildSectionRecord(tag, fullXml, absStart, absEnd, sectionKeyBase) {
  const title = extractTitleFromSection(fullXml);
  const sectionType = classifySectionType(tag, title);
  const sheetName = decideSectionSheetName(sectionType, title);
  const blocks = parseOrderedBlocks(fullXml, absStart);

  return {
    sectionKey: sectionKeyBase,
    sectionType: sectionType,
    sheetName: sheetName,
    title: title || sheetName,
    sourceStart: absStart,
    sourceEnd: absEnd,
    sourceXml: fullXml,
    blocks: blocks
  };
}

function splitSection2Notes(fullXml, absStart, sectionKeyBase) {
  const notes = [];
  const bodyBounds = getInnerBounds(fullXml, "SECTION-2");
  const blocks = parseOrderedBlocks(fullXml, absStart);
  const bookmarkBoundaries = collectBookmarkBoundaries(blocks);
  const boundaries = bookmarkBoundaries.length > 0 ? bookmarkBoundaries : collectHeadingBoundaries(blocks);

  if (boundaries.length === 0) {
    notes.push({
      sectionKey: sectionKeyBase + "_NOTE_0",
      sectionType: "SECTION2_NOTE",
      noteNumber: "0",
      sheetName: "주석_0",
      title: "주석",
      sourceStart: absStart,
      sourceEnd: absStart + fullXml.length,
      sourceXml: fullXml,
      blocks: blocks
    });
    return notes;
  }

  for (let i = 0; i < boundaries.length; i += 1) {
    const boundary = boundaries[i];
    const nextBoundary = i + 1 < boundaries.length ? boundaries[i + 1] : null;
    const startBlockIndex = i === 0 ? 0 : boundary.blockIndex;
    const endBlockIndex = nextBoundary ? nextBoundary.blockIndex : blocks.length;
    const noteBlocks = coalesceLeadingNoteTitleBlocks(blocks.slice(startBlockIndex, endBlockIndex));
    if (noteBlocks.length === 0) {
      continue;
    }
    const noteStart = noteBlocks[0].relStart;
    const noteEnd = nextBoundary ? nextBoundary.relStart : bodyBounds.innerEnd;
    const noteNo = boundary.noteNo;

    notes.push({
      sectionKey: sectionKeyBase + "_NOTE_" + noteNo,
      sectionType: "SECTION2_NOTE",
      noteNumber: noteNo,
      sheetName: "주석_" + noteNo,
      title: boundary.text || ("주석 " + noteNo),
      sourceStart: absStart + noteStart,
      sourceEnd: absStart + noteEnd,
      sourceXml: fullXml.slice(noteStart, noteEnd),
      blocks: noteBlocks
    });
  }

  return notes;
}

function collectBookmarkBoundaries(blocks) {
  const boundaries = [];
  for (let i = 0; i < blocks.length; i += 1) {
    const block = blocks[i];
    if (block.kind !== "서술문") {
      continue;
    }
    if (!/USERMARK\s*=\s*["']B[^"']*["']/i.test(block.rawXml || "")) {
      continue;
    }
    const bookmarkTitle = extractBookmarkTitle(block, blocks, i);
    const noteNo = extractNoteNumber(bookmarkTitle || block.text, boundaries.length + 1);
    boundaries.push({
      blockIndex: i,
      relStart: block.relStart,
      text: bookmarkTitle || normalizeText(block.text || ""),
      noteNo: String(noteNo)
    });
  }
  return boundaries;
}

function collectHeadingBoundaries(blocks) {
  const boundaries = [];
  for (let i = 0; i < blocks.length; i += 1) {
    if (blocks[i].kind !== "서술문") {
      continue;
    }
    const heading = parseTopLevelNoteHeading(blocks[i].text);
    if (!heading) {
      continue;
    }
    boundaries.push({
      blockIndex: i,
      relStart: blocks[i].relStart,
      text: heading.title,
      noteNo: heading.noteNo
    });
  }
  return boundaries;
}

function extractBookmarkTitle(block, blocks, blockIndex) {
  const spanIdMatch = /<SPAN\b[^>]*USERMARK\s*=\s*["']B[^"']*["'][^>]*\bID\s*=\s*["']([^"']+)["']/i.exec(block.rawXml || "");
  if (spanIdMatch && spanIdMatch[1]) {
    return normalizeText(stripXml(spanIdMatch[1]));
  }

  const currentText = normalizeText(block.text || "");
  if (/^(?:주석\s*)?\d+[\.)]?$/.test(currentText)) {
    const nextBlock = blocks[blockIndex + 1];
    if (nextBlock && nextBlock.kind === "서술문") {
      const nextText = normalizeText(nextBlock.text || "");
      if (nextText && !/^(?:주석\s*)?\d+[\.)]?$/.test(nextText)) {
        return (currentText + " " + nextText).trim();
      }
    }
  }

  return currentText;
}

function coalesceLeadingNoteTitleBlocks(noteBlocks) {
  if (noteBlocks.length < 2) {
    return noteBlocks;
  }
  const first = noteBlocks[0];
  const second = noteBlocks[1];
  if (!first || !second || first.kind !== "서술문" || second.kind !== "서술문") {
    return noteBlocks;
  }

  const firstText = normalizeText(first.text || "");
  const secondText = normalizeText(second.text || "");
  if (!/^(?:주석\s*)?\d+[\.)]?$/.test(firstText)) {
    return noteBlocks;
  }
  if (!secondText || /^(?:주석\s*)?\d+[\.)]?$/.test(secondText)) {
    return noteBlocks;
  }

  const merged = Object.assign({}, first);
  merged.text = (firstText + " " + secondText).trim();
  merged.rawXml = String(first.rawXml || "") + String(second.rawXml || "");
  merged.relEnd = second.relEnd;
  merged.absEnd = second.absEnd;

  const output = [merged];
  for (let i = 2; i < noteBlocks.length; i += 1) {
    output.push(noteBlocks[i]);
  }
  return output;
}

function parseTopLevelNoteHeading(text) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  const match = normalized.match(/^(?:주석\s*)?(\d+)[\.)]\s+(.+)$/);
  if (!match) {
    return null;
  }
  return {
    noteNo: match[1],
    title: normalized
  };
}

function parseOrderedBlocks(sectionXml, sectionAbsStart) {
  const bounds = getInnerBounds(sectionXml, getTopTagName(sectionXml));
  const inner = sectionXml.slice(bounds.innerStart, bounds.innerEnd);
  const tokenRegex = /<TABLE\b[\s\S]*?<\/TABLE>|<P\b[^>]*>[\s\S]*?<\/P>|<TITLE\b[^>]*>[\s\S]*?<\/TITLE>|<COVER-TITLE\b[^>]*>[\s\S]*?<\/COVER-TITLE>/gi;
  const blocks = [];

  let token = tokenRegex.exec(inner);
  let serial = 1;
  let tableSerial = 1;
  let textSerial = 1;

  while (token !== null) {
    const rawXml = token[0];
    const relStart = bounds.innerStart + token.index;
    const relEnd = bounds.innerStart + tokenRegex.lastIndex;
    const absStart = sectionAbsStart + relStart;
    const absEnd = sectionAbsStart + relEnd;
    const upper = rawXml.substring(0, 16).toUpperCase();

    if (upper.indexOf("<TABLE") === 0) {
      const tableInner = extractTableInner(rawXml);
      const rows = parseTableRows(tableInner);
      if (rows.length > 0) {
        blocks.push({
          marker: "[" + serial + "_표_" + tableSerial + "]",
          kind: "표",
          rawXml: rawXml,
          rows: rows,
          tableAttrs: parseXmlAttrs(rawXml.match(/<TABLE\s*([^>]*)>/i)?.[1] || ""),
          relStart: relStart,
          relEnd: relEnd,
          absStart: absStart,
          absEnd: absEnd
        });
        serial += 1;
        tableSerial += 1;
      }
    } else if (upper.indexOf("<TITLE") === 0 || upper.indexOf("<COVER-TITLE") === 0) {
      const text = normalizeText(stripXml(rawXml));
      if (text) {
        const titleKind = upper.indexOf("<COVER-TITLE") === 0 ? "표지제목" : "제목";
        blocks.push({
          marker: "[" + serial + "_" + titleKind + "_" + textSerial + "]",
          kind: titleKind,
          rawXml: rawXml,
          text: text,
          relStart: relStart,
          relEnd: relEnd,
          absStart: absStart,
          absEnd: absEnd
        });
        serial += 1;
        textSerial += 1;
      }
    } else {
      const paragraphText = normalizeText(stripXml(rawXml));
      if (paragraphText) {
        blocks.push({
          marker: "[" + serial + "_서술문_" + textSerial + "]",
          kind: "서술문",
          rawXml: rawXml,
          text: paragraphText,
          relStart: relStart,
          relEnd: relEnd,
          absStart: absStart,
          absEnd: absEnd
        });
        serial += 1;
        textSerial += 1;
      }
    }

    token = tokenRegex.exec(inner);
  }

  return blocks;
}

function classifySectionType(tag, title) {
  if (tag === "COVER") {
    return "COVER";
  }
  if (tag === "TOC") {
    return "TOC";
  }
  const compact = (title || "").replace(/\s+/g, "");
  if (compact.indexOf("독립된감사인의감사보고서") !== -1 || compact.indexOf("감사보고서") !== -1) {
    return "SECTION1_AUDIT";
  }
  if (compact.indexOf("(첨부)연결재무제표") !== -1 || compact.indexOf("연결재무제표") !== -1) {
    return "SECTION1_ATTACHED";
  }
  return "OTHER";
}

function decideSectionSheetName(sectionType, title) {
  if (sectionType === "COVER") {
    return "표지";
  }
  if (sectionType === "TOC") {
    return "목차";
  }
  if (sectionType === "SECTION1_AUDIT") {
    return "감사보고서";
  }
  if (sectionType === "SECTION1_ATTACHED") {
    return "(첨부)연결재무제표";
  }
  return (title || "기타").substring(0, 30);
}

function extractTitleFromSection(sectionXml) {
  const titleMatch = /<TITLE\b[^>]*>([\s\S]*?)<\/TITLE>/i.exec(sectionXml);
  if (!titleMatch) {
    const coverTitleMatch = /<COVER-TITLE\b[^>]*>([\s\S]*?)<\/COVER-TITLE>/i.exec(sectionXml);
    return coverTitleMatch ? normalizeText(stripXml(coverTitleMatch[1])) : "";
  }
  return normalizeText(stripXml(titleMatch[1]));
}

function getTopTagName(xml) {
  const m = /^\s*<([A-Z0-9-]+)/i.exec(xml);
  return m ? m[1] : "SECTION-1";
}

function getInnerBounds(sectionXml, tagName) {
  const openRegex = new RegExp("<" + tagName + "\\b[^>]*>", "i");
  const openMatch = openRegex.exec(sectionXml);
  const openEnd = openMatch ? openMatch.index + openMatch[0].length : 0;
  const closeTag = "</" + tagName + ">";
  const closeStart = sectionXml.toUpperCase().lastIndexOf(closeTag.toUpperCase());
  return {
    innerStart: openEnd,
    innerEnd: closeStart >= 0 ? closeStart : sectionXml.length
  };
}

function extractNoteNumber(title, fallbackNo) {
  const normalized = String(title || "").replace(/\s+/g, " ").trim();
  const m = normalized.match(/^(?:주석\s*)?(\d+)[\.)]?/);
  return m ? m[1] : String(fallbackNo);
}

function dedupeNotes(notes) {
  const totals = {};
  const seen = {};
  const output = [];

  for (let i = 0; i < notes.length; i += 1) {
    const noteNo = String(notes[i].noteNumber);
    totals[noteNo] = (totals[noteNo] || 0) + 1;
  }

  for (let i = 0; i < notes.length; i += 1) {
    const noteNo = String(notes[i].noteNumber);
    const count = seen[noteNo] || 0;
    seen[noteNo] = count + 1;

    if (totals[noteNo] > 1) {
      const suffix = count + 1;
      notes[i].sheetName = "주석_" + noteNo + "_" + suffix;
      notes[i].sectionKey = notes[i].sectionKey + "_" + suffix;
    } else {
      notes[i].sheetName = "주석_" + noteNo;
    }

    output.push(notes[i]);
  }
  return output;
}

function extractTableInner(tableXml) {
  const start = tableXml.indexOf(">");
  const end = tableXml.lastIndexOf("</TABLE>");
  if (start < 0 || end < 0 || end <= start) {
    return "";
  }
  return tableXml.slice(start + 1, end);
}

function parseTableRows(tableXml) {
  const trRegex = /<TR\s*([^>]*)>([\s\S]*?)<\/TR>/gi;
  const rows = [];
  let tr = trRegex.exec(tableXml);
  while (tr !== null) {
    rows.push({
      attrs: parseXmlAttrs(tr[1] || ""),
      cells: parseCells(tr[2]),
      rawXml: tr[0]
    });
    tr = trRegex.exec(tableXml);
  }
  return rows;
}

function parseCells(trXml) {
  const cellRegex = /<(TD|TH|TU)\s*([^>]*)>([\s\S]*?)<\/\1>/gi;
  const cells = [];
  let cell = cellRegex.exec(trXml);
  while (cell !== null) {
    cells.push({
      tag: cell[1],
      attrs: parseXmlAttrs(cell[2] || ""),
      value: normalizeText(stripXml(cell[3])),
      rawXml: cell[0]
    });
    cell = cellRegex.exec(trXml);
  }
  return cells;
}

function parseXmlAttrs(attrSource) {
  const attrs = {};
  const attrRegex = /([A-Za-z_:][-A-Za-z0-9_:.]*)="([^"]*)"/g;
  let attr = attrRegex.exec(attrSource);
  while (attr !== null) {
    attrs[attr[1]] = attr[2];
    attr = attrRegex.exec(attrSource);
  }
  return attrs;
}

function stripXml(fragment) {
  return String(fragment || "")
    .replace(/<BR\s*\/?\s*>/gi, "\n")
    .replace(/&amp;cr;/gi, "\n")
    .replace(/&cr;/gi, "\n")
    .replace(/&nbsp;/gi, " ")
    .replace(/<[^>]*>/g, " ")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function normalizeText(value) {
  return String(value || "")
    .split("\n")
    .map(function (line) {
      return line.replace(/\s+/g, " ").trim();
    })
    .filter(function (line) {
      return line !== "";
    })
    .join("\n");
}
