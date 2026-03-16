function extractDsd(fileId, options) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const zipBlob = Utilities.newBlob(blob.getBytes(), "application/zip", file.getName());
  const zippedEntries = Utilities.unzip(zipBlob);
  const opts = options || {};

  const result = {
    contentsXml: "",
    metaXml: "",
    resources: [],
    unzippedFolderId: "",
    unzippedFolderUrl: ""
  };

  const entryNames = [];
  const contentsCandidates = [];
  const xmlCandidates = [];

  for (let i = 0; i < zippedEntries.length; i += 1) {
    const entry = zippedEntries[i];
    const name = entry.getName();
    entryNames.push(name);

    const normalized = normalizeZipEntryName(name);
    if (normalized === "contents.xml") {
      contentsCandidates.unshift(entry);
    } else if (/^contents.*\.xml$/.test(normalized)) {
      contentsCandidates.push(entry);
    } else if (/\.xml$/.test(normalized)) {
      xmlCandidates.push(entry);
    } else if (normalized === "meta.xml") {
      result.metaXml = readXmlBlobFlexible(entry);
    } else {
      result.resources.push(entry);
    }
  }

  if (!result.contentsXml && contentsCandidates.length > 0) {
    result.contentsXml = chooseBestContentsXml(contentsCandidates);
  }

  if (!result.contentsXml && xmlCandidates.length > 0) {
    result.contentsXml = chooseBestContentsXml(xmlCandidates);
  }

  if (opts.sourceFolderId) {
    try {
      const extractedInfo = persistUnzippedEntries(opts.sourceFolderId, file.getName(), zippedEntries);
      result.unzippedFolderId = extractedInfo.folderId;
      result.unzippedFolderUrl = extractedInfo.folderUrl;
    } catch (error) {
      Logger.log("UNZIP_PERSIST_WARN: " + String(error));
    }
  }

  if (!result.contentsXml) {
    const fileInfo = file.getName() + " (" + file.getMimeType() + ", size: " + blob.getBytes().length + ")";
    throw new Error("Invalid DSD: contents.xml is missing. File: " + fileInfo + " zipEntries=" + entryNames.slice(0, 20).join(", "));
  }

  if (!result.metaXml) {
    result.metaXml = defaultMetaXml();
  }

  return result;
}

function normalizeZipEntryName(name) {
  return String(name || "")
    .replace(/\\/g, "/")
    .split("/")
    .pop()
    .trim()
    .toLowerCase();
}

function chooseBestContentsXml(candidates) {
  let fallback = "";
  for (let i = 0; i < candidates.length; i += 1) {
    const xml = readXmlBlobFlexible(candidates[i]);
    if (!fallback) {
      fallback = xml;
    }
    if (looksLikeContentsXml(xml)) {
      return xml;
    }
  }
  return fallback;
}

function readXmlBlobFlexible(blob) {
  const charsets = ["utf-8", "euc-kr", "x-windows-949"];
  for (let i = 0; i < charsets.length; i += 1) {
    try {
      const text = blob.getDataAsString(charsets[i]);
      if (text && text.trim()) {
        return text;
      }
    } catch (error) {
    }
  }
  try {
    return blob.getDataAsString();
  } catch (error) {
    return "";
  }
}

function looksLikeContentsXml(xml) {
  const text = String(xml || "").trim();
  if (!text) {
    return false;
  }
  if (text.indexOf("<DOCUMENT") === -1 && text.indexOf("<BODY") === -1 && text.indexOf("<SECTION") === -1) {
    return false;
  }
  try {
    XmlService.parse(text);
    return true;
  } catch (error) {
    return false;
  }
}

function persistUnzippedEntries(sourceFolderId, sourceFileName, entries) {
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const unzippedRoot = getOrCreateDriveChildFolder(sourceFolder, "unzipped");
  const baseName = String(sourceFileName || "upload").replace(/\.dsd$/i, "");
  const stampedName = sanitizeDriveName(baseName) + "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const runFolder = unzippedRoot.createFolder(stampedName);

  for (let i = 0; i < entries.length; i += 1) {
    const entry = entries[i];
    const path = normalizeZipEntryPath(entry.getName());
    if (!path || path.charAt(path.length - 1) === "/") {
      continue;
    }
    writeEntryBlobWithPath(runFolder, path, entry);
  }

  return {
    folderId: runFolder.getId(),
    folderUrl: runFolder.getUrl()
  };
}

function writeEntryBlobWithPath(rootFolder, entryPath, blob) {
  const safePath = String(entryPath || "").replace(/^\/+/, "");
  const parts = safePath.split("/");
  if (parts.length === 0) {
    return;
  }

  let targetFolder = rootFolder;
  for (let i = 0; i < parts.length - 1; i += 1) {
    const folderName = sanitizeDriveName(parts[i]);
    if (!folderName) {
      continue;
    }
    targetFolder = getOrCreateDriveChildFolder(targetFolder, folderName);
  }

  const fileName = sanitizeDriveName(parts[parts.length - 1]) || ("entry_" + new Date().getTime());
  targetFolder.createFile(blob.copyBlob().setName(fileName));
}

function getOrCreateDriveChildFolder(parentFolder, childName) {
  const iter = parentFolder.getFoldersByName(childName);
  if (iter.hasNext()) {
    return iter.next();
  }
  return parentFolder.createFolder(childName);
}

function normalizeZipEntryPath(name) {
  return String(name || "").replace(/\\/g, "/").trim();
}

function sanitizeDriveName(name) {
  return String(name || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .trim();
}

function createDsdFromXml(contentsXml, metaXml, resources, outputName, targetFolderId) {
  const blobs = [];
  blobs.push(Utilities.newBlob(contentsXml, "application/xml", "contents.xml"));
  blobs.push(Utilities.newBlob(metaXml || defaultMetaXml(), "application/xml", "meta.xml"));

  const safeResources = resources || [];
  for (let i = 0; i < safeResources.length; i += 1) {
    blobs.push(safeResources[i]);
  }

  const zipBlob = Utilities.zip(blobs, outputName + ".dsd");
  if (targetFolderId) {
    return DriveApp.getFolderById(targetFolderId).createFile(zipBlob);
  }

  return DriveApp.getRootFolder().createFile(zipBlob);
}

function defaultMetaXml() {
  return '<?xml version="1.0" encoding="utf-8"?><META><DOCUMENT-STATUS><OPENMARKUP>Y</OPENMARKUP><OPENREPORT>Y</OPENREPORT></DOCUMENT-STATUS></META>';
}
