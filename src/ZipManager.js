function extractDsd(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const zipBlob = blob.copyBlob().setContentType("application/zip");
  const zippedEntries = Utilities.unzip(zipBlob);

  const result = {
    contentsXml: "",
    metaXml: "",
    resources: []
  };

  const entryNames = [];

  for (let i = 0; i < zippedEntries.length; i += 1) {
    const entry = zippedEntries[i];
    const name = entry.getName();
    entryNames.push(name);

    const normalized = normalizeZipEntryName(name);
    if (normalized === "contents.xml" || normalized.indexOf("contents") === 0 && normalized.indexOf(".xml") > -1) {
      result.contentsXml = entry.getDataAsString("utf-8");
    } else if (normalized === "meta.xml") {
      result.metaXml = entry.getDataAsString("utf-8");
    } else {
      result.resources.push(entry);
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
