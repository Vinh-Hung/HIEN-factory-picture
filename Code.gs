// ── CẤU HÌNH ─────────────────────────────────────────────────────────────────
const SHEET_NAME = "PhotoLog";
const RETENTION_DAYS = 30;

// ── Web App entry point ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === "upload")       return jsonResponse(handleUpload(data));
    if (action === "getFactories") return jsonResponse(handleGetFactories());
    if (action === "addFactory")   return jsonResponse(handleAddFactory(data.name));
    if (action === "removeFactory")return jsonResponse(handleRemoveFactory(data.name));
    if (action === "search")       return jsonResponse(handleSearch(data));
    if (action === "purge")        return jsonResponse(handlePurge());

    return jsonResponse({ ok: false, error: "Unknown action" });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  // health check
  return jsonResponse({ ok: true, message: "Factory Photo API running" });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet helpers ─────────────────────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id","factory","date","engineer","uploadedAt","photoName","photoSize","photoData"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getFactorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Factories");
  if (!sheet) {
    sheet = ss.insertSheet("Factories");
    sheet.appendRow(["name"]);
    sheet.appendRow(["Nhà máy A"]);
    sheet.appendRow(["Nhà máy B"]);
  }
  return sheet;
}

// ── Handlers ──────────────────────────────────────────────────────────────────
function handleGetFactories() {
  const sheet = getFactorySheet();
  const data = sheet.getDataRange().getValues();
  const factories = data.slice(1).map(r => r[0]).filter(Boolean);
  return { ok: true, factories };
}

function handleAddFactory(name) {
  if (!name) return { ok: false, error: "Thiếu tên nhà máy" };
  const sheet = getFactorySheet();
  const existing = sheet.getDataRange().getValues().slice(1).map(r => r[0]);
  if (existing.includes(name)) return { ok: false, error: "Nhà máy đã tồn tại" };
  sheet.appendRow([name]);
  return { ok: true };
}

function handleRemoveFactory(name) {
  const sheet = getFactorySheet();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === name) { sheet.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

function handleUpload(data) {
  // purge old records first
  handlePurge();

  const sheet = getSheet();
  const { factory, date, engineer, photos, sessionId } = data;
  const uploadedAt = new Date().toISOString();

  photos.forEach(p => {
    sheet.appendRow([
      sessionId,
      factory,
      date,
      engineer,
      uploadedAt,
      p.name,
      p.size,
      p.data   // base64
    ]);
  });

  return { ok: true, count: photos.length };
}

function handleSearch(data) {
  const { factory, date } = data;
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, sessions: [] };

  const headers = rows[0]; // id,factory,date,engineer,uploadedAt,photoName,photoSize,photoData
  const filtered = rows.slice(1).filter(r => {
    if (factory && r[1] !== factory) return false;
    if (date && r[2] !== date) return false;
    return true;
  });

  // group by sessionId
  const sessionMap = {};
  filtered.forEach(r => {
    const sid = r[0];
    if (!sessionMap[sid]) {
      sessionMap[sid] = { id: sid, factory: r[1], date: r[2], engineer: r[3], uploadedAt: r[4], photos: [] };
    }
    sessionMap[sid].photos.push({ name: r[5], size: r[6], data: r[7] });
  });

  const sessions = Object.values(sessionMap).sort((a,b) => b.date.localeCompare(a.date));
  return { ok: true, sessions };
}

function handlePurge() {
  const sheet = getSheet();
  const cutoff = new Date(Date.now() - RETENTION_DAYS * 86400000);
  const rows = sheet.getDataRange().getValues();
  let deleted = 0;

  for (let i = rows.length - 1; i >= 1; i--) {
    const uploadedAt = new Date(rows[i][4]);
    if (uploadedAt < cutoff) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return { ok: true, deleted };
}

function getStats() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  return { totalRows: rows.length - 1 };
}
