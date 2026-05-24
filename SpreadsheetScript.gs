// ============================================================
//  レシート分類アプリ（無料版）— Google Apps Script
//  設置先: スプレッドシート → 拡張機能 → Apps Script
// ============================================================

const SHEET_NAME = "食費記録";
const HEADER = ["日付", "店名", "品目", "カテゴリ", "MFカテゴリ", "金額（円）", "メモ"];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "append") return respond(appendRows(payload.rows));
    if (payload.action === "getAll") return respond(getAllRows());
    return respond({ error: "unknown action" });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doGet() {
  return respond({ status: "ok", message: "レシート分類 GAS 稼働中" });
}

function appendRows(rows) {
  const sheet = getOrCreateSheet();
  rows.forEach(row => {
    sheet.appendRow([
      row.date || today(), row.store || "", row.name || "",
      row.cat || "", row.mfCat || "",
      Number(row.price) || 0, row.memo || "レシートスキャン"
    ]);
  });
  // 金額列を数値フォーマット
  const last = sheet.getLastRow();
  if (rows.length > 0) {
    sheet.getRange(last - rows.length + 1, 6, rows.length, 1).setNumberFormat("#,##0");
  }
  return { appended: rows.length, lastRow: sheet.getLastRow() };
}

function getAllRows() {
  const sheet = getOrCreateSheet();
  const last = sheet.getLastRow();
  if (last <= 1) return { rows: [] };
  const data = sheet.getRange(2, 1, last - 1, HEADER.length).getValues();
  return {
    rows: data.map(r => ({
      date:  r[0] ? Utilities.formatDate(new Date(r[0]), "Asia/Tokyo", "yyyy/MM/dd") : "",
      store: r[1], name: r[2], cat: r[3], mfCat: r[4], price: r[5], memo: r[6]
    }))
  };
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    const h = sheet.getRange(1, 1, 1, HEADER.length);
    h.setBackground("#1D9E75");
    h.setFontColor("#ffffff");
    h.setFontWeight("bold");
    sheet.setFrozenRows(1);
    [100, 140, 180, 120, 150, 100, 130].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

function today() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
}

function respond(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
