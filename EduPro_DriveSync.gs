// ╔══════════════════════════════════════════════════════════════════╗
// ║          EDUPRO — GOOGLE APPS SCRIPT BACKEND                   ║
// ║    Đồng bộ dữ liệu tất cả module lên Google Drive              ║
// ║                                                                 ║
// ║  HƯỚNG DẪN CÀI ĐẶT:                                            ║
// ║  1. Mở script.google.com → New project                         ║
// ║  2. Dán toàn bộ nội dung file này vào                          ║
// ║  3. Nhấn Deploy → New deployment → Web app                     ║
// ║  4. Execute as: Me | Who has access: Anyone                    ║
// ║  5. Sao chép URL deployment → dán vào EDUPRO_SCRIPT_URL        ║
// ║     trong file edupro-sync.js (hoặc mỗi module HTML)           ║
// ╚══════════════════════════════════════════════════════════════════╝

// ─── CẤU HÌNH ────────────────────────────────────────────────────
const FOLDER_NAME  = 'EduPro_Data';       // Tên thư mục trên Drive
const SHEET_NAME   = 'EduPro_Database';   // Tên Google Sheet chính

// ─── ENTRY POINT: xử lý mọi request từ các module HTML ───────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { action, module, key, data } = payload;

    let result;
    switch (action) {
      case 'save':   result = saveData(module, key, data); break;
      case 'load':   result = loadData(module, key);       break;
      case 'loadAll':result = loadAll(module);             break;
      case 'delete': result = deleteData(module, key);     break;
      case 'ping':   result = { ok: true, time: new Date().toISOString() }; break;
      default:       result = { error: 'Unknown action: ' + action };
    }

    return jsonResponse({ ok: true, ...result });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doGet(e) {
  // Health check endpoint
  return jsonResponse({ ok: true, service: 'EduPro DriveSync', time: new Date().toISOString() });
}

// ─── LƯU DỮ LIỆU ─────────────────────────────────────────────────
function saveData(module, key, data) {
  const sheet = getOrCreateSheet(module);
  const now   = new Date().toISOString();
  const json  = JSON.stringify(data);

  // Tìm hàng có key tương ứng
  const col1 = sheet.getRange('A:A').getValues().flat();
  const rowIdx = col1.indexOf(key);

  if (rowIdx > 0) {
    // Cập nhật hàng đã có (rowIdx+1 vì mảng 0-based, sheet 1-based)
    sheet.getRange(rowIdx + 1, 2).setValue(json);
    sheet.getRange(rowIdx + 1, 3).setValue(now);
  } else {
    // Thêm hàng mới
    sheet.appendRow([key, json, now]);
  }

  // Ghi log hoạt động
  logActivity(module, key, 'SAVE');
  return { saved: true, key, module, timestamp: now };
}

// ─── TẢI DỮ LIỆU ─────────────────────────────────────────────────
function loadData(module, key) {
  const sheet = getOrCreateSheet(module);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {  // bỏ qua header row
    if (values[i][0] === key) {
      try {
        return { found: true, key, data: JSON.parse(values[i][1]), updatedAt: values[i][2] };
      } catch {
        return { found: true, key, data: values[i][1], updatedAt: values[i][2] };
      }
    }
  }
  return { found: false, key };
}

// ─── TẢI TOÀN BỘ DỮ LIỆU CỦA 1 MODULE ──────────────────────────
function loadAll(module) {
  const sheet  = getOrCreateSheet(module);
  const values = sheet.getDataRange().getValues();
  const result = {};

  for (let i = 1; i < values.length; i++) {
    const key = values[i][0];
    if (!key) continue;
    try { result[key] = JSON.parse(values[i][1]); }
    catch { result[key] = values[i][1]; }
  }
  return { found: true, module, data: result };
}

// ─── XOÁ DỮ LIỆU ─────────────────────────────────────────────────
function deleteData(module, key) {
  const sheet  = getOrCreateSheet(module);
  const col1   = sheet.getRange('A:A').getValues().flat();
  const rowIdx = col1.indexOf(key);

  if (rowIdx > 0) {
    sheet.deleteRow(rowIdx + 1);
    return { deleted: true, key };
  }
  return { deleted: false, key, reason: 'Key not found' };
}

// ─── TIỆN ÍCH ────────────────────────────────────────────────────
function getOrCreateSheet(module) {
  const ss        = getOrCreateSpreadsheet();
  const sheetName = 'mod_' + module;
  let sheet       = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Header row
    sheet.getRange('A1:C1').setValues([['key', 'value_json', 'updated_at']]);
    sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#0d9488').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 500);
    sheet.setColumnWidth(3, 200);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateSpreadsheet() {
  const folder = getOrCreateFolder();
  const files  = folder.getFilesByName(SHEET_NAME);

  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  // Tạo mới spreadsheet trong thư mục EduPro
  const ss = SpreadsheetApp.create(SHEET_NAME);
  // Di chuyển vào thư mục EduPro_Data
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // Tạo sheet Overview
  const overview = ss.getSheets()[0];
  overview.setName('📊 Overview');
  overview.getRange('A1').setValue('EduPro DriveSync — Database').setFontSize(16).setFontWeight('bold');
  overview.getRange('A2').setValue('Tự động tạo bởi EduPro Apps Script');
  overview.getRange('A3').setValue('Cập nhật: ' + new Date().toLocaleString('vi-VN'));

  return ss;
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();

  const folder = DriveApp.createFolder(FOLDER_NAME);
  return folder;
}

function logActivity(module, key, action) {
  try {
    const ss    = getOrCreateSpreadsheet();
    let logSheet = ss.getSheetByName('📋 Activity Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('📋 Activity Log');
      logSheet.getRange('A1:D1').setValues([['Thời gian', 'Module', 'Key', 'Action']]);
      logSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([new Date().toLocaleString('vi-VN'), module, key, action]);
    // Giữ tối đa 500 dòng log
    const lastRow = logSheet.getLastRow();
    if (lastRow > 501) logSheet.deleteRows(2, lastRow - 501);
  } catch (e) { /* log không critical, bỏ qua lỗi */ }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
