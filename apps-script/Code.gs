/****************************************************
 * 南帝精密｜內部報修系統 GAS 安全版 Code.gs
 *
 * 重點：
 * 1. 員工頁仍由 GitHub Pages 提供。
 * 2. 員工以「姓名＋4 位查詢碼」查自己的案件。
 * 3. 員工 API 不回傳照片檔案 ID 或照片連結。
 * 4. 管理頁改由 GAS 的 Admin.html 提供。
 * 5. 管理密碼只經由 google.script.run 傳到 GAS，不放在網址參數中。
 * 6. 新上傳照片保留為私人 Drive 檔案，只在管理後台按需讀取。
 * 7. 每次管理者更新皆寫入「處理紀錄」。
 ****************************************************/

const SHEET_NAME = '報修單總表';
const LOG_SHEET_NAME = '處理紀錄';
const PHOTO_FOLDER_NAME = '南帝報修圖片';
const ADMIN_PASSWORD = '請改成正式管理密碼';
const ADMIN_NOTIFY_EMAIL = 'flt.roger@gmail.com';
const SPREADSHEET_ID = '';
const TIME_ZONE = 'Asia/Taipei';
const ADMIN_TOKEN_SECONDS = 60 * 60 * 6;

const HEADERS = [
  '報修編號','建立時間','姓名','查詢碼','部門','報修人Email','地點',
  '問題類型','緊急程度','問題說明',
  '照片檔案ID1','照片檔案ID2','照片檔案ID3',
  '照片連結','照片連結2','照片連結3',
  '狀態','負責人','預定完成日','備註','完成時間','最後更新時間'
];

const LOG_HEADERS = [
  '更新時間','報修編號','原狀態','新狀態','負責人','預定完成日','備註'
];

/******************** 初始化 ********************/
function setup() {
  const sheet = getSheet_();
  ensureHeaders_(sheet, HEADERS);
  reorderColumns_(sheet, HEADERS);
  styleSheet_(sheet);

  const logSheet = getLogSheet_();
  ensureHeaders_(logSheet, LOG_HEADERS);
  styleSheet_(logSheet);

  const map = getHeaderMap_(sheet);
  if (map['查詢碼'] !== undefined) {
    sheet.getRange(2, map['查詢碼'] + 1, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
  }

  getPhotoFolder_();
  Logger.log('安全版報修系統初始化完成。');
}

/******************** Web App 入口 ********************/
function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};

  if (String(p.page || '') === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('南帝精密報修管理後台')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    const action = String(p.action || '').trim();
    if (action === 'query') return queryRepair_(e);

    if (['listRecent', 'listRepairs', 'stats'].includes(action)) {
      return outputJson_({
        success: false,
        message: '此公開查詢功能已停用，管理者請由管理後台操作。'
      }, e);
    }

    return outputJson_({success: true, message: '南帝精密報修系統 GAS 正常運作'}, e);
  } catch (error) {
    return outputJson_({success: false, message: '系統讀取失敗：' + error.message}, e);
  }
}

function doPost(e) {
  try {
    const action = String(e.parameter.action || '').trim();
    if (action === 'create') return createRepair_(e);

    return outputJson_({
      success: false,
      message: '此公開送出功能不支援管理更新，請由管理後台操作。'
    });
  } catch (error) {
    return outputJson_({success: false, message: '系統處理失敗：' + error.message});
  }
}

/******************** 員工端：新增與查詢 ********************/
function createRepair_(e) {
  const queryCode = String(e.parameter.queryCode || '').trim();
  if (!/^\d{4}$/.test(queryCode)) {
    return outputJson_({success: false, message: '查詢碼必須是 4 位數字'});
  }

  const sheet = getSheet_();
  ensureHeaders_(sheet, HEADERS);
  const caseId = String(e.parameter.caseId || generateCaseId_()).trim();
  const now = new Date();

  const photoIds = [1, 2, 3].map(function(index) {
    const suffix = index === 1 ? '' : String(index);
    return savePrivatePhoto_(
      caseId + '_0' + index,
      e.parameter['photoData' + suffix] || '',
      e.parameter['photoName' + suffix] || '',
      e.parameter['photoType' + suffix] || ''
    );
  });

  const item = {
    '報修編號': caseId,
    '建立時間': now,
    '姓名': String(e.parameter.name || '').trim(),
    '查詢碼': queryCode,
    '部門': String(e.parameter.department || '').trim(),
    '報修人Email': String(e.parameter.requesterEmail || '').trim(),
    '地點': String(e.parameter.location || '').trim(),
    '問題類型': String(e.parameter.category || '').trim(),
    '緊急程度': String(e.parameter.urgency || '一般').trim(),
    '問題說明': String(e.parameter.description || '').trim(),
    '照片檔案ID1': photoIds[0],
    '照片檔案ID2': photoIds[1],
    '照片檔案ID3': photoIds[2],
    '照片連結': '', '照片連結2': '', '照片連結3': '',
    '狀態': '待處理', '負責人': '', '預定完成日': '', '備註': '',
    '完成時間': '', '最後更新時間': now
  };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    appendObjectRow_(sheet, item);
  } finally {
    lock.releaseLock();
  }

  try { sendNewRepairNotice_(item, photoIds.filter(Boolean).length); }
  catch (error) { Logger.log('新報修通知寄送失敗：' + error.message); }

  return outputJson_({success: true, message: '報修單已建立', caseId: caseId});
}

function queryRepair_(e) {
  const name = String(e.parameter.name || '').trim();
  const queryCode = String(e.parameter.queryCode || '').trim();
  if (!name) return outputJson_({success: false, message: '請輸入姓名'}, e);
  if (!/^\d{4}$/.test(queryCode)) return outputJson_({success: false, message: '請輸入 4 位數字查詢碼'}, e);

  const results = getDataObjects_()
    .filter(function(item) {
      return String(item['姓名'] || '').trim() === name && String(item['查詢碼'] || '').trim() === queryCode;
    })
    .map(toEmployeeSafeItem_)
    .reverse();

  return outputJson_({success: true, count: results.length, results: results}, e);
}

function toEmployeeSafeItem_(item) {
  return {
    '報修編號': item['報修編號'],
    '建立時間': item['建立時間'],
    '地點': item['地點'],
    '問題類型': item['問題類型'],
    '緊急程度': item['緊急程度'],
    '問題說明': item['問題說明'],
    '狀態': item['狀態'],
    '預定完成日': item['預定完成日'],
    '備註': item['備註'],
    '完成時間': item['完成時間']
  };
}

/******************** 管理後台：登入與工作階段 ********************/
function adminLogin(password) {
  if (String(password || '') !== ADMIN_PASSWORD || ADMIN_PASSWORD === '請改成正式管理密碼') {
    return {success: false, message: '管理密碼錯誤，或尚未設定正式管理密碼。'};
  }
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('ADMIN_TOKEN_' + token, 'OK', ADMIN_TOKEN_SECONDS);
  return {success: true, token: token, expiresMinutes: ADMIN_TOKEN_SECONDS / 60};
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove('ADMIN_TOKEN_' + token);
  return {success: true};
}

function requireAdmin_(token) {
  const ok = token && CacheService.getScriptCache().get('ADMIN_TOKEN_' + String(token));
  if (!ok) throw new Error('管理登入已逾時，請重新登入。');
}

/******************** 管理後台：清單、統計、更新、照片 ********************/
function adminListRepairs(token, filterStatus) {
  requireAdmin_(token);
  const status = String(filterStatus || '待處理').trim();
  const results = getDataObjects_().filter(function(item) {
    const value = String(item['狀態'] || '待處理').trim() || '待處理';
    return status === '全部' || value === status;
  }).map(function(item) {
    return {
      '報修編號': item['報修編號'], '建立時間': item['建立時間'], '姓名': item['姓名'],
      '部門': item['部門'], '報修人Email': item['報修人Email'], '地點': item['地點'],
      '問題類型': item['問題類型'], '緊急程度': item['緊急程度'], '問題說明': item['問題說明'],
      '狀態': item['狀態'], '負責人': item['負責人'], '預定完成日': item['預定完成日'],
      '備註': item['備註'], '完成時間': item['完成時間'],
      '照片數': countPhotos_(item)
    };
  }).reverse();
  return {success: true, count: results.length, results: results};
}

function adminGetStats(token) {
  requireAdmin_(token);
  const data = getDataObjects_();
  const today = startOfToday_();
  const currentMonth = Utilities.formatDate(today, TIME_ZONE, 'yyyy/MM');
  const stats = {'全部': data.length, '待處理': 0, '處理中': 0, '暫緩': 0, '已完成': 0, '逾期': 0, '本月新增': 0, '平均完成天數': 0};
  let days = 0, completed = 0;
  data.forEach(function(item) {
    const status = String(item['狀態'] || '待處理').trim() || '待處理';
    if (stats[status] !== undefined) stats[status]++;
    if (String(item['建立時間'] || '').indexOf(currentMonth) === 0) stats['本月新增']++;
    const due = parseDateOnly_(item['預定完成日']);
    if (due && status !== '已完成' && due.getTime() < today.getTime()) stats['逾期']++;
    const start = parseDateTime_(item['建立時間']);
    const finish = parseDateTime_(item['完成時間']);
    if (start && finish) { days += Math.max(0, (finish.getTime() - start.getTime()) / 86400000); completed++; }
  });
  stats['平均完成天數'] = completed ? Math.round(days / completed * 10) / 10 : 0;
  return {success: true, stats: stats};
}

function adminUpdateRepair(token, payload) {
  requireAdmin_(token);
  payload = payload || {};
  const caseId = String(payload.caseId || '').trim();
  const status = String(payload.status || '').trim();
  const owner = String(payload.owner || '').trim();
  const dueDate = String(payload.dueDate || '').trim();
  const note = String(payload.note || '').trim();
  if (!caseId) throw new Error('請先選擇案件。');
  if (!['待處理','處理中','暫緩','已完成'].includes(status)) throw new Error('狀態不正確。');
  if ((status === '處理中' || status === '已完成') && !owner) throw new Error('處理中或已完成案件，請填寫負責人。');
  if ((status === '暫緩' || status === '已完成') && !note) throw new Error('暫緩或已完成案件，請填寫備註說明。');

  const sheet = getSheet_();
  const map = getHeaderMap_(sheet);
  const data = sheet.getDataRange().getValues();
  let rowNo = -1, oldItem = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map['報修編號']] || '').trim() === caseId) {
      rowNo = i + 1; oldItem = rowArrayToObject_(data[0], data[i]); break;
    }
  }
  if (rowNo < 0) throw new Error('找不到此報修編號。');
  const now = new Date();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    setCellByHeader_(sheet, rowNo, map, '狀態', status);
    setCellByHeader_(sheet, rowNo, map, '負責人', owner);
    setCellByHeader_(sheet, rowNo, map, '預定完成日', dueDate);
    setCellByHeader_(sheet, rowNo, map, '備註', note);
    setCellByHeader_(sheet, rowNo, map, '最後更新時間', now);
    if (status === '已完成') setCellByHeader_(sheet, rowNo, map, '完成時間', oldItem['完成時間'] || now);
    if (status !== '已完成' && oldItem['狀態'] === '已完成') setCellByHeader_(sheet, rowNo, map, '完成時間', '');
    appendLog_(caseId, oldItem['狀態'], status, owner, dueDate, note);
  } finally { lock.releaseLock(); }

  const updated = getRepairByCaseId_(caseId);
  try { sendCompletionNotice_(oldItem, updated); }
  catch (error) { Logger.log('完成通知寄送失敗：' + error.message); }
  return {success: true, message: '案件已更新'};
}

function adminGetPhoto(token, caseId, photoNumber) {
  requireAdmin_(token);
  const index = Number(photoNumber);
  if (![1, 2, 3].includes(index)) throw new Error('照片編號不正確。');
  const item = getRepairByCaseId_(String(caseId || '').trim());
  if (!item) throw new Error('找不到此案件。');
  const fileId = resolvePhotoFileId_(item, index);
  if (!fileId) return {success: false, message: '此案件沒有照片 ' + index + '。'};
  const blob = DriveApp.getFileById(fileId).getBlob();
  const contentType = blob.getContentType() || 'image/jpeg';
  return {success: true, dataUrl: 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes())};
}

/******************** 照片 ********************/
function savePrivatePhoto_(prefix, dataUrl, fileName, mimeType) {
  if (!dataUrl) return '';
  const base64 = String(dataUrl).replace(/^data:.+;base64,/, '');
  const bytes = Utilities.base64Decode(base64);
  const ext = getExtension_(fileName, mimeType);
  const file = getPhotoFolder_().createFile(Utilities.newBlob(bytes, mimeType || 'image/jpeg', prefix + '_' + Date.now() + ext));
  file.setShareableByEditors(false);
  // 不呼叫 setSharing(ANYONE_WITH_LINK)：新照片維持 Drive 預設私人權限。
  return file.getId();
}

function countPhotos_(item) {
  let count = 0;
  [1,2,3].forEach(function(i) { if (resolvePhotoFileId_(item, i)) count++; });
  return count;
}

function resolvePhotoFileId_(item, number) {
  const newId = String(item['照片檔案ID' + number] || '').trim();
  if (newId) return newId;
  const legacyKey = number === 1 ? '照片連結' : '照片連結' + number;
  return extractDriveFileId_(String(item[legacyKey] || ''));
}

function extractDriveFileId_(url) {
  if (!url) return '';
  const match = url.match(/\/d\/([A-Za-z0-9_-]+)/) || url.match(/[?&]id=([A-Za-z0-9_-]+)/);
  return match ? match[1] : '';
}

function getExtension_(name, type) {
  if (String(name || '').indexOf('.') >= 0) return String(name).substring(String(name).lastIndexOf('.'));
  if (type === 'image/png') return '.png';
  if (type === 'image/webp') return '.webp';
  return '.jpg';
}

/******************** 通知與紀錄 ********************/
function sendNewRepairNotice_(item, photoCount) {
  if (!ADMIN_NOTIFY_EMAIL) return;
  MailApp.sendEmail({
    to: ADMIN_NOTIFY_EMAIL,
    subject: '【南帝報修】新案件 ' + item['報修編號'] + '｜' + item['地點'],
    name: '南帝精密報修系統',
    body: '有新的報修案件：\n\n報修編號：' + item['報修編號'] + '\n報修人：' + item['姓名'] + '\n部門：' + (item['部門'] || '-') + '\n地點：' + item['地點'] + '\n問題類型：' + item['問題類型'] + '\n緊急程度：' + item['緊急程度'] + '\n\n問題說明：\n' + item['問題說明'] + '\n\n附加照片：' + photoCount + ' 張\n\n請至管理後台查看與處理。'
  });
}

function sendCompletionNotice_(oldItem, newItem) {
  if (!newItem || !newItem['報修人Email']) return;
  if (newItem['狀態'] !== '已完成' || oldItem['狀態'] === '已完成') return;
  MailApp.sendEmail({
    to: String(newItem['報修人Email']).trim(),
    subject: '【南帝報修】您的案件已完成｜' + newItem['報修編號'],
    name: '南帝精密報修系統',
    body: '您好，您的報修案件已完成。\n\n報修編號：' + newItem['報修編號'] + '\n地點：' + newItem['地點'] + '\n問題：' + newItem['問題說明'] + '\n\n負責人：' + (newItem['負責人'] || '-') + '\n備註：' + (newItem['備註'] || '-')
  });
}

function appendLog_(caseId, oldStatus, newStatus, owner, dueDate, note) {
  getLogSheet_().appendRow([new Date(), caseId, oldStatus || '', newStatus, owner, dueDate, note]);
}

/******************** 試算表工具 ********************/
function getSpreadsheet_() {
  if (String(SPREADSHEET_ID || '').trim()) return SpreadsheetApp.openById(String(SPREADSHEET_ID).trim());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到 Google Sheet；獨立 GAS 專案請設定 SPREADSHEET_ID。');
  return ss;
}
function getSheet_() { const ss = getSpreadsheet_(); return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME); }
function getLogSheet_() { const ss = getSpreadsheet_(); return ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME); }
function getPhotoFolder_() { const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME); return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME); }
function ensureHeaders_(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
  if (!existing.some(function(v){ return v.trim(); })) { sheet.getRange(1, 1, 1, headers.length).setValues([headers]); return; }
  headers.forEach(function(header) { if (existing.indexOf(header) < 0) { sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header); existing.push(header); } });
}
function reorderColumns_(sheet, standardHeaders) {
  const data = sheet.getDataRange().getValues();
  if (!data.length) return;
  const current = data[0].map(function(v){ return String(v || '').trim(); });
  const extras = current.filter(function(h){ return h && standardHeaders.indexOf(h) < 0; });
  const ordered = standardHeaders.concat(extras), positions = {};
  current.forEach(function(h, i){ if (h) positions[h] = i; });
  const reordered = data.map(function(row, r){ return r === 0 ? ordered : ordered.map(function(h){ return positions[h] === undefined ? '' : row[positions[h]]; }); });
  sheet.clearContents();
  sheet.getRange(1, 1, reordered.length, ordered.length).setValues(reordered);
}
function styleSheet_(sheet) { sheet.setFrozenRows(1); sheet.getRange(1,1,1,sheet.getLastColumn()).setFontWeight('bold').setBackground('#d9eaf7').setHorizontalAlignment('center'); sheet.autoResizeColumns(1, sheet.getLastColumn()); }
function getHeaderMap_(sheet) { const map = {}; sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].forEach(function(h,i){ if (h) map[String(h).trim()] = i; }); return map; }
function appendObjectRow_(sheet, item) { const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]; sheet.appendRow(headers.map(function(h){ return item[String(h).trim()] === undefined ? '' : item[String(h).trim()]; })); }
function setCellByHeader_(sheet, row, map, name, value) { if (map[name] !== undefined) sheet.getRange(row, map[name] + 1).setValue(value); }
function getDataObjects_() { const data = getSheet_().getDataRange().getValues(); if (data.length <= 1) return []; return data.slice(1).filter(function(r){ return r[0]; }).map(function(r){ return rowArrayToObject_(data[0], r); }); }
function getRepairByCaseId_(id) { return getDataObjects_().filter(function(i){ return String(i['報修編號']) === String(id); })[0] || null; }
function rowArrayToObject_(headers, row) { const out = {}; headers.forEach(function(h,i){ if (h) out[String(h).trim()] = formatValue_(row[i]); }); return out; }
function formatValue_(value) { return Object.prototype.toString.call(value) === '[object Date]' ? Utilities.formatDate(value, TIME_ZONE, 'yyyy/MM/dd HH:mm') : value; }
function generateCaseId_() { return 'R' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyyMMdd-HHmmss') + Math.floor(Math.random() * 900 + 100); }
function parseDateTime_(value) { const m = String(value || '').match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2}))?/); return m ? new Date(+m[1], +m[2]-1, +m[3], +(m[4]||0), +(m[5]||0)) : null; }
function parseDateOnly_(value) { const m = String(value || '').match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/); return m ? new Date(+m[1], +m[2]-1, +m[3]) : null; }
function startOfToday_() { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }
function outputJson_(object, e) { const cb = e && e.parameter ? String(e.parameter.callback || '') : ''; return cb && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(cb) ? ContentService.createTextOutput(cb + '(' + JSON.stringify(object) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT) : ContentService.createTextOutput(JSON.stringify(object)).setMimeType(ContentService.MimeType.JSON); }
