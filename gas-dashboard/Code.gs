/**
 * 📊 國中生物課堂遊戲 - 後端服務與即時儀表板程式碼 (Code.gs)
 * 升級重點（V4.5 最新版）：
 * 1. 學生端提交成績 (POST)：相容大廳一與大廳二最新中介層 (seat, name, className, score, duration)
 * 2. 學生姓名後端智能檢驗與清洗：
 *    - 支援 2～5 字中文全名（完美支援「王拓」單名、三字姓名與複姓長名）
 *    - 自動過濾內部多餘空格（如「王 拓」自動修正為「王拓」）
 *    - 封鎖連續重複玩笑字（如「哈哈哈」）與英數亂碼，試算表自動標註方便老師追查
 *    - 內建試算表公式注入防禦（Formula Injection Defense）
 * 3. 10 秒防手震升級：修復中文系統「上午/下午」Date 解析為 NaN 的潛在 Bug，百分之百精準過濾連點
 * 4. 欄位純淨去班號：座號標準化為純兩位數字串 ('01'~'35')，班級為純三位數 ('701'~'710')，排序與試算表對齊零報錯
 * 5. 儀表板手動補登介面安全串接 (回傳純 JS 物件供 google.script.run 解析)
 * 6. 課堂截交鎖定開關控制 (PropertiesService 狀態儲存)
 * 7. 班級獨立工作表與「所有成績彙整」總表安全寫入
 */

// 1. 處理網頁開啟請求（大螢幕即時看板）
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('📊 生物課堂遊戲即時成績儀表板')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// 2. 處理遊戲端提交成績請求 (POST) - 含截交鎖定判斷與防手震防呆
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // 等候鎖定最多 10 秒，避免多台平板同時送出時搶佔試算表衝突
    lock.waitLock(10000);

    // 檢查老師端是否開啟了「截止交卷」鎖定開關
    var props = PropertiesService.getScriptProperties();
    var isLocked = props.getProperty('SUBMISSION_LOCKED') === 'true';
    if (isLocked) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: '🛑 課堂挑戰已由老師截止交卷，請專心聽講！' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = {};
      }
    } else {
      data = {
        className: '701',
        seat: '01',
        name: '王拓',
        gameTitle: '測試遊戲',
        score: 100,
        duration: '01:00'
      };
    }
    
    // 寫入試算表 (自動執行防手震比對與姓名檢驗)
    saveRecordToSheetsInternal(data);

    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'success', 
      message: 'Score saved successfully' 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

// 3. 處理儀表板前端呼叫的「手動補登成績」請求 (不受交卷鎖定限制，回傳純 JS 物件)
function saveManualScoreBackend(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    saveRecordToSheetsInternal(data);
    return { status: 'success', message: 'Score saved successfully' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

// 4. 切換截止交卷鎖定狀態
function setLockStatusBackend(isLocked) {
  PropertiesService.getScriptProperties().setProperty('SUBMISSION_LOCKED', String(isLocked));
  return { status: 'success', isLocked: isLocked };
}

// 核心輔助：學生姓名後端智能檢驗與清洗 (支援 2~5 字中文全名，包含王拓等單名)
function sanitizeStudentName(rawName, seat) {
  if (!rawName) return seat ? (seat + '號同學') : '未具名';
  var str = String(rawName).replace(/\s+/g, '').trim();

  // 防範試算表公式注入符號 (=, +, -, @)
  if (/^[=\+\-@]/.test(str)) {
    str = "'" + str;
  }

  // 檢驗繁體中文 2~5 字
  var isChinese = /^[\u4e00-\u9fa5]{2,5}$/.test(str);
  var isRepeated = /(.)\1{2,}/.test(str); // 三連字如 哈哈哈、啊啊啊

  // 合法中文姓名 (王拓、王小明、歐陽修德、張簡美麗)
  if (isChinese && !isRepeated) {
    return str;
  }

  // 非正規姓名後端標註處理，方便老師在試算表一眼抓出
  if (str.length > 0) {
    if (isRepeated) {
      return str + ' ⚠️(重複玩笑字)';
    }
    if (!isChinese) {
      return str + ' ⚠️(待核對名冊)';
    }
    return str;
  }

  return seat ? (seat + '號同學') : '未具名';
}

// 核心輔助：安全解析試算表中的時間字串為毫秒 timestamp（相容中英文與 Date 物件）
function parseTimestampMs(raw) {
  if (!raw) return 0;
  if (raw instanceof Date) return raw.getTime();
  if (typeof raw === 'number') return raw;
  var str = String(raw).trim();
  if (!str) return 0;

  // 處理繁體中文 '上午' / '下午'
  var clean = str.replace(/上午/g, 'AM ').replace(/下午/g, 'PM ');
  var parsed = Date.parse(clean);
  if (!isNaN(parsed)) return parsed;

  // 一般 ISO 格式或直接 new Date 嘗試
  var d = new Date(str);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

// 核心函式：安全寫入試算表（含姓名檢核與 10 秒內連點自動過濾）
function saveRecordToSheetsInternal(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 班級標準化 (純三位數 701~710)
  var rawClass = String(data.className || data.class || '701').trim();
  var matchClass = rawClass.match(/(?:\d{3}|7[甲乙丙])/);
  var className = matchClass ? matchClass[0] : (rawClass.replace(/[^0-9]/g, '') || '701');

  // 2. 座號標準化 (純兩位數 01~35)
  var rawSeat = String(data.seat || data.seatOrGroup || '').trim();
  var matchSeat = rawSeat.match(/\d+/);
  var seat = matchSeat ? String(matchSeat[0]).padStart(2, '0') : '';
  
  // 3. 姓名後端標準化與檢核（支援王拓單名，過濾亂碼）
  var rawName = String(data.name || data.studentName || data.nameOrMembers || '').trim();
  var name = sanitizeStudentName(rawName, seat);

  // 4. 遊戲名稱與單元簡化
  var gameTitle = String(data.gameTitle || '課堂遊戲').trim();
  var mode = (data.mode === 'group' || String(data.seatOrGroup || '').includes('組')) ? '👥 小組' : '👤 個人';
  var score = data.score !== undefined ? Number(data.score) : 0;
  var duration = String(data.duration || '--:--').trim();
  
  var now = new Date();
  var nowMs = now.getTime();
  var timestamp = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

  // A. 寫入按「班級」自動建立的工作表
  var targetSheet = sheet.getSheetByName(className);
  if (!targetSheet) {
    targetSheet = sheet.insertSheet(className);
    targetSheet.appendRow(['提交時間', '遊戲名稱', '模式', '座號', '姓名', '得分', '遊戲耗時']);
    targetSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e0f2fe');
    targetSheet.setFrozenRows(1);
  }

  // 🛡️ 學生緊張連點 / 網路 Lag 重複送出過濾機制 (10 秒防手震)
  var lastRow = targetSheet.getLastRow();
  if (lastRow > 1) {
    var lastRecord = targetSheet.getRange(lastRow, 1, 1, 7).getValues()[0];
    var lastTimeMs = parseTimestampMs(lastRecord[0]);
    var lastSeatRaw = String(lastRecord[3] || '').trim();
    var lastSeatMatch = lastSeatRaw.match(/\d+/);
    var lastSeat = lastSeatMatch ? String(lastSeatMatch[0]).padStart(2, '0') : lastSeatRaw;
    var lastScore = Number(lastRecord[5]);

    // 若同一座號、分數完全一致，且兩次送出間隔小於 10 秒 (10000 毫秒)，視為重複送出，直接略過寫入
    var timeDiff = Math.abs(nowMs - lastTimeMs);
    if (lastSeat === seat && lastScore === score && timeDiff < 10000) {
      return; // 成功攔截重複連點
    }
  }

  // 正常寫入班級分頁 (座號為純數字 01~35，姓名為正規中文名)
  targetSheet.appendRow([timestamp, gameTitle, mode, seat, name, score, duration]);

  // B. 彙整寫入「所有成績彙整」總表
  var allSheet = sheet.getSheetByName('所有成績彙整');
  if (!allSheet) {
    allSheet = sheet.insertSheet('所有成績彙整');
    allSheet.appendRow(['提交時間', '班級', '遊戲名稱', '模式', '座號', '姓名', '得分', '遊戲耗時']);
    allSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#fef08a');
    allSheet.setFrozenRows(1);
  }
  allSheet.appendRow([timestamp, className, gameTitle, mode, seat, name, score, duration]);
}

// 5. 供儀表板即時抓取資料與交卷鎖定狀態
function getScoresData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheet = sheet.getSheetByName('所有成績彙整');
  var props = PropertiesService.getScriptProperties();
  var isLocked = props.getProperty('SUBMISSION_LOCKED') === 'true';
  
  if (!allSheet) {
    return { records: [], classes: [], isLocked: isLocked };
  }

  var data = allSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { records: [], classes: [], isLocked: isLocked };
  }

  var records = [];
  var classSet = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rawCls = String(row[1] || '').trim();
    var matchCls = rawCls.match(/\d{3}/);
    var cls = matchCls ? matchCls[0] : rawCls;
    if (cls) classSet[cls] = true;

    // 座號補零確保 01~35 純數字
    var rawSeat = String(row[4] || '').trim();
    var matchSeat = rawSeat.match(/\d+/);
    var seatClean = matchSeat ? String(matchSeat[0]).padStart(2, '0') : rawSeat;

    records.push({
      timestamp: row[0],
      className: cls,
      gameTitle: row[2],
      mode: row[3],
      seatOrGroup: seatClean,
      nameOrMembers: row[5],
      score: Number(row[6]) || 0,
      duration: row[7]
    });
  }

  records.reverse(); // 最新資料放最上面

  return {
    records: records,
    classes: Object.keys(classSet).sort(),
    isLocked: isLocked
  };
}
