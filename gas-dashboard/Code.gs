/**
 * ==========================================================================
 * 🌿 國中生物快閃測驗後端服務與即時教師儀表板（復古琥珀金版 - 旗艦升級）
 * 包含：
 * 1. 方案 A：依測驗範圍（如「南一版 (2-3 植物如何製造養分)」）自動建立獨立分頁
 * 2. 智慧整理最優成績：前三欄依序為【座號】、【姓名】、【採計總分(最高分)】，方便老師一秒複製登記！
 * 3. 學生就地更新 (In-Place Update)：同生交卷/訂正/加分增量更新最高分，不製造多列重複
 * 4. 自動座號排序：自動按座號 01~35 順號排列，與學校點名冊 100% 吻合
 * 5. 班級格式全面相容：710 班與 7年10班 自動匹配，大螢幕點名卡片即時亮綠燈
 * 6. 原有完整功能全數保留：左上請假角標、深色模式、導師分數遮蔽、手動補登彈窗、CSV匯出、未交名單複製
 * ==========================================================================
 */

// 1. 清洗分頁名稱（排除試算表禁止字元 \ / ? * : [ ] 並限制 50 字）
function cleanSheetName(rawUnits) {
  if (!rawUnits) return "測驗成績紀錄";
  var clean = String(rawUnits).replace(/[\\\/?:*\[\]]/g, "-").trim();
  if (clean.length > 50) clean = clean.substring(0, 50);
  return clean || "測驗成績紀錄";
}

// 2. 座號補零標準化（例如 "2" -> "02號"，確保排序與比對一致）
function formatSeatWithPadding(seatStr) {
  if (!seatStr) return "";
  var match = String(seatStr).match(/\d+/);
  if (!match) return String(seatStr).trim();
  return String(match[0]).padStart(2, "0") + "號";
}

// 3. 班級字串正規化（相容 "710"、"7年10班"、"七年十班"、"710班"）
function normalizeClassStr(input) {
  if (!input) return "";
  var s = String(input).trim().replace(/班$/, "");
  var matchSpecial = s.match(/(?:7|七)[\s-]*([甲乙丙])/);
  if (matchSpecial) return "七" + matchSpecial[1];
  if (/^[甲乙丙]$/.test(s)) return "七" + s;

  var digits = s.match(/\d+/g);
  if (digits) {
    if (digits.length >= 2) return digits[0] + digits[1].padStart(2, "0");
    if (digits.length === 1 && digits[0].length === 3) return digits[0];
    if (digits.length <= 2) return "7" + digits[0].padStart(2, "0");
  }
  if (s.includes("十")) return "710";
  var cnMap = {"一":"01","二":"02","三":"03","四":"04","五":"05","六":"06","七":"07","八":"08","九":"09"};
  for (var k in cnMap) {
    if (s.includes(k)) return "7" + cnMap[k];
  }
  return s;
}

// 4. 輔助函式：確保分頁具備標準化表頭
function ensureStandardHeader(sheet, isClassSheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "座號", "姓名", "採計總分(最高分)", "班級", "交卷時間", 
      "原始得分(第1階段)", "訂正後得分(第2階段)", "加分後總分(第3階段)", 
      "錯題數", "錯題與補救訂正關鍵字", "跳出視窗次數", "狀態備註", "測驗主題/遊戲"
    ]);
    var headerColor = isClassSheet ? "#0f766e" : "#b45309"; // 班級分頁為翡翠綠，單元/彙整總表為琥珀金
    sheet.getRange(1, 1, 1, 13).setBackground(headerColor).setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(3); // 凍結座號、姓名、最高採計分
  }
}

// 4-1. 輔助函式：就地更新 (In-Place Update) 或新增資料至指定分頁
function saveRecordToSheet(sheet, data, isClassSheet) {
  if (!sheet) return { bestScore: 0, isUpdate: false };
  ensureStandardHeader(sheet, isClassSheet);

  var targetClass = String(data.className || "").trim();
  var targetClassNorm = normalizeClassStr(targetClass);
  var targetSeat = formatSeatWithPadding(data.seat);
  var targetName = String(data.name || "").trim();
  var nowTime = data.timestamp || new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" });
  var switchCount = data.switchCount || 0;
  var switchAlert = switchCount > 3 ? (switchCount + " 次 (警示)") : (switchCount > 0 ? switchCount + " 次" : "0 次 (正常)");
  var unitTitle = String(data.units || data.gameTitle || "課堂測驗").trim();

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return String(h).trim(); });
  var isNewFormat = headers[0] === "座號" && headers[2] && headers[2].includes("採計");

  var existingRowIndex = -1;
  var existingOriginal = 0;
  var existingCorrection = 0;
  var existingBonus = 0;
  var existingBest = 0;

  for (var i = 1; i < rows.length; i++) {
    if (isNewFormat) {
      var rSeat = formatSeatWithPadding(rows[i][0]);
      var rName = String(rows[i][1] || "").trim();
      var rClass = String(rows[i][3] || "").trim();
      
      // 比對條件：
      // 在「班級分頁」中，同座號 + 同姓名即為該學生（精確就地更新成績）
      // 在「單元分頁」或「總表」中，同班級 + 同座號 + 同姓名即為該學生
      var isMatch = isClassSheet ? (rSeat === targetSeat && rName === targetName)
                                 : (normalizeClassStr(rClass) === targetClassNorm && rSeat === targetSeat && rName === targetName);
      if (isMatch) {
        existingRowIndex = i + 1;
        existingBest = Number(rows[i][2]) || 0;
        existingOriginal = Number(rows[i][5]) || 0;
        existingCorrection = Number(rows[i][6]) || 0;
        existingBonus = Number(rows[i][7]) || 0;
        break;
      }
    } else {
      var rClassOld = String(rows[i][1] || "").trim();
      var rSeatOld = formatSeatWithPadding(rows[i][2]);
      var rNameOld = String(rows[i][3] || "").trim();
      if (normalizeClassStr(rClassOld) === targetClassNorm && rSeatOld === targetSeat && rNameOld === targetName) {
        existingRowIndex = i + 1;
        existingOriginal = Number(rows[i][5]) || 0;
        existingBest = Number(rows[i][6]) || existingOriginal;
        break;
      }
    }
  }

  var incomingScore = data.finalScore !== undefined ? Number(data.finalScore) : (data.score !== undefined ? Number(data.score) : (data.originalScore !== undefined ? Number(data.originalScore) : 0));
  var currentOriginal = data.originalScore !== undefined ? Number(data.originalScore) : (data.score !== undefined ? Number(data.score) : existingOriginal);
  var currentCorrection = data.isCorrectionPhase ? Number(data.finalScore || incomingScore) : existingCorrection;
  var currentBonus = data.isFinalBonus ? Number(data.finalScore || currentCorrection) : existingBonus;

  // 核心最高分採計邏輯：自動取學生所有階段達到的最優最高分
  var currentBest = Math.max(existingBest, currentOriginal, currentCorrection, currentBonus, incomingScore);

  var statusNote = "正常繳交";
  if (data.isFinalBonus) statusNote = "🏆 已完成加分題";
  else if (data.isCorrectionPhase) statusNote = "✏️ 訂正完成送出";
  else if (data.isInitialSync) statusNote = "⚡ 第1階段成績已送出";
  if (data.note) statusNote = data.note;

  // 情況 A：已有舊紀錄 -> 就地更新 (In-Place Update)
  if (existingRowIndex > 0) {
    if (isNewFormat) {
      sheet.getRange(existingRowIndex, 3).setValue(currentBest); // C欄最高採計分
      sheet.getRange(existingRowIndex, 5).setValue(nowTime);     // E欄交卷時間
      if (data.originalScore !== undefined) sheet.getRange(existingRowIndex, 6).setValue(currentOriginal);
      if (data.isCorrectionPhase) sheet.getRange(existingRowIndex, 7).setValue(currentCorrection);
      if (data.isFinalBonus) sheet.getRange(existingRowIndex, 8).setValue(currentBonus);
      if (data.errorCount !== undefined) sheet.getRange(existingRowIndex, 9).setValue(data.errorCount);
      if (data.keywords) sheet.getRange(existingRowIndex, 10).setValue(data.keywords);
      sheet.getRange(existingRowIndex, 11).setValue(switchAlert);
      sheet.getRange(existingRowIndex, 12).setValue(statusNote);
      if (sheet.getLastColumn() >= 13) {
        sheet.getRange(existingRowIndex, 13).setValue(unitTitle);
      }
      if (switchCount >= 3) {
        sheet.getRange(existingRowIndex, 1, 1, sheet.getLastColumn()).setBackground("#FEF3C7");
      }
    } else {
      sheet.getRange(existingRowIndex, 1).setValue(nowTime);
      sheet.getRange(existingRowIndex, 6).setValue(currentOriginal);
      sheet.getRange(existingRowIndex, 7).setValue(currentBest);
      sheet.getRange(existingRowIndex, 10).setValue(switchAlert);
      sheet.getRange(existingRowIndex, 11).setValue(statusNote);
    }
    return { bestScore: currentBest, isUpdate: true };
  }

  // 情況 B：全新紀錄 -> 新增一列
  if (isNewFormat) {
    var newRow = [
      targetSeat,                                              // A: 座號 (01號)
      targetName,                                              // B: 姓名
      currentBest,                                             // C: 採計總分(最高分)
      targetClass,                                             // D: 班級
      nowTime,                                                 // E: 交卷時間
      currentOriginal,                                         // F: 原始得分
      data.isCorrectionPhase ? currentCorrection : "-",        // G: 訂正後得分
      data.isFinalBonus ? currentBonus : "-",                  // H: 加分後總分
      data.errorCount !== undefined ? data.errorCount : 0,     // I: 錯題數
      data.keywords || "全對無錯題",                           // J: 錯題訂正心得
      switchAlert,                                             // K: 跳視窗次數
      statusNote,                                              // L: 狀態備註
      unitTitle                                                // M: 測驗主題/遊戲
    ];
    sheet.appendRow(newRow);

    if (switchCount >= 3) {
      sheet.getRange(sheet.getLastRow(), 1, 1, 13).setBackground("#FEF3C7");
    }

    // 自動按座號 (A 欄) 排序 (01, 02, 03... 30, 35)，與點名冊完全一致
    if (sheet.getLastRow() > 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({ column: 1, ascending: true });
    }
  } else {
    sheet.appendRow([
      nowTime, targetClass, targetSeat, targetName, unitTitle,
      currentOriginal, currentBest, data.errorCount || 0, data.keywords || "全對無錯題",
      switchAlert, statusNote
    ]);
  }

  return { bestScore: currentBest, isUpdate: false };
}

// 4. 學生測驗端提交成績 (POST) - 雙軌精準分流：自動更新班級分頁 + 彙整總表
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = {
        className: "710",
        seat: "01",
        name: "測試學生",
        units: "南一版 (2-3 植物如何製造養分)",
        originalScore: 100,
        finalScore: 100,
        errorCount: 0,
        keywords: "全對",
        switchCount: 0
      };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rawClass = String(data.className || "").trim();
    var unitTitle = String(data.units || data.gameTitle || "").trim();

    // 1️⃣ 第一軌：尋找或建立該學生的【班級專屬分頁】（如 701, 702, 七乙, 七丙）
    // 進行「座號就地更新 (In-Place Update)」，一人一列，保持乾淨排序！
    var classSheet = null;
    if (rawClass) {
      classSheet = ss.getSheetByName(rawClass);
      if (!classSheet) {
        var normCls = normalizeClassStr(rawClass);
        classSheet = ss.getSheetByName(normCls);
      }
      if (!classSheet) {
        classSheet = ss.insertSheet(rawClass);
      }
    }

    var classRes = { bestScore: 0, isUpdate: false };
    if (classSheet) {
      classRes = saveRecordToSheet(classSheet, data, true);
    }

    // 2️⃣ 第二軌：同步寫入【單元分頁】或【所有成績彙整】/【測驗成績紀錄】總表
    var secondarySheetName = cleanSheetName(unitTitle);
    if (!secondarySheetName || secondarySheetName === "測驗成績紀錄") {
      secondarySheetName = ss.getSheetByName("所有成績彙整") ? "所有成績彙整" : "測驗成績紀錄";
    }
    var secondarySheet = ss.getSheetByName(secondarySheetName);
    if (!secondarySheet) {
      secondarySheet = ss.insertSheet(secondarySheetName);
    }
    var secRes = saveRecordToSheet(secondarySheet, data, false);

    var finalBest = Math.max(classRes.bestScore, secRes.bestScore);

    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "success", 
      "message": "成績已成功存入班級分頁【" + (classSheet ? classSheet.getName() : "") + "】與彙整分頁【" + secondarySheetName + "】！最高採計分：" + finalBest + "分！",
      "bestScore": finalBest,
      "isUpdate": classRes.isUpdate
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "error", 
      "message": error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 5. 處理儀表板手動補登成績請求（同步就地更新班級分頁與彙整分頁）
function saveManualScoreBackend(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rawClass = String(data.className || "").trim();
    var unitTitle = String(data.units || data.gameTitle || "課堂實體手動補登").trim();

    var classSheet = null;
    if (rawClass) {
      classSheet = ss.getSheetByName(rawClass) || ss.getSheetByName(normalizeClassStr(rawClass));
      if (!classSheet) classSheet = ss.insertSheet(rawClass);
    }

    var manualData = {
      className: rawClass,
      seat: data.seat,
      name: data.name,
      originalScore: Number(data.score) || 0,
      finalScore: Number(data.score) || 0,
      units: unitTitle,
      errorCount: 0,
      keywords: "老師手動補登",
      switchCount: 0,
      note: data.note ? ("✍️ " + data.note) : "✍️ 老師手動補登"
    };

    if (classSheet) {
      saveRecordToSheet(classSheet, manualData, true);
    }

    var secSheet = ss.getSheetByName(cleanSheetName(unitTitle)) || ss.getSheetByName("所有成績彙整") || ss.getSheetByName("測驗成績紀錄");
    if (!secSheet) secSheet = ss.insertSheet("測驗成績紀錄");
    saveRecordToSheet(secSheet, manualData, false);

    return { status: "success", message: "手動補登成功！已同步更新至班級專屬分頁。" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// 6. 供儀表板讀取資料（自動跨所有測驗分頁彙整全部學生資料）
function getScoresData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var records = [];
  var classSet = {};
  
  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetName.includes("說明") || sheetName.includes("範本")) return;
    
    var rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;
    
    var headers = rows[0].map(function(h) { return String(h).trim(); });
    var isNewFormat = headers[0] === "座號" && headers[2] && headers[2].includes("採計");
    var isOldFormat = headers[0] && headers[0].includes("時間") && headers[1] === "班級";
    
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (isNewFormat) {
        var cls = String(r[3] || "").trim();
        if (cls) classSet[cls] = true;
        records.push({
          seat: String(r[0] || "").replace(/號/g, "").trim(),
          name: r[1],
          finalScore: Number(r[2]) || 0,
          className: r[3],
          timestamp: r[4],
          originalScore: Number(r[5]) || 0,
          correctionScore: r[6],
          bonusScore: r[7],
          errorCount: Number(r[8]) || 0,
          keywords: r[9],
          switchCount: r[10],
          statusNote: r[11],
          units: sheetName
        });
      } else if (isOldFormat) {
        var clsOld = String(r[1] || "").trim();
        if (clsOld) classSet[clsOld] = true;
        records.push({
          timestamp: r[0],
          className: r[1],
          seat: String(r[2] || "").replace(/號/g, "").trim(),
          name: r[3],
          units: r[4] || sheetName,
          originalScore: r[5],
          finalScore: Number(r[6]) !== undefined ? Number(r[6]) : (Number(r[5]) || 0),
          errorCount: r[7],
          keywords: r[8],
          switchCount: r[9],
          statusNote: r[10]
        });
      }
    }
  });

  records.reverse(); // 最新資料在最前
  return {
    records: records,
    classes: Object.keys(classSet).sort()
  };
}



// ==========================================================================
// 🌿 試算表自訂功能選單與一鍵美化優化引擎
// ==========================================================================

// 1. 試算表頂端選單擴充：方便老師在 Google 試算表一鍵執行美化、修復與重新自然排序
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("🌿 生物課堂試算表小幫手")
      .addItem("✨ 一鍵套用頂級美化排版 (當前分頁)", "美化當前分頁")
      .addItem("🎨 一鍵美化所有測驗分頁", "美化所有分頁")
      .addSeparator()
      .addItem("🚀 一鍵自動分流歷史成績到各班分頁", "一鍵將歷史成績分流至各班")
      .addItem("🔄 重新依座號排序當前分頁", "重新排序當前分頁")
      .addItem("🛠️ 一鍵修復並自動補齊所有分頁", "一鍵修復並美化所有分頁")
      .addToUi();
  } catch (e) {}
}

// 輔助功能：將「測驗成績紀錄」或「所有成績彙整」中的現有舊成績，自動分流填入對應班級分頁（701~710、七乙、七丙等）
function 一鍵將歷史成績分流至各班() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("測驗成績紀錄") || ss.getSheetByName("所有成績彙整");
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("找不到【測驗成績紀錄】或【所有成績彙整】分頁！");
    return;
  }
  
  var rows = masterSheet.getDataRange().getValues();
  if (rows.length <= 1) {
    SpreadsheetApp.getUi().alert("總表中無成績資料可分流！");
    return;
  }
  
  var headers = rows[0].map(function(h) { return String(h).trim(); });
  var isNewFormat = headers[0] === "座號" && headers[2] && headers[2].includes("採計");
  
  var count = 0;
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var data = {};
    if (isNewFormat) {
      data = {
        seat: r[0],
        name: r[1],
        finalScore: r[2],
        className: r[3],
        timestamp: r[4],
        originalScore: r[5],
        isCorrectionPhase: r[6] && r[6] !== "-",
        isFinalBonus: r[7] && r[7] !== "-",
        errorCount: r[8],
        keywords: r[9],
        switchCount: r[10],
        note: r[11],
        units: r[12] || masterSheet.getName()
      };
    } else {
      data = {
        timestamp: r[0],
        className: r[1],
        seat: r[2],
        name: r[3],
        units: r[4],
        originalScore: r[5],
        finalScore: r[6],
        errorCount: r[7],
        keywords: r[8],
        switchCount: r[9],
        note: r[10]
      };
    }
    
    var rawClass = String(data.className || "").trim();
    if (!rawClass) continue;
    
    var classSheet = ss.getSheetByName(rawClass) || ss.getSheetByName(normalizeClassStr(rawClass));
    if (!classSheet) {
      classSheet = ss.insertSheet(rawClass);
    }
    saveRecordToSheet(classSheet, data, true);
    count++;
  }
  
  SpreadsheetApp.getUi().alert("🎉 分流完成！已成功將 " + count + " 筆學生成績按座號就地歸納更新至所屬班級分頁！");
}

// 2. 一鍵美化當前工作表
function 美化當前分頁() {
  var sheet = SpreadsheetApp.getActiveSheet();
  applyUltraBeautification(sheet);
  try {
    SpreadsheetApp.getUi().alert("🎉 美化完成！已套用深翡翠綠奢華表頭、護眼斑馬紋、最適欄寬、座號姓名凍結與滿分高亮！");
  } catch (_) {}
}

// 3. 一鍵美化所有工作表
function 美化所有分頁() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var count = 0;
  sheets.forEach(function(s) {
    var name = s.getName();
    if (!name.includes("說明") && !name.includes("範本")) {
      applyUltraBeautification(s);
      count++;
    }
  });
  try {
    SpreadsheetApp.getUi().alert("🎉 全數完成！已為 " + count + " 個分頁套用頂級現代排版、斑馬紋與高低分智慧標色！");
  } catch (_) {}
}

// 4. 重新依「班級 ➔ 座號」排序當前分頁
function 重新排序當前分頁() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    var headers = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 12)).getValues()[0];
    var isNewFormat = headers[0] === "座號" && headers[2] && String(headers[2]).includes("採計");
    if (isNewFormat) {
      sheet.getRange(2, 1, lastRow - 1, 12).sort([
        { column: 4, ascending: true }, // 班級
        { column: 1, ascending: true }  // 座號
      ]);
    } else {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort([
        { column: 2, ascending: true }, // 班級
        { column: 3, ascending: true }  // 座號
      ]);
    }
  }
  applyUltraBeautification(sheet);
  try {
    SpreadsheetApp.getUi().alert("✅ 排序完成！已將當前分頁依「班級 ➔ 座號 (01~35)」精準自然升冪排列。");
  } catch (_) {}
}

// 5. 一鍵修復並美化所有分頁
function 一鍵修復並美化所有分頁() {
  美化所有分頁();
}

// 6. 頂級美化排版核心引擎（深翡翠綠/琥珀金、護眼斑馬紋、最適欄寬、座號姓名雙凍結、高低分條件格式化）
function applyUltraBeautification(sheet) {
  if (!sheet) return;
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var numCols = Math.max(sheet.getLastColumn(), 12);

  // 1. 表頭高度與雙向凍結 (表頭 40px 大氣舒適，凍結前兩欄座號與姓名)
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // 2. 表頭現代深翡翠綠風格 (#0F766E)
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground("#0F766E")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(10.5)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // 3. 黃金比例最適欄寬
  var colWidths = [
    65,  // A: 座號 (01~35)
    85,  // B: 姓名
    115, // C: 採計總分(最高分)
    65,  // D: 班級 (701~710)
    140, // E: 交卷時間
    95,  // F: 原始得分(第1階段)
    95,  // G: 訂正後得分(第2階段)
    95,  // H: 加分後總分(第3階段)
    70,  // I: 錯題數
    180, // J: 錯題與補救訂正關鍵字
    95,  // K: 跳出視窗次數
    125  // L: 狀態備註
  ];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // 4. 資料列排版、行高、斑馬紋與高低分條件格式化
  if (lastRow > 1) {
    sheet.setRowHeights(2, lastRow - 1, 28); // 舒適列高 28px

    var dataRange = sheet.getRange(2, 1, lastRow - 1, numCols);
    dataRange
      .setVerticalAlignment("middle")
      .setFontSize(10);

    // 細項欄位排版
    sheet.getRange(2, 1, lastRow - 1, 1).setHorizontalAlignment("center"); // 座號
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment("center"); // 姓名
    sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11); // 採計最高分
    sheet.getRange(2, 4, lastRow - 1, 1).setHorizontalAlignment("center"); // 班級
    sheet.getRange(2, 5, lastRow - 1, 1).setHorizontalAlignment("center").setFontColor("#64748B"); // 時間灰
    sheet.getRange(2, 6, lastRow - 1, 3).setHorizontalAlignment("center"); // 得分置中
    sheet.getRange(2, 9, lastRow - 1, 1).setHorizontalAlignment("center"); // 錯題數
    sheet.getRange(2, 10, lastRow - 1, 1).setHorizontalAlignment("left");  // 關鍵字靠左
    sheet.getRange(2, 11, lastRow - 1, 2).setHorizontalAlignment("center"); // 跳窗、備註

    // 清除舊斑馬紋並套用清新護眼斑馬紋
    try {
      var bandings = sheet.getBandings();
      for (var b = 0; b < bandings.length; b++) {
        bandings[b].remove();
      }
      var fullTableRange = sheet.getRange(1, 1, lastRow, numCols);
      var newBanding = fullTableRange.applyRowBanding();
      newBanding
        .setHeaderRowColor("#0F766E")
        .setFirstRowColor("#FFFFFF")
        .setSecondRowColor("#F8FAFC");
    } catch (_) {}

    // 條件式格式化 (滿分金黃高亮、不及格粉紅高亮、跳窗警示)
    try {
      var rules = [];
      var scoreRange = sheet.getRange("C2:C" + lastRow);
      
      // 滿分 (>=100)：香檳金底 + 琥珀金深字 + 粗體
      var rule100 = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(100)
        .setBackground("#FEF3C7")
        .setFontColor("#92400E")
        .setBold(true)
        .setRanges([scoreRange])
        .build();
      rules.push(rule100);

      // 不及格 (<60)：櫻粉色底 + 警示深紅字
      var ruleFail = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(60)
        .setBackground("#FEE2E2")
        .setFontColor("#991B1B")
        .setRanges([scoreRange])
        .build();
      rules.push(ruleFail);

      // 跳窗警示：柔和橙黃底 + 橘紅字
      var switchRange = sheet.getRange("K2:K" + lastRow);
      var ruleSwitch = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("警示")
        .setBackground("#FFEDD5")
        .setFontColor("#C2410C")
        .setBold(true)
        .setRanges([switchRange])
        .build();
      rules.push(ruleSwitch);

      sheet.setConditionalFormatRules(rules);
    } catch (_) {}

    // 自動開啟表頭下拉篩選箭頭
    try {
      if (!sheet.getFilter()) {
        sheet.getRange(1, 1, lastRow, numCols).createFilter();
      }
    } catch (_) {}
  }
}

// 7. 儀表板 HTML 介面與外部接口（支援 JSONP 跨域免疫、JSON API 與 直接大螢幕看板）
// 7. 儀表板 HTML 介面與外部接口（單一檔案自包含 100% 免手動建 index.html）
function doGet(e) {
  var callback = e && e.parameter ? e.parameter.callback : "";
  var action = e && e.parameter ? e.parameter.action : "";
  
  // 外部透過 JSONP 請求
  if (callback) {
    var dataJsonp = getScoresData();
    var output = callback + "(" + JSON.stringify(dataJsonp) + ");";
    return ContentService.createTextOutput(output)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  // 外部透過標準 JSON API 請求
  if (action === "getLiveScores") {
    var dataJson = getScoresData();
    return ContentService.createTextOutput(JSON.stringify(dataJson))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 優先嘗試載入獨立 index.html；若無則自動使用內建完整儀表板（單檔即開即用）
  try {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("🏆 國中生物測驗與遊戲即時儀表板")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  } catch (_) {
    return getBuiltinDashboardHtml();
  }
}

// 8. 內建儀表板完整網頁模組（真實換行與縮排排版，拒絕轉義混亂亂碼）
function getBuiltinDashboardHtml() {
  var html = `
<!DOCTYPE html>
<html lang="zh-TW" data-hall-theme="green" data-color-mode="light">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>🏆 國中生物測驗與遊戲即時儀表板</title>
  <link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>🏆</text></svg>">
  <style>
    /* ========================================================
       🎨 雙大廳色系定義：大廳一 (翠綠) vs 大廳二 (玫瑰粉)
       🌓 雙模式支援：明亮模式 (預設) vs 深色模式
       ======================================================== */
    
    /* 1. 大廳一 ‧ 翠綠主題 (明亮) */
    :root, [data-hall-theme="green"][data-color-mode="light"] {
      --bg-page: radial-gradient(circle at 50% 0%, #f0fdf4 0%, #f8fafc 60%, #e2e8f0 100%);
      --body-bg: #f8fafc;
      --card-bg: #ffffff;
      --card-border: #bbf7d0;
      --card-shadow: 0 4px 20px rgba(22, 163, 74, 0.08);
      --inner-bg: #f0fdf4;
      --inner-border: #dcfce7;
      
      --text-main: #14532d;
      --text-sub: #166534;
      --text-muted: #15803d;
      --text-light: #16a34a;
      
      --primary: #16a34a;
      --primary-hover: #15803d;
      --accent-green: #16a34a;
      --accent-red: #dc2626;
      --accent-blue: #2563eb;
      
      --card-submitted-bg: #dcfce7;
      --card-submitted-border: #16a34a;
      --card-submitted-text: #14532d;
      --card-submitted-tag: #15803d;
      
      --card-absent-bg: #fee2e2;
      --card-absent-border: #ef4444;
      --card-absent-text: #991b1b;
      --card-absent-tag: #dc2626;
      
      --card-pending-bg: #ffffff;
      --card-pending-border: #bbf7d0;
      --card-pending-text: #166534;
      --card-pending-tag: #15803d;

      --theme-toggle-bg: #ffffff;
      --theme-toggle-border: #86efac;
      --theme-toggle-text: #14532d;
    }

    /* 2. 大廳二 ‧ 星空魅紫主題 (明亮) */
    [data-hall-theme="purple"][data-color-mode="light"],
    [data-hall-theme="pink"][data-color-mode="light"] {
      --bg-page: radial-gradient(circle at 50% 0%, #f5f3ff 0%, #faf5ff 60%, #ede9fe 100%);
      --body-bg: #faf5ff;
      --card-bg: #ffffff;
      --card-border: #ddd6fe;
      --card-shadow: 0 4px 20px rgba(124, 58, 237, 0.08);
      --inner-bg: #f5f3ff;
      --inner-border: #ede9fe;
      
      --text-main: #4c1d95;
      --text-sub: #5b21b6;
      --text-muted: #6d28d9;
      --text-light: #7c3aed;
      
      --primary: #7c3aed;
      --primary-hover: #6d28d9;
      --accent-green: #7c3aed;
      --accent-red: #dc2626;
      --accent-blue: #0284c7;
      
      --card-submitted-bg: #ede9fe;
      --card-submitted-border: #7c3aed;
      --card-submitted-text: #4c1d95;
      --card-submitted-tag: #6d28d9;
      
      --card-absent-bg: #fee2e2;
      --card-absent-border: #ef4444;
      --card-absent-text: #991b1b;
      --card-absent-tag: #dc2626;
      
      --card-pending-bg: #ffffff;
      --card-pending-border: #ddd6fe;
      --card-pending-text: #6d28d9;
      --card-pending-tag: #7c3aed;

      --theme-toggle-bg: #ffffff;
      --theme-toggle-border: #c4b5fd;
      --theme-toggle-text: #4c1d95;
    }

    /* 3. 大廳一 ‧ 翠綠主題 (深色星艦) */
    [data-hall-theme="green"][data-color-mode="dark"] {
      --bg-page: radial-gradient(circle at top, #062413 0%, #031109 100%);
      --body-bg: #031109;
      --card-bg: rgba(6, 36, 19, 0.9);
      --card-border: rgba(34, 197, 94, 0.3);
      --card-shadow: 0 10px 30px rgba(0, 0, 0, 0.6);
      --inner-bg: rgba(4, 25, 13, 0.75);
      --inner-border: rgba(34, 197, 94, 0.2);
      
      --text-main: #dcfce7;
      --text-sub: #bbf7d0;
      --text-muted: #86efac;
      --text-light: #4ade80;
      
      --primary: #22c55e;
      --primary-hover: #16a34a;
      --accent-green: #22c55e;
      --accent-red: #f87171;
      --accent-blue: #38bdf8;
      
      --card-submitted-bg: rgba(34, 197, 94, 0.25);
      --card-submitted-border: #22c55e;
      --card-submitted-text: #dcfce7;
      --card-submitted-tag: #4ade80;
      
      --card-absent-bg: rgba(239, 68, 68, 0.22);
      --card-absent-border: #ef4444;
      --card-absent-text: #fca5a5;
      --card-absent-tag: #f87171;
      
      --card-pending-bg: rgba(255, 255, 255, 0.04);
      --card-pending-border: rgba(34, 197, 94, 0.2);
      --card-pending-text: #86efac;
      --card-pending-tag: #4ade80;

      --theme-toggle-bg: #062413;
      --theme-toggle-border: #16a34a;
      --theme-toggle-text: #dcfce7;
    }

    /* 4. 大廳二 ‧ 星空魅紫主題 (深色星艦) */
    [data-hall-theme="purple"][data-color-mode="dark"],
    [data-hall-theme="pink"][data-color-mode="dark"] {
      --bg-page: radial-gradient(circle at top, #1e1035 0%, #0c0517 100%);
      --body-bg: #0c0517;
      --card-bg: rgba(30, 16, 53, 0.9);
      --card-border: rgba(167, 139, 250, 0.35);
      --card-shadow: 0 10px 30px rgba(0, 0, 0, 0.6);
      --inner-bg: rgba(20, 10, 38, 0.75);
      --inner-border: rgba(167, 139, 250, 0.25);
      
      --text-main: #ede9fe;
      --text-sub: #ddd6fe;
      --text-muted: #c4b5fd;
      --text-light: #a78bfa;
      
      --primary: #a78bfa;
      --primary-hover: #8b5cf6;
      --accent-green: #a78bfa;
      --accent-red: #f87171;
      --accent-blue: #38bdf8;
      
      --card-submitted-bg: rgba(139, 92, 246, 0.25);
      --card-submitted-border: #a78bfa;
      --card-submitted-text: #ede9fe;
      --card-submitted-tag: #c4b5fd;
      
      --card-absent-bg: rgba(239, 68, 68, 0.22);
      --card-absent-border: #ef4444;
      --card-absent-text: #fca5a5;
      --card-absent-tag: #f87171;
      
      --card-pending-bg: rgba(255, 255, 255, 0.04);
      --card-pending-border: rgba(167, 139, 250, 0.2);
      --card-pending-text: #c4b5fd;
      --card-pending-tag: #a78bfa;

      --theme-toggle-bg: #1e1035;
      --theme-toggle-border: #7c3aed;
      --theme-toggle-text: #ede9fe;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background: var(--bg-page);
      background-color: var(--body-bg);
      color: var(--text-main);
      padding: 16px 20px 40px;
      min-height: 100vh;
      transition: background 0.3s ease, color 0.3s ease;
    }

    .container {
      max-width: 1260px;
      margin: 0 auto;
    }

    /* 頂部列 */
    .header-bar {
      background: var(--card-bg);
      border: 1.5px solid var(--card-border);
      border-radius: 18px;
      padding: 16px 22px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 14px;
      box-shadow: var(--card-shadow);
      margin-bottom: 20px;
    }

    .header-title h1 {
      font-size: 1.55rem;
      font-weight: 900;
      color: var(--text-main);
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .header-controls {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 10px;
    }

    .select-group {
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .control-label {
      font-size: 0.88rem;
      font-weight: 800;
      color: var(--text-main);
      white-space: nowrap;
    }

    .select-hall, .select-unit, .select-class {
      background: var(--inner-bg);
      border: 2px solid var(--primary);
      color: var(--text-main);
      font-size: 0.95rem;
      font-weight: 800;
      padding: 8px 12px;
      border-radius: 12px;
      cursor: pointer;
      outline: none;
      transition: all 0.2s ease;
    }
    .select-hall { font-weight: 900; }
    .select-unit { max-width: 280px; }

    .seat-count-box {
      display: flex;
      align-items: center;
      gap: 4px;
      font-size: 0.88rem;
      font-weight: 800;
      color: var(--text-main);
    }

    .input-seat-count {
      width: 58px;
      padding: 7px 6px;
      text-align: center;
      border: 2px solid var(--card-border);
      border-radius: 10px;
      background: var(--inner-bg);
      color: var(--text-main);
      font-size: 0.95rem;
      font-weight: 800;
      outline: none;
    }

    .btn-control {
      background: var(--theme-toggle-bg);
      border: 1.5px solid var(--theme-toggle-border);
      border-radius: 12px;
      padding: 8px 14px;
      font-size: 0.88rem;
      font-weight: 800;
      color: var(--text-main);
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s ease;
      user-select: none;
    }
    .btn-control:hover {
      transform: translateY(-1px);
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }

    .btn-hall-switch {
      background: linear-gradient(135deg, #16a34a, #db2777);
      color: #ffffff;
      border: none;
      box-shadow: 0 4px 14px rgba(219, 39, 119, 0.25);
    }

    .btn-refresh {
      background: linear-gradient(135deg, var(--primary), var(--primary-hover));
      color: #ffffff;
      border: none;
      box-shadow: 0 4px 14px rgba(0, 0, 0, 0.15);
    }

    /* 繳交率進度面板 */
    .progress-panel {
      background: var(--card-bg);
      border: 1.5px solid var(--card-border);
      border-radius: 18px;
      padding: 20px 24px;
      box-shadow: var(--card-shadow);
      margin-bottom: 22px;
    }

    .progress-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 12px;
    }

    .progress-title {
      font-size: 1.25rem;
      font-weight: 900;
      color: var(--text-main);
    }

    .progress-stats {
      font-size: 1.2rem;
      font-weight: 900;
      color: var(--text-main);
    }

    .progress-stats .stat-done { color: var(--accent-green); }
    .progress-stats .stat-req { color: var(--primary); }
    .progress-stats .stat-pct { color: var(--primary-hover); font-size: 1.25rem; }

    .progress-bar-bg {
      width: 100%;
      height: 18px;
      background: var(--inner-bg);
      border: 1.5px solid var(--inner-border);
      border-radius: 10px;
      overflow: hidden;
      margin-bottom: 16px;
    }

    .progress-bar-fill {
      height: 100%;
      width: 0%;
      background: linear-gradient(90deg, var(--primary), var(--primary-hover));
      border-radius: 10px;
      transition: width 0.6s cubic-bezier(0.4, 0, 0.2, 1);
    }

    /* 請假與操作按鈕列 */
    .action-subbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 14px;
      background: var(--inner-bg);
      border: 1px solid var(--inner-border);
      padding: 10px 16px;
      border-radius: 14px;
    }

    .absent-input-group {
      display: flex;
      align-items: center;
      gap: 8px;
      flex: 1;
      min-width: 320px;
    }

    .absent-label {
      font-size: 0.88rem;
      font-weight: 800;
      color: var(--text-main);
      white-space: nowrap;
    }

    .input-absent {
      flex: 1;
      padding: 7px 12px;
      border: 1.5px solid var(--card-border);
      border-radius: 10px;
      background: var(--card-bg);
      color: var(--text-main);
      font-size: 0.9rem;
      outline: none;
      transition: border-color 0.2s ease;
    }
    .input-absent:focus {
      border-color: var(--primary);
    }

    .absent-note {
      font-size: 0.8rem;
      color: var(--text-muted);
      white-space: nowrap;
    }

    .tool-btn-group {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }

    .btn-act {
      border: none;
      border-radius: 10px;
      padding: 8px 14px;
      font-size: 0.85rem;
      font-weight: 800;
      color: #ffffff;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 5px;
      transition: all 0.2s ease;
      box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    .btn-act:hover { transform: translateY(-1px); }

    .btn-act-green { background: #16a34a; }
    .btn-act-green:hover { background: #15803d; }
    .btn-act-amber { background: #d97706; }
    .btn-act-amber:hover { background: #b45309; }
    .btn-act-teal { background: #0d9488; }
    .btn-act-teal:hover { background: #0f766e; }

    /* 主區域雙欄格 */
    .main-grid {
      display: grid;
      grid-template-columns: 360px 1fr;
      gap: 22px;
      align-items: start;
    }

    @media (max-width: 980px) {
      .main-grid { grid-template-columns: 1fr; }
    }

    .panel-card {
      background: var(--card-bg);
      border: 1.5px solid var(--card-border);
      border-radius: 18px;
      padding: 20px;
      box-shadow: var(--card-shadow);
    }

    .panel-card h2 {
      font-size: 1.18rem;
      font-weight: 900;
      color: var(--text-main);
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 16px;
      border-bottom: 1.5px dashed var(--inner-border);
      padding-bottom: 10px;
    }

    .panel-card h2 .sub-tip {
      font-size: 0.8rem;
      font-weight: 700;
      color: var(--text-muted);
    }

    /* 榮譽榜 Top 5 */
    .top5-list {
      display: flex;
      flex-direction: column;
      gap: 10px;
    }

    .top5-item {
      background: var(--inner-bg);
      border: 1px solid var(--inner-border);
      border-radius: 14px;
      padding: 12px 14px;
      display: flex;
      align-items: center;
      gap: 12px;
      transition: transform 0.2s ease;
    }
    .top5-item:hover {
      transform: translateX(4px);
    }

    .top5-badge {
      width: 38px;
      height: 38px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.25rem;
      font-weight: 900;
      flex-shrink: 0;
    }
    .badge-gold { background: linear-gradient(135deg, #fef08a, #f59e0b); color: #78350f; }
    .badge-silver { background: linear-gradient(135deg, #f1f5f9, #94a3b8); color: #1e293b; }
    .badge-bronze { background: linear-gradient(135deg, #fed7aa, #ea580c); color: #7c2d12; }
    .badge-star { background: var(--inner-bg); border: 1.5px solid var(--card-border); color: var(--primary); font-size: 1rem; }

    .top5-info {
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 2px;
      overflow: hidden;
    }

    .top5-name-row {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .top5-name {
      font-size: 1.02rem;
      font-weight: 900;
      color: var(--text-main);
    }

    .top5-seat {
      font-size: 0.82rem;
      font-weight: 800;
      color: var(--text-muted);
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      padding: 1px 6px;
      border-radius: 6px;
    }

    .top5-units {
      font-size: 0.78rem;
      font-weight: 700;
      color: var(--text-muted);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .top5-score-box {
      text-align: right;
    }

    .top5-score {
      font-size: 1.25rem;
      font-weight: 900;
      color: var(--accent-green);
    }

    .top5-score-sub {
      font-size: 0.72rem;
      font-weight: 700;
      color: var(--text-muted);
    }

    .empty-state {
      text-align: center;
      padding: 36px 16px;
      color: var(--text-muted);
      font-size: 0.95rem;
      font-weight: 700;
    }

    /* 座號卡片網格 (點名矩陣) */
    .seat-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(105px, 1fr));
      gap: 12px;
    }

    .seat-card {
      border-radius: 14px;
      padding: 12px 8px;
      text-align: center;
      cursor: pointer;
      user-select: none;
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      position: relative;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      min-height: 80px;
    }
    .seat-card:hover {
      transform: translateY(-3px);
      box-shadow: 0 6px 16px rgba(0, 0, 0, 0.08);
    }

    /* 狀態 A: 已完成 */
    .seat-card.submitted {
      background: var(--card-submitted-bg);
      border: 1.5px solid var(--card-submitted-border);
      color: var(--card-submitted-text);
    }
    .seat-card.submitted .card-seat-num {
      color: var(--card-submitted-tag);
      font-weight: 900;
    }

    /* 狀態 B: 請假 */
    .seat-card.absent {
      background: var(--card-absent-bg);
      border: 1.5px solid var(--card-absent-border);
      color: var(--card-absent-text);
    }
    .seat-card.absent .card-seat-num {
      color: var(--card-absent-tag);
      font-weight: 900;
    }

    /* 狀態 C: 未交 (加油中) */
    .seat-card.pending {
      background: var(--card-pending-bg);
      border: 1.5px dashed var(--card-pending-border);
      color: var(--card-pending-text);
    }
    .seat-card.pending .card-seat-num {
      color: var(--card-pending-tag);
      font-weight: 900;
    }

    .card-top-icon {
      font-size: 1.05rem;
      margin-bottom: 2px;
    }

    .card-seat-num {
      font-size: 1.05rem;
      letter-spacing: -0.3px;
    }

    .card-status-text {
      font-size: 0.8rem;
      font-weight: 800;
      margin-top: 3px;
    }

    .card-student-name {
      font-size: 0.82rem;
      font-weight: 900;
      color: var(--text-main);
      margin-top: 2px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 90px;
    }

    .card-score-badge {
      display: none;
      font-size: 0.76rem;
      font-weight: 900;
      color: var(--primary);
      margin-top: 2px;
    }

    /* 導師分數切換顯示時 */
    body.show-teacher-scores .card-score-badge {
      display: block;
    }

    /* 彈窗樣式 (手動補登成績) */
    .modal-overlay {
      position: fixed;
      inset: 0;
      background: rgba(0, 0, 0, 0.55);
      backdrop-filter: blur(4px);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 9999;
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.25s ease;
    }
    .modal-overlay.active {
      opacity: 1;
      pointer-events: auto;
    }

    .modal-box {
      background: var(--card-bg);
      border: 2px solid var(--primary);
      border-radius: 20px;
      width: 90%;
      max-width: 460px;
      padding: 24px 28px;
      box-shadow: 0 20px 40px rgba(0, 0, 0, 0.25);
      transform: translateY(20px);
      transition: transform 0.25s ease;
    }
    .modal-overlay.active .modal-box {
      transform: translateY(0);
    }

    .modal-title {
      font-size: 1.35rem;
      font-weight: 900;
      color: var(--text-main);
      margin-bottom: 16px;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .modal-field {
      margin-bottom: 14px;
    }

    .modal-field label {
      display: block;
      font-size: 0.88rem;
      font-weight: 800;
      color: var(--text-main);
      margin-bottom: 5px;
    }

    .modal-input {
      width: 100%;
      padding: 9px 12px;
      border: 1.5px solid var(--card-border);
      border-radius: 10px;
      background: var(--inner-bg);
      color: var(--text-main);
      font-size: 0.95rem;
      font-weight: 700;
      outline: none;
    }
    .modal-input:focus {
      border-color: var(--primary);
    }

    .modal-actions {
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      margin-top: 20px;
    }

    /* 連線提示通知條 (Toast) */
    .toast-msg {
      position: fixed;
      bottom: 24px;
      right: 24px;
      background: #1e293b;
      color: #ffffff;
      padding: 12px 20px;
      border-radius: 12px;
      font-size: 0.92rem;
      font-weight: 800;
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
      transform: translateY(100px);
      opacity: 0;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      z-index: 10000;
    }
    .toast-msg.show {
      transform: translateY(0);
      opacity: 1;
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- 頂部導航與總控列 -->
    <header class="header-bar">
      <div class="header-title">
        <h1 id="dashboardMainTitle">🌿 大廳一：互動遊戲即時儀表板</h1>
        <p id="dashboardSubTitle">共用題庫遊戲 ‧ 即時大螢幕點名 ‧ 雙大廳色系聯動</p>
      </div>
      <div class="header-controls">
        <!-- 第一層：大廳切換 (大廳一 vs 大廳二) -->
        <div class="select-group">
          <label class="control-label" for="hallSelect">🏛️ 目標大廳：</label>
          <select class="select-hall" id="hallSelect" onchange="onHallChange()">
            <option value="hall1">🌿 大廳一：共用題庫遊戲 (綠)</option>
            <option value="hall2">🔮 大廳二：獨立單元遊戲 (紫)</option>
          </select>
        </div>

        <!-- 第二層：該大廳底下的遊戲主題下拉選單 -->
        <div class="select-group">
          <label class="control-label" for="unitSelect">🎮 遊戲主題：</label>
          <select class="select-unit" id="unitSelect" onchange="onUnitChange()">
            <option value="ALL">🌐 全遊戲彙整總覽</option>
          </select>
        </div>

        <!-- 第三層：班級下拉選單 -->
        <div class="select-group">
          <label class="control-label" for="classSelect">🏫 班級：</label>
          <select class="select-class" id="classSelect" onchange="onClassChange()">
            <option value="701">701 班</option>
            <option value="702">702 班</option>
            <option value="703">703 班</option>
            <option value="704">704 班</option>
            <option value="705">705 班</option>
            <option value="706">706 班</option>
            <option value="707">707 班</option>
            <option value="708">708 班</option>
            <option value="709">709 班</option>
            <option value="710" selected>710 班</option>
            <option value="七甲">七甲</option>
            <option value="七乙">七乙</option>
            <option value="七丙">七丙</option>
          </select>
        </div>

        <div class="seat-count-box">
          <span>應到：</span>
          <input type="number" class="input-seat-count" id="totalSeatsInput" value="30" min="1" max="50" onchange="saveClassConfig(); renderDashboard();" />
          <span>人</span>
        </div>

        <!-- 🔮 切換大廳色 (綠/紫) 快捷按鈕 -->
        <button type="button" class="btn-control btn-hall-switch" id="btnHallColorToggle" onclick="toggleHallColor()">
          🔮 切換大廳色 (綠/紫)
        </button>

        <!-- 🌓 明暗模式切換按鈕 -->
        <button type="button" class="btn-control" id="btnThemeToggle" onclick="toggleColorMode()">
          <span id="themeToggleIcon">🌙</span> <span id="themeToggleText">深色模式</span>
        </button>

        <button type="button" class="btn-control" id="btnScoreToggle" onclick="toggleTeacherScore()">
          👁️ 導師分數切換
        </button>

        <button type="button" class="btn-control btn-refresh" id="btnRefresh" onclick="fetchLiveScores()">
          🔄 刷新資料
        </button>
      </div>
    </header>

    <!-- 測驗繳交參與率統計面板 -->
    <section class="progress-panel">
      <div class="progress-header">
        <div class="progress-title" id="progressClassTitle">710 班 測驗繳交參與率</div>
        <div class="progress-stats">
          已交：<span class="stat-done" id="submittedCount">0</span> / 
          實到應交：<span class="stat-req" id="actualRequiredCount">30</span> 人 
          (<span class="stat-pct" id="progressPercent">0%</span>)
        </div>
      </div>

      <div class="progress-bar-bg">
        <div class="progress-bar-fill" id="progressBarFill"></div>
      </div>

      <!-- 請假輸入列與快速動作按鈕 -->
      <div class="action-subbar">
        <div class="absent-input-group">
          <label class="absent-label">🏖️ 本班今日請假座號：</label>
          <input type="text" class="input-absent" id="absentSeatsInput" placeholder="請輸入請假座號 (例如：05, 12)，或直接點擊下方座號卡" onchange="saveClassConfig(); renderDashboard();" />
          <span class="absent-note">(將自動扣除應交人數)</span>
        </div>

        <div class="tool-btn-group">
          <button type="button" class="btn-act btn-act-green" onclick="openManualGradeModal()">
            ✍️ 手動補登成績
          </button>
          <button type="button" class="btn-act btn-act-amber" onclick="copyUnsubmittedList()">
            📋 複製未交名單
          </button>
          <button type="button" class="btn-act btn-act-teal" onclick="exportScoresCsv()">
            📥 匯出CSV
          </button>
        </div>
      </div>
    </section>

    <!-- 雙欄主面板 -->
    <main class="main-grid">
      <!-- 左欄：測驗榮譽英雄榜 Top 5 -->
      <section class="panel-card">
        <h2>
          <span>🎉 測驗榮譽英雄榜 (Top 5)</span>
          <span class="sub-tip" id="rankUnitFilterTip">全班實力競技</span>
        </h2>
        <div class="top5-list" id="top5List">
          <div class="empty-state">尚無測驗紀錄</div>
        </div>
      </section>

      <!-- 右欄：座號完成狀態圖 -->
      <section class="panel-card">
        <h2>
          <span>🔢 全班完成狀態圖 (點座號卡切換請假 / 查看成績)</span>
          <span class="sub-tip">未交點擊切換請假 · 已交點擊看明細</span>
        </h2>
        <div class="seat-grid" id="seatGrid"></div>
      </section>
    </main>
  </div>

  <!-- 手動補登成績彈窗 -->
  <div class="modal-overlay" id="manualModal">
    <div class="modal-box">
      <div class="modal-title">✍️ 課堂手動補登成績</div>
      <div class="modal-field">
        <label>座號：</label>
        <input type="number" class="modal-input" id="manualSeat" min="1" max="50" placeholder="例如：08 或 8" />
      </div>
      <div class="modal-field">
        <label>學生姓名：</label>
        <input type="text" class="modal-input" id="manualName" placeholder="例如：林小明" />
      </div>
      <div class="modal-field">
        <label>得分：</label>
        <input type="number" class="modal-input" id="manualScore" min="0" max="100" value="100" />
      </div>
      <div class="modal-field">
        <label>遊戲主題 / 任務名稱：</label>
        <input type="text" class="modal-input" id="manualRemark" value="課堂紙筆測驗 / 實體補登" />
      </div>
      <div class="modal-actions">
        <button type="button" class="btn-control" onclick="closeManualGradeModal()">取消</button>
        <button type="button" class="btn-act btn-act-green" onclick="submitManualGrade()">確認補登並點燈 ✅</button>
      </div>
    </div>
  </div>

  <!-- Toast 訊息提示 -->
  <div class="toast-msg" id="toastMsg"></div>

  <script>
    let rawRecords = [];
    let isTeacherScoreShown = false;
    let currentHall = localStorage.getItem("bio_dash_selected_hall") || "hall1"; // 'hall1' (綠) | 'hall2' (粉)
    let currentColorMode = localStorage.getItem("bio_dash_color_mode") || "light"; // 'light' | 'dark'

    // 大廳一 17 款共用題庫遊戲清單
    const HALL1_GAMES = [
      "🏰 皇家古堡：生物密室逃脫",
      "🔬 微觀特攻隊：細胞世界大暴走",
      "🐇 奇幻仙境：愛麗絲生物密室逃脫",
      "🏫 現代校園：生物密室逃脫",
      "🐑 羊羊農場：生物大逃脫",
      "🫘 淘氣豆：生物敲敲研究所",
      "🐝 生物小蜜蜂：星際防衛隊",
      "🚀 生物星際大逃亡：逃離黑洞",
      "👾 生物小精靈塔防：微觀免疫大作戰",
      "🧬 生物大墜落：細胞深淵下樓梯",
      "🎣 生物大釣手：黃金深海大捕撈",
      "🌱 植物大戰殭屍：細胞防衛戰",
      "🗡️ 樂透大冒險：生物勇者傳奇",
      "🍄 超級馬力歐生物闖關",
      "🐰 動物方城市生物大富翁",
      "🧬 生物能量祖瑪：光合呼吸大冒險",
      "🦖 侏羅紀叢林逃生記"
    ];

    // 大廳二 3 款獨立單元遊戲清單
    const HALL2_GAMES = [
      "💩 人體消化大冒險 (54格完備版)",
      "🧋 珍珠奶茶大對決 (喝光奶茶擂台版)",
      "🔬 五界生物研究餐廳 (48種樣本・迴轉研究大餐)"
    ];

    // 0. 主題樣式應用與切換
    function applyThemes() {
      const root = document.documentElement;
      root.setAttribute("data-hall-theme", currentHall === "hall2" ? "pink" : "green");
      root.setAttribute("data-color-mode", currentColorMode);

      const titleEl = document.getElementById("dashboardMainTitle");
      const subTitleEl = document.getElementById("dashboardSubTitle");
      const hallSelect = document.getElementById("hallSelect");
      if (hallSelect) hallSelect.value = currentHall;

      if (currentHall === "hall2") {
        if (titleEl) titleEl.textContent = "🌸 大廳二：獨立遊戲即時儀表板";
        if (subTitleEl) subTitleEl.textContent = "獨立單元遊戲 ‧ 即時大螢幕點名 ‧ 專屬單元題庫";
      } else {
        if (titleEl) titleEl.textContent = "🌿 大廳一：互動遊戲即時儀表板";
        if (subTitleEl) subTitleEl.textContent = "共用題庫遊戲 ‧ 即時大螢幕點名 ‧ 雙大廳色系聯動";
      }

      const icon = document.getElementById("themeToggleIcon");
      const text = document.getElementById("themeToggleText");
      if (currentColorMode === "dark") {
        if (icon) icon.textContent = "☀️";
        if (text) text.textContent = "明亮模式";
      } else {
        if (icon) icon.textContent = "🌙";
        if (text) text.textContent = "深色模式";
      }
    }

    // 🔮 切換大廳色快捷鍵 (綠/紫一鍵切換)
    function toggleHallColor() {
      currentHall = (currentHall === "hall1") ? "hall2" : "hall1";
      localStorage.setItem("bio_dash_selected_hall", currentHall);
      applyThemes();
      updateUnitDropdown();
      renderDashboard();
      showToast(currentHall === "hall2" ? "🔮 已切換為【大廳二：星空紫】主題" : "🌿 已切換為【大廳一：森林綠】主題");
    }

    // 🌓 明暗色彩模式切換
    function toggleColorMode() {
      currentColorMode = (currentColorMode === "dark") ? "light" : "dark";
      localStorage.setItem("bio_dash_color_mode", currentColorMode);
      applyThemes();
      showToast(currentColorMode === "dark" ? "🌙 已啟用深色夜間模式" : "☀️ 已啟用明亮色彩模式");
    }

    // 第一層：手動改變大廳下拉
    function onHallChange() {
      const sel = document.getElementById("hallSelect");
      currentHall = sel.value;
      localStorage.setItem("bio_dash_selected_hall", currentHall);
      applyThemes();
      updateUnitDropdown();
      renderDashboard();
    }

    // 第二層：動態根據目前大廳填入專屬遊戲清單
    function updateUnitDropdown() {
      const unitSelect = document.getElementById("unitSelect");
      if (!unitSelect) return;

      const currentVal = unitSelect.value || "ALL";
      unitSelect.innerHTML = '<option value="ALL">🌐 全遊戲彙整總覽</option>';

      const targetList = (currentHall === "hall2") ? HALL2_GAMES : HALL1_GAMES;
      const groupLabel = (currentHall === "hall2") ? "🧩 大廳二：獨立遊戲 (3款)" : "🎮 大廳一：共用題庫遊戲 (17款)";

      const optGroup = document.createElement("optgroup");
      optGroup.label = "── " + groupLabel + " ──";
      targetList.forEach(g => {
        const opt = document.createElement("option");
        opt.value = g;
        opt.textContent = g;
        optGroup.appendChild(opt);
      });
      unitSelect.appendChild(optGroup);

      // 檢查試算表中是否有額外單元分頁
      const classRegex = /^(7\\d{2}|七[甲乙丙]|\\d+班|[一二三四五六七八九]班)/;
      const extraSet = new Set();
      rawRecords.forEach(r => {
        const u = String(r.units || "").trim();
        if (u && !classRegex.test(u) && !u.includes("班級") && !u.includes("座號") && u !== "ALL") {
          if (!HALL1_GAMES.some(g => g.includes(u)) && !HALL2_GAMES.some(g => g.includes(u))) {
            extraSet.add(u);
          }
        }
      });

      if (extraSet.size > 0) {
        const optGroupExtra = document.createElement("optgroup");
        optGroupExtra.label = "── 📚 試算表其他測驗分頁 ──";
        extraSet.forEach(u => {
          const opt = document.createElement("option");
          opt.value = u;
          opt.textContent = "📚 " + u;
          optGroupExtra.appendChild(opt);
        });
        unitSelect.appendChild(optGroupExtra);
      }

      // 若先前選中的選項在新清單中存在則保留，否則預設回 ALL
      const exists = Array.from(unitSelect.options).some(o => o.value === currentVal);
      unitSelect.value = exists ? currentVal : "ALL";
    }

    function onUnitChange() {
      const unitSelect = document.getElementById("unitSelect");
      const val = unitSelect.value;
      const rankTip = document.getElementById("rankUnitFilterTip");
      if (rankTip) {
        rankTip.textContent = (val === "ALL") ? "全班實力競技" : val;
      }
      renderDashboard();
    }

    // 1. 班級字串統一正規化對照函式 (相容 "710", "7年10班", "七年十班", "710班", "7-10")
    function normalizeClassStr(input) {
      if (!input) return "";
      const str = String(input).trim().replace(/班$/, "");
      const matchSpecial = str.match(/(?:7|七)[\\s-]*([甲乙丙])/);
      if (matchSpecial) return "七" + matchSpecial[1];
      if (/^[甲乙丙]$/.test(str)) return "七" + str;

      const digits = str.match(/\\d+/g);
      if (digits) {
        if (digits.length >= 2) {
          return digits[0] + digits[1].padStart(2, "0");
        } else if (digits.length === 1) {
          const d = digits[0];
          if (d.length === 3) return d;
          if (d.length <= 2) return "7" + d.padStart(2, "0");
        }
      }
      if (str.includes("十")) return "710";
      const cnMap = {"一":"01","二":"02","三":"03","四":"04","五":"05","六":"06","七":"07","八":"08","九":"09"};
      for (const [cn, num] of Object.entries(cnMap)) {
        if (str.includes(cn)) return "7" + num;
      }
      return str;
    }

    // 2. 座號數字純化函式 (例如 "35號" -> 35, "05號" -> 5)
    function normalizeSeatNum(seat) {
      if (seat === null || seat === undefined || seat === "") return null;
      const match = String(seat).match(/\\d+/);
      return match ? parseInt(match[0], 10) : null;
    }

    // 3. 連線抓取 Google 試算表最新成績
    function fetchLiveScores() {
      const btn = document.getElementById("btnRefresh");
      btn.textContent = "⏳ 刷新中...";
      btn.disabled = true;

      if (typeof google !== "undefined" && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(function(resp) {
            btn.textContent = "🔄 刷新資料";
            btn.disabled = false;
            if (resp && resp.records) {
              rawRecords = resp.records;
              showToast("✅ 已同步最新資料 (" + rawRecords.length + " 筆紀錄)");
              updateUnitDropdown();
              renderDashboard();
            }
          })
          .withFailureHandler(function(err) {
            btn.textContent = "🔄 刷新資料";
            btn.disabled = false;
            console.error("fetchLiveScores error:", err);
            showToast("⚠️ 取得資料發生問題：" + err);
          })
          .getScoresData();
      } else {
        btn.textContent = "🔄 刷新資料";
        btn.disabled = false;
      }
    }

    function isUnitMatching(recordUnit, targetUnit) {
      if (!targetUnit || targetUnit === 'ALL') return true;
      const u = String(recordUnit || '').trim();
      const t = String(targetUnit || '').trim();
      if (u === t || u.includes(t) || t.includes(u)) return true;
      // 萃取中文核心名稱比對（排除 Emoji 與副標題冒號前後）
      const cleanCore = (s) => s.replace(/^[\\p{Emoji}\\s]+/u, '').split(/[:：]/)[0].trim();
      const uCore = cleanCore(u);
      const tCore = cleanCore(t);
      if (uCore && tCore && (uCore.includes(tCore) || tCore.includes(uCore))) return true;
      // 若學生記錄在通用『測驗成績紀錄』分頁，且該遊戲也是近期測驗，寬容採計
      if (u === '測驗成績紀錄' || u === '自訂範圍') return true;
      return false;
    }

    // 4. 渲染儀表板所有卡片與統計數字
    function renderDashboard() {
      const selectedClassRaw = document.getElementById("classSelect").value;
      const targetClassNorm = normalizeClassStr(selectedClassRaw);
      const unitSelect = document.getElementById("unitSelect");
      const selectedUnit = unitSelect ? unitSelect.value : "ALL";

      const unitTitleStr = (selectedUnit && selectedUnit !== "ALL") ? ("【" + selectedUnit + "】") : "";
      document.getElementById("progressClassTitle").textContent = selectedClassRaw + " 班 " + unitTitleStr + " 測驗繳交參與率";

      // 依班級正規化過濾
      let classRecords = rawRecords.filter(r => normalizeClassStr(r.className) === targetClassNorm);

      // 若有選擇指定遊戲主題，精準過濾
      if (selectedUnit && selectedUnit !== "ALL") {
        classRecords = classRecords.filter(r => isUnitMatching(r.units, selectedUnit));
      } else {
        // 若選 ALL，保留此班所有紀錄或當前大廳紀錄
        const targetList = (currentHall === "hall2") ? HALL2_GAMES : HALL1_GAMES;
        classRecords = classRecords.filter(r => {
          const u = String(r.units || "").trim();
          if (u === '測驗成績紀錄' || u === '自訂範圍') return true;
          return targetList.some(g => isUnitMatching(u, g));
        });
      }

      // 取得請假座號集合
      const absentInput = document.getElementById("absentSeatsInput").value;
      const absentSet = new Set(
        absentInput.split(/[,，\\s]+/)
          .map(s => normalizeSeatNum(s))
          .filter(n => n !== null)
      );

      // 整理此班學生已交記錄 (以座號為 Key，取最高分)
      const studentMap = new Map();
      let maxSeatFound = 0;

      classRecords.forEach(r => {
        const seatNum = normalizeSeatNum(r.seat);
        if (seatNum !== null) {
          if (seatNum > maxSeatFound) maxSeatFound = seatNum;
          const prev = studentMap.get(seatNum);
          if (!prev || (Number(r.finalScore) > Number(prev.finalScore))) {
            studentMap.set(seatNum, r);
          }
        }
      });

      // 1. 渲染測驗榮譽英雄榜 (Top 5)
      const top5List = document.getElementById("top5List");
      const sortedStudents = Array.from(studentMap.values()).sort((a, b) => Number(b.finalScore) - Number(a.finalScore));
      const top5 = sortedStudents.slice(0, 5);

      if (top5.length === 0) {
        top5List.innerHTML = '<div class="empty-state">尚無測驗紀錄</div>';
      } else {
        const medals = ["🥇", "🥈", "🥉", "✨", "🌟"];
        const badgeClasses = ["badge-gold", "badge-silver", "badge-bronze", "badge-star", "badge-star"];

        top5List.innerHTML = top5.map((r, idx) => {
          const seatFormatted = String(normalizeSeatNum(r.seat) || "").padStart(2, "0") + " 號";
          return '<div class="top5-item">' +
            '<div class="top5-badge ' + badgeClasses[idx] + '">' + medals[idx] + '</div>' +
            '<div class="top5-info">' +
              '<div class="top5-name-row">' +
                '<span class="top5-name">' + escapeHtml(r.name || "學生") + '</span>' +
                '<span class="top5-seat">' + seatFormatted + '</span>' +
              '</div>' +
              '<div class="top5-units" title="' + escapeHtml(r.units || '') + '">' + escapeHtml(r.units || '課堂測驗') + '</div>' +
            '</div>' +
            '<div class="top5-score-box">' +
              '<div class="top5-score">' + r.finalScore + ' 分</div>' +
              '<div class="top5-score-sub">' + (r.errorCount > 0 ? ('錯 ' + r.errorCount + ' 題') : '全對滿分 💯') + '</div>' +
            '</div>' +
          '</div>';
        }).join("");
      }

      // 2. 渲染全班完成狀態圖 (座號卡片)
      const configuredTotal = parseInt(document.getElementById("totalSeatsInput").value, 10) || 30;
      const renderLimit = Math.max(configuredTotal, maxSeatFound);

      const grid = document.getElementById("seatGrid");
      grid.innerHTML = "";

      let submittedCount = 0;
      let absentCount = 0;

      for (let seat = 1; seat <= renderLimit; seat++) {
        const seatFormatted = String(seat).padStart(2, "0");
        const isAbsent = absentSet.has(seat);
        const record = studentMap.get(seat);
        const isSubmitted = !!record;

        let cardClass = "pending";
        let statusText = "加油中 💪";
        let icon = "🍰";
        let studentNameHtml = "";
        let scoreBadgeHtml = "";

        if (isAbsent) {
          absentCount++;
          cardClass = "absent";
          statusText = "請假 🏖️";
          icon = "🏖️";
        } else if (isSubmitted) {
          submittedCount++;
          cardClass = "submitted";
          statusText = "✅ 已完成";
          icon = "🎉";
          studentNameHtml = '<div class="card-student-name">' + escapeHtml(record.name || '學生') + '</div>';
          scoreBadgeHtml = '<div class="card-score-badge">' + record.finalScore + ' 分</div>';
        }

        grid.innerHTML += '<div class="seat-card ' + cardClass + '" onclick="handleSeatCardClick(' + seat + ', ' + isSubmitted + ', ' + isAbsent + ')" title="座號 ' + seatFormatted + ' 號 (點擊切換請假 / 查看成績)">' +
          '<div class="card-top-icon">' + icon + '</div>' +
          '<div class="card-seat-num">' + seatFormatted + ' 號</div>' +
          '<div class="card-status-text">' + statusText + '</div>' +
          studentNameHtml +
          scoreBadgeHtml +
        '</div>';
      }

      // 3. 計算實到應交與繳交率
      const actualRequired = Math.max(1, configuredTotal - absentCount);
      const percent = Math.min(100, Math.round((submittedCount / actualRequired) * 100));

      document.getElementById("submittedCount").textContent = submittedCount;
      document.getElementById("actualRequiredCount").textContent = actualRequired;
      document.getElementById("progressPercent").textContent = percent + "%";
      document.getElementById("progressBarFill").style.width = percent + "%";

      if (percent >= 100 && submittedCount >= actualRequired) {
        document.getElementById("progressPercent").textContent = "100% 🎉 應交全數完成！";
      }
    }

    // 5. 點擊座號卡：未交者點擊切換請假，已交者點擊查看學生詳情
    function handleSeatCardClick(seatNum, isSubmitted, isAbsent) {
      if (isSubmitted) {
        const record = rawRecords.find(r => 
          normalizeClassStr(r.className) === normalizeClassStr(document.getElementById("classSelect").value) &&
          normalizeSeatNum(r.seat) === seatNum
        );
        if (record) {
          const detail = "【" + record.className + "班 " + record.seat + "號 " + record.name + "】\\n" +
                         "• 測驗主題：" + (record.units || '生物測驗') + "\\n" +
                         "• 原始得分：" + record.originalScore + " 分\\n" +
                         "• 採計總分：" + record.finalScore + " 分\\n" +
                         "• 錯題題數：" + record.errorCount + " 題\\n" +
                         "• 訂正備註：" + (record.keywords || '無') + "\\n" +
                         "• 視窗警示：" + (record.switchCount || '正常') + "\\n" +
                         "• 繳交時間：" + (record.timestamp || '');
          alert(detail);
          return;
        }
      }

      // 未交者切換請假標記
      toggleAbsentSeat(seatNum);
    }

    // 6. 切換請假狀態
    function toggleAbsentSeat(seatNum) {
      const input = document.getElementById("absentSeatsInput");
      const currentSet = new Set(
        input.value.split(/[,，\\s]+/)
          .map(s => normalizeSeatNum(s))
          .filter(n => n !== null)
      );

      if (currentSet.has(seatNum)) {
        currentSet.delete(seatNum);
        showToast("已取消 " + String(seatNum).padStart(2, '0') + " 號請假標記");
      } else {
        currentSet.add(seatNum);
        showToast("已標記 " + String(seatNum).padStart(2, '0') + " 號為請假 (自動扣除應交)");
      }

      const sorted = Array.from(currentSet).sort((a, b) => a - b);
      input.value = sorted.map(n => String(n).padStart(2, "0")).join(", ");
      saveClassConfig();
      renderDashboard();
    }

    // 7. 班級組態儲存與讀取 (LocalStorage)
    function onClassChange() {
      loadClassConfig();
      renderDashboard();
    }

    function saveClassConfig() {
      const cls = document.getElementById("classSelect").value;
      const total = document.getElementById("totalSeatsInput").value;
      const absent = document.getElementById("absentSeatsInput").value;
      localStorage.setItem("bio_dash_cfg_" + cls + "_total", total);
      localStorage.setItem("bio_dash_cfg_" + cls + "_absent", absent);
    }

    function loadClassConfig() {
      const cls = document.getElementById("classSelect").value;
      const savedTotal = localStorage.getItem("bio_dash_cfg_" + cls + "_total") || "30";
      const savedAbsent = localStorage.getItem("bio_dash_cfg_" + cls + "_absent") || "";
      document.getElementById("totalSeatsInput").value = savedTotal;
      document.getElementById("absentSeatsInput").value = savedAbsent;
    }

    // 8. 導師分數切換
    function toggleTeacherScore() {
      isTeacherScoreShown = !isTeacherScoreShown;
      const btn = document.getElementById("btnScoreToggle");
      if (isTeacherScoreShown) {
        document.body.classList.add("show-teacher-scores");
        btn.textContent = "🙈 遮蔽隱藏分數";
        showToast("👁️ 導師模式：已顯示各座號成績");
      } else {
        document.body.classList.remove("show-teacher-scores");
        btn.textContent = "👁️ 導師分數切換";
        showToast("🔒 學生隱私保護模式：已遮蔽分數");
      }
    }

    // 9. 複製未交名單到剪貼簿
    function copyUnsubmittedList() {
      const selectedClassRaw = document.getElementById("classSelect").value;
      const targetClassNorm = normalizeClassStr(selectedClassRaw);
      const configuredTotal = parseInt(document.getElementById("totalSeatsInput").value, 10) || 30;

      const absentInput = document.getElementById("absentSeatsInput").value;
      const absentSet = new Set(
        absentInput.split(/[,，\\s]+/)
          .map(s => normalizeSeatNum(s))
          .filter(n => n !== null)
      );

      const submittedSeats = new Set(
        rawRecords
          .filter(r => normalizeClassStr(r.className) === targetClassNorm)
          .map(r => normalizeSeatNum(r.seat))
          .filter(n => n !== null)
      );

      const missingSeats = [];
      for (let seat = 1; seat <= configuredTotal; seat++) {
        if (!absentSet.has(seat) && !submittedSeats.has(seat)) {
          missingSeats.push(String(seat).padStart(2, "0"));
        }
      }

      const absentStr = absentSet.size > 0 
        ? Array.from(absentSet).sort((a,b)=>a-b).map(n => String(n).padStart(2, "0") + "號").join("、")
        : "無";

      const missingStr = missingSeats.length > 0 
        ? missingSeats.map(s => s + "號").join("、") 
        : "全數繳交完成 🎉";

      const text = "【" + selectedClassRaw + " 班 生物測驗繳交進度回報】\\n" +
                   "• 應到人數：" + configuredTotal + " 人 / 實到應交：" + Math.max(0, configuredTotal - absentSet.size) + " 人\\n" +
                   "• 已交人數：" + submittedSeats.size + " 人 (繳交率：" + Math.round((submittedSeats.size / Math.max(1, configuredTotal - absentSet.size)) * 100) + "%)\\n" +
                   "• 今日請假：" + absentStr + "\\n" +
                   "• 尚未交卷座號：" + missingStr + "\\n\\n" +
                   "請尚未繳交的同學確認送出，保護好個人作答成果！💪";

      navigator.clipboard.writeText(text).then(() => {
        showToast("📋 已成功複製未交名單至剪貼簿！");
      }).catch(() => {
        prompt("請手動複製以下未交名單：", text);
      });
    }

    // 10. 匯出 CSV 成績表
    function exportScoresCsv() {
      const selectedClassRaw = document.getElementById("classSelect").value;
      const targetClassNorm = normalizeClassStr(selectedClassRaw);
      const unitSelect = document.getElementById("unitSelect");
      const selectedUnit = unitSelect ? unitSelect.value : "ALL";

      let classRecords = rawRecords.filter(r => normalizeClassStr(r.className) === targetClassNorm);
      if (selectedUnit && selectedUnit !== "ALL") {
        classRecords = classRecords.filter(r => isUnitMatching(r.units, selectedUnit));
      }

      if (classRecords.length === 0) {
        showToast("⚠️ 目前選取的班級與單元尚無成績紀錄可供匯出");
        return;
      }

      let csvContent = "\\uFEFF交卷時間,班級,座號,姓名,測驗主題,原始得分,最終總分,錯題數,訂正心得,跳出次數,狀態備註\\n";
      classRecords.forEach(r => {
        const row = [
          r.timestamp || "",
          r.className || selectedClassRaw,
          r.seat || "",
          r.name || "",
          r.units || "",
          r.originalScore !== undefined ? r.originalScore : "",
          r.finalScore !== undefined ? r.finalScore : "",
          r.errorCount !== undefined ? r.errorCount : 0,
          (r.keywords || "").replace(/[\\n\\r,]/g, " "),
          r.switchCount || 0,
          r.statusNote || "正常"
        ];
        csvContent += row.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(",") + "\\n";
      });

      const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", selectedClassRaw + "班_生物課堂成績報表.csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast("📥 CSV 報表匯出成功！");
    }

    // 11. 手動補登彈窗邏輯
    function openManualGradeModal() {
      const modal = document.getElementById("manualModal");
      modal.classList.add("active");
      const unitSelect = document.getElementById("unitSelect");
      const curUnit = (unitSelect && unitSelect.value !== "ALL") ? unitSelect.value : "課堂實體手動補登";
      document.getElementById("manualRemark").value = curUnit;
      document.getElementById("manualSeat").value = "";
      document.getElementById("manualName").value = "";
      document.getElementById("manualScore").value = "100";
    }

    function closeManualGradeModal() {
      const modal = document.getElementById("manualModal");
      modal.classList.remove("active");
    }

    function submitManualGrade() {
      const cls = document.getElementById("classSelect").value;
      const seat = parseInt(document.getElementById("manualSeat").value, 10);
      const name = document.getElementById("manualName").value.trim();
      const score = parseFloat(document.getElementById("manualScore").value) || 0;
      const remark = document.getElementById("manualRemark").value.trim() || "課堂手動補登";

      if (!seat || seat < 1) {
        alert("請輸入有效的學生座號 (1~50)");
        return;
      }
      if (!name) {
        alert("請輸入學生姓名");
        return;
      }

      const seatFormatted = String(seat).padStart(2, "0");
      if (typeof google !== "undefined" && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(function(res) {
            if (res && res.status === "success") {
              showToast("✅ " + seatFormatted + " 號 " + name + " 成績補登成功！");
              closeManualGradeModal();
              fetchLiveScores();
            } else {
              alert("補登失敗：" + (res ? res.message : "未知錯誤"));
            }
          })
          .withFailureHandler(function(err) {
            alert("補登連線失敗：" + err);
          })
          .saveManualScoreBackend({
            className: cls,
            seat: seatFormatted,
            name: name,
            score: score,
            units: remark,
            note: "老師手動補登"
          });
      }
    }

    // 輔助函式：HTML 跳脫
    function escapeHtml(text) {
      if (!text) return "";
      return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }

    // 輔助函式：Toast 通知
    let toastTimer = null;
    function showToast(msg) {
      const toast = document.getElementById("toastMsg");
      if (!toast) return;
      toast.textContent = msg;
      toast.classList.add("show");
      if (toastTimer) clearTimeout(toastTimer);
      toastTimer = setTimeout(() => {
        toast.classList.remove("show");
      }, 3500);
    }

    // 頁面初始化
    applyThemes();
    loadClassConfig();
    updateUnitDropdown();
    fetchLiveScores();
    setInterval(fetchLiveScores, 15000); // 每 15 秒自動重新整理
  </script>
</body>
</html>
`;
  return HtmlService.createHtmlOutput(html)
    .setTitle("🏆 國中生物測驗與遊戲即時儀表板")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}
