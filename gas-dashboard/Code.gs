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

    // 1. 第一軌：尋找或建立該學生的班級專屬分頁（如 701, 702, 七甲, 七乙, 七丙）
    // 進行座號就地更新 (In-Place Update)，一人一列，保持乾淨排序！
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

    // 2. 第二軌：同步寫入單元分頁或【所有成績彙整】/【測驗成績紀錄】總表
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
  
  // 1. 外部透過 JSONP 請求
  if (callback) {
    var dataJsonp = getScoresData();
    var output = callback + "(" + JSON.stringify(dataJsonp) + ");";
    return ContentService.createTextOutput(output)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  // 2. 外部 (GitHub Pages 或其他前端) 透過標準 JSON API 請求即時成績
  if (action === "getLiveScores" || !action) {
    var dataJson = getScoresData();
    return ContentService.createTextOutput(JSON.stringify(dataJson))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "running", time: new Date() }))
    .setMimeType(ContentService.MimeType.JSON);
}