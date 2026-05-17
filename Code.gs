const SPREADSHEET_ID = '1wRpsLna87Hk9zM-DsIsBOYUJtwxqxh6YoZBNRISuU-M';

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('財產及物品管理系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    let requestData = JSON.parse(e.postData.contents);
    let action = requestData.action;
    
    if (action === 'saveProperty') {
      let res = saveProperty(requestData.data, requestData.p1Base64, requestData.p2Base64);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'updateProperty') {
      let res = updateProperty(requestData.rowIdx, requestData.data, requestData.p1Base64, requestData.p2Base64);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'searchProperties') {
      let res = searchProperties(requestData.query);
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: res })).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'getOptions') {
      let res = getOptions();
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: res })).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'getAppConfig') {
      let res = getAppConfig();
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: res })).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'generateScrapReport') {
      let res = generateScrapReport(requestData.targetDate, requestData.targetCategory);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'updateAuditStatus') {
      let res = updateAuditStatus(requestData.rowIdx, requestData.auditor);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'clearAuditData') {
      let res = clearAuditData(requestData.password);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 取得系統設定 (密碼)
 */
function getAppConfig() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("AppConfig");
  if (!sheet) {
    sheet = ss.insertSheet("AppConfig");
    sheet.getRange("A1:B1").setValues([["設定項目", "內容"]]);
    sheet.getRange("A2:B2").setValues([["系統密碼", "1234"]]); // 預設密碼
    sheet.setFrozenRows(1);
  }
  return {
    password: sheet.getRange("B2").getValue().toString()
  };
}

/**
 * 建立或取得存放照片的資料夾
 */
function getOrCreatePhotoFolder() {
  const folderName = "財產管理照片";
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

/**
 * 將 Base64 圖片上傳至 Google Drive 並回傳連結
 */
function uploadPhoto(base64Data, filename) {
  if (!base64Data) return "";
  try {
    const folder = getOrCreatePhotoFolder();
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    console.error("上傳圖片失敗: " + e);
    return "";
  }
}

/**
 * 初始化表頭
 */
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "類別", "編號", "名稱", "別名", "型式/廠牌", "數量單位", 
      "取得日期", "使用年限", "存置地點", "使用人使用單位", 
      "照片1", "照片2", "備註", "報廢日期", "是否已完成報廢流程",
      "照片1連結", "照片2連結", // 隱藏欄位供網頁讀取
      "最新查核時間", "查核人員"
    ]);
    sheet.setFrozenRows(1);
    // 隱藏連結欄位，並加寬圖片欄位
    sheet.hideColumns(16, 2);
    sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(12, 100);
  } else {
    // 確保既有試算表補上查核欄位
    if (sheet.getRange("R1").getValue() !== "最新查核時間") {
      sheet.getRange("R1:S1").setValues([["最新查核時間", "查核人員"]]);
    }
  }
}

/**
 * 新增財產資料
 */
function saveProperty(data, photo1Base64, photo2Base64) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("財產清單");
    if (!sheet) {
      sheet = ss.insertSheet("財產清單");
    }
    ensureHeaders(sheet);

    // 上傳照片
    const timestamp = new Date().getTime();
    const p1Url = uploadPhoto(photo1Base64, `${data.id}_photo1_${timestamp}.jpg`);
    const p2Url = uploadPhoto(photo2Base64, `${data.id}_photo2_${timestamp}.jpg`);

    // 準備圖片嵌入物件 (使用直接下載連結)
    const getDirectUrl = (url) => {
      if (!url) return "";
      const match = url.toString().match(/[-\w]{25,}/);
      return match ? `https://drive.google.com/thumbnail?id=${match[0]}&sz=w1000` : "";
    };
    
    const p1Direct = getDirectUrl(p1Url);
    const p2Direct = getDirectUrl(p2Url);

    const img1 = p1Direct ? SpreadsheetApp.newCellImage().setSourceUrl(p1Direct).setAltTextDescription(p1Url).build() : "";
    const img2 = p2Direct ? SpreadsheetApp.newCellImage().setSourceUrl(p2Direct).setAltTextDescription(p2Url).build() : "";

    const lastRow = sheet.getLastRow() + 1;
    const textRow = [
      data.category || "",
      data.id || "",
      data.name || "",
      data.alias || "",
      data.brand || "",
      data.unit || "",
      data.acquireDate || "",
      data.lifespan || "",
      data.location || "",
      data.userDept || "",
      "", // 照片1 佔位
      "", // 照片2 佔位
      data.notes || "",
      data.scrapDate || "",
      data.isScrapped ? "是" : "否",
      p1Url,
      p2Url
    ];

    // 先寫入文字資料
    sheet.getRange(lastRow, 1, 1, 17).setValues([textRow]);

    // 寫入圖片物件
    if (img1) sheet.getRange(lastRow, 11).setValue(img1);
    if (img2) sheet.getRange(lastRow, 12).setValue(img2);

    // 調整列高
    sheet.setRowHeight(lastRow, 80);
    return { success: true, message: "資料新增成功！" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 更新財產資料
 */
function updateProperty(rowIdx, data, photo1Base64, photo2Base64) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("財產清單");
    if (!sheet) throw new Error("找不到工作表");

    // 處理圖片 (如果有傳新圖片才更新)
    let p1Url = data.photo1;
    let p2Url = data.photo2;
    const timestamp = new Date().getTime();

    if (photo1Base64) p1Url = uploadPhoto(photo1Base64, `${data.id}_photo1_${timestamp}.jpg`);
    if (photo2Base64) p2Url = uploadPhoto(photo2Base64, `${data.id}_photo2_${timestamp}.jpg`);

    const getDirectUrl = (url) => {
      if (!url) return "";
      const match = url.toString().match(/[-\w]{25,}/);
      return match ? `https://drive.google.com/thumbnail?id=${match[0]}&sz=w1000` : "";
    };
    const p1Direct = getDirectUrl(p1Url);
    const p2Direct = getDirectUrl(p2Url);

    const textRow = [
      data.category || "",
      data.id || "",
      data.name || "",
      data.alias || "",
      data.brand || "",
      data.unit || "",
      data.acquireDate || "",
      data.lifespan || "",
      data.location || "",
      data.userDept || "",
      "", // 照片1 佔位
      "", // 照片2 佔位
      data.notes || "",
      data.scrapDate || "",
      data.isScrapped ? "是" : "否",
      p1Url,
      p2Url
    ];

    // 更新資料
    sheet.getRange(rowIdx, 1, 1, 17).setValues([textRow]);

    // 更新圖片物件
    if (p1Direct) sheet.getRange(rowIdx, 11).setValue(SpreadsheetApp.newCellImage().setSourceUrl(p1Direct).setAltTextDescription(p1Url).build());
    if (p2Direct) sheet.getRange(rowIdx, 12).setValue(SpreadsheetApp.newCellImage().setSourceUrl(p2Direct).setAltTextDescription(p2Url).build());

    return { success: true, message: "資料更新成功！" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 查詢財產資料
 */
function searchProperties(query) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("財產清單");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const results = [];
    
    const q = query.toString().toLowerCase();

    const getDirectUrl = (url) => {
      if (!url) return "";
      const match = url.toString().match(/[-\w]{25,}/);
      return match ? `https://drive.google.com/thumbnail?id=${match[0]}&sz=w1000` : "";
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const category = row[0];
      const id = row[1];
      const name = row[2];
      const alias = row[3];
      const isScrapped = row[14];

      // 檢查是否符合關鍵字 (編號、名稱、別名)
      if (id.toString().toLowerCase().includes(q) || 
          name.toString().toLowerCase().includes(q) || 
          alias.toString().toLowerCase().includes(q) ||
          q === "") {
        
        results.push({
          rowIdx: i + 1,
          category: category,
          id: id,
          name: name,
          alias: alias,
          brand: row[4],
          unit: row[5],
          acquireDate: row[6],
          lifespan: row[7],
          location: row[8],
          userDept: row[9],
          photo1: getDirectUrl(row[15]), // 轉換為直接連結
          photo2: getDirectUrl(row[16]), // 轉換為直接連結
          notes: row[12],
          scrapDate: row[13],
          isScrapped: isScrapped === "是",
          auditTime: row[17], // 第 18 欄
          auditor: row[18]    // 第 19 欄
        });
      }
    }
    return results;
  } catch (error) {
    console.error("搜尋錯誤", error);
    return [];
  }
}

/**
 * 取得現有的存置地點與使用人清單 (不重複)
 */
function getOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("財產清單");
    if (!sheet) return { locations: [], users: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { locations: [], users: [] };

    const locations = new Set();
    const users = new Set();

    for (let i = 1; i < data.length; i++) {
      if (data[i][8]) locations.add(data[i][8].toString());
      if (data[i][9]) users.add(data[i][9].toString());
    }

    return {
      locations: Array.from(locations).sort(),
      users: Array.from(users).sort()
    };
  } catch (e) {
    return { locations: [], users: [] };
  }
}

/**
 * 產生報廢單
 */
function generateScrapReport(targetDate, targetCategory) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 取得範本，如果沒有就建立一個陽春範本
    let templateSheet = ss.getSheetByName("報廢單範本");
    if (!templateSheet) {
      templateSheet = ss.insertSheet("報廢單範本");
      templateSheet.appendRow(["報廢單", "日期:", new Date().toLocaleDateString()]);
      templateSheet.appendRow(["編號", "名稱", "類別", "報廢日期"]);
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const newSheetName = `財產物品報廢單_${timestamp}`;
    const newSheet = templateSheet.copyTo(ss).setName(newSheetName);

    // 強制設定第 4 列 (表頭) 的高度與自動換行，
    // 解決 ODS 下載後在 Excel 開啟時，列高不會自動撐開導致「使用年限」等字樣被截斷的問題
    newSheet.setRowHeight(4, 45);
    newSheet.getRange("A4:K4").setWrap(true);

    // 尋找符合條件的資料
    const sourceSheet = ss.getSheetByName("財產清單");
    if (!sourceSheet) throw new Error("找不到財產清單");
    const data = sourceSheet.getDataRange().getValues();
    
    let currentRow = 5; // 從第5列開始填寫
    let matchCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const category = row[0];
      const id = row[1];
      const name = row[2];
      const explicitScrapDate = row[13];
      
      // 處理試算表中的手動報廢日期
      let scrapDateStr = "";
      if (explicitScrapDate) {
        const d = new Date(explicitScrapDate);
        if (!isNaN(d.getTime())) {
          scrapDateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        } else {
          scrapDateStr = explicitScrapDate.toString();
        }
      }

      // 只有當手動填寫的「報廢日期」符合時，才列入報廢單
      if (category === targetCategory && (scrapDateStr === targetDate || (explicitScrapDate && explicitScrapDate.toString().includes(targetDate)))) {
        
        let rocDateStr = "";
        let usedYears = "";
        
        const acquireDateRaw = row[6];
        if (acquireDateRaw) {
          const acq = new Date(acquireDateRaw);
          if (!isNaN(acq.getTime())) {
            const rocYear = acq.getFullYear() - 1911;
            const mm = String(acq.getMonth() + 1).padStart(2, '0');
            const dd = String(acq.getDate()).padStart(2, '0');
            rocDateStr = `${rocYear}/${mm}/${dd}`;
            
            const now = new Date();
            let years = now.getFullYear() - acq.getFullYear();
            if (now.getMonth() < acq.getMonth() || (now.getMonth() === acq.getMonth() && now.getDate() < acq.getDate())) {
              years--;
            }
            usedYears = Math.max(0, years);
          } else {
             rocDateStr = acquireDateRaw.toString();
          }
        }

        const unitQty = row[5]; // 數量
        const lifespan = row[7]; // 使用年限
        const writeOffReason = "不堪使用";

        // 依據範本欄位順序：項次, 財產編號, 財產名稱, 數量, 總價, 購置日期, 使用年限, 已使用年數, 減損原由, 減損後財產流向, 備註
        const rowData = [
          matchCount + 1, // 項次
          id,             // 財產編號
          name,           // 財產名稱
          unitQty,        // 數量
          "",             // 總價
          rocDateStr,     // 購置日期 (民國年)
          lifespan,       // 使用年限
          usedYears,      // 已使用年數
          writeOffReason, // 減損原由
          "",             // 減損後財產流向
          ""              // 備註
        ];

        // 如果資料已經填到第 19 列 (接近簽核欄位)，則自動插入一行新的空白列
        if (currentRow >= 19) {
          newSheet.insertRowBefore(currentRow);
        }

        newSheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
        currentRow++;
        matchCount++;
      }
    }

    if (matchCount === 0) {
      ss.deleteSheet(newSheet);
      return { success: false, message: "找不到符合條件的報廢資料" };
    }

    return { 
      success: true, 
      message: `成功產生報廢單！共 ${matchCount} 筆資料。`,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=ods&gid=${newSheet.getSheetId()}`
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 更新查核狀態
 */
function updateAuditStatus(rowIdx, auditor) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("財產清單");
    if (!sheet) throw new Error("找不到工作表");
    
    // 時間格式 yyyy/MM/dd HH:mm:ss
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    
    // R欄(18) = 時間, S欄(19) = 人員
    sheet.getRange(rowIdx, 18, 1, 2).setValues([[timestamp, auditor]]);
    
    return { success: true, message: "查核完成", auditTime: timestamp, auditor: auditor };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 一鍵清除所有查核資料
 */
function clearAuditData(password) {
  try {
    const config = getAppConfig();
    if (config.password !== password) {
      return { success: false, message: "密碼錯誤，無法清除資料" };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("財產清單");
    if (!sheet) throw new Error("找不到工作表");
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // 清除 R 欄與 S 欄 (行2到最後一行的第18,19欄)
      sheet.getRange(2, 18, lastRow - 1, 2).clearContent();
    }
    
    return { success: true, message: "查核資料已成功清除！" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
