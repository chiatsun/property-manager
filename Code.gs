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
    } else if (action === 'searchProperties') {
      let res = searchProperties(requestData.query);
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: res })).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'generateScrapReport') {
      let res = generateScrapReport(requestData.targetDate, requestData.targetCategory);
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
      "照片1連結", "照片2連結" // 隱藏欄位供網頁讀取
    ]);
    sheet.setFrozenRows(1);
    // 隱藏連結欄位，並加寬圖片欄位
    sheet.hideColumns(16, 2);
    sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(12, 100);
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
    const getDirectUrl = (url) => url ? url.replace("file/d/", "uc?export=download&id=").replace("/view?usp=drivesdk", "").replace("/view", "") : "";
    
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
          photo1: row[15], // 從隱藏欄位讀取連結
          photo2: row[16], // 從隱藏欄位讀取連結
          notes: row[12],
          scrapDate: row[13],
          isScrapped: isScrapped === "是"
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

    // 尋找符合條件的資料
    const sourceSheet = ss.getSheetByName("財產清單");
    if (!sourceSheet) throw new Error("找不到財產清單");
    const data = sourceSheet.getDataRange().getValues();
    
    let currentRow = templateSheet.getLastRow() + 1; // 接在範本下面填寫
    let matchCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const category = row[0];
      const name = row[2];
      const id = row[1];
      const sDate = row[13];
      
      // 假設 sDate 格式與 targetDate 可比較，或者做簡單字串比較
      if (category === targetCategory && sDate.toString().includes(targetDate)) {
        newSheet.getRange(currentRow, 1, 1, 4).setValues([[id, name, category, sDate]]);
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
      message: `成功產生報廢單: ${newSheetName}，共 ${matchCount} 筆資料。`,
      sheetUrl: `${ss.getUrl()}#gid=${newSheet.getSheetId()}`
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
