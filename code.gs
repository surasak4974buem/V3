// Google Apps Script สำหรับระบบฐานข้อมูลคลังอะไหล่แพทย์และแจ้งเตือน LINE OA โรงพยาบาลนครพิงค์
// วิธีใช้งาน: นำโค้ดทั้งหมดนี้ไปวางใน "ส่วนขยาย" -> "Apps Script" ของไฟล์ Google Sheets ของคุณ

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  
  if (data.action === "sync") {
    // 1. ซิงก์ชีตข้อมูลอะไหล่ (Parts)
    var partsSheet = ss.getSheetByName("Parts") || ss.insertSheet("Parts");
    partsSheet.clear();
    if (data.parts && data.parts.length > 0) {
      var partsHeaders = Object.keys(data.parts[0]);
      partsSheet.appendRow(partsHeaders);
      var partsRows = data.parts.map(function(p) {
        return partsHeaders.map(function(h) {
          // หากค่าเป็น Object/Array ให้แปลงเป็นสตริง JSON
          return typeof p[h] === 'object' ? JSON.stringify(p[h]) : p[h];
        });
      });
      partsSheet.getRange(2, 1, partsRows.length, partsHeaders.length).setValues(partsRows);
    }
    
    // 2. ซิงก์ชีตประวัติการทำรายการคลัง (Transactions)
    var txSheet = ss.getSheetByName("Transactions") || ss.insertSheet("Transactions");
    txSheet.clear();
    if (data.transactions && data.transactions.length > 0) {
      var txHeaders = Object.keys(data.transactions[0]);
      txSheet.appendRow(txHeaders);
      var txRows = data.transactions.map(function(t) {
        return txHeaders.map(function(h) {
          return typeof t[h] === 'object' ? JSON.stringify(t[h]) : t[h];
        });
      });
      txSheet.getRange(2, 1, txRows.length, txHeaders.length).setValues(txRows);
    }

    // 3. ซิงก์ชีตรายชื่อและสิทธิ์ผู้ใช้งาน (Users)
    var usersSheet = ss.getSheetByName("Users") || ss.insertSheet("Users");
    usersSheet.clear();
    if (data.users && data.users.length > 0) {
      var usersHeaders = Object.keys(data.users[0]);
      usersSheet.appendRow(usersHeaders);
      var usersRows = data.users.map(function(u) {
        return usersHeaders.map(function(h) {
          return typeof u[h] === 'object' ? JSON.stringify(u[h]) : u[h];
        });
      });
      usersSheet.getRange(2, 1, usersRows.length, usersHeaders.length).setValues(usersRows);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (data.action === "notify") {
    var lineResult = sendLineNotify(data.message, data.token, data.groupId);
    return ContentService.createTextOutput(JSON.stringify({status: "sent", result: lineResult}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};
  
  var partsSheet = ss.getSheetByName("Parts");
  if (partsSheet) {
    result.parts = getSheetDataJson(partsSheet);
  }
  
  var txSheet = ss.getSheetByName("Transactions");
  if (txSheet) {
    result.transactions = getSheetDataJson(txSheet);
  }

  var usersSheet = ss.getSheetByName("Users");
  if (usersSheet) {
    result.users = getSheetDataJson(usersSheet);
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันช่วยดึงข้อมูลจากชีตแปลงเป็น JSON Array
function getSheetDataJson(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  
  var headers = rows[0];
  var data = [];
  
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      // ลองตรวจเช็กว่าค่าในช่องนั้นเป็น JSON string ของ Array/Object หรือไม่
      if (typeof val === 'string' && (val.indexOf('[') === 0 || val.indexOf('{') === 0)) {
        try {
          val = JSON.parse(val);
        } catch(e) {}
      }
      obj[headers[j]] = val;
    }
    data.push(obj);
  }
  return data;
}

// ฟังก์ชันส่ง Flex Message ไปกลุ่ม LINE OA
function sendLineNotify(flexMessage, token, groupId) {
  var url = "https://api.line.me/v2/bot/message/push";
  var payload = {
    "to": groupId,
    "messages": [
      {
        "type": "flex",
        "altText": "🚨 แจ้งเตือนการทำรายการคลังอะไหล่เครื่องมือแพทย์ รพ.นครพิงค์",
        "contents": flexMessage
      }
    ]
  };
  
  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + token
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  return response.getContentText();
}
