// ============================================================
//  CHECK SCANNER — Backend (Google Apps Script)
//  รับ request จาก GitHub Pages Web App
//  action: "ocr"  → ส่งรูปให้ Claude อ่านข้อมูลเช็ค
//  action: "save" → บันทึก PDF ขึ้น Drive + Link ใน Sheet
// ============================================================

// ── ตั้งค่าก่อน Deploy ──────────────────────────────────────
var CONFIG = {
  CLAUDE_API_KEY   : "ใส่ Claude API Key ที่นี่",   // console.anthropic.com
  ROOT_FOLDER_NAME : "ทะเบียนเช็คสั่งจ่าย",
  SHEET_NAME       : "ทะเบียนเช็คหลัก",
  COL_CHECK_NO     : 1,   // A
  COL_LINK         : 27,  // AA
  DATA_START_ROW   : 5,
  GITHUB_PAGES_URL : "https://YOUR_USERNAME.github.io/check-scanner", // URL ของ GitHub Pages
};
// ──────────────────────────────────────────────────────────

// ── รับ POST request จาก Web App ──
function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin"  : CONFIG.GITHUB_PAGES_URL,
    "Access-Control-Allow-Methods" : "POST",
    "Access-Control-Allow-Headers" : "Content-Type",
    "Content-Type"                 : "application/json"
  };

  try {
    var body   = JSON.parse(e.postData.contents);
    var result = {};

    if (body.action === "ocr") {
      result = analyzeCheck(body.image, body.mime);
    } else if (body.action === "save") {
      result = saveCheckFile(body.image, body.mime, body.data);
    } else {
      result = { success:false, error:"Unknown action" };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success:false, error:err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// รองรับ OPTIONS preflight
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status:"ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── OCR: ส่งรูปให้ Claude อ่านข้อมูลเช็ค ──
function analyzeCheck(base64Image, mimeType) {
  try {
    var payload = {
      model      : "claude-sonnet-4-6",
      max_tokens : 800,
      messages   : [{
        role    : "user",
        content : [
          { type:"image", source:{ type:"base64", media_type:mimeType, data:base64Image } },
          { type:"text", text:
            "วิเคราะห์เช็คในรูปนี้ ตอบเป็น JSON เท่านั้น ไม่มีข้อความอื่น:\n" +
            '{"check_no":"เลขที่เช็ค 7-8 หลัก","bank":"ชื่อธนาคาร","amount":"ยอดเงินตัวเลข","date":"DD/MM/YYYY","payee":"ชื่อผู้รับเงิน","confidence":"high หรือ medium หรือ low"}'
          }
        ]
      }]
    };

    var res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method:"post", contentType:"application/json",
      headers:{ "x-api-key":CONFIG.CLAUDE_API_KEY, "anthropic-version":"2023-06-01" },
      payload:JSON.stringify(payload), muteHttpExceptions:true
    });

    var json = JSON.parse(res.getContentText());
    var text = (json.content && json.content[0] && json.content[0].text) || "{}";
    var data = JSON.parse(text.replace(/```json|```/g,"").trim());
    data.success  = true;
    data.filename = (data.check_no||"") + "_" + (data.bank||"") + ".pdf";
    return data;

  } catch(e) { return { success:false, error:e.message }; }
}

// ── บันทึก PDF ขึ้น Drive ──
function saveCheckFile(base64Image, mimeType, checkData) {
  try {
    var decoded  = Utilities.base64Decode(base64Image);
    var blob     = Utilities.newBlob(decoded, mimeType, checkData.filename);

    // แปลงรูปเป็น PDF
    var fileBlob = blob;
    if (mimeType.startsWith("image/")) {
      var tmp  = DriveApp.createFile(blob);
      fileBlob = tmp.getAs("application/pdf");
      fileBlob.setName(checkData.filename);
      tmp.setTrashed(true);
    }

    // หา/สร้างโฟลเดอร์ตามปี
    var folder    = getOrCreateYearFolder(checkData.date);
    var driveFile = folder.createFile(fileBlob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl   = driveFile.getUrl();

    // ใส่ Link ใน Sheet
    var linked = linkToSheet(checkData.check_no, fileUrl);

    return { success:true, fileUrl:fileUrl, filename:checkData.filename, linked:linked };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── หา/สร้างโฟลเดอร์ปี ──
function getOrCreateYearFolder(dateStr) {
  var fList = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  var root  = fList.hasNext() ? fList.next() : DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);

  var yearBE = new Date().getFullYear() + 543;
  if (dateStr) {
    var p = dateStr.split("/");
    if (p.length === 3) {
      var y = parseInt(p[2]);
      yearBE = y > 2400 ? y : y + 543;
    }
  }

  var sub = root.getFoldersByName(yearBE.toString());
  return sub.hasNext() ? sub.next() : root.createFolder(yearBE.toString());
}

// ── ใส่ Link ใน Sheet ──
function linkToSheet(checkNo, fileUrl) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return false;
    var last  = sheet.getLastRow();
    for (var r = CONFIG.DATA_START_ROW; r <= last; r++) {
      if (sheet.getRange(r, CONFIG.COL_CHECK_NO).getValue().toString().trim() === checkNo.toString().trim()) {
        var cell = sheet.getRange(r, CONFIG.COL_LINK);
        cell.setFormula('=HYPERLINK("'+fileUrl+'","📎 '+checkNo+'")');
        cell.setFontColor("#1565C0");
        return true;
      }
    }
    return false;
  } catch(e) { return false; }
}
