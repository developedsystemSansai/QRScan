// ============================================================
// ระบบลงทะเบียนและเช็คอินงานสัมมนา
// โรงพยาบาลสันทราย – HRD
// ============================================================
// วิธีใช้:
//   1. เปิด Spreadsheet → Extensions → Apps Script → วางโค้ดนี้
//   2. รัน setupSheet() เพื่อเตรียม Sheet
//   3. สร้าง Google Form ด้วยมือ แล้ว Link กับ Sheet นี้
//   4. ผูก Trigger: onFormSubmit (On form submit)
//   5. Deploy Web App (Execute as: Me, Access: Anyone)
// ============================================================

// ──────────────────────────────────────────────
// ตั้งค่าระบบ — แก้ตรงนี้อย่างเดียว
// ──────────────────────────────────────────────
var CONFIG = {
  SHEET_NAME       : "การลงทะเบียน",
  // ★ วางรหัส Spreadsheet ของคุณที่นี่ (เอาจาก URL ระหว่าง /d/ กับ /edit)
  // ตัวอย่าง: https://docs.google.com/spreadsheets/d/★ตรงนี้★/edit
  SPREADSHEET_ID   : "1Us8hTn68j49T8pX9T8GoQRUBRNTmFjPWkYKDW7Nmx5s",
  EVENT_NAME       : "สัมมนาแมชชีนเลิร์นนิงและปัญญาประดิษฐ์",
  EVENT_DATE       : "วันที่ 1 สิงหาคม 2568",
  EVENT_VENUE      : "ห้องประชุมโรงพยาบาลสันทราย",
  FORM_FIELD_NAME  : "ชื่อ-นามสกุล",   // ต้องตรงกับชื่อคำถามใน Form ทุกตัวอักษร
  FORM_FIELD_EMAIL : "อีเมล"            // ต้องตรงกับชื่อคำถามใน Form ทุกตัวอักษร
};

// ──────────────────────────────────────────────
// Helper: เปิด Spreadsheet ได้ทั้งจาก UI และ Web App
// ──────────────────────────────────────────────
function getSpreadsheet() {
  try {
    // ทำงานได้เมื่อเรียกจาก Spreadsheet UI / Trigger
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  // ทำงานได้เมื่อเรียกจาก Web App (ไม่มี active spreadsheet)
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet() {
  return getSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
}
/**
 * [1] SETUP SHEET & FORM — รันฟังก์ชันนี้เพียงครั้งเดียวเพื่อเริ่มใช้งาน
 * ฟังก์ชันจะสร้าง Google Form, เชื่อมต่อกับ Sheet, จัดรูปแบบ และตั้งค่า Trigger ให้อัตโนมัติ
 */
function setupSheet() {
  var ss = getSpreadsheet(); // [cite: 111]
  
  // 1. ตรวจสอบว่า Spreadsheet นี้มีการผูก Form ไว้แล้วหรือยัง
  if (ss.getFormUrl()) {
    SpreadsheetApp.getUi().alert(
      "⚠️ ตรวจพบฟอร์มเดิม", 
      "Spreadsheet นี้มีการผูก Google Form ไว้แล้วครับ หากต้องการสร้างใหม่กรุณาลบฟอร์มเดิมออกก่อน\nลิงก์ฟอร์ม: " + ss.getFormUrl(), 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // 2. สร้าง Google Form ใหม่ตามชื่อกิจกรรม
  var form = FormApp.create("ลงทะเบียน: " + CONFIG.EVENT_NAME); // [cite: 110]
  form.setDescription("กรุณากรอกข้อมูลเพื่อลงทะเบียนเข้าร่วม\n" + CONFIG.EVENT_DATE + " ณ " + CONFIG.EVENT_VENUE);
  
  // 3. เพิ่มคำถามตามค่าที่ตั้งไว้ใน CONFIG
  form.addTextItem().setTitle(CONFIG.FORM_FIELD_NAME).setRequired(true); // ชื่อ-นามสกุล [cite: 110]
  form.addTextItem().setTitle(CONFIG.FORM_FIELD_EMAIL).setRequired(true); // อีเมล [cite: 110]
  
  // 4. ตั้งค่าให้ Form ส่งข้อมูลมาที่ Spreadsheet นี้
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  // รอระบบประมวลผลการเชื่อมต่อ
  SpreadsheetApp.flush();
  Utilities.sleep(2000); 

  // 5. ค้นหา Sheet ที่ถูกสร้างขึ้นใหม่จากการเชื่อมต่อ Form
  var sheets = ss.getSheets();
  var formSheet = sheets.find(function(s) { return s.getFormUrl() != null; });
  
  if (formSheet) {
    // เปลี่ยนชื่อ Sheet ให้ตรงตามที่ระบบต้องการ
    formSheet.setName(CONFIG.SHEET_NAME); // "การลงทะเบียน" [cite: 110]
    
    // 6. เพิ่มหัวข้อคอลัมน์และจัดรูปแบบ [cite: 115]
    var headers = ["Timestamp", CONFIG.FORM_FIELD_NAME, CONFIG.FORM_FIELD_EMAIL, "Registration ID", "Email Status", "Attendance"];
    formSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground("#1a237e")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
      
    formSheet.setRowHeight(1, 36);
    formSheet.setFrozenRows(1);
    
    // ตั้งความกว้างคอลัมน์ [cite: 117]
    [160, 180, 200, 240, 110, 200].forEach(function(w, i){ 
      formSheet.setColumnWidth(i+1, w); 
    });

    // ตั้งค่าสีสถานะ (Conditional Formatting) [cite: 117]
    formSheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Present")
        .setBackground("#e8f5e9").setFontColor("#2e7d32").setBold(true)
        .setRanges([formSheet.getRange("F2:F1000")]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("ส่งแล้ว")
        .setBackground("#e3f2fd").setFontColor("#1565c0")
        .setRanges([formSheet.getRange("E2:E1000")]).build()
    ]);
  }

  // 7. สร้าง Trigger สำหรับส่งอีเมลอัตโนมัติ (onFormSubmit) [cite: 119]
  var triggers = ScriptApp.getProjectTriggers();
  var isTriggerSet = triggers.some(function(t) { return t.getHandlerFunction() === "onFormSubmit"; });
  
  if (!isTriggerSet) {
    ScriptApp.newTrigger("onFormSubmit")
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();
  }

  // 8. แจ้งผลสำเร็จและให้ลิงก์ฟอร์ม
  SpreadsheetApp.getUi().alert(
    "✅ ตั้งค่าระบบสำเร็จ!",
    "ระบบสร้าง Google Form และเชื่อมต่อเรียบร้อยแล้ว\n\n" +
    "🔗 ลิงก์ฟอร์มสำหรับส่งให้ผู้เข้าร่วม:\n" + form.getPublishedUrl() + "\n\n" +
    "อย่าลืมทำการ Deploy Web App เพื่อให้ระบบสแกนใช้งานได้ครับ",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ══════════════════════════════════════════════
// [2] ON FORM SUBMIT — ทำงานอัตโนมัติเมื่อมีการส่ง Form
// ══════════════════════════════════════════════
function onFormSubmit(e) {
  try {
    var sheet = getSheet();
    if (!sheet) { Logger.log("ไม่พบ Sheet: " + CONFIG.SHEET_NAME); return; }

    // อ่านข้อมูลจาก namedValues
    var values   = e.namedValues || {};
    var fullName = ((values[CONFIG.FORM_FIELD_NAME]  || [""])[0] || "").trim();
    var email    = ((values[CONFIG.FORM_FIELD_EMAIL] || [""])[0] || "").trim();

    if (!fullName || !email) {
      Logger.log("ไม่พบชื่อหรืออีเมล keys=" + Object.keys(values).join(","));
      return;
    }

    // หาแถวที่ Form เพิ่งเขียน (แถวที่ยังไม่มี Registration ID)
    var lastRow   = sheet.getLastRow();
    var targetRow = lastRow;
    if (lastRow >= 2) {
      var regCol = sheet.getRange(2, 4, lastRow-1, 1).getValues();
      for (var i = regCol.length-1; i >= 0; i--) {
        if (!String(regCol[i][0]).trim()) { targetRow = i+2; break; }
      }
    }

    // Registration ID
    var regId = "REG-" + Date.now() + "-" + targetRow;
    sheet.getRange(targetRow, 4).setValue(regId);

    // ส่งอีเมล
    sendConfirmationEmail(email, fullName, regId);

    // สถานะ
    sheet.getRange(targetRow, 5).setValue("ส่งแล้ว");
    SpreadsheetApp.flush();
    Logger.log("OK: " + fullName + " | " + regId);

  } catch(err) {
    Logger.log("onFormSubmit error: " + err.message);
  }
}


// ══════════════════════════════════════════════
// [3] SEND EMAIL
// ══════════════════════════════════════════════
function sendConfirmationEmail(toEmail, fullName, regId) {
  var qrUrl   = "https://quickchart.io/qr?text=" + encodeURIComponent(regId) + "&size=280&margin=2";
  var subject = "ยืนยันการลงทะเบียน – " + CONFIG.EVENT_NAME;

  var html =
    '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:Tahoma,sans-serif;background:#f0f4ff;color:#1a1a2e}' +
    '.w{max-width:540px;margin:24px auto;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 4px 20px rgba(26,35,126,.12)}' +
    '.h{background:linear-gradient(135deg,#1a237e,#3949ab);padding:32px;text-align:center;color:#fff}' +
    '.h .i{font-size:40px;margin-bottom:8px}.h h1{font-size:19px;font-weight:700;line-height:1.4}' +
    '.h p{font-size:12px;opacity:.85;margin-top:5px}' +
    '.b{padding:26px}' +
    '.c{background:#f5f7ff;border:1.5px solid #c5cae9;border-radius:10px;padding:16px;margin:16px 0}' +
    '.r{display:flex;gap:8px;margin-bottom:7px;align-items:baseline}.r:last-child{margin:0}' +
    '.l{font-size:10px;color:#7986cb;font-weight:700;text-transform:uppercase;min-width:90px}' +
    '.v{font-size:13px;color:#1a1a2e;font-weight:500}' +
    '.id{font-family:monospace;background:#e8eaf6;padding:2px 7px;border-radius:5px;color:#283593;font-weight:700}' +
    '.q{text-align:center;margin:20px 0}.q p{font-size:12px;color:#5c6bc0;margin-bottom:10px;font-weight:600}' +
    '.n{background:#fff8e1;border-left:4px solid #ffc107;border-radius:0 8px 8px 0;padding:10px 13px;font-size:11px;color:#5d4037;line-height:1.6}' +
    '.f{text-align:center;padding:16px;background:#f5f7ff;font-size:10px;color:#9fa8da;border-top:1px solid #e8eaf6}' +
    '</style></head><body>' +
    '<div class="w">' +
    '<div class="h"><div class="i">🎓</div>' +
    '<h1>' + CONFIG.EVENT_NAME + '</h1>' +
    '<p>' + CONFIG.EVENT_DATE + ' | ' + CONFIG.EVENT_VENUE + '</p></div>' +
    '<div class="b">' +
    '<p style="font-size:15px;font-weight:600;color:#1a237e;margin-bottom:14px">สวัสดีครับ/ค่ะ คุณ' + fullName + '</p>' +
    '<p style="font-size:13px;line-height:1.7;color:#37474f;margin-bottom:4px">ขอบคุณที่ลงทะเบียน กรุณาเก็บ QR Code นี้ไว้เพื่อใช้เช็คอินหน้างาน</p>' +
    '<div class="c">' +
    '<div class="r"><span class="l">ชื่อ-นามสกุล</span><span class="v">' + fullName + '</span></div>' +
    '<div class="r"><span class="l">รหัสลงทะเบียน</span><span class="v"><span class="id">' + regId + '</span></span></div>' +
    '<div class="r"><span class="l">วันที่</span><span class="v">' + CONFIG.EVENT_DATE + '</span></div>' +
    '<div class="r"><span class="l">สถานที่</span><span class="v">' + CONFIG.EVENT_VENUE + '</span></div>' +
    '</div>' +
    '<div class="q"><p>📱 QR Code สำหรับเช็คอินหน้างาน</p>' +
    '<img src="' + qrUrl + '" width="210" height="210" style="border:5px solid #e8eaf6;border-radius:12px"></div>' +
    '<div class="n">⚠️ <strong>หมายเหตุ:</strong> กรุณาแสดง QR Code นี้ต่อเจ้าหน้าที่ ณ จุดลงทะเบียนหน้างาน สามารถบันทึกภาพหน้าจอได้</div>' +
    '</div>' +
    '<div class="f">อีเมลนี้ส่งโดยอัตโนมัติ | ' + CONFIG.EVENT_VENUE + '</div>' +
    '</div></body></html>';

  GmailApp.sendEmail(toEmail, subject, "", { htmlBody: html });
}


// ══════════════════════════════════════════════
// [4] WEB APP
// ══════════════════════════════════════════════
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("เช็คอิน & Dashboard – " + CONFIG.EVENT_NAME)
    .addMetaTag("viewport","width=device-width,initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ══════════════════════════════════════════════
// [5] PROCESS CHECK-IN — เรียกจาก Web App
// ══════════════════════════════════════════════
function processCheckIn(registrationId) {
  try {
    var regId = String(registrationId || "").trim();
    if (!regId) return { status:"error", message:"ไม่พบรหัสลงทะเบียน" };

    var sheet = getSheet();
    if (!sheet) return { status:"error", message:"ไม่พบ Sheet ข้อมูล" };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status:"not_found", message:"ไม่พบรหัสลงทะเบียนนี้ในระบบ" };

    var data = sheet.getRange(2, 1, lastRow-1, 6).getValues();
    var now  = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][3]).trim() !== regId) continue;

      var fullName   = String(data[i][1]).trim();
      var attendance = String(data[i][5]).trim();

      if (!attendance) {
        sheet.getRange(i+2, 6).setValue("Present – " + now);
        SpreadsheetApp.flush();
        return { status:"checkin_ok", message:"เช็คอินสำเร็จ: " + fullName, name:fullName, time:now };
      } else {
        return { status:"already_checkin", message:fullName + " ได้เช็คอินไปแล้ว", name:fullName };
      }
    }

    return { status:"not_found", message:"ไม่พบรหัสลงทะเบียนนี้ในระบบ" };

  } catch(err) {
    Logger.log("processCheckIn: " + err.message);
    return { status:"error", message:"เกิดข้อผิดพลาด: " + err.message };
  }
}


// ══════════════════════════════════════════════
// [6] GET DASHBOARD DATA — เรียกจาก Web App
// ══════════════════════════════════════════════
function getDashboardData() {
  try {
    var sheet = getSheet();

    var result = {
      total        : 0,
      checkedIn    : 0,
      emailSent    : 0,
      recentCheckins: [],
      hourly       : [],
      pending      : [],
      lastUpdate   : Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss")
    };

    if (!sheet || sheet.getLastRow() < 2) return result;

    var lastRow = sheet.getLastRow();
    var data    = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    // ตัวแปรเก็บข้อมูล hourly (key = "HH")
    var hourlyMap = {};

    for (var i = 0; i < data.length; i++) {
      var row        = data[i];
      var fullName   = String(row[1]).trim();
      var regId      = String(row[3]).trim();
      var emailStat  = String(row[4]).trim();
      var attendance = String(row[5]).trim();

      if (!regId) continue; // ข้ามแถวที่ยังไม่มีข้อมูล
      result.total++;

      if (emailStat === "ส่งแล้ว") result.emailSent++;

      if (attendance && attendance.indexOf("Present") === 0) {
        result.checkedIn++;

        // แยกเวลา: "Present – dd/MM/yyyy HH:mm:ss"
        var timePart = attendance.replace(/^Present\s*[–-]\s*/, "").trim(); // "dd/MM/yyyy HH:mm:ss"
        var shortTime = "";
        var hourKey   = "";

        if (timePart) {
          // timePart = "dd/MM/yyyy HH:mm:ss"
          var parts = timePart.split(" ");  // ["dd/MM/yyyy", "HH:mm:ss"]
          if (parts.length >= 2) {
            shortTime = parts[1].substring(0, 5); // "HH:mm"
            hourKey   = parts[1].substring(0, 2); // "HH"
            hourlyMap[hourKey] = (hourlyMap[hourKey] || 0) + 1;
          }
        }

        // เก็บ 15 รายการล่าสุด (เรียง desc แล้วสไลซ์ทีหลัง)
        result.recentCheckins.push({ name: fullName, time: shortTime, rawTime: timePart });
      } else {
        // ยังไม่ได้เช็คอิน
        result.pending.push({ name: fullName, regId: regId });
      }
    }

    // เรียง recentCheckins ล่าสุดก่อน (เรียงตาม rawTime desc)
    result.recentCheckins.sort(function(a, b) {
      return b.rawTime.localeCompare(a.rawTime);
    });
    result.recentCheckins = result.recentCheckins.slice(0, 15).map(function(r) {
      return { name: r.name, time: r.time };
    });

    // สร้าง hourly array (ช่วงเวลา 7:00 – 18:00)
    for (var h = 7; h <= 18; h++) {
      var hStr = (h < 10 ? "0" : "") + h;
      result.hourly.push({ hour: hStr + ":00", count: hourlyMap[hStr] || 0 });
    }

    return result;

  } catch(err) {
    Logger.log("getDashboardData error: " + err.message);
    throw new Error("โหลดข้อมูล Dashboard ล้มเหลว: " + err.message);
  }
}


// ══════════════════════════════════════════════
// [7] RESEND EMAILS — ส่งอีเมลซ้ำ (รันด้วยมือ)
// ══════════════════════════════════════════════
function resendPendingEmails() {
  var ss    = getSpreadsheet();
  var sheet = getSheet();
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("ไม่พบข้อมูล");
    return;
  }

  var data  = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
  var count = 0;

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][4]).trim() === "ส่งแล้ว") continue;
    var row      = i + 2;
    var fullName = String(data[i][1]).trim();
    var email    = String(data[i][2]).trim();
    var regId    = String(data[i][3]).trim();

    if (!regId) {
      regId = "REG-" + Date.now() + "-" + row;
      sheet.getRange(row, 4).setValue(regId);
    }
    sendConfirmationEmail(email, fullName, regId);
    sheet.getRange(row, 5).setValue("ส่งแล้ว");
    count++;
    Utilities.sleep(1500);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("✅ ส่งอีเมลซ้ำ " + count + " รายการ");
}


// ══════════════════════════════════════════════
// [8] CUSTOM MENU
// ══════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎓 ระบบสัมมนา")
    .addItem("1️⃣  ตั้งค่า Sheet", "setupSheet")
    .addSeparator()
    .addItem("📧 ส่งอีเมลซ้ำ (คนที่ยังไม่ได้รับ)", "resendPendingEmails")
    .addToUi();
}
