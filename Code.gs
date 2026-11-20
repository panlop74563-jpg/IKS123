
/*******************************************************
 * ระบบรับสมัครนักเรียนประจำปีการศึกษา2570
 * โรงเรียนอิสลามวิทยานุเคราะห์
 *******************************************************/

// 1. ตั้งค่าระบบ
const CONFIG = {
  SCHOOL_NAME: "โรงเรียนอิสลามวิทยานุเคราะห์",
  SHEET_NAME: "Students",
  FOLDER_NAME: "ระบบรับสมัครนักเรียนประจำปีการศึกษา2570",
  ADMIN_PASSWORD: "CHANGE_ME_2828",
  MAX_FILE_SIZE: 5 * 1024 * 1024,
  TIMEZONE: "Asia/Bangkok"
};

// 2. หัวตาราง Google Sheets
const HEADERS = [
  "รหัสสมัคร", "วันเวลาสมัคร", "คำนำหน้า", "ชื่อ-สกุล",
  "วันเกิด", "เลขบัตรนักเรียน", "ชื่อเล่น", "จำนวนพี่น้อง",
  "ลำดับบุตร", "อายุ", "ที่อยู่", "ระดับชั้น", "โปรแกรม",
  "วุฒิสามัญ", "วุฒิศาสนา", "โรงเรียนเดิม",
  "ชื่อบิดา", "อาชีพบิดา", "โทรบิดา", "บัตรบิดา",
  "ชื่อมารดา", "อาชีพมารดา", "โทรมารดา", "บัตรมารดา",
  "ผู้ติดต่อฉุกเฉิน", "โทรฉุกเฉิน", "โรคประจำตัว",
  "ข้อมูลเพิ่มเติม", "ประเภทการพัก",
  "บัตรประชาชนหรือสูติบัตร", "ทะเบียนบ้าน",
  "สลิปชำระเงิน", "สถานะ", "หมายเหตุ", "Folder ID"
];

// 3. เปิดเว็บไซต์
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle(CONFIG.SCHOOL_NAME + " | ระบบรับสมัครนักเรียน ประจำปีการศึกษา 2570")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// 4. โลโก้และ QR Code
function getImageDataFromConfig(type) {
  if (type === "logo") {
    return "https://i.postimg.cc/ZJnPm3rK/images-Photoroom.png";
  }

  if (type === "qr") {
    return "https://i.postimg.cc/44n4yJwK/13.png";
  }

  return "";
}

// 5. สร้างหรือเชื่อมต่อชีต
function getStudentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("ไม่พบ Google Sheets กรุณาเปิด Apps Script จากไฟล์ฐานข้อมูล");
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);

    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground("#8F2429")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");

    sheet.autoResizeColumns(1, HEADERS.length);
  }

  return sheet;
}

// 6. สร้างโฟลเดอร์หลักใน Google Drive
function getMainFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty("MAIN_FOLDER_ID");

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (error) {
      props.deleteProperty("MAIN_FOLDER_ID");
    }
  }

  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  let folder;

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  }

  props.setProperty("MAIN_FOLDER_ID", folder.getId());
  return folder;
}

// 7. ติดตั้งระบบครั้งแรก
function setupSystem() {
  getStudentSheet();
  getMainFolder();
  return "ติดตั้งระบบเรียบร้อยแล้ว";
}

// 8. บันทึกไฟล์ลง Google Drive
function saveFile(file, folder, filename) {
  if (!file || !file.data) {
    throw new Error("กรุณาแนบไฟล์ " + filename);
  }

  const allowedTypes = [
    "image/jpeg",
    "image/png",
    "image/webp",
    "application/pdf"
  ];

  if (!allowedTypes.includes(file.mimeType)) {
    throw new Error(
      "ไฟล์ " + filename +
      " ต้องเป็น JPG, PNG, WEBP หรือ PDF (ไม่รองรับ HEIC)"
    );
  }

  const bytes = Utilities.base64Decode(file.data);

  if (bytes.length > CONFIG.MAX_FILE_SIZE) {
    throw new Error("ไฟล์ " + filename + " ต้องมีขนาดไม่เกิน 5 MB");
  }

  let extension = ".jpg";

  if (file.mimeType === "image/png") extension = ".png";
  if (file.mimeType === "image/webp") extension = ".webp";
  if (file.mimeType === "application/pdf") extension = ".pdf";

  const blob = Utilities.newBlob(
    bytes,
    file.mimeType,
    filename + extension
  );

  return folder.createFile(blob);
}

// 9. ส่งใบสมัครนักเรียน
function submitApplication(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let folder = null;

  try {
    if (!data || !data.documents) {
      throw new Error("ไม่พบข้อมูลหรือเอกสารแนบ");
    }

    const sheet = getStudentSheet();

    const id =
      "IW" +
      Utilities.formatDate(
        new Date(),
        CONFIG.TIMEZONE,
        "yyyyMMddHHmmss"
      ) +
      Math.floor(100 + Math.random() * 900);

    // สร้างโฟลเดอร์ประจำใบสมัคร
    folder = getMainFolder().createFolder(id);

    // บันทึกเอกสาร 3 รายการ
    const identity = saveFile(
      data.documents.identity,
      folder,
      "บัตรประชาชนหรือสูติบัตร"
    );

    const house = saveFile(
      data.documents.house,
      folder,
      "ทะเบียนบ้าน"
    );

    const payment = saveFile(
      data.documents.payment,
      folder,
      "สลิปชำระเงิน"
    );

    // เตรียมข้อมูลบันทึกลงชีต
    const row = [
      id,
      new Date(),
      data.prefix || "",
      data.name || "",
      data.birth || "",
      data.studentID || "",
      data.nickname || "",
      data.siblings || "",
      data.order || "",
      data.age || "",
      data.address || "",
      data.grade || "",
      data.program || "",
      data.generalEducation || "",
      data.religiousEducation || "",
      data.oldSchool || "",
      data.fatherName || "",
      data.fatherJob || "",
      data.fatherPhone || "",
      data.fatherID || "",
      data.motherName || "",
      data.motherJob || "",
      data.motherPhone || "",
      data.motherID || "",
      data.emergencyName || "",
      data.emergencyPhone || "",
      data.disease || "",
      (data.conditions || []).join(", "),
      data.stay || "",
      identity.getUrl(),
      house.getUrl(),
      payment.getUrl(),
      "รอตรวจสอบ",
      "",
      folder.getId()
    ];

    sheet.appendRow(row);

    return {
      success: true,
      applicationId: id
    };

  } catch (error) {
    if (folder) {
      try {
        folder.setTrashed(true);
      } catch (e) {}
    }

    throw new Error(error.message || "ไม่สามารถบันทึกใบสมัครได้");

  } finally {
    lock.releaseLock();
  }
}

// 10. เข้าสู่ระบบเจ้าหน้าที่
function adminLogin(password) {
  if (String(password) !== CONFIG.ADMIN_PASSWORD) {
    throw new Error("รหัสผ่านไม่ถูกต้อง");
  }

  const token = Utilities.getUuid();

  CacheService.getScriptCache().put(
    "ADMIN_" + token,
    "TRUE",
    21600
  );

  return token;
}

// 11. ตรวจสอบสิทธิ์เจ้าหน้าที่
function checkAdmin(token) {
  return !!token &&
    CacheService.getScriptCache().get("ADMIN_" + token) === "TRUE";
}

function requireAdmin(token) {
  if (!checkAdmin(token)) {
    throw new Error("กรุณาเข้าสู่ระบบใหม่");
  }
}

// 12. ข้อมูล Dashboard
function getDashboard(token) {
  requireAdmin(token);

  const sheet = getStudentSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      total: 0,
      pending: 0,
      approved: 0,
      rejected: 0,
      students: []
    };
  }

  const values = sheet.getRange(
    2, 1, lastRow - 1, HEADERS.length
  ).getValues();

  const students = values.map(function(r, index) {
    return {
      row: index + 2,
      id: String(r[0] || ""),
      date: r[1] instanceof Date
        ? Utilities.formatDate(r[1], CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm")
        : String(r[1] || ""),
      name: String(r[3] || ""),
      studentID: String(r[5] || ""),
      grade: String(r[11] || ""),
      program: String(r[12] || ""),
      status: String(r[32] || "รอตรวจสอบ"),
      identity: String(r[29] || ""),
      house: String(r[30] || ""),
      payment: String(r[31] || ""),
      note: String(r[33] || ""),
      birth: String(r[4] || ""),
      age: String(r[9] || ""),
      address: String(r[10] || ""),
      fatherName: String(r[16] || ""),
      fatherPhone: String(r[18] || ""),
      motherName: String(r[20] || ""),
      motherPhone: String(r[22] || ""),
      emergencyName: String(r[24] || ""),
      emergencyPhone: String(r[25] || ""),
      disease: String(r[26] || ""),
      conditions: String(r[27] || ""),
      stay: String(r[28] || "")
    };
  }).reverse();

  return {
    total: students.length,
    pending: students.filter(s => s.status === "รอตรวจสอบ").length,
    approved: students.filter(s => s.status === "อนุมัติ").length,
    rejected: students.filter(s => s.status === "ไม่อนุมัติ").length,
    students: students
  };
}

// 13. เปลี่ยนสถานะใบสมัคร
function updateApplication(token, rowNumber, status, note) {
  requireAdmin(token);

  const allowed = [
    "รอตรวจสอบ",
    "อนุมัติ",
    "ไม่อนุมัติ"
  ];

  if (!allowed.includes(status)) {
    throw new Error("สถานะไม่ถูกต้อง");
  }

  const sheet = getStudentSheet();
  const row = Number(rowNumber);

  if (
    !Number.isInteger(row) ||
    row < 2 ||
    row > sheet.getLastRow()
  ) {
    throw new Error("ไม่พบใบสมัคร");
  }

  sheet.getRange(row, 33).setValue(status);
  sheet.getRange(row, 34).setValue(note || "");

  return "บันทึกสถานะเรียบร้อย";
}

// 14. ทดสอบระบบรูปภาพ
function testImages() {
  Logger.log("โลโก้: " + getImageDataFromConfig("logo"));
  Logger.log("QR Code: " + getImageDataFromConfig("qr"));
  return "ตั้งค่ารูปภาพเรียบร้อย";
}
