// กำหนดค่าเริ่มต้นและแสดงผลหน้า Web App
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบบันทึกข้อมูลผู้เสียหาย Aussie Oil')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ดึงการตั้งค่าสำหรับ Dropdown และส่งคู่จับคู่สัญญา-หน่วยนับไปหน้าเว็บ
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  
  let invTypes = [];
  let unitTypes = [];
  let companyTypes = [];
  let unitMapping = {}; 
  
  for (let i = 1; i < data.length; i++) {
    let inv = data[i][0] ? data[i][0].toString().trim() : "";
    let unit = (data[i].length > 1 && data[i][1]) ? data[i][1].toString().trim() : "";
    let comp = (data[i].length > 2 && data[i][2]) ? data[i][2].toString().trim() : "";
    
    if (inv) {
      invTypes.push(inv);
      if (unit) {
        unitMapping[inv] = unit; 
      }
    }
    if (unit) unitTypes.push(unit);
    if (comp) companyTypes.push(comp);
  }
  
  return {
    invTypes: [...new Set(invTypes)].filter(String),
    unitTypes: [...new Set(unitTypes)].filter(String),
    companyTypes: [...new Set(companyTypes)].filter(String),
    unitMapping: unitMapping 
  };
}

// ฟังก์ชันตรวจสอบบัตรประชาชนฝั่ง Server ป้องกันการยิง API ขยะเข้ามา
function validateThaiIDServer(id) {
  if (!/^[0-9]{13}$/.test(id)) return false;
  let sum = 0;
  for (let i = 0; i < 12; i++) {
    sum += parseInt(id.charAt(i)) * (13 - i);
  }
  let checkDigit = (11 - (sum % 11)) % 10;
  return checkDigit === parseInt(id.charAt(12));
}

// บันทึก หรือ อัปเดตข้อมูล (รองรับ 22 คอลัมน์)
function saveData(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    // ตรวจสอบเลข ปชช. ก่อนเซฟ
    let cleanIdCard = formObj.idCard.toString().replace(/\D/g, '');
    if (!validateThaiIDServer(cleanIdCard)) {
       throw new Error("เลขประจำตัวประชาชนไม่ถูกต้องตามหลักเกณฑ์");
    }
    formObj.idCard = cleanIdCard; // มั่นใจว่าบันทึกแบบไม่มีขีด
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('Data');
    const settingsSheet = ss.getSheetByName('Settings');
    
    let isUpdate = formObj.id ? true : false;
    let id = isUpdate ? formObj.id : Utilities.getUuid();
    let timestamp = new Date();
    
    // โครงสร้างข้อมูล 22 คอลัมน์ (ดัชนี 0-21) บันทึกลงคอลัมน์ A ถึง V ใน Google Sheets
    let rowData = [
      id,                     // 0: A - ID
      formObj.fileNo,         // 1: B - แฟ้มที่
      formObj.idCard,         // 2: C - เลข ปชช.
      formObj.fullName,       // 3: D - ชื่อผู้เสียหาย
      formObj.phone,          // 4: E - ผู้รับมอบอำนาจ
      formObj.contractDate,   // 5: F - วันที่ทำสัญญา
      formObj.paymentDate,    // 6: G - วันที่โอนเงิน/จ่ายเงิน (JSON string ของสลิป)
      formObj.companyName,    // 7: H - บริษัทที่ทำสัญญา
      formObj.invType,        // 8: I - ประเภทการลงทุน
      formObj.unitAmount,     // 9: J - จำนวน
      formObj.unitType,       // 10: K - หน่วยนับ
      formObj.fullAmount,     // 11: L - มูลค่าเต็มตามสัญญา
      formObj.discount,       // 12: M - ส่วนลดเงินลงทุน
      formObj.netAmount,      // 13: N - ยอดเงินลงทุนสุทธิ
      formObj.monthlyReturn,  // 14: O - ผลตอบแทน/เดือน
      formObj.returnPercent,  // 15: P - ผลตอบแทนเพิ่ม (%)
      formObj.totalReturn,    // 16: Q - รวมผลตอบแทนที่ได้จริง/เดือน
      formObj.contractYears,  // 17: R - อายุสัญญา (ปี)
      formObj.actualInvest,   // 18: S - เงินลงทุนที่จ่ายจริงรวมทั้งหมด (ตามบัญชีทรัพย์) (บาท)
      formObj.actualReturn,   // 19: T - ผลตอบแทนที่ได้รับคืนจริงรวมทั้งหมด (ตามบัญชีทรัพย์) (บาท)
      formObj.netDamage,      // 20: U - มูลค่าความเสียหายสุทธิรายบุคคล (ตามบัญชีทรัพย์) (บาท)
      timestamp               // 21: V - วันที่กรอกข้อมูล (Timestamp)
    ];
    
    if (isUpdate) {
      let lastRow = dataSheet.getLastRow();
      if (lastRow > 0) {
        // ดึงแค่คอลัมน์ A มาเช็ค Index ทำให้ไวกว่าการดึงข้อมูลทั้งตาราง
        let idColumn = dataSheet.getRange(1, 1, lastRow, 1).getValues().flat();
        let rowIndex = idColumn.indexOf(id) + 1; 
        
        if (rowIndex > 1) { // สมมติว่าแถวที่ 1 คือ Header
          rowData[21] = dataSheet.getRange(rowIndex, 22).getValue(); // รักษา Timestamp เดิม
          dataSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
        } else {
          throw new Error("ไม่พบข้อมูลที่ต้องการแก้ไข");
        }
      } else {
        throw new Error("ตารางข้อมูลว่างเปล่า");
      }
    } else {
      dataSheet.appendRow(rowData);
    }
    
    updateSettingsIfNew(settingsSheet, formObj.invType, formObj.unitType, formObj.companyName);
    return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
    
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateSettingsIfNew(sheet, newInvType, newUnitType, newCompany) {
  const data = sheet.getDataRange().getValues();
  let existingInv = [];
  let existingUnit = [];
  let existingCompany = [];
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0]) existingInv.push(data[i][0].toString().trim());
    if(data[i].length > 1 && data[i][1]) existingUnit.push(data[i][1].toString().trim());
    if(data[i].length > 2 && data[i][2]) existingCompany.push(data[i][2].toString().trim());
  }
  
  let nextInvRow = existingInv.length + 2;
  let nextUnitRow = existingUnit.length + 2;
  let nextCompanyRow = existingCompany.length + 2;
  
  if (newInvType && !existingInv.includes(newInvType.trim())) {
    sheet.getRange(nextInvRow, 1).setValue(newInvType.trim());
  }
  if (newUnitType && !existingUnit.includes(newUnitType.trim())) {
    sheet.getRange(nextUnitRow, 2).setValue(newUnitType.trim());
  }
  if (newCompany && !existingCompany.includes(newCompany.trim())) {
    sheet.getRange(nextCompanyRow, 3).setValue(newCompany.trim());
  }
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data');
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length > 0) data.shift(); // ลบ Header ออก
  return data;
}

function deleteData(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Data');
    
    let lastRow = sheet.getLastRow();
    if (lastRow > 0) {
      let idColumn = sheet.getRange(1, 1, lastRow, 1).getValues().flat();
      let rowIndex = idColumn.indexOf(id) + 1;
      
      if (rowIndex > 1) {
        sheet.deleteRow(rowIndex);
        return { status: 'success', message: 'ลบข้อมูลเรียบร้อยแล้ว' };
      }
    }
    return { status: 'error', message: 'ไม่พบข้อมูลที่ต้องการลบ' };
  } catch (e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}
