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
  let salesNames = [];
  let unitMapping = {}; 
  
  for (let i = 1; i < data.length; i++) {
    let inv = data[i][0] ? data[i][0].toString().trim() : "";
    let unit = (data[i].length > 1 && data[i][1]) ? data[i][1].toString().trim() : "";
    let comp = (data[i].length > 2 && data[i][2]) ? data[i][2].toString().trim() : "";
    let sales = (data[i].length > 3 && data[i][3]) ? data[i][3].toString().trim() : "";
    
    if (inv) {
      invTypes.push(inv);
      if (unit) {
        unitMapping[inv] = unit; 
      }
    }
    if (unit) unitTypes.push(unit);
    if (comp) companyTypes.push(comp);
    if (sales) salesNames.push(sales);
  }
  
  return {
    invTypes: [...new Set(invTypes)].filter(String),
    unitTypes: [...new Set(unitTypes)].filter(String),
    companyTypes: [...new Set(companyTypes)].filter(String),
    salesNames: [...new Set(salesNames)].filter(String),
    unitMapping: unitMapping 
  };
}

// ฟังก์ชันตรวจสอบบัตรประชาชนฝั่ง Server
function validateThaiIDServer(id) {
  if (!/^[0-9]{13}$/.test(id)) return false;
  let sum = 0;
  for (let i = 0; i < 12; i++) {
    sum += parseInt(id.charAt(i)) * (13 - i);
  }
  let checkDigit = (11 - (sum % 11)) % 10;
  return checkDigit === parseInt(id.charAt(12));
}

// ----------------------------------------------------
// ฟังก์ชันค้นหาข้อมูลคนเดิม เพื่อดึงประวัติมาแสดง
// ----------------------------------------------------
function searchCitizenId(citizenIdRaw) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  if (!sheet) return { found: false };
  
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  let count = 0;
  let personData = {
    totalInvest: 0,
    totalReturn: 0,
    netDamage: 0
  };
  
  // Index: 1=แฟ้มที่, 2=เลขปชช, 3=ชื่อ, 4=ผู้รับมอบอำนาจ, 19=Invest, 20=Return, 21=Damage (มีการขยับ Index แล้ว)
  for (let i = 1; i < data.length; i++) {
    let idCol = String(data[i][2]).replace(/-/g, '');

    if (idCol === citizenIdRaw) {
      count++; // นับจำนวนสัญญาที่มีในระบบ
      
      // ดึงข้อมูลส่วนตัว (ดึงจากอันแรกที่เจอ)
      if (!found) {
        personData.fileNo = data[i][1];
        personData.fullName = data[i][3];
        personData.phone = data[i][4];
        found = true;
      }
      
      // ดึงยอดยกมา (มองหาแถวที่มีค่ายอดรวม เนื่องจากแถวใหม่ๆ จะถูกบันทึกเป็นค่าว่าง)
      if (data[i][19] !== "" && personData.totalInvest == 0) personData.totalInvest = data[i][19];
      if (data[i][20] !== "" && personData.totalReturn == 0) personData.totalReturn = data[i][20];
      if (data[i][21] !== "" && personData.netDamage == 0) personData.netDamage = data[i][21];
    }
  }
  
  if (found) {
    return { found: true, count: count, data: personData };
  }
  return { found: false };
}

// บันทึก หรือ อัปเดตข้อมูล (รองรับ 23 คอลัมน์)
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
    
    // โครงสร้างข้อมูล 23 คอลัมน์ (ดัชนี 0-22) 
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
      formObj.salesName,      // 11: L - รายชื่อเซลล์ (NEW)
      formObj.fullAmount,     // 12: M - มูลค่าเต็มตามสัญญา (SHIFTED)
      formObj.discount,       // 13: N - ส่วนลดเงินลงทุน
      formObj.netAmount,      // 14: O - ยอดเงินลงทุนสุทธิ
      formObj.monthlyReturn,  // 15: P - ผลตอบแทน/เดือน
      formObj.returnPercent,  // 16: Q - ผลตอบแทนเพิ่ม (%)
      formObj.totalReturn,    // 17: R - รวมผลตอบแทนที่ได้จริง/เดือน
      formObj.contractYears,  // 18: S - อายุสัญญา (ปี)
      formObj.actualInvest,   // 19: T - เงินลงทุนที่จ่ายจริงรวมทั้งหมด (ตามบัญชีทรัพย์) (บาท)
      formObj.actualReturn,   // 20: U - ผลตอบแทนที่ได้รับคืนจริงรวมทั้งหมด (ตามบัญชีทรัพย์) (บาท)
      formObj.netDamage,      // 21: V - มูลค่าความเสียหายสุทธิรายบุคคล (ตามบัญชีทรัพย์) (บาท)
      timestamp               // 22: W - วันที่กรอกข้อมูล (Timestamp)
    ];
    
    if (isUpdate) {
      let lastRow = dataSheet.getLastRow();
      if (lastRow > 0) {
        let idColumn = dataSheet.getRange(1, 1, lastRow, 1).getValues().flat();
        let rowIndex = idColumn.indexOf(id) + 1; 
        
        if (rowIndex > 1) { 
          rowData[22] = dataSheet.getRange(rowIndex, 23).getValue(); // รักษา Timestamp เดิม (คอลัมน์ 23)
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
    
    updateSettingsIfNew(settingsSheet, formObj.invType, formObj.unitType, formObj.companyName, formObj.salesName);
    return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
    
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateSettingsIfNew(sheet, newInvType, newUnitType, newCompany, newSalesName) {
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  let existingInv = [];
  let existingUnit = [];
  let existingCompany = [];
  let existingSales = [];
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0]) existingInv.push(data[i][0].toString().trim());
    if(data[i].length > 1 && data[i][1]) existingUnit.push(data[i][1].toString().trim());
    if(data[i].length > 2 && data[i][2]) existingCompany.push(data[i][2].toString().trim());
    if(data[i].length > 3 && data[i][3]) existingSales.push(data[i][3].toString().trim());
  }
  
  let nextInvRow = existingInv.length + 2;
  let nextUnitRow = existingUnit.length + 2;
  let nextCompanyRow = existingCompany.length + 2;
  let nextSalesRow = existingSales.length + 2;
  
  if (newInvType && !existingInv.includes(newInvType.trim())) {
    sheet.getRange(nextInvRow, 1).setValue(newInvType.trim());
  }
  if (newUnitType && !existingUnit.includes(newUnitType.trim())) {
    sheet.getRange(nextUnitRow, 2).setValue(newUnitType.trim());
  }
  if (newCompany && !existingCompany.includes(newCompany.trim())) {
    sheet.getRange(nextCompanyRow, 3).setValue(newCompany.trim());
  }
  if (newSalesName && !existingSales.includes(newSalesName.trim())) {
    sheet.getRange(nextSalesRow, 4).setValue(newSalesName.trim());
  }
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data');
  if(!sheet) return [];
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
