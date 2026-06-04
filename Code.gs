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
  let unitMapping = {}; // ตัวแปรสำหรับจับคู่ ประเภทการลงทุน -> หน่วยนับ
  
  for (let i = 1; i < data.length; i++) {
    let inv = data[i][0] ? data[i][0].toString().trim() : "";
    let unit = (data[i].length > 1 && data[i][1]) ? data[i][1].toString().trim() : "";
    let comp = (data[i].length > 2 && data[i][2]) ? data[i][2].toString().trim() : "";
    
    if (inv) {
      invTypes.push(inv);
      if (unit) {
        unitMapping[inv] = unit; // ทำการจับคู่ในระบบหลังบ้าน
      }
    }
    if (unit) unitTypes.push(unit);
    if (comp) companyTypes.push(comp);
  }
  
  return {
    invTypes: [...new Set(invTypes)].filter(String),
    unitTypes: [...new Set(unitTypes)].filter(String),
    companyTypes: [...new Set(companyTypes)].filter(String),
    unitMapping: unitMapping // ส่งคู่จับคู่ไปให้ฝั่ง HTML ใช้งาน
  };
}

// บันทึก หรือ อัปเดตข้อมูล
function saveData(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('Data');
    const settingsSheet = ss.getSheetByName('Settings');
    
    let isUpdate = formObj.id ? true : false;
    let id = isUpdate ? formObj.id : Utilities.getUuid();
    let timestamp = new Date();
    
// โครงสร้างข้อมูล 18 คอลัมน์ (ดัชนี 0-17)
    let rowData = [
      id,                     // 0: A
      formObj.fileNo,         // 1: B
      formObj.idCard,         // 2: C
      formObj.fullName,       // 3: D
      formObj.phone,          // 4: E
      formObj.contractDate,   // 5: F
      formObj.companyName,    // 6: G
      formObj.invType,        // 7: H
      formObj.unitAmount,     // 8: I
      formObj.unitType,       // 9: J
      formObj.fullAmount,     // 10: K
      formObj.discount,       // 11: L
      formObj.netAmount,      // 12: M
      formObj.monthlyReturn,  // 13: N
      formObj.returnPercent,  // 14: O (ตำแหน่งใหม่)
      formObj.totalReturn,    // 15: P (ตำแหน่งใหม่)
      formObj.contractYears,  // 16: Q (ตำแหน่งใหม่)
      timestamp               // 17: R
    ];
    
    if (isUpdate) {
      const data = dataSheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          rowIndex = i + 1; 
          break;
        }
      }
      if (rowIndex > -1) {
        rowData[17] = data[rowIndex - 1][17]; // รักษา Timestamp เดิม
        dataSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        throw new Error("ไม่พบข้อมูลที่ต้องการแก้ไข");
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

// เช็คและเพิ่ม Dropdown ใหม่
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

// ดึงข้อมูลทั้งหมดไปแสดงในตาราง
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data');
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length > 0) data.shift();
  return data;
}

// ลบข้อมูล
function deleteData(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Data');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
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

// ดึงข้อมูลสรุปสำหรับ Dashboard โดยดึงคำระบุหน่วยนับไปด้วย
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data');
  const data = sheet.getDataRange().getValues();
  
  let totalContracts = 0;
  let totalNetAmount = 0;
  let totalUnits = 0;
  let uniqueVictims = new Set();
  let typeSummary = {};
  
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      if (!row[0]) continue; // ข้ามแถวที่ไม่มี ID
      
      let idCard = row[2]; // C
      let invType = row[7] || 'ไม่ระบุ'; // H
      let unitAmount = parseFloat(row[8]) || 0; // I
      let unitText = row[9] || ''; // J: ดึงหน่วยนับ (เช่น ตู้, ลิตร)
      let netAmount = parseFloat(row[12]) || 0; // M
      
      totalContracts++;
      totalNetAmount += netAmount;
      totalUnits += unitAmount;
      if (idCard) uniqueVictims.add(idCard.toString().trim());
      
      if (!typeSummary[invType]) {
        typeSummary[invType] = { count: 0, units: 0, sum: 0, unitText: '' };
      }
      typeSummary[invType].count += 1;
      typeSummary[invType].units += unitAmount;
      typeSummary[invType].sum += netAmount;
      typeSummary[invType].unitText = unitText; // บันทึกข้อความหน่วยนับไว้ใน Object
    }
  }
  
  return {
    totalVictims: uniqueVictims.size,
    totalContracts: totalContracts,
    totalNetAmount: totalNetAmount,
    totalUnits: totalUnits,
    summaryByType: typeSummary
  };
}
