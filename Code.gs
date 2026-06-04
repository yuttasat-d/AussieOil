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
        unitMapping[inv] = unit; // จับคู่ข้อมูล
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

// บันทึก หรือ อัปเดตข้อมูล (รองรับ 22 คอลัมน์เพื่อความปลอดภัยสูงสุดในการทำสำนวนคดี)
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
    
    // โครงสร้างข้อมูล 22 คอลัมน์ (ดัชนี 0-21) บันทึกลงคอลัมน์ A ถึง V ใน Google Sheets
    let rowData = [
      id,                     // 0: A - ID
      formObj.fileNo,         // 1: B - แฟ้มที่
      formObj.idCard,         // 2: C - เลข ปชช.
      formObj.fullName,       // 3: D - ชื่อผู้เสียหาย
      formObj.phone,          // 4: E - ผู้รับมอบอำนาจ
      formObj.contractDate,   // 5: F - วันที่ทำสัญญา
      formObj.paymentDate,    // 6: G - วันที่โอนเงิน/จ่ายเงิน (ตามสลิป) -> เก็บเป็นข้อมูล JSON string ของรายการแบ่งชำระเงินงวดทั้งหมด
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
      const data = dataSheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          rowIndex = i + 1; 
          break;
        }
      }
      if (rowIndex > -1) {
        rowData[21] = data[rowIndex - 1][21]; // รักษา Timestamp เดิมไว้ที่ Index 21 (Col V)
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

// เช็คและเพิ่ม Dropdown ใหม่ในชีต Settings
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

// ดึงข้อมูลทั้งหมดในชีต Data ไปแสดงในตารางหน้าเว็บ
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data');
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length > 0) data.shift();
  return data;
}

// ลบข้อมูลรายสัญญา
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

// คำนวณสรุปแดชบอร์ด หาผลตอบแทนคืนจริง และมูลค่าความเสียหายสุทธิอย่างเป็นระบบสอบสวน
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data');
  const data = sheet.getDataRange().getValues();
  
  let totalContracts = 0;
  let totalNetAmount = 0;
  let totalUnits = 0;
  let uniqueVictims = new Set();
  let typeSummary = {};
  
  // โครงสร้างสำหรับดักจับยอดเงินคืนและยอดเงินลงทุนรายบุคคล เพื่อวิเคราะห์ไม่ให้บวกซ้ำซ้อน
  let victimActualReturns = {};
  let victimActualInvests = {};
  
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      if (!row[0]) continue;
      
      let idCard = row[2] ? row[2].toString().trim() : '';
      let invType = row[8] || 'ไม่ระบุ'; // ดัชนีเปลี่ยนเป็น 8 จากคอลัมน์ประเภทการลงทุน
      let unitAmount = parseFloat(row[9]) || 0; // ดัชนีเปลี่ยนเป็น 9
      let unitText = row[10] || ''; // ดัชนีเปลี่ยนเป็น 10
      let netAmount = parseFloat(row[13]) || 0; // ดัชนีเปลี่ยนเป็น 13 จากยอดเงินสุทธิ
      let actualInvest = parseFloat(row[18]) || 0; // Col S: Index 18 ยอดเงินลงทุนจริง
      let actualReturn = parseFloat(row[19]) || 0; // Col T: Index 19 ยอดเงินคืนจริง
      
      totalContracts++;
      totalNetAmount += netAmount;
      totalUnits += unitAmount;
      if (idCard) {
        uniqueVictims.add(idCard);
        // บันทึกค่าเงินสะสมแบบกลุ่มรายบุคคล ป้องกันการบวกเบิ้ลสะสมของคนที่มีหลายสัญญา
        if (!victimActualReturns[idCard] || actualReturn > victimActualReturns[idCard]) {
          victimActualReturns[idCard] = actualReturn;
        }
        if (!victimActualInvests[idCard] || actualInvest > victimActualInvests[idCard]) {
          victimActualInvests[idCard] = actualInvest;
        }
      }
      
      if (!typeSummary[invType]) {
        typeSummary[invType] = { count: 0, units: 0, sum: 0, unitText: '' };
      }
      typeSummary[invType].count += 1;
      typeSummary[invType].units += unitAmount;
      typeSummary[invType].sum += netAmount;
      typeSummary[invType].unitText = unitText;
    }
  }
  
  // รวมยอดเงินจริงสะสมของทุกคนแบบไม่ซ้ำตัวบุคคล
  let totalActualReturnSum = 0;
  for (let id in victimActualReturns) {
    totalActualReturnSum += victimActualReturns[id];
  }
  
  let totalActualInvestSum = 0;
  for (let id in victimActualInvests) {
    totalActualInvestSum += victimActualInvests[id];
  }
  
  // ความเสียหายสุทธิรวมคดี = เงินลงทุนทั้งหมดลบด้วยยอดคืนเงินทั้งหมดตาม Statement
  let totalNetDamage = totalActualInvestSum - totalActualReturnSum;
  
  return {
    totalVictims: uniqueVictims.size,
    totalContracts: totalContracts,
    totalNetAmount: totalNetAmount,
    totalUnits: totalUnits,
    totalActualInvest: totalActualInvestSum,
    totalActualReturn: totalActualReturnSum,
    totalNetDamage: totalNetDamage,
    summaryByType: typeSummary
  };
}
