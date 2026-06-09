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
  let payMethods = []; // คอลัมน์ E
  let destAccounts = []; // คอลัมน์ F
  let edcMachines = []; // คอลัมน์ G
  let senderBanks = []; // คอลัมน์ H (ธนาคารต้นทางผู้โอน/เช็ค)
  let unitMapping = {}; 
  
  for (let i = 1; i < data.length; i++) {
    let inv = data[i][0] ? data[i][0].toString().trim() : "";
    let unit = (data[i].length > 1 && data[i][1]) ? data[i][1].toString().trim() : "";
    let comp = (data[i].length > 2 && data[i][2]) ? data[i][2].toString().trim() : "";
    let sales = (data[i].length > 3 && data[i][3]) ? data[i][3].toString().trim() : "";
    let payM = (data[i].length > 4 && data[i][4]) ? data[i][4].toString().trim() : "";
    let dest = (data[i].length > 5 && data[i][5]) ? data[i][5].toString().trim() : "";
    let edc = (data[i].length > 6 && data[i][6]) ? data[i][6].toString().trim() : "";
    let sBank = (data[i].length > 7 && data[i][7]) ? data[i][7].toString().trim() : "";
    
    if (inv) { invTypes.push(inv); if (unit) unitMapping[inv] = unit; }
    if (unit) unitTypes.push(unit);
    if (comp) companyTypes.push(comp);
    if (sales) salesNames.push(sales);
    if (payM) payMethods.push(payM);
    if (dest) destAccounts.push(dest);
    if (edc) edcMachines.push(edc);
    if (sBank) senderBanks.push(sBank);
  }
  
  return {
    invTypes: [...new Set(invTypes)].filter(String),
    unitTypes: [...new Set(unitTypes)].filter(String),
    companyTypes: [...new Set(companyTypes)].filter(String),
    salesNames: [...new Set(salesNames)].filter(String),
    payMethods: [...new Set(payMethods)].filter(String),
    destAccounts: [...new Set(destAccounts)].filter(String),
    edcMachines: [...new Set(edcMachines)].filter(String),
    senderBanks: [...new Set(senderBanks)].filter(String),
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

// ค้นหาข้อมูลคนเดิม เพื่อดึงประวัติมาแสดง
function searchCitizenId(citizenIdRaw) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  if (!sheet) return { found: false };
  
  const data = sheet.getDataRange().getValues();
  let found = false;
  let count = 0;
  let personData = { totalInvest: 0, totalReturn: 0, netDamage: 0 };
  
  for (let i = 1; i < data.length; i++) {
    let idCol = String(data[i][2]).replace(/-/g, '');

    if (idCol === citizenIdRaw) {
      count++; 
      if (!found) {
        personData.fileNo = data[i][1];
        personData.fullName = data[i][3];
        personData.authPerson = data[i][4];
        personData.contactAddress = data[i][5];
        personData.contactPhone = data[i][6];
        found = true;
      }
      if (data[i][21] !== "" && personData.totalInvest == 0) personData.totalInvest = data[i][21];
      if (data[i][22] !== "" && personData.totalReturn == 0) personData.totalReturn = data[i][22];
      if (data[i][23] !== "" && personData.netDamage == 0) personData.netDamage = data[i][23];
    }
  }
  
  if (found) return { found: true, count: count, data: personData };
  return { found: false };
}

// บันทึก หรือ อัปเดตข้อมูล
function saveData(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    let cleanIdCard = formObj.idCard.toString().replace(/\D/g, '');
    if (!validateThaiIDServer(cleanIdCard)) {
       throw new Error("เลขประจำตัวประชาชนไม่ถูกต้องตามหลักเกณฑ์");
    }
    formObj.idCard = cleanIdCard; 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('Data');
    const settingsSheet = ss.getSheetByName('Settings');
    
    let isUpdate = formObj.id ? true : false;
    let id = isUpdate ? formObj.id : Utilities.getUuid();
    let timestamp = new Date();
    
    // ดึงค่าบัญชีปลายทาง เครื่อง EDC และธนาคาร ใหม่จาก JSON การชำระเงิน
    let newDestAccounts = [];
    let newEdcMachines = [];
    let newSenderBanks = [];
    try {
      let payments = JSON.parse(formObj.paymentDate);
      payments.forEach(p => {
        if (p.method === 'โอนเงิน' && p.details && p.details.destAccount) newDestAccounts.push(p.details.destAccount);
        if (p.method === 'โอนเงิน' && p.details && p.details.senderBank) newSenderBanks.push(p.details.senderBank);
        if (p.method === 'เช็ค' && p.details && p.details.chequeBank) newSenderBanks.push(p.details.chequeBank);
        if (p.method === 'บัตรเครดิต' && p.details && p.details.edc) newEdcMachines.push(p.details.edc);
      });
    } catch(e) {}

    let rowData = [
      id,                     // 0: A
      formObj.fileNo,         // 1: B
      formObj.idCard,         // 2: C
      formObj.fullName,       // 3: D
      formObj.authPerson,     // 4: E
      formObj.contactAddress, // 5: F
      formObj.contactPhone,   // 6: G
      formObj.contractDate,   // 7: H
      formObj.paymentDate,    // 8: I - JSON Payments
      formObj.companyName,    // 9: J
      formObj.invType,        // 10: K
      formObj.unitAmount,     // 11: L
      formObj.unitType,       // 12: M
      formObj.salesName,      // 13: N
      formObj.fullAmount,     // 14: O
      formObj.discount,       // 15: P
      formObj.netAmount,      // 16: Q
      formObj.monthlyReturn,  // 17: R
      formObj.returnPercent,  // 18: S
      formObj.totalReturn,    // 19: T
      formObj.contractYears,  // 20: U
      formObj.actualInvest,   // 21: V
      formObj.actualReturn,   // 22: W
      formObj.netDamage,      // 23: X
      timestamp               // 24: Y
    ];
    
    if (isUpdate) {
      let lastRow = dataSheet.getLastRow();
      if (lastRow > 0) {
        let idColumn = dataSheet.getRange(1, 1, lastRow, 1).getValues().flat();
        let rowIndex = idColumn.indexOf(id) + 1; 
        
        if (rowIndex > 1) { 
          rowData[24] = dataSheet.getRange(rowIndex, 25).getValue(); // เก็บ Timestamp เดิมไว้
          dataSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
        } else {
          throw new Error("ไม่พบข้อมูลที่ต้องการแก้ไข");
        }
      }
    } else {
      dataSheet.appendRow(rowData);
    }
    
    // อัพเดท Settings ถ้ามีค่าใหม่โผล่มา
    updateSettingsIfNew(settingsSheet, formObj.invType, formObj.unitType, formObj.companyName, formObj.salesName, newDestAccounts, newEdcMachines, newSenderBanks);
    return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
    
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// อัพเดทข้อมูลลงตาราง Settings ทีละคอลัมน์
function updateSettingsIfNew(sheet, newInvType, newUnitType, newCompany, newSalesName, newDestAccounts, newEdcMachines, newSenderBanks) {
  if (!sheet) return;
  
  function appendToColumn(colIndex, newValues) {
    if(!newValues || newValues.length === 0) return;
    let colData = sheet.getRange(1, colIndex, sheet.getMaxRows() || 1, 1).getValues().flat().filter(String);
    let toAdd = newValues.filter(v => v.trim() !== '' && !colData.includes(v.trim()));
    if (toAdd.length > 0) {
        let startRow = colData.length + 1;
        let addData = toAdd.map(v => [v.trim()]);
        sheet.getRange(startRow, colIndex, addData.length, 1).setValues(addData);
    }
  }

  if (newInvType) appendToColumn(1, [newInvType]);
  if (newUnitType) appendToColumn(2, [newUnitType]);
  if (newCompany) appendToColumn(3, [newCompany]);
  if (newSalesName) appendToColumn(4, [newSalesName]);
  if (newDestAccounts && newDestAccounts.length > 0) appendToColumn(6, newDestAccounts); // Col F
  if (newEdcMachines && newEdcMachines.length > 0) appendToColumn(7, newEdcMachines); // Col G
  if (newSenderBanks && newSenderBanks.length > 0) appendToColumn(8, newSenderBanks); // Col H
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data');
  if(!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length > 0) data.shift(); 
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
