const SPREADSHEET_ID = '1BkhC_02odW8OINve6c3Ec4QI4cr_DEQvFGCVWrgebfg'; // ตรวจสอบให้แน่ใจว่าใส่ ID Sheet ถูกต้อง

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAS Defect Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// 1. อัปเดตโครงสร้าง Sheet DEFECT ให้ตรงกับ Col A ถึง O (คอลัมน์ C,D ขออนุญาตเก็บเป็น TargetDate เดิมไว้เพื่อไม่ให้กระทบระบบอื่น)
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsInfo = {
    'JOB': ['JobID', 'Site', 'Owner', 'OwnerCompany', 'Staff', 'ReplyDueDate', 'Remark', 'Timestamp'],
    'TASK': ['TaskID', 'JobID', 'Scope', 'Building', 'Unit', 'Status', 'CustomerName', 'TargetFixDate', 'ActualStartDate', 'ActualEndDate', 'Duration', 'Remark', 'Timestamp'],
    'DEFECT': [
      'DefectID', 'TaskID', 'TargetStartDate', 'TargetEndDate', 'Status', 'MainCategory', 
      'SubCategory', 'Description', 'Major', 'Team', 'ImgUnit', 'ImgBefore', 'ImgDuring', 'ImgAfter', 'Timestamp', 
      'VOSteps', 'ActualStartDate', 'ActualEndDate', 'Remark' // นำฟิลด์ที่เหลือไปต่อท้าย
    ]
  };
  Object.keys(sheetsInfo).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheetsInfo[name]);
      sheet.getRange(1, 1, 1, sheetsInfo[name].length).setFontWeight("bold").setBackground("#f3f4f6");
    }
  });
}

// 2. อัปเดต getAllData ให้ดึงข้อมูล Major และ รูปภาพออกมาด้วย
function getAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const getSheetData = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    return data.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
  };

  const jobs = getSheetData('JOB');
  const tasks = getSheetData('TASK');
  const defects = getSheetData('DEFECT');

  const structuredJobs = jobs.map(job => {
    const jobTasks = tasks.filter(t => t.JobID === job.JobID).map(task => {
      const taskDefects = defects.filter(d => d.TaskID === task.TaskID).map(def => ({
          id: def.DefectID || def['DefectID'],
          mainCategory: def.MainCategory || def['ลักษณะงานหลัก'],
          subCategory: def.SubCategory || def['ลักษณะงานรอง'],
          description: def.Description || def['รายละเอียด'],
          major: def.Major || def['Major'], 
          team: def.Team || def['ทีมเข้าแก้ไข'],
          imgUnit: def.ImgUnit || def['รูปภาพเลขยูนิต'], 
          imgBefore: def.ImgBefore || def['รูปภาพก่อนแก้ไข'],
          imgDuring: def.ImgDuring || def['รูปภาพระหว่างแก้ไข'],
          imgAfter: def.ImgAfter || def['รูปภาพหลังแก้ไข'],
          status: def.Status || def['DefectStatus'] || def['สถานะ defect'],
          remark: def.Remark || def['หมายเหตุ'] || ''
        }));

      return {
        id: task.TaskID,
        scope: task.Scope,
        building: task.Building,
        unit: task.Unit,
        status: task.Status,
        customerName: task.CustomerName,
        targetFixDate: task.TargetFixDate,
        actualStartDate: task.ActualStartDate,
        actualEndDate: task.ActualEndDate,
        duration: task.Duration,
        remark: task.Remark,
        defects: taskDefects
      };
    });

    return {
      id: job.JobID,
      site: job.Site,
      owner: job.Owner,
      ownerCompany: job.OwnerCompany,
      staff: job.Staff,
      replyDueDate: job.ReplyDueDate,
      remark: job.Remark,
      tasks: jobTasks
    };
  });

  return JSON.stringify(structuredJobs);
}
// ฟังก์ชันสำหรับสร้างใบงานหลัก (JOB)
function addJob(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const newId = 'JOB-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  
  sheet.appendRow([
    newId,                        // JobID
    formData.site || '',          // Site
    formData.owner || '',         // Owner
    formData.ownerCompany || '',  // OwnerCompany
    formData.staff || '',         // Staff
    formData.replyDueDate || '',  // ReplyDueDate
    formData.remark || '',        // Remark
    new Date()                    // Timestamp
  ]);
  
  return newId;
}

// ฟังก์ชันสำหรับสร้างใบงานย่อย (TASK)
function addTask(jobId, formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('TASK');
  const newId = 'TSK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  
  sheet.appendRow([
    newId,                        // TaskID
    jobId,                        // JobID
    formData.scope || 'SAS',      // Scope
    formData.building || '',      // Building
    formData.unit || '',          // Unit
    'รอดำเนินการ',                  // Status
    formData.customerName || '',  // CustomerName
    formData.targetFixDate || '', // TargetFixDate
    '',                           // ActualStartDate
    '',                           // ActualEndDate
    '',                           // Duration
    formData.remark || '',        // Remark
    new Date()                    // Timestamp
  ]);
  
  return newId;
}

function addDefect(taskId, defectData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const newId = 'DEF-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');

  // สร้าง Array ขนาด 15 ช่อง (Index 0 ถึง 14) เพื่อล็อคคอลัมน์ A ถึง O ให้ตรงเป๊ะ
  const rowData = new Array(15).fill('');
  
  rowData[0] = newId;                        // Col A: DefectID
  rowData[1] = taskId;                       // Col B: TaskID
  // Col C (joB) และ D (Scope) ปล่อยว่างไว้ตาม Prompt หรือรอรับค่าจาก Task
  rowData[4] = 'ยังไม่แก้ไข';                  // Col E: สถานะ defect (DefectStatus)
  rowData[5] = defectData.mainCategory;      // Col F: ลักษณะงานหลัก
  rowData[6] = defectData.subCategory;       // Col G: ลักษณะงานรอง
  rowData[7] = defectData.description;       // Col H: รายละเอียด
  rowData[8] = defectData.major;             // Col I: Major (ใช่/ไม่ใช่)
  rowData[9] = defectData.team;              // Col J: ทีมเข้าแก้ไข
  // Col K - N ปล่อยว่างไว้รอระบบอัปโหลดรูปภาพทีหลัง
  rowData[10] = '';                          // Col K: รูปภาพเลขยูนิต
  rowData[11] = '';                          // Col L: รูปภาพก่อนแก้ไข
  rowData[12] = '';                          // Col M: รูปภาพระหว่างแก้ไข
  rowData[13] = '';                          // Col N: รูปภาพหลังแก้ไข
  rowData[14] = new Date();                  // Col O: Timestamp

  sheet.appendRow(rowData);
  return newId;
}

// 4. สร้างฟังก์ชันใหม่สำหรับบันทึกรูปภาพ 4 รูป
function uploadDefectImages(defectId, imagesPayload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  
  // หาแถวของ Defect นี้
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return "Defect not found";

  // ฟังก์ชันย่อยแปลง Base64 และสร้างไฟล์ใน Google Drive
  function uploadBase64(base64Str, filename) {
    if (!base64Str) return '';
    if (base64Str.startsWith('http')) return base64Str; // กรณีเป็น URL เดิมอยู่แล้ว
    try {
      const splitBase = base64Str.split(',');
      const contentType = splitBase[0].split(';')[0].replace('data:', '');
      const byteCharacters = Utilities.base64Decode(splitBase[1]);
      const blob = Utilities.newBlob(byteCharacters, contentType, filename);
      const file = DriveApp.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getUrl();
    } catch (e) {
      return '';
    }
  }

  const ts = new Date().getTime();
  
  // บันทึกไฟล์และรับ URL (ถ้าไม่ได้อัปโหลดใหม่ ให้ใช้ค่าเดิมในคอลัมน์ K,L,M,N)
  const imgUnitUrl = imagesPayload.imgUnit ? uploadBase64(imagesPayload.imgUnit, `Unit_${defectId}_${ts}`) : data[rowIndex-1][10];
  const imgBeforeUrl = imagesPayload.imgBefore ? uploadBase64(imagesPayload.imgBefore, `Before_${defectId}_${ts}`) : data[rowIndex-1][11];
  const imgDuringUrl = imagesPayload.imgDuring ? uploadBase64(imagesPayload.imgDuring, `During_${defectId}_${ts}`) : data[rowIndex-1][12];
  const imgAfterUrl = imagesPayload.imgAfter ? uploadBase64(imagesPayload.imgAfter, `After_${defectId}_${ts}`) : data[rowIndex-1][13];

  // อัปเดตข้อมูลลงชีต Col K(11), L(12), M(13), N(14)
  sheet.getRange(rowIndex, 11).setValue(imgUnitUrl);
  sheet.getRange(rowIndex, 12).setValue(imgBeforeUrl);
  sheet.getRange(rowIndex, 13).setValue(imgDuringUrl);
  sheet.getRange(rowIndex, 14).setValue(imgAfterUrl);

  return "Success";
}
