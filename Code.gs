const SPREADSHEET_ID = '1BkhC_02odW8OINve6c3Ec4QI4cr_DEQvFGCVWrgebfg';
const IMAGE_FOLDER_ID = '1pD5dfsyjrtoy7k3IUGaCGPMo6-SiCJPO'; // <--- เพิ่มบรรทัดนี้

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAS Defect Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. อัปเดตโครงสร้าง Sheet DEFECT ให้ตรงกับ Col A ถึง O
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsInfo = {
    'JOB': ['JobID', 'Site', 'Owner', 'OwnerCompany', 'Staff', 'ReplyDueDate', 'Remark', 'Timestamp', 'Status'],
    'TASK': ['TaskID', 'JobID', 'Scope', 'Building', 'Unit', 'Status', 'CustomerName', 'TargetFixDate', 'ActualStartDate', 'ActualEndDate', 'Duration', 'Remark', 'Timestamp'],
    'DEFECT': [
      'DefectID', 'TaskID', 'TargetStartDate', 'TargetEndDate', 'Status', 'MainCategory', 
      'SubCategory', 'Description', 'Major', 'Team', 'ImgUnit', 'ImgBefore', 'ImgDuring', 'ImgAfter', 'Timestamp', 
      'VOSteps', 'ActualStartDate', 'ActualEndDate', 'Remark' 
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

// 2. ฟังก์ชันดึงข้อมูลทั้งหมด
function getAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // ฟังก์ชันย่อยสำหรับแปลงข้อมูลจาก Sheet เป็น Object
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
        status: task.Status || task['TaskStatus'] || task['สถานะ'] || task['สถานะใบงาน'] || 'รอดำเนินการ',
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
      status: job.Status || job['JobStatus'] || job['สถานะ'] || job['สถานะใบงานหลัก'] || 'รอดำเนินการ',
      tasks: jobTasks
    };
  });

  return JSON.stringify(structuredJobs);
}

function addJob(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const newId = 'JOB-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  
  sheet.appendRow([
    newId,                        // Col A
    formData.site || '',          // Col B
    formData.owner || '',         // Col C
    formData.ownerCompany || '',  // Col D
    formData.staff || '',         // Col E
    formData.replyDueDate || '',  // Col F
    formData.remark || '',        // Col G
    new Date(),                   // Col H: Timestamp
    'รอดำเนินการ'                   // Col I: Status
  ]);
  return newId;
}

function addTask(jobId, formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('TASK');
  
  const newId = 'TSK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  const historyLog = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy, HH:mm:ss');
  
  let durationDate = '';
  if (formData.targetFixDate) {
    const parts = formData.targetFixDate.split('-'); 
    if (parts.length === 3) {
      const dateObj = new Date(parts[0], parts[1] - 1, parts[2]);
      dateObj.setDate(dateObj.getDate() + 14);
      durationDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd'); 
    }
  }

  sheet.appendRow([
    newId,                        // Col A: TaskID
    jobId,                        // Col B: JobID
    formData.scope || 'SAS',      // Col C: Scope
    formData.building || '',      // Col D: Building
    formData.unit || '',          // Col E: Unit
    'รอดำเนินการ',                  // Col F: Status
    formData.customerName || '',  // Col G: ชื่อลูกค้า
    formData.targetFixDate || '', // Col H: กำหนดวันเข้าแก้ไข
    durationDate,                 // Col I: Duration
    formData.remark || '',        // Col J: รายละเอียด
    historyLog                    // Col K: ประวัติ
  ]);
  
  return newId;
}

function addDefect(taskId, defectData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const newId = 'DEF-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');

  function uploadBase64(base64Str, filename) {
    if (!base64Str) return '';
    try {
      const splitBase = base64Str.split(',');
      const contentType = splitBase[0].split(';')[0].replace('data:', '');
      const byteCharacters = Utilities.base64Decode(splitBase[1]);
      const blob = Utilities.newBlob(byteCharacters, contentType, filename);
      const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return "https://drive.google.com/uc?export=view&id=" + file.getId();
    } catch (e) {
      return '';
    }
  }

  let imgBeforeUrl = '';
  if (defectData.imgBefore) {
    const ts = new Date().getTime();
    imgBeforeUrl = uploadBase64(defectData.imgBefore, `Before_${newId}_${ts}`);
  }

  const rowData = new Array(15).fill('');
  
  rowData[0] = newId;                        // Col A: DefectID
  rowData[1] = taskId;                       // Col B: TaskID
  rowData[4] = 'ยังไม่แก้ไข';                  // Col E: Status
  rowData[5] = defectData.mainCategory;      // Col F: ลักษณะงานหลัก
  rowData[6] = defectData.subCategory;       // Col G: ลักษณะงานรอง
  rowData[7] = defectData.description;       // Col H: รายละเอียด
  rowData[8] = defectData.major;             // Col I: Major
  rowData[9] = defectData.team;              // Col J: ทีมเข้าแก้ไข
  rowData[10] = '';                          // Col K: รูปภาพเลขยูนิต
  rowData[11] = imgBeforeUrl;                // Col L: รูปภาพก่อนแก้ไข
  rowData[12] = '';                          // Col M: รูปภาพระหว่างแก้ไข
  rowData[13] = '';                          // Col N: รูปภาพหลังแก้ไข
  rowData[14] = new Date();                  // Col O: Timestamp

  sheet.appendRow(rowData);
  return newId;
}

// ยังคงเก็บฟังก์ชันเดิมไว้เผื่อกรณีต้องการใช้ (ไม่กระทบการทำงานใหม่)
function uploadDefectImages(defectId, imagesPayload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return "Defect not found";

  function uploadBase64(base64Str, filename) {
    if (!base64Str) return '';
    if (base64Str.startsWith('http')) return base64Str; 
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
  
  const imgUnitUrl = imagesPayload.imgUnit ? uploadBase64(imagesPayload.imgUnit, `Unit_${defectId}_${ts}`) : data[rowIndex-1][10];
  const imgBeforeUrl = imagesPayload.imgBefore ? uploadBase64(imagesPayload.imgBefore, `Before_${defectId}_${ts}`) : data[rowIndex-1][11];
  const imgDuringUrl = imagesPayload.imgDuring ? uploadBase64(imagesPayload.imgDuring, `During_${defectId}_${ts}`) : data[rowIndex-1][12];
  const imgAfterUrl = imagesPayload.imgAfter ? uploadBase64(imagesPayload.imgAfter, `After_${defectId}_${ts}`) : data[rowIndex-1][13];

  sheet.getRange(rowIndex, 11).setValue(imgUnitUrl);
  sheet.getRange(rowIndex, 12).setValue(imgBeforeUrl);
  sheet.getRange(rowIndex, 13).setValue(imgDuringUrl);
  sheet.getRange(rowIndex, 14).setValue(imgAfterUrl);

  return "Success";
}

function updateJob(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === formData.id) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error("ไม่พบ JobID ที่ต้องการแก้ไขในฐานข้อมูล");
  }

  sheet.getRange(rowIndex, 2).setValue(formData.site || '');
  sheet.getRange(rowIndex, 3).setValue(formData.owner || '');
  sheet.getRange(rowIndex, 4).setValue(formData.ownerCompany || '');
  sheet.getRange(rowIndex, 5).setValue(formData.staff || '');
  sheet.getRange(rowIndex, 6).setValue(formData.replyDueDate || '');
  sheet.getRange(rowIndex, 7).setValue(formData.remark || '');

  return "Update Success";
}

function getMasterData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let result = { sites: [], owners: [] };

  const projectSheet = ss.getSheetByName('Project');
  if (projectSheet) {
    const pLastRow = projectSheet.getLastRow();
    if (pLastRow >= 2) {
      const pData = projectSheet.getRange(2, 2, pLastRow - 1, 1).getDisplayValues();
      const sites = pData.map(r => r[0]).filter(s => s !== '');
      result.sites = [...new Set(sites)];
    }
  }

  const ownerSheet = ss.getSheetByName('Owner');
  if (ownerSheet) {
    const oLastRow = ownerSheet.getLastRow();
    if (oLastRow >= 2) {
      const oData = ownerSheet.getRange(2, 2, oLastRow - 1, 3).getDisplayValues();
      result.owners = oData
        .filter(row => row[2] !== '') 
        .map(row => ({
          ownerCompany: row[0], 
          site: row[1],         
          owner: row[2]         
        }));
    }
  }

  return JSON.stringify(result);
}

function deleteJob(jobId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  throw new Error("ไม่พบ JobID ที่ต้องการลบ");
}

function deleteTask(taskId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('TASK');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  throw new Error("ไม่พบ TaskID ที่ต้องการลบ");
}

function deleteDefect(defectId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  throw new Error("ไม่พบ DefectID ที่ต้องการลบ");
}

// --- ฟังก์ชันเปลี่ยนสถานะ ---
function updateTaskStatusAndJob(taskId, newStatus) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('TASK');
  const taskData = taskSheet.getDataRange().getValues();

  let jobId = '';
  let taskRowIndex = -1;

  for (let i = 1; i < taskData.length; i++) {
    if (taskData[i][0] === taskId) {
      taskRowIndex = i + 1;
      jobId = taskData[i][1]; // ดึง JobID ในคอลัมน์ B (Index 1)
      taskData[i][5] = newStatus; // อัปเดตสถานะจำลองใน Array เพื่อใช้เช็คเงื่อนไขทันที
      break;
    }
  }
  
  if (taskRowIndex !== -1) {
    // 1. อัปเดตสถานะของ Task ในคอลัมน์ F (ตำแหน่งที่ 6)
    taskSheet.getRange(taskRowIndex, 6).setValue(newStatus);

    // 2. เงื่อนไขอัปเดต Job หลัก
    if (jobId) {
       const jobSheet = ss.getSheetByName('JOB');
       const jobData = jobSheet.getDataRange().getValues();
       
       // เช็คสถานะของทุกใบงานย่อยภายใต้ Job เดียวกัน
       let allTasksFinished = true;
       for (let i = 1; i < taskData.length; i++) {
         if (taskData[i][1] === jobId) {
           const status = taskData[i][5];
           // ถ้ายังมีงานที่ 'รอดำเนินการ', 'Active' หรือไม่มีสถานะ ถือว่างานหลักยังไม่จบ
           if (status === 'รอดำเนินการ' || status === 'Active' || status === '') {
             allTasksFinished = false;
             break;
           }
         }
       }

       let jobRowIndex = -1;
       for (let j = 1; j < jobData.length; j++) {
         if (jobData[j][0] === jobId) {
           jobRowIndex = j + 1;
           break;
         }
       }

       if (jobRowIndex !== -1) {
         if (allTasksFinished) {
           // เงื่อนไขใหม่: ถ้าทุกใบงานย่อยไม่มี รอดำเนินการ/Active แล้ว ให้ปิดใบงานหลัก
           jobSheet.getRange(jobRowIndex, 9).setValue('Closed');
         } else if (newStatus === 'Active') {
           // เงื่อนไขเดิม: ถ้ามีการเปลี่ยน Task เป็น Active ให้ Job หลักเป็น Active
           if (jobData[jobRowIndex - 1][8] !== 'Active') { 
             jobSheet.getRange(jobRowIndex, 9).setValue('Active');
           }
         }
       }
    }
    
    // บังคับบันทึกข้อมูลทันทีก่อนแจ้ง Frontend ว่าสำเร็จ
    SpreadsheetApp.flush();
    return "Success";
  }
  throw new Error("ไม่พบข้อมูลใบงานย่อยที่ต้องการเปลี่ยนสถานะ");
}

// --- ฟังก์ชันอัปโหลดรูปภาพทีละรูป (NEW) ---
function uploadSingleDefectImage(defectId, field, base64Str) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Defect not found");

  if (!base64Str) return '';
  if (base64Str.startsWith('http')) return base64Str; 

  try {
    const splitBase = base64Str.split(',');
    const contentType = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    
    const ts = new Date().getTime();
    const filename = `${field}_${defectId}_${ts}`;
    const blob = Utilities.newBlob(byteCharacters, contentType, filename);
    
    // --- ส่วนที่แก้ไข: เล็งเป้าหมายไปที่โฟลเดอร์ ---
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    const file = folder.createFile(blob);
    // ----------------------------------------
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
    
    const colMap = { 'imgUnit': 11, 'imgBefore': 12, 'imgDuring': 13, 'imgAfter': 14 };
    if (colMap[field]) {
      sheet.getRange(rowIndex, colMap[field]).setValue(url);
    }
    return url;
  } catch (e) {
    throw new Error('Upload failed: ' + e.toString());
  }
}

// --- ฟังก์ชันอัปเดตสถานะ Defect (NEW) ---
function updateDefectStatus(defectId, status) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      // คอลัมน์ Status ของ DEFECT อยู่ที่คอลัมน์ E (ตำแหน่งที่ 5)
      sheet.getRange(i + 1, 5).setValue(status);
      return "Success";
    }
  }
  throw new Error("ไม่พบข้อมูล DefectID ที่ต้องการเปลี่ยนสถานะ");
}
