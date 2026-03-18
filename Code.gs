const SPREADSHEET_ID = '1BkhC_02odW8OINve6c3Ec4QI4cr_DEQvFGCVWrgebfg';
const IMAGE_FOLDER_ID = '1pD5dfsyjrtoy7k3IUGaCGPMo6-SiCJPO'; // <--- เพิ่มบรรทัดนี้

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAS Defect Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsInfo = {
    'JOB': ['JobID', 'Site', 'Owner', 'OwnerCompany', 'Staff', 'ReplyDueDate', 'Remark', 'Timestamp', 'Status'],
    'TASK': ['TaskID', 'JobID', 'Scope', 'Building', 'Unit', 'Status', 'CustomerName', 'TargetFixDate', 'ActualStartDate', 'ActualEndDate', 'Duration', 'Remark', 'Timestamp'],
    'DEFECT': [
      'DefectID', 'TaskID', 'TargetStartDate', 'TargetEndDate', 'Status', 'MainCategory', 
      'SubCategory', 'Description', 'Major', 'Team', 'ImgUnit', 'ImgBefore', 'ImgDuring', 'ImgAfter', 'Timestamp', 
      'VOSteps', 'ActualStartDate', 'ActualEndDate', 'Remark' 
    ],
    // เพิ่ม Sheet User และตั้ง Timestamp ให้อยู่ที่ Col N (ตำแหน่งที่ 14)
    'User': ['UserID', 'Password', '', '', 'Position', '', '', '', '', 'Email', 'Line', 'Phone', '','','Timestamp'] 
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
        if (header) { // กันกรณีหัวตารางว่าง
          obj[header] = row[index];
        }
      });
      // เพิ่มตัวแปร _raw ไว้เก็บข้อมูลแถวดิบๆ สำหรับอ้างอิงด้วย Column (0=A, 1=B, ..., 6=G)
      obj['_raw'] = row; 
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
        // บังคับดึงข้อมูลจาก Column G (Index ที่ 6) โดยตรง
        customerName: task._raw[6] || task.CustomerName || task['ชื่อลูกค้า'] || '', 
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

// --- ฟังก์ชัน Export PDF แผนเข้าแก้ไข ---
function exportTaskPlansToPDF(jobId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const allDataStr = getAllData();
  const allJobs = JSON.parse(allDataStr);
  const job = allJobs.find(j => j.id === jobId);

  if (!job) throw new Error("ไม่พบข้อมูลใบงานหลัก (Job)");
  if (!job.tasks || job.tasks.length === 0) throw new Error("ไม่มีใบงานย่อยให้ Export");

  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  let exportedFiles = [];

  // แก้ปัญหาภาพไม่ขึ้น: ดึงไฟล์จาก Drive แปลงเป็น Base64 เพื่อฝังลงใน PDF โดยตรง
  const getPrintableImgUrl = (url) => {
    if (!url) return '';
    if (url.startsWith('data:image')) return url;
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      try {
        const fileId = match[1];
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        const base64 = Utilities.base64Encode(blob.getBytes());
        const mimeType = blob.getContentType();
        return `data:${mimeType};base64,${base64}`;
      } catch (e) {
        // Fallback กรณีดึงไฟล์ไม่ได้
        return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w500`;
      }
    }
    return url;
  };

  job.tasks.forEach(task => {
    let html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
          body { 
            font-family: 'Sarabun', sans-serif; 
            color: #1e293b; 
            line-height: 1.6; 
            font-size: 14px; 
            margin: 0; 
            padding: 10px;
          }
          .header-title { 
            text-align: center; 
            color: #0f172a; 
            margin-bottom: 25px; 
            font-size: 24px; 
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
          }
          table.header-table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-bottom: 30px; 
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
          }
          table.header-table th, table.header-table td { 
            border: 1px solid #cbd5e1; 
            padding: 10px 12px; 
            text-align: left; 
            vertical-align: top;
          }
          table.header-table th { 
            background-color: #f1f5f9; 
            width: 18%; 
            font-weight: 600; 
            color: #334155;
          }
          table.header-table td {
            width: 32%;
            color: #0f172a;
          }
          
          .section-title {
            font-size: 18px;
            font-weight: 600;
            color: #1e40af;
            border-bottom: 2px solid #93c5fd;
            padding-bottom: 8px;
            margin-bottom: 15px;
          }

          .defect-card { 
            border: 1px solid #e2e8f0; 
            margin-bottom: 20px; 
            padding: 15px; 
            page-break-inside: avoid; 
            border-radius: 8px; 
            background-color: #ffffff;
          }
          .defect-layout { 
            display: table; 
            width: 100%; 
          }
          .img-col { 
            display: table-cell; 
            width: 200px; 
            vertical-align: top; 
            text-align: center; 
            padding-right: 20px; 
            border-right: 1px dashed #cbd5e1; 
          }
          .info-col { 
            display: table-cell; 
            vertical-align: top; 
            padding-left: 20px; 
          }
          
          .img-col img { 
            max-width: 100%; 
            max-height: 200px; 
            border-radius: 6px; 
            border: 1px solid #e2e8f0; 
            padding: 3px;
          }
          .no-img { 
            width: 100%; 
            height: 120px; 
            background: #f8fafc; 
            border: 2px dashed #cbd5e1; 
            border-radius: 6px;
            display: flex; 
            align-items: center; 
            justify-content: center; 
            color: #94a3b8; 
            font-size: 13px; 
            margin-top: 10px;
          }
          
          .meta { margin-bottom: 6px; color: #475569; }
          .meta strong { color: #1e293b; font-weight: 600;}
          .major { color: #dc2626; font-weight: 700; background-color: #fef2f2; padding: 2px 6px; border-radius: 4px; font-size: 12px;}
          
          .desc-box { 
            background-color: #f8fafc; 
            border-left: 4px solid #3b82f6; 
            padding: 12px 15px; 
            margin: 15px 0; 
            font-size: 15px; 
            color: #334155; 
            display: block;
            border-radius: 0 6px 6px 0;
          }
          .desc-box strong { color: #0f172a; display: block; margin-bottom: 4px; font-size: 14px;}

          .date-badge { 
            display: inline-block; 
            background-color: #fffbeb; 
            border: 1px solid #fde68a; 
            color: #b45309; 
            padding: 6px 12px; 
            font-weight: 600; 
            font-size: 13px; 
            border-radius: 6px; 
          }

          /* --- ส่วนลายเซ็น (Signature Section) --- */
          .signature-container {
              width: 100%;
              margin-top: 50px;
              page-break-inside: avoid;
          }
          .signature-table {
              width: 100%;
              border-collapse: collapse;
              text-align: center;
          }
          .signature-table td {
              width: 50%;
              padding: 10px 20px;
              vertical-align: bottom;
          }
          .sign-line {
              border-bottom: 1px dashed #94a3b8;
              width: 70%;
              margin: 40px auto 10px auto;
          }
          .sign-text {
              color: #334155;
              font-size: 14px;
              line-height: 1.5;
          }
          .sign-name {
              font-weight: 600;
              color: #0f172a;
          }
        </style>
      </head>
      <body>
        <div class="header-title">แผนเข้าแก้ไข (Repair Plan)</div>
        <table class="header-table">
          <tr>
            <th>Job ID</th><td>${job.id}</td>
            <th>Task ID</th><td>${task.id}</td>
          </tr>
          <tr>
            <th>Site</th><td>${job.site}</td>
            <th>Scope</th><td>${task.scope}</td>
          </tr>
          <tr>
            <th>Owner / ผู้ดูแล</th><td>${job.owner}</td>
            <th>Building / Unit</th><td>${task.building} - ${task.unit}</td>
          </tr>
          <tr>
            <th>Company</th><td>${job.ownerCompany || '-'}</td>
            <th>ชื่อลูกค้า</th><td>${task.customerName || '-'}</td>
          </tr>
          <tr>
            <th>Staff / ผู้จัดทำ</th><td colspan="3">${job.staff || '-'}</td>
          </tr>
        </table>

        <div class="section-title">รายการ Defect ที่ต้องดำเนินการ</div>
    `;

    if (task.defects && task.defects.length > 0) {
      task.defects.forEach((def) => {
        let printImg = getPrintableImgUrl(def.imgBefore);
        let imgTag = printImg ? `<img src="${printImg}" />` : `<div class="no-img">ไม่มีรูปภาพก่อนแก้ไข</div>`;
        
        html += `
        <div class="defect-card">
          <div class="defect-layout">
            <div class="img-col">
              ${imgTag}
            </div>
            <div class="info-col">
              <div class="meta"><strong>ลักษณะงานหลัก:</strong> ${def.mainCategory} &nbsp;|&nbsp; <strong>ลักษณะงานรอง:</strong> ${def.subCategory}</div>
              <div class="meta"><strong>ทีมเข้าแก้ไข:</strong> ${def.team} &nbsp;|&nbsp; <strong>Major:</strong> <span class="${def.major === 'ใช่' ? 'major' : ''}">${def.major || 'ไม่ใช่'}</span></div>
              
              <div class="desc-box">
                <strong>รายละเอียดปัญหา:</strong>
                ${def.description}
              </div>
              
              <div class="date-badge">
                📅 กำหนดวันเข้าแก้ไข: ${task.targetFixDate || 'ยังไม่ได้ระบุวันที่'}
              </div>
            </div>
          </div>
        </div>
        `;
      });
    } else {
      html += `<p style="text-align:center; color:#94a3b8; padding: 30px 0; font-style: italic;">- ไม่มีรายการ Defect ในใบงานย่อยนี้ -</p>`;
    }

    // --- ส่วน HTML ลายเซ็นต์ท้ายเอกสาร ---
    html += `
        <div class="signature-container">
            <table class="signature-table">
                <tr>
                    <td>
                        <div class="sign-line"></div>
                        <div class="sign-text sign-name">( ${job.staff || '.........................................................'} )</div>
                        <div class="sign-text">ผู้จัดทำแผน (Staff)</div>
                        <div class="sign-text" style="margin-top: 5px;">วันที่: ......../......../..............</div>
                    </td>
                    <td>
                        <div class="sign-line"></div>
                        <div class="sign-text sign-name">( ${job.owner || '.........................................................'} )</div>
                        <div class="sign-text">ผู้อนุมัติ (Owner)</div>
                        <div class="sign-text" style="margin-top: 5px;">วันที่: ......../......../..............</div>
                    </td>
                </tr>
            </table>
        </div>
    `;

    html += `</body></html>`;

    // แปลงเนื้อหาเป็น PDF
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(`RepairPlan_${task.id}.pdf`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    exportedFiles.push({ taskId: task.id, url: file.getUrl() });
  });

  return JSON.stringify(exportedFiles);
}

// --- ฟังก์ชัน Export PDF เอกสารแก้ไข Defect (NEW) ---
function exportDefectReportToPDF(taskId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const allDataStr = getAllData();
  const allJobs = JSON.parse(allDataStr);
  
  let targetTask = null;
  let targetJob = null;
  
  // ค้นหา Job และ Task ที่ตรงกับ taskId
  for (const job of allJobs) {
    const foundTask = job.tasks.find(t => t.id === taskId);
    if (foundTask) {
      targetTask = foundTask;
      targetJob = job;
      break;
    }
  }

  if (!targetTask) throw new Error("ไม่พบข้อมูลใบงานย่อย (Task)");

  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  
  // ฟังก์ชันย่อยสำหรับแปลงรูปเป็น Base64 ฝัง PDF
  const getPrintableImgUrl = (url) => {
    if (!url) return '';
    if (url.startsWith('data:image')) return url;
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      try {
        const file = DriveApp.getFileById(match[1]);
        const blob = file.getBlob();
        const base64 = Utilities.base64Encode(blob.getBytes());
        const mimeType = blob.getContentType();
        return `data:${mimeType};base64,${base64}`;
      } catch (e) {
        return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w500`;
      }
    }
    return url;
  };

  let html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Sarabun', sans-serif; color: #1e293b; line-height: 1.6; font-size: 14px; margin: 0; padding: 10px; }
          .header-title { text-align: center; color: #0f172a; margin-bottom: 25px; font-size: 24px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
          table.header-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
          table.header-table th, table.header-table td { border: 1px solid #cbd5e1; padding: 10px 12px; text-align: left; vertical-align: top; }
          table.header-table th { background-color: #f1f5f9; width: 18%; font-weight: 600; color: #334155; }
          table.header-table td { width: 32%; color: #0f172a; }
          .section-title { font-size: 18px; font-weight: 600; color: #1e40af; border-bottom: 2px solid #93c5fd; padding-bottom: 8px; margin-bottom: 15px; }
          
          .defect-card { border: 1px solid #e2e8f0; margin-bottom: 25px; padding: 15px; page-break-inside: avoid; border-radius: 8px; background-color: #ffffff; }
          .defect-info { margin-bottom: 15px; padding: 12px; background-color: #f8fafc; border-radius: 6px; border-left: 4px solid #3b82f6; }
          .defect-info strong { color: #0f172a; }
          
          .img-grid { display: table; width: 100%; table-layout: fixed; margin-top: 15px; }
          .img-cell { display: table-cell; width: 25%; padding: 0 5px; text-align: center; vertical-align: top; }
          .img-cell img { width: 100%; max-height: 180px; object-fit: contain; border: 1px solid #cbd5e1; border-radius: 6px; padding: 2px; }
          .img-label { font-size: 13px; font-weight: 700; margin-bottom: 8px; color: #1e40af; background-color: #eff6ff; padding: 4px 0; border-radius: 4px; }
          .no-img { height: 120px; background: #f8fafc; border: 2px dashed #cbd5e1; border-radius: 6px; display: flex; align-items: center; justify-content: center; color: #94a3b8; font-size: 12px; margin-top: 5px; }

          .signature-container { width: 100%; margin-top: 50px; page-break-inside: avoid; }
          .signature-table { width: 100%; border-collapse: collapse; text-align: center; }
          .signature-table td { width: 50%; padding: 10px 20px; vertical-align: bottom; }
          .sign-line { border-bottom: 1px dashed #94a3b8; width: 70%; margin: 40px auto 10px auto; }
          .sign-text { color: #334155; font-size: 14px; line-height: 1.5; }
          .sign-name { font-weight: 600; color: #0f172a; }
        </style>
      </head>
      <body>
        <div class="header-title">เอกสารแก้ไข Defect</div>
        <table class="header-table">
          <tr>
            <th>Job ID</th><td>${targetJob.id}</td>
            <th>Task ID</th><td>${targetTask.id}</td>
          </tr>
          <tr>
            <th>Site</th><td>${targetJob.site}</td>
            <th>Scope</th><td>${targetTask.scope}</td>
          </tr>
          <tr>
            <th>Owner / ผู้ดูแล</th><td>${targetJob.owner}</td>
            <th>Building / Unit</th><td>${targetTask.building} - ${targetTask.unit}</td>
          </tr>
          <tr>
            <th>Company</th><td>${targetJob.ownerCompany || '-'}</td>
            <th>ชื่อลูกค้า</th><td>${targetTask.customerName || '-'}</td>
          </tr>
          <tr>
            <th>Staff / ผู้จัดทำ</th><td colspan="3">${targetJob.staff || '-'}</td>
          </tr>
        </table>

        <div class="section-title">รายละเอียดผลการแก้ไข Defect</div>
  `;

  if (targetTask.defects && targetTask.defects.length > 0) {
    targetTask.defects.forEach((def) => {
      let imgUnit = getPrintableImgUrl(def.imgUnit);
      let imgBefore = getPrintableImgUrl(def.imgBefore);
      let imgDuring = getPrintableImgUrl(def.imgDuring);
      let imgAfter = getPrintableImgUrl(def.imgAfter);

      const renderImg = (src, label) => `
        <div class="img-cell">
          <div class="img-label">${label}</div>
          ${src ? `<img src="${src}" />` : `<div class="no-img">ไม่มีรูปภาพ</div>`}
        </div>
      `;

      html += `
      <div class="defect-card">
        <div class="defect-info">
          <div style="margin-bottom: 6px;">
            <strong>สถานะ:</strong> <span style="color: #047857; font-weight: 600;">${def.status}</span> &nbsp;|&nbsp; 
            <strong>ลักษณะงานหลัก:</strong> ${def.mainCategory} &nbsp;|&nbsp; 
            <strong>ลักษณะงานรอง:</strong> ${def.subCategory}
          </div>
          <div style="margin-bottom: 6px;">
            <strong>ทีมเข้าแก้ไข:</strong> ${def.team}
          </div>
          <div>
            <strong>รายละเอียด:</strong> ${def.description}
          </div>
        </div>
        
        <div class="img-grid">
          ${renderImg(imgUnit, '1. รูปภาพเลขยูนิต')}
          ${renderImg(imgBefore, '2. รูปภาพก่อนแก้ไข')}
          ${renderImg(imgDuring, '3. รูปภาพระหว่างแก้ไข')}
          ${renderImg(imgAfter, '4. รูปภาพหลังแก้ไข')}
        </div>
      </div>
      `;
    });
  } else {
    html += `<p style="text-align:center; color:#94a3b8; padding: 30px 0; font-style: italic;">- ไม่มีรายการ Defect ในใบงานย่อยนี้ -</p>`;
  }

  // ลายเซ็นต์ Owner (จาก Job) และ ลูกค้า (จาก Task)
  html += `
        <div class="signature-container">
            <table class="signature-table">
                <tr>
                    <td>
                        <div class="sign-line"></div>
                        <div class="sign-text sign-name">( ${targetJob.owner || '.........................................................'} )</div>
                        <div class="sign-text">ผู้อนุมัติ (Owner)</div>
                        <div class="sign-text" style="margin-top: 5px;">วันที่: ......../......../..............</div>
                    </td>
                    <td>
                        <div class="sign-line"></div>
                        <div class="sign-text sign-name">( ${targetTask.customerName || '.........................................................'} )</div>
                        <div class="sign-text">ลูกค้า (Customer)</div>
                        <div class="sign-text" style="margin-top: 5px;">วันที่: ......../......../..............</div>
                    </td>
                </tr>
            </table>
        </div>
      </body>
    </html>
  `;

  const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(`DefectReport_${targetTask.id}.pdf`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return JSON.stringify({ taskId: targetTask.id, url: file.getUrl() });
}

// --- ฟังก์ชันสำหรับระบบ Auth ---
function registerUser(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('User');
  if (!sheet) {
    initSheets();
    sheet = ss.getSheetByName('User');
  }
  
  const data = sheet.getDataRange().getValues();
  const inputUserId = String(formData.userId).trim();

  // เช็คว่า User ID ซ้ำหรือไม่
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === inputUserId) {
      throw new Error('User ID นี้มีผู้ใช้งานแล้ว กรุณาใช้ชื่ออื่น');
    }
  }

  // บันทึกข้อมูลลง Sheet (Col A = UserID, Col B = Password, Col E = Position, Col J = Email, Col K = Line, Col L = Phone, Col N = Timestamp)
  // สร้าง Array เปล่าๆ ความยาว 14 เพื่อให้ Timestamp ไปตกที่คอลัมน์ที่ 14 (Col N)
  const newRow = new Array(15).fill('');
  newRow[0] = formData.userId;
  newRow[1] = "'" + formData.password; // เติม ' นำหน้า Password บังคับให้เป็น Text
  newRow[4] = formData.position;       // Col E: Position
  newRow[9] = formData.email;          // Col J: Email
  newRow[10] = formData.line;          // Col K: Line
  newRow[11] = formData.phone;         // Col L: Phone
  newRow[14] = new Date();             // Index ที่ 13 คือ Col N

  sheet.appendRow(newRow);
  
  return 'Success';
}

function loginUser(userId, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('User');
  if (!sheet) throw new Error('ไม่พบฐานข้อมูลผู้ใช้งาน กรุณาติดต่อผู้ดูแลระบบ');

  const data = sheet.getDataRange().getValues();
  
  // แปลงค่าที่ส่งมาเป็น String และตัดช่องว่างซ้ายขวา
  const inputUserId = String(userId).trim();
  const inputPassword = String(password).trim();

  for (let i = 1; i < data.length; i++) {
    const sheetUserId = String(data[i][0]).trim();
    let sheetPassword = String(data[i][1]).trim();
    
    // ลบเครื่องหมาย ' ออก หากมีติดมาจากการบันทึกแบบบังคับเป็นข้อความ
    if (sheetPassword.startsWith("'")) {
        sheetPassword = sheetPassword.substring(1);
    }

    // เช็ค User ID และ Password
    if (sheetUserId === inputUserId && sheetPassword === inputPassword) {
      return { 
        userId: data[i][0], 
        fullName: data[i][0] // เอาชื่อ-นามสกุลออก จึงส่ง UserID ไปแสดงผลแทน
      };
    }
  }
  
  throw new Error('User ID หรือ รหัสผ่าน ไม่ถูกต้อง');
}
