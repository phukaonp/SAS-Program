const SPREADSHEET_ID = '1BkhC_02odW8OINve6c3Ec4QI4cr_DEQvFGCVWrgebfg';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAS Defect Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// รันฟังก์ชันนี้ 1 ครั้ง เพื่อสร้าง Sheet และ Header อัตโนมัติ
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsInfo = {
    'JOB': ['JobID', 'Site', 'Owner', 'OwnerCompany', 'Staff', 'ReplyDueDate', 'Remark', 'Timestamp'],
    'TASK': ['TaskID', 'JobID', 'Scope', 'Building', 'Unit', 'Status', 'CustomerName', 'TargetFixDate', 'ActualStartDate', 'ActualEndDate', 'Duration', 'Remark', 'Timestamp'],
    'DEFECT': ['DefectID', 'TaskID', 'MainCategory', 'SubCategory', 'Description', 'Team', 'TargetStartDate', 'TargetEndDate', 'VOSteps', 'ActualStartDate', 'ActualEndDate', 'Status', 'Remark', 'Timestamp']
  };

  Object.keys(sheetsInfo).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheetsInfo[name]);
      // ปรับแต่ง Header เล็กน้อย
      sheet.getRange(1, 1, 1, sheetsInfo[name].length).setFontWeight("bold").setBackground("#f3f4f6");
    }
  });
}

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

  // ประกอบ Data กลับเป็น JSON โครงสร้างแบบเดิม
  const structuredJobs = jobs.map(job => {
    const jobTasks = tasks.filter(t => t.JobID === job.JobID).map(task => {
      const taskDefects = defects.filter(d => d.TaskID === task.TaskID).map(def => ({
        id: def.DefectID,
        mainCategory: def.MainCategory,
        subCategory: def.SubCategory,
        description: def.Description,
        team: def.Team,
        targetStartDate: def.TargetStartDate,
        targetEndDate: def.TargetEndDate,
        voSteps: def.VOSteps,
        actualStartDate: def.ActualStartDate,
        actualEndDate: def.ActualEndDate,
        status: def.Status,
        remark: def.Remark
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

function addJob(jobData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const newId = 'JOB-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  
  sheet.appendRow([
    newId, jobData.site, jobData.owner, jobData.ownerCompany, jobData.staff, 
    jobData.replyDueDate, jobData.remark, new Date()
  ]);
  return newId;
}

function addTask(jobId, taskData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('TASK');
  const newId = 'TSK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  const actualStartDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  sheet.appendRow([
    newId, jobId, taskData.scope, taskData.building, taskData.unit, 'Pending', 
    taskData.customerName, taskData.targetFixDate, actualStartDate, '', 0, taskData.remark, new Date()
  ]);
  return newId;
}

function addDefect(taskId, defectData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const newId = 'DEF-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  
  sheet.appendRow([
    newId, taskId, defectData.mainCategory, defectData.subCategory, defectData.description, 
    defectData.team, defectData.targetStartDate, defectData.targetEndDate, defectData.voSteps || '', 
    '', '', 'ยังไม่แก้ไข', defectData.remark, new Date()
  ]);
  return newId;
}
