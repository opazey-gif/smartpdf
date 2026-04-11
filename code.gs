/**
 * โครงการ: Smart PDF Governance & Knowledge Repository
 * ผู้พัฒนา: OCTO_IT (Senior Digital Justice Architect)
 * เวอร์ชัน: 3.4 (Master Governance + AI OCR Pipeline)
 */

// ==========================================
// [NEW] AUDIT TRAIL ENGINE (Immutable Logs)
// ==========================================
function logAuditEvent(actionType, resourceId, details) {
  try {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DB_SHEET_ID');
    if (!dbId) return;

    const ss = SpreadsheetApp.openById(dbId);
    let auditSheet = ss.getSheetByName('Table_AuditLogs');
    
    if (!auditSheet) {
      auditSheet = ss.insertSheet('Table_AuditLogs');
      auditSheet.appendRow(["log_id", "timestamp", "user_email", "action_type", "resource_id", "details"]);
    }

    const email = Session.getActiveUser().getEmail() || "anonymous@public.com";
    const timestamp = new Date().toISOString();
    const logId = "LOG-" + Date.now() + "-" + Math.floor(Math.random() * 1000);

    auditSheet.appendRow([logId, timestamp, email, actionType, resourceId, details]);
  } catch (e) {
    console.error("Audit Log Error: " + e.message);
  }
}

function logViewActionBackend(fileId, fileName) {
  logAuditEvent("VIEW_DOWNLOAD", fileId, `Accessed file: ${fileName}`);
  return true;
}

// ==========================================
// ABAC POLICY ENGINE (Policy-as-Code)
// ==========================================
function evaluatePolicy(user, document) {
  if (user.email === document.data_owner) return true;
  if (document.security_class === 'Secret' && user.clearance_level < 4) return false;
  if (user.project_context && document.project_context && 
      user.project_context === document.project_context && 
      document.project_context !== "-") return true;
  if (user.role === 'Supervisor' && user.department === document.department) return true;
  if (document.security_class === 'Public') return true;
  if (document.security_class === 'Internal' && user.department === document.department) return true;
  return false;
}

// ==========================================
// IDENTITY PROVIDER (Server-Side Auth)
// ==========================================
function getCurrentUserContext() {
  const email = Session.getActiveUser().getEmail() || "anonymous@public.com";
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty('DB_SHEET_ID');
  if (!dbId) return { email: email, role: "Guest", department: "-", clearance_level: 0, project_context: "-" };
  
  const ss = SpreadsheetApp.openById(dbId);
  const userSheet = ss.getSheetByName('Table_Users');
  
  let context = { email: email, role: "Staff", department: "Unknown", clearance_level: 1, project_context: "-" };

  if (userSheet) {
    const data = userSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        context.role = data[i][1];
        context.department = data[i][2];
        context.clearance_level = Number(data[i][3]) || 1;
        context.project_context = data[i][4];
        break;
      }
    }
  }
  return context;
}

function getUserIdentityAPI() {
  return getCurrentUserContext();
}

// ==========================================
// 1. SETUP ENVIRONMENT
// ==========================================
function setupEnvironment() {
  const rootFolder = DriveApp.createFolder("TIJ_Smart_PDF_Repo");
  const stagingFolder = rootFolder.createFolder("Staging_Area");
  const activeFolder = rootFolder.createFolder("Active_Files");
  
  const ss = SpreadsheetApp.create("TIJ_PDF_Governance_DB");
  const sheetTx = ss.getActiveSheet();
  sheetTx.setName("Table_Transactions");
  
  // Schema 24 คอลัมน์
  sheetTx.appendRow([
    "request_id", "file_name", "subject", "doc_date", "effective_date", "expiry_date", "pdf_creation_date", 
    "department", "unit", "document_type", "security_class", "pdpa_flag", "data_owner", "data_steward",
    "project_context", "tags", "description", "file_size", "page_count", 
    "hash_sha256", "status", "file_id", "timestamp", "ai_status"
  ]);
  
  const sheetConfig = ss.insertSheet("Table_Config");
  sheetConfig.appendRow(["config_type", "parent_value", "child_value"]);
  sheetConfig.appendRow(["Department", "-", "OED"]);
  sheetConfig.appendRow(["Department", "-", "DX"]);
  sheetConfig.appendRow(["Unit", "DX", "IT"]);
  sheetConfig.appendRow(["Unit", "DX", "Data"]);
  sheetConfig.appendRow(["Classification", "-", "Public"]);
  sheetConfig.appendRow(["Classification", "-", "Internal"]);
  sheetConfig.appendRow(["Classification", "-", "Confidential"]);
  sheetConfig.appendRow(["Classification", "-", "Secret"]);
  sheetConfig.appendRow(["Document_Type", "-", "Policy / Regulation"]);
  sheetConfig.appendRow(["Document_Type", "-", "Manual / Guideline"]);
  sheetConfig.appendRow(["Document_Type", "-", "Report / Research"]);
  sheetConfig.appendRow(["Document_Type", "-", "General Form"]);
  
  const sheetUsers = ss.insertSheet("Table_Users");
  sheetUsers.appendRow(["email", "role", "department", "clearance_level", "project_context"]);
  const adminEmail = Session.getActiveUser().getEmail();
  sheetUsers.appendRow([adminEmail, "Supervisor", "DX", 5, "P-2024-01"]); 

  const sheetAudit = ss.insertSheet("Table_AuditLogs");
  sheetAudit.appendRow(["log_id", "timestamp", "user_email", "action_type", "resource_id", "details"]);

  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'ROOT_FOLDER_ID': rootFolder.getId(),
    'STAGING_FOLDER_ID': stagingFolder.getId(),
    'DB_SHEET_ID': ss.getId()
  });

  return `Setup Complete!\nSpreadsheet ID: ${ss.getId()}`;
}

// ==========================================
// 2. WEB APP ROUTING
// ==========================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TIJ PDF Governance Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getConfigData() {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty('DB_SHEET_ID');
  if(!dbId) throw new Error("System not initialized.");
  
  const sheet = SpreadsheetApp.openById(dbId).getSheetByName('Table_Config');
  const data = sheet.getDataRange().getValues();
  let config = { departments: [], units: {}, classifications: [], docTypes: [] };
  
  for (let i = 1; i < data.length; i++) {
    let [type, parent, child] = data[i];
    if (type === 'Department') config.departments.push(child);
    if (type === 'Unit') {
      if (!config.units[parent]) config.units[parent] = [];
      config.units[parent].push(child);
    }
    if (type === 'Classification') config.classifications.push(child);
    if (type === 'Document_Type') config.docTypes.push(child);
  }
  return config;
}

// ==========================================
// 4. API: ASYNC UPLOAD (CRUD - Create)
// ==========================================
function uploadFile(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); 
  
  try {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DB_SHEET_ID');
    const stagingId = props.getProperty('STAGING_FOLDER_ID');
    
    const ss = SpreadsheetApp.openById(dbId);
    const txSheet = ss.getSheetByName('Table_Transactions');
    const data = txSheet.getDataRange().getValues();
    const hashColIndex = 19; 
    for (let i = 1; i < data.length; i++) {
      if (data[i][hashColIndex] === payload.file_hash) throw new Error("Duplicate Error: ไฟล์นี้ถูกจัดเก็บในระบบแล้ว (Hash ตรงกัน)");
    }
    
    const folder = DriveApp.getFolderById(stagingId);
    const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64Data), payload.mimeType, payload.file_name);
    const file = folder.createFile(blob);
    
    const reqId = "REQ-" + new Date().getFullYear() + "-" + Math.floor(1000 + Math.random() * 9000);
    const uploaderEmail = getCurrentUserContext().email; 
    
    txSheet.appendRow([
      reqId, payload.file_name, payload.subject, payload.doc_date, 
      payload.effective_date || "-", payload.expiry_date || "-", payload.pdf_creation_date || "-", 
      payload.department, payload.unit, payload.document_type || "-", 
      payload.security_class, payload.pdpa_flag, payload.data_owner, uploaderEmail, 
      payload.project_context || "-", payload.tags || "-", payload.description || "-", 
      payload.file_size || "-", payload.page_count || "-", 
      payload.file_hash, "QUEUED", file.getId(), new Date().toISOString(), "PENDING"
    ]);

    logAuditEvent("UPLOAD", file.getId(), `Uploaded file: ${payload.file_name} [Req: ${reqId}]`);
    return { success: true, requestId: reqId, message: "อัปโหลดสำเร็จ ระบบกำลังประมวลผล" };
    
  } catch (e) {
    logAuditEvent("UPLOAD_FAILED", "-", `Error: ${e.message} (File: ${payload.file_name})`);
    return { success: false, error: e.message };
  } finally { lock.releaseLock(); }
}

// ==========================================
// 5. ASYNC BACKGROUND WORKER (File Mover)
// ==========================================
function setupAsyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processQueuedFiles') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('processQueuedFiles').timeBased().everyMinutes(1).create();
  return "✅ สร้าง Background Trigger สำเร็จ";
}

function processQueuedFiles() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; 
  try {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DB_SHEET_ID');
    const rootFolderId = props.getProperty('ROOT_FOLDER_ID');
    if (!dbId || !rootFolderId) return;

    const ss = SpreadsheetApp.openById(dbId);
    const txSheet = ss.getSheetByName('Table_Transactions');
    const data = txSheet.getDataRange().getValues(); 
    const MAX_PROCESS = 10; 
    let processedCount = 0;
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    let activeFolder = getOrCreateSubFolder(rootFolder, "Active_Files");

    for (let i = 1; i < data.length; i++) {
      if (processedCount >= MAX_PROCESS) break;
      const row = data[i];
      const statusColIndex = 20; 
      if (row[statusColIndex] === 'QUEUED') {
        const fileId = row[21];  
        const year = row[3] ? new Date(row[3]).getFullYear().toString() : new Date().getFullYear().toString();
        try {
          const file = DriveApp.getFileById(fileId);
          let deptFolder = getOrCreateSubFolder(getOrCreateSubFolder(activeFolder, year), row[7]); 
          file.moveTo(deptFolder);
          txSheet.getRange(i + 1, statusColIndex + 1).setValue('ACTIVE');
          processedCount++;
        } catch (err) {
          txSheet.getRange(i + 1, statusColIndex + 1).setValue('ERROR');
        }
      }
    }
  } finally { lock.releaseLock(); }
}

function getOrCreateSubFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

// ==========================================
// 6. API: KNOWLEDGE CATALOG & SEARCH 
// ==========================================
function searchDocuments(filters) {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty('DB_SHEET_ID');
  if (!dbId) throw new Error("ระบบยังไม่ได้ Initialized");

  const userContext = getCurrentUserContext(); 
  logAuditEvent("SEARCH", "-", `Keyword: '${filters.keyword || "-"}', Dept: '${filters.department || "-"}', Type: '${filters.document_type || "-"}'`);

  const ss = SpreadsheetApp.openById(dbId);
  const data = ss.getSheetByName('Table_Transactions').getDataRange().getValues();
  let results = [];

  for(let i = 1; i < data.length; i++) {
    let row = data[i];
    if (row[20] !== 'ACTIVE' && row[20] !== 'QUEUED') continue;

    const docMeta = { department: row[7], security_class: row[10], data_owner: row[12], project_context: row[14] };
    if (!evaluatePolicy(userContext, docMeta)) continue; 

    let match = true;
    if (filters.keyword) {
      const kw = filters.keyword.toLowerCase();
      const subject = String(row[2]).toLowerCase();
      const filename = String(row[1]).toLowerCase();
      const tags = String(row[15]).toLowerCase();
      const description = String(row[16]).toLowerCase(); 
      if (!subject.includes(kw) && !filename.includes(kw) && !tags.includes(kw) && !description.includes(kw)) match = false;
    }
    if (filters.department && row[7] !== filters.department) match = false;
    if (filters.classification && row[10] !== filters.classification) match = false;
    if (filters.document_type && row[9] !== filters.document_type) match = false;

    if (match) {
      results.push({
        date: row[3] instanceof Date ? row[3].toISOString() : String(row[3]),
        effective_date: row[4] instanceof Date ? row[4].toISOString() : String(row[4]),
        expiry_date: row[5] instanceof Date ? row[5].toISOString() : String(row[5]),
        pdf_creation_date: String(row[6]),
        timestamp: row[22] instanceof Date ? row[22].toISOString() : String(row[22]),
        subject: row[2], department: row[7], unit: row[8], document_type: row[9],
        security_class: row[10], pdpa_flag: row[11], data_owner: row[12], data_steward: row[13],
        project_context: row[14], tags: row[15], description: row[16], file_size: row[17],
        page_count: row[18], hash_sha256: row[19], status: row[20], file_id: row[21],
        file_name: row[1], ai_status: row[23]
      });
    }
  }

  results.reverse();
  const limit = 50; const page = filters.page || 1; const startIndex = (page - 1) * limit;
  return { data: results.slice(startIndex, startIndex + limit), currentPage: page, totalPages: Math.ceil(results.length / limit) || 1, totalItems: results.length };
}

// ==========================================
// 7. API: DASHBOARD STATS
// ==========================================
function getDashboardStats() {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty('DB_SHEET_ID');
  if (!dbId) throw new Error("ระบบยังไม่ได้ Initialized");

  const ss = SpreadsheetApp.openById(dbId);
  const data = ss.getSheetByName('Table_Transactions').getDataRange().getValues();
  let stats = { totalFiles: 0, totalPages: 0, totalSizeMB: 0, byDept: {}, bySecurity: {} };

  for(let i = 1; i < data.length; i++) {
    let row = data[i]; if (row[20] !== 'ACTIVE') continue; 
    stats.totalFiles++; stats.totalPages += (parseInt(row[18]) || 0); 
    stats.totalSizeMB += (parseFloat(String(row[17]).replace(' MB', '').trim()) || 0);
    let dept = row[7] || 'Unknown'; stats.byDept[dept] = (stats.byDept[dept] || 0) + 1;
    let sec = row[10] || 'Unknown'; stats.bySecurity[sec] = (stats.bySecurity[sec] || 0) + 1;
  }
  stats.totalSizeMB = parseFloat(stats.totalSizeMB).toFixed(2);
  return stats;
}

// ==========================================
// 8. API: UPDATE DOCUMENT (CRUD - Update)
// ==========================================
function updateDocument(payload) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    const props = PropertiesService.getScriptProperties(); const dbId = props.getProperty('DB_SHEET_ID');
    if (!dbId) throw new Error("ระบบยังไม่ได้ Initialized");
    const ss = SpreadsheetApp.openById(dbId); const txSheet = ss.getSheetByName('Table_Transactions');
    const data = txSheet.getDataRange().getValues();
    const userContext = getCurrentUserContext(); let rowIndex = -1; let docMeta = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][21] === payload.file_id) { 
        rowIndex = i + 1; docMeta = { department: data[i][7], security_class: data[i][10], data_owner: data[i][12], project_context: data[i][14] }; break;
      }
    }
    if (rowIndex === -1) throw new Error("ไม่พบเอกสารในระบบ");

    let canEdit = false;
    if (userContext.email === docMeta.data_owner) canEdit = true;
    else if (userContext.role === 'Supervisor' && userContext.department === docMeta.department) canEdit = true;
    if (!canEdit) throw new Error("Access Denied: คุณไม่มีสิทธิ์แก้ไขเอกสารนี้");

    txSheet.getRange(rowIndex, 3).setValue(payload.subject); txSheet.getRange(rowIndex, 4).setValue(payload.doc_date);
    txSheet.getRange(rowIndex, 5).setValue(payload.effective_date || "-"); txSheet.getRange(rowIndex, 6).setValue(payload.expiry_date || "-");
    txSheet.getRange(rowIndex, 8).setValue(payload.department); txSheet.getRange(rowIndex, 9).setValue(payload.unit || "-");
    txSheet.getRange(rowIndex, 10).setValue(payload.document_type || "-"); txSheet.getRange(rowIndex, 11).setValue(payload.security_class);
    txSheet.getRange(rowIndex, 12).setValue(payload.pdpa_flag); txSheet.getRange(rowIndex, 13).setValue(payload.data_owner); 
    txSheet.getRange(rowIndex, 15).setValue(payload.project_context || "-"); txSheet.getRange(rowIndex, 16).setValue(payload.tags || "-");
    txSheet.getRange(rowIndex, 17).setValue(payload.description || "-");

    logAuditEvent("UPDATE", payload.file_id, `Updated metadata for: ${payload.subject}`);
    return { success: true, message: "อัปเดตข้อมูลเอกสารสำเร็จ" };
  } catch (e) { return { success: false, error: e.message }; } finally { lock.releaseLock(); }
}

// ==========================================
// 9. API: ARCHIVE DOCUMENT (CRUD - Soft Delete)
// ==========================================
function archiveDocument(fileId) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    const props = PropertiesService.getScriptProperties(); const dbId = props.getProperty('DB_SHEET_ID');
    if (!dbId) throw new Error("ระบบยังไม่ได้ Initialized");
    const ss = SpreadsheetApp.openById(dbId); const txSheet = ss.getSheetByName('Table_Transactions');
    const data = txSheet.getDataRange().getValues();
    const userContext = getCurrentUserContext(); let rowIndex = -1; let docMeta = null; let fileName = "";

    for (let i = 1; i < data.length; i++) {
      if (data[i][21] === fileId) { 
        rowIndex = i + 1; fileName = data[i][1]; docMeta = { department: data[i][7], data_owner: data[i][12] }; break;
      }
    }
    if (rowIndex === -1) throw new Error("ไม่พบเอกสารในระบบ");

    let canDelete = false;
    if (userContext.email === docMeta.data_owner) canDelete = true;
    else if (userContext.role === 'Supervisor' && userContext.department === docMeta.department) canDelete = true;
    if (!canDelete) throw new Error("Access Denied: คุณไม่มีสิทธิ์จัดเก็บเอกสารนี้");

    txSheet.getRange(rowIndex, 21).setValue("ARCHIVED"); 
    logAuditEvent("ARCHIVE", fileId, `Archived file: ${fileName}`);
    return { success: true, message: "นำเอกสารเข้าสู่คลังจัดเก็บถาวร (Archive) เรียบร้อยแล้ว" };
  } catch (e) { return { success: false, error: e.message }; } finally { lock.releaseLock(); }
}

// ==========================================
// 10. AI OCR ENGINE (Text Extraction Pipeline)
// ==========================================
function setupAIEnvironment() {
  const props = PropertiesService.getScriptProperties();
  const rootFolderId = props.getProperty('ROOT_FOLDER_ID');
  if (!rootFolderId) return "❌ Error: System not initialized. Run setupEnvironment first.";
  
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  let aiFolder;
  const folders = rootFolder.getFoldersByName("AI_Extracted_Text");
  if (folders.hasNext()) { aiFolder = folders.next(); } 
  else { aiFolder = rootFolder.createFolder("AI_Extracted_Text"); }
  
  props.setProperty('EXTRACTED_FOLDER_ID', aiFolder.getId());
  
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processAIExtraction') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('processAIExtraction').timeBased().everyMinutes(5).create();
  
  return "✅ AI OCR Environment & Trigger Setup Complete!";
}

function processAIExtraction() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; 

  try {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DB_SHEET_ID');
    const extractedFolderId = props.getProperty('EXTRACTED_FOLDER_ID');
    if (!dbId || !extractedFolderId) return;

    const ss = SpreadsheetApp.openById(dbId);
    const txSheet = ss.getSheetByName('Table_Transactions');
    const data = txSheet.getDataRange().getValues();
    const extractedFolder = DriveApp.getFolderById(extractedFolderId);
    
    let processedCount = 0;
    const MAX_PROCESS_PER_RUN = 3; 

    for (let i = 1; i < data.length; i++) {
      if (processedCount >= MAX_PROCESS_PER_RUN) break;
      
      // ดึงไฟล์ที่รอทำ (PENDING) หรือเคยล้มเหลวแต่เราอยากให้ลองใหม่ (RETRY)
      if (data[i][20] === 'ACTIVE' && (data[i][23] === 'PENDING' || data[i][23] === 'RETRY')) {
        const fileId = data[i][21]; 
        const fileName = data[i][1]; 
        
        try {
          const file = DriveApp.getFileById(fileId);
          // [FIX 1] บังคับประทับตราว่านี่คือ PDF 100% ป้องกัน Google API สับสน
          const pdfBlob = file.getBlob().setContentType(MimeType.PDF); 
          let tempDoc;

          try {
            // [FIX 2] สำหรับ Drive API V3
            tempDoc = Drive.Files.create({
              name: fileName + '_OCR', 
              mimeType: MimeType.GOOGLE_DOCS
            }, pdfBlob);
          } catch (apiErr) {
            // [FIX 3] สำหรับ Drive API V2 (ลบ mimeType ออกจาก metadata เพื่อไม่ให้ชนกับ ocr: true)
            tempDoc = Drive.Files.insert({
              title: fileName + '_OCR'
            }, pdfBlob, {
              ocr: true, 
              ocrLanguage: 'th'
            });
          }

          const doc = DocumentApp.openById(tempDoc.id);
          const extractedText = doc.getBody().getText();
          const textFile = extractedFolder.createFile(fileName.replace('.pdf', '.txt'), extractedText, MimeType.PLAIN_TEXT);
          
          DriveApp.getFileById(tempDoc.id).setTrashed(true);
          
          txSheet.getRange(i + 1, 24).setValue('EXTRACTED'); 
          processedCount++;
          
          logAuditEvent("AI_OCR_SUCCESS", fileId, `Extracted text saved to: ${textFile.getId()}`);

        } catch (err) {
          txSheet.getRange(i + 1, 24).setValue('FAILED');
          logAuditEvent("AI_OCR_FAILED", fileId, `OCR Error: ${err.message}`);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}
