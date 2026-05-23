// ============================================================
//  Project Cost Control - Google Apps Script Backend
//  วางไฟล์นี้ใน Google Apps Script แล้ว Deploy เป็น Web App
// ============================================================

// ===== CONFIG: แก้ไขค่าเหล่านี้ให้ตรงกับ Google Sheets ของคุณ =====
const SPREADSHEET_ID = '1OheGAygDZTlsLW5aCrQ79S9y8zmbqfPI0LELP04WBZE'; // ID ของ Google Sheets
const DRIVE_FOLDER_ID = '1sB8FVJaxYrLJzSBNv-Oa2Lt2xaW521_8'; // ID ของ Google Drive folder สำหรับเก็บไฟล์

// Sheet names
const SHEET_PROJECTS = 'Projects';
const SHEET_COSTS = 'ProjectCost';
const SHEET_DOCUMENTS = 'ProjectDocument';
const SHEET_DISBURSEMENTS = 'รายการเบิก';
const SHEET_USERS = 'Users';
const SHEET_PARTNERS = 'Partners';

// ===== CORS & ROUTING =====
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const params = e.parameter || {};
    const postData = e.postData ? JSON.parse(e.postData.contents) : {};
    const data = Object.assign({}, params, postData);

    const action = data.action;
    const token = data.token;

    // Public actions (no auth required)
    if (action === 'login') {
      output.setContent(JSON.stringify(handleLogin(data)));
      return addCorsHeaders(output);
    }

    // Auth check
    if (!validateToken(token)) {
      output.setContent(JSON.stringify({ success: false, error: 'Unauthorized' }));
      return addCorsHeaders(output);
    }

    // Protected actions
    let result;
    switch (action) {
      // Projects
      case 'getProjects':       result = getProjects(); break;
      case 'getProject':        result = getProject(data.id); break;
      case 'addProject':        result = addProject(data); break;
      case 'updateProject':     result = updateProject(data); break;
      case 'deleteProject':     result = deleteProject(data.id); break;

      // Project Costs
      case 'getCosts':          result = getCosts(data.projectId); break;
      case 'addCost':           result = addCost(data); break;
      case 'updateCost':        result = updateCost(data); break;
      case 'deleteCost':        result = deleteCost(data.id); break;

      // Project Documents
      case 'getDocuments':      result = getDocuments(data.projectId); break;
      case 'addDocument':       result = addDocument(data); break;
      case 'updateDocument':    result = updateDocument(data); break;
      case 'deleteDocument':    result = deleteDocument(data.id); break;

      // Disbursements
      case 'getDisbursements':  result = getDisbursements(); break;
      case 'addDisbursement':   result = addDisbursement(data); break;
      case 'updateDisbursement':result = updateDisbursement(data); break;
      case 'deleteDisbursement':result = deleteDisbursement(data.rowIndex); break;

      // Summary
      case 'getSummary':        result = getSummary(); break;

      // Partners
      case 'getPartners':       result = getPartners(); break;
      case 'addPartner':        result = addPartner(data); break;
      case 'updatePartner':     result = updatePartner(data); break;
      case 'deletePartner':     result = deletePartner(data.id); break;

      // File Upload
      case 'uploadFile':        result = uploadFile(data); break;

      // Users
      case 'getUsers':          result = getUsers(); break;
      case 'addUser':           result = addUser(data); break;
      case 'updateUser':        result = updateUser(data); break;
      case 'deleteUser':        result = deleteUser(data.id); break;

      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }

    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.toString() }));
  }

  return addCorsHeaders(output);
}

function addCorsHeaders(output) {
  return output; // Apps Script handles CORS automatically when deployed as web app
}

// ===== AUTH =====
function handleLogin(data) {
  const { username, password } = data;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_USERS);

  // สร้าง Users sheet ถ้ายังไม่มี
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.getRange(1, 1, 1, 5).setValues([['UserID', 'Username', 'Password', 'DisplayName', 'Role']]);
    // สร้าง admin user เริ่มต้น
    sheet.getRange(2, 1, 1, 5).setValues([[generateId(), 'admin', hashPassword('admin1234'), 'Administrator', 'admin']]);
  }

  const data2d = sheet.getDataRange().getValues();
  const headers = data2d[0];
  const rows = data2d.slice(1);

  const usernameIdx = headers.indexOf('Username');
  const passwordIdx = headers.indexOf('Password');
  const displayNameIdx = headers.indexOf('DisplayName');
  const roleIdx = headers.indexOf('Role');
  const userIdIdx = headers.indexOf('UserID');

  for (const row of rows) {
    if (row[usernameIdx] === username && row[passwordIdx] === hashPassword(password)) {
      const token = generateToken(row[userIdIdx]);
      return {
        success: true,
        token,
        user: {
          id: row[userIdIdx],
          username: row[usernameIdx],
          displayName: row[displayNameIdx],
          role: row[roleIdx]
        }
      };
    }
  }

  return { success: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function hashPassword(password) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');
}

function generateToken(userId) {
  const timestamp = new Date().getTime();
  const raw = userId + ':' + timestamp + ':' + Math.random();
  return Utilities.base64Encode(raw);
}

function validateToken(token) {
  if (!token) return false;
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split(':');
    if (parts.length < 2) return false;
    const timestamp = parseInt(parts[1]);
    const now = new Date().getTime();
    // Token valid for 24 hours
    return (now - timestamp) < 24 * 60 * 60 * 1000;
  } catch (e) {
    return false;
  }
}

// ===== UTILITIES =====
function generateId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((row, index) => {
    const obj = { _rowIndex: index + 2 }; // 1-based, +1 for header
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) val = val.toISOString().split('T')[0];
      obj[h] = val === '' ? null : val;
    });
    return obj;
  }).filter(obj => Object.values(obj).some(v => v !== null && v !== '' && v !== undefined));
}

// ===== PROJECTS =====
function getProjects() {
  const sheet = getSheet(SHEET_PROJECTS);
  const projects = sheetToObjects(sheet);
  return { success: true, data: projects };
}

function getProject(id) {
  const sheet = getSheet(SHEET_PROJECTS);
  const projects = sheetToObjects(sheet);
  const project = projects.find(p => p.ProjectID === id);
  if (!project) return { success: false, error: 'Project not found' };
  return { success: true, data: project };
}

function addProject(data) {
  const sheet = getSheet(SHEET_PROJECTS);
  const id = generateId();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    if (h === 'ProjectID') return id;
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  return { success: true, id };
}

function updateProject(data) {
  const sheet = getSheet(SHEET_PROJECTS);
  const projects = sheetToObjects(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const project = projects.find(p => p.ProjectID === data.ProjectID);
  if (!project) return { success: false, error: 'Project not found' };

  const row = headers.map(h => data[h] !== undefined ? data[h] : project[h] || '');
  sheet.getRange(project._rowIndex, 1, 1, headers.length).setValues([row]);
  return { success: true };
}

function deleteProject(id) {
  const sheet = getSheet(SHEET_PROJECTS);
  const projects = sheetToObjects(sheet);
  const project = projects.find(p => p.ProjectID === id);
  if (!project) return { success: false, error: 'Project not found' };
  sheet.deleteRow(project._rowIndex);
  return { success: true };
}

// ===== PROJECT COSTS =====
function getCosts(projectId) {
  const sheet = getSheet(SHEET_COSTS);
  const costs = sheetToObjects(sheet);
  const filtered = projectId ? costs.filter(c => c.ProjectID === projectId) : costs;
  return { success: true, data: filtered };
}

function addCost(data) {
  const sheet = getSheet(SHEET_COSTS);
  const id = generateId();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    if (h === 'CostID') return id;
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  return { success: true, id };
}

function updateCost(data) {
  const sheet = getSheet(SHEET_COSTS);
  const costs = sheetToObjects(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cost = costs.find(c => c.CostID === data.CostID);
  if (!cost) return { success: false, error: 'Cost not found' };
  const row = headers.map(h => data[h] !== undefined ? data[h] : cost[h] || '');
  sheet.getRange(cost._rowIndex, 1, 1, headers.length).setValues([row]);
  return { success: true };
}

function deleteCost(id) {
  const sheet = getSheet(SHEET_COSTS);
  const costs = sheetToObjects(sheet);
  const cost = costs.find(c => c.CostID === id);
  if (!cost) return { success: false, error: 'Cost not found' };
  sheet.deleteRow(cost._rowIndex);
  return { success: true };
}

// ===== PROJECT DOCUMENTS =====
function getDocuments(projectId) {
  const sheet = getSheet(SHEET_DOCUMENTS);
  const docs = sheetToObjects(sheet);
  const filtered = projectId ? docs.filter(d => d.ProjectID === projectId) : docs;
  return { success: true, data: filtered };
}

function addDocument(data) {
  const sheet = getSheet(SHEET_DOCUMENTS);
  const id = generateId();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    if (h === 'DocID') return id;
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  return { success: true, id };
}

function updateDocument(data) {
  const sheet = getSheet(SHEET_DOCUMENTS);
  const docs = sheetToObjects(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const doc = docs.find(d => d.DocID === data.DocID);
  if (!doc) return { success: false, error: 'Document not found' };
  const row = headers.map(h => data[h] !== undefined ? data[h] : doc[h] || '');
  sheet.getRange(doc._rowIndex, 1, 1, headers.length).setValues([row]);
  return { success: true };
}

function deleteDocument(id) {
  const sheet = getSheet(SHEET_DOCUMENTS);
  const docs = sheetToObjects(sheet);
  const doc = docs.find(d => d.DocID === id);
  if (!doc) return { success: false, error: 'Document not found' };
  sheet.deleteRow(doc._rowIndex);
  return { success: true };
}

// ===== DISBURSEMENTS =====
function getDisbursements() {
  const sheet = getSheet(SHEET_DISBURSEMENTS);
  const items = sheetToObjects(sheet);
  return { success: true, data: items };
}

function addDisbursement(data) {
  const sheet = getSheet(SHEET_DISBURSEMENTS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return { success: true };
}

function updateDisbursement(data) {
  const sheet = getSheet(SHEET_DISBURSEMENTS);
  const items = sheetToObjects(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const item = items.find(i => i._rowIndex === parseInt(data.rowIndex));
  if (!item) return { success: false, error: 'Disbursement not found' };
  const row = headers.map(h => data[h] !== undefined ? data[h] : item[h] || '');
  sheet.getRange(item._rowIndex, 1, 1, headers.length).setValues([row]);
  return { success: true };
}

function deleteDisbursement(rowIndex) {
  const sheet = getSheet(SHEET_DISBURSEMENTS);
  sheet.deleteRow(parseInt(rowIndex));
  return { success: true };
}

// ===== SUMMARY =====
function getSummary() {
  const projects = sheetToObjects(getSheet(SHEET_PROJECTS));
  const costs = sheetToObjects(getSheet(SHEET_COSTS));
  const disbursements = sheetToObjects(getSheet(SHEET_DISBURSEMENTS));
  const partners = getPartnersData();

  const done = projects.filter(p => p.Status === 'Done');
  const onProcess = projects.filter(p => p.Status === 'On Process');

  const totalRevenue = projects.reduce((sum, p) => sum + (parseFloat(p.NetPrice) || 0), 0);
  const totalCost = costs.reduce((sum, c) => sum + (parseFloat(c.Amount) || 0), 0);
  const totalProfit = totalRevenue - totalCost;

  const pendingPayment = projects.filter(p => p.PaymentStatus === 'Pending').length;
  const totalDisbursed = disbursements.reduce((sum, d) => sum + (parseFloat(d['จำนวนเงิน']) || 0), 0);
  const pendingDisbursements = disbursements.filter(d => d['สถานะ'] !== 'จ่ายแล้ว').length;

  // คำนวณส่วนแบ่งกำไรแบบ per-project (รองรับ ProjectOwner = 70%)
  const partnerShareMap = {};
  partners.forEach(p => { partnerShareMap[p.PartnerID] = 0; });

  projects.forEach(function(project) {
    var projectCosts = costs.filter(function(c) { return c.ProjectID === project.ProjectID; });
    var projectCostTotal = projectCosts.reduce(function(sum, c) { return sum + (parseFloat(c.Amount) || 0); }, 0);
    var projectProfit = (parseFloat(project.NetPrice) || 0) - projectCostTotal;
    var owner = (project.ProjectOwner || '').toString().trim();

    partners.forEach(function(p) {
      var share = 0;
      if (owner !== '' && owner !== 'ไม่มีเจ้าของ') {
        var isOwner = (p.Name || '').trim() === owner || (p.DisbursementName || '').trim() === owner;
        if (isOwner) {
          share = projectProfit * 0.70;
        } else {
          // 30% แบ่งระหว่างคนที่เหลือตาม % สัดส่วนปกติ
          var nonOwners = partners.filter(function(p2) {
            return (p2.Name || '').trim() !== owner && (p2.DisbursementName || '').trim() !== owner;
          });
          var totalNonOwnerPct = nonOwners.reduce(function(s, p2) { return s + (parseFloat(p2.Percentage) || 0); }, 0);
          var myPct = totalNonOwnerPct > 0 ? (parseFloat(p.Percentage) || 0) / totalNonOwnerPct : 0;
          share = projectProfit * 0.30 * myPct;
        }
      } else {
        // ไม่มีเจ้าของ → แบ่งตาม % ปกติ
        share = projectProfit * (parseFloat(p.Percentage) || 0) / 100;
      }
      partnerShareMap[p.PartnerID] += share;
    });
  });

  const profitSharing = partners.map(function(p) {
    var share = partnerShareMap[p.PartnerID] || 0;
    var disbursed = disbursements
      .filter(function(d) {
        var nameMatch = d['คนเบิก'] && d['คนเบิก'].toString().trim() === (p.DisbursementName || '').trim();
        var typeMatch = !p.DisbursementType || p.DisbursementType === '' || d['ประเภท'] === p.DisbursementType;
        return nameMatch && typeMatch;
      })
      .reduce(function(sum, d) { return sum + (parseFloat(d['จำนวนเงิน']) || 0); }, 0);
    return {
      id: p.PartnerID,
      name: p.Name,
      percentage: parseFloat(p.Percentage) || 0,
      share: share,
      disbursed: disbursed,
      remaining: share - disbursed
    };
  });

  return {
    success: true,
    data: {
      งานคงค้าง: onProcess.length,
      งานที่ทำเสร็จแล้ว: done.length,
      รายได้สุทธิ: totalRevenue,
      ต้นทุนรวม: totalCost,
      กำไร: totalProfit,
      รอพิจรณาเพื่อจ่าย: pendingPayment,
      อื่นๆที่รอจ่าย: pendingDisbursements,
      สรุปเบิกจ่ายตาม: totalDisbursed,
      profitSharing: profitSharing
    }
  };
}

// ===== PARTNERS =====
function getPartnersData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_PARTNERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PARTNERS);
    sheet.getRange(1, 1, 1, 5).setValues([['PartnerID', 'Name', 'Percentage', 'DisbursementName', 'DisbursementType']]);
  } else {
    // migrate: add DisbursementType column if missing
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('DisbursementType') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('DisbursementType');
    }
  }
  return sheetToObjects(sheet);
}

function getPartners() {
  return { success: true, data: getPartnersData() };
}

function addPartner(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_PARTNERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PARTNERS);
    sheet.getRange(1, 1, 1, 5).setValues([['PartnerID', 'Name', 'Percentage', 'DisbursementName', 'DisbursementType']]);
  }
  const id = generateId();
  sheet.appendRow([id, data.Name || '', data.Percentage || 0, data.DisbursementName || '', data.DisbursementType || '']);
  return { success: true, id };
}

function updatePartner(data) {
  const sheet = getSheet(SHEET_PARTNERS);
  if (!sheet) return { success: false, error: 'Partners sheet not found' };
  const partners = sheetToObjects(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const partner = partners.find(p => p.PartnerID === data.PartnerID);
  if (!partner) return { success: false, error: 'Partner not found' };
  const row = headers.map(function(h) {
    if (h === 'PartnerID') return partner.PartnerID;
    return data[h] !== undefined ? data[h] : (partner[h] || '');
  });
  sheet.getRange(partner._rowIndex, 1, 1, headers.length).setValues([row]);
  return { success: true };
}

function deletePartner(id) {
  const sheet = getSheet(SHEET_PARTNERS);
  if (!sheet) return { success: false, error: 'Partners sheet not found' };
  const partners = sheetToObjects(sheet);
  const partner = partners.find(p => p.PartnerID === id);
  if (!partner) return { success: false, error: 'Partner not found' };
  sheet.deleteRow(partner._rowIndex);
  return { success: true };
}

// ===== FILE UPLOAD =====
function uploadFile(data) {
  try {
    const { fileName, fileData, mimeType, subfolder } = data;
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    let targetFolder = folder;
    if (subfolder) {
      const existing = folder.getFoldersByName(subfolder);
      targetFolder = existing.hasNext() ? existing.next() : folder.createFolder(subfolder);
    }

    const decoded = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      fileId: file.getId(),
      fileUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
      fileName: fileName
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ===== USERS =====
function getUsers() {
  const sheet = getSheet(SHEET_USERS);
  if (!sheet) return { success: true, data: [] };
  const users = sheetToObjects(sheet);
  // Don't return passwords
  return { success: true, data: users.map(u => ({ ...u, Password: undefined })) };
}

function addUser(data) {
  const sheet = getSheet(SHEET_USERS);
  if (!sheet) return { success: false, error: 'Users sheet not found' };
  const id = generateId();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    if (h === 'UserID') return id;
    if (h === 'Password') return hashPassword(data.Password || 'changeme123');
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  return { success: true, id };
}

function updateUser(data) {
  const sheet = getSheet(SHEET_USERS);
  const users = sheetToObjects(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const user = users.find(u => u.UserID === data.UserID);
  if (!user) return { success: false, error: 'User not found' };
  const row = headers.map(h => {
    if (h === 'Password' && data.Password) return hashPassword(data.Password);
    if (h === 'Password') return user.Password;
    return data[h] !== undefined ? data[h] : user[h] || '';
  });
  sheet.getRange(user._rowIndex, 1, 1, headers.length).setValues([row]);
  return { success: true };
}

function deleteUser(id) {
  const sheet = getSheet(SHEET_USERS);
  const users = sheetToObjects(sheet);
  const user = users.find(u => u.UserID === id);
  if (!user) return { success: false, error: 'User not found' };
  sheet.deleteRow(user._rowIndex);
  return { success: true };
}
