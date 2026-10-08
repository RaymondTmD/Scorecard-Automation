// ==========================================
// CONFIGURATION & SETUP
// ==========================================
const CONFIG = {
  sheetNames: [
    'VIDEO MODERATION', 
    'VIDEO MODERATION II',
    'COMMENT MODERATION',
    'CURATION',
    'CONTRIBUTOR',
    'SR.COMMENT MODERATOR'
  ], 
  templateDocId: '1aVwo6o6EstyOzBtEJihMwnklQSfgiQYf6iyQr7WIIvE', // <-- PASTE YOUR ACTUAL GOOGLE DOC TEMPLATE ID HERE!
  adminEmails: ['raymond.tumanday@iopex.com', 'another.leader@iopex.com'] // <-- ADD AUTHORIZED ADMIN EMAILS HERE!
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 Scorecard App')
    .addItem('Open Web Dashboard', 'openWebDashboard')
    .addToUi();
}

function openWebDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1450)
    .setHeight(850)
    .setTitle('iOPEX Scorecard Automator');
  SpreadsheetApp.getUi().showModalDialog(html, '🏆 Scorecard Automation Dashboard');
}

// ==========================================
// AUDIT TRAIL LOGGING ENGINE
// ==========================================
function logAudit(action, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('AUDIT_LOG');
    if (!sheet) {
      sheet = ss.insertSheet('AUDIT_LOG');
      sheet.appendRow(['Timestamp', 'Admin Email', 'Action Performed', 'Details']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2C3E50').setFontColor('#FFFFFF');
    }
    const email = Session.getActiveUser().getEmail() || 'System/User';
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([time, email, action, details]);
  } catch(e) {
    console.error("Audit Log Error: " + e.message);
  }
}

// ==========================================
// E-SIGNATURE UI & WEB APP ROUTER (UPDATED WITH PERIOD & ROW INDEX)
// ==========================================
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'ack') {
    const empId = e.parameter.empId || '';
    const team = e.parameter.team || '';
    const period = e.parameter.period || '';
    const rowIndex = e.parameter.rowIndex || '';
    
    // Escape quotes and backslashes to ensure safe injection into the JS template
    const safeEmpId = String(empId).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const safeTeam = String(team).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const safePeriod = String(period).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const safeRowIndex = String(rowIndex).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0">
        <title>Acknowledge & Sign Scorecard</title>
        <style>
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #F4F7F9; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; }
          .card { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); text-align: center; border-top: 6px solid #E65100; max-width: 450px; width: 100%; }
          h2 { color: #2C3E50; margin-top: 0; margin-bottom: 5px; }
          p { color: #7F8C8D; font-size: 14px; margin-bottom: 15px; }
          .notice-box { background-color: #FFF8E1; border-left: 4px solid #FF8F00; color: #5D4037; padding: 12px; font-size: 13px; text-align: left; margin-bottom: 18px; border-radius: 6px; line-height: 1.4; }
          canvas { border: 2px dashed #CBD5E1; border-radius: 8px; background: #FAFAFA; cursor: crosshair; touch-action: none; width: 100%; height: 160px; }
          .btn-group { display: flex; gap: 10px; justify-content: center; margin-top: 20px; }
          button { border: none; padding: 12px 20px; border-radius: 6px; font-weight: bold; font-size: 14px; cursor: pointer; transition: 0.2s; flex: 1; }
          .btn-clear { background: #E2E8F0; color: #475569; }
          .btn-submit { background: #E65100; color: white; }
          .btn-submit:disabled { background: #ccc; cursor: not-allowed; }
          #success-screen { display: none; }
          .loader { display: none; margin-top: 15px; font-size: 14px; color: #E65100; font-weight: bold; }
        </style>
      </head>
      <body>
        <div class="card" id="sign-screen">
          <h2>Scorecard Acknowledgment</h2>
          <p>Please draw your signature below to acknowledge your performance scorecard for this period.</p>
          
          <div class="notice-box">
            <strong>⚠️ Formal Signature Required:</strong> Please draw your official, legal signature. This signature and timestamp will be permanently embedded into your performance scorecard.
          </div>

          <canvas id="sigPad"></canvas>
          <div class="loader" id="loader">Submitting your signature... ⏳</div>
          <div class="btn-group">
            <button class="btn-clear" onclick="clearPad()">Clear</button>
            <button class="btn-submit" id="submitBtn" onclick="submitPad()">Acknowledge & Sign</button>
          </div>
        </div>
        
        <div class="card" id="success-screen">
          <div style="font-size: 60px; margin-bottom: 10px;">🎉</div>
          <h1 style="color: #E65100; font-size: 26px; margin-bottom: 15px; line-height: 1.3;">Thanks for your Effort and Commitment! :)</h1>
          <p style="color: #475569; font-size: 16px;">Your scorecard has been successfully acknowledged and digitally signed.</p>
          <p style="color: #7F8C8D; font-size: 13px; margin-top: 30px; border-top: 1px solid #E0E0E0; padding-top: 15px;">You may now safely close this window.</p>
        </div>

        <script>
          const canvas = document.getElementById('sigPad');
          const ctx = canvas.getContext('2d');
          let drawing = false;

          function resizeCanvas() {
            const rect = canvas.getBoundingClientRect();
            canvas.width = rect.width;
            canvas.height = rect.height;
          }
          window.addEventListener('resize', resizeCanvas);
          resizeCanvas();

          function getPos(e) {
            const rect = canvas.getBoundingClientRect();
            const clientX = e.touches ? e.touches[0].clientX : e.clientX;
            const clientY = e.touches ? e.touches[0].clientY : e.clientY;
            return { x: clientX - rect.left, y: clientY - rect.top };
          }

          function startPos(e) { drawing = true; ctx.beginPath(); const pos = getPos(e); ctx.moveTo(pos.x, pos.y); e.preventDefault(); }
          function endPos() { drawing = false; }
          function draw(e) {
            if (!drawing) return;
            const pos = getPos(e);
            ctx.lineWidth = 3; ctx.lineCap = 'round'; ctx.strokeStyle = '#0F172A';
            ctx.lineTo(pos.x, pos.y); ctx.stroke();
            e.preventDefault();
          }

          canvas.addEventListener('mousedown', startPos); canvas.addEventListener('mouseup', endPos); canvas.addEventListener('mousemove', draw);
          canvas.addEventListener('touchstart', startPos, {passive: false}); canvas.addEventListener('touchend', endPos); canvas.addEventListener('touchmove', draw, {passive: false});

          function clearPad() { ctx.clearRect(0, 0, canvas.width, canvas.height); }

          function submitPad() {
            const blank = document.createElement('canvas');
            blank.width = canvas.width; blank.height = canvas.height;
            if (canvas.toDataURL() === blank.toDataURL()) { alert('Please draw your signature first!'); return; }

            document.getElementById('submitBtn').disabled = true;
            document.getElementById('loader').style.display = 'block';
            
            // Stamp the date at the bottom right of the canvas before exporting
            ctx.font = "bold 11px 'Segoe UI', sans-serif";
            ctx.fillStyle = "#475569";
            ctx.textAlign = "right";
            const now = new Date();
            const dateStamp = "Signed: " + now.toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' });
            ctx.fillText(dateStamp, canvas.width - 10, canvas.height - 8);

            const base64Data = canvas.toDataURL('image/png').split(',')[1];
            
            google.script.run
              .withSuccessHandler(() => {
                document.getElementById('sign-screen').style.display = 'none';
                document.getElementById('success-screen').style.display = 'block';
              })
              .withFailureHandler(err => {
                alert('Error submitting signature: ' + err.message);
                document.getElementById('submitBtn').disabled = false;
                document.getElementById('loader').style.display = 'none';
              })
              .saveSignatureToServer('${safeEmpId}', '${safeTeam}', '${safePeriod}', '${safeRowIndex}', base64Data);
          }
        </script>
      </body>
      </html>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('iOPEX Scorecard Automator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 3-TIER SMART SIGNATURE SAVER (FIXED BUG)
// ==========================================
function saveSignatureToServer(empId, sheetName, period, rowIndex, base64Data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  
  let empIdCol = headers.findIndex(h => h.includes('emp id') || h.includes('employee id'));
  let ackCol = headers.findIndex(h => h.includes('acknowledgment status'));
  let sigCol = headers.findIndex(h => h.includes('e-signature'));
  
  if (empIdCol === -1) throw new Error("Could not find 'Emp ID' column in sheet.");
  if (ackCol === -1) { ackCol = sheet.getLastColumn(); sheet.getRange(1, ackCol + 1).setValue("Acknowledgment Status"); }
  if (sigCol === -1) { sigCol = sheet.getLastColumn(); sheet.getRange(1, sigCol + 1).setValue("E-Signature"); }
  
  // Find the Period/Month/Year column dynamically just like getAgentData does
  let monthCol = -1;
  let count = 0;
  const keywords = ['month', 'year', 'period', 'date', 'review'];
  const excludeWords = ['email', 'days', 'count', 'live'];
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (keywords.some(k => header.includes(k.toLowerCase())) && !excludeWords.some(e => header.includes(e.toLowerCase()))) {
      count++;
      if (count === 1) { monthCol = i; break; }
    }
  }
  if (monthCol === -1) monthCol = 0; // fallback to Column A
  
  let targetRowIndex = -1;
  const targetEmpId = String(empId).trim().toLowerCase();
  const targetPeriod = String(period || '').trim().toLowerCase();
  const targetRowNum = parseInt(rowIndex, 10);
  
  // TIER 1: Check if exact rowIndex from email points directly to the right Emp ID and Period
  if (!isNaN(targetRowNum) && targetRowNum > 1 && targetRowNum <= data.length) {
    const rowEmpId = String(data[targetRowNum - 1][empIdCol]).trim().toLowerCase();
    const rowPeriod = String(data[targetRowNum - 1][monthCol] || data[targetRowNum - 1][0] || '').trim().toLowerCase();
    
    if (rowEmpId === targetEmpId && (!targetPeriod || rowPeriod === targetPeriod)) {
      targetRowIndex = targetRowNum;
    }
  }
  
  // TIER 2: If exact row match failed (e.g., sheet sorted/filtered), scan for exact match of Emp ID + Period
  if (targetRowIndex === -1 && targetPeriod !== '') {
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][empIdCol]).trim().toLowerCase();
      const rowPeriod = String(data[i][monthCol] || data[i][0] || '').trim().toLowerCase();
      
      if (rowEmpId === targetEmpId && rowPeriod === targetPeriod) {
        targetRowIndex = i + 1;
        break;
      }
    }
  }
  
  // TIER 3: Fallback for old email links without period. Loop from BOTTOM UP to update their newest/latest row!
  if (targetRowIndex === -1) {
    for (let i = data.length - 1; i >= 1; i--) {
      const rowEmpId = String(data[i][empIdCol]).trim().toLowerCase();
      if (rowEmpId === targetEmpId) {
        targetRowIndex = i + 1;
        break;
      }
    }
  }
  
  if (targetRowIndex !== -1) {
     sheet.getRange(targetRowIndex, ackCol + 1).setValue("Signed on " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy"));
     sheet.getRange(targetRowIndex, sigCol + 1).setValue(base64Data);
     const matchedPeriod = String(data[targetRowIndex - 1][monthCol] || data[targetRowIndex - 1][0] || 'Current Period');
     logAudit("E-Signature Completed", `Emp ID: ${empId} signed their scorecard for period '${matchedPeriod}' (Row ${targetRowIndex}) on team: ${sheetName}`);
  } else {
     throw new Error(`Could not find a matching record for Emp ID ${empId} (${period}) on tab ${sheetName}.`);
  }
}

// ==========================================
// TEMPLATE POPULATOR (SUPPORTS UNSIGNED OPTION)
// ==========================================
function populateTemplate(doc, agent, ignoreSignature = false) {
  const body = doc.getBody();
  const header = doc.getHeader();
  const footer = doc.getFooter();
  
  Object.keys(agent).forEach(key => {
    if (key !== '_rowIndex' && key !== '_sheetName' && key !== 'E-Signature' && agent[key] !== undefined && agent[key] !== null) {
      const escapedKey = key.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const searchPattern = `\\{\\{${escapedKey}\\}\\}`;
      const val = String(agent[key]);
      if (body) body.replaceText(searchPattern, val);
      if (header) header.replaceText(searchPattern, val);
      if (footer) footer.replaceText(searchPattern, val);
    }
  });

  if (body) {
    let sigRange = body.findText("\\{\\{Agent Signature\\}\\}");
    if (sigRange) {
      let textElement = sigRange.getElement();
      let parent = textElement.getParent();
      
      // If ignoreSignature is true, it skips placing the image entirely.
      if (!ignoreSignature && agent['E-Signature'] && agent['E-Signature'] !== '-') {
        try {
          let imageBlob = Utilities.newBlob(Utilities.base64Decode(agent['E-Signature']), 'image/png', 'signature.png');
          parent.asParagraph().insertInlineImage(parent.getChildIndex(textElement), imageBlob).setWidth(150).setHeight(50);
          textElement.setText(textElement.getText().replace("{{Agent Signature}}", "")); 
        } catch(e) {
          textElement.setText(textElement.getText().replace("{{Agent Signature}}", "[Signature Error]"));
        }
      } else {
        textElement.setText(textElement.getText().replace("{{Agent Signature}}", "[Pending E-Signature]"));
      }
    }
  }
}

function getValidatedTemplateFile() {
  if (!CONFIG.templateDocId || CONFIG.templateDocId === 'YOUR_TEMPLATE_DOC_ID_HERE') throw new Error("❌ You forgot to paste your Google Doc Template ID into CONFIG.templateDocId!");
  return DriveApp.getFileById(CONFIG.templateDocId.trim());
}

// ==========================================
// LOGIN & DASHBOARD PAYLOAD ROUTER
// ==========================================
function getInitialSetup() {
  const email = Session.getActiveUser().getEmail() || '';
  const globalTheme = PropertiesService.getScriptProperties().getProperty('GLOBAL_THEME') || 'default';
  return { email: email, theme: globalTheme };
}

function getDashboardPayload(userEmail) {
  const currentUser = (userEmail || Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  const isAdmin = CONFIG.adminEmails.map(e => e.toLowerCase()).includes(currentUser);
  const globalTheme = PropertiesService.getScriptProperties().getProperty('GLOBAL_THEME') || 'default';
  
  let agents = getAgentData(['ALL'], ['ALL']);
  
  // If non-admin, filter dataset strictly to only their own email!
  if (!isAdmin && currentUser !== '') {
    agents = agents.filter(a => String(a['Moderator Email']).toLowerCase().trim() === currentUser);
    if (agents.length === 0) {
      throw new Error("⛔ Access Denied: No scorecard records found for email address: " + currentUser);
    }
  }
  
  return { isAdmin: isAdmin, theme: globalTheme, agents: agents, userEmail: currentUser };
}

function verifyAdmin() {
  const currentUser = Session.getActiveUser().getEmail().toLowerCase();
  if (!CONFIG.adminEmails.map(e => e.toLowerCase()).includes(currentUser)) throw new Error("⛔ Unauthorized: You do not have Admin permission.");
}

function saveGlobalTheme(themeName) {
  verifyAdmin(); 
  PropertiesService.getScriptProperties().setProperty('GLOBAL_THEME', themeName);
  logAudit("Theme Change", `Global dashboard theme updated to: ${themeName}`);
  return themeName;
}

// ==========================================
// SMART HEADER SCANNER
// ==========================================
function getAgentData(teamFilters = ['ALL'], monthFilters = ['ALL']) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allAgents = [];
  
  if (!Array.isArray(teamFilters)) teamFilters = [teamFilters];
  if (!Array.isArray(monthFilters)) monthFilters = [monthFilters];
  
  CONFIG.sheetNames.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return; 
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;
    
    const rawHeaders = data.shift();
    const headers = rawHeaders.map(h => String(h).trim().toLowerCase());
    
    const findCol = (row, keywords, excludeWords = [], occurrence = 1) => {
      let count = 0;
      for (let i = 0; i < headers.length; i++) {
        const header = headers[i];
        if (keywords.some(k => header.includes(k.toLowerCase())) && !excludeWords.some(e => header.includes(e.toLowerCase()))) {
          count++;
          if (count === occurrence) return row[i] !== undefined && row[i] !== null && row[i] !== '' ? row[i] : '-';
        }
      }
      return null;
    };
    
    const tabAgents = data.map((row, index) => {
      let monthVal = findCol(row, ['month', 'year', 'period', 'date', 'review'], ['email', 'days', 'count', 'live']);
      if (!monthVal || monthVal === '-') monthVal = row[0] || 'Current Period'; 
      
      const finalScoringVal = findCol(row, ['final scoring', 'final score']) || '0';
      const scoreVal = parseFloat(String(finalScoringVal).replace(/[^0-9.-]/g, ''));
      const computedTopStatus = scoreVal >= 80 ? 'Pass' : 'Fail';
      const sheetGateStatus = findCol(row, ['gate (fail/pass)', 'gate status', 'gate'], ['before', 'keeper', 'incentive']) || '-';

      const ackStatus = findCol(row, ['acknowledgment status', 'ack']) || '-';
      const dateSignedVal = ackStatus.includes('Signed on ') ? ackStatus.replace('Signed on ', '') : '-';

      return {
        '_rowIndex': index + 2,
        '_sheetName': tabName, 
        'Team': tabName,       
        'Generated Date': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy"),
        'Top Status': computedTopStatus,         
        'Gate (Fail/Pass)': sheetGateStatus,     
        'Pass Box': computedTopStatus === 'Pass' ? '[ ✓ ] PASS' : '[   ] PASS',
        'Fail Box': computedTopStatus === 'Pass' ? '[   ] FAIL' : '[ ✓ ] FAIL',
        'Month/Year': monthVal,
        'Moderator Email': findCol(row, ['email'], ['status']) || '-',
        'Emp ID': findCol(row, ['emp id', 'employee id']) || '-',
        'Moderator': findCol(row, ['moderator', 'name'], ['email']) || '-', 
        'Go live': findCol(row, ['go live', 'go live date']) || '-',
        'Tenurity Days': findCol(row, ['tenurity days'], [], 1) || '-',
        'Login Count': findCol(row, ['login count', 'login']) || '-',
        'Productivity': findCol(row, ['productivity'], ['rating', 'weight'], 1) || '-',
        'Client Escalation': findCol(row, ['client escalation'], ['rating', 'weight'], 1) || '-',
        'Accuracy': findCol(row, ['accuracy'], ['rating', 'weight'], 1) || '-',
        'PKT': findCol(row, ['pkt'], ['rating', 'weight'], 1) || '-',
        'Personal Attendance': findCol(row, ['personal attendance', 'personal attendan'], ['rating', 'weight', '2'], 1) || '-',
        'Productivity Rating': findCol(row, ['productivity'], [], 2) || '-',
        'Client Escalation Rating': findCol(row, ['client escalation'], [], 2) || '-',
        'Accuracy Rating': findCol(row, ['accuracy'], [], 2) || '-',
        'PKT Rating': findCol(row, ['pkt'], [], 2) || '-',
        'Personal Attendance Rating': findCol(row, ['personal attendance', 'personal attendan'], ['2'], 2) || '-',
        'OverBreak': findCol(row, ['overbreak'], ['rating'], 1) || '-',
        'Scoreable': findCol(row, ['scoreable']) || '-',
        'Scorecard %': finalScoringVal || '-',
        'Final Scoring': finalScoringVal || '-',
        'Final Rating': findCol(row, ['final rating', 'final overall rating']) || '-',
        'Incentive Before Gatekeepers': findCol(row, ['incentive before g', 'incentive before gatekeepers']) || '-',
        'Tenurity': findCol(row, ['tenurity'], ['days'], 1) || '-', 
        'Personal attendance 2': findCol(row, ['personal attendance', 'personal attendan'], [], 3) || '-', 
        'Overbreak Rating': findCol(row, ['overbreak rating', 'overbreak'], [], 2) || '-', 
        'Final Bonus (Amount)': findCol(row, ['final bonus', 'bonus'], ['email', 'status', 'gate']) || '-',
        'Email Status': findCol(row, ['email status', 'status'], ['gate', 'bonus', 'live']) || '-',
        'Acknowledgment Status': ackStatus,
        'Date Signed': dateSignedVal,
        'Signed Date': dateSignedVal,
        'E-Signature': findCol(row, ['e-signature', 'signature']) || '-' 
      };
    }).filter(agent => agent['Moderator'] && agent['Moderator'] !== '-');
    allAgents = allAgents.concat(tabAgents);
  });
  
  if (allAgents.length === 0) throw new Error("❌ No valid agent data found across tabs.");

  if (!teamFilters.includes('ALL') && teamFilters.length > 0) {
    allAgents = allAgents.filter(a => teamFilters.includes(a._sheetName));
  }
  if (!monthFilters.includes('ALL') && monthFilters.length > 0) {
    allAgents = allAgents.filter(a => monthFilters.includes(String(a['Month/Year']).trim()));
  }
  
  allAgents.sort((a, b) => {
    const getNum = (val) => parseFloat(String(val).replace(/[^0-9.-]/g, '')) || 0;
    let ratingA = getNum(a['Final Rating']), ratingB = getNum(b['Final Rating']);
    if (ratingB !== ratingA) return ratingB - ratingA; 
    return getNum(b['Personal Attendance']) - getNum(a['Personal Attendance']);      
  });
  
  return allAgents;
}

// ==========================================
// FORMAL EMAIL GENERATOR
// ==========================================
function getFormalEmailHtml(recipientName, title, subtitle, summaryRowsHtml, buttonText, buttonUrl) {
  return `
    <div style="font-family: 'Segoe UI', sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #E0E0E0; border-radius: 8px; overflow: hidden; color: #333333;">
      <div style="background-color: #FFFFFF; padding: 20px 25px; border-bottom: 3px solid #E65100; display: flex; align-items: center; justify-content: space-between;">
        <h2 style="color: #E65100; font-size: 22px; font-weight: bold; margin: 0;">&Phi; iOPEX <span style="color: #7F8C8D; font-size: 16px;">TECHNOLOGIES</span></h2>
        <span style="font-size: 11px; color: #7F8C8D; font-weight: bold; text-transform: uppercase;">Private &amp; Confidential</span>
      </div>
      <div style="padding: 25px; background-color: #FAFAFA;">
        <p style="font-size: 16px; margin-top: 0; color: #2C3E50;">Hello <strong>${recipientName}</strong>,</p>
        <p style="font-size: 14px; line-height: 1.6; color: #475569;">${subtitle}</p>
        <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 5px solid #E65100; border-radius: 6px; padding: 18px 20px; margin: 25px 0;">
          <h3 style="margin: 0 0 12px 0; font-size: 13px; color: #7F8C8D; text-transform: uppercase; border-bottom: 1px solid #EDF2F7; padding-bottom: 8px;">${title}</h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">${summaryRowsHtml}</table>
        </div>
        ${buttonUrl ? `<div style="text-align: center; margin: 30px 0;"><a href="${buttonUrl}" target="_blank" style="background-color: #E65100; color: white; padding: 12px 28px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px;">${buttonText}</a></div>` : ''}
        <p style="font-size: 13px; color: #7F8C8D; margin-top: 30px; border-top: 1px solid #E2E8F0; padding-top: 15px;">Best regards,<br><strong style="color: #2C3E50;">Moderation Team Leadership</strong></p>
      </div>
      <div style="background-color: #F1F5F9; padding: 15px; text-align: center; font-size: 11px; color: #64748B;">
        E-SQUARE I.T PARK, TAGUIG CITY-1634<br><a href="https://www.iopex.com" style="color: #E65100; text-decoration: none; font-weight: bold;">www.iopex.com</a>
      </div>
    </div>
  `;
}

// ==========================================
// GENERATORS (ADMIN SECURED)
// ==========================================
function generateSignedIndividualDocs(teamFilters = ['ALL'], monthFilters = ['ALL']) {
  return processBatchGeneration(teamFilters, monthFilters, false, false, false);
}
function generateSignedCombinedDoc(teamFilters = ['ALL'], monthFilters = ['ALL']) {
  return processBatchGeneration(teamFilters, monthFilters, false, true, false);
}
function generateUnsignedIndividualDocs(teamFilters = ['ALL'], monthFilters = ['ALL']) {
  return processBatchGeneration(teamFilters, monthFilters, true, false, false);
}
function generateUnsignedCombinedDoc(teamFilters = ['ALL'], monthFilters = ['ALL']) {
  return processBatchGeneration(teamFilters, monthFilters, true, true, false);
}
function generateSignedOnlyCombinedDoc(teamFilters = ['ALL'], monthFilters = ['ALL']) {
  return processBatchGeneration(teamFilters, monthFilters, false, true, true);
}

function processBatchGeneration(teamFilters, monthFilters, ignoreSignature, isCombined, onlySigned = false) {
  verifyAdmin(); 
  let data = getAgentData(teamFilters, monthFilters);
  
  if (onlySigned) {
    data = data.filter(a => a['Acknowledgment Status'] && String(a['Acknowledgment Status']).includes('Signed'));
    if (data.length === 0) throw new Error("❌ No digitally signed scorecards found matching your selected filters!");
  }

  const templateFile = getValidatedTemplateFile();
  const myEmail = Session.getActiveUser().getEmail();
  const tempFilesToTrash = [];
  const attachments = [];
  
  const signType = onlySigned ? "Signed (Only)" : (ignoreSignature ? "Unsigned" : "Signed");

  if (isCombined) {
    const templateDoc = DocumentApp.openById(templateFile.getId());
    const templateElements = templateDoc.getBody().getNumChildren();
    const fileName = `Combined ${signType} Scorecards - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy")}`;
    const combinedDoc = DocumentApp.create(fileName);
    const combinedBody = combinedDoc.getBody();
    combinedBody.clear();
    
    data.forEach((agent, index) => {
      const copy = templateFile.makeCopy(`Temp_${index}`);
      const tempDoc = DocumentApp.openById(copy.getId());
      populateTemplate(tempDoc, agent, ignoreSignature);
      tempDoc.saveAndClose();
      
      const filledDoc = DocumentApp.openById(copy.getId());
      const elements = filledDoc.getBody().getNumChildren();
      for (let i = 0; i < elements; i++) {
        const element = filledDoc.getBody().getChild(i).copy();
        const type = element.getType();
        if (type == DocumentApp.ElementType.PARAGRAPH) combinedBody.appendParagraph(element);
        else if (type == DocumentApp.ElementType.TABLE) combinedBody.appendTable(element);
        else if (type == DocumentApp.ElementType.LIST_ITEM) combinedBody.appendListItem(element);
      }
      if (index < data.length - 1) combinedBody.appendPageBreak();
      copy.setTrashed(true);
    });
    
    combinedDoc.saveAndClose();
    const fileId = combinedDoc.getId();
    const file = DriveApp.getFileById(fileId);
    SpreadsheetApp.flush(); Utilities.sleep(4000);
    
    attachments.push(file.getAs('application/pdf').setName(`${fileName}.pdf`));
    attachments.push(UrlFetchApp.fetch(`https://docs.google.com/document/d/${fileId}/export?format=docx`, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }).getBlob().setName(`${fileName}.docx`));
    
    const summaryRows = `<tr><td style="padding: 6px 0; color: #475569; width: 55%;"><strong>Filtered Teams:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #2C3E50;">${teamFilters.join(', ')}</td></tr><tr><td style="padding: 6px 0; color: #475569;"><strong>Total Agents:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #2C3E50;">${data.length} Agents</td></tr>`;
    MailApp.sendEmail({ to: myEmail, subject: `🖨️ Combined ${signType} Scorecard (${data.length} Agents)`, htmlBody: getFormalEmailHtml("Leader", "Combined Document Summary", `Attached is your master document containing all ${data.length} agents (${signType}).`, summaryRows, null, null), attachments: attachments });
    file.setTrashed(true);
    
    logAudit("Batch Combined Generation", `Generated ${signType} master document for ${data.length} agents.`);
    return `✅ Combined ${signType} document sent to ${myEmail}!`;
    
  } else {
    data.forEach(agent => {
      const copy = templateFile.makeCopy(`Scorecard - ${agent['Team']} - ${agent['Moderator']}`);
      const doc = DocumentApp.openById(copy.getId());
      populateTemplate(doc, agent, ignoreSignature);
      doc.saveAndClose();
      tempFilesToTrash.push(copy);
    });
    
    SpreadsheetApp.flush(); Utilities.sleep(4000);
    tempFilesToTrash.forEach(copy => attachments.push(copy.getAs('application/pdf').setName(`${copy.getName()}.pdf`)));
    
    const summaryRows = `<tr><td style="padding: 6px 0; color: #475569; width: 55%;"><strong>Filtered Teams:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #2C3E50;">${teamFilters.join(', ')}</td></tr><tr><td style="padding: 6px 0; color: #475569;"><strong>Total Scorecards:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #2C3E50;">${data.length} Agents</td></tr>`;
    MailApp.sendEmail({ to: myEmail, subject: `🖨️ ${signType} Individual Scorecards (${data.length} Agents)`, htmlBody: getFormalEmailHtml("Leader", "Batch Summary", `Attached are ${data.length} ready-to-print PDFs (${signType}).`, summaryRows, null, null), attachments: attachments });
    tempFilesToTrash.forEach(file => file.setTrashed(true));
    
    logAudit("Batch Individual Generation", `Generated ${signType} individual PDFs for ${data.length} agents.`);
    return `✅ Generated ${data.length} ${signType} PDFs and sent them to ${myEmail}!`;
  }
}

// ==========================================
// INDIVIDUAL EMAIL GENERATOR (UPDATED WITH &period= & &rowIndex=)
// ==========================================
function sendSingleAgentEmail(rowIndex, sheetName) {
  verifyAdmin(); 
  const data = getAgentData(['ALL'], ['ALL']);
  const agent = data.find(a => a._rowIndex == rowIndex && a._sheetName === sheetName);
  
  if (!agent || !agent['Moderator Email'].includes('@')) throw new Error(`❌ Valid email missing for ${agent ? agent['Moderator'] : 'agent'}!`);
  
  const templateFile = getValidatedTemplateFile();
  const fileName = `Scorecard - ${agent['Team']} - ${agent['Moderator']}`;
  const copy = templateFile.makeCopy(fileName);
  const doc = DocumentApp.openById(copy.getId());
  
  // Always ignore signature when first emailing to agent, as they haven't signed it yet!
  populateTemplate(doc, agent, true); 
  doc.saveAndClose();
  
  SpreadsheetApp.flush(); Utilities.sleep(4000);
  const docxAttachment = UrlFetchApp.fetch(`https://docs.google.com/document/d/${copy.getId()}/export?format=docx`, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }).getBlob().setName(`${fileName}.docx`);
  
  const badgeColor = agent['Top Status'] === 'Pass' ? '#2E7D32' : '#C62828';
  const badgeBg = agent['Top Status'] === 'Pass' ? '#E8F5E9' : '#FFEBEE';
  const summaryRows = `
    <tr><td style="padding: 6px 0; color: #475569; width: 55%;"><strong>Period:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #2C3E50;">${agent['Month/Year']}</td></tr>
    <tr><td style="padding: 6px 0; color: #475569;"><strong>Status:</strong></td><td style="padding: 6px 0;"><span style="background-color: ${badgeBg}; color: ${badgeColor}; font-weight: bold; padding: 3px 10px; border-radius: 12px; font-size: 12px;">${agent['Top Status']}</span></td></tr>
    <tr><td style="padding: 6px 0; color: #475569;"><strong>Scorecard %:</strong></td><td style="padding: 6px 0; font-weight: bold; font-size: 15px; color: #2C3E50;">${agent['Scorecard %']}</td></tr>
    <tr><td style="padding: 6px 0; color: #475569;"><strong>Final Rating:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #2C3E50;">${agent['Final Rating']}</td></tr>
    <tr><td style="padding: 6px 0; color: #475569;"><strong>Incentive Bonus:</strong></td><td style="padding: 6px 0; font-weight: bold; font-size: 15px; color: #2E7D32;">${agent['Final Bonus (Amount)']}</td></tr>
  `;

  const webAppUrl = ScriptApp.getService().getUrl();
  let ackUrl = '';
  if (webAppUrl && webAppUrl.indexOf('exec') > -1) {
    // INJECTED PERIOD AND ROW INDEX TO PREVENT ROW OVERWRITING BUG
    ackUrl = `${webAppUrl}?action=ack&empId=${encodeURIComponent(agent['Emp ID'])}&team=${encodeURIComponent(sheetName)}&period=${encodeURIComponent(agent['Month/Year'])}&rowIndex=${agent._rowIndex}`;
  } else {
    throw new Error("❌ Web App Not Deployed! Please Deploy the script as a Web App to enable E-Signatures.");
  }

  MailApp.sendEmail({ to: agent['Moderator Email'].trim(), subject: `Performance Scorecard - ${agent['Moderator']} (${agent['Month/Year']})`, htmlBody: getFormalEmailHtml(agent['Moderator'], "Performance Summary", `Attached is your official Scorecard for the review period. Click below to review and digitally sign your scorecard.`, summaryRows, "Acknowledge & Sign Scorecard", ackUrl), attachments: [docxAttachment] });
  copy.setTrashed(true);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let statusColIndex = headers.findIndex(h => String(h).trim().toLowerCase().includes('email status')) + 1;
  if (statusColIndex === 0) { statusColIndex = sheet.getLastColumn() + 1; sheet.getRange(1, statusColIndex).setValue("Email Status"); }
  
  sheet.getRange(agent._rowIndex, statusColIndex).setValue("Sent on " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy"));
  
  logAudit("Email Distributed", `Sent scorecard to ${agent['Moderator']} (${agent['Emp ID']}) on team: ${sheetName}`);
  return `✅ Scorecard sent to ${agent['Moderator']}!`;
}

// ==========================================
// AUTOMATED SIGNATURE REMINDERS ENGINE (UPDATED WITH &period= & &rowIndex=)
// ==========================================
function sendReminderEmails(teamFilters = ['ALL'], monthFilters = ['ALL']) {
  verifyAdmin();
  const data = getAgentData(teamFilters, monthFilters);
  let reminderCount = 0;
  
  data.forEach(agent => {
    const hasAck = agent['Acknowledgment Status'] && agent['Acknowledgment Status'] !== '-' && String(agent['Acknowledgment Status']).includes('Signed');
    const hasEmail = agent['Moderator Email'] && String(agent['Moderator Email']).includes('@');
    const alreadySent = agent['Email Status'] && agent['Email Status'] !== '-' && (String(agent['Email Status']).includes('Sent') || String(agent['Email Status']).includes('Reminder'));
    
    if (hasEmail && alreadySent && !hasAck) {
      const webAppUrl = ScriptApp.getService().getUrl();
      // INJECTED PERIOD AND ROW INDEX TO PREVENT ROW OVERWRITING BUG
      const ackUrl = `${webAppUrl}?action=ack&empId=${encodeURIComponent(agent['Emp ID'])}&team=${encodeURIComponent(agent._sheetName)}&period=${encodeURIComponent(agent['Month/Year'])}&rowIndex=${agent._rowIndex}`;
      
      const summaryRows = `
        <tr><td style="padding: 6px 0; color: #475569; width: 55%;"><strong>Period:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #2C3E50;">${agent['Month/Year']}</td></tr>
        <tr><td style="padding: 6px 0; color: #475569;"><strong>Status:</strong></td><td style="padding: 6px 0;"><span style="color: #C62828; font-weight: bold;">Pending E-Signature</span></td></tr>
        <tr><td style="padding: 6px 0; color: #475569;"><strong>Scorecard %:</strong></td><td style="padding: 6px 0; font-weight: bold; font-size: 15px; color: #2C3E50;">${agent['Scorecard %']}</td></tr>
      `;
      
      MailApp.sendEmail({
        to: agent['Moderator Email'].trim(),
        subject: `⏰ Reminder: Pending Scorecard Signature - ${agent['Moderator']} (${agent['Month/Year']})`,
        htmlBody: getFormalEmailHtml(agent['Moderator'], "Signature Reminder", `This is a friendly reminder that your performance scorecard for <strong>${agent['Month/Year']}</strong> is awaiting your review and digital signature. Please click the link below to complete this action.`, summaryRows, "Acknowledge & Sign Scorecard", ackUrl)
      });
      
      // Update Sheet Status
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(agent._sheetName);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      let statusColIndex = headers.findIndex(h => String(h).trim().toLowerCase().includes('email status')) + 1;
      if (statusColIndex > 0) {
        sheet.getRange(agent._rowIndex, statusColIndex).setValue("Reminder sent on " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy"));
      }
      
      reminderCount++;
    }
  });
  
  logAudit("Sent Reminders", `Distributed automated signature reminders to ${reminderCount} pending agents.`);
  return `✅ Sent reminder emails to ${reminderCount} agents awaiting signature!`;
}
