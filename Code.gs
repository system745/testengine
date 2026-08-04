/**
 * =========================================================================
 * SMART GEM ASSESSMENT ENGINE - MASTER UNIFIED SCRIPT (SECURE EDITION)
 * =========================================================================
 */

// 🟢 1. CENTRAL LICENSE CHECKER (The Security Guard)
function verifySubscription() {
  const userEmail = Session.getActiveUser().getEmail();
  
  // ⚠️ ACTION REQUIRED: Paste YOUR Web App URL from Phase 1 inside the quotes below!
  const authUrl = "https://script.google.com/macros/s/AKfycbxbDymiERLVQBw62ey2J32MiEzZD-1EXLY_JZhaozkXlKHtKQi2YJchE5RHC2E1D9E0/exec" + "?email=" + encodeURIComponent(userEmail);
  
  try {
    const res = UrlFetchApp.fetch(authUrl, { muteHttpExceptions: true });
    const authData = JSON.parse(res.getContentText());
    if (!authData.authorized) {
      SpreadsheetApp.getUi().alert("⛔ Access Denied: " + authData.message + "\nPlease contact support to activate your license.");
      return false; // Blocks the script from running!
    }
    return true; // Allows the script to continue!
  } catch (e) {
    SpreadsheetApp.getUi().alert("⚠️ License Verification Error: Check your internet connection.");
    return false;
  }
}

// 🟢 2. TEACHER API KEY SETTER (Bring Your Own Key)
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Gemini Setup', 'Paste your free Google Gemini API Key below:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', resp.getResponseText().trim());
    ui.alert('✅ Gemini API Key saved successfully!');
  }
}

// 🟢 ROUTING THE WEB APP
function doGet(e) {
  if (e && e.parameter && e.parameter.adminKey === 'TarunSecure2026') {
    return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle("Mobile Dispatch Hub")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e && e.parameter && e.parameter.view === 'practice') {
    return HtmlService.createTemplateFromFile('PracticeList').evaluate()
      .setTitle("Smart Gem Free Practice Hub")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  let template = HtmlService.createTemplateFromFile('Exam');
  template.requestedTestId = (e && e.parameter && e.parameter.testId) ? e.parameter.testId : "";
  return template.evaluate()
    .setTitle("Smart Gem Assessment Portal")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 🟢 THE CUSTOM MENU
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quiz Engine')
    .addItem('Run Stage 1: Analyze PDF Document(s)', 'stage1_AnalyzeSource')
    .addItem('Run Stage 2: Initialize Matrix Blueprint', 'stage2_BuildMatrix')
    .addItem('Check Marks Validation', 'checkMarksDistribution')
    .addSeparator()
    .addItem('Run Stage 3: Generate Hybrid Review Queue', 'stage3_GenerateReviewQueue')
    .addSeparator()
    .addItem('🚀 Run Stage 5: Deploy Unique Exam Link', 'stage5_DeployExam')
    .addSeparator()
    .addItem('📁 Run Stage 6: Organize & Rename Files', 'stage6_OrganizeFiles')
    .addSeparator()
    .addItem('🛠️ Repair / Reset Sheet Layout', 'repairLayout')
    .addToUi();
}

// 🟢 SELF-HEALING FUNCTION (Repairs accidental damage)
function repairLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    "Dashboard", "Output Selection Matrix", "Draft Review Queue", 
    "Source Text Cache", "System Answer Keys", "Active Deployments", "Master Exam Responses"
  ];
  
  requiredSheets.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (sheetName === "Master Exam Responses") {
        sheet.appendRow(["Timestamp", "Exam Name", "Student Full Name", "Mobile Number", "Email Address", "Score", "📄 PDF Report", "💬 WhatsApp", "Evaluation Status", "📧 Delivery Status", "Structured Evaluation Payload"]);
        sheet.getRange("1:1").setFontWeight("bold").setBackground("#334155").setFontColor("white");
      }
    }
  });
  SpreadsheetApp.getUi().alert("✅ Repair Complete! Missing tabs and headers have been restored.");
}

function getReportCardFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  if (!dash) throw new Error("Dashboard sheet not found!");
  let folderId = dash.getRange("B13").getValue().toString().trim();
  if (!folderId) {
    var files = DriveApp.getFileById(ss.getId()).getParents();
    if (files.hasNext()) return files.next().getId();
    throw new Error("Error: PDF Folder ID is missing in Dashboard cell B13.");
  }
  return folderId;
}

function stage1_AnalyzeSource() {
  if (!verifySubscription()) return; // 🔒 Guard blocks unauthorized users
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) { SpreadsheetApp.getUi().alert("Error: API Key missing in Script Properties."); return; }
  const sourceUrls = dash.getRange("B1").getValue().toString();
  const subject = dash.getRange("B2").getValue();
  const chapter = dash.getRange("B3").getValue();
  const customInstructions = dash.getRange("B12").getValue() || "No specific board instructions provided.";
  const allFileIds = sourceUrls.match(/[-\w]{25,}/g);
  if (!allFileIds || allFileIds.length === 0) {
    SpreadsheetApp.getUi().alert("Error: Please paste at least one valid Google Drive PDF link in cell B1.");
    return;
  }
  dash.getRange("B6").setValue(`Stage 1: Accessing ${allFileIds.length} PDF File(s)...`);
  SpreadsheetApp.flush();
  let requestParts = [{
    text: `You are an expert curriculum developer. Visually analyze the attached document(s) for the chapter "${chapter}" in the subject "${subject}".\nCRITICAL BOARD/PATTERN INSTRUCTIONS:\n${customInstructions}\nKeeping those instructions in mind, identify the distinct core technical sub-topics present.\nReturn a JSON array of objects with exactly two fields: "topic" and "suggestedWeight".\nThe weights should be integers representing percentages, and the total sum of all weights must equal exactly 100.`
  }];
  let filesProcessed = 0;
  allFileIds.forEach(id => {
    try {
      const file = DriveApp.getFileById(id);
      const base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());
      requestParts.push({ inlineData: { mimeType: "application/pdf", data: base64Pdf } });
      filesProcessed++;
    } catch(e) { console.warn(`Could not read file ID: ${id}. Check sharing permissions.`); }
  });
  if (filesProcessed === 0) {
    SpreadsheetApp.getUi().alert("Error: Could not access any of the provided PDFs.");
    dash.getRange("B6").setValue("Stage 1: Error");
    return;
  }
  let cacheSheet = ss.getSheetByName("Source Text Cache") || ss.insertSheet("Source Text Cache");
  cacheSheet.clearContents();
  cacheSheet.getRange("A1:A2").setValues([[`Successfully Linked ${filesProcessed} PDF(s)`], [sourceUrls]]);
  dash.getRange("B6").setValue(`Stage 1: AI Vision Analyzing ${filesProcessed} PDF(s)...`);
  SpreadsheetApp.flush();
  try {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const payload = { contents: [{ parts: requestParts }], generationConfig: { responseMimeType: "application/json" } };
    const options = { method: "POST", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(endpoint, options);
    if (response.getResponseCode() !== 200) throw new Error(response.getContentText());
    const topics = JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);
    let weightSheet = ss.getSheetByName("Topic Weightage") || ss.insertSheet("Topic Weightage");
    weightSheet.clearContents();
    const header = [["Sub-Topic Identified", "AI Suggested Weight (%)", "Your Final Approved Weight (%)"]];
    const rowData = topics.map(t => [t.topic, t.suggestedWeight, t.suggestedWeight]);
    const fullData = header.concat(rowData);
    weightSheet.getRange(1, 1, fullData.length, 3).setValues(fullData);
    weightSheet.getRange("A1:C1").setFontWeight("bold");
    dash.getRange("B6").setValue("Stage 2: Awaiting Weightage Approval");
    SpreadsheetApp.getUi().alert(`Stage 1 Complete!\n\nSuccessfully scanned and bundled ${filesProcessed} document(s).`);
  } catch(error) { SpreadsheetApp.getUi().alert("API Error: " + error.message); }
}

function stage2_BuildMatrix() {
  if (!verifySubscription()) return; // 🔒 Guard
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weightSheet = ss.getSheetByName("Topic Weightage");
  const dash = ss.getSheetByName("Dashboard");
  if (!weightSheet) { SpreadsheetApp.getUi().alert("Error: Run Stage 1 first."); return; }
  let matrixSheet = ss.getSheetByName("Output Selection Matrix") || ss.insertSheet("Output Selection Matrix");
  matrixSheet.clear();
  const headers = ["Topic Area", "MCQ Count", "Marks / MCQ", "Total MCQ Marks", "Assertion Count", "Marks / Assertion", "Total Assertion Marks", "Theory Count", "Marks / Theory Q", "Total Theory Marks", "Case Study Count", "Marks / Case Study", "Total Case Marks", "Overall Topic Marks", "Difficulty Level"];
  const weightData = weightSheet.getDataRange().getValues();
  let matrixData = [headers];
  for (let i = 1; i < weightData.length; i++) {
    if (weightData[i][0]) {
      const row = i + 1;
      matrixData.push([weightData[i][0], 4, 1, `=B${row}*C${row}`, 1, 1, `=E${row}*F${row}`, 1, 3, `=H${row}*I${row}`, 0, 4, `=K${row}*L${row}`, `=D${row}+G${row}+J${row}+M${row}`, "Medium"]);
    }
  }
  const last = matrixData.length;
  const sumRow = last + 1;
  matrixData.push(["🏆 GRAND TOTAL", "", "", `=SUM(D2:D${last})`, "", "", `=SUM(G2:G${last})`, "", "", `=SUM(J2:J${last})`, "", "", `=SUM(M2:M${last})`, `=SUM(N2:N${last})`, ""]);
  matrixSheet.getRange(1, 1, matrixData.length, 15).setValues(matrixData);
  matrixSheet.getRange("A1:O1").setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
  if (last > 1) {
    matrixSheet.getRange(2, 4, last - 1, 1).setBackground("#f3f3f3");
    matrixSheet.getRange(2, 7, last - 1, 1).setBackground("#f3f3f3");
    matrixSheet.getRange(2, 10, last - 1, 1).setBackground("#f3f3f3");
    matrixSheet.getRange(2, 13, last - 1, 1).setBackground("#f3f3f3");
    matrixSheet.getRange(2, 14, last - 1, 1).setBackground("#d9ead3").setFontWeight("bold");
    const diffRule = SpreadsheetApp.newDataValidation().requireValueInList(["Easy", "Medium", "High"]).build();
    matrixSheet.getRange(2, 15, last - 1, 1).setDataValidation(diffRule);
  }
  matrixSheet.autoResizeColumns(1, 15);
  dash.getRange("B6").setValue("Stage 3: Awaiting Matrix Configuration");
  SpreadsheetApp.getUi().alert("Stage 2 Complete! Matrix generated.");
}

function checkMarksDistribution() {
  if (!verifySubscription()) return; // 🔒 Guard
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const matrix = ss.getSheetByName("Output Selection Matrix");
  const targetMarks = parseInt(dash.getRange("B5").getValue()) || 0;
  const currentTotal = parseInt(matrix.getRange(matrix.getLastRow(), 14).getValue()) || 0;
  const ui = SpreadsheetApp.getUi();
  if (currentTotal === targetMarks) ui.alert("✅ PERFECT MATCH!", `Your matrix totals exactly ${currentTotal} marks.`, ui.ButtonSet.OK);
  else if (currentTotal > targetMarks) ui.alert("⚠️ OVER LIMIT!", `Target: ${targetMarks}\nCurrent: ${currentTotal}\nYou are OVER by ${currentTotal - targetMarks} marks.`, ui.ButtonSet.OK);
  else ui.alert("⚠️ UNDER LIMIT!", `Target: ${targetMarks}\nCurrent: ${currentTotal}\nYou are SHORT by ${targetMarks - currentTotal} marks.`, ui.ButtonSet.OK);
}

function stage3_GenerateReviewQueue() {
  if (!verifySubscription()) return; // 🔒 Guard
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const matrixSheet = ss.getSheetByName("Output Selection Matrix");
  const cacheSheet = ss.getSheetByName("Source Text Cache");
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!matrixSheet) { SpreadsheetApp.getUi().alert("Error: Run Stage 2 first."); return; }
  if (!apiKey) { SpreadsheetApp.getUi().alert("Error: Gemini API Key missing."); return; }
  const subject = dash.getRange("B2").getValue().toString().trim();
  const vaultUrl = dash.getRange("B14").getValue().toString().trim();
  const customInstructions = dash.getRange("B12").getValue() || "No specific board instructions provided.";
  let vaultTargetPct = dash.getRange("B15").getValue();
  if (vaultTargetPct === "" || isNaN(vaultTargetPct)) vaultTargetPct = 100;
  if (vaultTargetPct <= 1 && vaultTargetPct > 0) vaultTargetPct = vaultTargetPct * 100;
  if (!vaultUrl) { SpreadsheetApp.getUi().alert("Error: Please paste your Master Vault URL in cell B14 first."); return; }
  let vaultSheet = null;
  try {
    let vaultId = vaultUrl.match(/[-\w]{25,}/g)[0];
    vaultSheet = SpreadsheetApp.openById(vaultId).getSheetByName(subject);
  } catch(e) { SpreadsheetApp.getUi().alert("Vault Access Error: Check your link in B14 or sharing permissions."); return; }

  const matrixData = matrixSheet.getDataRange().getValues();
  let finalReviewQueueRows = [];
  let vaultPool = [];
  if (vaultSheet) {
    let vData = vaultSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (vData[i][0]) {
        vaultPool.push({
          topic: String(vData[i][0]).trim().toLowerCase(), type: String(vData[i][1]).trim(), text: String(vData[i][2]).trim(), options: String(vData[i][3]).trim(), answer: String(vData[i][4]).trim(), explanation: String(vData[i][5]).trim(), difficulty: String(vData[i][6]).trim()
        });
      }
    }
  }

  let aiGenerationManifest = [];
  let totalFromVault = 0;
  let totalFromAI = 0;

  for (let i = 1; i < matrixData.length - 1; i++) {
    let topicName = matrixData[i][0];
    if (!topicName || topicName.includes("GRAND TOTAL")) continue;
    let targetMCQ = parseInt(matrixData[i][1]) || 0, targetAR = parseInt(matrixData[i][4]) || 0, targetTheory = parseInt(matrixData[i][7]) || 0, targetCase = parseInt(matrixData[i][10]) || 0;
    let difficulty = matrixData[i][14] || "Medium";
    let counts = { "MCQ": targetMCQ, "Assertion-Reason": targetAR, "Theory": targetTheory, "Case Study": targetCase };

    Object.keys(counts).forEach(type => {
      let totalNeeded = counts[type];
      if (totalNeeded <= 0) return;
      let matches = vaultPool.filter(q => q.topic === topicName.toLowerCase().trim() && q.type === type);
      let desiredVaultCount = Math.ceil(totalNeeded * (vaultTargetPct / 100));
      let vaultAllocation = Math.min(desiredVaultCount, matches.length);
      let aiAllocation = totalNeeded - vaultAllocation;

      for (let k = 0; k < vaultAllocation; k++) {
        finalReviewQueueRows.push(["Approved", matches[k].type, topicName, matches[k].text, matches[k].options, matches[k].answer, matches[k].explanation, "[Vault]"]);
        totalFromVault++;
      }
      if (aiAllocation > 0) {
        totalFromAI += aiAllocation;
        let existingManifestItem = aiGenerationManifest.find(m => m.topic === topicName);
        if (!existingManifestItem) {
          existingManifestItem = { topic: topicName, MCQs: 0, Assertions: 0, TheoryQuestions: 0, CaseStudies: 0, Difficulty: difficulty };
          aiGenerationManifest.push(existingManifestItem);
        }
        if (type === "MCQ") existingManifestItem.MCQs += aiAllocation;
        else if (type === "Assertion-Reason") existingManifestItem.Assertions += aiAllocation;
        else if (type === "Theory") existingManifestItem.TheoryQuestions += aiAllocation;
        else if (type === "Case Study") existingManifestItem.CaseStudies += aiAllocation;
      }
    });
  }

  if (totalFromAI > 0 && aiGenerationManifest.length > 0) {
    dash.getRange("B6").setValue(`Stage 3: Pulling ${totalFromVault} from Vault. Drafting ${totalFromAI} new items via AI...`);
    SpreadsheetApp.flush();
    let requestParts = [];
    const promptText = `You are an expert curriculum developer. Based ONLY on the attached document(s), generate questions to fill the remaining empty slots of this blueprint:\n${JSON.stringify(aiGenerationManifest)}\nCRITICAL INSTRUCTIONS:\n${customInstructions}\nReturn a strict JSON array of objects. Each object MUST contain:\n"type" (MCQ, Assertion-Reason, Theory, or Case Study)\n"topic" (Exact topic name string match)\n"questionText" (The actual question text)\n"options" (Choices separated by ' | '. Blank for Theory/Case)\n"correctAnswer" (Correct option letter or comprehensive marking string)\n"explanation" (Detailed concept or step-by-step mathematical logic)`;
    requestParts.push({ text: promptText });
    const sourceUrls = cacheSheet.getRange("A2").getValue().toString();
    const allFileIds = sourceUrls.match(/[-\w]{25,}/g);
    if (allFileIds) {
      allFileIds.forEach(id => {
        try { requestParts.push({ inlineData: { mimeType: "application/pdf", data: Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes()) } }); } catch(e) {}
      });
    }
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const payload = { contents: [{ parts: requestParts }], generationConfig: { responseMimeType: "application/json", temperature: 0.2 } };
    const options = { method: "POST", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(endpoint, options);

    if (response.getResponseCode() === 200) {
      const newQuestions = JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);
      newQuestions.forEach(q => finalReviewQueueRows.push(["Pending Review", q.type, q.topic, q.questionText, q.options, q.correctAnswer, q.explanation, "[New AI Draft]"]));
    } else { SpreadsheetApp.getUi().alert("Gemini Generation Error: " + response.getContentText()); return; }
  }

  let reviewSheet = ss.getSheetByName("Draft Review Queue") || ss.insertSheet("Draft Review Queue");
  reviewSheet.clear();
  reviewSheet.appendRow(["Action", "Question Type", "Topic", "Question Text", "Options", "Correct Answer", "Explanation / Marking Key", "Source Information"]);
  reviewSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
  if (finalReviewQueueRows.length > 0) {
    reviewSheet.getRange(2, 1, finalReviewQueueRows.length, 8).setValues(finalReviewQueueRows);
    const actionRule = SpreadsheetApp.newDataValidation().requireValueInList(["Pending Review", "Approved", "Regenerate", "Reject"]).build();
    reviewSheet.getRange(2, 1, reviewSheet.getLastRow() - 1, 1).setDataValidation(actionRule).setFontWeight("bold");
  }
  dash.getRange("B6").setValue("Stage 4: Awaiting Hybrid Data Review");
  SpreadsheetApp.getUi().alert("📊 Engine Compilation Summary", `Vault Target Configured: ${vaultTargetPct}%\n\n✅ Retrieved from Vault Inventory: ${totalFromVault} Question(s)\n🤖 Dynamically Drafted by Gemini: ${totalFromAI} Question(s)\n\nCheck your 'Draft Review Queue' tab to view the results!`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function stage5_DeployExam() {
  if (!verifySubscription()) return; // 🔒 Guard
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const reviewSheet = ss.getSheetByName("Draft Review Queue");
  const matrixSheet = ss.getSheetByName("Output Selection Matrix");

  const subject = dash.getRange("B2").getValue() || "Subject";
  const chapter = dash.getRange("B3").getValue() || "Chapter";
  const targetGrade = dash.getRange("B4").getValue() || "Grade";
  const vaultUrl = dash.getRange("B14").getValue().toString().trim();
  const instName = dash.getRange("B9").getValue() || "Smart Gem";
  const teacherPhone = dash.getRange("B21").getValue().toString().replace(/[^0-9]/g, ""); // Dynamic phone capture
  const fileName = `${instName} - ${chapter}`;

  const reviewData = reviewSheet.getDataRange().getValues();
  let approvedQuestions = [];
  let newQuestionsToVault = [];
  let marksMap = {};

  if (matrixSheet) {
    let mData = matrixSheet.getDataRange().getValues();
    for(let i=1; i<mData.length-1; i++) {
      if(mData[i][0]) marksMap[mData[i][0].toLowerCase().trim()] = { MCQ: parseInt(mData[i][2])||1, AR: parseInt(mData[i][5])||1, Theory: parseInt(mData[i][8])||3, Case: parseInt(mData[i][11])||4 };
    }
  }

  for (let i = 1; i < reviewData.length; i++) {
    if (reviewData[i][0] === "Approved") {
      let qTopic = String(reviewData[i][2]).toLowerCase().trim();
      let qType = reviewData[i][1];
      let pts = 1;
      if (marksMap[qTopic]) pts = marksMap[qTopic][qType === "Assertion-Reason" ? "AR" : qType] || 1;
      let qObj = { type: qType, topic: reviewData[i][2], text: reviewData[i][3], options: reviewData[i][4] || "", points: pts, correctAnswer: reviewData[i][5], explanation: reviewData[i][6], source: reviewData[i][7] };
      approvedQuestions.push(qObj);
      if (qObj.source === "[New AI Draft]") newQuestionsToVault.push([qObj.topic, qObj.type, qObj.text, qObj.options, qObj.correctAnswer, qObj.explanation, "Medium"]);
    }
  }

  if (approvedQuestions.length === 0) { SpreadsheetApp.getUi().alert("No questions marked 'Approved'."); return; }
  dash.getRange("B6").setValue("Stage 5: Packaging Custom Exam App...");
  SpreadsheetApp.flush();

  let uniqueTestId = "SG-" + Math.floor(10000 + Math.random() * 90000);
  let deploySheet = ss.getSheetByName("Active Deployments") || ss.insertSheet("Active Deployments");
  if (deploySheet.getLastRow() === 0) {
    deploySheet.appendRow(["Deploy Date", "Unique Test ID", "Exam Name", "Exam JSON Payload"]);
    deploySheet.getRange("A1:D1").setFontWeight("bold").setBackground("#334155").setFontColor("white");
  }

  const currentDuration = dash.getRange("B11").getValue() || 45;
  let examPayload = { 
    fileName: fileName, 
    description: `${subject} Mastery - ${chapter} (${targetGrade})`, 
    questions: approvedQuestions, 
    durationMinutes: currentDuration,
    teacherPhone: teacherPhone // Injected for app-side connection
  };
  deploySheet.appendRow([new Date(), uniqueTestId, fileName, JSON.stringify(examPayload)]);

  let keySheet = ss.getSheetByName("System Answer Keys") || ss.insertSheet("System Answer Keys");
  if (keySheet.getLastRow() === 0) {
    keySheet.appendRow(["Response Tab Name", "Form Edit URL", "Question Title Lookup Key", "Original Question Title", "Correct Answer", "Explanation", "Question Type", "Question Marks"]);
    keySheet.getRange("A1:H1").setFontWeight("bold").setBackground("#334155").setFontColor("white");
  }
  const cleanStr = function(str) { return String(str).toLowerCase().replace(/[^a-z0-9]/g, "").trim(); };
  approvedQuestions.forEach(q => {
    let fullTitle = `[${q.topic}] ${q.text}`;
    keySheet.appendRow([fileName, "Custom App", cleanStr(fullTitle), fullTitle, q.correctAnswer || "", q.explanation || "", q.type, q.points]);
  });

  if (newQuestionsToVault.length > 0 && vaultUrl) {
    try {
      let vaultId = vaultUrl.match(/[-\w]{25,}/g)[0];
      let vaultSs = SpreadsheetApp.openById(vaultId);
      let vaultSheet = vaultSs.getSheetByName(subject) || vaultSs.insertSheet(subject);
      if (vaultSheet.getLastRow() === 0) vaultSheet.appendRow(["Topic", "Type", "Question", "Options", "Answer", "Explanation", "Difficulty"]);
      vaultSheet.getRange(vaultSheet.getLastRow() + 1, 1, newQuestionsToVault.length, 7).setValues(newQuestionsToVault);
    } catch (e) { console.warn("Could not reach Master Vault."); }
  }

  // Uses ScriptApp to grab the live published script URL
  let baseAppUrl = "https://script.google.com/macros/s/AKfycbyxYiRD3TL-Rz5k8fZIff0uYsr2f__OfMEUTNrfrDB_EQRvaSZX11-yc4qdD8UE0kd8/exec";
  let finalLink = baseAppUrl + "?testId=" + uniqueTestId;

  appendToArchive(fileName, finalLink);
  dash.getRange("D8").setFormula(`=HYPERLINK("${finalLink}", "🔗 Copy Student Exam Link")`);
  dash.getRange("B6").setValue("Custom App Deployed Successfully!");
  SpreadsheetApp.getUi().alert(`Deployment Complete!\n\nYour test is live. Unique Link ID: ${uniqueTestId}\n\n${newQuestionsToVault.length} AI drafts sent to Vault.`);
}

function getActiveExamSchema(testId) {
  // 🟢 Fixed: Fetch directly from central database sheet ID (works seamlessly in standalone Web App)
  const ss = SpreadsheetApp.openById("1Mu_dodcjNtc890Iac313bZOY01atjn03WE9iPfIZW6c");
  const deploySheet = ss.getSheetByName("Active Deployments");
  if (!deploySheet) return JSON.stringify({ error: "Exam system sheets not found." });
  
  const data = deploySheet.getDataRange().getValues();
  const examRow = data.find(row => row[1] == testId);
  if (!examRow) return JSON.stringify({ error: "Exam not found or inactive." });

  let schema = {};
  try {
    let rawPayload = String(examRow[3]).trim();
    if (!rawPayload.endsWith("}")) {
      let lastGoodBracket = rawPayload.lastIndexOf("}");
      if (lastGoodBracket > -1) rawPayload = rawPayload.substring(0, lastGoodBracket + 1);
    }
    schema = JSON.parse(rawPayload);
  } catch (e) {
    return JSON.stringify({ error: "Corrupted exam data package string: " + e.message });
  }
  
  if (!schema.durationMinutes) schema.durationMinutes = 45;
  return JSON.stringify(schema);
}

function getActiveTestCatalog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const managerSheet = ss.getSheetByName("Test Manager");

    if (!managerSheet) {
      return { error: "The 'Test Manager' tab was not found. Please create it." };
    }

    const data = managerSheet.getDataRange().getValues();
    const activeCatalog = [];

    for (let i = 1; i < data.length; i++) {
      let status = data[i][5] ? data[i][5].toString().trim().toLowerCase() : "";
      if (status === "active") {
        activeCatalog.push({
          className: data[i][0],
          subject:   data[i][1],
          chapter:   data[i][2],
          level:     data[i][3],
          testCode:  data[i][4]
        });
      }
    }
    return activeCatalog;
  } catch (e) {
    return { error: "Failed to read catalog: " + e.message };
  }
}

// 🟢 THE UPDATED DATABASE SUBMISSION CATCHER
function processCustomExamSubmission(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let responseSheet = ss.getSheetByName("Master Exam Responses");
    
    if (!responseSheet) throw new Error("'Master Exam Responses' sheet is missing.");

    let rowData = [
      new Date(),                // 1. Timestamp
      payload.examName,          // 2. Exam Name
      payload.name,              // 3. Student Full Name
      payload.phone,             // 4. Mobile Number
      "",                        // 5. Email Address
      "",                        // 6. Score
      "",                        // 7. PDF Report
      "",                        // 8. WhatsApp
      "Pending",                 // 9. Evaluation Status
      "",                        // 10. Delivery Status
      ""                         // 11. Structured Evaluation Payload
    ];

    if (payload.answers && payload.answers.length > 0) {
      payload.answers.forEach(ans => {
        if (typeof ans === 'object') {
          rowData.push(JSON.stringify(ans)); 
        } else {
          rowData.push(ans);
        }
      });
    }

    responseSheet.appendRow(rowData);
    
    // 🟢 THE MOCK EVENT INJECTION
    try {
      let newRowNumber = responseSheet.getLastRow();
      let mockEvent = {
        source: ss,
        range: responseSheet.getRange(newRowNumber, 1, 1, responseSheet.getLastColumn()),
        namedValues: {
          "Student Full Name": [payload.name],
          "Mobile Number": [payload.phone],
          "Timestamp": [new Date().toLocaleString()]
        }
      };
      
      // Hand the mock event directly to the PDF Engine
      autoGenerateReportCard(mockEvent);
      
    } catch (triggerError) {
      console.warn("PDF Engine Trigger failed: " + triggerError.message);
    }
    // 🟢 END OF MOCK EVENT
    
    return { success: true };
  } catch (error) {
    return { error: error.message };
  }
}

function autoGenerateReportCard(e) {
  if (!e) return;
  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    const responses = e.namedValues;
    const name = responses["Student Full Name"] ? responses["Student Full Name"][0] : "Student";
    const phone = responses["Mobile Number"] ? responses["Mobile Number"][0] : "";
    const timestamp = responses["Timestamp"] ? responses["Timestamp"][0] : new Date().toLocaleString();
    let dynamicExamName = sheet.getRange(row, 2).getValue() || sheet.getName();

    const ss = e.source;
    const dash = ss.getSheetByName("Dashboard");
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    const cleanStr = function(str) { return String(str).toLowerCase().replace(/[^a-z0-9]/g, "").trim(); };

    let masterAnswerKey = {};
    const keySheet = ss.getSheetByName("System Answer Keys");
    if (keySheet) {
      let data = keySheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        let normalizedLookupKey = cleanStr(data[i][3]);
        if (normalizedLookupKey) {
          masterAnswerKey[normalizedLookupKey] = { correctAnswer: String(data[i][4]).trim(), explanation: String(data[i][5]).trim(), qType: String(data[i][6] || "Theory").trim(), qPoints: parseFloat(data[i][7]) || 1 };
        }
      }
    }

    const docName = `[Diagnostic] ${name} - ${dynamicExamName}`;
    const doc = DocumentApp.create(docName);
    const docBody = doc.getBody();
    doc.setMarginLeft(36); doc.setMarginRight(36); doc.setMarginTop(36); doc.setMarginBottom(36);

    let brandPara = docBody.appendParagraph(dash.getRange("B9").getValue() || "Smart Gem Institute");
    brandPara.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(24).setBold(true).setFontFamily("Cambria");
    docBody.appendParagraph("Guided by Excellence").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true).editAsText().setFontSize(11).setForegroundColor("#57606a").setFontFamily("Arial");
    docBody.appendParagraph("PERFORMANCE & DIAGNOSTIC SCORECARD").setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(14).setBold(true).setFontFamily("Arial").setForegroundColor("#0f172a");
    docBody.appendHorizontalRule();

    let metaTable = docBody.appendTable([ ["📋 Candidate Info", `Name: ${name}\nMobile: ${phone}`], ["📝 Assessment Info", `Exam Scope: ${dynamicExamName}\nDate: ${timestamp}`] ]);
    metaTable.setBorderWidth(0); metaTable.setColumnWidth(0, 150); metaTable.setColumnWidth(1, 370);
    for(let r=0; r<2; r++) {
      metaTable.getRow(r).getCell(0).setPaddingTop(6).setPaddingBottom(6).setBackgroundColor("#f8fafc").editAsText().setBold(true).setFontSize(11).setFontFamily("Arial");
      metaTable.getRow(r).getCell(1).setPaddingTop(6).setPaddingBottom(6).setBackgroundColor("#ffffff").editAsText().setFontSize(11).setFontFamily("Arial");
    }
    docBody.appendParagraph("").setSpacingBefore(10);

    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let studentRowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const logisticsHeaders = ["Timestamp", "Exam Name", "Student Full Name", "Mobile Number", "Email Address", "Score", "📄 PDF Report", "💬 WhatsApp", "Evaluation Status", "📧 Delivery Status", "Structured Evaluation Payload"];

    let totalEarnedMarks = 0, totalMaxMarks = 0, qCounter = 1, evalLog = [], topicStats = {};
    let uploadFolder;
    try { uploadFolder = DriveApp.getFolderById(getReportCardFolderId()); } catch(fErr) { uploadFolder = DriveApp.getRootFolder(); }

    for (let i = 0; i < headers.length; i++) {
      let headerName = headers[i];
      if (!headerName || logisticsHeaders.includes(headerName)) continue;
      let studentAnswer = studentRowValues[i] || "[No Response]";
      let lookupKey = cleanStr(headerName);
      let keyData = masterAnswerKey[lookupKey];
      let qMaxMarks = keyData ? keyData.qPoints : 1;
      let qType = keyData ? keyData.qType : "Theory";
      let topicMatch = headerName.match(/\[(.*?)\]/);
      let topicName = topicMatch ? topicMatch[1].trim() : "General Concepts";
      let earnedForQ = 0;
      let isObjective = (qType === "MCQ" || qType === "Assertion-Reason");
      let hasImageAttachment = false, attachedFileUrl = "", textContent = String(studentAnswer);

      if (!isObjective && typeof studentAnswer === 'string' && studentAnswer.trim().indexOf('{') === 0) {
        try {
          let parsedAns = JSON.parse(studentAnswer);
          if (parsedAns.base64) {
            hasImageAttachment = true;
            textContent = parsedAns.textNotes || "[Image Attached]";
            let imgBlob = Utilities.newBlob(Utilities.base64Decode(parsedAns.base64), parsedAns.mimeType || "image/jpeg", parsedAns.filename || "submission.jpg");
            let savedFile = uploadFolder.createFile(imgBlob);
            savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            attachedFileUrl = savedFile.getUrl();
            sheet.getRange(row, i + 1).setValue(attachedFileUrl);
          }
        } catch(jsonErr) { textContent = String(studentAnswer); }
      } else if (textContent.indexOf("http") > -1 && textContent.indexOf("drive.google.com") > -1) {
        hasImageAttachment = true; attachedFileUrl = textContent.trim();
      }

      if (isObjective) {
        let isCorrect = false;
        if (studentAnswer !== "[No Response]" && keyData) {
          let normStudent = cleanStr(studentAnswer), normMaster = cleanStr(keyData.correctAnswer);
          if (normStudent === normMaster || normStudent.includes(normMaster) || normMaster.includes(normStudent)) isCorrect = true;
        }
        if (isCorrect) {
          earnedForQ = qMaxMarks;
          evalLog.push({ num: qCounter, q: headerName, a: studentAnswer, res: `🟢 Correct (${qMaxMarks}/${qMaxMarks})`, note: "" });
        } else {
          evalLog.push({ num: qCounter, q: headerName, a: studentAnswer, res: `🔴 Incorrect (0/${qMaxMarks})`, note: `Expected: ${keyData ? keyData.correctAnswer : 'N/A'}` });
        }
      }
      else {
        if (textContent === "[No Response]" || textContent.trim() === "") {
          evalLog.push({ num: qCounter, q: headerName, a: "[No Response]", res: `🔴 Skipped (0/${qMaxMarks})`, note: `Expected: ${keyData ? keyData.explanation : 'Written response required.'}` });
        }
        else if (apiKey) {
          try {
            let aiParts = [{ text: `You are an expert examiner grading a professional accounting/math test.\nQuestion: "${headerName}"\nStudent Answer Context: "${textContent}"\n${hasImageAttachment ? 'CRITICAL: The student has attached a handwritten sheet image. Read the numbers and text from the image carefully.' : ''}\nMaster Key Marking Rubric: "${keyData ? keyData.explanation : 'Check general mathematical precision and step-by-step accuracy'}"\nMax Marks: ${qMaxMarks}.\nReturn strictly JSON formatting: {"awardedMarks": float, "justification": "short clear reason"}` }];
            if (hasImageAttachment && attachedFileUrl) {
              try {
                let fileId = attachedFileUrl.match(/[-\w]{25,}/)[0];
                let blob = DriveApp.getFileById(fileId).getBlob();
                aiParts.push({ inlineData: { mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } });
              } catch(visionErr) { console.warn("Vision fallback text parsing: " + visionErr.message); }
            }
            let aiPayload = { contents: [{ parts: aiParts }], generationConfig: { responseMimeType: "application/json", temperature: 0.1 } };
            let aiRes = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`, { method: "POST", contentType: "application/json", payload: JSON.stringify(aiPayload), muteHttpExceptions: true });
            if (aiRes.getResponseCode() === 200) {
              let aiData = JSON.parse(JSON.parse(aiRes.getContentText()).candidates[0].content.parts[0].text);
              earnedForQ = parseFloat(aiData.awardedMarks) || 0;
              evalLog.push({ num: qCounter, q: headerName, a: attachedFileUrl || textContent, res: `${earnedForQ === qMaxMarks ? '🟢' : '🟡'} Evaluated (${earnedForQ}/${qMaxMarks})`, note: aiData.justification, isUrl: hasImageAttachment });
            } else { throw new Error("API Offline"); }
          } catch(err) {
            evalLog.push({ num: qCounter, q: headerName, a: attachedFileUrl || textContent, res: `🟡 Awaiting Manual Review (0/${qMaxMarks})`, note: "Evaluation Processing Engine Timeout.", isUrl: hasImageAttachment });
          }
        }
        else {
          evalLog.push({ num: qCounter, q: headerName, a: textContent, res: `🟡 Awaiting Manual Review (0/${qMaxMarks})`, note: "No active evaluation API key configured." });
        }
      }

      if (!topicStats[topicName]) { topicStats[topicName] = { earned: 0, max: 0 }; }
      topicStats[topicName].max += qMaxMarks;
      topicStats[topicName].earned += earnedForQ;
      totalEarnedMarks += earnedForQ;
      totalMaxMarks += qMaxMarks;
      qCounter++;
    }

    docBody.appendParagraph(`Overall Performance Index: ${totalEarnedMarks} / ${totalMaxMarks}`).setBold(true).setHeading(DocumentApp.ParagraphHeading.HEADING3).editAsText().setFontSize(16).setFontFamily("Arial").setForegroundColor("#1e3a8a");
    docBody.appendHorizontalRule();
    docBody.appendParagraph("📊 Concept Mastery & Remediation Plan").setBold(true).setHeading(DocumentApp.ParagraphHeading.HEADING3).editAsText().setFontSize(14).setFontFamily("Arial");
    let diagTable = docBody.appendTable();
    diagTable.setBorderWidth(1).setBorderColor("#cbd5e1");
    let thr = diagTable.appendTableRow();
    thr.appendTableCell("Syllabus Sub-Topic").setBold(true).setBackgroundColor("#1e40af").editAsText().setForegroundColor("#ffffff").setFontSize(11).setFontFamily("Arial");
    thr.appendTableCell("Proficiency").setBold(true).setBackgroundColor("#1e40af").editAsText().setForegroundColor("#ffffff").setFontSize(11).setFontFamily("Arial");
    thr.appendTableCell("Strategic Action Plan").setBold(true).setBackgroundColor("#1e40af").editAsText().setForegroundColor("#ffffff").setFontSize(11).setFontFamily("Arial");
    diagTable.setColumnWidth(0, 180); diagTable.setColumnWidth(1, 95); diagTable.setColumnWidth(2, 245);

    let topicTableData = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Topic').addColumn(Charts.ColumnType.NUMBER, 'Score (%)');

    for (let t in topicStats) {
      let pct = Math.round((topicStats[t].earned / topicStats[t].max) * 100);
      topicTableData.addRow([t, pct]);
      let statusTag, advice, colorBg;
      if (pct >= 80) { statusTag = `🟢 Ready (${pct}%)`; advice = "Concept locked. Ready for higher difficulty tests."; colorBg = "#ecfdf5"; }
      else if (pct >= 50) { statusTag = `🟡 Review (${pct}%)`; advice = "Review active notebook formulas & practice 5 standard illustrations."; colorBg = "#fffbeb"; }
      else { statusTag = `🔴 Critical (${pct}%)`; advice = "Re-read chapter fundamentals immediately. Attend doubt clearing workspace."; colorBg = "#fef2f2"; }
      let rowCells = diagTable.appendTableRow();
      rowCells.appendTableCell(t).setBackgroundColor("#ffffff").setPaddingTop(8).setPaddingBottom(8).editAsText().setFontSize(11).setBold(true).setFontFamily("Arial").setForegroundColor("#0f172a");
      rowCells.appendTableCell(statusTag).setBackgroundColor(colorBg).setPaddingTop(8).setPaddingBottom(8).editAsText().setFontSize(10).setBold(true).setFontFamily("Arial");
      rowCells.appendTableCell(advice).setBackgroundColor("#ffffff").setPaddingTop(8).setPaddingBottom(8).editAsText().setFontSize(10).setItalic(true).setFontFamily("Arial").setForegroundColor("#475569");
    }

    docBody.appendParagraph("").setSpacingBefore(12);
    try {
      let barChart = Charts.newColumnChart().setDataTable(topicTableData).setTitle('Diagnostic Analytics Index (%)').setDimensions(520, 240).setColors(['#1e40af']).setYAxisRange(0, 100).build();
      docBody.appendImage(barChart.getAs('image/png')).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } catch(e) { }

    docBody.appendHorizontalRule();
    docBody.appendParagraph("📝 Itemized Question Breakdown").setHeading(DocumentApp.ParagraphHeading.HEADING4).editAsText().setFontSize(13).setBold(true).setFontFamily("Arial");

    evalLog.forEach(log => {
      docBody.appendParagraph(`Q${log.num}. ${log.q}`).setBold(true).setSpacingBefore(16).editAsText().setFontSize(11).setFontFamily("Arial");
      if (log.isUrl) {
        docBody.appendParagraph("Answer Field: [Handwritten Image Sheet Reference Logged Below]").setItalic(true).editAsText().setFontSize(10).setForegroundColor("#64748b");
        try {
          let fileId = log.a.match(/[-\w]{25,}/)[0];
          let img = docBody.appendImage(DriveApp.getFileById(fileId).getBlob());
          if (img.getWidth() > 480) { let height = img.getHeight(); img.setWidth(480); img.setHeight((480 / img.getWidth()) * height); }
        } catch(imgErr) { docBody.appendParagraph(`Link: ${log.a}`).setLineSpacing(1.1).editAsText().setFontSize(10); }
      } else {
        docBody.appendParagraph(`Answer Given: ${log.a}`).setLineSpacing(1.3).editAsText().setFontSize(10).setFontFamily("Arial").setForegroundColor("#1e293b");
      }
      docBody.appendParagraph(log.res).setBold(true).setSpacingBefore(4).editAsText().setFontSize(10).setFontFamily("Arial");
      if(log.note) docBody.appendParagraph(`💡 Evaluation Feedback: ${log.note}`).setItalic(true).setLineSpacing(1.3).editAsText().setFontSize(10).setFontFamily("Arial").setForegroundColor("#475569");
    });

    docBody.appendHorizontalRule();
    let promoTitleText = dash.getRange("B18").getValue(), promoBodyText = dash.getRange("B19").getValue(), promoContactText = dash.getRange("B20").getValue();
    if (promoTitleText) docBody.appendParagraph(promoTitleText).setHeading(DocumentApp.ParagraphHeading.HEADING3).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(12).setFontFamily("Arial");
    if (promoBodyText) docBody.appendParagraph(promoBodyText).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true).editAsText().setFontSize(10).setFontFamily("Arial");
    if (promoContactText) docBody.appendParagraph(promoContactText).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).editAsText().setFontSize(11).setFontFamily("Arial");

    let waContactNumber = dash.getRange("B21").getValue().toString().replace(/[^0-9]/g, "");
    if (waContactNumber) {
      docBody.appendParagraph("💬 Click Here to Connect Instantly on WhatsApp").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setLinkUrl(`https://wa.me/${waContactNumber}`).editAsText().setFontSize(11).setFontFamily("Arial").setForegroundColor("#16a34a");
    }

    doc.saveAndClose(); Utilities.sleep(2000);
    const pdfUrl = uploadFolder.createFile(DriveApp.getFileById(doc.getId()).getAs('application/pdf')).setName(`${docName}.pdf`).getUrl();
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    let pdfCol = 0, waCol = 0, scoreCol = 0, evalCol = 0;
    for (let c = 0; c < headers.length; c++) {
      let hText = String(headers[c]).trim();
      if (hText.includes("PDF Report")) pdfCol = c + 1;
      if (hText.includes("WhatsApp")) waCol = c + 1;
      if (hText === "Score") scoreCol = c + 1;
      if (hText.includes("Evaluation Status")) evalCol = c + 1;
    }

    let dynamicScoreStr = `${totalEarnedMarks} / ${totalMaxMarks}`;
    if (scoreCol > 0) sheet.getRange(row, scoreCol).setValue(dynamicScoreStr);
    if (pdfCol > 0) sheet.getRange(row, pdfCol).setFormula(`=HYPERLINK("${pdfUrl}", "📥 View PDF")`);
    if (waCol > 0) sheet.getRange(row, waCol).setFormula(`=HYPERLINK("https://api.whatsapp.com/send?phone=91${phone}&text=${encodeURIComponent(`Test Result: ${dynamicScoreStr}\n\nPDF: ${pdfUrl}`)}", "📱 Send WhatsApp")`);
    if (evalCol > 0) sheet.getRange(row, evalCol).setValue("✅ Done");
  } catch (err) { e.range.getSheet().getRange(e.range.getRow(), e.range.getSheet().getLastColumn() + 1).setValue("❌ ERROR: " + err.message); }
}

function appendToArchive(testName, testLink) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Test Archive") || ss.insertSheet("Test Archive");
  if (sheet.getLastRow() === 0) sheet.getRange("A1:C1").setValues([["Date", "Test Name", "Test Link"]]).setFontWeight("bold").setBackground("#e2e8f0");
  sheet.appendRow([new Date(), testName, testLink]);
}

function stage6_OrganizeFiles() {
  if (!verifySubscription()) return; // 🔒 Guard
  const dash = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  dash.getRange("B6").setValue("Stage 6: Cleaning workspace...");
  const coreTabs = ["Dashboard", "Output Selection Matrix", "Draft Review Queue", "Source Text Cache", "System Answer Keys", "Active Deployments", "Master Exam Responses"];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => { if (!coreTabs.includes(sheet.getName())) sheet.hideSheet(); });
  dash.getRange("B6").setValue("Workspace Cleanup Completed!");
}

function getAvailablePracticeTests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deploySheet = ss.getSheetByName("Active Deployments");
  if (!deploySheet) return JSON.stringify([]);
  const data = deploySheet.getDataRange().getValues();
  let activeHubList = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    let descriptionStr = "", parsedClass = "General", parsedTopic = "Mixed Concepts", calcLevel = "🟡 Medium";
    try {
      let payload = JSON.parse(data[i][3]);
      descriptionStr = payload.description || "";
      let match = descriptionStr.match(/(.*?) Mastery - (.*?) \((.*?)\)/);
      if (match) { parsedTopic = match[2].trim(); parsedClass = match[3].trim(); }
      else if (payload.fileName) { parsedTopic = payload.fileName.replace("Classes by CMA Tarun Singhal - ", "").trim(); }

      if (payload.questions && payload.questions.length > 0) {
        let avgPoints = payload.questions.reduce((sum, q) => sum + (q.points || 1), 0) / payload.questions.length;
        if (avgPoints <= 1.5) calcLevel = "🟢 Easy"; else if (avgPoints <= 2.8) calcLevel = "🟡 Medium";
        else calcLevel = "🔴 Tough";
      }
    } catch(err) { descriptionStr = "Practice Assessment Workspace Available."; }
    activeHubList.push({ testId: String(data[i][1]).toUpperCase(), examName: String(data[i][2]), targetClass: parsedClass, topic: parsedTopic, level: calcLevel });
  }
  return JSON.stringify(activeHubList.reverse());
}

function getMobileHubData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const instName = dash ? (dash.getRange("B9").getValue() || "Smart Gem") : "Smart Gem";
  let results = [];

  const responseSheet = ss.getSheetByName("Master Exam Responses");
  if (!responseSheet) return JSON.stringify({ instName: instName, records: [] });

  const data = responseSheet.getDataRange().getValues();
  const formulas = responseSheet.getDataRange().getFormulas();
  if (data.length <= 1) return JSON.stringify({ instName: instName, records: [] });

  const timeIdx = data[0].indexOf("Timestamp");
  const nameIdx = data[0].indexOf("Student Full Name");
  const mobileIdx = data[0].indexOf("Mobile Number"); 
  const waIdx = data[0].indexOf("💬 WhatsApp");
  const examNameIdx = data[0].indexOf("Exam Name");

  if (nameIdx !== -1 && waIdx !== -1) {
    for (let i = 1; i < data.length; i++) {
      let studentName = data[i][nameIdx];
      let formula = formulas[i][waIdx];
      let actualExamName = (examNameIdx !== -1 && data[i][examNameIdx]) ? data[i][examNameIdx] : "General Assessment";
      let mobileNum = (mobileIdx !== -1 && data[i][mobileIdx]) ? String(data[i][mobileIdx]) : "N/A";

      let rawTime = 0;
      if (timeIdx !== -1 && data[i][timeIdx]) {
        let parsedDate = new Date(data[i][timeIdx]).getTime();
        if (!isNaN(parsedDate)) rawTime = parsedDate;
      }

      if (!studentName || !formula) continue;
      let urlMatch = formula.match(/"([^"]+)"/);
      if (urlMatch && urlMatch[1] !== "#") {
        results.push({
          timestamp: rawTime,
          sheetName: actualExamName,
          studentName: studentName,
          mobile: mobileNum, 
          waUrl: urlMatch[1]
        });
      }
    }
  }
  results.sort((a, b) => b.timestamp - a.timestamp);
  return JSON.stringify({ instName: instName, records: results });
}
