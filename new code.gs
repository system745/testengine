/**
 * =========================================================================
 * SMART GEM CENTRAL LIBRARY (ENGINE) - PART 1: UTILITIES & ONBOARDING
 * =========================================================================
 */

// ⚠️ IMPORTANT: Paste your Central Web App URL (from Part 1) inside the quotes below!
const CENTRAL_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbyxYiRD3TL-Rz5k8fZIff0uYsr2f__OfMEUTNrfrDB_EQRvaSZX11-yc4qdD8UE0kd8/exec";

// 🟢 HELPER: EXTRACT CLEAN ID FROM URL OR RAW ID
function extractIdFromInput(input) {
  if (!input) return "";
  const str = String(input).trim();
  const match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

// 🟢 1. CENTRAL LICENSE CHECKER (The Security Guard)
function verifySubscription() {
  const userEmail = Session.getActiveUser().getEmail();
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

// 🟢 2. TEACHER API KEY SETTER
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Gemini Setup', 'Paste your free Google Gemini API Key below:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', resp.getResponseText().trim());
    ui.alert('✅ Gemini API Key saved successfully!');
  }
}

// 🟢 3. SELF-HEALING FUNCTION (Repairs accidental damage)
function repairLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    "Dashboard", "Output Selection Matrix", "Draft Review Queue", 
    "Source Text Cache", "System Answer Keys", "Master Exam Responses", "Test Archive"
  ];
  
  requiredSheets.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (sheetName === "Master Exam Responses") {
        sheet.appendRow(["Timestamp", "Exam Name", "Student Full Name", "Mobile Number", "Email Address", "Score", "📄 PDF Report", "💬 WhatsApp", "Evaluation Status", "📧 Delivery Status", "Structured Evaluation Payload"]);
        sheet.getRange("1:1").setFontWeight("bold").setBackground("#334155").setFontColor("white");
      } else if (sheetName === "Test Archive") {
        sheet.appendRow(["Date", "Test Name", "Test Link"]);
        sheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
      }
    }
  });
  SpreadsheetApp.getUi().alert("✅ Repair Complete! Missing tabs and headers have been restored.");
}

// 🟢 4. ONE-CLICK ONBOARDING: Builds folders & Vault automatically
function initializeQuizEngine() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  
  if (!dash) { ui.alert("Error: Dashboard tab missing."); return; }

  const existingFolder = dash.getRange("B13").getValue();
  const existingVault = dash.getRange("B14").getValue();
  
  if (existingFolder || existingVault) {
    const response = ui.alert("Already Initialized", "It looks like your workspace is already set up!\n\nDo you want to wipe the current links and generate brand new folders?", ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  dash.getRange("B6").setValue("Initializing Workspace... Please wait.");
  SpreadsheetApp.flush();

  try {
    const mainFolder = DriveApp.createFolder("Smart Gem Assessment Workspace");
    const outputFolder = mainFolder.createFolder("1. Output Question Papers");
    const scoreFolder = mainFolder.createFolder("2. Student Scorecards");
    
    const vaultApp = SpreadsheetApp.create("Master Question Vault");
    const vaultId = vaultApp.getId();
    DriveApp.getFileById(vaultId).moveTo(mainFolder);
    
    let defaultSheet = vaultApp.getSheets()[0];
    defaultSheet.setName("General Topics");
    defaultSheet.appendRow(["Topic", "Type", "Question", "Options", "Answer", "Explanation", "Difficulty"]);
    defaultSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#334155").setFontColor("white");
    vaultApp.setFrozenRows(1);
    
    dash.getRange("B13").setValue(outputFolder.getUrl());
    dash.getRange("B14").setValue(vaultApp.getUrl());
    dash.getRange("B22").setValue(scoreFolder.getUrl());
    
    repairLayout(); 
    dash.getRange("B6").setValue("System Ready!");
    ui.alert("✅ Initialization Complete!", "Your custom folders and Master Vault have been automatically created and linked to your dashboard. You are ready to start generating tests!", ui.ButtonSet.OK);
  } catch (error) {
    dash.getRange("B6").setValue("Error during initialization.");
    ui.alert("❌ Error: " + error.message);
  }
}

function getReportCardFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  if (!dash) throw new Error("Dashboard sheet not found!");
  let rawFolderInput = dash.getRange("B13").getValue().toString().trim();
  let folderId = extractIdFromInput(rawFolderInput);
  if (!folderId) {
    var files = DriveApp.getFileById(ss.getId()).getParents();
    if (files.hasNext()) return files.next().getId();
    throw new Error("Error: PDF Folder Link is missing in Dashboard cell B13.");
  }
  return folderId;
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

function stage6_OrganizeFiles() {
  if (!verifySubscription()) return; // 🔒 Guard
  const dash = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  dash.getRange("B6").setValue("Stage 6: Cleaning workspace...");
  const coreTabs = ["Dashboard", "Output Selection Matrix", "Draft Review Queue", "Source Text Cache", "System Answer Keys", "Master Exam Responses", "Test Archive"];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => { if (!coreTabs.includes(sheet.getName())) sheet.hideSheet(); });
  dash.getRange("B6").setValue("Workspace Cleanup Completed!");
}
/**
 * =========================================================================
 * SMART GEM CENTRAL LIBRARY (ENGINE) - PART 2: AI SCAN & MATRIX BUILDER
 * =========================================================================
 */

// 🟢 STAGE 1: AI DOCUMENT SCAN & WEIGHTAGE ASSIGNMENT
function stage1_AnalyzeSource() {
  if (!verifySubscription()) return; // 🔒 Guard
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const ui = SpreadsheetApp.getUi();
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    ui.alert("⚠️ Missing API Key", "Please run 'Quiz Engine > Setup > Set Gemini API Key' first.", ui.ButtonSet.OK);
    return;
  }
  
  const sourceUrls = dash.getRange("B1").getValue().toString();
  const subject = dash.getRange("B2").getValue();
  const chapter = dash.getRange("B3").getValue();
  
  if (!sourceUrls) {
    ui.alert("⚠️ Missing Source", "Please paste a Google Drive link to your PDF(s) in cell B1.", ui.ButtonSet.OK);
    return;
  }
  
  // Extract all file IDs from the comma-separated or multi-line URLs
  const allFileIds = sourceUrls.match(/[-\w]{25,}/g);
  if (!allFileIds || allFileIds.length === 0) {
    ui.alert("⚠️ Invalid Link", "Could not extract a valid Google Drive File ID. Ensure it is a valid shareable link.", ui.ButtonSet.OK);
    return;
  }
  
  dash.getRange("B6").setValue("Stage 1: AI Vision Analyzing Document(s)...");
  SpreadsheetApp.flush();
  
  // Prepare Gemini Request
  let requestParts = [{
    text: `Analyze the attached document(s) for the chapter "${chapter}" in the subject "${subject}". Identify the distinct core technical sub-topics present. 
    Return a JSON array of objects with exactly two fields: "topic" and "suggestedWeight". 
    The weights should be integers representing percentages, summing to 100. DO NOT use markdown formatting, output pure JSON.`
  }];
  
  let validFilesCount = 0;
  
  // Attach all PDFs to the prompt
  allFileIds.forEach(id => {
    try {
      let file = DriveApp.getFileById(id);
      let blob = file.getBlob();
      let bytes = blob.getBytes();
      let encoded = Utilities.base64Encode(bytes);
      requestParts.push({
        inlineData: {
          mimeType: "application/pdf",
          data: encoded
        }
      });
      validFilesCount++;
    } catch(e) {
      console.warn("Failed to attach file ID: " + id + " - Error: " + e.message);
    }
  });
  
  if (validFilesCount === 0) {
    ui.alert("⚠️ Permission Error", "Could not read the PDF(s). Please ensure 'system@smartgem.in' or your account has viewer access to the files.", ui.ButtonSet.OK);
    dash.getRange("B6").setValue("Stage 1 Failed - Permission Error");
    return;
  }
  
  try {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const payload = {
      contents: [{ parts: requestParts }],
      generationConfig: { responseMimeType: "application/json" }
    };
    
    const response = UrlFetchApp.fetch(endpoint, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(response.getContentText());
    }
    
    const responseObj = JSON.parse(response.getContentText());
    let aiText = responseObj.candidates[0].content.parts[0].text;
    
    // Clean up potential markdown formatting from Gemini
    aiText = aiText.replace(/```json/g, "").replace(/```/g, "").trim();
    const topics = JSON.parse(aiText);
    
    // Create or clear Topic Weightage tab
    let weightSheet = ss.getSheetByName("Topic Weightage");
    if (!weightSheet) {
      weightSheet = ss.insertSheet("Topic Weightage");
    } else {
      weightSheet.clearContents();
    }
    
    const header = [["Sub-Topic Identified", "AI Suggested Weight (%)", "Your Final Approved Weight (%)"]];
    const rowData = topics.map(t => [t.topic, t.suggestedWeight, t.suggestedWeight]);
    
    weightSheet.getRange(1, 1, header.length + rowData.length, 3).setValues(header.concat(rowData));
    weightSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#cfe2f3");
    weightSheet.autoResizeColumns(1, 3);
    
    // Save to Cache for Stage 3
    let cacheSheet = ss.getSheetByName("Source Text Cache");
    if (!cacheSheet) {
      cacheSheet = ss.insertSheet("Source Text Cache");
    } else {
      cacheSheet.clearContents();
    }
    cacheSheet.getRange("A1:A2").setValues([
      [`Successfully Linked ${validFilesCount} PDF(s)`],
      [sourceUrls] // Store raw input to extract IDs later
    ]);
    cacheSheet.hideSheet(); // Hide it to keep UI clean
    
    dash.getRange("B6").setValue("Stage 2: Awaiting Weightage Approval");
    ui.alert("✅ Stage 1 Complete!", `Successfully analyzed ${validFilesCount} document(s).\n\nPlease review the 'Topic Weightage' tab and adjust the final percentages if needed before running Stage 2.`, ui.ButtonSet.OK);
    
  } catch(error) {
    dash.getRange("B6").setValue("Stage 1 Error");
    ui.alert("❌ API Error: " + error.message);
  }
}

// 🟢 STAGE 2: BUILD SELECTION MATRIX
function stage2_BuildMatrix() {
  if (!verifySubscription()) return; // 🔒 Guard
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weightSheet = ss.getSheetByName("Topic Weightage");
  const dash = ss.getSheetByName("Dashboard");
  const ui = SpreadsheetApp.getUi();
  
  if (!weightSheet) {
    ui.alert("⚠️ Missing Data", "Please run Stage 1 first to generate the topic weightage.", ui.ButtonSet.OK);
    return;
  }
  
  dash.getRange("B6").setValue("Stage 2: Constructing Matrix...");
  
  let matrixSheet = ss.getSheetByName("Output Selection Matrix");
  if (!matrixSheet) {
    matrixSheet = ss.insertSheet("Output Selection Matrix");
  } else {
    matrixSheet.clear();
  }
  
  const headers = [
    "Topic Area", 
    "MCQ Count", "Marks / MCQ", "Total MCQ Marks",
    "Assertion Count", "Marks / Assertion", "Total Assertion Marks",
    "Theory Count", "Marks / Theory Q", "Total Theory Marks",
    "Case Study Count", "Marks / Case Study", "Total Case Marks",
    "Overall Topic Marks", "Difficulty Level"
  ];
  
  const weightData = weightSheet.getDataRange().getValues();
  let matrixData = [headers];
  
  // Starting rows at 2 because headers are row 1
  for (let i = 1; i < weightData.length; i++) {
    let topic = weightData[i][0];
    if (topic) {
      const row = i + 1;
      matrixData.push([
        topic, 
        4, 1, `=B${row}*C${row}`,  // MCQ default: 4 Qs, 1 Mark
        1, 1, `=E${row}*F${row}`,  // Assertion default: 1 Q, 1 Mark
        1, 3, `=H${row}*I${row}`,  // Theory default: 1 Q, 3 Marks
        0, 4, `=K${row}*L${row}`,  // Case Study default: 0 Qs, 4 Marks
        `=D${row}+G${row}+J${row}+M${row}`, // Total Marks per topic
        "Medium" // Default Difficulty
      ]);
    }
  }
  
  const last = matrixData.length;
  // Grand total row
  matrixData.push([
    "🏆 GRAND TOTAL", 
    "", "", `=SUM(D2:D${last})`, 
    "", "", `=SUM(G2:G${last})`, 
    "", "", `=SUM(J2:J${last})`, 
    "", "", `=SUM(M2:M${last})`, 
    `=SUM(N2:N${last})`, 
    ""
  ]);
  
  matrixSheet.getRange(1, 1, matrixData.length, 15).setValues(matrixData);
  matrixSheet.getRange("A1:O1").setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
  matrixSheet.getRange(matrixData.length, 1, 1, 15).setFontWeight("bold").setBackground("#d9ead3");
  
  // Apply data validation for Difficulty
  const diffRule = SpreadsheetApp.newDataValidation().requireValueInList(["Easy", "Medium", "Hard"]).build();
  matrixSheet.getRange(2, 15, matrixData.length - 2, 1).setDataValidation(diffRule);
  
  matrixSheet.autoResizeColumns(1, 15);
  dash.getRange("B6").setValue("Stage 3: Awaiting Matrix Configuration");
  
  ui.alert("✅ Stage 2 Complete!", "Your Assessment Blueprint Matrix has been generated.\n\nPlease review the 'Output Selection Matrix' tab and adjust the question counts per topic before running Stage 3.", ui.ButtonSet.OK);
}

/**
 * =========================================================================
 * SMART GEM CENTRAL LIBRARY (ENGINE) - PART 3: HYBRID QUESTION GENERATION
 * =========================================================================
 */

// 🟢 STAGE 3: GENERATE HYBRID QUEUE (VAULT + AI)
function stage3_GenerateReviewQueue() {
  if (!verifySubscription()) return; // 🔒 Guard
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const ui = SpreadsheetApp.getUi();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    ui.alert("⚠️ Missing API Key", "Please set your Gemini API key in the setup menu.", ui.ButtonSet.OK);
    return;
  }
  
  let matrixSheet = ss.getSheetByName("Output Selection Matrix");
  if (!matrixSheet) {
    ui.alert("⚠️ Missing Matrix", "Please run Stage 2 first.", ui.ButtonSet.OK);
    return;
  }
  
  dash.getRange("B6").setValue("Stage 3: Searching Vault & Drafting new AI items...");
  SpreadsheetApp.flush();
  
  const matrixData = matrixSheet.getDataRange().getValues();
  const vaultUrl = dash.getRange("B14").getValue();
  let vaultQuestions = [];
  
  // 1. Try to pull existing questions from the Teacher's Master Vault
  try {
    if (vaultUrl) {
      let vaultId = extractIdFromInput(vaultUrl);
      let vaultApp = SpreadsheetApp.openById(vaultId);
      let vaultSheet = vaultApp.getSheets()[0]; // Assuming first tab is the active one
      let vaultData = vaultSheet.getDataRange().getValues();
      
      // Vault Headers assumed: Topic | Type | Question | Options | Answer | Explanation | Difficulty
      for (let i = 1; i < vaultData.length; i++) {
        if (vaultData[i][2]) { // If question text exists
          vaultQuestions.push({
            topic: String(vaultData[i][0]).trim(),
            type: String(vaultData[i][1]).trim(),
            text: String(vaultData[i][2]).trim(),
            options: String(vaultData[i][3]).trim(),
            answer: String(vaultData[i][4]).trim(),
            explanation: String(vaultData[i][5]).trim(),
            difficulty: String(vaultData[i][6]).trim()
          });
        }
      }
    }
  } catch (err) {
    console.warn("Vault access error (continuing with AI only): " + err.message);
  }
  
  let aiGenerationManifest = [];
  let finalReviewQueueRows = [];
  
  // 2. Cross-check Matrix against Vault, find shortages for AI to generate
  for (let i = 1; i < matrixData.length - 1; i++) { // Skip header and Grand Total
    let topicName = matrixData[i][0];
    if (!topicName || topicName.includes("GRAND TOTAL")) continue;
    
    let reqMCQ = parseInt(matrixData[i][1]) || 0;
    let reqAssertion = parseInt(matrixData[i][4]) || 0;
    let reqTheory = parseInt(matrixData[i][7]) || 0;
    let reqCase = parseInt(matrixData[i][10]) || 0;
    let targetDiff = String(matrixData[i][14]).trim();
    
    let missingMCQ = reqMCQ, missingAssertion = reqAssertion, missingTheory = reqTheory, missingCase = reqCase;
    
    // Check Vault for matches
    for (let v = 0; v < vaultQuestions.length; v++) {
      let q = vaultQuestions[v];
      if (q.topic.toLowerCase() === topicName.toLowerCase() && (targetDiff === "Any" || q.difficulty.toLowerCase() === targetDiff.toLowerCase())) {
        if (q.type.toLowerCase().includes("mcq") && missingMCQ > 0) {
          finalReviewQueueRows.push(["Pending Review", "MCQ", q.topic, q.text, q.options, q.answer, q.explanation, "[Vault]"]);
          missingMCQ--;
        } else if (q.type.toLowerCase().includes("assertion") && missingAssertion > 0) {
          finalReviewQueueRows.push(["Pending Review", "Assertion-Reason", q.topic, q.text, q.options, q.answer, q.explanation, "[Vault]"]);
          missingAssertion--;
        } else if (q.type.toLowerCase().includes("theory") && missingTheory > 0) {
          finalReviewQueueRows.push(["Pending Review", "Theory", q.topic, q.text, q.options, q.answer, q.explanation, "[Vault]"]);
          missingTheory--;
        } else if (q.type.toLowerCase().includes("case") && missingCase > 0) {
          finalReviewQueueRows.push(["Pending Review", "Case Study", q.topic, q.text, q.options, q.answer, q.explanation, "[Vault]"]);
          missingCase--;
        }
      }
    }
    
    // If Vault didn't have enough, ask AI to make the rest
    if (missingMCQ > 0 || missingAssertion > 0 || missingTheory > 0 || missingCase > 0) {
      aiGenerationManifest.push({
        topic: topicName,
        MCQs: missingMCQ,
        Assertions: missingAssertion,
        TheoryQuestions: missingTheory,
        CaseStudies: missingCase,
        Difficulty: targetDiff
      });
    }
  }
  
  // 3. Generate Missing Questions via Gemini AI
  if (aiGenerationManifest.length > 0) {
    dash.getRange("B6").setValue("Drafting " + aiGenerationManifest.length + " topic(s) via Gemini AI...");
    SpreadsheetApp.flush();
    
    let requestParts = [{
      text: `You are an expert curriculum developer. Generate questions to exactly match this shortage blueprint:\n${JSON.stringify(aiGenerationManifest)}\n
      CRITICAL INSTRUCTIONS:
      1. Return a strict JSON array of objects.
      2. Required fields for each object: "type" (MCQ, Assertion-Reason, Theory, Case Study), "topic", "questionText", "options" (separated by | if applicable, else empty), "correctAnswer", "explanation".
      3. Ensure high academic quality suitable for the difficulty level requested. DO NOT wrap in markdown formatting.`
    }];
    
    // Attach PDFs from Cache again so AI has context
    const cacheSheet = ss.getSheetByName("Source Text Cache");
    if (cacheSheet) {
      const sourceUrls = cacheSheet.getRange("A2").getValue().toString();
      const allFileIds = sourceUrls.match(/[-\w]{25,}/g);
      if (allFileIds) {
        allFileIds.forEach(id => {
          try {
            requestParts.push({
              inlineData: { mimeType: "application/pdf", data: Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes()) }
            });
          } catch(e) {}
        });
      }
    }
    
    try {
      const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
      const payload = {
        contents: [{ parts: requestParts }],
        generationConfig: { responseMimeType: "application/json" }
      };
      
      const response = UrlFetchApp.fetch(endpoint, {
        method: "POST",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        let aiText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
        aiText = aiText.replace(/```json/g, "").replace(/```/g, "").trim();
        let aiQuestions = JSON.parse(aiText);
        
        aiQuestions.forEach(q => {
          finalReviewQueueRows.push([
            "Pending Review", 
            q.type, 
            q.topic, 
            q.questionText, 
            q.options || "", 
            q.correctAnswer || "", 
            q.explanation || "", 
            "[New AI Draft]"
          ]);
        });
      } else {
        throw new Error(response.getContentText());
      }
    } catch(err) {
      ui.alert("⚠️ AI Generation Warning", "Some AI questions failed to generate: " + err.message, ui.ButtonSet.OK);
    }
  }
  
  // 4. Output everything to Draft Review Queue
  let reviewSheet = ss.getSheetByName("Draft Review Queue");
  if (!reviewSheet) {
    reviewSheet = ss.insertSheet("Draft Review Queue");
  } else {
    reviewSheet.clear();
  }
  
  const headers = ["Action", "Question Type", "Topic", "Question Text", "Options", "Correct Answer", "Explanation / Marking Key", "Source Information"];
  reviewSheet.appendRow(headers);
  reviewSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
  
  if (finalReviewQueueRows.length > 0) {
    reviewSheet.getRange(2, 1, finalReviewQueueRows.length, 8).setValues(finalReviewQueueRows);
    
    // Add dropdowns for the Action column
    const actionRule = SpreadsheetApp.newDataValidation().requireValueInList(["Pending Review", "Approved", "Regenerate", "Reject"]).build();
    reviewSheet.getRange(2, 1, reviewSheet.getLastRow() - 1, 1).setDataValidation(actionRule).setFontWeight("bold");
    
    // Auto format
    reviewSheet.setColumnWidth(4, 300); // Make Question text wider
    reviewSheet.setColumnWidth(7, 300); // Make explanation wider
    reviewSheet.getRange(2, 1, reviewSheet.getLastRow(), reviewSheet.getLastColumn()).setWrap(true).setVerticalAlignment("middle");
  }
  
  dash.getRange("B6").setValue("Stage 4: Awaiting Hybrid Data Review");
  ui.alert("✅ Stage 3 Complete!", `Drafted ${finalReviewQueueRows.length} total questions.\n\nPlease go to the 'Draft Review Queue' tab and change the action to 'Approved' for the questions you want in the final exam.`, ui.ButtonSet.OK);
}

/**
 * =========================================================================
 * SMART GEM CENTRAL LIBRARY (ENGINE) - PART 4: DEPLOYMENT & SYNC
 * =========================================================================
 */

// 🟢 STAGE 5: DEPLOY EXAM TO CENTRAL SERVER
function stage5_DeployExam() {
  if (!verifySubscription()) return; // 🔒 Guard
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const reviewSheet = ss.getSheetByName("Draft Review Queue");
  const ui = SpreadsheetApp.getUi();
  
  if (!reviewSheet) {
    ui.alert("⚠️ Missing Data", "Draft Review Queue not found. Run Stage 3 first.", ui.ButtonSet.OK);
    return;
  }
  
  const reviewData = reviewSheet.getDataRange().getValues();
  const teacherEmail = Session.getActiveUser().getEmail(); // 🔐 SECURE TEACHER ID
  const subject = dash.getRange("B2").getValue() || "Subject";
  const chapter = dash.getRange("B3").getValue() || "Chapter";
  const fileName = `${dash.getRange("B9").getValue() || "Smart Gem"} - ${chapter}`;
  const uniqueTestId = "SG-" + Math.floor(10000 + Math.random() * 90000);
  
  let approvedQuestions = [];
  
  // Extract only approved questions
  for (let i = 1; i < reviewData.length; i++) {
    if (reviewData[i][0] === "Approved") {
      approvedQuestions.push({
        type: reviewData[i][1],
        topic: reviewData[i][2],
        text: reviewData[i][3],
        options: reviewData[i][4] || "",
        points: (String(reviewData[i][1]).includes("MCQ") || String(reviewData[i][1]).includes("Assertion")) ? 1 : (String(reviewData[i][1]).includes("Theory") ? 3 : 4),
        correctAnswer: reviewData[i][5],
        explanation: reviewData[i][6]
      });
    }
  }

  if (approvedQuestions.length === 0) {
    ui.alert("⚠️ No Approved Questions", "Please mark at least one question as 'Approved' in the Draft Review Queue.", ui.ButtonSet.OK);
    return;
  }
  
  dash.getRange("B6").setValue("Packaging & Sending to Central Server...");
  SpreadsheetApp.flush();

  const examPayload = { 
    fileName: fileName, 
    description: `${subject} Mastery - ${chapter}`, 
    questions: approvedQuestions, 
    durationMinutes: dash.getRange("B11").getValue() || 45,
    teacherPhone: dash.getRange("B21").getValue().toString().replace(/[^0-9]/g, "")
  };

  // 📡 SEND TO CENTRAL WEB APP
  try {
    const res = UrlFetchApp.fetch(CENTRAL_WEBAPP_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ 
        action: "deployExam", 
        testId: uniqueTestId, 
        fileName: fileName, 
        examPayload: examPayload, 
        teacherEmail: teacherEmail 
      }),
      muteHttpExceptions: true
    });
    
    if (res.getResponseCode() !== 200 || JSON.parse(res.getContentText()).status !== "success") {
      throw new Error("Server rejected deployment. Response: " + res.getContentText());
    }
  } catch(e) { 
    dash.getRange("B6").setValue("Deployment Failed");
    return ui.alert("❌ Deployment Error: " + e.message); 
  }

  // Update Answer Keys Locally for grading later (So the PDF Generator knows the answers)
  let keySheet = ss.getSheetByName("System Answer Keys");
  if (!keySheet) {
    keySheet = ss.insertSheet("System Answer Keys");
  }
  
  if (keySheet.getLastRow() === 0) {
    keySheet.appendRow(["Response Tab Name", "Form Edit URL", "Question Title Lookup Key", "Original Question Title", "Correct Answer", "Explanation", "Question Type", "Question Marks"]);
    keySheet.getRange("A1:H1").setFontWeight("bold").setBackground("#334155").setFontColor("white");
  }
  
  const cleanStr = function(str) { return String(str).toLowerCase().replace(/[^a-z0-9]/g, "").trim(); };
  
  approvedQuestions.forEach(q => {
    let fullTitle = `[${q.topic}] ${q.text}`;
    keySheet.appendRow([
      uniqueTestId, // Use TestID to link answers exactly
      "Web App Interface", 
      cleanStr(fullTitle), 
      fullTitle, 
      q.correctAnswer || "", 
      q.explanation || "", 
      q.type, 
      q.points
    ]);
  });

  // Log in Archive
  let archiveSheet = ss.getSheetByName("Test Archive");
  if (!archiveSheet) { archiveSheet = ss.insertSheet("Test Archive"); }
  
  let finalLink = CENTRAL_WEBAPP_URL + "?testId=" + uniqueTestId;
  archiveSheet.appendRow([new Date(), fileName, finalLink]);
  
  // Update Dashboard link
  dash.getRange("D8").setFormula(`=HYPERLINK("${finalLink}", "🔗 Copy Student Exam Link")`);
  dash.getRange("B6").setValue("Deployed Successfully!");
  
  ui.alert("🚀 Test is Live!", `Your assessment has been successfully deployed.\n\nTest ID: ${uniqueTestId}\n\nStudents can now take the exam using the link in cell D8.`, ui.ButtonSet.OK);
}

// 🟢 📥 PULL DATA FROM SERVER (SYNC NEW SUBMISSIONS)
function syncStudentResults() {
  if (!verifySubscription()) return; // 🔒 Guard
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard");
  const ui = SpreadsheetApp.getUi();
  const teacherEmail = Session.getActiveUser().getEmail();
  
  dash.getRange("B6").setValue("Syncing new results from server...");
  SpreadsheetApp.flush();

  try {
    const res = UrlFetchApp.fetch(CENTRAL_WEBAPP_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ action: "syncResults", teacherEmail: teacherEmail }),
      muteHttpExceptions: true
    });
    
    const responseData = JSON.parse(res.getContentText());
    if (responseData.status !== "success") {
      throw new Error(responseData.message || "Unknown Server Error");
    }

    const records = responseData.records;
    if (!records || records.length === 0) {
      dash.getRange("B6").setValue("System Ready!");
      ui.alert("📥 No New Submissions", "Your local database is completely up to date.", ui.ButtonSet.OK);
      return;
    }

    let responseSheet = ss.getSheetByName("Master Exam Responses");
    if (!responseSheet) {
      repairLayout();
      responseSheet = ss.getSheetByName("Master Exam Responses");
    }

    let successCount = 0;

    // Process each new record fetched from the Central Server
    records.forEach(record => {
      let rowData = [
        record.timestamp,          // 1. Timestamp
        record.testId,             // 2. Exam Name (Test ID)
        record.studentName,        // 3. Student Full Name
        record.studentPhone,       // 4. Mobile Number
        "",                        // 5. Email Address
        "",                        // 6. Score
        "",                        // 7. PDF Report
        "",                        // 8. WhatsApp
        "Pending",                 // 9. Evaluation Status
        "",                        // 10. Delivery Status
        JSON.stringify(record.rawAnswers) // 11. Structured Evaluation Payload
      ];

      responseSheet.appendRow(rowData);
      let newRowNumber = responseSheet.getLastRow();
      
      // Expand raw answers into columns starting at column L (12)
      let currentCols = responseSheet.getLastColumn();
      if (record.rawAnswers && record.rawAnswers.length > 0) {
         let answerColStart = 12; 
         record.rawAnswers.forEach((ans, idx) => {
           responseSheet.getRange(newRowNumber, answerColStart + idx).setValue(
             typeof ans === 'object' ? JSON.stringify(ans) : ans
           );
         });
      }

      // ⚙️ Mock the Event Object to pass to the PDF Generator
      let mockEvent = {
        source: ss,
        range: responseSheet.getRange(newRowNumber, 1, 1, responseSheet.getLastColumn()),
        namedValues: {
          "Student Full Name": [record.studentName],
          "Mobile Number": [record.studentPhone],
          "Timestamp": [new Date(record.timestamp).toLocaleString()]
        }
      };
      
      // 🚀 Trigger PDF Generation & AI Grading locally (This uses the code from Part 5)
      try {
        autoGenerateReportCard(mockEvent);
      } catch(pdfErr) {
        responseSheet.getRange(newRowNumber, 9).setValue("❌ PDF Error: " + pdfErr.message);
      }
      
      successCount++;
    });

    dash.getRange("B6").setValue("System Ready!");
    ui.alert("✅ Sync Complete!", `Successfully downloaded and processed ${successCount} new student submission(s).\n\nPDF Scorecards are generating into your Google Drive folder.`, ui.ButtonSet.OK);

  } catch(error) {
    dash.getRange("B6").setValue("Sync Error");
    ui.alert("❌ Sync Error: " + error.message);
  }
}
/**
 * =========================================================================
 * SMART GEM CENTRAL LIBRARY (ENGINE) - PART 5: AI GRADER & PDF GENERATOR
 * =========================================================================
 */

// 🟢 HELPER: AI SUBJECTIVE GRADER
function gradeSubjectiveWithAI(questionText, studentAnswer, correctAnswer, maxMarks, apiKey) {
  if (!studentAnswer || studentAnswer.trim() === "" || studentAnswer === "[No Response]") {
    return { marks: 0, feedback: "No response provided by the student." };
  }
  
  const prompt = `You are a strict but fair academic evaluator. 
  Question: "${questionText}"
  Official Answer/Marking Scheme: "${correctAnswer}"
  Max Marks Possible: ${maxMarks}
  Student's Answer: "${studentAnswer}"
  
  Evaluate the student's answer based on the official scheme. Give partial credit if applicable.
  Respond ONLY in this exact JSON format, with no markdown formatting:
  {"awardedMarks": (number), "feedback": "(1-2 short sentences of constructive feedback)"}`;

  try {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: "application/json" } };
    const response = UrlFetchApp.fetch(endpoint, { method: "POST", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      let aiText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
      let result = JSON.parse(aiText.replace(/```json/g, "").replace(/```/g, "").trim());
      return { marks: parseFloat(result.awardedMarks) || 0, feedback: result.feedback };
    }
  } catch(e) {
    console.warn("AI Grading Error: " + e.message);
  }
  return { marks: 0, feedback: "Evaluation requires manual teacher review (AI Error)." };
}

// 🟢 MAIN CORE: PDF SCORECARD GENERATOR
function autoGenerateReportCard(e) {
  if (!e) return;
  
  const ss = e.source;
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  
  try {
    // 1. Core Data Extraction
    const responses = e.namedValues;
    const name = responses["Student Full Name"] ? responses["Student Full Name"][0] : "Student";
    const phone = responses["Mobile Number"] ? responses["Mobile Number"][0] : "";
    const timestamp = responses["Timestamp"] ? responses["Timestamp"][0] : new Date().toLocaleString();
    let dynamicExamName = sheet.getRange(row, 2).getValue() || "Assessment";
    const testId = sheet.getRange(row, 2).getValue(); // Assuming Test ID is logged here
    
    const dash = ss.getSheetByName("Dashboard");
    let folderId = extractIdFromInput(dash.getRange("B22").getValue());
    let uploadFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    
    // 2. Build Master Answer Key Dictionary (filtered by Test ID)
    let masterAnswerKey = {};
    const keySheet = ss.getSheetByName("System Answer Keys");
    if (keySheet) {
      let data = keySheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        // Match the Test ID to ensure we grab the right answers
        if (String(data[i][0]).trim() === String(testId).trim()) {
          let rawLookupKey = String(data[i][3]);
          masterAnswerKey[rawLookupKey] = {
            correctAnswer: String(data[i][4]).trim(),
            explanation: String(data[i][5]).trim(),
            qType: String(data[i][6] || "Theory").trim(),
            qPoints: parseFloat(data[i][7]) || 1,
            topic: rawLookupKey.match(/\[(.*?)\]/) ? rawLookupKey.match(/\[(.*?)\]/)[1] : "General"
          };
        }
      }
    }

    // 3. Initialize Grading Variables
    let totalEarnedMarks = 0, totalMaxMarks = 0;
    let evalLog = [];
    let conceptScores = {}; // Tracks marks per topic for the Chart
    
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let studentRowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const logisticsHeaders = ["Timestamp", "Exam Name", "Student Full Name", "Mobile Number", "Email Address", "Score", "📄 PDF Report", "💬 WhatsApp", "Evaluation Status", "📧 Delivery Status", "Structured Evaluation Payload"];

    // 4. Evaluation Loop (Start from column 12 - index 11)
    let qCounter = 1;
    for (let i = 11; i < headers.length; i++) {
      let headerName = headers[i];
      if (!headerName || logisticsHeaders.includes(headerName)) continue;
      
      let studentAnswer = studentRowValues[i] || "[No Response]";
      let keyData = masterAnswerKey[headerName];
      if (!keyData) continue; // Skip if question not found in key
      
      let qMaxMarks = keyData.qPoints;
      let earnedForQ = 0;
      let feedback = "";
      let isObjective = (keyData.qType.includes("MCQ") || keyData.qType.includes("Assertion"));
      
      // Initialize Topic tracking
      if (!conceptScores[keyData.topic]) conceptScores[keyData.topic] = { earned: 0, total: 0 };
      conceptScores[keyData.topic].total += qMaxMarks;

      if (isObjective) {
        // Objective String Matching
        let cleanStudent = String(studentAnswer).toLowerCase().replace(/[^a-z0-9]/g, "").trim();
        let cleanMaster = String(keyData.correctAnswer).toLowerCase().replace(/[^a-z0-9]/g, "").trim();
        
        if (cleanStudent === cleanMaster || cleanStudent.includes(cleanMaster) || cleanMaster.includes(cleanStudent)) {
          earnedForQ = qMaxMarks;
          feedback = "Correct.";
        } else {
          feedback = `Incorrect. Correct Answer: ${keyData.correctAnswer}`;
        }
      } else {
        // Subjective AI Grading
        let aiEvaluation = gradeSubjectiveWithAI(headerName, studentAnswer, keyData.correctAnswer, qMaxMarks, apiKey);
        earnedForQ = aiEvaluation.marks;
        feedback = aiEvaluation.feedback;
      }
      
      // Safety cap marks
      if (earnedForQ > qMaxMarks) earnedForQ = qMaxMarks;
      
      totalEarnedMarks += earnedForQ;
      totalMaxMarks += qMaxMarks;
      conceptScores[keyData.topic].earned += earnedForQ;
      
      evalLog.push({
        question: `Q${qCounter}: ${headerName}`,
        studentAns: `Candidate's Response: ${studentAnswer}`,
        marksStr: `[Marks Awarded: ${earnedForQ} / ${qMaxMarks}]`,
        feedback: `Feedback: ${feedback}`
      });
      qCounter++;
    }

    // 5. Generate Concept Mastery Google Chart
    let chartBlob = null;
    try {
      var dataBuilder = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, 'Concept')
        .addColumn(Charts.ColumnType.NUMBER, 'Score (%)');
      
      let needsRemediation = [];
      for (let topic in conceptScores) {
        let percentage = (conceptScores[topic].total > 0) ? Math.round((conceptScores[topic].earned / conceptScores[topic].total) * 100) : 0;
        dataBuilder.addRow([topic.substring(0, 20) + (topic.length > 20 ? "..." : ""), percentage]);
        if (percentage < 60) needsRemediation.push(topic);
      }
      
      var chart = Charts.newBarChart()
        .setDataTable(dataBuilder.build())
        .setDimensions(550, 300)
        .setColors(['#1e3a8a'])
        .setTitle('Concept Mastery Breakdown (%)')
        .setOption('vAxis', {textStyle: {fontSize: 10}})
        .setOption('hAxis', {viewWindow: {min: 0, max: 100}})
        .build();
      chartBlob = chart.getAs('image/png');
    } catch (chartErr) {
      console.warn("Chart generation failed: " + chartErr.message);
    }

    // 6. Build the PDF Document
    const docName = `[Scorecard] ${name} - ${dynamicExamName}`;
    const doc = DocumentApp.create(docName);
    const docBody = doc.getBody();
    
    // Formatting
    doc.setMarginLeft(40); doc.setMarginRight(40); doc.setMarginTop(40); doc.setMarginBottom(40);
    
    // Header
    let brandName = dash.getRange("B9").getValue() || "Smart Gem Assessment System";
    docBody.appendParagraph(brandName).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    docBody.appendParagraph("DIAGNOSTIC SCORECARD & REMEDIATION PLAN").setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    docBody.appendHorizontalRule();

    // Meta Data Table
    let metaTable = docBody.appendTable([
      ["Candidate Name:", name, "Overall Score:", `${totalEarnedMarks} / ${totalMaxMarks}`],
      ["Mobile / ID:", phone, "Percentage:", `${Math.round((totalEarnedMarks/totalMaxMarks)*100 || 0)}%`],
      ["Test Scope:", dynamicExamName, "Evaluation Date:", timestamp]
    ]);
    metaTable.setBorderWidth(0); 
    
    // Chart Insertion
    if (chartBlob) {
      docBody.appendParagraph("").setSpacingBefore(12);
      let p = docBody.appendParagraph("");
      p.appendImage(chartBlob).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    // AI Remediation Plan
    docBody.appendParagraph("Targeted Remediation Plan").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    let weakTopics = [];
    for (let topic in conceptScores) {
      if ((conceptScores[topic].earned / conceptScores[topic].total) < 0.6) weakTopics.push(topic);
    }
    
    if (weakTopics.length > 0) {
      docBody.appendParagraph(`Based on the assessment, the candidate requires focused revision in the following areas:`).setItalic(true);
      weakTopics.forEach(t => docBody.appendListItem(t).setGlyphType(DocumentApp.GlyphType.BULLET));
    } else {
      docBody.appendParagraph("Excellent performance! The candidate has demonstrated strong mastery across all tested concepts.");
    }
    docBody.appendHorizontalRule();

    // Detailed Log
    docBody.appendParagraph("Detailed Question Analysis").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    evalLog.forEach(log => {
      docBody.appendParagraph(log.question).setBold(true);
      docBody.appendParagraph(log.studentAns).setBold(false);
      docBody.appendParagraph(log.marksStr).setForegroundColor("#0b5394").setBold(true);
      docBody.appendParagraph(log.feedback).setItalic(true);
      docBody.appendParagraph(""); // spacer
    });

    // 7. Save Document to PDF
    doc.saveAndClose(); 
    Utilities.sleep(3000); // Wait for Google to process the doc
    const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
    const finalPdf = uploadFolder.createFile(pdfBlob).setName(`${docName}.pdf`);
    const pdfUrl = finalPdf.getUrl();
    
    // Trash the temporary Google Doc
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    // 8. Write Results Back to the Master Sheet
    let pdfCol = headers.indexOf("📄 PDF Report") + 1;
    let scoreCol = headers.indexOf("Score") + 1;
    let waCol = headers.indexOf("💬 WhatsApp") + 1;
    let evalCol = headers.indexOf("Evaluation Status") + 1;

    let dynamicScoreStr = `${totalEarnedMarks} / ${totalMaxMarks}`;
    if (scoreCol > 0) sheet.getRange(row, scoreCol).setValue(dynamicScoreStr);
    if (pdfCol > 0) sheet.getRange(row, pdfCol).setFormula(`=HYPERLINK("${pdfUrl}", "📥 View PDF")`);
    
    // Format WhatsApp Link
    if (waCol > 0) {
      let waText = `Hello ${name},\nYour assessment results for '${dynamicExamName}' are ready.\n\nScore: ${dynamicScoreStr}\n\nDownload your detailed PDF Scorecard here:\n${pdfUrl}`;
      sheet.getRange(row, waCol).setFormula(`=HYPERLINK("https://api.whatsapp.com/send?phone=91${phone}&text=${encodeURIComponent(waText)}", "📱 Send WhatsApp")`);
    }
    
    if (evalCol > 0) sheet.getRange(row, evalCol).setValue("✅ Done");

  } catch (err) { 
    sheet.getRange(row, 9).setValue("❌ ERROR: " + err.message); // Log error in Evaluation Status
  }
}
