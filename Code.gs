const SHEET_NAME = PropertiesService.getScriptProperties().getProperty("WATTBIKE_SHEET") || "Wattbike Sessions";
const SENDER_EMAIL = "hub@wattbike.com";
const SEARCH_SUBJECT = "session summary";
const HOURS_LOOKBACK = 30;
const API_TOKEN = "RideHard2026!"; // Change this to your chosen password

function logWattbikeSession() {
  const sheet = getOrCreateSheet();
  const cutoffTime = new Date(Date.now() - HOURS_LOOKBACK * 60 * 60 * 1000);
  const query = `from:(${SENDER_EMAIL}) subject:(${SEARCH_SUBJECT}) after:${formatDateForQuery(cutoffTime)}`;
  const threads = GmailApp.search(query);
  if (threads.length === 0) return;

  let sessionsLogged = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      if (isAlreadyProcessed(message)) continue;
      
      const body = message.getBody(); 
      const data = parseSessionData(body, message);
      
      if (data) {
          appendToSheet(sheet, data);
          sessionsLogged++;
          markAsProcessed(message);
          
          CacheService.getScriptCache().remove("analyticsData");
      } 
    }
  }
}

function parseSessionData(html, message) {
  const text = html
    .replace(/<style[\s\S]*?<\/style>/gi, " ")  
    .replace(/<script[\s\S]*?<\/script>/gi, " ") 
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/td>/gi, "\t")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&#160;/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();

  const steps = {};
  try { steps.date     = extractDate(text); }     catch(e) { steps.dateErr     = e.message; }
  try { steps.duration = extractDuration(text); } catch(e) { steps.durationErr = e.message; }
  try { steps.metrics  = parseAveragesAndPeaks(text); } catch(e) { steps.metricsErr = e.message; }

  const failed = steps.dateErr || steps.durationErr || steps.metricsErr;
  
  if (failed) {
    const msgDate = message ? message.getDate().toString() : "unknown";
    logDebugError(msgDate, steps, text);
    return null;
  }

  const workDoneKJ = calculateWorkDone(steps.duration, steps.metrics.avg.power);

  return {
    date:          steps.date,
    duration:      steps.duration,
    avgPower:      steps.metrics.avg.power,
    peakPower:     steps.metrics.peak.power,
    avgHeartRate:  steps.metrics.avg.heartRate,
    peakHeartRate: steps.metrics.peak.heartRate,
    avgCadence:    steps.metrics.avg.cadence,
    peakCadence:   steps.metrics.peak.cadence,
    workDoneKJ:    workDoneKJ
  };
}

function calculateWorkDone(durationStr, avgPower) {
  if (!durationStr || !avgPower) return null;
  const timeParts = durationStr.split(':');
  const durationSeconds = (parseInt(timeParts[0], 10) * 3600) + 
                          (parseInt(timeParts[1], 10) * 60) + 
                          parseInt(timeParts[2], 10);
  return Math.round((avgPower * durationSeconds) / 1000); 
}

function extractDate(text) {
  const match = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
  if (!match) throw new Error("Date not found");
  const parts = match[1].split("/");
  return `${parts[0].padStart(2,"0")}/${parts[1].padStart(2,"0")}/${parts[2]}`;
}

function extractDuration(text) {
  const hoursMatch = text.match(/(\d+):(\d+):(\d+)\s*h\b/i);
  if (hoursMatch) return `${hoursMatch[1]}:${hoursMatch[2].padStart(2, "0")}:${hoursMatch[3].padStart(2, "0")}`;

  const minsMatch = text.match(/(\d+):(\d+)\s*min\b/i);
  if (minsMatch) return `0:${minsMatch[1].padStart(2, "0")}:${minsMatch[2].padStart(2, "0")}`;

  const secsMatch = text.match(/(\d+)\s*sec\b/i);
  if (secsMatch) return `0:00:${secsMatch[1].padStart(2, "0")}`;

  throw new Error("Duration not found");
}

function parseAveragesAndPeaks(text) {
  const sectionMatch = text.match(/Averages and Peaks([\s\S]*?)(?:PES|Leg Balance|$)/i);
  if (!sectionMatch) throw new Error("Averages and Peaks section not found");
  const section = sectionMatch[1];
  
  const rowPattern = /(\d+)\s*w\s+(--|\d+)\s*bpm\s+(\d+)\s*rpm/gi;
  const rows = [];
  let m;
  
  while ((m = rowPattern.exec(section)) !== null) {
    rows.push({
      power:     parseInt(m[1], 10),
      heartRate: m[2] === "--" ? "" : parseInt(m[2], 10), 
      cadence:   parseInt(m[3], 10)
    });
  }

  if (rows.length < 2) throw new Error("Could not find both avg and peak rows");
  return { avg: rows[0], peak: rows[1] };
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ["Date", "Duration", "Avg Power (w)", "Peak Power (w)", "Avg Heart Rate (bpm)", "Peak Heart Rate (bpm)", "Avg Cadence (rpm)", "Peak Cadence (rpm)", "Work Done (kJ)"];
    sheet.appendRow(headers);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 100); 
    for (let i = 3; i <= headers.length; i++) sheet.setColumnWidth(i, 160);
  }
  return sheet;
}

function appendToSheet(sheet, data) {
  sheet.appendRow([
    data.date,
    data.duration,
    data.avgPower,
    data.peakPower,
    data.avgHeartRate,
    data.peakHeartRate,
    data.avgCadence,
    data.peakCadence,
    data.workDoneKJ
  ]);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 9).setHorizontalAlignment("center");
}

function logDebugError(msgDate, steps, text) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debugSheet = ss.getSheetByName("Debug");
  if (!debugSheet) {
    debugSheet = ss.insertSheet("Debug");
    debugSheet.appendRow(["Timestamp", "Message Date", "Errors", "Raw Text Snapshot"]);
  }
  const errorObj = {
    date: steps.dateErr || "OK",
    duration: steps.durationErr || "OK",
    metrics: steps.metricsErr || "OK"
  };
  debugSheet.appendRow([new Date(), msgDate, JSON.stringify(errorObj), text.substring(0, 800)]);
}

const PROCESSED_SHEET_NAME = "WattbikeProcessed";

function getOrCreateProcessedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PROCESSED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PROCESSED_SHEET_NAME);
    sheet.hideSheet();
  }
  return sheet;
}

function markAsProcessed(message) {
  getOrCreateProcessedSheet().appendRow([message.getId(), new Date().toISOString()]);
}

function isAlreadyProcessed(message) {
  const sheet = getOrCreateProcessedSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return false;
  return sheet.getRange(1, 1, lastRow, 1).getValues().flat().includes(message.getId());
}

function formatDateForQuery(date) {
  return Utilities.formatDate(date, "Europe/London", "yyyy/MM/dd");
}

// --- NEW OFFLINE API FUNCTIONS ---

function doGet(e) {
  if (!e.parameter || e.parameter.token !== API_TOKEN) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Access Denied" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  logWattbikeSession(); 

  const payload = {
    today: getTodayWorkout(),
    upcoming: getUpcomingWorkouts(),
    analytics: getAnalyticsData()
  };

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e.parameter || e.parameter.token !== API_TOKEN) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Access Denied" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const queue = JSON.parse(e.postData.contents);
    for (const item of queue) {
      saveWorkout(item.rowNum, item.goal, item.duration, item.intensity, item.hrCap, item.notes);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
