const PLAN_SHEET_NAME = "Training Plan";
const WATTBIKE_SHEET = PropertiesService.getScriptProperties().getProperty("WATTBIKE_SHEET") || "Wattbike Sessions";

const COL = {
  DATE:      1,  
  DAY:       2,  
  GOAL:      3,  
  DURATION:  4,  
  INTENSITY: 5,  
  HR_CAP:    6,  
  AVG_WATTS: 7,  
  AVG_HR:    8,  
  AVG_RPM:   9,  
  WORK_KJ:   10,
  NOTES:     11,
  COMPLETED: 12 
};

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Today's Workout")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTodayWorkout() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PLAN_SHEET_NAME);
  if (!sheet) return { error: "Sheet '" + PLAN_SHEET_NAME + "' not found." };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: "No data in Training Plan sheet." };

  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const cellDate = new Date(data[i][COL.DATE - 1]);
    cellDate.setHours(0, 0, 0, 0);

    if (cellDate.getTime() === today.getTime()) {
      return {
        rowNum:    i + 2,
        date:      formatDate(today),
        dayOfWeek: getDayName(today),
        goal:      data[i][COL.GOAL - 1]      || "",
        duration:  data[i][COL.DURATION - 1]  || "",
        intensity: data[i][COL.INTENSITY - 1] || "",
        hrCap:     data[i][COL.HR_CAP - 1]    || "",
        avgWatts:  data[i][COL.AVG_WATTS - 1] || "",
        avgHR:     data[i][COL.AVG_HR - 1]    || "",
        avgRPM:    data[i][COL.AVG_RPM - 1]   || "",
        workKJ:    data[i][COL.WORK_KJ - 1]   || "",
        notes:     data[i][COL.NOTES - 1]     || ""
      };
    }
  }

  return { error: "No workout planned for today (" + formatDate(today) + ")." };
}

function saveWorkout(rowNum, goal, duration, intensity, hrCap, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PLAN_SHEET_NAME);
  if (!sheet) return { success: false, message: "Sheet not found." };

  try {
    sheet.getRange(rowNum, COL.GOAL).setValue(goal);
    sheet.getRange(rowNum, COL.DURATION).setValue(duration);
    sheet.getRange(rowNum, COL.INTENSITY).setValue(intensity);
    sheet.getRange(rowNum, COL.HR_CAP).setValue(hrCap);
    sheet.getRange(rowNum, COL.NOTES).setValue(notes);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function formatDate(date) {
  return Utilities.formatDate(date, "Europe/London", "dd/MM/yyyy");
}

function getDayName(date) {
  return ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][date.getDay()];
}

function getUpcomingWorkouts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PLAN_SHEET_NAME);
  if (!sheet) return { error: "Sheet '" + PLAN_SHEET_NAME + "' not found." };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { days: [] };
  
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  const upcoming = [];
  
  for (let i = 0; i < data.length && upcoming.length < 5; i++) {
    const cellDate = new Date(data[i][COL.DATE - 1]);
    cellDate.setHours(0, 0, 0, 0);

    const diffDays = Math.round((cellDate - today) / (1000 * 60 * 60 * 24));
    if (diffDays < 1 || diffDays > 10) continue; 

    upcoming.push({
      date:      formatDate(cellDate),
      dayOfWeek: getDayName(cellDate),
      daysAway:  diffDays,
      goal:      data[i][COL.GOAL - 1]      || "",
      duration:  data[i][COL.DURATION - 1]  || "",
      intensity: data[i][COL.INTENSITY - 1] || "",
      hrCap:     data[i][COL.HR_CAP - 1]    || "",
      avgWatts:  data[i][COL.AVG_WATTS - 1] || ""
    });
  }

  return { days: upcoming };
}

function getAnalyticsData() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("analyticsData");
  if (cachedData) return JSON.parse(cachedData);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(WATTBIKE_SHEET);

  if (!sheet) return { error: "Wattbike Sessions sheet not found." };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { sessions: [], sheetName: sheet.getName() };

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const sessions = [];

  for (const row of data) {
    const rawDate = row[0];
    if (!rawDate) continue;
    
    let date;
    if (rawDate instanceof Date) {
      date = rawDate;
    } else {
      const parts = String(rawDate).split('/');
      if (parts.length === 3) date = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    }
    
    if (!date || isNaN(date.getTime())) continue;
    
    sessions.push({
      date:         formatDate(date),
      dateMs:       date.getTime(),
      durationMins: parseDurationToMins(row[1]),
      avgWatts:     Number(row[2]) || null,
      peakWatts:    Number(row[3]) || null,
      avgHR:        Number(row[4]) || null,
      peakHR:       Number(row[5]) || null,
      avgRPM:       Number(row[6]) || null,
      workKJ:       Number(row[8]) || null
    });
  }

  sessions.sort((a, b) => a.dateMs - b.dateMs);
  
  const payload = { sessions: sessions };
  cache.put("analyticsData", JSON.stringify(payload), 14400); 
  return payload;
}

function parseDurationToMins(val) {
  if (!val) return 0;
  const str = String(val).trim();
  const parts = str.split(':');
  if (parts.length === 3) return parseInt(parts[0]) * 60 + parseInt(parts[1]) + parseInt(parts[2]) / 60;
  else if (parts.length === 2) return parseInt(parts[0]) + parseInt(parts[1]) / 60;
  return 0;
}
