/**
 * Computer Science & IT Educational WebAR - Backend API Engine
 * School: Chonprathan Phataek School
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ---------------- API ROUTERS & RESPONSE HELPER ----------------
function buildJSONResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'initDb') {
      return buildJSONResponse(initDatabase());
    } else if (action === 'getLeaderboard') {
      return buildJSONResponse({ status: 'success', data: getLeaderboardData() });
    } else if (action === 'getDashboard') {
      return buildJSONResponse({ status: 'success', data: getDashboardMetrics() });
    } else if (action === 'getQuestions') {
      return buildJSONResponse({ status: 'success', data: getQuestionsData() }); // ดึงข้อมูลโจทย์
    }

    return buildJSONResponse({ status: 'error', message: 'Invalid GET action' });
  } catch (err) {
    return buildJSONResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'login') {
      const user = handleUserLogin(body.userData);
      return buildJSONResponse({ status: 'success', data: user });
    } else if (action === 'addExp') {
      const result = addPlayerEXP(body.email, body.expGained, body.activityName);
      return buildJSONResponse({ status: 'success', data: result });
    }

    return buildJSONResponse({ status: 'error', message: 'Invalid POST action' });
  } catch (err) {
    return buildJSONResponse({ status: 'error', message: err.toString() });
  }
}

// ---------------- DATABASE INITIALIZER ----------------
function initDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = [
    { name: 'Users', headers: ['UID', 'Email', 'Name', 'Class', 'Room', 'EXP', 'Level', 'Badges', 'CreatedAt'] },
    { name: 'Scores', headers: ['ID', 'Email', 'Activity', 'Score', 'Timestamp'] },
    { name: 'Quiz', headers: ['ID', 'ModelID', 'Question', 'OptA', 'OptB', 'OptC', 'OptD', 'CorrectOpt', 'Explain'] },
    { name: 'Missions', headers: ['ID', 'Title', 'ReqType', 'ReqCount', 'RewardEXP', 'BadgeID'] },
    { name: 'Logs', headers: ['ID', 'Email', 'Action', 'Timestamp'] },
    { name: 'Badges', headers: ['ID', 'Name', 'Icon', 'Description'] },
    { name: 'DATA', headers: ['โจทย์', 'คำตอบถูก', 'คำตอบผิด', 'คะแนน', 'เวลา(วินาที)', 'ความเร็ว(ช้า/ปานกลาง/เร็ว)'] } // เพิ่มชีต DATA
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold").setBackground("#1e3c72").setFontColor("#ffffff");
    }
  });
  return { status: "success", message: "Database initialized successfully" };
}

// ---------------- USER & AUTH MANAGEMENT ----------------
function handleUserLogin(userData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  const email = userData.email;
  let userRow = -1;
  let user = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      userRow = i + 1;
      user = {
        uid: data[i][0], email: data[i][1], name: data[i][2],
        class: data[i][3], room: data[i][4], exp: data[i][5],
        level: data[i][6], badges: data[i][7] ? JSON.parse(data[i][7]) : []
      };
      break;
    }
  }

  if (userRow === -1) {
    const uid = 'USR-' + Date.now();
    user = {
      uid: uid, email: email, name: userData.name,
      class: userData.class || 'M.1', room: userData.room || '1',
      exp: 0, level: 'LV1 Beginner', badges: []
    };
    sheet.appendRow([user.uid, user.email, user.name, user.class, user.room, 0, 'LV1 Beginner', '[]', new Date()]);
  }

  logActivity(email, 'LOGIN');
  return sanitizeData(user);
}

// ---------------- GAME PROGRESS & EXP SYSTEM ----------------
function addPlayerEXP(email, expGained, activityName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      let currentExp = Number(data[i][5]) + Number(expGained);
      let newLevel = calculateLevel(currentExp);
      
      sheet.getRange(i + 1, 6).setValue(currentExp);
      sheet.getRange(i + 1, 7).setValue(newLevel);
      
      // Record Score Log
      ss.getSheetByName('Scores').appendRow(['SCR-' + Date.now(), email, activityName, expGained, new Date()]);
      
      return { exp: currentExp, level: newLevel };
    }
  }
  return null;
}

function calculateLevel(exp) {
  if (exp >= 1000) return 'LV5 Master';
  if (exp >= 600)  return 'LV4 Expert';
  if (exp >= 300)  return 'LV3 Advanced';
  if (exp >= 100)  return 'LV2 Student';
  return 'LV1 Beginner';
}

// ---------------- DATA FETCHING SERVICES ----------------
function getLeaderboardData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove header

  return data
    .map(r => ({ name: r[2], class: `ม.${r[3]}/${r[4]}`, exp: r[5], level: r[6], badges: r[7] ? JSON.parse(r[7]).length : 0 }))
    .sort((a, b) => b.exp - a.exp)
    .slice(0, 100);
}

function getDashboardMetrics() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const users = ss.getSheetByName('Users').getDataRange().getValues();
  const scores = ss.getSheetByName('Scores').getDataRange().getValues();
  const logs = ss.getSheetByName('Logs').getDataRange().getValues();

  users.shift(); scores.shift(); logs.shift();

  return {
    totalStudents: users.length,
    totalPlays: scores.length,
    recentLogs: logs.slice(-20).reverse(),
    levelDistribution: users.reduce((acc, curr) => {
      acc[curr[6]] = (acc[curr[6]] || 0) + 1;
      return acc;
    }, {})
  };
}

// ระบบดึงข้อมูลโจทย์จากชีต 'DATA'
function getQuestionsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DATA');
  
  if (!sheet) return []; // คืนค่าว่างถ้ายังไม่มีชีต
  
  const data = sheet.getDataRange().getValues();
  let questions = [];
  
  // ลูปข้ามแถวแรก (แถวหัวข้อ)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== "") { // ตรวจสอบว่ามีโจทย์ ไม่ใช่บรรทัดว่าง
      questions.push({
        question: data[i][0],              // คอลัมน์ A (โจทย์)
        correct: data[i][1],               // คอลัมน์ B (คำตอบถูก)
        wrong: data[i][2],                 // คอลัมน์ C (คำตอบผิด)
        score: data[i][3] || 10,           // คอลัมน์ D (คะแนน) - ค่าเร่ิมต้น 10
        time: data[i][4] || 30,            // คอลัมน์ E (เวลาวินาที) - ค่าเริ่มต้น 30
        speed: data[i][5] || "ปานกลาง"      // คอลัมน์ F (ความเร็ว) - ค่าเริ่มต้น ปานกลาง
      });
    }
  }
  
  return questions;
}

// ---------------- UTILS & SECURITY ----------------
function logActivity(email, action) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getSheetByName('Logs').appendRow(['LOG-' + Date.now(), email, action, new Date()]);
}

function sanitizeData(input) {
  return JSON.parse(JSON.stringify(input));
}
