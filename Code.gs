/**
 * School Slot Booking System - Google Apps Script Backend
 * 
 * Setup:
 * 1. Create a Google Sheet with sheets named: "Teachers", "Config", "Bookings"
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code and save
 * 4. Deploy > New deployment > Web app
 * 5. Execute as: Me, Who has access: Anyone
 * 6. Copy the deployment URL to your frontend
 */

// Sheet names
const TEACHERS_SHEET = 'Teachers';
const CONFIG_SHEET = 'Config';
const BOOKINGS_SHEET = 'Bookings';
const SESSIONS_SHEET = 'Sessions';

/**
 * Handle GET requests - fetch data (also handles POST fallback)
 */
function doGet(e) {
  const action = e.parameter.action;
  let result;
  
  try {
    // Check if this is a POST fallback (has 'data' parameter)
    if (e.parameter.data) {
      const data = JSON.parse(e.parameter.data);
      result = handlePostAction(data);
    } else {
      switch(action) {
        case 'getTeachers':
          result = getTeachers();
          break;
        case 'getConfig':
          result = getConfig();
          break;
        case 'getBookings':
          result = getBookings();
          break;
        case 'getSessions':
          result = getSessions();
          break;
        case 'getAll':
          result = {
            teachers: getTeachers(),
            config: getConfig(),
            bookings: getBookings(),
            sessions: getSessions()
          };
          break;
        default:
          result = { error: 'Unknown action' };
      }
    }
  } catch (error) {
    result = { error: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle POST action (shared between doPost and doGet fallback)
 */
function handlePostAction(data) {
  const action = data.action;
  
  switch(action) {
    case 'addTeacher':
      return addTeacher(data.name, data.availableFrom, data.availableUntil, data.slotsBeforeBreak, data.breakDuration);
    case 'updateTeacher':
      return updateTeacher(data.row, data.name, data.availableFrom, data.availableUntil, data.slotsBeforeBreak, data.breakDuration);
    case 'deleteTeacher':
      return deleteTeacher(data.row);
    case 'updateConfig':
      return updateConfig(data.slotDuration, data.slotsBeforeBreak, data.breakDuration, data.feedbackDate, data.lunchStart, data.lunchEnd);
    case 'bookSlot':
      return bookSlot(data.teacher, data.slotTime, data.studentName, data.studentEmail);
    case 'cancelBooking':
      return cancelBooking(data.row);
    case 'addSession':
      return addSession(data.name, data.date, data.sheetUrl);
    case 'updateSession':
      return updateSession(data.row, data.name, data.date, data.sheetUrl);
    case 'deleteSession':
      return deleteSession(data.row);
    default:
      return { error: 'Unknown action' };
  }
}

/**
 * Handle POST requests - create/update data
 */
function doPost(e) {
  let result;
  
  try {
    const data = JSON.parse(e.postData.contents);
    result = handlePostAction(data);
  } catch (error) {
    result = { error: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Get all teachers
 */
function getTeachers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHERS_SHEET);
  const data = sheet.getDataRange().getValues();
  const teachers = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Only if name exists
      teachers.push({
        row: i + 1,
        name: data[i][0],
        availableFrom: data[i][1],
        availableUntil: data[i][2],
        slotsBeforeBreak: data[i][3] || '',
        breakDuration: data[i][4] !== undefined ? data[i][4] : ''
      });
    }
  }
  
  return teachers;
}

/**
 * Add a new teacher
 */
function addTeacher(name, availableFrom, availableUntil, slotsBeforeBreak, breakDuration) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHERS_SHEET);
  sheet.appendRow([name, availableFrom, availableUntil, slotsBeforeBreak || '', breakDuration !== undefined ? breakDuration : '']);
  return { success: true };
}

/**
 * Update an existing teacher
 */
function updateTeacher(row, name, availableFrom, availableUntil, slotsBeforeBreak, breakDuration) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHERS_SHEET);
  sheet.getRange(row, 1, 1, 5).setValues([[name, availableFrom, availableUntil, slotsBeforeBreak || '', breakDuration !== undefined ? breakDuration : '']]);
  return { success: true };
}

/**
 * Delete a teacher
 */
function deleteTeacher(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHERS_SHEET);
  sheet.deleteRow(row);
  return { success: true };
}

/**
 * Get configuration
 */
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = sheet.getDataRange().getValues();
  
  return {
    slotDuration: data[0] ? data[0][1] || 15 : 15,
    slotsBeforeBreak: data[1] ? data[1][1] || 4 : 4,
    breakDuration: data[2] ? data[2][1] || 15 : 15,
    feedbackDate: data[3] ? data[3][1] || '' : '',
    lunchStart: data[4] ? data[4][1] || '' : '',
    lunchEnd: data[5] ? data[5][1] || '' : ''
  };
}

/**
 * Update configuration
 */
function updateConfig(slotDuration, slotsBeforeBreak, breakDuration, feedbackDate, lunchStart, lunchEnd) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  sheet.getRange('A1:B6').setValues([
    ['Slot Duration (min)', slotDuration],
    ['Slots Before Break', slotsBeforeBreak],
    ['Break Duration (min)', breakDuration],
    ['Feedback Date', feedbackDate || ''],
    ['Lunch Start', lunchStart || ''],
    ['Lunch End', lunchEnd || '']
  ]);
  return { success: true };
}

/**
 * Get all bookings
 */
function getBookings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET);
  const data = sheet.getDataRange().getValues();
  const bookings = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Only if teacher exists
      bookings.push({
        row: i + 1,
        teacher: data[i][0],
        slotTime: data[i][1],
        studentName: data[i][2],
        studentEmail: data[i][3],
        timestamp: data[i][4]
      });
    }
  }
  
  return bookings;
}

/**
 * Book a slot
 */
function bookSlot(teacher, slotTime, studentName, studentEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET);
  
  // Check if slot is already booked
  const bookings = getBookings();
  const existingBooking = bookings.find(b => b.teacher === teacher && b.slotTime === slotTime);
  
  if (existingBooking) {
    return { success: false, error: 'Slot already booked' };
  }
  
  // Add the booking
  sheet.appendRow([teacher, slotTime, studentName, studentEmail, new Date().toISOString()]);
  return { success: true };
}

/**
 * Cancel a booking
 */
function cancelBooking(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET);
  sheet.deleteRow(row);
  return { success: true };
}

/**
 * Get all sessions
 */
function getSessions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SESSIONS_SHEET);
  if (!sheet) {
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const sessions = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      sessions.push({
        row: i + 1,
        name: data[i][0],
        date: data[i][1],
        sheetUrl: data[i][2] || ''
      });
    }
  }
  
  return sessions;
}

/**
 * Add a new session
 */
function addSession(name, date, sheetUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SESSIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SESSIONS_SHEET);
    sheet.getRange('A1:C1').setValues([['Session Name', 'Date', 'Sheet URL']]);
  }
  sheet.appendRow([name, date, sheetUrl || '']);
  return { success: true };
}

/**
 * Update a session
 */
function updateSession(row, name, date, sheetUrl) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
  sheet.getRange(row, 1, 1, 3).setValues([[name, date, sheetUrl || '']]);
  return { success: true };
}

/**
 * Delete a session
 */
function deleteSession(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
  sheet.deleteRow(row);
  return { success: true };
}

/**
 * Initialize sheets with headers (run once manually)
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sessions sheet (new)
  let sheet = ss.getSheetByName(SESSIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SESSIONS_SHEET);
  }
  sheet.getRange('A1:C1').setValues([['Session Name', 'Date', 'Sheet URL']]);
  
  // Teachers sheet
  sheet = ss.getSheetByName(TEACHERS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TEACHERS_SHEET);
  }
  sheet.getRange('A1:E1').setValues([['Teacher Name', 'Available From', 'Available Until', 'Slots Before Break', 'Break Duration']]);
  
  // Config sheet
  sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET);
  }
  sheet.getRange('A1:B6').setValues([
    ['Slot Duration (min)', 15],
    ['Slots Before Break', 4],
    ['Break Duration (min)', 15],
    ['Feedback Date', ''],
    ['Lunch Start', '12:00'],
    ['Lunch End', '13:00']
  ]);
  
  // Bookings sheet
  sheet = ss.getSheetByName(BOOKINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BOOKINGS_SHEET);
  }
  sheet.getRange('A1:E1').setValues([['Teacher', 'Slot Time', 'Student Name', 'Student Email', 'Timestamp']]);
}
