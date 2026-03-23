const SPREADSHEET_ID = "1tYlXDTrIoMYajLDJZa_7apaEpvVzB0rRld7K6Ymbhck";

/**
 * Handle GET requests from the web app
 * /exec?action=getData
 */
function doGet(e) {
  var action = e.parameter.action;
  
  if (action === 'getData') {
    var data = {
      plants: getPlants(),
      bookings: getBookings()
    };
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'saveBooking') {
    var payloadStr = e.parameter.payload;
    if (payloadStr) {
      try {
        var payload = JSON.parse(payloadStr);
        return ContentService.createTextOutput(JSON.stringify(saveBooking(payload)))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  } else if (action === 'deleteBooking') {
    var payloadStr = e.parameter.payload;
    if (payloadStr) {
      try {
        var payload = JSON.parse(payloadStr);
        return ContentService.createTextOutput(JSON.stringify(deleteBooking(payload)))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  } else if (action === 'logAction') {
    var payloadStr = e.parameter.payload;
    if (payloadStr) {
      try {
        var payload = JSON.parse(payloadStr);
        return ContentService.createTextOutput(JSON.stringify(saveLog(payload)))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  
  // Default response if no action matched
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "API is active. Supported actions: getData, saveBooking, deleteBooking" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle POST requests (Save, Edit, Delete) (Retained for backwards compatibility)
 */
function doPost(e) {
  var request;
  try {
    request = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid JSON format" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = request.action;
  var payload = request.payload;

  var result = { status: "success" };

  try {
    if (action === 'saveBooking') {
      result = saveBooking(payload);
    } else if (action === 'deleteBooking') {
      result = deleteBooking(payload);
    } else {
      result = { status: "error", message: "Unknown action" };
    }
  } catch (err) {
    result = { status: "error", message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------
// HELPER FUNCTIONS
// ----------------------------------------------------------------------

function getPlants() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("PlantTranfer");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only headers
  
  var headers = data[0];
  var idPlantIdx = headers.indexOf("IDPlant");
  var plantIdx = headers.indexOf("Plant");
  var roundIdx = headers.indexOf("Round");
  var dateStartIdx = headers.indexOf("DateStart");
  var dateBeginIdx = headers.indexOf("Datebegin");
  
  // Fallback indices if header names differ slightly
  if (idPlantIdx === -1) idPlantIdx = 0;
  if (plantIdx === -1) plantIdx = 1;
  if (roundIdx === -1) roundIdx = 2;
  if (dateStartIdx === -1) dateStartIdx = 3;
  if (dateBeginIdx === -1) dateBeginIdx = 4;
  
  var plants = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[plantIdx]) continue; // Skip empty rows
    
    var dBegin = row[dateBeginIdx];
    var dateBeginStr = "";
    if (dBegin) {
      if (Object.prototype.toString.call(dBegin) === '[object Date]') {
        dateBeginStr = dBegin.toISOString();
      } else {
        dateBeginStr = String(dBegin);
      }
    }
    
    plants.push({
      idPlant: String(row[idPlantIdx]),
      plant: String(row[plantIdx]),
      round: String(row[roundIdx]),
      dateStart: String(row[dateStartIdx]),
      dateBegin: dateBeginStr
    });
  }
  return plants;
}

function getBookingsSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Bookings");
  if (!sheet) {
    // Create 'Bookings' sheet if it doesn't exist
    sheet = ss.insertSheet("Bookings");
    sheet.appendRow(["ID", "Date", "Plant", "DeliveryNumber", "ScrapNumber", "Timestamp", "Status"]);
  }
  return sheet;
}

function getBookings() {
  var sheet = getBookingsSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var bookings = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    bookings.push({
      id: String(row[0]),
      date: String(row[1]),
      plant: String(row[2]),
      deliveryNumber: String(row[3]),
      scrapNumber: String(row[4]),
      timestamp: String(row[5]),
      status: String(row[6] || "")
    });
  }
  return bookings;
}

function saveBooking(payload) {
  var sheet = getBookingsSheet();
  // payload: { id, date, plant, deliveryNumber, scrapNumber }
  
  // If editing, find existing
  var id = payload.id;
  if (!id) {
    id = "B" + new Date().getTime() + Math.floor(Math.random() * 1000);
  }
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      rowIndex = i + 1; // 1-based index
      break;
    }
  }
  
  var dateStr = payload.date; // e.g. "2026-03-25"
  var plant = payload.plant || "";
  var dNum = payload.deliveryNumber || "";
  var sNum = payload.scrapNumber || "";
  var ts = new Date().toISOString();
  var statusVal = payload.status || "";
  
  if (rowIndex !== -1) {
    // Edit existing
    sheet.getRange(rowIndex, 2, 1, 6).setValues([[dateStr, plant, dNum, sNum, ts, statusVal]]);
  } else {
    // Create new
    sheet.appendRow([id, dateStr, plant, dNum, sNum, ts, statusVal]);
  }
  
  return { status: "success", id: id };
}

function deleteBooking(payload) {
  var sheet = getBookingsSheet();
  var id = payload.id;
  if (!id) return { status: "error", message: "No ID provided" };
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex);
    return { status: "success" };
  } else {
    return { status: "error", message: "Booking not found" };
  }
}

/**
 * Save user action logs
 */
function saveLog(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Logs");
  if (!sheet) {
    sheet = ss.insertSheet("Logs");
    sheet.appendRow(["Timestamp", "UserID", "Name", "Action", "Details"]);
  }
  
  var ts = new Date().toISOString();
  var userId = payload.userId || "System";
  var name = payload.name || "Unknown";
  var actionStr = payload.action || "";
  var details = payload.details || "";
  
  sheet.appendRow([ts, userId, name, actionStr, details]);
  return { status: "success" };
}

/**
 * ----------------------------------------------------
 * AUTOMATED TELEGRAM NOTIFICATIONS
 * This function should be set up as a Daily Trigger 
 * (Edit -> Current project's triggers -> Add Trigger -> Time-driven -> Day timer)
 * ----------------------------------------------------
 */
function checkMissingBookingsAndNotify() {
  var TELEGRAM_TOKEN = "7787184267:AAGdxqbkQ6Xfb_G1mm_220tPNdvogGSNog4";
  var TELEGRAM_CHAT_ID = "-4562197204";
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  var plantSheet = ss.getSheetByName("PlantTranfer");
  if (!plantSheet) return;
  var plantData = plantSheet.getDataRange().getValues();
  if (plantData.length < 2) return;
  var headers = plantData[0];
  var idIdx=headers.indexOf("IDPlant"), pIdx=headers.indexOf("Plant");
  var rIdx=headers.indexOf("Round"), sIdx=headers.indexOf("DateStart"), bIdx=headers.indexOf("Datebegin");
  
  var masterPlants = [];
  for (var i=1; i<plantData.length; i++) {
    var row = plantData[i];
    masterPlants.push({
      idPlant: String(row[idIdx]||""),
      plant: String(row[pIdx]||""),
      round: String(row[rIdx]||"Non"),
      dateStart: String(row[sIdx]||""),
      dateBegin: row[bIdx] instanceof Date ? Utilities.formatDate(row[bIdx], "GMT+07:00", "dd/MM/yyyy") : String(row[bIdx]||"")
    });
  }
  
  var bookSheet = ss.getSheetByName("Bookings");
  var bookings = [];
  if (bookSheet) {
    var bookData = bookSheet.getDataRange().getValues();
    if (bookData.length > 1) {
      var bHeaders = bookData[0];
      var bdIdx=bHeaders.indexOf("Date"), bpIdx=bHeaders.indexOf("Plant");
      var dNumIdx=bHeaders.indexOf("DeliveryNumber"), sNumIdx=bHeaders.indexOf("ScrapNumber");
      for (var j=1; j<bookData.length; j++) {
        var brow = bookData[j];
        bookings.push({
          date: String(brow[bdIdx]||""),
          plant: String(brow[bpIdx]||""),
          hasData: ((brow[dNumIdx] && brow[dNumIdx]!=="") || (brow[sNumIdx] && brow[sNumIdx]!==""))
        });
      }
    }
  }

  function getWeekNumber(d) {
    if (!d) return 0;
    var date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    var dayNum = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - dayNum);
    var yearStart = new Date(Date.UTC(date.getUTCFullYear(),0,1));
    return Math.ceil((((date - yearStart) / 86400000) + 1)/7);
  }

  function parseDateStr(ds) {
    if (!ds) return null;
    var clean = ds.split(" ")[0].trim();
    if (clean.indexOf("/") > -1) {
        var parts = clean.split("/");
        if (parts.length >= 3) {
            var day = parseInt(parts[0], 10);
            var month = parseInt(parts[1], 10) - 1;
            var year = parseInt(parts[2], 10);
            return new Date(year, month, day);
        }
    }
    return null;
  }

  var today = new Date();
  var messageLines = ["🚨 *แจ้งเตือน: ข้อมูลการจองรถโอนอะไหล่ตามรอบคลังล่วงหน้าที่ยังไม่สมบูรณ์* 🚨\n"];
  var foundMissing = false;

  // Check today up to +3 days
  for (var offset = 0; offset <= 3; offset++) {
    var checkDate = new Date(today.getTime());
    checkDate.setDate(today.getDate() + offset);
    
    var dateStr = Utilities.formatDate(checkDate, "GMT+07:00", "dd/MM/yyyy");
    var dayOfWeek = checkDate.getDay() === 0 ? 7 : checkDate.getDay();
    var currentWM = getWeekNumber(checkDate);
    
    var missingPlants = [];
    
    for (var k=0; k<masterPlants.length; k++) {
        var p = masterPlants[k];
        if (p.round === "Non") continue; // Not active
        
        var starts = p.dateStart.split(',').map(function(s){return parseInt(s.trim(),10);});
        var isWeekDayMatch = starts.indexOf(dayOfWeek) !== -1;
        if (!isWeekDayMatch) continue;

        var isActive = false;
        if (p.round === "1") {
            isActive = true;
        } else if (p.round === "2") {
            var sDate = parseDateStr(p.dateBegin);
            if (sDate) {
                var sWeek = getWeekNumber(sDate);
                if (Math.abs(currentWM - sWeek) % 2 === 0) {
                    if (checkDate >= sDate) {
                        isActive = true;
                    }
                }
            }
        }
        
        if (isActive) {
            // Check if there is data in Bookings
            var hasBooking = false;
            for (var m=0; m<bookings.length; m++) {
                if (bookings[m].date === dateStr && bookings[m].plant === p.plant && bookings[m].hasData) {
                    hasBooking = true;
                    break;
                }
            }
            if (!hasBooking) {
                missingPlants.push("- " + p.idPlant + " " + p.plant);
            }
        }
    }
    
    if (missingPlants.length > 0) {
        foundMissing = true;
        messageLines.push("📅 *วันที่:* " + dateStr);
        for (var n=0; n<missingPlants.length; n++) {
            messageLines.push(missingPlants[n]);
        }
        messageLines.push(" ");
    }
  }
  
  if (foundMissing) {
    var text = messageLines.join("\n");
    var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
    var payloadObj = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "Markdown"
    };
    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payloadObj)
    };
    try {
      UrlFetchApp.fetch(url, options);
    } catch(e) {
      Logger.log("Telegram Error: " + e.toString());
    }
  }
}
