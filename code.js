function generateSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var availabilitySheet = ss.getSheetByName('availability');
  var scheduleSheet = ss.getSheetByName('final schedule');
  var logSheet = ss.getSheetByName('log');
  var ledgerSheet = ss.getSheetByName('ledger');
  
  // === Setup ===
  var days = ["Tuesday", "Thursday", "Friday"];
  var timestamp = Utilities.formatDate(new Date(), "America/Denver", "yyyy-MM-dd HH:mm");
  


  // === Clear schedule (but keep headers) ===
  scheduleSheet.getRange("A3:C").clearContent();
  
  // === Load data ===
  var { availability, people, partners } = parseAvailability(availabilitySheet);
  var ledger = loadLedger(ledgerSheet);
  var schedule = [];

  // === Assign volunteers for each day ===
  days.forEach(function(day) {
    var chosen = assignVolunteersToDay(day, availability, partners, ledger);
    if (chosen.length > 0) {
      schedule.push([day, chosen[0], chosen[1]]);

    } else {
      schedule.push([day, "Not enough volunteers", ""]);
    }
  });
  
  // === Save new schedule ===
  scheduleSheet.getRange("A1").setValue("Last Updated");
  scheduleSheet.getRange("B1").setValue(timestamp);
  scheduleSheet.getRange(3, 1, schedule.length, 3).setValues(schedule);

  // === Log update ===
  logSheet.appendRow([timestamp, "Schedule Updated", JSON.stringify(schedule)]);
}

/**
 * Parse availability into structures we can use
 */
function parseAvailability(sheet) {
  var data = sheet.getDataRange().getValues();
  var availability = { "Tuesday": [], "Thursday": [], "Friday": [] };
  var people = [];
  var partners = {};
  
  for (var i = 1; i < data.length; i++) {
    var name = (data[i][0] || "").trim();
    if (!name) continue;
    people.push(name);
    
    if ((data[i][1] || "").toLowerCase() == 'yes') availability["Tuesday"].push(name);
    if ((data[i][2] || "").toLowerCase() == 'yes') availability["Thursday"].push(name);
    if ((data[i][3] || "").toLowerCase() == 'yes') availability["Friday"].push(name);
    
    if (data[i][4]) {
      partners[name] = data[i][4].split(',').map(p => p.trim());
    }
  }
  return { availability, people, partners };
}

/**
 * Assign two volunteers to a given day
 * Weighted by lowest ledger counts (fairness)
 */
function assignVolunteersToDay(day, availability, partners, ledger) {
  var candidates = availability[day] || [];
  if (candidates.length < 2) return [];
  
  // Sort candidates by how many times they’ve served (ascending)
  candidates.sort(function(a, b) {
    return (ledger[a] || 0) - (ledger[b] || 0);
  });
  
  // Try to honor partner preferences
  for (var i = 0; i < candidates.length; i++) {
    var p = candidates[i];
    if (partners[p]) {
      for (var j = 0; j < partners[p].length; j++) {
        var partner = partners[p][j];
        if (candidates.includes(partner)) {
          return [p, partner];
        }
      }
    }
  }

  // Otherwise, pick the two lowest-count people
  return [candidates[0], candidates[1]];
}

/**
 * === Ledger Helpers ===
 */

// Load ledger as { name: count }
function loadLedger(sheet) {
  var data = sheet.getDataRange().getValues();
  var ledger = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) ledger[data[i][0]] = data[i][1];
  }
  return ledger;
}

/**
 * Fisher-Yates shuffle (unused in fairness mode but kept for flexibility)
 */
function shuffle(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
}

/**
 * Extend the schedule by adding Tuesdays, Thursdays, and Fridays for one month forward
 * This function reads the existing schedule and continues from the last entry
 */
function extendScheduleOneMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName('final schedule');
  
  if (!scheduleSheet) {
    Logger.log('Error: "final schedule" sheet not found');
    return;
  }
  
  // Get all data from the schedule sheet
  var data = scheduleSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('No existing schedule data found');
    return;
  }
  
  // Find the last entry with a date (skip header row)
  var lastDate = null;
  var lastRowIndex = -1;
  
  for (var i = data.length - 1; i >= 1; i--) {
    var dateValue = data[i][1];
    if (dateValue) {
      // Try to parse the date regardless of format
      var parsedDate = parseDate(dateValue);
      if (parsedDate) {
        lastDate = parsedDate;
        lastRowIndex = i;
        break;
      }
    }
  }
  
  if (!lastDate) {
    Logger.log('No valid date found in existing schedule');
    Logger.log('Available data in column B: ' + JSON.stringify(data.map(function(row) { return row[1]; })));
    return;
  }
  
  Logger.log('Last date found: ' + lastDate);
  
  // Generate dates for the next month (Tuesdays, Thursdays, Fridays)
  var newEntries = generateNextMonthDates(lastDate);
  
  if (newEntries.length === 0) {
    Logger.log('No new dates to add');
    return;
  }
  
  // Add new entries to the schedule
  var startRow = lastRowIndex + 2; // Start after the last entry
  
  for (var i = 0; i < newEntries.length; i++) {
    var entry = newEntries[i];
    var row = startRow + i;
    
    // Set Month, Date, Day columns
    scheduleSheet.getRange(row, 1).setValue(entry.month);
    scheduleSheet.getRange(row, 2).setValue(entry.date);
    scheduleSheet.getRange(row, 3).setValue(entry.day);
    
    // Leave Who and Last Updated columns empty for now
    scheduleSheet.getRange(row, 4).setValue('');
    scheduleSheet.getRange(row, 5).setValue('');
  }
  
  // Update the timestamp
  var timestamp = Utilities.formatDate(new Date(), "America/Denver", "yyyy-MM-dd HH:mm");
  scheduleSheet.getRange("F1").setValue(timestamp);
  
  Logger.log('Added ' + newEntries.length + ' new schedule entries');
}

/**
 * Parse various date formats that might be in the spreadsheet
 */
function parseDate(dateValue) {
  if (!dateValue) return null;

  // If it's already a Date object
  if (dateValue instanceof Date) {
    return dateValue;
  }

  // Generic attempt: works for ISO strings and other parseable inputs
  var generic = new Date(dateValue);
  if (!isNaN(generic.getTime())) {
    return generic;
  }

  // If it's a number (Excel serial number)
  if (typeof dateValue === 'number') {
    var millis = Math.round((dateValue - 25569) * 86400 * 1000);
    var serialDate = new Date(millis);
    if (!isNaN(serialDate.getTime())) return serialDate;
  }

  // As a final fallback, try common US formats by normalizing separators
  if (typeof dateValue === 'string') {
    var normalized = dateValue.replace(/-/g, '/');
    var fallback = new Date(normalized);
    if (!isNaN(fallback.getTime())) return fallback;
  }

  return null;
}

/**
 * Generate the next month's worth of Tuesdays, Thursdays, and Fridays
 * starting from the day after the last date
 */
function generateNextMonthDates(lastDate) {
  var targetDays = [2, 4, 5]; // Tuesday=2, Thursday=4, Friday=5
  var entries = [];
  
  // Start from the day after the last date
  var currentDate = new Date(lastDate);
  currentDate.setDate(currentDate.getDate() + 1);
  
  // Generate dates for approximately one month (30 days)
  var endDate = new Date(currentDate);
  endDate.setDate(endDate.getDate() + 30);
  
  while (currentDate <= endDate) {
    var dayOfWeek = currentDate.getDay();
    
    if (targetDays.includes(dayOfWeek)) {
      var monthName = currentDate.toLocaleDateString('en-US', { month: 'long' });
      var formattedDate = Utilities.formatDate(currentDate, "America/Denver", "M/d/yyyy");
      var dayName = currentDate.toLocaleDateString('en-US', { weekday: 'long' });
      
      entries.push({
        month: monthName,
        date: formattedDate,
        day: dayName
      });
    }
    
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  return entries;
}
