// generateSchedule was deprecated in favor of main* entry points

// Single entry point for triggers or manual runs
function mainSchedule() {
  mainExtendScheduleOneMonth();
  mainFillMissingAssignments();
}


/**
 * Fill any rows in the final schedule that have dates but are missing
 * one or both names in columns D and E. Uses availability and a working
 * copy of the ledger for fairness (do not write to the ledger sheet).
 * Partner preference is a soft constraint and is only applied when it
 * does not violate the fairness guard (avoid differences > 1 from min).
 */
function mainFillMissingAssignments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var availabilitySheet = ss.getSheetByName('availability');
  var scheduleSheet = ss.getSheetByName('final schedule');
  var ledgerSheet = ss.getSheetByName('ledger');
  var logSheet = ss.getSheetByName('log');

  if (!availabilitySheet || !scheduleSheet || !ledgerSheet) {
    Logger.log('Required sheets not found');
    return;
  }

  var parsed = parseAvailability(availabilitySheet);
  var availability = parsed.availability;
  var partners = parsed.partners;
  var baseLedger = loadLedger(ledgerSheet);

  // Working copy used only for selection fairness within this run
  var workingLedger = {};
  Object.keys(baseLedger).forEach(function(name) { workingLedger[name] = baseLedger[name]; });

  var data = scheduleSheet.getDataRange().getValues();
  var updates = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateValue = row[1];
    var dayName = (row[2] || '').toString().trim();
    if (!dateValue || !dayName) continue;

    var who1 = (row[3] || '').toString().trim();
    var who2 = (row[4] || '').toString().trim();
    if (who1 && who2) continue;

    var existing = who1 || who2 || '';
    var candidates = getEligibleCandidatesForRow(dayName, availability, [who1, who2], workingLedger);
    if (candidates.length === 0) continue;

    if (!who1 && !who2) {
      var pair = selectPairFromCandidates(candidates, partners, workingLedger);
      if (pair[0]) scheduleSheet.getRange(i + 1, 4).setValue(pair[0]);
      if (pair[1]) scheduleSheet.getRange(i + 1, 5).setValue(pair[1]);
      updates.push({ row: i + 1, who: pair.slice(0, 2) });
    } else {
      // Only one blank; fill only that cell
      var second = selectSecondWithPreference(existing, candidates, partners, workingLedger);
      if (!second) continue;
      if (!who1) scheduleSheet.getRange(i + 1, 4).setValue(second);
      if (!who2) scheduleSheet.getRange(i + 1, 5).setValue(second);
      updates.push({ row: i + 1, who: [existing, second].filter(function(x){return x;}) });
    }
  }

  // Sheet-level timestamp only
  var timestamp = Utilities.formatDate(new Date(), 'America/Denver', 'yyyy-MM-dd HH:mm');
  scheduleSheet.getRange('F1').setValue(timestamp);
  if (logSheet) {
    logSheet.appendRow([timestamp, 'Filled Missing Assignments', JSON.stringify(updates)]);
  }
}


/**
 * Extend the schedule by adding Tuesdays, Thursdays, and Fridays for one month forward
 * This function reads the existing schedule and continues from the last entry
 */
function mainExtendScheduleOneMonth() {
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


// === Helpers used by fillMissingAssignments ===
function getEligibleCandidatesForRow(dayName, availability, excludeNames, workingLedger) {
  var excluded = {};
  (excludeNames || []).forEach(function(n){ if (n) excluded[n] = true; });
  var candidates = (availability[dayName] || []).filter(function(name){ return !excluded[name]; });
  candidates.sort(function(a, b) { return (workingLedger[a] || 0) - (workingLedger[b] || 0); });
  return candidates;
}

function selectPairFromCandidates(candidates, partners, workingLedger) {
  if (candidates.length === 0) return [];
  var first = candidates[0];
  incrementWorking(workingLedger, first);

  var remaining = candidates.slice(1);
  remaining.sort(function(a, b) { return (workingLedger[a] || 0) - (workingLedger[b] || 0); });
  var second = chooseSecondFromList(first, remaining, partners, workingLedger);
  if (second) incrementWorking(workingLedger, second);
  return [first, second].filter(function(x){return x;});
}

function selectSecondWithPreference(existingName, candidates, partners, workingLedger) {
  var second = chooseSecondFromList(existingName, candidates, partners, workingLedger);
  if (second) incrementWorking(workingLedger, second);
  return second;
}

function chooseSecondFromList(anchor, candidates, partners, workingLedger) {
  if (candidates.length === 0) return null;
  var minCount = (workingLedger[candidates[0]] || 0);
  var maxAllowed = minCount + 1;
  var eligible = candidates.filter(function(n){ return (workingLedger[n] || 0) <= maxAllowed; });

  if (eligible.length > 0 && partners[anchor]) {
    var preferred = eligible.filter(function(n){ return partners[anchor].indexOf(n) !== -1; });
    if (preferred.length > 0) {
      preferred.sort(function(a, b){ return (workingLedger[a] || 0) - (workingLedger[b] || 0); });
      return preferred[0];
    }
  }
  return candidates[0];
}

function incrementWorking(workingLedger, name) {
  workingLedger[name] = (workingLedger[name] || 0) + 1;
}

