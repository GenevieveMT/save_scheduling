function generateSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var availabilitySheet = ss.getSheetByName('availability');
  var scheduleSheet = ss.getSheetByName('final schedule');
  var logSheet = ss.getSheetByName('log');
  
  // Clear previous schedule (but keep headers)
  scheduleSheet.getRange("A3:C").clearContent();
  
  // Get data from the availability sheet
  var { availability, people, partners } = parseAvailability(availabilitySheet);
  
  var days = ["Tuesday", "Thursday", "Friday"];
  var schedule = [];
  var cooldown = [...people];  // Copy of all people to avoid repeats
  
  // Assign volunteers to each day
  days.forEach(function(day) {
    var chosen = assignVolunteersToDay(day, availability, partners, cooldown);
    if (chosen.length > 0) {
      schedule.push([day, chosen[0], chosen[1]]);
      cooldown = cooldown.filter(p => !chosen.includes(p));  // Remove chosen ones
    } else {
      schedule.push([day, "Not enough volunteers", ""]);
    }
  });
  
  // Write the "Last Updated" timestamp
  var timestamp = Utilities.formatDate(new Date(), "America/Denver", "yyyy-MM-dd HH:mm");
  scheduleSheet.getRange("A1").setValue("Last Updated");
  scheduleSheet.getRange("B1").setValue(timestamp);
  
  // Write the schedule starting at row 3
  scheduleSheet.getRange(3, 1, schedule.length, 3).setValues(schedule);
  
  // Log update in the log sheet with details
  logSheet.appendRow([timestamp, "Schedule Updated", JSON.stringify(schedule)]);
}

/**
 * Parse availability data into structures we can use
 */
function parseAvailability(sheet) {
  var data = sheet.getDataRange().getValues();
  var availability = { "Tuesday": [], "Thursday": [], "Friday": [] };
  var people = [];
  var partners = {};
  
  for (var i = 1; i < data.length; i++) {
    var name = (data[i][0] || "").trim();
    if (!name) continue; // skip empty rows
    people.push(name);
    
    // Mark availability
    if ((data[i][1] || "").toLowerCase() == 'yes') availability["Tuesday"].push(name);
    if ((data[i][2] || "").toLowerCase() == 'yes') availability["Thursday"].push(name);
    if ((data[i][3] || "").toLowerCase() == 'yes') availability["Friday"].push(name);
    
    // Preferred partners
    if (data[i][4]) {
      partners[name] = data[i][4].split(',').map(p => p.trim());
    }
  }
  
  return { availability, people, partners };
}

/**
 * Assign two volunteers to a given day
 */
function assignVolunteersToDay(day, availability, partners, cooldown) {
  var candidates = availability[day] || [];
  candidates = candidates.filter(p => cooldown.includes(p)); // respect cooldown
  
  if (candidates.length < 2) return []; // not enough
  
  // Try to find a preferred partner pairing
  for (var i = 0; i < candidates.length; i++) {
    var p = candidates[i];
    if (partners[p]) {
      for (var j = 0; j < partners[p].length; j++) {
        var partner = partners[p][j];
        if (candidates.includes(partner)) {
          return [p, partner]; // success
        }
      }
    }
  }
  
  // Otherwise pick two random people
  shuffle(candidates);
  return [candidates[0], candidates[1]];
}

/**
 * Fisher-Yates shuffle
 */
function shuffle(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
}
