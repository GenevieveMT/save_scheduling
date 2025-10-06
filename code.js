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
