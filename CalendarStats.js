'use strict';

// Calendar Stats
// This Apps Script uses the Sheets and Calendar APIs to generate statistics
// about the owners calendar, and throws them into a handy spreadsheet.

/////////////////////////////////////////////////////////////////////////////////
// Constants
//
// Edit the entries in the following section to tailor the script
// to your preferences.

// Your Google Calendar username and domain.
const OWNER_USERNAME = 'andrew';
const OWNER_DOMAIN = 'hurstdog.org';
const OWNER_EMAIL = OWNER_USERNAME + '@' + OWNER_DOMAIN;
const MEETING_TAG = 'TAG: ';

// Counting days from today, forward or backwards.
// Note that both of these values should be positive
const RANGE_DAYS_PAST = 30;
const RANGE_DAYS_FUTURE = 30;

// Name of the sheet to show the results.
// Will create if it doesn't exist, otherwise will re-use the existing.
const ONE_ON_ONE_STATS_SHEET = '1:1 Stats';
const ONE_ON_ONE_HDR_RANGE = "A1:E1";

// Headers for the stats rows. Note that this is the order needed in the stats
// frequency dict as well.
const ONE_ON_ONE_STATS_HDRS = [["Who",
                                "Time Since Last 1:1",
                                "Time Until next 1:1",
                                "SLO",
                                "Overdue"]];

// End Constants. Below is just code, and bad code at that. Ignore it.
/////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////
//
// OneOnOneStatCollector
//
// This class contains all of the knowledge for parsing the one on one stats
// spreadsheet, collating the stats from a series of CalendarEvents, and then
// updating the stats spreadsheet.
class OneOnOneStatCollector {

  // Given a Spreadsheet, constructs & populates the stats collector based on the
  // data in that spreadsheet.
  //
  // Arguments:
  //   sheet: Spreadsheet containing 1:1 information. Stored in the class.
  constructor(sheet) {
    this.sheet = sheet;

    // Map of '1:1 partner email' => [
    //        days since last 1:1,
    //        days until next 1:1,
    //        frequency SLO,
    //        ]
    this.oneOnOneFreq = {};
    this._populateOneOnOneFreq()
  }

  // Populates the oneOnOneFreq dictionary with data from the Stats Sheet.
  // NOTE: This assumes that the sheet has the columns in ONE_ON_ONE_STATS_HDRS,
  //  in that order.
  //
  _populateOneOnOneFreq() {
    var r = this.sheet.getRange('A2:D200');

    var freq = {}

    // Loop over the range, populating the frequency map as we go.
    for (const row of r.getValues()) {
      var email = row[0];
      var lastO = row[1];
      var nextO = row[2];
      var slo = row[3];
      var over = row[4];

      if (email == "") {
        break;
      }

      freq[email] = [lastO, nextO, slo, over];
      //Logger.log(email + ' = [' + freq[email] + ']');
    }

    //this._printFreq(freq);
    this.oneOnOneFreq = freq;
  }

  // Given a CalendarEvent that is assumed to be a 1:1, this extracts out the 1:1 partner name and
  // updates the statistics for 1:1s with that person.
  //
  // TODO: Collect stats on 1:1 frequency as well.
  //
  // Arguments:
  //   event: CalendarEvent
  trackOneOnOne(event) {
    const now = new Date();
    const guest = cleanGuestEmail(getOneOnOneGuestEmail(event));

    // Temp variables so I don't have to spend all my mental energy with array indices
    var guestStats = this.oneOnOneFreq[guest]
    if (guestStats == undefined) {
      guestStats = []
    }

    var lastOneOnOneDelta = guestStats[0]; // days in the past
    var nextOneOnOneDelta = guestStats[1]; // days in the future

    const diffMs = now - event.getStartTime();
    const daysToEvent = Math.floor(diffMs / 1000 / 60 / 60 / 24);

    if (diffMs > 0) {
      // Past events
      // Update the last 1:1 time if
      //    a) it hasn't been defined or
      //    b) it's farther away than the current event.
      if (lastOneOnOneDelta == undefined || lastOneOnOneDelta > daysToEvent) {
        lastOneOnOneDelta = daysToEvent;
      }
    } else {
      // Future events
      // Update the last 1:1 time if
      //    a) it hasn't been defined or
      //    b) it's farther away (more negative) than the current event.
      if (nextOneOnOneDelta == undefined || nextOneOnOneDelta < daysToEvent) {
        nextOneOnOneDelta = daysToEvent;
      }
    }

    guestStats[0] = lastOneOnOneDelta;
    guestStats[1] = nextOneOnOneDelta;
    this.oneOnOneFreq[guest] = guestStats;

    //Logger.log('Most recent 1:1 with ' + guest + ': ' + this.oneOnOneFreq[guest]);
  }

  // Populates the stored Spreadsheet with the statistics stored in this class.
  //
  updateStatsSheet() {
    var freqEntries = Object.entries(this.oneOnOneFreq);

    // Set and freeze the column headers
    var r = this.sheet.getRange(ONE_ON_ONE_HDR_RANGE);
    r.setValues(ONE_ON_ONE_STATS_HDRS);
    r.setFontWeight('bold');
    this.sheet.setFrozenRows(1);

    // Generate the range
    var range = ['A2:E', freqEntries.length + 1].join("");

    // Populate the data
    r = this.sheet.getRange(range);

    r.setValues(this.getFlatFreq());

    // Sort the data
    r.sort({column: 4, ascending: false});

    // Resize columns last, to match the data we just added.
    this.sheet.autoResizeColumns(1, 4);
  }

  // Returns the email frequency data structure as an array of arrays, ready to
  // pass to Range.setValues()
  //
  getFlatFreq() {
    var f = [];
    for (const [k, v] of Object.entries(this.oneOnOneFreq)) {

      var guest = k;
      var last = v[0];
      var next = v[1];
      var slo = v[2];
      var outslo = last - slo;

      // If we don't have useful variables to calculate how far out of SLO we are,
      // try to show something reasonable.
      if (slo == undefined) {
        outslo = "";
      } else {
        if (last == undefined) {
          outslo = slo;
        }
      }

      f.push([guest, last, next, slo, outslo]);
    }
    //this._printFlatFreq(f);
    return f;
  }

  _printFlatFreq(freq) {
    Logger.log('Printing the flat freq');
    for (const row of freq) {
      Logger.log('row is [' + row + ']');
    }
  }

  _printFreq(freq) {
    Logger.log('Printing the 1:1 frequency');
    for (const [k, v] of Object.entries(freq)) {
      if (k && v) {
        Logger.log(k + ' -> [' + v + ']');
      }
    }
  }

}

// End OneOnOneStatCollector
/////////////////////////////////////////////////////////////////////////////////

// Adds a custom menu item to run the script
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu('Calendar Script',
             [{name: 'Update Meeting Stats', functionName: 'ListMeetings'}]);
}

function ListMeetings() {
  var events = CalendarApp.getDefaultCalendar().getEvents(getDateByDays(RANGE_DAYS_PAST * -1),
                                                          getDateByDays(RANGE_DAYS_FUTURE));
  
  reportStats(events)
}

// Reports statistics on the CalendarEvent[] passed in to `events`, by percentage of time broken down.
// 1:1s
// Group meetings
// Focused time?
function reportStats(events) {
  var numEvents = events.length;

  var oneOnOnes = 0;
  var blockedTime = 0;
  var meetings = 0;

  // 'tag name' => count
  var tags = {};
  // '1:1 partner' => [days since last 1:1, 1:1 frequency SLO]
  var stats = new OneOnOneStatCollector(getStatsSheet());
  for (const event of events) {
    var tag = extractTag(event);
    var guests = event.getGuestList(true);
    if (tag != null) {
      inc(tags, tag);
    } else if (guests.length == 0) {
      blockedTime++;
    } else if (guests.length == 1 && guests[0].getEmail() == OWNER_EMAIL) {
      blockedTime++;
    } else if (guests.length == 2) {
      oneOnOnes++;
      stats.trackOneOnOne(event);
      //Logger.log('Found a 1:1 with guests! ' + event.getTitle());
      //printGuests(guests);
    } else {
      meetings++;
      //Logger.log('OTHER Title: ' + event.getTitle());
      //printGuests(guests)
    }
  }

  Logger.log('Total: ' + numEvents);
  Logger.log('1:1s: ' + oneOnOnes);
  Logger.log('Blocked Time: ' + blockedTime);
  Logger.log('Meetings: ' + meetings);
  for (const [tag, count] of Object.entries(tags)) {
    Logger.log(tag + ': ' + count);
  }

  stats.updateStatsSheet();
}

// Given and event with two guests, returns the guest email that isn't OWNER_EMAIL
// Prints an error and returns null on lists that don't contain two entries.
function getOneOnOneGuestEmail(event) {
  var guestList = event.getGuestList(true);
  if (guestList.length != 2) {
    Logger.log('Too many guests in purported 1:1 (Title: ' + event.getTitle() + ', skipping');
    return null;
  }

  var guest = "";
  for (const g of guestList) {
    if (g.getEmail() != OWNER_EMAIL) {
      guest = g.getEmail();
    }
  }

  return guest;
}

// Given an email address, strips off the domain if it's the same as OWNER_DOMAIN
function cleanGuestEmail(email) {
  return email.replace('@' + OWNER_DOMAIN, '');
}

// Returns the Spreadsheet object used to store statistics, and creates it if one
// doesn't exist yet.
function getStatsSheet() {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONE_ON_ONE_STATS_SHEET);
  if (sheet == null) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(ONE_ON_ONE_STATS_SHEET);
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  return sheet;
}

// Given a CalendarEvent, will read the description and return any of the text
// on a line after the keyword MEETING_TAG (currently 'TAG: ')
// e.g. return $1 from "^\w*TAG: (.*)\w*$"
function extractTag(event) {
  var tag = null;
  var desc = event.getDescription();
  var lines = desc.split("\n");
  for (const line of lines) {
    var trimline = line.trim();
    if (trimline.startsWith(MEETING_TAG)) {
      tag = trimline.substr(MEETING_TAG.length);
      break;
    }
  }
  return tag;
}


// Increments element `tag` in dictionary `dict`
function inc(dict, tag) {
  if (tag in dict) {
    dict[tag]++;
  } else {
    dict[tag] = 1;
  }
}

function printGuests(guests) {
  for (const guest of guests) {
    Logger.log('-> Guest: ' + guest.getEmail());
  }
}

function printTitles(events) {
  for (const event of events) {
    Logger.log('  Title: ' + event.getTitle());
  }
}

// Returns a date object that is `days` in the future or past, depending on positive or
// negative value.
function getDateByDays(days) {
  // Determines how many events are happening in the next two hours.
  var now = new Date();
  return new Date(now.getTime() + (days * 24 * 60 * 60 * 1000));
}
