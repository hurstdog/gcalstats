'use strict';

// Calendar Stats
// This Apps Script uses the Sheets and Calendar APIs to generate statistics
// about the owners calendar, and throws them into a handy spreadsheet.

/////////////////////////////////////////////////////////////////////////////////
// Constants.
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
const ONE_ON_ONE_HDR_RANGE = "A1:D1";

// Headers for the stats rows. Note that this is the order needed in the stats
// frequency dict as well.
const ONE_ON_ONE_STATS_HDRS = [["Who",
                                "Time Since Last 1:1",
                                "SLO",
                                "Overdue"]];

// End Constants. Below is just code, and bad code at that. Ignore it.
/////////////////////////////////////////////////////////////////////////////////


class OneOnOneStatCollector {
  constructor(sheet) {
    // Map of '1:1 partner => [days since last 1:1 (`freqMap`), sla]
    this.oneOnOneFreq = {};
    this._populateOneOnOneFreq(sheet)
  }

  // Returns the oneOnOneFreq dictionary populated with data from the Stats Sheet.
  // '1:1 partner' => [undef, 1:1 frequency SLO]
  _populateOneOnOneFreq(sheet) {
    var r = sheet.getRange('A2:C200');

    var freq = {}

    // Loop over the range, populating the frequency map as we go.
    for (const row of r.getValues()) {
      if (row[0] == "") {
        break;
      }
      var sla = "";
      if (row[2]) {
        sla = row[2];
      };
      freq[row[0]] = [row[1], sla];
      //Logger.log(row[0] + ' = [' + freq[row[0]] + ']');
    }

    //printFreq(freq);
    this.oneOnOneFreq = freq;
  }

  // This extracts out the 1:1 partner name and updates it with the minimum gap since the last 1:1.
  // Skips any events in the future.
  // event: CalendarEvent
  trackOneOnOne(event) {
    const now = new Date();
    const guest = cleanGuestEmail(getOneOnOneGuestEmail(event));

    var diffMs = now - event.getStartTime();
    if (diffMs < 0) {
      //Logger.log(event.getStartTime() + ' is in the future so I\'m skipping it');
      return;
    }

    const daysSinceEvent = Math.floor(diffMs / 1000 / 60 / 60 / 24);

    if (guest in this.oneOnOneFreq) {
      // Only update if this event is newer
      if (this.oneOnOneFreq[guest][0] > daysSinceEvent) {
        this.oneOnOneFreq[guest][0] = daysSinceEvent;
      }
    } else {
      if (this.oneOnOneFreq[guest] == undefined) {
        this.oneOnOneFreq[guest] = []
      }
      this.oneOnOneFreq[guest][0] = daysSinceEvent;
    }

    //Logger.log('Most recent 1:1 with ' + guest + ': ' + this.oneOnOneFreq[guest]);
  }

  // Takes a Spreadsheet and populates it with the 1:1 statistics
  updateStatsSheet(sheet) {
    var freqEntries = Object.entries(this.oneOnOneFreq);

    // Set and freeze the column headers
    var r = sheet.getRange(ONE_ON_ONE_HDR_RANGE);
    r.setValues(ONE_ON_ONE_STATS_HDRS);
    r.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 4);

    // Generate the range
    var range = ['A2:D', freqEntries.length + 1].join("");

    // Populate the data
    r = sheet.getRange(range);

    //r.setValues(Object.entries(flattenFreq(oneOnOneFreq)));
    r.setValues(flattenFreq(this.oneOnOneFreq));

    // Sort the data
    r.sort({column: 4, ascending: false});
  }
}

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
// TODO: Update stats in a sheet, with columns based on the tags
// TODO: Collect stats on 1:1 frequency as well.
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

  stats.updateStatsSheet(getStatsSheet());
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

// Returns the Spreadsheet object used to store statistics
function getStatsSheet() {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONE_ON_ONE_STATS_SHEET);
  if (sheet == null) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(ONE_ON_ONE_STATS_SHEET);
    sheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  return sheet;
}

function flattenFreq(freq) {
  var f = [];
  for (const [k, v] of Object.entries(freq)) {
    f.push([k, v[0], v[1], v[0] - v[1]]);
  }
  //printFlatFreq(f);
  return f;
}

// Increments element `tag` in dictionary `dict`
function inc(dict, tag) {
  if (tag in dict) {
    dict[tag]++;
  } else {
    dict[tag] = 1;
  }
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

function printFreq(freq) {
  Logger.log('Printing the 1:1 frequency');
  for (const [k, v] of Object.entries(freq)) {
    if (k && v) {
      Logger.log(k + ' -> [' + v + ']');
    }
  }
}

function printFlatFreq(freq) {
  Logger.log('Printing the flat freq');
  for (const row of freq) {
    Logger.log('row is [' + row + ']');
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
