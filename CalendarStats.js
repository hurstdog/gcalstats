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
const RANGE_DAYS_PAST = 60;
const RANGE_DAYS_FUTURE = 30;

// How many hours a day do you work?  For now, 9-6pm, or 9 hours a day (45 hour week).
// This is used to exclude events outside work hours from tracking.
// Note that if you regularly have meetings outside of working hours, I can't help you :P
const START_HOUR = 9;
const END_HOUR = 18;
const WORK_HOURS_PER_DAY = END_HOUR - START_HOUR;
const WORK_DAYS_PER_WEEK = 5;

// String for tracking unscheduled time. If a week has no meetings, you'll have
// WORK_HOURS_PER_DAY * WORK_DAYS_PER_WEEK hours of UNSCHEDULED_TIME.
const UNSCHEDULED_TIME = "Unscheduled";

/////////////////////////////////////////////////////////////////////////////////
// End user configurable Constants.
// Below here are constants for the running of the script, and you probalby
// shouldn't change them without fussing with the script code as well.

// Name of the sheet to show the 1:1 ranking results.
// Will create if it doesn't exist, otherwise will re-use the existing.
const ONE_ON_ONE_LIST_SHEET = "1:1 List";
const ONE_ON_ONE_LIST_HDR_RANGE = "A1:F1";
const ONE_ON_ONE_LIST_DATA_RANGE_COLS = "A2:F";
const ONE_ON_ONE_LIST_DATA_RANGE_MAX = "A2:F200";

// Headers for the stats rows. Note that this is the order needed in the stats
// frequency dict as well.
const ONE_ON_ONE_LIST_HDRS = [["Who",
                                "Time Since Last 1:1",
                                "Time Until next 1:1",
                                "SLO",
                                "Days Overdue",
                                "Notes"]];

// Which column to sort the results by. This corresponds to ONE_ON_ONE_LIST_HDRS.
const ONE_ON_ONE_LIST_SORT_COLUMN = 5;

// Name of the sheet to show the 1:1 stats results.
// Will create if it doesn't exist, otherwise will re-use the existing.
const MEETING_STATS_SHEET = "Meeting Stats";
const MEETING_STATS_HDR_RANGE = "A1:E1";
const MEETING_STATS_DATA_RANGE_COLS = "A2:E";

const MEETING_STATS_DATA_HOUR_RANGE_COLS = "B2:E";

// Headers for the stats rows. Note that this is the order needed in the stats
// frequency dict as well.
// TODO: Support tags.
const MEETING_STATS_HDRS = [["Week",
                             "1:1s",
                             "Meetings",
                             "Blocked",
                             UNSCHEDULED_TIME]];

// Which column to sort the results by. This corresponds to MEETING_STATS_HDRS.
const MEETING_STATS_SORT_COLUMN = 1;

// End Constants. Below is just code, and bad code at that. Ignore it.
/////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////
//
// OneOnOneListCollector
//
// This class contains all of the knowledge for parsing the one on one stats
// spreadsheet, collating the stats from a series of CalendarEvents, and then
// updating the stats spreadsheet.
class OneOnOneListCollector {

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
    //        overdue,
    //        Notes,
    //        ]
    this.oneOnOneFreq = {};
    this._populateOneOnOneFreq()
  }

  // Populates the oneOnOneFreq dictionary with data from the Stats Sheet.
  // NOTE: This assumes that the sheet has the columns in ONE_ON_ONE_LIST_HDRS,
  //  in that order.
  //
  _populateOneOnOneFreq() {
    var r = this.sheet.getRange(ONE_ON_ONE_LIST_DATA_RANGE_MAX);

    var freq = {}

    // Loop over the range, populating the frequency map as we go.
    for (const row of r.getValues()) {
      var email = row[0];
      var lastO = row[1];
      var nextO = row[2];
      var slo = row[3];
      var over = row[4];
      var notes = row[5];

      if (email == "") {
        break;
      }

      freq[email] = [lastO, nextO, slo, over, notes];
    }

    //this._printFreq(freq);
    this.oneOnOneFreq = freq;
  }

  // Given a CalendarEvent that is might be a 1:1, extracts out the 1:1 partner name and
  // updates the statistics for 1:1s with that person (if it's a 1:1, naturally).
  //
  // TODO: Collect stats on 1:1 frequency as well.
  //
  // Arguments:
  //   event: CalendarEvent
  maybeTrackOneOnOne(event) {
    if (!this._isOneOnOne(event)) {
      return;
    }

    const now = new Date();
    const guest = this._cleanGuestEmail(this._getOneOnOneGuestEmail(event));

    // Temp variables so I don't have to spend all my mental energy with array indices
    var guestStats = this.oneOnOneFreq[guest];
    if (guestStats == undefined) {
      guestStats = [];
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
  updateListSheet() {
    var freqEntries = Object.entries(this.oneOnOneFreq);

    // Set and freeze the column headers
    var r = this.sheet.getRange(ONE_ON_ONE_LIST_HDR_RANGE);
    r.setValues(ONE_ON_ONE_LIST_HDRS);
    r.setFontWeight('bold');
    this.sheet.setFrozenRows(1);

    // Generate the range
    var range = [ONE_ON_ONE_LIST_DATA_RANGE_COLS, freqEntries.length + 1].join("");

    // Populate the data
    r = this.sheet.getRange(range);

    r.setValues(this._getFlatFreq());

    // Sort the data
    r.sort({column: ONE_ON_ONE_LIST_SORT_COLUMN, ascending: false});

    // Resize columns last, to match the data we just added.
    this.sheet.autoResizeColumns(1, ONE_ON_ONE_LIST_HDRS[0].length);
  }

  // Given a CalendarEvent, this will return true if it's a 1:1, false otherwise.
  //
  // Argument:
  //   event: CalendarEvent
  _isOneOnOne(event) {
    var guests = event.getGuestList(true);
    return guests.length == 2;
  }

  // Returns the email frequency data structure as an array of arrays, ready to
  // pass to Range.setValues()
  //
  _getFlatFreq() {
    var f = [];
    for (const [k, v] of Object.entries(this.oneOnOneFreq)) {

      var guest = k;
      var last = v[0];
      var next = v[1];
      var slo = v[2];
      var overdue = undefined;
      var notes = v[4];

      // Clean up next, so that it displays positive values rather than negative.
      if (next != undefined && next < 0) {
        next = Math.abs(next);
      }

      // If we don't have useful variables to calculate how far out of SLO we are,
      // try to show something reasonable.
      if (!slo) {
        overdue = "";
      } else {
        if (!last) {
          overdue = slo;
        } else {
          overdue = last - slo;
        }

        // If it's less than zero, we're totally fine, don't show anything.
        if (overdue <= 0) {
          overdue = "";
        }
      }

      f.push([guest, last, next, slo, overdue, notes]);
    }
    return f;
  }

  // Given an event with two guests, returns the guest email that isn't OWNER_EMAIL
  // Prints an error and returns null on lists that don't contain two entries.
  _getOneOnOneGuestEmail(event) {
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
  _cleanGuestEmail(email) {
    return email.replace('@' + OWNER_DOMAIN, '');
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

// End OneOnOneListCollector
/////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////
//
// MeetingStatsCollector
//
// This class contains all of the knowledge for parsing meetings and tracking
// some basic statistics about them.
class MeetingStatsCollector {

  // Constructs the MeetingStatsCollector.
  //
  // Arguments:
  //   sheet: Spreadsheet to put the results into. Stored in the class.
  constructor(sheet) {
    this.sheet = sheet;

  /*
   const MEETING_STATS_HDRS = [["Week",
                                "1:1s",
                                "Meetings",
                                "Blocked",
                                "Unscheduled"]];
                             */
    // Map of 'Meeting type' => count
    this.meetings = {};

    // Map of 'Meeting week' => {'Meeting type' => hours}
    // Note the special "Unscheduled" meeting type, which starts out with the time available
    // in the whole week.
  }

  // Given a CalendarEvent, will extract basic information about what type of event it
  // is and store the statistics in this class.
  //
  // Arguments:
  //   event: A Calendar Event
  trackEvent(event) {
    var tag = this._extractTag(event);
    var guests = event.getGuestList(true);

    var eventWeek = this._extractWeek(event);
    this._ensureWeekIsInitialized(eventWeek);

    if (tag != null) {
      this._recordHours(event, eventWeek, tag);
    } else if (guests.length == 0) {
      this._recordHours(event, eventWeek, "Blocked Time");
    } else if (guests.length == 1 && guests[0].getEmail() == OWNER_EMAIL) {
      this._recordHours(event, eventWeek, "Blocked Time");
    } else if (guests.length == 2) {
      this._recordHours(event, eventWeek, "1:1s");
    } else {
      this._recordHours(event, eventWeek, "Meetings");
    }
  }

  // Records the hours for a given meeting, of a given type, while subtracting from Unscheduled time.
  //
  // Arguments:
  //   event: A CalendarEvent
  //   eventWeek: a String representing the Sunday of the week of the event, in YYYY-MM-DD format.
  //   tag: the tag to associate with this event.
  _recordHours(event, eventWeek, tag) {
    if (event.getStartTime().getHours() < START_HOUR) {
      //Logger.log('Skipping event that starts before hour ' + START_HOUR + '. It starts at ' + event.getStartTime().getHours());
      return;
    } else if (event.getEndTime().getHours() > END_HOUR) {
      //Logger.log('Skipping event that ends after hour ' + END_HOUR + '. It ends at ' + event.getEndTime().getHours());
      return;
    }

    const hours = (event.getEndTime() - event.getStartTime()) / 1000 / 60 / 60;
    this._inc(this.meetings[eventWeek], tag, hours);
    this._inc(this.meetings[eventWeek], UNSCHEDULED_TIME, -1 * hours);
  }

  // Given an event, returns the date of the Sunday of the week in the format "YYYY-MM-DD"
  _extractWeek(event) {
    // Subtract as many days as needed from the event date to get to Sunday to create a new date object
    // from that, and then convert that to a string.
    const week = new Date(event.getStartTime() - event.getStartTime().getDay() * 24 * 60 * 60 * 1000);
    const weekstr = week.getFullYear() + "-" + (week.getMonth() + 1) + "-" + week.getDate()
    return weekstr;
  }

  // Ensures that we have a structure populated for tracking the stats for the given week.
  // Note that this also creates a basic starting count for "Unscheduled" time, calculated as
  // the amount of time between your start and end hour each day.
  //
  // Arguments:
  //   weekStr: A string representing the week in "YYYY-MM-DD" format.
  _ensureWeekIsInitialized(weekStr) {
    if (!(weekStr in this.meetings)) {
      this.meetings[weekStr] = {};
      this.meetings[weekStr][UNSCHEDULED_TIME] = WORK_HOURS_PER_DAY * WORK_DAYS_PER_WEEK;
    }
  }

  // Populates the stored Spreadsheet with the meeting statistics stored in this class.
  //
  //updateStatsSheet(total, oneOnOnes, blockedTime, meetings, taggedMeetings) {
  updateStatsSheet() {
    var stats = [];
    for (const [tag, count] of Object.entries(this.meetings)) {
      stats.push([tag,
                  this.meetings[tag]["1:1s"],
                  this.meetings[tag]["Meetings"],
                  this.meetings[tag]["Blocked Time"],
                  this.meetings[tag][UNSCHEDULED_TIME]]);
    }

    // Set and freeze the column headers
    var r = this.sheet.getRange(MEETING_STATS_HDR_RANGE);
    r.setValues(MEETING_STATS_HDRS);
    r.setFontWeight('bold');
    this.sheet.setFrozenRows(1);

    // Generate the range
    const range = [MEETING_STATS_DATA_RANGE_COLS, stats.length + 1].join("");
    r = this.sheet.getRange(range);
    r.setValues(stats);

    // Sort the data
    r.sort({column: MEETING_STATS_SORT_COLUMN, ascending: true});

    // Resize columns last, to match the data we just added.
    this.sheet.autoResizeColumns(1, MEETING_STATS_HDRS[0].length);

    // Format the data cells
    this.sheet.getRange([MEETING_STATS_DATA_HOUR_RANGE_COLS, stats.length + 1].join("")).setNumberFormat("0.0");
  }

  // Given a CalendarEvent, will read the description and return any of the text
  // on a line after the keyword MEETING_TAG (currently 'TAG: ')
  // e.g. return $1 from "^\w*TAG: (.*)\w*$"
  _extractTag(event) {
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

  // Increments element `tag` in dictionary `dict` by value `value`
  _inc(dict, tag, value) {
    if (tag in dict) {
      dict[tag] = dict[tag] + value;
    } else {
      dict[tag] = value;
    }
  }
}

// End MeetingStatsCollector
/////////////////////////////////////////////////////////////////////////////////

// Adds a custom menu item to run the script
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu('Calendar Script',
             [{name: 'Collect Meeting Stats', functionName: 'CollectMeetingStats'}]);
}

// Collects statistics on 1:1s and general meetings on the user's default calendar.
//
function CollectMeetingStats() {
  var events = CalendarApp.getDefaultCalendar().getEvents(getDateByDays(RANGE_DAYS_PAST * -1),
                                                          getDateByDays(RANGE_DAYS_FUTURE));
  
  var oneOnOneList = new OneOnOneListCollector(getListSheet());
  var stats = new MeetingStatsCollector(getStatsSheet());
  for (const event of events) {
    oneOnOneList.maybeTrackOneOnOne(event);
    stats.trackEvent(event);
  }

  oneOnOneList.updateListSheet();
  stats.updateStatsSheet();
}

// Returns the Spreadsheet object used to store 1:1 lists, and creates it if one
// doesn't exist yet.
function getListSheet() {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONE_ON_ONE_LIST_SHEET);
  if (sheet == null) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(ONE_ON_ONE_LIST_SHEET);
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  return sheet;
}

// Returns the Spreadsheet object used to store meeting stats, and creates it if one
// doesn't exist yet.
function getStatsSheet() {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEETING_STATS_SHEET);
  if (sheet == null) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(MEETING_STATS_SHEET);
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  return sheet;
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
