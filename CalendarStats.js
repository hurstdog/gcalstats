'use strict';

// Calendar Stats
// This Apps Script uses the Sheets and Calendar APIs to generate statistics
// about the owners calendar, and throws them into a handy spreadsheet.

/////////////////////////////////////////////////////////////////////////////////
// Constants.
//
// Edit the entries in the following section to tailor the script
// to your preferences.

var OWNER_USERNAME = 'andrew';
var OWNER_DOMAIN = 'hurstdog.org';
var OWNER_EMAIL = OWNER_USERNAME + '@' + OWNER_DOMAIN;
var MEETING_TAG = 'TAG: ';

// Counting days from today, forward or backwards.
// Note that both of these values should be positive
var RANGE_DAYS_PAST = 30;
var RANGE_DAYS_FUTURE = 30;

var range = 'A1:B10';

var outputCell = 'C2';

// End Constants. Below is just code, and bad code at that. Ignore it.
/////////////////////////////////////////////////////////////////////////////////


/**
 * Adds a custom menu item to run the script
 */
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
  var tags = {};
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
