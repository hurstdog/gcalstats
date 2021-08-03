# gcalstats

This is a simple Apps Script for Spreadsheets and Calendar to pull data from
your Google Calendar and report various statistics on it in the spreadsheet.

This was inspired by a Google Eng Manager who wrote a similar script and shared
it on the `eng-managers@` mailing list in the mid-2010s. Thanks :)

## Setup

  1. [Create a new spreadsheet](https://sheets.new/)
  1. Click `Tools -> Script Editor`
  1. Change the title to whatever you like (I used 'Calendar Stats')
  1. Delete the example function `myFunction` and paste in the contents of `CalendarStats.js`
  1. Edit `OWNER_USERNAME` and `OWNER_DOMAIN` to match the Google account you
     want calendar stats for.
  1. In the function pull-down menu, choose `ListMeetings`, then click `Run`
  1. Authorize the access (if you trust the code, naturally).
  1. Choose your spreadsheet tab again and view the results. They'll be in
     sheets named `1:1 List` and `Meeting Stats` (names chosen in the constants
     section of the script)

That's pretty much it :)

If you run it again, it'll update the fields that already exist in the
spreadsheet instead of stomping on their values.

## Meeting SLOs

The most useful bit for me is the Meeting SLO. This script will analyze all of
the 1:1s in your calendar, put them into a spreadsheet, and sort according to
who you're farthest out of SLO with. Edit the values in the SLO column and
re-run the script for this to work.

The SLO is set in days, so for instance if you want to meet with someone every
30 days, but it's been 90 since you met with them, you're 60 days out of SLO.
This script will sort that person at the top of your list to show that you need
to check in with them.

## Meeting Stats

This is just a rough breakdown of the time spent in various types of meetings.
It can pick out 1:1s, general meetings (more than 2 participants), and supports
custom tags. Put the tag 'TAG: blah' in a meeting to get that meeting counted
as 'blah' instead of 'Meetings'.
