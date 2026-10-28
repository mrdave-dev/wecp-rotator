# WECP Rotator

A Google Sheets script (via clasp) that fills out a weekly station-rotation
grid from a simple, non-technical Config sheet, and rotates rosters so the
same people don't always land on the same station.

Document: https://drive.google.com/open?id=1Vvhs6LTdsmvRbgvtWRHJhenCZE6p8iOFAJzIhlgVOOE
Script: https://script.google.com/d/1CL1SL6NdrfEQqkJlJ9SfI53gsXX7t43GoVFu3kia6_8V6bgakgHCO-tL/edit

## How it works

Open the spreadsheet and use the **WECP** menu:

- **Set Up Workbook** creates a `Config` sheet and a `Schedule` sheet.
  The Config sheet has a start date, a single **Stations** list (defined
  once — the same stations apply to every day), and a single **Roster**
  table headed `Name | Day`: one row per person per workday, e.g.

  ```
  Name     Day
  Adam     M
  Brenda   T
  Charlie  R
  Diana    T
  ```

  Day is one letter — `M`, `T`, `W`, `R`, or `F` for Mon-Fri (`R` is
  Thursday, to avoid clashing with Tuesday's `T`). Someone who works
  multiple days just gets one row per day. A day is only scheduled if at
  least one roster row names it — there's no separate day toggle. Comes
  pre-filled with the example above.
- Fill in the Config sheet: list your stations (in the order they should
  appear on the schedule), set the start date (the first day of the
  upcoming week), and list who works which day in the Roster table.
- **Fill Grid for Upcoming Week** reads the Config sheet and writes the
  `Schedule` grid: one column per included day (in calendar order from the
  start date), one row per station (in the order they're listed in Config),
  and each cell is the person assigned to that station that day, filling
  stations top-to-bottom in the order that day's roster is listed. If a
  day's roster is shorter than the station list, the remaining stations are
  left blank for that day rather than repeating someone; extra people
  beyond the station count just aren't assigned anywhere that day. Only the
  grid itself (title, day headers, and station rows) is touched — any other
  content on the Schedule sheet, like announcements below the grid, is left
  alone.
- **Next Week** moves the whole workbook forward one week in a single
  click: the Config start date advances by 7 days, each day's roster
  rotates (the last person moves to the front, everyone else moves down a
  spot), and the `Schedule` grid is filled for the resulting week.
- **Previous Week** is the mirror of Next Week: the start date moves back 7
  days, each day's roster un-rotates (the first person moves to the end,
  everyone else moves up a spot), and the grid is filled for that week.
- **Add 16 Test Names** is a development/testing helper: appends 16 sample
  names to the Roster table, cycling through M-F, so Fill Grid/Next Week
  can be exercised against a nontrivial roster without typing names in by
  hand. It only adds rows — it never touches what's already there.

Fill Grid, Next Week, and Previous Week all check whether the Schedule grid
was hand-edited since it was last generated, and confirm before
overwriting. They also confirm separately if a larger grid than last time
(a new station or day) would grow into rows or columns that already have
other content on them, since that's outside the area the hand-edit check
watches. For Next Week/Previous Week, these checks (and the date/roster
changes) only happen after you confirm — a "No" or a config problem leaves
the start date and rosters untouched.

## Development

```
npm install
npx tsc        # compiles src/*.ts to dist/*.js
npx clasp push # deploys dist/ to the bound Apps Script project
```

Source files under `src/` have no imports/exports by design — Apps Script
runs every file's top-level declarations in one shared global scope, so
each `.ts` file is its own script file, matching that model.
