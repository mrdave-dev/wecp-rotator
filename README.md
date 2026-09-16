# WECP Rotator

A Google Sheets script (via clasp) that fills out a weekly station-rotation
grid from a simple, non-technical Config sheet, and rotates rosters so the
same people don't always land on the same station.

Document: https://drive.google.com/open?id=1Vvhs6LTdsmvRbgvtWRHJhenCZE6p8iOFAJzIhlgVOOE
Script: https://script.google.com/d/1CL1SL6NdrfEQqkJlJ9SfI53gsXX7t43GoVFu3kia6_8V6bgakgHCO-tL/edit

## How it works

Open the spreadsheet and use the **Rotation** menu:

- **Set Up Workbook** creates a `Config` sheet and a `Schedule` sheet. The
  Config sheet has one block per day of the week: a checkbox to include that
  day, a start date, and two columns to list that day's roster (people) and
  stations. Monday comes pre-filled with an example.
- Fill in the Config sheet: check the days you use, set the start date (the
  first day of the upcoming week), and list each included day's roster and
  stations.
- **Fill Grid for Upcoming Week** reads the Config sheet and writes the
  `Schedule` grid: one column per included day (in calendar order from the
  start date), one row per station, and each cell is the person assigned to
  that station that day. If a day has more people than stations (or vice
  versa), the roster cycles across the stations in order.
- **Rotate Rosters** shifts each included day's roster by one: the last
  person moves to the front and everyone else moves down a spot. This only
  changes the saved Config order — run Fill Grid afterwards to apply it to
  the schedule.
- **Check for Manual Edits** reports whether the Schedule sheet has been
  hand-edited since it was last generated. Fill Grid and Set Up Workbook
  both check this automatically and confirm before overwriting a
  hand-edited schedule.

## Development

```
npm install
npx tsc        # compiles src/*.ts to dist/*.js
npx clasp push # deploys dist/ to the bound Apps Script project
```

Source files under `src/` have no imports/exports by design — Apps Script
runs every file's top-level declarations in one shared global scope, so
each `.ts` file is its own script file, matching that model.
