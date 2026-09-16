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
  once — the same stations apply to every day), and one block per day of
  the week with a checkbox to include that day and a roster of who works
  it. Monday comes pre-filled with an example roster.
- Fill in the Config sheet: list your stations (in the order they should
  appear on the schedule), check the days you use, set the start date (the
  first day of the upcoming week), and list each included day's roster.
- **Fill Grid for Upcoming Week** reads the Config sheet and writes the
  `Schedule` grid: one column per included day (in calendar order from the
  start date), one row per station (in the order they're listed in Config),
  and each cell is the person assigned to that station that day. If a day's
  roster is longer or shorter than the station list, it cycles across the
  stations in order. Only the grid itself (title, day headers, and station
  rows) is touched — any other content on the Schedule sheet, like
  announcements below the grid, is left alone.
- **Rotate & Fill for Upcoming Week** does Rotate Rosters and Fill Grid in
  one click.
- **Rotate Rosters** shifts each included day's roster by one: the last
  person moves to the front and everyone else moves down a spot. This only
  changes the saved Config order — run Fill Grid afterwards to apply it to
  the schedule.
- **Unrotate Rosters** undoes one Rotate Rosters step: the first person
  moves to the end and everyone else moves up a spot.

Fill Grid and Rotate & Fill both check whether the Schedule grid was
hand-edited since it was last generated, and confirm before overwriting.

## Development

```
npm install
npx tsc        # compiles src/*.ts to dist/*.js
npx clasp push # deploys dist/ to the bound Apps Script project
```

Source files under `src/` have no imports/exports by design — Apps Script
runs every file's top-level declarations in one shared global scope, so
each `.ts` file is its own script file, matching that model.
