# LPH Timesheet Report

Weekly hours per employee (Sunday–Saturday) from ServiceTitan's scheduled "Timesheet Report", with department filters and a "still clocked in" flag.

Live page: https://mattsbaker1980-dev.github.io/LPH-Timesheet-Report/

## How it updates
A scheduled Claude task runs at 8:15 AM and 2:15 PM Eastern:
1. Finds the "Employee Time" emails from ServiceTitan in Gmail and downloads the .xlsx attachments.
2. `sync/ts_ingest.py` turns each report into a snapshot (one day, one run time).
3. `sync/ts_publish.js` adds new snapshots to `data/weeks/<sunday>.json` and updates `data/index.json`.

Each day shows its newest report (the next-morning "Yesterday" run is final). Anyone whose hours went up between two same-day runs (1:45 and 2:00 PM) is flagged as still clocked in.
