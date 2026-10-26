# LPH Timesheet Report

Weekly hours per employee (Sunday–Saturday) from ServiceTitan's scheduled "Timesheet Report", with department filters and a "still clocked in" flag.

Live page: https://mattsbaker1980-dev.github.io/LPH-Timesheet-Report/

## How it updates
A scheduled Claude task runs daily at 8:15, 10:15, 12:15, 2:15, 4:15, 6:15 and 10:15 PM Eastern (15 minutes after each ServiceTitan "Today" run):
1. Finds the "Employee Time" emails from ServiceTitan in Gmail and downloads the .xlsx attachments.
2. `sync/ts_ingest.py` turns each report into a snapshot (one day, one run time).
3. `sync/ts_publish.js` adds new snapshots to `data/weeks/<sunday>.json` and updates `data/index.json`.

Weeks run Monday–Sunday to match payroll. Each day shows its newest report (the next-morning "Yesterday" run is final).

## Importing the Technician Timesheet Summary
Run the legacy Technician Timesheet Summary for all employees in ServiceTitan, export to Excel, and click **Import timesheet** on the page. It adds start/end times, lunch, PTO (Vac), Non-PTO and Holiday hours for that week, and flags anyone whose End matches the export time as still clocked in. The first import on a computer asks for a fine-grained GitHub token with Contents: Read and write on this repo only; it is kept in that browser (localStorage), never in the code.
