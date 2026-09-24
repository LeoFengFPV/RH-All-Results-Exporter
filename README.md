<h1 align="center">All Results Exporter for RotorHazard</h1>

<p align="center">
  <b>Export every saved race, pilot and lap time to one formatted Excel workbook.</b><br>
  <a href="./README_CN.md">[🇨🇳 中文文档]</a>
</p>

<p align="center">
  <a href="https://github.com/LeoFengFPV/RH-All-Results-Exporter/actions/workflows/rhfest.yml">
    <img src="https://github.com/LeoFengFPV/RH-All-Results-Exporter/actions/workflows/rhfest.yml/badge.svg" alt="RHFest">
  </a>
  <a href="https://github.com/LeoFengFPV/RH-All-Results-Exporter/releases">
    <img src="https://img.shields.io/github/v/release/LeoFengFPV/RH-All-Results-Exporter" alt="Release">
  </a>
  <a href="./LICENSE">
    <img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT">
  </a>
  <img src="https://img.shields.io/badge/RotorHazard-4.4.0-blue" alt="Tested on RotorHazard 4.4.0">
  <img src="https://img.shields.io/badge/RHAPI-1.4-blue" alt="Tested on RHAPI 1.4">
</p>

---

## 1. Overview

**All Results Exporter** is a RotorHazard plugin that collects **every saved
race** in an event and writes them into a single Excel (`.xlsx`) workbook, with
one row per pilot per race and every lap time expanded into its own column. It
is designed for the end-of-event review, when the race director needs one
human-readable sheet instead of many JSON files.

## 2. Why this plugin exists

RotorHazard's built-in export produces JSON, which is complete but not easy for
everyone to open and read. Some community exporters only export the **last**
race, forcing the operator to export after every single race. This plugin
fills the gap: **one click exports the whole event at once**, in a formatted
spreadsheet, including per-lap times and the frequency used during each race.

## 3. Features

- One-click export of **all saved races** to a single worksheet.
- One row for every pilot who actually raced.
- Per-lap times expanded across columns (Lap 0, Lap 1, Lap 2, …).
- Race-time **frequency** read from the saved record (one frequency per pilot).
- Fully empty columns are hidden automatically.
- Class, heat and round cells merge only within each real race boundary.
- Unassigned empty seats produce no result rows.
- Names starting with `=` are stored as text, not Excel formulas.
- Runs on demand — there is no background polling or scheduled task.

## 4. Output columns

The worksheet contains these columns (lap columns are added as needed, and the
total time is always the last column):

| Column | Meaning |
|---|---|
| Class | Race class name |
| Heat | Heat / group name |
| Round | Round number |
| Position | Finishing position |
| Pilot | Pilot callsign |
| Team | Team name (if any) |
| Frequency (MHz) | Frequency saved for that pilot in that race |
| Laps | Number of counted laps |
| Average Lap | Average lap time |
| Fastest Lap | Fastest lap time |
| Lap 0 … Lap N | Each individual lap time |
| Total Time | Overall time (always the last column) |

Lap numbering and the total-time field follow the race start format:

- **Hole Shot** and **Staggered Start** number laps from **Lap 0** (Lap 0 is the
  start/holeshot lap).
- **First Lap** numbers laps from **Lap 1** (there is no Lap 0).
- For **Staggered Start**, the total time uses `total_time_laps`, which excludes
  the start lap.

A column that is empty for every row (for example Class when it is never used)
is not exported. Class, Heat and Round are merged only inside the boundary of a
single real race and never across different races.

## 5. Installation

1. Open RotorHazard and go to **Plugins → Upload**.
2. Select the release asset `all_results_exporter_v1.1.1.zip` and upload it.
3. Restart RotorHazard when prompted.
4. The exporter appears under **Data Management → Exporter** as
   **Export All Results (XLSX)**.

The plugin requires [openpyxl](https://pypi.org/project/openpyxl/). Manual
install is also possible: place the `all_results_exporter` folder under
`~/rh-data/plugins/` and restart.

## 6. Usage

1. Fly and **save** your races as usual.
2. When the event (or a session) is finished, open
   **Data Management → Exporter**.
3. Choose **Export All Results (XLSX)** and download the file.
4. Open it in Excel, WPS or any spreadsheet application to review.

Only **saved races** are included. Races that were never saved, or live lap data
that was discarded, are out of scope.

## 7. Compatibility

- Tested on **RotorHazard v4.4.0** with **RHAPI 1.4**.
- The manifest therefore declares a minimum RHAPI version of **1.4**, because no
  earlier version was tested.
- A standard installation includes Python 3 and the `openpyxl` dependency.

## 8. Verified test results

Real-hardware verification (Raspberry Pi):

- RotorHazard v4.4.0 / RHAPI 1.4
- Raspberry Pi 4 Model B Rev 1.5, Python 3.13.5, openpyxl 3.1.5
- Real event: **10 saved races, 30 pilot rows**
- Export took about **0.2 seconds**, page stayed responsive
- Hole Shot results were checked cell-by-cell against real hardware races

Automated regression tests (`tests/`), covering duplicate callsigns, empty
seats, placeholders, deleted laps, all three start formats, merge boundaries,
formula-like names and workbook re-opening, all pass.

## 9. Known limitations

- **First Lap** and **Staggered Start** were verified using RotorHazard's real
  scoring code, but the lap times came from temporarily injected data rather
  than real flights. When you first use these formats in a real event, please
  double-check the starting lap number and the total-time column.
- In the original test environment the web **"plugin upload" button was not
  clicked directly**; an equivalent command-line install plus a real export were
  used to verify loading and output.
- Unassigned seats (a module frequency with no pilot and no laps) are omitted;
  this is intentional for a human-readable result sheet, even though the native
  JSON backup keeps those entries.

## 10. Performance notes

- Real event scale (10 races / 30 rows): roughly **0.2 s**, negligible memory
  increase, normal page response.
- A simulated stress test of about **4,000 rows** (mock data, **not** a real
  event) took about **19.6 s**. For large events, export during a non-timing
  period.
- The plugin does not claim zero resource use; time and memory grow with the
  number of saved races and laps.

## 11. Troubleshooting

- **No download appears:** there may be no saved races, or an internal error was
  logged. Check the RotorHazard log for details.
- **A pilot is missing:** confirm the race was saved and that a pilot was
  assigned to that seat; empty seats are skipped by design.
- **A lap is missing:** deleted laps are not exported. Re-check saved races if a
  lap was removed or added manually.
- **File will not open:** ensure the download completed and openpyxl is
  installed on the timer.

## 12. Privacy

The repository and release contain no real race databases, customer details,
real pilot names, local IP addresses, logs, credentials, tokens or cookies.
Sample data and screenshots use anonymous pilots such as **Pilot A** and
**Pilot B**. Release packages are cleaned of caches and metadata before
uploading.

## 13. License

Released under the [MIT License](./LICENSE).

## 14. Contributing

Issues and pull requests are welcome. When filing a bug report, please include
the RotorHazard/RHAPI versions and remove any private information from logs.
