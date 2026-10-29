# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/),
and this project follows [Semantic Versioning](https://semver.org/).

## [1.1.1] - 2026-09-23

First public stable release.

### Added

- One-click export of **every saved race** to a single formatted XLSX workbook.
- One row per pilot per race, with per-lap times expanded into columns.
- Race-time **frequency** read from the saved records.
- Automatic hiding of fully empty columns.
- Class / heat / round cells merged only within each real race boundary.
- Unassigned empty-seat filtering.
- Formula-injection protection (callsigns starting with `=` stay as text).
- RHFest validation workflow.
- Tag-triggered release workflow with SHA-256 checksums.
- Automated regression tests under `tests/`.

### Verified

- RotorHazard v4.4.0 / RHAPI 1.4 on a Raspberry Pi 4 Model B.
- Real event: 10 saved races, 30 pilot rows, export ~0.2 s.
