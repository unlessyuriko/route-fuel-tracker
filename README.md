# Route &amp; Fuel Tracker

A single-page web app that imports field-visit data from Excel, calculates the
**real road distance** each user travelled per day per territory using
[OSRM](http://project-osrm.org/), visualizes the daily route on a map, and
compares the **calculated** distance &amp; fuel against the **actual** values from
telematics and fuel records.

The core question it answers: *did the vehicle actually travel the distance the
visit plan implies — and did it burn the fuel that distance should consume?*

## Features

1. **Import** — drag &amp; drop an `.xlsx` / `.xls` / `.csv`. Flexible header
   matching, coordinate parsing from the `[lon, lat]` format, and Excel
   date/time handling.
2. **Distance** — stops are ordered by `visited_day` timestamp, grouped per
   `territory_id` per day. Each leg is the real driving distance from OSRM; the
   daily total is the sum. Per-leg detail (distance + drive time) is shown too.
3. **Route map** — Leaflet map drawing the actual OSRM road geometry with
   numbered, time-stamped stop markers per day.
4. **Fuel check** — per user, enter car number, engine power, fuel economy
   (km/L), telematics distance and monthly fuel. The app computes:
   - **Calculated distance** = Σ OSRM daily distances
   - **Calculated fuel** = calculated distance ÷ fuel economy
   - Distance Δ and fuel Δ (actual vs. calculated, absolute + %)
   - A status flag: **On plan** (≤15%), **Review** (≤35%), **Anomaly** (>35%)

## Expected columns

| Column | Example |
|---|---|
| `visited_day` | `2026-06-15 08:10` (date **and** time) |
| `outlet_id` | `OUT-1001` |
| `outlet_name` | `City Mart Junction` |
| `location` | `[96.06444, 16.86281]`  (lon, lat) |
| `username` | `Aung` |
| `email` | `aung@demo.com` |
| `territory_id` | `T-01` |

Header names are matched loosely (case/spacing/underscores ignored, plus common
aliases). Use **⬇ Excel Template** in the header to download a starter file, or
**Load Sample** to try the app instantly.

## Running

No build step. Either:

- **Just open** `index.html` in a modern browser, **or**
- Serve the folder (recommended, avoids any `file://` quirks):

  ```bash
  # Python
  python -m http.server 8080
  # then open http://localhost:8080
  ```

An internet connection is required for OSRM routing, OpenStreetMap tiles, and
the CDN libraries (SheetJS, Leaflet).

## Notes

- Routing uses the public OSRM demo server (`router.project-osrm.org`), which is
  rate-limited. Results are cached per coordinate pair. If a leg fails, the app
  falls back to a straight-line (haversine) estimate, marks it with `*` in the
  leg table and a dashed amber line on the map.
- The `driving` profile is used for all routes.

## Tech

Vanilla JS · [SheetJS](https://sheetjs.com/) · [Leaflet](https://leafletjs.com/)
· [OSRM](http://project-osrm.org/) · OpenStreetMap. Heineken-green accented UI.
