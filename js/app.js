/* ==========================================================================
   Route & Fuel Tracker
   - Imports Excel/CSV visit data
   - Calculates real road distances per territory per day via OSRM
   - Visualizes daily routes on a Leaflet map
   - Compares calculated vs. actual distance & fuel per user
   ========================================================================== */

"use strict";

const OSRM_BASE = "https://router.project-osrm.org/route/v1/driving";

/* ---- application state ---- */
const state = {
  rows: [],        // parsed visit rows
  trips: [],       // computed { key, territoryId, username, email, date, stops[], legs[], totalKm }
  routed: false,
};

/* ---------- small DOM helpers ---------- */
const $ = (sel) => document.querySelector(sel);
const el = (tag, cls, html) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (html != null) n.innerHTML = html;
  return n;
};
const fmt = (n, d = 2) => (n == null || isNaN(n) ? "—" : Number(n).toFixed(d));

function toast(msg, kind = "") {
  const t = $("#toast");
  t.textContent = msg;
  t.className = "toast " + kind;
  setTimeout(() => t.classList.add("hidden"), 3600);
}
function overlay(show, text = "Working…") {
  $("#overlayText").textContent = text;
  $("#overlay").classList.toggle("hidden", !show);
}

/* ---------- tab navigation ---------- */
$("#tabs").addEventListener("click", (e) => {
  const tab = e.target.closest(".tab");
  if (!tab) return;
  document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
  tab.classList.add("active");
  $("#tab-" + tab.dataset.tab).classList.add("active");
  if (tab.dataset.tab === "map" && map) setTimeout(() => map.invalidateSize(), 60);
});

/* ==========================================================================
   1. PARSING
   ========================================================================== */

/** Pull two floats out of values like "[96.06444, 16.86281]" -> {lon, lat} */
function parseLocation(val) {
  if (val == null) return null;
  const nums = String(val).match(/-?\d+(\.\d+)?/g);
  if (!nums || nums.length < 2) return null;
  // Input order is [lon, lat] (e.g. 96.x = longitude in Myanmar, 16.x = latitude)
  return { lon: parseFloat(nums[0]), lat: parseFloat(nums[1]) };
}

/** Normalize a header into a known field name */
function normKey(h) {
  const k = String(h).toLowerCase().replace(/[\s_\-]+/g, "");
  const map = {
    visitedday: "visitedDay", visiteddate: "visitedDay", visitday: "visitedDay", date: "visitedDay", visitedat: "visitedDay",
    outletid: "outletId", outletcode: "outletId",
    outletname: "outletName", outlet: "outletName",
    location: "location", coordinates: "location", coords: "location", latlon: "location", geo: "location",
    username: "username", user: "username", name: "username",
    email: "email", mail: "email",
    territoryid: "territoryId", territory: "territoryId", terr: "territoryId",
  };
  return map[k] || null;
}

function toDate(v) {
  if (v instanceof Date) return v;
  if (typeof v === "number") {
    // Excel serial date fallback
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d) ? null : d;
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function parseWorkbook(data) {
  const wb = XLSX.read(data, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false, cellDates: true });
  if (!raw.length) throw new Error("No rows found in the first sheet.");

  // build a header map from the first row's keys
  const sample = raw[0];
  const colMap = {};
  Object.keys(sample).forEach((h) => {
    const f = normKey(h);
    if (f) colMap[h] = f;
  });
  const required = ["visitedDay", "location", "territoryId"];
  const haveFields = new Set(Object.values(colMap));
  const missing = required.filter((r) => !haveFields.has(r));
  if (missing.length) {
    throw new Error("Missing required column(s): " + missing.join(", ") +
      ". Found: " + Object.keys(sample).join(", "));
  }

  const rows = [];
  raw.forEach((r, i) => {
    const o = { _row: i + 2 };
    Object.keys(colMap).forEach((h) => (o[colMap[h]] = r[h]));
    const loc = parseLocation(o.location);
    const when = toDate(o.visitedDay);
    if (!loc || !when) return; // skip unusable rows
    rows.push({
      visitedDay: when,
      outletId: String(o.outletId ?? "").trim(),
      outletName: String(o.outletName ?? "").trim(),
      lon: loc.lon, lat: loc.lat,
      username: String(o.username ?? "").trim() || "(unknown)",
      email: String(o.email ?? "").trim(),
      territoryId: String(o.territoryId ?? "").trim() || "(none)",
    });
  });
  if (!rows.length) throw new Error("No rows had a valid location + visited day.");
  return rows;
}

/* ---------- file handling ---------- */
function handleFile(file) {
  overlay(true, "Reading " + file.name + "…");
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const rows = parseWorkbook(new Uint8Array(e.target.result));
      onDataLoaded(rows, file.name);
    } catch (err) {
      console.error(err);
      toast(err.message, "bad");
    } finally {
      overlay(false);
    }
  };
  reader.onerror = () => { overlay(false); toast("Could not read file.", "bad"); };
  reader.readAsArrayBuffer(file);
}

function onDataLoaded(rows, sourceName) {
  state.rows = rows;
  state.trips = [];
  state.routed = false;
  buildTrips();
  renderImportSummary(sourceName);
  renderRawTable();
  renderDayTable();
  renderFuelTable();
  $("#btnCalc").disabled = false;
  toast("Loaded " + rows.length + " rows.", "ok");
}

/* ---------- dropzone wiring ---------- */
const dz = $("#dropZone");
$("#fileInput").addEventListener("change", (e) => e.target.files[0] && handleFile(e.target.files[0]));
["dragover", "dragenter"].forEach((ev) =>
  dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
["dragleave", "drop"].forEach((ev) =>
  dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
dz.addEventListener("drop", (e) => e.dataTransfer.files[0] && handleFile(e.dataTransfer.files[0]));

/* ==========================================================================
   2. TRIP GROUPING & DISTANCE
   ========================================================================== */

const dayKey = (d) => d.toISOString().slice(0, 10);

function buildTrips() {
  const groups = new Map();
  state.rows.forEach((r) => {
    const key = r.territoryId + "||" + dayKey(r.visitedDay);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });

  const trips = [];
  groups.forEach((stops, key) => {
    stops.sort((a, b) => a.visitedDay - b.visitedDay);
    const [territoryId, date] = key.split("||");
    trips.push({
      key, territoryId, date,
      username: stops[0].username,
      email: stops[0].email,
      stops,
      legs: [],
      totalKm: null,
    });
  });
  trips.sort((a, b) => (a.territoryId + a.date).localeCompare(b.territoryId + b.date));
  state.trips = trips;
}

/* ---- OSRM with caching + limited concurrency ---- */
const routeCache = new Map();

async function osrmRoute(from, to) {
  const ck = `${from.lon},${from.lat};${to.lon},${to.lat}`;
  if (routeCache.has(ck)) return routeCache.get(ck);
  const url = `${OSRM_BASE}/${ck}?overview=full&geometries=geojson`;
  let lastErr;
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error("OSRM HTTP " + res.status);
      const json = await res.json();
      if (json.code !== "Ok" || !json.routes?.length) throw new Error("OSRM: " + json.code);
      const route = json.routes[0];
      const out = {
        km: route.distance / 1000,
        durationMin: route.duration / 60,
        geometry: route.geometry.coordinates, // [lon,lat][]
      };
      routeCache.set(ck, out);
      return out;
    } catch (err) {
      lastErr = err;
      await new Promise((r) => setTimeout(r, 400));
    }
  }
  // Fallback: straight-line (haversine) so one failure doesn't break the run
  const km = haversine(from, to);
  const out = { km, durationMin: null, geometry: [[from.lon, from.lat], [to.lon, to.lat]], fallback: true };
  routeCache.set(ck, out);
  console.warn("OSRM failed, used straight-line:", lastErr?.message);
  return out;
}

function haversine(a, b) {
  const R = 6371, toRad = (x) => (x * Math.PI) / 180;
  const dLat = toRad(b.lat - a.lat), dLon = toRad(b.lon - a.lon);
  const s = Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(s));
}

async function calculateAll() {
  // count legs for progress
  const legCount = state.trips.reduce((s, t) => s + Math.max(0, t.stops.length - 1), 0);
  if (!legCount) { toast("No multi-stop days to route.", "warn"); return; }

  $("#btnCalc").disabled = true;
  const prog = $("#calcProgress");
  prog.classList.remove("hidden");
  let done = 0;
  let usedFallback = false;
  const setProg = () => {
    const pct = Math.round((done / legCount) * 100);
    prog.querySelector(".bar").style.setProperty("--p", pct + "%");
    prog.querySelector(".label").textContent = `Routing ${done}/${legCount} legs…`;
  };
  setProg();

  for (const trip of state.trips) {
    trip.legs = [];
    let total = 0;
    for (let i = 0; i < trip.stops.length - 1; i++) {
      const a = trip.stops[i], b = trip.stops[i + 1];
      const r = await osrmRoute(a, b);
      if (r.fallback) usedFallback = true;
      trip.legs.push({ from: a, to: b, ...r });
      total += r.km;
      done++; setProg();
    }
    trip.totalKm = total;
  }

  state.routed = true;
  prog.classList.add("hidden");
  $("#btnCalc").disabled = false;
  renderDayTable();
  renderLegTable();
  renderFuelTable();
  populateMapSelect();
  toast(usedFallback
    ? "Done — some legs used straight-line fallback (OSRM unavailable)."
    : "Routing complete.", usedFallback ? "warn" : "ok");
}

$("#btnCalc").addEventListener("click", () => calculateAll());

/* ==========================================================================
   3. RENDERING — tables
   ========================================================================== */

function renderImportSummary(sourceName) {
  const box = $("#importSummary");
  const users = new Set(state.rows.map((r) => r.username));
  const terrs = new Set(state.rows.map((r) => r.territoryId));
  const days = new Set(state.rows.map((r) => dayKey(r.visitedDay)));
  box.innerHTML = "";
  const stats = [
    ["Source", sourceName, true],
    ["Rows", state.rows.length],
    ["Users", users.size],
    ["Territories", terrs.size],
    ["Days", days.size],
    ["Trips (terr×day)", state.trips.length],
  ];
  stats.forEach(([k, v, isText]) => {
    const s = el("div", "stat");
    s.innerHTML = `<div class="n" style="${isText ? "font-size:14px;word-break:break-all" : ""}">${v}</div><div class="k">${k}</div>`;
    box.appendChild(s);
  });
  box.classList.remove("hidden");
}

function renderRawTable() {
  const t = $("#rawTable");
  const head = ["Visited day", "Territory", "User", "Outlet ID", "Outlet name", "Lon", "Lat"];
  let html = "<thead><tr>" + head.map((h) => `<th>${h}</th>`).join("") + "</tr></thead><tbody>";
  state.rows.slice(0, 300).forEach((r) => {
    html += `<tr>
      <td>${r.visitedDay.toLocaleString()}</td>
      <td>${r.territoryId}</td>
      <td>${r.username}</td>
      <td>${r.outletId}</td>
      <td>${r.outletName}</td>
      <td class="num">${fmt(r.lon, 5)}</td>
      <td class="num">${fmt(r.lat, 5)}</td>
    </tr>`;
  });
  html += "</tbody>";
  t.innerHTML = html;
  $("#rawCard").classList.toggle("hidden", state.rows.length === 0);
  if (state.rows.length > 300) {
    $("#rawCard").querySelector("h2").textContent = `Parsed rows (showing 300 of ${state.rows.length})`;
  }
}

function renderDayTable() {
  const t = $("#dayTable");
  const head = ["Territory", "Date", "User", "Stops", "Total distance (km)", ""];
  let html = "<thead><tr>" + head.map((h, i) => `<th class="${i >= 3 ? "num" : ""}">${h}</th>`).join("") + "</tr></thead><tbody>";
  state.trips.forEach((tr) => {
    const total = tr.totalKm == null ? "—" : fmt(tr.totalKm);
    html += `<tr>
      <td>${tr.territoryId}</td>
      <td>${tr.date}</td>
      <td>${tr.username}</td>
      <td class="num">${tr.stops.length}</td>
      <td class="num">${total}</td>
      <td>${state.routed && tr.stops.length > 1 ? `<span class="row-link" data-trip="${tr.key}">view map →</span>` : ""}</td>
    </tr>`;
  });
  html += "</tbody>";
  t.innerHTML = html;
  t.querySelectorAll(".row-link").forEach((lnk) =>
    lnk.addEventListener("click", () => openTripOnMap(lnk.dataset.trip)));
}

function renderLegTable() {
  const t = $("#legTable");
  let html = `<thead><tr>
    <th>Territory</th><th>Date</th><th>#</th><th>From</th><th>To</th>
    <th class="num">Leg km</th><th class="num">Drive min</th></tr></thead><tbody>`;
  state.trips.forEach((tr) => {
    tr.legs.forEach((leg, i) => {
      html += `<tr>
        <td>${tr.territoryId}</td>
        <td>${tr.date}</td>
        <td class="num">${i + 1}</td>
        <td>${leg.from.outletName || leg.from.outletId}</td>
        <td>${leg.to.outletName || leg.to.outletId}</td>
        <td class="num">${fmt(leg.km)}${leg.fallback ? " *" : ""}</td>
        <td class="num">${leg.durationMin == null ? "—" : fmt(leg.durationMin, 0)}</td>
      </tr>`;
    });
  });
  html += "</tbody>";
  t.innerHTML = html;
  $("#legCard").classList.toggle("hidden", !state.routed);
}

/* ==========================================================================
   4. FUEL CHECK  (per user)
   ========================================================================== */

const fuelInputs = {}; // username -> {carNo, enginePower, economy, actualKm, actualFuel}

function userCalcKm(username) {
  return state.trips
    .filter((t) => t.username === username && t.totalKm != null)
    .reduce((s, t) => s + t.totalKm, 0);
}

function renderFuelTable() {
  const users = [...new Set(state.rows.map((r) => r.username))].sort();
  const t = $("#fuelTable");
  let html = `<thead><tr>
    <th>User</th><th>Car no.</th><th>Engine (hp)</th><th>Fuel econ (km/L)</th>
    <th class="num">Calc dist (km)</th><th class="num">Actual dist (km)</th>
    <th class="num">Calc fuel (L)</th><th class="num">Actual fuel (L)</th>
    <th>Distance Δ</th><th>Fuel Δ</th><th>Status</th>
    </tr></thead><tbody>`;
  users.forEach((u) => {
    const f = fuelInputs[u] || (fuelInputs[u] = { economy: 10 });
    const calcKm = userCalcKm(u);
    html += `<tr data-user="${encodeURIComponent(u)}">
      <td><strong>${u}</strong></td>
      <td><input class="text" data-k="carNo" value="${f.carNo ?? ""}" placeholder="ABC-123" /></td>
      <td><input data-k="enginePower" value="${f.enginePower ?? ""}" placeholder="hp" /></td>
      <td><input data-k="economy" value="${f.economy ?? ""}" placeholder="10" /></td>
      <td class="num calc cell-calcKm">${state.routed ? fmt(calcKm) : "—"}</td>
      <td class="num"><input data-k="actualKm" value="${f.actualKm ?? ""}" placeholder="km" /></td>
      <td class="num calc cell-calcFuel">—</td>
      <td class="num"><input data-k="actualFuel" value="${f.actualFuel ?? ""}" placeholder="L" /></td>
      <td class="cell-distDelta">—</td>
      <td class="cell-fuelDelta">—</td>
      <td class="cell-status">—</td>
    </tr>`;
  });
  html += "</tbody>";
  t.innerHTML = html;

  t.querySelectorAll("tr[data-user]").forEach((tr) => {
    const u = decodeURIComponent(tr.dataset.user);
    tr.querySelectorAll("input").forEach((inp) => {
      inp.addEventListener("input", () => {
        fuelInputs[u][inp.dataset.k] = inp.value;
        recomputeFuelRow(tr, u);
      });
    });
    recomputeFuelRow(tr, u);
  });
}

function recomputeFuelRow(tr, u) {
  const f = fuelInputs[u];
  const calcKm = userCalcKm(u);
  const econ = parseFloat(f.economy);
  const actualKm = parseFloat(f.actualKm);
  const actualFuel = parseFloat(f.actualFuel);

  const calcFuel = econ > 0 && state.routed ? calcKm / econ : null;

  tr.querySelector(".cell-calcKm").textContent = state.routed ? fmt(calcKm) : "—";
  tr.querySelector(".cell-calcFuel").textContent = calcFuel == null ? "—" : fmt(calcFuel);

  // distance delta (actual vs calc)
  const dDelta = tr.querySelector(".cell-distDelta");
  if (state.routed && !isNaN(actualKm) && calcKm > 0) {
    const diff = actualKm - calcKm;
    const pct = (diff / calcKm) * 100;
    dDelta.innerHTML = badge(diff, pct, "km");
  } else dDelta.textContent = "—";

  // fuel delta (actual vs calc)
  const fDelta = tr.querySelector(".cell-fuelDelta");
  if (calcFuel != null && !isNaN(actualFuel) && calcFuel > 0) {
    const diff = actualFuel - calcFuel;
    const pct = (diff / calcFuel) * 100;
    fDelta.innerHTML = badge(diff, pct, "L");
  } else fDelta.textContent = "—";

  // overall status: does telematics distance match the routed plan?
  const status = tr.querySelector(".cell-status");
  if (state.routed && !isNaN(actualKm) && calcKm > 0) {
    const pct = Math.abs((actualKm - calcKm) / calcKm) * 100;
    if (pct <= 15) status.innerHTML = `<span class="badge ok">On plan</span>`;
    else if (pct <= 35) status.innerHTML = `<span class="badge warn">Review</span>`;
    else status.innerHTML = `<span class="badge bad">Anomaly</span>`;
  } else status.textContent = "—";
}

function badge(diff, pct, unit) {
  const sign = diff >= 0 ? "+" : "−";
  const cls = Math.abs(pct) <= 15 ? "ok" : Math.abs(pct) <= 35 ? "warn" : "bad";
  return `<span class="badge ${cls}">${sign}${fmt(Math.abs(diff))} ${unit} (${sign}${fmt(Math.abs(pct), 0)}%)</span>`;
}

/* ==========================================================================
   5. MAP
   ========================================================================== */

let map, routeLayer;

function ensureMap() {
  if (map) return;
  map = L.map("map").setView([16.86, 96.06], 12);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);
  routeLayer = L.layerGroup().addTo(map);
}

function populateMapSelect() {
  const sel = $("#mapTrip");
  sel.innerHTML = "";
  const routable = state.trips.filter((t) => t.stops.length > 1 && t.totalKm != null);
  if (!routable.length) {
    sel.innerHTML = `<option>No routable days</option>`;
    return;
  }
  routable.forEach((t) => {
    const o = el("option");
    o.value = t.key;
    o.textContent = `${t.territoryId} · ${t.date} · ${t.username} (${fmt(t.totalKm)} km)`;
    sel.appendChild(o);
  });
  sel.onchange = () => drawTrip(sel.value);
  drawTrip(routable[0].key);
}

function openTripOnMap(key) {
  // switch to map tab and select
  document.querySelector('.tab[data-tab="map"]').click();
  const sel = $("#mapTrip");
  sel.value = key;
  drawTrip(key);
}

function drawTrip(key) {
  const trip = state.trips.find((t) => t.key === key);
  if (!trip) return;
  ensureMap();
  setTimeout(() => map.invalidateSize(), 50);
  routeLayer.clearLayers();

  const allLatLng = [];

  // route polyline from concatenated leg geometries
  trip.legs.forEach((leg) => {
    const latlngs = leg.geometry.map(([lon, lat]) => [lat, lon]);
    L.polyline(latlngs, {
      color: leg.fallback ? "#e3a008" : "#007a33",
      weight: 5, opacity: .85,
      dashArray: leg.fallback ? "8 8" : null,
    }).addTo(routeLayer);
    allLatLng.push(...latlngs);
  });

  // numbered stop markers
  trip.stops.forEach((s, i) => {
    const icon = L.divIcon({
      className: "",
      html: `<div class="leaflet-marker-num">${i + 1}</div>`,
      iconSize: [26, 26], iconAnchor: [13, 13],
    });
    L.marker([s.lat, s.lon], { icon })
      .bindPopup(`<b>${i + 1}. ${s.outletName || s.outletId}</b><br>${s.visitedDay.toLocaleString()}<br>${s.territoryId}`)
      .addTo(routeLayer);
    allLatLng.push([s.lat, s.lon]);
  });

  if (allLatLng.length) map.fitBounds(L.latLngBounds(allLatLng).pad(0.15));

  const fb = trip.legs.some((l) => l.fallback);
  $("#mapMeta").innerHTML =
    `<b>${trip.territoryId}</b> · ${trip.date} · ${trip.username} — ${trip.stops.length} stops, ` +
    `<b>${fmt(trip.totalKm)} km</b>${fb ? ' · <span style="color:#8a6300">dashed = straight-line fallback</span>' : ""}`;
}

/* ==========================================================================
   6. TEMPLATE + SAMPLE DATA
   ========================================================================== */

const SAMPLE = [
  // territory T-01, user Aung — 2026-06-15
  ["2026-06-15 08:10", "OUT-1001", "City Mart Junction", "[96.1561, 16.8050]", "Aung", "aung@demo.com", "T-01"],
  ["2026-06-15 09:25", "OUT-1002", "Sky Bar Yangon", "[96.1402, 16.7969]", "Aung", "aung@demo.com", "T-01"],
  ["2026-06-15 11:00", "OUT-1003", "Ocean Supercenter", "[96.1739, 16.8409]", "Aung", "aung@demo.com", "T-01"],
  ["2026-06-15 13:40", "OUT-1004", "Golden Valley Mart", "[96.1490, 16.8190]", "Aung", "aung@demo.com", "T-01"],
  // territory T-01, user Aung — 2026-06-16
  ["2026-06-16 08:30", "OUT-1005", "Hledan Center", "[96.1300, 16.8290]", "Aung", "aung@demo.com", "T-01"],
  ["2026-06-16 10:15", "OUT-1006", "Junction Square", "[96.1352, 16.8268]", "Aung", "aung@demo.com", "T-01"],
  ["2026-06-16 12:05", "OUT-1007", "Myaynigone Plaza", "[96.1394, 16.8136]", "Aung", "aung@demo.com", "T-01"],
  // territory T-02, user Su — 2026-06-15
  ["2026-06-15 08:05", "OUT-2001", "Mandalay City Mart", "[96.0894, 21.9750]", "Su", "su@demo.com", "T-02"],
  ["2026-06-15 09:50", "OUT-2002", "Zay Cho Market", "[96.0852, 21.9810]", "Su", "su@demo.com", "T-02"],
  ["2026-06-15 11:30", "OUT-2003", "Diamond Plaza MDY", "[96.1010, 21.9640]", "Su", "su@demo.com", "T-02"],
  ["2026-06-16 08:20", "OUT-2004", "Ocean MDY", "[96.0950, 21.9900]", "Su", "su@demo.com", "T-02"],
  ["2026-06-16 10:40", "OUT-2005", "Mingalar Market", "[96.0790, 21.9700]", "Su", "su@demo.com", "T-02"],
];
const SAMPLE_HEAD = ["visited_day", "outlet_id", "outlet_name", "location", "username", "email", "territory_id"];

function buildSampleSheet() {
  const aoa = [SAMPLE_HEAD, ...SAMPLE];
  return XLSX.utils.aoa_to_sheet(aoa);
}

$("#btnTemplate").addEventListener("click", () => {
  const ws = XLSX.utils.aoa_to_sheet([SAMPLE_HEAD, SAMPLE[0]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "visits");
  XLSX.writeFile(wb, "route_fuel_template.xlsx");
});

$("#btnSample").addEventListener("click", () => {
  try {
    onDataLoaded(parseWorkbook(buildSampleArrayBuffer()), "sample-data.xlsx");
  } catch (err) {
    console.error(err);
    toast(err.message, "bad");
  }
});

function buildSampleArrayBuffer() {
  const ws = buildSampleSheet();
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "visits");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}
