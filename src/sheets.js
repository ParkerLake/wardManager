import config from "./config";

// ── Worker proxy — all Sheets API calls go through here ──────────────────────
// The Worker holds the service account key securely server-side.
// No Google OAuth token needed from the user at all.

const WORKER_URL = "https://ward-manager-sheets.parkerllake.workers.dev";

const SID  = () => config.SPREADSHEET_ID;
const SSID = () => config.SACRAMENT_SHEET_ID;
const WCID = () => config.WARD_COUNCIL_SHEET_ID;
const PLID = () => config.PRAYER_LIST_SHEET_ID;

async function sheetsReq(sheetId, range, method = "GET", body = null) {
  const params = new URLSearchParams({ spreadsheetId: sheetId, range });
  const url = `${WORKER_URL}/sheets?${params}`;
  const res = await fetch(url, {
    method,
    headers: { "Content-Type": "application/json" },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  if (!res.ok) {
    const t = await res.text();
    console.error(`Worker ${res.status} on ${method} ${range}:`, t);
    throw new Error(`Sheets API ${res.status}: ${t}`);
  }
  return res.json();
}

const bishopricReq = (range, method, body) => sheetsReq(SID(),  range, method, body);
const sacramentReq = (range, method, body) => sheetsReq(SSID(), range, method, body);
const wcReq        = (range, method, body) => sheetsReq(WCID(), range, method, body);

// ── Pad rows so deletions clear trailing rows in the sheet ───────────────────
function padRows(values, minRows = 500) {
  const cols = values[0]?.length || 1;
  const empty = Array(cols).fill("");
  const padded = [...values];
  while (padded.length < minRows) padded.push(empty);
  return padded;
}

async function clearAndWrite(sheetId, range, values) {
  return sheetsReq(sheetId, range, "PUT", { values: padRows(values) });
}

// ── Row converters ────────────────────────────────────────────────────────────

function appointmentsToRows(data) {
  return [
    ["Name", "Status", "Owner", "Purpose", "Appt Date", "Notes"],
    ...data.map(a => [a.name, a.status, a.owner, a.purpose, a.apptDate, a.notes]),
  ];
}

function rowsToAppointments(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map((r, i) => ({
    id: `a_${(r[0]||"").replace(/\s+/g,"_").toLowerCase()}_${(r[3]||"").replace(/\s+/g,"_").toLowerCase()}_${(r[4]||"").replace(/\s+/g,"_").toLowerCase()}_${i}`,
    name: r[0]||"", status: r[1]||"Need to Schedule",
    owner: r[2]||"", purpose: r[3]||"", apptDate: r[4]||"", notes: r[5]||"",
  }));
}

function callingsToRows(data) {
  return [
    ["Name", "Calling", "Stage", "Notes"],
    ...data.map(c => [c.name, c.calling, c.stage, c.notes||""]),
  ];
}

function rowsToCallings(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1)
    .filter(r => r[0] || r[1])
    .map((r, i) => ({
      id: `c_${i}`, name: r[0]||"", calling: r[1]||"",
      stage: r[2]||"Proposed", notes: r[3]||"",
    }));
}

function notesToRows(data) {
  return [
    ["ID", "Title", "Body", "Date", "Author"],
    ...data.map(n => [n.id||"", n.title||"", n.body||"", n.date||"", n.author||""]),
  ];
}

function rowsToNotes(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1)
    .filter(r => r[0] || r[1] || r[2])
    .map((r, i) => ({
      id: r[0] || `n_${i}`, title: r[1]||"",
      body: r[2]||"", date: r[3]||"", author: r[4]||"",
    }));
}

// Legacy aliases — kept so any remaining members refs don't break
const membersToRows = notesToRows;
const rowsToMembers = rowsToNotes;

function rowsToMeeting(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1)
    .filter(r => r[0] || r[1] || r[2])
    .map((r, i) => ({
      id: `mt_${i}`, topic: r[0]||"", owner: r[1]||"",
      status: r[2]||"", notes: r[3]||"",
    }));
}

function meetingToRows(data) {
  return [
    ["Topic", "Owner", "Status", "Notes"],
    ...data.map(m => [m.topic, m.owner, m.status||"", m.notes||""]),
  ];
}

function rowsToLinks(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1)
    .filter(r => r[0] || r[2])
    .map((r, i) => ({
      id: r[0] || `lnk_${i}`, name: r[1]||"", url: r[2]||"", description: r[3]||"",
    }));
}

// ── Pull functions ────────────────────────────────────────────────────────────

export async function pullAll() {
  const [appts, callings, releasings, members, meeting] = await Promise.all([
    bishopricReq("Appointments!A:F").then(r => rowsToAppointments(r.values)),
    bishopricReq("Callings!A:D").then(r => rowsToCallings(r.values)),
    bishopricReq("Releasings!A:D").then(r => rowsToCallings(r.values)),
    bishopricReq("Notes!A:E").then(r => rowsToNotes(r.values)),
    bishopricReq("BishopricMeeting!A:D").then(r => rowsToMeeting(r.values)),
  ]);
  return { appointments: appts, callings, releasings, members, bishopricMeeting: meeting };
}

export async function pullSacramentProgram() {
  const r = await sacramentReq("SacramentProgram!A:F");
  if (!r.values || r.values.length < 2) return { sacramentProgram: [] };
  const hasDate = r.values[0][0] === "Date";
  const rows = r.values.slice(1).filter(row => row.some(c => c)).map((row, i) => {
    if (hasDate) {
      // Format: Date | GlobalOrder | Section | Label | Value | Notes
      return {
        id: `sp_${i}`,
        date: row[0]||"",
        globalOrder: parseInt(row[1]||i),
        section: row[2]||"",
        label: row[3]||"",
        value: row[4]||"",
        notes: row[5]||"",
      };
    } else {
      // Legacy format (old Role/Name columns) — map best we can
      return {
        id: `sp_${i}`, date: "", globalOrder: i,
        section: row[0]||"", label: row[1]||"",
        value: row[2]||"", notes: row[5]||"",
      };
    }
  });
  return { sacramentProgram: rows };
}

export async function pullWardCouncilMeeting() {
  const r = await wcReq("WardCouncilMeeting!A:D");
  return { wardCouncilMeeting: rowsToMeeting(r.values) };
}

export async function pullBishopricLinks() {
  const r = await bishopricReq("Links!A:D");
  return { links: rowsToLinks(r.values) };
}

export async function pullWardCouncilLinks() {
  const r = await wcReq("Links!A:D");
  return { links: rowsToLinks(r.values) };
}

export async function sheetsLightPull() {
  const [appts, callings, releasings] = await Promise.all([
    bishopricReq("Appointments!A:F").then(r => rowsToAppointments(r.values)),
    bishopricReq("Callings!A:D").then(r => rowsToCallings(r.values)),
    bishopricReq("Releasings!A:D").then(r => rowsToCallings(r.values)),
  ]);
  return { appointments: appts, callings, releasings };
}

// ── Push functions ────────────────────────────────────────────────────────────

export async function pushAll({ appointments, callings, releasings, members }) {
  const sid = SID();
  const ops = [
    clearAndWrite(sid, "Appointments!A:F", appointmentsToRows(appointments)),
    clearAndWrite(sid, "Callings!A:D",     callingsToRows(callings)),
    clearAndWrite(sid, "Releasings!A:D",   callingsToRows(releasings)),
  ];
  if (members !== undefined) {
    ops.push(clearAndWrite(sid, "Notes!A:E", notesToRows(members)));
  }
  await Promise.all(ops);
}

export async function pushBishopricMeeting(data) {
  await clearAndWrite(SID(), "BishopricMeeting!A:D", meetingToRows(data));
}

export async function pushWardCouncilMeeting(data) {
  await clearAndWrite(WCID(), "WardCouncilMeeting!A:D", meetingToRows(data));
}

export async function pushSacramentProgram(rows) {
  const values = [
    ["Date", "GlobalOrder", "Section", "Label", "Value", "Notes"],
    ...rows.map(r => [r.date||"", r.globalOrder??0, r.section||"", r.label||"", r.value||"", r.notes||""]),
  ];
  await clearAndWrite(SSID(), "SacramentProgram!A:F", values);
}

// ── Calendar ──────────────────────────────────────────────────────────────────

export async function pullCalendar() {
  const r = await wcReq("Calendar!A:E");
  if (!r.values || r.values.length < 2) return { calendar: [] };
  return {
    calendar: r.values.slice(1).filter(r => r[0]).map((r, i) => ({
      id: `cal_${i}`, date: r[0]||"", time: r[1]||"",
      event: r[2]||"", location: r[3]||"", notes: r[4]||"",
    })),
  };
}

export async function pushCalendar(data) {
  const values = [
    ["Date", "Time", "Event", "Location", "Notes"],
    ...data.map(e => [e.date||"", e.time||"", e.event||"", e.location||"", e.notes||""]),
  ];
  await clearAndWrite(WCID(), "Calendar!A:E", values);
}

// ── Roster ────────────────────────────────────────────────────────────────────

export async function pullRoster() {
  const r = await bishopricReq("Roster!A:B");
  if (!r.values || r.values.length < 2) return { roster: [] };
  return {
    roster: r.values.slice(1).filter(r => r[0]).map(r => ({
      role: r[0]||"", name: r[1]||"",
    })),
  };
}

export async function pushRoster(data) {
  const values = [
    ["Role", "Name"],
    ...data.map(r => [r.role||"", r.name||""]),
  ];
  await clearAndWrite(SID(), "Roster!A:B", values);
}

// ── Aliases for backward compatibility ───────────────────────────────────────
export const pullSacrament = pullSacramentProgram;
export const pushSacrament = pushSacramentProgram;

// ── Health check ──────────────────────────────────────────────────────────────
export async function testConnection() {
  const res = await fetch(`${WORKER_URL}/health`);
  if (!res.ok) throw new Error(`Worker health check failed: ${res.status}`);
  return res.json();
}

// ── Prayer List ───────────────────────────────────────────────────────────────
export async function pullPrayerList() {
  try {
    const plid = PLID();
    if (!plid || plid.includes("YOUR_")) return [];
    // Try "Sheet1" first (default tab name), then "PrayerList"
    let r;
    try {
      r = await sheetsReq(plid, "Sheet1!A:E");
    } catch {
      r = await sheetsReq(plid, "PrayerList!A:E");
    }
    if (!r.values || r.values.length < 2) return [];
    return r.values.slice(1).filter(r => r[0]).map((r, i) => ({
      id: `pl_${i}`, category: r[0]||"", name: r[1]||"",
      notes: r[2]||"", date: r[3]||"", addedBy: r[4]||"",
    }));
  } catch(e) {
    console.error("pullPrayerList error:", e);
    return [];
  }
}

// ── Ward Council Roster ───────────────────────────────────────────────────────
export async function pullRosterFromWardCouncil() {
  try {
    const r = await wcReq("Roster!A:B");
    if (!r.values || r.values.length < 2) return { roster: [] };
    return {
      roster: r.values.slice(1).filter(r => r[0]).map(r => ({
        role: r[0]||"", name: r[1]||"",
      })),
    };
  } catch { return { roster: [] }; }
}
