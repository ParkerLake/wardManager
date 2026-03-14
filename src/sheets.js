import config from "./config";

const SID  = () => config.SPREADSHEET_ID;
const SSID = () => config.SACRAMENT_SHEET_ID;

async function sheetsReq(token, sheetId, range, method = "GET", body = null) {
  const base = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}`;
  // POST = append (requires :append suffix), PUT = update existing row
  const url  = method === "GET"  ? base
             : method === "POST" ? `${base}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`
             :                     `${base}?valueInputOption=USER_ENTERED`;
  const res  = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  if (!res.ok) {
    const t = await res.text();
    console.error(`Sheets API ${res.status} on ${method} ${range}:`, t);
    throw new Error(`Sheets API ${res.status}: ${t}`);
  }
  return res.json();
}

const bishopricReq = (token, range, method, body) => sheetsReq(token, SID(),  range, method, body);
const sacramentReq = (token, range, method, body) => sheetsReq(token, SSID(), range, method, body);

// Throttled sequential updates — avoids rate limits without needing batchUpdate scope
async function sheetsBatchUpdate(token, sheetId, data) {
  if (!data || data.length === 0) return;
  // Run updates in chunks of 5, sequentially, to stay well under API rate limits
  const CHUNK = 5;
  const reqFn = (token, range, method, body) => sheetsReq(token, sheetId, range, method, body);
  for (let i = 0; i < data.length; i += CHUNK) {
    const chunk = data.slice(i, i + CHUNK);
    await Promise.all(chunk.map(({ range, values }) => reqFn(token, range, "PUT", { values })));
  }
}

export async function testConnection(token) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SID()}?fields=sheets.properties.title`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Cannot reach spreadsheet (${res.status}). Check your SPREADSHEET_ID in config.js.\n${t}`);
  }
  const data    = await res.json();
  const tabs    = (data.sheets || []).map(s => s.properties.title);
  const need    = ["Appointments", "Callings", "Releasings", "Members", "BishopricMeeting", "Roster"];
  const missing = need.filter(n => !tabs.includes(n));
  return { tabs, missing };
}

// ── Row converters ─────────────────────────────────────────────────────────────

export function rowsToAppointments(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map((r, i) => ({
    id: `a_${i}`, name: r[0]||"", status: r[1]||"Need to Schedule",
    owner: r[2]||"", purpose: r[3]||"", apptDate: r[4]||"", notes: r[5]||"",
  }));
}

export function rowsToCallings(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map((r, i) => ({
    id: `c_${i}`, calling: r[0]||"", name: r[1]||"", stage: r[2]||"Discuss", notes: r[3]||"",
  }));
}

export function rowsToMembers(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map((r, i) => ({
    id: `mb_${i}`, name: r[0]||"", calling: r[1]||"",
    phone: r[2]||"", email: r[3]||"", notes: r[4]||"",
  }));
}

export function rowsToBishopricMeeting(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map((r, i) => ({
    id:           `bm_${i}`,
    date:         r[0]||"",
    itemKey:      r[1]||"",
    assignee:     r[2]||"",
    done:         r[3]==="TRUE",
    notes:        r[4]||"",
    customLabel:  r[5]||"",
    spiritToggle: r[6]||"",   // "spiritual_thought" | "handbook_review" | ""
  }));
}

export function rowsToSacramentProgram(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map((r, i) => ({
    id: `sp_${i}`,
    date:        r[0]||"",
    section:     r[1]||"",
    globalOrder: parseInt(r[2]||"0", 10),
    label:       r[3]||"",
    value:       r[4]||"",
    notes:       r[5]||"",
  }));
}

export function rowsToRoster(rows) {
  // rows: [[Role, Name], ...]
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({ role: r[0]||"", name: r[1]||"" }));
}

export function rosterToRows(data) {
  return [
    ["Role", "Name"],
    ...data.map(r => [r.role, r.name||""]),
  ];
}

// ── Row serializers ────────────────────────────────────────────────────────────

export function appointmentsToRows(data) {
  return [
    ["Name", "Status", "Owner", "Purpose", "Appt Date", "Notes"],
    ...data.map(a => [a.name, a.status, a.owner, a.purpose, a.apptDate, a.notes]),
  ];
}

export function callingsToRows(data) {
  return [
    ["Calling", "Name", "Stage", "Notes"],
    ...data.map(c => [c.calling, c.name, c.stage, c.notes]),
  ];
}

export function membersToRows(data) {
  return [
    ["Name", "Calling", "Phone", "Email", "Notes"],
    ...data.map(m => [m.name, m.calling||"", m.phone||"", m.email||"", m.notes||""]),
  ];
}

export function bishopricMeetingToRows(data) {
  return [
    ["Date", "ItemKey", "Assignee", "Done", "Notes", "CustomLabel", "SpiritualToggle"],
    ...data.map(r => [
      r.date, r.itemKey, r.assignee||"",
      r.done ? "TRUE" : "FALSE",
      r.notes||"", r.customLabel||"", r.spiritToggle||"",
    ]),
  ];
}

export function sacramentProgramToRows(data) {
  return [
    ["Date", "Section", "Order", "Label", "Value", "Notes"],
    ...data.map(r => [r.date, r.section, r.globalOrder ?? 0, r.label, r.value, r.notes]),
  ];
}

// ── Pull / Push ────────────────────────────────────────────────────────────────

export async function pullAll(token) {
  const [a, c, r, mb, bm, ros] = await Promise.all([
    bishopricReq(token, "Appointments!A:F"),
    bishopricReq(token, "Callings!A:D"),
    bishopricReq(token, "Releasings!A:D"),
    bishopricReq(token, "Members!A:E").catch(() => ({ values: [] })),
    bishopricReq(token, "BishopricMeeting!A:G").catch(() => ({ values: [] })),
    bishopricReq(token, "Roster!A:B").catch(() => ({ values: [] })),
  ]);
  return {
    appointments:    rowsToAppointments(a.values),
    callings:        rowsToCallings(c.values),
    releasings:      rowsToCallings(r.values),
    members:         rowsToMembers(mb.values),
    bishopricMeeting: rowsToBishopricMeeting(bm.values),
    roster:          rowsToRoster(ros.values),
  };
}

export async function pullSacrament(token) {
  if (!SSID() || SSID().includes("YOUR_")) return { sacramentProgram: [] };
  const sp = await sacramentReq(token, "SacramentProgram!A:F").catch(() => ({ values: [] }));
  return { sacramentProgram: rowsToSacramentProgram(sp.values) };
}

export async function pushAll(token, { appointments, callings, releasings, members }) {
  const ops = [
    bishopricReq(token, "Appointments!A:F", "PUT", { values: appointmentsToRows(appointments) }),
    bishopricReq(token, "Callings!A:D",     "PUT", { values: callingsToRows(callings) }),
    bishopricReq(token, "Releasings!A:D",   "PUT", { values: callingsToRows(releasings) }),
  ];
  if (members !== undefined) ops.push(bishopricReq(token, "Members!A:E", "PUT", { values: membersToRows(members) }));
  await Promise.all(ops);
}

export async function pushBishopricMeeting(token, data) {
  if (!token) throw new Error("Not authenticated — please refresh and sign in again");
  if (!data || !Array.isArray(data) || data.length === 0) return;
  // Pull current sheet to find row positions
  const current = await bishopricReq(token, "BishopricMeeting!A:G").catch(() => ({ values: [] }));
  const rows = current.values || [];

  const rowIndex = {};
  rows.slice(1).forEach((r, i) => {
    if (r[0] && r[1]) rowIndex[`${r[0]}|${r[1]}`] = i + 2;
  });

  const toUpdate = []; // batchUpdate data entries
  const toAppend = []; // rows to append

  data.forEach(r => {
    const key = `${r.date}|${r.itemKey}`;
    const rowValues = [r.date, r.itemKey, r.assignee||"", r.done?"TRUE":"FALSE", r.notes||"", r.customLabel||"", r.spiritToggle||""];
    if (rowIndex[key]) {
      const n = rowIndex[key];
      toUpdate.push({ range: `BishopricMeeting!A${n}:G${n}`, values: [rowValues] });
    } else {
      toAppend.push(rowValues);
    }
  });

  // Single batch call for all updates, then one append call
  await sheetsBatchUpdate(token, SID(), toUpdate);

  if (toAppend.length > 0) {
    await bishopricReq(token, "BishopricMeeting!A:G", "POST", { values: toAppend });
  }
}

export async function pullWardCouncilMeeting(token) {
  if (!WCID() || WCID().includes("YOUR_")) return { wardCouncilMeeting: [] };
  const res = await wardCouncilReq(token, "WardCouncilMeeting!A:G").catch(() => ({ values: [] }));
  return { wardCouncilMeeting: rowsToWardCouncilMeeting(res.values) };
}

export async function pushWardCouncilMeeting(token, data) {
  if (!token) throw new Error("Not authenticated — please refresh and sign in again");
  if (!data || !Array.isArray(data) || data.length === 0) return;
  // Pull current sheet to find row positions
  const current = await wardCouncilReq(token, "WardCouncilMeeting!A:G").catch(() => ({ values: [] }));
  const rows = current.values || [];

  const rowIndex = {};
  rows.slice(1).forEach((r, i) => {
    if (r[0] && r[1]) rowIndex[`${r[0]}|${r[1]}`] = i + 2;
  });

  const toUpdate = [];
  const toAppend = [];

  data.forEach(r => {
    const key = `${r.date}|${r.itemKey}`;
    const rowValues = [r.date, r.itemKey, r.assignee||"", r.done?"TRUE":"FALSE", r.notes||"", r.customLabel||"", r.spiritToggle||""];
    if (rowIndex[key]) {
      const n = rowIndex[key];
      toUpdate.push({ range: `WardCouncilMeeting!A${n}:G${n}`, values: [rowValues] });
    } else {
      toAppend.push(rowValues);
    }
  });

  // Single batch call for all updates, then one append call
  await sheetsBatchUpdate(token, WCID(), toUpdate);

  if (toAppend.length > 0) {
    await wardCouncilReq(token, "WardCouncilMeeting!A:G", "POST", { values: toAppend });
  }
}

export async function pushRoster(token, data) {
  await bishopricReq(token, "Roster!A:B", "PUT", { values: rosterToRows(data) });
}

export async function pushSacrament(token, sacramentProgram) {
  if (!SSID() || SSID().includes("YOUR_")) return;
  await sacramentReq(token, "SacramentProgram!A:F", "PUT", {
    values: sacramentProgramToRows(sacramentProgram),
  });
}

// ── Prayer List (separate sheet, read-only) ──────────────────────────────────
const prayerListReq = (token, range, method, body) =>
  sheetsReq(token, config.PRAYER_LIST_SHEET_ID, range, method, body);

// ── Ward Council / Calendar ───────────────────────────────────────────────────
const WCID = () => config.WARD_COUNCIL_SHEET_ID;
const wardCouncilReq = (token, range, method, body) => sheetsReq(token, WCID(), range, method, body);

export function rowsToWardCouncilMeeting(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map((r, i) => ({
    id:           `wc_${i}`,
    date:         r[0]||"",
    itemKey:      r[1]||"",
    assignee:     r[2]||"",
    done:         r[3]==="TRUE",
    notes:        r[4]||"",
    customLabel:  r[5]||"",
    spiritToggle: r[6]||"",
  }));
}

export function wardCouncilMeetingToRows(data) {
  return [
    ["Date", "ItemKey", "Assignee", "Done", "Notes", "CustomLabel", "SpiritualToggle"],
    ...data.map(r => [
      r.date, r.itemKey, r.assignee||"",
      r.done ? "TRUE" : "FALSE",
      r.notes||"", r.customLabel||"", r.spiritToggle||"",
    ]),
  ];
}

export function rowsToCalendar(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1)
    .filter(r => r[0] || r[2]) // must have date or event
    .map((r, i) => ({
      id:    `cal_${i}`,
      date:  r[0] || "",
      time:  r[1] || "",
      event: r[2] || "",
    }));
}

export function calendarToRows(data) {
  return [
    ["Date", "Time", "Event"],
    ...data.map(r => [r.date, r.time || "", r.event]),
  ];
}

export async function pullRosterFromWardCouncil(token) {
  const wcid = config.WARD_COUNCIL_SHEET_ID;
  if (!wcid || wcid.includes("YOUR_")) return { roster: [] };
  const res = await wardCouncilReq(token, "Roster!A:B").catch(() => ({ values: [] }));
  return { roster: rowsToRoster(res.values || []) };
}

export async function pullCalendar(token) {
  const wcid = config.WARD_COUNCIL_SHEET_ID;
  if (!wcid || wcid.includes("YOUR_")) return { calendar: [] };
  const data = await wardCouncilReq(token, "Calendar!A:C");
  return { calendar: rowsToCalendar(data.values || []) };
}

export async function pushCalendar(token, data) {
  const rows = calendarToRows(data);
  await wardCouncilReq(token, "Calendar!A1", "PUT", { values: rows });
}

export async function pullPrayerList(token) {
  const plid = config.PRAYER_LIST_SHEET_ID;
  if (!plid || plid.includes("YOUR_")) return [];
  const d = await prayerListReq(token, "Sheet1!A:B").catch(() => ({ values: [] }));
  const rows = d.values || [];
  if (rows.length < 2) return [];
  // Skip header row, return [{category, name}]
  return rows.slice(1)
    .filter(r => r[0] || r[1])
    .map(r => ({ category: r[0] || "", name: r[1] || "" }));
}

// ─── Links ────────────────────────────────────────────────────────────────────
function rowsToLinks(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1)
    .filter(r => r[1] || r[2]) // must have name or URL
    .map(r => ({
      id:          r[0] || `link_${Math.random().toString(36).slice(2,9)}`,
      name:        r[1] || "",
      url:         r[2] || "",
      description: r[3] || "",
    }));
}

function linksToRows(data) {
  return [
    ["ID", "Name", "URL", "Description"],
    ...data.map(r => [r.id, r.name, r.url, r.description || ""]),
  ];
}

export async function pullBishopricLinks(token) {
  if (!SID() || SID().includes("YOUR_")) return { links: [] };
  try {
    const data = await bishopricReq(token, "Links!A:D");
    return { links: rowsToLinks(data.values || []) };
  } catch(_) { return { links: [] }; }
}

export async function pushBishopricLinks(token, data) {
  if (!token) throw new Error("Not authenticated");
  await bishopricReq(token, "Links!A1", "PUT", { values: linksToRows(data) });
}

export async function pullWardCouncilLinks(token) {
  if (!WCID() || WCID().includes("YOUR_")) return { links: [] };
  try {
    const data = await wardCouncilReq(token, "Links!A:D");
    return { links: rowsToLinks(data.values || []) };
  } catch(_) { return { links: [] }; }
}

export async function pushWardCouncilLinks(token, data) {
  if (!token) throw new Error("Not authenticated");
  await wardCouncilReq(token, "Links!A1", "PUT", { values: linksToRows(data) });
}
