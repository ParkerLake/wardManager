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
  const header = rows[0];
  // Detect old format (Topic|Owner|Status|Notes) vs new (Date|ItemKey|...)
  const isNewFormat = header[0] === "Date" && header[1] === "ItemKey";
  if (isNewFormat) {
    return rows.slice(1)
      .filter(r => r[0] || r[1]) // must have date or itemKey
      .map((r, i) => ({
        id: `mt_${i}`,
        date: r[0]||"",
        itemKey: r[1]||"",
        assignee: r[2]||"",
        done: r[3] === true || (typeof r[3] === "string" && r[3].toLowerCase() === "true"),
        notes: r[4]||"",
        customLabel: r[5]||"",
        spiritToggle: r[6]||"",
        topicOrder: r[7] !== undefined && r[7] !== "" ? parseInt(r[7]) : undefined,
      }));
  }
  // Legacy format — return empty so app creates fresh agenda
  return [];
}

function meetingToRows(data) {
  const rows = [
    ["Date", "ItemKey", "Assignee", "Done", "Notes", "CustomLabel", "SpiritToggle", "TopicOrder"],
    ...data.map(r => [
      r.date||"", r.itemKey||"", r.assignee||"",
      r.done ? "true" : "false",
      r.notes||"", r.customLabel||"", r.spiritToggle||"",
      r.topicOrder !== undefined ? String(r.topicOrder) : "",
    ]),
  ];
  // Pad to 500 rows to clear any trailing stale rows
  while (rows.length < 500) rows.push(["","","","","","","",""]);
  return rows;
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
    bishopricReq("BishopricMeeting!A:H").then(r => rowsToMeeting(r.values)),
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
  const r = await wcReq("WardCouncilMeeting!A:H");
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


// ── 3-way merge for same-date collision protection ────────────────────────────
// local    = User B's current edits for this date
// remote   = what's currently in the sheet (may include User A's edits)
// base     = what User B loaded when they first opened the tab
//
// For each row, field by field:
//   if User B changed the field (local != base) → use local (User B's intent)
//   if User B didn't change the field           → use remote (may have User A's edit)
//
// For dynamic rows (topic_*, task_*):
//   rows only in local (not in remote, not in base) → User B added them → keep
//   rows in remote but not in local AND not in base → User A added them → include
//   rows in remote + base but not in local          → User B deleted them → exclude
function threewayMergeDate(localRows, remoteRows, baseRows) {
  const MERGE_FIELDS = ["assignee", "done", "notes", "customLabel", "spiritToggle", "topicOrder"];

  const remoteMap = new Map(remoteRows.map(r => [r.itemKey, r]));
  const baseMap   = new Map(baseRows.map(r  => [r.itemKey, r]));
  const localMap  = new Map(localRows.map(r  => [r.itemKey, r]));

  const result = [];

  // Process all local rows (User B's version of existing + new rows)
  localRows.forEach(localRow => {
    const remote = remoteMap.get(localRow.itemKey);
    const base   = baseMap.get(localRow.itemKey);

    if (!remote || !base) {
      // Row is new to User B (not in remote or base) — keep as-is
      result.push(localRow);
      return;
    }

    // Both users have this row — merge field by field
    const merged = { ...localRow };
    MERGE_FIELDS.forEach(field => {
      const localChanged = localRow[field] !== base[field];
      if (!localChanged && remote[field] !== undefined) {
        // User B didn't change this field → defer to remote (may have User A's edit)
        merged[field] = remote[field];
      }
      // If User B changed it → keep User B's value (already in merged via spread)
    });
    result.push(merged);
  });

  // Include rows User A added after User B loaded (in remote, not in local, not in base)
  remoteRows.forEach(remoteRow => {
    const inLocal = localMap.has(remoteRow.itemKey);
    const inBase  = baseMap.has(remoteRow.itemKey);
    if (!inLocal && !inBase) {
      // User A added this row after User B loaded — include it
      result.push(remoteRow);
    }
    // If in base but not in local → User B deleted it → respect deletion, don't include
  });

  return result;
}

// ── Merge other-date rows: prefer sheet data (preserves peer edits),
//    fall back to local memory (preserves unsaved local agendas) ─────────────
function mergeOtherDates(remoteAll, localAll, currentDate) {
  const remoteOthers = remoteAll.filter(r => r.date && r.date !== currentDate);
  const localOthers  = localAll.filter(r  => r.date && r.date !== currentDate);

  if (remoteAll.length === 0) {
    // Sheet was empty or in legacy format — use local as the only source
    return localOthers;
  }

  // Sheet has parseable data: use sheet's other dates as primary
  // Also include any local-only dates not yet saved to the sheet
  const remoteDates = new Set(remoteOthers.map(r => r.date));
  const localOnlyOthers = localOthers.filter(r => !remoteDates.has(r.date));
  return [...remoteOthers, ...localOnlyOthers];
}

export async function pushBishopricMeeting(localDateRows, date, baseRows = [], localAllRows = []) {
  let remoteAll = [];
  try {
    const r = await sheetsReq("BishopricMeeting!A:H");
    remoteAll = rowsToMeeting(r.values);
  } catch(_) {}
  // Other dates: prefer sheet (peer edits), fall back to local memory (unsaved agendas)
  const otherDates = mergeOtherDates(remoteAll, localAllRows, date);
  // Current date: 3-way field merge if we have a base snapshot, otherwise local wins
  const remoteDate = remoteAll.filter(r => r.date === date);
  const mergedDate = baseRows.length > 0 && remoteDate.length > 0
    ? threewayMergeDate(localDateRows, remoteDate, baseRows)
    : localDateRows;
  await clearAndWrite(SID(), "BishopricMeeting!A:H", meetingToRows([...otherDates, ...mergedDate]));
}

export async function pushWardCouncilMeeting(localDateRows, date, baseRows = [], localAllRows = []) {
  let remoteAll = [];
  try {
    const r = await wcReq("WardCouncilMeeting!A:H");
    remoteAll = rowsToMeeting(r.values);
  } catch(_) {}
  const otherDates = mergeOtherDates(remoteAll, localAllRows, date);
  const remoteDate = remoteAll.filter(r => r.date === date);
  const mergedDate = baseRows.length > 0 && remoteDate.length > 0
    ? threewayMergeDate(localDateRows, remoteDate, baseRows)
    : localDateRows;
  await clearAndWrite(WCID(), "WardCouncilMeeting!A:H", meetingToRows([...otherDates, ...mergedDate]));
}

export async function pushSacramentProgram(localDateRows, date, baseRows = [], localAllRows = []) {
  // Pull full current sheet for merge
  let remoteAll = [];
  try {
    const r = await sacramentReq("SacramentProgram!A:F");
    if (r.values && r.values.length > 1 && r.values[0][0] === "Date") {
      remoteAll = r.values.slice(1)
        .filter(row => row.some(c => c))
        .map((row, i) => ({
          id: `sp_remote_${i}`, date: row[0]||"",
          globalOrder: parseInt(row[1]||i), section: row[2]||"",
          label: row[3]||"", value: row[4]||"", notes: row[5]||"",
        }));
    }
  } catch(_) {}

  // Other dates: prefer sheet (peer edits), fall back to local memory (unsaved programs)
  const otherDateRows = mergeOtherDates(remoteAll, localAllRows, date);
  const remoteDate    = remoteAll.filter(r => r.date === date);

  // Current date: 3-way merge on value/notes fields, matched by globalOrder
  const SACR_FIELDS = ["value", "notes", "label", "globalOrder"];
  let mergedDate = localDateRows;
  if (baseRows.length > 0 && remoteDate.length > 0) {
    const remoteMap = new Map(remoteDate.map(r => [r.globalOrder, r]));
    const baseMap   = new Map(baseRows.map(r   => [r.globalOrder, r]));
    mergedDate = localDateRows.map(localRow => {
      const remote = remoteMap.get(localRow.globalOrder);
      const base   = baseMap.get(localRow.globalOrder);
      if (!remote || !base) return localRow;
      const merged = { ...localRow };
      SACR_FIELDS.forEach(field => {
        const localChanged = localRow[field] !== base[field];
        if (!localChanged && remote[field] !== undefined) {
          merged[field] = remote[field];
        }
      });
      return merged;
    });
  }

  const merged = [...otherDateRows, ...mergedDate];
  const values = [
    ["Date", "GlobalOrder", "Section", "Label", "Value", "Notes"],
    ...merged.map(r => [r.date||"", r.globalOrder??0, r.section||"", r.label||"", r.value||"", r.notes||""]),
  ];
  while (values.length < 500) values.push(["","","","","",""]);
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


// ── Narrative Templates ───────────────────────────────────────────────────────

const DEFAULT_NARRATIVE_TEMPLATES = {
  intro:         "Welcome brothers and sisters. We are glad you could join us today for sacrament meeting.",
  organist_intro:"We'd like to thank our organist {organist} and music leader {chorister} for their help with the music today.",
  releasing:     "The following individuals have been released and we propose they be given a vote of thanks for their excellent service.\n{names}\nAll who wish to show their appreciation, please manifest it by the uplifted hand.",
  calling:       "The following have accepted a call to serve in the ward. (We ask that you please stand as your name is called.)\n{names_callings}\nWe propose that they be sustained. Those in favor may manifest it by the uplifted hand. Those opposed, if any, may manifest it.",
  new_member:    "Would the following new members in our ward please stand:\n{names}\nWe propose they be welcomed into the [name of ward] Ward, [name of stake] Stake, of The Church of Jesus Christ of Latter-day Saints. Those in favor may manifest it by the uplifted hand.",
  ordination:    "We propose the following receive the Aaronic Priesthood and be ordained.\n{names_offices}\nThose in favor may manifest it by the uplifted hand. (Pause for vote) Those opposed, if any, may manifest it.",
};

export async function pullNarrativeTemplates() {
  try {
    const r = await bishopricReq("NarrativeTemplates!A:B");
    if (!r.values || r.values.length < 2) return { templates: { ...DEFAULT_NARRATIVE_TEMPLATES } };
    const templates = { ...DEFAULT_NARRATIVE_TEMPLATES };
    r.values.slice(1).forEach(row => {
      if (row[0] && row[1] !== undefined) templates[row[0]] = row[1];
    });
    return { templates };
  } catch(_) {
    return { templates: { ...DEFAULT_NARRATIVE_TEMPLATES } };
  }
}

export async function pushNarrativeTemplates(templates) {
  const values = [
    ["Key", "Template"],
    ...Object.entries(templates).map(([k, v]) => [k, v]),
  ];
  await clearAndWrite(SID(), "NarrativeTemplates!A:B", values);
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
