/**
 * Central RO Dashboard — Access Control backend.
 *
 * Deploy this bound to a NEW, PRIVATE Google Sheet (do NOT share it —
 * unlike the data sheet, this one must stay locked down since it holds
 * user records). See setup steps in the chat reply that shipped this file.
 *
 * Sheet tabs expected on that private spreadsheet:
 *   "Users": ID | Name | Role | PasswordHash | Active | CreatedAt
 *   "Logs":  Timestamp | UserID | Role | Event | Details
 *
 * Roles used by the dashboard: Admin, Technician, Viewer (case-insensitive).
 */

const SESSION_TTL_SECONDS = 8 * 60 * 60; // 8 hours

function getUsersSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
}
function getLogsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
}

function hashPassword(plain) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain, Utilities.Charset.UTF_8);
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function findUserRow(id) {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      return { rowIndex: i + 1, id: data[i][0], name: data[i][1], role: data[i][2], hash: data[i][3], active: data[i][4] };
    }
  }
  return null;
}

function appendLog(userId, role, event, details) {
  getLogsSheet().appendRow([new Date(), userId || "", role || "", event || "", details || ""]);
}

function putSession(token, payload) {
  CacheService.getScriptCache().put(token, JSON.stringify(payload), SESSION_TTL_SECONDS);
}
function getSession(token) {
  const raw = CacheService.getScriptCache().get(token);
  return raw ? JSON.parse(raw) : null;
}
function removeSession(token) {
  CacheService.getScriptCache().remove(token);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: "Bad request" });
  }

  const action = body.action;

  if (action === "login") {
    const user = findUserRow(body.id);
    if (!user || user.active === false || user.active === "FALSE") {
      return jsonOut({ ok: false, error: "Invalid ID or password" });
    }
    if (hashPassword(String(body.password || "")) !== user.hash) {
      return jsonOut({ ok: false, error: "Invalid ID or password" });
    }
    const token = Utilities.getUuid();
    putSession(token, { id: user.id, name: user.name, role: user.role });
    appendLog(user.id, user.role, "login", "");
    return jsonOut({ ok: true, token, name: user.name, role: user.role });
  }

  if (action === "logout") {
    const session = getSession(body.token);
    if (session) {
      appendLog(session.id, session.role, "logout", "");
      removeSession(body.token);
    }
    return jsonOut({ ok: true });
  }

  if (action === "log") {
    const session = getSession(body.token);
    if (!session) return jsonOut({ ok: false, error: "Session expired" });
    appendLog(session.id, session.role, body.event || "event", body.details || "");
    return jsonOut({ ok: true });
  }

  if (action === "listLogs") {
    const session = getSession(body.token);
    if (!session || String(session.role).toLowerCase() !== "admin") {
      return jsonOut({ ok: false, error: "Not authorized" });
    }
    const data = getLogsSheet().getDataRange().getValues();
    const rows = data.slice(1).map(r => ({
      timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      userId: r[1], role: r[2], event: r[3], details: r[4]
    }));
    rows.reverse();
    const limit = Number(body.limit) || 100;
    return jsonOut({ ok: true, logs: rows.slice(0, limit) });
  }

  return jsonOut({ ok: false, error: "Unknown action" });
}

/**
 * Adds a menu to the private sheet's UI so the admin (non-engineer) can
 * create/update accounts without ever touching code or the raw password hash.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("RO Dashboard Access")
    .addItem("Add / update user", "addOrUpdateUserPrompt")
    .addItem("Deactivate user", "deactivateUserPrompt")
    .addToUi();
}

function addOrUpdateUserPrompt() {
  const ui = SpreadsheetApp.getUi();
  const idResp = ui.prompt("Add / update user", "Staff ID (e.g. PW-014):", ui.ButtonSet.OK_CANCEL);
  if (idResp.getSelectedButton() !== ui.Button.OK || !idResp.getResponseText().trim()) return;
  const id = idResp.getResponseText().trim();

  const nameResp = ui.prompt("Add / update user", "Full name:", ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  const name = nameResp.getResponseText().trim();

  const roleResp = ui.prompt("Add / update user", "Role — type exactly one of: Admin, Technician, Viewer", ui.ButtonSet.OK_CANCEL);
  if (roleResp.getSelectedButton() !== ui.Button.OK) return;
  const role = roleResp.getResponseText().trim();
  if (!["Admin", "Technician", "Viewer"].includes(role)) {
    ui.alert("Role must be exactly: Admin, Technician, or Viewer.");
    return;
  }

  const pwResp = ui.prompt("Add / update user", "Password for this user:", ui.ButtonSet.OK_CANCEL);
  if (pwResp.getSelectedButton() !== ui.Button.OK || !pwResp.getResponseText()) return;
  const hash = hashPassword(pwResp.getResponseText());

  const sheet = getUsersSheet();
  const existing = findUserRow(id);
  if (existing) {
    sheet.getRange(existing.rowIndex, 1, 1, 6).setValues([[id, name, role, hash, true, existing.rowIndex ? sheet.getRange(existing.rowIndex, 6).getValue() : new Date()]]);
    ui.alert(`Updated user ${id}.`);
  } else {
    sheet.appendRow([id, name, role, hash, true, new Date()]);
    ui.alert(`Added user ${id} as ${role}.`);
  }
}

function deactivateUserPrompt() {
  const ui = SpreadsheetApp.getUi();
  const idResp = ui.prompt("Deactivate user", "Staff ID to deactivate:", ui.ButtonSet.OK_CANCEL);
  if (idResp.getSelectedButton() !== ui.Button.OK) return;
  const user = findUserRow(idResp.getResponseText().trim());
  if (!user) { ui.alert("No such user."); return; }
  getUsersSheet().getRange(user.rowIndex, 5).setValue(false);
  ui.alert(`Deactivated ${user.id}.`);
}
