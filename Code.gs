// ============================================================
// RAFFLE PLATFORM - Google Apps Script Backend
// Deploy as: Extensions > Apps Script > Deploy > Web App
//   Execute as: Me | Who has access: Anyone
// ============================================================

const MASTER_SHEET_ID = "YOUR_MASTER_GOOGLE_SHEET_ID_HERE"; // <-- Replace this

// ---- CORS helper ----
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const body = e.postData ? JSON.parse(e.postData.contents || "{}") : {};
  const action = params.action || body.action;

  let result;
  try {
    switch (action) {
      case "createRaffle":      result = createRaffle(body); break;
      case "getRaffle":         result = getRaffle(body); break;
      case "addDistributor":    result = addDistributor(body); break;
      case "getDistributors":   result = getDistributors(body); break;
      case "loginDistributor":  result = loginDistributor(body); break;
      case "createTicket":      result = createTicket(body); break;
      case "getTickets":        result = getTickets(body); break;
      case "closeRaffle":       result = closeRaffle(body); break;
      case "adminLogin":        result = adminLogin(body); break;
      default: result = { error: "Unknown action: " + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Utilities ----
function getMasterSheet(tabName) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }
  return sheet;
}

function getRaffleSheet(raffleSheetId, tabName) {
  const ss = SpreadsheetApp.openById(raffleSheetId);
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }
  return sheet;
}

function generateId(prefix) {
  return prefix + "_" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5).toUpperCase();
}

// ---- Actions ----

// Create a new raffle — creates a new Google Sheet for this raffle
function createRaffle(body) {
  const { raffleName, eventName, adminPassword, templateId, logoUrl, bodyImageUrl, googleSheetLink } = body;
  if (!raffleName || !adminPassword) throw new Error("raffleName and adminPassword required");

  // Use provided sheet link, or create a new one
  let raffleSheetId;
  if (googleSheetLink) {
    const match = googleSheetLink.match(/\/d\/([\w-]+)/);
    if (!match) throw new Error("Invalid Google Sheet link");
    raffleSheetId = match[1];
  } else {
    const newSS = SpreadsheetApp.create("Raffle: " + raffleName);
    raffleSheetId = newSS.getId();
  }

  const raffleId = generateId("R");
  const passcode = adminPassword;

  // Init the raffle-level sheet tabs
  const ticketSheet = getRaffleSheet(raffleSheetId, "Tickets");
  if (ticketSheet.getLastRow() === 0) {
    ticketSheet.appendRow(["TicketNumber", "FirstName", "LastName", "Contact", "CareOf", "DistributorId", "DistributorName", "CreatedAt"]);
    ticketSheet.getRange(1, 1, 1, 8).setFontWeight("bold");
  }

  const distSheet = getRaffleSheet(raffleSheetId, "Distributors");
  if (distSheet.getLastRow() === 0) {
    distSheet.appendRow(["DistributorId", "Name", "Password", "CreatedAt", "Active"]);
    distSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }

  const metaSheet = getRaffleSheet(raffleSheetId, "Meta");
  if (metaSheet.getLastRow() === 0) {
    metaSheet.appendRow(["Key", "Value"]);
  }
  metaSheet.appendRow(["raffleId", raffleId]);
  metaSheet.appendRow(["raffleName", raffleName]);
  metaSheet.appendRow(["eventName", eventName || ""]);
  metaSheet.appendRow(["adminPassword", passcode]);
  metaSheet.appendRow(["templateId", templateId || "classic"]);
  metaSheet.appendRow(["logoUrl", logoUrl || ""]);
  metaSheet.appendRow(["bodyImageUrl", bodyImageUrl || ""]);
  metaSheet.appendRow(["status", "active"]);
  metaSheet.appendRow(["createdAt", new Date().toISOString()]);

  // Register in master sheet
  const masterSheet = getMasterSheet("Raffles");
  if (masterSheet.getLastRow() === 0) {
    masterSheet.appendRow(["RaffleId", "RaffleName", "EventName", "SheetId", "Status", "CreatedAt"]);
    masterSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  masterSheet.appendRow([raffleId, raffleName, eventName || "", raffleSheetId, "active", new Date().toISOString()]);

  return { success: true, raffleId, raffleSheetId };
}

// Get raffle meta info (for ticketing screen)
function getRaffle(body) {
  const { raffleSheetId } = body;
  if (!raffleSheetId) throw new Error("raffleSheetId required");

  const metaSheet = getRaffleSheet(raffleSheetId, "Meta");
  const data = metaSheet.getDataRange().getValues();
  const meta = {};
  data.forEach(row => { if (row[0]) meta[row[0]] = row[1]; });

  const distSheet = getRaffleSheet(raffleSheetId, "Distributors");
  const distCount = Math.max(0, distSheet.getLastRow() - 1);

  const ticketSheet = getRaffleSheet(raffleSheetId, "Tickets");
  const ticketCount = Math.max(0, ticketSheet.getLastRow() - 1);

  return { success: true, meta, distCount, ticketCount };
}

// Admin login
function adminLogin(body) {
  const { raffleSheetId, password } = body;
  if (!raffleSheetId || !password) throw new Error("raffleSheetId and password required");

  const metaSheet = getRaffleSheet(raffleSheetId, "Meta");
  const data = metaSheet.getDataRange().getValues();
  const meta = {};
  data.forEach(row => { if (row[0]) meta[row[0]] = row[1]; });

  if (meta.adminPassword !== password) return { success: false, error: "Invalid passcode" };

  const ticketSheet = getRaffleSheet(raffleSheetId, "Tickets");
  const tickets = [];
  const tData = ticketSheet.getDataRange().getValues();
  const headers = tData[0];
  for (let i = 1; i < tData.length; i++) {
    const obj = {};
    headers.forEach((h, j) => obj[h] = tData[i][j]);
    tickets.push(obj);
  }

  const distSheet = getRaffleSheet(raffleSheetId, "Distributors");
  const distributors = [];
  const dData = distSheet.getDataRange().getValues();
  const dHeaders = dData[0];
  for (let i = 1; i < dData.length; i++) {
    const obj = {};
    dHeaders.forEach((h, j) => obj[h] = dData[i][j]);
    distributors.push(obj);
  }

  return { success: true, meta, tickets, distributors };
}

// Add distributor
function addDistributor(body) {
  const { raffleSheetId, adminPassword, name, distPassword } = body;
  if (!raffleSheetId || !adminPassword || !name || !distPassword) throw new Error("Missing fields");

  // Verify admin
  const metaSheet = getRaffleSheet(raffleSheetId, "Meta");
  const data = metaSheet.getDataRange().getValues();
  const meta = {};
  data.forEach(row => { if (row[0]) meta[row[0]] = row[1]; });
  if (meta.adminPassword !== adminPassword) return { success: false, error: "Invalid admin password" };

  const distSheet = getRaffleSheet(raffleSheetId, "Distributors");
  const count = Math.max(0, distSheet.getLastRow() - 1);
  const distId = "D" + String(count + 1).padStart(2, "0");

  distSheet.appendRow([distId, name, distPassword, new Date().toISOString(), "true"]);
  return { success: true, distId };
}

// Get distributors list (public — no passwords returned)
function getDistributors(body) {
  const { raffleSheetId } = body;
  const distSheet = getRaffleSheet(raffleSheetId, "Distributors");
  const data = distSheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, distributors: [] };
  const headers = data[0];
  const distributors = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, j) => { if (h !== "Password") obj[h] = data[i][j]; });
    distributors.push(obj);
  }
  return { success: true, distributors };
}

// Distributor login
function loginDistributor(body) {
  const { raffleSheetId, distPassword } = body;
  if (!raffleSheetId || !distPassword) throw new Error("raffleSheetId and distPassword required");

  const distSheet = getRaffleSheet(raffleSheetId, "Distributors");
  const data = distSheet.getDataRange().getValues();
  const headers = data[0];
  const pwIdx = headers.indexOf("Password");
  const idIdx = headers.indexOf("DistributorId");
  const nameIdx = headers.indexOf("Name");
  const activeIdx = headers.indexOf("Active");

  for (let i = 1; i < data.length; i++) {
    if (data[i][pwIdx] === distPassword && data[i][activeIdx] === "true") {
      return { success: true, distId: data[i][idIdx], distName: data[i][nameIdx] };
    }
  }
  return { success: false, error: "Invalid distributor password" };
}

// Create ticket — uses lock to prevent race condition
function createTicket(body) {
  const { raffleSheetId, distId, distName, firstName, lastName, contact, careOf } = body;
  if (!raffleSheetId || !distId || !firstName || !lastName) throw new Error("Missing required fields");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // wait up to 10s

  try {
    const metaSheet = getRaffleSheet(raffleSheetId, "Meta");
    const metaData = metaSheet.getDataRange().getValues();
    const meta = {};
    metaData.forEach(row => { if (row[0]) meta[row[0]] = row[1]; });

    if (meta.status === "closed") return { success: false, error: "Raffle is closed" };

    const ticketSheet = getRaffleSheet(raffleSheetId, "Tickets");
    const lastRow = ticketSheet.getLastRow();

    // Count tickets by this distributor to get their sequence
    let distSeq = 0;
    if (lastRow > 1) {
      const existingData = ticketSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      existingData.forEach(row => {
        if (row[5] === distId) distSeq++;
      });
    }

    const raffleId = meta.raffleId || "R";
    const seqNum = String(distSeq + 1).padStart(4, "0");
    const ticketNumber = `${raffleId.slice(-6)}-${distId}-${seqNum}`;

    ticketSheet.appendRow([
      ticketNumber,
      firstName,
      lastName,
      contact || "",
      careOf || "",
      distId,
      distName || "",
      new Date().toISOString()
    ]);

    return { success: true, ticketNumber, raffleName: meta.raffleName, eventName: meta.eventName };
  } finally {
    lock.releaseLock();
  }
}

// Get all tickets (admin only — gated by caller)
function getTickets(body) {
  const { raffleSheetId } = body;
  const ticketSheet = getRaffleSheet(raffleSheetId, "Tickets");
  const data = ticketSheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, tickets: [] };
  const headers = data[0];
  const tickets = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, j) => obj[h] = data[i][j]);
    tickets.push(obj);
  }
  return { success: true, tickets };
}

// Close raffle
function closeRaffle(body) {
  const { raffleSheetId, adminPassword } = body;
  const metaSheet = getRaffleSheet(raffleSheetId, "Meta");
  const data = metaSheet.getDataRange().getValues();
  const meta = {};
  data.forEach((row, i) => { if (row[0]) meta[row[0]] = { val: row[1], idx: i + 1 }; });

  if (meta.adminPassword.val !== adminPassword) return { success: false, error: "Invalid password" };

  // Update status row
  const statusRow = meta.status.idx;
  metaSheet.getRange(statusRow, 2).setValue("closed");

  // Update master sheet too
  const masterSheet = getMasterSheet("Raffles");
  const mData = masterSheet.getDataRange().getValues();
  const raffleId = meta.raffleId.val;
  for (let i = 1; i < mData.length; i++) {
    if (mData[i][0] === raffleId) {
      masterSheet.getRange(i + 1, 5).setValue("closed");
      break;
    }
  }

  return { success: true };
}
