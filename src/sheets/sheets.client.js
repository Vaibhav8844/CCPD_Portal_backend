import { google } from "googleapis";
import dotenv from "dotenv";
import fs from "fs";
import fsp from "fs/promises";
import path from "path";
import { oauth2Client } from "../auth/googleAuth.js";
import { loadTokens } from "../utils/tokenStore.js";

dotenv.config();

// ---------- SERVICE-ACCOUNT FALLBACK ----------
// Prepare service-account GoogleAuth options in case no user OAuth tokens are present.
const authOptions = {
  scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
};

if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  const saPath = path.resolve(process.cwd(), "service-account.json");
  if (fs.existsSync(saPath)) {
    try {
      const raw = fs.readFileSync(saPath, "utf8");
      const creds = JSON.parse(raw);
      if (creds && creds.client_email) {
        authOptions.credentials = creds;
        console.log("Using local service-account.json for GoogleAuth");
      } else {
        console.warn("service-account.json found but missing client_email field");
      }
    } catch (err) {
      console.warn("Failed to parse service-account.json:", err.message);
    }
  }
}

const serviceAuth = new google.auth.GoogleAuth(authOptions);

function getSheetsClient() {
  const tokens = loadTokens();
  if (tokens) {
    oauth2Client.setCredentials(tokens);
    return google.sheets({ version: "v4", auth: oauth2Client });
  }
  return google.sheets({ version: "v4", auth: serviceAuth });
}

const DEFAULT_RANGE = "!A1:Z";

const normalizeRange = (sheetName) =>
  sheetName.includes("!") ? sheetName : `${sheetName}${DEFAULT_RANGE}`;

const stateFile = path.resolve(process.cwd(), "data", "calendar_state.json");

// Return the calendar workbook id if initialized, otherwise null.
async function getSpreadsheetId() {
  try {
    const raw = await fsp.readFile(stateFile, "utf8");
    const state = JSON.parse(raw);
    if (state && state.workbookId) return state.workbookId;
  } catch (err) {
    // ignore - treat as not initialized
  }
  return null;
}

export { getSpreadsheetId };

// ---------- HELPERS ----------
export async function getSheet(sheetName) {
  const range = normalizeRange(sheetName);
  const spreadsheetId = await getSpreadsheetId();

  if (!spreadsheetId) throw new Error("No calendar workbook initialized. Initialize the calendar first.");

  const sheetsClient = getSheetsClient();
  const res = await sheetsClient.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values || [];
}

export async function appendRow(sheetName, row) {
  const range = normalizeRange(sheetName);
  const spreadsheetId = await getSpreadsheetId();

  if (!spreadsheetId) throw new Error("No calendar workbook initialized. Initialize the calendar first.");

  const sheetsClient = getSheetsClient();
  await sheetsClient.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [row] },
  });
}

export async function updateCell(sheetName, row, col, value) {
  const range = `${sheetName}!${String.fromCharCode(65 + col)}${row}`;
  const spreadsheetId = await getSpreadsheetId();

  if (!spreadsheetId) throw new Error("No calendar workbook initialized. Initialize the calendar first.");

  const sheetsClient = getSheetsClient();
  await sheetsClient.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[value]] },
  });
}

export async function updateRow(sheetName, row, values) {
  const endCol = String.fromCharCode(65 + values.length - 1);
  const range = `${sheetName}!A${row}:${endCol}${row}`;
  const spreadsheetId = await getSpreadsheetId();

  if (!spreadsheetId) throw new Error("No calendar workbook initialized. Initialize the calendar first.");

  const sheetsClient = getSheetsClient();
  await sheetsClient.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [values] },
  });
}

/**
 * Delete the first row in `sheetName` where any cell (trimmed, lowercased)
 * matches the provided `matchValue` (case-insensitive). Does not delete header row.
 */
export async function deleteRowByEmail(sheetName, matchValue) {
  const spreadsheetId = await getSpreadsheetId();
  if (!spreadsheetId) throw new Error("No spreadsheet id available");

  // fetch sheet metadata to get sheetId
  const sheetsClient = getSheetsClient();
  const meta = await sheetsClient.spreadsheets.get({ spreadsheetId });
  const sheet = (meta.data.sheets || []).find((s) => s.properties && s.properties.title === sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
  const sheetId = sheet.properties.sheetId;

  // fetch values
  const r = await sheetsClient.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A1:Z1000` });
  const vals = r.data.values || [];
  if (vals.length <= 1) return { deleted: false, reason: "No data rows" };

  const target = String((matchValue || "").trim()).toLowerCase();
  let rowIndex = -1;
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i] || [];
    for (const cell of row) {
      if (String(cell || "").trim().toLowerCase() === target) {
        rowIndex = i; // 0-based index into vals; sheet row number = i+1
        break;
      }
    }
    if (rowIndex !== -1) break;
  }

  if (rowIndex === -1) return { deleted: false, reason: "Not found" };

  // Prevent deleting header row (index 0)
  const startIndex = rowIndex; // zero-based index in sheet rows

  await sheetsClient.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: { sheetId, dimension: "ROWS", startIndex: startIndex, endIndex: startIndex + 1 },
          },
        },
      ],
    },
  });

  return { deleted: true };
}
