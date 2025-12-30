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

let serviceAuth;

// Production: Use environment variables
if (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
  console.log("[sheets] Using service account from environment variables");
  const credentials = {
    type: "service_account",
    project_id: process.env.GOOGLE_PROJECT_ID || "placement-system",
    private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    client_id: process.env.GOOGLE_SA_CLIENT_ID,
    auth_uri: "https://accounts.google.com/o/oauth2/auth",
    token_uri: "https://oauth2.googleapis.com/token",
    auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
  };
  authOptions.credentials = credentials;
  serviceAuth = new google.auth.GoogleAuth(authOptions);
} 
// Development: Use service-account.json file or GOOGLE_APPLICATION_CREDENTIALS
else if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  const saPath = path.resolve(process.cwd(), "service-account.json");
  if (fs.existsSync(saPath)) {
    try {
      const raw = fs.readFileSync(saPath, "utf8");
      const creds = JSON.parse(raw);
      if (creds && creds.client_email) {
        authOptions.credentials = creds;
        console.log("[sheets] Using local service-account.json for GoogleAuth");
      } else {
        console.warn("service-account.json found but missing client_email field");
      }
    } catch (err) {
      console.warn("Failed to parse service-account.json:", err.message);
    }
  }
  serviceAuth = new google.auth.GoogleAuth(authOptions);
} else {
  serviceAuth = new google.auth.GoogleAuth(authOptions);
}

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
  // Priority 1: Use environment variable (for production)
  if (process.env.CALENDAR_WORKBOOK_ID) {
    return process.env.CALENDAR_WORKBOOK_ID;
  }
  
  // Priority 2: Use state file (for development)
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
 * Batch update multiple cells at once to reduce API calls
 * @param {string} sheetName - The sheet name
 * @param {Array<{row: number, col: number, value: any}>} updates - Array of updates
 */
export async function batchUpdateCells(sheetName, updates) {
  if (!updates || updates.length === 0) return;
  
  const spreadsheetId = await getSpreadsheetId();
  if (!spreadsheetId) throw new Error("No calendar workbook initialized. Initialize the calendar first.");

  const sheetsClient = getSheetsClient();
  const data = updates.map(({ row, col, value }) => ({
    range: `${sheetName}!${String.fromCharCode(65 + col)}${row}`,
    values: [[value]],
  }));

  await sheetsClient.spreadsheets.values.batchUpdate({
    spreadsheetId,
    valueInputOption: "RAW",
    requestBody: { data },
  });
}

// Simple in-memory cache for sheet data
const sheetCache = new Map();
const CACHE_TTL = 30 * 1000; // 30 seconds

/**
 * Get sheet with optional caching
 * @param {string} sheetName - The sheet name
 * @param {boolean} useCache - Whether to use cache (default: false)
 */
export async function getSheetCached(sheetName, useCache = false) {
  if (useCache) {
    const cached = sheetCache.get(sheetName);
    if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
      return cached.data;
    }
  }
  
  const data = await getSheet(sheetName);
  
  if (useCache) {
    sheetCache.set(sheetName, { data, timestamp: Date.now() });
  }
  
  return data;
}

export function invalidateCache(sheetName) {
  if (sheetName) {
    sheetCache.delete(sheetName);
  } else {
    sheetCache.clear();
  }
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
