import { google } from "googleapis";
import dotenv from "dotenv";

dotenv.config();

// ---------- CONFIG ----------
export const spreadsheetId = process.env.GOOGLE_SHEET_ID;

if (!spreadsheetId) {
  throw new Error("GOOGLE_SHEET_ID not set in environment variables");
}

// ---------- AUTH ----------
const auth = new google.auth.GoogleAuth({
  credentials: {
    // type: "service_account",
    // project_id: process.env.GOOGLE_PROJECT_ID,
    // private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
    // private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    // client_email: process.env.GOOGLE_CLIENT_EMAIL,
    // client_id: process.env.GOOGLE_CLIENT_ID,
    // auth_uri: process.env.GOOGLE_AUTH_URI,
    // token_uri: process.env.GOOGLE_TOKEN_URI,
    // auth_provider_x509_cert_url: process.env.GOOGLE_AUTH_PROVIDER_X509_CERT_URL,
    // client_x509_cert_url: process.env.GOOGLE_CLIENT_X509_CERT_URL,
  },
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// 🔥 EXPORT THIS
export const sheets = google.sheets({
  version: "v4",
  auth,
});

const DEFAULT_RANGE = "!A1:Z";

const normalizeRange = (sheetName) =>
  sheetName.includes("!") ? sheetName : `${sheetName}${DEFAULT_RANGE}`;

// ---------- HELPERS ----------
export async function getSheet(sheetName) {
  const range = normalizeRange(sheetName);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });

  return res.data.values || [];
}

export async function appendRow(sheetName, row) {
  const range = normalizeRange(sheetName);

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [row],
    },
  });
}

export async function updateCell(sheetName, row, col, value) {
  const range = `${sheetName}!${String.fromCharCode(65 + col)}${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    requestBody: {
      values: [[value]],
    },
  });
}

export async function updateRow(sheetName, row, values) {
  const endCol = String.fromCharCode(65 + values.length - 1);
  const range = `${sheetName}!A${row}:${endCol}${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    requestBody: {
      values: [values],
    },
  });
}
