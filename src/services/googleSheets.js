import { google } from "googleapis";
import { oauth2Client } from "../auth/googleAuth.js";
import { loadTokens } from "../utils/tokenStore.js";

function getAuthorizedClient() {
  const tokens = loadTokens();
  if (!tokens) {
    throw new Error("Google not authenticated. Visit /auth/google");
  }
  oauth2Client.setCredentials(tokens);
  return oauth2Client;
}

// CREATE SPREADSHEET (USES YOUR DRIVE)
export async function createSpreadsheet(name) {
  const auth = getAuthorizedClient();

  const drive = google.drive({ version: "v3", auth });

  const res = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.spreadsheet",
    },
    fields: "id",
  });

  return res.data.id;
}

// APPEND DATA
export async function appendRows(spreadsheetId, sheetName, values) {
  const auth = getAuthorizedClient();

  const sheets = google.sheets({ version: "v4", auth });

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values },
  });
}
