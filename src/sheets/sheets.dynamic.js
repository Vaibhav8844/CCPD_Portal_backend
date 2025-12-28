import { google } from "googleapis";
import { oauth2Client } from "../auth/googleAuth.js";
import { loadTokens } from "../utils/tokenStore.js";
import fs from "fs";
import path from "path";

async function getAuthClient() {
  const tokens = loadTokens();
  if (tokens) {
    oauth2Client.setCredentials(tokens);
    return oauth2Client;
  }

  // Try service account fallback via service-account.json or ADC
  const saPath = path.resolve(process.cwd(), "service-account.json");
  if (fs.existsSync(saPath)) {
    const raw = fs.readFileSync(saPath, "utf8");
    const creds = JSON.parse(raw);
    const auth = new google.auth.GoogleAuth({
      credentials: creds,
      scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
    });
    return await auth.getClient();
  }

  if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    const auth = new google.auth.GoogleAuth({
      scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
    });
    return await auth.getClient();
  }

  throw new Error("Google not authenticated. Visit /auth/google");
}

export async function getSheets() {
  const auth = await getAuthClient();
  return google.sheets({ version: "v4", auth });
}

export async function getDrive() {
  const auth = await getAuthClient();
  return google.drive({ version: "v3", auth });
}
