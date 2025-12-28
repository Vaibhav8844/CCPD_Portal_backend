import { google } from "googleapis";
import { oauth2Client } from "../auth/googleAuth.js";
import { loadTokens } from "../utils/tokenStore.js";

function getAuthClient() {
  const tokens = loadTokens();
  if (!tokens) {
    throw new Error("Google not authenticated. Visit /auth/google");
  }

  oauth2Client.setCredentials(tokens);
  return oauth2Client;
}

export function getSheets() {
  return google.sheets({
    version: "v4",
    auth: getAuthClient(),
  });
}

export function getDrive() {
  return google.drive({
    version: "v3",
    auth: getAuthClient(),
  });
}
