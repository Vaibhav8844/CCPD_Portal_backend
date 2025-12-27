import { google } from "googleapis";
import { sheets, spreadsheetId } from "../sheets/sheets.client.js";

/**
 * Ensure sheet exists, else create
 */
export async function ensureSheet(sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });

  const exists = meta.data.sheets.some(
    (s) => s.properties.title === sheetName
  );

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          { addSheet: { properties: { title: sheetName } } },
        ],
      },
    });
  }
}

/**
 * Ensure headers exist (idempotent)
 */
export async function ensureHeaders(sheetName, expectedHeaders) {
  await ensureSheet(sheetName);

  const range = `${sheetName}!A1:Z1`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });

  const existing = res.data.values?.[0] || [];

  // Normalize for comparison
  const normalize = (s) =>
    String(s).replace(/\u00A0/g, " ").trim().toLowerCase();

  const existingNorm = existing.map(normalize);

  // Find missing headers
  const missing = expectedHeaders.filter(
    (h) => !existingNorm.includes(normalize(h))
  );

  // If no headers exist at all → write full header row
  if (existing.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [expectedHeaders] },
    });
    return;
  }

  // If some headers are missing → append them
  if (missing.length > 0) {
    const updated = [...existing, ...missing];

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [updated] },
    });
  }
}
