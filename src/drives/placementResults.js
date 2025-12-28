import { getSheets } from "../sheets/sheets.dynamic.js";
import { getSpreadsheetId } from "../sheets/sheets.client.js";
import { ensureHeaders } from "../utils/sheetBootstrap.js";

const PLACEMENT_RESULTS_HEADERS = ["Company", "Request ID", "Roll Numbers", "Last Updated"];

/**
 * Ensure Placement_Results sheet exists in CCPD Calendar workbook
 */
async function ensurePlacementResultsSheet() {
  const calendarId = await getSpreadsheetId();
  await ensureHeaders("Placement_Results", PLACEMENT_RESULTS_HEADERS, calendarId);
  return calendarId;
}

/**
 * Get existing roll numbers for a company from Placement_Results sheet
 */
export async function getPlacementResults(requestId) {
  const sheets = await getSheets();
  const calendarId = await ensurePlacementResultsSheet();

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: calendarId,
    range: "Placement_Results!A:D",
  });

  const rows = result.data.values || [];
  const row = rows.find((r, i) => i > 0 && r[1] === requestId);

  if (!row) {
    return { exists: false, rollNumbers: [] };
  }

  const rollNumbers = row[2] ? row[2].split(",").map(r => r.trim()).filter(Boolean) : [];
  return { exists: true, rollNumbers, rowIndex: rows.indexOf(row) };
}

/**
 * Update Placement_Results sheet with new roll numbers
 */
export async function updatePlacementResults(requestId, company, rollNumbers) {
  const sheets = await getSheets();
  const calendarId = await ensurePlacementResultsSheet();

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: calendarId,
    range: "Placement_Results!A:D",
  });

  const rows = result.data.values || [];
  const rowIndex = rows.findIndex((r, i) => i > 0 && r[1] === requestId);

  const rollNumbersStr = rollNumbers.join(", ");
  const now = new Date().toISOString();

  if (rowIndex > 0) {
    // Update existing row
    await sheets.spreadsheets.values.update({
      spreadsheetId: calendarId,
      range: `Placement_Results!A${rowIndex + 1}:D${rowIndex + 1}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[company, requestId, rollNumbersStr, now]],
      },
    });
  } else {
    // Append new row
    await sheets.spreadsheets.values.append({
      spreadsheetId: calendarId,
      range: "Placement_Results!A:D",
      valueInputOption: "RAW",
      requestBody: {
        values: [[company, requestId, rollNumbersStr, now]],
      },
    });
  }

  console.log(`[PlacementResults] Updated ${company} with ${rollNumbers.length} students`);
}

/**
 * Revoke offer for a student in Students_<branch> sheet
 */
export async function revokeStudentOffer(workbookId, branchCode, rollNo) {
  const sheets = await getSheets();
  const sheetName = `Students_${branchCode}`;

  try {
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: workbookId,
      range: `${sheetName}!A:K`,
    });

    const rows = result.data.values || [];
    const studentRowIndex = rows.findIndex((r, i) => i > 0 && r[0] === rollNo);

    if (studentRowIndex === -1) {
      console.warn(`[revokeOffer] Student ${rollNo} not found in ${sheetName}`);
      return;
    }

    // Clear placement fields (G-K): Placement Status, Placement Type, Company, Highest CTC, Offer Revoked
    await sheets.spreadsheets.values.update({
      spreadsheetId: workbookId,
      range: `${sheetName}!G${studentRowIndex + 1}:K${studentRowIndex + 1}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [["", "", "", "", ""]],
      },
    });

    console.log(`[revokeOffer] Revoked offer for ${rollNo} in ${sheetName}`);
  } catch (err) {
    console.error(`[revokeOffer] Error:`, err.message);
  }
}
