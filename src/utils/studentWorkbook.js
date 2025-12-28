import { getSheets, getDrive } from "../sheets/sheets.dynamic.js";
import { STUDENT_HEADERS } from "../constants/studentHeaders.js";

/**
 * Get or create UG_Students_<YEAR> or PG_Students_<YEAR>
 */
export async function getOrCreateStudentWorkbook(year, degreeType) {
  const drive = getDrive();
  const name = `${degreeType}_Students_${year}`;

  const res = await drive.files.list({
    q: `name='${name}' and mimeType='application/vnd.google-apps.spreadsheet'`,
    fields: "files(id, name)",
  });

  if (res.data.files.length > 0) {
    return res.data.files[0].id;
  }

  const created = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.spreadsheet",
    },
  });

  return created.data.id;
}

/**
 * Ensure branch sheet exists with headers
 */
export async function ensureBranchSheet(spreadsheetId, branch) {
  const sheets = getSheets();
  const meta = await sheets.spreadsheets.get({ spreadsheetId });

  const exists = meta.data.sheets.some(
    (s) => s.properties.title === branch
  );

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: branch,
                gridProperties: { frozenRowCount: 1 },
              },
            },
          },
        ],
      },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${branch}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [STUDENT_HEADERS] },
    });
  }
}

/**
 * Append students to branch sheet
 */
export async function appendStudents(spreadsheetId, branch, rows) {
  const sheets = getSheets();

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${branch}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: rows },
  });
}
