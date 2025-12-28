import { getSheets, getDrive } from "../sheets/sheets.dynamic.js";
import { STUDENT_HEADERS } from "../constants/studentHeaders.js";

/**
 * Normalize branch code to 2 letters (CS, EC, EE, ME)
 */
function normalizeBranchCode(branch) {
  const code = String(branch || "").toUpperCase().trim();
  // If already 2 letters, return as is
  if (code.length === 2) return code;
  // If 3 letters (CSE, ECE, etc.), take first 2
  if (code.length === 3) return code.slice(0, 2);
  // Otherwise return as is
  return code;
}

/**
 * Get or create UG_Students_<ACADEMIC_YEAR> or PG_Students_<ACADEMIC_YEAR>
 */
export async function getOrCreateStudentWorkbook(academicYear, degreeType) {
  const drive = await getDrive();
  const name = `${degreeType}_Students_${academicYear}`;

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
  const sheets = await getSheets();
  const normalizedBranch = normalizeBranchCode(branch);
  const meta = await sheets.spreadsheets.get({ spreadsheetId });

  const exists = meta.data.sheets.some(
    (s) => s.properties.title === normalizedBranch
  );

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: normalizedBranch,
                gridProperties: { frozenRowCount: 1 },
              },
            },
          },
        ],
      },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${normalizedBranch}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [STUDENT_HEADERS] },
    });
  }
}

/**
 * Append students to branch sheet (with duplicate prevention)
 */
export async function appendStudents(spreadsheetId, branch, rows) {
  const sheets = await getSheets();
  const normalizedBranch = normalizeBranchCode(branch);

  // Get existing roll numbers to prevent duplicates
  const existingData = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${normalizedBranch}!A:A`,
  });

  const existingRollNumbers = new Set(
    (existingData.data.values || []).slice(1).flat().filter(Boolean)
  );

  // Filter out duplicate students
  const newRows = rows.filter(row => {
    const rollNo = row[0];
    if (existingRollNumbers.has(rollNo)) {
      console.log(`[appendStudents] Skipping duplicate student: ${rollNo}`);
      return false;
    }
    return true;
  });

  if (newRows.length === 0) {
    console.log(`[appendStudents] No new students to add (all duplicates)`);
    return { added: 0, skipped: rows.length };
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${normalizedBranch}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: newRows },
  });

  console.log(`[appendStudents] Added ${newRows.length} students, skipped ${rows.length - newRows.length} duplicates`);
  return { added: newRows.length, skipped: rows.length - newRows.length };
}
