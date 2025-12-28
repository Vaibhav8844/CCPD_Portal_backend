import { getSheets, getDrive } from "../sheets/sheets.dynamic.js";
import { SHEET_TEMPLATES } from "../constants/sheets.js";

export async function ensurePlacementSheets({ year, program, branch }) {
  const sheets = getSheets();
  const drive = getDrive();

  const workbookName = `Placement_Data_${year}_${program}`;
  const branchCode = branch.toUpperCase();

  const spreadsheetId = await getOrCreateWorkbook(drive, workbookName);

  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const existingSheets = meta.data.sheets.map(
    (s) => s.properties.title
  );

  const requiredSheets = [
    { name: `Students_${branchCode}`, headers: SHEET_TEMPLATES.students },
    { name: `Offers_${branchCode}`, headers: SHEET_TEMPLATES.offers },
    { name: `Company_Drives_${branchCode}`, headers: SHEET_TEMPLATES.companyDrives },
    { name: `Placement_Stats_${branchCode}`, headers: SHEET_TEMPLATES.placementStats },
    { name: `CTC_Distribution_${branchCode}`, headers: SHEET_TEMPLATES.ctcDistribution },
  ];

  for (const sheet of requiredSheets) {
    if (!existingSheets.includes(sheet.name)) {
      await createSheetWithHeader(
        sheets,
        spreadsheetId,
        sheet.name,
        sheet.headers
      );
    }
  }

  return { spreadsheetId, workbookName };
}

/* ---------- HELPERS ---------- */

async function getOrCreateWorkbook(drive, name) {
  const folderId = process.env.PLACEMENT_FOLDER_ID;
  if (!folderId) throw new Error("PLACEMENT_FOLDER_ID not set");

  const res = await drive.files.list({
    q: `name='${name}' and mimeType='application/vnd.google-apps.spreadsheet' and '${folderId}' in parents`,
    fields: "files(id, name)",
  });

  if (res.data.files.length > 0) {
    return res.data.files[0].id;
  }

  const create = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [folderId],
    },
  });

  return create.data.id;
}

async function createSheetWithHeader(sheets, spreadsheetId, title, headers) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ addSheet: { properties: { title } } }],
    },
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${title}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: [headers] },
  });
}
