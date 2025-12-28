import { getSheets } from "../sheets/sheets.dynamic.js";

/* --------- HELPERS --------- */

function mean(arr) {
  if (!arr.length) return 0;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

function median(arr) {
  if (!arr.length) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2
    ? sorted[mid]
    : (sorted[mid - 1] + sorted[mid]) / 2;
}

/* --------- MAIN ANALYTICS --------- */

export async function getPlacementOverview({ spreadsheetId }) {
  const sheets = getSheets();

  /* 1️⃣ Get sheet metadata */
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheetTitles = meta.data.sheets.map(
    (s) => s.properties.title
  );

  const studentSheets = sheetTitles.filter((s) =>
    s.startsWith("Students_")
  );
  const offerSheets = sheetTitles.filter((s) =>
    s.startsWith("Offers_")
  );

  let totalStudents = 0;
  let branchStats = {};
  let placedRolls = new Set();
  let ctcs = [];

  /* 2️⃣ Read Students sheets */
  for (const sheet of studentSheets) {
    const branch = sheet.replace("Students_", "");

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheet}!A2:Z`,
    });

    const rows = res.data.values || [];
    totalStudents += rows.length;

    branchStats[branch] = {
      branch,
      total: rows.length,
      placed: 0,
    };
  }

  /* 3️⃣ Read Offers sheets */
  for (const sheet of offerSheets) {
    const branch = sheet.replace("Offers_", "");

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheet}!A2:Z`,
    });

    const rows = res.data.values || [];

    for (const row of rows) {
      const roll = row[0];          // Roll No
      const ctc = Number(row[2]);   // CTC

      if (roll) placedRolls.add(roll);
      if (!isNaN(ctc)) ctcs.push(ctc);

      if (branchStats[branch]) {
        branchStats[branch].placed++;
      }
    }
  }

  /* 4️⃣ Final metrics */
  const placedStudents = placedRolls.size;

  const overview = {
    totalStudents,
    placedStudents,
    placementPercentage:
      totalStudents === 0
        ? 0
        : Number(
            ((placedStudents / totalStudents) * 100).toFixed(2)
          ),
    avgCTC: Number(mean(ctcs).toFixed(2)),
    medianCTC: Number(median(ctcs).toFixed(2)),
    highestCTC: ctcs.length ? Math.max(...ctcs) : 0,
    branchWise: Object.values(branchStats),
  };

  return overview;
}
