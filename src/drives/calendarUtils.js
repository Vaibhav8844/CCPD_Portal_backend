import { getSheet, appendRow, updateCell } from "../sheets/sheets.client.js";
import { idxOf } from "../utils/sheetUtils.js";
export async function upsertCompanyDrivesEntry({
  requestId,
  slot,              // "PPT" | "OT" | "INTERVIEW"
  sheetRow,          // row from Drive_Requests
  header,            // Drive_Requests header
}) {
  const rows = await getSheet("Company_Drives");
  const cdHeader = rows[0];

  const cdCol = (name) => idxOf(cdHeader, name);
  const drCol = (name) => idxOf(header, name);

  // ---- sanity check (once) ----
  [
    "Company",
    "SPOC",
    "Request ID",
    "Type",
    "Eligible Pool",
    "CGPA Cutoff",
    "PPT Datetime",
    "OT Datetime",
    "Interview Datetime",
    "PPT Status",
    "OT Status",
    "INTERVIEW Status",
    "Internship Stipend",
    "FTE CTC",
    "FTE Base",
    "Expected Hires",
    "Drive Status",
    "Last Updated",
  ].forEach((h) => {
    if (cdCol(h) === -1) {
      throw new Error(`Missing column in Company_Drives: ${h}`);
    }
  });

  // ---- find by Request ID ----
  let rowIndex = rows.findIndex(
    (r, i) => i > 0 && r[cdCol("Request ID")] === requestId
  );

  const now = new Date().toISOString();

  // ---- CREATE ----
  if (rowIndex === -1) {
    const newRow = Array(cdHeader.length).fill("");

    newRow[cdCol("Company")] =
      sheetRow[drCol("Company")] || "";

    newRow[cdCol("SPOC")] =
      sheetRow[drCol("SPOC")] || "";

    newRow[cdCol("Request ID")] = requestId;

    newRow[cdCol("Type")] =
      sheetRow[drCol("Type")] || "";

    newRow[cdCol("Eligible Pool")] =
      sheetRow[drCol("Eligible Pool")] || "";

    newRow[cdCol("CGPA Cutoff")] =
      sheetRow[drCol("CGPA Cutoff")] || "";

    newRow[cdCol("Internship Stipend")] =
      sheetRow[drCol("Internship Stipend")] || "";

    newRow[cdCol("FTE CTC")] =
      sheetRow[drCol("FTE CTC")] || "";

    newRow[cdCol("FTE Base")] =
      sheetRow[drCol("FTE Base")] || "";

    newRow[cdCol("Expected Hires")] =
      sheetRow[drCol("Expected Hires")] || "";

    newRow[cdCol("Drive Status")] =
      sheetRow[drCol("Drive Status")] || "Scheduled";

    // ---- slot-specific ----
    if (slot === "PPT") {
      newRow[cdCol("PPT Datetime")] =
        sheetRow[drCol("PPT Datetime")] || "";
      newRow[cdCol("PPT Status")] = "APPROVED";
    }

    if (slot === "OT") {
      newRow[cdCol("OT Datetime")] =
        sheetRow[drCol("OT Datetime")] || "";
      newRow[cdCol("OT Status")] = "APPROVED";
    }

    if (slot === "INTERVIEW") {
      newRow[cdCol("Interview Datetime")] =
        sheetRow[drCol("Interview Datetime")] || "";
      newRow[cdCol("INTERVIEW Status")] = "APPROVED";
    }

    newRow[cdCol("Last Updated")] = now;

    await appendRow("Company_Drives", newRow);
    return;
  }

  // ---- UPDATE ----
  const row = rowIndex + 1;

  if (slot === "PPT") {
    await updateCell(
      "Company_Drives",
      row,
      cdCol("PPT Datetime"),
      sheetRow[drCol("PPT Datetime")] || ""
    );
    await updateCell(
      "Company_Drives",
      row,
      cdCol("PPT Status"),
      "APPROVED"
    );
  }

  if (slot === "OT") {
    await updateCell(
      "Company_Drives",
      row,
      cdCol("OT Datetime"),
      sheetRow[drCol("OT Datetime")] || ""
    );
    await updateCell(
      "Company_Drives",
      row,
      cdCol("OT Status"),
      "APPROVED"
    );
  }

  if (slot === "INTERVIEW") {
    await updateCell(
      "Company_Drives",
      row,
      cdCol("Interview Datetime"),
      sheetRow[drCol("Interview Datetime")] || ""
    );
    await updateCell(
      "Company_Drives",
      row,
      cdCol("INTERVIEW Status"),
      "APPROVED"
    );
  }

  await updateCell(
    "Company_Drives",
    row,
    cdCol("Last Updated"),
    now
  );
}
