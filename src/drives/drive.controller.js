import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";
import {
  getSheet,
  appendRow,
  updateCell,
  updateRow,
  batchUpdateCells,
  getSheetCached,
  invalidateCache,
} from "../sheets/sheets.client.js";
import { getSheets } from "../sheets/sheets.dynamic.js";
import { ensurePlacementSheets } from "../utils/placementWorkbook.js";
import { getAcademicYear } from "../config/academicYear.js";
import { idxOf } from "../utils/sheetUtils.js";
import { ensureHeaders } from "../utils/sheetBootstrap.js";
import {
  DRIVE_REQUEST_HEADERS,
  COMPANY_DRIVES_HEADERS,
} from "../constants/sheets.js";
import { upsertCompanyDrivesEntry } from "./calendarUtils.js";
import { updatePlacementWorkbook } from "../analytics/updatePlacementWorkbook.js";
import { 
  getPlacementResults, 
  updatePlacementResults, 
  revokeStudentOffer 
} from "./placementResults.js";

import crypto from "crypto";

const router = express.Router();
const hasValue = (v) => v !== undefined && v !== null && v !== "";

/**
 * Helper: map headers → column index
 */
function getColumnMap(header) {
  const map = {};
  header.forEach((h, i) => {
    if (h) map[h.trim().toLowerCase()] = i;
  });
  return map;
}

/**
 * SPOC submits request
 */
router.post("/request", authenticate, roleGuard("SPOC"), async (req, res) => {
  const spoc = req.user.username;
  let {
    request_id,
    company,
    type,
    eligible_pool,
    cgpa_cutoff,
    ppt_datetime,
    ot_datetime,
    interview_datetime,
    internship_stipend,
    fte_ctc,
    fte_base,
    expected_hires,
  } = req.body;

  const rows = await getSheet("Drive_Requests");
  const header = rows[0];

  const idx = {
    request_id: idxOf(header, "Request ID"),
    company: idxOf(header, "Company"),
    spoc: idxOf(header, "SPOC"),
    type: idxOf(header, "Type"),
    eligible_pool: idxOf(header, "Eligible Pool"),
    cgpa: idxOf(header, "CGPA Cutoff"),

    ppt_dt: idxOf(header, "PPT Datetime"),
    ot_dt: idxOf(header, "OT Datetime"),
    interview_dt: idxOf(header, "Interview Datetime"),

    ppt_status: idxOf(header, "PPT Status"),
    ot_status: idxOf(header, "OT Status"),
    interview_status: idxOf(header, "INTERVIEW Status"),

    stipend: idxOf(header, "Internship Stipend"),
    fte_ctc: idxOf(header, "FTE CTC"),
    fte_base: idxOf(header, "FTE Base"),
    hires: idxOf(header, "Expected Hires"),
    drive_status: idxOf(header, "Drive Status"),
  };

  /* ---------- FIND BY request_id ONLY ---------- */
  let rowIndex = -1;
  if (request_id) {
    rowIndex = rows.findIndex(
      (r, i) => i > 0 && r[idx.request_id] === request_id
    );
  }

  /* ---------- CREATE (ONLY ONCE) ---------- */
  if (!request_id || rowIndex === -1) {
    request_id = crypto.randomUUID();

    const newRow = Array(header.length).fill("");

    newRow[idx.request_id] = request_id;
    newRow[idx.company] = company;
    newRow[idx.spoc] = spoc;
    newRow[idx.type] = type;
    newRow[idx.eligible_pool] = eligible_pool;
    newRow[idx.cgpa] = cgpa_cutoff;
    newRow[idx.stipend] = internship_stipend;
    newRow[idx.fte_ctc] = fte_ctc;
    newRow[idx.fte_base] = fte_base;
    newRow[idx.hires] = expected_hires;
    newRow[idx.drive_status] = "Scheduled";

    if (hasValue(ppt_datetime)) {
      newRow[idx.ppt_dt] = ppt_datetime;
      newRow[idx.ppt_status] = "PENDING";
    }

    if (hasValue(ot_datetime)) {
      newRow[idx.ot_dt] = ot_datetime;
      newRow[idx.ot_status] = "PENDING";
    }

    if (hasValue(interview_datetime)) {
      newRow[idx.interview_dt] = interview_datetime;
      newRow[idx.interview_status] = "PENDING";
    }

    await appendRow("Drive_Requests", newRow);
    return res.json({ request_id });
  }

  /* ---------- UPDATE EXISTING ---------- */
  const row = rowIndex + 1;

  // Collect all updates to batch them
  const cellUpdates = [];

  // ---- NON-DATE FIELDS ----
  const baseUpdates = [
    [idx.type, type],
    [idx.eligible_pool, eligible_pool],
    [idx.cgpa, cgpa_cutoff],
    [idx.stipend, internship_stipend],
    [idx.fte_ctc, fte_ctc],
    [idx.fte_base, fte_base],
    [idx.hires, expected_hires],
  ];
  
  for (const [col, val] of baseUpdates) {
    if (hasValue(val)) {
      cellUpdates.push({ row, col, value: val });
    }
  }

  /* ----------------------------------------------------
       🔥 FIXED STATUS LOGIC (MOST IMPORTANT PART)
       Status resets ONLY if date actually changes
    ---------------------------------------------------- */

  // ---- PPT ----
  if (hasValue(ppt_datetime)) {
    const prevDate = rows[rowIndex][idx.ppt_dt];
    const prevStatus = rows[rowIndex][idx.ppt_status];

    if (ppt_datetime !== prevDate) {
      cellUpdates.push({ row, col: idx.ppt_dt, value: ppt_datetime });
      if (prevStatus !== "APPROVED") {
        cellUpdates.push({ row, col: idx.ppt_status, value: "PENDING" });
      }
    }
  }

  // ---- OT ----
  if (hasValue(ot_datetime)) {
    const prevDate = rows[rowIndex][idx.ot_dt];
    const prevStatus = rows[rowIndex][idx.ot_status];

    if (ot_datetime !== prevDate) {
      cellUpdates.push({ row, col: idx.ot_dt, value: ot_datetime });
      if (prevStatus !== "APPROVED") {
        cellUpdates.push({ row, col: idx.ot_status, value: "PENDING" });
      }
    }
  }

  // ---- INTERVIEW ----
  if (hasValue(interview_datetime)) {
    const prevDate = rows[rowIndex][idx.interview_dt];
    const prevStatus = rows[rowIndex][idx.interview_status];
    if (interview_datetime !== prevDate) {
      cellUpdates.push({ row, col: idx.interview_dt, value: interview_datetime });
      if (prevStatus !== "APPROVED") {
        cellUpdates.push({ row, col: idx.interview_status, value: "PENDING" });
      }
    }
  }

  // Apply all updates in one batch operation
  if (cellUpdates.length > 0) {
    await batchUpdateCells("Drive_Requests", cellUpdates);
    invalidateCache("Drive_Requests");
  }

  res.json({ request_id });
});

/**
 * Calendar team views pending approvals
 */
router.get(
  "/pending",
  authenticate,
  roleGuard("CALENDAR_TEAM"),
  async (req, res) => {
    const rows = await getSheetCached("Drive_Requests", true); // Use cache
    const header = rows[0];
    const idx = (c) => idxOf(header, c);

    const hasDate = (v) => v && v !== "";
    const isPending = (date, status) =>
      hasDate(date) && (!status || status === "PENDING");

    const pending = rows
      .slice(1)
      .filter(
        (r) =>
          isPending(r[idx("PPT Datetime")], r[idx("PPT Status")]) ||
          isPending(r[idx("OT Datetime")], r[idx("OT Status")]) ||
          isPending(r[idx("Interview Datetime")], r[idx("INTERVIEW Status")])
      )
      .map((r) => ({
        request_id: r[idx("Request ID")],
        company: r[idx("Company")],
        ppt_datetime: r[idx("PPT Datetime")],
        ot_datetime: r[idx("OT Datetime")],
        interview_datetime: r[idx("Interview Datetime")],
        ppt_status: r[idx("PPT Status")] || "PENDING",
        ot_status: r[idx("OT Status")] || "PENDING",
        interview_status: r[idx("INTERVIEW Status")] || "PENDING",
      }));

    res.json({ pending });
  }
);

/**
 * Calendar team approves ONE slot
 */
router.post(
  "/approve",
  authenticate,
  roleGuard("CALENDAR_TEAM"),
  async (req, res) => {
    const { request_id, slot, action, suggested_datetime } = req.body;

    const rows = await getSheet("Drive_Requests");
    const header = rows[0];
    const col = (name) => idxOf(header, name);

    const rowIndex = rows.findIndex(
      (r, i) => i > 0 && r[col("Request ID")] === request_id
    );

    if (rowIndex === -1) {
      return res.status(404).json({ message: "Request not found" });
    }

    const sheetRow = rows[rowIndex];
    const row = rowIndex + 1;

    const company = sheetRow[col("Company")];
    const spoc = sheetRow[col("SPOC")];

    const dateColumnMap = {
      PPT: "PPT Datetime",
      OT: "OT Datetime",
      INTERVIEW: "Interview Datetime",
    };

    const statusColumn = `${slot} Status`;
    const suggestColumn = `${slot} Suggested Datetime`;

    const cellUpdates = [];

    /* ================= APPROVE ================= */
    if (action === "APPROVE") {
      const dateCol = dateColumnMap[slot];
      const approvedDate = sheetRow[col(dateCol)];

      if (approvedDate) {
        await upsertCompanyDrivesEntry({
          requestId: request_id,
          slot,
          sheetRow,
          header,
        });
      }

      cellUpdates.push({ row, col: col(statusColumn), value: "APPROVED" });
    }

    /* ================= REJECT ================= */
    if (action === "REJECT") {
      cellUpdates.push({ row, col: col(statusColumn), value: "REJECTED" });
    }

    /* ================= SUGGEST ================= */
    if (action === "SUGGEST") {
      cellUpdates.push({ row, col: col(statusColumn), value: "SUGGESTED" });
      cellUpdates.push({ row, col: col(suggestColumn), value: suggested_datetime });
    }

    // Batch update all changes
    if (cellUpdates.length > 0) {
      await batchUpdateCells("Drive_Requests", cellUpdates);
      invalidateCache("Drive_Requests");
    }

    res.json({ success: true });
  }
);

router.get(
  "/my",
  authenticate,
  roleGuard("SPOC", "CALENDAR_TEAM", "DATA_TEAM", "ADMIN"),
  async (req, res) => {
    const rows = await getSheetCached("Drive_Requests", true); // Use cache
    const header = rows[0];

    const idx = {
      request_id: idxOf(header, "Request ID"),
      company: idxOf(header, "Company"),
      type: idxOf(header, "Type"),
      eligible_pool: idxOf(header, "Eligible Pool"),
      cgpa: idxOf(header, "CGPA Cutoff"),

      ppt_dt: idxOf(header, "PPT Datetime"),
      ot_dt: idxOf(header, "OT Datetime"),
      interview_dt: idxOf(header, "Interview Datetime"),

      ppt_status: idxOf(header, "PPT Status"),
      ot_status: idxOf(header, "OT Status"),
      interview_status: idxOf(header, "INTERVIEW Status"),

      ppt_suggested: idxOf(header, "PPT Suggested Datetime"),
      ot_suggested: idxOf(header, "OT Suggested Datetime"),
      interview_suggested: idxOf(header, "INTERVIEW Suggested Datetime"),

      internship_stipend: idxOf(header, "Internship Stipend"),
      fte_ctc: idxOf(header, "FTE CTC"),
      fte_base: idxOf(header, "FTE Base"),
      expected_hires: idxOf(header, "Expected Hires"),

      drive_status: idxOf(header, "Drive Status"),
    };

    const drives = rows.slice(1).map((r) => ({
      request_id: r[idx.request_id],
      company: r[idx.company],
      type: r[idx.type],
      eligible_pool: r[idx.eligible_pool],
      cgpa_cutoff: r[idx.cgpa],

      ppt_datetime: r[idx.ppt_dt],
      ot_datetime: r[idx.ot_dt],
      interview_datetime: r[idx.interview_dt],

      ppt_status: r[idx.ppt_status],
      ot_status: r[idx.ot_status],
      interview_status: r[idx.interview_status],

      ppt_suggested_datetime: r[idx.ppt_suggested],
      ot_suggested_datetime: r[idx.ot_suggested],
      interview_suggested_datetime: r[idx.interview_suggested],

      internship_stipend: r[idx.internship_stipend],
      fte_ctc: r[idx.fte_ctc],
      fte_base: r[idx.fte_base],
      expected_hires: r[idx.expected_hires],

      drive_status:
        r[idx.drive_status] && r[idx.drive_status].trim()
          ? r[idx.drive_status]
          : "Scheduled",
    }));

    res.json({ drives });
  }
);

router.get(
  "/completed",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN"),
  async (req, res) => {
    const rows = await getSheetCached("Drive_Requests", true); // Use cache
    const header = rows[0];

    const idx = {
      company: header.indexOf("Company"),
      ppt_status: header.indexOf("PPT Status"),
      ot_status: header.indexOf("OT Status"),
      interview_status: header.indexOf("INTERVIEW Status"),
    };

    const completed = rows
      .slice(1)
      .filter(
        (r) =>
          r[idx.ppt_status] === "APPROVED" &&
          r[idx.ot_status] === "APPROVED" &&
          r[idx.interview_status] === "APPROVED"
      )
      .map((r) => ({
        company: r[idx.company],
        ppt_status: r[idx.ppt_status],
        ot_status: r[idx.ot_status],
        interview_status: r[idx.interview_status],
      }));

    res.json({ completed });
  }
);

router.post(
  "/status",
  authenticate,
  roleGuard("SPOC", "ADMIN"),
  async (req, res) => {
    const { request_id, status } = req.body;

    if (!request_id) {
      return res.status(400).json({ message: "request_id required" });
    }

    const rows = await getSheet("Drive_Requests");
    const header = rows[0];

    const idx = {
      request_id: idxOf(header, "Request ID"),
      drive_status: idxOf(header, "Drive Status"),
    };

    const rowIndex = rows.findIndex(
      (r, i) => i > 0 && r[idx.request_id] === request_id
    );

    if (rowIndex === -1) {
      return res.status(404).json({ message: "Drive not found" });
    }

    await updateCell("Drive_Requests", rowIndex + 1, idx.drive_status, status);

    // sync Company_Drives record to reflect the explicit drive status (and derived approval state)
    try {
      const rows = await getSheet("Drive_Requests");
      const header = rows[0];
      const sheetRow = rows[rowIndex];
      await upsertCompanyDrivesEntry({ requestId: request_id, slot: null, sheetRow, header });
    } catch (err) {
      console.warn("Failed to sync Company_Drives after drive status change:", err.message || err);
    }

    res.json({ success: true });
  }
);

router.post(
  "/results",
  authenticate,
  roleGuard("SPOC", "ADMIN"),
  async (req, res) => {
    const { request_id, results } = req.body;

    // ---------- VALIDATION ----------
    if (!request_id) {
      return res.status(400).json({ message: "request_id is required" });
    }

    if (typeof results !== "string" || !results.trim()) {
      return res.status(400).json({
        message: "results must be comma-separated roll numbers",
      });
    }

    const newRollNumbers = results
      .split(",")
      .map((r) => r.trim())
      .filter(Boolean);

    if (newRollNumbers.length === 0) {
      return res.status(400).json({ message: "No valid roll numbers" });
    }

    console.log(`[publish][start] actor=${req.user?.username || "unknown"} request_id=${request_id} rolls=${newRollNumbers.join(",")}`);

    // ---------- GET PREVIOUS RESULTS ----------
    const previousResults = await getPlacementResults(request_id);
    const previousRollNumbers = previousResults.rollNumbers;

    // Determine additions and removals
    const addedRollNumbers = newRollNumbers.filter(r => !previousRollNumbers.includes(r));
    const removedRollNumbers = previousRollNumbers.filter(r => !newRollNumbers.includes(r));

    console.log(`[publish] Added: ${addedRollNumbers.length}, Removed: ${removedRollNumbers.length}`);

    // ---------- LOAD ONLY DRIVE REQUESTS SHEET ----------
    const driveRows = await getSheet("Drive_Requests");
    const dHeader = driveRows[0];

    // ---------- INDEX MAPS ----------
    const dIdx = {
      request_id: idxOf(dHeader, "Request ID"),
      company: idxOf(dHeader, "Company"),
      spoc: idxOf(dHeader, "SPOC"),
      type: idxOf(dHeader, "Type"),
      eligible_pool: idxOf(dHeader, "Eligible Pool"),
      cgpa_cutoff: idxOf(dHeader, "CGPA Cutoff"),
      ppt_datetime: idxOf(dHeader, "PPT Datetime"),
      ot_datetime: idxOf(dHeader, "OT Datetime"),
      interview_datetime: idxOf(dHeader, "Interview Datetime"),
      ppt_status: idxOf(dHeader, "PPT Status"),
      ot_status: idxOf(dHeader, "OT Status"),
      interview_status: idxOf(dHeader, "INTERVIEW Status"),
      internship_stipend: idxOf(dHeader, "Internship Stipend"),
      fte_ctc: idxOf(dHeader, "FTE CTC"),
      fte_base: idxOf(dHeader, "FTE Base"),
      expected_hires: idxOf(dHeader, "Expected Hires"),
      drive_status: idxOf(dHeader, "Drive Status"),
    };

    // ---------- FIND DRIVE ----------
    const driveRow = driveRows.find(
      (r, i) => i > 0 && r[dIdx.request_id] === request_id
    );

    if (!driveRow) {
      return res.status(404).json({ message: "Drive not found" });
    }

    console.log(`[publish] Processing ${addedRollNumbers.length} additions and ${removedRollNumbers.length} removals`);

    const now = new Date().toISOString();

    function parseRoll(roll) {
      const s = String(roll || "");
      const yy = s.slice(0, 2);
      const year = 2000 + Number(yy);
      const branchCode = s.slice(2, 4);
      const degreeChar = s.slice(6, 7) || s.slice(4, 5);
      return { year, branchCode: (branchCode || "").toUpperCase(), degreeChar: (degreeChar || "").toUpperCase() };
    }

    // ---------- PROCESS ADDED STUDENTS IN PARALLEL ----------
    const processAddedStudent = async (roll) => {
      console.log(`[publish][roll] request_id=${request_id} roll=${roll} - parsing`);
      const parsed = parseRoll(roll);
      try {
        const degreeType = parsed.degreeChar === "M" ? "PG" : "UG";
        const academicYear = getAcademicYear();
        console.log(`[publish][roll] ${roll} -> admission=${parsed.year} academic=${academicYear} branch=${parsed.branchCode} degree=${degreeType}`);

        // Get placement workbook (single source of truth)
        const { spreadsheetId: placementId } = await ensurePlacementSheets({ program: degreeType, branch: parsed.branchCode });
        const sheetsApi = await getSheets();
        
        // Read from Students_<branch> sheet in placement workbook
        const studentsSheetName = `Students_${parsed.branchCode}`;
        const studentsRange = `${studentsSheetName}!A1:K1000`;
        const sres = await sheetsApi.spreadsheets.values.get({ 
          spreadsheetId: placementId, 
          range: studentsRange 
        });
        
        const svals = sres.data.values || [];
        console.log(`[publish][roll] ${roll} reading from ${studentsSheetName}, found ${svals.length} rows`);
        
        if (svals.length <= 1) {
          console.log(`[publish][roll][skip] ${roll} no students in ${studentsSheetName}`);
          return { success: false, roll };
        }

        const sheader = svals[0].map((h) => String(h || "").trim());
        const rollIdx = sheader.findIndex((h) => /roll/i.test(h));
        const nameIdx = sheader.findIndex((h) => /name/i.test(h));
        const branchIdx = sheader.findIndex((h) => /branch/i.test(h));
        const cgpaIdx = sheader.findIndex((h) => /cgpa/i.test(h));

        const studentRow = svals.slice(1).find((r) => (r[rollIdx] || "").toString().trim() === roll.toString().trim());
        if (!studentRow) {
          console.log(`[publish][roll][skip] ${roll} student not found in ${studentsSheetName}`);
          return { success: false, roll };
        }

        const studentName = studentRow[nameIdx] || "";
        const studentBranch = studentRow[branchIdx] || parsed.branchCode;
        const studentCgpa = studentRow[cgpaIdx] || "";

        console.log(`[publish][roll] ${roll} found student name=${studentName} branch=${studentBranch} cgpa=${studentCgpa}`);

        // Update placement workbook directly (single source of truth)
        await updatePlacementWorkbook({
          rollNo: roll,
          company: driveRow[dIdx.company],
          branch: parsed.branchCode,
          degreeType: degreeType,
          ctc: parseFloat(driveRow[dIdx.fte_ctc]) || 0,
          offerType: driveRow[dIdx.type] || "FTE",
          requestId: request_id,
          driveInfo: {
            spoc: driveRow[dIdx.spoc],
            driveType: driveRow[dIdx.type],
            eligiblePool: driveRow[dIdx.eligible_pool],
            pptDatetime: driveRow[dIdx.ppt_datetime],
            otDatetime: driveRow[dIdx.ot_datetime],
            interviewDatetime: driveRow[dIdx.interview_datetime],
            pptStatus: driveRow[dIdx.ppt_status],
            otStatus: driveRow[dIdx.ot_status],
            interviewStatus: driveRow[dIdx.interview_status],
            internshipStipend: driveRow[dIdx.internship_stipend],
            fteCTC: driveRow[dIdx.fte_ctc],
            fteBase: driveRow[dIdx.fte_base],
            expectedHires: driveRow[dIdx.expected_hires],
            driveStatus: driveRow[dIdx.drive_status],
            resultsPublished: true,
          },
        });
        console.log(`[placement] Updated placement workbook for ${roll}`);
        return { success: true, roll };
      } catch (err) {
        console.warn(`[publish][roll][error] Failed to process roll ${roll}:`, err.message || err);
        return { success: false, roll, error: err.message };
      }
    };

    // Process students in batches of 5 to avoid overwhelming the API
    const batchSize = 5;
    const addResults = [];
    for (let i = 0; i < addedRollNumbers.length; i += batchSize) {
      const batch = addedRollNumbers.slice(i, i + batchSize);
      const batchResults = await Promise.all(batch.map(processAddedStudent));
      addResults.push(...batchResults);
    }

    // ---------- PROCESS REMOVED STUDENTS IN PARALLEL ----------
    const processRemovedStudent = async (roll) => {
      console.log(`[publish][remove] Revoking offer for ${roll}`);
      const parsed = parseRoll(roll);
      const degreeType = parsed.degreeChar === "M" ? "PG" : "UG";
      
      try {
        // Get placement workbook
        const { spreadsheetId: placementId } = await ensurePlacementSheets({ 
          program: degreeType, 
          branch: parsed.branchCode 
        });
        
        // Revoke offer in Students sheet
        await revokeStudentOffer(placementId, parsed.branchCode, roll);
        console.log(`[publish][remove] Revoked offer for ${roll}`);
        return { success: true, roll };
      } catch (err) {
        console.warn(`[publish][remove] Failed to revoke for ${roll}:`, err.message);
        return { success: false, roll, error: err.message };
      }
    };

    const removeResults = await Promise.all(removedRollNumbers.map(processRemovedStudent));

    // ---------- UPDATE PLACEMENT_RESULTS SHEET ----------
    await updatePlacementResults(request_id, driveRow[dIdx.company], newRollNumbers);

    // ---------- LOCK DRIVE ----------
    await updateCell(
      "Drive_Requests",
      driveRows.indexOf(driveRow) + 1,
      dIdx.drive_status,
      "Completed"
    );
    invalidateCache("Drive_Requests");

    const successfulAdds = addResults.filter(r => r.success).length;
    const successfulRemoves = removeResults.filter(r => r.success).length;

    res.json({
      success: true,
      company: driveRow[dIdx.company],
      selected: newRollNumbers.length,
      added: successfulAdds,
      removed: successfulRemoves,
      failedAdds: addResults.filter(r => !r.success).map(r => r.roll),
      failedRemoves: removeResults.filter(r => !r.success).map(r => r.roll),
    });
  }
);

router.get(
  "/results/:request_id",
  authenticate,
  roleGuard("SPOC", "ADMIN"),
  async (req, res) => {
    const { request_id } = req.params;

    try {
      const placementResults = await getPlacementResults(request_id);
      
      res.json({
        results: placementResults.rollNumbers.join(", "),
        rollNumbers: placementResults.rollNumbers,
        count: placementResults.rollNumbers.length,
      });
    } catch (err) {
      console.error("[results/:request_id] Error:", err);
      res.status(500).json({ error: "Failed to fetch results" });
    }
  }
);

export default router;
