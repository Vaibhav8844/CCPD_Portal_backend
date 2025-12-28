import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";
import {
  getSheet,
  appendRow,
  updateCell,
  updateRow,
} from "../sheets/sheets.client.js";
import { getSheets } from "../sheets/sheets.dynamic.js";
import { getOrCreateStudentWorkbook } from "../utils/studentWorkbook.js";
import { ensurePlacementSheets } from "../utils/placementWorkbook.js";
import { getAcademicYear } from "../config/academicYear.js";
import { idxOf } from "../utils/sheetUtils.js";
import { ensureHeaders } from "../utils/sheetBootstrap.js";
import {
  DRIVE_REQUEST_HEADERS,
  COMPANY_DRIVES_HEADERS,
} from "../constants/sheets.js";
import { upsertCompanyDrivesEntry } from "../drives/calendarUtils.js";

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
      await updateCell("Drive_Requests", row, col, val);
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
      await updateCell("Drive_Requests", row, idx.ppt_dt, ppt_datetime);
      if (prevStatus !== "APPROVED") {
        await updateCell("Drive_Requests", row, idx.ppt_status, "PENDING");
      }
    }
  }

  // ---- OT ----
  if (hasValue(ot_datetime)) {
    const prevDate = rows[rowIndex][idx.ot_dt];
    const prevStatus = rows[rowIndex][idx.ot_status];

    if (ot_datetime !== prevDate) {
      await updateCell("Drive_Requests", row, idx.ot_dt, ot_datetime);
      if (prevStatus !== "APPROVED") {
        await updateCell("Drive_Requests", row, idx.ot_status, "PENDING");
      }
    }
  }

  // ---- INTERVIEW ----
  if (hasValue(interview_datetime)) {
    const prevDate = rows[rowIndex][idx.interview_dt];
    const prevStatus = rows[rowIndex][idx.interview_status];
    if (interview_datetime !== prevDate) {
      await updateCell(
        "Drive_Requests",
        row,
        idx.interview_dt,
        interview_datetime
      );
      if (prevStatus !== "APPROVED") {
        await updateCell(
          "Drive_Requests",
          row,
          idx.interview_status,
          "PENDING"
        );
      }
    }
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
    const rows = await getSheet("Drive_Requests");
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

      await updateCell("Drive_Requests", row, col(statusColumn), "APPROVED");
    }

    /* ================= REJECT ================= */
    if (action === "REJECT") {
      await updateCell("Drive_Requests", row, col(statusColumn), "REJECTED");
    }

    /* ================= SUGGEST ================= */
    if (action === "SUGGEST") {
      await updateCell("Drive_Requests", row, col(statusColumn), "SUGGESTED");

      await updateCell(
        "Drive_Requests",
        row,
        col(suggestColumn),
        suggested_datetime
      );
    }

    res.json({ success: true });
  }
);

router.get(
  "/my",
  authenticate,
  roleGuard("SPOC", "CALENDAR_TEAM", "DATA_TEAM", "ADMIN"),
  async (req, res) => {
    const rows = await getSheet("Drive_Requests");
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
    const rows = await getSheet("Drive_Requests");
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

    const rollNumbers = results
      .split(",")
      .map((r) => r.trim())
      .filter(Boolean);

    if (rollNumbers.length === 0) {
      return res.status(400).json({ message: "No valid roll numbers" });
    }

    console.log(`[publish][start] actor=${req.user?.username || "unknown"} request_id=${request_id} rolls=${rollNumbers.join(",")}`);

    // ---------- LOAD SHEETS ----------
    const driveRows = await getSheet("Drive_Requests");
    const studentRows = await getSheet("Students_Data");
    const placementRows = await getSheet("Placement_Data");

    const dHeader = driveRows[0];
    const sHeader = studentRows[0];
    const pHeader = placementRows[0];

    // ---------- INDEX MAPS ----------
    const dIdx = {
      request_id: idxOf(dHeader, "Request ID"),
      company: idxOf(dHeader, "Company"),
      type: idxOf(dHeader, "Type"),
      eligible_pool: idxOf(dHeader, "Eligible Pool"),
      internship_stipend: idxOf(dHeader, "Internship Stipend"),
      fte_ctc: idxOf(dHeader, "FTE CTC"),
      fte_base: idxOf(dHeader, "FTE Base"),
      drive_status: idxOf(dHeader, "Drive Status"),
    };

    const sIdx = {
      roll: idxOf(sHeader, "Roll Number"),
      name: idxOf(sHeader, "Student Name"),
      branch: idxOf(sHeader, "Branch"),
      cgpa: idxOf(sHeader, "CGPA"),
    };

    const pIdx = {
      request_id: idxOf(pHeader, "Request ID"),
      company: idxOf(pHeader, "Company"),
      drive_type: idxOf(pHeader, "Drive Type"),
      eligible_pool: idxOf(pHeader, "Eligible Pool"),
      internship_stipend: idxOf(pHeader, "Internship Stipend"),
      fte_ctc: idxOf(pHeader, "FTE CTC"),
      fte_base: idxOf(pHeader, "FTE Base"),
      roll: idxOf(pHeader, "Roll Number"),
      student_name: idxOf(pHeader, "Student Name"),
      branch: idxOf(pHeader, "Branch"),
      cgpa: idxOf(pHeader, "CGPA"),
      result: idxOf(pHeader, "Result"),
      published_at: idxOf(pHeader, "Published At"),
    };

    // ---------- FIND DRIVE ----------
    const driveRow = driveRows.find(
      (r, i) => i > 0 && r[dIdx.request_id] === request_id
    );

    if (!driveRow) {
      return res.status(404).json({ message: "Drive not found" });
    }

    console.log(`[publish] loading drive/student/placement sheets`);

    // ---------- APPEND RESULTS (per-roll lookup in per-year student workbook) ----------
    const now = new Date().toISOString();

    function parseRoll(roll) {
      const s = String(roll || "");
      const yy = s.slice(0, 2);
      const year = 2000 + Number(yy);
      const branchCode = s.slice(2, 4);
      const degreeChar = s.slice(6, 7) || s.slice(4, 5);
      return { year, branchCode: (branchCode || "").toUpperCase(), degreeChar: (degreeChar || "").toUpperCase() };
    }

    function findBranchSheetName(sheetsMeta, branchCode) {
      const candidates = (sheetsMeta || []).map((s) => s.properties && s.properties.title).filter(Boolean);
      const exact = candidates.find((t) => t.toLowerCase() === branchCode.toLowerCase());
      if (exact) return exact;
      const starts = candidates.find((t) => t.toLowerCase().startsWith(branchCode.toLowerCase()));
      if (starts) return starts;
      const inc = candidates.find((t) => t.toLowerCase().includes(branchCode.toLowerCase()));
      if (inc) return inc;
      return candidates.find((t) => t.toLowerCase().startsWith("students_")) || candidates[0];
    }

    for (const roll of rollNumbers) {
      console.log(`[publish][roll] request_id=${request_id} roll=${roll} - parsing`);
      const parsed = parseRoll(roll);
      try {
        const degreeType = parsed.degreeChar === "M" ? "PG" : "UG";
        // Use academic year from config for all placement workbooks
        const academicYear = getAcademicYear();
        console.log(`[publish][roll] ${roll} -> admission=${parsed.year} academic=${academicYear} branch=${parsed.branchCode} degree=${degreeType}`);

        // For student workbook, use the first 4 digits of academicYear (e.g., 2025-26 -> 2025)
        const studentYear = parseInt((academicYear || "2025").slice(0, 4), 10);
        const studentWorkbookId = await getOrCreateStudentWorkbook(studentYear, degreeType);
        console.log(`[publish][roll] ${roll} using studentWorkbookId=${studentWorkbookId}`);

        const sheetsApi = await getSheets();
        const meta = await sheetsApi.spreadsheets.get({ spreadsheetId: studentWorkbookId });
        const sheetName = findBranchSheetName(meta.data.sheets, parsed.branchCode);
        console.log(`[publish][roll] ${roll} matched sheetName=${sheetName}`);
        if (!sheetName) {
          console.log(`[publish][roll][skip] ${roll} no branch sheet found`);
          continue;
        }

        const studentsRange = `${sheetName}!A1:Z1000`;
        const sres = await sheetsApi.spreadsheets.values.get({ spreadsheetId: studentWorkbookId, range: studentsRange });
        const svals = sres.data.values || [];
        console.log(`[publish][roll] ${roll} student rows found=${svals.length}`);
        if (svals.length <= 1) {
          console.log(`[publish][roll][skip] ${roll} no student rows in sheet`);
          continue;
        }

        const sheader = svals[0].map((h) => String(h || "").trim());
        const rollIdx = sheader.findIndex((h) => /roll/i.test(h));
        const nameIdx = sheader.findIndex((h) => /name/i.test(h));
        const programIdx = sheader.findIndex((h) => /program|degree/i.test(h));
        const cgpaIdx = sheader.findIndex((h) => /cgpa/i.test(h));

        const studentRow = svals.slice(1).find((r) => (r[rollIdx] || "").toString().trim() === roll.toString().trim());
        if (!studentRow) {
          console.log(`[publish][roll][skip] ${roll} student not found in sheet`);
          continue;
        }

        const studentName = studentRow[nameIdx] || "";
        const studentBranch = studentRow[programIdx] || parsed.branchCode;
        const studentCgpa = studentRow[cgpaIdx] || "";
        const program = studentRow[programIdx] || degreeType;

        console.log(`[publish][roll] ${roll} found student name=${studentName} branch=${studentBranch} cgpa=${studentCgpa}`);

        const prow = Array(pHeader.length).fill("");
        prow[pIdx.request_id] = request_id;
        prow[pIdx.company] = driveRow[dIdx.company];
        prow[pIdx.drive_type] = driveRow[dIdx.type];
        prow[pIdx.eligible_pool] = driveRow[dIdx.eligible_pool];
        prow[pIdx.internship_stipend] = driveRow[dIdx.internship_stipend];
        prow[pIdx.fte_ctc] = driveRow[dIdx.fte_ctc];
        prow[pIdx.fte_base] = driveRow[dIdx.fte_base];
        prow[pIdx.roll] = roll;
        prow[pIdx.student_name] = studentName;
        prow[pIdx.branch] = studentBranch;
        prow[pIdx.cgpa] = studentCgpa;
        prow[pIdx.result] = "SELECTED";
        prow[pIdx.published_at] = now;

        try {
          await appendRow("Placement_Data", prow);
          console.log(`[publish][roll] ${roll} appended to Placement_Data (central workbook)`);
        } catch (err) {
          console.warn(`[publish][roll][error] Failed to append to Placement_Data:`, err.message || err);
        }

        try {
          const { spreadsheetId: placementId, workbookName } = await ensurePlacementSheets({ program, branch: parsed.branchCode });
          const offersSheet = `Offers_${parsed.branchCode}`;
          const offerRow = [roll, driveRow[dIdx.company], driveRow[dIdx.type], driveRow[dIdx.fte_ctc] || "", "SELECTED"];
          const placementSheetsApi = await getSheets();
          console.log(`[placement] Attempting to append offer for roll ${roll} to workbook '${workbookName}' (ID: ${placementId}), sheet '${offersSheet}'`);
          await placementSheetsApi.spreadsheets.values.append({
            spreadsheetId: placementId,
            range: `${offersSheet}!A1`,
            valueInputOption: "RAW",
            requestBody: { values: [offerRow] },
          });
          console.log(`[placement] Successfully appended offer for roll ${roll} to workbook '${workbookName}', sheet '${offersSheet}'`);
        } catch (err) {
          console.warn(`[placement][error] Failed to append offer for roll ${roll}:`, err.message || err);
        }
      } catch (err) {
        console.warn(`[publish][roll][error] Failed to process roll ${roll}:`, err.message || err);
      }
    }

    // ---------- LOCK DRIVE ----------
    await updateCell(
      "Drive_Requests",
      driveRows.indexOf(driveRow) + 1,
      dIdx.drive_status,
      "Completed"
    );

    res.json({
      success: true,
      company: driveRow[dIdx.company],
      selected: rollNumbers.length,
    });
  }
);

router.get(
  "/results/:request_id",
  authenticate,
  roleGuard("SPOC", "ADMIN"),
  async (req, res) => {
    const { request_id } = req.params;

    const rows = await getSheet("Placement_Data");
    const header = rows[0];

    const idx = {
      request_id: idxOf(header, "Request ID"),
      roll: idxOf(header, "Roll Number"),
      result: idxOf(header, "Result"),
    };

    const selectedRolls = rows
      .slice(1)
      .filter(
        (r) => r[idx.request_id] === request_id && r[idx.result] === "SELECTED"
      )
      .map((r) => r[idx.roll]);

    res.json({
      results: selectedRolls.join(", "),
    });
  }
);

export default router;
