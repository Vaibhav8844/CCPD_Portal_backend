// src/placements/placement.controller.js
import { getSheetCached } from "../sheets/sheets.client.js";
import { idxOf } from "../utils/sheetUtils.js";
import { ensurePlacementSheets } from "../utils/placementWorkbook.js";

export async function getCalendarData(req, res) {
  const rows = await getSheetCached("Company_Drives", true); // Use cache
  const header = rows[0];

  const idx = {
    company: idxOf(header, "Company"),
    spoc: idxOf(header, "SPOC"),
    eligible: idxOf(header, "Eligible Pool"),
    ctc: idxOf(header, "FTE CTC"),
    base: idxOf(header, "FTE Base"),
    hires: idxOf(header, "Expected Hires"),
    ppt: idxOf(header, "PPT Datetime"),
    ot: idxOf(header, "OT Datetime"),
    interview: idxOf(header, "Interview Datetime"),
  };

  const data = rows.slice(1).map((r) => ({
    company: r[idx.company],
    spoc: r[idx.spoc],
    eligible_pool: r[idx.eligible],
    fte_ctc: r[idx.ctc],
    fte_base: r[idx.base],
    expected_hires: r[idx.hires],
    ppt: r[idx.ppt],
    ot: r[idx.ot],
    interview: r[idx.interview],
  }));

  res.json({ data });
}

export async function enrollBatch(req, res) {
  const { year, program, branches } = req.body;

  if (!year || !program || !Array.isArray(branches) || branches.length === 0) {
    return res.status(400).json({
      message: "Year, program and at least one branch are required",
    });
  }

  try {
    // Process all branches in parallel for faster enrollment
    const results = await Promise.all(
      branches.map(async (branch) => {
        try {
          const result = await ensurePlacementSheets({
            year,
            program,
            branch,
          });
          return {
            branch,
            workbook: result.workbookName,
            success: true,
          };
        } catch (err) {
          console.error(`Failed to enroll ${branch}:`, err);
          return {
            branch,
            success: false,
            error: err.message,
          };
        }
      })
    );

    const successful = results.filter(r => r.success);
    const failed = results.filter(r => !r.success);

    res.json({
      success: failed.length === 0,
      message: `Enrolled ${successful.length} branches${failed.length > 0 ? `, ${failed.length} failed` : ''}`,
      data: successful.map(r => ({ branch: r.branch, workbook: r.workbook })),
      errors: failed.map(r => ({ branch: r.branch, error: r.error })),
    });
  } catch (err) {
    console.error("Enroll batch failed:", err);
    res.status(500).json({
      message: "Failed to enroll batch",
    });
  }
}
