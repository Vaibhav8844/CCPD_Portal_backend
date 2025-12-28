// src/placements/placement.controller.js
import { getSheet } from "../sheets/sheets.client.js";
import { idxOf } from "../utils/sheetUtils.js";
import { ensurePlacementSheets } from "../utils/placementWorkbook.js";

export async function getCalendarData(req, res) {
  const rows = await getSheet("Company_Drives");
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
    const results = [];

    for (const branch of branches) {
      const result = await ensurePlacementSheets({
        year,
        program,
        branch,
      });

      results.push({
        branch,
        workbook: result.workbookName,
      });
    }

    res.json({
      success: true,
      message: "Batch enrolled successfully",
      data: results,
    });
  } catch (err) {
    console.error("Enroll batch failed:", err);
    res.status(500).json({
      message: "Failed to enroll batch",
    });
  }
}
