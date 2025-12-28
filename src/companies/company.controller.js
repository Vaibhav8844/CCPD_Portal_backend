import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";
import { appendRow, getSheet } from "../sheets/sheets.client.js";

const router = express.Router();

router.get(
  "/my",
  authenticate,
  roleGuard("SPOC", "CALENDAR_TEAM", "DATA_TEAM", "ADMIN"),
  async (req, res) => {
    const username = req.user.username.toLowerCase().trim();
    const rows = await getSheet("Company_SPOC_Map");

    const companies = rows
      .filter((r, i) => i > 0 && r[1])
      .filter(r => r[1].toLowerCase().trim() === username)
      .map(r => r[0]?.trim());

    res.json({ companies });
  }
);

router.post(
  "/assign",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN"),
  async (req, res) => {
    const { company, spoc_email } = req.body;

    if (!company || !spoc_email) {
      return res.status(400).json({ message: "Company and SPOC email required" });
    }

    const rows = await getSheet("Company_SPOC_Map");

    // Prevent duplicate assignment
    const exists = rows
      .slice(1)
      .some(
        r =>
          r[0]?.trim() === company.trim() &&
          r[1]?.trim() === spoc_email.trim()
      );

    if (exists) {
      return res.status(400).json({ message: "SPOC already assigned to this company" });
    }

    await appendRow("Company_SPOC_Map", [
      company.trim(),
      spoc_email.trim(),
    ]);

    res.json({ success: true });
  }
);

router.get(
  "/company-drives/:company",
  authenticate,
  roleGuard("SPOC"),
  async (req, res) => {
    const rows = await getSheet("Company_Drives");
    const header = rows[0];
    const idx = (c) => idxOf(header, c);

    const row = rows.find(
      (r, i) => i > 0 && r[idx("Company")] === req.params.company
    );

    if (!row) return res.json({ drive: null });

    res.json({
      drive: {
        company: row[idx("Company")],
        request_id: row[idx("Request ID")],
        type: row[idx("Type")],
        eligible_pool: row[idx("Eligible Pool")],
        cgpa_cutoff: row[idx("CGPA Cutoff")],
        ppt_datetime: row[idx("PPT Datetime")],
        ot_datetime: row[idx("OT Datetime")],
        interview_datetime: row[idx("Interview Datetime")],
        ppt_status: row[idx("PPT Status")],
        ot_status: row[idx("OT Status")],
        interview_status: row[idx("INTERVIEW Status")],
        internship_stipend: row[idx("Internship Stipend")],
        fte_ctc: row[idx("FTE CTC")],
        fte_base: row[idx("FTE Base")],
        expected_hires: row[idx("Expected Hires")],
        drive_status: row[idx("Drive Status")],
        results_published: row[idx("Results Published")],
      },
    });
  }
);



export default router;
