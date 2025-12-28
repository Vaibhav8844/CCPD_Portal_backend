import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";
import {
  getOverallStats,
  getBranchwiseStats,
  getCTCDistribution,
  getDemographicSplit,
  getTrendAnalysis,
  getCompanyStats,
  getPlacementSnapshot,
  recalculateAllPlacementWorkbooks,
} from "../analytics/stats.service.js";
import { recalculateBranchStats } from "../analytics/updatePlacementWorkbook.js";
import { getAcademicYear } from "../config/academicYear.js";
import { getDrive, getSheets } from "../sheets/sheets.dynamic.js";

const router = express.Router();

// Get overall statistics
router.get(
  "/overall",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const stats = await getOverallStats();
      res.json(stats);
    } catch (err) {
      console.error("[analytics] /overall error:", err);
      res.status(500).json({ error: "Failed to fetch overall stats" });
    }
  }
);

// Get branch-wise statistics (default UG)
router.get(
  "/branchwise",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const stats = await getBranchwiseStats("UG");
      res.json(stats);
    } catch (err) {
      console.error("[analytics] /branchwise error:", err);
      res.status(500).json({ error: "Failed to fetch branch-wise stats" });
    }
  }
);

// Get branch-wise statistics by degree type
router.get(
  "/branchwise/:degreeType",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const degreeType = req.params.degreeType;
      const stats = await getBranchwiseStats(degreeType);
      res.json(stats);
    } catch (err) {
      console.error("[analytics] /branchwise error:", err);
      res.status(500).json({ error: "Failed to fetch branch-wise stats" });
    }
  }
);

// Get CTC distribution
router.get(
  "/ctc-distribution",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const distribution = await getCTCDistribution();
      res.json(distribution);
    } catch (err) {
      console.error("[analytics] /ctc-distribution error:", err);
      res.status(500).json({ error: "Failed to fetch CTC distribution" });
    }
  }
);

// Get demographic split
router.get(
  "/demographic",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const demographic = await getDemographicSplit();
      res.json(demographic);
    } catch (err) {
      console.error("[analytics] /demographic error:", err);
      res.status(500).json({ error: "Failed to fetch demographic data" });
    }
  }
);

// Get trend analysis
router.get(
  "/trends",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const trends = await getTrendAnalysis();
      res.json(trends);
    } catch (err) {
      console.error("[analytics] /trends error:", err);
      res.status(500).json({ error: "Failed to fetch trend analysis" });
    }
  }
);

// Get company-wise statistics
router.get(
  "/companies",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const companies = await getCompanyStats();
      res.json(companies);
    } catch (err) {
      console.error("[analytics] /companies error:", err);
      res.status(500).json({ error: "Failed to fetch company stats" });
    }
  }
);

// Get complete placement snapshot
router.get(
  "/snapshot",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const snapshot = await getPlacementSnapshot();
      res.json(snapshot);
    } catch (err) {
      console.error("[analytics] /snapshot error:", err);
      res.status(500).json({ error: "Failed to fetch placement snapshot" });
    }
  }
);

// Recalculate all placement statistics
router.post(
  "/recalculate",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const academicYear = getAcademicYear();
      const drive = await getDrive();
      const sheets = await getSheets();
      
      const results = {
        academicYear,
        recalculated: [],
        errors: [],
      };

      // Recalculate for both UG and PG
      for (const degreeType of ["UG", "PG"]) {
        const workbookName = `Placement_Data_${academicYear}_${degreeType}`;
        
        const res = await drive.files.list({
          q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
          fields: "files(id, name)",
        });

        if (res.data.files.length === 0) {
          console.log(`[recalculate] No workbook found: ${workbookName}`);
          continue;
        }

        const workbookId = res.data.files[0].id;

        // Get all branches with Students sheets
        const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
        const studentSheets = meta.data.sheets
          .map(s => s.properties.title)
          .filter(name => name.startsWith("Students_"));

        for (const sheetName of studentSheets) {
          const branchCode = sheetName.replace("Students_", "");
          
          try {
            await recalculateBranchStats(workbookId, branchCode, degreeType, academicYear);
            results.recalculated.push(`${degreeType} - ${branchCode}`);
          } catch (err) {
            console.error(`[recalculate] Error for ${degreeType}-${branchCode}:`, err.message);
            results.errors.push({
              branch: `${degreeType} - ${branchCode}`,
              error: err.message,
            });
          }
        }
      }

      res.json({
        success: true,
        message: `Recalculated stats for ${results.recalculated.length} branches`,
        ...results,
      });
    } catch (err) {
      console.error("[analytics] /recalculate error:", err);
      res.status(500).json({ error: "Failed to recalculate statistics" });
    }
  }
);

// Force recalculate all formulas in placement workbooks
router.post(
  "/recalculate-formulas",
  authenticate,
  roleGuard("ADMIN", "CALENDAR_TEAM", "DATA_TEAM"),
  async (req, res) => {
    try {
      const results = await recalculateAllPlacementWorkbooks();
      
      const successCount = results.filter(r => r.success).length;
      const failCount = results.filter(r => !r.success).length;
      
      res.json({
        success: failCount === 0,
        message: `Recalculated ${successCount} workbooks${failCount > 0 ? `, ${failCount} failed` : ''}`,
        results
      });
    } catch (err) {
      console.error("[analytics] /recalculate-formulas error:", err);
      res.status(500).json({ error: "Failed to recalculate formulas" });
    }
  }
);

export default router;
