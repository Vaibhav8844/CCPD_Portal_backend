import express from "express";
import { getPlacementOverview } from "../analytics/overview.analytics.js";

const router = express.Router();

/**
 * GET /api/analytics/overview?spreadsheetId=XXXX
 */
router.get("/overview", async (req, res) => {
  try {
    const { spreadsheetId } = req.query;

    if (!spreadsheetId) {
      return res
        .status(400)
        .json({ error: "spreadsheetId is required" });
    }

    const data = await getPlacementOverview({ spreadsheetId });
    res.json(data);
  } catch (err) {
    console.error("Analytics error:", err);
    res.status(500).json({ error: err.message });
  }
});

export default router;
