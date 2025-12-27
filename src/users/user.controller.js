import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import { roleGuard } from "../middleware/roleGuard.js";
import { getSheet } from "../sheets/sheets.client.js";

const router = express.Router();

/**
 * Search SPOCs by name (autocomplete)
 */
router.get(
  "/spocs",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN"),
  async (req, res) => {
    const query = (req.query.q || "").toLowerCase();

    const rows = await getSheet("Associates");
    if (rows.length < 2) return res.json({ users: [] });

    const users = rows
      .slice(1)
      .filter(r =>
        r[2] === "SPOC" &&
        r[0]?.toLowerCase().includes(query)
      )
      .map(r => ({
        name: r[0],
        email: r[1],
      }));

    res.json({ users });
  }
);

export default router;
