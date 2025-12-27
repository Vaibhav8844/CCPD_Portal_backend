import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";
import { getSheet } from "../sheets/sheets.client.js";

const router = express.Router();

const ROLE_HIERARCHY = {
  SPOC: ["SPOC", "CALENDAR_TEAM", "ADMIN"],
  CALENDAR_TEAM: ["CALENDAR_TEAM", "ADMIN"],
  ADMIN: ["ADMIN"],
};

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

    const header = rows[0];

    const idx = {
      name: header.indexOf("Name"),
      email: header.indexOf("Email"),
      role: header.indexOf("Role"),
    };

    const users = rows
      .slice(1)
      .filter(r =>
        ROLE_HIERARCHY["SPOC"].includes(
          r[idx.role]?.trim()
        )
      )
      .filter(r =>
        r[idx.name]?.toLowerCase().includes(query) ||
        r[idx.email]?.toLowerCase().includes(query)
      )
      .map(r => ({
        name: r[idx.name],
        email: r[idx.email],
        role: r[idx.role],
      }));

    res.json({ users });
  }
);
export default router;