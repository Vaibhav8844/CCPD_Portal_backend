import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";
import { getSheet } from "../sheets/sheets.client.js";

const router = express.Router();

const ROLE_HIERARCHY = {
  SPOC: ["SPOC", "CALENDAR_TEAM","DATA_TEAM", "ADMIN"],
  CALENDAR_TEAM: ["CALENDAR_TEAM", "ADMIN"],
  DATA_TEAM: ["DATA_TEAM", "ADMIN"],
  ADMIN: ["ADMIN"],
};

/**
 * Search SPOCs by name (autocomplete)
 */
router.get(
  "/spocs",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN", "DATA_TEAM"),
  async (req, res) => {
    const query = (req.query.q || "").toLowerCase();

    let rows;
    try {
      rows = await getSheet("Associates");
    } catch (err) {
      console.error("Failed to read Associates sheet:", err.message || err);
      // don't throw — return empty list so frontend continues to work
      return res.json({ users: [], warning: "Failed to read Associates sheet. Check Google credentials and calendar initialization." });
    }

    if (!rows || rows.length < 1) {
      return res.json({ users: [], warning: "Associates sheet missing. Initialize calendar to create required sheets." });
    }

    const header = rows[0] || [];

    // normalize header names (trim, replace NBSP, lowercase)
    const normalize = (s) => String(s || "").replace(/\u00A0/g, " ").trim().toLowerCase();
    const headerNorm = header.map(normalize);

    // verify required headers exist (case/whitespace insensitive)
    const reqHeaders = ["name", "email", "role"];
    const missingHeaders = reqHeaders.filter((h) => headerNorm.indexOf(h) === -1);
    if (missingHeaders.length > 0) {
      return res.json({ users: [], warning: `Associates sheet missing headers: ${missingHeaders.join(", ")}. Expected: ${reqHeaders.join(", ")}` });
    }

    const idx = {
      name: headerNorm.indexOf("name"),
      email: headerNorm.indexOf("email"),
      role: headerNorm.indexOf("role"),
    };

    const users = rows
      .slice(1)
      .map((r) => {
        // ensure row has at least header length and trim cells
        const out = header.map((_, i) => String(r[i] || "").trim());
        return out;
      })
      .filter((r) => {
        const roleVal = String(r[idx.role] || "").toUpperCase();
        return ROLE_HIERARCHY["SPOC"].includes(roleVal);
      })
      .filter((r) => {
        const name = String(r[idx.name] || "").toLowerCase();
        const email = String(r[idx.email] || "").toLowerCase();
        return name.includes(query) || email.includes(query);
      })
      .map((r) => ({
        name: r[idx.name],
        email: r[idx.email],
        role: r[idx.role],
      }));

    if (!users || users.length === 0) {
      const sample = rows.slice(0, 6).map((r) =>
        (r || []).map((c) => (typeof c === "string" ? c.trim() : c))
      );
      return res.json({
        users: [],
        warning: "No matching SPOCs found for query",
        diagnostics: {
          rowsLength: rows.length,
          header,
          headerNorm: header.map((h) => String(h || "").replace(/\u00A0/g, " ").trim().toLowerCase()),
          sampleRows: sample,
        },
      });
    }

    res.json({ users });
  }
);
export default router;