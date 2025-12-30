import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";
import { getSheet } from "../sheets/sheets.client.js";
import { buildUserSearchIndex, deduplicateResults } from "../utils/trieSearch.js";

const router = express.Router();

// Cache for Trie index to avoid rebuilding on every search
let userSearchCache = {
  trie: null,
  lastUpdated: null,
  ttl: 5 * 60 * 1000, // 5 minutes cache TTL
};

const ROLE_HIERARCHY = {
  SPOC: ["SPOC", "CALENDAR_TEAM","DATA_TEAM", "ADMIN"],
  CALENDAR_TEAM: ["CALENDAR_TEAM", "ADMIN"],
  DATA_TEAM: ["DATA_TEAM", "ADMIN"],
  ADMIN: ["ADMIN"],
};

/**
 * Search SPOCs by name (autocomplete) - OPTIMIZED with Trie data structure
 * Time Complexity: O(m) where m is the query length (vs O(n*m) with includes())
 * Space Complexity: O(n*k) where n is users and k is avg name/email length
 */
router.get(
  "/spocs",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN", "DATA_TEAM"),
  async (req, res) => {
    const query = (req.query.q || "").toLowerCase();

    // Return empty if query is too short
    if (!query || query.length < 1) {
      return res.json({ users: [] });
    }

    let rows;
    try {
      rows = await getSheet("Associates");
    } catch (err) {
      console.error("Failed to read Associates sheet:", err.message || err);
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

    // Build or retrieve cached Trie index
    const now = Date.now();
    const cacheExpired = !userSearchCache.lastUpdated || (now - userSearchCache.lastUpdated) > userSearchCache.ttl;

    if (!userSearchCache.trie || cacheExpired) {
      // Extract and filter SPOC users
      const spocUsers = rows
        .slice(1)
        .map((r) => {
          const out = header.map((_, i) => String(r[i] || "").trim());
          return out;
        })
        .filter((r) => {
          const roleVal = String(r[idx.role] || "").toUpperCase();
          return ROLE_HIERARCHY["SPOC"].includes(roleVal);
        })
        .map((r) => ({
          name: r[idx.name],
          email: r[idx.email],
          role: r[idx.role],
        }));

      // Build Trie index from SPOC users
      userSearchCache.trie = buildUserSearchIndex(spocUsers);
      userSearchCache.lastUpdated = now;
      console.log(`🔍 Trie index rebuilt with ${spocUsers.length} SPOC users`);
    }

    // Perform ultra-fast prefix search using Trie
    let results = userSearchCache.trie.searchPrefix(query);
    
    // Deduplicate results (same user can match on name/email)
    results = deduplicateResults(results);

    // Limit results to prevent overwhelming the frontend
    const MAX_RESULTS = 20;
    if (results.length > MAX_RESULTS) {
      results = results.slice(0, MAX_RESULTS);
    }

    if (results.length === 0) {
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

    res.json({ users: results });
  }
);
export default router;