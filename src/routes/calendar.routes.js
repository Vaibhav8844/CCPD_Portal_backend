import express from "express";
import fs from "fs/promises";
import path from "path";
import { getSheets, getDrive } from "../sheets/sheets.dynamic.js";
import { ensureHeaders } from "../utils/sheetBootstrap.js";
import { DRIVE_REQUEST_HEADERS, COMPANY_DRIVES_HEADERS, SHEET_TEMPLATES } from "../constants/sheets.js";
import { appendRow, getSpreadsheetId, deleteRowByEmail } from "../sheets/sheets.client.js";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";

const router = express.Router();
const stateFile = path.resolve(process.cwd(), "data", "calendar_state.json");

async function readState() {
  try {
    const raw = await fs.readFile(stateFile, "utf8");
    return JSON.parse(raw);
  } catch (err) {
    return { initialized: false };
  }
}

async function writeState(state) {
  try {
    await fs.mkdir(path.dirname(stateFile), { recursive: true });
    await fs.writeFile(stateFile, JSON.stringify(state, null, 2), "utf8");
  } catch (err) {
    throw err;
  }
}

router.get("/status", async (req, res) => {
  try {
    const state = await readState();

    // If a workbookId exists, verify it still exists in Drive.
    if (state.workbookId) {
      try {
        const drive = await getDrive();
        await drive.files.get({ fileId: state.workbookId, fields: "id,name" });
        // file exists — return current state
        return res.json({ initialized: !!state.initialized, ...state });
      } catch (err) {
        // workbook not found on Drive (deleted) — clear persisted state
        try {
          const cleared = { initialized: false };
          await writeState(cleared);
          return res.json({ initialized: false });
        } catch (writeErr) {
          console.error("Failed to clear calendar state after missing workbook:", writeErr);
          return res.status(500).json({ error: "Calendar state inconsistent" });
        }
      }
    }

    res.json({ initialized: !!state.initialized, ...state });
  } catch (err) {
    console.error("Failed to read calendar state:", err);
    res.status(500).json({ error: "Failed to read state" });
  }
});
router.post(
  "/initialize",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN"),
  async (req, res) => {
    try {
      const state = await readState();
      if (state.initialized) {
        return res.json({ success: true, message: "Already initialized", state });
      }

      const { year } = req.body;

      // compute default academic year if not provided
      function computeDefaultAcademicYear() {
        const now = new Date();
        const yr = now.getFullYear();
        const month = now.getMonth() + 1; // 1-12
        const start = month >= 6 ? yr : yr - 1;
        const end = start + 1;
        return `${start}-${String(end).slice(-2)}`;
      }

      const academicYear = year || computeDefaultAcademicYear();
      const workbookName = `CCPD Calendar ${academicYear}`;

      const folderId = process.env.PLACEMENT_FOLDER_ID;
      if (!folderId) throw new Error("PLACEMENT_FOLDER_ID not set");

      const drive = await getDrive();
      const sheets = await getSheets();

      // check if workbook already exists
      const listRes = await drive.files.list({
        q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet' and '${folderId}' in parents`,
        fields: "files(id, name)",
      });

      let spreadsheetId;
      if (listRes.data.files.length > 0) {
        spreadsheetId = listRes.data.files[0].id;
      } else {
        const create = await drive.files.create({
          requestBody: {
            name: workbookName,
            mimeType: "application/vnd.google-apps.spreadsheet",
            parents: [folderId],
          },
        });
        spreadsheetId = create.data.id;
      }

      // Ensure the calendar workbook is accessible to the local service account
      // (if present) so server-side operations using service-account auth can succeed.
      try {
        const saPath = path.resolve(process.cwd(), "service-account.json");
        if (fs.existsSync(saPath)) {
          const raw = await fs.readFile(saPath, "utf8");
          const creds = JSON.parse(raw);
          const saEmail = creds.client_email;
          if (saEmail) {
            try {
              // create permission if it does not already exist
              await drive.permissions.create({
                fileId: spreadsheetId,
                requestBody: { role: "writer", type: "user", emailAddress: saEmail },
                sendNotificationEmail: false,
              });
              console.log("Granted service account access to calendar workbook", saEmail);
            } catch (permErr) {
              // if permission already exists or API rejects, just warn
              console.warn("Failed to grant service account permission:", permErr.message || permErr);
            }
          }
        }
      } catch (err) {
        console.warn("Error while attempting to share workbook with service account:", err.message || err);
      }

      // Ensure required sheets + headers exist in the workbook
      try {
        // core calendar sheet
        await ensureHeaders("Calendar", ["Date", "Company", "Drive Title", "Status"], spreadsheetId);

        // drive requests + company drives
        await ensureHeaders("Drive_Requests", DRIVE_REQUEST_HEADERS, spreadsheetId);
        await ensureHeaders("Company_Drives", COMPANY_DRIVES_HEADERS, spreadsheetId);

        // company <-> spoc mapping
        await ensureHeaders("Company_SPOC_Map", ["Company", "SPOC"], spreadsheetId);

        // associates (SPOC user list) used by /users/spocs
        await ensureHeaders("Associates", ["Name", "Email", "Role"], spreadsheetId);

        // student & placement data sheets
        await ensureHeaders(
          "Students_Data",
          ["Roll Number", "Student Name", "Branch", "CGPA"],
          spreadsheetId
        );

        await ensureHeaders(
          "Placement_Data",
          [
            "Request ID",
            "Company",
            "Drive Type",
            "Eligible Pool",
            "Internship Stipend",
            "FTE CTC",
            "FTE Base",
            "Roll Number",
            "Student Name",
            "Branch",
            "CGPA",
            "Result",
            "Published At",
          ],
          spreadsheetId
        );

        await ensureHeaders("Placement_Results", ["Company", "Roll"], spreadsheetId);

        // If there is an existing default spreadsheet with data (GOOGLE_SHEET_ID),
        // copy over common sheets like Associates so new calendar has existing SPOC entries.
        const defaultId = process.env.GOOGLE_SHEET_ID;
        if (defaultId && defaultId !== spreadsheetId) {
          try {
            const sheetsApi = await getSheets();
            const copySheets = ["Associates", "Company_SPOC_Map", "Company_Drives", "Drive_Requests"];
            for (const name of copySheets) {
              try {
                const r = await sheetsApi.spreadsheets.values.get({ spreadsheetId: defaultId, range: `${name}!A1:Z1000` });
                const vals = r.data.values;
                if (vals && vals.length > 0) {
                  await sheetsApi.spreadsheets.values.update({
                    spreadsheetId,
                    range: `${name}!A1`,
                    valueInputOption: "RAW",
                    requestBody: { values: vals },
                  });
                }
              } catch (inner) {
                // ignore missing sheets in default spreadsheet
              }
            }
          } catch (err) {
            console.warn("Failed to copy data from default spreadsheet:", err.message || err);
          }
        }
      } catch (err) {
        console.warn("Failed to ensure required calendar sheets:", err.message || err);
      }

      const newState = {
        initialized: true,
        academicYear,
        workbookName,
        workbookId: spreadsheetId,
        initializedAt: new Date().toISOString(),
      };

      await writeState(newState);

      res.json({ success: true, message: "Calendar initialized", state: newState });
    } catch (err) {
      console.error("Failed to initialize calendar:", err);
      res.status(500).json({ error: err.message || "Initialization failed" });
    }
  }
);

/**
 * Delete an associate row by email
 */
router.post(
  "/associates/delete",
  authenticate,
  roleGuard("ADMIN"),
  async (req, res) => {
    try {
      const { email } = req.body;
      if (!email) return res.status(400).json({ message: "email is required" });

      const calendarId = await getSpreadsheetId();
      if (!calendarId) return res.status(400).json({ message: "Calendar workbook not initialized" });

      const result = await deleteRowByEmail("Associates", email);
      if (result.deleted) return res.json({ success: true });
      return res.status(404).json({ success: false, message: result.reason || "Not found" });
    } catch (err) {
      console.error("Failed to delete associate:", err);
      res.status(500).json({ error: err.message || "Delete failed" });
    }
  }
);

/**
 * Append a single associate row into the calendar workbook
 */
router.post(
  "/associates/append",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN"),
  async (req, res) => {
    try {
      const { name, email, role } = req.body;
      if (!name || !email || !role) {
        return res.status(400).json({ message: "name, email and role are required" });
      }

      await appendRow("Associates", [name.trim(), email.trim(), role.trim()]);
      res.json({ success: true });
    } catch (err) {
      console.error("Failed to append associate:", err);
      res.status(500).json({ error: err.message || "Failed to append" });
    }
  }
);

/**
 * Import Associates (and optionally other sheets) from the default spreadsheet into calendar workbook
 */
router.post(
  "/associates/import-from-default",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN"),
  async (req, res) => {
    try {
      const defaultId = process.env.GOOGLE_SHEET_ID;
      if (!defaultId) return res.status(400).json({ message: "No default spreadsheet configured (GOOGLE_SHEET_ID)" });

      const calendarId = await getSpreadsheetId();
      if (!calendarId) return res.status(400).json({ message: "Calendar workbook not initialized" });

      const sheetsApi = await getSheets();
      const copySheets = req.body.sheets || ["Associates"];
      const results = [];

      for (const name of copySheets) {
        try {
          const r = await sheetsApi.spreadsheets.values.get({ spreadsheetId: defaultId, range: `${name}!A1:Z1000` });
          const vals = r.data.values;
          if (vals && vals.length > 0) {
            await sheetsApi.spreadsheets.values.update({
              spreadsheetId: calendarId,
              range: `${name}!A1`,
              valueInputOption: "RAW",
              requestBody: { values: vals },
            });
            results.push({ sheet: name, copied: vals.length });
          } else {
            results.push({ sheet: name, copied: 0 });
          }
        } catch (inner) {
          results.push({ sheet: name, error: inner.message || String(inner) });
        }
      }

      res.json({ success: true, results });
    } catch (err) {
      console.error("Failed to import associates from default:", err);
      res.status(500).json({ error: err.message || "Import failed" });
    }
  }
);

export default router;

