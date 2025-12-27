import express from "express";
import { authenticate } from "../auth/auth.middleware.js";
import { roleGuard } from "../middleware/roleGuard.js";
import { appendRow } from "../sheets/sheets.client.js";

const router = express.Router();

router.post(
  "/submit",
  authenticate,
  roleGuard("SPOC"),
  async (req, res) => {
    const { company, roll } = req.body;
    await appendRow("Placement_Results", [company, roll]);
    res.json({ success: true });
  }
);

export default router;
