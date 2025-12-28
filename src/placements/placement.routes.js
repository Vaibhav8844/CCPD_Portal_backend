import express from "express";
import { getCalendarData } from "./placement.controller.js";
import { authenticate } from "../auth/auth.middleware.js";
import { enrollBatch } from "./placement.controller.js";
import roleGuard from "../middleware/roleGuard.js";

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

router.get(
  "/calendar",
  authenticate,
  roleGuard("CALENDAR_TEAM", "ADMIN"),
  getCalendarData
);

router.post(
  "/enroll-batch",
  authenticate,
  roleGuard("DATA_TEAM", "ADMIN"),
  enrollBatch
);

export default router;
