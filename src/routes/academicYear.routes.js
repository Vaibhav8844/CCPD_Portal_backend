import express from "express";
import { getAcademicYear, setAcademicYear } from "../config/academicYear.js";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";

const router = express.Router();

// Get current academic year (accessible to all authenticated users)
router.get("/", authenticate, (req, res) => {
  res.json({ academicYear: getAcademicYear() });
});

// Set academic year (ADMIN only)
router.post("/", authenticate, roleGuard("ADMIN"), (req, res) => {
  const { academicYear } = req.body;
  if (!academicYear || typeof academicYear !== "string") {
    return res.status(400).json({ message: "academicYear is required" });
  }
  setAcademicYear(academicYear);
  res.json({ academicYear });
});

export default router;
