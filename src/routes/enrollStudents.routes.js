import express from "express";
import multer from "multer";
import { enrollStudents } from "../services/enrollStudents.service.js";
import { authenticate } from "../auth/auth.middleware.js";
import roleGuard from "../middleware/roleGuard.js";

const router = express.Router();
const upload = multer();

router.post(
  "/students",
  authenticate,
  roleGuard("DATA_TEAM", "ADMIN"),
  upload.single("file"),
  async (req, res) => {
    try {
      const { year, branch, degreeType, program } = req.body;

      if (!req.file) {
        return res.status(400).json({ error: "Excel file required" });
      }

      const result = await enrollStudents({
        fileBuffer: req.file.buffer,
        year,
        branch,
        degreeType,
        program,
      });

      res.json({
        success: true,
        ...result,
      });
    } catch (err) {
      console.error("Enroll students error:", err);
      res.status(500).json({ error: err.message });
    }
  }
);

export default router;
