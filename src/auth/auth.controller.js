import express from "express";
import { generateToken } from "./auth.service.js";

const router = express.Router();

// TEMP login (replace with DB later)
router.post("/login", (req, res) => {
  const { email, role } = req.body;

  const token = generateToken({ email, role });
  res.json({ token });
});

export default router;
