import express from "express";
import { login, verifyEmail, register } from "./auth.controller.js";

const router = express.Router();

router.post("/login", login);
router.post("/verify-email", verifyEmail);
router.post("/register", register);

export default router;
