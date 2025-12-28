import express from "express";
import { oauth2Client } from "../auth/googleAuth.js";
import { saveTokens } from "../utils/tokenStore.js";

const router = express.Router();

// STEP 1: Login
router.get("/google", (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: [
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/spreadsheets",
    ],
  });

  res.redirect(url);
});

// STEP 2: Callback
router.get("/google/callback", async (req, res) => {
  try {
    const { code } = req.query;

    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    saveTokens(tokens);

    res.send("✅ Google OAuth successful. You can close this tab.");
  } catch (err) {
    console.error(err);
    res.status(500).send("OAuth failed");
  }
});

export default router;
