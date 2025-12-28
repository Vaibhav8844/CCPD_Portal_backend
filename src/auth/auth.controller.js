import jwt from "jsonwebtoken";
import bcrypt from "bcryptjs";
import { users, saveUsers, getRegisteredUsers } from "./users.local.js";
import { getSheet } from "../sheets/sheets.client.js";
import { idxOf } from "../utils/sheetUtils.js";

const JWT_SECRET = process.env.JWT_SECRET || "super_secret_key";

export async function login(req, res) {
  const { username, password, rememberMe } = req.body;

  const user = users.find((u) => u.username === username);
  if (!user) {
    return res.status(401).json({ message: "Invalid credentials" });
  }

  const valid = await bcrypt.compare(password, user.passwordHash);
  if (!valid) {
    return res.status(401).json({ message: "Invalid credentials" });
  }

  const token = jwt.sign(
    {
      id: user.id,
      role: user.role,
      username: user.username,
    },
    JWT_SECRET,
    {
      expiresIn: rememberMe ? "7d" : "2h",
    }
  );

  res.json({
    token,
    role: user.role,
    name: user.name,
  });
}

export async function verifyEmail(req, res) {
  const { email } = req.body;

  if (!email) {
    return res.status(400).json({ message: "Email is required" });
  }

  try {
    const rows = await getSheet("Associates");
    const header = rows[0];
    const emailIdx = idxOf(header, "Email");
    const nameIdx = idxOf(header, "Name");
    const roleIdx = idxOf(header, "Role");

    const associate = rows.slice(1).find(
      (r) => r[emailIdx]?.toLowerCase() === email.toLowerCase()
    );

    if (!associate) {
      return res.status(404).json({ message: "Email not Registered.Please contact Admin" });
    }

    // Check if user already registered
    const existingUser = users.find((u) => u.username === email);
    if (existingUser) {
      return res.status(409).json({ message: "User already registered. Please login." });
    }

    res.json({
      email,
      name: associate[nameIdx] || "",
      role: associate[roleIdx] || "SPOC",
    });
  } catch (err) {
    console.error("Verify email error:", err);
    res.status(500).json({ message: "Failed to verify email" });
  }
}

export async function register(req, res) {
  const { email, username, password, name, role } = req.body;

  if (!email || !username || !password) {
    return res.status(400).json({ message: "Email, username, and password are required" });
  }

  try {
    // Verify email exists in Associates
    const rows = await getSheet("Associates");
    const header = rows[0];
    const emailIdx = idxOf(header, "Email");

    const associate = rows.slice(1).find(
      (r) => r[emailIdx]?.toLowerCase() === email.toLowerCase()
    );

    if (!associate) {
      return res.status(403).json({ message: "Email not authorized" });
    }

    // Check if username already exists
    const existingUser = users.find((u) => u.username === username);
    if (existingUser) {
      return res.status(409).json({ message: "Username already exists" });
    }

    // Create new user
    const passwordHash = await bcrypt.hash(password, 10);
    const newUser = {
      id: String(users.length + 1),
      name: name || email,
      username,
      passwordHash,
      role: role || "SPOC",
    };

    users.push(newUser);

    // Save registered users to file
    const registeredUsers = getRegisteredUsers();
    saveUsers(registeredUsers);

    const token = jwt.sign(
      {
        id: newUser.id,
        role: newUser.role,
        username: newUser.username,
      },
      JWT_SECRET,
      { expiresIn: "2h" }
    );

    res.json({
      token,
      role: newUser.role,
      name: newUser.name,
      message: "Registration successful",
    });
  } catch (err) {
    console.error("Register error:", err);
    res.status(500).json({ message: "Registration failed" });
  }
}
