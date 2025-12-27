import jwt from "jsonwebtoken";
import bcrypt from "bcryptjs";
import { users } from "./users.local.js";

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
