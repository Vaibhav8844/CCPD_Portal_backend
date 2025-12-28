import fs from "fs";
import path from "path";

const TOKEN_PATH = process.env.GOOGLE_TOKEN_PATH;

export function saveTokens(tokens) {
  fs.mkdirSync(path.dirname(TOKEN_PATH), { recursive: true });
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
}

export function loadTokens() {
  if (!fs.existsSync(TOKEN_PATH)) return null;
  return JSON.parse(fs.readFileSync(TOKEN_PATH));
}
