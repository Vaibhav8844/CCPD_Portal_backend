import fs from "fs";
import path from "path";

const CONFIG_PATH = path.resolve(process.cwd(), "data", "config.json");

export function getAcademicYear() {
  try {
    const raw = fs.readFileSync(CONFIG_PATH, "utf8");
    const config = JSON.parse(raw);
    return config.academicYear || "2025-26";
  } catch (e) {
    return "2025-26";
  }
}

export function setAcademicYear(academicYear) {
  let config = {};
  try {
    config = JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
  } catch (e) {}
  config.academicYear = academicYear;
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
}
