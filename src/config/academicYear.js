import fs from "fs";
import path from "path";

const CONFIG_PATH = path.resolve(process.cwd(), "data", "config.json");

export function getAcademicYear() {
  // Priority 1: Use environment variable (for production)
  if (process.env.ACADEMIC_YEAR) {
    return process.env.ACADEMIC_YEAR;
  }
  
  // Priority 2: Use config file (for development)
  try {
    const raw = fs.readFileSync(CONFIG_PATH, "utf8");
    const config = JSON.parse(raw);
    return config.academicYear || "2025-26";
  } catch (e) {
    return "2025-26";
  }
}

export function setAcademicYear(academicYear) {
  // If using environment variable, don't write to file
  if (process.env.ACADEMIC_YEAR) {
    console.warn("[academicYear] Using ACADEMIC_YEAR from environment. File update skipped.");
    return;
  }
  
  // Update config file in development
  let config = {};
  try {
    config = JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
  } catch (e) {}
  config.academicYear = academicYear;
  
  // Ensure directory exists
  fs.mkdirSync(path.dirname(CONFIG_PATH), { recursive: true });
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
}
