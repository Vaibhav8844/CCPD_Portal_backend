import { parseStudentsExcel } from "../utils/parseStudentsExcel.js";
import {
  getOrCreateStudentWorkbook,
  ensureBranchSheet,
  appendStudents,
} from "../utils/studentWorkbook.js";
import { getAcademicYear } from "../config/academicYear.js";

export async function enrollStudents({
  fileBuffer,
  branch,
  degreeType,
  program,
}) {
  const records = parseStudentsExcel(fileBuffer);

  // Always use admin-set academic year for file naming
  const academicYear = getAcademicYear();

  const spreadsheetId = await getOrCreateStudentWorkbook(academicYear, degreeType);
  await ensureBranchSheet(spreadsheetId, branch);

  const rows = records.map((r) => [
    r["Roll No."] || "",
    r["Student Name"] || "",
    r["Gender"] || "",
    r["Phone No."] || "",
    r["Institute Email"] || "",
    r["Personal Email"] || "",
    branch,
    degreeType,
    program,
    r["CGPA"] || "",
    r["Session"] || "",
    r["Semester/Quarter"] || "",
  ]);

  await appendStudents(spreadsheetId, branch, rows);

  return {
    inserted: rows.length,
    workbook: `${degreeType}_Students_${academicYear}`,
    sheet: branch,
  };
}
