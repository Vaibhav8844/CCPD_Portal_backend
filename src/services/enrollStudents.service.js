import { parseStudentsExcel } from "../utils/parseStudentsExcel.js";
import {
  getOrCreateStudentWorkbook,
  ensureBranchSheet,
  appendStudents,
} from "../utils/studentWorkbook.js";

export async function enrollStudents({
  fileBuffer,
  year,
  branch,
  degreeType,
  program,
}) {
  const records = parseStudentsExcel(fileBuffer);

  const spreadsheetId =
    await getOrCreateStudentWorkbook(year, degreeType);

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
    workbook: `${degreeType}_Students_${year}`,
    sheet: branch,
  };
}
