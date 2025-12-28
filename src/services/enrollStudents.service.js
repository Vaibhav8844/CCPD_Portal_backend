import { parseStudentsExcel } from "../utils/parseStudentsExcel.js";
import { getAcademicYear } from "../config/academicYear.js";
import { populatePlacementStudentsSheet } from "../analytics/populatePlacementSheets.js";

export async function enrollStudents({
  fileBuffer,
  branch,
  degreeType,
  program,
}) {
  const records = parseStudentsExcel(fileBuffer);
  const academicYear = getAcademicYear();

  // Populate placement workbook Students_<branch> sheet directly
  const result = await populatePlacementStudentsSheet({
    academicYear,
    degreeType,
    branch,
    students: records,
  });

  return {
    inserted: result.added,
    duplicates: result.skipped,
    workbook: `Placement_Data_${academicYear}_${degreeType}`,
    sheet: `Students_${branch}`,
  };
}
