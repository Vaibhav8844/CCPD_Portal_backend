import { getSheets, getDrive } from "../sheets/sheets.dynamic.js";
import { getAcademicYear } from "../config/academicYear.js";
import { getOrCreateStudentWorkbook } from "../utils/studentWorkbook.js";

/**
 * Helper: Get student data from workbook
 * STUDENT_HEADERS: Roll No.(0), Student Name(1), Gender(2), Phone No.(3), Institute Email(4), Personal Email(5), Department(6), Degree Type(7), Program(8), CGPA(9), Session(10), Semester/Quarter(11)
 */
async function getStudentData(degreeType) {
  const academicYear = getAcademicYear();
  const sheets = await getSheets();
  
  try {
    const workbookId = await getOrCreateStudentWorkbook(academicYear, degreeType);
    
    const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
    const branches = meta.data.sheets
      .map(s => s.properties.title)
      .filter(name => !name.startsWith("Sheet"));
    
    const students = [];
    
    for (const branch of branches) {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: workbookId,
          range: `${branch}!A2:Z`,
        });
        
        const rows = result.data.values || [];
        for (const row of rows) {
          if (row[0]) { // Has roll number
            students.push({
              rollNo: row[0],
              name: row[1],
              branch,
              degreeType,
              gender: row[2],
              phone: row[3],
              email: row[4] || row[5],
              cgpa: parseFloat(row[9]) || 0,
            });
          }
        }
      } catch (err) {
        console.error(`Error reading branch ${branch}:`, err.message);
      }
    }
    
    return students;
  } catch (err) {
    console.error(`[getStudentData] Error for ${degreeType}:`, err.message);
    return [];
  }
}

/**
 * Helper: Get placement data from Offers sheets
 * Offers sheet columns: Roll No(0), Company(1), Offer Type(2), CTC (LPA)(3), Offer Status(4)
 */
async function getPlacementData(degreeType) {
  const academicYear = getAcademicYear();
  const drive = await getDrive();
  const sheets = await getSheets();
  
  const workbookName = `Placement_Data_${academicYear}_${degreeType}`;
  
  const res = await drive.files.list({
    q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
    fields: "files(id, name)",
  });
  
  if (res.data.files.length === 0) {
    console.log(`[getPlacementData] No workbook found: ${workbookName}`);
    return [];
  }
  
  const workbookId = res.data.files[0].id;
  
  try {
    // Get all Offers_* sheets
    const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
    const offersSheets = meta.data.sheets
      .map(s => s.properties.title)
      .filter(name => name.startsWith("Offers_"));
    
    const allOffers = [];
    
    for (const sheetName of offersSheets) {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: workbookId,
          range: `${sheetName}!A2:Z`,
        });
        
        const rows = result.data.values || [];
        const branch = sheetName.replace("Offers_", "");
        
        // Columns: Roll No(0), Company(1), Offer Type(2), CTC (LPA)(3), Offer Status(4)
        for (const row of rows) {
          if (row[0] && row[4] === "Active") { // Only active offers
            allOffers.push({
              rollNo: row[0],
              company: row[1] || "",
              offerType: row[2] || "",
              ctc: parseFloat(row[3]) || 0,
              branch,
            });
          }
        }
      } catch (err) {
        console.error(`Error reading ${sheetName}:`, err.message);
      }
    }
    
    return allOffers.filter(p => p.rollNo && p.company);
  } catch (err) {
    console.error("Error reading placement data:", err);
    return [];
  }
}

/**
 * Get overall placement statistics from Overall sheet
 */
export async function getOverallStats() {
  try {
    const academicYear = getAcademicYear();
    const drive = await getDrive();
    const sheets = await getSheets();
    
    const workbookName = `Placement_Data_${academicYear}_UG`;
    
    const res = await drive.files.list({
      q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: "files(id, name)",
    });
    
    if (res.data.files.length === 0) {
      console.log(`[getOverallStats] No workbook found: ${workbookName}`);
      return { academicYear, error: "No data available" };
    }
    
    const workbookId = res.data.files[0].id;
    
    // Read Overall sheet (rows 2-5 are CS, EC, EE, ME)
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: workbookId,
      range: "Overall!A2:AA5", // 27 columns: A-AA
    });
    
    const rows = result.data.values || [];
    
    // Aggregate stats across all branches
    let totalStudents = 0;
    let totalEligible = 0;
    let totalPlaced = 0;
    let ctcValues = [];
    let onlyInternship = 0;
    let onlyFTE = 0;
    let bothOffers = 0;
    
    for (const row of rows) {
      if (!row[0]) continue; // Skip empty rows
      
      totalStudents += parseInt(row[1]) || 0;
      totalEligible += parseInt(row[4]) || 0;
      totalPlaced += parseInt(row[7]) || 0;
      
      // Collect CTC values (columns Q-T: Highest, Average, Lowest, Median)
      const highest = parseFloat(row[16]) || 0;
      const lowest = parseFloat(row[18]) || 0;
      if (highest > 0 && lowest > 0) {
        ctcValues.push(highest, lowest);
      }
      
      onlyInternship += parseInt(row[20]) || 0;
      onlyFTE += parseInt(row[21]) || 0;
      bothOffers += parseInt(row[22]) || 0;
    }
    
    // Calculate overall CTC stats
    const highestCTC = ctcValues.length > 0 ? Math.max(...ctcValues).toFixed(2) : 0;
    const lowestCTC = ctcValues.length > 0 ? Math.min(...ctcValues).toFixed(2) : 0;
    const averageCTC = ctcValues.length > 0 
      ? (ctcValues.reduce((sum, c) => sum + c, 0) / ctcValues.length).toFixed(2)
      : 0;
    
    const stats = {
      academicYear,
      totalStudents,
      eligiblePool: totalEligible,
      totalPlaced,
      placementPercentage: totalEligible > 0 
        ? ((totalPlaced / totalEligible) * 100).toFixed(2)
        : 0,
      averageCTC,
      highestCTC,
      lowestCTC,
      medianCTC: ctcValues.length > 0 ? calculateMedian(ctcValues) : 0,
      onlyInternship,
      onlyFTE,
      bothOffers,
      offersReceived: totalPlaced, // Each placed student has at least 1 offer
    };

    return stats;
  } catch (err) {
    console.error("[stats] getOverallStats error:", err);
    throw err;
  }
}

function calculateMedian(values) {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? ((sorted[mid - 1] + sorted[mid]) / 2).toFixed(2)
    : sorted[mid].toFixed(2);
}

/**
 * Get branch-wise statistics from Placement_Stats sheets
 */
export async function getBranchwiseStats(degreeType = "UG") {
  try {
    const academicYear = getAcademicYear();
    const drive = await getDrive();
    const sheets = await getSheets();
    
    const workbookName = `Placement_Data_${academicYear}_${degreeType}`;
    
    const res = await drive.files.list({
      q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: "files(id, name)",
    });
    
    if (res.data.files.length === 0) {
      console.log(`[getBranchwiseStats] No workbook found: ${workbookName}`);
      return [];
    }
    
    const workbookId = res.data.files[0].id;
    
    // Get all Placement_Stats sheets
    const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
    const statsSheets = meta.data.sheets
      .map(s => s.properties.title)
      .filter(name => name.startsWith("Placement_Stats_"));
    
    if (statsSheets.length === 0) {
      console.log(`[getBranchwiseStats] No Placement_Stats sheets found in ${workbookName}`);
      return [];
    }
    
    const branchStats = [];
    
    for (const sheetName of statsSheets) {
      try {
        // Read the Placement_Stats sheet (key-value pairs in columns A and B, rows 2-27)
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: workbookId,
          range: `${sheetName}!A2:B27`,
        });
        
        const rows = result.data.values || [];
        const statsMap = {};
        
        // Parse key-value pairs
        for (const row of rows) {
          if (row[0] && row[1] !== undefined) {
            statsMap[row[0]] = row[1];
          }
        }
        
        console.log(`[DEBUG] ${sheetName} raw data:`, {
          medianCTC: statsMap["Median CTC (LPA)"],
          avgCTC: statsMap["Average CTC (LPA)"],
          totalStudents: statsMap["Total Students"],
          maleStudents: statsMap["Total M Students"],
          placed: statsMap["Total Placed"]
        });
        
        const branch = sheetName.replace("Placement_Stats_", "");
        
        // Helper function to parse numeric values from sheets
        const parseNumeric = (value) => {
          if (value === undefined || value === null || value === "") return 0;
          const parsed = parseFloat(String(value).replace(/,/g, ''));
          return isNaN(parsed) ? 0 : parsed;
        };
        
        const parseInteger = (value) => {
          if (value === undefined || value === null || value === "") return 0;
          const parsed = Number(String(value).replace(/,/g, ''));
          return isNaN(parsed) ? 0 : Math.floor(parsed);
        };
        
        branchStats.push({
          branch,
          degreeType,
          totalStudents: parseInteger(statsMap["Total Students"]),
          maleStudents: parseInteger(statsMap["Total M Students"]),
          femaleStudents: parseInteger(statsMap["Total F Students"]),
          eligibleStudents: parseInteger(statsMap["Total Eligible"]),
          maleEligible: parseInteger(statsMap["Total M Eligible"]),
          femaleEligible: parseInteger(statsMap["Total F Eligible"]),
          placed: parseInteger(statsMap["Total Placed"]),
          malePlaced: parseInteger(statsMap["No of M Placed"]),
          femalePlaced: parseInteger(statsMap["No of F Placed"]),
          placementPercentageOfTotal: parseNumeric(statsMap["% Students Placed (of Total)"]),
          placementPercentage: parseNumeric(statsMap["% Students Placed (of Eligible)"]),
          malePlacedPercentageOfTotal: parseNumeric(statsMap["% M Placed (of Total)"]),
          malePlacedPercentageOfEligible: parseNumeric(statsMap["% M Placed (of Eligible)"]),
          femalePlacedPercentageOfTotal: parseNumeric(statsMap["% F Placed (of Total)"]),
          femalePlacedPercentageOfEligible: parseNumeric(statsMap["% F Placed (of Eligible)"]),
          highestCTC: parseNumeric(statsMap["Highest CTC (LPA)"]),
          averageCTC: parseNumeric(statsMap["Average CTC (LPA)"]),
          lowestCTC: parseNumeric(statsMap["Lowest CTC (LPA)"]),
          medianCTC: parseNumeric(statsMap["Median CTC (LPA)"]),
          onlyInternship: parseInteger(statsMap["Only Internship Offers"]),
          onlyFTE: parseInteger(statsMap["Only FTE Offers"]),
          bothOffers: parseInteger(statsMap["Both Offers"]),
          unplacedCGPA8Plus: parseInteger(statsMap["Unplaced (CGPA >= 8)"]),
          unplacedCGPA7_5Plus: parseInteger(statsMap["Unplaced (CGPA >= 7.5)"]),
          unplacedCGPA7Plus: parseInteger(statsMap["Unplaced (CGPA >= 7)"]),
          unplacedCGPA6_5Plus: parseInteger(statsMap["Unplaced (CGPA >= 6.5)"]),
        });
      } catch (err) {
        console.error(`[getBranchwiseStats] Error reading ${sheetName}:`, err);
      }
    }

    return branchStats.sort((a, b) => a.branch.localeCompare(b.branch));
  } catch (err) {
    console.error("[stats] getBranchwiseStats error:", err);
    throw err;
  }
}

/**
 * Fallback: Calculate branch-wise stats from raw data
 */
async function calculateBranchwiseStats(degreeType) {
  const [students, placements] = await Promise.all([
    getStudentData(degreeType),
    getPlacementData(degreeType),
  ]);
  
  // Group students by branch
  const branchMap = new Map();
  
  for (const student of students) {
    if (!branchMap.has(student.branch)) {
      branchMap.set(student.branch, {
        branch: student.branch,
        degreeType,
        totalStudents: 0,
        eligibleStudents: 0,
        placed: 0,
        offers: [],
      });
    }
    
    const branchData = branchMap.get(student.branch);
    branchData.totalStudents++;
    if (student.cgpa >= 6.5) {
      branchData.eligibleStudents++;
    }
  }
  
  // Add placement data
  for (const placement of placements) {
    if (branchMap.has(placement.branch)) {
      branchMap.get(placement.branch).offers.push(placement);
    }
  }
  
  // Calculate stats for each branch
  const branchStats = [];
  for (const [branch, data] of branchMap.entries()) {
    const placedRollNos = new Set(data.offers.map(o => o.rollNo));
    const placed = placedRollNos.size;
    const ctcValues = data.offers.map(o => o.ctc).filter(c => c > 0);
    
    branchStats.push({
      branch,
      degreeType,
      totalStudents: data.totalStudents,
      eligibleStudents: data.eligibleStudents,
      placed,
      placementPercentage: data.eligibleStudents > 0
        ? ((placed / data.eligibleStudents) * 100).toFixed(2)
        : 0,
      averageCTC: ctcValues.length > 0
        ? (ctcValues.reduce((sum, c) => sum + c, 0) / ctcValues.length).toFixed(2)
        : 0,
      highestCTC: ctcValues.length > 0 ? Math.max(...ctcValues).toFixed(2) : 0,
      offersReceived: data.offers.length,
    });
  }

  return branchStats.sort((a, b) => a.branch.localeCompare(b.branch));
}

/**
 * Get CTC distribution from CTC_Distribution sheets
 */
export async function getCTCDistribution() {
  try {
    const academicYear = getAcademicYear();
    const drive = await getDrive();
    const sheets = await getSheets();
    
    const workbookName = `Placement_Data_${academicYear}_UG`;
    
    const res = await drive.files.list({
      q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: "files(id, name)",
    });
    
    if (res.data.files.length === 0) {
      console.log(`[getCTCDistribution] No workbook found: ${workbookName}`);
      return [];
    }
    
    const workbookId = res.data.files[0].id;
    
    // Get all CTC_Distribution sheets
    const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
    const ctcSheets = meta.data.sheets
      .map(s => s.properties.title)
      .filter(name => name.startsWith("CTC_Distribution_"));
    
    // Aggregate data from all branches
    const rangeMap = new Map();
    
    for (const sheetName of ctcSheets) {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: workbookId,
          range: `${sheetName}!A2:B7`, // 6 ranges: 0-5, 5-10, 10-15, 15-20, 20-30, 30+
        });
        
        const rows = result.data.values || [];
        
        for (const row of rows) {
          if (!row[0]) continue;
          
          const range = row[0];
          const count = parseInt(row[1]) || 0;
          
          if (rangeMap.has(range)) {
            rangeMap.set(range, rangeMap.get(range) + count);
          } else {
            rangeMap.set(range, count);
          }
        }
      } catch (err) {
        console.error(`[getCTCDistribution] Error reading ${sheetName}:`, err);
      }
    }
    
    // Convert to array format
    const ranges = [
      { range: "0-5 LPA", count: rangeMap.get("0-5 LPA") || 0 },
      { range: "5-10 LPA", count: rangeMap.get("5-10 LPA") || 0 },
      { range: "10-15 LPA", count: rangeMap.get("10-15 LPA") || 0 },
      { range: "15-20 LPA", count: rangeMap.get("15-20 LPA") || 0 },
      { range: "20-30 LPA", count: rangeMap.get("20-30 LPA") || 0 },
      { range: "30+ LPA", count: rangeMap.get("30+ LPA") || 0 },
    ];

    return ranges;
  } catch (err) {
    console.error("[stats] getCTCDistribution error:", err);
    throw err;
  }
}

/**
 * Get gender/degree split from Overall sheet
 */
export async function getDemographicSplit() {
  try {
    const academicYear = getAcademicYear();
    const drive = await getDrive();
    const sheets = await getSheets();
    
    const workbookName = `Placement_Data_${academicYear}_UG`;
    
    const res = await drive.files.list({
      q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: "files(id, name)",
    });
    
    if (res.data.files.length === 0) {
      console.log(`[getDemographicSplit] No workbook found: ${workbookName}`);
      return { gender: {}, degree: {}, placementByGender: {}, placementByDegree: {} };
    }
    
    const workbookId = res.data.files[0].id;
    
    // Read Overall sheet (rows 2-5 are CS, EC, EE, ME)
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: workbookId,
      range: "Overall!A2:J5", // Columns A-J contain gender data
    });
    
    const rows = result.data.values || [];
    
    let totalMale = 0;
    let totalFemale = 0;
    let placedMale = 0;
    let placedFemale = 0;
    
    for (const row of rows) {
      if (!row[0]) continue;
      
      totalMale += parseInt(row[2]) || 0; // Column C: M
      totalFemale += parseInt(row[3]) || 0; // Column D: F
      placedMale += parseInt(row[8]) || 0; // Column I: M Placed
      placedFemale += parseInt(row[9]) || 0; // Column J: F Placed
    }
    
    const data = {
      gender: {
        M: totalMale,
        F: totalFemale,
        other: 0,
      },
      degree: {
        UG: totalMale + totalFemale, // For now, only reading UG data
        PG: 0,
      },
      placementByGender: {
        M: placedMale,
        F: placedFemale,
        other: 0,
      },
      placementByDegree: {
        UG: placedMale + placedFemale,
        PG: 0,
      },
    };

    return data;
  } catch (err) {
    console.error("[stats] getDemographicSplit error:", err);
    throw err;
  }
}

/**
 * Get trend analysis (year-over-year)
 */
export async function getTrendAnalysis() {
  try {
    const currentYear = parseInt(getAcademicYear().slice(0, 4), 10);
    const drive = await getDrive();
    const trends = [];

    // Check last 5 years
    for (let i = 0; i < 5; i++) {
      const year = currentYear - i;
      const nextYear = year + 1;
      const academicYear = `${year}-${String(nextYear).slice(-2)}`;
      
      try {
        // Try to get data for this year
        const [ugStudents, pgStudents, ugPlacements, pgPlacements] = await Promise.all([
          getStudentDataForYear(academicYear, "UG"),
          getStudentDataForYear(academicYear, "PG"),
          getPlacementDataForYear(academicYear, "UG"),
          getPlacementDataForYear(academicYear, "PG"),
        ]);
        
        const allStudents = [...ugStudents, ...pgStudents];
        const allPlacements = [...ugPlacements, ...pgPlacements];
        
        if (allStudents.length === 0) continue; // Skip if no data
        
        const placedRollNos = new Set(allPlacements.map(p => p.rollNo));
        const ctcValues = allPlacements.map(p => p.ctc).filter(c => c > 0);
        const companies = new Set(allPlacements.map(p => p.company));
        
        trends.push({
          academicYear,
          totalStudents: allStudents.length,
          placed: placedRollNos.size,
          averageCTC: ctcValues.length > 0
            ? (ctcValues.reduce((sum, c) => sum + c, 0) / ctcValues.length).toFixed(2)
            : 0,
          highestCTC: ctcValues.length > 0 ? Math.max(...ctcValues).toFixed(2) : 0,
          companiesVisited: companies.size,
        });
      } catch (err) {
        console.log(`No data for ${academicYear}`);
      }
    }

    return trends.reverse(); // Oldest to newest
  } catch (err) {
    console.error("[stats] getTrendAnalysis error:", err);
    throw err;
  }
}

/**
 * Helper: Get student data for specific academic year
 */
async function getStudentDataForYear(academicYear, degreeType) {
  try {
    const sheets = await getSheets();
    const drive = await getDrive();
    const name = `${degreeType}_Students_${academicYear}`;
    
    const res = await drive.files.list({
      q: `name='${name}' and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: "files(id, name)",
    });
    
    if (res.data.files.length === 0) return [];
    
    const workbookId = res.data.files[0].id;
    const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
    const branches = meta.data.sheets
      .map(s => s.properties.title)
      .filter(name => !name.startsWith("Sheet"));
    
    const students = [];
    
    for (const branch of branches) {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: workbookId,
          range: `${branch}!A2:Z`,
        });
        
        const rows = result.data.values || [];
        for (const row of rows) {
          if (row[0]) {
            students.push({
              rollNo: row[0],
              name: row[1],
              branch,
              degreeType,
              cgpa: parseFloat(row[2]) || 0,
            });
          }
        }
      } catch (err) {
        console.error(`Error reading branch ${branch}:`, err);
      }
    }
    
    return students;
  } catch (err) {
    return [];
  }
}

/**
 * Helper: Get placement data for specific academic year
 */
async function getPlacementDataForYear(academicYear, degreeType) {
  try {
    const drive = await getDrive();
    const sheets = await getSheets();
    const workbookName = `Placement_Data_${academicYear}_${degreeType}`;
    
    const res = await drive.files.list({
      q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: "files(id, name)",
    });
    
    if (res.data.files.length === 0) return [];
    
    const workbookId = res.data.files[0].id;
    
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: workbookId,
      range: "Placement_Data!A2:Z",
    });
    
    const rows = result.data.values || [];
    return rows.map(row => ({
      rollNo: row[0],
      company: row[3],
      ctc: parseFloat(row[4]) || 0,
    })).filter(p => p.rollNo && p.company);
  } catch (err) {
    return [];
  }
}

/**
 * Get company-wise offers
 */
export async function getCompanyStats() {
  try {
    const [ugPlacements, pgPlacements] = await Promise.all([
      getPlacementData("UG"),
      getPlacementData("PG"),
    ]);
    
    const allPlacements = [...ugPlacements, ...pgPlacements];
    
    // Group by company
    const companyMap = new Map();
    
    for (const placement of allPlacements) {
      if (!companyMap.has(placement.company)) {
        companyMap.set(placement.company, {
          company: placement.company,
          offers: [],
        });
      }
      companyMap.get(placement.company).offers.push(placement);
    }
    
    const companyStats = [];
    
    for (const [company, data] of companyMap.entries()) {
      const ctcValues = data.offers.map(o => o.ctc).filter(c => c > 0);
      
      companyStats.push({
        company,
        offers: data.offers.length,
        averageCTC: ctcValues.length > 0
          ? (ctcValues.reduce((sum, c) => sum + c, 0) / ctcValues.length).toFixed(2)
          : 0,
        highestCTC: ctcValues.length > 0 ? Math.max(...ctcValues).toFixed(2) : 0,
      });
    }

    return companyStats.sort((a, b) => b.offers - a.offers);
  } catch (err) {
    console.error("[stats] getCompanyStats error:", err);
    throw err;
  }
}

/**
 * Get placement snapshot (one-page summary)
 */
export async function getPlacementSnapshot() {
  try {
    const [overall, ugBranch, pgBranch, ctc, demographic, trends, companies] = await Promise.all([
      getOverallStats(),
      getBranchwiseStats("UG"),
      getBranchwiseStats("PG"),
      getCTCDistribution(),
      getDemographicSplit(),
      getTrendAnalysis(),
      getCompanyStats(),
    ]);

    return {
      academicYear: getAcademicYear(),
      generatedAt: new Date().toISOString(),
      overall,
      branchwise: { UG: ugBranch, PG: pgBranch },
      ctcDistribution: ctc,
      demographic,
      trends: trends.slice(-3), // Last 3 years
      topCompanies: companies.slice(0, 10),
    };
  } catch (err) {
    console.error("[stats] getPlacementSnapshot error:", err);
    throw err;
  }
}

/**
 * Force recalculate all formulas in a workbook
 * This fixes #N/A and #REF! errors by forcing Google Sheets to re-evaluate all formulas
 */
export async function forceRecalculateWorkbook(workbookId) {
  try {
    const sheets = await getSheets();
    
    console.log(`[recalculate] Starting recalculation for workbook ${workbookId}`);
    
    // Get all sheets in the workbook
    const workbook = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
    
    const updates = [];
    
    for (const sheet of workbook.data.sheets) {
      const sheetName = sheet.properties.title;
      
      try {
        // Read all formulas from the sheet
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: workbookId,
          range: `${sheetName}!A1:ZZ`,
          valueRenderOption: 'FORMULA'
        });
        
        if (response.data.values && response.data.values.length > 0) {
          // Write formulas back to force recalculation
          updates.push({
            range: `${sheetName}!A1:ZZ`,
            values: response.data.values
          });
        }
      } catch (err) {
        console.warn(`[recalculate] Error reading sheet ${sheetName}:`, err.message);
      }
    }
    
    if (updates.length > 0) {
      // Batch update all sheets at once
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: workbookId,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: updates
        }
      });
      
      console.log(`[recalculate] Successfully recalculated ${updates.length} sheets`);
    }
    
    return { success: true, sheetsRecalculated: updates.length };
  } catch (err) {
    console.error("[recalculate] Error:", err);
    throw err;
  }
}

/**
 * Recalculate all Placement_Data workbooks for the current academic year
 */
export async function recalculateAllPlacementWorkbooks() {
  try {
    const academicYear = getAcademicYear();
    const drive = await getDrive();
    const results = [];
    
    // Find all Placement_Data workbooks for current year
    const workbookNames = [
      `Placement_Data_${academicYear}_UG`,
      `Placement_Data_${academicYear}_PG`
    ];
    
    for (const workbookName of workbookNames) {
      try {
        const res = await drive.files.list({
          q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
          fields: "files(id, name)",
        });
        
        if (res.data.files.length > 0) {
          const workbookId = res.data.files[0].id;
          const result = await forceRecalculateWorkbook(workbookId);
          results.push({
            workbook: workbookName,
            ...result
          });
        } else {
          console.log(`[recalculateAll] Workbook not found: ${workbookName}`);
        }
      } catch (err) {
        console.error(`[recalculateAll] Error with ${workbookName}:`, err.message);
        results.push({
          workbook: workbookName,
          success: false,
          error: err.message
        });
      }
    }
    
    return results;
  } catch (err) {
    console.error("[recalculateAll] Error:", err);
    throw err;
  }
}
