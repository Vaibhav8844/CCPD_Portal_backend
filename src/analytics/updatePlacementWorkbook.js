import { getSheets, getDrive } from "../sheets/sheets.dynamic.js";
import { getAcademicYear } from "../config/academicYear.js";

/**
 * Normalize branch code to 2 letters (CS, EC, EE, ME)
 */
function normalizeBranchCode(branch) {
  const code = String(branch || "").toUpperCase().trim();
  if (code.length === 2) return code;
  if (code.length === 3) return code.slice(0, 2);
  return code;
}

// Cache for sheet metadata to avoid repeated API calls
const sheetCache = new Map();
const CACHE_TTL = 5 * 60 * 1000; // 5 minutes

async function getCachedSheetMetadata(workbookId) {
  const cached = sheetCache.get(workbookId);
  if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
    return cached.data;
  }

  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
  sheetCache.set(workbookId, { data: meta.data, timestamp: Date.now() });
  return meta.data;
}

function invalidateCache(workbookId) {
  sheetCache.delete(workbookId);
}

/**
 * Update placement workbook with offer and student data
 * Students are pre-populated during enrollment
 * Stats will be calculated separately via recalculate endpoint
 */
export async function updatePlacementWorkbook(offerData) {
  const {
    rollNo,
    company,
    branch,
    degreeType, // UG or PG
    ctc,
    offerType, // FTE/Internship/Both
    requestId, // request ID for company drives tracking
  } = offerData;

  const academicYear = getAcademicYear();
  const sheets = await getSheets();
  const drive = await getDrive();

  console.log(`[updatePlacementWorkbook] Processing ${rollNo} for ${branch} (${degreeType})`);

  // 1. Get placement workbook
  const workbookName = `Placement_Data_${academicYear}_${degreeType}`;
  const res = await drive.files.list({
    q: `name='${workbookName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
    fields: "files(id, name)",
  });

  if (res.data.files.length === 0) {
    throw new Error(`Placement workbook ${workbookName} not found. Students must be enrolled first.`);
  }

  const workbookId = res.data.files[0].id;
  const branchCode = normalizeBranchCode(branch);

  // 2. Append to Offers sheet
  await appendToOffersSheet(workbookId, branchCode, {
    rollNo,
    company,
    ctc,
    offerType,
    timestamp: new Date().toISOString(),
  });

  // 3. Update existing student's placement info in Students sheet
  await updateStudentPlacementInfo(workbookId, branchCode, {
    rollNo,
    company,
    ctc,
    offerType,
  });

  // 5. Update Company_Drives sheet (if offerData has drive info)
  if (offerData.driveInfo) {
    await updateCompanyDrivesSheet(workbookId, branchCode, {
      company,
      requestId,
      rollNo,
      ...offerData.driveInfo, // Pass all drive metadata
    });
  } else if (requestId) {
    // Fallback: minimal data if driveInfo not provided
    await updateCompanyDrivesSheet(workbookId, branchCode, {
      company,
      requestId,
      rollNo,
    });
  }

  console.log(`[updatePlacementWorkbook] Completed for ${rollNo} - stats will be calculated on refresh`);
}

/**
 * Append to Offers_<branch> sheet
 */
async function appendToOffersSheet(workbookId, branchCode, offerData) {
  const sheets = await getSheets();
  const sheetName = `Offers_${branchCode}`;

  // Ensure sheet exists (with cache)
  let meta = await getCachedSheetMetadata(workbookId);
  let sheetExists = meta.sheets.some(
    s => s.properties.title === sheetName
  );

  if (!sheetExists) {
    console.log(`[appendToOffersSheet] Creating ${sheetName}`);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: workbookId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { 
                title: sheetName,
                gridProperties: { frozenRowCount: 1 },
              },
            },
          },
          {
            updateCells: {
              range: {
                sheetId: 0, // Will be updated by API
              },
              rows: [{
                values: [
                  { userEnteredValue: { stringValue: "Roll No" } },
                  { userEnteredValue: { stringValue: "Company" } },
                  { userEnteredValue: { stringValue: "Offer Type" } },
                  { userEnteredValue: { stringValue: "CTC (LPA)" } },
                  { userEnteredValue: { stringValue: "Offer Status" } },
                ],
              }],
              fields: "userEnteredValue",
            },
          },
        ],
      },
    });
    invalidateCache(workbookId);
  }

  // Check for duplicate offer (same roll + company + offer type)
  const existingOffers = await sheets.spreadsheets.values.get({
    spreadsheetId: workbookId,
    range: `${sheetName}!A:E`,
  });

  const offerExists = (existingOffers.data.values || []).some((row, i) => {
    if (i === 0) return false; // Skip header
    return row[0] === offerData.rollNo && 
           row[1] === offerData.company && 
           row[2] === offerData.offerType &&
           row[4] === "Active"; // Only check active offers
  });

  if (offerExists) {
    console.log(`[appendToOffersSheet] Duplicate offer detected for ${offerData.rollNo} from ${offerData.company}, skipping`);
    return;
  }

  // Append offer - columns: Roll No, Company, Offer Type, CTC (LPA), Offer Status
  await sheets.spreadsheets.values.append({
    spreadsheetId: workbookId,
    range: `${sheetName}!A:E`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[
        offerData.rollNo,
        offerData.company,
        offerData.offerType,
        offerData.ctc,
        "Active",
      ]],
    },
  });

  console.log(`[appendToOffersSheet] Added offer for ${offerData.rollNo} to ${sheetName}`);
}

/**
 * Update placement info for existing student in Students_<branch> sheet
 * Students are pre-populated during enrollment, so we only update placement fields
 */
async function updateStudentPlacementInfo(workbookId, branchCode, placementData) {
  const sheets = await getSheets();
  const sheetName = `Students_${branchCode}`;

  try {
    // Find student row by roll number
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: workbookId,
      range: `${sheetName}!A:K`,
    });

    const rows = result.data.values || [];
    if (rows.length === 0) {
      console.warn(`[updateStudentPlacementInfo] Sheet ${sheetName} is empty`);
      return;
    }

    const studentRowIndex = rows.findIndex((r, i) => i > 0 && r[0] === placementData.rollNo);

    if (studentRowIndex === -1) {
      console.warn(`[updateStudentPlacementInfo] Student ${placementData.rollNo} not found in ${sheetName}`);
      return;
    }

    // Update only placement-related columns (G-K): Placement Status, Placement Type, Company, Highest CTC, Offer Revoked
    // Keep existing values for basic info (Roll No, Name, Gender, Branch, CGPA, Eligible)
    const existingRow = rows[studentRowIndex];
    const currentCTC = parseFloat(existingRow[9]) || 0;
    const newCTC = parseFloat(placementData.ctc) || 0;
    const highestCTC = Math.max(currentCTC, newCTC);

    await sheets.spreadsheets.values.update({
      spreadsheetId: workbookId,
      range: `${sheetName}!G${studentRowIndex + 1}:K${studentRowIndex + 1}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          "Placed", // Placement Status
          placementData.offerType || "FTE", // Placement Type
          placementData.company, // Company
          highestCTC, // Highest CTC
          "No", // Offer Revoked
        ]],
      },
    });

    console.log(`[updateStudentPlacementInfo] Updated placement info for ${placementData.rollNo} in ${sheetName}`);
  } catch (err) {
    console.error(`[updateStudentPlacementInfo] Error:`, err.message);
  }
}

/**
 * Update Company_Drives_<branch> sheet
 * Matches SHEET_TEMPLATES.companyDrives (20 columns)
 */
async function updateCompanyDrivesSheet(workbookId, branchCode, driveData) {
  const sheets = await getSheets();
  const sheetName = `Company_Drives_${branchCode}`;

  try {
    // Ensure sheet exists (with cache)
    let meta = await getCachedSheetMetadata(workbookId);
    let sheetExists = meta.sheets.some(
      s => s.properties.title === sheetName
    );

    if (!sheetExists) {
      console.log(`[updateCompanyDrivesSheet] Creating ${sheetName}`);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: workbookId,
        requestBody: {
          requests: [{
            addSheet: {
              properties: { 
                title: sheetName,
                gridProperties: { frozenRowCount: 1 },
              },
            },
          }],
        },
      });

      // Add headers - Match SHEET_TEMPLATES.companyDrives
      await sheets.spreadsheets.values.update({
        spreadsheetId: workbookId,
        range: `${sheetName}!A1:T1`,
        valueInputOption: "RAW",
        requestBody: {
          values: [[
            "Company", "SPOC", "Request ID", "Drive Type", "Eligible Pool",
            "PPT Datetime", "OT Datetime", "Interview Datetime",
            "PPT Status", "OT Status", "INTERVIEW Status",
            "Internship Stipend", "FTE CTC", "FTE Base", "Expected Hires",
            "Actual Hires", "Drive Status", "Results Published",
            "Results Published At", "Last Updated"
          ]],
        },
      });
      invalidateCache(workbookId);
    }

    // Check if drive already exists for this request
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: workbookId,
      range: `${sheetName}!A:T`,
    });

    const rows = result.data.values || [];
    const driveRow = rows.findIndex((r, i) => 
      i > 0 && r[2] === driveData.requestId // Request ID is column 2 (index 2)
    );

    const now = new Date().toISOString();

    if (driveRow > 0) {
      // Update actual hires count and last updated
      const currentHires = parseInt(rows[driveRow][15]) || 0; // Actual Hires at index 15
      await sheets.spreadsheets.values.update({
        spreadsheetId: workbookId,
        range: `${sheetName}!P${driveRow + 1}:T${driveRow + 1}`, // Columns P-T (Actual Hires through Last Updated)
        valueInputOption: "RAW",
        requestBody: {
          values: [[
            currentHires + 1, // Actual Hires
            driveData.driveStatus || "In Progress", // Drive Status
            driveData.resultsPublished ? "Yes" : "No", // Results Published
            driveData.resultsPublished ? now : "", // Results Published At
            now, // Last Updated
          ]],
        },
      });
    } else {
      // Append new drive entry
      await sheets.spreadsheets.values.append({
        spreadsheetId: workbookId,
        range: `${sheetName}!A:T`,
        valueInputOption: "RAW",
        requestBody: {
          values: [[
            driveData.company,
            driveData.spoc || "",
            driveData.requestId,
            driveData.driveType || "",
            driveData.eligiblePool || "",
            driveData.pptDatetime || "",
            driveData.otDatetime || "",
            driveData.interviewDatetime || "",
            driveData.pptStatus || "",
            driveData.otStatus || "",
            driveData.interviewStatus || "",
            driveData.internshipStipend || "",
            driveData.fteCTC || "",
            driveData.fteBase || "",
            driveData.expectedHires || "",
            1, // Actual Hires (first offer)
            driveData.driveStatus || "In Progress",
            driveData.resultsPublished ? "Yes" : "No",
            driveData.resultsPublished ? now : "",
            now, // Last Updated
          ]],
        },
      });
    }

    console.log(`[updateCompanyDrivesSheet] Updated ${sheetName} for ${driveData.company}`);
  } catch (err) {
    console.error(`[updateCompanyDrivesSheet] Error:`, err.message);
  }
}

/**
 * Recalculate and update Placement_Stats_<branch> sheet
 * EXPORTED for manual recalculation via API
 */
export async function recalculateBranchStats(workbookId, branchCode, degreeType, academicYear) {
  const sheets = await getSheets();

  console.log(`[recalculateBranchStats] Setting up formulas for ${branchCode}`);

  // Set up formula-based sheets
  await setupPlacementStatsFormulas(sheets, workbookId, branchCode);
  await setupCTCDistributionFormulas(sheets, workbookId, branchCode);
  
  console.log(`[recalculateBranchStats] Formulas configured for ${branchCode}`);
}

/**
 * Setup Placement_Stats sheet with formulas
 */
async function setupPlacementStatsFormulas(sheets, workbookId, branchCode) {
  const sheetName = `Placement_Stats_${branchCode}`;
  const studentSheet = `Students_${branchCode}`;

  const statsRows = [
    ['Total Students', `=COUNTA(${studentSheet}!A2:A)`],
    ['Total M Students', `=COUNTIF(${studentSheet}!C2:C,"M")`],
    ['Total F Students', `=COUNTIF(${studentSheet}!C2:C,"F")`],
    ['Total Eligible', `=COUNTIF(${studentSheet}!F2:F,"Yes")`],
    ['Total M Eligible', `=COUNTIFS(${studentSheet}!F2:F,"Yes",${studentSheet}!C2:C,"M")`],
    ['Total F Eligible', `=COUNTIFS(${studentSheet}!F2:F,"Yes",${studentSheet}!C2:C,"F")`],
    ['Total Placed', `=COUNTIF(${studentSheet}!G2:G,"Placed")`],
    ['No of M Placed', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!C2:C,"M")`],
    ['No of F Placed', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!C2:C,"F")`],
    ['% Students Placed (of Total)', `=IF(B2=0,0,ROUND(B8/B2*100,2))`],
    ['% Students Placed (of Eligible)', `=IF(B5=0,0,ROUND(B8/B5*100,2))`],
    ['% M Placed (of Total)', `=IF(B3=0,0,ROUND(B9/B3*100,2))`],
    ['% M Placed (of Eligible)', `=IF(B6=0,0,ROUND(B9/B6*100,2))`],
    ['% F Placed (of Total)', `=IF(B4=0,0,ROUND(B10/B4*100,2))`],
    ['% F Placed (of Eligible)', `=IF(B7=0,0,ROUND(B10/B7*100,2))`],
    ['Highest CTC (LPA)', `=IF(B8=0,"",MAXIFS(${studentSheet}!J2:J,${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">0"))`],
    ['Average CTC (LPA)', `=IF(B8=0,"",ROUND(AVERAGEIFS(${studentSheet}!J2:J,${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">0"),2))`],
    ['Lowest CTC (LPA)', `=IF(B8=0,"",MINIFS(${studentSheet}!J2:J,${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">0"))`],
    ['Median CTC (LPA)', `=IF(B8=0,"",ROUND(PERCENTILE(${studentSheet}!J:J,0.5),2))`],
    ['Only Internship Offers', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!H2:H,"Internship")`],
    ['Only FTE Offers', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!H2:H,"FTE")`],
    ['Both Offers', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!H2:H,"Both")`],
    ['Unplaced (CGPA >= 8)', `=COUNTIFS(${studentSheet}!F2:F,"Yes",${studentSheet}!G2:G,"<>Placed",${studentSheet}!E2:E,">=8")`],
    ['Unplaced (CGPA >= 7.5)', `=COUNTIFS(${studentSheet}!F2:F,"Yes",${studentSheet}!G2:G,"<>Placed",${studentSheet}!E2:E,">=7.5")`],
    ['Unplaced (CGPA >= 7)', `=COUNTIFS(${studentSheet}!F2:F,"Yes",${studentSheet}!G2:G,"<>Placed",${studentSheet}!E2:E,">=7")`],
    ['Unplaced (CGPA >= 6.5)', `=COUNTIFS(${studentSheet}!F2:F,"Yes",${studentSheet}!G2:G,"<>Placed",${studentSheet}!E2:E,">=6.5")`],
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: workbookId,
    range: `${sheetName}!A2:B27`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: statsRows },
  });
  
  console.log(`[setupPlacementStatsFormulas] Configured formulas for ${sheetName}`);
}

/**
 * Setup CTC_Distribution sheet with formulas
 */
async function setupCTCDistributionFormulas(sheets, workbookId, branchCode) {
  const sheetName = `CTC_Distribution_${branchCode}`;
  const studentSheet = `Students_${branchCode}`;

  const ctcRanges = [
    ['0-5 LPA', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">=0",${studentSheet}!J2:J,"<5")`],
    ['5-10 LPA', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">=5",${studentSheet}!J2:J,"<10")`],
    ['10-15 LPA', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">=10",${studentSheet}!J2:J,"<15")`],
    ['15-20 LPA', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">=15",${studentSheet}!J2:J,"<20")`],
    ['20-30 LPA', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">=20",${studentSheet}!J2:J,"<30")`],
    ['30+ LPA', `=COUNTIFS(${studentSheet}!G2:G,"Placed",${studentSheet}!J2:J,">=30")`],
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: workbookId,
    range: `${sheetName}!A2:B7`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: ctcRanges },
  });
  
  console.log(`[setupCTCDistributionFormulas] Configured formulas for ${sheetName}`);
}

/**
 * Helper: Get sheet data
 */
async function getSheetData(workbookId, sheetName) {
  try {
    const sheets = await getSheets();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: workbookId,
      range: `${sheetName}!A:Z`,
    });
    return result.data.values || [];
  } catch (err) {
    console.warn(`[getSheetData] Sheet ${sheetName} not found or empty`);
    return [];
  }
}
