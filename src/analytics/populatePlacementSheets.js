import { getSheets, getDrive } from "../sheets/sheets.dynamic.js";
import { SHEET_TEMPLATES } from "../constants/sheets.js";

/**
 * Normalize branch code to 2 letters (CS, EC, EE, ME)
 */
function normalizeBranchCode(branch) {
  const code = String(branch || "").toUpperCase().trim();
  if (code.length === 2) return code;
  if (code.length === 3) return code.slice(0, 2);
  return code;
}

/**
 * Populate Students_<branch> sheet in placement workbook during enrollment
 * Creates the placement workbook and sheets if they don't exist
 */
export async function populatePlacementStudentsSheet({ academicYear, degreeType, branch, students }) {
  const sheets = await getSheets();
  const drive = await getDrive();
  
  const workbookName = `Placement_Data_${academicYear}_${degreeType}`;
  // Normalize to 2 letters: CS, EC, EE, ME
  const branchCode = normalizeBranchCode(branch);
  
  console.log(`[populatePlacementStudents] Creating/updating ${workbookName} for branch ${branchCode}`);

  // Get or create placement workbook
  const workbookId = await getOrCreatePlacementWorkbook(drive, workbookName);
  
  // Ensure all placement sheets exist for this branch
  await ensurePlacementSheetsForBranch(sheets, workbookId, branchCode);
  
  // Get existing students to prevent duplicates
  const existingData = await sheets.spreadsheets.values.get({
    spreadsheetId: workbookId,
    range: `Students_${branchCode}!A:A`,
  });

  const existingRollNumbers = new Set(
    (existingData.data.values || []).slice(1).flat().filter(Boolean)
  );

  // Populate Students sheet with full student details (skip duplicates)
  const studentRows = students
    .filter(student => {
      const rollNo = student["Roll No."] || "";
      if (existingRollNumbers.has(rollNo)) {
        console.log(`[populatePlacementStudents] Skipping duplicate student in placement sheet: ${rollNo}`);
        return false;
      }
      return true;
    })
    .map(student => {
      const cgpa = parseFloat(student["CGPA"]) || 0; // Convert to number
      return [
        student["Roll No."] || "",
        student["Student Name"] || "",
        student["Gender"] || "",
        branchCode,
        cgpa, // Store as number, not string
        (cgpa >= 6.5) ? "Yes" : "No", // Eligible
        "", // Placement Status - filled when results published
        "", // Placement Type
        "", // Company
        "", // CTC
        "No", // Offer Revoked
      ];
    });

  if (studentRows.length > 0) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: workbookId,
      range: `Students_${branchCode}!A:K`,
      valueInputOption: "RAW",
      requestBody: {
        values: studentRows,
      },
    });
    
    console.log(`[populatePlacementStudents] Added ${studentRows.length} students to Students_${branchCode}`);
  }

  const skipped = students.length - studentRows.length;
  
  return { added: studentRows.length, skipped };
}

/**
 * Get or create placement workbook
 */
async function getOrCreatePlacementWorkbook(drive, name) {
  const folderId = process.env.PLACEMENT_FOLDER_ID;
  if (!folderId) {
    console.warn("PLACEMENT_FOLDER_ID not set, creating in root");
  }

  const query = folderId
    ? `name='${name}' and mimeType='application/vnd.google-apps.spreadsheet' and '${folderId}' in parents`
    : `name='${name}' and mimeType='application/vnd.google-apps.spreadsheet'`;

  const res = await drive.files.list({
    q: query,
    fields: "files(id, name)",
  });

  if (res.data.files.length > 0) {
    console.log(`[populatePlacementStudents] Found existing workbook ${name}`);
    return res.data.files[0].id;
  }

  const requestBody = {
    name,
    mimeType: "application/vnd.google-apps.spreadsheet",
  };

  if (folderId) {
    requestBody.parents = [folderId];
  }

  const create = await drive.files.create({
    requestBody,
  });

  console.log(`[populatePlacementStudents] Created new workbook ${name}`);
  return create.data.id;
}

/**
 * Ensure all placement sheets exist for a branch
 */
async function ensurePlacementSheetsForBranch(sheets, workbookId, branchCode) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: workbookId });
  const existingSheets = meta.data.sheets.map(s => s.properties.title);

  const requiredSheets = [
    { name: `Students_${branchCode}`, headers: SHEET_TEMPLATES.students },
    { name: `Offers_${branchCode}`, headers: SHEET_TEMPLATES.offers },
    { name: `Company_Drives_${branchCode}`, headers: SHEET_TEMPLATES.companyDrives },
    { name: `Placement_Stats_${branchCode}`, headers: SHEET_TEMPLATES.placementStats },
    { name: `CTC_Distribution_${branchCode}`, headers: SHEET_TEMPLATES.ctcDistribution },
  ];

  // Also ensure Overall sheet exists (only once per workbook)
  if (!existingSheets.includes('Overall')) {
    requiredSheets.push({ name: 'Overall', headers: SHEET_TEMPLATES.overall, isOverall: true });
  }

  for (const sheet of requiredSheets) {
    if (!existingSheets.includes(sheet.name)) {
      console.log(`[populatePlacementStudents] Creating sheet ${sheet.name}`);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: workbookId,
        requestBody: {
          requests: [{
            addSheet: {
              properties: { 
                title: sheet.name,
                gridProperties: { frozenRowCount: 1 },
              },
            },
          }],
        },
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: workbookId,
        range: `${sheet.name}!A1`,
        valueInputOption: "RAW",
        requestBody: { values: [sheet.headers] },
      });
      
      // If Overall sheet, populate with formula rows for all branches
      if (sheet.isOverall) {
        await setupOverallFormulas(sheets, workbookId);
      }
      
      // If Placement_Stats sheet, populate with formula rows
      if (sheet.name.startsWith('Placement_Stats_')) {
        await setupPlacementStatsFormulas(sheets, workbookId, branchCode);
      }
      
      // If CTC_Distribution sheet, populate with formula rows
      if (sheet.name.startsWith('CTC_Distribution_')) {
        await setupCTCDistributionFormulas(sheets, workbookId, branchCode);
      }
    }
  }
}

/**
 * Setup Overall sheet with formulas for all branches
 */
async function setupOverallFormulas(sheets, workbookId) {
  const branches = ['CS', 'EC', 'EE', 'ME'];
  const formulaRows = [];

  for (let i = 0; i < branches.length; i++) {
    const branch = branches[i];
    const row = i + 2; // Starting from row 2 (after header)
    const studSheet = `Students_${branch}`;

    formulaRows.push([
      branch,
      // Total counts
      `=COUNTA(${studSheet}!A2:A)`, // Total Students
      `=COUNTIF(${studSheet}!C2:C,"M")`, // Total M Students  
      `=COUNTIF(${studSheet}!C2:C,"F")`, // Total F Students
      
      // Eligible counts (Column F = "Yes")
      `=COUNTIF(${studSheet}!F2:F,"Yes")`, // Total Eligible
      `=COUNTIFS(${studSheet}!F2:F,"Yes",${studSheet}!C2:C,"M")`, // Total M Eligible
      `=COUNTIFS(${studSheet}!F2:F,"Yes",${studSheet}!C2:C,"F")`, // Total F Eligible
      
      // Placed counts (Column G = "Placed")
      `=COUNTIF(${studSheet}!G2:G,"Placed")`, // Total Placed
      `=COUNTIFS(${studSheet}!G2:G,"Placed",${studSheet}!C2:C,"M")`, // No of M Placed
      `=COUNTIFS(${studSheet}!G2:G,"Placed",${studSheet}!C2:C,"F")`, // No of F Placed
      
      // Placement percentages (row references: B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10)
      `=IF(B${row}=0,0,ROUND(H${row}/B${row}*100,2))`, // % Students Placed (of Total) = Total Placed / Total Students
      `=IF(E${row}=0,0,ROUND(H${row}/E${row}*100,2))`, // % Students Placed (of Eligible) = Total Placed / Total Eligible
      `=IF(C${row}=0,0,ROUND(I${row}/C${row}*100,2))`, // % M Placed (of Total) = M Placed / Total M
      `=IF(F${row}=0,0,ROUND(I${row}/F${row}*100,2))`, // % M Placed (of Eligible) = M Placed / M Eligible
      `=IF(D${row}=0,0,ROUND(J${row}/D${row}*100,2))`, // % F Placed (of Total) = F Placed / Total F
      `=IF(G${row}=0,0,ROUND(J${row}/G${row}*100,2))`, // % F Placed (of Eligible) = F Placed / F Eligible
      
      // CTC statistics (Column J has CTC values)
      `=IF(H${row}=0,"",MAXIFS(${studSheet}!J2:J,${studSheet}!G2:G,"Placed",${studSheet}!J2:J,">0"))`, // Highest CTC
      `=IF(H${row}=0,"",ROUND(AVERAGEIFS(${studSheet}!J2:J,${studSheet}!G2:G,"Placed",${studSheet}!J2:J,">0"),2))`, // Average CTC
      `=IF(H${row}=0,"",MINIFS(${studSheet}!J2:J,${studSheet}!G2:G,"Placed",${studSheet}!J2:J,">0"))`, // Lowest CTC
      `=IF(H${row}=0,"",ROUND((PERCENTILE(${studSheet}!J:J,0.5)),2))`, // Median CTC
      
      // Offer types (Column H has offer type)
      `=COUNTIFS(${studSheet}!G2:G,"Placed",${studSheet}!H2:H,"Internship")`, // Only Internship Offers
      `=COUNTIFS(${studSheet}!G2:G,"Placed",${studSheet}!H2:H,"FTE")`, // Only FTE Offers
      `=COUNTIFS(${studSheet}!G2:G,"Placed",${studSheet}!H2:H,"Both")`, // Both Offers
      
      // Unplaced by CGPA (Eligible but not placed, Column E has CGPA)
      `=COUNTIFS(${studSheet}!F2:F,"Yes",${studSheet}!G2:G,"<>Placed",${studSheet}!E2:E,">=8")`, // Unplaced CGPA >= 8
      `=COUNTIFS(${studSheet}!F2:F,"Yes",${studSheet}!G2:G,"<>Placed",${studSheet}!E2:E,">=7.5")`, // Unplaced CGPA >= 7.5
      `=COUNTIFS(${studSheet}!F2:F,"Yes",${studSheet}!G2:G,"<>Placed",${studSheet}!E2:E,">=7")`, // Unplaced CGPA >= 7
      `=COUNTIFS(${studSheet}!F2:F,"Yes",${studSheet}!G2:G,"<>Placed",${studSheet}!E2:E,">=6.5")`, // Unplaced CGPA >= 6.5
    ]);
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: workbookId,
    range: 'Overall!A2:AA5',
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: formulaRows },
  });

  console.log('[setupOverallFormulas] Created formula-based Overall sheet with comprehensive statistics');
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

  console.log(`[setupPlacementStatsFormulas] Created formulas for ${sheetName}`);
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

  console.log(`[setupCTCDistributionFormulas] Created formulas for ${sheetName}`);
}
