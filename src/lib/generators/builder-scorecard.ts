import ExcelJS from 'exceljs';

// Color scheme
const COLORS = {
  primaryBlue: 'FF3378FF',
  orange: 'FFf97316',
  green: 'FF22c55e',
  slateHeader: 'FF1e293b',
  inputCell: 'FFfff7ed',
  white: 'FFFFFFFF',
  lightGray: 'FFf1f5f9',
  red: 'FFef4444',
  yellow: 'FFfbbf24',
};

const headerStyle: Partial<ExcelJS.Style> = {
  font: { bold: true, color: { argb: COLORS.white }, size: 11 },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } },
  alignment: { horizontal: 'center', vertical: 'middle' },
  border: {
    top: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    left: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    right: { style: 'thin', color: { argb: 'FFe2e8f0' } },
  },
};

const inputStyle: Partial<ExcelJS.Style> = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.inputCell } },
  border: {
    top: { style: 'thin', color: { argb: COLORS.orange } },
    left: { style: 'thin', color: { argb: COLORS.orange } },
    bottom: { style: 'thin', color: { argb: COLORS.orange } },
    right: { style: 'thin', color: { argb: COLORS.orange } },
  },
  alignment: { horizontal: 'center', vertical: 'middle' },
};

const currencyFormat = '"£"#,##0';
const percentFormat = '0%';

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    left: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    right: { style: 'thin', color: { argb: 'FFe2e8f0' } },
  };
}

// Scoring criteria with weights
const scoringCriteria = [
  { category: 'QUALIFICATIONS & EXPERIENCE', items: [] },
  { item: 'Years in business', weight: 8 },
  { item: 'Relevant project experience', weight: 10 },
  { item: 'Technical knowledge demonstrated', weight: 8 },
  { item: 'Portfolio quality', weight: 7 },
  { category: 'INSURANCE & ACCREDITATION', items: [] },
  { item: 'Public liability insurance (min £2m)', weight: 10, isChecklist: true },
  { item: 'Employers liability insurance', weight: 8, isChecklist: true },
  { item: 'Professional indemnity insurance', weight: 6, isChecklist: true },
  { item: 'FMB / TrustMark / Which? Trusted', weight: 7, isChecklist: true },
  { item: 'Gas Safe registered (if applicable)', weight: 5, isChecklist: true },
  { item: 'NICEIC/NAPIT registered (if applicable)', weight: 5, isChecklist: true },
  { category: 'QUOTE & CONTRACT', items: [] },
  { item: 'Quote detail and clarity', weight: 9 },
  { item: 'Competitive pricing', weight: 8 },
  { item: 'Payment terms reasonableness', weight: 9 },
  { item: 'Contract terms clarity', weight: 8 },
  { item: 'Warranty offered', weight: 7 },
  { category: 'COMMUNICATION & PROFESSIONALISM', items: [] },
  { item: 'Response time', weight: 7 },
  { item: 'Communication clarity', weight: 8 },
  { item: 'Professionalism during visit', weight: 7 },
  { item: 'Willingness to answer questions', weight: 6 },
  { category: 'REFERENCES & REVIEWS', items: [] },
  { item: 'Quality of references provided', weight: 9 },
  { item: 'Online reviews rating', weight: 7 },
  { item: 'Verified past project visits', weight: 8 },
  { category: 'AVAILABILITY & TIMELINE', items: [] },
  { item: 'Start date availability', weight: 6 },
  { item: 'Proposed timeline realism', weight: 8 },
  { item: 'Team/subcontractor availability', weight: 6 },
];

function createBuilderSheet(workbook: ExcelJS.Workbook, builderNum: number) {
  const sheetName = `Builder ${builderNum}`;
  const sheet = workbook.addWorksheet(sheetName, {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  sheet.columns = [
    { width: 3 },
    { width: 35 },
    { width: 10 },
    { width: 12 },
    { width: 12 },
    { width: 30 },
  ];

  // Title
  sheet.mergeCells('B2:F2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = `BUILDER ${builderNum} SCORECARD`;
  titleCell.font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  // Builder details section
  sheet.getCell('B4').value = 'Company Name:';
  sheet.getCell('B4').font = { bold: true };
  const nameCell = sheet.getCell('C4');
  Object.assign(nameCell, { style: inputStyle });
  sheet.mergeCells('C4:E4');

  sheet.getCell('B5').value = 'Contact Person:';
  sheet.getCell('B5').font = { bold: true };
  const contactCell = sheet.getCell('C5');
  Object.assign(contactCell, { style: inputStyle });
  sheet.mergeCells('C5:E5');

  sheet.getCell('B6').value = 'Phone:';
  sheet.getCell('B6').font = { bold: true };
  const phoneCell = sheet.getCell('C6');
  Object.assign(phoneCell, { style: inputStyle });

  sheet.getCell('D6').value = 'Email:';
  sheet.getCell('D6').font = { bold: true };
  const emailCell = sheet.getCell('E6');
  Object.assign(emailCell, { style: inputStyle });

  sheet.getCell('B7').value = 'Quote Amount:';
  sheet.getCell('B7').font = { bold: true };
  const quoteCell = sheet.getCell('C7');
  Object.assign(quoteCell, { style: inputStyle });
  quoteCell.numFmt = currencyFormat;

  sheet.getCell('D7').value = 'Quote Date:';
  sheet.getCell('D7').font = { bold: true };
  const quoteDateCell = sheet.getCell('E7');
  Object.assign(quoteDateCell, { style: inputStyle });
  quoteDateCell.numFmt = 'dd/mm/yyyy';

  // Headers
  const headerRow = sheet.getRow(9);
  ['', 'Criteria', 'Weight', 'Score (1-5)', 'Weighted', 'Notes'].forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Scoring rows
  let rowNum = 10;
  let scoreRowsStart = 0;
  let scoreRowsEnd = 0;

  scoringCriteria.forEach((criterion) => {
    const row = sheet.getRow(rowNum);

    if ('category' in criterion) {
      // Category header
      sheet.mergeCells(`B${rowNum}:F${rowNum}`);
      row.getCell(2).value = criterion.category;
      row.getCell(2).font = { bold: true, size: 11, color: { argb: COLORS.slateHeader } };
      row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
    } else {
      // Scoring item
      if (scoreRowsStart === 0) scoreRowsStart = rowNum;
      scoreRowsEnd = rowNum;

      row.getCell(2).value = criterion.item;
      applyBorder(row.getCell(2));

      const weightCell = row.getCell(3);
      weightCell.value = criterion.weight;
      weightCell.alignment = { horizontal: 'center' };
      weightCell.numFmt = '0';
      applyBorder(weightCell);

      const scoreCell = row.getCell(4);
      if (criterion.isChecklist) {
        // Yes/No dropdown for checklist items
        scoreCell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"Yes,No,N/A"'],
        };
        // Convert Yes=5, No=1, N/A=0 in weighted column
      } else {
        // 1-5 score dropdown
        scoreCell.dataValidation = {
          type: 'whole',
          operator: 'between',
          allowBlank: true,
          showErrorMessage: true,
          errorTitle: 'Invalid Score',
          error: 'Please enter a score between 1 and 5',
          formulae: [1, 5],
        };
      }
      Object.assign(scoreCell, { style: inputStyle });
      if (!criterion.isChecklist) {
        scoreCell.numFmt = '0';
      }

      const weightedCell = row.getCell(5);
      if (criterion.isChecklist) {
        weightedCell.value = { formula: `IF(D${rowNum}="Yes",C${rowNum}*5,IF(D${rowNum}="No",C${rowNum}*1,0))` };
      } else {
        weightedCell.value = { formula: `IF(D${rowNum}="",0,C${rowNum}*D${rowNum})` };
      }
      weightedCell.alignment = { horizontal: 'center' };
      weightedCell.numFmt = '0';
      applyBorder(weightedCell);

      const notesCell = row.getCell(6);
      Object.assign(notesCell, { style: inputStyle });
      notesCell.alignment = { horizontal: 'left' };
    }

    rowNum++;
  });

  // Total row
  rowNum += 1;
  const totalRow = sheet.getRow(rowNum);
  totalRow.getCell(2).value = 'TOTAL SCORE';
  totalRow.getCell(2).font = { bold: true, size: 12 };
  totalRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
  totalRow.getCell(2).font = { bold: true, color: { argb: COLORS.white } };
  applyBorder(totalRow.getCell(2));

  const totalWeightCell = totalRow.getCell(3);
  totalWeightCell.value = { formula: `SUM(C${scoreRowsStart}:C${scoreRowsEnd})` };
  totalWeightCell.font = { bold: true, color: { argb: COLORS.white } };
  totalWeightCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
  totalWeightCell.alignment = { horizontal: 'center' };
  applyBorder(totalWeightCell);

  totalRow.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
  applyBorder(totalRow.getCell(4));

  const totalScoreCell = totalRow.getCell(5);
  totalScoreCell.value = { formula: `SUM(E${scoreRowsStart}:E${scoreRowsEnd})` };
  totalScoreCell.font = { bold: true, color: { argb: COLORS.white } };
  totalScoreCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
  totalScoreCell.alignment = { horizontal: 'center' };
  applyBorder(totalScoreCell);

  totalRow.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
  applyBorder(totalRow.getCell(6));

  // Percentage score
  rowNum += 1;
  const pctRow = sheet.getRow(rowNum);
  pctRow.getCell(2).value = 'PERCENTAGE SCORE';
  pctRow.getCell(2).font = { bold: true };

  const pctCell = pctRow.getCell(5);
  pctCell.value = { formula: `E${rowNum - 1}/(C${rowNum - 1}*5)` };
  pctCell.numFmt = percentFormat;
  pctCell.font = { bold: true, size: 14, color: { argb: COLORS.primaryBlue } };
  pctCell.alignment = { horizontal: 'center' };

  // Rating interpretation
  rowNum += 2;
  sheet.getCell(`B${rowNum}`).value = 'Rating:';
  sheet.getCell(`B${rowNum}`).font = { bold: true };
  const ratingCell = sheet.getCell(`C${rowNum}`);
  ratingCell.value = { formula: `IF(E${rowNum - 3}/(C${rowNum - 3}*5)>=0.8,"Excellent",IF(E${rowNum - 3}/(C${rowNum - 3}*5)>=0.6,"Good",IF(E${rowNum - 3}/(C${rowNum - 3}*5)>=0.4,"Fair","Poor")))` };
  ratingCell.font = { bold: true, size: 12 };

  // Conditional formatting for percentage
  sheet.addConditionalFormatting({
    ref: `E${rowNum - 2}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0.79'],
        style: { font: { color: { argb: COLORS.green } } },
        priority: 1,
      },
      {
        type: 'cellIs',
        operator: 'between',
        formulae: ['0.6', '0.79'],
        style: { font: { color: { argb: COLORS.primaryBlue } } },
        priority: 2,
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0.6'],
        style: { font: { color: { argb: COLORS.red } } },
        priority: 3,
      },
    ],
  });

  // Additional notes section
  rowNum += 2;
  sheet.getCell(`B${rowNum}`).value = 'ADDITIONAL NOTES';
  sheet.getCell(`B${rowNum}`).font = { bold: true, size: 12 };

  rowNum += 1;
  sheet.mergeCells(`B${rowNum}:F${rowNum + 3}`);
  const notesArea = sheet.getCell(`B${rowNum}`);
  Object.assign(notesArea, { style: inputStyle });
  notesArea.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };

  // Red flags section
  rowNum += 5;
  sheet.getCell(`B${rowNum}`).value = 'RED FLAGS (tick if observed)';
  sheet.getCell(`B${rowNum}`).font = { bold: true, size: 12, color: { argb: COLORS.red } };

  const redFlags = [
    'Requests large upfront payment (>25%)',
    'No written contract offered',
    'Cannot provide insurance certificates',
    'No references available',
    'Pressure to sign quickly',
    'Vague about subcontractors',
    'Cash-only payment requests',
  ];

  redFlags.forEach((flag, index) => {
    rowNum++;
    const flagRow = sheet.getRow(rowNum);
    const checkCell = flagRow.getCell(2);
    checkCell.value = flag;
    applyBorder(checkCell);

    const tickCell = flagRow.getCell(3);
    tickCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Yes,No"'],
    };
    Object.assign(tickCell, { style: inputStyle });
  });

  return sheet;
}

export async function generateBuilderScorecard(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Extension Survival Guide';
  workbook.created = new Date();

  // ===== INSTRUCTIONS SHEET =====
  const instructionsSheet = workbook.addWorksheet('Instructions', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  instructionsSheet.columns = [
    { width: 5 },
    { width: 80 },
  ];

  const instructions = [
    ['', ''],
    ['', 'BUILDER VETTING SCORECARD'],
    ['', ''],
    ['', 'This scorecard helps you objectively evaluate and compare builders for your'],
    ['', 'extension project. Use it to make an informed decision based on weighted criteria.'],
    ['', ''],
    ['', 'HOW TO USE:'],
    ['', ''],
    ['', '1. Complete one scorecard for each builder you are considering'],
    ['', '2. Score each criterion from 1-5 (or Yes/No for checklist items)'],
    ['', '   1 = Very Poor, 2 = Poor, 3 = Average, 4 = Good, 5 = Excellent'],
    ['', '3. Add notes for important observations'],
    ['', '4. Check any red flags observed'],
    ['', '5. Review the Comparison sheet for automatic ranking'],
    ['', ''],
    ['', 'SCORING GUIDE:'],
    ['', ''],
    ['', 'Excellent (80%+): Strong candidate, likely good choice'],
    ['', 'Good (60-79%): Solid option, verify any weak areas'],
    ['', 'Fair (40-59%): Proceed with caution, significant gaps'],
    ['', 'Poor (<40%): Not recommended, look elsewhere'],
    ['', ''],
    ['', 'IMPORTANT CHECKS BEFORE APPOINTING:'],
    ['', ''],
    ['', '- Verify insurance certificates are current and adequate'],
    ['', '- Check Companies House for company registration'],
    ['', '- Visit at least one completed project if possible'],
    ['', '- Get at least 3 quotes for comparison'],
    ['', '- Read and understand the contract before signing'],
    ['', '- Never pay more than 10-15% upfront'],
  ];

  instructions.forEach((row, index) => {
    const excelRow = instructionsSheet.addRow(row);
    if (index === 1) {
      excelRow.getCell(2).font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };
    } else if (row[1].startsWith('HOW TO USE') || row[1].startsWith('SCORING GUIDE') || row[1].startsWith('IMPORTANT')) {
      excelRow.getCell(2).font = { bold: true, size: 12, color: { argb: COLORS.slateHeader } };
    }
  });

  // ===== CREATE 10 BUILDER SHEETS =====
  for (let i = 1; i <= 10; i++) {
    createBuilderSheet(workbook, i);
  }

  // ===== COMPARISON SHEET =====
  const comparisonSheet = workbook.addWorksheet('Comparison', {
    properties: { tabColor: { argb: COLORS.green } },
  });

  comparisonSheet.columns = [
    { width: 5 },
    { width: 8 },
    { width: 25 },
    { width: 15 },
    { width: 12 },
    { width: 15 },
    { width: 20 },
  ];

  // Title
  comparisonSheet.mergeCells('B2:G2');
  const titleCell = comparisonSheet.getCell('B2');
  titleCell.value = 'BUILDER COMPARISON & RANKING';
  titleCell.font = { bold: true, size: 16, color: { argb: COLORS.green } };

  // Headers
  const compHeaders = ['', 'Rank', 'Builder Name', 'Quote', 'Score %', 'Rating', 'Notes'];
  const compHeaderRow = comparisonSheet.getRow(4);
  compHeaders.forEach((header, index) => {
    const cell = compHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Builder comparison rows
  for (let i = 1; i <= 10; i++) {
    const rowNum = i + 4;
    const row = comparisonSheet.getRow(rowNum);

    // Rank (will be calculated)
    const rankCell = row.getCell(2);
    rankCell.value = { formula: `IF(E${rowNum}="","-",RANK(E${rowNum},$E$5:$E$14,0))` };
    rankCell.alignment = { horizontal: 'center' };
    applyBorder(rankCell);

    // Builder name from sheet
    const nameCell = row.getCell(3);
    nameCell.value = { formula: `IF('Builder ${i}'!C4="","Builder ${i}",'Builder ${i}'!C4)` };
    applyBorder(nameCell);

    // Quote
    const quoteCell = row.getCell(4);
    quoteCell.value = { formula: `'Builder ${i}'!C7` };
    quoteCell.numFmt = currencyFormat;
    applyBorder(quoteCell);

    // Score percentage - find the percentage cell (it's at row 43 after all criteria)
    const scoreCell = row.getCell(5);
    scoreCell.value = { formula: `IF('Builder ${i}'!E42=0,"",'Builder ${i}'!E43)` };
    scoreCell.numFmt = percentFormat;
    applyBorder(scoreCell);

    // Rating
    const ratingCell = row.getCell(6);
    ratingCell.value = { formula: `IF(E${rowNum}="","",IF(E${rowNum}>=0.8,"Excellent",IF(E${rowNum}>=0.6,"Good",IF(E${rowNum}>=0.4,"Fair","Poor"))))` };
    applyBorder(ratingCell);

    // Notes
    const notesCell = row.getCell(7);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left' };
  }

  // Conditional formatting for ranking
  comparisonSheet.addConditionalFormatting({
    ref: 'B5:B14',
    rules: [
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['1'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { bold: true, color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['2'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
        },
        priority: 2,
      },
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['3'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfed7aa' } },
        },
        priority: 3,
      },
    ],
  });

  // Score color coding
  comparisonSheet.addConditionalFormatting({
    ref: 'E5:E14',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0.79'],
        style: { font: { color: { argb: COLORS.green }, bold: true } },
        priority: 1,
      },
      {
        type: 'cellIs',
        operator: 'between',
        formulae: ['0.6', '0.79'],
        style: { font: { color: { argb: COLORS.primaryBlue } } },
        priority: 2,
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0.6'],
        style: { font: { color: { argb: COLORS.red } } },
        priority: 3,
      },
    ],
  });

  // Summary section
  comparisonSheet.getCell('B17').value = 'RECOMMENDATION';
  comparisonSheet.getCell('B17').font = { bold: true, size: 14 };

  comparisonSheet.getCell('B19').value = 'Highest Scored:';
  comparisonSheet.getCell('B19').font = { bold: true };
  const topBuilderCell = comparisonSheet.getCell('C19');
  topBuilderCell.value = { formula: 'INDEX(C5:C14,MATCH(MAX(E5:E14),E5:E14,0))' };
  topBuilderCell.font = { bold: true, color: { argb: COLORS.green } };

  comparisonSheet.getCell('B20').value = 'Lowest Quote:';
  comparisonSheet.getCell('B20').font = { bold: true };
  const lowestQuoteCell = comparisonSheet.getCell('C20');
  lowestQuoteCell.value = { formula: 'INDEX(C5:C14,MATCH(MIN(IF(D5:D14>0,D5:D14)),IF(D5:D14>0,D5:D14),0))' };

  comparisonSheet.getCell('B21').value = 'Best Value:';
  comparisonSheet.getCell('B21').font = { bold: true };
  comparisonSheet.getCell('C21').value = '(Top scored with competitive price)';
  comparisonSheet.getCell('C21').font = { italic: true, color: { argb: 'FF64748b' } };

  // ===== SCORING GUIDE SHEET =====
  const guideSheet = workbook.addWorksheet('Scoring Guide', {
    properties: { tabColor: { argb: 'FF94a3b8' } },
  });

  guideSheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 60 },
  ];

  guideSheet.mergeCells('B2:C2');
  guideSheet.getCell('B2').value = 'SCORING GUIDE';
  guideSheet.getCell('B2').font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  const guideData = [
    ['', 'Criterion', 'Scoring Guidance'],
    ['', '', ''],
    ['', 'Years in business', '1: <1yr, 2: 1-2yrs, 3: 3-5yrs, 4: 5-10yrs, 5: 10+yrs'],
    ['', 'Relevant project experience', '1: None, 3: Some similar, 5: Extensive similar projects'],
    ['', 'Technical knowledge', '1: Basic, 3: Competent, 5: Expert understanding'],
    ['', 'Portfolio quality', '1: None/poor, 3: Adequate, 5: Impressive quality'],
    ['', '', ''],
    ['', 'Quote detail and clarity', '1: Vague, 3: Basic breakdown, 5: Fully itemized'],
    ['', 'Competitive pricing', '1: Very high, 3: Average, 5: Best value'],
    ['', 'Payment terms', '1: >30% upfront, 3: 20%, 5: Stage payments only'],
    ['', 'Contract clarity', '1: No contract, 3: Basic, 5: Comprehensive'],
    ['', 'Warranty offered', '1: None, 3: 1yr, 5: 2yr+ with insurance backing'],
    ['', '', ''],
    ['', 'Response time', '1: >1 week, 3: 2-3 days, 5: Same/next day'],
    ['', 'Communication clarity', '1: Poor, 3: Clear, 5: Excellent and proactive'],
    ['', 'Professionalism', '1: Unprofessional, 3: Professional, 5: Exemplary'],
    ['', '', ''],
    ['', 'Reference quality', '1: None, 3: Generic, 5: Detailed and verifiable'],
    ['', 'Online reviews', '1: Poor/none, 3: Good, 5: Excellent across platforms'],
    ['', 'Verified projects', '1: None offered, 3: Photos only, 5: Site visit arranged'],
  ];

  guideData.forEach((row, index) => {
    const excelRow = guideSheet.addRow(row);
    if (index === 0) {
      excelRow.getCell(2).font = { bold: true };
      excelRow.getCell(3).font = { bold: true };
    }
  });

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
