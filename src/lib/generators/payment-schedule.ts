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

const warningStyle: Partial<ExcelJS.Style> = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
  font: { color: { argb: COLORS.red }, bold: true },
  border: {
    top: { style: 'thin', color: { argb: COLORS.red } },
    left: { style: 'thin', color: { argb: COLORS.red } },
    bottom: { style: 'thin', color: { argb: COLORS.red } },
    right: { style: 'thin', color: { argb: COLORS.red } },
  },
};

const currencyFormat = '"£"#,##0.00';
const percentFormat = '0%';
const dateFormat = 'dd/mm/yyyy';

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    left: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    right: { style: 'thin', color: { argb: 'FFe2e8f0' } },
  };
}

export async function generatePaymentSchedule(): Promise<Buffer> {
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
    ['', 'PAYMENT SCHEDULE TEMPLATE'],
    ['', ''],
    ['', 'This template helps you manage payments to your builder based on completed work stages.'],
    ['', 'It includes built-in protection warnings to avoid paying ahead of work completed.'],
    ['', ''],
    ['', 'KEY PRINCIPLES:'],
    ['', ''],
    ['', '1. NEVER pay more than 10-15% upfront (materials deposit maximum)'],
    ['', '2. Always tie payments to completed and inspected work stages'],
    ['', '3. Retain 2.5-5% until ALL snagging is complete'],
    ['', '4. Get receipts and keep records of all payments'],
    ['', '5. If using staged payments, verify completion before paying'],
    ['', ''],
    ['', 'HOW TO USE:'],
    ['', ''],
    ['', '1. CONTRACT SUMMARY: Enter your contract details'],
    ['', '2. PAYMENT SCHEDULE: Set up milestones and payment amounts'],
    ['', '3. PAYMENT LOG: Record actual payments made'],
    ['', '4. WARNINGS: Review any payment protection alerts'],
    ['', ''],
    ['', 'RETENTION EXPLAINED:'],
    ['', ''],
    ['', 'Retention is a percentage held back from each payment until the project is complete.'],
    ['', 'Typically 2.5-5% of the contract value, released after:'],
    ['', '- All work is finished to your satisfaction'],
    ['', '- Snagging items have been addressed'],
    ['', '- Building control sign-off obtained'],
    ['', '- All warranties and certificates provided'],
    ['', ''],
    ['', 'WARNING INDICATORS:'],
    ['', ''],
    ['', 'RED: Payments exceed work completion - STOP and verify'],
    ['', 'AMBER: Payments approaching work completion - monitor closely'],
    ['', 'GREEN: Payments in line with or behind work completion - healthy'],
  ];

  instructions.forEach((row, index) => {
    const excelRow = instructionsSheet.addRow(row);
    if (index === 1) {
      excelRow.getCell(2).font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };
    } else if (row[1].startsWith('KEY PRINCIPLES') || row[1].startsWith('HOW TO USE') ||
               row[1].startsWith('RETENTION') || row[1].startsWith('WARNING')) {
      excelRow.getCell(2).font = { bold: true, size: 12, color: { argb: COLORS.slateHeader } };
    } else if (row[1].match(/^\d\./)) {
      excelRow.getCell(2).font = { bold: true };
    } else if (row[1].startsWith('RED:')) {
      excelRow.getCell(2).font = { color: { argb: COLORS.red } };
    } else if (row[1].startsWith('AMBER:')) {
      excelRow.getCell(2).font = { color: { argb: COLORS.yellow } };
    } else if (row[1].startsWith('GREEN:')) {
      excelRow.getCell(2).font = { color: { argb: COLORS.green } };
    }
  });

  // ===== CONTRACT SUMMARY SHEET =====
  const contractSheet = workbook.addWorksheet('Contract Summary', {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  contractSheet.columns = [
    { width: 5 },
    { width: 30 },
    { width: 25 },
    { width: 5 },
    { width: 30 },
    { width: 25 },
  ];

  // Title
  contractSheet.mergeCells('B2:F2');
  const titleCell = contractSheet.getCell('B2');
  titleCell.value = 'CONTRACT SUMMARY';
  titleCell.font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  // Contract details
  const contractData = [
    ['', '', '', '', '', ''],
    ['', 'Builder/Contractor:', '', '', 'Contract Date:', ''],
    ['', 'Contact Name:', '', '', 'Start Date:', ''],
    ['', 'Phone:', '', '', 'Expected End:', ''],
    ['', 'Email:', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['', 'CONTRACT VALUE', '', '', '', ''],
    ['', 'Original Contract Sum:', '', '', '', ''],
    ['', 'Agreed Variations (+/-):', '', '', '', ''],
    ['', 'Current Contract Value:', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['', 'RETENTION DETAILS', '', '', '', ''],
    ['', 'Retention Percentage:', '', '', '', ''],
    ['', 'Retention Amount:', '', '', '', ''],
    ['', 'Retention Release Terms:', '', '', '', ''],
  ];

  contractData.forEach((row, index) => {
    const excelRow = contractSheet.addRow(row);
    const rowNum = index + 4;

    // Style section headers
    if (row[1] === 'CONTRACT VALUE' || row[1] === 'RETENTION DETAILS') {
      excelRow.getCell(2).font = { bold: true, size: 12, color: { argb: COLORS.slateHeader } };
    }

    // Input cells - left column
    if ([4, 5, 6, 7, 11, 12, 16].includes(rowNum)) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      if (rowNum === 11 || rowNum === 12) {
        inputCell.numFmt = currencyFormat;
      }
    }

    // Input cells - right column
    if ([4, 5, 6].includes(rowNum - 4 + 4)) {
      if (row[4]) {
        const inputCell = excelRow.getCell(6);
        Object.assign(inputCell, { style: inputStyle });
        inputCell.numFmt = dateFormat;
      }
    }

    // Calculated cells
    if (rowNum === 13) {
      const calcCell = excelRow.getCell(3);
      calcCell.value = { formula: 'C11+C12' };
      calcCell.numFmt = currencyFormat;
      calcCell.font = { bold: true, color: { argb: COLORS.primaryBlue } };
      calcCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
    }

    if (rowNum === 16) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.numFmt = percentFormat;
      inputCell.value = 0.05; // Default 5%
    }

    if (rowNum === 17) {
      const calcCell = excelRow.getCell(3);
      calcCell.value = { formula: 'C13*C16' };
      calcCell.numFmt = currencyFormat;
      calcCell.font = { color: { argb: COLORS.primaryBlue } };
    }

    if (rowNum === 18) {
      contractSheet.mergeCells(`C${rowNum}:E${rowNum}`);
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.alignment = { horizontal: 'left', vertical: 'middle' };
    }
  });

  // Payment summary
  contractSheet.getCell('B21').value = 'PAYMENT SUMMARY';
  contractSheet.getCell('B21').font = { bold: true, size: 12, color: { argb: COLORS.slateHeader } };

  const summaryLabels = [
    ['Total Scheduled:', 'SUM(\'Payment Schedule\'!E7:E20)'],
    ['Total Paid:', 'SUM(\'Payment Log\'!D7:D50)'],
    ['Remaining:', 'C13-C24'],
    ['Retention Held:', 'C17'],
  ];

  summaryLabels.forEach((item, index) => {
    const rowNum = 23 + index;
    contractSheet.getCell(`B${rowNum}`).value = item[0];
    contractSheet.getCell(`B${rowNum}`).font = { bold: true };
    const valueCell = contractSheet.getCell(`C${rowNum}`);
    valueCell.value = { formula: item[1] };
    valueCell.numFmt = currencyFormat;
    applyBorder(valueCell);
    if (index === 2) {
      valueCell.font = { bold: true, color: { argb: COLORS.primaryBlue } };
    }
  });

  // ===== PAYMENT SCHEDULE SHEET =====
  const scheduleSheet = workbook.addWorksheet('Payment Schedule', {
    properties: { tabColor: { argb: COLORS.green } },
  });

  scheduleSheet.columns = [
    { width: 5 },
    { width: 8 },
    { width: 30 },
    { width: 12 },
    { width: 15 },
    { width: 12 },
    { width: 15 },
    { width: 12 },
    { width: 20 },
  ];

  // Title
  scheduleSheet.mergeCells('B2:I2');
  const schedTitleCell = scheduleSheet.getCell('B2');
  schedTitleCell.value = 'PAYMENT SCHEDULE';
  schedTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.green } };

  scheduleSheet.getCell('B4').value = 'Define payment milestones and track work completion vs payments';
  scheduleSheet.getCell('B4').font = { italic: true, color: { argb: 'FF64748b' } };

  // Headers
  const schedHeaders = ['', 'Stage', 'Milestone Description', '% of Contract', 'Amount', 'Work %', 'Paid %', 'Status', 'Notes'];
  const schedHeaderRow = scheduleSheet.getRow(6);
  schedHeaders.forEach((header, index) => {
    const cell = schedHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Default milestones
  const milestones = [
    { stage: 1, desc: 'Contract Signing / Deposit', pct: 0.10 },
    { stage: 2, desc: 'Foundations Complete', pct: 0.15 },
    { stage: 3, desc: 'Walls to Plate Level', pct: 0.15 },
    { stage: 4, desc: 'Roof Watertight', pct: 0.15 },
    { stage: 5, desc: 'First Fix Complete', pct: 0.15 },
    { stage: 6, desc: 'Second Fix Complete', pct: 0.15 },
    { stage: 7, desc: 'Practical Completion', pct: 0.10 },
    { stage: 8, desc: 'Retention Release', pct: 0.05 },
  ];

  milestones.forEach((milestone, index) => {
    const rowNum = index + 7;
    const row = scheduleSheet.getRow(rowNum);

    row.getCell(2).value = milestone.stage;
    row.getCell(2).alignment = { horizontal: 'center' };
    row.getCell(2).numFmt = '0';
    applyBorder(row.getCell(2));

    const descCell = row.getCell(3);
    descCell.value = milestone.desc;
    Object.assign(descCell, { style: inputStyle });
    descCell.alignment = { horizontal: 'left' };

    const pctCell = row.getCell(4);
    pctCell.value = milestone.pct;
    pctCell.numFmt = percentFormat;
    Object.assign(pctCell, { style: inputStyle });

    const amtCell = row.getCell(5);
    amtCell.value = { formula: `D${rowNum}*'Contract Summary'!C13` };
    amtCell.numFmt = currencyFormat;
    applyBorder(amtCell);

    // Work completion %
    const workPctCell = row.getCell(6);
    workPctCell.dataValidation = {
      type: 'decimal',
      operator: 'between',
      allowBlank: true,
      formulae: [0, 1],
    };
    Object.assign(workPctCell, { style: inputStyle });
    workPctCell.numFmt = percentFormat;

    // Paid %
    const paidPctCell = row.getCell(7);
    paidPctCell.value = { formula: `IF('Contract Summary'!C13=0,0,SUMIF('Payment Log'!C7:C50,B${rowNum},'Payment Log'!D7:D50)/E${rowNum})` };
    paidPctCell.numFmt = percentFormat;
    applyBorder(paidPctCell);

    // Status
    const statusCell = row.getCell(8);
    statusCell.value = { formula: `IF(F${rowNum}="","Not Started",IF(G${rowNum}>F${rowNum}+0.05,"OVERPAID",IF(G${rowNum}>=F${rowNum},"OK","Behind")))` };
    applyBorder(statusCell);

    // Notes
    const notesCell = row.getCell(9);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left' };
  });

  // Add more empty rows for additional milestones
  for (let i = milestones.length + 7; i <= 20; i++) {
    const row = scheduleSheet.getRow(i);
    row.getCell(2).value = i - 6;
    row.getCell(2).alignment = { horizontal: 'center' };
    row.getCell(2).numFmt = '0';
    applyBorder(row.getCell(2));

    Object.assign(row.getCell(3), { style: inputStyle });
    row.getCell(3).alignment = { horizontal: 'left' };

    const pctCell = row.getCell(4);
    Object.assign(pctCell, { style: inputStyle });
    pctCell.numFmt = percentFormat;

    const amtCell = row.getCell(5);
    amtCell.value = { formula: `D${i}*'Contract Summary'!C13` };
    amtCell.numFmt = currencyFormat;
    applyBorder(amtCell);

    const workPctCell = row.getCell(6);
    workPctCell.dataValidation = {
      type: 'decimal',
      operator: 'between',
      allowBlank: true,
      formulae: [0, 1],
    };
    Object.assign(workPctCell, { style: inputStyle });
    workPctCell.numFmt = percentFormat;

    const paidPctCell = row.getCell(7);
    paidPctCell.value = { formula: `IF(E${i}=0,0,SUMIF('Payment Log'!C7:C50,B${i},'Payment Log'!D7:D50)/E${i})` };
    paidPctCell.numFmt = percentFormat;
    applyBorder(paidPctCell);

    const statusCell = row.getCell(8);
    statusCell.value = { formula: `IF(F${i}="","",IF(G${i}>F${i}+0.05,"OVERPAID",IF(G${i}>=F${i},"OK","Behind")))` };
    applyBorder(statusCell);

    Object.assign(row.getCell(9), { style: inputStyle });
    row.getCell(9).alignment = { horizontal: 'left' };
  }

  // Totals
  const totalRowNum = 21;
  const totalRow = scheduleSheet.getRow(totalRowNum);
  totalRow.getCell(3).value = 'TOTALS';
  totalRow.getCell(3).font = { bold: true };
  applyBorder(totalRow.getCell(3));

  const totalPctCell = totalRow.getCell(4);
  totalPctCell.value = { formula: 'SUM(D7:D20)' };
  totalPctCell.numFmt = percentFormat;
  totalPctCell.font = { bold: true };
  totalPctCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  applyBorder(totalPctCell);

  const totalAmtCell = totalRow.getCell(5);
  totalAmtCell.value = { formula: 'SUM(E7:E20)' };
  totalAmtCell.numFmt = currencyFormat;
  totalAmtCell.font = { bold: true };
  totalAmtCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  applyBorder(totalAmtCell);

  // Conditional formatting for status
  scheduleSheet.addConditionalFormatting({
    ref: 'H7:H20',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'OVERPAID',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'OK',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Behind',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
          font: { color: { argb: COLORS.yellow } },
        },
        priority: 3,
      },
    ],
  });

  // Warning if total % doesn't equal 100%
  scheduleSheet.getCell('D23').value = { formula: 'IF(D21<>1,"Warning: Schedule does not total 100%","")' };
  scheduleSheet.getCell('D23').font = { color: { argb: COLORS.red }, bold: true };

  // ===== PAYMENT LOG SHEET =====
  const logSheet = workbook.addWorksheet('Payment Log', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  logSheet.columns = [
    { width: 5 },
    { width: 12 },
    { width: 10 },
    { width: 15 },
    { width: 20 },
    { width: 18 },
    { width: 25 },
  ];

  // Title
  logSheet.mergeCells('B2:G2');
  const logTitleCell = logSheet.getCell('B2');
  logTitleCell.value = 'PAYMENT LOG';
  logTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  logSheet.getCell('B4').value = 'Record all payments made. Keep receipts for all transactions.';
  logSheet.getCell('B4').font = { italic: true, color: { argb: 'FF64748b' } };

  // Headers
  const logHeaders = ['', 'Date', 'Stage', 'Amount', 'Payment Method', 'Reference', 'Notes'];
  const logHeaderRow = logSheet.getRow(6);
  logHeaders.forEach((header, index) => {
    const cell = logHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Payment rows
  for (let i = 7; i <= 50; i++) {
    const row = logSheet.getRow(i);

    const dateCell = row.getCell(2);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    const stageCell = row.getCell(3);
    stageCell.dataValidation = {
      type: 'whole',
      operator: 'between',
      allowBlank: true,
      formulae: [1, 14],
    };
    Object.assign(stageCell, { style: inputStyle });
    stageCell.numFmt = '0';

    const amtCell = row.getCell(4);
    Object.assign(amtCell, { style: inputStyle });
    amtCell.numFmt = currencyFormat;

    const methodCell = row.getCell(5);
    methodCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Bank Transfer,Cheque,Card,Cash,Other"'],
    };
    Object.assign(methodCell, { style: inputStyle });

    const refCell = row.getCell(6);
    Object.assign(refCell, { style: inputStyle });
    refCell.alignment = { horizontal: 'left' };

    const notesCell = row.getCell(7);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left' };
  }

  // Total
  const logTotalRow = logSheet.getRow(51);
  logTotalRow.getCell(3).value = 'TOTAL PAID';
  logTotalRow.getCell(3).font = { bold: true };
  applyBorder(logTotalRow.getCell(3));

  const logTotalCell = logTotalRow.getCell(4);
  logTotalCell.value = { formula: 'SUM(D7:D50)' };
  logTotalCell.numFmt = currencyFormat;
  logTotalCell.font = { bold: true };
  logTotalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  applyBorder(logTotalCell);

  // ===== WARNINGS SHEET =====
  const warningsSheet = workbook.addWorksheet('Warnings', {
    properties: { tabColor: { argb: COLORS.red } },
  });

  warningsSheet.columns = [
    { width: 5 },
    { width: 25 },
    { width: 50 },
    { width: 15 },
  ];

  // Title
  warningsSheet.mergeCells('B2:D2');
  const warnTitleCell = warningsSheet.getCell('B2');
  warnTitleCell.value = 'PAYMENT PROTECTION WARNINGS';
  warnTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.red } };

  warningsSheet.getCell('B4').value = 'This sheet automatically flags potential payment issues';
  warningsSheet.getCell('B4').font = { italic: true, color: { argb: 'FF64748b' } };

  // Warning checks
  const warnings = [
    {
      check: 'Upfront Payment > 15%',
      formula: 'IF(\'Payment Schedule\'!D7>0.15,"WARNING: Initial payment exceeds 15% - this is higher than recommended","OK")',
    },
    {
      check: 'Total Paid > Work Complete',
      formula: 'IF(SUM(\'Payment Log\'!D7:D50)/\'Contract Summary\'!C13>AVERAGE(\'Payment Schedule\'!F7:F20),"WARNING: Total payments exceed average work completion","OK")',
    },
    {
      check: 'Schedule Totals 100%',
      formula: 'IF(SUM(\'Payment Schedule\'!D7:D20)<>1,"WARNING: Payment schedule does not total 100%","OK")',
    },
    {
      check: 'Retention Included',
      formula: 'IF(\'Contract Summary\'!C16<0.025,"WARNING: Retention less than 2.5% - consider increasing","OK")',
    },
    {
      check: 'Payment Method Safety',
      formula: 'IF(COUNTIF(\'Payment Log\'!E7:E50,"Cash")>2,"WARNING: Multiple cash payments - ensure you have receipts","OK")',
    },
  ];

  // Headers
  const warnHeaders = ['', 'Check', 'Status', 'Action'];
  const warnHeaderRow = warningsSheet.getRow(6);
  warnHeaders.forEach((header, index) => {
    const cell = warnHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  warnings.forEach((warning, index) => {
    const rowNum = index + 7;
    const row = warningsSheet.getRow(rowNum);

    row.getCell(2).value = warning.check;
    applyBorder(row.getCell(2));

    const statusCell = row.getCell(3);
    statusCell.value = { formula: warning.formula };
    applyBorder(statusCell);

    const actionCell = row.getCell(4);
    actionCell.value = { formula: `IF(LEFT(C${rowNum},7)="WARNING","Review","None")` };
    applyBorder(actionCell);
  });

  // Conditional formatting for warnings
  warningsSheet.addConditionalFormatting({
    ref: 'C7:C11',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'WARNING',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'OK',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 2,
      },
    ],
  });

  // Summary warning
  warningsSheet.getCell('B14').value = 'OVERALL STATUS';
  warningsSheet.getCell('B14').font = { bold: true, size: 12 };

  const overallCell = warningsSheet.getCell('C14');
  overallCell.value = { formula: 'IF(COUNTIF(C7:C11,"WARNING*")>0,"ACTION REQUIRED - Review warnings above","ALL CLEAR")' };
  overallCell.font = { bold: true, size: 14 };

  warningsSheet.addConditionalFormatting({
    ref: 'C14',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'ACTION',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'CLEAR',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green }, bold: true },
        },
        priority: 2,
      },
    ],
  });

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
