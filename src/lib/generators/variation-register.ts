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

const currencyFormat = '"£"#,##0.00';
const dateFormat = 'dd/mm/yyyy';

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    left: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    right: { style: 'thin', color: { argb: 'FFe2e8f0' } },
  };
}

export async function generateVariationRegister(): Promise<Buffer> {
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
    ['', 'VARIATION REQUEST REGISTER'],
    ['', ''],
    ['', 'This register tracks all changes to the original contract scope, including additions,'],
    ['', 'omissions, and modifications to the agreed works.'],
    ['', ''],
    ['', 'WHY TRACK VARIATIONS?'],
    ['', ''],
    ['', '- Maintain accurate project costs'],
    ['', '- Provide audit trail for disputes'],
    ['', '- Manage scope creep'],
    ['', '- Support final account reconciliation'],
    ['', '- Document agreed changes'],
    ['', ''],
    ['', 'HOW TO USE:'],
    ['', ''],
    ['', '1. Log every change request, no matter how small'],
    ['', '2. Get written approval before work proceeds'],
    ['', '3. Track cost impact (additions AND omissions)'],
    ['', '4. Update status as variations progress'],
    ['', '5. Review running totals regularly'],
    ['', ''],
    ['', 'VARIATION TYPES:'],
    ['', ''],
    ['', 'ADDITION: New work not in original scope'],
    ['', 'OMISSION: Work removed from original scope (negative cost)'],
    ['', 'MODIFICATION: Change to specified work'],
    ['', 'PROVISIONAL SUM: Adjustment to PS allowance'],
    ['', ''],
    ['', 'APPROVAL PROCESS:'],
    ['', ''],
    ['', '1. Builder submits variation request with cost'],
    ['', '2. Client reviews and queries if needed'],
    ['', '3. Client approves/rejects in writing'],
    ['', '4. Work proceeds only after approval'],
    ['', '5. Keep copies of all approvals'],
    ['', ''],
    ['', 'IMPORTANT: Never approve verbal variations. Always get it in writing!'],
  ];

  instructions.forEach((row, index) => {
    const excelRow = instructionsSheet.addRow(row);
    if (index === 1) {
      excelRow.getCell(2).font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };
    } else if (row[1].startsWith('WHY') || row[1].startsWith('HOW') ||
               row[1].startsWith('VARIATION TYPES') || row[1].startsWith('APPROVAL')) {
      excelRow.getCell(2).font = { bold: true, size: 12, color: { argb: COLORS.slateHeader } };
    } else if (row[1].startsWith('IMPORTANT')) {
      excelRow.getCell(2).font = { bold: true, color: { argb: COLORS.red } };
    } else if (row[1].match(/^[A-Z]+:/)) {
      excelRow.getCell(2).font = { bold: true };
    }
  });

  // ===== SUMMARY SHEET =====
  const summarySheet = workbook.addWorksheet('Summary', {
    properties: { tabColor: { argb: COLORS.green } },
  });

  summarySheet.columns = [
    { width: 5 },
    { width: 30 },
    { width: 18 },
    { width: 5 },
    { width: 30 },
    { width: 18 },
  ];

  // Title
  summarySheet.mergeCells('B2:F2');
  const sumTitleCell = summarySheet.getCell('B2');
  sumTitleCell.value = 'VARIATION REGISTER - SUMMARY';
  sumTitleCell.font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };

  // Project details
  summarySheet.getCell('B4').value = 'Project:';
  summarySheet.getCell('B4').font = { bold: true };
  const projectCell = summarySheet.getCell('C4');
  Object.assign(projectCell, { style: inputStyle });
  projectCell.alignment = { horizontal: 'left', vertical: 'middle' };

  summarySheet.getCell('E4').value = 'Last Updated:';
  summarySheet.getCell('E4').font = { bold: true };
  summarySheet.getCell('F4').value = { formula: 'TODAY()' };
  summarySheet.getCell('F4').numFmt = dateFormat;

  summarySheet.getCell('B5').value = 'Original Contract Sum:';
  summarySheet.getCell('B5').font = { bold: true };
  const contractCell = summarySheet.getCell('C5');
  Object.assign(contractCell, { style: inputStyle });
  contractCell.numFmt = currencyFormat;

  // Summary statistics
  summarySheet.getCell('B8').value = 'VARIATION SUMMARY';
  summarySheet.getCell('B8').font = { bold: true, size: 14 };

  const summaryData = [
    { label: 'Total Variations Logged:', formula: 'COUNTA(Register!B8:B107)' },
    { label: 'Approved Variations:', formula: 'COUNTIF(Register!J8:J107,"Approved")' },
    { label: 'Pending Approval:', formula: 'COUNTIF(Register!J8:J107,"Pending")' },
    { label: 'Rejected:', formula: 'COUNTIF(Register!J8:J107,"Rejected")' },
  ];

  summaryData.forEach((item, index) => {
    const rowNum = 10 + index;
    summarySheet.getCell(`B${rowNum}`).value = item.label;
    summarySheet.getCell(`B${rowNum}`).font = { bold: true };
    const valueCell = summarySheet.getCell(`C${rowNum}`);
    valueCell.value = { formula: item.formula };
    applyBorder(valueCell);
  });

  // Cost summary
  summarySheet.getCell('B16').value = 'COST IMPACT';
  summarySheet.getCell('B16').font = { bold: true, size: 14 };

  const costData = [
    { label: 'Total Additions (Approved):', formula: 'SUMIFS(Register!G8:G107,Register!J8:J107,"Approved",Register!E8:E107,"Addition")+SUMIFS(Register!G8:G107,Register!J8:J107,"Approved",Register!E8:E107,"Modification")' },
    { label: 'Total Omissions (Approved):', formula: 'SUMIFS(Register!G8:G107,Register!J8:J107,"Approved",Register!E8:E107,"Omission")' },
    { label: 'Net Approved Variations:', formula: 'SUMIF(Register!J8:J107,"Approved",Register!G8:G107)' },
    { label: 'Pending Variations:', formula: 'SUMIF(Register!J8:J107,"Pending",Register!G8:G107)' },
  ];

  costData.forEach((item, index) => {
    const rowNum = 18 + index;
    summarySheet.getCell(`B${rowNum}`).value = item.label;
    summarySheet.getCell(`B${rowNum}`).font = { bold: true };
    const valueCell = summarySheet.getCell(`C${rowNum}`);
    valueCell.value = { formula: item.formula };
    valueCell.numFmt = currencyFormat;
    applyBorder(valueCell);

    if (item.label.includes('Omissions')) {
      valueCell.font = { color: { argb: COLORS.green } };
    } else if (item.label.includes('Additions')) {
      valueCell.font = { color: { argb: COLORS.red } };
    }
  });

  // Current contract value
  summarySheet.getCell('B24').value = 'CURRENT CONTRACT POSITION';
  summarySheet.getCell('B24').font = { bold: true, size: 14 };

  summarySheet.getCell('B26').value = 'Original Contract Sum:';
  summarySheet.getCell('C26').value = { formula: 'C5' };
  summarySheet.getCell('C26').numFmt = currencyFormat;
  applyBorder(summarySheet.getCell('C26'));

  summarySheet.getCell('B27').value = 'Approved Variations:';
  summarySheet.getCell('C27').value = { formula: 'C20' };
  summarySheet.getCell('C27').numFmt = currencyFormat;
  applyBorder(summarySheet.getCell('C27'));

  summarySheet.getCell('B28').value = 'REVISED CONTRACT SUM:';
  summarySheet.getCell('B28').font = { bold: true, size: 12 };
  const revisedCell = summarySheet.getCell('C28');
  revisedCell.value = { formula: 'C26+C27' };
  revisedCell.numFmt = currencyFormat;
  revisedCell.font = { bold: true, size: 14, color: { argb: COLORS.primaryBlue } };
  revisedCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  applyBorder(revisedCell);

  summarySheet.getCell('B30').value = 'Potential Additional (Pending):';
  summarySheet.getCell('C30').value = { formula: 'C21' };
  summarySheet.getCell('C30').numFmt = currencyFormat;
  summarySheet.getCell('C30').font = { italic: true, color: { argb: 'FF64748b' } };
  applyBorder(summarySheet.getCell('C30'));

  // By type breakdown
  summarySheet.getCell('E8').value = 'BY TYPE';
  summarySheet.getCell('E8').font = { bold: true, size: 14 };

  const types = ['Addition', 'Omission', 'Modification', 'Provisional Sum'];
  types.forEach((type, index) => {
    const rowNum = 10 + index;
    summarySheet.getCell(`E${rowNum}`).value = type;
    const countCell = summarySheet.getCell(`F${rowNum}`);
    countCell.value = { formula: `COUNTIF(Register!E8:E107,"${type}")` };
    applyBorder(countCell);
  });

  // By category breakdown
  summarySheet.getCell('E16').value = 'BY CATEGORY';
  summarySheet.getCell('E16').font = { bold: true, size: 14 };

  const categories = ['Structure', 'Finishes', 'M&E', 'External', 'Kitchen', 'Bathroom', 'Other'];
  categories.forEach((cat, index) => {
    const rowNum = 18 + index;
    summarySheet.getCell(`E${rowNum}`).value = cat;
    const valueCell = summarySheet.getCell(`F${rowNum}`);
    valueCell.value = { formula: `SUMIF(Register!F8:F107,"${cat}",Register!G8:G107)` };
    valueCell.numFmt = currencyFormat;
    applyBorder(valueCell);
  });

  // Warning if variations exceed threshold
  summarySheet.getCell('B33').value = 'VARIATION ALERT:';
  summarySheet.getCell('B33').font = { bold: true, size: 12, color: { argb: COLORS.red } };
  const alertCell = summarySheet.getCell('C33');
  alertCell.value = { formula: 'IF(C5=0,"Enter contract sum",IF(C27/C5>0.15,"WARNING: Variations exceed 15% of contract",IF(C27/C5>0.1,"CAUTION: Variations exceed 10% of contract","Variations within normal range")))' };
  alertCell.font = { bold: true };
  summarySheet.mergeCells('C33:F33');

  summarySheet.addConditionalFormatting({
    ref: 'C33',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'WARNING',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'CAUTION',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
          font: { color: { argb: COLORS.yellow }, bold: true },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'normal',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 3,
      },
    ],
  });

  // ===== REGISTER SHEET =====
  const registerSheet = workbook.addWorksheet('Register', {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  registerSheet.columns = [
    { width: 5 },
    { width: 10 },
    { width: 12 },
    { width: 35 },
    { width: 14 },
    { width: 12 },
    { width: 14 },
    { width: 10 },
    { width: 12 },
    { width: 12 },
    { width: 20 },
    { width: 12 },
  ];

  // Title
  registerSheet.mergeCells('B2:L2');
  const regTitleCell = registerSheet.getCell('B2');
  regTitleCell.value = 'VARIATION REQUEST REGISTER';
  regTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.orange } };

  // Running total display
  registerSheet.getCell('B4').value = 'Running Total (Approved):';
  registerSheet.getCell('B4').font = { bold: true };
  const runningTotalCell = registerSheet.getCell('D4');
  runningTotalCell.value = { formula: 'SUMIF(J8:J107,"Approved",G8:G107)' };
  runningTotalCell.numFmt = currencyFormat;
  runningTotalCell.font = { bold: true, size: 14, color: { argb: COLORS.primaryBlue } };

  registerSheet.getCell('F4').value = 'Pending:';
  registerSheet.getCell('F4').font = { bold: true };
  const pendingTotalCell = registerSheet.getCell('G4');
  pendingTotalCell.value = { formula: 'SUMIF(J8:J107,"Pending",G8:G107)' };
  pendingTotalCell.numFmt = currencyFormat;
  pendingTotalCell.font = { italic: true, color: { argb: 'FF64748b' } };

  // Headers
  const headers = ['', 'Ref', 'Date', 'Description', 'Type', 'Category', 'Cost Impact', 'Days', 'Requested By', 'Status', 'Approved By', 'Approval Date'];
  const headerRow = registerSheet.getRow(7);
  headers.forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Enable auto-filter
  registerSheet.autoFilter = {
    from: { row: 7, column: 2 },
    to: { row: 107, column: 12 },
  };

  // Data rows
  for (let i = 8; i <= 107; i++) {
    const row = registerSheet.getRow(i);

    // Reference number
    const refCell = row.getCell(2);
    refCell.value = { formula: `IF(C${i}="","","VR-"&TEXT(ROW()-7,"000"))` };
    refCell.alignment = { horizontal: 'center' };
    applyBorder(refCell);

    // Date
    const dateCell = row.getCell(3);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    // Description
    const descCell = row.getCell(4);
    Object.assign(descCell, { style: inputStyle });
    descCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // Type dropdown
    const typeCell = row.getCell(5);
    typeCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Addition,Omission,Modification,Provisional Sum"'],
    };
    Object.assign(typeCell, { style: inputStyle });

    // Category dropdown
    const catCell = row.getCell(6);
    catCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Structure,Finishes,M&E,External,Kitchen,Bathroom,Other"'],
    };
    Object.assign(catCell, { style: inputStyle });

    // Cost impact
    const costCell = row.getCell(7);
    Object.assign(costCell, { style: inputStyle });
    costCell.numFmt = currencyFormat;

    // Days impact
    const daysCell = row.getCell(8);
    Object.assign(daysCell, { style: inputStyle });
    daysCell.numFmt = '0';

    // Requested by
    const reqByCell = row.getCell(9);
    reqByCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Builder,Client,Architect,SE,Other"'],
    };
    Object.assign(reqByCell, { style: inputStyle });

    // Status dropdown
    const statusCell = row.getCell(10);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Pending,Approved,Rejected,On Hold"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    // Approved by
    const appByCell = row.getCell(11);
    Object.assign(appByCell, { style: inputStyle });
    appByCell.alignment = { horizontal: 'left', vertical: 'middle' };

    // Approval date
    const appDateCell = row.getCell(12);
    Object.assign(appDateCell, { style: inputStyle });
    appDateCell.numFmt = dateFormat;
  }

  // Conditional formatting for type
  registerSheet.addConditionalFormatting({
    ref: 'E8:E107',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Omission',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Addition',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
        },
        priority: 2,
      },
    ],
  });

  // Conditional formatting for status
  registerSheet.addConditionalFormatting({
    ref: 'J8:J107',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Approved',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Rejected',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red } },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Pending',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
        },
        priority: 3,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'On Hold',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } },
        },
        priority: 4,
      },
    ],
  });

  // Conditional formatting for cost - negative values in green
  registerSheet.addConditionalFormatting({
    ref: 'G8:G107',
    rules: [
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: { font: { color: { argb: COLORS.green } } },
        priority: 1,
      },
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['1000'],
        style: { font: { color: { argb: COLORS.red }, bold: true } },
        priority: 2,
      },
    ],
  });

  // ===== VARIATION DETAIL TEMPLATE SHEET =====
  const detailSheet = workbook.addWorksheet('Variation Detail Template', {
    properties: { tabColor: { argb: 'FF94a3b8' } },
  });

  detailSheet.columns = [
    { width: 5 },
    { width: 25 },
    { width: 35 },
    { width: 20 },
  ];

  // Title
  detailSheet.mergeCells('B2:D2');
  const detTitleCell = detailSheet.getCell('B2');
  detTitleCell.value = 'VARIATION REQUEST FORM';
  detTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  detailSheet.getCell('B4').value = '(Duplicate this sheet for each variation requiring detailed documentation)';
  detailSheet.getCell('B4').font = { italic: true, color: { argb: 'FF64748b' } };

  // Form fields
  const formFields = [
    { label: 'Variation Reference:', cell: 'C6' },
    { label: 'Date Raised:', cell: 'C7', format: dateFormat },
    { label: 'Raised By:', cell: 'C8' },
    { label: '', cell: '' },
    { label: 'DESCRIPTION OF CHANGE', cell: '', isHeader: true },
    { label: 'Original Specification:', cell: 'C11', merge: 'C11:D13' },
    { label: '', cell: '' },
    { label: '', cell: '' },
    { label: 'Proposed Change:', cell: 'C15', merge: 'C15:D17' },
    { label: '', cell: '' },
    { label: '', cell: '' },
    { label: '', cell: '' },
    { label: 'REASON FOR CHANGE', cell: '', isHeader: true },
    { label: 'Reason:', cell: 'C20', merge: 'C20:D21' },
    { label: '', cell: '' },
    { label: '', cell: '' },
    { label: 'COST BREAKDOWN', cell: '', isHeader: true },
    { label: 'Labour:', cell: 'C24', format: currencyFormat },
    { label: 'Materials:', cell: 'C25', format: currencyFormat },
    { label: 'Plant/Equipment:', cell: 'C26', format: currencyFormat },
    { label: 'Preliminaries:', cell: 'C27', format: currencyFormat },
    { label: 'Margin:', cell: 'C28', format: currencyFormat },
    { label: 'TOTAL:', cell: 'C29', format: currencyFormat, formula: 'SUM(C24:C28)' },
    { label: '', cell: '' },
    { label: 'PROGRAMME IMPACT', cell: '', isHeader: true },
    { label: 'Days Impact:', cell: 'C32' },
    { label: 'Reason:', cell: 'C33', merge: 'C33:D34' },
    { label: '', cell: '' },
    { label: '', cell: '' },
    { label: 'APPROVAL', cell: '', isHeader: true },
    { label: 'Status:', cell: 'C37', dropdown: '"Pending,Approved,Rejected"' },
    { label: 'Approved By:', cell: 'C38' },
    { label: 'Date:', cell: 'C39', format: dateFormat },
    { label: 'Comments:', cell: 'C40', merge: 'C40:D42' },
  ];

  let currentRow = 6;
  formFields.forEach((field) => {
    if (field.isHeader) {
      detailSheet.getCell(`B${currentRow}`).value = field.label;
      detailSheet.getCell(`B${currentRow}`).font = { bold: true, size: 12, color: { argb: COLORS.slateHeader } };
      currentRow++;
      return;
    }

    if (field.label === '') {
      currentRow++;
      return;
    }

    detailSheet.getCell(`B${currentRow}`).value = field.label;
    detailSheet.getCell(`B${currentRow}`).font = { bold: true };

    if (field.cell) {
      const inputCell = detailSheet.getCell(field.cell);

      if (field.merge) {
        detailSheet.mergeCells(field.merge);
        Object.assign(inputCell, { style: inputStyle });
        inputCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      } else {
        Object.assign(inputCell, { style: inputStyle });
        inputCell.alignment = { horizontal: 'left', vertical: 'middle' };
      }

      if (field.format) {
        inputCell.numFmt = field.format;
      }

      if (field.formula) {
        inputCell.value = { formula: field.formula };
        inputCell.font = { bold: true };
        inputCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
      }

      if (field.dropdown) {
        inputCell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: [field.dropdown],
        };
      }
    }

    currentRow++;
  });

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
