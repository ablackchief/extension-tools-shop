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

const subHeaderStyle: Partial<ExcelJS.Style> = {
  font: { bold: true, color: { argb: COLORS.slateHeader }, size: 11 },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } },
  alignment: { horizontal: 'left', vertical: 'middle' },
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
  alignment: { horizontal: 'right', vertical: 'middle' },
};

const currencyFormat = '"£"#,##0.00';
const percentFormat = '0.0%';
const dateFormat = 'dd/mm/yyyy';

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    left: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    right: { style: 'thin', color: { argb: 'FFe2e8f0' } },
  };
}

function createCategorySheet(
  workbook: ExcelJS.Workbook,
  name: string,
  color: string,
  items: Array<{ item: string; estimated: number; notes?: string }>
) {
  const sheet = workbook.addWorksheet(name, {
    properties: { tabColor: { argb: color } },
  });

  sheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 12 },
    { width: 30 },
  ];

  // Title
  sheet.mergeCells('B2:G2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = name.toUpperCase();
  titleCell.font = { bold: true, size: 14, color: { argb: COLORS.primaryBlue } };

  // Headers
  const headerRow = sheet.getRow(4);
  ['', 'Item', 'Estimated', 'Actual', 'Variance', 'Status', 'Notes'].forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Data rows
  items.forEach((item, index) => {
    const rowNum = index + 5;
    const row = sheet.getRow(rowNum);

    row.getCell(2).value = item.item;
    applyBorder(row.getCell(2));

    const estCell = row.getCell(3);
    estCell.value = item.estimated;
    estCell.numFmt = currencyFormat;
    Object.assign(estCell, { style: inputStyle });

    const actCell = row.getCell(4);
    Object.assign(actCell, { style: inputStyle });
    actCell.numFmt = currencyFormat;

    const varCell = row.getCell(5);
    varCell.value = { formula: `D${rowNum}-C${rowNum}` };
    varCell.numFmt = currencyFormat;
    applyBorder(varCell);

    const statusCell = row.getCell(6);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Pending,Paid,Partial"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    const notesCell = row.getCell(7);
    notesCell.value = item.notes || '';
    Object.assign(notesCell, { style: inputStyle });
  });

  // Totals row
  const totalRowNum = items.length + 5;
  const totalRow = sheet.getRow(totalRowNum);
  totalRow.getCell(2).value = 'TOTAL';
  totalRow.getCell(2).font = { bold: true };
  applyBorder(totalRow.getCell(2));

  const totalEstCell = totalRow.getCell(3);
  totalEstCell.value = { formula: `SUM(C5:C${totalRowNum - 1})` };
  totalEstCell.numFmt = currencyFormat;
  totalEstCell.font = { bold: true };
  totalEstCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  applyBorder(totalEstCell);

  const totalActCell = totalRow.getCell(4);
  totalActCell.value = { formula: `SUM(D5:D${totalRowNum - 1})` };
  totalActCell.numFmt = currencyFormat;
  totalActCell.font = { bold: true };
  totalActCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  applyBorder(totalActCell);

  const totalVarCell = totalRow.getCell(5);
  totalVarCell.value = { formula: `SUM(E5:E${totalRowNum - 1})` };
  totalVarCell.numFmt = currencyFormat;
  totalVarCell.font = { bold: true };
  totalVarCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  applyBorder(totalVarCell);

  // Conditional formatting for variance
  sheet.addConditionalFormatting({
    ref: `E5:E${totalRowNum}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0'],
        style: { font: { color: { argb: COLORS.red } } },
        priority: 1,
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: { font: { color: { argb: COLORS.green } } },
        priority: 2,
      },
    ],
  });

  return sheet;
}

export async function generateBudgetPlanner(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Extension Survival Guide';
  workbook.created = new Date();

  // ===== DASHBOARD SHEET =====
  const dashboard = workbook.addWorksheet('Dashboard', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  dashboard.columns = [
    { width: 3 },
    { width: 25 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 15 },
    { width: 3 },
    { width: 25 },
    { width: 18 },
  ];

  // Title
  dashboard.mergeCells('B2:F2');
  const titleCell = dashboard.getCell('B2');
  titleCell.value = 'EXTENSION BUDGET PLANNER';
  titleCell.font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };

  // Project info
  dashboard.getCell('B4').value = 'Project Name:';
  dashboard.getCell('B4').font = { bold: true };
  const projectNameCell = dashboard.getCell('C4');
  Object.assign(projectNameCell, { style: inputStyle });
  dashboard.mergeCells('C4:D4');

  dashboard.getCell('B5').value = 'Start Date:';
  dashboard.getCell('B5').font = { bold: true };
  const startDateCell = dashboard.getCell('C5');
  Object.assign(startDateCell, { style: inputStyle });
  startDateCell.numFmt = dateFormat;

  dashboard.getCell('D5').value = 'Target End:';
  dashboard.getCell('D5').font = { bold: true };
  const endDateCell = dashboard.getCell('E5');
  Object.assign(endDateCell, { style: inputStyle });
  endDateCell.numFmt = dateFormat;

  // Summary Cards
  dashboard.getCell('B7').value = 'BUDGET SUMMARY';
  dashboard.getCell('B7').font = { bold: true, size: 14 };

  // Headers for summary
  const summaryHeaders = ['Category', 'Estimated', 'Actual', 'Variance', '% Spent'];
  const summaryHeaderRow = dashboard.getRow(9);
  summaryHeaders.forEach((header, index) => {
    const cell = summaryHeaderRow.getCell(index + 2);
    cell.value = header;
    Object.assign(cell, { style: headerStyle });
  });

  // Category summaries with formulas
  const categories = [
    { name: 'Professional Fees', sheet: 'Professional Fees' },
    { name: 'Planning & Regs', sheet: 'Planning & Regs' },
    { name: 'Construction', sheet: 'Construction' },
    { name: 'Utilities', sheet: 'Utilities' },
    { name: 'Making Good', sheet: 'Making Good' },
    { name: 'Landscaping', sheet: 'Landscaping' },
    { name: 'Contingency', sheet: 'Contingency' },
  ];

  categories.forEach((cat, index) => {
    const rowNum = index + 10;
    const row = dashboard.getRow(rowNum);

    row.getCell(2).value = cat.name;
    applyBorder(row.getCell(2));

    // Get the total row for each category sheet
    const estCell = row.getCell(3);
    estCell.value = { formula: `'${cat.sheet}'!C${getLastRowForCategory(cat.name)}` };
    estCell.numFmt = currencyFormat;
    applyBorder(estCell);

    const actCell = row.getCell(4);
    actCell.value = { formula: `'${cat.sheet}'!D${getLastRowForCategory(cat.name)}` };
    actCell.numFmt = currencyFormat;
    applyBorder(actCell);

    const varCell = row.getCell(5);
    varCell.value = { formula: `D${rowNum}-C${rowNum}` };
    varCell.numFmt = currencyFormat;
    applyBorder(varCell);

    const pctCell = row.getCell(6);
    pctCell.value = { formula: `IF(C${rowNum}=0,0,D${rowNum}/C${rowNum})` };
    pctCell.numFmt = percentFormat;
    applyBorder(pctCell);
  });

  // Grand total
  const grandTotalRow = dashboard.getRow(17);
  grandTotalRow.getCell(2).value = 'GRAND TOTAL';
  grandTotalRow.getCell(2).font = { bold: true, size: 12 };
  grandTotalRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
  grandTotalRow.getCell(2).font = { bold: true, color: { argb: COLORS.white } };

  const gtEstCell = grandTotalRow.getCell(3);
  gtEstCell.value = { formula: 'SUM(C10:C16)' };
  gtEstCell.numFmt = currencyFormat;
  gtEstCell.font = { bold: true, color: { argb: COLORS.white } };
  gtEstCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };

  const gtActCell = grandTotalRow.getCell(4);
  gtActCell.value = { formula: 'SUM(D10:D16)' };
  gtActCell.numFmt = currencyFormat;
  gtActCell.font = { bold: true, color: { argb: COLORS.white } };
  gtActCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };

  const gtVarCell = grandTotalRow.getCell(5);
  gtVarCell.value = { formula: 'D17-C17' };
  gtVarCell.numFmt = currencyFormat;
  gtVarCell.font = { bold: true, color: { argb: COLORS.white } };
  gtVarCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };

  const gtPctCell = grandTotalRow.getCell(6);
  gtPctCell.value = { formula: 'IF(C17=0,0,D17/C17)' };
  gtPctCell.numFmt = percentFormat;
  gtPctCell.font = { bold: true, color: { argb: COLORS.white } };
  gtPctCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };

  // Status indicators
  dashboard.getCell('H7').value = 'BUDGET STATUS';
  dashboard.getCell('H7').font = { bold: true, size: 14 };

  dashboard.getCell('H9').value = 'Overall Status:';
  dashboard.getCell('H9').font = { bold: true };
  const statusCell = dashboard.getCell('I9');
  statusCell.value = { formula: 'IF(E17>C16,"OVER BUDGET",IF(E17>0,"CAUTION","ON TRACK"))' };
  statusCell.font = { bold: true };

  dashboard.getCell('H11').value = 'Contingency Used:';
  dashboard.getCell('H11').font = { bold: true };
  const contUsedCell = dashboard.getCell('I11');
  contUsedCell.value = { formula: 'Contingency!D10' };
  contUsedCell.numFmt = currencyFormat;

  dashboard.getCell('H12').value = 'Contingency Remaining:';
  dashboard.getCell('H12').font = { bold: true };
  const contRemCell = dashboard.getCell('I12');
  contRemCell.value = { formula: 'Contingency!C10-Contingency!D10' };
  contRemCell.numFmt = currencyFormat;

  // Contingency warning
  dashboard.getCell('H14').value = 'Contingency Alert:';
  dashboard.getCell('H14').font = { bold: true };
  const alertCell = dashboard.getCell('I14');
  alertCell.value = { formula: 'IF(I12<0,"DEPLETED!",IF(I11/Contingency!C10>0.5,"Over 50% used","OK"))' };

  // Conditional formatting for variance column
  dashboard.addConditionalFormatting({
    ref: 'E10:E17',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0'],
        style: { font: { color: { argb: COLORS.red } } },
        priority: 1,
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: { font: { color: { argb: COLORS.green } } },
        priority: 2,
      },
    ],
  });

  // Conditional formatting for status
  dashboard.addConditionalFormatting({
    ref: 'I9',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'OVER',
        style: { font: { color: { argb: COLORS.red }, bold: true } },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'CAUTION',
        style: { font: { color: { argb: COLORS.yellow }, bold: true } },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'ON TRACK',
        style: { font: { color: { argb: COLORS.green }, bold: true } },
        priority: 3,
      },
    ],
  });

  // ===== CATEGORY SHEETS =====

  // Professional Fees
  createCategorySheet(workbook, 'Professional Fees', COLORS.primaryBlue, [
    { item: 'Architect - Initial Design', estimated: 2500 },
    { item: 'Architect - Planning Drawings', estimated: 1500 },
    { item: 'Architect - Building Regs Drawings', estimated: 2000 },
    { item: 'Structural Engineer', estimated: 1200 },
    { item: 'Party Wall Surveyor', estimated: 1500 },
    { item: 'Building Control (Private)', estimated: 800 },
    { item: 'CDM Advisor (if required)', estimated: 500 },
    { item: 'Project Manager', estimated: 0, notes: 'Optional' },
  ]);

  // Planning & Regs
  createCategorySheet(workbook, 'Planning & Regs', COLORS.orange, [
    { item: 'Planning Application Fee', estimated: 528 },
    { item: 'Lawful Development Certificate', estimated: 129, notes: 'If applicable' },
    { item: 'Building Regulations Fee', estimated: 500 },
    { item: 'Building Control Inspections', estimated: 400 },
    { item: 'Thames Water Build Over (if required)', estimated: 0 },
    { item: 'Party Wall Notices & Awards', estimated: 750 },
  ]);

  // Construction
  createCategorySheet(workbook, 'Construction', COLORS.green, [
    { item: 'Groundworks & Foundations', estimated: 8000 },
    { item: 'Structural Steel', estimated: 3500 },
    { item: 'Brickwork / Block work', estimated: 6000 },
    { item: 'Roofing', estimated: 4500 },
    { item: 'Windows & Doors', estimated: 5000 },
    { item: 'First Fix Electrics', estimated: 2000 },
    { item: 'First Fix Plumbing', estimated: 1800 },
    { item: 'Plastering', estimated: 2500 },
    { item: 'Second Fix Electrics', estimated: 1500 },
    { item: 'Second Fix Plumbing', estimated: 1200 },
    { item: 'Kitchen Installation', estimated: 8000 },
    { item: 'Bathroom Installation', estimated: 4000 },
    { item: 'Flooring', estimated: 3000 },
    { item: 'Decoration', estimated: 2000 },
    { item: 'Skirting & Architraves', estimated: 800 },
  ]);

  // Utilities
  createCategorySheet(workbook, 'Utilities', COLORS.primaryBlue, [
    { item: 'Electrical Supply Upgrade', estimated: 0, notes: 'If required' },
    { item: 'Gas Connection/Modification', estimated: 0 },
    { item: 'Water Connection/Modification', estimated: 0 },
    { item: 'Drainage Connection', estimated: 500 },
    { item: 'Temporary Electrics (site)', estimated: 200 },
    { item: 'Skip Hire', estimated: 1500 },
    { item: 'Scaffolding', estimated: 1200 },
  ]);

  // Making Good
  createCategorySheet(workbook, 'Making Good', COLORS.orange, [
    { item: 'Matching Existing Flooring', estimated: 500 },
    { item: 'Matching Existing Skirting', estimated: 200 },
    { item: 'Touch-up Decoration', estimated: 300 },
    { item: 'Cleaning', estimated: 400 },
    { item: 'Repairs to Existing', estimated: 500 },
  ]);

  // Landscaping
  createCategorySheet(workbook, 'Landscaping', COLORS.green, [
    { item: 'Patio/Paving', estimated: 2000 },
    { item: 'Fencing', estimated: 800 },
    { item: 'Turfing/Planting', estimated: 500 },
    { item: 'Garden Clearance', estimated: 300 },
    { item: 'External Lighting', estimated: 400 },
  ]);

  // Contingency
  createCategorySheet(workbook, 'Contingency', COLORS.red, [
    { item: 'General Contingency (10-15%)', estimated: 7500 },
    { item: 'Design Changes Reserve', estimated: 2000 },
    { item: 'Unforeseen Works', estimated: 2500 },
    { item: 'Price Increases', estimated: 1500 },
    { item: 'Weather Delays', estimated: 500 },
  ]);

  // ===== VARIATION LOG SHEET =====
  const variationLog = workbook.addWorksheet('Variation Log', {
    properties: { tabColor: { argb: COLORS.red } },
  });

  variationLog.columns = [
    { width: 5 },
    { width: 8 },
    { width: 12 },
    { width: 30 },
    { width: 15 },
    { width: 12 },
    { width: 15 },
    { width: 20 },
  ];

  // Title
  variationLog.mergeCells('B2:H2');
  const varTitleCell = variationLog.getCell('B2');
  varTitleCell.value = 'VARIATION LOG';
  varTitleCell.font = { bold: true, size: 14, color: { argb: COLORS.red } };

  // Headers
  const varHeaders = ['', 'No.', 'Date', 'Description', 'Cost Impact', 'Status', 'Approved By', 'Notes'];
  const varHeaderRow = variationLog.getRow(4);
  varHeaders.forEach((header, index) => {
    const cell = varHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Empty rows for data entry
  for (let i = 5; i <= 24; i++) {
    const row = variationLog.getRow(i);
    row.getCell(2).value = i - 4;
    applyBorder(row.getCell(2));

    const dateCell = row.getCell(3);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    Object.assign(row.getCell(4), { style: inputStyle });

    const costCell = row.getCell(5);
    Object.assign(costCell, { style: inputStyle });
    costCell.numFmt = currencyFormat;

    const statusCell = row.getCell(6);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Pending,Approved,Rejected"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    Object.assign(row.getCell(7), { style: inputStyle });
    Object.assign(row.getCell(8), { style: inputStyle });
  }

  // Total variations
  const varTotalRow = variationLog.getRow(25);
  varTotalRow.getCell(4).value = 'TOTAL APPROVED VARIATIONS';
  varTotalRow.getCell(4).font = { bold: true };
  const varTotalCell = varTotalRow.getCell(5);
  varTotalCell.value = { formula: 'SUMIF(F5:F24,"Approved",E5:E24)' };
  varTotalCell.numFmt = currencyFormat;
  varTotalCell.font = { bold: true, color: { argb: COLORS.red } };
  varTotalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

// Helper function to get the total row number for each category
function getLastRowForCategory(category: string): number {
  const rowCounts: Record<string, number> = {
    'Professional Fees': 13,
    'Planning & Regs': 11,
    'Construction': 20,
    'Utilities': 12,
    'Making Good': 10,
    'Landscaping': 10,
    'Contingency': 10,
  };
  return rowCounts[category] || 10;
}
