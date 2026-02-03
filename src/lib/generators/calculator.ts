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

export async function generateExtensionCalculator(): Promise<Buffer> {
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
    ['', 'EXTENSION DECISION CALCULATOR'],
    ['', ''],
    ['', 'Welcome to the Extension Decision Calculator. This tool helps you compare the financial'],
    ['', 'implications of extending your home versus moving to a new property.'],
    ['', ''],
    ['', 'HOW TO USE THIS CALCULATOR:'],
    ['', ''],
    ['', '1. YOUR PROPERTY SHEET'],
    ['', '   - Enter your current property details including value, mortgage, and location'],
    ['', '   - Select your region for accurate cost adjustments'],
    ['', ''],
    ['', '2. EXTENSION OPTION SHEET'],
    ['', '   - Enter details about your planned extension'],
    ['', '   - Adjust the value uplift percentage based on your research'],
    ['', '   - Review the total costs and expected value increase'],
    ['', ''],
    ['', '3. MOVING OPTION SHEET'],
    ['', '   - Enter the target property price'],
    ['', '   - View calculated stamp duty and other moving costs'],
    ['', ''],
    ['', '4. COMPARISON SHEET'],
    ['', '   - See a side-by-side comparison of both options'],
    ['', '   - Review the financial summary and recommendation'],
    ['', ''],
    ['', 'IMPORTANT NOTES:'],
    ['', '- Orange cells are for your input'],
    ['', '- All calculations update automatically'],
    ['', '- Regional adjustments affect construction costs'],
    ['', '- Stamp duty is calculated based on current UK rates'],
    ['', ''],
    ['', 'Color Key:'],
    ['', '  Orange background = Enter your data here'],
    ['', '  Blue text = Calculated result'],
    ['', '  Green = Positive/favorable outcome'],
    ['', '  Red = Warning or unfavorable outcome'],
  ];

  instructions.forEach((row, index) => {
    const excelRow = instructionsSheet.addRow(row);
    if (index === 1) {
      excelRow.getCell(2).font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };
    } else if (row[1].startsWith('HOW TO USE') || row[1].startsWith('IMPORTANT')) {
      excelRow.getCell(2).font = { bold: true, size: 12, color: { argb: COLORS.slateHeader } };
    } else if (row[1].match(/^\d\./)) {
      excelRow.getCell(2).font = { bold: true, size: 11 };
    }
  });

  // ===== YOUR PROPERTY SHEET =====
  const propertySheet = workbook.addWorksheet('Your Property', {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  propertySheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 20 },
    { width: 15 },
    { width: 25 },
  ];

  // Title
  propertySheet.mergeCells('B2:D2');
  const titleCell = propertySheet.getCell('B2');
  titleCell.value = 'YOUR PROPERTY DETAILS';
  titleCell.font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  // Property details section
  const propertyData = [
    ['', '', '', '', ''],
    ['', 'Current Property Value', '', '', 'Enter your best estimate of current market value'],
    ['', 'Outstanding Mortgage', '', '', 'Current mortgage balance'],
    ['', 'Available Equity', '', '', 'Calculated: Value - Mortgage'],
    ['', '', '', '', ''],
    ['', 'Region', '', '', 'Select for regional cost adjustment'],
    ['', 'Regional Cost Multiplier', '', '', 'Automatically calculated'],
  ];

  propertyData.forEach((row, index) => {
    const excelRow = propertySheet.addRow(row);
    if (index === 1 || index === 2 || index === 5) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      if (index === 1) inputCell.value = 450000;
      if (index === 2) inputCell.value = 200000;
      if (index === 5) {
        inputCell.dataValidation = {
          type: 'list',
          allowBlank: false,
          formulae: ['"London,South East,South West,East,West Midlands,East Midlands,Yorkshire,North West,North East,Wales,Scotland,Northern Ireland"'],
        };
        inputCell.value = 'South East';
      }
      inputCell.numFmt = index === 5 ? '@' : currencyFormat;
    }
    if (index === 3) {
      const calcCell = excelRow.getCell(3);
      calcCell.value = { formula: 'C5-C6' };
      calcCell.numFmt = currencyFormat;
      calcCell.font = { bold: true, color: { argb: COLORS.primaryBlue } };
    }
    if (index === 6) {
      const multiplierCell = excelRow.getCell(3);
      multiplierCell.value = { formula: 'VLOOKUP(C9,\'Reference Data\'!A:B,2,FALSE)' };
      multiplierCell.numFmt = '0.00';
      multiplierCell.font = { color: { argb: COLORS.primaryBlue } };
    }
    excelRow.getCell(5).font = { italic: true, color: { argb: 'FF64748b' }, size: 10 };
  });

  // ===== EXTENSION OPTION SHEET =====
  const extensionSheet = workbook.addWorksheet('Extension Option', {
    properties: { tabColor: { argb: COLORS.green } },
  });

  extensionSheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 18 },
    { width: 18 },
    { width: 30 },
  ];

  // Title
  extensionSheet.mergeCells('B2:D2');
  const extTitleCell = extensionSheet.getCell('B2');
  extTitleCell.value = 'EXTENSION OPTION ANALYSIS';
  extTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.green } };

  // Extension type header
  extensionSheet.getCell('B4').value = 'Extension Details';
  extensionSheet.getCell('B4').font = { bold: true, size: 12 };

  const extensionData = [
    ['', 'Extension Type', '', '', ''],
    ['', 'Size (sqm)', '', '', 'Floor area of extension'],
    ['', 'Base Cost per sqm', '', '', 'Average for extension type'],
    ['', 'Regional Adjusted Cost/sqm', '', '', 'Adjusted for your region'],
    ['', '', '', '', ''],
    ['', 'CONSTRUCTION COSTS', '', '', ''],
    ['', 'Base Construction', '', '', 'Size x Adjusted cost'],
    ['', 'Professional Fees (12%)', '', '', 'Architect, SE, etc.'],
    ['', 'Planning & Building Regs', '', '', 'Applications and inspections'],
    ['', 'Contingency (15%)', '', '', 'Recommended buffer'],
    ['', '', '', '', ''],
    ['', 'TOTAL EXTENSION COST', '', '', ''],
    ['', '', '', '', ''],
    ['', 'VALUE ANALYSIS', '', '', ''],
    ['', 'Value Uplift %', '', '', 'Typical: 50-80% of cost'],
    ['', 'Expected Value Increase', '', '', 'Based on uplift percentage'],
    ['', 'New Property Value', '', '', 'Current + Value Increase'],
    ['', '', '', '', ''],
    ['', 'Net Cost of Extension', '', '', 'Total Cost - Value Increase'],
  ];

  extensionData.forEach((row, index) => {
    const excelRow = extensionSheet.addRow(row);
    const rowNum = index + 5;

    if (index === 0) {
      // Extension type dropdown
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.dataValidation = {
        type: 'list',
        allowBlank: false,
        formulae: ['"Single Storey Rear,Double Storey Rear,Side Return,Loft Conversion,Basement,Wrap Around"'],
      };
      inputCell.value = 'Single Storey Rear';
    }
    if (index === 1) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 25;
      inputCell.numFmt = '0';
    }
    if (index === 2) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'VLOOKUP(C5,\'Reference Data\'!D:E,2,FALSE)' };
      cell.numFmt = currencyFormat;
    }
    if (index === 3) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'C7*\'Your Property\'!C10' };
      cell.numFmt = currencyFormat;
      cell.font = { color: { argb: COLORS.primaryBlue } };
    }
    if (index === 5) {
      excelRow.getCell(2).font = { bold: true, size: 11 };
    }
    if (index === 6) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'C6*C8' };
      cell.numFmt = currencyFormat;
    }
    if (index === 7) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'C11*0.12' };
      cell.numFmt = currencyFormat;
    }
    if (index === 8) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 3500;
      inputCell.numFmt = currencyFormat;
    }
    if (index === 9) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: '(C11+C12+C13)*0.15' };
      cell.numFmt = currencyFormat;
    }
    if (index === 11) {
      excelRow.getCell(2).font = { bold: true, size: 12 };
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'SUM(C11:C14)' };
      cell.numFmt = currencyFormat;
      cell.font = { bold: true, size: 12, color: { argb: COLORS.primaryBlue } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
    }
    if (index === 13) {
      excelRow.getCell(2).font = { bold: true, size: 11 };
    }
    if (index === 14) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 0.7;
      inputCell.numFmt = percentFormat;
      inputCell.dataValidation = {
        type: 'decimal',
        operator: 'between',
        allowBlank: false,
        showErrorMessage: true,
        errorTitle: 'Invalid Value',
        error: 'Please enter a percentage between 0% and 150%',
        formulae: [0, 1.5],
      };
    }
    if (index === 15) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'C16*C19' };
      cell.numFmt = currencyFormat;
      cell.font = { color: { argb: COLORS.green } };
    }
    if (index === 16) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: '\'Your Property\'!C5+C20' };
      cell.numFmt = currencyFormat;
      cell.font = { bold: true, color: { argb: COLORS.green } };
    }
    if (index === 18) {
      excelRow.getCell(2).font = { bold: true };
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'C16-C20' };
      cell.numFmt = currencyFormat;
      cell.font = { bold: true };
    }
    excelRow.getCell(5).font = { italic: true, color: { argb: 'FF64748b' }, size: 10 };
  });

  // ===== MOVING OPTION SHEET =====
  const movingSheet = workbook.addWorksheet('Moving Option', {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  movingSheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 18 },
    { width: 18 },
    { width: 30 },
  ];

  // Title
  movingSheet.mergeCells('B2:D2');
  const movTitleCell = movingSheet.getCell('B2');
  movTitleCell.value = 'MOVING OPTION ANALYSIS';
  movTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.orange } };

  movingSheet.getCell('B4').value = 'Target Property';
  movingSheet.getCell('B4').font = { bold: true, size: 12 };

  const movingData = [
    ['', 'Target Property Price', '', '', 'Price of property you want to move to'],
    ['', '', '', '', ''],
    ['', 'STAMP DUTY CALCULATION', '', '', ''],
    ['', 'Stamp Duty Land Tax', '', '', 'Calculated based on UK rates'],
    ['', '', '', '', ''],
    ['', 'SELLING COSTS', '', '', ''],
    ['', 'Estate Agent Fees (1.5%)', '', '', 'On current property'],
    ['', 'Conveyancing (Sale)', '', '', 'Legal fees for sale'],
    ['', 'EPC Certificate', '', '', ''],
    ['', '', '', '', ''],
    ['', 'BUYING COSTS', '', '', ''],
    ['', 'Conveyancing (Purchase)', '', '', 'Legal fees for purchase'],
    ['', 'Survey', '', '', 'Homebuyer or full structural'],
    ['', 'Mortgage Fees', '', '', 'Arrangement and valuation'],
    ['', 'Removal Costs', '', '', ''],
    ['', '', '', '', ''],
    ['', 'TOTAL MOVING COSTS', '', '', ''],
    ['', '', '', '', ''],
    ['', 'NET POSITION AFTER MOVE', '', '', ''],
    ['', 'Sale Proceeds', '', '', 'After mortgage and selling costs'],
    ['', 'Required Mortgage', '', '', 'Target price + costs - proceeds'],
  ];

  movingData.forEach((row, index) => {
    const excelRow = movingSheet.addRow(row);

    if (index === 0) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 650000;
      inputCell.numFmt = currencyFormat;
    }
    if ([2, 5, 10, 15, 18].includes(index)) {
      excelRow.getCell(2).font = { bold: true, size: 11 };
    }
    if (index === 3) {
      // Complex stamp duty formula
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'IF(C5<=250000,0,IF(C5<=925000,(MIN(C5,925000)-250000)*0.05,IF(C5<=1500000,(675000*0.05)+(MIN(C5,1500000)-925000)*0.1,(675000*0.05)+(575000*0.1)+(C5-1500000)*0.12)))' };
      cell.numFmt = currencyFormat;
      cell.font = { color: { argb: COLORS.red } };
    }
    if (index === 6) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: '\'Your Property\'!C5*0.015' };
      cell.numFmt = currencyFormat;
    }
    if (index === 7) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 1500;
      inputCell.numFmt = currencyFormat;
    }
    if (index === 8) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 120;
      inputCell.numFmt = currencyFormat;
    }
    if (index === 11) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 2000;
      inputCell.numFmt = currencyFormat;
    }
    if (index === 12) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 600;
      inputCell.numFmt = currencyFormat;
    }
    if (index === 13) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 1500;
      inputCell.numFmt = currencyFormat;
    }
    if (index === 14) {
      const inputCell = excelRow.getCell(3);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.value = 1500;
      inputCell.numFmt = currencyFormat;
    }
    if (index === 16) {
      excelRow.getCell(2).font = { bold: true, size: 12 };
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'C8+C11+C12+C13+C16+C17+C18+C19' };
      cell.numFmt = currencyFormat;
      cell.font = { bold: true, size: 12, color: { argb: COLORS.red } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
    }
    if (index === 19) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: '\'Your Property\'!C5-\'Your Property\'!C6-C11-C12-C13' };
      cell.numFmt = currencyFormat;
    }
    if (index === 20) {
      const cell = excelRow.getCell(3);
      cell.value = { formula: 'C5+C21-C24' };
      cell.numFmt = currencyFormat;
      cell.font = { bold: true, color: { argb: COLORS.primaryBlue } };
    }
    excelRow.getCell(5).font = { italic: true, color: { argb: 'FF64748b' }, size: 10 };
  });

  // ===== COMPARISON SHEET =====
  const comparisonSheet = workbook.addWorksheet('Comparison', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  comparisonSheet.columns = [
    { width: 5 },
    { width: 30 },
    { width: 20 },
    { width: 20 },
    { width: 25 },
  ];

  // Title
  comparisonSheet.mergeCells('B2:D2');
  const compTitleCell = comparisonSheet.getCell('B2');
  compTitleCell.value = 'EXTEND VS MOVE COMPARISON';
  compTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  // Headers
  comparisonSheet.getCell('C4').value = 'EXTEND';
  comparisonSheet.getCell('C4').style = headerStyle;
  comparisonSheet.getCell('D4').value = 'MOVE';
  comparisonSheet.getCell('D4').style = headerStyle;

  const comparisonData = [
    ['', 'Total Cost', '', ''],
    ['', 'New Property Value', '', ''],
    ['', 'Net Cost / Position', '', ''],
    ['', 'Mortgage Required', '', ''],
    ['', '', '', ''],
    ['', 'Disruption Level', 'Medium-High', 'Very High'],
    ['', 'Time to Complete', '4-8 months', '3-6 months'],
    ['', 'Risk Level', 'Medium', 'Low-Medium'],
    ['', '', '', ''],
    ['', 'FINANCIAL ADVANTAGE', '', ''],
  ];

  comparisonData.forEach((row, index) => {
    const excelRow = comparisonSheet.addRow(row);
    const rowNum = index + 5;

    if (index === 0) {
      const extCell = excelRow.getCell(3);
      extCell.value = { formula: '\'Extension Option\'!C16' };
      extCell.numFmt = currencyFormat;
      const movCell = excelRow.getCell(4);
      movCell.value = { formula: '\'Moving Option\'!C21' };
      movCell.numFmt = currencyFormat;
    }
    if (index === 1) {
      const extCell = excelRow.getCell(3);
      extCell.value = { formula: '\'Extension Option\'!C21' };
      extCell.numFmt = currencyFormat;
      const movCell = excelRow.getCell(4);
      movCell.value = { formula: '\'Moving Option\'!C5' };
      movCell.numFmt = currencyFormat;
    }
    if (index === 2) {
      const extCell = excelRow.getCell(3);
      extCell.value = { formula: '\'Extension Option\'!C23' };
      extCell.numFmt = currencyFormat;
      const movCell = excelRow.getCell(4);
      movCell.value = { formula: '\'Moving Option\'!C21' };
      movCell.numFmt = currencyFormat;
    }
    if (index === 3) {
      const extCell = excelRow.getCell(3);
      extCell.value = { formula: '\'Extension Option\'!C16-\'Your Property\'!C7' };
      extCell.numFmt = currencyFormat;
      const movCell = excelRow.getCell(4);
      movCell.value = { formula: '\'Moving Option\'!C25' };
      movCell.numFmt = currencyFormat;
    }
    if (index === 9) {
      excelRow.getCell(2).font = { bold: true, size: 12 };
      const resultCell = excelRow.getCell(3);
      comparisonSheet.mergeCells(`C${rowNum}:D${rowNum}`);
      resultCell.value = { formula: 'IF(\'Extension Option\'!C23<\'Moving Option\'!C21,"EXTEND - Save "&TEXT(\'Moving Option\'!C21-\'Extension Option\'!C23,"£#,##0"),"MOVE - Better value by "&TEXT(\'Extension Option\'!C23-\'Moving Option\'!C21,"£#,##0"))' };
      resultCell.font = { bold: true, size: 14 };
      resultCell.alignment = { horizontal: 'center' };
    }
  });

  // Conditional formatting for comparison
  comparisonSheet.addConditionalFormatting({
    ref: 'C5:C8',
    rules: [{
      type: 'expression',
      formulae: ['$C5<$D5'],
      style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } } },
      priority: 1,
    }],
  });

  comparisonSheet.addConditionalFormatting({
    ref: 'D5:D8',
    rules: [{
      type: 'expression',
      formulae: ['$D5<$C5'],
      style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } } },
      priority: 1,
    }],
  });

  // ===== REFERENCE DATA SHEET (HIDDEN) =====
  const refSheet = workbook.addWorksheet('Reference Data', {
    properties: { tabColor: { argb: 'FF94a3b8' } },
    state: 'hidden',
  });

  refSheet.columns = [
    { width: 20 },
    { width: 15 },
    { width: 5 },
    { width: 25 },
    { width: 15 },
  ];

  // Regional multipliers
  refSheet.getCell('A1').value = 'Region';
  refSheet.getCell('B1').value = 'Multiplier';
  const regions = [
    ['London', 1.35],
    ['South East', 1.15],
    ['South West', 1.05],
    ['East', 1.08],
    ['West Midlands', 0.95],
    ['East Midlands', 0.92],
    ['Yorkshire', 0.90],
    ['North West', 0.88],
    ['North East', 0.85],
    ['Wales', 0.88],
    ['Scotland', 0.92],
    ['Northern Ireland', 0.85],
  ];

  regions.forEach((region, index) => {
    refSheet.getCell(`A${index + 2}`).value = region[0];
    refSheet.getCell(`B${index + 2}`).value = region[1];
  });

  // Extension type costs
  refSheet.getCell('D1').value = 'Extension Type';
  refSheet.getCell('E1').value = 'Cost/sqm';
  const extensionTypes = [
    ['Single Storey Rear', 2200],
    ['Double Storey Rear', 1800],
    ['Side Return', 2400],
    ['Loft Conversion', 1500],
    ['Basement', 3500],
    ['Wrap Around', 2000],
  ];

  extensionTypes.forEach((type, index) => {
    refSheet.getCell(`D${index + 2}`).value = type[0];
    refSheet.getCell(`E${index + 2}`).value = type[1];
  });

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
