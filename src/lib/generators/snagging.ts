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

const dateFormat = 'dd/mm/yyyy';

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    left: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    right: { style: 'thin', color: { argb: 'FFe2e8f0' } },
  };
}

function createRoomSheet(
  workbook: ExcelJS.Workbook,
  roomName: string,
  checklistItems: string[]
) {
  const sheet = workbook.addWorksheet(roomName, {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  sheet.columns = [
    { width: 5 },
    { width: 8 },
    { width: 35 },
    { width: 12 },
    { width: 35 },
    { width: 12 },
    { width: 15 },
    { width: 12 },
    { width: 20 },
  ];

  // Title
  sheet.mergeCells('B2:I2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = `${roomName.toUpperCase()} - SNAGGING CHECKLIST`;
  titleCell.font = { bold: true, size: 14, color: { argb: COLORS.primaryBlue } };

  // Room photo reference
  sheet.getCell('B4').value = 'Photo Folder Reference:';
  sheet.getCell('B4').font = { bold: true };
  const photoRefCell = sheet.getCell('C4');
  Object.assign(photoRefCell, { style: inputStyle });
  sheet.mergeCells('C4:E4');
  photoRefCell.alignment = { horizontal: 'left', vertical: 'middle' };

  // Headers
  const headers = ['', 'Item', 'Check Area', 'Status', 'Issue Description', 'Severity', 'Photo Ref', 'Resolved', 'Resolution Date'];
  const headerRow = sheet.getRow(6);
  headers.forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Checklist items
  checklistItems.forEach((item, index) => {
    const rowNum = index + 7;
    const row = sheet.getRow(rowNum);

    row.getCell(2).value = index + 1;
    row.getCell(2).alignment = { horizontal: 'center' };
    applyBorder(row.getCell(2));

    row.getCell(3).value = item;
    applyBorder(row.getCell(3));

    const statusCell = row.getCell(4);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"OK,Issue,N/A"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    const descCell = row.getCell(5);
    Object.assign(descCell, { style: inputStyle });
    descCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    const severityCell = row.getCell(6);
    severityCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Minor,Moderate,Major"'],
    };
    Object.assign(severityCell, { style: inputStyle });

    const photoCell = row.getCell(7);
    Object.assign(photoCell, { style: inputStyle });
    photoCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const resolvedCell = row.getCell(8);
    resolvedCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Yes,No"'],
    };
    Object.assign(resolvedCell, { style: inputStyle });

    const dateCell = row.getCell(9);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;
  });

  // Additional rows for custom items
  const startAdditional = checklistItems.length + 7;
  for (let i = startAdditional; i < startAdditional + 10; i++) {
    const row = sheet.getRow(i);

    row.getCell(2).value = i - 6;
    row.getCell(2).alignment = { horizontal: 'center' };
    applyBorder(row.getCell(2));

    const areaCell = row.getCell(3);
    Object.assign(areaCell, { style: inputStyle });
    areaCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const statusCell = row.getCell(4);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"OK,Issue,N/A"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    const descCell = row.getCell(5);
    Object.assign(descCell, { style: inputStyle });
    descCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    const severityCell = row.getCell(6);
    severityCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Minor,Moderate,Major"'],
    };
    Object.assign(severityCell, { style: inputStyle });

    const photoCell = row.getCell(7);
    Object.assign(photoCell, { style: inputStyle });
    photoCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const resolvedCell = row.getCell(8);
    resolvedCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Yes,No"'],
    };
    Object.assign(resolvedCell, { style: inputStyle });

    const dateCell = row.getCell(9);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;
  }

  const totalRowNum = startAdditional + 10;

  // Summary for this room
  sheet.getCell(`B${totalRowNum + 1}`).value = 'ROOM SUMMARY';
  sheet.getCell(`B${totalRowNum + 1}`).font = { bold: true, size: 12 };

  sheet.getCell(`B${totalRowNum + 2}`).value = 'Total Issues:';
  sheet.getCell(`C${totalRowNum + 2}`).value = { formula: `COUNTIF(D7:D${totalRowNum - 1},"Issue")` };

  sheet.getCell(`B${totalRowNum + 3}`).value = 'Resolved:';
  sheet.getCell(`C${totalRowNum + 3}`).value = { formula: `COUNTIF(H7:H${totalRowNum - 1},"Yes")` };

  sheet.getCell(`B${totalRowNum + 4}`).value = 'Outstanding:';
  const outstandingCell = sheet.getCell(`C${totalRowNum + 4}`);
  outstandingCell.value = { formula: `C${totalRowNum + 2}-C${totalRowNum + 3}` };
  outstandingCell.font = { bold: true };

  sheet.getCell(`E${totalRowNum + 2}`).value = 'Major Issues:';
  sheet.getCell(`F${totalRowNum + 2}`).value = { formula: `COUNTIF(F7:F${totalRowNum - 1},"Major")` };
  sheet.getCell(`F${totalRowNum + 2}`).font = { color: { argb: COLORS.red }, bold: true };

  sheet.getCell(`E${totalRowNum + 3}`).value = 'Moderate Issues:';
  sheet.getCell(`F${totalRowNum + 3}`).value = { formula: `COUNTIF(F7:F${totalRowNum - 1},"Moderate")` };
  sheet.getCell(`F${totalRowNum + 3}`).font = { color: { argb: COLORS.yellow } };

  sheet.getCell(`E${totalRowNum + 4}`).value = 'Minor Issues:';
  sheet.getCell(`F${totalRowNum + 4}`).value = { formula: `COUNTIF(F7:F${totalRowNum - 1},"Minor")` };
  sheet.getCell(`F${totalRowNum + 4}`).font = { color: { argb: COLORS.green } };

  // Conditional formatting for status
  sheet.addConditionalFormatting({
    ref: `D7:D${totalRowNum - 1}`,
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Issue',
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

  // Conditional formatting for severity
  sheet.addConditionalFormatting({
    ref: `F7:F${totalRowNum - 1}`,
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Major',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Moderate',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Minor',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
        },
        priority: 3,
      },
    ],
  });

  // Conditional formatting for resolved
  sheet.addConditionalFormatting({
    ref: `H7:H${totalRowNum - 1}`,
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Yes',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'No',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red } },
        },
        priority: 2,
      },
    ],
  });

  return sheet;
}

export async function generateSnaggingChecklist(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Extension Survival Guide';
  workbook.created = new Date();

  // ===== SUMMARY SHEET =====
  const summarySheet = workbook.addWorksheet('Summary', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  summarySheet.columns = [
    { width: 5 },
    { width: 25 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
  ];

  // Title
  summarySheet.mergeCells('B2:H2');
  const titleCell = summarySheet.getCell('B2');
  titleCell.value = 'SNAGGING CHECKLIST - SUMMARY';
  titleCell.font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };

  // Project info
  summarySheet.getCell('B4').value = 'Project Address:';
  summarySheet.getCell('B4').font = { bold: true };
  const addrCell = summarySheet.getCell('C4');
  Object.assign(addrCell, { style: inputStyle });
  summarySheet.mergeCells('C4:F4');
  addrCell.alignment = { horizontal: 'left', vertical: 'middle' };

  summarySheet.getCell('B5').value = 'Builder:';
  summarySheet.getCell('B5').font = { bold: true };
  const builderCell = summarySheet.getCell('C5');
  Object.assign(builderCell, { style: inputStyle });
  summarySheet.mergeCells('C5:D5');
  builderCell.alignment = { horizontal: 'left', vertical: 'middle' };

  summarySheet.getCell('E5').value = 'Inspection Date:';
  summarySheet.getCell('E5').font = { bold: true };
  const dateCell = summarySheet.getCell('F5');
  Object.assign(dateCell, { style: inputStyle });
  dateCell.numFmt = dateFormat;

  summarySheet.getCell('B6').value = 'Inspector:';
  summarySheet.getCell('B6').font = { bold: true };
  const inspCell = summarySheet.getCell('C6');
  Object.assign(inspCell, { style: inputStyle });
  summarySheet.mergeCells('C6:D6');
  inspCell.alignment = { horizontal: 'left', vertical: 'middle' };

  // Overall summary
  summarySheet.getCell('B8').value = 'OVERALL SUMMARY';
  summarySheet.getCell('B8').font = { bold: true, size: 14 };

  // Headers
  const sumHeaders = ['', 'Room/Area', 'Total Issues', 'Major', 'Moderate', 'Minor', 'Resolved', 'Outstanding'];
  const sumHeaderRow = summarySheet.getRow(10);
  sumHeaders.forEach((header, index) => {
    const cell = sumHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Room list - will reference room sheets
  const rooms = [
    'Kitchen',
    'Bathroom 1',
    'Bathroom 2',
    'Bedroom 1',
    'Bedroom 2',
    'Bedroom 3',
    'Living Room',
    'Dining Room',
    'Hallway',
    'External',
    'General',
  ];

  rooms.forEach((room, index) => {
    const rowNum = index + 11;
    const row = summarySheet.getRow(rowNum);

    row.getCell(2).value = room;
    applyBorder(row.getCell(2));

    // Total issues
    const totalCell = row.getCell(3);
    totalCell.value = { formula: `COUNTIF('${room}'!D7:D50,"Issue")` };
    applyBorder(totalCell);

    // Major
    const majorCell = row.getCell(4);
    majorCell.value = { formula: `COUNTIF('${room}'!F7:F50,"Major")` };
    applyBorder(majorCell);

    // Moderate
    const modCell = row.getCell(5);
    modCell.value = { formula: `COUNTIF('${room}'!F7:F50,"Moderate")` };
    applyBorder(modCell);

    // Minor
    const minorCell = row.getCell(6);
    minorCell.value = { formula: `COUNTIF('${room}'!F7:F50,"Minor")` };
    applyBorder(minorCell);

    // Resolved
    const resCell = row.getCell(7);
    resCell.value = { formula: `COUNTIF('${room}'!H7:H50,"Yes")` };
    applyBorder(resCell);

    // Outstanding
    const outCell = row.getCell(8);
    outCell.value = { formula: `C${rowNum}-G${rowNum}` };
    outCell.font = { bold: true };
    applyBorder(outCell);
  });

  // Grand total row
  const totalRowNum = rooms.length + 11;
  const totalRow = summarySheet.getRow(totalRowNum);
  totalRow.getCell(2).value = 'TOTAL';
  totalRow.getCell(2).font = { bold: true };
  totalRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
  totalRow.getCell(2).font = { bold: true, color: { argb: COLORS.white } };

  for (let col = 3; col <= 8; col++) {
    const cell = totalRow.getCell(col);
    const colLetter = String.fromCharCode(64 + col);
    cell.value = { formula: `SUM(${colLetter}11:${colLetter}${totalRowNum - 1})` };
    cell.font = { bold: true, color: { argb: COLORS.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.slateHeader } };
    applyBorder(cell);
  }

  // Final payment gate
  summarySheet.getCell(`B${totalRowNum + 2}`).value = 'FINAL PAYMENT GATE';
  summarySheet.getCell(`B${totalRowNum + 2}`).font = { bold: true, size: 14, color: { argb: COLORS.red } };

  summarySheet.getCell(`B${totalRowNum + 4}`).value = 'Major Issues Outstanding:';
  summarySheet.getCell(`B${totalRowNum + 4}`).font = { bold: true };
  const majorOutCell = summarySheet.getCell(`C${totalRowNum + 4}`);
  majorOutCell.value = { formula: `D${totalRowNum}-SUMPRODUCT(('Kitchen'!F7:F50="Major")*('Kitchen'!H7:H50="Yes"))-SUMPRODUCT(('Bathroom 1'!F7:F50="Major")*('Bathroom 1'!H7:H50="Yes"))-SUMPRODUCT(('Bathroom 2'!F7:F50="Major")*('Bathroom 2'!H7:H50="Yes"))-SUMPRODUCT(('Bedroom 1'!F7:F50="Major")*('Bedroom 1'!H7:H50="Yes"))-SUMPRODUCT(('Bedroom 2'!F7:F50="Major")*('Bedroom 2'!H7:H50="Yes"))-SUMPRODUCT(('Bedroom 3'!F7:F50="Major")*('Bedroom 3'!H7:H50="Yes"))-SUMPRODUCT(('Living Room'!F7:F50="Major")*('Living Room'!H7:H50="Yes"))-SUMPRODUCT(('Dining Room'!F7:F50="Major")*('Dining Room'!H7:H50="Yes"))-SUMPRODUCT(('Hallway'!F7:F50="Major")*('Hallway'!H7:H50="Yes"))-SUMPRODUCT(('External'!F7:F50="Major")*('External'!H7:H50="Yes"))-SUMPRODUCT(('General'!F7:F50="Major")*('General'!H7:H50="Yes"))` };

  summarySheet.getCell(`B${totalRowNum + 5}`).value = 'Total Outstanding:';
  summarySheet.getCell(`B${totalRowNum + 5}`).font = { bold: true };
  summarySheet.getCell(`C${totalRowNum + 5}`).value = { formula: `H${totalRowNum}` };

  summarySheet.getCell(`B${totalRowNum + 7}`).value = 'PAYMENT RECOMMENDATION:';
  summarySheet.getCell(`B${totalRowNum + 7}`).font = { bold: true, size: 12 };
  const recommendCell = summarySheet.getCell(`C${totalRowNum + 7}`);
  recommendCell.value = { formula: `IF(C${totalRowNum + 4}>0,"DO NOT RELEASE - Major issues outstanding",IF(H${totalRowNum}>5,"HOLD PARTIAL - More than 5 items outstanding",IF(H${totalRowNum}>0,"RELEASE WITH RETENTION - Minor items only","CLEAR TO RELEASE FINAL PAYMENT")))` };
  recommendCell.font = { bold: true, size: 12 };
  summarySheet.mergeCells(`C${totalRowNum + 7}:G${totalRowNum + 7}`);

  // Conditional formatting for recommendation
  summarySheet.addConditionalFormatting({
    ref: `C${totalRowNum + 7}`,
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'DO NOT',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'HOLD',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
          font: { color: { argb: COLORS.yellow }, bold: true },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'CLEAR',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green }, bold: true },
        },
        priority: 3,
      },
    ],
  });

  // Conditional formatting for outstanding column
  summarySheet.addConditionalFormatting({
    ref: `H11:H${totalRowNum}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
    ],
  });

  // Sign-off section
  summarySheet.getCell(`B${totalRowNum + 10}`).value = 'SIGN-OFF';
  summarySheet.getCell(`B${totalRowNum + 10}`).font = { bold: true, size: 12 };

  summarySheet.getCell(`B${totalRowNum + 11}`).value = 'All snags resolved:';
  const signoffCell = summarySheet.getCell(`C${totalRowNum + 11}`);
  signoffCell.dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"Yes,No"'],
  };
  Object.assign(signoffCell, { style: inputStyle });

  summarySheet.getCell(`B${totalRowNum + 12}`).value = 'Client Signature:';
  const sigCell = summarySheet.getCell(`C${totalRowNum + 12}`);
  Object.assign(sigCell, { style: inputStyle });
  summarySheet.mergeCells(`C${totalRowNum + 12}:D${totalRowNum + 12}`);

  summarySheet.getCell(`E${totalRowNum + 12}`).value = 'Date:';
  const sigDateCell = summarySheet.getCell(`F${totalRowNum + 12}`);
  Object.assign(sigDateCell, { style: inputStyle });
  sigDateCell.numFmt = dateFormat;

  // ===== CREATE ROOM SHEETS =====

  // Kitchen
  createRoomSheet(workbook, 'Kitchen', [
    'Worktop - surface condition',
    'Worktop - joins and edges',
    'Sink - installation and seals',
    'Tap - operation and finish',
    'Cabinet doors - alignment',
    'Cabinet doors - handles',
    'Drawers - operation',
    'Soft close mechanisms',
    'Kickboards - fit and finish',
    'Splashback - installation',
    'Appliance housing/fit',
    'Extractor - operation',
    'Lighting - operation',
    'Sockets - operation',
    'Flooring - condition',
    'Flooring - edges/thresholds',
    'Walls - finish',
    'Ceiling - finish',
    'Window - operation',
    'Window - seals',
    'Door - operation',
    'Skirting - fit and finish',
  ]);

  // Bathroom 1
  createRoomSheet(workbook, 'Bathroom 1', [
    'Bath - installation',
    'Bath - seals',
    'Shower - installation',
    'Shower - seals',
    'Shower screen/door',
    'Basin - installation',
    'Basin - seals',
    'Toilet - installation',
    'Toilet - operation',
    'Taps - operation',
    'Tiles - grouting',
    'Tiles - condition',
    'Silicone seals',
    'Towel rail - installation',
    'Heated towel rail - operation',
    'Extractor fan - operation',
    'Lighting - operation',
    'Mirror/cabinet',
    'Flooring - condition',
    'Door - operation',
    'Window - operation',
  ]);

  // Bathroom 2
  createRoomSheet(workbook, 'Bathroom 2', [
    'Bath/Shower - installation',
    'Bath/Shower - seals',
    'Basin - installation',
    'Basin - seals',
    'Toilet - installation',
    'Toilet - operation',
    'Taps - operation',
    'Tiles - grouting',
    'Tiles - condition',
    'Silicone seals',
    'Towel rail - installation',
    'Extractor fan - operation',
    'Lighting - operation',
    'Flooring - condition',
    'Door - operation',
    'Window - operation',
  ]);

  // Bedroom 1
  createRoomSheet(workbook, 'Bedroom 1', [
    'Walls - finish',
    'Walls - corners',
    'Ceiling - finish',
    'Flooring - condition',
    'Flooring - edges',
    'Skirting - fit',
    'Skirting - finish',
    'Architrave - fit',
    'Door - operation',
    'Door - finish',
    'Door handles/locks',
    'Windows - operation',
    'Windows - seals',
    'Window board - finish',
    'Radiator - installation',
    'Radiator - operation',
    'Lighting - operation',
    'Light switches - operation',
    'Sockets - operation',
    'Built-in wardrobes',
  ]);

  // Bedroom 2
  createRoomSheet(workbook, 'Bedroom 2', [
    'Walls - finish',
    'Ceiling - finish',
    'Flooring - condition',
    'Skirting - fit and finish',
    'Architrave - fit',
    'Door - operation',
    'Door - finish',
    'Windows - operation',
    'Windows - seals',
    'Radiator - installation',
    'Radiator - operation',
    'Lighting - operation',
    'Sockets - operation',
    'Built-in wardrobes',
  ]);

  // Bedroom 3
  createRoomSheet(workbook, 'Bedroom 3', [
    'Walls - finish',
    'Ceiling - finish',
    'Flooring - condition',
    'Skirting - fit and finish',
    'Architrave - fit',
    'Door - operation',
    'Door - finish',
    'Windows - operation',
    'Windows - seals',
    'Radiator - installation',
    'Radiator - operation',
    'Lighting - operation',
    'Sockets - operation',
  ]);

  // Living Room
  createRoomSheet(workbook, 'Living Room', [
    'Walls - finish',
    'Walls - corners',
    'Ceiling - finish',
    'Flooring - condition',
    'Flooring - edges',
    'Skirting - fit and finish',
    'Architrave - fit',
    'Door(s) - operation',
    'Door(s) - finish',
    'Bi-fold/Sliding doors',
    'Windows - operation',
    'Windows - seals',
    'Radiator(s) - installation',
    'Radiator(s) - operation',
    'Lighting - operation',
    'Dimmer switches',
    'Sockets - operation',
    'TV/data points',
    'Fireplace (if applicable)',
  ]);

  // Dining Room
  createRoomSheet(workbook, 'Dining Room', [
    'Walls - finish',
    'Ceiling - finish',
    'Flooring - condition',
    'Skirting - fit and finish',
    'Door - operation',
    'Windows - operation',
    'Windows - seals',
    'Radiator - installation',
    'Radiator - operation',
    'Lighting - operation',
    'Sockets - operation',
  ]);

  // Hallway
  createRoomSheet(workbook, 'Hallway', [
    'Walls - finish',
    'Ceiling - finish',
    'Flooring - condition',
    'Stairs - finish',
    'Stairs - balustrade',
    'Handrail - secure',
    'Skirting - fit and finish',
    'Front door - operation',
    'Front door - locks',
    'Internal doors - operation',
    'Lighting - operation',
    'Smoke detector - installation',
    'Sockets - operation',
    'Understairs storage',
  ]);

  // External
  createRoomSheet(workbook, 'External', [
    'Brickwork - pointing',
    'Render - finish',
    'Cladding - finish',
    'Roof tiles - condition',
    'Guttering - installation',
    'Downpipes - installation',
    'Fascias - finish',
    'Soffits - finish',
    'External doors - operation',
    'External doors - seals',
    'Windows - external finish',
    'Window sills - finish',
    'Driveway - condition',
    'Patio - condition',
    'Fencing - condition',
    'Garden - cleared',
    'External lighting',
    'External taps',
    'Drainage - function',
    'Meter boxes - access',
  ]);

  // General
  createRoomSheet(workbook, 'General', [
    'Heating system - operation',
    'Hot water - temperature',
    'Water pressure - adequate',
    'Boiler - accessible',
    'Consumer unit - labelled',
    'Certificates provided',
    'Manuals provided',
    'Keys - all provided',
    'Cleaning - standard',
    'Rubbish - removed',
  ]);

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
