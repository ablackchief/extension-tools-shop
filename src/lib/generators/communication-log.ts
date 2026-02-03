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
const timeFormat = 'hh:mm';

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    left: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
    right: { style: 'thin', color: { argb: 'FFe2e8f0' } },
  };
}

export async function generateCommunicationLog(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Extension Survival Guide';
  workbook.created = new Date();

  // ===== COMMUNICATION LOG SHEET =====
  const logSheet = workbook.addWorksheet('Communication Log', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  logSheet.columns = [
    { width: 5 },
    { width: 8 },
    { width: 12 },
    { width: 8 },
    { width: 12 },
    { width: 20 },
    { width: 15 },
    { width: 45 },
    { width: 12 },
    { width: 12 },
    { width: 25 },
  ];

  // Title
  logSheet.mergeCells('B2:K2');
  const titleCell = logSheet.getCell('B2');
  titleCell.value = 'PROJECT COMMUNICATION LOG';
  titleCell.font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };

  // Project info
  logSheet.getCell('B4').value = 'Project:';
  logSheet.getCell('B4').font = { bold: true };
  const projectCell = logSheet.getCell('C4');
  Object.assign(projectCell, { style: inputStyle });
  logSheet.mergeCells('C4:E4');

  logSheet.getCell('G4').value = 'Last Updated:';
  logSheet.getCell('G4').font = { bold: true };
  logSheet.getCell('H4').value = { formula: 'TODAY()' };
  logSheet.getCell('H4').numFmt = dateFormat;

  // Instructions
  logSheet.getCell('B6').value = 'Use this log to record all project communications. Filter by column headers to find specific entries.';
  logSheet.getCell('B6').font = { italic: true, color: { argb: 'FF64748b' } };

  // Headers
  const headers = ['', 'Ref', 'Date', 'Time', 'Type', 'Person', 'Company/Role', 'Summary', 'Follow-up', 'Due Date', 'Status'];
  const headerRow = logSheet.getRow(8);
  headers.forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Enable auto-filter
  logSheet.autoFilter = {
    from: { row: 8, column: 2 },
    to: { row: 208, column: 11 },
  };

  // Data rows
  for (let i = 9; i <= 208; i++) {
    const row = logSheet.getRow(i);
    const refNum = i - 8;

    // Reference number
    const refCell = row.getCell(2);
    refCell.value = refNum;
    refCell.alignment = { horizontal: 'center' };
    applyBorder(refCell);

    // Date
    const dateCell = row.getCell(3);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    // Time
    const timeCell = row.getCell(4);
    Object.assign(timeCell, { style: inputStyle });
    timeCell.numFmt = timeFormat;

    // Type dropdown
    const typeCell = row.getCell(5);
    typeCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Email,Phone,Meeting,Site Visit,Video Call,WhatsApp,Letter,Other"'],
    };
    Object.assign(typeCell, { style: inputStyle });

    // Person
    const personCell = row.getCell(6);
    Object.assign(personCell, { style: inputStyle });
    personCell.alignment = { horizontal: 'left', vertical: 'middle' };

    // Company/Role
    const companyCell = row.getCell(7);
    companyCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Builder,Architect,Structural Engineer,Planning Officer,Building Control,Electrician,Plumber,Supplier,Neighbour,Other"'],
    };
    Object.assign(companyCell, { style: inputStyle });

    // Summary
    const summaryCell = row.getCell(8);
    Object.assign(summaryCell, { style: inputStyle });
    summaryCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // Follow-up required
    const followupCell = row.getCell(9);
    followupCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Yes,No"'],
    };
    Object.assign(followupCell, { style: inputStyle });

    // Due Date
    const dueCell = row.getCell(10);
    Object.assign(dueCell, { style: inputStyle });
    dueCell.numFmt = dateFormat;

    // Status
    const statusCell = row.getCell(11);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Open,In Progress,Complete,Cancelled"'],
    };
    Object.assign(statusCell, { style: inputStyle });
  }

  // Conditional formatting for follow-up required
  logSheet.addConditionalFormatting({
    ref: 'I9:I208',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Yes',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
          font: { bold: true },
        },
        priority: 1,
      },
    ],
  });

  // Conditional formatting for status
  logSheet.addConditionalFormatting({
    ref: 'K9:K208',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Complete',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Open',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red } },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'In Progress',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
        },
        priority: 3,
      },
    ],
  });

  // Conditional formatting for overdue items
  logSheet.addConditionalFormatting({
    ref: 'J9:J208',
    rules: [
      {
        type: 'expression',
        formulae: ['AND(J9<TODAY(),J9<>"",K9<>"Complete")'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
    ],
  });

  // ===== SUMMARY SHEET =====
  const summarySheet = workbook.addWorksheet('Summary', {
    properties: { tabColor: { argb: COLORS.green } },
  });

  summarySheet.columns = [
    { width: 5 },
    { width: 25 },
    { width: 15 },
    { width: 5 },
    { width: 25 },
    { width: 15 },
  ];

  // Title
  summarySheet.mergeCells('B2:F2');
  const sumTitleCell = summarySheet.getCell('B2');
  sumTitleCell.value = 'COMMUNICATION SUMMARY';
  sumTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.green } };

  // Statistics section
  summarySheet.getCell('B4').value = 'OVERVIEW';
  summarySheet.getCell('B4').font = { bold: true, size: 12 };

  const stats = [
    { label: 'Total Communications', formula: 'COUNTA(\'Communication Log\'!C9:C208)' },
    { label: 'Open Items', formula: 'COUNTIF(\'Communication Log\'!K9:K208,"Open")' },
    { label: 'In Progress', formula: 'COUNTIF(\'Communication Log\'!K9:K208,"In Progress")' },
    { label: 'Complete', formula: 'COUNTIF(\'Communication Log\'!K9:K208,"Complete")' },
    { label: 'Follow-ups Required', formula: 'COUNTIF(\'Communication Log\'!I9:I208,"Yes")' },
    { label: 'Overdue Items', formula: 'SUMPRODUCT((\'Communication Log\'!J9:J208<TODAY())*(\'Communication Log\'!J9:J208<>"")*(\'Communication Log\'!K9:K208<>"Complete")*1)' },
  ];

  stats.forEach((stat, index) => {
    const rowNum = 6 + index;
    summarySheet.getCell(`B${rowNum}`).value = stat.label;
    summarySheet.getCell(`B${rowNum}`).font = { bold: true };

    const valueCell = summarySheet.getCell(`C${rowNum}`);
    valueCell.value = { formula: stat.formula };
    valueCell.font = { size: 14 };
    applyBorder(valueCell);

    if (stat.label === 'Overdue Items') {
      valueCell.font = { size: 14, color: { argb: COLORS.red }, bold: true };
    }
  });

  // By Type section
  summarySheet.getCell('B14').value = 'BY COMMUNICATION TYPE';
  summarySheet.getCell('B14').font = { bold: true, size: 12 };

  const types = ['Email', 'Phone', 'Meeting', 'Site Visit', 'Video Call', 'WhatsApp', 'Letter', 'Other'];
  types.forEach((type, index) => {
    const rowNum = 16 + index;
    summarySheet.getCell(`B${rowNum}`).value = type;
    const countCell = summarySheet.getCell(`C${rowNum}`);
    countCell.value = { formula: `COUNTIF('Communication Log'!E9:E208,"${type}")` };
    applyBorder(countCell);
  });

  // By Person/Role section
  summarySheet.getCell('E4').value = 'BY ROLE';
  summarySheet.getCell('E4').font = { bold: true, size: 12 };

  const roles = ['Builder', 'Architect', 'Structural Engineer', 'Planning Officer', 'Building Control', 'Electrician', 'Plumber', 'Supplier', 'Neighbour', 'Other'];
  roles.forEach((role, index) => {
    const rowNum = 6 + index;
    summarySheet.getCell(`E${rowNum}`).value = role;
    const countCell = summarySheet.getCell(`F${rowNum}`);
    countCell.value = { formula: `COUNTIF('Communication Log'!G9:G208,"${role}")` };
    applyBorder(countCell);
  });

  // Recent activity
  summarySheet.getCell('E18').value = 'RECENT ACTIVITY';
  summarySheet.getCell('E18').font = { bold: true, size: 12 };

  summarySheet.getCell('E19').value = 'Last 7 days:';
  summarySheet.getCell('F19').value = { formula: 'SUMPRODUCT((\'Communication Log\'!C9:C208>=TODAY()-7)*(\'Communication Log\'!C9:C208<>""))' };
  applyBorder(summarySheet.getCell('F19'));

  summarySheet.getCell('E20').value = 'Last 30 days:';
  summarySheet.getCell('F20').value = { formula: 'SUMPRODUCT((\'Communication Log\'!C9:C208>=TODAY()-30)*(\'Communication Log\'!C9:C208<>""))' };
  applyBorder(summarySheet.getCell('F20'));

  // ===== CONTACTS DIRECTORY SHEET =====
  const contactsSheet = workbook.addWorksheet('Contacts Directory', {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  contactsSheet.columns = [
    { width: 5 },
    { width: 25 },
    { width: 20 },
    { width: 18 },
    { width: 25 },
    { width: 15 },
    { width: 25 },
  ];

  // Title
  contactsSheet.mergeCells('B2:G2');
  const contTitleCell = contactsSheet.getCell('B2');
  contTitleCell.value = 'PROJECT CONTACTS DIRECTORY';
  contTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.orange } };

  // Headers
  const contHeaders = ['', 'Name', 'Company', 'Phone', 'Email', 'Role', 'Notes'];
  const contHeaderRow = contactsSheet.getRow(4);
  contHeaders.forEach((header, index) => {
    const cell = contHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Pre-populated roles
  const prePopulatedContacts = [
    { role: 'Main Contractor/Builder' },
    { role: 'Site Manager' },
    { role: 'Architect' },
    { role: 'Structural Engineer' },
    { role: 'Planning Officer' },
    { role: 'Building Control' },
    { role: 'Electrician' },
    { role: 'Plumber' },
    { role: 'Gas Engineer' },
    { role: 'Roofer' },
    { role: 'Window Supplier' },
    { role: 'Kitchen Supplier' },
    { role: 'Bathroom Supplier' },
    { role: 'Party Wall Surveyor' },
  ];

  prePopulatedContacts.forEach((contact, index) => {
    const rowNum = index + 5;
    const row = contactsSheet.getRow(rowNum);

    const nameCell = row.getCell(2);
    Object.assign(nameCell, { style: inputStyle });
    nameCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const companyCell = row.getCell(3);
    Object.assign(companyCell, { style: inputStyle });
    companyCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const phoneCell = row.getCell(4);
    Object.assign(phoneCell, { style: inputStyle });
    phoneCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const emailCell = row.getCell(5);
    Object.assign(emailCell, { style: inputStyle });
    emailCell.alignment = { horizontal: 'left', vertical: 'middle' };

    row.getCell(6).value = contact.role;
    applyBorder(row.getCell(6));

    const notesCell = row.getCell(7);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left', vertical: 'middle' };
  });

  // Additional blank rows
  for (let i = prePopulatedContacts.length + 5; i <= 30; i++) {
    const row = contactsSheet.getRow(i);

    const nameCell = row.getCell(2);
    Object.assign(nameCell, { style: inputStyle });
    nameCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const companyCell = row.getCell(3);
    Object.assign(companyCell, { style: inputStyle });
    companyCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const phoneCell = row.getCell(4);
    Object.assign(phoneCell, { style: inputStyle });
    phoneCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const emailCell = row.getCell(5);
    Object.assign(emailCell, { style: inputStyle });
    emailCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const roleCell = row.getCell(6);
    Object.assign(roleCell, { style: inputStyle });
    roleCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const notesCell = row.getCell(7);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left', vertical: 'middle' };
  }

  // ===== ACTION ITEMS SHEET =====
  const actionsSheet = workbook.addWorksheet('Action Items', {
    properties: { tabColor: { argb: COLORS.red } },
  });

  actionsSheet.columns = [
    { width: 5 },
    { width: 12 },
    { width: 35 },
    { width: 20 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 25 },
  ];

  // Title
  actionsSheet.mergeCells('B2:H2');
  const actTitleCell = actionsSheet.getCell('B2');
  actTitleCell.value = 'ACTION ITEMS & FOLLOW-UPS';
  actTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.red } };

  actionsSheet.getCell('B4').value = 'Track follow-up actions from communications';
  actionsSheet.getCell('B4').font = { italic: true, color: { argb: 'FF64748b' } };

  // Headers
  const actHeaders = ['', 'Date Created', 'Action Required', 'Assigned To', 'Priority', 'Due Date', 'Status', 'Notes'];
  const actHeaderRow = actionsSheet.getRow(6);
  actHeaders.forEach((header, index) => {
    const cell = actHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Enable auto-filter
  actionsSheet.autoFilter = {
    from: { row: 6, column: 2 },
    to: { row: 56, column: 8 },
  };

  // Action rows
  for (let i = 7; i <= 56; i++) {
    const row = actionsSheet.getRow(i);

    const dateCell = row.getCell(2);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    const actionCell = row.getCell(3);
    Object.assign(actionCell, { style: inputStyle });
    actionCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    const assignedCell = row.getCell(4);
    Object.assign(assignedCell, { style: inputStyle });
    assignedCell.alignment = { horizontal: 'left', vertical: 'middle' };

    const priorityCell = row.getCell(5);
    priorityCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"High,Medium,Low"'],
    };
    Object.assign(priorityCell, { style: inputStyle });

    const dueCell = row.getCell(6);
    Object.assign(dueCell, { style: inputStyle });
    dueCell.numFmt = dateFormat;

    const statusCell = row.getCell(7);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Not Started,In Progress,Complete,On Hold"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    const notesCell = row.getCell(8);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left', vertical: 'middle' };
  }

  // Conditional formatting for priority
  actionsSheet.addConditionalFormatting({
    ref: 'E7:E56',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'High',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Medium',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Low',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
        },
        priority: 3,
      },
    ],
  });

  // Conditional formatting for status
  actionsSheet.addConditionalFormatting({
    ref: 'G7:G56',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Complete',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Not Started',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
        },
        priority: 2,
      },
    ],
  });

  // Conditional formatting for overdue
  actionsSheet.addConditionalFormatting({
    ref: 'F7:F56',
    rules: [
      {
        type: 'expression',
        formulae: ['AND(F7<TODAY(),F7<>"",G7<>"Complete")'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red }, bold: true },
        },
        priority: 1,
      },
    ],
  });

  // Summary at bottom
  actionsSheet.getCell('B58').value = 'SUMMARY';
  actionsSheet.getCell('B58').font = { bold: true, size: 12 };

  actionsSheet.getCell('B59').value = 'Open Actions:';
  actionsSheet.getCell('C59').value = { formula: 'COUNTIF(G7:G56,"Not Started")+COUNTIF(G7:G56,"In Progress")' };

  actionsSheet.getCell('B60').value = 'Overdue:';
  actionsSheet.getCell('C60').value = { formula: 'SUMPRODUCT((F7:F56<TODAY())*(F7:F56<>"")*(G7:G56<>"Complete")*1)' };
  actionsSheet.getCell('C60').font = { color: { argb: COLORS.red }, bold: true };

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
