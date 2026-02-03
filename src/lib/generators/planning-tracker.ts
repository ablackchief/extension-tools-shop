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

export async function generatePlanningTracker(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Extension Survival Guide';
  workbook.created = new Date();

  // ===== APPLICATION DETAILS SHEET =====
  const appSheet = workbook.addWorksheet('Application Details', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  appSheet.columns = [
    { width: 5 },
    { width: 30 },
    { width: 25 },
    { width: 5 },
    { width: 30 },
    { width: 25 },
  ];

  // Title
  appSheet.mergeCells('B2:F2');
  const titleCell = appSheet.getCell('B2');
  titleCell.value = 'PLANNING APPLICATION TRACKER';
  titleCell.font = { bold: true, size: 18, color: { argb: COLORS.primaryBlue } };

  // Application summary
  appSheet.getCell('B4').value = 'APPLICATION SUMMARY';
  appSheet.getCell('B4').font = { bold: true, size: 12 };

  const appDetails = [
    { label: 'Property Address:', cell: 'C5', merge: 'C5:F5' },
    { label: 'Postcode:', cell: 'C6' },
    { label: 'Local Authority:', cell: 'C7' },
    { label: 'Application Type:', cell: 'C8', dropdown: '"Full Planning,Householder,Prior Approval,Lawful Development Certificate,Listed Building Consent"' },
    { label: '', cell: '' },
    { label: 'Application Reference:', cell: 'C10' },
    { label: 'Submission Date:', cell: 'C11', format: dateFormat },
    { label: 'Validation Date:', cell: 'C12', format: dateFormat },
    { label: 'Target Decision Date:', cell: 'C13', format: dateFormat },
    { label: 'Actual Decision Date:', cell: 'C14', format: dateFormat },
    { label: 'Decision:', cell: 'C15', dropdown: '"Pending,Approved,Approved with Conditions,Refused,Withdrawn"' },
  ];

  appDetails.forEach((detail, index) => {
    const rowNum = index + 5;
    if (detail.label) {
      appSheet.getCell(`B${rowNum}`).value = detail.label;
      appSheet.getCell(`B${rowNum}`).font = { bold: true };

      const inputCell = appSheet.getCell(detail.cell);
      Object.assign(inputCell, { style: inputStyle });
      inputCell.alignment = { horizontal: 'left', vertical: 'middle' };

      if (detail.merge) {
        appSheet.mergeCells(detail.merge);
      }
      if (detail.format) {
        inputCell.numFmt = detail.format;
      }
      if (detail.dropdown) {
        inputCell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: [detail.dropdown],
        };
      }
    }
  });

  // Key dates calculation
  appSheet.getCell('E10').value = 'Statutory Period:';
  appSheet.getCell('E10').font = { bold: true };
  appSheet.getCell('F10').value = { formula: 'IF(C8="Full Planning","13 weeks",IF(C8="Householder","8 weeks","8 weeks"))' };

  appSheet.getCell('E11').value = 'Days Since Submission:';
  appSheet.getCell('E11').font = { bold: true };
  appSheet.getCell('F11').value = { formula: 'IF(C11="","",TODAY()-C11)' };
  appSheet.getCell('F11').numFmt = '0';

  appSheet.getCell('E12').value = 'Days Until Target:';
  appSheet.getCell('E12').font = { bold: true };
  const daysUntilCell = appSheet.getCell('F12');
  daysUntilCell.value = { formula: 'IF(C13="","",C13-TODAY())' };
  daysUntilCell.numFmt = '0';

  // Conditional formatting for days until target
  appSheet.addConditionalFormatting({
    ref: 'F12',
    rules: [
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['7'],
        style: { font: { color: { argb: COLORS.red }, bold: true } },
        priority: 1,
      },
      {
        type: 'cellIs',
        operator: 'between',
        formulae: ['7', '14'],
        style: { font: { color: { argb: COLORS.yellow }, bold: true } },
        priority: 2,
      },
    ],
  });

  // Officer contact
  appSheet.getCell('B18').value = 'CASE OFFICER';
  appSheet.getCell('B18').font = { bold: true, size: 12 };

  const officerDetails = [
    { label: 'Name:', cell: 'C19' },
    { label: 'Phone:', cell: 'C20' },
    { label: 'Email:', cell: 'C21' },
    { label: 'Last Contact:', cell: 'C22', format: dateFormat },
  ];

  officerDetails.forEach((detail, index) => {
    const rowNum = 19 + index;
    appSheet.getCell(`B${rowNum}`).value = detail.label;
    appSheet.getCell(`B${rowNum}`).font = { bold: true };

    const inputCell = appSheet.getCell(detail.cell);
    Object.assign(inputCell, { style: inputStyle });
    inputCell.alignment = { horizontal: 'left', vertical: 'middle' };
    appSheet.mergeCells(`C${rowNum}:D${rowNum}`);

    if (detail.format) {
      inputCell.numFmt = detail.format;
    }
  });

  // Application status
  appSheet.getCell('B25').value = 'APPLICATION STATUS';
  appSheet.getCell('B25').font = { bold: true, size: 12 };

  const statusCell = appSheet.getCell('C25');
  statusCell.value = { formula: 'IF(C15="Pending",IF(F12<0,"OVERDUE",IF(F12<7,"Decision Imminent","In Progress")),C15)' };
  statusCell.font = { bold: true, size: 14 };

  appSheet.addConditionalFormatting({
    ref: 'C25',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'OVERDUE',
        style: { font: { color: { argb: COLORS.red }, bold: true } },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Approved',
        style: { font: { color: { argb: COLORS.green }, bold: true } },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Refused',
        style: { font: { color: { argb: COLORS.red }, bold: true } },
        priority: 3,
      },
    ],
  });

  // ===== TIMELINE SHEET =====
  const timelineSheet = workbook.addWorksheet('Timeline', {
    properties: { tabColor: { argb: COLORS.green } },
  });

  timelineSheet.columns = [
    { width: 5 },
    { width: 12 },
    { width: 35 },
    { width: 15 },
    { width: 12 },
    { width: 30 },
  ];

  // Title
  timelineSheet.mergeCells('B2:F2');
  const timeTitleCell = timelineSheet.getCell('B2');
  timeTitleCell.value = 'APPLICATION TIMELINE';
  timeTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.green } };

  // Headers
  const timeHeaders = ['', 'Date', 'Milestone/Event', 'Status', 'Days', 'Notes'];
  const timeHeaderRow = timelineSheet.getRow(4);
  timeHeaders.forEach((header, index) => {
    const cell = timeHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Pre-populated timeline events
  const timelineEvents = [
    { event: 'Pre-application advice submitted', status: 'Pending' },
    { event: 'Pre-application response received', status: 'Pending' },
    { event: 'Application submitted', status: 'Pending' },
    { event: 'Application validated', status: 'Pending' },
    { event: 'Neighbour consultation period starts', status: 'Pending' },
    { event: 'Neighbour consultation period ends', status: 'Pending' },
    { event: 'Case officer assigned', status: 'Pending' },
    { event: 'Site visit conducted', status: 'Pending' },
    { event: 'Additional information requested', status: 'Pending' },
    { event: 'Additional information submitted', status: 'Pending' },
    { event: 'Committee date (if applicable)', status: 'Pending' },
    { event: 'Decision issued', status: 'Pending' },
    { event: 'Conditions discharged', status: 'Pending' },
  ];

  timelineEvents.forEach((event, index) => {
    const rowNum = index + 5;
    const row = timelineSheet.getRow(rowNum);

    const dateCell = row.getCell(2);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    row.getCell(3).value = event.event;
    applyBorder(row.getCell(3));

    const statusCell = row.getCell(4);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Pending,Complete,N/A"'],
    };
    statusCell.value = event.status;
    Object.assign(statusCell, { style: inputStyle });

    const daysCell = row.getCell(5);
    daysCell.value = { formula: `IF(B${rowNum}="","",IF(B${rowNum + 1}="",TODAY()-B${rowNum},B${rowNum + 1}-B${rowNum}))` };
    daysCell.numFmt = '0';
    applyBorder(daysCell);

    const notesCell = row.getCell(6);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left' };
  });

  // Additional rows for custom events
  for (let i = timelineEvents.length + 5; i <= 30; i++) {
    const row = timelineSheet.getRow(i);

    const dateCell = row.getCell(2);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    const eventCell = row.getCell(3);
    Object.assign(eventCell, { style: inputStyle });
    eventCell.alignment = { horizontal: 'left' };

    const statusCell = row.getCell(4);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Pending,Complete,N/A"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    applyBorder(row.getCell(5));

    const notesCell = row.getCell(6);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left' };
  }

  // Conditional formatting for status
  timelineSheet.addConditionalFormatting({
    ref: 'D5:D30',
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
        text: 'Pending',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
        },
        priority: 2,
      },
    ],
  });

  // ===== DOCUMENT CHECKLIST SHEET =====
  const docSheet = workbook.addWorksheet('Document Checklist', {
    properties: { tabColor: { argb: COLORS.orange } },
  });

  docSheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 12 },
    { width: 15 },
    { width: 30 },
    { width: 20 },
  ];

  // Title
  docSheet.mergeCells('B2:F2');
  const docTitleCell = docSheet.getCell('B2');
  docTitleCell.value = 'DOCUMENT CHECKLIST';
  docTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.orange } };

  // Headers
  const docHeaders = ['', 'Document', 'Required', 'Status', 'Notes', 'File Reference'];
  const docHeaderRow = docSheet.getRow(4);
  docHeaders.forEach((header, index) => {
    const cell = docHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Document list
  const documents = [
    { doc: 'Application Form', required: 'Yes' },
    { doc: 'Planning Fee Payment', required: 'Yes' },
    { doc: 'Site Location Plan (1:1250 or 1:2500)', required: 'Yes' },
    { doc: 'Block Plan (1:500 or 1:200)', required: 'Yes' },
    { doc: 'Existing Floor Plans', required: 'Yes' },
    { doc: 'Proposed Floor Plans', required: 'Yes' },
    { doc: 'Existing Elevations', required: 'Yes' },
    { doc: 'Proposed Elevations', required: 'Yes' },
    { doc: 'Existing & Proposed Sections', required: 'Maybe' },
    { doc: 'Design & Access Statement', required: 'Maybe' },
    { doc: 'Heritage Statement', required: 'Maybe' },
    { doc: 'Tree Survey', required: 'Maybe' },
    { doc: 'Flood Risk Assessment', required: 'Maybe' },
    { doc: 'Ecology Report', required: 'Maybe' },
    { doc: 'Structural Survey', required: 'Maybe' },
    { doc: 'CIL Form', required: 'Yes' },
    { doc: 'Ownership Certificate', required: 'Yes' },
    { doc: 'Agricultural Holdings Certificate', required: 'Yes' },
    { doc: 'Photographs of Site', required: 'No' },
    { doc: 'Pre-application Response', required: 'No' },
  ];

  documents.forEach((doc, index) => {
    const rowNum = index + 5;
    const row = docSheet.getRow(rowNum);

    row.getCell(2).value = doc.doc;
    applyBorder(row.getCell(2));

    const reqCell = row.getCell(3);
    reqCell.value = doc.required;
    reqCell.alignment = { horizontal: 'center' };
    applyBorder(reqCell);

    const statusCell = row.getCell(4);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Not Started,In Progress,Ready,Submitted,N/A"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    const notesCell = row.getCell(5);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left' };

    const refCell = row.getCell(6);
    Object.assign(refCell, { style: inputStyle });
    refCell.alignment = { horizontal: 'left' };
  });

  // Summary
  const summaryRowNum = documents.length + 6;
  docSheet.getCell(`B${summaryRowNum}`).value = 'COMPLETION STATUS';
  docSheet.getCell(`B${summaryRowNum}`).font = { bold: true, size: 12 };

  docSheet.getCell(`B${summaryRowNum + 1}`).value = 'Documents Submitted:';
  const submittedCell = docSheet.getCell(`C${summaryRowNum + 1}`);
  submittedCell.value = { formula: 'COUNTIF(D5:D24,"Submitted")' };
  submittedCell.numFmt = '0';

  docSheet.getCell(`B${summaryRowNum + 2}`).value = 'Documents Required:';
  const requiredCell = docSheet.getCell(`C${summaryRowNum + 2}`);
  requiredCell.value = { formula: 'COUNTIF(C5:C24,"Yes")' };
  requiredCell.numFmt = '0';

  // Conditional formatting for status
  docSheet.addConditionalFormatting({
    ref: 'D5:D24',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Submitted',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'In Progress',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } },
        },
        priority: 2,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Not Started',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
        },
        priority: 3,
      },
    ],
  });

  // Conditional formatting for required column
  docSheet.addConditionalFormatting({
    ref: 'C5:C24',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Yes',
        style: { font: { color: { argb: COLORS.red }, bold: true } },
        priority: 1,
      },
    ],
  });

  // ===== CONTACT LOG SHEET =====
  const contactSheet = workbook.addWorksheet('Contact Log', {
    properties: { tabColor: { argb: COLORS.primaryBlue } },
  });

  contactSheet.columns = [
    { width: 5 },
    { width: 12 },
    { width: 12 },
    { width: 20 },
    { width: 40 },
    { width: 12 },
    { width: 20 },
  ];

  // Title
  contactSheet.mergeCells('B2:G2');
  const contactTitleCell = contactSheet.getCell('B2');
  contactTitleCell.value = 'OFFICER CONTACT LOG';
  contactTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.primaryBlue } };

  // Headers
  const contactHeaders = ['', 'Date', 'Type', 'Person', 'Summary', 'Follow-up', 'Due Date'];
  const contactHeaderRow = contactSheet.getRow(4);
  contactHeaders.forEach((header, index) => {
    const cell = contactHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Contact log rows
  for (let i = 5; i <= 30; i++) {
    const row = contactSheet.getRow(i);

    const dateCell = row.getCell(2);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    const typeCell = row.getCell(3);
    typeCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Phone,Email,Meeting,Portal Message,Letter"'],
    };
    Object.assign(typeCell, { style: inputStyle });

    const personCell = row.getCell(4);
    Object.assign(personCell, { style: inputStyle });
    personCell.alignment = { horizontal: 'left' };

    const summaryCell = row.getCell(5);
    Object.assign(summaryCell, { style: inputStyle });
    summaryCell.alignment = { horizontal: 'left', wrapText: true };

    const followupCell = row.getCell(6);
    followupCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Yes,No"'],
    };
    Object.assign(followupCell, { style: inputStyle });

    const dueCell = row.getCell(7);
    Object.assign(dueCell, { style: inputStyle });
    dueCell.numFmt = dateFormat;
  }

  // Conditional formatting for follow-up required
  contactSheet.addConditionalFormatting({
    ref: 'F5:F30',
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

  // ===== DECISION & CONDITIONS SHEET =====
  const decisionSheet = workbook.addWorksheet('Decision & Conditions', {
    properties: { tabColor: { argb: COLORS.green } },
  });

  decisionSheet.columns = [
    { width: 5 },
    { width: 8 },
    { width: 45 },
    { width: 15 },
    { width: 12 },
    { width: 25 },
  ];

  // Title
  decisionSheet.mergeCells('B2:F2');
  const decTitleCell = decisionSheet.getCell('B2');
  decTitleCell.value = 'DECISION & CONDITIONS';
  decTitleCell.font = { bold: true, size: 16, color: { argb: COLORS.green } };

  // Decision section
  decisionSheet.getCell('B4').value = 'DECISION DETAILS';
  decisionSheet.getCell('B4').font = { bold: true, size: 12 };

  decisionSheet.getCell('B5').value = 'Decision:';
  decisionSheet.getCell('B5').font = { bold: true };
  const decCell = decisionSheet.getCell('C5');
  decCell.value = { formula: '\'Application Details\'!C15' };
  decCell.font = { bold: true, size: 12 };

  decisionSheet.getCell('B6').value = 'Decision Date:';
  decisionSheet.getCell('B6').font = { bold: true };
  decisionSheet.getCell('C6').value = { formula: '\'Application Details\'!C14' };
  decisionSheet.getCell('C6').numFmt = dateFormat;

  decisionSheet.getCell('B7').value = 'Permission Expires:';
  decisionSheet.getCell('B7').font = { bold: true };
  const expiryCell = decisionSheet.getCell('C7');
  Object.assign(expiryCell, { style: inputStyle });
  expiryCell.numFmt = dateFormat;
  decisionSheet.getCell('D7').value = '(Usually 3 years from decision)';
  decisionSheet.getCell('D7').font = { italic: true, color: { argb: 'FF64748b' }, size: 10 };

  // Conditions section
  decisionSheet.getCell('B10').value = 'CONDITIONS TO DISCHARGE';
  decisionSheet.getCell('B10').font = { bold: true, size: 12 };

  const condHeaders = ['', 'No.', 'Condition Description', 'Discharge Status', 'Date', 'Reference/Notes'];
  const condHeaderRow = decisionSheet.getRow(12);
  condHeaders.forEach((header, index) => {
    const cell = condHeaderRow.getCell(index + 1);
    cell.value = header;
    if (index > 0) {
      Object.assign(cell, { style: headerStyle });
    }
  });

  // Condition rows
  for (let i = 13; i <= 27; i++) {
    const row = decisionSheet.getRow(i);

    row.getCell(2).value = i - 12;
    row.getCell(2).alignment = { horizontal: 'center' };
    row.getCell(2).numFmt = '0';
    applyBorder(row.getCell(2));

    const descCell = row.getCell(3);
    Object.assign(descCell, { style: inputStyle });
    descCell.alignment = { horizontal: 'left', wrapText: true };

    const statusCell = row.getCell(4);
    statusCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Pre-commencement,Pre-occupation,Ongoing,Discharged,N/A"'],
    };
    Object.assign(statusCell, { style: inputStyle });

    const dateCell = row.getCell(5);
    Object.assign(dateCell, { style: inputStyle });
    dateCell.numFmt = dateFormat;

    const notesCell = row.getCell(6);
    Object.assign(notesCell, { style: inputStyle });
    notesCell.alignment = { horizontal: 'left' };
  }

  // Summary
  decisionSheet.getCell('B29').value = 'Conditions Discharged:';
  decisionSheet.getCell('B29').font = { bold: true };
  const dischargedCell = decisionSheet.getCell('C29');
  dischargedCell.value = { formula: 'COUNTIF(D13:D27,"Discharged")' };
  dischargedCell.numFmt = '0';

  decisionSheet.getCell('B30').value = 'Conditions Remaining:';
  decisionSheet.getCell('B30').font = { bold: true };
  const remainingCell = decisionSheet.getCell('C30');
  remainingCell.value = { formula: 'COUNTIF(D13:D27,"Pre-commencement")+COUNTIF(D13:D27,"Pre-occupation")+COUNTIF(D13:D27,"Ongoing")' };
  remainingCell.numFmt = '0';

  // Conditional formatting for condition status
  decisionSheet.addConditionalFormatting({
    ref: 'D13:D27',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Discharged',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } },
          font: { color: { argb: COLORS.green } },
        },
        priority: 1,
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Pre-commencement',
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfee2e2' } },
          font: { color: { argb: COLORS.red } },
        },
        priority: 2,
      },
    ],
  });

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
