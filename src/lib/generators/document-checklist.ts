import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  HeadingLevel,
  ShadingType,
  convertInchesToTwip,
  PageBreak,
  Header,
  Footer,
} from 'docx';

// Brand Colors
const COLORS = {
  primaryBlue: '1a56f5',
  headerDark: '1e293b',
  inputOrange: 'fff7ed',
  sectionBlue: '3b82f6',
  white: 'ffffff',
  lightGray: 'f1f5f9',
  mediumGray: '94a3b8',
  greenSuccess: '22c55e',
  greenLight: 'dcfce7',
  amberWarning: 'f59e0b',
  amberLight: 'fef3c7',
  redDanger: 'ef4444',
  redLight: 'fee2e2',
};

// Checkbox character
const CHECKBOX = '\u2610'; // ☐

// Helper to create styled heading
function createHeading(text: string, level: (typeof HeadingLevel)[keyof typeof HeadingLevel] = HeadingLevel.HEADING_1): Paragraph {
  return new Paragraph({
    heading: level,
    spacing: { before: 400, after: 200 },
    children: [
      new TextRun({
        text,
        font: 'Calibri',
        color: COLORS.headerDark,
        bold: true,
        size: level === HeadingLevel.HEADING_1 ? 36 : level === HeadingLevel.HEADING_2 ? 28 : 24,
      }),
    ],
  });
}

// Helper to create body text
function createBodyText(text: string, options: { bold?: boolean; italic?: boolean; color?: string } = {}): Paragraph {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    children: [
      new TextRun({
        text,
        font: 'Calibri',
        size: 22,
        color: options.color || COLORS.headerDark,
        bold: options.bold,
        italics: options.italic,
      }),
    ],
  });
}

// Helper to create section divider
function createSectionDivider(): Paragraph {
  return new Paragraph({
    spacing: { before: 300, after: 300 },
    border: {
      bottom: {
        color: COLORS.sectionBlue,
        size: 24,
        style: BorderStyle.SINGLE,
      },
    },
    children: [],
  });
}

// Helper to create checklist table with status columns
function createChecklistTable(items: { item: string; description?: string; required?: boolean }[]): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 5, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 40, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ children: [new TextRun({ text: 'Document', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ children: [new TextRun({ text: 'Status / Location', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Date Obtained', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
        ],
      }),
      ...items.map(
        (item) =>
          new TableRow({
            children: [
              new TableCell({
                shading: { fill: item.required ? COLORS.amberLight : COLORS.white, type: ShadingType.SOLID, color: item.required ? COLORS.amberLight : COLORS.white },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })],
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: item.item, font: 'Calibri', size: 22, bold: item.required }),
                      item.required ? new TextRun({ text: ' *', font: 'Calibri', size: 22, color: COLORS.redDanger }) : new TextRun({ text: '', font: 'Calibri', size: 22 }),
                    ],
                  }),
                  item.description
                    ? new Paragraph({
                        children: [new TextRun({ text: item.description, font: 'Calibri', size: 18, italics: true, color: COLORS.mediumGray })],
                      })
                    : new Paragraph({ children: [] }),
                ],
              }),
              new TableCell({
                shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
              }),
              new TableCell({
                shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
              }),
            ],
          })
      ),
    ],
  });
}

// Helper to create warranty tracking table
function createWarrantyTable(): Table {
  const warranties = [
    { item: 'Structural Warranty (NHBC/Premier/etc.)', years: '10 years' },
    { item: 'Roofing Guarantee', years: '10-25 years' },
    { item: 'Damp Proofing Guarantee', years: '10-30 years' },
    { item: 'Window/Door Guarantee', years: '10 years' },
    { item: 'Boiler Warranty', years: '5-10 years' },
    { item: 'Kitchen Appliance Warranties', years: '1-5 years' },
    { item: 'Bathroom Sanitary Ware', years: '1-25 years' },
    { item: 'Flooring Warranty', years: '5-25 years' },
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 5, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ children: [new TextRun({ text: 'Warranty Item', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Typical Term', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Start Date', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Expiry Date', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
            children: [new Paragraph({ children: [new TextRun({ text: 'Provider', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
          }),
        ],
      }),
      ...warranties.map(
        (w) =>
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })],
              }),
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: w.item, font: 'Calibri', size: 22 })] })],
              }),
              new TableCell({
                shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: w.years, font: 'Calibri', size: 20, color: COLORS.mediumGray })] })],
              }),
              new TableCell({
                shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
              }),
              new TableCell({
                shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
              }),
              new TableCell({
                shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
              }),
            ],
          })
      ),
    ],
  });
}

// Create title page
function createTitlePage(): Paragraph[] {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 2000 },
      children: [
        new TextRun({
          text: 'EXTENSION SURVIVAL GUIDE',
          font: 'Calibri',
          size: 56,
          bold: true,
          color: COLORS.primaryBlue,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 800 },
      children: [
        new TextRun({
          text: 'Document Checklist',
          font: 'Calibri',
          size: 40,
          color: COLORS.headerDark,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400 },
      children: [
        new TextRun({
          text: 'Track every certificate, warranty, and document for your extension project.',
          font: 'Calibri',
          size: 24,
          italics: true,
          color: COLORS.mediumGray,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 800 },
      children: [
        new TextRun({
          text: 'Essential for selling your home or making insurance claims.',
          font: 'Calibri',
          size: 22,
          color: COLORS.headerDark,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1600 },
      children: [
        new TextRun({
          text: 'Version 1.0  |  www.extensionsurvivalguide.co.uk',
          font: 'Calibri',
          size: 20,
          color: COLORS.mediumGray,
        }),
      ],
    }),
    new Paragraph({
      children: [new PageBreak()],
    }),
  ];
}

// Planning & Approvals section
function createPlanningApprovals(): (Paragraph | Table)[] {
  return [
    createHeading('Planning & Approvals', HeadingLevel.HEADING_1),
    createBodyText('Core permissions and approvals for your extension. Items marked * are essential.'),
    new Paragraph({ spacing: { before: 200 } }),
    createChecklistTable([
      { item: 'Planning Permission Decision Notice', description: 'If full planning was required', required: false },
      { item: 'Lawful Development Certificate', description: 'Confirms PD rights (recommended)', required: false },
      { item: 'Prior Approval Notification', description: 'For larger rear extensions', required: false },
      { item: 'Building Regulations Approval', description: 'Initial approval document', required: true },
      { item: 'Building Control Completion Certificate', description: 'Final sign-off (essential for sale)', required: true },
      { item: 'Party Wall Agreement(s)', description: 'If work affects party wall', required: false },
      { item: 'Party Wall Award', description: 'Surveyor\'s formal decision', required: false },
    ]),
    createSectionDivider(),
  ];
}

// Required Certificates section
function createRequiredCertificates(): (Paragraph | Table)[] {
  return [
    createHeading('Required Certificates', HeadingLevel.HEADING_1),
    createBodyText('Statutory certificates that must be obtained for compliance.'),
    new Paragraph({ spacing: { before: 200 } }),
    createChecklistTable([
      { item: 'Building Control Completion Certificate', description: 'Final building regs sign-off', required: true },
      { item: 'Part P Electrical Certificate (BS 7671)', description: 'For all notifiable electrical work', required: true },
      { item: 'Gas Safe Certificate', description: 'For any gas appliance installation/alteration', required: true },
      { item: 'Unvented Hot Water Certificate', description: 'If unvented cylinder installed (G3)', required: true },
      { item: 'EPC (Energy Performance Certificate)', description: 'May need updating after extension', required: false },
    ]),
    createSectionDivider(),
  ];
}

// Conditional Certificates section
function createConditionalCertificates(): (Paragraph | Table)[] {
  return [
    createHeading('Conditional Certificates', HeadingLevel.HEADING_1),
    createBodyText('Certificates needed depending on your specific project scope.'),
    new Paragraph({ spacing: { before: 200 } }),
    createChecklistTable([
      { item: 'FENSA Certificate', description: 'For replacement windows/doors by FENSA installer', required: false },
      { item: 'CERTASS Certificate', description: 'Alternative to FENSA for glazing', required: false },
      { item: 'Structural Engineer\'s Calculations', description: 'For steelwork, underpinning, etc.', required: false },
      { item: 'Structural Engineer\'s Sign-off', description: 'Confirmation of correct installation', required: false },
      { item: 'Drainage Sign-off', description: 'If connected to main sewer', required: false },
      { item: 'Water Authority Connection Approval', description: 'For new water/sewer connections', required: false },
      { item: 'Air Tightness Test Results', description: 'Required for some extensions', required: false },
      { item: 'SAP Calculations', description: 'Energy assessment for Part L compliance', required: false },
      { item: 'Sound Insulation Test', description: 'If Party Wall/floor constructed', required: false },
    ]),
    createSectionDivider(),
  ];
}

// Warranties & Guarantees section
function createWarrantiesGuarantees(): (Paragraph | Table)[] {
  return [
    createHeading('Warranties & Guarantees', HeadingLevel.HEADING_1),
    createBodyText('Track warranty expiry dates. Essential for future claims and property sale.'),
    new Paragraph({ spacing: { before: 200 } }),
    createWarrantyTable(),
    new Paragraph({ spacing: { before: 200 } }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 8, color: COLORS.amberWarning },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: COLORS.amberWarning },
        left: { style: BorderStyle.SINGLE, size: 8, color: COLORS.amberWarning },
        right: { style: BorderStyle.SINGLE, size: 8, color: COLORS.amberWarning },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.amberLight, type: ShadingType.SOLID, color: COLORS.amberLight },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 100, after: 100 },
                  children: [
                    new TextRun({
                      text: 'TIP: Set calendar reminders 1 month before warranty expiry dates!',
                      font: 'Calibri',
                      bold: true,
                      size: 22,
                      color: COLORS.amberWarning,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    createSectionDivider(),
  ];
}

// Drawings & Specifications section
function createDrawingsSpecifications(): (Paragraph | Table)[] {
  return [
    createHeading('Drawings & Specifications', HeadingLevel.HEADING_1),
    createBodyText('Technical documents showing what was built. Essential for future work.'),
    new Paragraph({ spacing: { before: 200 } }),
    createChecklistTable([
      { item: 'Approved Architectural Drawings', description: 'As-approved by planning/building control', required: true },
      { item: 'As-Built Drawings', description: 'Showing actual construction (if different)', required: false },
      { item: 'Structural Drawings', description: 'Steel/foundation details', required: false },
      { item: 'Electrical Layout Drawings', description: 'Socket/switch/lighting positions', required: false },
      { item: 'Plumbing/Drainage Drawings', description: 'Pipe runs and connections', required: false },
      { item: 'Heating System Schematic', description: 'Boiler, radiator, underfloor heating', required: false },
      { item: 'Specification Document', description: 'Materials and finishes specified', required: false },
      { item: 'Product Data Sheets', description: 'For key materials/products used', required: false },
    ]),
    createSectionDivider(),
  ];
}

// Photographs section
function createPhotographs(): (Paragraph | Table)[] {
  return [
    createHeading('Photographs', HeadingLevel.HEADING_1),
    createBodyText('Visual record of construction. Critical for disputes and insurance claims.'),
    new Paragraph({ spacing: { before: 200 } }),
    createChecklistTable([
      { item: 'Before (Pre-construction)', description: 'Existing property condition', required: false },
      { item: 'Foundations', description: 'Before concrete pour', required: true },
      { item: 'Structural Steelwork', description: 'Beams before covering', required: true },
      { item: 'First Fix - Electrical', description: 'Cables before plastering', required: true },
      { item: 'First Fix - Plumbing', description: 'Pipes before covering', required: true },
      { item: 'Insulation Installation', description: 'Before plasterboard', required: true },
      { item: 'Membrane/DPM Installation', description: 'Damp proofing visible', required: false },
      { item: 'Roof Construction', description: 'Before tiles/covering', required: false },
      { item: 'After (Completion)', description: 'Finished extension', required: false },
    ]),
    new Paragraph({ spacing: { before: 200 } }),
    createBodyText('Store photos with date stamps. Include a tape measure for scale where possible.', { italic: true, color: COLORS.mediumGray }),
    createSectionDivider(),
  ];
}

// Financial Records section
function createFinancialRecords(): (Paragraph | Table)[] {
  return [
    createHeading('Financial Records', HeadingLevel.HEADING_1),
    createBodyText('Payment evidence and contracts. Essential for disputes and capital gains calculations.'),
    new Paragraph({ spacing: { before: 200 } }),
    createChecklistTable([
      { item: 'Building Contract (signed)', description: 'Main contractor agreement', required: true },
      { item: 'Architect/Designer Contract', description: 'Design services agreement', required: false },
      { item: 'All Paid Invoices', description: 'With proof of payment', required: true },
      { item: 'Payment Schedule Record', description: 'Stage payments vs. completion', required: true },
      { item: 'Variation Orders (signed)', description: 'Any approved changes', required: false },
      { item: 'Final Account Statement', description: 'Total project cost breakdown', required: true },
      { item: 'VAT Receipts', description: 'For all materials and labour', required: false },
      { item: 'Professional Fee Invoices', description: 'Architect, SE, surveyors', required: false },
    ]),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// Sale-Ready Pack section
function createSaleReadyPack(): (Paragraph | Table)[] {
  return [
    createHeading('Sale-Ready Pack', HeadingLevel.HEADING_1),
    createBodyText('When selling your home, solicitors will request these documents. Missing items can delay or derail sales.'),
    new Paragraph({ spacing: { before: 200 } }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            new TableCell({
              width: { size: 5, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.greenSuccess, type: ShadingType.SOLID, color: COLORS.greenSuccess },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 45, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.greenSuccess, type: ShadingType.SOLID, color: COLORS.greenSuccess },
              children: [new Paragraph({ children: [new TextRun({ text: 'Essential Document', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.greenSuccess, type: ShadingType.SOLID, color: COLORS.greenSuccess },
              children: [new Paragraph({ children: [new TextRun({ text: 'Why Required', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        ...[
          { doc: 'Building Regulations Completion Certificate', why: 'Proves work was inspected and approved' },
          { doc: 'Planning Permission (if applicable)', why: 'Proves legal consent obtained' },
          { doc: 'FENSA/CERTASS Certificate', why: 'For any new/replacement windows' },
          { doc: 'Part P Electrical Certificate', why: 'Proves electrical work is safe and compliant' },
          { doc: 'Gas Safe Certificate', why: 'Proves gas work done by registered engineer' },
          { doc: 'Party Wall Agreement/Award', why: 'Proves neighbour consent obtained' },
          { doc: 'Structural Engineer Sign-off', why: 'Proves steelwork correctly installed' },
          { doc: 'Approved Drawings', why: 'Shows what was approved to be built' },
          { doc: 'Warranties (transferable)', why: 'Provides ongoing protection for buyer' },
          { doc: 'Indemnity Insurance', why: 'May be needed if certificates missing' },
        ].map(
          (item) =>
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: item.doc, font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
                  children: [new Paragraph({ children: [new TextRun({ text: item.why, font: 'Calibri', size: 20, color: COLORS.mediumGray })] })],
                }),
              ],
            })
        ),
      ],
    }),
    createSectionDivider(),
  ];
}

// Missing Documents Action Plan section
function createMissingDocumentsActionPlan(): (Paragraph | Table)[] {
  return [
    createHeading('Missing Documents Action Plan', HeadingLevel.HEADING_1),
    createBodyText('If you\'re missing essential documents, here\'s how to obtain them:'),
    new Paragraph({ spacing: { before: 200 } }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Missing Document', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 40, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'How to Obtain', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Alternative', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        ...[
          { doc: 'Building Control Completion', how: 'Contact Local Authority Building Control', alt: 'Regularisation application (retrospective)' },
          { doc: 'Planning Permission', how: 'Contact Local Planning Authority', alt: 'Certificate of Lawfulness (if PD)' },
          { doc: 'Part P Certificate', how: 'Contact NICEIC/NAPIT/ELECSA', alt: 'New inspection by registered electrician' },
          { doc: 'Gas Safe Certificate', how: 'Contact Gas Safe Register', alt: 'New inspection by Gas Safe engineer' },
          { doc: 'FENSA Certificate', how: 'Contact FENSA directly', alt: 'Building Control inspection (£££)' },
          { doc: 'Any Certificate', how: 'Search for original installer', alt: 'Indemnity insurance policy' },
        ].map(
          (item) =>
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: item.doc, font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: item.how, font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  shading: { fill: COLORS.amberLight, type: ShadingType.SOLID, color: COLORS.amberLight },
                  children: [new Paragraph({ children: [new TextRun({ text: item.alt, font: 'Calibri', size: 20 })] })],
                }),
              ],
            })
        ),
      ],
    }),
    new Paragraph({ spacing: { before: 400 } }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 8, color: COLORS.redDanger },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: COLORS.redDanger },
        left: { style: BorderStyle.SINGLE, size: 8, color: COLORS.redDanger },
        right: { style: BorderStyle.SINGLE, size: 8, color: COLORS.redDanger },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.redLight, type: ShadingType.SOLID, color: COLORS.redLight },
              children: [
                new Paragraph({
                  spacing: { before: 150, after: 150 },
                  children: [
                    new TextRun({
                      text: 'WARNING: ',
                      font: 'Calibri',
                      bold: true,
                      size: 22,
                      color: COLORS.redDanger,
                    }),
                    new TextRun({
                      text: 'Indemnity insurance does not fix the underlying issue - it only covers legal costs if problems arise. It may not satisfy all buyers or lenders. Always try to obtain the actual certificate first.',
                      font: 'Calibri',
                      size: 22,
                      color: COLORS.headerDark,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    new Paragraph({ spacing: { before: 400 } }),
    createHeading('My Action Items', HeadingLevel.HEADING_2),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            new TableCell({
              width: { size: 5, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 35, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Missing Document', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 35, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Action Required', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Deadline', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              width: { size: 10, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Done', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        ...[1, 2, 3, 4, 5].map(
          () =>
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                  children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                  children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                  children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })],
                }),
              ],
            })
        ),
      ],
    }),
  ];
}

export async function generateDocumentChecklist(): Promise<Buffer> {
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: 'Calibri',
            size: 22,
          },
        },
      },
    },
    sections: [
      {
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: 'Extension Survival Guide  |  Document Checklist',
                    font: 'Calibri',
                    size: 18,
                    color: COLORS.mediumGray,
                  }),
                ],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: 'www.extensionsurvivalguide.co.uk',
                    font: 'Calibri',
                    size: 18,
                    color: COLORS.mediumGray,
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          ...createTitlePage(),
          ...createPlanningApprovals(),
          ...createRequiredCertificates(),
          ...createConditionalCertificates(),
          ...createWarrantiesGuarantees(),
          ...createDrawingsSpecifications(),
          ...createPhotographs(),
          ...createFinancialRecords(),
          ...createSaleReadyPack(),
          ...createMissingDocumentsActionPlan(),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
}
