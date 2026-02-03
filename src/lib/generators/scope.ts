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
};

// Checkbox character
const CHECKBOX = '\u2610'; // ☐
const CHECKBOX_CHECKED = '\u2611'; // ☑

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

// Helper to create input field with orange background
function createInputField(label: string, width: number = 100): Paragraph {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    children: [
      new TextRun({
        text: `${label}: `,
        font: 'Calibri',
        size: 22,
        bold: true,
        color: COLORS.headerDark,
      }),
      new TextRun({
        text: '                                                            ',
        font: 'Calibri',
        size: 22,
        shading: { type: ShadingType.SOLID, fill: COLORS.inputOrange, color: COLORS.inputOrange },
      }),
    ],
  });
}

// Helper to create checkbox item
function createCheckboxItem(text: string, checked: boolean = false): Paragraph {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    indent: { left: convertInchesToTwip(0.25) },
    children: [
      new TextRun({
        text: `${checked ? CHECKBOX_CHECKED : CHECKBOX}  ${text}`,
        font: 'Calibri',
        size: 22,
        color: COLORS.headerDark,
      }),
    ],
  });
}

// Helper to create numbered item
function createNumberedItem(number: number, text: string = ''): Paragraph {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    indent: { left: convertInchesToTwip(0.25) },
    children: [
      new TextRun({
        text: `${number}. `,
        font: 'Calibri',
        size: 22,
        bold: true,
        color: COLORS.headerDark,
      }),
      new TextRun({
        text: text || '                                                                                    ',
        font: 'Calibri',
        size: 22,
        color: COLORS.headerDark,
        shading: text ? undefined : { type: ShadingType.SOLID, fill: COLORS.inputOrange, color: COLORS.inputOrange },
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
          text: 'Project Brief & Scope Template',
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
          text: 'Define your requirements. Lock down your scope. Control your project.',
          font: 'Calibri',
          size: 24,
          italics: true,
          color: COLORS.mediumGray,
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

// Create project details section
function createProjectDetails(): (Paragraph | Table)[] {
  return [
    createHeading('Project Details', HeadingLevel.HEADING_1),
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
          children: [
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
              children: [new Paragraph({ children: [new TextRun({ text: 'Project Address', font: 'Calibri', bold: true, size: 22 })] })],
            }),
            new TableCell({
              width: { size: 70, type: WidthType.PERCENTAGE },
              shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
              children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
              children: [new Paragraph({ children: [new TextRun({ text: 'Property Owner(s)', font: 'Calibri', bold: true, size: 22 })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
              children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
              children: [new Paragraph({ children: [new TextRun({ text: 'Project Start Date', font: 'Calibri', bold: true, size: 22 })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
              children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
              children: [new Paragraph({ children: [new TextRun({ text: 'Target Completion', font: 'Calibri', bold: true, size: 22 })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
              children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
              children: [new Paragraph({ children: [new TextRun({ text: 'Budget Range', font: 'Calibri', bold: true, size: 22 })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
              children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
            }),
          ],
        }),
      ],
    }),
    createSectionDivider(),
  ];
}

// Create project overview section
function createProjectOverview(): Paragraph[] {
  return [
    createHeading('Project Overview', HeadingLevel.HEADING_1),
    createBodyText('Select all extension types that apply to your project:'),
    createHeading('Extension Type', HeadingLevel.HEADING_3),
    createCheckboxItem('Single-storey rear extension'),
    createCheckboxItem('Single-storey side extension'),
    createCheckboxItem('Single-storey wrap-around'),
    createCheckboxItem('Two-storey rear extension'),
    createCheckboxItem('Two-storey side extension'),
    createCheckboxItem('Loft conversion (dormer)'),
    createCheckboxItem('Loft conversion (hip-to-gable)'),
    createCheckboxItem('Loft conversion (mansard)'),
    createCheckboxItem('Basement conversion'),
    createCheckboxItem('Garage conversion'),
    createCheckboxItem('Outbuilding/garden room'),
    createCheckboxItem('Other: ___________________'),
    createSectionDivider(),
  ];
}

// Create room requirements section
function createRoomRequirements(): (Paragraph | Table)[] {
  return [
    createHeading('Room Requirements', HeadingLevel.HEADING_1),
    createBodyText('Define specifications for each new or modified room:'),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.sectionBlue },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.sectionBlue },
        left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.sectionBlue },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.sectionBlue },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Category', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Specification', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Priority', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        ...['Dimensions (min)', 'Natural Light', 'Flooring Type', 'Heating', 'Electrical Outlets', 'Storage', 'Access/Doors', 'Special Features'].map(
          (category) =>
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
                  children: [new Paragraph({ children: [new TextRun({ text: category, font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                  children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })],
                }),
                new TableCell({
                  shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX + ' Must  ' + CHECKBOX + ' Nice', font: 'Calibri', size: 20 })] })],
                }),
              ],
            })
        ),
      ],
    }),
    createBodyText('(Copy this table for each room in your extension)', { italic: true, color: COLORS.mediumGray }),
    createSectionDivider(),
  ];
}

// Create must have list
function createMustHaveList(): Paragraph[] {
  return [
    createHeading('Must Have List', HeadingLevel.HEADING_1),
    createBodyText('These are non-negotiable requirements. The project is a failure without them.'),
    createNumberedItem(1),
    createNumberedItem(2),
    createNumberedItem(3),
    createNumberedItem(4),
    createNumberedItem(5),
    createSectionDivider(),
  ];
}

// Create nice to have list
function createNiceToHaveList(): Paragraph[] {
  return [
    createHeading('Nice to Have List', HeadingLevel.HEADING_1),
    createBodyText('Desirable features, but the project succeeds without them. These are first to be cut if budget is tight.'),
    createNumberedItem(1),
    createNumberedItem(2),
    createNumberedItem(3),
    createNumberedItem(4),
    createNumberedItem(5),
    createSectionDivider(),
  ];
}

// Create explicitly excluded section
function createExcludedSection(): Paragraph[] {
  return [
    createHeading('Explicitly Excluded', HeadingLevel.HEADING_1),
    createBodyText('These items are NOT part of this project scope. Document them to prevent scope creep.'),
    createCheckboxItem('Landscaping/garden work'),
    createCheckboxItem('Interior decoration (beyond basic)'),
    createCheckboxItem('Furniture and fixtures'),
    createCheckboxItem('Smart home systems'),
    createCheckboxItem('Solar panels/renewables'),
    createCheckboxItem('Driveway work'),
    createCheckboxItem('Boundary walls/fencing'),
    createInputField('Other exclusions'),
    createInputField(''),
    createSectionDivider(),
  ];
}

// Create specification standards section
function createSpecificationStandards(): (Paragraph | Table)[] {
  return [
    createHeading('Specification Standards', HeadingLevel.HEADING_1),
    createBodyText('Define the quality level for key elements:'),
    createHeading('Kitchen Specifications', HeadingLevel.HEADING_2),
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
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Units', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Worktop', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Appliances', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Sink/Tap', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
      ],
    }),
    createHeading('Bathroom Specifications', HeadingLevel.HEADING_2),
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
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Sanitaryware', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Shower/Bath', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Tiles', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Heated Towel Rail', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
      ],
    }),
    createHeading('Electrical Specifications', HeadingLevel.HEADING_2),
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
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Sockets per Room', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'USB Sockets', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Lighting Type', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'External Lighting', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray }, children: [new Paragraph({ children: [new TextRun({ text: 'Data/Network Points', font: 'Calibri', bold: true, size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ children: [new TextRun({ text: '', font: 'Calibri', size: 22 })] })] }),
          ],
        }),
      ],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// Create scope lock agreement section
function createScopeLockAgreement(): Paragraph[] {
  return [
    createHeading('Scope Lock Agreement', HeadingLevel.HEADING_1),
    createBodyText(
      'This document defines the agreed scope for the extension project. Any changes to items listed as "Must Have" or specifications marked as agreed require formal variation approval as per the Variation Control Process below.'
    ),
    new Paragraph({ spacing: { before: 400 } }),
    createBodyText('By signing below, all parties agree that:', { bold: true }),
    createNumberedItem(1, 'The requirements documented above accurately reflect project expectations'),
    createNumberedItem(2, 'Items in "Must Have" list are non-negotiable without written agreement'),
    createNumberedItem(3, 'Items in "Nice to Have" may be adjusted based on budget/progress'),
    createNumberedItem(4, 'Items in "Explicitly Excluded" will not be added without formal variation'),
    createNumberedItem(5, 'Changes after this date require the Variation Control Process'),
    new Paragraph({ spacing: { before: 600 } }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Property Owner 1:', font: 'Calibri', bold: true, size: 22 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 200 },
      children: [
        new TextRun({ text: 'Signature: ________________________________    Date: ________________', font: 'Calibri', size: 22 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 200 },
      children: [
        new TextRun({ text: 'Print Name: ', font: 'Calibri', size: 22 }),
        new TextRun({ text: '                                        ', font: 'Calibri', size: 22, shading: { type: ShadingType.SOLID, fill: COLORS.inputOrange, color: COLORS.inputOrange } }),
      ],
    }),
    new Paragraph({ spacing: { before: 400 } }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Property Owner 2 (if applicable):', font: 'Calibri', bold: true, size: 22 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 200 },
      children: [
        new TextRun({ text: 'Signature: ________________________________    Date: ________________', font: 'Calibri', size: 22 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 200 },
      children: [
        new TextRun({ text: 'Print Name: ', font: 'Calibri', size: 22 }),
        new TextRun({ text: '                                        ', font: 'Calibri', size: 22, shading: { type: ShadingType.SOLID, fill: COLORS.inputOrange, color: COLORS.inputOrange } }),
      ],
    }),
    createSectionDivider(),
  ];
}

// Create variation control process section
function createVariationControlProcess(): Paragraph[] {
  return [
    createHeading('Variation Control Process', HeadingLevel.HEADING_1),
    createBodyText('When a change to scope is required, follow this process:'),
    new Paragraph({ spacing: { before: 200 } }),
    createHeading('Step 1: Document the Change', HeadingLevel.HEADING_3),
    createBodyText('Complete a Variation Request Form (see Variation Register) including:'),
    createCheckboxItem('Description of the change'),
    createCheckboxItem('Reason for the change'),
    createCheckboxItem('Impact on budget (quote required)'),
    createCheckboxItem('Impact on timeline'),
    new Paragraph({ spacing: { before: 200 } }),
    createHeading('Step 2: Review & Approval', HeadingLevel.HEADING_3),
    createBodyText('Both parties must review and sign off on the variation before work begins.'),
    new Paragraph({ spacing: { before: 200 } }),
    createHeading('Step 3: Update Project Documents', HeadingLevel.HEADING_3),
    createBodyText('Update the following documents:'),
    createCheckboxItem('Budget Planner'),
    createCheckboxItem('Payment Schedule'),
    createCheckboxItem('This Scope Document (dated addendum)'),
    new Paragraph({ spacing: { before: 200 } }),
    createHeading('Step 4: Communicate', HeadingLevel.HEADING_3),
    createBodyText('Ensure all parties (builder, architect, building control) receive updated documentation.'),
    new Paragraph({ spacing: { before: 400 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
      children: [
        new TextRun({
          text: 'TIP: No verbal agreements. Every change must be documented and signed.',
          font: 'Calibri',
          size: 22,
          bold: true,
          color: COLORS.primaryBlue,
        }),
      ],
    }),
  ];
}

export async function generateScope(): Promise<Buffer> {
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
                    text: 'Extension Survival Guide  |  Project Brief & Scope',
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
          ...createProjectDetails(),
          ...createProjectOverview(),
          ...createRoomRequirements(),
          ...createMustHaveList(),
          ...createNiceToHaveList(),
          ...createExcludedSection(),
          ...createSpecificationStandards(),
          ...createScopeLockAgreement(),
          ...createVariationControlProcess(),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
}
