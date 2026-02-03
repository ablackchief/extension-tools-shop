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
  VerticalAlign,
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
function createBodyText(text: string, options: { bold?: boolean; italic?: boolean; color?: string; center?: boolean } = {}): Paragraph {
  return new Paragraph({
    alignment: options.center ? AlignmentType.CENTER : AlignmentType.LEFT,
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

// Helper to create decision box
function createDecisionBox(question: string, yesResult: string, noResult: string): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 12, color: COLORS.sectionBlue },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.sectionBlue },
      left: { style: BorderStyle.SINGLE, size: 12, color: COLORS.sectionBlue },
      right: { style: BorderStyle.SINGLE, size: 12, color: COLORS.sectionBlue },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.lightGray },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 2,
            shading: { fill: COLORS.sectionBlue, type: ShadingType.SOLID, color: COLORS.sectionBlue },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 100, after: 100 },
                children: [
                  new TextRun({
                    text: question,
                    font: 'Calibri',
                    bold: true,
                    size: 24,
                    color: COLORS.white,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.greenLight, type: ShadingType.SOLID, color: COLORS.greenLight },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 80, after: 40 },
                children: [
                  new TextRun({
                    text: `${CHECKBOX} YES`,
                    font: 'Calibri',
                    bold: true,
                    size: 22,
                    color: COLORS.greenSuccess,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 40, after: 80 },
                children: [
                  new TextRun({
                    text: yesResult,
                    font: 'Calibri',
                    size: 20,
                    color: COLORS.headerDark,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            shading: { fill: COLORS.redLight, type: ShadingType.SOLID, color: COLORS.redLight },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 80, after: 40 },
                children: [
                  new TextRun({
                    text: `${CHECKBOX} NO`,
                    font: 'Calibri',
                    bold: true,
                    size: 22,
                    color: COLORS.redDanger,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 40, after: 80 },
                children: [
                  new TextRun({
                    text: noResult,
                    font: 'Calibri',
                    size: 20,
                    color: COLORS.headerDark,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// Helper to create constraint checkbox
function createConstraintCheck(text: string): Paragraph {
  return new Paragraph({
    spacing: { before: 100, after: 100 },
    indent: { left: convertInchesToTwip(0.3) },
    children: [
      new TextRun({
        text: `${CHECKBOX}  ${text}`,
        font: 'Calibri',
        size: 22,
        color: COLORS.headerDark,
      }),
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
          text: 'Permitted Development Flowchart',
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
          text: 'Do you need planning permission? Work through this guide to find out.',
          font: 'Calibri',
          size: 24,
          italics: true,
          color: COLORS.mediumGray,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1000 },
      children: [
        new TextRun({
          text: 'For England & Wales only. Based on The Town and Country Planning',
          font: 'Calibri',
          size: 20,
          color: COLORS.mediumGray,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: '(General Permitted Development) (England) Order 2015 (as amended)',
          font: 'Calibri',
          size: 20,
          color: COLORS.mediumGray,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1200 },
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

// Step 1: Check for Constraints
function createStep1(): (Paragraph | Table)[] {
  return [
    createHeading('STEP 1: Check for Constraints', HeadingLevel.HEADING_1),
    createBodyText('First, check if your property has any special designations that restrict permitted development rights.'),
    new Paragraph({ spacing: { before: 200 } }),
    createBodyText('Check ALL that apply to your property:', { bold: true }),
    new Paragraph({ spacing: { before: 150 } }),
    createConstraintCheck('Conservation Area'),
    createConstraintCheck('Listed Building (Grade I, II*, or II)'),
    createConstraintCheck('Area of Outstanding Natural Beauty (AONB)'),
    createConstraintCheck('National Park'),
    createConstraintCheck('The Broads'),
    createConstraintCheck('World Heritage Site'),
    createConstraintCheck('Site of Special Scientific Interest (SSSI)'),
    createConstraintCheck('Article 4 Direction applies'),
    createConstraintCheck('Previous planning condition removing PD rights'),
    new Paragraph({ spacing: { before: 300 } }),
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
                  spacing: { before: 150, after: 150 },
                  children: [
                    new TextRun({
                      text: 'If you checked ANY box above, your PD rights may be limited or removed.',
                      font: 'Calibri',
                      bold: true,
                      size: 22,
                      color: COLORS.amberWarning,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 50, after: 150 },
                  children: [
                    new TextRun({
                      text: 'You should contact your Local Planning Authority for guidance.',
                      font: 'Calibri',
                      size: 20,
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
    new Paragraph({ spacing: { before: 200 } }),
    createBodyText('How to check: Search your Local Planning Authority website for "planning constraints map" or use the government\'s online planning portal.', { italic: true, color: COLORS.mediumGray }),
    createSectionDivider(),
  ];
}

// Step 2: Property Type Assessment
function createStep2(): (Paragraph | Table)[] {
  return [
    createHeading('STEP 2: Property Type Assessment', HeadingLevel.HEADING_1),
    createBodyText('Permitted development rules only apply to certain property types.'),
    new Paragraph({ spacing: { before: 200 } }),
    createDecisionBox(
      'Is your property a house (not a flat or maisonette)?',
      'Continue to Step 3',
      'PD rights do not apply to flats. Planning permission required.'
    ),
    new Paragraph({ spacing: { before: 300 } }),
    createDecisionBox(
      'Was the property built as a house (not converted from commercial)?',
      'Continue to Step 3',
      'Check if "prior approval" was granted - PD rights may be limited.'
    ),
    new Paragraph({ spacing: { before: 300 } }),
    createDecisionBox(
      'Is this your only or main residence (not a second home)?',
      'Continue to Step 3',
      'PD rights still apply, but check local restrictions.'
    ),
    createSectionDivider(),
  ];
}

// Step 3: Extension Type Rules
function createStep3(): (Paragraph | Table)[] {
  return [
    createHeading('STEP 3: Extension Type Rules', HeadingLevel.HEADING_1),
    createBodyText('Select your extension type and check the specific rules below.'),
    new Paragraph({ spacing: { before: 200 } }),

    // Single-storey rear extension
    createHeading('Single-Storey Rear Extension', HeadingLevel.HEADING_2),
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
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Rule', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Detached', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Semi/Terrace', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Check', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max depth from rear wall', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '8m (or 4m standard)', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '6m (or 3m standard)', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max height (eaves)', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '3m', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max overall height', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '4m', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'No more than 50% of garden covered', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Yes', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
      ],
    }),
    createBodyText('Note: 8m/6m limits require "prior approval" notification to neighbours.', { italic: true, color: COLORS.mediumGray }),
    new Paragraph({ spacing: { before: 300 } }),

    // Two-storey extension
    createHeading('Two-Storey Rear Extension', HeadingLevel.HEADING_2),
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
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Rule', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Requirement', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Check', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max depth from rear wall', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '3m', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max eaves height', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Match existing house', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Min 7m to rear boundary', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Yes', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Matching materials', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Yes', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'No balconies/verandas', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Correct', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
      ],
    }),
    new Paragraph({ spacing: { before: 300 } }),

    // Side extension
    createHeading('Side Extension (Single-Storey)', HeadingLevel.HEADING_2),
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
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Rule', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Requirement', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Check', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max width', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '50% of original house width', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max height', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '4m', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max eaves height', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '3m', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
      ],
    }),
    new Paragraph({ spacing: { before: 300 } }),

    // Loft conversion
    createHeading('Loft Conversion (Dormer)', HeadingLevel.HEADING_2),
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
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ children: [new TextRun({ text: 'Rule', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Detached', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Semi/Terrace', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
            new TableCell({
              shading: { fill: COLORS.primaryBlue, type: ShadingType.SOLID, color: COLORS.primaryBlue },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Check', font: 'Calibri', bold: true, size: 22, color: COLORS.white })] })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Max additional roof space', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '50m\u00B3', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '40m\u00B3', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'No higher than existing roof', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Correct', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Dormer not on front roof slope', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Correct', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Side windows obscured & non-opening', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Below 1.7m from floor', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Matching materials', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Yes', font: 'Calibri', size: 22 })] })] }),
            new TableCell({ shading: { fill: COLORS.inputOrange, type: ShadingType.SOLID, color: COLORS.inputOrange }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: CHECKBOX, font: 'Calibri', size: 22 })] })] }),
          ],
        }),
      ],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// Assessment Result Section
function createAssessmentResult(): (Paragraph | Table)[] {
  return [
    createHeading('Assessment Result', HeadingLevel.HEADING_1),
    createBodyText('Based on your answers above, determine your assessment result:'),
    new Paragraph({ spacing: { before: 300 } }),

    // Green result
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 12, color: COLORS.greenSuccess },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.greenSuccess },
        left: { style: BorderStyle.SINGLE, size: 12, color: COLORS.greenSuccess },
        right: { style: BorderStyle.SINGLE, size: 12, color: COLORS.greenSuccess },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.greenLight, type: ShadingType.SOLID, color: COLORS.greenLight },
              children: [
                new Paragraph({
                  spacing: { before: 150 },
                  children: [
                    new TextRun({ text: `${CHECKBOX}  `, font: 'Calibri', size: 28 }),
                    new TextRun({ text: 'GREEN - Likely Permitted Development', font: 'Calibri', bold: true, size: 28, color: COLORS.greenSuccess }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 100, after: 150 },
                  indent: { left: convertInchesToTwip(0.5) },
                  children: [
                    new TextRun({ text: 'No constraints checked in Step 1, AND all rules met in Step 3.', font: 'Calibri', size: 22 }),
                  ],
                }),
                new Paragraph({
                  indent: { left: convertInchesToTwip(0.5) },
                  spacing: { after: 150 },
                  children: [
                    new TextRun({ text: 'Action: ', font: 'Calibri', bold: true, size: 22 }),
                    new TextRun({ text: 'Consider applying for a Lawful Development Certificate for peace of mind.', font: 'Calibri', size: 22 }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    new Paragraph({ spacing: { before: 300 } }),

    // Amber result
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 12, color: COLORS.amberWarning },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.amberWarning },
        left: { style: BorderStyle.SINGLE, size: 12, color: COLORS.amberWarning },
        right: { style: BorderStyle.SINGLE, size: 12, color: COLORS.amberWarning },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.amberLight, type: ShadingType.SOLID, color: COLORS.amberLight },
              children: [
                new Paragraph({
                  spacing: { before: 150 },
                  children: [
                    new TextRun({ text: `${CHECKBOX}  `, font: 'Calibri', size: 28 }),
                    new TextRun({ text: 'AMBER - Requires Further Investigation', font: 'Calibri', bold: true, size: 28, color: COLORS.amberWarning }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 100, after: 150 },
                  indent: { left: convertInchesToTwip(0.5) },
                  children: [
                    new TextRun({ text: 'Some constraints apply OR "prior approval" needed for larger extensions.', font: 'Calibri', size: 22 }),
                  ],
                }),
                new Paragraph({
                  indent: { left: convertInchesToTwip(0.5) },
                  spacing: { after: 150 },
                  children: [
                    new TextRun({ text: 'Action: ', font: 'Calibri', bold: true, size: 22 }),
                    new TextRun({ text: 'Contact your Local Planning Authority for pre-application advice.', font: 'Calibri', size: 22 }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    new Paragraph({ spacing: { before: 300 } }),

    // Red result
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 12, color: COLORS.redDanger },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.redDanger },
        left: { style: BorderStyle.SINGLE, size: 12, color: COLORS.redDanger },
        right: { style: BorderStyle.SINGLE, size: 12, color: COLORS.redDanger },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.redLight, type: ShadingType.SOLID, color: COLORS.redLight },
              children: [
                new Paragraph({
                  spacing: { before: 150 },
                  children: [
                    new TextRun({ text: `${CHECKBOX}  `, font: 'Calibri', size: 28 }),
                    new TextRun({ text: 'RED - Planning Permission Required', font: 'Calibri', bold: true, size: 28, color: COLORS.redDanger }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 100, after: 150 },
                  indent: { left: convertInchesToTwip(0.5) },
                  children: [
                    new TextRun({ text: 'Extension exceeds PD limits OR property has restrictions removing PD rights.', font: 'Calibri', size: 22 }),
                  ],
                }),
                new Paragraph({
                  indent: { left: convertInchesToTwip(0.5) },
                  spacing: { after: 150 },
                  children: [
                    new TextRun({ text: 'Action: ', font: 'Calibri', bold: true, size: 22 }),
                    new TextRun({ text: 'Submit a full planning application. Consider using a planning consultant.', font: 'Calibri', size: 22 }),
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

// Disclaimer section
function createDisclaimer(): (Paragraph | Table)[] {
  return [
    createHeading('Important Disclaimer', HeadingLevel.HEADING_1),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 8, color: COLORS.mediumGray },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: COLORS.mediumGray },
        left: { style: BorderStyle.SINGLE, size: 8, color: COLORS.mediumGray },
        right: { style: BorderStyle.SINGLE, size: 8, color: COLORS.mediumGray },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.lightGray, type: ShadingType.SOLID, color: COLORS.lightGray },
              children: [
                new Paragraph({
                  spacing: { before: 150, after: 100 },
                  children: [
                    new TextRun({
                      text: 'This flowchart is for general guidance only and does not constitute legal or professional advice.',
                      font: 'Calibri',
                      bold: true,
                      size: 22,
                      color: COLORS.headerDark,
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { after: 100 },
                  children: [
                    new TextRun({
                      text: 'Permitted development rules are complex and subject to change. This guide:',
                      font: 'Calibri',
                      size: 22,
                      color: COLORS.headerDark,
                    }),
                  ],
                }),
                new Paragraph({
                  indent: { left: convertInchesToTwip(0.3) },
                  children: [
                    new TextRun({ text: '\u2022 Applies to England only (Wales has different rules)', font: 'Calibri', size: 22 }),
                  ],
                }),
                new Paragraph({
                  indent: { left: convertInchesToTwip(0.3) },
                  children: [
                    new TextRun({ text: '\u2022 Is based on regulations current as of 2024', font: 'Calibri', size: 22 }),
                  ],
                }),
                new Paragraph({
                  indent: { left: convertInchesToTwip(0.3) },
                  children: [
                    new TextRun({ text: '\u2022 Does not cover every possible scenario', font: 'Calibri', size: 22 }),
                  ],
                }),
                new Paragraph({
                  indent: { left: convertInchesToTwip(0.3) },
                  children: [
                    new TextRun({ text: '\u2022 Should be verified with your Local Planning Authority', font: 'Calibri', size: 22 }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 150, after: 150 },
                  children: [
                    new TextRun({
                      text: 'Always confirm your permitted development rights with your local council before starting work.',
                      font: 'Calibri',
                      bold: true,
                      size: 22,
                      color: COLORS.primaryBlue,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
}

export async function generatePdFlowchart(): Promise<Buffer> {
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
                    text: 'Extension Survival Guide  |  PD Flowchart',
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
          ...createStep1(),
          ...createStep2(),
          ...createStep3(),
          ...createAssessmentResult(),
          ...createDisclaimer(),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
}
