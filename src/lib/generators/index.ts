// XLSX Generators - Extension Survival Guide
// Extension Decision Calculator
export { generateExtensionCalculator } from './calculator';
export { generateExtensionCalculator as generateCalculator } from './calculator';

// Budget Planner
export { generateBudgetPlanner } from './budget';
export { generateBudgetPlanner as generateBudget } from './budget';

// Builder Vetting Scorecard
export { generateBuilderScorecard } from './builder-scorecard';

// Payment Schedule Template
export { generatePaymentSchedule } from './payment-schedule';

// Planning Application Tracker
export { generatePlanningTracker } from './planning-tracker';

// Communication Log
export { generateCommunicationLog } from './communication-log';

// Snagging Checklist
export { generateSnaggingChecklist } from './snagging';
export { generateSnaggingChecklist as generateSnagging } from './snagging';

// Variation Request Register
export { generateVariationRegister } from './variation-register';

// DOCX Generators
export { generateScope } from './scope';
export { generatePdFlowchart } from './pd-flowchart';
export { generateDocumentChecklist } from './document-checklist';

// Type definitions for generator functions
export type GeneratorFunction = () => Promise<Buffer>;

// Map of Excel generator names to functions for dynamic access
export const excelGenerators: Record<string, GeneratorFunction> = {
  'extension-calculator': async () => {
    const { generateExtensionCalculator } = await import('./calculator');
    return generateExtensionCalculator();
  },
  'budget-planner': async () => {
    const { generateBudgetPlanner } = await import('./budget');
    return generateBudgetPlanner();
  },
  'builder-scorecard': async () => {
    const { generateBuilderScorecard } = await import('./builder-scorecard');
    return generateBuilderScorecard();
  },
  'payment-schedule': async () => {
    const { generatePaymentSchedule } = await import('./payment-schedule');
    return generatePaymentSchedule();
  },
  'planning-tracker': async () => {
    const { generatePlanningTracker } = await import('./planning-tracker');
    return generatePlanningTracker();
  },
  'communication-log': async () => {
    const { generateCommunicationLog } = await import('./communication-log');
    return generateCommunicationLog();
  },
  'snagging-checklist': async () => {
    const { generateSnaggingChecklist } = await import('./snagging');
    return generateSnaggingChecklist();
  },
  'variation-register': async () => {
    const { generateVariationRegister } = await import('./variation-register');
    return generateVariationRegister();
  },
};

// File names for downloads
export const excelFileNames: Record<string, string> = {
  'extension-calculator': 'Extension-Decision-Calculator.xlsx',
  'budget-planner': 'Extension-Budget-Planner.xlsx',
  'builder-scorecard': 'Builder-Vetting-Scorecard.xlsx',
  'payment-schedule': 'Payment-Schedule-Template.xlsx',
  'planning-tracker': 'Planning-Application-Tracker.xlsx',
  'communication-log': 'Project-Communication-Log.xlsx',
  'snagging-checklist': 'Snagging-Checklist.xlsx',
  'variation-register': 'Variation-Request-Register.xlsx',
};
