export interface Tool {
  id: string;
  name: string;
  description: string;
  price: string;
  bundle: string;
  icon: string;
  link: string;
}

export const tools: Record<string, Tool> = {
  'budget-planner': {
    id: 'budget-planner',
    name: 'Extension Budget Planner',
    description: 'Forces you to include the 23 cost categories most people forget.',
    price: '£19.99',
    bundle: 'essential',
    icon: 'Calculator',
    link: '/#pricing',
  },
  'builder-scorecard': {
    id: 'builder-scorecard',
    name: 'Builder Vetting Scorecard',
    description: 'The exact criteria architects use to assess builders before recommending them.',
    price: '£14.99',
    bundle: 'essential',
    icon: 'ClipboardCheck',
    link: '/#pricing',
  },
  'scope-document': {
    id: 'scope-document',
    name: 'Scope of Works Template',
    description: 'Define exactly what is included. No ambiguity. No extras.',
    price: '£24.99',
    bundle: 'essential',
    icon: 'FileText',
    link: '/#pricing',
  },
  'payment-schedule': {
    id: 'payment-schedule',
    name: 'Payment Schedule Template',
    description: 'Stage payments tied to completed work, not calendar dates.',
    price: '£14.99',
    bundle: 'essential',
    icon: 'Wallet',
    link: '/#pricing',
  },
  'snagging-checklist': {
    id: 'snagging-checklist',
    name: 'Snagging Checklist',
    description: 'The 200+ point checklist architects use for final inspections.',
    price: '£19.99',
    bundle: 'essential',
    icon: 'CheckSquare',
    link: '/#pricing',
  },
  'variation-register': {
    id: 'variation-register',
    name: 'Variation Register',
    description: 'Track every change request with cost implications documented.',
    price: '£14.99',
    bundle: 'essential',
    icon: 'FileEdit',
    link: '/#pricing',
  },
  'communication-log': {
    id: 'communication-log',
    name: 'Site Communication Log',
    description: 'Document every conversation. Essential if disputes arise.',
    price: '£9.99',
    bundle: 'essential',
    icon: 'MessageSquare',
    link: '/#pricing',
  },
  'planning-tracker': {
    id: 'planning-tracker',
    name: 'Planning Tracker',
    description: 'Track your planning application from submission to approval.',
    price: '£14.99',
    bundle: 'essential',
    icon: 'Calendar',
    link: '/#pricing',
  },
  'pd-flowchart': {
    id: 'pd-flowchart',
    name: 'Permitted Development Flowchart',
    description: 'Check if your extension needs planning permission.',
    price: '£9.99',
    bundle: 'essential',
    icon: 'GitBranch',
    link: '/#pricing',
  },
  'extend-or-move': {
    id: 'extend-or-move',
    name: 'Extend or Move Calculator',
    description: 'The financial calculation that determines whether to extend or relocate.',
    price: '£14.99',
    bundle: 'essential',
    icon: 'Scale',
    link: '/#pricing',
  },
};

export function getTool(toolId: string): Tool | undefined {
  return tools[toolId];
}

export function getAllTools(): Tool[] {
  return Object.values(tools);
}
