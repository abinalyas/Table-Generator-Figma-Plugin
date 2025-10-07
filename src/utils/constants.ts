// Simple constants module - no complex imports or exports
import type { PropertyDefinition } from '../types';

export const HEADER_CELL_MODEL: PropertyDefinition[] = [
  { name: "Cell text#12234:32", label: "Header Text", type: "TEXT" },
  { name: "State", label: "State", type: "VARIANT", options: ["Enable", "Hover", "Focus"] },
  { name: "Sortable", label: "Sortable", type: "VARIANT", options: ["True", "False"], defaultValue: "False" },
  { name: "Sorted", label: "Sorted", type: "VARIANT", options: ["Ascending", "Descending", "None"], defaultValue: "Ascending", dependsOn: "Sortable", showWhen: "True" }
];

export const FOOTER_MODEL: PropertyDefinition[] = [
  { name: "Total items#12006:49", label: "Total items", type: "TEXT" },
  { name: "Current page#12006:39", label: "Current page", type: "TEXT" },
  { name: "Total pages#12006:29", label: "Total pages", type: "TEXT" },
  { name: "Type", label: "Type", type: "VARIANT", options: ["Advanced", "Simple"] },
  { name: "Size", label: "Size", type: "VARIANT", options: ["Large", "Small"] }
];