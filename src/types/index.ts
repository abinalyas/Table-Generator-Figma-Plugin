// Type definitions for the application

export interface ComponentInfo {
  id: string | null;
  name: string;
  width?: number;
  properties: { [key: string]: any };
  availableProperties: string[];
  propertyTypes: { [key: string]: any };
}

export interface State {
  selectedCells: Set<string>;
  isDragging: boolean;
  startCell: string | null;
  endCell: string | null;
  componentProps: { [key: string]: any };
  cellProperties: Map<string, any>;
  gridCols: number;
  gridRows: number;
  hasComponent: boolean;
  mode: 'selection' | 'edit';
  currentEditingCell: string | null;
  applyMode: 'cell' | 'row' | 'column';
  hoverTimeout: number | null;
  tooltip: HTMLDivElement;
  selectedSize: { width: number; height: number } | null;
  selectedComponent: ComponentInfo | null;
  sizeConfirmed: boolean;
  componentWidth: number | undefined;
  headerCellComponent: ComponentInfo | null;
  footerComponent: ComponentInfo | null;
  tableFrameId: string | undefined;
}

export interface Elements {
  grid: HTMLElement;
  gridHighlight: HTMLElement;
  statusMessage: HTMLElement;
  gridContainer: HTMLElement;
  actionButtons: HTMLElement;
  createTableBtn: HTMLButtonElement;
  clearSelectionBtn: HTMLButtonElement;
  propertyEditor: HTMLElement;
  propertyEditorTitle: HTMLElement;
  editingCellCoords: HTMLElement;
  showText: HTMLInputElement;
  secondTextLine: HTMLInputElement;
  secondCellText: HTMLInputElement;
  state: HTMLSelectElement;
  cancelPropsBtn: HTMLButtonElement;
  savePropsBtn: HTMLButtonElement;
  secondLineContainer: HTMLElement;
  secondCellTextContainer: HTMLElement;
  loader: HTMLElement;
  loaderText: HTMLElement;
  aiConfirmationDialog: HTMLElement;
  cancelAiBtn: HTMLButtonElement;
  confirmAiBtn: HTMLButtonElement;
  generateSampleCheckbox: HTMLInputElement;
  aiPromptContainer: HTMLElement;
  colWidthInput: HTMLInputElement;
  colWidthContainer: HTMLElement;
  slotCheckbox: HTMLInputElement;
  customCellTextToggle: HTMLInputElement;
  customCellText: HTMLInputElement;
  customCellTextContainer: HTMLElement;
  landingPage: HTMLElement;
  componentModeBtn: HTMLButtonElement | null;
  propertyEditorOverlay: HTMLElement;
  componentDisplay: HTMLElement;
  scanOptionsContainer: HTMLElement | null;
  scanTableBtn?: HTMLButtonElement;
  scanTableSection?: HTMLElement;
}

export interface PropertyDefinition {
  name: string;
  label: string;
  type: 'TEXT' | 'VARIANT' | 'BOOLEAN';
  options?: string[];
  defaultValue?: string;
  dependsOn?: string;
  showWhen?: string;
}