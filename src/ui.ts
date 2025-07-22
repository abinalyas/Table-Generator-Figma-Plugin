import { fakerMethods } from './faker-methods';

interface State {
  selectedCells: Set<string>;
  isDragging: boolean;
  startCell: HTMLDivElement | null;
  endCell: HTMLDivElement | null;
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
  selectedSize: string | null;
  selectedComponent: {
    id: string | null;
    name: string;
    width?: number;
    properties: { [key: string]: any };
    availableProperties: string[];
    propertyTypes: { [key: string]: any };
  } | null;
  sizeConfirmed: boolean;
  componentWidth?: number;
  headerCellComponent?: {
    id: string | null;
    name: string;
    width?: number;
    properties: { [key: string]: any };
    availableProperties: string[];
    propertyTypes: { [key: string]: any };
  } | null;
  footerComponent?: {
    id: string | null;
    name: string;
    width?: number;
    properties: { [key: string]: any };
    availableProperties: string[];
    propertyTypes: { [key: string]: any };
  } | null;
}

interface Elements {
  grid: HTMLElement;
  gridHighlight: HTMLElement;
  dimensionsDisplay: HTMLElement;
  statusMessage: HTMLElement;
  gridContainer: HTMLElement;
  actionButtons: HTMLElement;
  createTableBtn: HTMLButtonElement;
  clearSelectionBtn: HTMLButtonElement;
  selectedCount: HTMLElement;
  propertyEditor: HTMLElement;
  propertyEditorTitle: HTMLElement;
  editingCellCoords: HTMLElement;
  showText: HTMLInputElement;
  cellText: HTMLInputElement;
  secondTextLine: HTMLInputElement;
  secondCellText: HTMLInputElement;
  state: HTMLSelectElement;
  selectionModeBtn: HTMLButtonElement;
  editModeBtn: HTMLButtonElement;
  confirmSizeBtn: HTMLButtonElement;
  cancelSelectionBtn: HTMLButtonElement;
  generateDataBtn: HTMLButtonElement;
  cancelPropsBtn: HTMLButtonElement;
  savePropsBtn: HTMLButtonElement;
  cellTextContainer: HTMLElement;
  secondLineContainer: HTMLElement;
  secondCellTextContainer: HTMLElement;
  selectionActions: HTMLElement;
  dimensionPresets: HTMLElement;
  loader: HTMLElement;
  loaderText: HTMLElement;
  aiConfirmationDialog: HTMLElement;
  cancelAiBtn: HTMLButtonElement;
  confirmAiBtn: HTMLButtonElement;
  generateSampleCheckbox: HTMLInputElement;
  aiPromptContainer: HTMLElement;
  rowWidthInput: HTMLInputElement;
  colWidthInput: HTMLInputElement;
  colWidthContainer: HTMLElement;
  slotCheckbox: HTMLInputElement;
  landingPage: HTMLElement;
  componentModeBtn: HTMLButtonElement;
  propertyEditorOverlay: HTMLElement;
  scanTableSection: HTMLElement | null;
  scanTableBtn: HTMLButtonElement | null;
  scanTableInstructions: HTMLElement | null;
  componentDisplay: HTMLElement;
  scanOptionsContainer: HTMLElement | null;
}

// State Management
const state: State = {
  selectedCells: new Set(),
  isDragging: false,
  startCell: null,
  endCell: null,
  componentProps: {},
  cellProperties: new Map(),
  gridCols: 5,
  gridRows: 5,
  hasComponent: false,
  mode: 'selection',
  currentEditingCell: null,
  applyMode: 'column',
  hoverTimeout: null,
  tooltip: document.createElement('div'),
  selectedSize: null,
  selectedComponent: null,
  sizeConfirmed: false,
  componentWidth: undefined,
  headerCellComponent: null,
  footerComponent: null
};

// DOM Elements
let elements = {} as Elements;

// DOM Elements for faker dropdown
const fakerMethodInput = document.getElementById('fakerMethodInput') as HTMLInputElement;
const fakerDropdown = document.getElementById('fakerDropdown')!;

// Create tooltip element
state.tooltip.className = 'cell-tooltip';
// document.body.appendChild(state.tooltip);

let domReady = false;

// Static property models for header and footer
const HEADER_CELL_MODEL = [
  { name: "Cell text#12234:32", label: "Header Text", type: "TEXT" },
  { name: "Size", label: "Size", type: "VARIANT", options: ["Extra large", "Large", "Small"] },
  { name: "State", label: "State", type: "VARIANT", options: ["Enabled", "Disabled", "Focus"] },
  { name: "Sorted", label: "Sorted", type: "VARIANT", options: ["None", "Ascending", "Descending"] },
  { name: "Sortable", label: "Sortable", type: "VARIANT", options: ["True", "False"] }
];
const FOOTER_MODEL = [
  { name: "Total items#12006:49", label: "Total items", type: "TEXT" },
  { name: "Current page#12006:39", label: "Current page", type: "TEXT" },
  { name: "Total pages#12006:29", label: "Total pages", type: "TEXT" },
  { name: "Type", label: "Type", type: "VARIANT", options: ["Advanced", "Simple"] },
  { name: "Size", label: "Size", type: "VARIANT", options: ["Large", "Small"] }
];

// Setup landing page button logic and initialize elements after DOM is ready
window.addEventListener('DOMContentLoaded', () => {
  // Assign only existing DOM elements
  elements.grid = document.getElementById('grid')!;
  elements.gridHighlight = document.getElementById('gridHighlight')!;
  elements.dimensionsDisplay = document.getElementById('dimensionsDisplay')!;
  elements.statusMessage = document.getElementById('statusMessage')!;
  elements.gridContainer = document.getElementById('gridContainer')!;
  elements.actionButtons = document.getElementById('actionButtons')!;
  elements.createTableBtn = document.getElementById('createTableBtn') as HTMLButtonElement;
  elements.clearSelectionBtn = document.getElementById('clearSelectionBtn') as HTMLButtonElement;
  elements.selectedCount = document.getElementById('selectedCount')!;
  elements.propertyEditor = document.getElementById('propertyEditor')!;
  elements.propertyEditorTitle = document.getElementById('propertyEditorTitle')!;
  elements.editingCellCoords = document.getElementById('editingCellCoords')!;
  elements.showText = document.getElementById('showText') as HTMLInputElement;
  elements.cellText = document.getElementById('cellText') as HTMLInputElement;
  elements.secondTextLine = document.getElementById('secondTextLine') as HTMLInputElement;
  elements.secondCellText = document.getElementById('secondCellText') as HTMLInputElement;
  elements.state = document.getElementById('state') as HTMLSelectElement;
  elements.generateDataBtn = document.getElementById('generateDataBtn') as HTMLButtonElement;
  elements.cancelPropsBtn = document.getElementById('cancelPropsBtn') as HTMLButtonElement;
  elements.savePropsBtn = document.getElementById('savePropsBtn') as HTMLButtonElement;
  elements.cellTextContainer = document.getElementById('cellTextContainer')!;
  elements.secondLineContainer = document.getElementById('secondLineContainer')!;
  elements.secondCellTextContainer = document.getElementById('secondCellTextContainer')!;
  elements.loader = document.getElementById('loader')!;
  elements.loaderText = document.getElementById('loaderText')!;
  elements.aiConfirmationDialog = document.getElementById('aiConfirmationDialog')!;
  elements.cancelAiBtn = document.getElementById('cancelAiBtn') as HTMLButtonElement;
  elements.confirmAiBtn = document.getElementById('confirmAiBtn') as HTMLButtonElement;
  elements.generateSampleCheckbox = document.getElementById('generateSampleCheckbox') as HTMLInputElement;
  elements.aiPromptContainer = document.getElementById('aiPromptContainer')!;
  elements.rowWidthInput = document.getElementById('rowWidthInput') as HTMLInputElement;
  elements.colWidthInput = document.getElementById('colWidthInput') as HTMLInputElement;
  elements.colWidthContainer = document.getElementById('colWidthContainer')!;
  elements.slotCheckbox = document.getElementById('slotCheckbox') as HTMLInputElement;
  elements.landingPage = document.getElementById('landingPage')!;
  elements.componentModeBtn = document.getElementById('componentModeBtn') as HTMLButtonElement;
  elements.propertyEditorOverlay = document.getElementById('propertyEditorOverlay')!;
  elements.scanTableSection = null;
  elements.scanTableBtn = null;
  elements.scanTableInstructions = null;
  elements.componentDisplay = document.getElementById('componentDisplay')!;
  elements.scanOptionsContainer = null;

  domReady = true;

  // Now safe to call setup functions
  createGrid();
  setupEventListeners();
  setMode('selection');

  // Show colWidthInput if default mode is column
  if (state.applyMode === 'column' && elements.colWidthInput) {
    elements.colWidthInput.style.display = 'inline-block';
  }

  // Set default values
  const defaultWidth = state.selectedComponent?.width || 100; // Use component width or fallback to 100
  const defaultHeight = 64;
  const defaultText = "content";

  // Setup landing page button logic
  elements.componentModeBtn.onclick = () => {
    elements.landingPage.style.display = 'none';
  };
});

function renderHeaderFooterGrids() {
  // Header grid
  const headerGrid = document.getElementById('headerGrid');
  if (headerGrid) {
    headerGrid.innerHTML = '';
    headerGrid.style.setProperty('--header-cols', String(state.gridCols));
    for (let c = 1; c <= state.gridCols; c++) {
      const cell = document.createElement('div');
      cell.className = 'header-cell';
      cell.textContent = `H${c}`;
      cell.addEventListener('click', () => {
        openPropertyEditor(`header-${c}`);
      });
      headerGrid.appendChild(cell);
    }
  }
  // Footer grid
  const footerGrid = document.getElementById('footerGrid');
  if (footerGrid) {
    footerGrid.innerHTML = '';
    footerGrid.style.setProperty('--footer-cols', String(state.gridCols));
    const cell = document.createElement('div');
    cell.className = 'footer-cell';
    cell.textContent = 'Footer';
    cell.style.gridColumn = `span ${state.gridCols}`;
    cell.addEventListener('click', () => {
      openPropertyEditor('footer');
    });
    footerGrid.appendChild(cell);
  }
}

// Initialize
function createGrid() {
  const grid = elements.grid;
  grid.innerHTML = '';
  grid.style.gridTemplateColumns = `repeat(${state.gridCols}, 30px)`;

  for (let r = 1; r <= state.gridRows; r++) {
    for (let c = 1; c <= state.gridCols; c++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.dataset.row = String(r);
      cell.dataset.col = String(c);
      cell.textContent = `${r},${c}`;
      cell.addEventListener('click', () => {
        const key = `${r},${c}`;
        openPropertyEditor(key);
      });
      grid.appendChild(cell);
    }
  }
  renderHeaderFooterGrids();
  updateCreateButtonState();
}

function setupEventListeners() {
  elements.clearSelectionBtn.addEventListener('click', resetTableProperties);
  elements.createTableBtn.addEventListener('click', createTable);
  elements.cancelPropsBtn.addEventListener('click', closePropertyEditor);
  elements.propertyEditorOverlay.addEventListener('click', closePropertyEditor);
  elements.savePropsBtn.addEventListener('click', saveCellProperties);
  elements.showText.addEventListener('change', function (this: HTMLInputElement) {
    const show = this.checked;
    elements.cellTextContainer.style.display = show ? 'block' : 'none';
    elements.secondLineContainer.style.display = show ? 'flex' : 'none';
    if (!show && elements.secondTextLine) {
      elements.secondTextLine.checked = false;
      elements.secondCellTextContainer.style.display = 'none';
    }
  });
  elements.secondTextLine.addEventListener('change', function (this: HTMLInputElement) {
    elements.secondCellTextContainer.style.display = this.checked ? 'block' : 'none';
  });
  document.querySelectorAll('.apply-option').forEach(option => {
    option.addEventListener('click', function (this: HTMLElement) {
      document.querySelectorAll('.apply-option').forEach(opt => opt.classList.remove('active'));
      this.classList.add('active');
      state.applyMode = this.dataset.apply as 'cell' | 'row' | 'column';
      // Show/hide col width input
      if (state.applyMode === 'column') {
        elements.colWidthContainer.style.display = 'block';
      } else {
        elements.colWidthContainer.style.display = 'none';
      }
      // Re-render property editor if open
      if (state.currentEditingCell && elements.propertyEditor.style.display === 'block') {
        openPropertyEditor(state.currentEditingCell);
      }
    });
  });
  elements.generateSampleCheckbox.addEventListener('change', () => {
    const isChecked = elements.generateSampleCheckbox.checked;
    elements.cellText.disabled = isChecked;
    elements.aiPromptContainer.style.display = isChecked ? 'block' : 'none';
  });
  document.addEventListener('mouseover', function (e) {
    const target = e.target as HTMLElement;
    if (target.classList.contains('cell')) {
      showCellTooltip(target as HTMLDivElement);
    }
  });
  document.addEventListener('mouseout', function (e) {
    const target = e.target as HTMLElement;
    if (target.classList.contains('cell')) {
      hideCellTooltip();
    }
  });
}

function setMode(newMode: State['mode']) {
  state.mode = newMode;
  document.body.classList.remove('selection-mode', 'edit-mode');
  document.body.classList.add(`${newMode}-mode`);

  state.isDragging = false;
  elements.gridHighlight.style.display = 'none';

  updateModeDependentVisibility();

  // Show colWidthInput if default mode is column
  if (state.applyMode === 'column' && elements.colWidthInput) {
    elements.colWidthInput.style.display = 'inline-block';
  }
}

function updateModeDependentVisibility() {
  // Only show action buttons (Create Table, Clear Selection) if a valid component is selected and table size is confirmed
  if (state.hasComponent && state.sizeConfirmed) {
    elements.actionButtons.style.display = 'flex';
  } else {
    elements.actionButtons.style.display = 'none';
  }
}

function resetTableProperties() {
  state.cellProperties.clear();
  updateCellVisuals();
  showMessage("All cell properties have been reset.", "success");
}

function commitTableSize() {
    if (state.selectedCells.size === 0) {
        showMessage("Please select cells to define table size.", "error");
        return;
    }
    state.sizeConfirmed = true;
    elements.createTableBtn.disabled = false;
    showMessage("Table size confirmed. You can now create the table or switch to edit mode.", "success");
}

function selectDimensionPreset(rows: number, cols: number) {
    const startCell = document.querySelector(`.cell[data-row='1'][data-col='1']`) as HTMLDivElement;
    const endCell = document.querySelector(`.cell[data-row='${rows}'][data-col='${cols}']`) as HTMLDivElement;

    if (startCell && endCell) {
        // Removed drag/selection logic
  }
  updateModeDependentVisibility();
    commitTableSize();
}

function closePropertyEditor() {
    elements.propertyEditor.style.opacity = '0';
    elements.propertyEditorOverlay.style.display = 'none';
    setTimeout(() => {
        elements.propertyEditor.style.display = 'none';
        state.currentEditingCell = null;
    }, 300);
}

function getCellState(key: string) {
    if (!state.cellProperties.has(key)) {
        state.cellProperties.set(key, { properties: {} });
    }
    return state.cellProperties.get(key);
}

function createTable() {
  showLoader('Generating table...');

  // Consolidate property collection
  const propsForFigma: { [key: string]: any } = {};
  state.selectedCells.forEach(key => {
    const [row, col] = key.split(',').map(Number);
    const backendKey = `${row - 1}-${col - 1}`;
    const cellData = state.cellProperties.get(key);
    if (cellData) {
      propsForFigma[backendKey] = cellData;
    } else {
      propsForFigma[backendKey] = { properties: {} };
    }
  });

  // --- Add header cell properties ---
  for (let c = 1; c <= state.gridCols; c++) {
    const key = `header-${c}`;
    const cellData = state.cellProperties.get(key);
    if (cellData) {
      propsForFigma[key] = cellData;
    }
  }
  // --- Add footer cell properties ---
  const footerData = state.cellProperties.get('footer');
  if (footerData) {
    propsForFigma['footer'] = footerData;
  }

  // Only keep scan flow
  const rows = parseInt((document.getElementById('scanRowsInput') as HTMLInputElement)?.value || String(state.gridRows), 10);
  const cols = parseInt((document.getElementById('scanColsInput') as HTMLInputElement)?.value || String(state.gridCols), 10);
  const includeHeader = (document.getElementById('scanHeaderToggle') as HTMLInputElement)?.checked;
  const includeFooter = (document.getElementById('scanFooterToggle') as HTMLInputElement)?.checked;
  const includeSelectable = (document.getElementById('scanSelectableToggle') as HTMLInputElement)?.checked;
  const includeExpandable = (document.getElementById('scanExpandableToggle') as HTMLInputElement)?.checked;
  
  console.log('🚀 Sending to Figma (Scan Flow):', {
      rows,
      cols,
      cellProps: propsForFigma,
      includeHeader,
      includeFooter,
      includeSelectable,
      includeExpandable
  });

  parent.postMessage({
    pluginMessage: {
          type: "create-table-from-scan",
          rows,
          cols,
      cellProps: propsForFigma,
          includeHeader,
          includeFooter,
          includeSelectable,
          includeExpandable
    }
  }, '*');
}

function showLoader(message: string) {
  elements.loader.style.display = 'flex';
  elements.loaderText.textContent = message;
}

function hideLoader() {
  elements.loader.style.display = 'none';
  elements.loaderText.textContent = '';
}

function showMessage(text: string, type: 'success' | 'error') {
  elements.statusMessage.textContent = text;
  elements.statusMessage.className = `status ${type}`;
}

function showCellTooltip(cell: HTMLDivElement) {
  const key = `${cell.dataset.row},${cell.dataset.col}`;
  const props = state.cellProperties.get(key);
  const availableProps: string[] = state.selectedComponent?.availableProperties || [];

  if (props && props.properties && Object.keys(props.properties).length > 0) {
    clearTimeout(state.hoverTimeout!);
    state.hoverTimeout = window.setTimeout(() => {
      let tooltipContent = '';
      let cellTextProp = availableProps.find(p => p.toLowerCase().includes('text') && !p.toLowerCase().includes('second'));
      if (!cellTextProp) cellTextProp = Object.keys(props.properties).find(p => p.toLowerCase().includes('text') && !p.toLowerCase().includes('second')) || 'cellText';
      let secondTextProp = availableProps.find(p => p.toLowerCase().includes('second')) || Object.keys(props.properties).find(p => p.toLowerCase().includes('second')) || 'secondCellText';
      let showTextProp = availableProps.find(p => p.toLowerCase().includes('showtext')) || Object.keys(props.properties).find(p => p.toLowerCase().includes('showtext')) || 'showText';
      let stateProp = availableProps.find(p => p.toLowerCase().includes('state')) || Object.keys(props.properties).find(p => p.toLowerCase().includes('state')) || 'state';
      if (props.properties[showTextProp]) {
        tooltipContent += `<strong>Text:</strong> ${props.properties[cellTextProp] || 'Empty'}<br>`;
      }
      if (props.properties['secondTextLine'] || props.properties[secondTextProp]) {
        tooltipContent += `<strong>Second line:</strong> ${props.properties[secondTextProp] || 'Empty'}<br>`;
      }
      tooltipContent += `<strong>State:</strong> ${props.properties[stateProp] || 'Enabled'}<br>`;
      state.tooltip.innerHTML = tooltipContent;
      state.tooltip.style.display = 'block';
      const rect = cell.getBoundingClientRect();
      state.tooltip.style.left = `${rect.left}px`;
      state.tooltip.style.top = `${rect.bottom + 5}px`;
    }, 300);
  } else {
    hideCellTooltip();
  }
}

function hideCellTooltip() {
  clearTimeout(state.hoverTimeout!);
  state.tooltip.style.display = 'none';
}

function findMatchingProperty(availableProperties: string[], uiPropName: string): string | null {
    console.log('🔍 Finding match for:', uiPropName);
    console.log('Available properties:', availableProperties);
    // First try exact match
    const exactMatch = availableProperties.find((prop: string) => prop === uiPropName);
    if (exactMatch) {
        console.log('✅ Found exact match:', exactMatch);
        return exactMatch;
    }
    // Try match at start of property name
    const startsWithMatch = availableProperties.find((prop: string) => prop.toLowerCase().startsWith(uiPropName.toLowerCase()));
    if (startsWithMatch) {
        console.log('✅ Found starts-with match:', startsWithMatch);
        return startsWithMatch;
    }
    // Try contains match
    const containsMatch = availableProperties.find((prop: string) => prop.toLowerCase().includes(uiPropName.toLowerCase()));
    if (containsMatch) {
        console.log('✅ Found contains match:', containsMatch);
        return containsMatch;
    }
    console.log('❌ No match found');
    return null;
}

// Add a helper to render property fields dynamically
function renderDynamicPropertyFields(availableProps: string[], propertyTypes: { [key: string]: any }, props: any) {
  const container = document.getElementById('dynamicPropertyFields');
  if (!container) return;
  container.innerHTML = '';
  container.style.border = '';
  container.style.background = '';
  let fieldCount = 0;
  console.log('[DEBUG] renderDynamicPropertyFields', { availableProps, propertyTypes, props });

  // Map Figma property names to user-friendly labels and dropdown options
  const labelMap: { [key: string]: string } = {
    'Cell text#12234:32': 'Header Text',
    'Size': 'Size',
    'State': 'State',
    'Sorted': 'Sorted',
    'Sortable': 'Sortable',
    'Total items#12006:49': 'Total items',
    'Current page#12006:39': 'Current page',
    'Total pages#12006:29': 'Total pages',
    'Type': 'Type',
  };
  const optionsMap: { [key: string]: string[] } = {
    'Size': ['Extra large', 'Large', 'Small'],
    'State': ['Enabled', 'Disabled', 'Focus'],
    'Sorted': ['None', 'Ascending', 'Descending'],
    'Sortable': ['True', 'False'],
    'Type': ['Advanced', 'Simple'],
  };

  for (const propName of availableProps) {
    if (propName === 'Size' && state.applyMode !== 'row') continue;
    const type = propertyTypes[propName];
    const label = labelMap[propName] || cleanPropName(propName);
    const value = props[propName] ?? '';
    let field: HTMLElement | null = null;
    if (type === 'VARIANT') {
      field = document.createElement('div');
      field.className = 'property-field';
      const select = document.createElement('select');
      select.className = 'styled-input';
      select.id = `dynamic-${propName}`;
      // Use mapped options if available, else fallback
      const options = optionsMap[propName] || [String(value)];
      for (const opt of options) {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        if (String(value) === opt) option.selected = true;
        select.appendChild(option);
      }
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      field.appendChild(labelEl);
      field.appendChild(select);
      // --- Attach event for Show text ---
      if (label.toLowerCase().includes('show text')) {
        select.addEventListener('change', function () {
          const cellTextInput = document.getElementById('dynamic-Cell text#12234:16') as HTMLInputElement;
          if (cellTextInput) {
            const cellTextField = cellTextInput.parentElement as HTMLElement;
            if (cellTextField) {
              cellTextField.style.display = this.value === 'Cell text#12234:16' ? '' : 'none';
            }
          }
        });
      }
      // --- Attach event for Second text line ---
      if (label.toLowerCase().includes('second text line')) {
        select.addEventListener('change', function () {
          const secondCellTextInput = document.getElementById('dynamic-Second cell text#105573:16') as HTMLInputElement;
          if (secondCellTextInput) {
            const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
            if (secondCellTextField) {
              secondCellTextField.style.display = this.value === 'Second cell text#105573:16' ? '' : 'none';
            }
          }
        });
      }
    } else if (type === 'BOOLEAN') {
      field = document.createElement('div');
      field.className = 'property-field checkbox';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.id = `dynamic-${propName}`;
      input.checked = !!value;
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      field.appendChild(input);
      field.appendChild(labelEl);
      // --- Attach event for Show text ---
      if (label.toLowerCase().includes('show text')) {
        input.addEventListener('change', function () {
          const cellTextInput = document.getElementById('dynamic-Cell text#12234:16') as HTMLInputElement;
          if (cellTextInput) {
            const cellTextField = cellTextInput.parentElement as HTMLElement;
            if (cellTextField) {
              cellTextField.style.display = this.checked ? '' : 'none';
            }
          }
        });
      }
      // --- Attach event for Second text line ---
      if (label.toLowerCase().includes('second text line')) {
        input.addEventListener('change', function () {
          const secondCellTextInput = document.getElementById('dynamic-Second cell text#105573:16') as HTMLInputElement;
          if (secondCellTextInput) {
            const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
            if (secondCellTextField) {
              secondCellTextField.style.display = this.checked ? '' : 'none';
            }
          }
        });
      }
    } else if (type === 'TEXT') {
      field = document.createElement('div');
      field.className = 'property-field';
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'styled-input';
      input.id = `dynamic-${propName}`;
      input.value = value;
      field.appendChild(labelEl);
      field.appendChild(input);
    }
    if (field) {
      container.appendChild(field);
      fieldCount++;
    }
  }
  // After all fields are rendered, ensure Cell text is hidden if Show text is unchecked
  const showTextCheckbox = container.querySelector('input[type="checkbox"][id^="dynamic-Show text"]') as HTMLInputElement;
  const cellTextInput = container.querySelector('input[type="text"][id^="dynamic-Cell text"]') as HTMLInputElement;
  if (showTextCheckbox && cellTextInput) {
    const cellTextField = cellTextInput.parentElement as HTMLElement;
    if (cellTextField) {
      cellTextField.style.display = showTextCheckbox.checked ? '' : 'none';
    }
  }
  // After all fields are rendered, ensure Second cell text is hidden if Second text line is unchecked
  const secondTextLineCheckbox = container.querySelector('input[type="checkbox"][id^="dynamic-Second text line"]') as HTMLInputElement;
  const secondCellTextInput = container.querySelector('input[type="text"][id^="dynamic-Second cell text"]') as HTMLInputElement;
  if (secondTextLineCheckbox && secondCellTextInput) {
    const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
    if (secondCellTextField) {
      secondCellTextField.style.display = secondTextLineCheckbox.checked ? '' : 'none';
    }
  }
  // --- Hide dynamic Second text line checkbox if Show text is not checked ---
  const secondLineContainerDyn = secondTextLineCheckbox?.parentElement as HTMLElement;
  if (secondLineContainerDyn && showTextCheckbox) {
    secondLineContainerDyn.style.display = showTextCheckbox.checked ? '' : 'none';
    if (!showTextCheckbox.checked) {
      secondTextLineCheckbox.checked = false;
      if (secondCellTextInput) {
        const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
        if (secondCellTextField) secondCellTextField.style.display = 'none';
      }
    }
    showTextCheckbox.addEventListener('change', function () {
      secondLineContainerDyn.style.display = this.checked ? '' : 'none';
      if (!this.checked) {
        secondTextLineCheckbox.checked = false;
        if (secondCellTextInput) {
          const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
          if (secondCellTextField) secondCellTextField.style.display = 'none';
        }
      }
    });
  }
  console.log('[DEBUG] renderDynamicPropertyFields created fields:', fieldCount);
}

function renderDynamicPropertyFieldsFromModel(model: any[], props: any) {
  const container = document.getElementById('dynamicPropertyFields');
  if (!container) return;
  container.innerHTML = '';
  container.style.border = '';
  container.style.background = '';
  let fieldCount = 0;
  for (const fieldDef of model) {
    if (fieldDef.name === 'Size' && state.applyMode !== 'row') continue;
    const { name, label, type, options } = fieldDef;
    const value = props[name] ?? '';
    let field: HTMLElement | null = null;
    if (type === 'VARIANT') {
      field = document.createElement('div');
      field.className = 'property-field';
      const select = document.createElement('select');
      select.className = 'styled-input';
      select.id = `dynamic-${name}`;
      for (const opt of options || []) {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        if (String(value) === opt) option.selected = true;
        select.appendChild(option);
      }
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      field.appendChild(labelEl);
      field.appendChild(select);
    } else if (type === 'BOOLEAN') {
      field = document.createElement('div');
      field.className = 'property-field checkbox';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.id = `dynamic-${name}`;
      input.checked = !!value;
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      field.appendChild(input);
      field.appendChild(labelEl);
    } else if (type === 'TEXT') {
      field = document.createElement('div');
      field.className = 'property-field';
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'styled-input';
      input.id = `dynamic-${name}`;
      input.value = value;
      field.appendChild(labelEl);
      field.appendChild(input);
    }
    if (field) {
      container.appendChild(field);
      fieldCount++;
    }
  }
  console.log('[DEBUG] renderDynamicPropertyFieldsFromModel created fields:', fieldCount);
}

function renderBodyCellProperties(availableProps: string[], propertyTypes: { [key: string]: any }, props: any) {
  const container = document.getElementById('dynamicPropertyFields');
  if (!container) return;

  // Reset visibility
  container.style.display = '';

  // Check if text-related properties exist in availableProperties
  const hasShowTextProp = availableProps.some(p => {
    const type = propertyTypes[p];
    return type === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot');
  });
  const hasCellTextProp = availableProps.some(p => {
    const type = propertyTypes[p];
    return type === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second');
  });
  const hasSecondTextProp = availableProps.some(p => p.toLowerCase().includes('second'));
  const hasSlotProp = availableProps.some(p => p.toLowerCase().includes('slot'));

  // Hide text-related fields that duplicate dynamic properties
  if (hasShowTextProp) {
    const showTextField = elements.showText.parentElement;
    if (showTextField) showTextField.style.display = 'none';
  }
  
  if (hasCellTextProp) {
    const cellTextField = elements.cellTextContainer;
    if (cellTextField) cellTextField.style.display = 'none';
  }
  
  if (hasSecondTextProp) {
    const secondTextLineField = elements.secondLineContainer;
    if (secondTextLineField) secondTextLineField.style.display = 'none';
    const secondCellTextField = elements.secondCellTextContainer;
    if (secondCellTextField) secondCellTextField.style.display = 'none';
  }
  
  if (hasSlotProp) {
    const slotField = elements.slotCheckbox.parentElement;
    if (slotField) slotField.style.display = 'none';
  }
  
  // Ensure "Generate sample data using AI" and "Column width" are always visible for body cells.
  elements.generateSampleCheckbox.parentElement!.style.display = '';
  elements.aiPromptContainer.style.display = elements.generateSampleCheckbox.checked ? 'block' : 'none';
  elements.colWidthContainer.style.display = 'block';
}

// Helper to update static fields visibility and values for body cells
function updateStaticFieldsVisibilityAndValues(availableProps: string[], propertyTypes: { [key: string]: any }, props: any) {
  // Find dynamic property keys
  const showTextPropKey = availableProps.find(p => propertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot'));
  const cellTextPropKey = availableProps.find(p => propertyTypes[p] === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second'));
  const secondTextPropKey = availableProps.find(p => p.toLowerCase().includes('second'));
  const slotPropKey = availableProps.find(p => propertyTypes[p] === 'BOOLEAN' && p.toLowerCase().includes('slot'));
  const statePropKey = availableProps.find(p => p.toLowerCase() === 'state' && propertyTypes[p] === 'VARIANT');

  // Hide static 'Second text line' if dynamic one exists
  if (secondTextPropKey && elements.secondLineContainer) {
    elements.secondLineContainer.style.display = 'none';
  }

  // Show/hide static fields
  const showTextField = elements.showText.parentElement;
  if (showTextField) showTextField.style.display = showTextPropKey ? 'none' : '';
  elements.cellTextContainer.style.display = cellTextPropKey ? 'none' : (elements.showText.checked ? '' : 'none');
  elements.secondLineContainer.style.display = secondTextPropKey ? 'none' : (elements.showText.checked ? '' : 'none');
  elements.secondCellTextContainer.style.display = secondTextPropKey ? 'none' : (elements.showText.checked && !!elements.secondCellText.value ? 'block' : 'none');
  const slotField = elements.slotCheckbox.parentElement;
  if (slotField) slotField.style.display = slotPropKey ? 'none' : '';
  const stateField = elements.state.parentElement;
  if(stateField) stateField.style.display = statePropKey ? 'none' : '';

  // Populate static fields from props or reset to default
  elements.showText.checked = showTextPropKey ? (props[showTextPropKey] || false) : false;
  elements.cellText.value = cellTextPropKey ? (props[cellTextPropKey] || '') : '';
  const secondTextValue = secondTextPropKey ? (props[secondTextPropKey] || '') : '';
  elements.secondCellText.value = secondTextValue;
  elements.secondTextLine.checked = !!secondTextValue;
  elements.slotCheckbox.checked = slotPropKey ? (props[slotPropKey] || false) : false;
  elements.state.value = props['State'] || 'Enabled';

  // Always reset AI-related fields
  elements.generateSampleCheckbox.checked = false;
  (document.getElementById('fakerMethodInput') as HTMLInputElement).value = '';
  elements.aiPromptContainer.style.display = 'none';
  elements.cellText.disabled = false;

  // --- For static property fields ---
  if (elements.secondLineContainer && !secondTextPropKey) {
    elements.secondLineContainer.style.display = elements.showText.checked ? '' : 'none';
    if (!elements.showText.checked) {
      elements.secondTextLine.checked = false;
      elements.secondCellTextContainer.style.display = 'none';
    }
  }
}

function openPropertyEditor(key: string) {
  if (state.mode !== 'edit') return;
  try {
    // Ensure colWidthContainer visibility is correct when opening the modal
      if (state.applyMode === 'column') {
      elements.colWidthContainer.style.display = 'block';
    } else {
      elements.colWidthContainer.style.display = 'none';
    }
    // Hide 'apply to column' option for header cells, show for body cells
    const applyToColumnOption = document.querySelector('.apply-option[data-apply="column"]') as HTMLElement;
    if (applyToColumnOption) {
      if (key.startsWith('header-')) {
        applyToColumnOption.style.display = 'none';
      } else {
        applyToColumnOption.style.display = '';
      }
    }
    state.currentEditingCell = key;
    let availableProps: string[] = state.selectedComponent?.availableProperties || [];
    let propertyTypes: { [key: string]: any } = state.selectedComponent?.propertyTypes || {};
    let props: any = {};
    let label = '';
    // Get static property section element
    const staticSection = document.getElementById('staticPropertySection');
    const bodyOnlySection = document.getElementById('bodyOnlyPropertySection');

    // Hide static 'Second text line' if dynamic one exists
    const secondTextPropKey = availableProps.find(p => p.toLowerCase().includes('second'));
    if (secondTextPropKey && elements.secondLineContainer) {
      elements.secondLineContainer.style.display = 'none';
    }

    if (key.startsWith('header-') || key === 'footer') {
      // Hide static property fields for header/footer
      if (staticSection) staticSection.style.display = 'none';
      if (bodyOnlySection) bodyOnlySection.style.display = 'none';
    } else {
      // Show static property fields for body cells
      if (staticSection) staticSection.style.display = '';
      if (bodyOnlySection) bodyOnlySection.style.display = '';
      // Get the current cell's properties, or default to an empty object
      const cellState = getCellState(key);
      props = cellState.properties || {};
      // Centralized static field logic
      updateStaticFieldsVisibilityAndValues(availableProps, propertyTypes, props);
    }

    if (key.startsWith('header-')) {
      const cellState = state.cellProperties.get(key);
      props = (cellState && cellState.properties) ? cellState.properties : {};
      label = `Header ${key.split('-')[1]}`;
      console.log('[DEBUG] openPropertyEditor header', { props });
      renderDynamicPropertyFieldsFromModel(HEADER_CELL_MODEL, props);
    } else if (key === 'footer') {
      const cellState = state.cellProperties.get(key);
      props = (cellState && cellState.properties) ? cellState.properties : {};
      label = 'Footer';
      console.log('[DEBUG] openPropertyEditor footer', { props });
      renderDynamicPropertyFieldsFromModel(FOOTER_MODEL, props);
    } else {
      const cellState = getCellState(key);
      props = cellState.properties || {};
      const [row, col] = key.split(',');
      label = `(${row},${col})`;
      console.log('[DEBUG] openPropertyEditor body', { availableProps, propertyTypes, props });
      renderDynamicPropertyFields(availableProps, propertyTypes, props);
    }

    // Show col width checkbox for body cells, hide for header/footer
    if (key.startsWith('header-') || key === 'footer') {
      elements.colWidthContainer.style.display = 'none';
    } else {
      let width: number | undefined = undefined;
      const col = key.split(',')[1];
        const colCells = Array.from({length: state.gridRows}, (_, r) => `${r+1},${col}`);
        for (const k of colCells) {
          const cellState = state.cellProperties.get(k);
          if (cellState && cellState.colWidth) {
            width = cellState.colWidth;
            break;
          }
        }
      if (width === undefined) {
        width = state.selectedComponent?.width || 100;
      }
      elements.colWidthInput.value = String(width);
    }

    // Set title and show editor
    elements.propertyEditorTitle.innerHTML = `Cell Properties <span>${label}</span>`;
    elements.editingCellCoords.textContent = label;
    elements.propertyEditorOverlay.style.display = 'block';
    elements.propertyEditor.style.display = 'block';
    setTimeout(() => {
      elements.propertyEditor.style.opacity = '1';
    }, 10);
  } catch (error) {
    console.error('Error in openPropertyEditor:', error);
    if (typeof figma !== 'undefined') {
      figma.notify('Failed to open property editor');
    }
  }
}

let pendingFakerContext: null | { mode: 'cell' | 'row' | 'column', key: string, fakerMethod: string } = null;

async function saveCellProperties() {
  try {
    if (!state.currentEditingCell) return;
    const key = state.currentEditingCell;
    let availableProps: string[] = state.selectedComponent?.availableProperties || [];
    let propertyTypes: { [key: string]: any } = state.selectedComponent?.propertyTypes || {};
    if (key.startsWith('header-')) {
      // Use HEADER_CELL_MODEL
      const newProps: any = {};
      for (const fieldDef of HEADER_CELL_MODEL) {
        const { name, type } = fieldDef;
        const input = document.getElementById(`dynamic-${name}`) as HTMLInputElement | HTMLSelectElement;
        if (!input) continue;
        if (type === 'BOOLEAN') {
          newProps[name] = (input as HTMLInputElement).checked;
        } else {
          newProps[name] = input.value;
        }
      }
      console.log('[DEBUG] saveCellProperties header', key, newProps);
      state.cellProperties.set(key, { properties: newProps });
    } else if (key === 'footer') {
      // Use FOOTER_MODEL
      const newProps: any = {};
      for (const fieldDef of FOOTER_MODEL) {
        const { name, type } = fieldDef;
        const input = document.getElementById(`dynamic-${name}`) as HTMLInputElement | HTMLSelectElement;
        if (!input) continue;
        if (type === 'BOOLEAN') {
          newProps[name] = (input as HTMLInputElement).checked;
        } else {
          newProps[name] = input.value;
        }
      }
      console.log('[DEBUG] saveCellProperties footer', key, newProps);
      state.cellProperties.set(key, { properties: newProps });
    } else {
      // Body cell: save properties without duplication
      const props: any = {};
      
      // Check if text-related properties exist in availableProperties
      const hasShowTextProp = availableProps.some(p => propertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot'));
      const hasCellTextProp = availableProps.some(p => {
        const type = propertyTypes[p];
        return type === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second');
      });
      const hasSecondTextProp = availableProps.some(p => p.toLowerCase().includes('second'));
      const hasSlotProp = availableProps.some(p => p.toLowerCase().includes('slot'));
      const hasStateProp = availableProps.some(p => p.toLowerCase() === 'state' && propertyTypes[p] === 'VARIANT');
      
      // Save static properties only if they don't exist in dynamic properties
      if (!hasShowTextProp) {
        const showText = elements.showText.checked;
        if (showText && elements.cellText.value) {
          // Find the correct property key for cell text
          let cellTextProp = availableProps.find(p => {
            const type = propertyTypes[p];
            return type === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second');
          }) || availableProps.find(p => p.toLowerCase().includes('text')) || 'Cell text#12234:16';
          props[cellTextProp] = elements.cellText.value;
        }
        if (showText && elements.secondTextLine.checked && elements.secondCellText.value) {
          // Find the correct property key for second text
          let secondTextProp = availableProps.find(p => p.toLowerCase().includes('second')) || 'secondCellText';
          props[secondTextProp] = elements.secondCellText.value;
        }
      }
      if (!hasStateProp) {
        props['State'] = elements.state.value;
      }
      
      // Save dynamic properties
      for (const propName of availableProps) {
        const type = propertyTypes[propName];
        const input = document.getElementById(`dynamic-${propName}`) as HTMLInputElement | HTMLSelectElement;
        if (!input) continue;
        if (type === 'BOOLEAN') {
          props[propName] = (input as HTMLInputElement).checked;
        } else {
          props[propName] = input.value;
        }
      }
      
      // Col width logic: always set colWidth as a top-level property on first row of column
    let colWidth: number | undefined = undefined;
      if (elements.colWidthInput.value) {
      colWidth = parseInt(elements.colWidthInput.value, 10);
    }
      const applyProps = (key: string, props: any, colWidthTopLevel?: number) => {
        const cellState = getCellState(key);
        const existingProps = { ...cellState };
        cellState.properties = props;
        if (existingProps.isCheckbox !== undefined) {
          cellState.isCheckbox = existingProps.isCheckbox;
        }
        // Set colWidth as a top-level property if provided
        if (colWidthTopLevel !== undefined) {
          cellState.colWidth = colWidthTopLevel;
        } else {
          delete cellState.colWidth;
        }
      };
      const [row, col] = key.split(',').map(Number);
      
      // Always save the properties first (including column width, slot, etc.)
    if (state.applyMode === 'cell') {
        if (colWidth && row === 1) {
          applyProps(state.currentEditingCell, props, colWidth);
        } else {
      applyProps(state.currentEditingCell, props);
        }
    } else if (state.applyMode === 'row') {
      for (let c = 1; c <= state.gridCols; c++) {
        const key = `${row},${c}`;
        if (state.selectedCells.has(key)) {
            if (colWidth && row === 1) {
              applyProps(key, { ...props }, colWidth);
            } else {
          applyProps(key, { ...props });
            }
            const cellState = getCellState(key);
            if (cellState.isCheckbox !== undefined) {
              cellState.isCheckbox = cellState.isCheckbox;
          }
        }
      }
    } else if (state.applyMode === 'column') {
      for (let r = 1; r <= state.gridRows; r++) {
        const key = `${r},${col}`;
        if (state.selectedCells.has(key)) {
            if (colWidth && r === 1) {
              applyProps(key, { ...props }, colWidth);
            } else {
          applyProps(key, { ...props });
            }
            const cellState = getCellState(key);
            if (cellState.isCheckbox !== undefined) {
              cellState.isCheckbox = cellState.isCheckbox;
            }
          }
        }
      }
      
      // Now handle AI sample data generation
      const generateAI = elements.generateSampleCheckbox.checked;
      const fakerMethod = fakerMethodInput.value.trim();
      if (generateAI && fakerMethod) {
        let count = 1;
        let mode: 'cell' | 'row' | 'column' = 'cell';
        if (state.applyMode === 'row') {
          count = state.gridCols;
          mode = 'row';
        } else if (state.applyMode === 'column') {
          count = state.gridRows;
          mode = 'column';
        }
        pendingFakerContext = { mode, key, fakerMethod };
        parent.postMessage({ pluginMessage: { type: 'generate-fake-data', dataType: fakerMethod, count } }, '*');
        return; // Wait for response before closing editor
      }
    }
    closePropertyEditor();
    updateCellVisuals();
    // Visually select the affected cells in the grid
    document.querySelectorAll('.cell').forEach(cell => {
      const divCell = cell as HTMLDivElement;
      const key = `${divCell.dataset.row},${divCell.dataset.col}`;
      if (state.selectedCells.has(key)) {
        divCell.classList.add('selected');
      }
    });
  } catch (error) {
    console.error('Error saving properties:', error);
    if (typeof figma !== 'undefined') figma.notify('Failed to save properties');
  }
}

function updateCellVisuals() {
  const availableProps: string[] = state.selectedComponent?.availableProperties || [];
  // Update body cells
  document.querySelectorAll('.cell').forEach(cell => {
    const divCell = cell as HTMLDivElement;
    const key = `${divCell.dataset.row},${divCell.dataset.col}`;
    const props = state.cellProperties.get(key);
    if (props && props.properties && Object.keys(props.properties).length > 0) {
      let displayText = '';
      // Find any TEXT property to display as cell text
      for (const propName of availableProps) {
        const type = state.selectedComponent?.propertyTypes[propName];
        if (type === 'TEXT' && props.properties[propName]) {
          displayText += props.properties[propName];
          break; // Use the first TEXT property found
        }
      }
      // If no text and slot is checked, show empty
      const slotProp = availableProps.find(p => p.toLowerCase().includes('slot')) || 'slot';
      if (!displayText && props.properties[slotProp]) {
        displayText = '';
      }
      if (!displayText) {
        displayText = `${divCell.dataset.row},${divCell.dataset.col}`;
        divCell.classList.remove('edited');
        divCell.style.fontWeight = 'normal';
      } else {
      divCell.classList.add('edited');
        divCell.style.fontWeight = 'bold';
      }
      divCell.textContent = displayText;
    } else {
      divCell.textContent = `${divCell.dataset.row},${divCell.dataset.col}`;
      divCell.classList.remove('edited');
      divCell.style.fontWeight = 'normal';
    }
  });

  // Update header cells
  const headerGrid = document.getElementById('headerGrid');
  if (headerGrid) {
    for (let c = 1; c <= state.gridCols; c++) {
      const cell = headerGrid.querySelector(`.header-cell:nth-child(${c})`) as HTMLDivElement;
      if (cell) {
        const key = `header-${c}`;
        const props = state.cellProperties.get(key);
        if (props && props.properties && props.properties['Cell text#12234:32']) {
          cell.textContent = props.properties['Cell text#12234:32'];
          cell.classList.add('edited');
          cell.style.fontWeight = 'bold';
        } else {
          cell.textContent = `H${c}`;
          cell.classList.remove('edited');
          cell.style.fontWeight = 'normal';
        }
      }
    }
  }

  // Update footer cell
  const footerGrid = document.getElementById('footerGrid');
  if (footerGrid) {
    const cell = footerGrid.querySelector('.footer-cell') as HTMLDivElement;
    if (cell) {
      const props = state.cellProperties.get('footer');
      if (props && props.properties) {
        // Show Total items, Current page, or Total pages if present, else 'Footer'
        const text = props.properties['Total items#12006:49'] || props.properties['Current page#12006:39'] || props.properties['Total pages#12006:29'] || 'Footer';
        cell.textContent = text;
        cell.classList.add('edited');
        cell.style.fontWeight = text !== 'Footer' ? 'bold' : 'normal';
      } else {
        cell.textContent = 'Footer';
        cell.classList.remove('edited');
        cell.style.fontWeight = 'normal';
      }
    }
  }
}

function updateCreateButtonState() {
    elements.createTableBtn.disabled = !state.sizeConfirmed;
}

// Advanced mapping from user prompt to faker type
function mapPromptToFakerType(prompt: string): string {
  const lower = prompt.toLowerCase();
  if (lower.includes('name') && lower.includes('first')) return 'first name';
  if (lower.includes('name') && lower.includes('last')) return 'last name';
  if (lower.includes('name') || lower.includes('person')) return 'people name';
  if (lower.includes('brand') || lower.includes('company')) return 'brand name';
  if (lower.includes('email')) return 'email';
  if (lower.includes('product')) return 'product';
  if (lower.includes('price')) return 'price';
  if (lower.includes('mobile') || lower.includes('phone')) return 'mobile number';
  if (lower.includes('date')) return 'date';
  if (lower.includes('color')) return 'color';
  if (lower.includes('number') || lower.includes('random')) return 'random number';
  return 'word'; // fallback
}

// Message handling
window.onmessage = (event) => {
  if (!domReady || !elements.grid) return;
  const msg = event.data.pluginMessage;
  if (!msg) return;

  switch (msg.type) {
    case "component-selected":
      state.hasComponent = msg.isValidComponent;
      showMessage(msg.isValidComponent ? `Selected: ${msg.componentName}` : "Please select a component.", msg.isValidComponent ? "success" : "error");
      elements.gridContainer.style.display = "none";
      elements.actionButtons.style.display = "none";
      state.selectedComponent = {
        id: msg.componentId,
        name: msg.componentName,
        width: msg.componentWidth,
        properties: msg.properties || {},
        availableProperties: msg.availableProperties || [],
        propertyTypes: msg.propertyTypes || {}
      };
      if (msg.isValidComponent) {
        showLoader('Scanning table...');
        parent.postMessage({ pluginMessage: { type: 'scan-table' } }, '*');
      }
      break;

    case "selection-cleared":
      state.hasComponent = false;
      showMessage("Please select a table cell component.", "error");
      elements.gridContainer.style.display = "none";
      elements.actionButtons.style.display = "none";
      state.selectedComponent = null;
      elements.landingPage.style.display = 'block';
      break;

    case "table-created":
      hideLoader();
      showMessage("Table created successfully!", "success");
      // Show landing page and componentModeBtn again for new table generation
      elements.landingPage.style.display = 'flex';
      elements.componentModeBtn.style.display = 'inline-block';
      elements.gridContainer.style.display = 'none';
      elements.actionButtons.style.display = 'none';
      elements.propertyEditor.style.display = 'none';
      // Clear all cell properties after table is generated
      state.cellProperties.clear();
      break;

    case "fake-data-response":
      hideLoader();
      const sample = msg.data;
      console.log('[DEBUG] Received fake-data-response', sample);
      if (!Array.isArray(sample) || sample.length === 0) {
        showMessage('No data generated.', 'error');
        return;
      }
      // Apply faker data to the correct cells
      if (pendingFakerContext) {
        const { mode, key, fakerMethod } = pendingFakerContext;
        const availableProps: string[] = state.selectedComponent?.availableProperties || [];
        const propertyTypes: { [key: string]: any } = state.selectedComponent?.propertyTypes || {};

        // Helper to find the correct TEXT property key for a cell
        function getCellTextProp() {
          // Always use the first TEXT property from availableProps
          return availableProps.find(p => propertyTypes[p] === 'TEXT');
        }
        
        // Helper to find the property that controls text visibility
        function getTextVisibilityProp() {
          // Look for a boolean prop that likely controls text visibility
          return availableProps.find(p => 
            propertyTypes[p] === 'BOOLEAN' && 
            (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) &&
            !p.toLowerCase().includes('slot') // Exclude slot toggles
          );
        }

        const applyAiData = (cellKey: string, data: string) => {
          const cellState = getCellState(cellKey);
          const textProp = getCellTextProp();
          const visibilityProp = getTextVisibilityProp();

          if (textProp) {
            cellState.properties[textProp] = data;
          }
          if (visibilityProp) {
            // Ensure the text is visible
            cellState.properties[visibilityProp] = true;
          }
          state.selectedCells.add(cellKey);
        };

        if (mode === 'cell') {
          applyAiData(key, sample[0]);
        } else if (mode === 'row') {
          const [row] = key.split(',').map(Number);
        for (let c = 1; c <= state.gridCols; c++) {
            const k = `${row},${c}`;
            applyAiData(k, sample[(c-1) % sample.length]);
          }
        } else if (mode === 'column') {
          const [, col] = key.split(',').map(Number);
        for (let r = 1; r <= state.gridRows; r++) {
            const k = `${r},${col}`;
            applyAiData(k, sample[(r-1) % sample.length]);
          }
        }

        pendingFakerContext = null;
      updateCellVisuals();
        closePropertyEditor();
        document.querySelectorAll('.cell').forEach(cell => {
          const divCell = cell as HTMLDivElement;
          const key = `${divCell.dataset.row},${divCell.dataset.col}`;
          if (state.selectedCells.has(key)) {
            divCell.classList.add('selected');
          }
        });
      }
      break;

    case "prompt-fallback-notice":
      showMessage(`Could not find a specific category for "${msg.prompt}". Using general text.`, "error");
      break;

    case "creation-error":
      hideLoader();
      showMessage(`Error: ${msg.message}`, "error");
      break;

    case "component-properties":
      state.hasComponent = true;
      state.componentProps = msg.props;
      state.selectedComponent = msg.component;
      state.componentWidth = msg.component.width;
      elements.statusMessage.textContent = `Selected component: ${msg.component.name}`;
      elements.statusMessage.className = 'status success';
      elements.gridContainer.style.display = 'flex';
      break;

    case "scan-table":
      console.log('[PLUGIN] Received scan-table message from UI');
      // ... (your scanning logic here, or just a placeholder for now)
      figma.ui.postMessage({
        type: 'scan-table-result',
        success: true,
        message: 'Scan complete (placeholder).'
      });
      console.log('[PLUGIN] Sent scan-table-result back to UI');
      break;

    case "scan-table-result":
      hideLoader();
      if (elements.scanTableBtn) {
        elements.scanTableBtn.disabled = false;
        elements.scanTableBtn.textContent = 'Scan Table';
      }
      showMessage(msg.message, msg.success ? 'success' : 'error');

      if (!msg.success) return;

      // Hide the initial landing page message
      elements.landingPage.style.display = 'none';

      if (elements.scanTableSection) {
          elements.scanTableSection.style.display = 'none';
      }
      const details = msg.details;
      if (details.bodyCellComponent) {
          state.selectedComponent = details.bodyCellComponent;
          state.hasComponent = true;
          // Store header/footer components for property editing
          state.headerCellComponent = details.headerCellComponent || null;
          state.footerComponent = details.footerComponent || null;
          console.log('Component template set from scan:', state.selectedComponent);
      } else {
          showMessage('Could not find a body cell template in the scanned table.', 'error');
          if(elements.scanTableSection) elements.scanTableSection.style.display = 'flex';
          return;
      }
      // Setup grid
      state.gridCols = details.numCols || 5;
      state.gridRows = 5; // default
      createGrid();
      // Automatically select all cells
      document.querySelectorAll('.cell').forEach(cell => {
          const key = `${(cell as HTMLDivElement).dataset.row},${(cell as HTMLDivElement).dataset.col}`;
          state.selectedCells.add(key);
          cell.classList.add('selected');
      });
      state.sizeConfirmed = true;
      // Show grid UI
      elements.gridContainer.style.display = 'flex';
      elements.actionButtons.style.display = 'flex';
      
      // Remove old scan options if present
      let optionsDiv = document.getElementById('scanOptionsContainer') as HTMLDivElement | null;
      if (optionsDiv) optionsDiv.remove();

      optionsDiv = document.createElement('div');
      optionsDiv.id = 'scanOptionsContainer';
      optionsDiv.style.display = 'flex';
      optionsDiv.style.flexDirection = 'column';
      optionsDiv.style.alignItems = 'flex-start';
      optionsDiv.style.marginBottom = '12px';
      optionsDiv.style.gap = '0px';
      optionsDiv.innerHTML = `
          <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 8px; width: 100%; justify-content: flex-start;">
              <label style="margin-right: 8px;">Rows: <input type="number" id="scanRowsInput" class="styled-input" value="${state.gridRows}" min="1" style="width: 60px;"></label>
              <label>Cols: <input type="number" id="scanColsInput" class="styled-input" value="${state.gridCols}" min="1" style="width: 60px;"></label>
          </div>
          <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 8px;">
              <input type="checkbox" id="scanHeaderToggle" checked> <label for="scanHeaderToggle" style="margin-right: 16px;">Header</label>
              <input type="checkbox" id="scanFooterToggle" ${details.footer ? 'checked' : ''}> <label for="scanFooterToggle" style="margin-right: 16px;">Footer</label>
              <input type="checkbox" id="scanSelectableToggle" ${details.selectCellComponent ? 'checked' : ''}> <label for="scanSelectableToggle" style="margin-right: 16px;">Selectable</label>
              <input type="checkbox" id="scanExpandableToggle" ${details.expandCellComponent ? 'checked' : ''}> <label for="scanExpandableToggle">Expandable</label>
          </div>
      `;
      // Insert scan options before the grid in the grid container
      elements.gridContainer.insertBefore(optionsDiv, elements.gridContainer.firstChild);
      elements.scanOptionsContainer = optionsDiv;

      const updateGridFromInputs = () => {
          const newRows = parseInt((document.getElementById('scanRowsInput') as HTMLInputElement).value, 10);
          const newCols = parseInt((document.getElementById('scanColsInput') as HTMLInputElement).value, 10);
          if (newRows !== state.gridRows || newCols !== state.gridCols) {
              state.gridRows = newRows;
              state.gridCols = newCols;
              createGrid();
              // Re-select all cells after recreating grid
              document.querySelectorAll('.cell').forEach(cell => {
                  const key = `${(cell as HTMLDivElement).dataset.row},${(cell as HTMLDivElement).dataset.col}`;
                  state.selectedCells.add(key);
                  cell.classList.add('selected');
              });
              // --- Filter cellProperties to keep only valid cells/headers/footers ---
              for (const key of Array.from(state.cellProperties.keys())) {
                // Match cell keys like "row,col"
                const match = key.match(/^([0-9]+),([0-9]+)$/);
                if (match) {
                  const row = parseInt(match[1], 10);
                  const col = parseInt(match[2], 10);
                  if (row > state.gridRows || col > state.gridCols) {
                    state.cellProperties.delete(key);
                  }
                }
                // Remove header properties for columns that no longer exist
                if (key.startsWith('header-')) {
                  const col = parseInt(key.split('-')[1], 10);
                  if (col > state.gridCols) {
                    state.cellProperties.delete(key);
                  }
                }
                // Optionally, handle footer if you want to remove it when footer is not present
                // (No-op for now)
              }
              updateCellVisuals();
          }
      };

      document.getElementById('scanRowsInput')?.addEventListener('change', updateGridFromInputs);
      document.getElementById('scanColsInput')?.addEventListener('change', updateGridFromInputs);
      
      // --- Add listeners for header/footer toggles ---
      const headerToggle = document.getElementById('scanHeaderToggle') as HTMLInputElement;
      const footerToggle = document.getElementById('scanFooterToggle') as HTMLInputElement;
      const headerGrid = document.getElementById('headerGrid');
      const footerGrid = document.getElementById('footerGrid');

      const updateHeaderFooterVisibility = () => {
        if (headerGrid) {
          headerGrid.style.display = headerToggle.checked ? 'grid' : 'none';
        }
        if (footerGrid) {
          footerGrid.style.display = footerToggle.checked ? 'grid' : 'none';
        }
      };

      // Set initial visibility based on checkbox state
      updateHeaderFooterVisibility();

      // Add listeners to update visibility on change
      headerToggle?.addEventListener('change', updateHeaderFooterVisibility);
      footerToggle?.addEventListener('change', updateHeaderFooterVisibility);
      
      setMode('edit');
      updateCreateButtonState();
      break;
  }
};

// Setup event listener for the faker dropdown directly
fakerMethodInput.addEventListener('input', () => {
  console.log(`[UI] Found ${fakerMethods.length} total methods from import.`); // DEBUG LOG

  const filter = fakerMethodInput.value.toLowerCase();
  if (!filter) {
    fakerDropdown.style.display = 'none';
    return;
  }
  const filteredMethods = fakerMethods.filter((method: string) => method.toLowerCase().includes(filter));
  fakerDropdown.innerHTML = '';
  filteredMethods.slice(0, 100).forEach(method => {
    const a = document.createElement('a');
    a.href = '#';
    a.textContent = method;
    a.onclick = (e) => {
      e.preventDefault();
      fakerMethodInput.value = method;
      fakerDropdown.style.display = 'none';
    };
    fakerDropdown.appendChild(a);
  });
  fakerDropdown.style.display = 'block';
});

// Hide dropdown when clicking outside
document.addEventListener('click', (e) => {
  if (fakerDropdown && !fakerDropdown.contains(e.target as Node) && e.target !== fakerMethodInput) {
    fakerDropdown.style.display = 'none';
  }
});

function cleanPropName(name: string): string {
  return name.split('#')[0].trim();
}

// Request initial selection state
parent.postMessage({ pluginMessage: { type: "request-selection-state" } }, "*");

// Show colWidthInput if default mode is column
if (state.applyMode === 'column' && elements.colWidthInput) {
  elements.colWidthInput.style.display = 'inline-block';
}

// Set default values
const defaultWidth = state.selectedComponent?.width || 100; // Use component width or fallback to 100
const defaultHeight = 64;
const defaultText = "content";

export {}; // Treat this file as a module 